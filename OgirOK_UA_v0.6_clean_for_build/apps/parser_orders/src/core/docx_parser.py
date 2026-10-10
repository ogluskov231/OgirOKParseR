"""DOCX parser for orders.

Used by:
- GUI: per-person view (mentions + linked blocks)
- Excel autofill: flattened records

This version improves:
- Block segmentation (numbered points/subpoints)
- PIB extraction (strict surname + mixed case + initials)
- Category assignment at parse-time (CategoryEngine)
- Stable spans for UI highlighting (cat_span)

Backward compatibility:
- Keeps function name build_combined_v2
- Keeps legacy record keys used by excel_autofill
"""

from __future__ import annotations

import json
import re
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

from docx import Document
from openpyxl import load_workbook

try:
    from apps.parser_orders.src.core.json_schema import validate_order
except Exception:  # pragma: no cover
    validate_order = None  # type: ignore

try:
    from apps.parser_orders.src.core.category_engine import CategoryEngine
except Exception:  # pragma: no cover
    CategoryEngine = None  # type: ignore


UA_MONTHS = {
    "січня": 1,
    "лютого": 2,
    "березня": 3,
    "квітня": 4,
    "травня": 5,
    "червня": 6,
    "липня": 7,
    "серпня": 8,
    "вересня": 9,
    "жовтня": 10,
    "листопада": 11,
    "грудня": 12,
}

RANK_WORDS = [
    "солдат",
    "старший солдат",
    "молодший сержант",
    "сержант",
    "старший сержант",
    "головний сержант",
    "штаб-сержант",
    "старшина",
    "молодший лейтенант",
    "лейтенант",
    "старший лейтенант",
    "капітан",
    "майор",
    "підполковник",
    "полковник",
]

# Strict PIB in typical orders: UPPERCASE surname + Name + Patronymic
PIB_UA_STRICT_RE = re.compile(
    r"\b(?P<surname>[А-ЯІЇЄҐ]{2,}(?:[-'’][А-ЯІЇЄҐ]{2,})?)\s+"
    r"(?P<name>[А-ЯІЇЄҐ][а-яіїєґ]+)\s+"
    r"(?P<pat>[А-ЯІЇЄҐ][а-яіїєґ]+)\b"
)

# Mixed-case PIB
PIB_RE = re.compile(
    r"\b([А-ЯІЇЄҐA-Z][А-ЯІЇЄҐA-Zа-яіїєґA-Za-z'’\-]{2,})\s+"
    r"([А-ЯІЇЄҐA-Z][а-яіїєґA-Za-z'’\-]{1,})\s+"
    r"([А-ЯІЇЄҐA-Z][а-яіїєґA-Za-z'’\-]{1,})\b"
)

# Surname + initials: ІВАНЕНКО І.О.
PIB_INITIALS_RE = re.compile(
    r"\b(?P<surname>[А-ЯІЇЄҐ]{2,}(?:[-'’][А-ЯІЇЄҐ]{2,})?)\s+"
    r"(?P<i1>[А-ЯІЇЄҐ])\.\s*(?P<i2>[А-ЯІЇЄҐ])\.\b"
)

# surname variants like: "МЕНЮКА - Менкевича" (no name/patronymic)
SURNAME_VARIANT_RE = re.compile(
    r"\b(?P<s1>[А-ЯІЇЄҐ]{2,})\s*[-–—]\s*(?P<s2>[А-ЯІЇЄҐ][а-яіїєґ]+(?:[-'’][А-ЯІЇЄҐ][а-яіїєґ]+)?)\b"
)

# block boundary signals
POINT_RE = re.compile(r"^\s*(?P<num>\d{1,3})[.)]\s+")
SUBPOINT_RE = re.compile(r"^\s*(?P<ch>[а-яіїєґ])\)\s+")
DASH_RE = re.compile(r"^\s*[–—-]\s+")
ORDER_NO_RE = re.compile(r"(?:^|\s)№\s*([0-9]+(?:[-–—][A-ZА-ЯІЇЄҐ0-9]+)?)", re.IGNORECASE)


# Heuristics to reduce false positives (org/geo names matching PIB patterns).
_NON_PERSON_WORDS = {
    # geo/admin
    "республіка",
    "область",
    "район",
    "місто",
    "село",
    "громада",
    "україна",
    "україни",
    "крим",
    # org keywords
    "міністерство",
    "оборони",
    "генерального",
    "штабу",
    "військова",
    "військової",
    "частина",
    "частини",
    "бригада",
    "батальйон",
    "рота",
    "в/ч",
}

_NON_PERSON_SURNAMES = {
    "АВТОНОМНА",
    "ДЕРЖАВНА",
    "ВІЙСЬКОВА",
    "МІНІСТЕРСТВО",
    "ГЕНЕРАЛЬНОГО",
}


# Blocks containing these phrases are ignored (do not parse people/actions).
IGNORE_BLOCK_PHRASES = [
    "постанова кабінету міністрів",
]

# Special (non-standard) categories driven by headings/section context.
# These are *not* parsed from words inside the paragraph; they are inferred from nearby headings.
_SPECIAL_SECTION_RULES = [
    {
        "id": "special_detach_med",
        "label": "Прикомандировано до медичної роти (в/ч А0989)",
        "event_type": "departure",
        "match_all": ["прик", "медичн", "а0989"],
    },
    {
        "id": "special_annual_leave_return",
        "label": "З щорічної основної відпустки",
        "event_type": "arrival",
        "match_all": ["щорічн", "основн", "відпустк"],
    },
    {
        "id": "special_trip_departure",
        "label": "У відрядження",
        "event_type": "departure",
        "match_any": ["у відрядження", "вважати такими, що вибули"],
    },
]


def _special_category_from_section_path(section_path: List[str]) -> Optional[Dict[str, str]]:
    """Infer special category from recent heading context."""
    if not section_path:
        return None
    sp = " | ".join([str(s) for s in section_path if s]).lower()

    for rule in _SPECIAL_SECTION_RULES:
        ma = [str(x).lower() for x in (rule.get("match_all") or []) if x]
        my = [str(x).lower() for x in (rule.get("match_any") or []) if x]

        ok_all = True if not ma else all(k in sp for k in ma)
        ok_any = True if not my else any(k in sp for k in my)

        if ok_all and ok_any:
            return {"id": rule["id"], "label": rule["label"], "event_type": rule["event_type"]}
    return None


def _is_plausible_person(surname: str, name: str, pat: str, context_text: str) -> bool:
    """Cheap filters to suppress obvious non-person matches."""
    s = (surname or "").strip().upper()
    n = (name or "").strip().lower()
    p = (pat or "").strip().lower()
    if not s or not n or not p:
        return False
    if s in _NON_PERSON_SURNAMES:
        return False
    if n in _NON_PERSON_WORDS or p in _NON_PERSON_WORDS:
        return False

    # If context looks like a header/organization line, be stricter.
    ctx = (context_text or "").lower()
    if any(w in ctx for w in ("міністер", "генераль", "штаб", "україни")) and s in {"АВТОНОМНА", "ДЕРЖАВНА"}:
        return False

    # Common patronymic endings; if absent, allow but reduce noise by requiring rank words.
    if not re.search(r"(ич|вич|йович|ьович|івна|ївна|овна)$", p):
        if not any(rw in ctx for rw in ("солдат", "сержант", "лейтенант", "капітан", "майор", "підполковник", "полковник")):
            return False
    return True


def _norm_space(s: str) -> str:
    s = (s or "").replace("\u00a0", " ")
    s = s.replace("\u2019", "'").replace("\u2018", "'")
    return re.sub(r"\s+", " ", s).strip()


def _uppercase_ratio(s: str) -> float:
    letters = [c for c in s if c.isalpha()]
    if not letters:
        return 0.0
    ups = [c for c in letters if c == c.upper()]
    return len(ups) / max(1, len(letters))



def _is_heading(line: str) -> bool:
    t = (line or "").strip()
    if not t:
        return False
    up = t.upper()

    # Explicit ignore
    if "ПОСТАНОВА КАБІНЕТУ МІНІСТРІВ" in up:
        return False

    # Explicit section markers (single-word headings)
    if up in {"НАДХОДЖЕННЯ", "НАДХООДЖЕННЯ", "НАДІЙШЛО", "НАДІЙШЛИ"}:
        return True
    if up in {"ВИБУТТЯ", "ВИБУТТЯ З ЗАП РОТ", "ВИБУТТЯ ІЗ ЗАП РОТ"} or up.startswith("ВИБУТТЯ"):
        return True

    # Special headings required by workflow
    if up.startswith("З ЩОРІЧНОЇ ОСНОВНОЇ ВІДПУСТКИ"):
        return True
    if up.startswith("У ВІДРЯДЖЕННЯ"):
        return True

    if t.endswith(":"):
        return True
    if up in {"НАКАЗ", "МІНІСТЕРСТВО ОБОРОНИ УКРАЇНИ"}:
        return True
    if len(t) <= 110 and _uppercase_ratio(t) >= 0.75 and len(t.split()) >= 2:
        return True
    if t.startswith("Вважати ") or t.startswith("Нижчепойменованих"):
        return True
    return False


def _extract_text(docx_path: str) -> List[str]:
    """Extract text from DOCX preserving paragraph breaks.

    - empty paragraphs are preserved as "" (needed for splitting)
    - tables are added as separate lines (prefixed)
    """
    doc = Document(docx_path)
    lines: List[str] = []

    for p in doc.paragraphs:
        raw = p.text or ""
        parts = raw.splitlines() if "\n" in raw else [raw]
        for part in parts:
            lines.append(_norm_space(part))

    for table in doc.tables:
        for row in table.rows:
            cells = [_norm_space(c.text) for c in row.cells]
            t = " | ".join([c for c in cells if c])
            if t:
                lines.append("TABLE: " + t)

    # collapse excessive blank runs
    cleaned: List[str] = []
    blank = False
    for t in lines:
        if not t:
            if not blank:
                cleaned.append("")
            blank = True
        else:
            cleaned.append(t)
            blank = False
    return cleaned


def _split_blocks(lines: List[str]) -> List[List[str]]:
    """Split extracted lines into logical blocks.

    Boundaries:
    - blank line
    - headings are isolated
    - numbered points/subpoints/dash start new blocks
    """
    blocks: List[List[str]] = []
    cur: List[str] = []

    def flush() -> None:
        nonlocal cur
        if cur:
            blocks.append(cur)
            cur = []

    for ln in lines:
        if not ln:
            flush()
            continue

        if _is_heading(ln):
            flush()
            blocks.append([ln])
            continue

        # start new block when we see explicit enumerations
        if cur and (POINT_RE.match(ln) or SUBPOINT_RE.match(ln) or DASH_RE.match(ln)):
            flush()

        cur.append(ln)

    flush()
    return blocks


def _canon_pib(surname: str, name: str = "", pat: str = "") -> str:
    surname = (surname or "").strip()
    name = (name or "").strip()
    pat = (pat or "").strip()
    if surname:
        surname = surname.upper()
    if name:
        name = name.capitalize()
    if pat:
        pat = pat.capitalize()
    parts = [p for p in [surname, name, pat] if p]
    return " ".join(parts)


def _extract_people(block_text: str, surname_to_full: Dict[str, str]) -> List[str]:
    """Extract people identifiers, prioritizing declined military ranks."""
    out: List[str] = []
    occupied: List[Tuple[int, int]] = []

    rank_alt = "|".join(re.escape(x) for x in sorted(set(RANK_WORDS), key=len, reverse=True))
    ranked_re = re.compile(
        rf"\b(?P<rank>{rank_alt})[а-яіїєґA-ZА-ЯІЇЄҐ-]*"
        rf"(?:\s+(?P<reserve>запасу|резерву))?"
        rf"\s+(?P<surname>[А-ЯІЇЄҐ][а-яіїєґ'’\-]+)"
        rf"\s+(?P<name>[А-ЯІЇЄҐ][а-яіїєґ'’\-]+)"
        rf"\s+(?P<pat>[А-ЯІЇЄҐ][а-яіїєґ'’\-]+)\b",
        re.IGNORECASE
    )

    for m in ranked_re.finditer(block_text):
        s, n, p = m.group("surname"), m.group("name"), m.group("pat")
        if not _is_plausible_person(s, n, p, block_text):
            continue
        pib = _canon_pib(s, n, p)
        out.append(pib)
        surname_to_full[(s or "").upper()] = pib
        occupied.append(m.span())

    def overlaps(span: Tuple[int, int]) -> bool:
        return any(span[0] < y and span[1] > x for x, y in occupied)

    for m in PIB_UA_STRICT_RE.finditer(block_text):
        if overlaps(m.span()):
            continue
        s, n, p = m.group("surname"), m.group("name"), m.group("pat")
        if not _is_plausible_person(s, n, p, block_text):
            continue
        pib = _canon_pib(s, n, p)
        out.append(pib)
        surname_to_full[(s or "").upper()] = pib

    for m in PIB_RE.finditer(block_text):
        if overlaps(m.span()):
            continue
        s, n, p = m.group(1), m.group(2), m.group(3)
        if not _is_plausible_person(s, n, p, block_text):
            continue
        pib = _canon_pib(s, n, p)
        out.append(pib)
        surname_to_full[(s or "").upper()] = pib

    for m in PIB_INITIALS_RE.finditer(block_text):
        s = (m.group("surname") or "").upper().strip()
        if s and s in surname_to_full:
            out.append(surname_to_full[s])
        else:
            out.append(f"{s} {m.group('i1')}.{m.group('i2')}. (ініціали)")

    for m in SURNAME_VARIANT_RE.finditer(block_text):
        s1 = (m.group("s1") or "").strip()
        s2 = (m.group("s2") or "").strip()
        if s1 and s2:
            out.append(f"{s1.upper()}-{s2.capitalize()} (без імені)")

    seen = set()
    uniq: List[str] = []
    for p in out:
        if p and p not in seen:
            seen.add(p)
            uniq.append(p)
    return uniq


def _parse_date_any(s: str) -> str:
    s = (s or "").strip()
    m = re.search(r"(\d{1,2})[./](\d{1,2})[./](\d{2,4})", s)
    if m:
        d, mo, y = int(m.group(1)), int(m.group(2)), int(m.group(3))
        if y < 100:
            y += 2000
        return f"{d:02d}.{mo:02d}.{y:04d}"
    m = re.search(r"(\d{1,2})\s+([а-яіїє]+)\s+(\d{4})", s.lower())
    if m and m.group(2) in UA_MONTHS:
        d = int(m.group(1))
        mo = UA_MONTHS[m.group(2)]
        y = int(m.group(3))
        return f"{d:02d}.{mo:02d}.{y:04d}"
    return ""


def _parse_order_ref(lines: List[str]) -> str:
    # Prefer the line containing the actual order header, e.g.
    # ``15.06.2025 м. Київ № 101``.
    for line in lines[:25]:
        m = re.search(
            r"(\d{1,2}[./]\d{1,2}[./]\d{4})\s+м\.\s+(.+?)\s+№\s*([0-9]+(?:[-–—][A-ZА-ЯІЇЄҐ0-9]+)?)\s*$",
            line,
            re.IGNORECASE
        )
        if m:
            dt = _parse_date_any(m.group(1))
            return f"№{m.group(3)} від {dt}" if dt else f"№{m.group(3)}"

    header = " ".join([x for x in lines[:25] if x])
    # Fallback: use a number from the header, but never use a legal reference
    # with slash-separated identifiers such as 322/1/78.
    nums = re.findall(r"(?:^|\s)№\s*([0-9]+(?:[-–—][A-ZА-ЯІЇЄҐ0-9]+)?)", header, re.IGNORECASE)
    num = nums[-1].strip() if nums else ""
    if num:
        return f"№{num}"

    return ""


def _rank_category(rank: str) -> str:
    r = (rank or "").lower()
    if any(x in r for x in ["лейтенант", "капітан", "майор", "підполковник", "полковник"]):
        return "Офіцер"
    if "серж" in r or "старшина" in r:
        return "Сержант"
    if r:
        return "Солдат"
    return ""


def _load_events_cfg(learning_dir: str) -> Dict[str, List[str]]:
    p = Path(learning_dir) / "events.json"
    if p.exists():
        try:
            with open(p, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {"arrival": [], "departure": []}
    return {"arrival": [], "departure": []}


def _load_template_dicts(template_path: str) -> Dict[str, List[str]]:
    """Source-of-truth dictionaries from sheet 'Категорії' (optional)."""
    try:
        wb = load_workbook(template_path, data_only=True)
    except Exception:
        return {"units": [], "arrivals": [], "reasons": []}

    if "Категорії" not in wb.sheetnames:
        return {"units": [], "arrivals": [], "reasons": []}

    ws = wb["Категорії"]
    units, arrivals, reasons = [], [], []

    for r in range(3, 600):
        v = ws[f"A{r}"].value
        if v:
            units.append(str(v).strip())

    for r in range(2, 600):
        v = ws[f"B{r}"].value
        if v:
            arrivals.append(str(v).strip())

    for r in range(2, 600):
        v = ws[f"C{r}"].value
        if v:
            reasons.append(str(v).strip())

    return {"units": units, "arrivals": arrivals, "reasons": reasons}


def _best_match_from_list(text: str, items: List[str]) -> str:
    low = (text or "").lower()
    for it in sorted(items, key=len, reverse=True):
        if it and it.lower() in low:
            return it
    return ""


def _extract_position(text: str) -> str:
    low = (text or "").lower()
    for pat in [
        r"(?:на посаду)\s+([^,.]+)",
        r"(?:призначити)\s+на\s+посаду\s+([^,.]+)",
        r"(?:посада)\s+([^,.]+)",
    ]:
        m = re.search(pat, low)
        if m:
            return _norm_space(m.group(1))[:120]
    return ""


def _rank_from_text(text: str) -> str:
    low = (text or "").lower()
    for rw in sorted(RANK_WORDS, key=len, reverse=True):
        if rw in low:
            return rw
    return ""



def _update_section_hint(line: str) -> Optional[str]:
    up = (line or "").upper()

    # Ignore legal/boilerplate references
    if "ПОСТАНОВА КАБІНЕТУ МІНІСТРІВ" in up:
        return None

    # Headings
    if "НАДХОДЖ" in up or "НАДІЙШ" in up:
        return "arrival"
    if "ВИБУТТ" in up:
        return "departure"
    if "ВІДРЯДЖЕНН" in up:
        return "departure"
    if "ВІДПУСТК" in up and up.strip().startswith("З "):
        return "arrival"
    if "ПРИКОМАНДИР" in up:
        return "departure"

    if any(k in up for k in ["ПРИЗНАЧИТИ", "ЗАРАХУВАТИ", "ПРИЙНЯТИ", "ВКЛЮЧИТИ", "ПРИБУВ", "ПРИБУЛИ"]):
        return "arrival"
    if any(k in up for k in ["ЗВІЛЬНИТИ", "ВИКЛЮЧИТИ", "ВИБУТИ", "ПЕРЕВЕСТИ", "ВИРАХУВАТИ", "СЗЧ", "ПОЛОН"]):
        return "departure"
    return None


def _event_type_from_keywords(text: str, section_hint: str, events_cfg: Dict[str, List[str]]) -> str:
    low = (text or "").lower()
    if section_hint in {"arrival", "departure"}:
        return section_hint

    for kw in events_cfg.get("arrival", []) or []:
        if kw and kw.lower() in low:
            return "arrival"
    for kw in events_cfg.get("departure", []) or []:
        if kw and kw.lower() in low:
            return "departure"

    if any(k in low for k in ["зарахув", "признач", "прибув", "включит", "прийняти"]):
        return "arrival"
    if any(k in low for k in ["виключ", "вибув", "перевед", "звільн", "вирахув", "сзч", "зник", "полон"]):
        return "departure"
    return "unknown"


def _continuation_allowed(text: str) -> bool:
    t = (text or "").lstrip()
    if not t:
        return False
    low = t.lower()
    starters = (
        "підстава:",
        "підстави:",
        "з ",
        "у ",
        "відповідно",
        "на виконання",
        "виплатити",
        "щоріч",
        "відпуст",
        "грош",
        "матері",
        "допомог",
        "продовжити",
        "увільнити",
    )
    if any(low.startswith(s) for s in starters):
        return True
    if len(t) <= 260 and any(k in low for k in ("грив", "%", "по ", "рік", "місяц")):
        return True
    return False


def _make_category_engine(
    rules_path: str = "config/categories/event_categories_rules_v1.json",
    user_rules_path: str = "config/categories/user_category_rules.json",
) -> Optional[Any]:
    if CategoryEngine is None:
        return None
    try:
        rp = Path(rules_path)
        if not rp.exists():
            return None
        urp = Path(user_rules_path)
        return CategoryEngine(str(rp), str(urp) if urp.exists() else None)
    except Exception:
        return None



def _parse_pib_parts(pib: str) -> Tuple[str, str, str]:
    """Split canonical PIB (SURNAME Name Patronymic) without inventing values."""
    parts = _norm_space(pib).split()
    if len(parts) >= 3:
        return parts[0], parts[1], " ".join(parts[2:])
    if len(parts) == 2:
        return parts[0], parts[1], ""
    return (parts[0] if parts else "", "", "")


def _action_type_for_schema(text: str, event_type: str, category: Optional[str]) -> str:
    low = (text or "").lower()

    if "припинити доступ" in low or ("доступ" in low and "державн" in low and "таємниц" in low):
        return "security_access"
    if "присвоїти" in low and "зван" in low:
        return "rank_assignment"
    if "прийняти на роботу" in low:
        return "employment"
    if "відпустк" in low:
        return "leave"
    if "відряджен" in low or "відрядження" in low:
        return "business_trip"
    if "навчальн" in low and "збор" in low:
        return "training_assembly"
    if "виключити" in low and ("списків" in low or "забезпечення" in low):
        return "exclusion"
    if re.search(r"\\b(звільнити|звільнений|звільнено)\\b", low):
        return "dismissal"
    if "перевести" in low or "переведенн" in low:
        return "transfer"
    if any(k in low for k in ("здав і вибув", "здав справи", "вибув до нового місця", "вважати таким, що вибув")):
        return "departure"
    if any(k in low for k in ("призначити на посаду", "призначений", "призначено")) and "вважати" not in low:
        return "appointment"
    if "зарахув" in low:
        return "enrollment"
    if "прибув" in low or "приступив до виконання" in low:
        return "arrival"
    if any(k in low for k in ("виплатити", "виплачувати", "премі", "надбав", "компенсац", "грошову допомогу")):
        return "payment"
    if "резерв" in low and "контракт" in low:
        return "reserve_enrollment"
    if event_type in ("arrival", "departure"):
        return event_type
    return "other"


def _person_category_for_schema(pib: str, rank: str, text: str) -> str:
    low = (text or "").lower()
    if "резерв" in low and "військовозобов" in low:
        return "military_liable"
    if "резерв" in low or "резервіст" in low:
        return "reservist"
    if "прийняти на роботу" in low or "неповним робочим" in low or "посадовим окладом" in low and not rank:
        return "civilian_employee"
    return "military_personnel"


def _empty_schema_item() -> Dict[str, Any]:
    return {
        "subnumber": None,
        "person": {"category": None,"rank": None,"last_name": None,"first_name": None,"middle_name": None,"full_name": None,"position": None,"previous_position": None,"new_position": None,"unit": None,"previous_unit": None,"new_unit": None},
        "dates": {"effective": None,"start": None,"end": None},
        "service": {"service_years": None,"calendar_service_years": None,"calendar_service_months": None,"preferential_service_years": None,"preferential_service_months": None,"military_specialty": None,"shpk": None,"tariff_grade": None,"rank_previous": None,"rank_new": None,"contract_duration_years": None,"military_registration": None},
        "employment": {"position": None,"salary": None,"currency": None,"employment_type": None,"working_hours": None,"working_days": None,"probation": None,"probation_enabled": None,"probation_months": None,"payment_description": None},
        "leave": {"type": None,"type_text": None,"duration_days": None,"destination": None,"paid": None,"exclude_from_provisions": None,"travel_documents": None,"travel_money": None},
        "business_trip": {"destination": None,"purpose": None,"duration_days": None,"transport": None},
        "payments": [],
        "security": {"secret_access": False,"secret_levels": [],"access_action": None,"allowance_percent": None},
        "documents": [],
        "basis": None,
        "raw_text": None,
    }


def _to_schema_document(doc_struct: Dict[str, Any]) -> Dict[str, Any]:
    """Map legacy parser output to the embedded JSON Schema without losing raw text."""
    out = {"order": {"number": None,"date": None,"city": None,"unit": None,"type": "наказ"}, "items": []}
    order_ref = str(doc_struct.get("order_ref") or "")
    m = re.search(r"№\s*([^\s]+)", order_ref)
    d = _parse_date_any(order_ref)
    out["order"]["number"] = m.group(1) if m else None
    out["order"]["date"] = d or None
    for idx, rec in enumerate(doc_struct.get("records") or [], start=1):
        text = str(rec.get("note") or "")
        item = _empty_schema_item()
        # Preserve the actual order point/subpoint from the source text.
        nm = re.match(r"^\s*(\d+(?:\.\d+)?)\.\s+", text)
        if nm:
            point = nm.group(1)
            if "." in point:
                item["number"], item["subnumber"] = point.split(".", 1)[0], point
            else:
                item["number"] = point
        else:
            item["number"] = str(rec.get("block_id") or idx)
        item["action"] = {"type": _action_type_for_schema(text, str(rec.get("event_type") or ""), rec.get("category")), "text": text}
        surname, name, pat = _parse_pib_parts(str(rec.get("pib") or ""))
        rank = str(rec.get("rank") or "") or None
        item["person"].update({"category": _person_category_for_schema(str(rec.get("pib") or ""), rank or "", text),"rank": rank,"last_name": surname or None,"first_name": name or None,"middle_name": pat or None,"full_name": str(rec.get("pib") or "") or None,"position": rec.get("position") or None,"unit": rec.get("from_unit") or rec.get("to_unit") or None,"previous_unit": rec.get("from_unit") or None,"new_unit": rec.get("to_unit") or None})
        dt = rec.get("arrival_date") or rec.get("departure_date") or None
        if dt:
            if rec.get("event_type") == "arrival": item["dates"]["start"] = dt
            elif rec.get("event_type") == "departure": item["dates"]["effective"] = dt
            else: item["dates"]["effective"] = dt
        item["service"]["military_specialty"] = None
        item["employment"]["position"] = rec.get("position") or None
        item["basis"] = None
        item["raw_text"] = text or None
        # Preserve parser-specific category/rule in a schema-safe document-level extension is intentionally avoided.
        out["items"].append(item)
    return out


def parse_docx_to_structured(
    docx_path: str,
    template_path: str,
    learning_dir: str = "config/learning",
    category_rules_path: str = "config/categories/event_categories_rules_v1.json",
    user_category_rules_path: str = "config/categories/user_category_rules.json",
) -> Tuple[Dict[str, Any], Dict[str, Any]]:
    """Parse a DOCX order into structured blocks.

    Output contains:
      - blocks: list of blocks with section context + inferred action + date + category fields
      - people_index: pib -> list of block_ids where the person is mentioned
      - records: flattened records for Excel autofill

    New fields (non-breaking):
      - block.category / rule_id / cat_span
      - mention.category / rule_id / cat_span
      - record.category / rule_id / cat_span
    """

    events_cfg = _load_events_cfg(learning_dir)
    dicts = _load_template_dicts(template_path)
    cat_engine = _make_category_engine(category_rules_path, user_category_rules_path)

    lines = _extract_text(docx_path)
    non_empty = [x for x in lines if x]
    order_ref = _parse_order_ref(non_empty)

    blocks_raw = _split_blocks(lines)

    section_path: List[str] = []
    section_hint = ""
    last_date = ""

    blocks: List[Dict[str, Any]] = []
    people_index: Dict[str, List[int]] = {}
    records: List[Dict[str, Any]] = []

    surname_to_full: Dict[str, str] = {}

    # Active context (continuation blocks): attach to previously mentioned person(s)
    active_pibs: List[str] = []
    active_block_id: Optional[int] = None
    active_record_idxs: List[int] = []

    def set_heading(h: str) -> None:
        nonlocal section_path
        h2 = _norm_space(h).rstrip(":")
        if not h2:
            return
        section_path = (section_path + [h2])[-3:]

    def attach_to_active(extra_text: str, dt_eff: str) -> bool:
        nonlocal active_pibs, active_block_id, active_record_idxs
        if not active_pibs or active_block_id is None:
            return False

        # append to last block text
        for blk in blocks:
            if blk.get("id") == active_block_id:
                blk["text"] = (blk.get("text") or "") + "\n" + extra_text
                if dt_eff and not blk.get("date"):
                    blk["date"] = dt_eff
                # reclassify category with new text to keep span consistent
                if cat_engine:
                    res = cat_engine.classify(blk["text"], event_hint=blk.get("event_type") or None)
                    blk["event_type"] = res.get("event_type") or blk.get("event_type")
                    blk["category"] = res.get("category")
                    blk["rule_id"] = res.get("rule_id")
                    blk["cat_span"] = res.get("match_span")
                    blk["cat_match_text"] = res.get("match_text")
                break

        for idx in active_record_idxs:
            if 0 <= idx < len(records):
                records[idx]["note"] = (records[idx].get("note") or "") + "\n" + extra_text
        return True

    for b_id, b_lines in enumerate(blocks_raw, start=1):
        if len(b_lines) == 1 and _is_heading(b_lines[0]):
            set_heading(b_lines[0])
            sec = _update_section_hint(b_lines[0])
            if sec:
                section_hint = sec
            continue

        text = "\n".join(b_lines)

        # Ignore boilerplate/legal paragraphs (must not create records)
        low_text = (text or "").lower()
        if any(ph in low_text for ph in IGNORE_BLOCK_PHRASES):
            active_pibs = []
            active_block_id = None
            active_record_idxs = []
            continue


        # update section hint from strong verbs inside block
        for ln in b_lines:
            sec = _update_section_hint(ln)
            if sec:
                section_hint = sec

        # contextual date
        dt = ""
        for ln in b_lines:
            dt = _parse_date_any(ln) or dt
        if dt:
            last_date = dt
        dt_eff = dt or last_date

        # people
        pibs = _extract_people(text, surname_to_full)

        if not pibs:
            if _continuation_allowed(text) and attach_to_active(text, dt_eff):
                continue
            active_pibs = []
            active_block_id = None
            active_record_idxs = []
            continue

        # event type + category
        event_type_hint = _event_type_from_keywords(text, section_hint, events_cfg)
        event_type = event_type_hint
        category = None
        rule_id = None
        cat_span = None
        cat_match_text = None

        if cat_engine:
            res = cat_engine.classify(text, event_hint=event_type_hint)
            event_type = res.get("event_type") or event_type_hint
            category = res.get("category")
            rule_id = res.get("rule_id")
            cat_span = res.get("match_span")
            cat_match_text = res.get("match_text")

        # Special categories based on headings/section context (override rule-based category).
        sp = _special_category_from_section_path(section_path)
        if sp:
            category = sp.get("label")
            rule_id = sp.get("id")
            cat_span = None
            cat_match_text = None
            if sp.get("event_type") in ("arrival", "departure"):
                event_type = sp.get("event_type")

        rank_raw = _rank_from_text(text)
        unit = _best_match_from_list(text, dicts.get("units", []))
        reason = _best_match_from_list(text, dicts.get("reasons", [])) if event_type == "departure" else ""
        arrival_src = _best_match_from_list(text, dicts.get("arrivals", [])) if event_type == "arrival" else ""
        pos = _extract_position(text)

        block = {
            "id": b_id,
            "section_path": list(section_path),
            "event_type": event_type,
            "date": dt_eff,
            "text": text,
            "category": category,
            "rule_id": rule_id,
            "cat_span": cat_span,
            "cat_match_text": cat_match_text,
        }
        blocks.append(block)

        # set active
        active_pibs = list(pibs)
        active_block_id = b_id
        active_record_idxs = []

        for pib in pibs:
            people_index.setdefault(pib, []).append(b_id)

            # Backward-compatible record used by autofill.
            rec = {
                "pib": pib,
                "event_type": event_type if event_type in ("arrival", "departure") else "arrival",
                "is_attached": bool("придан" in text.lower()),
                "order_number": order_ref,
                "rank": rank_raw.title() if rank_raw else "",
                "rank_category": _rank_category(rank_raw),
                "service_type": "",
                "from_unit": unit if event_type == "departure" else "",
                "to_unit": unit if event_type == "arrival" else "",
                "position": pos,
                "arrival_date": dt_eff if event_type == "arrival" else "",
                "departure_date": dt_eff if event_type == "departure" else "",
                "arrival_source": arrival_src,
                "departure_reason": reason,
                "birth_date": "",
                "note": text,
                "block_id": b_id,
                "section_path": list(section_path),
                # New (non-breaking)
                "doc": Path(docx_path).name,
                "source_docx": str(docx_path),
                "category": category,
                "rule_id": rule_id,
                "cat_span": cat_span,
            }
            records.append(rec)
            active_record_idxs.append(len(records) - 1)

    structured_schema = _to_schema_document({
        "order_ref": order_ref,
        "records": records,
    })
    schema_valid = None
    schema_errors: List[str] = []
    if validate_order is not None:
        schema_valid, schema_errors = validate_order(structured_schema)

    structured = {
        "schema": "doc_struct_v4",
        "source_docx": str(docx_path),
        "doc_name": Path(docx_path).name,
        "order_ref": order_ref,
        "blocks": blocks,
        "people_index": people_index,
        "records": records,
        "structured_schema": structured_schema,
        "schema_validation": {"valid": bool(schema_valid) if schema_valid is not None else False, "errors": schema_errors},
    }

    meta = {
        "source_docx": str(docx_path),
        "order_ref": order_ref,
        "blocks": len(blocks),
        "records": len(records),
    }

    return structured, meta


def build_combined_v2(docx_files: List[str], template_path: str, learning_dir: str = "config/learning") -> Dict[str, Any]:
    """Combined parser output used by GUI + autofill.

    Kept as build_combined_v2 for backward compatibility.
    """
    people: Dict[str, Dict[str, Any]] = {}
    meta_docs: List[Dict[str, Any]] = []
    schema_documents: List[Dict[str, Any]] = []

    all_blocks: Dict[str, Dict[int, Dict[str, Any]]] = {}
    errors: List[Dict[str, str]] = []
    processed_files: List[str] = []

    # Process every DOCX independently. A damaged/locked file must not stop
    # the remaining files in a selected folder from being parsed.
    for f in docx_files:
        f = str(Path(f))
        try:
            doc_struct, meta = parse_docx_to_structured(f, template_path=template_path, learning_dir=learning_dir)
        except Exception as exc:
            errors.append({
                "file": f,
                "error_type": type(exc).__name__,
                "error": str(exc),
            })
            continue

        meta["source_docx"] = f
        meta_docs.append(meta)
        processed_files.append(f)
        schema_documents.append(doc_struct.get("structured_schema") or {"order": {}, "items": []})

        doc_name = doc_struct.get("doc_name") or Path(f).name
        all_blocks[doc_name] = {b["id"]: b for b in doc_struct.get("blocks", []) if isinstance(b, dict) and b.get("id")}

        # Build people containers and append records
        for rec in doc_struct.get("records", []):
            pib = (rec.get("pib") or "").strip()
            if not pib:
                continue
            if pib not in people:
                people[pib] = {"pib": pib, "variants": [], "ranks": [], "records": [], "mentions": []}
            if rec.get("rank") and rec["rank"] not in people[pib]["ranks"]:
                people[pib]["ranks"].append(rec["rank"])
            people[pib]["records"].append(rec)

        # Mentions (block-level)
        for pib, b_ids in (doc_struct.get("people_index") or {}).items():
            if pib not in people:
                people[pib] = {"pib": pib, "variants": [], "ranks": [], "records": [], "mentions": []}
            for b_id in b_ids:
                blk = all_blocks.get(doc_name, {}).get(b_id)
                if not blk:
                    continue
                people[pib]["mentions"].append(
                    {
                        "doc": doc_name,
                        "block_id": b_id,
                        "section_path": blk.get("section_path", []),
                        "event_type": blk.get("event_type") or "",
                        "category": blk.get("category"),
                        "rule_id": blk.get("rule_id"),
                        "cat_span": blk.get("cat_span"),
                        "date": blk.get("date") or "",
                        "text": blk.get("text") or "",
                    }
                )

    people_list = sorted(people.values(), key=lambda p: (p.get("pib") or "").upper())

    return {
        "schema": "combined_v4",
        "generated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "documents": meta_docs,
        "processed_files": processed_files,
        "errors": errors,
        "people": people_list,
        "structured_documents": schema_documents,
    }
