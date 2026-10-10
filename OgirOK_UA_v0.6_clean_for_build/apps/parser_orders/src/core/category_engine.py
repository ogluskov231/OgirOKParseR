import json
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple


def _same_length_normalize(text: str) -> str:
    """Normalize text while preserving length (so regex spans map back).
    Only 1:1 character substitutions are allowed here.
    """
    if not text:
        return ""
    # NBSP -> space, smart apostrophe -> apostrophe
    # keep length identical (single char -> single char)
    return (
        text.replace("\u00A0", " ")
            .replace("\u2019", "'")
            .replace("\u2018", "'")
    )


def make_regex_from_phrase(phrase: str) -> str:
    """Convert a user-selected phrase into a robust regex.
    - escape metacharacters
    - spaces become \\s+ to tolerate variable whitespace
    - apostrophes tolerate both ' and ’
    """
    phrase = (phrase or "").strip()
    # trim wrapping quotes
    phrase = phrase.strip('"\u201c\u201d“”')
    phrase = phrase.strip()
    if not phrase:
        return ""
    # normalize apostrophes in the phrase for regex
    phrase = phrase.replace("\u2019", "'").replace("\u2018", "'")
    phrase = re.sub(r"\s+", " ", phrase)
    esc = re.escape(phrase)
    # make spaces flexible: re.escape turns space into '\ ', so replace those sequences
    esc = esc.replace(r"\ ", r"\s+")
    # apostrophe: accept both ' and ’ in input
    esc = esc.replace(r"\'", r"[\'\u2019]")
    return esc


@dataclass(frozen=True)
class Rule:
    id: str
    label: str
    patterns: List[re.Pattern]
    fallback: bool = False


class CategoryEngine:
    """Rule-based classifier for 'Надійшло/Вибуло' taxonomy."""

    def __init__(self, rules_path: str, user_rules_path: Optional[str] = None):
        self.rules_path = str(rules_path)
        self.user_rules_path = str(user_rules_path) if user_rules_path else None

        self._arrival_rules: List[Rule] = []
        self._departure_rules: List[Rule] = []
        self._dep_triggers: List[re.Pattern] = []
        self._arr_triggers: List[re.Pattern] = []
        self._policy: str = "departure_wins_if_both"

        self.load()

    @property
    def arrival_labels(self) -> List[str]:
        return [r.label for r in self._arrival_rules]

    @property
    def departure_labels(self) -> List[str]:
        return [r.label for r in self._departure_rules]

    def load(self) -> None:
        rules = self._load_json(self.rules_path)
        # optional user_rules (future extension)
        user_rules = None
        if self.user_rules_path and Path(self.user_rules_path).exists():
            user_rules = self._load_json(self.user_rules_path)

        etd = rules.get("event_type_detection", {}) or {}
        self._policy = (etd.get("policy") or "departure_wins_if_both").strip()
        self._dep_triggers = [re.compile(p, re.IGNORECASE) for p in (etd.get("departure_triggers") or [])]
        self._arr_triggers = [re.compile(p, re.IGNORECASE) for p in (etd.get("arrival_triggers") or [])]

        self._arrival_rules = self._compile_rules(rules.get("arrival_rules") or [])
        self._departure_rules = self._compile_rules(rules.get("departure_rules") or [])

        # User rules: allow appending new rules without editing base file.
        # Supports:
        #   - enabled: false (skip)
        #   - priority: high|low (high is prepended)
        if user_rules and isinstance(user_rules.get("rules"), list):
            extra = user_rules["rules"]
            for r in extra:
                try:
                    if r.get("enabled", True) is False:
                        continue
                    event = (r.get("event_type") or "").strip()
                    compiled_list = self._compile_rules([r])
                    if not compiled_list:
                        continue
                    compiled = compiled_list[0]
                    if event == "arrival":
                        self._arrival_rules.insert(0, compiled) if r.get("priority") == "high" else self._arrival_rules.append(compiled)
                    elif event == "departure":
                        self._departure_rules.insert(0, compiled) if r.get("priority") == "high" else self._departure_rules.append(compiled)
                except Exception:
                    continue

    def detect_event_type(self, text: str) -> str:
        t = _same_length_normalize(text)
        dep = any(rx.search(t) for rx in self._dep_triggers)
        arr = any(rx.search(t) for rx in self._arr_triggers)
        if dep and arr and self._policy == "departure_wins_if_both":
            return "departure"
        if dep:
            return "departure"
        if arr:
            return "arrival"
        return "unknown"

    def classify(self, raw_text: str, event_hint: Optional[str] = None) -> Dict[str, Any]:
        raw_text = raw_text or ""
        t = _same_length_normalize(raw_text)

        event = (event_hint or "").strip()
        if event not in ("arrival", "departure"):
            event = self.detect_event_type(t)

        rules = self._departure_rules if event == "departure" else self._arrival_rules if event == "arrival" else []
        for rule in rules:
            for rx in rule.patterns:
                m = rx.search(t)
                if not m:
                    continue
                s, e = m.span()
                return {
                    "event_type": event,
                    "category": rule.label,
                    "rule_id": rule.id,
                    "match_span": (s, e),
                    "match_text": t[s:e],
                    "fallback": rule.fallback,
                }

        return {"event_type": event, "category": None, "rule_id": None, "match_span": None, "match_text": None}

    # ---------------- internal
    def _compile_rules(self, items: List[Dict[str, Any]]) -> List[Rule]:
        out: List[Rule] = []
        for r in items:
            rid = (r.get("id") or "").strip() or "rule"
            label = (r.get("label") or "").strip() or rid
            fallback = bool(r.get("fallback"))
            match = r.get("match") or {}
            any_list = match.get("any") or []
            patterns = []
            for p in any_list:
                try:
                    patterns.append(re.compile(p, re.IGNORECASE))
                except re.error:
                    # ignore broken patterns instead of crashing UI
                    continue
            if patterns:
                out.append(Rule(id=rid, label=label, patterns=patterns, fallback=fallback))
        return out

    @staticmethod
    def _load_json(path: str) -> Dict[str, Any]:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
