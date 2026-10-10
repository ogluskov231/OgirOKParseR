"""CLI runner for apps/parser_orders.

This provides a "ready program" path that doesn't require opening the Tkinter GUI.

Examples (Windows PowerShell):
  py run_parser_orders_cli.py --input "input\documents" --excel
  py run_parser_orders_cli.py --input "C:\\path\\to\\order.docx" --excel --json

Examples (Linux/WSL):
  python3 run_parser_orders_cli.py --input ./input/documents --excel
"""

from __future__ import annotations

import argparse
import json
import os
from pathlib import Path
from typing import List

from apps.parser_orders.src.core.docx_parser import build_combined_v2
from apps.parser_orders.src.core.excel_autofill import autofill_template


PROJECT_ROOT = Path(__file__).resolve().parent


def collect_docx(p: str) -> List[str]:
    path = Path(p)
    if path.is_file() and path.suffix.lower() == ".docx":
        return [str(path)]
    if path.is_dir():
        out: List[str] = []
        for f in path.rglob("*.docx"):
            # Always ignore MS Office temp files
            if f.name.startswith("~$"):
                continue
            out.append(str(f))
        return sorted(out)
    raise FileNotFoundError(f"Input path not found: {p}")


def main() -> int:
    ap = argparse.ArgumentParser(description="Parse DOCX orders -> combined JSON + (optional) Excel autofill")
    ap.add_argument("--input", required=True, help="DOCX file or a directory with .docx")
    ap.add_argument(
        "--template",
        default=str(PROJECT_ROOT / "templates" / "N_fixed.xlsx"),
        help="Excel template path (default: templates/N_fixed.xlsx)",
    )
    ap.add_argument(
        "--out",
        default=str(PROJECT_ROOT / "output"),
        help="Output directory (default: ./output)",
    )
    ap.add_argument("--excel", action="store_true", help="Generate filled Excel")
    ap.add_argument("--json", action="store_true", help="Write combined JSON")
    ap.add_argument("--txt", action="store_true", help="Write TXT summary")
    args = ap.parse_args()

    docx_files = collect_docx(args.input)
    if not docx_files:
        print("No DOCX found.")
        return 2

    template_path = str(Path(args.template))
    if not os.path.exists(template_path):
        raise FileNotFoundError(f"Template not found: {template_path}")

    out_root = Path(args.out)
    exports_json = out_root / "exports" / "json"
    exports_txt = out_root / "exports" / "txt"
    exports_excel = out_root / "exports" / "excel"
    diagnostics_dir = out_root / "diagnostics"
    exports_json.mkdir(parents=True, exist_ok=True)
    exports_txt.mkdir(parents=True, exist_ok=True)
    exports_excel.mkdir(parents=True, exist_ok=True)
    diagnostics_dir.mkdir(parents=True, exist_ok=True)

    mapping_path = str(PROJECT_ROOT / "config" / "autofill" / "mapping.json")
    template_sources_path = str(PROJECT_ROOT / "config" / "autofill" / "template_sources.json")
    from apps.parser_orders.src.core.excel_autofill import ensure_mapping_file
    ensure_mapping_file(mapping_path, template_sources_path)
    rules_path = str(PROJECT_ROOT / "config" / "autofill" / "passport_rules.json")
    learning_dir = str(PROJECT_ROOT / "config" / "learning")

    combined = build_combined_v2(docx_files, template_path=template_path, learning_dir=learning_dir)

    # flatten records like GUI does
    records = []
    for person in combined.get("people", []):
        pib = person.get("pib", "")
        for r in person.get("records", []):
            if not r.get("pib"):
                r["pib"] = pib
            records.append(r)

    ts = __import__("datetime").datetime.now().strftime("%Y%m%d_%H%M%S")
    out_json = exports_json / f"combined_{ts}.json"
    out_txt = exports_txt / f"summary_{ts}.txt"

    if args.json:
        out_json.write_text(json.dumps(combined, ensure_ascii=False, indent=2), encoding="utf-8")
        print(f"JSON: {out_json}")

    if args.txt:
        # basic summary
        lines = [f"DOCX: {len(docx_files)}", f"PEOPLE: {len(combined.get('people', []))}"]
        for d in combined.get("documents", []):
            lines.append(f"- {d.get('source_docx')}: order {d.get('order_ref')} records {d.get('records')}")
        out_txt.write_text("\n".join(lines), encoding="utf-8")
        print(f"TXT: {out_txt}")

    if args.excel:
        diag_path = str(diagnostics_dir / f"unknowns_{ts}.json")
        excel_path = autofill_template(
            records,
            template_path=template_path,
            output_dir=str(exports_excel),
            rules_path=rules_path,
            mapping_path=mapping_path,
            diagnostics_path=diag_path,
            learning_dir=learning_dir,
        )
        print(f"EXCEL: {excel_path}")
        print(f"DIAG:  {diag_path}")

    if not (args.excel or args.json or args.txt):
        print("Nothing to do. Specify at least one flag: --excel and/or --json and/or --txt")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
