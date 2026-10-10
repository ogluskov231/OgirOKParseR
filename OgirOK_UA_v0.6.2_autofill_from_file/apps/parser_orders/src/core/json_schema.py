"""Validation and normalization for the project JSON parser schema."""
from __future__ import annotations
import json
from pathlib import Path
from typing import Any, Dict, Tuple

SCHEMA_PATH = Path(__file__).resolve().parents[4] / "config" / "schema" / "military_order.schema.json"

try:
    from jsonschema import Draft7Validator
except Exception:  # pragma: no cover
    Draft7Validator = None


def load_schema() -> Dict[str, Any]:
    return json.loads(SCHEMA_PATH.read_text(encoding="utf-8"))


def validate_order(data: Dict[str, Any]) -> Tuple[bool, list[str]]:
    if Draft7Validator is None:
        return False, ["Пакет jsonschema не встановлено"]
    validator = Draft7Validator(load_schema())
    errors = []
    for err in sorted(validator.iter_errors(data), key=lambda e: list(e.path)):
        path = ".".join(str(x) for x in err.path) or "$"
        errors.append(f"{path}: {err.message}")
    return not errors, errors
