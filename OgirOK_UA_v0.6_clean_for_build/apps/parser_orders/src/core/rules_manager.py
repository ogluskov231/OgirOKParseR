
import json
import os
from pathlib import Path


def _project_root() -> Path:
    # .../apps/parser_orders/src/core -> project root
    return Path(__file__).resolve().parents[4]


LEARNING_FILES = {
    "units": str(_project_root() / "config" / "learning" / "units.json"),
    "reasons": str(_project_root() / "config" / "learning" / "reasons.json"),
    "arrival_sources": str(_project_root() / "config" / "learning" / "arrival_sources.json"),
    "events": str(_project_root() / "config" / "learning" / "events.json"),
}

def _load(path):
    if os.path.exists(path):
        return json.load(open(path, "r", encoding="utf-8"))
    return {}

def add_rule(category, match_text, maps_to):
    path = LEARNING_FILES[category]
    data = _load(path)
    data[match_text] = maps_to
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    return path

def add_event_keyword(kind, keyword):
    path = LEARNING_FILES["events"]
    data = _load(path)
    arr = data.get(kind, [])
    if keyword not in arr:
        arr.append(keyword)
    data[kind] = arr
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    return path
