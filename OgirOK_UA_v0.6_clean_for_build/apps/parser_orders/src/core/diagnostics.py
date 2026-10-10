
import json, os

def load_unknowns(path):
    if not os.path.exists(path):
        return {"unknown_units":[], "unknown_departure_reasons":[], "unknown_arrival_sources":[], "skipped":[]}
    return json.load(open(path, "r", encoding="utf-8"))
