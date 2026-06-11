# Kompatibilitätsmodul für Versionen vor 0.9.
# Die eigentliche Lernlogik liegt ab 0.9 in learning_db.py.
from .learning_db import clean_hersteller, import_learning_list, lookup_hersteller


def import_hersteller_liste(path):
    result = import_learning_list(path)
    return {
        "imported": result.get("hersteller", 0),
        "updated": 0,
        "skipped": result.get("skipped", 0),
        "ek": result.get("ek", 0),
    }
