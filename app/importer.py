from pathlib import Path
from .db import import_nmg_artikelliste


def import_excel(path: str | Path, typ: str = "nmg_stamm") -> dict:
    """Importiert eine NMG-Artikelliste/APU-HAP-Liste in tbl_nmg_stamm."""
    stats = import_nmg_artikelliste(path)
    return {"imported": stats.get("imported", 0), "skipped": stats.get("skipped", 0)}
