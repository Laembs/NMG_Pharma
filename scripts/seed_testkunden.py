"""Legt Test-Kunden (+ ein paar Beispiel-PK-Konditionen) in der Datenbank an,
zum Ausprobieren der Kasse-Kundensuche und der Rabatt-Kaskade.

Reine Testdaten - die DB ist nicht versioniert (data/ ist gitignored).
Mehrfach ausfuehrbar (Upsert ueber kundennummer):
    python scripts/seed_testkunden.py
"""
import sqlite3
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from app.config import DB_PATH
from app.kasse_import import _ensure_kunden_center

KUNDEN = [
    # kundennummer, name, plz, ort, strasse, inhaber, email
    ("10001", "Adler-Apotheke", "10115", "Berlin", "Hauptstraße 5", "Dr. Anna Adler", "anna.adler@adler-apo.de"),
    ("10002", "Bären-Apotheke", "80331", "München", "Marienplatz 2", "Bernd Bär", "info@baeren-apo.de"),
    ("10003", "Central-Apotheke", "50667", "Köln", "Domkloster 4", "Clara Cremer", "kontakt@central-apo.de"),
    ("10004", "Dom-Apotheke", "20095", "Hamburg", "Mönckebergstr. 17", "David Dorn", "david.dorn@dom-apo.de"),
    ("10005", "Engel-Apotheke", "70173", "Stuttgart", "Königstraße 28", "Eva Engel", "service@engel-apo.de"),
    ("10006", "Forsthaus-Apotheke", "01067", "Dresden", "Prager Str. 10", "Frank Forst", "frank@forsthaus-apo.de"),
]


def main():
    con = sqlite3.connect(DB_PATH)
    _ensure_kunden_center(con)
    neu = aktualisiert = 0
    for knr, name, plz, ort, strasse, inhaber, email in KUNDEN:
        ex = con.execute("SELECT id FROM tbl_kunden_center WHERE kundennummer=?", (knr,)).fetchone()
        if ex:
            con.execute(
                "UPDATE tbl_kunden_center SET kundenname=?, plz=?, ort=?, strasse=?, inhaber=?, "
                "email=?, status='aktiv' WHERE id=?",
                (name, plz, ort, strasse, inhaber, email, ex[0]))
            aktualisiert += 1
        else:
            con.execute(
                "INSERT INTO tbl_kunden_center(kundennummer, kundenname, plz, ort, strasse, "
                "inhaber, email, status, bearbeiter) VALUES(?,?,?,?,?,?,?,?,?)",
                (knr, name, plz, ort, strasse, inhaber, email, "aktiv", "seed"))
            neu += 1

    # Beispiel-PK-Konditionen (feste Kundenrabatte) auf echte NMG-PZNs
    con.execute(
        """CREATE TABLE IF NOT EXISTS tbl_pk_konditionen(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            kundennummer TEXT, kundenname TEXT, pzn TEXT, rabatt_prozent REAL,
            gueltigkeit TEXT, quelle TEXT, importdatum TEXT,
            letzte_aktualisierung TEXT DEFAULT CURRENT_TIMESTAMP,
            UNIQUE(kundennummer, pzn))""")
    pzns = [r[0] for r in con.execute("SELECT pzn FROM tbl_nmg_stamm LIMIT 2")]
    pk_neu = 0
    for pzn, rab in zip(pzns, (25.0, 30.0)):
        con.execute(
            "INSERT OR REPLACE INTO tbl_pk_konditionen(kundennummer, kundenname, pzn, rabatt_prozent, quelle) "
            "VALUES('10001','Adler-Apotheke',?,?,'Testdaten')", (pzn, rab))
        pk_neu += 1

    con.commit()
    con.close()
    print(f"Test-Kunden: {neu} neu, {aktualisiert} aktualisiert.")
    print(f"PK-Konditionen (Adler-Apotheke, fester Rabatt): {pk_neu} auf PZN {pzns}.")
    print(f"DB: {DB_PATH}")


if __name__ == "__main__":
    main()
