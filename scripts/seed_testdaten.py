"""Legt einen vollstaendigen Testdatensatz fuer die Kasse-App an, damit man alle
Ansichten sauber pruefen kann: Kunden, Lagerbestand (mehrere Chargen), Verkaeufe
(auch aeltere fuer Top-10), Vorbestellungen, stornierte Verkaeufe und abgesagte
Vorbestellungen.

ACHTUNG: loescht vorher die Transaktionsdaten (Bestellungen/Positionen/Lager/
Wareneingang) - Kunden + PK-Konditionen bleiben. Reine Testdaten, DB ist nicht
versioniert. Mehrfach ausfuehrbar (immer derselbe saubere Stand):
    python scripts/seed_testdaten.py
"""
import sqlite3
import sys
from datetime import date, datetime, timedelta
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(Path(__file__).resolve().parent))  # fuer seed_testkunden

from app.config import DB_PATH
from app.migrations import run_migrations
import seed_testkunden


def _d(tage):
    return (date.today() - timedelta(days=tage)).isoformat()


def main():
    run_migrations()
    seed_testkunden.main()  # Kunden + PK-Konditionen sicherstellen

    con = sqlite3.connect(DB_PATH)
    # Transaktionsdaten leeren (Kunden/PK bleiben).
    for t in ("tbl_bestellpositionen", "tbl_bestellungen", "tbl_wareneingang_positionen",
              "tbl_wareneingang", "tbl_lagerbestand"):
        con.execute(f"DELETE FROM {t}")

    art = con.execute("SELECT pzn, artikelname, apu FROM tbl_nmg_stamm ORDER BY pzn LIMIT 5").fetchall()
    if len(art) < 5:
        print("Zu wenige NMG-Artikel in der DB."); return
    a0, a1, a2, a3, a4 = art
    jetzt = datetime.now().isoformat(timespec="seconds")

    def apotheke(knr):
        r = con.execute("SELECT kundenname FROM tbl_kunden_center WHERE kundennummer=?", (knr,)).fetchone()
        return r[0] if r else knr

    def lager(a, charge, verfall, menge):
        con.execute("INSERT INTO tbl_lagerbestand(pzn,artikelname,charge,verfall,menge,aktualisiert_am) "
                    "VALUES(?,?,?,?,?,?)", (a[0], a[1], charge, verfall, menge, jetzt))

    def bestellung(datum, knr, status="offen"):
        cur = con.execute("INSERT INTO tbl_bestellungen(datum,kundennummer,apotheke,status,bestellart,bearbeiter) "
                          "VALUES(?,?,?,?,?,?)", (datum, knr, apotheke(knr), status, "Bestellung", "seed"))
        return cur.lastrowid

    def pos(bid, a, menge, bestellart="Bestellung", charge="", verfall="", rabatt=0,
            lz="10 Uhr", termin=""):
        con.execute("INSERT INTO tbl_bestellpositionen(bestell_id,pzn,artikelname,apu,menge,"
                    "rabatt_prozent,rabatt_quelle,charge,verfall,bestellart,lieferzeit,liefertermin) "
                    "VALUES(?,?,?,?,?,?,?,?,?,?,?,?)",
                    (bid, a[0], a[1], a[2], menge, rabatt, "Testdaten", charge, verfall,
                     bestellart, lz, termin))

    # --- Lagerbestand (a3 bewusst OHNE Bestand -> Vorbestellung testbar) ---
    lager(a0, "CH-A", "10/2027", 20)
    lager(a0, "CH-B", "03/2029", 15)
    lager(a1, "L-2026", "12/2026", 30)
    lager(a2, "L-2028", "06/2028", 8)
    lager(a4, "L-2030", "01/2030", 5)

    # --- Verkaeufe (abgeschlossen) ---
    b = bestellung(_d(0), "10001"); pos(b, a0, 5, charge="CH-A", verfall="10/2027"); pos(b, a1, 2, charge="L-2026", verfall="12/2026")
    b = bestellung(_d(5), "10001"); pos(b, a0, 3, charge="CH-B", verfall="03/2029")
    b = bestellung(_d(2), "10002"); pos(b, a2, 4, charge="L-2028", verfall="06/2028")
    b = bestellung(_d(180), "10003"); pos(b, a1, 10, charge="L-2026", verfall="12/2026")  # aelter, aber < 12 Mon
    b = bestellung(_d(400), "10001"); pos(b, a4, 9, charge="L-2030", verfall="01/2030")   # > 12 Mon (Top-10-Grenze)

    # --- Stornierter Verkauf ---
    b = bestellung(_d(20), "10001", status="storniert"); pos(b, a0, 7, charge="CH-A", verfall="10/2027")

    # --- Offene Vorbestellungen ---
    b = bestellung(_d(1), "10002"); pos(b, a3, 6, bestellart="Vorbestellung", termin=_d(-2))
    b = bestellung(_d(1), "10004"); pos(b, a0, 4, bestellart="Vorbestellung", termin=_d(-3))

    # --- Abgesagte Vorbestellung ---
    b = bestellung(_d(8), "10001"); pos(b, a2, 3, bestellart="abgesagt")

    # Ein paar Verkaeufe als "in MSK erfasst" markieren (Rest bleibt MSK offen).
    ids = [r[0] for r in con.execute(
        "SELECT b.id FROM tbl_bestellungen b JOIN tbl_bestellpositionen p ON p.bestell_id=b.id "
        "WHERE p.bestellart='Bestellung' AND COALESCE(b.status,'offen')<>'storniert' "
        "GROUP BY b.id ORDER BY b.id LIMIT 2")]
    for i in ids:
        con.execute("UPDATE tbl_bestellungen SET msk_erfasst=1, msk_von='seed', msk_am=? WHERE id=?",
                    (datetime.now().isoformat(timespec="seconds"), i))

    # Ein paar Protokoll-Beispiele (sonst fuellt sich das Log erst im Betrieb).
    con.execute("DELETE FROM tbl_kasse_log")
    vk = con.execute("SELECT b.id, COALESCE(b.apotheke,'') FROM tbl_bestellungen b "
                     "JOIN tbl_bestellpositionen p ON p.bestell_id=b.id "
                     "WHERE p.bestellart='Bestellung' GROUP BY b.id ORDER BY b.id LIMIT 3").fetchall()
    beispiele = [("anna", "Verkauf gespeichert"), ("bernd", "MSK erfasst"), ("anna", "Verkauf storniert")]
    for (bid, apo), (wer, aktion) in zip(vk, beispiele):
        con.execute("INSERT INTO tbl_kasse_log(zeitpunkt,bearbeiter,aktion,bestell_id,kunde,details) "
                    "VALUES(?,?,?,?,?,?)", (datetime.now().isoformat(timespec="seconds"),
                    wer, aktion, bid, apo, "Testdaten"))

    con.commit()
    # Zusammenfassung
    nb = con.execute("SELECT COUNT(*) FROM tbl_bestellungen").fetchone()[0]
    nstorno = con.execute("SELECT COUNT(*) FROM tbl_bestellungen WHERE status='storniert'").fetchone()[0]
    nvb = con.execute("SELECT COUNT(*) FROM tbl_bestellpositionen WHERE bestellart='Vorbestellung'").fetchone()[0]
    nab = con.execute("SELECT COUNT(*) FROM tbl_bestellpositionen WHERE bestellart='abgesagt'").fetchone()[0]
    nl = con.execute("SELECT COUNT(*) FROM tbl_lagerbestand").fetchone()[0]
    con.close()
    print("Testdaten angelegt:")
    print(f"  Bestellungen: {nb} (davon {nstorno} storniert)")
    print(f"  Offene Vorbestellungen: {nvb} · Abgesagte: {nab}")
    print(f"  Lagerbestand-Zeilen: {nl} (a0 hat 2 Chargen, a3 ohne Bestand)")
    print(f"  DB: {DB_PATH}")
    print("Pruefen: Verkaeufe (Storniert-Filter), Vorbestellungen (Abgesagt-Filter),")
    print("  Top-Artikel fuer Adler-Apotheke (10001), Charge-Auswahl bei a0.")


if __name__ == "__main__":
    main()
