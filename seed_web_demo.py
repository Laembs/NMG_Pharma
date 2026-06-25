"""Demo-Daten für den Web-Pilot anlegen.

Erzeugt zwei Firmen, um Mandanten-Trennung UND Lizenz-Gate zu beweisen:

  • Firma A „Muster-Pharma GmbH"  – Modul *personal* LIZENZIERT, mit Demo-Mitarbeitern.
  • Firma B „Beta-Distribution KG" – Modul *personal* NICHT lizenziert (nur faktura,
    das im Pilot noch nicht umgesetzt ist) → Dashboard zeigt keine Personal-Kachel.

Logins:
  Firma A:  admin / demo123
  Firma B:  admin / demo123

Aufruf:  python seed_web_demo.py
"""
from __future__ import annotations

from datetime import date, timedelta

from web.auth import create_user
from web.tenancy import create_firma, init_platform_db, tenant_con
from web.services import personal_service as svc
from app.personal_app import DEMO_EMPS  # (id, vorname, name, abteilung, position, x, y)


def _seed_mitarbeiter(slug: str) -> None:
    con = tenant_con(slug)
    try:
        if con.execute("SELECT COUNT(*) FROM tbl_mitarbeiter").fetchone()[0]:
            return  # schon befüllt
        for _id, vorname, name, abteilung, position, _x, _y in DEMO_EMPS:
            con.execute(
                """INSERT INTO tbl_mitarbeiter(vorname, name, abteilung, position)
                   VALUES (?,?,?,?)""", (vorname, name, abteilung, position))
        con.commit()
    finally:
        con.close()


def _seed_workflow(slug: str) -> None:
    """Beispiele für Abwesenheits-Workflow + Zeiterfassung (nur einmal)."""
    con = tenant_con(slug)
    try:
        if con.execute("SELECT COUNT(*) FROM tbl_abwesenheit").fetchone()[0]:
            return
        ids = [r["id"] for r in con.execute(
            "SELECT id FROM tbl_mitarbeiter ORDER BY id LIMIT 3").fetchall()]
        if not ids:
            return
        jahr = date.today().year
        # 1 genehmigter Urlaub, 1 offener Antrag, 1 Krankheit (direkt)
        svc.create_abwesenheit(con, mitarbeiter_id=ids[0], art="Urlaub",
                               von=f"{jahr}-07-01", bis=f"{jahr}-07-12",
                               notiz="Sommerurlaub", status="genehmigt")
        svc.create_abwesenheit(con, mitarbeiter_id=ids[min(1, len(ids)-1)], art="Urlaub",
                               von=f"{jahr}-08-05", bis=f"{jahr}-08-09",
                               notiz="Brückentage", status="beantragt")
        svc.create_abwesenheit(con, mitarbeiter_id=ids[min(2, len(ids)-1)], art="Krankheit",
                               von=f"{jahr}-06-03", bis=f"{jahr}-06-04",
                               notiz="", status="genehmigt")
        # Zeiterfassung: zwei Tage für den ersten Mitarbeiter
        heute = date.today()
        for delta, (k, g) in enumerate([("08:00", "16:30"), ("08:15", "17:00")]):
            tag = (heute - timedelta(days=delta + 1)).isoformat()
            svc.create_zeit(con, mitarbeiter_id=ids[0], datum=tag,
                            kommt=k, geht=g, pause_minuten=30)
    finally:
        con.close()


def main() -> None:
    init_platform_db()

    # Firma A: Personal freigeschaltet
    a_id, a_slug = create_firma("Muster-Pharma GmbH", modules=["personal"])
    create_user(a_id, "admin", "demo123", anzeigename="Admin A", ist_admin=True)
    _seed_mitarbeiter(a_slug)
    _seed_workflow(a_slug)

    # Firma B: nur faktura (im Pilot nicht umgesetzt) → kein Personal
    b_id, b_slug = create_firma("Beta-Distribution KG", modules=["faktura"])
    create_user(b_id, "admin", "demo123", anzeigename="Admin B", ist_admin=True)

    print("Demo-Daten angelegt:")
    print(f"  - {a_slug:24s}  Login admin/demo123  | Personal LIZENZIERT (+Demo-Mitarbeiter)")
    print(f"  - {b_slug:24s}  Login admin/demo123  | Personal NICHT lizenziert")
    print("\nStart:  uvicorn web.app:app --reload   ->  http://localhost:8000")


if __name__ == "__main__":
    main()
