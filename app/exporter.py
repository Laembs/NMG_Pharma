from pathlib import Path
from datetime import datetime
import sqlite3
import re
import csv
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter
from .config import OUTPUT_DIR, DB_PATH
from .db import init_db
from .file_loader import load_worksheet
from .learning_db import clean_hersteller, find_columns, register_hersteller, lookup_hersteller, parse_number, register_basisdaten, lookup_basisdaten, register_ek, lookup_latest_ek
from .abgleich_protocol import trace_lookup_source, write_abgleichartikel_protocol

LINDEN_HEADERS = [
    "PZN", "Artikelname", "DF", "Pck", "Herst", "EK",
    "Abverkäufe 6 Monate", "im Sortiment", "PZN NMG", "APU NMG",
    "NMG Rabatt", "lieferbar", "Bevorratung angeraten", "Liefervor- schlag",
    "austauschbar gegen", "NMG Rabatt in Euro", "NMG Rabatt Gesamt nach Absatz", "Umsatz"
]


class UnknownInputFormatError(Exception):
    """Wird ausgelöst, wenn eine Rohdatei nicht sicher erkannt wurde."""
    def __init__(self, message: str, report_path: Path | None = None):
        super().__init__(message)
        self.report_path = report_path


def _col_letter(idx):
    try:
        return get_column_letter(idx) if idx else ""
    except Exception:
        return ""


def _diagnose_format(input_file: Path, ws, mapping):
    """Erzeugt einen Diagnosebericht, wenn die Datei nicht sicher auswertbar ist."""
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    safe = ''.join(ch if ch.isalnum() or ch in '-_' else '_' for ch in input_file.stem)[:60]
    report = OUTPUT_DIR / f"Diagnose_Format_nicht_erkannt_{safe}_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
    wb = Workbook()
    sh = wb.active
    sh.title = "Diagnose"
    sh.append(["Datei", input_file.name])
    sh.append(["Status", "Format nicht sicher erkannt – Datei wurde nicht ausgewertet."])
    sh.append([])

    mapping = mapping or {}
    checks = [
        ("PZN", mapping.get("pzn"), True),
        ("Artikelname", mapping.get("artikel"), False),
        ("Hersteller/Herst", mapping.get("hersteller"), False),
        ("Absatz/Verbrauch", mapping.get("absatz"), True),
        ("EK", mapping.get("ek"), False),
    ]
    sh.append(["Pflicht-/Zielfeld", "Gefunden", "Spalte", "Hinweis"])
    for name, col, required in checks:
        sh.append([name, "Ja" if col else "Nein", _col_letter(col), "Pflichtfeld" if required else "optional"])
    sh.append([])
    sh.append(["Gefundene Überschriften / erste Zeilen"])

    # Zeilen 1-15 als Diagnose ausgeben.
    max_col = min(ws.max_column, 40)
    for r in range(1, min(ws.max_row, 15) + 1):
        values = [ws.cell(r, c).value for c in range(1, max_col + 1)]
        sh.append([f"Zeile {r}"] + values)

    sh2 = wb.create_sheet("Manuelle Zuordnung")
    sh2.append(["Feld", "Trage hier Spaltenbuchstabe ein", "Beispiel", "Pflicht"])
    rows = [
        ("PZN", "", "A", "Ja"),
        ("Artikelname", "", "B oder N", "Nein, aber empfohlen"),
        ("Hersteller", "", "E oder O", "Nein"),
        ("EK", "", "F / Apo-EK / Taxe-EK", "Nein"),
        ("Absatz 6 Monate", "", "G oder P", "Ja"),
        ("DF", "", "C", "Nein"),
        ("Pck", "", "D", "Nein"),
    ]
    for row in rows:
        sh2.append(row)
    sh2.append([])
    sh2.append(["Hinweis", "Diese Version erzeugt die Vorlage. Die automatische Nutzung der manuellen Zuordnung kommt im nächsten Schritt."])

    for sheet in wb.worksheets:
        for col in range(1, min(sheet.max_column, 12) + 1):
            sheet.column_dimensions[get_column_letter(col)].width = 22
    wb.save(report)
    return report


def _is_mapping_usable(mapping):
    if not mapping:
        return False
    # Für eine sichere Auswertung müssen PZN und Absatz vorhanden sein.
    # EK ist optional, Hersteller und Artikel können aus DB/leer ergänzt werden.
    return bool(mapping.get("pzn") and mapping.get("absatz"))


def _pzn(value):
    if value is None:
        return ""
    text = str(value).strip()
    if text.endswith('.0'):
        text = text[:-2]
    return text.zfill(8) if text.isdigit() else text


def _to_number(value, default=0):
    if value in (None, ""):
        return default
    if isinstance(value, (int, float)):
        return value
    try:
        return float(str(value).strip().replace(".", "").replace(",", "."))
    except Exception:
        return default


def _rowdict(row):
    return dict(row) if row else None


def _lookup(con: sqlite3.Connection, original_pzn: str):
    """Liest die gelernte Zuordnung für die Neue Auswertung.

    Wichtig für die Schulbank-/Editor-Logik:
    Die zentrale Wissensquelle ist tbl_austauschdatenbank. Sie muss vor alten
    Referenz-/Importtabellen gewinnen, sonst erscheinen bereits gelernte Fälle
    nach einer neuen Auswertung wieder als Abweichung.
    """
    original_pzn = _pzn(original_pzn)
    if not original_pzn:
        return {}

    def _norm_db_pzn(value):
        return _pzn(value)

    def _article_name_by_pzn(pzn):
        pzn = _pzn(pzn)
        if not pzn:
            return ""
        for table, col in (
            ("tbl_nmg_stamm", "artikelname"),
            ("tbl_artikelstamm", "artikel"),
            ("tbl_artikelstamm", "artikelname"),
            ("tbl_pzn_basisdaten", "artikelname"),
        ):
            try:
                row = con.execute(f"SELECT {col} FROM {table} WHERE pzn=? LIMIT 1", (pzn,)).fetchone()
                if row and row[0]:
                    return str(row[0]).strip()
            except Exception:
                continue
        return ""

    # 1) Höchste Priorität: neue Austauschdatenbank / Schulbank-Lernstand.
    # PZN robust vergleichen, weil Excel und alte Importe führende Nullen verlieren können.
    # Wichtig: Wenn zu einer PZN alt mehrere aktive Zielvarianten existieren, darf die
    # Neue Auswertung NICHT automatisch irgendeinen Treffer nehmen. Dieser Fall gehört
    # in die Schulbank -> Manuelle Prüfung und bleibt in der Auswertung bewusst offen.
    austausch = None
    try:
        rows = con.execute("""
            SELECT *
            FROM tbl_austauschdatenbank
            WHERE COALESCE(status, 'aktiv') = 'aktiv'
              AND COALESCE(pzn_alt, '') <> ''
            ORDER BY datetime(COALESCE(aktualisiert_am, erstellt_am, gueltig_ab, '1970-01-01')) DESC, id DESC
        """).fetchall()
        matches = []
        for candidate in rows:
            data = _rowdict(candidate)
            if _norm_db_pzn(data.get("pzn_alt")) == original_pzn:
                matches.append(data)

        def _norm_text(value):
            return " ".join(str(value or "").strip().lower().split())

        variants = {
            (
                _pzn(item.get("pzn_nmg")),
                _norm_text(item.get("freitext_austausch")),
            )
            for item in matches
            if _pzn(item.get("pzn_nmg")) or _norm_text(item.get("freitext_austausch"))
        }
        if len(variants) > 1:
            # Mehrdeutiger Lernstand: bewusst keinen automatischen Treffer verwenden.
            return {}
        if matches:
            austausch = matches[0]
    except Exception:
        austausch = None

    if austausch:
        nmg_pzn = _pzn(austausch.get("pzn_nmg"))
        freitext = str(austausch.get("freitext_austausch") or "").strip()
        artikel_nmg = str(austausch.get("artikel_nmg") or "").strip()
        if nmg_pzn and (not freitext or freitext.lower().startswith("pzn nmg:")):
            freitext = artikel_nmg or _article_name_by_pzn(nmg_pzn) or freitext

        stamm = _rowdict(con.execute("SELECT * FROM tbl_nmg_stamm WHERE pzn=?", (nmg_pzn or '',)).fetchone()) if nmg_pzn else None
        rabatt = _rowdict(con.execute("SELECT * FROM nmg_rabatte WHERE nmg_pzn=?", (nmg_pzn or '',)).fetchone()) if nmg_pzn else None
        liefer = _rowdict(con.execute("SELECT * FROM tbl_lieferfaehigkeit WHERE nmg_pzn=?", (nmg_pzn or '',)).fetchone()) if nmg_pzn else None

        return {
            "im_sortiment": "X" if nmg_pzn else "X Austausch mögl",
            "nmg_pzn": nmg_pzn or None,
            "apu_nmg": stamm.get("apu") if stamm and stamm.get("apu") is not None else None,
            "rabatt": rabatt.get("rabatt") if rabatt else None,
            "lieferbar": liefer.get("lieferbar") if liefer else None,
            "bevorratung": liefer.get("bevorratung_angeraten") if liefer else None,
            "liefervorschlag": liefer.get("liefervorschlag") if liefer else None,
            "austauschbar_gegen": freitext or None,
        }

    # 2) Sicherheitsnetz: bereits übernommene Lernvorschläge, falls ältere Datenbanken
    # noch nicht vollständig in tbl_austauschdatenbank gespiegelt wurden.
    try:
        rows = con.execute("""
            SELECT *
            FROM tbl_lernvorschlaege
            WHERE status = 'uebernommen'
              AND COALESCE(pzn_alt, '') <> ''
            ORDER BY datetime(COALESCE(bearbeitet_am, erstellt_am, '1970-01-01')) DESC, id DESC
        """).fetchall()
        for candidate in rows:
            data = _rowdict(candidate)
            if _norm_db_pzn(data.get("pzn_alt")) != original_pzn:
                continue
            nmg_pzn = _pzn(data.get("pzn_nmg"))
            freitext = str(data.get("freitext_austausch") or data.get("produkt_neu") or "").strip()
            if nmg_pzn and (not freitext or freitext.lower().startswith("pzn nmg:")):
                freitext = _article_name_by_pzn(nmg_pzn) or freitext
            stamm = _rowdict(con.execute("SELECT * FROM tbl_nmg_stamm WHERE pzn=?", (nmg_pzn or '',)).fetchone()) if nmg_pzn else None
            rabatt = _rowdict(con.execute("SELECT * FROM nmg_rabatte WHERE nmg_pzn=?", (nmg_pzn or '',)).fetchone()) if nmg_pzn else None
            liefer = _rowdict(con.execute("SELECT * FROM tbl_lieferfaehigkeit WHERE nmg_pzn=?", (nmg_pzn or '',)).fetchone()) if nmg_pzn else None
            return {
                "im_sortiment": "X" if nmg_pzn else "X Austausch mögl",
                "nmg_pzn": nmg_pzn or None,
                "apu_nmg": stamm.get("apu") if stamm and stamm.get("apu") is not None else None,
                "rabatt": rabatt.get("rabatt") if rabatt else None,
                "lieferbar": liefer.get("lieferbar") if liefer else None,
                "bevorratung": liefer.get("bevorratung_angeraten") if liefer else None,
                "liefervorschlag": liefer.get("liefervorschlag") if liefer else None,
                "austauschbar_gegen": freitext or None,
            }
    except Exception:
        pass

    # 3) Danach erst alte geprüfte Referenz, damit neue Schulbank-Korrekturen nicht
    # durch veraltete Referenzwerte überschrieben werden.
    ref = _rowdict(con.execute(
        "SELECT * FROM tbl_referenz_h_o WHERE original_pzn=?", (original_pzn,)
    ).fetchone())
    if ref:
        nmg_pzn_ref = _pzn(ref.get("nmg_pzn"))
        stamm_ref = _rowdict(con.execute("SELECT * FROM tbl_nmg_stamm WHERE pzn=?", (nmg_pzn_ref or '',)).fetchone())
        return {
            "im_sortiment": ref.get("im_sortiment"),
            "nmg_pzn": nmg_pzn_ref,
            "apu_nmg": stamm_ref.get("apu") if stamm_ref and stamm_ref.get("apu") is not None else ref.get("apu_nmg"),
            "rabatt": ref.get("rabatt"),
            "lieferbar": ref.get("lieferbar"),
            "bevorratung": ref.get("bevorratung_angeraten"),
            "liefervorschlag": ref.get("liefervorschlag"),
            "austauschbar_gegen": ref.get("austauschbar_gegen"),
        }

    # 4) Falls die abgegebene PZN selbst NMG-Artikel ist oder Rabatt-/Lieferdaten hat.
    nmg_pzn = original_pzn
    stamm = _rowdict(con.execute("SELECT * FROM tbl_nmg_stamm WHERE pzn=?", (nmg_pzn or '',)).fetchone())
    rabatt = _rowdict(con.execute("SELECT * FROM nmg_rabatte WHERE nmg_pzn=?", (nmg_pzn or '',)).fetchone())
    liefer = _rowdict(con.execute("SELECT * FROM tbl_lieferfaehigkeit WHERE nmg_pzn=?", (nmg_pzn or '',)).fetchone())

    if not any([stamm, rabatt, liefer]):
        return {}

    return {
        "im_sortiment": "X",
        "nmg_pzn": nmg_pzn,
        "apu_nmg": stamm.get("apu") if stamm and stamm.get("apu") is not None else None,
        "rabatt": rabatt.get("rabatt") if rabatt else None,
        "lieferbar": liefer.get("lieferbar") if liefer else None,
        "bevorratung": liefer.get("bevorratung_angeraten") if liefer else None,
        "liefervorschlag": liefer.get("liefervorschlag") if liefer else None,
        "austauschbar_gegen": None,
    }

def _lookup_hersteller(con: sqlite3.Connection, pzn: str) -> str:
    return lookup_hersteller(con, pzn)




def _strip_leading_manufacturer_number(value):
    """Entfernt EKM-artige numerische Anbieter-Codes vor dem Herstellertext.
    Beispiel: "17760 Aristo Pha" -> "Aristo Pha".
    Alphanumerische Hersteller wie "1A Pharma" bleiben erhalten, weil nach der 1 kein Leerzeichen folgt.
    """
    if value is None:
        return None
    text = str(value).strip()
    m = re.match(r"^\d+\s+(.+)$", text)
    return m.group(1).strip() if m else value



class _CSVCell:
    def __init__(self, value):
        self.value = value

class _CSVWorksheet:
    """Kleine openpyxl-kompatible Hülle für CSV-Dateien.
    Sie stellt genau die Eigenschaften bereit, die find_columns(), Diagnose und
    die Export-Routine benötigen: max_row, max_column, ws[row], cell() und iter_rows().
    """
    def __init__(self, rows):
        self._rows = rows
        self.max_row = len(rows)
        self.max_column = max((len(r) for r in rows), default=0)

    def __getitem__(self, row_idx):
        row = self._rows[row_idx - 1] if 1 <= row_idx <= self.max_row else []
        padded = list(row) + [None] * max(0, self.max_column - len(row))
        return [_CSVCell(v) for v in padded]

    def cell(self, row, column):
        try:
            return _CSVCell(self._rows[row - 1][column - 1])
        except Exception:
            return _CSVCell(None)

    def iter_rows(self, min_row=1, values_only=False, **kwargs):
        max_row = kwargs.get('max_row') or self.max_row
        min_col = kwargs.get('min_col') or 1
        max_col = kwargs.get('max_col') or self.max_column
        for r in range(min_row, max_row + 1):
            source = self._rows[r - 1] if 1 <= r <= self.max_row else []
            vals = []
            for c in range(min_col, max_col + 1):
                vals.append(source[c - 1] if c - 1 < len(source) else None)
            if values_only:
                yield tuple(vals)
            else:
                yield tuple(_CSVCell(v) for v in vals)


def _read_csv_rows(path: Path):
    """Liest CSV robust mit typischen deutschen Export-Einstellungen.
    Unterstützt Semikolon, Komma und Tab sowie UTF-8/UTF-8-SIG/CP1252.
    """
    encodings = ['utf-8-sig', 'utf-8', 'cp1252', 'latin1']
    last_error = None
    for enc in encodings:
        try:
            text = path.read_text(encoding=enc)
            sample = text[:4096]
            try:
                dialect = csv.Sniffer().sniff(sample, delimiters=';,	|')
                delimiter = dialect.delimiter
            except Exception:
                delimiter = ';' if sample.count(';') >= sample.count(',') else ','
            rows = list(csv.reader(text.splitlines(), delimiter=delimiter))
            return [[cell.strip() if isinstance(cell, str) else cell for cell in row] for row in rows]
        except Exception as exc:
            last_error = exc
    raise ValueError(f'CSV-Datei konnte nicht gelesen werden: {last_error}')


def _load_input_worksheet(input_file: Path):
    suffix = input_file.suffix.lower()
    if suffix in {'.xlsx', '.xlsm', '.csv', '.txt'}:
        return load_worksheet(input_file)
    raise UnknownInputFormatError(
        'Nicht unterstütztes Dateiformat. Erlaubt sind: .xlsx, .xlsm, .csv und .txt'
    )

def _read_input_rows(input_file: Path):
    """
    Liest Rohdaten flexibel ein und baut exakt die Linden-Spalten A-G:
    A PZN, B Artikelname, C DF/DAR, D Pck/Pack.Gr, E Herst, F EK, G Absatz.

    Unterstützte Formate:
    - Klassische Hochpreiser-/Abverkaufslisten mit Header-Erkennung
    - CSV-Dateien mit automatischer Trennzeichen-/Encoding-Erkennung
    - Monatsverbrauchslisten wie EKM: PZN + YYYYMM-Spalten + Artikel/Hersteller/6 Monate/12 Monate
      Hier werden nur die Rohdaten-Spalten gelesen; vorhandene manuelle Auswertungsspalten rechts davon werden ignoriert.

    EK kommt zuerst aus der aktuellen Rohdatei. Wenn kein EK vorhanden ist, wird der letzte gespeicherte echte Rohdaten-EK als Fallback genutzt.
    """
    ws, input_typ = _load_input_worksheet(input_file)
    mapping = find_columns(ws)
    if not _is_mapping_usable(mapping):
        report = _diagnose_format(input_file, ws, mapping)
        raise UnknownInputFormatError(
            "Dateiformat nicht erkannt. Die Datei wurde nicht ausgewertet. "
            f"Bitte Diagnosebericht prüfen und Datei/Spalten anpassen: {report}",
            report
        )
    header_row = mapping.get("header_row", 1)
    format_typ = mapping.get("format", "standard")

    def get(row, key, fallback_idx=None):
        col = mapping.get(key)
        if col and len(row) >= col:
            return row[col - 1]
        if fallback_idx is not None and len(row) > fallback_idx:
            return row[fallback_idx]
        return None

    rows = []
    for row in ws.iter_rows(min_row=header_row + 1, values_only=True):
        if not any(v not in (None, "") for v in row):
            continue

        source = [None] * 7
        if format_typ == "monatsverbrauch":
            # EKM-/Monatsverbrauchsformat: keine Auswertungsspalten rechts übernehmen.
            source[0] = get(row, "pzn", 0)
            source[1] = get(row, "artikel")
            source[2] = None
            source[3] = None
            source[4] = _strip_leading_manufacturer_number(get(row, "hersteller"))
            source[5] = None  # EK gibt es in diesem Format nicht: bewusst leer lassen.
            source[6] = get(row, "absatz")
        else:
            # Fallback-Indizes entsprechen nur dann dem alten Linden-Layout,
            # wenn keine bessere Header-Erkennung vorhanden ist.
            source[0] = get(row, "pzn", 0)
            source[1] = get(row, "artikel", 1)
            source[2] = get(row, "df", 2)
            source[3] = get(row, "packung", 3)
            source[4] = get(row, "hersteller", 4 if not mapping.get("ek") == 5 else None)
            source[5] = get(row, "ek", 5)
            source[6] = get(row, "absatz", 6)

            # Falls EK per Header in Spalte E erkannt wurde und keine Hersteller-Spalte existiert,
            # darf der Wert aus Spalte E nicht als Hersteller-Fallback landen.
            if not mapping.get("hersteller"):
                source[4] = None

        # Nur echte PZN-Zeilen exportieren.
        if not _pzn(source[0]):
            continue
        rows.append(source)

    return rows, mapping

def create_linden_export(input_file: str | Path, apotheke: str) -> Path:
    input_file = Path(input_file)
    if not DB_PATH.exists():
        init_db(DB_PATH)
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    safe_name = "".join(ch if ch.isalnum() or ch in "-_" else "_" for ch in apotheke.strip()) or "Apotheke"
    out = OUTPUT_DIR / f"Auswertung_{safe_name}_{datetime.now():%Y%m%d_%H%M%S}.xlsx"

    rows, input_mapping = _read_input_rows(input_file)
    wb = Workbook()
    ws = wb.active
    ws.title = apotheke[:31] if apotheke else "Auswertung"
    wb.calculation.fullCalcOnLoad = True
    wb.calculation.forceFullCalc = True
    wb.calculation.calcMode = "auto"

    start_row = 1
    for col, header in enumerate(LINDEN_HEADERS, start=1):
        cell = ws.cell(start_row, col, header)
        cell.font = Font(bold=True)
        cell.fill = PatternFill("solid", fgColor="D9EAF7")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    with sqlite3.connect(DB_PATH) as con:
        con.row_factory = sqlite3.Row
        auswertung_id = None
        statistik = {"positionen": 0, "nmg_treffer": 0, "nicht_nmg": 0, "gesamt_absatz": 0.0}
        abgleich_rows = []
        abgleich_protocol_path = None
        try:
            cur = con.execute(
                """INSERT INTO tbl_auswertungen(apotheke, quelldatei, ausgabedatei, bemerkung)
                   VALUES (?, ?, ?, ?)""",
                (apotheke, input_file.name, str(out), "automatisch beim Export gespeichert")
            )
            auswertung_id = cur.lastrowid
        except Exception:
            auswertung_id = None
        try:
            con.execute("""INSERT INTO tbl_rohdaten_mapping(dateiname, pzn_spalte, hersteller_spalte, ek_spalte, absatz_spalte, header_zeile, format_typ, letzte_aktualisierung)
                         VALUES (?, ?, ?, ?, ?, ?, ?, CURRENT_TIMESTAMP)
                         ON CONFLICT(dateiname) DO UPDATE SET
                           pzn_spalte=excluded.pzn_spalte, hersteller_spalte=excluded.hersteller_spalte,
                           ek_spalte=excluded.ek_spalte, absatz_spalte=excluded.absatz_spalte,
                           header_zeile=excluded.header_zeile, format_typ=excluded.format_typ, letzte_aktualisierung=CURRENT_TIMESTAMP""",
                        (input_file.name, str(input_mapping.get("pzn_header", input_mapping.get("pzn", ""))),
                         str(input_mapping.get("hersteller_header", input_mapping.get("hersteller", ""))),
                         str(input_mapping.get("ek_header", input_mapping.get("ek", ""))),
                         str(input_mapping.get("absatz_header", input_mapping.get("absatz", ""))),
                         int(input_mapping.get("header_row", 1)), str(input_mapping.get("format", "standard"))))
            con.commit()
        except Exception:
            pass
        for r_idx, source_row in enumerate(rows, start=2):
            original_pzn = _pzn(source_row[0] if source_row else None)

            # A-G aus Rohdatei übernehmen. Fehlende Artikel/DF/PCK/Hersteller werden aus der PZN-Basisdatenbank ergänzt.
            basis = lookup_basisdaten(con, original_pzn)
            artikel_roh = source_row[1] if len(source_row) > 1 else None
            df_roh = source_row[2] if len(source_row) > 2 else None
            pck_roh = source_row[3] if len(source_row) > 3 else None
            hersteller_roh = clean_hersteller(source_row[4] if len(source_row) > 4 else None)

            artikel_final = artikel_roh or basis.get("artikelname")
            df_final = df_roh or basis.get("df")
            pck_final = pck_roh or basis.get("pck")
            hersteller_final = hersteller_roh or basis.get("herstellerkuerzel") or _lookup_hersteller(con, original_pzn)

            register_basisdaten(con, original_pzn, artikel_roh, hersteller_roh, df_roh, pck_roh, input_file.name)

            ek_roh = source_row[5] if len(source_row) > 5 else None
            ek_num = parse_number(ek_roh)
            if ek_num is not None:
                register_ek(con, original_pzn, ek_num, input_file.name, str(input_mapping.get("ek_header") or input_mapping.get("ek") or ""))
                ek_final = ek_num
            else:
                ek_final = lookup_latest_ek(con, original_pzn)

            output_values = [
                original_pzn, artikel_final, df_final, pck_final, hersteller_final,
                ek_final if ek_final is not None else None,
                source_row[6] if len(source_row) > 6 else None,
            ]
            for c_idx, value in enumerate(output_values, start=1):
                ws.cell(r_idx, c_idx, value)

            hit = _lookup(con, original_pzn)
            abgleich_trace = trace_lookup_source(con, original_pzn, hit)
            abgleich_trace.update({
                "zeile": r_idx,
                "artikel_alt": artikel_final or "",
                "df": df_final or "",
                "pck": pck_final or "",
                "hersteller": hersteller_final or "",
                "absatz_6m": source_row[6] if len(source_row) > 6 else "",
            })
            abgleich_rows.append(abgleich_trace)

            ws.cell(r_idx, 8, hit.get("im_sortiment"))
            ws.cell(r_idx, 9, hit.get("nmg_pzn"))
            ws.cell(r_idx, 10, hit.get("apu_nmg"))
            ws.cell(r_idx, 11, hit.get("rabatt"))
            ws.cell(r_idx, 12, hit.get("lieferbar"))
            ws.cell(r_idx, 13, hit.get("bevorratung"))
            ws.cell(r_idx, 14, hit.get("liefervorschlag"))
            ws.cell(r_idx, 15, hit.get("austauschbar_gegen"))

            # 0.7: Werte direkt berechnen, damit sie sofort sichtbar sind und auch beim
            # automatischen Vergleich nicht als leer erscheinen.
            # Logik bleibt: P = J × K, Q = P × G, R = J × G
            apu_wert = _to_number(hit.get("apu_nmg"), 0)
            rabatt_wert = _to_number(hit.get("rabatt"), 0)
            menge_wert = _to_number(source_row[6] if len(source_row) > 6 else 0, 0)
            ek_wert = _to_number(ek_final, 0)
            rabatt_euro = apu_wert * rabatt_wert
            ws.cell(r_idx, 16, rabatt_euro)
            ws.cell(r_idx, 17, rabatt_euro * menge_wert)
            # Umsatz wurde im geprüften Rosen/Linden-Stand aus EK × Absatz gebildet.
            # Das überschreibt die frühere Annahme R = J × G.
            umsatz_wert = ek_wert * menge_wert
            ws.cell(r_idx, 18, umsatz_wert)

            ist_nmg = 1 if (hit.get("nmg_pzn") or hit.get("im_sortiment") or hit.get("apu_nmg") is not None) else 0
            statistik["positionen"] += 1
            statistik["gesamt_absatz"] += menge_wert
            if ist_nmg:
                statistik["nmg_treffer"] += 1
            else:
                statistik["nicht_nmg"] += 1
            if auswertung_id is not None:
                try:
                    con.execute(
                        """INSERT INTO tbl_auswertungspositionen(
                            auswertung_id, pzn, artikelname, df, pck, herstellerkuerzel, ek, absatz_6m,
                            im_sortiment, pzn_nmg, apu_nmg, nmg_rabatt, lieferbar, bevorratung_angeraten,
                            liefervorschlag, austauschbar_gegen, nmg_rabatt_euro, nmg_rabatt_gesamt, umsatz,
                            ist_nmg_treffer, quelle
                        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                        (auswertung_id, original_pzn, artikel_final, df_final, pck_final, hersteller_final,
                         ek_final, menge_wert, hit.get("im_sortiment"), hit.get("nmg_pzn"),
                         hit.get("apu_nmg"), hit.get("rabatt"), hit.get("lieferbar"),
                         hit.get("bevorratung"), hit.get("liefervorschlag"), hit.get("austauschbar_gegen"),
                         rabatt_euro, rabatt_euro * menge_wert, umsatz_wert, ist_nmg, input_file.name)
                    )
                except Exception:
                    pass
        try:
            abgleich_protocol_path = write_abgleichartikel_protocol(
                abgleich_rows,
                analyse_name=apotheke,
                input_file=input_file,
                output_file=out,
                auswertung_id=auswertung_id,
            )
        except Exception:
            abgleich_protocol_path = None

        if auswertung_id is not None:
            try:
                bemerkung = "automatisch beim Export gespeichert"
                if abgleich_protocol_path:
                    bemerkung += f" | Abgleichartikel-Protokoll: {abgleich_protocol_path}"
                con.execute(
                    """UPDATE tbl_auswertungen
                       SET anzahl_positionen=?, nmg_treffer=?, nicht_nmg=?, gesamt_absatz=?, bemerkung=?
                       WHERE id=?""",
                    (statistik["positionen"], statistik["nmg_treffer"], statistik["nicht_nmg"],
                     statistik["gesamt_absatz"], bemerkung, auswertung_id)
                )
                con.commit()
            except Exception:
                pass

    last_row = max(2, len(rows) + 1)

    # Zahlenformat wie Linden: Preise/Rabatte/Umsatz sichtbar und berechenbar.
    for row in range(2, last_row + 1):
        for col in [6, 10, 16, 17, 18]:
            ws.cell(row, col).number_format = '#,##0.00'
        ws.cell(row, 11).number_format = '0.00%'
        ws.cell(row, 7).number_format = '0'

    thin = Side(style="thin", color="B7B7B7")
    for row in ws.iter_rows(min_row=1, max_row=last_row, min_col=1, max_col=18):
        for cell in row:
            cell.border = Border(left=thin, right=thin, top=thin, bottom=thin)
            cell.alignment = Alignment(vertical="top", wrap_text=True)

    widths = [10, 32, 8, 10, 10, 12, 18, 16, 12, 12, 12, 10, 16, 14, 28, 14, 18, 14]
    for idx, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(idx)].width = width
    ws.row_dimensions[1].height = 42
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:R{last_row}"
    wb.save(out)
    return out
