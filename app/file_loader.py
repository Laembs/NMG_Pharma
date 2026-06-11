from __future__ import annotations

import csv
from pathlib import Path
from typing import Iterable

from openpyxl import load_workbook


SUPPORTED_DATA_EXTENSIONS = {".xlsx", ".xlsm", ".csv", ".txt"}
SUPPORTED_DATA_FILETYPES = [
    ("Daten", "*.xlsx *.xlsm *.csv *.txt"),
    ("Excel", "*.xlsx *.xlsm"),
    ("CSV", "*.csv"),
    ("Text", "*.txt"),
    ("Alle Dateien", "*.*"),
]


class UnsupportedFileFormatError(ValueError):
    pass


class TableData:
    """Einheitliche Tabellendaten fuer Excel, CSV und TXT.

    headers: erste erkannte Zeile
    rows: Datenzeilen danach
    source_type: excel/csv/txt
    delimiter: bei csv/txt erkanntes Trennzeichen
    """

    def __init__(self, headers, rows, source_type="", delimiter=None, sheet_name=""):
        self.headers = list(headers or [])
        self.rows = [list(r) for r in (rows or [])]
        self.source_type = source_type
        self.delimiter = delimiter
        self.sheet_name = sheet_name

    @property
    def row_count(self) -> int:
        return len(self.rows)

    @property
    def column_count(self) -> int:
        return max([len(self.headers)] + [len(r) for r in self.rows] + [0])


class _Cell:
    def __init__(self, value):
        self.value = value


class TableWorksheet:
    """Kleine openpyxl-kompatible Huelle.

    Damit koennen bestehende Funktionen wie find_columns(ws), Diagnose und
    iter_rows(values_only=True) auch CSV/TXT verarbeiten.
    """

    def __init__(self, table: TableData):
        self._rows = [table.headers] + table.rows
        self.max_row = len(self._rows)
        self.max_column = max((len(r) for r in self._rows), default=0)
        self.title = table.sheet_name or table.source_type or "Tabelle"

    def __getitem__(self, row_idx):
        row = self._rows[row_idx - 1] if 1 <= row_idx <= self.max_row else []
        padded = list(row) + [None] * max(0, self.max_column - len(row))
        return [_Cell(v) for v in padded]

    def cell(self, row, column):
        try:
            return _Cell(self._rows[row - 1][column - 1])
        except Exception:
            return _Cell(None)

    def iter_rows(self, min_row=1, max_row=None, min_col=1, max_col=None, values_only=False, **kwargs):
        max_row = max_row or self.max_row
        max_col = max_col or self.max_column
        for r in range(min_row, max_row + 1):
            source = self._rows[r - 1] if 1 <= r <= self.max_row else []
            vals = []
            for c in range(min_col, max_col + 1):
                vals.append(source[c - 1] if c - 1 < len(source) else None)
            if values_only:
                yield tuple(vals)
            else:
                yield tuple(_Cell(v) for v in vals)


def _read_text_with_fallback(path: Path) -> str:
    last_error = None
    for enc in ("utf-8-sig", "utf-8", "cp1252", "latin1"):
        try:
            return path.read_text(encoding=enc)
        except UnicodeDecodeError as exc:
            last_error = exc
    if last_error:
        return path.read_text(encoding="latin1", errors="replace")
    return path.read_text(errors="replace")


def _detect_delimiter(sample: str) -> str:
    """Erkennt uebliche Trenner fuer CSV/TXT.

    Reihenfolge/Logik:
    1. csv.Sniffer
    2. haeufigster Trenner aus ; , TAB |
    3. Semikolon als deutscher Standard
    """
    if not sample:
        return ";"

    try:
        dialect = csv.Sniffer().sniff(sample, delimiters=";\t,|")
        if dialect.delimiter:
            return dialect.delimiter
    except Exception:
        pass

    candidates = [";", "\t", "|", ","]
    counts = {d: sample.count(d) for d in candidates}
    best = max(counts, key=counts.get)
    return best if counts[best] > 0 else ";"


def _strip_empty_tail(row):
    row = list(row)
    while row and row[-1] in (None, ""):
        row.pop()
    return row


def _rows_from_text(path: Path) -> tuple[list[list[str]], str]:
    text = _read_text_with_fallback(path)
    sample = text[:8192]
    delimiter = _detect_delimiter(sample)

    rows = []
    reader = csv.reader(text.splitlines(), delimiter=delimiter)
    for row in reader:
        cleaned = [cell.strip() if isinstance(cell, str) else cell for cell in row]
        if any(cell not in (None, "") for cell in cleaned):
            rows.append(_strip_empty_tail(cleaned))
    return rows, delimiter


def load_table(path: str | Path, sheet_name: str | None = None) -> TableData:
    """Laedt xlsx/xlsm/csv/txt als einheitliche Tabelle.

    Bei Excel wird das erste aktive Blatt verwendet, falls kein sheet_name angegeben ist.
    Bei CSV/TXT wird die erste nicht-leere Zeile als Header verwendet.
    """
    path = Path(path)
    suffix = path.suffix.lower()

    if suffix not in SUPPORTED_DATA_EXTENSIONS:
        raise UnsupportedFileFormatError(
            f"Nicht unterstütztes Dateiformat: {suffix}. Erlaubt sind .xlsx, .xlsm, .csv und .txt."
        )

    if suffix in {".xlsx", ".xlsm"}:
        wb = load_workbook(path, data_only=True, read_only=True)
        ws = wb[sheet_name] if sheet_name else wb.active
        all_rows = []
        for row in ws.iter_rows(values_only=True):
            values = _strip_empty_tail(row)
            if any(v not in (None, "") for v in values):
                all_rows.append(values)
        headers = all_rows[0] if all_rows else []
        rows = all_rows[1:] if len(all_rows) > 1 else []
        return TableData(headers=headers, rows=rows, source_type="excel", sheet_name=ws.title)

    rows, delimiter = _rows_from_text(path)
    headers = rows[0] if rows else []
    data_rows = rows[1:] if len(rows) > 1 else []
    return TableData(headers=headers, rows=data_rows, source_type="txt" if suffix == ".txt" else "csv", delimiter=delimiter)


def load_worksheet(path: str | Path, sheet_name: str | None = None):
    """Laedt Dateien fuer bestehende openpyxl-nahe Verarbeitung.

    Rueckgabe:
    - Excel: echtes Worksheet
    - CSV/TXT: TableWorksheet
    - source_type: excel/csv/txt
    """
    path = Path(path)
    suffix = path.suffix.lower()

    if suffix in {".xlsx", ".xlsm"}:
        wb = load_workbook(path, data_only=True)
        ws = wb[sheet_name] if sheet_name else wb.active
        return ws, "excel"

    table = load_table(path, sheet_name=sheet_name)
    return TableWorksheet(table), table.source_type


def normalize_header(value) -> str:
    import re
    text = str(value or "").strip().lower()
    text = text.replace("ä", "ae").replace("ö", "oe").replace("ü", "ue").replace("ß", "ss")
    text = text.replace("_", " ").replace("-", " ")
    text = re.sub(r"\s+", " ", text)
    return text


def compact_header(value) -> str:
    import re
    return re.sub(r"[^a-z0-9]+", "", normalize_header(value))


def find_column(headers: Iterable, aliases: Iterable[str], contains: Iterable[str] = ()):
    normalized = [normalize_header(h) for h in headers]
    compacted = [compact_header(h) for h in headers]
    alias_norm = {normalize_header(a) for a in aliases}
    alias_compact = {compact_header(a) for a in aliases}
    contains_compact = [compact_header(c) for c in contains]

    for idx, h in enumerate(normalized):
        if h in alias_norm:
            return idx

    for idx, h in enumerate(compacted):
        if h in alias_compact:
            return idx
        if any(part and part in h for part in contains_compact):
            return idx

    return None
