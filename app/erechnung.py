# -*- coding: utf-8 -*-
"""Gemeinsamer eRechnungs-Kern fuer die NMGone-Programmfamilie.

Eine Quelle fuer beide Richtungen:
  * VERSAND  (Faktura): aus Beleg+Positionen ein EN-16931-konformes XML
    erzeugen – Profil Factur-X/ZUGFeRD (CII). Optional in ein PDF/A-3
    einbetten (ZUGFeRD), falls die Bibliothek 'facturx' vorhanden ist;
    sonst wird eine eigenstaendige .xml-Datei geschrieben.
  * EMPFANG  (Wareneingang/Produktion, Buchhaltung): eine eingehende
    eRechnung einlesen – sowohl ZUGFeRD/Factur-X (CII) als auch
    XRechnung (UBL) – und in eine einheitliche dict-Struktur normalisieren.

Bewusst OHNE Fremdabhaengigkeiten (nur Standardbibliothek), damit der Kern
offline und ohne zusaetzliche Installation laeuft. Die KoSIT-Schematron-
Vollvalidierung ist ein spaeterer Schritt; hier gibt es eine Pruefung der
EN-16931-Pflichtfelder.

Normalisierte Datenstruktur ('daten'):
{
  'rechnungsnr': str,
  'typ': 'rechnung' | 'gutschrift',
  'datum': 'YYYY-MM-DD',
  'leistungsdatum': 'YYYY-MM-DD' | None,
  'waehrung': 'EUR',
  'verkaeufer': {name, strasse, plz, ort, land, ustid, steuernr, email},
  'kaeufer':    {name, strasse, plz, ort, land, ustid},
  'positionen': [{nr, pzn, bezeichnung, menge, einheit, einzelpreis, ust_satz, netto}],
  'summen': {netto, ust, brutto, steuer: [{satz, netto, ust, kategorie}]},
  'zahlung': {iban, bic, bank},
}
"""
from __future__ import annotations

import xml.etree.ElementTree as ET
from datetime import date, datetime
from pathlib import Path

# ── Namensraeume ─────────────────────────────────────────────────────────────
NS = {
    "rsm": "urn:un:unece:uncefact:data:standard:CrossIndustryInvoice:100",
    "ram": "urn:un:unece:uncefact:data:standard:ReusableAggregateBusinessInformationEntity:100",
    "udt": "urn:un:unece:uncefact:data:standard:UnqualifiedDataType:100",
}
# XRechnung / UBL (nur Empfang)
UBL = {
    "inv": "urn:oasis:names:specification:ubl:schema:xsd:Invoice-2",
    "cn":  "urn:oasis:names:specification:ubl:schema:xsd:CreditNote-2",
    "cac": "urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2",
    "cbc": "urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2",
}

EN16931_GUIDELINE = "urn:cen.eu:en16931:2017"
TYPECODE = {"rechnung": "380", "gutschrift": "381", "storno": "381"}
STD_EINHEIT = "C62"  # UN/ECE Rec 20: Stueck

for _p, _u in {**NS, **UBL}.items():
    ET.register_namespace(_p, _u)


# ── kleine Helfer ────────────────────────────────────────────────────────────
def _m(v) -> str:
    """Betrag mit 2 Nachkommastellen, Punkt als Dezimaltrenner (XML-Norm)."""
    try:
        return f"{float(v):.2f}"
    except (TypeError, ValueError):
        return "0.00"


def _pct(v) -> str:
    try:
        return f"{float(v):.2f}"
    except (TypeError, ValueError):
        return "0.00"


def _yyyymmdd(iso: str | None) -> str:
    s = str(iso or "").strip()
    for fmt in ("%Y-%m-%d", "%d.%m.%Y"):
        try:
            return datetime.strptime(s, fmt).strftime("%Y%m%d")
        except ValueError:
            continue
    return date.today().strftime("%Y%m%d")


def _q(prefix: str, tag: str) -> str:
    return f"{{{NS[prefix]}}}{tag}"


def _sub(parent, prefix, tag, text=None, **attrs):
    el = ET.SubElement(parent, _q(prefix, tag))
    if text is not None:
        el.text = str(text)
    for k, v in attrs.items():
        el.set(k, str(v))
    return el


def _num(s) -> float:
    """Robust gegen DE/EN-Zahlformat: '1.234,56' / '1234.56' / '' -> float."""
    s = str(s or "").replace("€", "").replace("%", "").replace(" ", "").strip()
    if not s:
        return 0.0
    if "," in s and "." in s:               # 1.234,56 -> 1234.56
        s = s.replace(".", "").replace(",", ".")
    elif "," in s:                          # 1234,56 -> 1234.56
        s = s.replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return 0.0


# ── Steueraufschluesselung aus Positionen ────────────────────────────────────
def _steuergruppen(positionen: list[dict]) -> list[dict]:
    """Gruppiert Positionen je USt-Satz -> [{satz, netto, ust, kategorie}]."""
    gruppen: dict[float, dict] = {}
    for p in positionen:
        satz = round(_num(p.get("ust_satz", 0)), 2)
        netto = _num(p.get("netto", 0))
        g = gruppen.setdefault(satz, {"satz": satz, "netto": 0.0, "ust": 0.0,
                                      "kategorie": "S" if satz > 0 else "Z"})
        g["netto"] += netto
    for g in gruppen.values():
        g["ust"] = round(g["netto"] * g["satz"] / 100.0, 2)
        g["netto"] = round(g["netto"], 2)
    return [gruppen[k] for k in sorted(gruppen)]


def summen_aus_positionen(positionen: list[dict]) -> dict:
    steuer = _steuergruppen(positionen)
    netto = round(sum(g["netto"] for g in steuer), 2)
    ust = round(sum(g["ust"] for g in steuer), 2)
    return {"netto": netto, "ust": ust, "brutto": round(netto + ust, 2), "steuer": steuer}


# ════════════════════════════════════════════════════════════════════════════
# VERSAND – EN-16931-CII (Factur-X / ZUGFeRD) erzeugen
# ════════════════════════════════════════════════════════════════════════════
def baue_factur_x_xml(daten: dict) -> bytes:
    """Erzeugt ein EN-16931-konformes CrossIndustryInvoice-XML (CII)."""
    d = daten
    pos = d.get("positionen") or []
    summen = d.get("summen") or summen_aus_positionen(pos)
    waehrung = d.get("waehrung") or "EUR"
    typ = (d.get("typ") or "rechnung").lower()

    root = ET.Element(_q("rsm", "CrossIndustryInvoice"))

    # 1) Kontext / Leitlinie
    ctx = _sub(root, "rsm", "ExchangedDocumentContext")
    gp = _sub(ctx, "ram", "GuidelineSpecifiedDocumentContextParameter")
    _sub(gp, "ram", "ID", EN16931_GUIDELINE)

    # 2) Dokumentkopf
    doc = _sub(root, "rsm", "ExchangedDocument")
    _sub(doc, "ram", "ID", d.get("rechnungsnr", ""))
    _sub(doc, "ram", "TypeCode", TYPECODE.get(typ, "380"))
    idt = _sub(doc, "ram", "IssueDateTime")
    _sub(idt, "udt", "DateTimeString", _yyyymmdd(d.get("datum")), format="102")

    # 3) Transaktion
    tx = _sub(root, "rsm", "SupplyChainTradeTransaction")

    # 3a) Positionen
    for i, p in enumerate(pos, start=1):
        li = _sub(tx, "ram", "IncludedSupplyChainTradeLineItem")
        adl = _sub(li, "ram", "AssociatedDocumentLineDocument")
        _sub(adl, "ram", "LineID", p.get("nr", i))
        prod = _sub(li, "ram", "SpecifiedTradeProduct")
        if p.get("pzn"):
            _sub(prod, "ram", "SellerAssignedID", p["pzn"])
        _sub(prod, "ram", "Name", p.get("bezeichnung", "") or "Position")
        agr = _sub(li, "ram", "SpecifiedLineTradeAgreement")
        npp = _sub(agr, "ram", "NetPriceProductTradePrice")
        _sub(npp, "ram", "ChargeAmount", _m(p.get("einzelpreis", 0)))
        deli = _sub(li, "ram", "SpecifiedLineTradeDelivery")
        _sub(deli, "ram", "BilledQuantity", _m(p.get("menge", 1)),
             unitCode=p.get("einheit") or STD_EINHEIT)
        sett = _sub(li, "ram", "SpecifiedLineTradeSettlement")
        tax = _sub(sett, "ram", "ApplicableTradeTax")
        satz = _num(p.get("ust_satz", 0))
        _sub(tax, "ram", "TypeCode", "VAT")
        _sub(tax, "ram", "CategoryCode", "S" if satz > 0 else "Z")
        _sub(tax, "ram", "RateApplicablePercent", _pct(satz))
        ms = _sub(sett, "ram", "SpecifiedTradeSettlementLineMonetarySummation")
        _sub(ms, "ram", "LineTotalAmount", _m(p.get("netto", 0)))

    # 3b) Beteiligte
    agr = _sub(tx, "ram", "ApplicableHeaderTradeAgreement")
    _party(agr, "SellerTradeParty", d.get("verkaeufer") or {}, mit_ustid=True)
    _party(agr, "BuyerTradeParty", d.get("kaeufer") or {}, mit_ustid=True)

    # 3c) Lieferung / Leistungsdatum
    deli = _sub(tx, "ram", "ApplicableHeaderTradeDelivery")
    if d.get("leistungsdatum"):
        ev = _sub(deli, "ram", "ActualDeliverySupplyChainEvent")
        odt = _sub(ev, "ram", "OccurrenceDateTime")
        _sub(odt, "udt", "DateTimeString", _yyyymmdd(d["leistungsdatum"]), format="102")

    # 3d) Abrechnung
    se = _sub(tx, "ram", "ApplicableHeaderTradeSettlement")
    _sub(se, "ram", "InvoiceCurrencyCode", waehrung)
    for g in summen["steuer"]:
        tax = _sub(se, "ram", "ApplicableTradeTax")
        _sub(tax, "ram", "CalculatedAmount", _m(g["ust"]))
        _sub(tax, "ram", "TypeCode", "VAT")
        _sub(tax, "ram", "BasisAmount", _m(g["netto"]))
        _sub(tax, "ram", "CategoryCode", g.get("kategorie", "S"))
        _sub(tax, "ram", "RateApplicablePercent", _pct(g["satz"]))
    mon = _sub(se, "ram", "SpecifiedTradeSettlementHeaderMonetarySummation")
    _sub(mon, "ram", "LineTotalAmount", _m(summen["netto"]))
    _sub(mon, "ram", "TaxBasisTotalAmount", _m(summen["netto"]))
    _sub(mon, "ram", "TaxTotalAmount", _m(summen["ust"]), currencyID=waehrung)
    _sub(mon, "ram", "GrandTotalAmount", _m(summen["brutto"]))
    _sub(mon, "ram", "DuePayableAmount", _m(summen["brutto"]))

    ET.indent(root, space="  ")
    return b'<?xml version="1.0" encoding="UTF-8"?>\n' + ET.tostring(root, encoding="utf-8")


def _party(parent, tag, info: dict, mit_ustid=False):
    party = _sub(parent, "ram", tag)
    _sub(party, "ram", "Name", info.get("name", "") or "")
    addr = _sub(party, "ram", "PostalTradeAddress")
    if info.get("plz"):
        _sub(addr, "ram", "PostcodeCode", info["plz"])
    if info.get("strasse"):
        _sub(addr, "ram", "LineOne", info["strasse"])
    if info.get("ort"):
        _sub(addr, "ram", "CityName", info["ort"])
    _sub(addr, "ram", "CountryID", (info.get("land") or "DE"))
    if mit_ustid and info.get("ustid"):
        reg = _sub(party, "ram", "SpecifiedTaxRegistration")
        _sub(reg, "ram", "ID", info["ustid"], schemeID="VA")


def schreibe_xml(daten: dict, ziel_pfad: str | Path) -> str:
    """Erzeugt das eRechnungs-XML und schreibt es nach ziel_pfad."""
    ziel = Path(ziel_pfad)
    ziel.parent.mkdir(parents=True, exist_ok=True)
    ziel.write_bytes(baue_factur_x_xml(daten))
    return str(ziel)


def betten_in_pdf(pdf_pfad: str | Path, daten: dict,
                  ziel_pdf: str | Path | None = None) -> str:
    """ZUGFeRD: bettet das XML in ein vorhandenes PDF (PDF/A-3) ein.

    Nutzt 'facturx', falls installiert. Ist die Bibliothek nicht vorhanden,
    wird stattdessen eine eigenstaendige .xml-Datei neben dem PDF geschrieben
    (gueltige eRechnung als XRechnung-Datentraeger) und deren Pfad geliefert.
    """
    pdf_pfad = Path(pdf_pfad)
    xml = baue_factur_x_xml(daten)
    try:
        from facturx import generate_from_file  # type: ignore
        ziel = Path(ziel_pdf or pdf_pfad)
        generate_from_file(str(pdf_pfad), xml, output_pdf_file=str(ziel),
                           flavor="factur-x", level="en16931", check_xsd=False)
        return str(ziel)
    except Exception:
        side = pdf_pfad.with_suffix(".xml")
        side.write_bytes(xml)
        return str(side)


# ════════════════════════════════════════════════════════════════════════════
# EMPFANG – eingehende eRechnung einlesen (CII oder UBL)
# ════════════════════════════════════════════════════════════════════════════
def lies_erechnung(quelle) -> dict:
    """Liest eine eRechnung ein. 'quelle' = Pfad, bytes oder XML-String.

    Erkennt ZUGFeRD/Factur-X (CII) und XRechnung (UBL). Bei einem PDF wird –
    sofern moeglich – das eingebettete XML extrahiert. Liefert die
    normalisierte dict-Struktur (siehe Modulkopf).
    """
    roh = _als_xml_bytes(quelle)
    if roh is None:
        raise ValueError("Keine eRechnung-XML-Daten gefunden (XML oder ZUGFeRD-PDF erwartet).")
    root = ET.fromstring(roh)
    tag = root.tag.rsplit("}", 1)[-1]
    if tag == "CrossIndustryInvoice":
        return _parse_cii(root)
    if tag in ("Invoice", "CreditNote"):
        return _parse_ubl(root, tag)
    raise ValueError(f"Unbekanntes eRechnungs-Format: <{tag}>")


def _als_xml_bytes(quelle) -> bytes | None:
    if isinstance(quelle, bytes):
        return quelle if b"<" in quelle[:200] else _xml_aus_pdf_bytes(quelle)
    if isinstance(quelle, str) and quelle.lstrip().startswith("<"):
        return quelle.encode("utf-8")
    p = Path(quelle)
    if p.suffix.lower() == ".pdf":
        return _xml_aus_pdf(p)
    return p.read_bytes()


def _xml_aus_pdf(pdf_pfad: Path) -> bytes | None:
    try:
        from facturx import get_facturx_xml_from_pdf  # type: ignore
        _name, xml = get_facturx_xml_from_pdf(str(pdf_pfad))
        if xml:
            return xml
    except Exception:
        pass
    return _xml_aus_pdf_bytes(pdf_pfad.read_bytes())


def _xml_aus_pdf_bytes(roh: bytes) -> bytes | None:
    """Best-effort: unkomprimiert eingebettetes CII/UBL-XML aus PDF-Bytes ziehen."""
    for start_tag, end_tag in (
        (b"<rsm:CrossIndustryInvoice", b"</rsm:CrossIndustryInvoice>"),
        (b"<ubl:Invoice", b"</ubl:Invoice>"),
        (b"<Invoice", b"</Invoice>"),
    ):
        i = roh.find(start_tag)
        j = roh.find(end_tag)
        if i != -1 and j != -1:
            return roh[i:j + len(end_tag)]
    return None


def _txt(el) -> str:
    return (el.text or "").strip() if el is not None else ""


def _parse_cii(root) -> dict:
    def f(path):
        return root.find(path, NS)

    doc = f(".//rsm:ExchangedDocument")
    typecode = _txt(doc.find("ram:TypeCode", NS)) if doc is not None else "380"
    datum = ""
    dt = f(".//rsm:ExchangedDocument/ram:IssueDateTime/udt:DateTimeString")
    if dt is not None:
        datum = _iso(_txt(dt))

    seller = f(".//ram:SellerTradeParty")
    buyer = f(".//ram:BuyerTradeParty")

    mon = f(".//ram:SpecifiedTradeSettlementHeaderMonetarySummation")
    netto = _num(_txt(mon.find("ram:TaxBasisTotalAmount", NS))) if mon is not None else 0.0
    ust = _num(_txt(mon.find("ram:TaxTotalAmount", NS))) if mon is not None else 0.0
    brutto = _num(_txt(mon.find("ram:GrandTotalAmount", NS))) if mon is not None else 0.0

    positionen = []
    for i, li in enumerate(root.findall(".//ram:IncludedSupplyChainTradeLineItem", NS), start=1):
        name = _txt(li.find("ram:SpecifiedTradeProduct/ram:Name", NS))
        pzn = _txt(li.find("ram:SpecifiedTradeProduct/ram:SellerAssignedID", NS))
        menge = _num(_txt(li.find("ram:SpecifiedLineTradeDelivery/ram:BilledQuantity", NS)))
        zeile = _num(_txt(li.find(".//ram:SpecifiedTradeSettlementLineMonetarySummation/"
                                  "ram:LineTotalAmount", NS)))
        satz = _txt(li.find(".//ram:ApplicableTradeTax/ram:RateApplicablePercent", NS))
        einzel = _num(_txt(li.find("ram:SpecifiedLineTradeAgreement/"
                                   "ram:NetPriceProductTradePrice/ram:ChargeAmount", NS)))
        if not einzel and menge:
            einzel = round(zeile / menge, 4)
        positionen.append({"nr": i, "pzn": pzn, "bezeichnung": name, "menge": menge,
                           "einheit": STD_EINHEIT, "einzelpreis": einzel,
                           "ust_satz": _num(satz), "netto": zeile})

    steuer = []
    for tax in root.findall(".//ram:ApplicableHeaderTradeSettlement/ram:ApplicableTradeTax", NS):
        steuer.append({"satz": _num(_txt(tax.find("ram:RateApplicablePercent", NS))),
                       "netto": _num(_txt(tax.find("ram:BasisAmount", NS))),
                       "ust": _num(_txt(tax.find("ram:CalculatedAmount", NS))),
                       "kategorie": _txt(tax.find("ram:CategoryCode", NS)) or "S"})

    return {
        "rechnungsnr": _txt(doc.find("ram:ID", NS)) if doc is not None else "",
        "typ": "gutschrift" if typecode == "381" else "rechnung",
        "datum": datum, "leistungsdatum": None, "waehrung": "EUR",
        "verkaeufer": _cii_party(seller),
        "kaeufer": _cii_party(buyer),
        "positionen": positionen,
        "summen": {"netto": netto, "ust": ust, "brutto": brutto, "steuer": steuer},
        "zahlung": {"iban": "", "bic": "", "bank": ""},
        "_format": "ZUGFeRD/Factur-X (CII)",
    }


def _cii_party(party) -> dict:
    if party is None:
        return {"name": "", "strasse": "", "plz": "", "ort": "", "land": "DE",
                "ustid": "", "steuernr": "", "email": ""}
    addr = party.find("ram:PostalTradeAddress", NS)
    ustid = ""
    for reg in party.findall("ram:SpecifiedTaxRegistration/ram:ID", NS):
        if reg.get("schemeID") == "VA":
            ustid = _txt(reg)
    return {
        "name": _txt(party.find("ram:Name", NS)),
        "strasse": _txt(addr.find("ram:LineOne", NS)) if addr is not None else "",
        "plz": _txt(addr.find("ram:PostcodeCode", NS)) if addr is not None else "",
        "ort": _txt(addr.find("ram:CityName", NS)) if addr is not None else "",
        "land": (_txt(addr.find("ram:CountryID", NS)) if addr is not None else "") or "DE",
        "ustid": ustid, "steuernr": "", "email": "",
    }


def _parse_ubl(root, wurzel="Invoice") -> dict:
    def t(parent, path):
        return _txt(parent.find(path, UBL)) if parent is not None else ""

    nr = t(root, "cbc:ID")
    datum = _iso(t(root, "cbc:IssueDate"))
    typecode = t(root, "cbc:InvoiceTypeCode") or ("381" if wurzel == "CreditNote" else "380")

    seller = root.find("cac:AccountingSupplierParty/cac:Party", UBL)
    buyer = root.find("cac:AccountingCustomerParty/cac:Party", UBL)

    tot = root.find("cac:LegalMonetaryTotal", UBL)
    netto = _num(t(tot, "cbc:TaxExclusiveAmount"))
    brutto = _num(t(tot, "cbc:TaxInclusiveAmount"))
    ust = round(brutto - netto, 2)
    tax_tot = root.find("cac:TaxTotal/cbc:TaxAmount", UBL)
    if tax_tot is not None:
        ust = _num(_txt(tax_tot))

    positionen = []
    zeilen = root.findall("cac:InvoiceLine", UBL) or root.findall("cac:CreditNoteLine", UBL)
    for i, ln in enumerate(zeilen, start=1):
        name = t(ln, "cac:Item/cbc:Name")
        pzn = t(ln, "cac:Item/cac:SellersItemIdentification/cbc:ID")
        menge = _num(t(ln, "cbc:InvoicedQuantity") or t(ln, "cbc:CreditedQuantity"))
        zeile = _num(t(ln, "cbc:LineExtensionAmount"))
        satz = _num(t(ln, "cac:Item/cac:ClassifiedTaxCategory/cbc:Percent"))
        einzel = _num(t(ln, "cac:Price/cbc:PriceAmount"))
        if not einzel and menge:
            einzel = round(zeile / menge, 4)
        positionen.append({"nr": i, "pzn": pzn, "bezeichnung": name, "menge": menge,
                           "einheit": STD_EINHEIT, "einzelpreis": einzel,
                           "ust_satz": satz, "netto": zeile})

    steuer = []
    for sub in root.findall("cac:TaxTotal/cac:TaxSubtotal", UBL):
        steuer.append({"satz": _num(t(sub, "cac:TaxCategory/cbc:Percent")),
                       "netto": _num(t(sub, "cbc:TaxableAmount")),
                       "ust": _num(t(sub, "cbc:TaxAmount")),
                       "kategorie": t(sub, "cac:TaxCategory/cbc:ID") or "S"})

    return {
        "rechnungsnr": nr,
        "typ": "gutschrift" if typecode == "381" else "rechnung",
        "datum": datum, "leistungsdatum": None, "waehrung": t(root, "cbc:DocumentCurrencyCode") or "EUR",
        "verkaeufer": _ubl_party(seller),
        "kaeufer": _ubl_party(buyer),
        "positionen": positionen,
        "summen": {"netto": netto, "ust": ust, "brutto": brutto, "steuer": steuer},
        "zahlung": {"iban": "", "bic": "", "bank": ""},
        "_format": "XRechnung (UBL)",
    }


def _ubl_party(party) -> dict:
    if party is None:
        return {"name": "", "strasse": "", "plz": "", "ort": "", "land": "DE",
                "ustid": "", "steuernr": "", "email": ""}
    name = _txt(party.find("cac:PartyLegalEntity/cbc:RegistrationName", UBL)) \
        or _txt(party.find("cac:PartyName/cbc:Name", UBL))
    addr = party.find("cac:PostalAddress", UBL)
    ustid = _txt(party.find("cac:PartyTaxScheme/cbc:CompanyID", UBL))
    return {
        "name": name,
        "strasse": _txt(addr.find("cbc:StreetName", UBL)) if addr is not None else "",
        "plz": _txt(addr.find("cbc:PostalZone", UBL)) if addr is not None else "",
        "ort": _txt(addr.find("cbc:CityName", UBL)) if addr is not None else "",
        "land": (_txt(addr.find("cac:Country/cbc:IdentificationCode", UBL))
                 if addr is not None else "") or "DE",
        "ustid": ustid, "steuernr": "", "email": "",
    }


def _iso(s: str) -> str:
    s = (s or "").strip()
    if len(s) == 8 and s.isdigit():                 # 20260630
        return f"{s[:4]}-{s[4:6]}-{s[6:]}"
    return s                                          # 2026-06-30 schon ISO


# ════════════════════════════════════════════════════════════════════════════
# PRUEFUNG – EN-16931-Pflichtfelder (kein Schematron, aber die wichtigsten BT)
# ════════════════════════════════════════════════════════════════════════════
def pruefe_en16931(daten: dict) -> list[str]:
    """Liefert eine Liste von Beanstandungen. Leere Liste = ok (Pflichtfelder)."""
    f = []
    if not daten.get("rechnungsnr"):
        f.append("Rechnungsnummer fehlt (BT-1).")
    if not daten.get("datum"):
        f.append("Rechnungsdatum fehlt (BT-2).")
    v = daten.get("verkaeufer") or {}
    if not v.get("name"):
        f.append("Verkaeufer-Name fehlt (BT-27).")
    if not (v.get("ustid") or v.get("steuernr")):
        f.append("Verkaeufer USt-IdNr. oder Steuernummer fehlt (BT-31/BT-32).")
    if not v.get("land"):
        f.append("Verkaeufer-Land fehlt (BT-40).")
    k = daten.get("kaeufer") or {}
    if not k.get("name"):
        f.append("Kaeufer-Name fehlt (BT-44).")
    if not (daten.get("positionen")):
        f.append("Keine Rechnungsposition vorhanden (BG-25).")
    s = daten.get("summen") or {}
    if not s.get("steuer"):
        f.append("Keine USt-Aufschluesselung vorhanden (BG-23).")
    # rechnerische Konsistenz
    if s:
        ist = round(_num(s.get("netto")) + _num(s.get("ust")), 2)
        soll = round(_num(s.get("brutto")), 2)
        if abs(ist - soll) > 0.02:
            f.append(f"Summen inkonsistent: netto+USt ({ist}) != brutto ({soll}).")
    return f


# ════════════════════════════════════════════════════════════════════════════
# ADAPTER – Faktura-Beleg -> normalisierte Daten (Versand)
# ════════════════════════════════════════════════════════════════════════════
def faktura_beleg_zu_daten(beleg: dict, positionen: list[dict], firma: dict,
                           kaeufer: dict | None = None) -> dict:
    """Mappt das Faktura-Datenmodell auf die normalisierte eRechnungs-Struktur."""
    pos = []
    for i, p in enumerate(positionen, start=1):
        netto = _num(p.get("netto_zeile"))
        menge = _num(p.get("menge")) or 1.0
        einzel = _num(p.get("apu_einzel"))
        pos.append({
            "nr": p.get("pos_nr") or i,
            "pzn": p.get("pzn") or "",
            "bezeichnung": p.get("bezeichnung") or "",
            "menge": menge,
            "einheit": STD_EINHEIT,
            "einzelpreis": einzel if einzel else (round(netto / menge, 4) if menge else 0.0),
            "ust_satz": _num(p.get("ust_satz")),
            "netto": netto,
        })
    summen = summen_aus_positionen(pos)
    # Stored beleg totals bevorzugen, falls vorhanden
    if _num(beleg.get("netto")):
        summen["netto"] = _num(beleg["netto"])
    if _num(beleg.get("ust_betrag")):
        summen["ust"] = _num(beleg["ust_betrag"])
    if _num(beleg.get("brutto")):
        summen["brutto"] = _num(beleg["brutto"])

    belegart = (beleg.get("belegart") or "rechnung").lower()
    kdaten = kaeufer or {
        "name": beleg.get("kunde_name") or "",
        "strasse": (beleg.get("kunde_adresse") or "").replace("\n", ", "),
        "plz": "", "ort": "", "land": "DE",
        "ustid": beleg.get("kunde_ustid") or "",
    }
    return {
        "rechnungsnr": beleg.get("beleg_nr") or "",
        "typ": "gutschrift" if belegart in ("gutschrift", "storno") else "rechnung",
        "datum": beleg.get("beleg_datum") or "",
        "leistungsdatum": beleg.get("leistungsdatum") or None,
        "waehrung": "EUR",
        "verkaeufer": firma,
        "kaeufer": kdaten,
        "positionen": pos,
        "summen": summen,
        "zahlung": {"iban": firma.get("iban", ""), "bic": firma.get("bic", ""),
                    "bank": firma.get("bank", "")},
    }


def faktura_firma_aus_settings() -> dict:
    """Liest die Firmenstammdaten aus den Faktura-Einstellungen (lazy import)."""
    from . import faktura_app as F  # lokal, um Zirkelimporte zu vermeiden
    g = F.get_setting
    return {
        "name": g("firma_name"), "strasse": g("firma_strasse"),
        "plz": g("firma_plz"), "ort": g("firma_ort"), "land": "DE",
        "ustid": g("firma_ustid"), "steuernr": g("firma_steuernr"),
        "email": g("firma_email"), "iban": g("firma_iban"),
        "bic": g("firma_bic"), "bank": g("firma_bank"),
    }


# ── Selbsttest: Round-trip Erzeugen -> Einlesen -> Pruefen ───────────────────
if __name__ == "__main__":
    demo = {
        "rechnungsnr": "RE-2026-00042",
        "typ": "rechnung",
        "datum": "2026-06-25",
        "leistungsdatum": "2026-06-20",
        "waehrung": "EUR",
        "verkaeufer": {"name": "NMG Pharma GmbH", "strasse": "Industriestr. 1",
                       "plz": "12345", "ort": "Musterstadt", "land": "DE",
                       "ustid": "DE123456789", "iban": "DE02100100100006820101"},
        "kaeufer": {"name": "Apotheke am Markt", "strasse": "Marktplatz 5",
                    "plz": "54321", "ort": "Beispielheim", "land": "DE",
                    "ustid": "DE987654321"},
        "positionen": [
            {"nr": 1, "pzn": "12345678", "bezeichnung": "Arzneimittel A",
             "menge": 10, "einzelpreis": 12.50, "ust_satz": 19, "netto": 125.00},
            {"nr": 2, "pzn": "87654321", "bezeichnung": "Arzneimittel B (7%)",
             "menge": 5, "einzelpreis": 8.00, "ust_satz": 7, "netto": 40.00},
        ],
    }
    demo["summen"] = summen_aus_positionen(demo["positionen"])

    xml = baue_factur_x_xml(demo)
    print("── erzeugtes EN-16931-CII-XML (Auszug) ──")
    print(xml.decode("utf-8")[:600], "...\n")

    daten = lies_erechnung(xml)
    print("── wieder eingelesen ──")
    print("Format     :", daten["_format"])
    print("Rechnungsnr:", daten["rechnungsnr"], "| Typ:", daten["typ"], "| Datum:", daten["datum"])
    print("Verkaeufer :", daten["verkaeufer"]["name"], "/", daten["verkaeufer"]["ustid"])
    print("Kaeufer    :", daten["kaeufer"]["name"])
    print("Summen     :", daten["summen"]["netto"], "+", daten["summen"]["ust"],
          "=", daten["summen"]["brutto"])
    print("Positionen :", len(daten["positionen"]),
          "| Steuergruppen:", [(g["satz"], g["ust"]) for g in daten["summen"]["steuer"]])

    fehler = pruefe_en16931(daten)
    print("\n── EN-16931-Pflichtfeldpruefung ──")
    print("OK" if not fehler else "\n".join("- " + x for x in fehler))
