"""buecherlisten/core: Laden, Fächerauswahl und PDF-Erzeugung ohne Netz.

Die Website-Zuordnungen werden über ``Zuordnungen`` vorgegeben; ein Test, der
sie vergäße, würde live die TRG-Website abfragen. Das Layout selbst prüft
diese Datei nicht - dafür war der Pixelvergleich beim Herauslösen aus
generate_booklists.py da (2026-09-17: 48 Seiten identisch).
"""
from __future__ import annotations

import io
import re

import pytest

pytest.importorskip("reportlab")

from bestand.core.testing import FakeClient  # noqa: E402
from buecherlisten.core.daten import lade_buecherdaten, waehle_faecher, waehle_gruppen  # noqa: E402
from buecherlisten.core.erzeugen import (  # noqa: E402
    Zuordnungen,
    erzeuge_buecherlisten_pdfs,
    erzeuge_schuelerlisten_pdfs,
)

OHNE_NETZ = Zuordnungen(fkl={}, kollegium={}, aufgabenfeld={"Erdkunde": "B", "Deutsch": "A"})


@pytest.fixture()
def daten():
    return lade_buecherdaten(FakeClient())


def _seiten(inhalt: bytes) -> int:
    return len(re.findall(rb"/Type\s*/Page[^s]", inhalt))


def test_faecher_und_schuljahr(daten):
    assert daten.faecher == ["Deutsch", "Erdkunde"]
    assert daten.schuljahr_id == "2026/2027"


def test_faecherauswahl_ohne_gross_klein(daten):
    gefunden, unbekannt = waehle_faecher(daten.faecher, ["erdkunde", "Latein", "ERDKUNDE"])
    assert gefunden == ["Erdkunde"]
    assert unbekannt == ["Latein"]


def test_ein_pdf_fuer_alle_faecher(daten):
    (pdf,) = erzeuge_buecherlisten_pdfs(daten, zuordnungen=OHNE_NETZ)
    assert pdf.inhalt.startswith(b"%PDF")
    assert pdf.dateiname == "Bücherliste Fächer 2026-2027.pdf"
    assert _seiten(pdf.inhalt) == 2
    assert pdf.warnungen == []


def test_split_liefert_ein_pdf_je_fach(daten):
    pdfs = erzeuge_buecherlisten_pdfs(daten, faecher=["Erdkunde"], modus="split", zuordnungen=OHNE_NETZ)
    assert [p.dateiname for p in pdfs] == ["Bücherliste Erdkunde 2026-2027.pdf"]


def test_bestaetigung_setzt_praefix(daten):
    (pdf,) = erzeuge_buecherlisten_pdfs(
        daten, bestaetigung=True, rueckgabe_bis="08.09.2026", rueckgabe_an="Ml", zuordnungen=OHNE_NETZ,
    )
    assert pdf.dateiname.startswith("Bestätigung ")
    assert pdf.inhalt.startswith(b"%PDF")


def test_doppelseitig_polstert_jedes_fach_auf_gerade_seitenzahl(daten):
    (pdf,) = erzeuge_buecherlisten_pdfs(daten, doppelseitig=True, zuordnungen=OHNE_NETZ)
    assert _seiten(pdf.inhalt) == 4


def test_falls_noetig_polstert_nicht_wenn_alles_einseitig_ist(daten):
    (pdf,) = erzeuge_buecherlisten_pdfs(
        daten, doppelseitig=True, nur_falls_noetig=True, zuordnungen=OHNE_NETZ,
    )
    assert _seiten(pdf.inhalt) == 2


def test_fehlende_website_wird_warnung_statt_abbruch(daten, monkeypatch):
    from buecherlisten.core import erzeugen

    def kaputt() -> dict[str, str]:
        raise OSError("kein Netz")

    monkeypatch.setattr(erzeugen, "fetch_fkl_mapping", kaputt)
    monkeypatch.setattr(erzeugen, "fetch_kollegium_kuerzel_mapping", kaputt)
    (pdf,) = erzeuge_buecherlisten_pdfs(daten, bestaetigung=True)
    assert pdf.inhalt.startswith(b"%PDF")
    assert len(pdf.warnungen) == 2
    assert all("kein Netz" in w for w in pdf.warnungen)


# ── Verlag, Jahrgang und Schülerliste ────────────────────────────────────────

def test_verlage_und_jahrgaenge_stehen_neben_den_faechern(daten):
    assert daten.verlage == ["Cornelsen", "Klett", "Westermann"]
    # Jahrgänge zählen numerisch, nicht alphabetisch ("Jahrgang 12" zuletzt).
    assert daten.jahrgaenge == ["Jahrgang 5", "Jahrgang 6", "Jahrgang 7", "Jahrgang 12"]
    assert daten.gruppen("verlag") == daten.verlage
    assert set(daten.listen_ids) == set(daten.jahrgaenge)


def test_jahrgangstabelle_behaelt_die_reihenfolge_der_iserv_liste(daten):
    zeilen = daten.je_jahrgang["Jahrgang 5"]["leih"]
    assert [z["titel"] for z in zeilen] == ["Deutschbuch 5", "Terra 5/6", "Deutschbuch eBook"]
    # Fach statt Klasse ist die Spalte, die die Jahrgangsliste braucht.
    assert zeilen[0]["fach"] == "Deutsch"


def test_mehrjahresband_steht_beim_verlag_einmal_mit_allen_klassen(daten):
    (terra,) = [z for z in daten.je_verlag["Klett"]["leih"] if z["titel"] == "Terra 5/6"]
    assert terra["klasse"] == "5, 6"


@pytest.mark.parametrize("ansicht, dateiname", [
    ("verlag", "Bücherliste Verlage 2026-2027.pdf"),
    ("jahrgang", "Bücherliste Jahrgänge 2026-2027.pdf"),
])
def test_ein_pdf_je_ansicht(daten, ansicht, dateiname):
    (pdf,) = erzeuge_buecherlisten_pdfs(daten, ansicht=ansicht, zuordnungen=OHNE_NETZ)
    assert pdf.dateiname == dateiname
    assert _seiten(pdf.inhalt) == len(daten.gruppen(ansicht))


def test_auswahl_folgt_der_reihenfolge_der_ansicht(daten):
    gefunden, unbekannt = waehle_gruppen(daten, "jahrgang", ["12", "jahrgang 5", "Jahrgang 3"])
    assert gefunden == ["Jahrgang 5", "Jahrgang 12"]
    assert unbekannt == ["Jahrgang 3"]


@pytest.mark.parametrize("optionen", [{"bestaetigung": True}, {"modus": "aufgabenfeld"}])
def test_bestaetigung_und_aufgabenfeld_nur_fuer_faecher(daten, optionen):
    with pytest.raises(ValueError):
        erzeuge_buecherlisten_pdfs(daten, ansicht="verlag", zuordnungen=OHNE_NETZ, **optionen)


def _iserv_pdf(seiten: int) -> bytes:
    from reportlab.lib.pagesizes import A4
    from reportlab.pdfgen.canvas import Canvas

    puffer = io.BytesIO()
    c = Canvas(puffer, pagesize=A4)
    for _ in range(seiten):
        c.drawString(72, 700, "IServ")
        c.showPage()
    c.save()
    return puffer.getvalue()


def test_schuelerliste_haengt_die_iserv_pdfs_aneinander(daten):
    (pdf,) = erzeuge_schuelerlisten_pdfs(daten, lambda listen_id: _iserv_pdf(1))
    assert pdf.dateiname == "Bücherliste Jahrgänge 2026-2027 (Schülerliste).pdf"
    assert _seiten(pdf.inhalt) == len(daten.jahrgaenge)


def test_schuelerliste_split_und_auswahl(daten):
    pdfs = erzeuge_schuelerlisten_pdfs(
        daten, lambda listen_id: _iserv_pdf(1), jahrgaenge=["Jahrgang 12", "Jahrgang 5"], modus="split",
    )
    assert [p.dateiname for p in pdfs] == [
        "Bücherliste Jahrgang 5 2026-2027 (Schülerliste).pdf",
        "Bücherliste Jahrgang 12 2026-2027 (Schülerliste).pdf",
    ]


def test_schuelerliste_polstert_nur_bei_bedarf_auf_gerade_seitenzahl(daten):
    einseitig = erzeuge_schuelerlisten_pdfs(
        daten, lambda listen_id: _iserv_pdf(1), doppelseitig=True, nur_falls_noetig=True,
    )[0]
    assert _seiten(einseitig.inhalt) == 4

    def gemischt(listen_id: int) -> bytes:
        return _iserv_pdf(3 if listen_id == daten.listen_ids["Jahrgang 5"] else 1)

    gepolstert = erzeuge_schuelerlisten_pdfs(
        daten, gemischt, doppelseitig=True, nur_falls_noetig=True,
    )[0]
    # 3 + 1 Leerseite für Jahrgang 5, je 1 + 1 für die drei anderen.
    assert _seiten(gepolstert.inhalt) == 10
