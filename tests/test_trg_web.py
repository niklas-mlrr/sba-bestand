"""Tests für buecherlisten/trg_web.py — die drei TRG-Website-Scraper.

Alle Tests laufen gegen Fixtures unter tests/fixtures/trg/ (shape-genaue,
2026-09-05 von der Live-Website abgezogene HTML-Ausschnitte mit erfundenen
Namen, siehe Kommentarkopf jeder Fixture) und machen KEINEN Netzzugriff — es
werden ausschließlich die parse_*()-Funktionen aufgerufen, nie fetch_*()
(letztere sind nur requests.get()+raise_for_status() vor demselben parse_*,
siehe trg_web.py-Modul-Docstring, und brauchen deshalb keinen eigenen Test:
ein Mock von requests.get würde nur denselben parse_*()-Aufruf ein zweites Mal
prüfen)."""
from __future__ import annotations

from pathlib import Path

import pytest

from buecherlisten.trg_web import (
    _normalize_person_name,
    find_kollegium_kuerzel,
    find_mapped_value,
    parse_aufgabenfeld_mapping,
    parse_aufgabenfeld_mapping_from_faecher_page,
    parse_fkl_mapping,
    parse_kollegium_kuerzel_mapping,
    subject_sort_key,
)

_FIXTURES = Path(__file__).parent / "fixtures" / "trg"


def _read(name: str) -> str:
    return (_FIXTURES / name).read_text(encoding="utf-8")


@pytest.fixture()
def fkl_html() -> str:
    return _read("fkl_table.html")


@pytest.fixture()
def faecher_html() -> str:
    return _read("faecher_table.html")


@pytest.fixture()
def kollegium_html() -> str:
    return _read("kollegium_table.html")


# ── parse_fkl_mapping ────────────────────────────────────────────────────

def test_parse_fkl_mapping_liest_fach_zu_name(fkl_html: str) -> None:
    mapping = parse_fkl_mapping(fkl_html)
    # Stichprobe über alle drei Aufgabenfelder der Fixture, nicht nur das erste.
    assert mapping["Deutsch"] == "Frauke Ostmann"
    assert mapping["Politik-Wirtschaft"] == "Marlene Baumgart"
    assert mapping["Mathematik"] == "Dr. Andreas Wehrmann"
    assert len(mapping) == 18  # 18 Fächer in der Fixture, keine Überschriftszeile mitgezählt


def test_parse_fkl_mapping_ueberspringt_aufgabenfeld_ueberschriften(fkl_html: str) -> None:
    mapping = parse_fkl_mapping(fkl_html)
    assert "Aufgabenfeld A" not in mapping
    assert "Aufgabenfeld B" not in mapping
    assert "Aufgabenfeld C" not in mapping


def test_parse_fkl_mapping_ohne_table_gibt_leeres_dict(fkl_html: str) -> None:
    assert parse_fkl_mapping("<div>keine Tabelle hier</div>") == {}


def test_parse_fkl_mapping_verkraftet_table_mit_klasse(fkl_html: str) -> None:
    """Regressionstest für den in trg_web.py dokumentierten Bug: das Regex war
    früher für die FKL-Parser attributlos (`<table>`), während die anderen
    beiden TRG-Tabellen bereits `<table class="...">` schreiben. Bekäme die
    FKL-Seite eines Tages ebenfalls eine Klasse verpasst, hätte das alte
    Regex lautlos {} geliefert. Diese Fixture-Variante simuliert genau das."""
    variant = fkl_html.replace(
        "<table>", '<table class="has-background" style="background-color:#e7f5fe">', 1,
    )
    assert parse_fkl_mapping(variant) == parse_fkl_mapping(fkl_html)
    assert parse_fkl_mapping(variant)["Deutsch"] == "Frauke Ostmann"


# ── parse_aufgabenfeld_mapping (FKL-Seite) ──────────────────────────────

def test_parse_aufgabenfeld_mapping_gruppiert_nach_feld(fkl_html: str) -> None:
    mapping = parse_aufgabenfeld_mapping(fkl_html)
    assert mapping["Deutsch"] == "Aufgabenfeld A"
    assert mapping["Werte und Normen"] == "Aufgabenfeld B"
    assert mapping["Sport"] == "Aufgabenfeld C"
    assert len(mapping) == 18


def test_parse_aufgabenfeld_mapping_ueberspringt_ueberschriftszeilen_selbst(fkl_html: str) -> None:
    # Die Überschriftszeilen selbst ("Aufgabenfeld A" in Spalte 1, Rest leer)
    # dürfen nicht als Fach in der Zuordnung landen.
    mapping = parse_aufgabenfeld_mapping(fkl_html)
    assert "Aufgabenfeld A" not in mapping
    assert "Aufgabenfeld B" not in mapping
    assert "Aufgabenfeld C" not in mapping


def test_parse_aufgabenfeld_mapping_ohne_table_gibt_leeres_dict() -> None:
    assert parse_aufgabenfeld_mapping("<p>kein table-Tag</p>") == {}


def test_parse_aufgabenfeld_mapping_verkraftet_table_mit_klasse(fkl_html: str) -> None:
    variant = fkl_html.replace(
        "<table>", '<table class="has-background" style="background-color:#e7f5fe">', 1,
    )
    assert parse_aufgabenfeld_mapping(variant) == parse_aufgabenfeld_mapping(fkl_html)


# ── parse_aufgabenfeld_mapping_from_faecher_page (Fallback-Quelle) ──────

def test_parse_aufgabenfeld_mapping_from_faecher_page_liest_spalten(faecher_html: str) -> None:
    mapping = parse_aufgabenfeld_mapping_from_faecher_page(faecher_html)
    assert mapping["Deutsch"] == "Aufgabenfeld A"
    assert mapping["Erdkunde"] == "Aufgabenfeld B"
    assert mapping["Biologie"] == "Aufgabenfeld C"
    # Die API kennt nur "Politik", die Website hier aber auch nur "Politik"
    # (anders als auf der FKL-Seite, die "Politik-Wirtschaft" führt) — beide
    # Schreibweisen kommen in der Praxis vor, siehe find_mapped_value-Tests.
    assert mapping["Politik"] == "Aufgabenfeld B"


def test_parse_aufgabenfeld_mapping_from_faecher_page_verkraftet_unterschiedlich_lange_spalten(
    faecher_html: str,
) -> None:
    # Spalte C ("Aufgabenfeld C") hat in der Fixture nur 5 gefüllte Zeilen,
    # Spalte A/B je 7 — die überzähligen leeren Zellen dürfen keine Einträge
    # erzeugen (leerer subject-Wert wird übersprungen).
    mapping = parse_aufgabenfeld_mapping_from_faecher_page(faecher_html)
    assert "" not in mapping
    assert sum(1 for v in mapping.values() if v == "Aufgabenfeld C") == 5


def test_parse_aufgabenfeld_mapping_from_faecher_page_ohne_table_gibt_leeres_dict() -> None:
    assert parse_aufgabenfeld_mapping_from_faecher_page("<span>nichts</span>") == {}


# ── parse_kollegium_kuerzel_mapping ─────────────────────────────────────

def test_parse_kollegium_kuerzel_mapping_liest_name_zu_kuerzel(kollegium_html: str) -> None:
    mapping = parse_kollegium_kuerzel_mapping(kollegium_html)
    assert mapping["Dr. Andreas Wehrmann"] == "We"
    assert mapping["Maik Osterhoff"] == "Ost"
    assert len(mapping) == 7  # Kopfzeile ("Name") ist bewusst nicht mitgezählt


def test_parse_kollegium_kuerzel_mapping_ueberspringt_kopfzeile(kollegium_html: str) -> None:
    mapping = parse_kollegium_kuerzel_mapping(kollegium_html)
    assert "Name" not in mapping


def test_parse_kollegium_kuerzel_mapping_ohne_table_gibt_leeres_dict() -> None:
    assert parse_kollegium_kuerzel_mapping("<p>kein table</p>") == {}


# ── find_mapped_value ────────────────────────────────────────────────────

def test_find_mapped_value_exakter_treffer(fkl_html: str) -> None:
    mapping = parse_fkl_mapping(fkl_html)
    assert find_mapped_value(mapping, "Deutsch") == "Frauke Ostmann"


def test_find_mapped_value_praefix_website_laenger_als_api() -> None:
    # Ausleihe-API kennt nur "Politik", FKL-Seite führt "Politik-Wirtschaft".
    mapping = {"Politik-Wirtschaft": "Marlene Baumgart"}
    assert find_mapped_value(mapping, "Politik") == "Marlene Baumgart"


def test_find_mapped_value_praefix_api_laenger_als_website() -> None:
    # Umgekehrter Fall: die Zuordnung kennt nur das kurze Wort, gesucht wird
    # mit dem längeren Namen.
    mapping = {"Politik": "Marlene Baumgart"}
    assert find_mapped_value(mapping, "Politik-Wirtschaft") == "Marlene Baumgart"


def test_find_mapped_value_leere_zuordnung_gibt_none() -> None:
    assert find_mapped_value({}, "Deutsch") is None


def test_find_mapped_value_kein_treffer_gibt_none() -> None:
    assert find_mapped_value({"Deutsch": "Frauke Ostmann"}, "Chemie") is None


# ── subject_sort_key ─────────────────────────────────────────────────────

def test_subject_sort_key_bekanntes_feld_zuerst_dann_alphabetisch(fkl_html: str) -> None:
    aufgabenfeld_map = parse_aufgabenfeld_mapping(fkl_html)
    subjects = ["Sport", "Deutsch", "Zebra-AG", "Chemie"]
    ranked = sorted(subjects, key=lambda s: subject_sort_key(s, aufgabenfeld_map))
    # Deutsch (Feld A) vor Chemie (Feld C) vor Sport (auch Feld C, aber nach
    # Chemie alphabetisch) vor Zebra-AG (kein Feld bekannt -> ganz zuletzt).
    assert ranked == ["Deutsch", "Chemie", "Sport", "Zebra-AG"]


def test_subject_sort_key_leere_zuordnung_ist_rein_alphabetisch() -> None:
    # Docstring-Versprechen: keine Quelle erreichbar -> jedes Fach landet in
    # derselben Gruppe, das Ergebnis ist rein alphabetisch.
    subjects = ["Zebra-AG", "Anton-Kurs", "Mathematik"]
    ranked = sorted(subjects, key=lambda s: subject_sort_key(s, {}))
    assert ranked == ["Anton-Kurs", "Mathematik", "Zebra-AG"]


# ── _normalize_person_name / find_kollegium_kuerzel ─────────────────────

def test_normalize_person_name_vereinheitlicht_bindestrich_und_leerzeichen() -> None:
    assert _normalize_person_name("Birgit Sanders-Kühne") == _normalize_person_name(
        "Birgit Sanders Kühne",
    )


def test_find_kollegium_kuerzel_exakter_treffer(kollegium_html: str) -> None:
    mapping = parse_kollegium_kuerzel_mapping(kollegium_html)
    assert find_kollegium_kuerzel(mapping, "Maik Osterhoff") == "Ost"


def test_find_kollegium_kuerzel_normalisierter_treffer(fkl_html: str, kollegium_html: str) -> None:
    # FKL-Seite (Fixture): "Birgit Sanders-Kühne" (Bindestrich)
    # Kollegiumsseite (Fixture): "Birgit Sanders Kühne" (Leerzeichen)
    # — exakte Buchstaben sonst identisch, nur der Bindestrich/Leerzeichen-
    # Unterschied trennt sie, deshalb greift hier bereits der Normalisierungs-
    # Schritt (nicht erst der Fuzzy-Fallback).
    fkl_mapping = parse_fkl_mapping(fkl_html)
    kollegium_mapping = parse_kollegium_kuerzel_mapping(kollegium_html)
    fkl_name = fkl_mapping["Religion"]
    assert fkl_name == "Birgit Sanders-Kühne"
    assert fkl_name not in kollegium_mapping  # kein exakter Treffer möglich
    assert find_kollegium_kuerzel(kollegium_mapping, fkl_name) == "SaK"


def test_find_kollegium_kuerzel_fuzzy_fallback_bei_abweichender_schreibweise(
    fkl_html: str, kollegium_html: str,
) -> None:
    """Bildet den in find_kollegium_kuerzel() dokumentierten, real
    beobachteten Fall nach (dort: FKL "Kirscht-Nörthmann" vs. Kollegiumsseite
    "Kirscht Nörthemann" — Bindestrich UND eine abweichende Schreibvariante,
    sodass selbst die Normalisierung keinen Treffer liefert). Diese Fixture
    verwendet dieselbe Form mit erfundenen Namen: FKL "Petra Wilkens-Ahlgrim"
    vs. Kollegiumsseite "Petra Wilkens Ahlgrimm" (Bindestrich vs. Leerzeichen
    plus ein zusätzlicher Buchstabe am Ende)."""
    fkl_mapping = parse_fkl_mapping(fkl_html)
    kollegium_mapping = parse_kollegium_kuerzel_mapping(kollegium_html)
    fkl_name = fkl_mapping["Spanisch"]
    assert fkl_name == "Petra Wilkens-Ahlgrim"
    assert fkl_name not in kollegium_mapping
    assert _normalize_person_name(fkl_name) not in {
        _normalize_person_name(n) for n in kollegium_mapping
    }
    assert find_kollegium_kuerzel(kollegium_mapping, fkl_name) == "WA"


def test_find_kollegium_kuerzel_ohne_treffer_und_ohne_namen() -> None:
    mapping = {"Maik Osterhoff": "Ost"}
    assert find_kollegium_kuerzel(mapping, "Völlig Unbekannt") is None
    assert find_kollegium_kuerzel(mapping, None) is None
    assert find_kollegium_kuerzel({}, "Maik Osterhoff") is None
