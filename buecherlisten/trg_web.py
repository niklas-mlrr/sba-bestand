"""Scraper für die drei öffentlichen Personen-/Fach-Tabellen der TRG-Website
(trg-osterode.de), aus generate_booklists.py herausgelöst.

Der Grund für die Trennung: generate_booklists.py importiert reportlab,
ausleihe, dotenv und isbnlib und lädt beim Import Umgebungsvariablen
(``load_dotenv``) — nichts davon lässt sich in einem Testprozess ohne Netz und
ohne Geschwister-Repo unfallfrei importieren. Dieses Modul importiert bewusst
nur ``html``/``re``/``difflib``/``requests`` und hat keine Import-Zeit-
Nebenwirkungen, ist also isoliert testbar (siehe tests/test_trg_web.py).

Jede öffentliche ``fetch_*``-Funktion ist deshalb nur eine dünne Hülle aus
Netzabruf (``_hole()``) und der zugehörigen ``parse_*``-Funktion — die
``parse_*``-Funktion enthält die eigentliche Logik, bekommt bereits geladenen
HTML-Text und macht selbst keinen Netzzugriff. Getestet wird ausschließlich
gegen die ``parse_*``-Funktionen (Fixtures statt Live-Abruf); ``fetch_*``
bleibt in Signatur und Rückgabewert unverändert, weil generate_booklists.py
sie unverändert weiter aufruft.

Rein lesend (nur GET) gegen die öffentliche TRG-Website — keine IServ-
Produktionsdaten.
"""
from __future__ import annotations

import difflib
import html
import re

import requests

# Fachkonferenzleitungen: liefert die Fach->Name-Zuordnung für die Kopfzeile
# bei --confirmation, sowie (als primäre Quelle) die Fach->Aufgabenfeld-
# Zuordnung für --mode aufgabenfeld.
FKL_URL = "https://trg-osterode.de/wir-am-trg/kollegium/fachkonferenzleitungen/"

# Fallback-Quelle für die Fach->Aufgabenfeld-Zuordnung (--mode aufgabenfeld),
# falls FKL_URL nicht (mehr) erreichbar ist oder deren Tabelle keine
# Aufgabenfeld-Zuordnung mehr enthält.
FAECHER_URL = "https://trg-osterode.de/unterricht-und-ganztagsangebot/faecher/"

# Kollegiumsseite: Name -> persönliches Lehrerkürzel (Spalte "Kürzel", z.B.
# "Mk" für Meike Menkens) — nicht zu verwechseln mit den Fach-Kürzeln in der
# "Fächer"-Spalte derselben Tabelle.
KOLLEGIUM_URL = "https://trg-osterode.de/wir-am-trg/kollegium/"


def _hole(url: str, timeout: float) -> str:
    """Lädt eine TRG-Seite als rohen HTML-Text.

    Einziger Netzzugriff dieses Moduls — mit derselben User-Agent-Kopfzeile
    wie zuvor in generate_booklists.py (ohne sie liefern manche WordPress-
    Installationen reduzierten/anderen HTML-Code an offensichtliche Bot-
    Requests aus). Nur GET, ``raise_for_status`` lässt HTTP-Fehler als
    Exception nach oben durchreichen — die Aufrufer in generate_booklists.py
    fangen sie bereits ab und sortieren im Fehlerfall alphabetisch bzw. lassen
    Name/Kürzel in der PDF-Kopfzeile weg.
    """
    resp = requests.get(url, timeout=timeout, headers={"User-Agent": "Mozilla/5.0"})
    resp.raise_for_status()
    return resp.text


def _clean_cell(raw_html: str) -> str:
    text = re.sub(r"<[^>]+>", " ", raw_html)
    text = html.unescape(text).replace("\xa0", " ")
    return re.sub(r"\s+", " ", text).strip()


def parse_fkl_mapping(html_text: str) -> dict[str, str]:
    """Fach -> Name der Fachkonferenzleitung, aus bereits geladenem HTML-Text.

    Die Seite enthält eine einzelne HTML-Tabelle (Fach, Name, Amtsbezeichnung)
    mit leeren Trenn-/Überschriftszeilen ("Aufgabenfeld A/B/C") dazwischen —
    die werden hier übersprungen (leere oder fehlende Zellen).

    Regex-Hinweis: Das Tabellen-Pattern ist bewusst ``<table[^>]*>`` und
    nicht mehr das früher hier verwendete attributlose ``<table>`` — die
    Live-FKL-Seite lieferte bislang ein bares ``<table>``, während die beiden
    anderen TRG-Tabellen (FAECHER_URL, KOLLEGIUM_URL) auf demselben
    WordPress-Install bereits ``<table class="has-background" ...>``
    schreiben. Bekäme die FKL-Seite künftig ebenfalls eine Klasse verpasst,
    hätte das attributlose Pattern lautlos {} zurückgegeben — keine
    Fachkonferenzleitungs-Namen mehr in den PDFs, ohne jede Fehlermeldung.
    Mit ``[^>]*`` matcht dieselbe Funktion beide Formen.
    """
    table_match = re.search(r"<table[^>]*>(.*?)</table>", html_text, re.S)
    if not table_match:
        return {}
    mapping: dict[str, str] = {}
    for row in re.findall(r"<tr>(.*?)</tr>", table_match.group(1), re.S):
        cells = re.findall(r"<td[^>]*>(.*?)</td>", row, re.S)
        if len(cells) < 2:
            continue
        subject, name = _clean_cell(cells[0]), _clean_cell(cells[1])
        if subject and name:
            mapping[subject] = name
    return mapping


def fetch_fkl_mapping(*, timeout: float = 10.0) -> dict[str, str]:
    """Fach -> Name der Fachkonferenzleitung, live von der TRG-Website gelesen.

    Siehe parse_fkl_mapping() für die eigentliche (netzlos testbare) Logik.
    """
    return parse_fkl_mapping(_hole(FKL_URL, timeout))


def find_mapped_value(mapping: dict[str, str], subject: str) -> str | None:
    """Wert zum Fach (Name oder Aufgabenfeld, je nach übergebener Zuordnung) —
    exakter Treffer zuerst, sonst Präfix-Abgleich in beide Richtungen
    (Ausleihe-API kennt z.B. nur "Politik", die Website führt
    "Politik-Wirtschaft")."""
    if not mapping:
        return None
    cf = subject.casefold()
    for key, name in mapping.items():
        if key.casefold() == cf:
            return name
    for key, name in mapping.items():
        key_cf = key.casefold()
        if key_cf.startswith(cf) or cf.startswith(key_cf):
            return name
    return None


def parse_aufgabenfeld_mapping(html_text: str) -> dict[str, str]:
    """Fach -> Aufgabenfeld ("Aufgabenfeld A"/"B"/"C"), aus derselben
    Fachkonferenzleitungen-Tabelle wie parse_fkl_mapping() gelesen.

    Die Tabelle gliedert die Fächer durch eigene Überschriftszeilen
    ("Aufgabenfeld A" in der ersten Spalte, restliche Spalten leer), gefolgt
    von den Fachzeilen dieses Aufgabenfelds bis zur nächsten Überschrift.

    Regex-Hinweis: siehe parse_fkl_mapping() — dasselbe ``<table[^>]*>``-
    Pattern, aus demselben Grund (bricht sonst lautlos weg, sobald die FKL-
    Seite eine Klasse am ``<table>`` bekommt).
    """
    table_match = re.search(r"<table[^>]*>(.*?)</table>", html_text, re.S)
    if not table_match:
        return {}
    mapping: dict[str, str] = {}
    current_field: str | None = None
    for row in re.findall(r"<tr>(.*?)</tr>", table_match.group(1), re.S):
        cells = [_clean_cell(c) for c in re.findall(r"<td[^>]*>(.*?)</td>", row, re.S)]
        if not cells or not cells[0]:
            continue
        if re.match(r"Aufgabenfeld\s+\S+", cells[0]) and not any(cells[1:]):
            current_field = cells[0]
            continue
        if current_field is not None:
            mapping[cells[0]] = current_field
    return mapping


def fetch_aufgabenfeld_mapping(*, timeout: float = 10.0) -> dict[str, str]:
    """Fach -> Aufgabenfeld, live von der TRG-Website gelesen.

    Siehe parse_aufgabenfeld_mapping() für die eigentliche (netzlos testbare)
    Logik.
    """
    return parse_aufgabenfeld_mapping(_hole(FKL_URL, timeout))


def parse_aufgabenfeld_mapping_from_faecher_page(html_text: str) -> dict[str, str]:
    """Fallback für parse_aufgabenfeld_mapping(): die TRG-Fächerübersichtsseite
    (FAECHER_URL) listet die Aufgabenfelder als drei Tabellenspalten (Kopfzeile
    "Aufgabenfeld A"/"B"/"C", darunter je Spalte die zugehörigen Fächer,
    unterschiedlich viele Zeilen pro Spalte — kürzere Spalten haben leere
    Zellen in den überzähligen Zeilen)."""
    table_match = re.search(r"<table[^>]*>(.*?)</table>", html_text, re.S)
    if not table_match:
        return {}
    rows = re.findall(r"<tr>(.*?)</tr>", table_match.group(1), re.S)
    if not rows:
        return {}
    headers = [_clean_cell(c) for c in re.findall(r"<td[^>]*>(.*?)</td>", rows[0], re.S)]
    mapping: dict[str, str] = {}
    for row in rows[1:]:
        cells = [_clean_cell(c) for c in re.findall(r"<td[^>]*>(.*?)</td>", row, re.S)]
        for i, subject in enumerate(cells):
            if subject and i < len(headers) and headers[i].startswith("Aufgabenfeld"):
                mapping[subject] = headers[i]
    return mapping


def fetch_aufgabenfeld_mapping_from_faecher_page(*, timeout: float = 10.0) -> dict[str, str]:
    """Fallback-Fetch für fetch_aufgabenfeld_mapping().

    Siehe parse_aufgabenfeld_mapping_from_faecher_page() für die eigentliche
    (netzlos testbare) Logik.
    """
    return parse_aufgabenfeld_mapping_from_faecher_page(_hole(FAECHER_URL, timeout))


def subject_sort_key(subject: str, aufgabenfeld_map: dict[str, str]) -> tuple[int, str, str]:
    """Sortierschlüssel für --mode aufgabenfeld: Fächer mit bekanntem
    Aufgabenfeld zuerst (gruppiert und alphabetisch je Feld), unbekannte
    Fächer danach, alphabetisch. Bei leerer `aufgabenfeld_map` (keine der
    Quellen erreichbar / kein Aufgabenfeld dort gefunden) landet jedes Fach in
    der zweiten Gruppe — das Ergebnis ist dann rein alphabetisch, wie
    gefordert."""
    field = find_mapped_value(aufgabenfeld_map, subject)
    if field is None:
        return (1, "", subject.casefold())
    return (0, field, subject.casefold())


def parse_kollegium_kuerzel_mapping(html_text: str) -> dict[str, str]:
    """Name -> Lehrerkürzel, aus bereits geladenem HTML-Text der TRG-
    Kollegiumsseite (Tabellenspalten: Name, Kürzel, Fächer, Amtsbezeichnung)."""
    table_match = re.search(r"<table[^>]*>(.*?)</table>", html_text, re.S)
    if not table_match:
        return {}
    mapping: dict[str, str] = {}
    for row in re.findall(r"<tr>(.*?)</tr>", table_match.group(1), re.S):
        cells = re.findall(r"<td[^>]*>(.*?)</td>", row, re.S)
        if len(cells) < 2:
            continue
        name, kuerzel = _clean_cell(cells[0]), _clean_cell(cells[1])
        if name and kuerzel and name != "Name":  # Kopfzeile der Tabelle überspringen
            mapping[name] = kuerzel
    return mapping


def fetch_kollegium_kuerzel_mapping(*, timeout: float = 10.0) -> dict[str, str]:
    """Name -> Lehrerkürzel, live von der TRG-Kollegiumsseite gelesen.

    Siehe parse_kollegium_kuerzel_mapping() für die eigentliche (netzlos
    testbare) Logik.
    """
    return parse_kollegium_kuerzel_mapping(_hole(KOLLEGIUM_URL, timeout))


def _normalize_person_name(name: str) -> str:
    return re.sub(r"[\s-]+", " ", name).strip().casefold()


def find_kollegium_kuerzel(mapping: dict[str, str], person_name: str | None) -> str | None:
    """Lehrerkürzel zu einem von find_fkl_name() gelieferten Namen — exakter
    Treffer, sonst Namensabgleich ohne Bindestrich-/Leerzeichen-Unterschiede
    (die beiden TRG-Seiten schreiben Namen nicht immer identisch)."""
    if not person_name or not mapping:
        return None
    if person_name in mapping:
        return mapping[person_name]
    target = _normalize_person_name(person_name)
    for name, abbrev in mapping.items():
        if _normalize_person_name(name) == target:
            return abbrev
    # Fuzzy-Fallback: die beiden TRG-Seiten schreiben denselben Namen nicht
    # immer identisch (beobachtet 2026-08-19: FKL-Seite "Kirscht-Nörthmann"
    # vs. Kollegiumsseite "Kirscht Nörthemann") — bei sehr hoher Ähnlichkeit
    # trotzdem zuordnen, statt das Kürzel ganz wegzulassen.
    best_name, best_ratio = None, 0.0
    for name in mapping:
        ratio = difflib.SequenceMatcher(None, _normalize_person_name(name), target).ratio()
        if ratio > best_ratio:
            best_name, best_ratio = name, ratio
    if best_name is not None and best_ratio >= 0.85:
        return mapping[best_name]
    return None
