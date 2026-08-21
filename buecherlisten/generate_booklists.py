#!/usr/bin/env python3
"""Bücherlisten eines Schuljahrs nach Fach aufbereitet als PDF.

Holt alle Bücherlisten (eine je Jahrgang) eines Schuljahrs über die
IServ-Ausleihe-API und stellt sie fachweise neu zusammen: pro Fach eine
Tabelle "Leihbare Bücher" (Spalten Klasse, Titel, Verlag, ISBN, Neupreis,
Leihgebühr) und eine Tabelle "Selbst anzuschaffende Bücher" (dieselben Spalten
ohne Leihgebühr — genau wie in den offiziellen IServ-Bücherlisten-PDFs, an
deren Aufmachung sich dieses Layout orientiert). Bücher, die in mehreren
Jahrgängen angeboten werden (Mehrjahresbände), erscheinen einmal pro Fach mit
allen betroffenen Klassen (z.B. "5, 6") und werden zuerst nach der untersten,
dann nach der zweituntersten Klasse einsortiert.

Spaltenbreiten: Klasse/ISBN/Neupreis/Leihgebühr brechen nie um. Titel und
Verlag dürfen umbrechen, tun es aber nur, wenn der Inhalt nicht einzeilig in
die Tabellenbreite passt — dann wird der Platz zwischen beiden so aufgeteilt,
dass die Tabelle insgesamt möglichst wenig Zeilen braucht ("Leihgebühr" wird
dabei zu "Leihgeb.", "Klasse" bei Bedarf zu "Kl."). Alle Spaltenzwischenräume
sind danach gleich breit, jede Spalte ist maximal so breit wie ihr
tatsächlich benötigter Inhalt (Details: `buecherlisten_layout.md` im Wiki).

Es gibt für diesen Anwendungsfall keinen eigenen API-Endpunkt — die Zusammen-
stellung passiert clientseitig aus den regulären Bücherlisten-Daten
(GET /schoolyears/:id/booklists/:bl_id, siehe
~/wiki/wiki/30_projects/sba/ausleihe_api/api_reference.md). Die ISBN-Formatierung
mit Bindestrichen nutzt dieselbe isbnlib-Maskierung wie das Bestand-Tooling
unter "tools/bestand/"; auch dafür gibt es keinen API-Weg, die
Ausleihe-API liefert ISBNs immer ohne Trennzeichen.

Rein lesend (nur GET). Kein Schreibzugriff auf die IServ-Produktionsdatenbank.

Verwendung:
  python3 generate_booklists.py [--schoulyear 2026/2027] [--mode combined|split]
                                 [--subjects "Fach1" "Fach2" ...] [--list-subjects]
                                 [--output-dir PFAD] [--confirmation] [--duplex | --duplex-if-needed]

  --schoolyear     Schuljahr wie "2026/2027" (Default: laufendes Schuljahr)
  --mode           combined = eine PDF-Datei mit einer neuen Seite pro Fach,
                              benannt "Bücherliste Fächer <Schuljahr>.pdf"
                   split    = eine PDF-Datei pro Fach,
                              benannt "Bücherliste <Fach> <Schuljahr>.pdf"
                   (Default: combined)
  --subjects       Nur diese Fächer aufnehmen (ein oder mehrere Namen, exakt
                    wie in der Bücherliste, z.B. --subjects Deutsch Mathematik).
                    Default: alle Fächer, die im Schuljahr vorkommen.
  --list-subjects  Nur die verfügbaren Fächer des Schuljahrs auflisten und
                    beenden (keine PDF-Erzeugung).
  --output-dir     Zielordner für die PDF(s) (Default: dieser Skriptordner)
  --confirmation   Bestätigungs-Block (Ankreuzfelder + Ort/Datum/Unterschrift
                    Fachkonferenzleitung <Fach>) am Ende jeder Fach-Liste
                    ergänzen. Lädt zusätzlich live die Fach->Name-Zuordnung
                    von der TRG-Website (Fachkonferenzleitungen) und zeigt den
                    Namen mittig in der Kopfzeile; schlägt der Abruf fehl oder
                    ist das Fach dort nicht gelistet, steht dort ersatzweise
                    "Bestätigung".
  --duplex         Für doppelseitigen Druck vorbereiten: jedes Fach bekommt
                    nötigenfalls eine leere Endseite, damit seine Seitenzahl
                    gerade ist (sonst würde beim doppelseitigen Druck das
                    nächste Fach auf der Rückseite der letzten Seite des
                    vorherigen beginnen).
  --duplex-if-needed  Wie --duplex, aber nur wirksam, wenn mindestens ein Fach
                    von Natur aus (ohne jede Polsterung) mehr als eine Seite
                    braucht. Sind alle Fächer ohnehin einseitig, bleibt die
                    Ausgabe unverändert (keine Leerseiten). Schließt sich mit
                    --duplex gegenseitig aus.
"""
from __future__ import annotations

import argparse
import difflib
import html
import io
import re
import sys
from collections import defaultdict
from datetime import datetime
from pathlib import Path

import requests

_HERE = Path(__file__).parent
_ROOT = _HERE.parent

# sba-bestand hält keine eigenen Secrets: Die IServ-Credentials liegen in der
# ``.env`` des Geschwister-Repos ausleihe-api. Beide Repos werden nebeneinander
# geklont (``<irgendein-ordner>/{ausleihe-api,sba-bestand}``).
_API_ROOT = _ROOT.parent / "ausleihe-api"

# ``ausleihe`` kommt normalerweise aus dem Venv (editable-Install, siehe
# pyproject). Fallback auf das Geschwister-Repo, damit das Skript auch ohne
# Installation läuft (Nachfolger-Pfad: nur klonen, nichts installieren).
if _API_ROOT.is_dir():
    sys.path.insert(0, str(_API_ROOT))

from dotenv import load_dotenv

load_dotenv(_API_ROOT / ".env")

from ausleihe import AusleiheClient  # noqa: E402
from ausleihe.exceptions import NotFoundError  # noqa: E402

try:
    import isbnlib as _isbnlib

    def format_isbn(isbn: str) -> str:
        try:
            masked = _isbnlib.mask(isbn)
            return masked if masked else isbn
        except Exception:
            return isbn
except ImportError:  # pragma: no cover

    def format_isbn(isbn: str) -> str:  # type: ignore[misc]
        return isbn


from reportlab.lib import colors  # noqa: E402
from reportlab.lib.enums import TA_JUSTIFY  # noqa: E402
from reportlab.lib.pagesizes import A4  # noqa: E402
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet  # noqa: E402
from reportlab.lib.units import mm  # noqa: E402
from reportlab.pdfbase.pdfmetrics import stringWidth  # noqa: E402
from reportlab.pdfgen.canvas import Canvas  # noqa: E402
from reportlab.platypus import (  # noqa: E402
    BaseDocTemplate,
    Flowable,
    Frame,
    NextPageTemplate,
    PageBreak,
    PageTemplate,
    Paragraph,
    Table,
    TableStyle,
)

# Schulname für die Fußzeile — an der Aufmachung der offiziellen IServ-
# Bücherlisten-PDFs orientiert (dort im Footer geführt).
SCHOOL_NAME = "Tilman-Riemenschneider-Gymnasium Osterode am Harz"

# ── Maße aus der offiziellen IServ-Bücherliste ───────────────────────────────
# Alle Werte unten sind aus "Bücherliste Jahrgang 5.pdf" ausgemessen
# (pdfplumber, 2026-08-18) und bewusst als absolute Punkt-Angaben gepflegt:
# Kopfbereich, Überschrift und Tabellen sollen deckungsgleich mit dem Original
# sitzen. Alles unterhalb der Tabellen (Fußzeile) ist davon ausgenommen.
PAGE_W, PAGE_H = A4

LEFT_MARGIN = 72.0                          # Textkante links (Original: x0 = 72.00)
RIGHT_EDGE = 537.28                         # Textkante rechts (Original: x1 = 537.28)
RIGHT_MARGIN = PAGE_W - RIGHT_EDGE          # = 58.0 (Original ist asymmetrisch)
CONTENT_WIDTH = RIGHT_EDGE - LEFT_MARGIN    # = 465.28 — auch die Tabellenbreite

# Grundlinien (Baselines) gemessen vom Seitenoberrand.
HEADER_LABEL_FONT, HEADER_LABEL_SIZE = "Helvetica", 8.0
HEADER_LABEL_BASELINE = PAGE_H - 45.74      # "Liste für" / "gültig für"
HEADER_VALUE_FONT, HEADER_VALUE_SIZE = "Helvetica-Bold", 12.0
HEADER_VALUE_BASELINE = PAGE_H - 56.02      # "Schuljahr 26/27" / "<Fach>"
TITLE_FONT, TITLE_SIZE = "Helvetica-Bold", 24.0
TITLE_BASELINE = PAGE_H - 105.73            # "Bücherliste <Fach>"

# Der Fließtext-Rahmen beginnt oben auf der Seite; der Kopf-/Titelblock wird
# absolut positioniert gezeichnet und reserviert nur seine Höhe (siehe
# SubjectHeading). INTRO_TOP ist die Oberkante des Einleitungstexts.
FRAME_TOP = PAGE_H - 30.0
INTRO_TOP = 714.0
HEADING_BLOCK_HEIGHT = FRAME_TOP - INTRO_TOP
BOTTOM_MARGIN = 18 * mm
# Die Fußzeile fluchtet mit dem Inhalt (LEFT_MARGIN/RIGHT_EDGE) — sie darf
# nicht breiter sein als Tabellen/Text darüber.
FOOTER_CENTER_X = (LEFT_MARGIN + RIGHT_EDGE) / 2

# Fußzeile exakt wie im Original ("Bücherliste Jahrgang 5.pdf", pdfplumber
# 2026-08-21): zwei Zeilen Helvetica 6pt in Schwarz, Grundlinien absolut vom
# Seitenunterrand aus gemessen. Obere Zeile rechtsbündig "Seite n von N",
# untere Zeile links das Erstelldatum (im Original: die URL) und rechtsbündig
# Schule + Kontext.
FOOTER_FONT, FOOTER_SIZE = "Helvetica", 6.0
FOOTER_PAGE_BASELINE = 25.692               # Grundlinie "Seite n von N"
FOOTER_INFO_BASELINE = 20.142               # Grundlinie Datum / Schule+Kontext
FOOTER_COLOR = colors.black
# Ortszusatz, den das Original hinter dem Schulnamen führt.
SCHOOL_CITY = "Osterode am Harz"

ACCENT_COLOR = colors.Color(0.6627, 0.2667, 0.2588)  # Original-Rot der Abschnittstitel
RULE_COLOR = colors.black                            # Linie unter der Kopfzeile (1.0pt schwarz)
ROW_RULE_COLOR = colors.Color(0.6667, 0.6667, 0.6667)  # Zeilentrenner (1.0pt grau)
GREY = colors.HexColor("#666666")

# Tabellen-Typografie exakt wie im Original: durchgehend Helvetica 10, nur die
# ISBN-Werte 8pt; Kopfzeile ist ebenfalls Helvetica 10 (nicht fett).
BODY_FONT = "Helvetica"
HEADER_FONT = "Helvetica"
CELL_FONT_SIZE = 10.0
ISBN_FONT_SIZE = 8.0
# Mindestbreite für Titel/Verlag beim Aufteilen des Restplatzes, damit keine
# der beiden Spalten auf ein einzelnes-Zeichen-pro-Zeile schrumpft.
MIN_WRAP_COL = 20 * mm
# Schrittweite der Breiten-Suche für die Titel/Verlag-Aufteilung (siehe
# _split_titel_verlag). 1pt ist für Tabellen dieser Größe schnell genug.
SPLIT_SEARCH_STEP = 1.0
# Mindestabstand zwischen zwei Spalten (14pt = 7pt je Seite), auch wenn der
# Inhalt die Tabellenbreite fast ausfüllt — sonst kann der rechnerisch
# gleichmäßig verteilte Rest pro Lücke gegen 0 gehen und Text ohne sichtbaren
# Abstand aneinanderstoßen (beobachtet 2026-08-18, "Klasse"-Wert direkt vor
# dem Titel).
MIN_GAP = 14.0
# Abstand einer Tabellenzeile zur Trennlinie darüber/darunter (Original: 4pt).
CELL_VPAD = 2.0
# Siehe TOPPADDING/BOTTOMPADDING in render_table.
VPAD_SHIFT = 0.66
# Siehe ISBN-TOPPADDING in render_table.
ISBN_VSHIFT = 1.278

STYLES = getSampleStyleSheet()
INTRO_STYLE = ParagraphStyle(
    "Intro", parent=STYLES["Normal"], fontName=BODY_FONT, fontSize=10, leading=12.5,
    textColor=colors.black, spaceAfter=6 * mm,
)
SECTION_STYLE = ParagraphStyle(
    "Abschnitt", parent=STYLES["Heading2"], fontName="Helvetica-Bold", fontSize=16,
    textColor=ACCENT_COLOR, spaceBefore=6 * mm, spaceAfter=3 * mm, leading=19,
)
EMPTY_STYLE = ParagraphStyle("Leer", parent=STYLES["Normal"], fontSize=9, textColor=GREY)
CONFIRM_STYLE = ParagraphStyle(
    "Bestaetigung", parent=STYLES["Normal"], fontName=BODY_FONT, fontSize=9, leading=11.5,
    alignment=TA_JUSTIFY,
)
SIGNATURE_LABEL_STYLE = ParagraphStyle(
    "SignaturLabel", parent=STYLES["Normal"], fontName=BODY_FONT, fontSize=8, leading=10,
)
# Kantenlänge des Ankreuzkästchens vor dem Bestätigungstext.
CHECKBOX_SIZE = 9.0
# Höhe der Unterschriftslinie über dem Label (Platz für die handschriftliche
# Unterschrift).
SIGNATURE_LINE_HEIGHT = 14 * mm

# Titel/Verlag brechen um (Paragraph); alle anderen Spalten bleiben Klartext
# in exakt passend berechneten Breiten (siehe render_table). splitLongWords=0
# verhindert, dass reportlab ein Wort mitten im Zeichen umbricht, wenn die
# zugewiesene Breite durch Rundung minimal knapper ist als das Wort selbst
# (beobachtet 2026-08-18, "Westermann" bei Erdkunde) — ein Wort, das nicht
# passt, läuft dann lieber minimal über, statt aufgetrennt zu werden.
CELL_STYLE = ParagraphStyle(
    "Zelle", parent=STYLES["Normal"], fontName=BODY_FONT, fontSize=CELL_FONT_SIZE, leading=11.5,
    splitLongWords=0,
)


def fmt_price(value: float | None) -> str:
    if value is None:
        return "–"
    try:
        value = float(value)
    except (TypeError, ValueError):
        return "–"
    if value == 0:
        return "–"
    return f"{value:.2f}".replace(".", ",") + " €"


def short_schoolyear(schoolyear_id: str) -> str:
    # "2026/2027" -> "26/27", wie im Kopfbereich der offiziellen
    # IServ-Bücherlisten-PDFs ("Liste für / Schuljahr 26/27").
    parts = schoolyear_id.split("/")
    if len(parts) == 2 and all(p.isdigit() for p in parts):
        return "/".join(p[-2:] for p in parts)
    return schoolyear_id


def fmt_grades(grades: tuple[int, ...]) -> str:
    # Komma + Leerzeichen statt "/": erlaubt Zeilenumbruch in der schmalen
    # Klasse-Spalte, wenn ein Mehrjahresband viele Klassen abdeckt.
    return ", ".join(str(g) for g in grades)


# Öffentliche TRG-Website (kein IServ, keine Produktionsdaten) — liefert die
# Fach->Name-Zuordnung der Fachkonferenzleitungen für die Kopfzeile bei
# --confirmation. Nur GET, rein lesend.
FKL_URL = "https://trg-osterode.de/wir-am-trg/kollegium/fachkonferenzleitungen/"


def _clean_cell(raw_html: str) -> str:
    text = re.sub(r"<[^>]+>", " ", raw_html)
    text = html.unescape(text).replace("\xa0", " ")
    return re.sub(r"\s+", " ", text).strip()


def fetch_fkl_mapping(*, timeout: float = 10.0) -> dict[str, str]:
    """Fach -> Name der Fachkonferenzleitung, live von der TRG-Website gelesen.

    Die Seite enthält eine einzelne HTML-Tabelle (Fach, Name, Amtsbezeichnung)
    mit leeren Trenn-/Überschriftszeilen ("Aufgabenfeld A/B/C") dazwischen —
    die werden hier übersprungen (leere oder fehlende Zellen).
    """
    resp = requests.get(FKL_URL, timeout=timeout, headers={"User-Agent": "Mozilla/5.0"})
    resp.raise_for_status()
    table_match = re.search(r"<table>(.*?)</table>", resp.text, re.S)
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


def find_fkl_name(mapping: dict[str, str], subject: str) -> str | None:
    """Name zum Fach — exakter Treffer zuerst, sonst Präfix-Abgleich in beide
    Richtungen (Ausleihe-API kennt z.B. nur "Politik", die Website führt
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


# Kollegiumsseite: Name -> persönliches Lehrerkürzel (Spalte "Kürzel", z.B.
# "Mk" für Meike Menkens) — nicht zu verwechseln mit den Fach-Kürzeln in der
# "Fächer"-Spalte derselben Tabelle.
KOLLEGIUM_URL = "https://trg-osterode.de/wir-am-trg/kollegium/"


def fetch_kollegium_kuerzel_mapping(*, timeout: float = 10.0) -> dict[str, str]:
    """Name -> Lehrerkürzel, live von der TRG-Kollegiumsseite gelesen
    (Tabellenspalten: Name, Kürzel, Fächer, Amtsbezeichnung)."""
    resp = requests.get(KOLLEGIUM_URL, timeout=timeout, headers={"User-Agent": "Mozilla/5.0"})
    resp.raise_for_status()
    table_match = re.search(r"<table[^>]*>(.*?)</table>", resp.text, re.S)
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


def collect_entries(client: AusleiheClient, schoolyear_id: str) -> dict[tuple[str, str], dict]:
    """Alle Bücherlisten-Items eines Schuljahrs, gruppiert nach (Fach, ISBN).

    Ein Buch, das in mehreren Jahrgangs-Bücherlisten desselben Fachs auftaucht
    (Mehrjahresband), wird zu einem Eintrag mit der Vereinigung der Klassen
    zusammengeführt — nicht anhand von series_data.gradesFlat (das ist ein
    globales Serien-Attribut und kann von den tatsächlichen Bücherlisten-
    Vorkommen abweichen, verifiziert 2026-08-18), sondern anhand der
    Bücherlisten-Jahrgänge, in denen das Item tatsächlich erscheint.
    """
    booklists = client.schoolyears.get_booklists(schoolyear_id)
    by_grade = {bl["grade"]: bl for bl in booklists if bl.get("grade") is not None}

    entries: dict[tuple[str, str], dict] = {}
    for grade in sorted(by_grade):
        bl = client.schoolyears.get_booklist(schoolyear_id, by_grade[grade]["id"])
        for section in bl.get("sections", []):
            for option in section.get("options", []):
                for item in option.get("items", []):
                    sd = item.get("series_data", {}) or {}
                    isbn = sd.get("isbn") or item.get("series")
                    if not isbn:
                        continue
                    subjects = sd.get("subjectsFlat") or ["(ohne Fach)"]
                    for subject in subjects:
                        key = (subject, isbn)
                        entry = entries.setdefault(
                            key,
                            {
                                "title": sd.get("title", "?"),
                                "publisher": sd.get("publisher", ""),
                                "price": sd.get("price"),
                                "fee": sd.get("fee"),
                                "borrowable": bool(item.get("borrowable")),
                                "grades": set(),
                            },
                        )
                        entry["grades"].add(grade)
    return entries


def build_subject_tables(entries: dict[tuple[str, str], dict]) -> dict[str, dict[str, list[dict]]]:
    """subject -> {"leih": [Zeilen...], "kauf": [Zeilen...]}, jeweils fertig sortiert."""
    by_subject: dict[str, dict[str, list[dict]]] = defaultdict(lambda: {"leih": [], "kauf": []})
    for (subject, isbn), e in entries.items():
        grades_sorted = tuple(sorted(e["grades"]))
        row = {
            "sort_key": (grades_sorted, e["title"].lower()),
            "klasse": fmt_grades(grades_sorted),
            "titel": e["title"],
            "verlag": e["publisher"],
            "isbn": format_isbn(isbn),
            "neupreis": fmt_price(e["price"]),
            "leihgebuehr": fmt_price(e["fee"]),
        }
        bucket = "leih" if e["borrowable"] else "kauf"
        by_subject[subject][bucket].append(row)

    for tables in by_subject.values():
        for bucket in ("leih", "kauf"):
            tables[bucket].sort(key=lambda r: r["sort_key"])
    return by_subject


# Kleiner Sicherheitszuschlag auf jede berechnete Inhaltsbreite: reportlabs
# Paragraph-Layout kann bei der Wortabstands-/Kerning-Behandlung minimal von
# unserer stringWidth-Schätzung abweichen. Ohne Puffer reicht das, um ein
# Wort exakt an der Kante nicht mehr passen zu lassen — mit splitLongWords=0
# (siehe CELL_STYLE) würde es dann zwar nicht mitten im Wort umgebrochen,
# aber minimal über den Zellenrand hinausragen.
WIDTH_EPSILON = 0.5


def _raw_width(header: str, values: list[str], *, value_size: float = CELL_FONT_SIZE) -> float:
    """Reine Inhaltsbreite (kein Zellenpolster) = längster Kopf- oder Zellentext.

    "Länge" heißt hier immer gemessene Textbreite (`stringWidth`), nie
    Zeichenanzahl — ein kurzes "M" ist breiter als ein langes "iii".
    """
    widths = [stringWidth(header, HEADER_FONT, CELL_FONT_SIZE)]
    widths += [stringWidth(v, BODY_FONT, value_size) for v in values]
    return (max(widths) + WIDTH_EPSILON) if widths else 0.0


def _wrap_lines(text: str, width: float, *, font: str = BODY_FONT, size: float = CELL_FONT_SIZE) -> list[str]:
    """Greedy Wortumbruch wie reportlabs Paragraph (Standard, ohne Silbentrennung).

    Bricht ausschließlich an Leerzeichen — ein einzelnes Wort, das breiter als
    `width` ist, bleibt trotzdem als Ganzes auf einer Zeile (siehe Gotcha zu
    splitLongWords bei CELL_STYLE)."""
    words = text.split()
    if not words:
        return [""]
    lines: list[str] = []
    current = words[0]
    for word in words[1:]:
        candidate = f"{current} {word}"
        if stringWidth(candidate, font, size) <= width:
            current = candidate
        else:
            lines.append(current)
            current = word
    lines.append(current)
    return lines


def _wrap_info(text: str, width: float, *, font: str = BODY_FONT, size: float = CELL_FONT_SIZE) -> tuple[int, float]:
    """(Zeilenzahl, breiteste tatsächlich benötigte Zeile) beim Umbruch auf `width`."""
    lines = _wrap_lines(text, width, font=font, size=size)
    max_line_w = max((stringWidth(line, font, size) for line in lines), default=0.0)
    return len(lines), max_line_w + WIDTH_EPSILON


def _max_word_width(values: list[str], *, font: str = BODY_FONT, size: float = CELL_FONT_SIZE) -> float:
    """Breitestes einzelnes Wort über alle Werte — `_wrap_lines` bricht nie
    innerhalb eines Worts um, also kann eine Spalte nie schmaler werden als
    das breiteste unteilbare Wort, das in ihr vorkommt."""
    widest = 0.0
    for v in values:
        for word in v.split():
            widest = max(widest, stringWidth(word, font, size))
    return widest + WIDTH_EPSILON if widest else 0.0


def _split_titel_verlag(
    titel_vals: list[str], verlag_vals: list[str], available: float,
) -> tuple[float, float, float, float]:
    """Titel/Verlag-Breite so wählen, dass die Tabelle insgesamt am wenigsten
    Zeilen braucht (= am kleinsten ist); bei Gleichstand die Aufteilung, die
    zusätzlich am wenigsten Breite tatsächlich braucht.

    Rückgabe: (titel_budget, verlag_budget, titel_actual, verlag_actual) —
    *_actual ist die nach dem Umbruch bei diesem Budget tatsächlich breiteste
    Zeile (meist etwas schmaler als das Budget, siehe "eine Spalte nur
    maximal so lang wie der längste Inhalt einer Zeile").
    """
    if not titel_vals:
        return 0.0, 0.0, 0.0, 0.0

    # Kein Budget darf unter das breiteste unteilbare Wort der jeweiligen
    # Spalte fallen — sonst wird die tatsächlich gerenderte Zeile breiter als
    # das der anderen Spalte zugestandene Budget, und die Tabelle läuft über
    # die rechte Kante hinaus (beobachtet 2026-08-18, "Bibliographisches").
    min_titel = max(MIN_WRAP_COL, _max_word_width(titel_vals))
    min_verlag = max(MIN_WRAP_COL, _max_word_width(verlag_vals))

    if available <= min_titel + min_verlag:
        # Nicht genug Platz für beide Mindestbreiten — beide bekommen exakt
        # ihr Minimum; die Tabelle kann dann geringfügig breiter werden als
        # CONTENT_WIDTH (unvermeidbar bei einem einzelnen sehr breiten Wort).
        return min_titel, min_verlag, min_titel, min_verlag

    lo, hi = min_titel, available - min_verlag
    best: tuple[tuple[int, float], float, float, float] | None = None
    w = lo
    while w <= hi + 1e-6:
        verlag_w = available - w
        total_lines = 0
        max_titel_line = 0.0
        max_verlag_line = 0.0
        for t, v in zip(titel_vals, verlag_vals):
            lt, mt = _wrap_info(t, w)
            lv, mv = _wrap_info(v, verlag_w)
            total_lines += max(lt, lv)
            max_titel_line = max(max_titel_line, mt)
            max_verlag_line = max(max_verlag_line, mv)
        key = (total_lines, max_titel_line + max_verlag_line)
        if best is None or key < best[0]:
            best = (key, w, max_titel_line, max_verlag_line)
        w += SPLIT_SEARCH_STEP

    assert best is not None
    _, titel_w, titel_actual, verlag_actual = best
    return titel_w, available - titel_w, titel_actual, verlag_actual


def render_table(rows: list[dict], *, with_fee: bool) -> Table:
    isbn_idx, neupreis_idx = 3, 4
    n_cols = 6 if with_fee else 5
    n_gaps = n_cols - 1

    # Für jeden der n_gaps Zwischenräume wird vorab MIN_GAP reserviert, bevor
    # Spaltenbreiten überhaupt berechnet werden — sonst kann der nach Schritt 3
    # rechnerisch übrige Platz pro Lücke gegen 0 gehen, wenn der Inhalt die
    # Tabellenbreite fast ausfüllt (beobachtet 2026-08-18: "12Duden..." ohne
    # sichtbaren Abstand). effective_width ist das Budget, mit dem Schritt 1/2
    # rechnen — der danach ermittelte Gesamtinhalt passt dadurch garantiert
    # mit mindestens MIN_GAP Luft pro Lücke in CONTENT_WIDTH.
    effective_width = CONTENT_WIDTH - MIN_GAP * n_gaps

    klasse_vals = [r["klasse"] for r in rows]
    titel_vals = [r["titel"] for r in rows]
    verlag_vals = [r["verlag"] for r in rows]
    isbn_vals = [r["isbn"] for r in rows]
    neupreis_vals = [r["neupreis"] for r in rows]
    leihgebuehr_vals = [r["leihgebuehr"] for r in rows] if with_fee else []

    klasse_header = "Klasse"
    leihgebuehr_header = "Leihgebühr"

    # 1) Passt alles einzeilig (jede Spalte auf ihre natürliche Breite,
    #    Titel/Verlag inklusive) in die Tabellenbreite? Dann muss nichts
    #    umbrechen und nichts abgekürzt werden.
    isbn_w = _raw_width("ISBN", isbn_vals, value_size=ISBN_FONT_SIZE)
    neupreis_w = _raw_width("Neupreis", neupreis_vals)
    natural_klasse_w = _raw_width(klasse_header, klasse_vals)
    natural_leihgebuehr_w = _raw_width(leihgebuehr_header, leihgebuehr_vals) if with_fee else 0.0
    natural_titel_w = _raw_width("Titel", titel_vals)
    natural_verlag_w = _raw_width("Verlag", verlag_vals)
    natural_total = (
        natural_klasse_w + natural_titel_w + natural_verlag_w + isbn_w + neupreis_w + natural_leihgebuehr_w
    )

    if natural_total <= effective_width:
        klasse_w, titel_w, verlag_w, leihgebuehr_w = (
            natural_klasse_w, natural_titel_w, natural_verlag_w, natural_leihgebuehr_w,
        )
    else:
        # 2) Reicht nicht — zuerst Platz durch Abkürzen der Kopfzeilen
        #    zurückgewinnen: "Leihgebühr" wird immer zu "Leihgeb."; "Klasse"
        #    nur, wenn der Spaltentitel selbst breiter ist als der breiteste
        #    Klassen-Wert (sonst würde die Abkürzung nichts bringen).
        if with_fee:
            leihgebuehr_header = "Leihgeb."
        klasse_data_w = max((stringWidth(v, BODY_FONT, CELL_FONT_SIZE) for v in klasse_vals), default=0.0)
        if stringWidth(klasse_header, HEADER_FONT, CELL_FONT_SIZE) > klasse_data_w:
            klasse_header = "Kl."
        klasse_w = _raw_width(klasse_header, klasse_vals)
        leihgebuehr_w = _raw_width(leihgebuehr_header, leihgebuehr_vals) if with_fee else 0.0

        fixed_w = klasse_w + isbn_w + neupreis_w + leihgebuehr_w
        available_tv = effective_width - fixed_w
        _, _, titel_w, verlag_w = _split_titel_verlag(titel_vals, verlag_vals, available_tv)

        # 2b) Abkürzungen zurücknehmen, wenn nach der Titel/Verlag-Aufteilung
        #     wieder Platz dafür ist — zuerst "Klasse", danach (mit dem dann
        #     schon etwas größeren Gesamtinhalt) "Leihgebühr". Erst danach
        #     wird verteilt, damit die dadurch länger gewordenen Spalten in
        #     der Gleichverteilung berücksichtigt sind.
        total_now = klasse_w + titel_w + verlag_w + isbn_w + neupreis_w + leihgebuehr_w
        if klasse_header == "Kl.":
            full_klasse_w = _raw_width("Klasse", klasse_vals)
            if total_now + (full_klasse_w - klasse_w) <= effective_width:
                total_now += full_klasse_w - klasse_w
                klasse_w = full_klasse_w
                klasse_header = "Klasse"
        if with_fee and leihgebuehr_header == "Leihgeb.":
            full_leihgebuehr_w = _raw_width("Leihgebühr", leihgebuehr_vals)
            if total_now + (full_leihgebuehr_w - leihgebuehr_w) <= effective_width:
                total_now += full_leihgebuehr_w - leihgebuehr_w
                leihgebuehr_w = full_leihgebuehr_w
                leihgebuehr_header = "Leihgebühr"

    # 3) Restplatz gleichmäßig auf alle Spaltenzwischenräume verteilen —
    #    jede Spalte bleibt exakt so breit wie ihr tatsächlich benötigter
    #    Inhalt (keine Spalte länger als der längste Inhalt einer Zeile).
    #    Dank effective_width in Schritt 1/2 ist gap hier immer >= MIN_GAP.
    total_content = klasse_w + titel_w + verlag_w + isbn_w + neupreis_w + leihgebuehr_w
    gap = max(CONTENT_WIDTH - total_content, 0.0) / n_gaps if n_gaps else 0.0
    half_gap = gap / 2

    # colWidths braucht die VOLLE Spaltenbreite inkl. ihres Anteils an der
    # Lücke (reportlab zieht das TableStyle-Padding von colWidth ab, um die
    # Textfläche zu bestimmen — reine Inhaltsbreite hier würde bei jeder
    # Spalte außer der ersten/letzten ins Negative laufen).
    content_widths = [klasse_w, titel_w, verlag_w, isbn_w, neupreis_w]
    if with_fee:
        content_widths.append(leihgebuehr_w)
    col_widths = [
        w + (0.0 if i == 0 else half_gap) + (0.0 if i == n_cols - 1 else half_gap)
        for i, w in enumerate(content_widths)
    ]

    cols = [klasse_header, "Titel", "Verlag", "ISBN", "Neupreis"]
    if with_fee:
        cols.append(leihgebuehr_header)

    data = [list(cols)]
    for r in rows:
        line = [
            r["klasse"],
            Paragraph(r["titel"], CELL_STYLE),
            Paragraph(r["verlag"], CELL_STYLE),
            r["isbn"],
            r["neupreis"],
        ]
        if with_fee:
            line.append(r["leihgebuehr"])
        data.append(line)

    last_row = len(data) - 1
    # Tabellenbild wie im Original: keine Füllfarben, kein Gitternetz, 1.0pt
    # schwarze Linie unter der Kopfzeile, 1.0pt graue Trennlinie unter jeder
    # Datenzeile. Klasse/Titel/Verlag/ISBN linksbündig, Neupreis/Leihgebühr
    # rechtsbündig (nur Datenzeilen — die Kopfzeile bleibt linksbündig).
    style = [
        ("FONTNAME", (0, 0), (-1, 0), HEADER_FONT),
        ("FONTNAME", (0, 1), (-1, -1), BODY_FONT),
        ("FONTSIZE", (0, 0), (-1, -1), CELL_FONT_SIZE),
        ("FONTSIZE", (isbn_idx, 1), (isbn_idx, -1), ISBN_FONT_SIZE),
        ("ALIGN", (0, 0), (-1, 0), "LEFT"),
        ("ALIGN", (neupreis_idx, 1), (-1, -1), "RIGHT"),
        ("LINEBELOW", (0, 0), (-1, 0), 1.0, RULE_COLOR),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        # Reportlab nimmt für Klartext-Zellen (Klasse/ISBN/Neupreis/Leihgebühr)
        # standardmäßig eine andere Zeilenhöhe an (fontSize*1.2) als für die
        # Paragraph-Zellen Titel/Verlag (CELL_STYLE.leading=11.5) — das ergab
        # bei einzeiligen Zeilen eine andere Zeilenhöhe als bei umbrochenen und
        # dadurch einen row-abhängigen Rest-Versatz. LEADING gleicht das an.
        ("LEADING", (0, 0), (-1, -1), CELL_STYLE.leading),
        # TOPPADDING/BOTTOMPADDING sind bewusst NICHT symmetrisch: reportlab
        # reserviert bei TOP-VALIGN über der Textzeile die volle Oberlänge und
        # darunter die volle Unterlänge. Optisch maßgeblich ist aber die
        # Versalhöhe oben (Oberkante Großbuchstabe) und die Grundlinie unten
        # (Unterkante ohne g/p/q/Komma) — dazwischen sitzt der Text sonst zu
        # hoch. VPAD_SHIFT verschiebt ihn so weit nach unten, dass der Abstand
        # Versalhöhe→Linie darüber und Grundlinie→Linie darunter gleich groß
        # ist (kalibriert 2026-08-19: gemessen wurde die Grundlinie aus der
        # Text-Matrix des PDFs plus Helvetica-CapHeight 718/1000 — pdfplumbers
        # char-Bounding-Box taugt dafür nicht, sie ist für JEDES Zeichen exakt
        # 1 em hoch, egal ob "D" oder "g"). Summe TOP+BOTTOM bleibt 2*CELL_VPAD,
        # die Zeilenhöhe ändert sich also nicht. Reine Y-Position —
        # Schriftgröße/-farbe bleiben unverändert.
        ("TOPPADDING", (0, 0), (-1, -1), CELL_VPAD - VPAD_SHIFT),
        ("BOTTOMPADDING", (0, 0), (-1, -1), CELL_VPAD + VPAD_SHIFT),
        # ISBN ist mit 8pt kleiner als der restliche 10pt-Zeilentext (siehe
        # ISBN_FONT_SIZE) und hat damit eine kleinere Versalhöhe. Damit ihr
        # Versalhöhe/Grundlinie-Kasten auf derselben Mitte sitzt wie der der
        # 10pt-Spalten, muss ihre Grundlinie um die halbe Versalhöhen-Differenz
        # höher liegen — ISBN_VSHIFT stellt genau das ein. Reine Y-Position,
        # Schriftgröße bleibt 8pt.
        # ISBN_VSHIFT wird der ISBN-Spalte oben aufgeschlagen und unten wieder
        # abgezogen — die Zellenhöhe (und damit ggf. die Zeilenhöhe bei
        # einzeiligen Zeilen) bleibt dadurch identisch zu den anderen Spalten.
        ("TOPPADDING", (isbn_idx, 1), (isbn_idx, -1), CELL_VPAD - VPAD_SHIFT + ISBN_VSHIFT),
        ("BOTTOMPADDING", (isbn_idx, 1), (isbn_idx, -1), CELL_VPAD + VPAD_SHIFT - ISBN_VSHIFT),
        # Gleicher Abstand zwischen allen Spalten: jede Innenkante bekommt die
        # halbe Lücke, außen (erste/letzte Spalte) bleibt 0 — Tabelle fluchtet
        # weiterhin links mit "Liste für"/Überschrift, rechts mit "gültig für".
        ("LEFTPADDING", (0, 0), (-1, -1), half_gap),
        ("RIGHTPADDING", (0, 0), (-1, -1), half_gap),
        ("LEFTPADDING", (0, 0), (0, -1), 0),
        ("RIGHTPADDING", (-1, 0), (-1, -1), 0),
    ]
    if last_row >= 1:
        style.append(("LINEBELOW", (0, 1), (-1, last_row), 1.0, ROW_RULE_COLOR))
    table = Table(data, colWidths=col_widths, repeatRows=1, hAlign="LEFT")
    table.setStyle(TableStyle(style))
    return table


class SubjectHeading(Flowable):
    """Kopfbereich + Überschrift an den exakten Positionen des Originals.

    Die drei Textzeilen werden **absolut** auf der Seite gezeichnet (Baselines
    und x-Kanten aus der offiziellen IServ-Bücherliste ausgemessen), nicht im
    normalen Textfluss positioniert. Der Flowable reserviert im Fluss nur die
    Höhe des Blocks, damit der Einleitungstext darunter beginnt.
    """

    def __init__(self, schoolyear_id: str, subject: str, *, confirm_value: str | None = None) -> None:
        super().__init__()
        self.schoolyear_id = schoolyear_id
        self.subject = subject
        self.confirm_value = confirm_value
        self.width = CONTENT_WIDTH
        self.height = HEADING_BLOCK_HEIGHT

    def draw(self) -> None:
        c = self.canv
        # draw() zeichnet in lokalen Koordinaten — Ursprung auf die absoluten
        # Seitenkoordinaten zurückrechnen, damit der Block unabhängig von
        # seiner Position im Fluss immer exakt gleich sitzt.
        origin_x, origin_y = c.absolutePosition(0, 0)
        dx, dy = -origin_x, -origin_y

        c.saveState()
        c.setFillColor(colors.black)
        c.setFont(HEADER_LABEL_FONT, HEADER_LABEL_SIZE)
        c.drawString(dx + LEFT_MARGIN, dy + HEADER_LABEL_BASELINE, "Liste für")
        c.drawRightString(dx + RIGHT_EDGE, dy + HEADER_LABEL_BASELINE, "gültig für")

        c.setFont(HEADER_VALUE_FONT, HEADER_VALUE_SIZE)
        c.drawString(
            dx + LEFT_MARGIN, dy + HEADER_VALUE_BASELINE,
            f"Schuljahr {short_schoolyear(self.schoolyear_id)}",
        )
        c.drawRightString(dx + RIGHT_EDGE, dy + HEADER_VALUE_BASELINE, self.subject)

        if self.confirm_value:
            center_x = dx + FOOTER_CENTER_X
            confirm_label = "Bücherlisten zu Bestätigen durch"

            # Label-Zeile ("Liste für"/"gültig für"-Höhe): kurzer, fester
            # Text — bei den üblichen Fach-/Jahr-Kürzeln links/rechts nie eng.
            c.setFont(HEADER_LABEL_FONT, HEADER_LABEL_SIZE)
            c.drawCentredString(center_x, dy + HEADER_LABEL_BASELINE, confirm_label)

            # Wert-Zeile ("Schuljahr .."/<Fach>-Höhe): Kürzel mittig — bei
            # langem Fachnamen + langem Ersatzwert (Fallback: Name statt
            # Kürzel) notfalls verkleinern, damit nichts mit den beiden
            # Rand-Werten kollidiert (analog Fußzeile).
            left_val_w = c.stringWidth(
                f"Schuljahr {short_schoolyear(self.schoolyear_id)}", HEADER_VALUE_FONT, HEADER_VALUE_SIZE,
            )
            right_val_w = c.stringWidth(self.subject, HEADER_VALUE_FONT, HEADER_VALUE_SIZE)
            left_end = dx + LEFT_MARGIN + left_val_w
            right_start = dx + RIGHT_EDGE - right_val_w
            max_half_width = max(min(center_x - left_end, right_start - center_x) - 6, 10)

            center_fs = HEADER_VALUE_SIZE
            while center_fs > 7.0 and c.stringWidth(self.confirm_value, HEADER_VALUE_FONT, center_fs) / 2 > max_half_width:
                center_fs -= 0.25

            c.setFont(HEADER_VALUE_FONT, center_fs)
            c.drawCentredString(center_x, dy + HEADER_VALUE_BASELINE, self.confirm_value)

        c.setFont(TITLE_FONT, TITLE_SIZE)
        c.drawString(dx + LEFT_MARGIN, dy + TITLE_BASELINE, f"Bücherliste {self.subject}")
        c.restoreState()


class ConfirmationBlock(Flowable):
    """Einleitungssatz + zwei Ankreuzfelder + Unterschrift-Zeile am Fach-Listenende.

    Einleitungssatz: "Hiermit bestätige ich im Namen der Fachschaft <Fach>,
    dass ...". Oberes Kreuz: die Liste ist NICHT korrekt und soll um die
    handschriftlich eingetragenen Änderungen ergänzt werden. Unteres Kreuz:
    die Bücher sind richtig. Fließt normal im Textfluss mit — bei kurzen
    Listen landet der Block dadurch am Ende der Tabellen, bei sehr langen
    ggf. auf einer eigenen Folgeseite.
    """

    # Abstand zwischen den beiden Ankreuzzeilen und zwischen der zweiten
    # Ankreuzzeile und der Unterschriftszeile darunter.
    CHECKBOX_ROW_GAP = 4 * mm
    # Abstand zwischen dem einleitenden Satz und der ersten Ankreuzzeile.
    INTRO_GAP = 2 * mm

    def __init__(self, subject: str, schoolyear_id: str, *, teacher_kuerzel: str | None = None) -> None:
        super().__init__()
        self.width = CONTENT_WIDTH
        signature_label = f"Unterschrift Fachkonferenzleitung {subject}"
        if teacher_kuerzel:
            signature_label += f" ({teacher_kuerzel})"
        self._intro_par = Paragraph(
            f"Hiermit bestätige ich im Namen der Fachschaft {subject}, dass die oben "
            f"aufgeführten Bücher der Bücherliste {subject} und insbesondere deren ISBN "
            f"für das Schuljahr {schoolyear_id}",
            CONFIRM_STYLE,
        )
        _, self._intro_h = self._intro_par.wrap(self.width, 0xFFFFFF)

        text_width = self.width - CHECKBOX_SIZE - 6
        self._checkbox_pars = [
            Paragraph(
                "nicht korrekt ist und um die handschriftlichen Anmerkungen "
                "(Durchstreichungen; Eintragungen neuer Bücher, Klassen, ...) "
                "verändert werden muss.",
                CONFIRM_STYLE,
            ),
            Paragraph(
                "korrekt ist, an die Schüler übermittelt werden kann und eine "
                "nachträgliche Änderung unter Umständen nicht mehr gestattet werden "
                "kann.",
                CONFIRM_STYLE,
            ),
        ]
        self._checkbox_heights = [p.wrap(text_width, 0xFFFFFF)[1] for p in self._checkbox_pars]
        checkbox_block_h = (
            sum(max(h, CHECKBOX_SIZE) for h in self._checkbox_heights)
            + self.CHECKBOX_ROW_GAP
        )

        # Linke Hälfte: "Ort, Datum" mit einer Linie. Rechte Hälfte:
        # "Unterschrift Fachkonferenzleitung <Fach>" mit einer eigenen Linie.
        # Eine leere Spalter dazwischen (ohne Linie) sorgt für einen
        # sichtbaren Spalt zwischen den beiden Linien.
        sig_col_w = (self.width - MIN_GAP) / 2
        self._sig_table = Table(
            [
                ["", "", ""],
                [
                    Paragraph("Ort, Datum", SIGNATURE_LABEL_STYLE),
                    "",
                    Paragraph(signature_label, SIGNATURE_LABEL_STYLE),
                ],
            ],
            colWidths=[sig_col_w, MIN_GAP, sig_col_w],
            rowHeights=[SIGNATURE_LINE_HEIGHT, None],
        )
        self._sig_table.setStyle(TableStyle([
            ("LINEABOVE", (0, 1), (0, 1), 0.75, colors.black),
            ("LINEABOVE", (2, 1), (2, 1), 0.75, colors.black),
            ("TOPPADDING", (0, 1), (-1, 1), 2.0),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
            ("TOPPADDING", (0, 0), (-1, 0), 0),
            ("LEFTPADDING", (0, 0), (-1, -1), 0),
            ("RIGHTPADDING", (0, 0), (-1, -1), 0),
            ("VALIGN", (0, 0), (-1, -1), "BOTTOM"),
        ]))
        _, self._sig_h = self._sig_table.wrap(self.width, 0xFFFFFF)
        self.height = self._intro_h + self.INTRO_GAP + checkbox_block_h + 8 * mm + self._sig_h

    def wrap(self, availWidth: float, availHeight: float) -> tuple[float, float]:
        return self.width, self.height

    def draw(self) -> None:
        c = self.canv
        c.saveState()
        c.setLineWidth(0.75)
        c.setStrokeColor(colors.black)
        self._intro_par.drawOn(c, 0, self.height - self._intro_h)
        y = self.height - self._intro_h - self.INTRO_GAP
        for par, par_h in zip(self._checkbox_pars, self._checkbox_heights):
            row_h = max(par_h, CHECKBOX_SIZE)
            c.rect(0, y - CHECKBOX_SIZE, CHECKBOX_SIZE, CHECKBOX_SIZE)
            par.drawOn(c, CHECKBOX_SIZE + 6, y - par_h)
            y -= row_h + self.CHECKBOX_ROW_GAP
        self._sig_table.drawOn(c, 0, 0)
        c.restoreState()


class BottomAnchor(Flowable):
    """Drückt `inner` an das untere Ende des verbleibenden Frame-Bereichs.

    Beansprucht immer die komplette auf der aktuellen Seite noch verfügbare
    Höhe (statt nur die eigene Inhaltshöhe), sodass `inner` beim Zeichnen
    unten in dieser Fläche landet — also direkt über der Fußzeile, egal wie
    viel Text/Tabellen vorher im Frame stehen. Passt `inner` nicht mehr auf
    die aktuelle Seite, wird der komplette Block (reportlab-Standardverhalten
    für nicht teilbare Flowables) auf die nächste Seite verschoben, wo er
    dann entsprechend am unteren Rand der neuen, leeren Fläche sitzt.
    """

    def __init__(self, inner: Flowable) -> None:
        super().__init__()
        self.inner = inner

    def wrap(self, availWidth: float, availHeight: float) -> tuple[float, float]:
        _, inner_h = self.inner.wrap(availWidth, availHeight)
        self._inner_h = inner_h
        # Passt inner nicht in den Rest der aktuellen Seite: mehr Höhe
        # melden, als zur Verfügung steht, damit reportlab den gesamten Block
        # auf die nächste Seite verschiebt statt ihn hier abzuschneiden.
        self.height = availHeight if inner_h <= availHeight else availHeight + inner_h
        return availWidth, self.height

    def draw(self) -> None:
        self.inner.drawOn(self.canv, 0, 0)


class _RecordPage(Flowable):
    """Nullgroßes Flowable: merkt sich beim Layout die Seitenzahl an dieser
    Stelle im Textfluss (in `holder[0]`) — Grundlage für `_PadToEvenPages`.

    `_ZEROSIZE` weist reportlabs Frame an, `wrap()` auch dann noch
    aufzurufen, wenn im Frame kein Platz mehr übrig ist (z.B. weil ein
    vorangehendes `BottomAnchor` bereits die komplette Resthöhe der Seite für
    sich beansprucht hat) — ohne das würde Frame._add() bei Resthöhe 0 sofort
    (ohne wrap()-Aufruf) auf die nächste Seite ausweichen, und die hier
    gemessene Seitenzahl wäre dann schon die der falschen, nächsten Seite."""

    _ZEROSIZE = 1

    def __init__(self, holder: list) -> None:
        super().__init__()
        self.width = self.height = 0
        self.holder = holder

    def wrap(self, availWidth: float, availHeight: float) -> tuple[float, float]:
        self.holder[0] = self.canv.getPageNumber()
        return 0, 0

    def draw(self) -> None:
        pass


class _PadToEvenPages(Flowable):
    """Erzwingt eine zusätzliche Leerseite, wenn das Fach seit dem
    zugehörigen `_RecordPage` eine ungerade Seitenzahl belegt (--duplex: jedes
    Fach soll mit gerader Seitenzahl enden, damit beim doppelseitigen Druck
    kein Fach auf der Rückseite eines anderen beginnt).

    Nutzt denselben Trick wie `BottomAnchor`: mehr Höhe melden, als auf der
    aktuellen Seite noch verfügbar ist, damit reportlab dieses (unsichtbare)
    Flowable auf eine neue, sonst leere Seite verschiebt.

    `blank_pages`, falls übergeben, wird um die Seitenzahl der so erzwungenen
    Leerseite ergänzt — `make_footer()` lässt deren Fußzeile dadurch weg,
    damit die Seite wirklich komplett leer bleibt (weißes Blatt, keine
    Kopf-/Fußzeile), wie es für doppelseitigen Druck erwartet wird.

    `_ZEROSIZE` (siehe `_RecordPage`) ist hier essenziell: ohne sie würde
    reportlab bei Resthöhe 0 (typischer Fall direkt nach `BottomAnchor`)
    automatisch und ungefragt eine neue Seite beginnen, bevor `wrap()`
    überhaupt zum Zuge kommt — die Entscheidung "muss gepolstert werden"
    fiele dann immer auf der bereits (fälschlich) neuen Seite, nie auf der
    tatsächlich letzten Inhaltsseite."""

    _ZEROSIZE = 1

    def __init__(self, start_page_holder: list, blank_pages: set[int] | None = None) -> None:
        super().__init__()
        self.width = self.height = 0
        self.start_page_holder = start_page_holder
        self.blank_pages = blank_pages

    def wrap(self, availWidth: float, availHeight: float) -> tuple[float, float]:
        current_page = self.canv.getPageNumber()
        pages_used = current_page - self.start_page_holder[0] + 1
        if pages_used % 2 == 1:
            if self.blank_pages is not None:
                self.blank_pages.add(current_page + 1)
            return availWidth, availHeight + 1
        return 0, 0

    def draw(self) -> None:
        pass


class _RecordEndPage(Flowable):
    """Nullgroßes Flowable (siehe `_RecordPage`): merkt sich am Fach-Ende die
    aktuelle Seitenzahl in `results[index]`, ohne — anders als
    `_PadToEvenPages` — je eine Leerseite zu erzwingen. Grundlage für den
    Messlauf hinter `--duplex-if-needed` (`measure_subject_pages()`): dort
    soll nur die *natürliche* Seitenzahl je Fach ermittelt werden.

    Braucht `_ZEROSIZE` aus demselben Grund wie `_PadToEvenPages`: ohne sie
    würde reportlab bei Resthöhe 0 (typisch direkt nach `BottomAnchor`) schon
    selbst auf eine neue Seite wechseln, bevor `wrap()` aufgerufen wird — die
    hier gemessene Seitenzahl wäre dann um eins zu hoch."""

    _ZEROSIZE = 1

    def __init__(self, start_page_holder: list, results: list[int], index: int) -> None:
        super().__init__()
        self.width = self.height = 0
        self.start_page_holder = start_page_holder
        self.results = results
        self.index = index

    def wrap(self, availWidth: float, availHeight: float) -> tuple[float, float]:
        current_page = self.canv.getPageNumber()
        self.results[self.index] = current_page - self.start_page_holder[0] + 1
        return 0, 0

    def draw(self) -> None:
        pass


def subject_story(
    subject: str, tables: dict[str, list[dict]], schoolyear_id: str, *,
    confirmation: bool = False, fkl_map: dict[str, str] | None = None,
    kollegium_map: dict[str, str] | None = None, duplex: bool = False,
    blank_pages: set[int] | None = None,
) -> list:
    confirm_value = None
    teacher_kuerzel = None
    if confirmation:
        fkl_name = find_fkl_name(fkl_map or {}, subject)
        teacher_kuerzel = find_kollegium_kuerzel(kollegium_map or {}, fkl_name)
        # Kopfzeile zeigt bevorzugt das Kürzel; ist das nicht auflösbar,
        # ersatzweise der Name, sonst ein Platzhalter (nie komplett leer).
        confirm_value = teacher_kuerzel or fkl_name or "–"

    start_page_holder: list = [None]
    story: list = []
    if duplex:
        story.append(_RecordPage(start_page_holder))
    story += [
        SubjectHeading(schoolyear_id, subject, confirm_value=confirm_value),
        Paragraph(
            f"Die folgenden Bücher können für das Fach {subject} über die Schule ausgeliehen werden. "
            "Bücher, die selbst anzuschaffen sind, werden gesondert in der zweiten Tabelle ausgewiesen.",
            INTRO_STYLE,
        ),
    ]

    story.append(Paragraph("Leihbare Bücher", SECTION_STYLE))
    if tables["leih"]:
        story.append(render_table(tables["leih"], with_fee=True))
    else:
        story.append(Paragraph("Keine leihbaren Bücher in diesem Fach.", EMPTY_STYLE))

    story.append(Paragraph("Selbst anzuschaffende Bücher", SECTION_STYLE))
    if tables["kauf"]:
        story.append(render_table(tables["kauf"], with_fee=False))
    else:
        story.append(Paragraph("Keine selbst anzuschaffenden Bücher in diesem Fach.", EMPTY_STYLE))

    if confirmation:
        story.append(
            BottomAnchor(ConfirmationBlock(subject, schoolyear_id, teacher_kuerzel=teacher_kuerzel))
        )

    if duplex:
        story.append(_PadToEvenPages(start_page_holder, blank_pages))

    return story


def measure_subject_pages(
    subjects: list[str], by_subject: dict[str, dict[str, list[dict]]], schoolyear_id: str, *,
    confirmation: bool, fkl_map: dict[str, str], kollegium_map: dict[str, str],
) -> list[int]:
    """Baut alle Fächer einmal probeweise in einen verworfenen Speicherpuffer
    (kein Datei-Output), jedes mit eigenem, frisch beginnendem PageTemplate —
    damit ist die ermittelte Seitenzahl je Fach unabhängig vom später
    tatsächlich gewählten `--mode` (combined/split) und entspricht exakt der
    natürlichen (ungepolsterten) Seitenzahl, die dieses Fach auch im
    Enddokument bräuchte.

    Grundlage für `--duplex-if-needed`: nur wenn dabei mindestens ein Fach
    bereits mehr als eine Seite braucht, lohnt sich das Anhängen von
    Leerseiten für den doppelseitigen Druck (siehe `main()`)."""
    doc = BaseDocTemplate(
        io.BytesIO(), pagesize=A4, leftMargin=LEFT_MARGIN, rightMargin=RIGHT_MARGIN,
        topMargin=PAGE_H - FRAME_TOP, bottomMargin=BOTTOM_MARGIN,
    )
    page_counts: list[int] = [0] * len(subjects)
    templates = []
    story: list = []
    for i, subject in enumerate(subjects):
        frame = Frame(
            LEFT_MARGIN, BOTTOM_MARGIN, CONTENT_WIDTH, FRAME_TOP - BOTTOM_MARGIN,
            leftPadding=0, rightPadding=0, topPadding=0, bottomPadding=0, id=f"measure-{i}",
        )
        template_id = f"measure-{i}"
        templates.append(PageTemplate(id=template_id, frames=[frame]))
        if i > 0:
            story.append(NextPageTemplate(template_id))
            story.append(PageBreak())
        start_page_holder: list = [None]
        story.append(_RecordPage(start_page_holder))
        story.extend(
            subject_story(
                subject, by_subject[subject], schoolyear_id,
                confirmation=confirmation, fkl_map=fkl_map, kollegium_map=kollegium_map, duplex=False,
            )
        )
        story.append(_RecordEndPage(start_page_holder, page_counts, i))
    doc.addPageTemplates(templates)
    doc.build(story)
    return page_counts


def footer_context(subject_or_label: str, schoolyear_id: str) -> str:
    """Rechtsbündiger Fußzeilentext, Wort für Wort wie im Original-PDF
    aufgebaut ("<Schule>, <Ort> – Bücherliste <Kontext> (Schuljahr 26/27)") —
    nur der Kontext ist hier das Fach bzw. "Fächer" statt "Jahrgang X"."""
    return (
        f"{SCHOOL_NAME}, {SCHOOL_CITY} – Bücherliste {subject_or_label} "
        f"(Schuljahr {short_schoolyear(schoolyear_id)})"
    )


def make_footer(center_text: str, *, page_offset_holder: list | None = None, blank_pages: set[int] | None = None):
    """onPage-Callback: Fußzeile wie in den offiziellen IServ-Bücherlisten.

    Gezeichnet wird hier noch nichts — die Fußzeile nennt "Seite n von N", und
    N steht erst fest, wenn das Dokument fertig gesetzt ist. Der Callback legt
    darum nur die Angaben der aktuellen Seite auf dem Canvas ab; `FooterCanvas`
    zeichnet sie am Ende des Builds nach (siehe dort).

    `page_offset_holder` ist ein einelementiger mutabler Fach-eigener
    Zwischenspeicher (`[None]`): wird beim ersten Aufruf mit der aktuellen
    (dokumentweiten) `doc.page` belegt, sodass die angezeigte Seitenzahl für
    dieses Fach wieder bei 1 beginnt, statt über das gesamte kombinierte PDF
    durchzuzählen (siehe write_combined_confirmation_pdf).

    `blank_pages` (--duplex) listet die von `_PadToEvenPages` erzwungenen
    Leerseiten dieses Dokuments — auf ihnen wird gar nichts gezeichnet, auch
    keine Fußzeile, damit sie beim doppelseitigen Druck wirklich komplett
    leer (weiß) bleiben."""
    generated = datetime.now().strftime("%d.%m.%Y")

    def _draw(c: Canvas, doc: BaseDocTemplate) -> None:
        if blank_pages is not None and doc.page in blank_pages:
            return
        if page_offset_holder is not None:
            if page_offset_holder[0] is None:
                page_offset_holder[0] = doc.page
            page_num = doc.page - page_offset_holder[0] + 1
        else:
            page_num = doc.page
        # Gruppenschlüssel für "von N": im kombinierten Bestätigungs-PDF zählt
        # jedes Fach eigenständig, und dort ist center_text je Fach verschieden.
        c._footer_spec = (center_text, page_num, generated)

    return _draw


def draw_footer(c: Canvas, page_num: int, total_pages: int, center_text: str, generated: str) -> None:
    """Zeichnet die zweizeilige Fußzeile an den aus dem Original übernommenen
    absoluten Grundlinien (FOOTER_PAGE_BASELINE / FOOTER_INFO_BASELINE)."""
    c.saveState()
    c.setFillColor(FOOTER_COLOR)
    c.setFont(FOOTER_FONT, FOOTER_SIZE)
    c.drawRightString(RIGHT_EDGE, FOOTER_PAGE_BASELINE, f"Seite {page_num} von {total_pages}")
    c.drawString(LEFT_MARGIN, FOOTER_INFO_BASELINE, f"Erstellt: {generated}")
    c.drawRightString(RIGHT_EDGE, FOOTER_INFO_BASELINE, center_text)
    c.restoreState()


class FooterCanvas(Canvas):
    """Canvas, der die Fußzeilen erst nach dem Setzen des Dokuments zeichnet.

    Nötig für "Seite n von N": die Gesamtseitenzahl ist während des Builds noch
    unbekannt. Jede Seite wird darum zunächst zwischengespeichert; `save()`
    stellt sie einzeln wieder her, zeichnet die Fußzeile (jetzt mit bekanntem
    N) und gibt sie erst dann aus.

    N ist die Seitenzahl der jeweiligen Fußzeilen-Gruppe (`center_text`), damit
    im kombinierten Bestätigungs-PDF jedes Fach für sich zählt. Seiten ohne
    hinterlegte Angaben (`--duplex`-Leerseiten) bleiben leer und zählen nicht
    mit."""

    def __init__(self, *args, **kwargs) -> None:
        super().__init__(*args, **kwargs)
        self._saved_states: list[dict] = []
        self._footer_specs: list[tuple | None] = []
        self._footer_spec: tuple | None = None

    def showPage(self) -> None:
        self._footer_specs.append(self._footer_spec)
        self._footer_spec = None
        self._saved_states.append(dict(self.__dict__))
        self._startPage()

    def save(self) -> None:
        # Lokale Kopien: __dict__.update() unten überschreibt die Attribute
        # mit dem Stand der jeweiligen Seite.
        states, specs = self._saved_states, self._footer_specs
        totals: dict[str, int] = {}
        for spec in specs:
            if spec is not None:
                totals[spec[0]] = max(totals.get(spec[0], 0), spec[1])
        for state, spec in zip(states, specs):
            self.__dict__.update(state)
            if spec is not None:
                center_text, page_num, generated = spec
                draw_footer(self, page_num, totals[center_text], center_text, generated)
            super().showPage()
        super().save()


def write_pdf(
    path: Path, story: list, title: str, footer_center: str, *, blank_pages: set[int] | None = None,
) -> None:
    # BaseDocTemplate statt SimpleDocTemplate: dessen Frame hat 6pt Innenrand,
    # der den Inhalt gegenüber den Seitenrändern verschiebt. Hier soll der Text
    # exakt auf LEFT_MARGIN/RIGHT_EDGE sitzen → Frame-Padding auf 0.
    doc = BaseDocTemplate(
        str(path),
        pagesize=A4,
        leftMargin=LEFT_MARGIN,
        rightMargin=RIGHT_MARGIN,
        topMargin=PAGE_H - FRAME_TOP,
        bottomMargin=BOTTOM_MARGIN,
        title=title,
    )
    frame = Frame(
        LEFT_MARGIN, BOTTOM_MARGIN, CONTENT_WIDTH, FRAME_TOP - BOTTOM_MARGIN,
        leftPadding=0, rightPadding=0, topPadding=0, bottomPadding=0, id="content",
    )
    doc.addPageTemplates(
        [PageTemplate(id="fach", frames=[frame], onPage=make_footer(footer_center, blank_pages=blank_pages))]
    )
    doc.build(story, canvasmaker=FooterCanvas)


def write_combined_confirmation_pdf(
    path: Path, subjects: list[str], by_subject: dict[str, dict[str, list[dict]]],
    schoolyear_id: str, *, fkl_map: dict[str, str], kollegium_map: dict[str, str], title: str,
    duplex: bool = False,
) -> None:
    """Wie write_pdf, aber ein eigenes PageTemplate je Fach: mit --confirmation
    soll die Seitenzahl je Fach wieder bei 1 beginnen und die Fußzeile das
    jeweilige Fach statt "Fächer" nennen (Bestätigungs-Vorlage ist pro Fach an
    eine reale Person adressiert, da wirkt eine durchlaufende Gesamt-
    Seitenzahl/"Fächer"-Fußzeile über das ganze Dokument fehl am Platz)."""
    doc = BaseDocTemplate(
        str(path), pagesize=A4, leftMargin=LEFT_MARGIN, rightMargin=RIGHT_MARGIN,
        topMargin=PAGE_H - FRAME_TOP, bottomMargin=BOTTOM_MARGIN, title=title,
    )
    # Eine gemeinsame blank_pages-Menge über das ganze (kombinierte) Dokument
    # hinweg — Seitenzahlen sind dokumentweit eindeutig, auch über die
    # per-Fach-PageTemplates hinweg, daher genügt ein einziges Set.
    blank_pages: set[int] = set()
    templates = []
    story: list = []
    for i, subject in enumerate(subjects):
        frame = Frame(
            LEFT_MARGIN, BOTTOM_MARGIN, CONTENT_WIDTH, FRAME_TOP - BOTTOM_MARGIN,
            leftPadding=0, rightPadding=0, topPadding=0, bottomPadding=0, id=f"content-{i}",
        )
        footer_center = footer_context(subject, schoolyear_id)
        template_id = f"fach-{i}"
        templates.append(
            PageTemplate(
                id=template_id, frames=[frame],
                onPage=make_footer(footer_center, page_offset_holder=[None], blank_pages=blank_pages),
            )
        )
        if i > 0:
            story.append(NextPageTemplate(template_id))
            story.append(PageBreak())
        story.extend(
            subject_story(
                subject, by_subject[subject], schoolyear_id,
                confirmation=True, fkl_map=fkl_map, kollegium_map=kollegium_map, duplex=duplex,
                blank_pages=blank_pages,
            )
        )
    doc.addPageTemplates(templates)
    doc.build(story, canvasmaker=FooterCanvas)


def sanitize_filename(name: str) -> str:
    return name.replace("/", "-")


def main() -> None:
    parser = argparse.ArgumentParser(description="Bücherlisten nach Fach als PDF.")
    parser.add_argument("--schoolyear", default=None, help='Schuljahr, z.B. "2026/2027" (Default: laufendes)')
    parser.add_argument(
        "--mode", choices=["combined", "split"], default="combined",
        help="combined = 1 PDF mit Seite pro Fach, split = 1 PDF je Fach (Default: combined)",
    )
    parser.add_argument(
        "--subjects", nargs="+", default=None, metavar="FACH",
        help="Nur diese Fächer aufnehmen (Default: alle vorhandenen Fächer)",
    )
    parser.add_argument(
        "--list-subjects", action="store_true",
        help="Nur die verfügbaren Fächer auflisten und beenden",
    )
    parser.add_argument("--output-dir", default=None, help="Zielordner (Default: dieser Skriptordner)")
    parser.add_argument(
        "--confirmation", action="store_true",
        help="Ankreuzfeld + Ort/Datum/Unterschrift Fachkonferenzleitung am Ende jeder Fach-Liste ergänzen",
    )
    duplex_group = parser.add_mutually_exclusive_group()
    duplex_group.add_argument(
        "--duplex", action="store_true",
        help="Für doppelseitigen Druck vorbereiten: jedes Fach bekommt bei Bedarf eine leere Endseite, "
             "damit seine Seitenzahl gerade ist",
    )
    duplex_group.add_argument(
        "--duplex-if-needed", action="store_true",
        help="Wie --duplex, aber nur wirksam, wenn mindestens ein Fach von Natur aus (ungepolstert) "
             "mehr als eine Seite braucht; sind alle Fächer ohnehin einseitig, bleibt die Ausgabe "
             "unverändert (keine Leerseiten)",
    )
    args = parser.parse_args()

    client = AusleiheClient(allow_writes=False)

    schoolyear_id = args.schoolyear or client.schoolyears.get_current()["id"]
    try:
        entries = collect_entries(client, schoolyear_id)
    except NotFoundError:
        print(f"Fehler: Schuljahr nicht gefunden: {schoolyear_id}", file=sys.stderr)
        sys.exit(1)

    by_subject = build_subject_tables(entries)
    subjects = sorted(by_subject, key=str.casefold)
    if not subjects:
        print(f"Keine Bücher für Schuljahr {schoolyear_id} gefunden.", file=sys.stderr)
        sys.exit(1)

    if args.list_subjects:
        for subject in subjects:
            print(subject)
        return

    if args.subjects:
        by_casefold = {s.casefold(): s for s in subjects}
        selected: list[str] = []
        unknown: list[str] = []
        for wanted in args.subjects:
            match = by_casefold.get(wanted.casefold())
            if match is None:
                unknown.append(wanted)
            elif match not in selected:
                selected.append(match)
        if unknown:
            print(f"Fehler: Unbekannte Fächer für Schuljahr {schoolyear_id}: {', '.join(unknown)}", file=sys.stderr)
            print(f"Verfügbare Fächer: {', '.join(subjects)}", file=sys.stderr)
            sys.exit(1)
        subjects = sorted(selected, key=str.casefold)

    fkl_map: dict[str, str] = {}
    kollegium_map: dict[str, str] = {}
    if args.confirmation:
        try:
            fkl_map = fetch_fkl_mapping()
        except Exception as exc:  # Netzwerk/Parsing-Fehler sollen die PDF-Erzeugung nicht abbrechen
            print(
                f"Warnung: Fachkonferenzleitungen konnten nicht von {FKL_URL} geladen werden ({exc}) "
                "— Kopfzeile/Unterschriftszeile bleiben ohne Namen/Kürzel.",
                file=sys.stderr,
            )
        try:
            kollegium_map = fetch_kollegium_kuerzel_mapping()
        except Exception as exc:
            print(
                f"Warnung: Lehrerkürzel konnten nicht von {KOLLEGIUM_URL} geladen werden ({exc}) "
                "— Kopfzeile/Unterschriftszeile zeigen ersatzweise den vollen Namen bzw. bleiben ohne Kürzel.",
                file=sys.stderr,
            )

    effective_duplex = args.duplex
    if args.duplex_if_needed:
        page_counts = measure_subject_pages(
            subjects, by_subject, schoolyear_id,
            confirmation=args.confirmation, fkl_map=fkl_map, kollegium_map=kollegium_map,
        )
        effective_duplex = any(count > 1 for count in page_counts)

    out_dir = Path(args.output_dir) if args.output_dir else _HERE
    out_dir.mkdir(parents=True, exist_ok=True)
    sy_label = sanitize_filename(schoolyear_id)

    if args.mode == "combined":
        out_path = out_dir / f"Bücherliste Fächer {sy_label}.pdf"
        if args.confirmation:
            # Bestätigungs-Vorlage ist pro Fach an eine reale Person adressiert
            # (Fachkonferenzleitung) — Seitenzahl zählt daher je Fach neu, und
            # die Fußzeile nennt das jeweilige Fach statt pauschal "Fächer".
            write_combined_confirmation_pdf(
                out_path, subjects, by_subject, schoolyear_id,
                fkl_map=fkl_map, kollegium_map=kollegium_map, title=f"Bücherliste Fächer {schoolyear_id}",
                duplex=effective_duplex,
            )
        else:
            # Eine gemeinsame blank_pages-Menge über das ganze kombinierte
            # Dokument (ein einziges PageTemplate für alle Fächer hier).
            blank_pages: set[int] = set()
            story: list = []
            for i, subject in enumerate(subjects):
                if i > 0:
                    story.append(PageBreak())
                story.extend(
                    subject_story(
                        subject, by_subject[subject], schoolyear_id,
                        confirmation=False, fkl_map=fkl_map, kollegium_map=kollegium_map, duplex=effective_duplex,
                        blank_pages=blank_pages,
                    )
                )
            footer_center = footer_context("Fächer", schoolyear_id)
            write_pdf(
                out_path, story, title=f"Bücherliste Fächer {schoolyear_id}", footer_center=footer_center,
                blank_pages=blank_pages,
            )
        print(f"PDF gespeichert: {out_path}")
    else:
        for subject in subjects:
            # Jedes Fach ist im split-Modus ein eigenes Dokument -> eigene
            # blank_pages-Menge.
            blank_pages = set()
            story = subject_story(
                subject, by_subject[subject], schoolyear_id,
                confirmation=args.confirmation, fkl_map=fkl_map, kollegium_map=kollegium_map, duplex=effective_duplex,
                blank_pages=blank_pages,
            )
            out_path = out_dir / f"Bücherliste {subject} {sy_label}.pdf"
            footer_center = footer_context(subject, schoolyear_id)
            write_pdf(
                out_path, story, title=f"Bücherliste {subject} {schoolyear_id}", footer_center=footer_center,
                blank_pages=blank_pages,
            )
            print(f"PDF gespeichert: {out_path}")


if __name__ == "__main__":
    main()
