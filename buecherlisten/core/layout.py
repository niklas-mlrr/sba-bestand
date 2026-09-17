"""Das PDF-Layout der Bücherlisten nach Fach (reportlab).

Unverändert aus ``generate_booklists.py`` verschoben (2026-09-17); die
Ausmessungen und Begründungen in den Kommentaren gelten weiter. Neu ist nur,
dass ``write_pdf`` und ``write_combined_confirmation_pdf`` auch in einen
Puffer schreiben können (``BinaryIO``), damit das Dashboard das PDF ohne
Umweg über eine Datei ausliefern kann.
"""
from __future__ import annotations

import io
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import BinaryIO

from reportlab.lib import colors
from reportlab.lib.enums import TA_JUSTIFY
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.pdfbase.pdfmetrics import stringWidth
from reportlab.pdfgen.canvas import Canvas
from reportlab.platypus import (
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

from buecherlisten.trg_web import find_kollegium_kuerzel, find_mapped_value

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

# Mindestlücke zwischen den nebeneinander stehenden Kopf-Blöcken ("Liste für",
# "Zu bestätigen durch", Rückgabe-Angaben, "gültig für"); wird sie unterschritten,
# schrumpft die große Schrift der mittleren Blöcke (siehe SubjectHeading).
MIN_HEADER_BLOCK_GAP = 6.0

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
# Schriftgröße wie die Fließtexte ("Die folgenden Bücher können …") in den
# offiziellen IServ-Jahrgangslisten: Helvetica 10 pt. Deren Zeilenabstand
# (9,25 pt) ist aber so eng, dass Unter- und Oberlängen sich fast berühren.
# 12,2 pt: vertikale Lücke zwischen dem tiefsten Punkt einer Zeile (g, -220)
# und dem höchsten der nächsten (Ä, 901) = 12,2 - 11,21 = ~0,9 pt, also etwa
# die mittlere horizontale Lücke zwischen zwei Buchstaben bei 10 pt
# (gerendert gemessen 2026-09-16: ~0,9 pt über a-z, A-Z, äöü, ÄÖÜ, ß).
CONFIRM_STYLE = ParagraphStyle(
    "Bestaetigung", parent=STYLES["Normal"], fontName=BODY_FONT, fontSize=10, leading=12.2,
    alignment=TA_JUSTIFY,
)
# Prüfauftrag im Bestätigungs-Lauf: gesetzt wie der Bestätigungstext
# (Schriftgröße, Zeilenabstand, Blocksatz), nach dem letzten Absatz aber mit
# dem vollen INTRO_STYLE-Abstand zur Tabelle. Absätze innerhalb dieser
# mehrteiligen Einleitung haben nur eine halbe Leerzeile Abstand.
INTRO_CONFIRM_STYLE = ParagraphStyle(
    "IntroBestaetigung", parent=CONFIRM_STYLE, spaceAfter=INTRO_STYLE.spaceAfter,
)
INTRO_PART_STYLE = ParagraphStyle(
    "IntroTeil", parent=INTRO_CONFIRM_STYLE, spaceAfter=CONFIRM_STYLE.leading / 2,
)
SIGNATURE_LABEL_STYLE = ParagraphStyle(
    "SignaturLabel", parent=STYLES["Normal"], fontName=BODY_FONT, fontSize=8, leading=10,
)
# Cap-Height (Höhe eines Großbuchstabens) des Bestätigungstexts — Helvetica
# laut AFM: 718/1000 der Schriftgröße.
CONFIRM_CAP_HEIGHT = CONFIRM_STYLE.fontSize * 0.718
# Kantenlänge des quadratischen Ankreuzkästchens vor dem Bestätigungstext:
# etwa halb so hoch wie der zweizeilige Text daneben, auf einen glatten Wert
# gerundet. Es sitzt mittig in diesen zwei Zeilen (siehe ConfirmationBlock).
CHECKBOX_SIZE = 9.0
# Freier Platz über den Unterschriftslinien für die handschriftliche
# Unterschrift: zwei Schreibzeilen im Zeilenabstand des Bestätigungstexts.
SIGNATURE_LINE_HEIGHT = 2 * CONFIRM_STYLE.leading

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


def _wrap_info(
    text: str, width: float, *, font: str = BODY_FONT, size: float = CELL_FONT_SIZE,
) -> tuple[int, float]:
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
    zusatz_zeilen: list[int] | None = None,
) -> tuple[float, float, float, float]:
    """Titel/Verlag-Breite so wählen, dass die Tabelle insgesamt am wenigsten
    Zeilen braucht (= am kleinsten ist); bei Gleichstand die Aufteilung, die
    zusätzlich am wenigsten Breite tatsächlich braucht.

    Rückgabe: (titel_budget, verlag_budget, titel_actual, verlag_actual) —
    *_actual ist die nach dem Umbruch bei diesem Budget tatsächlich breiteste
    Zeile (meist etwas schmaler als das Budget, siehe "eine Spalte nur
    maximal so lang wie der längste Inhalt einer Zeile").

    ``zusatz_zeilen`` sind die Zeilenzahlen einer dritten, bereits fest
    umbrochenen Spalte je Zeile (siehe :func:`_split_drei`).
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
        for i, (t, v) in enumerate(zip(titel_vals, verlag_vals)):
            lt, mt = _wrap_info(t, w)
            lv, mv = _wrap_info(v, verlag_w)
            total_lines += max(lt, lv) if zusatz_zeilen is None else max(lt, lv, zusatz_zeilen[i])
            max_titel_line = max(max_titel_line, mt)
            max_verlag_line = max(max_verlag_line, mv)
        key = (total_lines, max_titel_line + max_verlag_line)
        if best is None or key < best[0]:
            best = (key, w, max_titel_line, max_verlag_line)
        w += SPLIT_SEARCH_STEP

    assert best is not None
    _, titel_w, titel_actual, verlag_actual = best
    return titel_w, available - titel_w, titel_actual, verlag_actual


def _breiten_kandidaten(values: list[str]) -> list[float]:
    """Alle Breiten, bei denen sich der Umbruch einer Spalte ändern kann: die
    Breiten jeder zusammenhängenden Wortfolge eines Werts."""
    kandidaten = set()
    for v in values:
        woerter = v.split()
        for a in range(len(woerter)):
            for b in range(a + 1, len(woerter) + 1):
                kandidaten.add(stringWidth(" ".join(woerter[a:b]), BODY_FONT, CELL_FONT_SIZE) + WIDTH_EPSILON)
    return sorted(kandidaten)


def _split_drei(
    werte: tuple[list[str], list[str], list[str]], available: float,
) -> tuple[float, float, float]:
    """Wie :func:`_split_titel_verlag`, mit einer dritten umbrechenden Spalte.

    Die dritte (z. B. Fach in der Jahrgangsliste) hat kurze Werte und damit nur
    wenige sinnvolle Breiten; für jede wird der Rest wie bei zwei Spalten
    aufgeteilt. Gewählt wird die Aufteilung mit den wenigsten Zeilen, bei
    Gleichstand die schmalste. Rückgabe: tatsächliche Breiten in Spaltenfolge.
    """
    erste, zweite, dritte = werte
    # Untergrenze ist hier nur das breiteste unteilbare Wort, nicht MIN_WRAP_COL:
    # die Kandidaten sind gemessene Umbruchbreiten, und die kurzen Werte dieser
    # Spalte liegen regelmäßig darunter ("Mathematik" ist schmaler als 20 mm).
    minimum = _max_word_width(dritte)
    natuerlich = _raw_width("", dritte)
    best: tuple[tuple[int, float], tuple[float, float, float]] | None = None
    for kandidat in [k for k in _breiten_kandidaten(dritte) if minimum <= k <= natuerlich] or [natuerlich]:
        infos = [_wrap_info(v, kandidat) for v in dritte]
        dritte_w = max((w for _, w in infos), default=0.0)
        zeilen = [n for n, _ in infos]
        _, _, erste_w, zweite_w = _split_titel_verlag(erste, zweite, available - dritte_w, zeilen)
        gesamt = sum(
            max(_wrap_info(a, erste_w)[0], _wrap_info(b, zweite_w)[0], n)
            for a, b, n in zip(erste, zweite, zeilen)
        )
        key = (gesamt, erste_w + zweite_w + dritte_w)
        if best is None or key < best[0]:
            best = (key, (erste_w, zweite_w, dritte_w))
    assert best is not None
    return best[1]


@dataclass(frozen=True)
class Spalte:
    """Eine Tabellenspalte. ``umbruch``: darf umbrechen (genau zwei je Tabelle).

    ``kurz`` ist der Kopf, auf den abgekürzt wird, wenn nicht alles einzeilig
    passt; ``kurz_nur_wenn_kopf_breiter`` kürzt nur, wenn der Kopf selbst
    breiter ist als der breiteste Wert (sonst brächte es nichts).
    """

    key: str
    kopf: str
    umbruch: bool = False
    kurz: str | None = None
    kurz_nur_wenn_kopf_breiter: bool = False
    # Dritte umbrechende Spalte mit kurzen Werten (Fach in der Jahrgangsliste):
    # bricht nur um, wenn die Tabelle dadurch weniger Zeilen braucht.
    nachrangig: bool = False
    schrift: float = CELL_FONT_SIZE
    rechts: bool = False


KLASSE = Spalte("klasse", "Klasse", kurz="Kl.", kurz_nur_wenn_kopf_breiter=True)
TITEL = Spalte("titel", "Titel", umbruch=True)
VERLAG = Spalte("verlag", "Verlag", umbruch=True)
ISBN = Spalte("isbn", "ISBN", schrift=ISBN_FONT_SIZE)
NEUPREIS = Spalte("neupreis", "Neupreis", rechts=True)
LEIHGEBUEHR = Spalte("leihgebuehr", "Leihgebühr", kurz="Leihgeb.", rechts=True)


FACH = Spalte("fach", "Fach", umbruch=True)
FACH_NACHRANGIG = Spalte("fach", "Fach", umbruch=True, nachrangig=True)


@dataclass(frozen=True)
class Listenart:
    """Was eine Fach-, Verlags- und Jahrgangsliste unterscheidet: Texte und Spalten.

    Kopf, Schrift, Abstände und Tabellenbild sind für alle gleich; die Texte
    enthalten ``{name}`` für das Fach, den Verlag oder "Jahrgang N".
    """

    einleitung: str
    leer_leih: str
    leer_kauf: str
    spalten: tuple[Spalte, ...]


_ZWEITE_TABELLE = "Bücher, die selbst anzuschaffen sind, werden gesondert in der zweiten Tabelle ausgewiesen."
FACH_LISTE = Listenart(
    "Die folgenden Bücher können für das Fach {name} über die Schule ausgeliehen werden. "
    + _ZWEITE_TABELLE,
    "Keine leihbaren Bücher in diesem Fach.",
    "Keine selbst anzuschaffenden Bücher in diesem Fach.",
    (KLASSE, TITEL, VERLAG, ISBN, NEUPREIS, LEIHGEBUEHR),
)
# Verlag: der Verlag steht im Kopf, dafür das Fach - das neben dem Titel
# umbrechen darf, weil fächerübergreifende Bücher mehrere Fächer tragen.
VERLAG_LISTE = Listenart(
    "Die folgenden Bücher des Verlags {name} können über die Schule ausgeliehen werden. "
    + _ZWEITE_TABELLE,
    "Keine leihbaren Bücher dieses Verlags.",
    "Keine selbst anzuschaffenden Bücher dieses Verlags.",
    (TITEL, FACH, KLASSE, ISBN, NEUPREIS, LEIHGEBUEHR),
)
# Jahrgang: Spalten wie in der IServ-Druckversion. Titel und Verlag brechen
# um; das Fach nur, wenn es Zeilen spart ("Werte und Normen" hielt sonst die
# ganze Spalte breit und brach dafür die Hälfte der Titel um).
JAHRGANG_LISTE = Listenart(
    "Die folgenden Bücher können für den {name} über die Schule ausgeliehen werden. "
    + _ZWEITE_TABELLE,
    "Keine leihbaren Bücher in diesem Jahrgang.",
    "Keine selbst anzuschaffenden Bücher in diesem Jahrgang.",
    (TITEL, FACH_NACHRANGIG, VERLAG, ISBN, NEUPREIS, LEIHGEBUEHR),
)


def render_table(rows: list[dict], *, with_fee: bool, spalten: tuple[Spalte, ...] | None = None) -> Table:
    """Tabelle mit ``spalten`` (Default: die der Fachliste); ohne ``with_fee``
    entfällt die Leihgebühr."""
    if spalten is None:
        spalten = (KLASSE, TITEL, VERLAG, ISBN, NEUPREIS, LEIHGEBUEHR)
    spalten = tuple(sp for sp in spalten if with_fee or sp.key != "leihgebuehr")
    n_cols = len(spalten)
    n_gaps = n_cols - 1
    umbruch = [i for i, sp in enumerate(spalten) if sp.umbruch]
    assert len(umbruch) in (2, 3), "zwei oder drei umbrechende Spalten"

    # Für jeden der n_gaps Zwischenräume wird vorab MIN_GAP reserviert, bevor
    # Spaltenbreiten überhaupt berechnet werden — sonst kann der nach Schritt 3
    # rechnerisch übrige Platz pro Lücke gegen 0 gehen, wenn der Inhalt die
    # Tabellenbreite fast ausfüllt (beobachtet 2026-08-18: "12Duden..." ohne
    # sichtbaren Abstand). effective_width ist das Budget, mit dem Schritt 1/2
    # rechnen — der danach ermittelte Gesamtinhalt passt dadurch garantiert
    # mit mindestens MIN_GAP Luft pro Lücke in CONTENT_WIDTH.
    effective_width = CONTENT_WIDTH - MIN_GAP * n_gaps

    werte = [[r[sp.key] for r in rows] for sp in spalten]
    koepfe = [sp.kopf for sp in spalten]

    # 1) Passt alles einzeilig (jede Spalte auf ihre natürliche Breite,
    #    umbrechende Spalten inklusive) in die Tabellenbreite? Dann muss nichts
    #    umbrechen und nichts abgekürzt werden.
    breiten = [_raw_width(sp.kopf, werte[i], value_size=sp.schrift) for i, sp in enumerate(spalten)]
    natural_total = sum(breiten)

    if natural_total > effective_width:
        # 2) Reicht nicht — zuerst Platz durch Abkürzen der Kopfzeilen
        #    zurückgewinnen: "Leihgebühr" wird immer zu "Leihgeb."; "Klasse"
        #    nur, wenn der Spaltentitel selbst breiter ist als der breiteste
        #    Klassen-Wert (sonst würde die Abkürzung nichts bringen).
        gekuerzt: list[int] = []
        for i, sp in enumerate(spalten):
            if sp.kurz is None:
                continue
            if sp.kurz_nur_wenn_kopf_breiter:
                data_w = max((stringWidth(v, BODY_FONT, sp.schrift) for v in werte[i]), default=0.0)
                if stringWidth(sp.kopf, HEADER_FONT, CELL_FONT_SIZE) <= data_w:
                    continue
            koepfe[i] = sp.kurz
            gekuerzt.append(i)
            breiten[i] = _raw_width(sp.kurz, werte[i], value_size=sp.schrift)

        fixed_w = sum(w for i, w in enumerate(breiten) if i not in umbruch)
        available_tv = effective_width - fixed_w
        if len(umbruch) == 2:
            erste_u, zweite_u = umbruch
            _, _, breiten[erste_u], breiten[zweite_u] = _split_titel_verlag(
                werte[erste_u], werte[zweite_u], available_tv,
            )
        else:
            haupt = [i for i in umbruch if not spalten[i].nachrangig]
            (dritte_u,) = [i for i in umbruch if spalten[i].nachrangig]
            breiten[haupt[0]], breiten[haupt[1]], breiten[dritte_u] = _split_drei(
                (werte[haupt[0]], werte[haupt[1]], werte[dritte_u]), available_tv,
            )

        # 2b) Abkürzungen zurücknehmen, wenn nach der Aufteilung der
        #     umbrechenden Spalten wieder Platz dafür ist — in Spalten-
        #     reihenfolge ("Klasse" vor "Leihgebühr"). Erst danach wird
        #     verteilt, damit die dadurch länger gewordenen Spalten in der
        #     Gleichverteilung berücksichtigt sind.
        total_now = sum(breiten)
        for i in gekuerzt:
            sp = spalten[i]
            voll_w = _raw_width(sp.kopf, werte[i], value_size=sp.schrift)
            if total_now + (voll_w - breiten[i]) <= effective_width:
                total_now += voll_w - breiten[i]
                breiten[i] = voll_w
                koepfe[i] = sp.kopf

    # 3) Restplatz gleichmäßig auf alle Spaltenzwischenräume verteilen —
    #    jede Spalte bleibt exakt so breit wie ihr tatsächlich benötigter
    #    Inhalt (keine Spalte länger als der längste Inhalt einer Zeile).
    #    Dank effective_width in Schritt 1/2 ist gap hier immer >= MIN_GAP.
    total_content = sum(breiten)
    gap = max(CONTENT_WIDTH - total_content, 0.0) / n_gaps if n_gaps else 0.0
    half_gap = gap / 2

    # colWidths braucht die VOLLE Spaltenbreite inkl. ihres Anteils an der
    # Lücke (reportlab zieht das TableStyle-Padding von colWidth ab, um die
    # Textfläche zu bestimmen — reine Inhaltsbreite hier würde bei jeder
    # Spalte außer der ersten/letzten ins Negative laufen).
    col_widths = [
        w + (0.0 if i == 0 else half_gap) + (0.0 if i == n_cols - 1 else half_gap)
        for i, w in enumerate(breiten)
    ]

    data: list[list] = [list(koepfe)]
    for r in rows:
        data.append([
            Paragraph(r[sp.key], CELL_STYLE) if sp.umbruch else r[sp.key]
            for sp in spalten
        ])

    last_row = len(data) - 1
    # Tabellenbild wie im Original: keine Füllfarben, kein Gitternetz, 1.0pt
    # schwarze Linie unter der Kopfzeile, 1.0pt graue Trennlinie unter jeder
    # Datenzeile. Klasse/Titel/Verlag/ISBN linksbündig, Neupreis/Leihgebühr
    # rechtsbündig (nur Datenzeilen — die Kopfzeile bleibt linksbündig).
    style = [
        ("FONTNAME", (0, 0), (-1, 0), HEADER_FONT),
        ("FONTNAME", (0, 1), (-1, -1), BODY_FONT),
        ("FONTSIZE", (0, 0), (-1, -1), CELL_FONT_SIZE),
        ("ALIGN", (0, 0), (-1, 0), "LEFT"),
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
        # Gleicher Abstand zwischen allen Spalten: jede Innenkante bekommt die
        # halbe Lücke, außen (erste/letzte Spalte) bleibt 0 — Tabelle fluchtet
        # weiterhin links mit "Liste für"/Überschrift, rechts mit "gültig für".
        ("LEFTPADDING", (0, 0), (-1, -1), half_gap),
        ("RIGHTPADDING", (0, 0), (-1, -1), half_gap),
        ("LEFTPADDING", (0, 0), (0, -1), 0),
        ("RIGHTPADDING", (-1, 0), (-1, -1), 0),
    ]
    for i, sp in enumerate(spalten):
        if sp.schrift != CELL_FONT_SIZE:
            style.append(("FONTSIZE", (i, 1), (i, -1), sp.schrift))
        if sp.rechts:
            style.append(("ALIGN", (i, 1), (i, -1), "RIGHT"))
        if sp.key == "isbn":
            # Siehe ISBN_VSHIFT oben.
            style += [
                ("TOPPADDING", (i, 1), (i, -1), CELL_VPAD - VPAD_SHIFT + ISBN_VSHIFT),
                ("BOTTOMPADDING", (i, 1), (i, -1), CELL_VPAD + VPAD_SHIFT - ISBN_VSHIFT),
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

    def __init__(
        self, schoolyear_name: str, subject: str, *, confirm_value: str | None = None,
        return_by: str | None = None, return_to: str | None = None,
    ) -> None:
        super().__init__()
        self.schoolyear_name = schoolyear_name
        self.subject = subject
        self.confirm_value = confirm_value
        self.return_by = return_by
        self.return_to = return_to
        self.width = CONTENT_WIDTH
        self.height = HEADING_BLOCK_HEIGHT

    def _return_columns(self) -> list[tuple[str, str]]:
        """Spalten des Rückgabe-Blocks als (Label, Wert)-Paare.

        "Rückgabe" gehört immer zum ersten vorhandenen Wert: mit Datum heißen
        die Labels "Rückgabe bis" / "an", fehlt das Datum, steht über dem
        Kürzel "Rückgabe an". Fehlt ein Wert, entfällt seine Spalte ganz."""
        columns: list[tuple[str, str]] = []
        if self.return_by:
            columns.append(("Rückgabe bis", self.return_by))
        if self.return_to:
            columns.append(("an" if self.return_by else "Rückgabe an", self.return_to))
        return columns

    def _draw_middle_blocks(self, c: Canvas, dx: float, dy: float) -> None:
        """"Zu bestätigen durch" und Rückgabe-Angaben zwischen den Rand-Spalten.

        Beide stehen auf denselben beiden Grundlinien wie "Liste für"/
        <Schuljahr> links und "gültig für"/<Fach> rechts: Label in der kleinen,
        Wert in der großen Kopf-Schrift, Label jeweils mittig über seinem Wert.
        Die vier Teile (links, Bestätigung, Rückgabe, rechts) bekommen
        untereinander denselben Abstand — der nach allen Blockbreiten
        verbleibende Platz wird gleichmäßig auf die drei Lücken verteilt. Wird
        es eng (langer Fachname, langes Datum), schrumpft die große Schrift der
        beiden mittleren Blöcke, bis die Lücken wieder mindestens
        MIN_HEADER_BLOCK_GAP breit sind.

        Der Rückgabe-Block selbst besteht aus bis zu zwei Spalten (Datum,
        Kürzel), die untereinander nur zwei Leerzeichen der großen Schrift
        Abstand halten — sie gehören sichtbar zusammen.

        Ohne Rückgabe-Angaben bleibt die Bestätigung wie bisher mittig auf der
        Seite (unveränderte Ausgabe für Läufe ohne --return-by/--return-to)."""
        blocks = [[("Zu bestätigen durch", self.confirm_value)]]
        return_columns = self._return_columns()
        if not return_columns:
            self._draw_centered_confirm(c, dx, dy)
            return
        blocks.append(return_columns)

        # Die Rand-Blöcke sind so breit wie ihr breiterer Text (Label oder
        # Wert) — sie sind links- bzw. rechtsbündig gesetzt und rücken nicht.
        left_w = max(
            c.stringWidth("Liste für", HEADER_LABEL_FONT, HEADER_LABEL_SIZE),
            c.stringWidth(self.schoolyear_name, HEADER_VALUE_FONT, HEADER_VALUE_SIZE),
        )
        right_w = max(
            c.stringWidth("gültig für", HEADER_LABEL_FONT, HEADER_LABEL_SIZE),
            c.stringWidth(self.subject, HEADER_VALUE_FONT, HEADER_VALUE_SIZE),
        )

        value_fs = HEADER_VALUE_SIZE
        while True:
            inner_gap = c.stringWidth("  ", HEADER_VALUE_FONT, value_fs)
            col_widths = [
                [
                    max(
                        c.stringWidth(value, HEADER_VALUE_FONT, value_fs),
                        c.stringWidth(label, HEADER_LABEL_FONT, HEADER_LABEL_SIZE),
                    )
                    for label, value in block
                ]
                for block in blocks
            ]
            block_widths = [
                sum(widths) + inner_gap * (len(widths) - 1) for widths in col_widths
            ]
            gap = (
                CONTENT_WIDTH - left_w - right_w - sum(block_widths)
            ) / (len(blocks) + 1)
            if gap >= MIN_HEADER_BLOCK_GAP or value_fs <= 7.0:
                break
            value_fs -= 0.25

        x = dx + LEFT_MARGIN + left_w
        for block, widths in zip(blocks, col_widths):
            x += gap
            for (label, value), col_w in zip(block, widths):
                center = x + col_w / 2
                c.setFont(HEADER_LABEL_FONT, HEADER_LABEL_SIZE)
                c.drawCentredString(center, dy + HEADER_LABEL_BASELINE, label)
                c.setFont(HEADER_VALUE_FONT, value_fs)
                c.drawCentredString(center, dy + HEADER_VALUE_BASELINE, value)
                x += col_w + inner_gap
            x -= inner_gap

    def _draw_centered_confirm(self, c: Canvas, dx: float, dy: float) -> None:
        """Nur "Zu bestätigen durch" + Kürzel, mittig auf der Seite.

        Wert-Zeile: bei langem Fachnamen + langem Ersatzwert (Fallback: Name
        statt Kürzel) notfalls verkleinern, damit nichts mit den beiden
        Rand-Werten kollidiert (analog Fußzeile)."""
        center_x = dx + FOOTER_CENTER_X

        # Label-Zeile ("Liste für"/"gültig für"-Höhe): kurzer, fester
        # Text — bei den üblichen Fach-/Jahr-Kürzeln links/rechts nie eng.
        c.setFont(HEADER_LABEL_FONT, HEADER_LABEL_SIZE)
        c.drawCentredString(center_x, dy + HEADER_LABEL_BASELINE, "Zu bestätigen durch")

        left_val_w = c.stringWidth(
            self.schoolyear_name, HEADER_VALUE_FONT, HEADER_VALUE_SIZE,
        )
        right_val_w = c.stringWidth(self.subject, HEADER_VALUE_FONT, HEADER_VALUE_SIZE)
        left_end = dx + LEFT_MARGIN + left_val_w
        right_start = dx + RIGHT_EDGE - right_val_w
        max_half_width = max(min(center_x - left_end, right_start - center_x) - 6, 10)

        center_fs = HEADER_VALUE_SIZE
        while (
            center_fs > 7.0
            and c.stringWidth(self.confirm_value, HEADER_VALUE_FONT, center_fs) / 2 > max_half_width
        ):
            center_fs -= 0.25

        c.setFont(HEADER_VALUE_FONT, center_fs)
        c.drawCentredString(center_x, dy + HEADER_VALUE_BASELINE, self.confirm_value)

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
            self.schoolyear_name,
        )
        c.drawRightString(dx + RIGHT_EDGE, dy + HEADER_VALUE_BASELINE, self.subject)

        if self.confirm_value:
            self._draw_middle_blocks(c, dx, dy)

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

    # Abstände zwischen einleitendem Satz, den beiden Ankreuzzeilen und dem
    # Unterschriftsfeld: überall der Absatzabstand des Prüfauftrags (halbe
    # Leerzeile, siehe INTRO_PART_STYLE).
    INTRO_GAP = CHECKBOX_ROW_GAP = SIGNATURE_GAP = INTRO_PART_STYLE.spaceAfter

    def __init__(
        self, subject: str, schoolyear_name: str, book_count: int, grade_count: int | None = None, *,
        teacher_kuerzel: str | None = None, confirm_page: int | None = None,
    ) -> None:
        super().__init__()
        self.width = CONTENT_WIDTH
        signature_label = f"Unterschrift Fachkonferenzleitung {subject}"
        if teacher_kuerzel:
            signature_label += f" ({teacher_kuerzel})"
        # Verweis auf die Tabelle(n) weiter oben: "oben", solange das Fach nur
        # 1 Seite braucht; "umseitig", wenn die Bestätigung auf Seite 2 des
        # Fachs landet (die Tabelle also komplett auf der vorigen Seite
        # steht); ab Seite 3 "auf den vorliegenden Seiten" (Plural, da
        # mehrere Seiten vorausgehen können).
        if confirm_page is None or confirm_page <= 1:
            location = "oben"
        elif confirm_page == 2:
            location = "umseitig"
        else:
            location = "auf den vorliegenden Seiten"
        if book_count == 1:
            # Einzelnes Buch: Singular ("das ... Buch", "seine", "dessen")
            # statt Plural ("die ... Bücher", "ihre", "deren"). Die
            # Klassenstufe(n) richten sich dabei nach den tatsächlich diesem
            # einen Buch zugeordneten Klassenstufen, nicht nach der Bücherzahl.
            klassenstufe_word = "Klassenstufe" if grade_count == 1 else "Klassenstufen"
            intro_text = (
                f"Hiermit bestätige ich im Namen der Fachschaft {subject}, dass die {location} "
                f"aufgeführte <b>Bücherliste {subject}</b>, das heißt das durch seine "
                f"<b>ISBN</b> beschriebene Buch und dessen zugeordnete <b>{klassenstufe_word}</b>, "
                f"für das <b>{schoolyear_name}</b> ..."
            )
        else:
            intro_text = (
                f"Hiermit bestätige ich im Namen der Fachschaft {subject}, dass die {location} "
                f"aufgeführte <b>Bücherliste {subject}</b>, das heißt die durch ihre <b>ISBN</b> "
                f"beschriebenen Bücher und deren zugeordnete <b>Klassenstufen</b>, für das "
                f"<b>{schoolyear_name}</b> ..."
            )
        self._intro_par = Paragraph(intro_text, CONFIRM_STYLE)
        _, self._intro_h = self._intro_par.wrap(self.width, 0xFFFFFF)

        text_width = self.width - CHECKBOX_SIZE - 6
        self._checkbox_pars = [
            Paragraph(
                "... <b>nicht korrekt</b> ist und um die <b>handschriftlichen Anmerkungen</b> "
                "(Durchstreichungen, Eintragungen neuer Bücher, Klassen, ...) "
                "verändert werden muss.",
                CONFIRM_STYLE,
            ),
            Paragraph(
                "... <b>korrekt</b> ist, <b>an</b> die <b>Schüler übermittelt</b> werden kann und eine "
                "<b>nachträgliche Änderung</b> unter Umständen <b>nicht</b> mehr <b>gestattet</b> werden "
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
        self.height = self._intro_h + self.INTRO_GAP + checkbox_block_h + self.SIGNATURE_GAP + self._sig_h

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
            # Mittig in dem Band, das die beiden Textzeilen aufspannen: von der
            # Versalhöhe der ersten Zeile bis zur Grundlinie der zweiten.
            # reportlab setzt die erste Grundlinie fontSize unter die
            # Paragraph-Oberkante (rl_config paraFontSizeHeightOffset=1, siehe
            # Paragraph.drawPara).
            first_baseline = y - CONFIRM_STYLE.fontSize
            band_top = first_baseline + CONFIRM_CAP_HEIGHT
            band_bottom = first_baseline - CONFIRM_STYLE.leading
            box_bottom = (band_top + band_bottom - CHECKBOX_SIZE) / 2
            c.rect(0, box_bottom, CHECKBOX_SIZE, CHECKBOX_SIZE)
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
    subject: str, tables: dict[str, list[dict]], schoolyear_name: str, *,
    confirmation: bool = False, fkl_map: dict[str, str] | None = None,
    kollegium_map: dict[str, str] | None = None, duplex: bool = False,
    blank_pages: set[int] | None = None, confirmation_page_count: int | None = None,
    return_by: str | None = None, return_to: str | None = None, art: Listenart = FACH_LISTE,
) -> list:
    confirm_value = None
    teacher_kuerzel = None
    if confirmation:
        fkl_name = find_mapped_value(fkl_map or {}, subject)
        teacher_kuerzel = find_kollegium_kuerzel(kollegium_map or {}, fkl_name)
        # Kopfzeile zeigt bevorzugt das Kürzel; ist das nicht auflösbar,
        # ersatzweise der Name, sonst ein Platzhalter (nie komplett leer).
        confirm_value = teacher_kuerzel or fkl_name or "–"

    start_page_holder: list = [None]
    story: list = []
    if duplex:
        story.append(_RecordPage(start_page_holder))
    story += [
        SubjectHeading(
            schoolyear_name, subject, confirm_value=confirm_value,
            return_by=return_by if confirmation else None,
            return_to=return_to if confirmation else None,
        ),
        Paragraph(art.einleitung.format(name=subject), INTRO_STYLE),
    ]
    book_count = len(tables["leih"]) + len(tables["kauf"])
    # Bei genau einem Buch entscheidet die Anzahl seiner Klassenstufen
    # (Spalte "klasse", z.B. "5" oder "5, 6") über Singular/Plural.
    grade_count = None
    if book_count == 1:
        only_row = (tables["leih"] or tables["kauf"])[0]
        grade_count = only_row["klasse"].count(",") + 1
    if confirmation:
        # Bestätigungs-Lauf: statt des Einleitungssatzes ein Prüfauftrag an die
        # Fachkonferenzleitung, dessen letzter Absatz auf die beiden
        # Ankreuzfelder im ConfirmationBlock abgestimmt ist. Singular/Plural
        # wie im ConfirmationBlock.
        if book_count == 1:
            klassenstufe_word = "Klassenstufe" if grade_count == 1 else "Klassenstufen"
            books_text = f"das durch seine ISBN beschriebene Buch und dessen zugeordnete {klassenstufe_word}"
        else:
            books_text = "die durch ihre ISBN beschriebenen Bücher und deren zugeordnete Klassenstufen"
        story[-1:] = [
            Paragraph(
                f"Bitte prüfen Sie als Fachkonferenzleitung {subject} im Namen der "
                f"Fachschaft {subject} die folgende Bücherliste {subject}, das heißt "
                f"{books_text}, für das {schoolyear_name}.",
                INTRO_PART_STYLE,
            ),
            Paragraph(
                "Da ein Buch durch seine ISBN eindeutig beschrieben wird, kann der Titel vom "
                "Originaltitel des Buches abweichen und gegebenenfalls auch im nachhinein noch einmal "
                "verändert werden, um z. B. Bücher für verschiedene Leistungsniveaus leichter "
                "auseinanderhalten zu können. Die Leihgebühr dient nur der Information. Sie ergibt "
                "sich aus dem Neupreis des Buches und der Anzahl der Jahrgänge, für die es "
                "vorgesehen ist.",
                INTRO_PART_STYLE,
            ),
            Paragraph(
                "Bitte korrigieren Sie, falls ISBN falsch sind, falsche Klassenstufen zugeordnet "
                "sind, neue Bücher angeschafft oder alte ausgemustert werden sollen, dies "
                "handschriftlich und bestätigen Sie am Ende die Gültigkeit dieser Änderungen, oder "
                "notieren Sie, falls gewünscht, handschriftlich Änderungsvorschläge für Buchtitel "
                "und bestätigen Sie am Ende die Gültigkeit der Bücherliste. Beachten Sie bitte, dass "
                "Änderungsvorschläge für Buchtitel gegebenenfalls nicht oder auch nur abgeändert "
                "übernommen werden.",
                INTRO_CONFIRM_STYLE,
            ),
        ]

    story.append(Paragraph("Leihbare Bücher", SECTION_STYLE))
    if tables["leih"]:
        story.append(render_table(tables["leih"], with_fee=True, spalten=art.spalten))
    else:
        story.append(Paragraph(art.leer_leih, EMPTY_STYLE))

    story.append(Paragraph("Selbst anzuschaffende Bücher", SECTION_STYLE))
    if tables["kauf"]:
        story.append(render_table(tables["kauf"], with_fee=False, spalten=art.spalten))
    else:
        story.append(Paragraph(art.leer_kauf, EMPTY_STYLE))

    if confirmation:
        story.append(
            BottomAnchor(
                ConfirmationBlock(
                    subject, schoolyear_name, book_count, grade_count,
                    teacher_kuerzel=teacher_kuerzel, confirm_page=confirmation_page_count,
                )
            )
        )

    if duplex:
        story.append(_PadToEvenPages(start_page_holder, blank_pages))

    return story


def measure_subject_pages(
    subjects: list[str], by_subject: dict[str, dict[str, list[dict]]], schoolyear_name: str, *,
    confirmation: bool, fkl_map: dict[str, str], kollegium_map: dict[str, str],
    return_by: str | None = None, return_to: str | None = None, art: Listenart = FACH_LISTE,
) -> list[int]:
    """Baut alle Fächer einmal probeweise in einen verworfenen Speicherpuffer
    (kein Datei-Output), jedes mit eigenem, frisch beginnendem PageTemplate —
    damit ist die ermittelte Seitenzahl je Fach unabhängig vom später
    tatsächlich gewählten `--mode` (split/alphabet/aufgabenfeld) und entspricht exakt der
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
                subject, by_subject[subject], schoolyear_name,
                confirmation=confirmation, fkl_map=fkl_map, kollegium_map=kollegium_map, duplex=False,
                return_by=return_by, return_to=return_to, art=art,
            )
        )
        story.append(_RecordEndPage(start_page_holder, page_counts, i))
    doc.addPageTemplates(templates)
    doc.build(story)
    return page_counts


def footer_context(subject_or_label: str, schoolyear_name: str) -> str:
    """Rechtsbündiger Fußzeilentext, Wort für Wort wie im Original-PDF
    aufgebaut ("<Schule>, <Ort> – Bücherliste <Kontext> (Schuljahr 26/27)") —
    nur der Kontext ist hier das Fach bzw. "Fächer" statt "Jahrgang X"."""
    return (
        f"{SCHOOL_NAME}, {SCHOOL_CITY} – Bücherliste {subject_or_label} "
        f"({schoolyear_name})"
    )


def make_footer(
    center_text: str, *, page_offset_holder: list | None = None, blank_pages: set[int] | None = None,
):
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


def _ziel(path: str | Path | BinaryIO) -> str | BinaryIO:
    return path if hasattr(path, "write") else str(path)  # type: ignore[return-value]


def write_pdf(
    path: str | Path | BinaryIO, story: list, title: str, footer_center: str, *,
    blank_pages: set[int] | None = None,
) -> None:
    # BaseDocTemplate statt SimpleDocTemplate: dessen Frame hat 6pt Innenrand,
    # der den Inhalt gegenüber den Seitenrändern verschiebt. Hier soll der Text
    # exakt auf LEFT_MARGIN/RIGHT_EDGE sitzen → Frame-Padding auf 0.
    doc = BaseDocTemplate(
        _ziel(path),
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
    path: str | Path | BinaryIO, subjects: list[str], by_subject: dict[str, dict[str, list[dict]]],
    schoolyear_name: str, *, fkl_map: dict[str, str], kollegium_map: dict[str, str], title: str,
    duplex: bool = False, page_counts: list[int] | None = None,
    return_by: str | None = None, return_to: str | None = None,
) -> None:
    """Wie write_pdf, aber ein eigenes PageTemplate je Fach: mit --confirmation
    soll die Seitenzahl je Fach wieder bei 1 beginnen und die Fußzeile das
    jeweilige Fach statt "Fächer" nennen (Bestätigungs-Vorlage ist pro Fach an
    eine reale Person adressiert, da wirkt eine durchlaufende Gesamt-
    Seitenzahl/"Fächer"-Fußzeile über das ganze Dokument fehl am Platz)."""
    doc = BaseDocTemplate(
        _ziel(path), pagesize=A4, leftMargin=LEFT_MARGIN, rightMargin=RIGHT_MARGIN,
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
        footer_center = footer_context(subject, schoolyear_name)
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
                subject, by_subject[subject], schoolyear_name,
                confirmation=True, fkl_map=fkl_map, kollegium_map=kollegium_map, duplex=duplex,
                blank_pages=blank_pages,
                confirmation_page_count=page_counts[i] if page_counts else None,
                return_by=return_by, return_to=return_to,
            )
        )
    doc.addPageTemplates(templates)
    doc.build(story, canvasmaker=FooterCanvas)


def sanitize_filename(name: str) -> str:
    return name.replace("/", "-")

