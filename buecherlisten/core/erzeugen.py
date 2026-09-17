"""Bücherlisten-PDFs erzeugen - die Logik, die bis 2026-09-17 in ``main()`` stand.

Kommandozeile (``generate_booklists.py``) und Dashboard rufen beide
:func:`erzeuge_buecherlisten_pdfs` auf; die Kommandozeile schreibt die
Ergebnisse in Dateien, das Dashboard liefert sie im Browser aus.

Die Website-Zuordnungen (Fachkonferenzleitungen, Kürzel, Aufgabenfelder)
werden nur geladen, wenn sie gebraucht werden, und lassen sich über
:class:`Zuordnungen` vorgeben - so laufen Tests ohne Netz. Scheitert ein
Abruf, wird das PDF trotzdem erzeugt; die Meldung steht dann in
``ErzeugtesPdf.warnungen`` statt wie früher auf stderr.
"""
from __future__ import annotations

import io
from collections.abc import Callable
from dataclasses import dataclass, field
from typing import Literal

from buecherlisten.trg_web import (
    FAECHER_URL,
    FKL_URL,
    KOLLEGIUM_URL,
    fetch_aufgabenfeld_mapping,
    fetch_aufgabenfeld_mapping_from_faecher_page,
    fetch_fkl_mapping,
    fetch_kollegium_kuerzel_mapping,
    subject_sort_key,
)

from .daten import Ansicht, Buecherdaten
from .layout import (
    FACH_LISTE,
    JAHRGANG_LISTE,
    VERLAG_LISTE,
    Listenart,
    PageBreak,
    footer_context,
    measure_subject_pages,
    sanitize_filename,
    subject_story,
    write_combined_confirmation_pdf,
    write_pdf,
)

Modus = Literal["alphabet", "aufgabenfeld", "split"]

# Ansicht -> (Listenart, Bezeichnung im Titel des Sammel-PDFs)
_ARTEN: dict[str, tuple[Listenart, str]] = {
    "fach": (FACH_LISTE, "Fächer"),
    "verlag": (VERLAG_LISTE, "Verlage"),
    "jahrgang": (JAHRGANG_LISTE, "Jahrgänge"),
}


@dataclass
class Zuordnungen:
    """Vorgegebene Website-Zuordnungen; ``None`` heißt: bei Bedarf live laden."""

    fkl: dict[str, str] | None = None
    kollegium: dict[str, str] | None = None
    aufgabenfeld: dict[str, str] | None = None


@dataclass
class ErzeugtesPdf:
    inhalt: bytes
    dateiname: str
    titel: str
    warnungen: list[str] = field(default_factory=list)


def _aufgabenfelder(vorgabe: dict[str, str] | None, warnungen: list[str]) -> dict[str, str]:
    if vorgabe is not None:
        return vorgabe
    zuordnung: dict[str, str] = {}
    # FKL_URL zuerst, FAECHER_URL als Fallback (existiert FKL_URL nicht mehr
    # oder enthält ihre Tabelle keine Aufgabenfeld-Zuordnung mehr) — erst
    # wenn auch das scheitert, wird rein alphabetisch sortiert.
    quellen: tuple[tuple[str, Callable[[], dict[str, str]]], ...] = (
        (FKL_URL, fetch_aufgabenfeld_mapping),
        (FAECHER_URL, fetch_aufgabenfeld_mapping_from_faecher_page),
    )
    for url, hole in quellen:
        try:
            zuordnung = hole()
        except Exception as exc:  # Netzwerk/Parsing-Fehler -> nächste Quelle versuchen
            warnungen.append(f"Warnung: Aufgabenfelder konnten nicht von {url} geladen werden ({exc}).")
            continue
        if zuordnung:
            break
    if not zuordnung:
        warnungen.append(
            "Warnung: Aufgabenfeld-Zuordnung von keiner Quelle verfügbar "
            "— Fächer werden stattdessen alphabetisch sortiert."
        )
    return zuordnung


def _bestaetigungs_zuordnungen(
    vorgabe: Zuordnungen, warnungen: list[str],
) -> tuple[dict[str, str], dict[str, str]]:
    fkl = vorgabe.fkl
    if fkl is None:
        try:
            fkl = fetch_fkl_mapping()
        except Exception as exc:  # Netzwerk/Parsing-Fehler sollen die PDF-Erzeugung nicht abbrechen
            fkl = {}
            warnungen.append(
                f"Warnung: Fachkonferenzleitungen konnten nicht von {FKL_URL} geladen werden ({exc}) "
                "— Kopfzeile/Unterschriftszeile bleiben ohne Namen/Kürzel."
            )
    kollegium = vorgabe.kollegium
    if kollegium is None:
        try:
            kollegium = fetch_kollegium_kuerzel_mapping()
        except Exception as exc:
            kollegium = {}
            warnungen.append(
                f"Warnung: Lehrerkürzel konnten nicht von {KOLLEGIUM_URL} geladen werden ({exc}) "
                "— Kopfzeile/Unterschriftszeile zeigen ersatzweise den vollen Namen bzw. bleiben "
                "ohne Kürzel."
            )
    return fkl, kollegium


def erzeuge_buecherlisten_pdfs(
    daten: Buecherdaten,
    *,
    faecher: list[str] | None = None,
    modus: Modus = "alphabet",
    bestaetigung: bool = False,
    rueckgabe_bis: str | None = None,
    rueckgabe_an: str | None = None,
    doppelseitig: bool = False,
    nur_falls_noetig: bool = False,
    zuordnungen: Zuordnungen | None = None,
    ansicht: Ansicht = "fach",
) -> list[ErzeugtesPdf]:
    """Ein PDF (``alphabet``/``aufgabenfeld``) oder eines je Gruppe (``split``).

    ``ansicht`` wählt Fach-, Verlags- oder Jahrgangslisten; ``faecher`` sind
    dann exakte Namen aus ``daten.gruppen(ansicht)`` (Zuordnung ohne
    Groß-/Kleinschreibung: :func:`~buecherlisten.core.daten.waehle_gruppen`);
    ``None`` heißt alle. Bestätigung und ``aufgabenfeld`` gibt es nur für
    Fächer. ``doppelseitig`` entspricht ``--duplex``, zusammen mit
    ``nur_falls_noetig`` ``--duplex-if-needed``. Rückgabe-Angaben wirken nur
    mit ``bestaetigung``.
    """
    if ansicht != "fach" and (bestaetigung or modus == "aufgabenfeld"):
        raise ValueError("Bestätigung und Sortierung nach Aufgabenfeld gibt es nur für Fächer.")
    art, sammel_label = _ARTEN[ansicht]
    vorgabe = zuordnungen or Zuordnungen()
    warnungen: list[str] = []
    by_subject = daten.tabellen(ansicht)
    schoolyear_name = daten.schuljahr_name
    reihenfolge = daten.gruppen(ansicht)
    if faecher is None:
        subjects = reihenfolge
    elif ansicht == "fach":
        subjects = sorted(faecher, key=str.casefold)
    else:
        subjects = sorted(faecher, key=reihenfolge.index)
    if not bestaetigung:
        rueckgabe_bis = rueckgabe_an = None
    rueckgabe_bis = rueckgabe_bis or None
    rueckgabe_an = rueckgabe_an or None

    if modus == "aufgabenfeld":
        aufgabenfeld_map = _aufgabenfelder(vorgabe.aufgabenfeld, warnungen)
        subjects = sorted(subjects, key=lambda s: subject_sort_key(s, aufgabenfeld_map))

    fkl_map: dict[str, str] = {}
    kollegium_map: dict[str, str] = {}
    if bestaetigung:
        fkl_map, kollegium_map = _bestaetigungs_zuordnungen(vorgabe, warnungen)

    # Seitenzahlen je Fach werden für zwei Dinge gebraucht: --duplex-if-needed
    # (Leerseiten nur einfügen, wenn mind. ein Fach mehr als 1 Seite braucht)
    # und --confirmation (Verweis im Bestätigungssatz auf "oben"/"umseitig"/
    # "auf den vorliegenden Seiten" je nachdem, wie viele Seiten die
    # Bücherliste vor dem Bestätigungsblock einnimmt). Ein einziger Messlauf
    # deckt beide Fälle ab.
    duplex_if_needed = doppelseitig and nur_falls_noetig
    page_counts: list[int] | None = None
    if duplex_if_needed or bestaetigung:
        page_counts = measure_subject_pages(
            subjects, by_subject, schoolyear_name,
            confirmation=bestaetigung, fkl_map=fkl_map, kollegium_map=kollegium_map,
            return_by=rueckgabe_bis, return_to=rueckgabe_an, art=art,
        )
    effective_duplex = doppelseitig
    if duplex_if_needed:
        effective_duplex = any(count > 1 for count in page_counts or [])

    sy_label = sanitize_filename(daten.schuljahr_id)
    # Bestätigung hängt "Bestätigung " vor Titel/Dateinamen, damit ein
    # Bestätigungs-Lauf die Datei eines normalen Laufs nicht überschreibt und
    # Bestätigungs-PDFs sofort als solche erkennbar sind.
    title_prefix = "Bestätigung " if bestaetigung else ""

    if modus != "split":
        label = sammel_label if modus == "alphabet" else "Fächer (nach Aufgabenfeld)"
        dateiname = f"{title_prefix}Bücherliste {sammel_label} {sy_label}.pdf"
        title = f"{title_prefix}Bücherliste {label} {daten.schuljahr_id}"
        puffer = io.BytesIO()
        if bestaetigung:
            # Bestätigungs-Vorlage ist pro Fach an eine reale Person adressiert
            # (Fachkonferenzleitung) — Seitenzahl zählt daher je Fach neu, und
            # die Fußzeile nennt das jeweilige Fach statt pauschal "Fächer".
            write_combined_confirmation_pdf(
                puffer, subjects, by_subject, schoolyear_name,
                fkl_map=fkl_map, kollegium_map=kollegium_map, title=title,
                duplex=effective_duplex, page_counts=page_counts,
                return_by=rueckgabe_bis, return_to=rueckgabe_an,
            )
        else:
            blank_pages: set[int] = set()
            story: list = []
            for i, subject in enumerate(subjects):
                if i > 0:
                    story.append(PageBreak())
                story.extend(
                    subject_story(
                        subject, by_subject[subject], schoolyear_name,
                        confirmation=False, fkl_map=fkl_map, kollegium_map=kollegium_map,
                        duplex=effective_duplex, blank_pages=blank_pages, art=art,
                    )
                )
            write_pdf(
                puffer, story, title=title, footer_center=footer_context(label, schoolyear_name),
                blank_pages=blank_pages,
            )
        return [ErzeugtesPdf(puffer.getvalue(), dateiname, title, warnungen)]

    ergebnisse = []
    for i, subject in enumerate(subjects):
        # Jedes Fach ist im split-Modus ein eigenes Dokument -> eigene
        # blank_pages-Menge.
        einzel_leer: set[int] = set()
        story = subject_story(
            subject, by_subject[subject], schoolyear_name,
            confirmation=bestaetigung, fkl_map=fkl_map, kollegium_map=kollegium_map,
            duplex=effective_duplex, blank_pages=einzel_leer,
            confirmation_page_count=page_counts[i] if page_counts else None,
            return_by=rueckgabe_bis, return_to=rueckgabe_an, art=art,
        )
        title = f"{title_prefix}Bücherliste {subject} {daten.schuljahr_id}"
        puffer = io.BytesIO()
        write_pdf(
            puffer, story, title=title, footer_center=footer_context(subject, schoolyear_name),
            blank_pages=einzel_leer,
        )
        ergebnisse.append(ErzeugtesPdf(
            puffer.getvalue(), f"{title_prefix}Bücherliste {subject} {sy_label}.pdf", title,
            # Die Warnungen betreffen den ganzen Lauf; nur einmal melden.
            warnungen if i == 0 else [],
        ))
    return ergebnisse


def erzeuge_schuelerlisten_pdfs(
    daten: Buecherdaten,
    hole_pdf: Callable[[int], bytes],
    *,
    jahrgaenge: list[str] | None = None,
    modus: Literal["alphabet", "split"] = "alphabet",
    doppelseitig: bool = False,
    nur_falls_noetig: bool = False,
) -> list[ErzeugtesPdf]:
    """Die IServ-Druckversionen ("Schülerliste") der Jahrgangslisten.

    ``hole_pdf`` bekommt die ID einer Bücherliste und liefert deren PDF, z. B.
    ``lambda i: client.admin.get_booklist_pdf(daten.schuljahr_id, i)``.
    ``alphabet`` hängt die Jahrgänge aufsteigend zu einem PDF zusammen,
    ``split`` liefert eines je Jahrgang. Doppelseitig wie bei den eigenen
    Listen: jeder Jahrgang wird mit einer leeren Seite auf gerade Seitenzahl
    gebracht (mit ``nur_falls_noetig`` nur, wenn einer mehr als eine Seite hat).
    """
    from pypdf import PdfReader, PdfWriter  # Extra "pdf", wie reportlab

    namen = daten.jahrgaenge
    if jahrgaenge is not None:
        namen = sorted(jahrgaenge, key=namen.index)
    leser = [PdfReader(io.BytesIO(hole_pdf(daten.listen_ids[name]))) for name in namen]
    auffuellen = doppelseitig and (
        not nur_falls_noetig or any(len(r.pages) > 1 for r in leser)
    )
    sy_label = sanitize_filename(daten.schuljahr_id)

    def schreibe(teile: list[PdfReader], titel: str) -> bytes:
        writer = PdfWriter()
        for reader in teile:
            for seite in reader.pages:
                writer.add_page(seite)
            if auffuellen and len(reader.pages) % 2 == 1:
                letzte = reader.pages[-1].mediabox
                writer.add_blank_page(width=letzte.width, height=letzte.height)
        writer.add_metadata({"/Title": titel})
        puffer = io.BytesIO()
        writer.write(puffer)
        return puffer.getvalue()

    if modus == "split":
        ergebnisse = []
        for name, reader in zip(namen, leser):
            titel = f"Bücherliste {name} {daten.schuljahr_id} (Schülerliste)"
            ergebnisse.append(ErzeugtesPdf(
                schreibe([reader], titel), f"Bücherliste {name} {sy_label} (Schülerliste).pdf", titel,
            ))
        return ergebnisse
    titel = f"Bücherliste Jahrgänge {daten.schuljahr_id} (Schülerliste)"
    return [ErzeugtesPdf(
        schreibe(leser, titel), f"Bücherliste Jahrgänge {sy_label} (Schülerliste).pdf", titel,
    )]
