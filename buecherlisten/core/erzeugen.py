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

from .daten import Buecherdaten
from .layout import (
    PageBreak,
    footer_context,
    measure_subject_pages,
    sanitize_filename,
    subject_story,
    write_combined_confirmation_pdf,
    write_pdf,
)

Modus = Literal["alphabet", "aufgabenfeld", "split"]


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
) -> list[ErzeugtesPdf]:
    """Ein PDF (``alphabet``/``aufgabenfeld``) oder eines je Fach (``split``).

    ``faecher`` sind exakte Fachnamen aus ``daten.faecher`` (Zuordnung ohne
    Groß-/Kleinschreibung: :func:`~buecherlisten.core.daten.waehle_faecher`);
    ``None`` heißt alle. ``doppelseitig`` entspricht ``--duplex``, zusammen
    mit ``nur_falls_noetig`` ``--duplex-if-needed``. Rückgabe-Angaben wirken
    nur mit ``bestaetigung``.
    """
    vorgabe = zuordnungen or Zuordnungen()
    warnungen: list[str] = []
    by_subject = daten.je_fach
    schoolyear_name = daten.schuljahr_name
    subjects = sorted(faecher, key=str.casefold) if faecher is not None else daten.faecher
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
            return_by=rueckgabe_bis, return_to=rueckgabe_an,
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
        label = "Fächer" if modus == "alphabet" else "Fächer (nach Aufgabenfeld)"
        dateiname = f"{title_prefix}Bücherliste Fächer {sy_label}.pdf"
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
                        duplex=effective_duplex, blank_pages=blank_pages,
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
            return_by=rueckgabe_bis, return_to=rueckgabe_an,
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
