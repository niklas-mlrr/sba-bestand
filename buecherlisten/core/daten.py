"""Bücherlisten-Daten aus der IServ-Ausleihe-API, nach Fach zusammengestellt.

Aus ``generate_booklists.py`` herausgelöst (2026-09-17), damit das Dashboard
dieselbe Zusammenstellung nutzen kann wie das Kommandozeilen-Skript - nach dem
Vorbild von ``bestand/core/``. Rein lesend, ohne Netz- oder Pfad-Seiteneffekte
beim Import.
"""
from __future__ import annotations

from collections import defaultdict
from dataclasses import dataclass
from typing import Any, Protocol

import isbnlib


class _Schuljahre(Protocol):
    def get_current(self) -> dict: ...
    def get_by_id(self, schoolyear_id: str) -> dict: ...
    def get_booklists(self, schoolyear_id: str) -> list[dict]: ...
    def get_booklist(self, schoolyear_id: str, booklist_id: int) -> dict: ...


class BuecherlistenClient(Protocol):
    """Was von ``ausleihe.AusleiheClient`` gebraucht wird."""

    @property
    def schoolyears(self) -> Any: ...


def format_isbn(isbn: str) -> str:
    try:
        masked = isbnlib.mask(isbn)
        return masked if masked else isbn
    except Exception:
        return isbn


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


def fmt_grades(grades: tuple[int, ...]) -> str:
    # Komma + Leerzeichen statt "/": erlaubt Zeilenumbruch in der schmalen
    # Klasse-Spalte, wenn ein Mehrjahresband viele Klassen abdeckt.
    return ", ".join(str(g) for g in grades)


def collect_entries(client: BuecherlistenClient, schoolyear_id: str) -> dict[tuple[str, str], dict]:
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


@dataclass(frozen=True)
class Buecherdaten:
    """Ein geladenes Schuljahr: Kennung, Anzeigename und die Tabellen je Fach."""

    schuljahr_id: str
    schuljahr_name: str
    je_fach: dict[str, dict[str, list[dict]]]

    @property
    def faecher(self) -> list[str]:
        return sorted(self.je_fach, key=str.casefold)


def lade_buecherdaten(client: BuecherlistenClient, schuljahr: str | None = None) -> Buecherdaten:
    """Lädt ein Schuljahr (Default: das laufende). ``NotFoundError`` fliegt durch."""
    if schuljahr:
        name = client.schoolyears.get_by_id(schuljahr)["name"]
        schuljahr_id = schuljahr
    else:
        aktuell = client.schoolyears.get_current()
        schuljahr_id, name = aktuell["id"], aktuell.get("name") or aktuell["id"]
    tabellen = build_subject_tables(collect_entries(client, schuljahr_id))
    return Buecherdaten(schuljahr_id=schuljahr_id, schuljahr_name=name, je_fach=dict(tabellen))


def waehle_faecher(verfuegbar: list[str], gewuenscht: list[str]) -> tuple[list[str], list[str]]:
    """Ordnet gewünschte Fachnamen ohne Rücksicht auf Groß-/Kleinschreibung zu.

    Gibt (gefundene, alphabetisch; unbekannte, in Eingabereihenfolge) zurück.
    """
    je_casefold = {s.casefold(): s for s in verfuegbar}
    gefunden: list[str] = []
    unbekannt: list[str] = []
    for wunsch in gewuenscht:
        treffer = je_casefold.get(wunsch.casefold())
        if treffer is None:
            unbekannt.append(wunsch)
        elif treffer not in gefunden:
            gefunden.append(treffer)
    return sorted(gefunden, key=str.casefold), unbekannt
