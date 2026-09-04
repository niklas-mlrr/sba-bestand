"""Alles, was aus IServ kommt - gebuendelt in einem Snapshot.

Der Client wird **injiziert**, nie hier gebaut: Wer die Zugangsdaten besorgt
(CLI aus der ``.env``, Dashboard aus dem Formular), entscheidet der Aufrufer.
Ohne Client laesst sich jede nachgelagerte Funktion offline testen.

Nur GET-Zugriffe.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime
from typing import Any, Callable, Iterator, Mapping, Optional

try:  # ausleihe ist beim reinen Offline-Test nicht noetig
    from ausleihe import NotFoundError
except Exception:  # pragma: no cover - nur ohne installierte Bibliothek
    class NotFoundError(Exception):  # type: ignore[no-redef]
        pass

# Fortschritt: (Ereignis, Nutzlast). Ereignisnamen sind stabil, die Texte nicht -
# jede Oberflaeche formuliert selbst.
Progress = Optional[Callable[[str, dict], None]]

EV_BOOKLISTS = "booklists"
EV_ENROLLMENTS = "enrollments"
EV_SERIES = "series"
EV_GRADE_BOOKS = "grade_books"
EV_NO_BOOKLIST = "no_booklist"


def load_grade_books(client, sy_id: str, bl_id: int) -> list[dict]:
    """Alle ausleihbaren Buecher einer Buecherliste (nur GET, gefiltert).

    Gefiltert: borrowable, Leihpreis > 0, kein 'eBook' im Titel, ISBN eindeutig.
    """
    try:
        detail = client.schoolyears.get_booklist(sy_id, bl_id)
    except NotFoundError:
        return []

    seen_isbns: set[str] = set()
    books: list[dict] = []
    for sec in detail.get("sections", []):
        for opt in sec.get("options", []):
            for item in opt.get("items", []):
                if not item.get("borrowable"):
                    continue
                sd = item.get("series_data", {}) or {}
                if (sd.get("fee") or 0) <= 0:
                    continue
                title = sd.get("title", "") or ""
                if "eBook" in title:
                    continue
                isbn = sd.get("isbn") or item.get("series") or ""
                if isbn in seen_isbns:
                    continue
                seen_isbns.add(isbn)
                books.append({
                    "isbn": isbn,
                    "title": title,
                    "subjects": sd.get("subjectsFlat", []) or [],
                })
    return books


def fetch_enrollment_counts_by_grade(
    client, sy_id: str
) -> tuple[dict[tuple[int, str], int], dict[tuple[int, str], int]]:
    """Zaehlt Anmeldungen und Bezahlungen pro (Jahrgang, ISBN).

    enrolled[(grade, isbn)] = Anzahl nicht-geloeschter Anmeldungen mit diesem Buch
    paid[(grade, isbn)]     = davon mit amountOpen == 0 (vollstaendig bezahlt)
    """
    enrollments = client.admin.get_enrollments(sy_id)
    enrolled: dict[tuple[int, str], int] = {}
    paid: dict[tuple[int, str], int] = {}

    for enr in enrollments:
        if enr.get("deleted_at"):
            continue
        grade = (enr.get("Booklist") or {}).get("grade")
        if grade is None:
            continue
        is_paid = enr.get("amountOpen", 1) == 0
        for item in enr.get("booklistItems", []):
            isbn = item.get("series")
            if not isbn:
                continue
            key = (grade, isbn)
            enrolled[key] = enrolled.get(key, 0) + 1
            if is_paid:
                paid[key] = paid.get(key, 0) + 1

    return enrolled, paid


def fetch_series_data(client) -> dict[str, dict]:
    """Laedt alle Serien; gibt pro ISBN 'total', 'publisher', 'price', 'title'."""
    return {
        s.isbn: {
            "total": s.total or 0,
            "publisher": s.publisher or "",
            "price": s.price or 0.0,
            "title": s.title or "",
        }
        for s in client.series.get_all(detailed=True)
        if s.isbn
    }


class LazyGradeBooks(Mapping):
    """Buecherlisten je Jahrgang, erst beim ersten Zugriff geladen.

    Der Abruf einer Buecherliste ist eine eigene HTTP-Runde je Jahrgang. Das
    Blatt sagt erst waehrend des Durchlaufs, welche Jahrgaenge ueberhaupt
    vorkommen - also wird hier gecacht statt vorgeladen. Ein Dict tut es fuer
    Tests genauso.
    """

    def __init__(self, client, sy_id: str, booklists_by_grade: dict[int, dict], progress: Progress = None):
        self._client = client
        self._sy_id = sy_id
        self._booklists = booklists_by_grade
        self._progress = progress
        self._cache: dict[int, list[dict]] = {}

    def _emit(self, event: str, **payload: Any) -> None:
        if self._progress is not None:
            self._progress(event, payload)

    def __getitem__(self, grade: int) -> list[dict]:
        if grade not in self._cache:
            bl = self._booklists.get(grade)
            if bl:
                self._emit(EV_GRADE_BOOKS, grade=grade)
                self._cache[grade] = load_grade_books(self._client, self._sy_id, bl["id"])
            else:
                self._emit(EV_NO_BOOKLIST, grade=grade, schoolyear_id=self._sy_id)
                self._cache[grade] = []
        return self._cache[grade]

    def __iter__(self) -> Iterator[int]:
        return iter(self._booklists)

    def __len__(self) -> int:
        return len(self._booklists)

    def materialize(self) -> dict[int, list[dict]]:
        """Laedt alle bekannten Jahrgaenge und gibt ein gewoehnliches Dict zurueck."""
        return {grade: self[grade] for grade in self._booklists}


@dataclass(frozen=True)
class Snapshot:
    """Ein vollstaendiger IServ-Stand. Enthaelt kein Excel und keinen Client."""
    schoolyear_id: str
    fetched_at: datetime
    booklists_by_grade: dict[int, dict] = field(default_factory=dict)
    grade_books: Mapping[int, list[dict]] = field(default_factory=dict)
    enrolled: dict[tuple[int, str], int] = field(default_factory=dict)
    paid: dict[tuple[int, str], int] = field(default_factory=dict)
    series_data: dict[str, dict] = field(default_factory=dict)

    @property
    def bestand_by_isbn(self) -> dict[str, int]:
        return {isbn: data["total"] for isbn, data in self.series_data.items()}

    def books_for_grade(self, grade: int) -> list[dict]:
        try:
            return self.grade_books[grade]
        except KeyError:
            return []


def fetch_snapshot(
    client,
    sy_id: str | None = None,
    *,
    progress: Progress = None,
    fetched_at: datetime | None = None,
    eager: bool = False,
) -> Snapshot:
    """Holt Buecherlisten, Anmeldungen und Serien; Buecherlisten je Jahrgang traege.

    ``eager=True`` laedt auch die Buecherlisten aller Jahrgaenge sofort - das
    Dashboard will einen abgeschlossenen Stand, die CLI will die Meldungen
    an derselben Stelle wie bisher.
    """
    def emit(event: str, **payload: Any) -> None:
        if progress is not None:
            progress(event, payload)

    if sy_id is None:
        sy_id = client.schoolyears.get_current()["id"]

    emit(EV_BOOKLISTS)
    booklists = client.schoolyears.get_booklists(sy_id)
    booklists_by_grade = {bl["grade"]: bl for bl in booklists if bl.get("grade") is not None}

    emit(EV_ENROLLMENTS)
    enrolled, paid = fetch_enrollment_counts_by_grade(client, sy_id)

    emit(EV_SERIES)
    series_data = fetch_series_data(client)

    grade_books: Mapping[int, list[dict]] = LazyGradeBooks(
        client, sy_id, booklists_by_grade, progress
    )
    if eager:
        grade_books = grade_books.materialize()  # type: ignore[union-attr]

    return Snapshot(
        schoolyear_id=sy_id,
        fetched_at=fetched_at or datetime.now().replace(microsecond=0),
        booklists_by_grade=booklists_by_grade,
        grade_books=grade_books,
        enrolled=enrolled,
        paid=paid,
        series_data=series_data,
    )
