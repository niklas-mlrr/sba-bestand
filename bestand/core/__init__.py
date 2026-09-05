"""Bibliothekskern der Bestandsliste - netzfrei testbar, ohne CLI-Anteile.

Nichts hier liest ``os.environ``, laedt eine ``.env``, parst Argumente oder
schreibt nach stdout. Die CLI (``update_bestand_auto.py``) und das Dashboard
(``sba-dashboard``) sind zwei duenne Schalen um dieselben Funktionen.
"""
from .config import BestandConfig, ConfigError
from .excel_io import atomic_save_workbook, replace_with_retry
from .grid import (
    Grid,
    GridCell,
    GridEntry,
    GridSlot,
    classify_row,
    extract_grade,
    find_blocks,
    find_fach_for_col,
    find_zustand_for_col,
    parse_grid,
    resolve_anchor,
    strip_hint,
)
from .iserv import (
    EV_BOOKLISTS,
    EV_ENROLLMENTS,
    EV_GRADE_BOOKS,
    EV_NO_BOOKLIST,
    EV_SERIES,
    LazyGradeBooks,
    Snapshot,
    fetch_enrollment_counts_by_grade,
    fetch_series_data,
    fetch_snapshot,
    load_grade_books,
)
from .update import (
    STAND_NUMBER_FORMAT,
    CellChange,
    UpdateResult,
    ZuBestellenRow,
    apply_snapshot,
    compute_zu_bestellen_rows,
    format_isbn,
    load_bestellt_counts,
    rebuild_zu_bestellen,
    write_stand,
)

__all__ = [
    "BestandConfig", "ConfigError",
    "atomic_save_workbook", "replace_with_retry",
    "Grid", "GridCell", "GridEntry", "GridSlot",
    "classify_row", "extract_grade", "find_blocks", "find_fach_for_col",
    "find_zustand_for_col", "parse_grid", "resolve_anchor", "strip_hint",
    "EV_BOOKLISTS", "EV_ENROLLMENTS", "EV_GRADE_BOOKS", "EV_NO_BOOKLIST", "EV_SERIES",
    "LazyGradeBooks", "Snapshot", "fetch_enrollment_counts_by_grade",
    "fetch_series_data", "fetch_snapshot", "load_grade_books",
    "STAND_NUMBER_FORMAT", "CellChange", "UpdateResult", "ZuBestellenRow",
    "apply_snapshot", "compute_zu_bestellen_rows", "format_isbn",
    "load_bestellt_counts", "rebuild_zu_bestellen", "write_stand",
]
