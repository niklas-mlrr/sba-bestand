"""Atomares Speichern: die Zieldatei aendert sich nur im Erfolgsfall.

Die Faelle standen bis 2026-09-04 in ``ausleihe-api/tests/test_inventory_excel.py``
und sind mit der Funktion hierher gewandert. Das Prueffalz braucht bewusst kein
openpyxl: ``atomic_save_workbook`` verlangt vom Workbook nur ``save(pfad)``.
"""
from __future__ import annotations

from pathlib import Path

import pytest

from bestand.core import atomic_save_workbook, excel_io


class Workbook:
    def __init__(self, content: bytes = b"new"):
        self.content = content

    def save(self, path):
        Path(path).write_bytes(self.content)


def test_ersetzt_erst_nach_erfolg_und_legt_sicherung_an(tmp_path):
    ziel = tmp_path / "bestand.xlsx"
    ziel.write_bytes(b"old")
    sicherung = atomic_save_workbook(Workbook(), ziel, backup_dir=tmp_path / "backups")
    assert ziel.read_bytes() == b"new"
    assert sicherung and sicherung.read_bytes() == b"old"


def test_fehler_beim_speichern_laesst_das_original_unberuehrt(tmp_path):
    ziel = tmp_path / "bestand.xlsx"
    ziel.write_bytes(b"old")

    class FailingWorkbook:
        def save(self, _):
            raise RuntimeError("boom")

    with pytest.raises(RuntimeError):
        atomic_save_workbook(FailingWorkbook(), ziel)
    assert ziel.read_bytes() == b"old"


def test_fehlgeschlagener_speichervorgang_laesst_keine_temporaerdatei_zurueck(tmp_path):
    """Ein Rest im Ordner der Mappe faende sich sonst spaeter auf dem Netzlaufwerk."""
    ziel = tmp_path / "bestand.xlsx"
    ziel.write_bytes(b"old")

    class FailingWorkbook:
        def save(self, _):
            raise RuntimeError("boom")

    with pytest.raises(RuntimeError):
        atomic_save_workbook(FailingWorkbook(), ziel)
    assert [p.name for p in tmp_path.iterdir()] == ["bestand.xlsx"]


def test_zwei_sicherungen_in_derselben_sekunde_ueberschreiben_sich_nicht(tmp_path):
    ziel = tmp_path / "bestand.xlsx"
    ziel.write_bytes(b"eins")
    backups = tmp_path / "backups"
    erste = atomic_save_workbook(Workbook(b"zwei"), ziel, backup_dir=backups)
    zweite = atomic_save_workbook(Workbook(b"drei"), ziel, backup_dir=backups)
    assert erste != zweite
    assert erste.read_bytes() == b"eins"
    assert zweite.read_bytes() == b"zwei"


def test_fehlende_zieldatei_ist_ein_fehler(tmp_path):
    with pytest.raises(FileNotFoundError):
        atomic_save_workbook(Workbook(), tmp_path / "gibt-es-nicht.xlsx")


# ── Windows: gleichzeitiger Leser blockiert das Ersetzen ─────────────────────

def test_ersetzen_sitzt_einen_gleichzeitigen_leser_aus(tmp_path, monkeypatch):
    """Ein kurzer ``PermissionError`` ist unter Windows ein Leser, kein Schreibschutz.

    Dort scheitert ``os.replace``, solange irgendein Handle auf die Zieldatei
    offen ist. Das Dashboard liest die Mappe bewusst ohne Sperre, ein zweites
    Fenster kann also genau waehrend des Speicherns lesen. Unter POSIX passiert
    das nie, deshalb wird es hier nachgestellt.
    """
    ziel = tmp_path / "bestand.xlsx"
    ziel.write_bytes(b"old")
    echtes_replace = excel_io.os.replace
    versuche = []

    def zickig(quelle, dest):
        versuche.append(1)
        if len(versuche) <= 2:
            raise PermissionError(5, "Access is denied")
        return echtes_replace(quelle, dest)

    monkeypatch.setattr(excel_io.os, "replace", zickig)
    monkeypatch.setattr(excel_io.time, "sleep", lambda _: None)

    atomic_save_workbook(Workbook(), ziel)

    assert ziel.read_bytes() == b"new"
    assert len(versuche) == 3


def test_dauerhafter_permission_error_wird_durchgereicht(tmp_path, monkeypatch):
    """Haelt der Fehler an, ist die Mappe wirklich belegt - dann muss er hochkommen.

    Das Dashboard macht daraus "Die Datei ist gerade in Excel geoeffnet", und
    diese Meldung ist dann richtig. Die alte Mappe bleibt unveraendert, und die
    Nachbardatei wird aufgeraeumt.
    """
    ziel = tmp_path / "bestand.xlsx"
    ziel.write_bytes(b"old")

    def immer(quelle, dest):
        raise PermissionError(5, "Access is denied")

    monkeypatch.setattr(excel_io.os, "replace", immer)
    monkeypatch.setattr(excel_io.time, "sleep", lambda _: None)

    with pytest.raises(PermissionError):
        atomic_save_workbook(Workbook(), ziel)

    assert ziel.read_bytes() == b"old"
    assert list(tmp_path.glob(".bestand.*")) == []
