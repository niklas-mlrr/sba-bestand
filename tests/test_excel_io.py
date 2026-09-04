"""Atomares Speichern: die Zieldatei aendert sich nur im Erfolgsfall.

Die Faelle standen bis 2026-09-04 in ``ausleihe-api/tests/test_inventory_excel.py``
und sind mit der Funktion hierher gewandert. Das Prueffalz braucht bewusst kein
openpyxl: ``atomic_save_workbook`` verlangt vom Workbook nur ``save(pfad)``.
"""
from __future__ import annotations

from pathlib import Path

import pytest

from bestand.core import atomic_save_workbook


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
