"""Friert die Konsolenausgabe von update_bestand_auto.py ein.

Der Refactor auf ``bestand/core/`` darf die Ausgabe nicht verändern. Der
Vergleich läuft gegen die Dateien in ``tests/golden/``, die vor dem Refactor
vom alten Skript aufgezeichnet wurden. Zeitpunkt, Pfade und Backup-Stempel
werden normalisiert - alles andere muss zeichengleich sein.
"""
from __future__ import annotations

import contextlib
import importlib.util
import io
import os
import re
from pathlib import Path

import pytest
from conftest import FakeClient, build_workbook

_ROOT = Path(__file__).resolve().parent.parent
_GOLDEN = Path(__file__).parent / "golden"
_TIMESTAMP = re.compile(r"\d{2}\.\d{2}\.\d{4} \d{2}:\d{2}:\d{2}")
_BACKUP_STAMP = re.compile(r"Bestand-Test\.\d{8}-\d{6}(-\d+)?\.xlsx")


def _load_cli():
    """Lädt update_bestand_auto.py als Modul - es ist ein Skript, kein Paket-Modul."""
    os.environ["ISERV_DOMAIN"] = "beispiel-schule.de"
    spec = importlib.util.spec_from_file_location(
        "update_bestand_auto", _ROOT / "bestand" / "update_bestand_auto.py"
    )
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def run_cli(tmp_path: Path, extra_args: list[str], *, skip_blocked: bool) -> tuple[str, int, Path]:
    cli = _load_cli()
    cli.SKIP_BLOCKED = skip_blocked
    xlsx = build_workbook(tmp_path / "Bestand-Test.xlsx")
    buffer = io.StringIO()
    code = 0
    try:
        with contextlib.redirect_stdout(buffer):
            cli.main(["--excel", str(xlsx)] + extra_args, client_factory=FakeClient)
    except SystemExit as exc:
        code = exc.code or 0
    text = buffer.getvalue()
    text = _TIMESTAMP.sub("<STAND>", text)
    text = text.replace(str(xlsx), "<XLSX>").replace(str(tmp_path), "<TMP>")
    text = _BACKUP_STAMP.sub("Bestand-Test.<STAMP>.xlsx", text)
    return text, code, xlsx


@pytest.mark.parametrize(
    ("golden_name", "extra_args"),
    [("cli_dry_run.txt", ["--dry-run"]), ("cli_save.txt", [])],
)
def test_cli_output_unchanged(tmp_path, golden_name, extra_args):
    """Ohne Regel 4 ist die Ausgabe zeichengleich mit dem Stand vor dem Refactor."""
    text, code, _ = run_cli(tmp_path, extra_args, skip_blocked=False)
    assert code == 0
    expected = (_GOLDEN / golden_name).read_text(encoding="utf-8")
    assert text == expected


def test_blocked_areas_disappear_from_skipped(tmp_path):
    """Regel 4: Sperrflächen tauchen nicht mehr als 'kein Buch' auf.

    Im Testblatt ist Latein in Jg 5/6 (J3:M4) und Jg 12 (J8:M13) als Merge über
    die volle Blockbreite gesperrt; in Jg 7 fehlt schlicht das Buch. Nur die
    Jg-7-Meldung darf übrig bleiben.
    """
    text, code, _ = run_cli(tmp_path, ["--dry-run"], skip_blocked=True)
    assert code == 0
    skipped = [line for line in text.splitlines() if line.startswith("  - Sp.")]
    assert skipped == [
        "  - Sp.J/Zeile 5: Kein Buch-Match für Fach 'Latein'.",
        "  - Sp.K/Zeile 5: Kein Buch-Match für Fach 'Latein'.",
        "  - Sp.L/Zeile 5: Kein Buch-Match für Fach 'Latein'.",
    ]
    # Die geschriebenen Zellen bleiben identisch - Sperrflächen waren ohnehin leer.
    ohne_regel, _, _ = run_cli(tmp_path, ["--dry-run"], skip_blocked=False)
    changes = [line for line in text.splitlines() if re.match(r"^  [A-Z]+\d+: ", line)]
    changes_ohne = [line for line in ohne_regel.splitlines() if re.match(r"^  [A-Z]+\d+: ", line)]
    assert changes == changes_ohne
