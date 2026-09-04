"""Dauerhaftes Speichern einer Arbeitsmappe - atomar, mit Sicherungskopie.

Diese Funktion lag bis 2026-09-04 in ``ausleihe.inventory_excel``, also im
IServ-Client. Dort war sie falsch: sie kennt weder IServ noch HTTP, sondern
ausschliesslich Dateisystem und openpyxl. Ein Aufrufer, der nur eine Mappe
sicher speichern will, musste dafuer den kompletten API-Client installieren.

Sie gehoert hierher, weil beide Schalen um dieselbe Excel-Logik sie brauchen:
die CLI (``update_bestand*.py``) und das Dashboard (``sba-dashboard``).

Das Verfahren ist bewusst dreistufig:

1. **Sicherungskopie** der bestehenden Datei, mit Zeitstempel im Namen. Zwei
   Speichervorgaenge in derselben Sekunde bekommen einen Zaehler, damit keiner
   den anderen still ueberschreibt.
2. **Vollstaendig in eine Nachbardatei schreiben** und ``fsync``. Die Nachbardatei
   liegt im selben Verzeichnis, weil ``os.replace`` nur innerhalb eines
   Dateisystems atomar ist.
3. **``os.replace``**. Erst hier aendert sich die Zieldatei, und zwar in einem
   Schritt. Ein Abbruch davor - WLAN weg, Akku leer - laesst die alte Mappe
   unberuehrt.
"""
from __future__ import annotations

import os
import shutil
import tempfile
from datetime import datetime
from pathlib import Path
from typing import Any, Optional


def atomic_save_workbook(
    workbook: Any, destination: Path, *, backup_dir: Optional[Path] = None
) -> Optional[Path]:
    """Ersetzt ``destination`` dauerhaft und legt optional eine Sicherung an.

    Gibt den Pfad der Sicherungskopie zurueck, oder ``None``, wenn keine
    angelegt wurde. ``workbook`` muss lediglich ``save(pfad)`` koennen - das
    haelt die Funktion ohne openpyxl-Import testbar.
    """
    destination = destination.resolve()
    if not destination.exists():
        raise FileNotFoundError(f"Excel-Datei nicht gefunden: {destination}")
    backup: Optional[Path] = None
    if backup_dir is not None:
        backup_dir.mkdir(parents=True, exist_ok=True)
        stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
        backup = backup_dir / f"{destination.stem}.{stamp}{destination.suffix}"
        # Zwei Speichervorgaenge in derselben Sekunde duerfen sich nicht
        # gegenseitig ueberschreiben.
        counter = 1
        while backup.exists():
            backup = backup_dir / f"{destination.stem}.{stamp}-{counter}{destination.suffix}"
            counter += 1
        shutil.copy2(destination, backup)
    fd, tmp_name = tempfile.mkstemp(
        prefix=f".{destination.stem}.", suffix=destination.suffix, dir=destination.parent
    )
    try:
        os.close(fd)
        workbook.save(tmp_name)
        # r+b (nicht rb): os.fsync() braucht ein zum Schreiben geoeffnetes
        # Handle, sonst OSError: [Errno 9] Bad file descriptor unter Windows.
        with open(tmp_name, "r+b") as handle:
            os.fsync(handle.fileno())
        shutil.copymode(destination, tmp_name)
        os.replace(tmp_name, destination)
    except Exception:
        try:
            os.unlink(tmp_name)
        except FileNotFoundError:
            pass
        raise
    return backup
