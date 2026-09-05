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
import time
from datetime import datetime
from pathlib import Path
from typing import Any, Optional

# Wie oft und wie lange ``os.replace`` wiederholt wird, wenn Windows die
# Zieldatei gerade als geoeffnet meldet; zusammen gut eine halbe Sekunde
# (10+20+40+80+160+320 ms).
_ERSETZ_VERSUCHE = 7
_ERSETZ_WARTE_START = 0.01


def replace_with_retry(quelle: str, ziel: Path) -> None:
    """``os.replace``, das einen gleichzeitigen *Leser* unter Windows aussitzt.

    Unter POSIX ersetzt ``rename`` eine Datei auch dann, wenn sie jemand
    geoeffnet hat. Windows verweigert das mit ``PermissionError``
    (``WinError 5``), solange irgendein Handle auf die Zieldatei offen ist -
    und Python oeffnet ohne ``FILE_SHARE_DELETE``, ein lesendes ``open()``
    genuegt also.

    Das trifft hier den Fall, der auf dem Schul-Laptop am ehesten vorkommt:
    das Dashboard laedt die Mappe zum *Lesen* bewusst **ohne** Sperre (damit
    ein Leser nie auf einen Schreiber warten muss), waehrend ein zweites
    Fenster gerade speichert. ``load_workbook`` haelt die Datei fuer die Dauer
    des Lesens offen; faellt das Ersetzen genau hinein, scheitert das
    Speichern. Der Aufrufer im Dashboard uebersetzt jeden
    ``PermissionError`` in "Die Datei ist gerade in Excel geoeffnet" - eine
    Meldung, die dann schlicht falsch waere, weil niemand Excel offen hat.

    Ein Leser haelt die Mappe nur fuer die Dauer eines ``load_workbook``.
    Kurzes Wiederholen loest den Konflikt deshalb, ohne irgendwo zu sperren.
    Haelt der Fehler an, ist es *wirklich* Excel (oder ein Schreibschutz) -
    dann fliegt er weiter und die Meldung stimmt wieder.

    Bis 2026-09-04 gab es dieselbe Funktion ein zweites Mal, fast wortgleich,
    in ``sba-dashboard/app/cache.py`` - fuer den Sidecar-Cache statt fuer die
    Mappe. Zwei Kopien derselben sieben Zeilen Wiederholungslogik heisst zwei
    Stellen, an denen jemand die Wartezeit oder die Versuchszahl aendern
    kann, ohne die andere zu bemerken. Diese Funktion ist deshalb jetzt
    public (englischer Name, Export aus ``bestand/core/__init__.py``) und
    wird vom Dashboard importiert statt erneut abgeschrieben - siehe
    ``sba-dashboard/app/dateien.py``.
    """
    warte = _ERSETZ_WARTE_START
    for versuch in range(_ERSETZ_VERSUCHE):
        try:
            os.replace(quelle, ziel)
            return
        except PermissionError:
            if versuch == _ERSETZ_VERSUCHE - 1:
                raise
            time.sleep(warte)
            warte *= 2


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
        replace_with_retry(tmp_name, destination)
    except Exception:
        try:
            os.unlink(tmp_name)
        except FileNotFoundError:
            pass
        raise
    return backup
