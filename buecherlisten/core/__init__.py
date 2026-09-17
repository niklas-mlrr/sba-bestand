"""Bücherlisten nach Fach als PDF - die Bibliothek hinter ``generate_booklists.py``.

Nach dem Vorbild von ``bestand/core/``: das Skript ist nur noch Kommandozeile,
und ``sba-dashboard`` importiert dieselben Funktionen.

* :mod:`.daten` - Laden aus der Ausleihe-API und Zusammenstellen nach Fach.
* :mod:`.layout` - das ausgemessene reportlab-Layout.
* :mod:`.erzeugen` - ein PDF (oder eines je Fach) als Bytes.

``layout`` und ``erzeugen`` brauchen reportlab (Extra ``pdf``).
"""
