# Bücherlisten nach Fach — PDF-Export

Erzeugt aus den regulären Jahrgangs-Bücherlisten der IServ-Ausleihe-API eine
fachweise sortierte PDF-Bücherliste. Rein lesend (nur GET) — kein Schreibzugriff
auf die Produktionsdatenbank.

Es gibt keinen eigenen API-Endpunkt für diese Sicht. Das Skript holt alle
Jahrgangs-Bücherlisten eines Schuljahrs (`GET /schoolyears/:id/booklists/:bl_id`)
und stellt die Bücher clientseitig nach Fach neu zusammen — inklusive korrektem
Zusammenführen von Mehrjahresbänden (z.B. "Elemente Chemie 5/6"), die in
mehreren Jahrgangs-Bücherlisten gleichzeitig auftauchen.

## Schnellstart

```bash
# einmalig im sba-bestand-Root, falls noch nicht geschehen.
# --extra pdf ist hier Pflicht: reportlab ist ein Extra und wird nur von
# diesem Erzeuger gebraucht (siehe ../README.md).
uv sync --extra pdf

cd buecherlisten

# 1 PDF, neue Seite pro Fach, alphabetisch sortiert (laufendes Schuljahr, Default):
uv run python3 generate_booklists.py --mode alphabet

# 1 PDF, neue Seite pro Fach, nach Aufgabenfeld sortiert (laut TRG-Website):
uv run python3 generate_booklists.py --mode aufgabenfeld

# 1 PDF-Datei pro Fach:
uv run python3 generate_booklists.py --mode split

# bestimmtes Schuljahr, eigener Zielordner:
uv run python3 generate_booklists.py --schoolyear "2025/2026" --mode split --output-dir ~/Downloads
```

`uv run` sorgt dafür, dass die von `uv sync` installierten Abhängigkeiten
(reportlab, requests, …) auch tatsächlich verwendet werden — ein einfaches
`python3 generate_booklists.py` mit dem System-Python schlägt sonst mit
`ModuleNotFoundError: No module named 'reportlab'` fehl. Alternativ direkt mit
`.venv/bin/python3 generate_booklists.py ...` oder nach Aktivieren von
`.venv` (`source ../.venv/bin/activate`) mit `python3 ...`.

## Inhalt pro Fach

Zwei Tabellen — **Leihbare Bücher** (`borrowable=True`) und **Selbst
anzuschaffende Bücher** (`borrowable=False`: Arbeitshefte, "1x pro
Familie"-Anschaffungen, Digitallizenzen etc.) — mit den Spalten Klasse, Titel,
Verlag, ISBN, Neupreis, Leihgebühr.

- **ISBN** wird mit Bindestrichen dargestellt (`isbnlib.mask`, wie im
  Bestand-Tooling unter `../bestand/`) — die API liefert
  ISBNs immer ohne Trennzeichen, dafür gibt es keinen direkten Endpunkt.
- **Klasse:** Ein Buch, das in mehreren Jahrgängen angeboten wird
  (Mehrjahresband), erscheint einmal mit allen Klassen (z.B. "5/6"). Sortiert
  wird zuerst nach der untersten, dann nach der zweituntersten Klasse usw.
  Innerhalb gleicher Klassen-Kombination zusätzlich alphabetisch nach Titel.
- Bücher, die zu mehreren Fächern gehören (z.B. fächerübergreifende
  Formelsammlungen), erscheinen auf jeder betroffenen Fach-Seite/-Datei.

Voraussetzungen: das Geschwister-Repo `ausleihe-api` liegt **neben** diesem Repo
und enthält eine `.env` mit den IServ-Zugangsdaten. Abhängigkeiten via `uv sync`
im `sba-bestand`-Root (siehe `../README.md`).

## `trg_web.py` — die TRG-Website-Scraper

`generate_booklists.py` holt drei Zuordnungen von der öffentlichen TRG-Website
(Fachkonferenzleitungen, Aufgabenfelder, Kollegiumskürzel; für `--confirmation`
und `--mode aufgabenfeld`). Diese drei Scraper stehen seit 2026-09-05 in
`trg_web.py` statt im Hauptskript, weil sie so — anders als der Rest der Datei —
netzlos testbar sind: jede `fetch_*`-Funktion ist nur `_hole()` (Netzabruf) plus
eine reine `parse_*`-Funktion, und genau die `parse_*`-Funktionen prüft
`../tests/test_trg_web.py` gegen Fixtures unter `../tests/fixtures/trg/`
(shape-genaue HTML-Ausschnitte mit erfundenen Namen statt echter
Kollegiumsdaten).
