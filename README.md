# sba-bestand — Excel-Tooling der Schulbuchausleihe

Eine Bibliothek und zwei Werkzeuge, die die IServ-Ausleihe-API **nur lesend**
(GET) abfragen und daraus Dateien erzeugen:

| Ordner | Werkzeug | Ergebnis |
|--------|----------|----------|
| `bestand/core/` | Bibliothek | Excel-Raster lesen, IServ-Snapshot holen, Zahlen eintragen — netzfrei testbar |
| `bestand/` | `update_bestand_auto.py` | Kommandozeilen-Schale um `core/`: trägt Bestands-/Anmeldezahlen in die Excel-Liste ein |
| `buecherlisten/` | `generate_booklists.py` | erzeugt die Bücherlisten-PDFs je Fach/Jahrgang |

Es wird **nie** nach IServ geschrieben.

## `bestand/core/` — die Bibliothek

Seit 2026-09-04 steckt die Logik nicht mehr im Skript, sondern in einem Paket.
Das war nötig, damit [`sba-dashboard`](https://github.com/niklas-mlrr/sba-dashboard)
dieselbe Excel-Behandlung benutzt, statt sie ein zweites Mal nachzubauen.

| Modul | Inhalt |
|-------|--------|
| `core/config.py` | `config.json` laden und prüfen |
| `core/grid.py` | Excel-Struktur: Fachblöcke, Mehrjahresbänder, „nicht angeboten"-Sperrflächen |
| `core/iserv.py` | `Snapshot` aus IServ — der Client wird **injiziert**, nie selbst gebaut |
| `core/update.py` | Snapshot anwenden, Blatt „zu Bestellen" neu aufbauen |
| `core/testing.py` | synthetisches Prüf-Workbook + Fake-IServ, von beiden Repos benutzt |

Nichts davon liest `os.environ`, lädt eine `.env`, parst Argumente oder schreibt
nach stdout. Die Tests laufen ohne Netz und ohne die echte Mappe:

```bash
uv sync --all-groups
uv run pytest
```

`tests/test_cli_golden.py` friert die Konsolenausgabe von
`update_bestand_auto.py` ein — der Refactor hat sie zeichengleich gelassen.

Die Struktur-Befunde, die den Entwurf bestimmen (keine Bezahlt-Spalte,
Formelspalten, Merge-Topologie, Sperrflächen), stehen in
[`sba-dashboard/docs/architektur.md`](https://github.com/niklas-mlrr/sba-dashboard/blob/main/docs/architektur.md).

## Voraussetzung: Geschwister-Layout

Dieses Repo braucht das Repo [`ausleihe-api`](https://github.com/niklas-mlrr/ausleihe-api)
**direkt daneben** — dort liegen der Python-Client (`ausleihe`) und die `.env`
mit den IServ-Zugangsdaten:

```
<irgendein-ordner>/
├── ausleihe-api/        <- Client + .env (Credentials)
└── sba-bestand/         <- dieses Repo
```

`sba-bestand` hält bewusst **keine eigenen Secrets**. Fehlt `ausleihe-api/.env`,
brechen die Skripte mit einer Meldung zu fehlenden `ISERV_*`-Variablen ab.

## Einrichtung

```bash
git clone https://github.com/niklas-mlrr/ausleihe-api.git
git clone https://github.com/niklas-mlrr/sba-bestand.git
cd sba-bestand
uv sync            # zieht ausleihe-api als editable-Install mit
```

Ohne `uv` genügt auch ein einfaches `pip install openpyxl isbnlib python-dotenv
reportlab requests` — die Skripte finden `ausleihe` dann über einen Fallback auf
den Geschwister-Ordner.

## Benutzung

Siehe `bestand/README.md` und `buecherlisten/README.md`. Die vollständige
Anleitung für Nachfolger liegt in
[`ausleihe-ausgabe/docs/nachfolge-anleitung.md`](https://github.com/niklas-mlrr/ausleihe-ausgabe/blob/main/docs/nachfolge-anleitung.md)
(Teil 3).

## Herkunft

Bis 2026-08-21 lagen beide Werkzeuge im Repo `ausleihe-api` (als
`bestand- und nachbestellungen/` und `buecherlisten-nach-fach/`). Sie wurden
herausgelöst, weil `ausleihe-api` eine wiederverwendbare Bibliothek ist,
während dies hier schulspezifische Anwendungen sind.
