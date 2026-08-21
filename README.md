# sba-bestand — Excel-Tooling der Schulbuchausleihe

Zwei Werkzeuge, die die IServ-Ausleihe-API **nur lesend** (GET) abfragen und
daraus Dateien erzeugen:

| Ordner | Werkzeug | Ergebnis |
|--------|----------|----------|
| `bestand/` | `update_bestand_auto.py` | trägt Bestands-/Anmeldezahlen in die Excel-Bestands- und Nachbestellungsliste ein |
| `buecherlisten/` | `generate_booklists.py` | erzeugt die Bücherlisten-PDFs je Fach/Jahrgang |

Es wird **nie** nach IServ geschrieben.

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
