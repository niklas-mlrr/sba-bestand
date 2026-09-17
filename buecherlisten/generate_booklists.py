#!/usr/bin/env python3
"""Bücherlisten eines Schuljahrs nach Fach, Verlag oder Jahrgang als PDF.

Holt alle Bücherlisten (eine je Jahrgang) eines Schuljahrs über die
IServ-Ausleihe-API und stellt sie fachweise neu zusammen: pro Fach eine
Tabelle "Leihbare Bücher" (Spalten Klasse, Titel, Verlag, ISBN, Neupreis,
Leihgebühr) und eine Tabelle "Selbst anzuschaffende Bücher" (dieselben Spalten
ohne Leihgebühr — genau wie in den offiziellen IServ-Bücherlisten-PDFs, an
deren Aufmachung sich dieses Layout orientiert). Bücher, die in mehreren
Jahrgängen angeboten werden (Mehrjahresbände), erscheinen einmal pro Fach mit
allen betroffenen Klassen (z.B. "5, 6") und werden zuerst nach der untersten,
dann nach der zweituntersten Klasse einsortiert.

Spaltenbreiten: Klasse/ISBN/Neupreis/Leihgebühr brechen nie um. Titel und
Verlag dürfen umbrechen, tun es aber nur, wenn der Inhalt nicht einzeilig in
die Tabellenbreite passt — dann wird der Platz zwischen beiden so aufgeteilt,
dass die Tabelle insgesamt möglichst wenig Zeilen braucht ("Leihgebühr" wird
dabei zu "Leihgeb.", "Klasse" bei Bedarf zu "Kl."). Alle Spaltenzwischenräume
sind danach gleich breit, jede Spalte ist maximal so breit wie ihr
tatsächlich benötigter Inhalt (Details: `buecherlisten_layout.md` im Wiki).

Es gibt für diesen Anwendungsfall keinen eigenen API-Endpunkt — die Zusammen-
stellung passiert clientseitig aus den regulären Bücherlisten-Daten
(GET /schoolyears/:id/booklists/:bl_id, siehe
~/wiki/wiki/30_projects/sba/ausleihe_api/api_reference.md). Die ISBN-Formatierung
mit Bindestrichen nutzt dieselbe isbnlib-Maskierung wie das Bestand-Tooling
unter "tools/bestand/"; auch dafür gibt es keinen API-Weg, die
Ausleihe-API liefert ISBNs immer ohne Trennzeichen.

Rein lesend (nur GET). Kein Schreibzugriff auf die IServ-Produktionsdatenbank.

Seit 2026-09-17 ist dieses Skript nur noch die Kommandozeile: Laden,
Layout und Erzeugen stehen in ``buecherlisten/core/`` (wie ``bestand/core/``
für die Bestandsliste), damit sba-dashboard dieselben PDFs erzeugen kann.

Seit 2026-09-17 gibt es neben den Fächern zwei weitere Ansichten (--view):
nach Verlag (Spalten Titel, Fach, Klasse, ISBN, Neupreis, Leihgebühr) und nach
Jahrgang (Spalten wie die IServ-Liste, Grundpaket und Wahlbereiche in einer
Tabelle). Kopf, Schriftgrößen und Tabellenbild sind überall dieselben; es
unterscheiden sich nur Spalten und Einleitungssatz. --student-list holt beim
Jahrgang stattdessen die Druckversion aus IServ.

Verwendung:
  python3 generate_booklists.py [--view fach|verlag|jahrgang] [--student-list]
                                 [--schoulyear 2026/2027]
                                 [--mode split|alphabet|aufgabenfeld]
                                 [--subjects "Fach1" "Fach2" ...] [--list-subjects]
                                 [--output-dir PFAD] [--confirmation]
                                 [--return-by DATUM] [--return-to KÜRZEL]
                                 [--duplex | --duplex-if-needed]

  --view           fach (Default) = eine Liste je Fach; verlag = je Verlag;
                    jahrgang = je Jahrgang. --confirmation und --mode
                    aufgabenfeld gibt es nur mit --view fach.
  --student-list   Nur mit --view jahrgang: statt der eigenen Liste die
                    IServ-Druckversion ("Schülerliste") holen; mehrere
                    Jahrgänge werden zu einer Datei zusammengehängt
                    (--mode split: eine Datei je Jahrgang), --duplex schiebt
                    Leerseiten ein. Dateiname mit "(Schülerliste)".
  --schoolyear     Schuljahr wie "2026/2027" (Default: laufendes Schuljahr)
  --mode           split       = eine PDF-Datei pro Fach,
                                  benannt "Bücherliste <Fach> <Schuljahr>.pdf"
                   alphabet    = eine PDF-Datei mit einer neuen Seite pro Fach
                                  (Fächer alphabetisch sortiert), benannt
                                  "Bücherliste Fächer <Schuljahr>.pdf"
                   aufgabenfeld = wie alphabet, aber die Fächer erst nach
                                  Aufgabenfeld (A/B/C, laut Tabelle auf
                                  trg-osterode.de/.../fachkonferenzleitungen/,
                                  ersatzweise .../unterricht-und-ganztags-
                                  angebot/faecher/, falls Erstere nicht mehr
                                  existiert) sortiert, dann alphabetisch;
                                  Fächer, die in keiner der beiden Quellen
                                  gelistet sind, kommen zuletzt, alphabetisch.
                                  Ist keine der beiden Quellen erreichbar,
                                  wird stattdessen alphabetisch sortiert.
                                  Ebenfalls benannt "Bücherliste Fächer
                                  <Schuljahr>.pdf" (wie alphabet)
                   (Default: alphabet)
  --subjects       Nur diese Fächer/Verlage/Jahrgänge aufnehmen (ein oder
                    mehrere Namen, exakt wie in der Bücherliste, z.B.
                    --subjects Deutsch Mathematik); Jahrgänge auch als Zahl
                    ("--grades 5 6"). Aliasse: --publishers, --grades.
                    Default: alles, was im Schuljahr vorkommt.
  --list-subjects  Nur die verfügbaren Fächer/Verlage/Jahrgänge der gewählten
                    Ansicht auflisten und beenden (keine PDF-Erzeugung).
                    Alias: --list.
  --output-dir     Zielordner für die PDF(s) (Default: dieser Skriptordner)
  --confirmation   Bestätigungs-Block (Ankreuzfelder + Ort/Datum/Unterschrift
                    Fachkonferenzleitung <Fach>) am Ende jeder Fach-Liste
                    ergänzen. Lädt zusätzlich live die Fach->Name-Zuordnung
                    von der TRG-Website (Fachkonferenzleitungen) und zeigt den
                    Namen mittig in der Kopfzeile; schlägt der Abruf fehl oder
                    ist das Fach dort nicht gelistet, steht dort ersatzweise
                    "Bestätigung". Hängt "Bestätigung " vor den Titel/
                    Dateinamen (z.B. "Bestätigung Bücherliste Deutsch
                    <Schuljahr>.pdf").
  --return-by      Rückgabedatum, das mittig im Kopf unter dem Kürzel steht.
                    Freier Text, also z.B. "08.09.2026" oder "Montag, den
                    08.09.2026". Nur mit --confirmation wirksam.
  --return-to      Kürzel, an das zurückgegeben wird; steht im Kopf neben dem
                    Rückgabedatum. Fehlt das Datum, lautet das Label darüber
                    "Rückgabe an" statt "an"; fehlt eines von beiden, entfällt
                    es ganz. Nur mit --confirmation wirksam.
  --duplex         Für doppelseitigen Druck vorbereiten: jedes Fach bekommt
                    nötigenfalls eine leere Endseite, damit seine Seitenzahl
                    gerade ist (sonst würde beim doppelseitigen Druck das
                    nächste Fach auf der Rückseite der letzten Seite des
                    vorherigen beginnen).
  --duplex-if-needed  Wie --duplex, aber nur wirksam, wenn mindestens ein Fach
                    von Natur aus (ohne jede Polsterung) mehr als eine Seite
                    braucht. Sind alle Fächer ohnehin einseitig, bleibt die
                    Ausgabe unverändert (keine Leerseiten). Schließt sich mit
                    --duplex gegenseitig aus.
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

_HERE = Path(__file__).parent
_ROOT = _HERE.parent

# sba-bestand hält keine eigenen Secrets: Die IServ-Credentials liegen in der
# ``.env`` des Geschwister-Repos ausleihe-api. Beide Repos werden nebeneinander
# geklont (``<irgendein-ordner>/{ausleihe-api,sba-bestand}``).
_API_ROOT = _ROOT.parent / "ausleihe-api"

# ``ausleihe`` kommt normalerweise aus dem Venv (editable-Install, siehe
# pyproject). Fallback auf das Geschwister-Repo, damit das Skript auch ohne
# Installation läuft (Nachfolger-Pfad: nur klonen, nichts installieren).
if _API_ROOT.is_dir():
    sys.path.insert(0, str(_API_ROOT))

# Dieses Skript liegt unterhalb des Repo-Roots (buecherlisten/) und importiert
# ``buecherlisten.trg_web`` als absoluten Paketimport (mirroring dessen, was
# tests/conftest.py über den Repo-Root tut) — ohne diesen Eintrag würde
# ``from buecherlisten.trg_web import ...`` beim Direktaufruf
# (``python3 generate_booklists.py``) mit ModuleNotFoundError scheitern, weil
# nur der buecherlisten/-Ordner selbst, nicht dessen Elternordner auf
# sys.path steht.
if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))

from dotenv import load_dotenv  # noqa: E402

load_dotenv(_API_ROOT / ".env")

from ausleihe import AusleiheClient  # noqa: E402
from ausleihe.exceptions import NotFoundError  # noqa: E402

# Laden, Layout und Erzeugen stehen seit 2026-09-17 in buecherlisten/core/.
from buecherlisten.core.daten import ANSICHTEN, lade_buecherdaten, waehle_gruppen  # noqa: E402
from buecherlisten.core.erzeugen import (  # noqa: E402
    erzeuge_buecherlisten_pdfs,
    erzeuge_schuelerlisten_pdfs,
)

# Wie die Gruppen einer Ansicht in Meldungen heißen.
_GRUPPEN_WORT = {"fach": "Fächer", "verlag": "Verlage", "jahrgang": "Jahrgänge"}


def main() -> None:
    parser = argparse.ArgumentParser(description="Bücherlisten nach Fach, Verlag oder Jahrgang als PDF.")
    parser.add_argument(
        "--view", choices=ANSICHTEN, default="fach",
        help="fach = eine Liste je Fach; verlag = je Verlag (Spalten Titel, Fach, Klasse, ISBN, "
             "Neupreis, Leihgebühr); jahrgang = je Jahrgang (Spalten wie die IServ-Liste). "
             "Default: fach",
    )
    parser.add_argument(
        "--student-list", action="store_true",
        help="Nur mit --view jahrgang: statt der eigenen Liste die IServ-Druckversion "
             "(Schülerliste) holen",
    )
    parser.add_argument("--schoolyear", default=None, help='Schuljahr, z.B. "2026/2027" (Default: laufendes)')
    parser.add_argument(
        "--mode", choices=["split", "alphabet", "aufgabenfeld"], default="alphabet",
        help="split = 1 PDF je Fach; alphabet = 1 PDF mit Seite pro Fach, alphabetisch "
             "sortiert; aufgabenfeld = wie alphabet, aber die Fächer erst nach "
             "Aufgabenfeld (laut TRG-Fachkonferenzleitungen-Seite, ersatzweise "
             "TRG-Fächerübersicht) sortiert, dann alphabetisch; Fächer ohne bekanntes "
             "Aufgabenfeld zuletzt, alphabetisch. Ist keine der beiden Quellen "
             "erreichbar, wird stattdessen alphabetisch sortiert. (Default: alphabet)",
    )
    parser.add_argument(
        "--subjects", "--publishers", "--grades", dest="subjects", nargs="+", default=None,
        metavar="NAME",
        help="Nur diese Fächer/Verlage/Jahrgänge aufnehmen, passend zu --view; Jahrgänge auch "
             'als Zahl ("5"). Default: alle',
    )
    parser.add_argument(
        "--list-subjects", "--list", dest="list_subjects", action="store_true",
        help="Nur die verfügbaren Fächer/Verlage/Jahrgänge (je nach --view) auflisten und beenden",
    )
    parser.add_argument("--output-dir", default=None, help="Zielordner (Default: dieser Skriptordner)")
    parser.add_argument(
        "--confirmation", action="store_true",
        help="Ankreuzfeld + Ort/Datum/Unterschrift Fachkonferenzleitung am Ende jeder Fach-Liste ergänzen",
    )
    parser.add_argument(
        "--return-by", default=None, metavar="DATUM",
        help='Rückgabedatum im Kopf der Bestätigung, frei formatierbar (z.B. "08.09.2026" '
             'oder "Montag, den 08.09.2026"); nur mit --confirmation',
    )
    parser.add_argument(
        "--return-to", default=None, metavar="KÜRZEL",
        help="Kürzel, an das zurückgegeben wird, im Kopf der Bestätigung; nur mit --confirmation",
    )
    duplex_group = parser.add_mutually_exclusive_group()
    duplex_group.add_argument(
        "--duplex", action="store_true",
        help="Für doppelseitigen Druck vorbereiten: jedes Fach bekommt bei Bedarf eine leere Endseite, "
             "damit seine Seitenzahl gerade ist",
    )
    duplex_group.add_argument(
        "--duplex-if-needed", action="store_true",
        help="Wie --duplex, aber nur wirksam, wenn mindestens ein Fach von Natur aus (ungepolstert) "
             "mehr als eine Seite braucht; sind alle Fächer ohnehin einseitig, bleibt die Ausgabe "
             "unverändert (keine Leerseiten)",
    )
    args = parser.parse_args()
    ansicht = args.view
    if ansicht != "fach" and (args.confirmation or args.mode == "aufgabenfeld"):
        parser.error("--confirmation und --mode aufgabenfeld gibt es nur mit --view fach")
    if args.student_list and ansicht != "jahrgang":
        parser.error("--student-list gibt es nur mit --view jahrgang")

    client = AusleiheClient(allow_writes=False)

    schoolyear_id = args.schoolyear or "current"
    try:
        daten = lade_buecherdaten(client, args.schoolyear)
    except NotFoundError:
        print(f"Fehler: Schuljahr nicht gefunden: {schoolyear_id}", file=sys.stderr)
        sys.exit(1)
    schoolyear_id = daten.schuljahr_id

    wort = _GRUPPEN_WORT[ansicht]
    subjects = daten.gruppen(ansicht)
    if not subjects:
        print(f"Keine Bücher für Schuljahr {schoolyear_id} gefunden.", file=sys.stderr)
        sys.exit(1)

    if args.list_subjects:
        for subject in subjects:
            print(subject)
        return

    if args.subjects:
        selected, unknown = waehle_gruppen(daten, ansicht, args.subjects)
        if unknown:
            print(
                f"Fehler: Unbekannte {wort} für Schuljahr {schoolyear_id}: {', '.join(unknown)}",
                file=sys.stderr,
            )
            print(f"Verfügbare {wort}: {', '.join(subjects)}", file=sys.stderr)
            sys.exit(1)
        subjects = selected

    if args.student_list:
        ergebnisse = erzeuge_schuelerlisten_pdfs(
            daten,
            lambda listen_id: client.admin.get_booklist_pdf(daten.schuljahr_id, listen_id),
            jahrgaenge=subjects,
            modus="split" if args.mode == "split" else "alphabet",
            doppelseitig=args.duplex or args.duplex_if_needed,
            nur_falls_noetig=args.duplex_if_needed,
        )
    else:
        ergebnisse = erzeuge_buecherlisten_pdfs(
            daten,
            ansicht=ansicht,
            faecher=subjects,
            modus=args.mode,
            bestaetigung=args.confirmation,
            rueckgabe_bis=args.return_by,
            rueckgabe_an=args.return_to,
            doppelseitig=args.duplex or args.duplex_if_needed,
            nur_falls_noetig=args.duplex_if_needed,
        )

    out_dir = Path(args.output_dir) if args.output_dir else _HERE
    out_dir.mkdir(parents=True, exist_ok=True)
    for ergebnis in ergebnisse:
        for warnung in ergebnis.warnungen:
            print(warnung, file=sys.stderr)
        out_path = out_dir / ergebnis.dateiname
        out_path.write_bytes(ergebnis.inhalt)
        print(f"PDF gespeichert: {out_path}")


if __name__ == "__main__":
    main()
