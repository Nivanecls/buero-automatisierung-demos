# Automatische Dokumentenerstellung aus Excel

Aus einer Excel-Tabelle wird per Knopfdruck ein Stapel personalisierter
Dokumente erzeugt – hier Teilnahmebescheinigungen, jeweils als DOCX und PDF.
Alle Daten sind synthetisch, alle Firmen und Personen frei erfunden.

Das Tool braucht Excel **und Word**: Word füllt im Hintergrund die Vorlage
aus und übernimmt den PDF-Export.

Das Prinzip funktioniert für beliebige Seriendokumente nach Vorlage –
Bescheinigungen, Arbeitsnachweise, Angebote, Serienbriefe. Es ändern sich nur
die Word-Vorlage und die Spalten der Tabelle.

## Vorher → Nachher

`input/Teilnehmerliste.xlsx` hat 12 Zeilen. In `output/` liegen danach 24
Dateien: pro Person eine Bescheinigung als DOCX und PDF, der Dateiname ist
der Teilnehmername. Von Hand wären das 12-mal Kopieren, Einsetzen, Speichern
und PDF-Export.

## Eingabeformat

**Teilnehmerliste (`.xlsx`)** – erstes Tabellenblatt, Kopfzeile in Zeile 1:

| Name | Kurs | Datum | Stunden |
|---|---|---|---|
| Lena Böhm | Datenschutz am Arbeitsplatz (DSGVO) | 18.03.2026 | 8 |

**Word-Vorlage (`.docx`)** – frei gestaltbares Dokument mit Platzhaltern:

- `{{NAME}}` – Name des Teilnehmers
- `{{KURS}}` – Kursbezeichnung
- `{{DATUM}}` – Kursdatum (Ausgabe immer als TT.MM.JJJJ)
- `{{STUNDEN}}` – Stundenumfang
- `{{AUSSTELLUNGSDATUM}}` – heutiges Datum, wird automatisch eingesetzt

Beide Beispieldateien liegen in `input/` und werden durch `generate_input.py`
erzeugt; für den Betrieb des Tools ist das Skript nicht nötig.

## So funktioniert der VBA-Code

Der Code liegt in [`src/modDokumente.bas`](src/modDokumente.bas), lesbarer
Export, identisch mit dem Code in der `.xlsm`.

1. Beim Start über den Button fragt das Tool nacheinander Teilnehmerliste,
   Word-Vorlage und Zielordner ab.
2. Die Liste wird schreibgeschützt geöffnet, komplett in ein Array gelesen
   und sofort wieder geschlossen.
3. Word startet unsichtbar (`CreateObject`, Late Binding – läuft daher ohne
   Verweis-Einstellungen mit jeder Office-Version).
4. Für jede Zeile entsteht ein neues Dokument auf Basis der Vorlage (die
   Vorlage selbst bleibt unverändert). Die Platzhalter werden per
   Suchen/Ersetzen gefüllt, gespeichert wird als
   `Bescheinigung_<Name>.docx` plus PDF-Export.
5. Word wird beendet, eine Meldung zeigt die Anzahl der erstellten Dokumente.

Leere Zeilen werden übersprungen, unzulässige Zeichen fliegen aus dem
Dateinamen, Datumswerte werden unabhängig von den Regionaleinstellungen im
deutschen Format ausgegeben.

## Verwendung

1. `Bescheinigungsgenerator.xlsm` öffnen, Makros erlauben.
2. Auf "Bescheinigungen erstellen" klicken.
3. Teilnehmerliste, Vorlage und Zielordner auswählen.

Für Tests ohne Dialoge:
`ErstelleBescheinigungen listePfad, vorlagePfad, zielOrdner`.

## Voraussetzungen

- Microsoft Excel und Microsoft Word (Windows), Makros erlaubt

## Projektstruktur

```
04_Dokumentenerstellung/
├── Bescheinigungsgenerator.xlsm   ← das Tool (VBA, eine Datei)
├── src/modDokumente.bas           ← VBA-Code als lesbarer Export
├── generate_input.py              ← erzeugt die synthetischen Demo-Daten
├── input/                         ← Vorher: Teilnehmerliste + Word-Vorlage
└── output/                        ← Nachher: fertige Bescheinigungen (DOCX + PDF)
```
