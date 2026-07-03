# Automatische Dokumentenerstellung aus Excel

**Demo-Projekt** – aus einer Excel-Tabelle wird per Knopfdruck ein ganzer Stapel
personalisierter Dokumente erzeugt (hier: Teilnahmebescheinigungen), jeweils als
**DOCX und PDF**. Alle Daten sind synthetisch, alle Firmen und Personen frei erfunden.

Das Prinzip funktioniert für beliebige Seriendokumente nach Vorlage:
Bescheinigungen, Arbeitsnachweise, Angebote, Serienbriefe – nur die Word-Vorlage
und die Spalten ändern sich.

## Vorher → Nachher

| Vorher | Nachher |
|---|---|
| Eine Excel-Liste mit 12 Teilnehmern (`input/Teilnehmerliste.xlsx`) | 24 fertige Dokumente in `output/` – pro Person eine Bescheinigung als DOCX **und** PDF, Dateiname = Teilnehmername |

Manuell wären das 12× Kopieren, Einsetzen, Speichern, PDF-Export – mit dem Tool
ein einziger Klick.

## Die Lösung ist eine einzige Excel-Datei

`Bescheinigungsgenerator.xlsm` enthält den kompletten VBA-Code. Es wird **kein
Python, keine Installation, kein Zusatzprogramm** benötigt – nur Microsoft Office
(Excel + Word), das in nahezu jedem Büro bereits vorhanden ist.

## Eingabeformat

**1. Teilnehmerliste (`.xlsx`)** – erstes Tabellenblatt, Kopfzeile in Zeile 1:

| Name | Kurs | Datum | Stunden |
|---|---|---|---|
| Lena Böhm | Datenschutz am Arbeitsplatz (DSGVO) | 18.03.2026 | 8 |

**2. Word-Vorlage (`.docx`)** – beliebig gestaltbares Dokument mit Platzhaltern:

- `{{NAME}}` – Name des Teilnehmers
- `{{KURS}}` – Kursbezeichnung
- `{{DATUM}}` – Kursdatum (wird immer als TT.MM.JJJJ ausgegeben)
- `{{STUNDEN}}` – Stundenumfang
- `{{AUSSTELLUNGSDATUM}}` – heutiges Datum (wird automatisch eingesetzt)

Beide Beispieldateien liegen in `input/` und werden durch `generate_input.py`
erzeugt (nur für die Demo-Daten nötig, nicht für den Betrieb des Tools).

## So funktioniert der VBA-Code

Der Code liegt im Modul [`src/modDokumente.bas`](src/modDokumente.bas)
(als Textdatei exportiert, damit er auf GitHub lesbar ist – identisch mit dem
Code in der `.xlsm`).

1. **Dateiauswahl:** Beim Start über den Button fragt das Tool nacheinander die
   Teilnehmerliste, die Word-Vorlage und den Zielordner ab.
2. **Einlesen:** Die Teilnehmerliste wird schreibgeschützt geöffnet, alle Zeilen
   werden in einem Rutsch in ein Array gelesen, danach wird die Datei sofort
   wieder geschlossen.
3. **Word im Hintergrund:** Das Tool startet Word unsichtbar (`CreateObject`,
   Late Binding – funktioniert daher ohne Verweis-Einstellungen mit jeder
   Office-Version).
4. **Pro Zeile ein Dokument:** Für jeden Teilnehmer wird ein neues Dokument auf
   Basis der Vorlage erzeugt (die Vorlage selbst bleibt unverändert), alle
   Platzhalter werden per Suchen/Ersetzen gefüllt, dann wird als
   `Bescheinigung_<Name>.docx` gespeichert und zusätzlich als PDF exportiert.
5. **Abschluss:** Word wird beendet, eine Meldung zeigt die Anzahl der
   erstellten Dokumente.

Leere Zeilen werden übersprungen; unzulässige Zeichen im Namen werden für den
Dateinamen automatisch entfernt. Datumswerte werden unabhängig von den
Regionaleinstellungen im deutschen Format ausgegeben.

## Verwendung

1. `Bescheinigungsgenerator.xlsm` öffnen, Makros erlauben.
2. Auf **„Bescheinigungen erstellen“** klicken.
3. Teilnehmerliste, Vorlage und Zielordner auswählen – fertig.

Für automatisierte Tests kann das Makro auch ohne Dialoge aufgerufen werden:
`ErstelleBescheinigungen listePfad, vorlagePfad, zielOrdner`.

## Voraussetzungen

- Microsoft Excel und Microsoft Word (Desktop, Windows)
- Makros müssen erlaubt sein

## Projektstruktur

```
04_Dokumentenerstellung/
├── Bescheinigungsgenerator.xlsm   ← das Tool (VBA, eine Datei)
├── src/modDokumente.bas           ← VBA-Code als lesbarer Export
├── generate_input.py              ← erzeugt die synthetischen Demo-Daten
├── input/                         ← Vorher: Teilnehmerliste + Word-Vorlage
└── output/                        ← Nachher: fertige Bescheinigungen (DOCX + PDF)
```
