# Verarbeitung von Eingangsrechnungen (PDF → Excel)

Ein Ordner voller PDF-Rechnungen von verschiedenen Lieferanten, jede anders
aufgebaut, wird in eine Excel-Tabelle überführt: Lieferant, Rechnungsnummer,
Datum, Netto, MwSt, Brutto, Zahlungsziel, Status. Überfällige Rechnungen sind
rot markiert. Alle Rechnungen sind synthetisch, alle Firmen frei erfunden.

Das Tool braucht Excel **und Word**: Word wird im Hintergrund gestartet und
dient nur als PDF-Textkonverter, es öffnet dabei keine sichtbaren Fenster.

## Vorher → Nachher

`input/` enthält 9 PDF-Rechnungen in drei unterschiedlichen Layouts
(klassischer Briefkopf, moderner Stil, schlichte Dienstleister-Rechnung).
`output/Rechnungsuebersicht.xlsx` enthält alle Felder extrahiert, mit
Autofilter, Summenzeile und Statusspalte (Offen / Überfällig / Prüfen).

## Eingabeformat

Ein Ordner mit textbasierten PDF-Rechnungen – keine Scans, dafür wäre
zusätzlich OCR nötig. Die Extraktion sucht nach Schlüsselwörtern und
regulären Ausdrücken statt nach festen Koordinaten und erkennt deshalb
unterschiedliche Formulierungen:

- Rechnungsnummer: `Rechnungsnummer:`, `RECHNUNG Nr.`, `Rechnung R-…`
- Beträge: `Nettobetrag` / `Zwischensumme (netto)` / `Summe netto`,
  `MwSt` / `Umsatzsteuer` / `Mehrwertsteuer`, `Rechnungsbetrag` /
  `Gesamtbetrag` / `Brutto zu zahlen`
- Fälligkeit: `Zahlbar bis …`, `Zahlungsziel: …`, `Fällig am …`
- deutsche Beträge (`1.234,56 €`) und Daten (`TT.MM.JJJJ`)

Die 9 Demo-PDFs erzeugt `generate_input.py` (reportlab, drei
Layout-Generatoren).

## So funktioniert der VBA-Code

Der Code liegt in [`src/modRechnungen.bas`](src/modRechnungen.bas), lesbarer
Export, identisch mit dem Code in der `.xlsm`.

1. Alle `*.pdf` im gewählten Ordner werden nacheinander verarbeitet.
2. Jede Datei wird von unsichtbarem Word geöffnet und in Text umgewandelt.
   Zwei Dinge werden dabei im Code abgefangen: der "PDF wird
   konvertiert"-Hinweis wird deaktiviert, sonst wartet der unsichtbare
   Word-Prozess auf einen Klick, den niemand sieht; und neben dem Haupttext
   werden auch die Kopfzeilen-Bereiche ausgelesen, weil Words PDF-Konverter
   den oberen Seitenrand (Briefkopf) dort ablegt statt im Fließtext.
3. Reguläre Ausdrücke (VBScript.RegExp) suchen die Schlüsselwort-Varianten
   und lesen den zugehörigen Wert, tolerant gegenüber Zeilenumbrüchen und
   Tabellenstruktur im konvertierten Text. Beträge und Daten werden ohne
   `CDbl`/`CDate` geparst, damit die Windows-Regionaleinstellung keine Rolle
   spielt.
4. Plausibilitätsprüfung: Netto + MwSt muss (bis auf 1 Cent) Brutto ergeben.
   Stimmt das nicht oder fehlt ein Feld, wird die Zeile gelb markiert und als
   "Prüfen" gekennzeichnet, statt einen falschen Wert einzutragen.
5. Die Übersichtstabelle bekommt Autofilter, Summenzeile und die Statusspalte
   Offen / Überfällig (rot) / Prüfen (gelb).

## Verwendung

1. `Rechnungserfassung.xlsm` öffnen, Makros erlauben.
2. Auf "Rechnungen erfassen" klicken.
3. PDF-Ordner und Zieldatei wählen.

Für Tests ohne Dialoge: `ErfasseRechnungen pdfOrdner, zielDatei`.

## Voraussetzungen

- Microsoft Excel und Microsoft Word ab 2013 (Windows), Makros erlaubt
- textbasierte PDFs, keine gescannten Bilder

## Projektstruktur

```
01_Eingangsrechnungen/
├── Rechnungserfassung.xlsm    ← das Tool (VBA, eine Datei)
├── src/modRechnungen.bas      ← VBA-Code als lesbarer Export
├── generate_input.py          ← erzeugt die 9 synthetischen PDF-Rechnungen
├── input/                     ← Vorher: PDF-Rechnungen (3 Layouts)
└── output/                    ← Nachher: Rechnungsuebersicht.xlsx
```
