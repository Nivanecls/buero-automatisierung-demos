# Verarbeitung von Eingangsrechnungen (PDF → Excel)

**Demo-Projekt** – ein Ordner voller PDF-Eingangsrechnungen von verschiedenen
Lieferanten (unterschiedliche Layouts!) wird automatisch in **eine**
übersichtliche Excel-Tabelle überführt: Lieferant, Rechnungsnummer, Datum,
Netto, MwSt, Brutto, Zahlungsziel, Status. Überfällige Rechnungen erscheinen
rot. Alle Rechnungen sind synthetisch, alle Firmen frei erfunden.

## Vorher → Nachher

| Vorher | Nachher |
|---|---|
| `input/` – 9 PDF-Rechnungen, drei völlig verschiedene Layouts (klassischer Briefkopf, moderner Stil, schlichte Dienstleister-Rechnung), verschiedene Feldbezeichnungen | `output/Rechnungsuebersicht.xlsx` – eine Tabelle, alle Felder extrahiert, überfällige Posten rot, unsichere Extraktionen gelb („Prüfen"), Summenzeile |

Statt 9 PDFs zu öffnen und abzutippen: ein Klick.

## Die Lösung ist eine einzige Excel-Datei

`Rechnungserfassung.xlsm` enthält den kompletten VBA-Code. Kein Python, kein
OCR-Dienst, keine Cloud – nur Microsoft Office. **Word dient dabei als
PDF-Textkonverter:** Word (ab 2013) kann PDFs öffnen und in Text umwandeln;
das Tool steuert Word unsichtbar im Hintergrund.

## Eingabeformat

Ein Ordner mit textbasierten PDF-Rechnungen (keine Scans – für gescannte
Rechnungen wäre zusätzlich OCR nötig). Die Extraktion arbeitet mit
**Schlüsselwörtern und regulären Ausdrücken statt fester Koordinaten** und
versteht daher unterschiedliche Formulierungen:

- Rechnungsnummer: `Rechnungsnummer:`, `RECHNUNG Nr.`, `Rechnung R-…`
- Beträge: `Nettobetrag`, `Zwischensumme (netto)`, `Summe netto`;
  `MwSt`, `Umsatzsteuer`, `Mehrwertsteuer`; `Rechnungsbetrag`,
  `Gesamtbetrag`, `Brutto zu zahlen`
- Fälligkeit: `Zahlbar bis …`, `Zahlungsziel: …`, `Fällig am …`
- Deutsche Beträge (`1.234,56 €`) und Daten (`TT.MM.JJJJ`)

Die 9 Demo-PDFs erzeugt `generate_input.py` (reportlab, drei
Layout-Generatoren).

## So funktioniert der VBA-Code

Der Code liegt in [`src/modRechnungen.bas`](src/modRechnungen.bas) (lesbarer
Export, identisch mit dem Code in der `.xlsm`).

1. **Ordner durchlaufen:** alle `*.pdf` im gewählten Ordner.
2. **PDF → Text:** jede Datei wird von unsichtbarem Word geöffnet und
   konvertiert. Zwei Stolperfallen werden im Code abgefangen: der
   „PDF wird konvertiert"-Hinweis wird deaktiviert (sonst wartet der
   unsichtbare Prozess ewig auf einen unsichtbaren Klick), und neben dem
   Haupttext werden auch die Kopfzeilen-Bereiche ausgelesen – Words
   PDF-Konverter legt Briefköpfe vom oberen Seitenrand dort ab.
3. **Felder extrahieren:** reguläre Ausdrücke (VBScript.RegExp) suchen
   Schlüsselwort-Varianten und greifen den zugehörigen Wert – tolerant
   gegenüber Zeilenumbrüchen und Tabellenstruktur im konvertierten Text.
   Beträge und Daten werden locale-unabhängig geparst (funktioniert auf
   jedem Windows, egal welche Regionaleinstellung).
4. **Plausibilitätsprüfung:** Netto + MwSt = Brutto (±1 Cent). Stimmt etwas
   nicht oder fehlt ein Feld → Zeile gelb, Status **Prüfen** – das Tool rät
   nicht, es markiert.
5. **Übersicht:** formatierte Tabelle mit Autofilter, Summenzeile und
   Statusspalte: **Offen** / **Überfällig** (rot) / **Prüfen** (gelb).

## Verwendung

1. `Rechnungserfassung.xlsm` öffnen, Makros erlauben.
2. Auf **„Rechnungen erfassen"** klicken.
3. PDF-Ordner und Zieldatei wählen – fertig.

Für automatisierte Tests: `ErfasseRechnungen pdfOrdner, zielDatei`
(ohne Dialoge).

## Voraussetzungen

- Microsoft Excel und Microsoft Word ab 2013 (Desktop, Windows), Makros erlaubt
- Textbasierte PDFs (keine gescannten Bilder)

## Projektstruktur

```
01_Eingangsrechnungen/
├── Rechnungserfassung.xlsm    ← das Tool (VBA, eine Datei)
├── src/modRechnungen.bas      ← VBA-Code als lesbarer Export
├── generate_input.py          ← erzeugt die 9 synthetischen PDF-Rechnungen
├── input/                     ← Vorher: PDF-Rechnungen (3 Layouts)
└── output/                    ← Nachher: Rechnungsuebersicht.xlsx
```
