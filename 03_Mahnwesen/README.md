# Automatisiertes Mahnwesen

Aus einer Offene-Posten-Liste (CSV) entstehen fertige Mahnschreiben als
Word-Dokumente, mit Eskalationsstufe je nach Verzugsdauer, dazu eine
Excel-Übersicht mit Ampelfarben. Alle Daten sind synthetisch, alle Firmen
frei erfunden.

Das Tool braucht Excel **und Word** (die Briefe werden als .docx erzeugt).
Versendet wird nichts – die Schreiben liegen als Dateien bereit, der Versand
bleibt eine manuelle Entscheidung.

## Eskalationslogik

| Verzug | Stufe | Ton |
|---|---|---|
| 1–7 Tage | Stufe 1 | freundliche Zahlungserinnerung |
| 8–21 Tage | Stufe 2 | Mahnung mit Fristsetzung |
| ab 22 Tagen | Stufe 3 | letzte Mahnung vor rechtlichen Schritten |

Bezahlte und noch nicht fällige Rechnungen bekommen kein Schreiben.

## Vorher → Nachher

`input/Offene_Posten.csv` enthält 18 Rechnungen: bezahlte, offene, 2 bis 41
Tage überfällige. In `output/` liegen danach 11 Mahnschreiben (der Dateiname
zeigt Stufe und Kunde) und `Mahnuebersicht.xlsx` – alle Posten mit
Verzugstagen und Stufe, farblich markiert, darunter die Zusammenfassung:
Schreiben je Stufe, offene und überfällige Gesamtsumme.

## Eingabeformat

CSV, Semikolon-getrennt, UTF-8, deutsche Zahlen- und Datumsformate – das
übliche Exportformat deutscher Buchhaltungsprogramme:

```
Kunde;Rechnungsnummer;Rechnungsdatum;Betrag;Faelligkeitsdatum;Zahlungsstatus
Bäckerei Vogt;AR-2026-1000;02.06.2026;1.154,61;02.07.2026;offen
```

`Zahlungsstatus` ist `offen` oder `bezahlt`. Die Spaltenreihenfolge ist egal,
die Spalten werden über die Kopfzeile gefunden. Die Demo-Datei erzeugt
`generate_input.py`.

## So funktioniert der VBA-Code

Der Code liegt in [`src/modMahnwesen.bas`](src/modMahnwesen.bas), lesbarer
Export, identisch mit dem Code in der `.xlsm`.

1. Die CSV wird über einen UTF-8-Stream gelesen; Beträge (`1.234,56`) und
   Daten (`TT.MM.JJJJ`) werden ohne `CDbl`/`CDate` geparst, damit die
   Windows-Regionaleinstellung keine Rolle spielt.
2. Verzugstage = heute minus Fälligkeitsdatum, daraus die Stufe nach der
   Tabelle oben.
3. Word läuft unsichtbar im Hintergrund und erzeugt pro überfälliger Rechnung
   ein Schreiben mit Absender, Empfänger, Datum, Betreff und stufengerechtem
   Text. Rechnungsnummer, Betrag und neue Frist (Stufe 1: +10 Tage, Stufe 2:
   +7, Stufe 3: +5) sind bereits eingesetzt.
4. Die Übersicht listet alle Posten mit Verzugstagen und Stufe, gelb/orange/
   rot markiert, bezahlte Posten ausgegraut, darunter die Zusammenfassung.

Der Briefabsender steht als Konstante im Modulkopf und ist schnell an die
eigene Firma angepasst.

## Verwendung

1. `Mahnwesen.xlsm` öffnen, Makros erlauben.
2. Auf "Mahnungen erstellen" klicken.
3. CSV und Zielordner wählen. Die Meldung am Ende zeigt die Anzahl der
   Schreiben je Stufe und die überfällige Summe.

Für Tests ohne Dialoge: `ErstelleMahnungen csvPfad, zielOrdner`.

## Voraussetzungen

- Microsoft Excel und Microsoft Word (Windows), Makros erlaubt

## Projektstruktur

```
03_Mahnwesen/
├── Mahnwesen.xlsm            ← das Tool (VBA, eine Datei)
├── src/modMahnwesen.bas      ← VBA-Code als lesbarer Export
├── generate_input.py         ← erzeugt die synthetische Offene-Posten-Liste
├── input/                    ← Vorher: Offene_Posten.csv
└── output/                   ← Nachher: Mahnschreiben + Mahnuebersicht.xlsx
```
