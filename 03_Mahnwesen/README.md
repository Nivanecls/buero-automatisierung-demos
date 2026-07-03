# Automatisiertes Mahnwesen

**Demo-Projekt** – aus einer Offene-Posten-Liste (CSV) entstehen automatisch
fertige Mahnschreiben als Word-Dokumente, mit **automatischer Eskalation** nach
Tagen des Verzugs, plus eine farbcodierte Excel-Übersicht. Alle Daten sind
synthetisch, alle Firmen frei erfunden.

**Wichtig:** Es wird nichts versendet. Das Tool erzeugt geprüfbare Entwürfe –
der Versand bleibt eine bewusste, manuelle Entscheidung.

## Eskalationslogik

| Verzug | Stufe | Ton |
|---|---|---|
| 1–7 Tage | Stufe 1 | freundliche Zahlungserinnerung |
| 8–21 Tage | Stufe 2 | bestimmte Mahnung mit Fristsetzung |
| ab 22 Tagen | Stufe 3 | letzte Mahnung vor rechtlichen Schritten / Inkasso |

Bezahlte und noch nicht fällige Rechnungen bekommen kein Schreiben.

## Vorher → Nachher

| Vorher | Nachher |
|---|---|
| `input/Offene_Posten.csv` – 18 Rechnungen, unsortiert: bezahlt, offen, 2 bis 41 Tage überfällig | `output/` – 11 versandfertige Mahnschreiben (Word), Dateiname zeigt Stufe und Kunde, plus `Mahnuebersicht.xlsx` mit Ampelfarben und Zusammenfassung (Schreiben je Stufe, überfällige Gesamtsumme) |

## Die Lösung ist eine einzige Excel-Datei

`Mahnwesen.xlsm` enthält den kompletten VBA-Code. Kein Python, keine
Installation – nur Microsoft Office.

## Eingabeformat

CSV, Semikolon-getrennt, UTF-8 (typisches ERP-Exportformat), deutsche
Zahlen- und Datumsformate:

```
Kunde;Rechnungsnummer;Rechnungsdatum;Betrag;Faelligkeitsdatum;Zahlungsstatus
Bäckerei Vogt;AR-2026-1000;02.06.2026;1.154,61;02.07.2026;offen
```

`Zahlungsstatus`: `offen` oder `bezahlt`. Die Spaltenreihenfolge ist egal –
die Spalten werden über die Kopfzeile gefunden. Die Demo-Datei erzeugt
`generate_input.py`.

## So funktioniert der VBA-Code

Der Code liegt in [`src/modMahnwesen.bas`](src/modMahnwesen.bas) (lesbarer
Export, identisch mit dem Code in der `.xlsm`).

1. **CSV einlesen:** über einen UTF-8-Stream – funktioniert unabhängig von den
   Windows-Regionaleinstellungen; deutsche Beträge (`1.234,56`) und Daten
   (`TT.MM.JJJJ`) werden locale-sicher geparst.
2. **Stufe bestimmen:** Verzugstage = heute − Fälligkeitsdatum, daraus die
   Eskalationsstufe nach obiger Tabelle.
3. **Briefe erzeugen:** Word läuft unsichtbar im Hintergrund; pro überfälliger
   Rechnung entsteht ein formatiertes Schreiben mit Absender, Empfänger, Datum,
   Betreff und stufengerechtem Text – Rechnungsnummer, Betrag und neue Frist
   (Stufe 1: +10 Tage, Stufe 2: +7, Stufe 3: +5) sind bereits eingesetzt.
4. **Übersicht schreiben:** alle Posten mit Verzugstagen und Stufe,
   Ampelfarben (gelb/orange/rot), bezahlte Posten ausgegraut, darunter die
   Zusammenfassung: Anzahl Schreiben je Stufe, offene und überfällige
   Gesamtsumme.

Der Briefabsender ist im Modulkopf als Konstante hinterlegt und in einer
Minute an die eigene Firma angepasst.

## Verwendung

1. `Mahnwesen.xlsm` öffnen, Makros erlauben.
2. Auf **„Mahnungen erstellen"** klicken.
3. CSV und Zielordner wählen – die Ergebnismeldung zeigt die Anzahl der
   Schreiben je Stufe und die überfällige Summe.

Für automatisierte Tests: `ErstelleMahnungen csvPfad, zielOrdner`
(ohne Dialoge).

## Voraussetzungen

- Microsoft Excel und Microsoft Word (Desktop, Windows), Makros erlaubt

## Projektstruktur

```
03_Mahnwesen/
├── Mahnwesen.xlsm            ← das Tool (VBA, eine Datei)
├── src/modMahnwesen.bas      ← VBA-Code als lesbarer Export
├── generate_input.py         ← erzeugt die synthetische Offene-Posten-Liste
├── input/                    ← Vorher: Offene_Posten.csv
└── output/                   ← Nachher: Mahnschreiben + Mahnuebersicht.xlsx
```
