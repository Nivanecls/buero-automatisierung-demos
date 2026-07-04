# Büroautomatisierung mit Excel/VBA – Demo-Portfolio

Vier Demo-Projekte für typische Büroaufgaben kleiner und mittlerer Unternehmen.
Jede Lösung ist eine einzelne Excel-Datei (.xlsm) mit VBA-Code, ohne Anbindung
an ein bestimmtes ERP oder CRM. Drei der vier Tools steuern zusätzlich Word im
Hintergrund (PDF-Text lesen, Briefe und Dokumente erzeugen) – nur die
Datenbereinigung in Projekt 2 kommt ohne Word aus. Genaueres steht jeweils
unter "Voraussetzungen" im README des Projekts.

Alle Daten sind synthetisch: Firmen, Personen, Adressen und Beträge sind frei
erfunden und werden pro Projekt durch `generate_input.py` erzeugt. Diese
Skripte gehören zur Demo, nicht zur eigentlichen Lösung.

| # | Projekt | Braucht | Vorher → Nachher |
|---|---|---|---|
| 1 | [Verarbeitung von Eingangsrechnungen](01_Eingangsrechnungen/) | Excel + Word | Ordner mit PDF-Rechnungen unterschiedlicher Layouts → eine Excel-Übersicht, überfällige Posten rot |
| 2 | [Bereinigung von Kundenstammdaten](02_Kundenstammdaten/) | Excel | verschmutzte Kundenliste (Dubletten, Tippfehler, Formatchaos) → saubere Liste plus Änderungsprotokoll |
| 3 | [Automatisiertes Mahnwesen](03_Mahnwesen/) | Excel + Word | Offene-Posten-Liste (CSV) → Mahnschreiben in drei Eskalationsstufen plus Übersicht |
| 4 | [Automatische Dokumentenerstellung](04_Dokumentenerstellung/) | Excel + Word | Excel-Teilnehmerliste und Word-Vorlage → personalisierte Dokumente als DOCX und PDF |

## Aufbau jedes Projekts

```
0X_Projektname/
├── <Tool>.xlsm          ← die Lösung: ein Excel-File mit Start-Button
├── src/<Modul>.bas      ← der VBA-Code als lesbare Textdatei (identisch mit der .xlsm)
├── generate_input.py    ← erzeugt die synthetischen Demo-Eingangsdaten
├── input/               ← „Vorher“
├── output/              ← „Nachher“
└── README.md            ← Funktionsweise, Eingabeformat, Voraussetzungen
```

## Gemeinsamkeiten

Jedes Tool startet über einen Button, fragt Datei bzw. Ordner ab und meldet am
Ende ein Ergebnis mit Zahlen (z.B. "9 Rechnungen verarbeitet, 5 überfällig").
Deutsche Zahlen- und Datumsformate (`1.234,56 €`, `TT.MM.JJJJ`) werden
unabhängig von den Windows-Regionaleinstellungen korrekt verarbeitet. Wo eine
Extraktion oder Zuordnung unsicher ist, wird die Zeile gelb markiert statt
geraten. Jedes Makro lässt sich auch mit festen Pfaden als Parameter ohne
Dialoge aufrufen – so wurden alle vier Tools automatisiert getestet.

## Voraussetzungen

Windows mit Microsoft Excel, für die Projekte 1, 3 und 4 zusätzlich Word;
Makros müssen erlaubt sein. Python wird nur zum Erzeugen der Demo-Daten
gebraucht (`pip install reportlab openpyxl python-docx`).
