# Büroautomatisierung mit Excel/VBA – Demo-Portfolio

Vier Demo-Projekte für typische Büroaufgaben kleiner und mittlerer Unternehmen.
Jedes Tool ist **eine einzige Excel-Datei (.xlsm)** – keine Installation, kein
Python, keine Cloud, keine Anbindung an ein bestimmtes ERP/CRM. Benötigt wird
nur Microsoft Office, das ohnehin auf fast jedem Büro-PC vorhanden ist.

**Alle Daten sind synthetisch.** Sämtliche Firmen, Personen, Adressen und
Beträge sind frei erfunden und werden pro Projekt durch ein Python-Skript
(`generate_input.py`) erzeugt – die Skripte gehören zur Demo, nicht zum Tool.

| # | Projekt | Vorher → Nachher |
|---|---|---|
| 1 | [Verarbeitung von Eingangsrechnungen](01_Eingangsrechnungen/) | Ordner mit PDF-Rechnungen (verschiedene Layouts) → eine Excel-Übersicht, Überfälliges rot markiert |
| 2 | [Bereinigung von Kundenstammdaten](02_Kundenstammdaten/) | verschmutzte Kundenliste (Dubletten, Tippfehler, Formatchaos) → saubere Liste + lückenloses Änderungsprotokoll |
| 3 | [Automatisiertes Mahnwesen](03_Mahnwesen/) | Offene-Posten-Liste (CSV) → versandfertige Mahnschreiben in 3 Eskalationsstufen + Ampel-Übersicht |
| 4 | [Automatische Dokumentenerstellung](04_Dokumentenerstellung/) | Excel-Teilnehmerliste + Word-Vorlage → Stapel personalisierter Bescheinigungen als DOCX und PDF |

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

## Technische Gemeinsamkeiten

- **Ein Klick, klare Dialoge:** Datei/Ordner auswählen, Ergebnismeldung mit
  Zahlen („X Rechnungen verarbeitet, Y überfällig").
- **Locale-fest:** deutsche Beträge (`1.234,56 €`) und Daten (`TT.MM.JJJJ`)
  werden unabhängig von den Windows-Regionaleinstellungen korrekt verarbeitet.
- **Word-Automatisierung** (Projekte 1, 3, 4): Word läuft unsichtbar im
  Hintergrund – als PDF-Textkonverter bzw. Dokumentgenerator. Late Binding,
  daher ohne Verweis-Einstellungen mit jeder Office-Version ab 2013 lauffähig.
- **Nachvollziehbar statt magisch:** Unsicheres wird markiert (gelb,
  „Prüfen"), Änderungen werden protokolliert, nichts wird ungefragt versendet
  oder gelöscht.
- Jedes Makro läuft auch parameterisiert ohne Dialoge – so wurden alle vier
  Tools automatisiert getestet.

## Voraussetzungen

Microsoft Office (Excel, für Projekte 1/3/4 zusätzlich Word) unter Windows,
Makros erlaubt. Python wird **nur** zum Erzeugen der Demo-Daten gebraucht
(`pip install reportlab openpyxl python-docx`).
