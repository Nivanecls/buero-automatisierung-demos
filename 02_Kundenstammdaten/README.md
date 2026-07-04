# Bereinigung von Kundenstammdaten

Eine "verschmutzte" Kundenliste wird bereinigt: Duplikate mit abweichender
Schreibweise, uneinheitliche Telefon- und E-Mail-Formate, Leerzeichen-Chaos
und fehlende Pflichtfelder. Zu jeder Änderung entsteht ein Protokolleintrag
(vorher → nachher → Grund). Alle Daten sind synthetisch, alle Firmen frei
erfunden.

Das Tool braucht nur Excel, kein Word und keine weitere Software.

## Vorher → Nachher

`input/Kundenliste_roh.xlsx` hat 48 Zeilen mit Dubletten (Müller GmbH /
Mueller GmbH / MÜLLER GMBH), kaputten E-Mails, fünf verschiedenen
Telefonformaten und Lücken. `output/Kundenliste_bereinigt.xlsx` enthält 40
saubere Datensätze und ein zweites Blatt Änderungsprotokoll mit 45
dokumentierten Korrekturen.

Das Protokoll ist Teil des Ergebnisses, nicht nur die saubere Liste: jede
Änderung lässt sich nachvollziehen.

## Eingabeformat

Excel-Datei, erstes Tabellenblatt, Kopfzeile in Zeile 1:

| Firma | Ansprechpartner | Straße | PLZ | Ort | Telefon | E-Mail |
|---|---|---|---|---|---|---|
| Müller GmbH | Hans Weber | Hauptstraße 12 | 01067 | Dresden | 0351/456789 | Info@Mueller,example |

Die Demo-Datei erzeugt `generate_input.py`.

## So funktioniert der VBA-Code

Der Code liegt in [`src/modStammdaten.bas`](src/modStammdaten.bas), lesbarer
Export, identisch mit dem Code in der `.xlsm`.

**Feldbereinigung** (jede Korrektur wird protokolliert): führende, folgende
und doppelte Leerzeichen raus; `MÜLLER GMBH` / `müller gmbh` → `Müller GmbH`
(Großschreibung wird nur angefasst, wenn der Name komplett groß oder komplett
klein ist, Rechtsformen wie GmbH/KG/AG/UG/e.K. werden korrekt gesetzt);
Telefonnummern auf `+49 Vorwahl Nummer` vereinheitlicht (erkennt `0351/…`,
`(0351) …`, `0049…`, reine Ziffernfolgen); E-Mails kleingeschrieben,
Leerzeichen entfernt, Komma zu Punkt korrigiert und per Regex validiert –
weiterhin ungültige Adressen werden markiert statt geraten.

**Duplikatsuche.** Für jede Firma entsteht ein Vergleichsschlüssel:
Kleinschreibung, Umlaut-Transliteration (ü→ue, ß→ss), Rechtsformen und
Sonderzeichen entfernt. Die Schlüssel werden paarweise per
Levenshtein-Distanz verglichen, so werden auch Tippfehler gefunden
(`Müler GmbH`). Ein einfacher `.drop_duplicates()` auf exakte Übereinstimmung
würde das nicht schaffen.

**Zusammenführen statt Löschen.** Pro Duplikatgruppe bleibt der
vollständigste Datensatz stehen, fehlende Felder werden aus den Duplikaten
übernommen und protokolliert, erst dann werden die Duplikate entfernt.

**Pflichtfeld-Prüfung.** Datensätze ohne Firma, PLZ, Telefon oder E-Mail
werden gelb markiert, mit Begründung in der Spalte Hinweis.

## Verwendung

1. `Stammdaten_Bereinigung.xlsm` öffnen, Makros erlauben.
2. Auf "Daten bereinigen" klicken.
3. Rohe Liste und Zieldatei wählen. Die Meldung am Ende zeigt eingelesene
   Zeilen, entfernte Duplikate und protokollierte Änderungen.

Für Tests ohne Dialoge: `BereinigeKundendaten quellPfad, zielPfad`.

## Voraussetzungen

- Microsoft Excel (Windows), Makros erlaubt

## Projektstruktur

```
02_Kundenstammdaten/
├── Stammdaten_Bereinigung.xlsm    ← das Tool (VBA, eine Datei)
├── src/modStammdaten.bas          ← VBA-Code als lesbarer Export
├── generate_input.py              ← erzeugt die synthetische "schmutzige" Liste
├── input/                         ← Vorher: Kundenliste_roh.xlsx
└── output/                        ← Nachher: bereinigte Liste + Änderungsprotokoll
```
