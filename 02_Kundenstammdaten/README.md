# Bereinigung von Kundenstammdaten

**Demo-Projekt** – eine „verschmutzte" Kundenliste wird automatisch bereinigt:
Duplikate mit abweichender Schreibweise, uneinheitliche Telefon- und
E-Mail-Formate, Leerzeichen-Chaos und fehlende Pflichtfelder. Zu **jeder**
Änderung entsteht ein Protokolleintrag (vorher → nachher → Grund).
Alle Daten sind synthetisch, alle Firmen frei erfunden.

## Vorher → Nachher

| Vorher | Nachher |
|---|---|
| `input/Kundenliste_roh.xlsx` – 48 Zeilen mit Dubletten (Müller GmbH / Mueller GmbH / MÜLLER GMBH), kaputten E-Mails, fünf verschiedenen Telefonformaten, Lücken | `output/Kundenliste_bereinigt.xlsx` – 40 saubere Datensätze + Blatt **Änderungsprotokoll** mit 45 dokumentierten Korrekturen |

Das Änderungsprotokoll ist bewusst Teil des Ergebnisses: Der Kunde sieht nicht
nur die saubere Liste, sondern kann jede einzelne Entscheidung nachvollziehen –
nichts verschwindet stillschweigend.

## Die Lösung ist eine einzige Excel-Datei

`Stammdaten_Bereinigung.xlsm` enthält den kompletten VBA-Code. Kein Python,
keine Installation – nur Excel.

## Eingabeformat

Excel-Datei, erstes Tabellenblatt, Kopfzeile in Zeile 1:

| Firma | Ansprechpartner | Straße | PLZ | Ort | Telefon | E-Mail |
|---|---|---|---|---|---|---|
| Müller GmbH | Hans Weber | Hauptstraße 12 | 01067 | Dresden | 0351/456789 | Info@Mueller,example |

Die Demo-Datei wird durch `generate_input.py` erzeugt (nur für die
synthetischen Testdaten nötig).

## So funktioniert der VBA-Code

Der Code liegt in [`src/modStammdaten.bas`](src/modStammdaten.bas) (lesbarer
Export, identisch mit dem Code in der `.xlsm`).

**1. Feldbereinigung** (jede Korrektur wird protokolliert):
- Führende, folgende und doppelte Leerzeichen entfernen
- `MÜLLER GMBH` / `müller gmbh` → `Müller GmbH` – Großschreibung wird nur
  angefasst, wenn der Name komplett groß oder komplett klein geschrieben ist;
  Rechtsformen (GmbH, KG, AG, UG, e.K.) werden korrekt gesetzt
- Telefonnummern → einheitlich `+49 Vorwahl Nummer` (erkennt `0351/…`,
  `(0351) …`, `0049…`, reine Ziffernfolgen; Vorwahlen über eine Prüfliste)
- E-Mails → Kleinschreibung, Leerzeichen raus, Komma → Punkt; danach
  Regex-Validierung – weiterhin ungültige Adressen werden **markiert statt
  geraten**

**2. Duplikatsuche – nicht nur exakte Treffer.** Für jede Firma wird ein
Vergleichsschlüssel gebildet: Kleinschreibung, Umlaut-Transliteration
(ü→ue, ß→ss), Rechtsformen und Sonderzeichen entfernt. Schlüssel werden
paarweise mit der **Levenshtein-Distanz** verglichen, so dass auch Tippfehler
(`Müler GmbH`) gefunden werden. Genau das kann `.drop_duplicates()` bzw.
„Duplikate entfernen" in Excel **nicht**.

**3. Zusammenführen statt Löschen.** Pro Duplikatgruppe bleibt der
vollständigste Datensatz erhalten; fehlende Felder werden aus den Duplikaten
übernommen (protokolliert), erst dann werden die Duplikate entfernt.

**4. Pflichtfeld-Prüfung.** Datensätze ohne Firma, PLZ, Telefon oder E-Mail
werden gelb markiert und in der Spalte *Hinweis* begründet.

## Verwendung

1. `Stammdaten_Bereinigung.xlsm` öffnen, Makros erlauben.
2. Auf **„Daten bereinigen"** klicken.
3. Rohe Liste und Zieldatei wählen – die Ergebnismeldung fasst zusammen:
   eingelesene Zeilen, entfernte Duplikate, protokollierte Änderungen.

Für automatisierte Tests: `BereinigeKundendaten quellPfad, zielPfad`
(ohne Dialoge).

## Voraussetzungen

- Microsoft Excel (Desktop, Windows), Makros erlaubt
- Keine weiteren Programme oder Bibliotheken

## Projektstruktur

```
02_Kundenstammdaten/
├── Stammdaten_Bereinigung.xlsm    ← das Tool (VBA, eine Datei)
├── src/modStammdaten.bas          ← VBA-Code als lesbarer Export
├── generate_input.py              ← erzeugt die synthetische "schmutzige" Liste
├── input/                         ← Vorher: Kundenliste_roh.xlsx
└── output/                        ← Nachher: bereinigte Liste + Änderungsprotokoll
```
