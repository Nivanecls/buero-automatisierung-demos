# -*- coding: utf-8 -*-
"""
Generates a synthetic open-invoice list (CSV) for the demo project
"Automatisiertes Mahnwesen".

18 invoices with a realistic mix:
  - paid invoices                          (no reminder)
  - open, not yet due                      (no reminder)
  - overdue  1-7 days   -> Stufe 1: freundliche Zahlungserinnerung
  - overdue  8-21 days  -> Stufe 2: Mahnung mit Fristsetzung
  - overdue 22+ days    -> Stufe 3: letzte Mahnung

CSV: semicolon separated, UTF-8 with BOM, German number/date formats
(typical German ERP export). Run:  python generate_input.py
"""

import csv
import os
import random
from datetime import date, timedelta

random.seed(23)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
INPUT_DIR = os.path.join(BASE_DIR, "input")
os.makedirs(INPUT_DIR, exist_ok=True)

HEUTE = date.today()

KUNDEN = [
    "Bäckerei Vogt", "Autohaus Lindner GmbH", "Elektro Brenner",
    "Gartenbau Wiese GmbH", "Metallbau Krause AG", "Steuerbüro Albrecht",
    "Spedition Falkner GmbH", "Optik Sonnberg", "Tischlerei Holzmann",
    "Apotheke am Markt", "Kanzlei Weber & Partner", "Sanitär Quelle GmbH",
    "Café Morgenrot", "Werbeagentur Pixelhaus GmbH", "Fahrschule Startklar",
    "Immobilien Schlüssel GmbH", "Physiotherapie Balance", "Modehaus Eleganza",
]

# Verteilung der Faelle: (Zahlungsstatus, Tage relativ zu heute)
# negative Tage = Faelligkeit liegt in der Zukunft (nicht ueberfaellig)
FAELLE = (
    [("bezahlt", random.randint(-10, 30)) for _ in range(4)]
    + [("offen", -random.randint(3, 20)) for _ in range(3)]      # noch nicht faellig
    + [("offen", random.randint(1, 7)) for _ in range(4)]        # Stufe 1
    + [("offen", random.randint(8, 21)) for _ in range(4)]       # Stufe 2
    + [("offen", random.randint(22, 45)) for _ in range(3)]      # Stufe 3
)
random.shuffle(FAELLE)


def eur_de(x):
    return f"{x:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


zeilen = []
for i, (kunde, (status, tage_ueberfaellig)) in enumerate(zip(KUNDEN, FAELLE)):
    faellig = HEUTE - timedelta(days=tage_ueberfaellig)
    rg_datum = faellig - timedelta(days=random.choice([14, 21, 30]))
    betrag = round(random.uniform(150, 4800), 2)
    zeilen.append({
        "Kunde": kunde,
        "Rechnungsnummer": f"AR-2026-{1000 + i * 7}",
        "Rechnungsdatum": rg_datum.strftime("%d.%m.%Y"),
        "Betrag": eur_de(betrag),
        "Faelligkeitsdatum": faellig.strftime("%d.%m.%Y"),
        "Zahlungsstatus": status,
    })

pfad = os.path.join(INPUT_DIR, "Offene_Posten.csv")
with open(pfad, "w", newline="", encoding="utf-8-sig") as f:
    w = csv.DictWriter(f, fieldnames=list(zeilen[0].keys()), delimiter=";")
    w.writeheader()
    w.writerows(zeilen)

print(f"OK: {pfad} ({len(zeilen)} Rechnungen)")
offen = [z for z in zeilen if z["Zahlungsstatus"] == "offen"]
print(f"  davon offen: {len(offen)}, bezahlt: {len(zeilen) - len(offen)}")
