# -*- coding: utf-8 -*-
"""
Generates 9 synthetic German supplier invoices as PDF (three different layouts)
for the demo project "Verarbeitung von Eingangsrechnungen".

All companies, addresses and amounts are fictional. Amounts use the German
number format (1.234,56 EUR). Some invoices are already overdue relative to
the generation date so that the result table shows a realistic mix.

Run:  python generate_input.py
"""

import os
import random
from datetime import date, timedelta

from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas

random.seed(7)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
INPUT_DIR = os.path.join(BASE_DIR, "input")
os.makedirs(INPUT_DIR, exist_ok=True)

W, H = A4
HEUTE = date.today()


def eur(x):
    """1234.56 -> '1.234,56 €' (German format)"""
    s = f"{x:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return s + " €"


def dstr(d):
    return d.strftime("%d.%m.%Y")


LIEFERANTEN = [
    ("Bürobedarf Sachse GmbH", "Königsbrücker Str. 45", "01099 Dresden"),
    ("Elbe IT-Service GmbH & Co. KG", "Hafenstraße 8", "01067 Dresden"),
    ("Druckerei Morgenstern e.K.", "Gutenbergplatz 3", "04103 Leipzig"),
    ("Gebäudereinigung Blitz & Blank GmbH", "Industrieweg 21", "01139 Dresden"),
    ("Kaffee & Co. Handels GmbH", "Rösterallee 7", "01309 Dresden"),
    ("Werkzeugtechnik Richter AG", "Stahlstraße 112", "09112 Chemnitz"),
    ("Papierwerk Lausitz GmbH", "Am Wehr 2", "02625 Bautzen"),
    ("CleanOffice Dienstleistungen UG", "Servicegasse 14", "01159 Dresden"),
    ("Medientechnik Falke GmbH", "Funkring 33", "01237 Dresden"),
]

POSITIONEN = [
    ("Kopierpapier A4 80g, Karton", 24.90, 34.50),
    ("Tonerkartusche schwarz", 59.00, 89.00),
    ("Wartung Netzwerkdrucker", 85.00, 140.00),
    ("Unterhaltsreinigung Büroetage", 180.00, 420.00),
    ("Kaffeebohnen Premium 1kg", 18.50, 26.00),
    ("Visitenkarten 4/4-farbig, 500 Stk.", 39.00, 69.00),
    ("IT-Support vor Ort (Std.)", 78.00, 110.00),
    ("Briefumschläge C4 m. Fenster, 250 Stk.", 21.00, 32.00),
    ("Beamer-Installation Konferenzraum", 240.00, 480.00),
    ("Aktenvernichtung nach DSGVO", 45.00, 95.00),
]

EMPFAENGER = ["Muster & Partner Steuerberatung", "Beispielweg 1", "01069 Dresden"]


def make_positionen():
    n = random.randint(2, 4)
    zeilen = []
    for text, lo, hi in random.sample(POSITIONEN, n):
        menge = random.randint(1, 5)
        preis = round(random.uniform(lo, hi), 2)
        zeilen.append((text, menge, preis, round(menge * preis, 2)))
    return zeilen


def kopf_empfaenger(c, y):
    c.setFont("Helvetica", 10)
    for zeile in EMPFAENGER:
        c.drawString(25 * mm, y, zeile)
        y -= 5 * mm
    return y


# ----------------------------------------------------------------- Layout A
# klassischer Briefkopf, Metadaten rechts, Summenblock rechts unten
def layout_a(c, d):
    c.setFont("Helvetica-Bold", 14)
    c.drawString(25 * mm, H - 25 * mm, d["name"])
    c.setFont("Helvetica", 9)
    c.setFillGray(0.35)
    c.drawString(25 * mm, H - 31 * mm, f"{d['strasse']} · {d['ort']}")
    c.setFillGray(0)

    kopf_empfaenger(c, H - 55 * mm)

    c.setFont("Helvetica", 10)
    c.drawRightString(W - 25 * mm, H - 55 * mm, f"Rechnungsnummer: {d['nr']}")
    c.drawRightString(W - 25 * mm, H - 61 * mm, f"Rechnungsdatum: {dstr(d['datum'])}")
    c.drawRightString(W - 25 * mm, H - 67 * mm, f"Kundennummer: K-{random.randint(1000, 9999)}")

    c.setFont("Helvetica-Bold", 13)
    c.drawString(25 * mm, H - 90 * mm, "Rechnung")

    y = H - 102 * mm
    c.setFont("Helvetica-Bold", 9)
    c.drawString(25 * mm, y, "Bezeichnung")
    c.drawRightString(120 * mm, y, "Menge")
    c.drawRightString(150 * mm, y, "Einzelpreis")
    c.drawRightString(W - 25 * mm, y, "Gesamt")
    c.line(25 * mm, y - 2 * mm, W - 25 * mm, y - 2 * mm)
    y -= 8 * mm
    c.setFont("Helvetica", 9)
    for text, menge, preis, summe in d["positionen"]:
        c.drawString(25 * mm, y, text)
        c.drawRightString(120 * mm, y, str(menge))
        c.drawRightString(150 * mm, y, eur(preis))
        c.drawRightString(W - 25 * mm, y, eur(summe))
        y -= 6 * mm

    y -= 6 * mm
    c.line(110 * mm, y + 3 * mm, W - 25 * mm, y + 3 * mm)
    c.setFont("Helvetica", 10)
    c.drawString(110 * mm, y - 3 * mm, "Nettobetrag:")
    c.drawRightString(W - 25 * mm, y - 3 * mm, eur(d["netto"]))
    c.drawString(110 * mm, y - 9 * mm, f"zzgl. {d['satz']} % MwSt:")
    c.drawRightString(W - 25 * mm, y - 9 * mm, eur(d["mwst"]))
    c.setFont("Helvetica-Bold", 11)
    c.drawString(110 * mm, y - 17 * mm, "Rechnungsbetrag:")
    c.drawRightString(W - 25 * mm, y - 17 * mm, eur(d["brutto"]))

    c.setFont("Helvetica", 10)
    c.drawString(25 * mm, y - 32 * mm, f"Zahlbar bis {dstr(d['ziel'])} ohne Abzug.")
    c.setFont("Helvetica", 8)
    c.setFillGray(0.4)
    c.drawString(25 * mm, 20 * mm,
                 f"{d['name']} · USt-IdNr. DE{random.randint(100000000, 999999999)} · Dies ist eine fiktive Demo-Rechnung.")
    c.setFillGray(0)


# ----------------------------------------------------------------- Layout B
# moderner Stil: grosser Titel, Metadaten links untereinander, andere Labels
def layout_b(c, d):
    c.setFillColorRGB(0.12, 0.31, 0.47)
    c.rect(0, H - 18 * mm, W, 18 * mm, stroke=0, fill=1)
    c.setFillColorRGB(1, 1, 1)
    c.setFont("Helvetica-Bold", 13)
    c.drawString(25 * mm, H - 12 * mm, d["name"])
    c.setFillColorRGB(0, 0, 0)
    c.setFont("Helvetica", 9)
    c.drawString(25 * mm, H - 24 * mm, f"{d['strasse']}, {d['ort']}")

    c.setFont("Helvetica-Bold", 20)
    c.drawString(25 * mm, H - 45 * mm, f"RECHNUNG Nr. {d['nr']}")

    c.setFont("Helvetica", 10)
    c.drawString(25 * mm, H - 55 * mm, f"Datum: {dstr(d['datum'])}")
    c.drawString(25 * mm, H - 61 * mm, f"Leistungszeitraum: {d['datum'].strftime('%m/%Y')}")

    kopf_empfaenger(c, H - 75 * mm)

    y = H - 95 * mm
    c.setFont("Helvetica-Bold", 9)
    c.drawString(25 * mm, y, "Leistung")
    c.drawRightString(W - 25 * mm, y, "Betrag")
    y -= 7 * mm
    c.setFont("Helvetica", 9)
    for text, menge, preis, summe in d["positionen"]:
        c.drawString(25 * mm, y, f"{menge} x {text}")
        c.drawRightString(W - 25 * mm, y, eur(summe))
        y -= 6 * mm

    y -= 8 * mm
    c.setFont("Helvetica", 10)
    c.drawString(25 * mm, y, "Zwischensumme (netto)")
    c.drawRightString(W - 25 * mm, y, eur(d["netto"]))
    y -= 6 * mm
    c.drawString(25 * mm, y, f"Umsatzsteuer {d['satz']} %")
    c.drawRightString(W - 25 * mm, y, eur(d["mwst"]))
    y -= 8 * mm
    c.setFont("Helvetica-Bold", 12)
    c.drawString(25 * mm, y, "Gesamtbetrag")
    c.drawRightString(W - 25 * mm, y, eur(d["brutto"]))

    y -= 15 * mm
    c.setFont("Helvetica", 10)
    c.drawString(25 * mm, y, f"Zahlungsziel: {dstr(d['ziel'])}")
    c.drawString(25 * mm, y - 6 * mm, "Bitte überweisen Sie den Betrag unter Angabe der Rechnungsnummer.")
    c.setFont("Helvetica", 8)
    c.setFillGray(0.4)
    c.drawString(25 * mm, 20 * mm, "Fiktive Demo-Rechnung – alle Angaben ohne Gewähr.")
    c.setFillGray(0)


# ----------------------------------------------------------------- Layout C
# schlichte Dienstleister-Rechnung, Metadaten im Fliesstext, "Fällig am"
def layout_c(c, d):
    c.setFont("Courier-Bold", 12)
    c.drawString(25 * mm, H - 25 * mm, d["name"])
    c.setFont("Courier", 9)
    c.drawString(25 * mm, H - 31 * mm, f"{d['strasse']} | {d['ort']}")
    c.line(25 * mm, H - 34 * mm, W - 25 * mm, H - 34 * mm)

    kopf_empfaenger(c, H - 48 * mm)

    c.setFont("Courier-Bold", 12)
    c.drawString(25 * mm, H - 75 * mm, f"Rechnung {d['nr']}")
    c.setFont("Courier", 10)
    c.drawString(25 * mm, H - 83 * mm, f"vom {dstr(d['datum'])}")

    y = H - 98 * mm
    c.setFont("Courier", 9)
    for text, menge, preis, summe in d["positionen"]:
        c.drawString(25 * mm, y, f"{text[:40]:<42}{menge:>3}  {eur(summe):>14}")
        y -= 6 * mm

    y -= 8 * mm
    c.setFont("Courier", 10)
    c.drawString(25 * mm, y, f"{'Summe netto':<30}{eur(d['netto']):>16}")
    y -= 6 * mm
    c.drawString(25 * mm, y, f"{'MwSt ' + str(d['satz']) + ' %':<30}{eur(d['mwst']):>16}")
    y -= 6 * mm
    c.setFont("Courier-Bold", 10)
    c.drawString(25 * mm, y, f"{'Brutto zu zahlen':<30}{eur(d['brutto']):>16}")

    y -= 14 * mm
    c.setFont("Courier", 10)
    c.drawString(25 * mm, y, f"Fällig am {dstr(d['ziel'])}.")
    c.setFont("Courier", 8)
    c.drawString(25 * mm, 20 * mm, "Synthetische Rechnung fuer Demo-Zwecke.")


LAYOUTS = [layout_a, layout_b, layout_c]
NR_FORMATE = [
    lambda i: f"RE-2026-{400 + i * 17}",
    lambda i: f"2026/{500 + i * 31}",
    lambda i: f"R-26{random.randint(10, 99)}-{i + 1}",
]

uebersicht = []
for i, (name, strasse, ort) in enumerate(LIEFERANTEN):
    positionen = make_positionen()
    netto = round(sum(p[3] for p in positionen), 2)
    satz = 7 if i == 4 else 19          # eine Rechnung mit ermässigtem Satz
    mwst = round(netto * satz / 100, 2)
    brutto = round(netto + mwst, 2)

    # Mischung aus ueberfaelligen und offenen Rechnungen (relativ zu heute)
    if i % 2 == 0:
        ziel = HEUTE - timedelta(days=random.randint(3, 25))    # ueberfaellig
    else:
        ziel = HEUTE + timedelta(days=random.randint(5, 24))    # noch offen
    datum = ziel - timedelta(days=random.choice([14, 21, 30]))

    d = {
        "name": name, "strasse": strasse, "ort": ort,
        "nr": NR_FORMATE[i % 3](i), "datum": datum, "ziel": ziel,
        "positionen": positionen, "netto": netto, "satz": satz,
        "mwst": mwst, "brutto": brutto,
    }

    pfad = os.path.join(INPUT_DIR, f"Rechnung_{i + 1:02d}.pdf")
    c = canvas.Canvas(pfad, pagesize=A4)
    LAYOUTS[i % 3](c, d)
    c.save()
    uebersicht.append((os.path.basename(pfad), name, d["nr"], eur(brutto), dstr(ziel)))
    print(f"OK: {pfad}")

print("\nKontrollwerte (Soll-Daten fuer den Test):")
for zeile in uebersicht:
    print("  " + " | ".join(zeile))
