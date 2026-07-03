Attribute VB_Name = "modRechnungen"
Option Explicit

' =====================================================================
' Verarbeitung von Eingangsrechnungen (PDF -> Excel)
' ---------------------------------------------------------------------
' Liest alle PDF-Rechnungen aus einem Ordner, extrahiert die wichtigsten
' Felder (Lieferant, Rechnungsnummer, Datum, Netto, MwSt, Brutto,
' Zahlungsziel) und schreibt alles in eine formatierte Excel-Tabelle.
' Ueberfaellige Rechnungen werden rot markiert.
'
' Die Texterkennung arbeitet mit Schluesselwoertern und regulaeren
' Ausdruecken - NICHT mit festen Koordinaten. Dadurch funktioniert sie
' fuer unterschiedliche Rechnungslayouts verschiedener Lieferanten.
'
' PDF-Text wird ueber Word gewonnen: Word (2013+) kann PDF-Dateien
' oeffnen und in Text umwandeln. Word laeuft dabei unsichtbar.
'
' Aufruf ohne Parameter  -> Dialogmodus (Ordnerauswahl + Ergebnismeldung)
' Aufruf mit Parametern  -> Automatikmodus (keine Dialoge, fuer Tests)
' =====================================================================

Private Const wdDoNotSaveChanges As Long = 0

' Ergebnis der Extraktion einer einzelnen Rechnung
Private Type TRechnung
    Datei As String
    Lieferant As String
    Nummer As String
    Datum As Variant        ' Date oder leer
    Netto As Variant        ' Double oder leer
    MwSt As Variant
    Brutto As Variant
    Ziel As Variant         ' Date oder leer
End Type

Public Sub ErfasseRechnungen(Optional ByVal pdfOrdner As String = "", _
                             Optional ByVal zielDatei As String = "")

    Dim automatik As Boolean
    automatik = (Len(pdfOrdner) > 0)

    ' ----------------------------------------------------------- Dialogmodus
    If Not automatik Then
        With Application.FileDialog(msoFileDialogFolderPicker)
            .Title = "Ordner mit den PDF-Rechnungen auswaehlen"
            If .Show = 0 Then Exit Sub
            pdfOrdner = .SelectedItems(1)
        End With

        Dim auswahl As Variant
        auswahl = Application.GetSaveAsFilename( _
            InitialFileName:="Rechnungsuebersicht.xlsx", _
            FileFilter:="Excel-Datei (*.xlsx), *.xlsx", _
            Title:="Zieldatei fuer die Uebersicht")
        If auswahl = False Then Exit Sub
        zielDatei = CStr(auswahl)
    End If

    If Right$(pdfOrdner, 1) <> "\" Then pdfOrdner = pdfOrdner & "\"

    ' ------------------------------------------------------- PDFs durchgehen
    Dim wordApp As Object
    Set wordApp = CreateObject("Word.Application")
    wordApp.Visible = False
    wordApp.DisplayAlerts = 0

    ' Words "PDF wird konvertiert"-Hinweis abschalten, sonst haengt der
    ' unsichtbare Word-Prozess an einem unsichtbaren Dialog
    On Error Resume Next
    CreateObject("WScript.Shell").RegWrite _
        "HKCU\Software\Microsoft\Office\" & wordApp.Version & _
        "\Word\Options\DisableConvertPdfWarning", 1, "REG_DWORD"
    On Error GoTo 0

    Dim rechnungen() As TRechnung
    Dim anzahl As Long
    ReDim rechnungen(1 To 100)

    Dim datei As String
    datei = Dir(pdfOrdner & "*.pdf")

    On Error GoTo Fehler

    Do While Len(datei) > 0
        anzahl = anzahl + 1
        rechnungen(anzahl) = ExtrahiereRechnung(wordApp, pdfOrdner & datei)
        datei = Dir()
    Loop

    wordApp.Quit wdDoNotSaveChanges
    Set wordApp = Nothing

    If anzahl = 0 Then
        If Not automatik Then MsgBox "Keine PDF-Dateien im Ordner gefunden.", vbExclamation
        Exit Sub
    End If

    ' -------------------------------------------------- Ergebnis-Tabelle bauen
    SchreibeUebersicht rechnungen, anzahl, zielDatei

    If Not automatik Then
        MsgBox "Fertig! " & anzahl & " Rechnungen verarbeitet." & vbCrLf & _
               "Uebersicht gespeichert unter:" & vbCrLf & zielDatei, _
               vbInformation, "Rechnungserfassung"
    End If
    Exit Sub

Fehler:
    On Error Resume Next
    If Not wordApp Is Nothing Then wordApp.Quit wdDoNotSaveChanges
    If Not automatik Then
        MsgBox "Fehler: " & Err.Description, vbCritical, "Rechnungserfassung"
    Else
        Debug.Print "FEHLER: " & Err.Description
    End If
End Sub

' ---------------------------------------------------------------------
' Extrahiert alle Felder aus einer PDF-Datei (ueber Word als Textquelle)
' ---------------------------------------------------------------------
Private Function ExtrahiereRechnung(ByVal wordApp As Object, ByVal pdfPfad As String) As TRechnung
    Dim r As TRechnung
    r.Datei = Mid$(pdfPfad, InStrRev(pdfPfad, "\") + 1)

    Dim doc As Object, text As String
    Set doc = wordApp.Documents.Open(FileName:=pdfPfad, ConfirmConversions:=False, _
                                     ReadOnly:=True, AddToRecentFiles:=False, Visible:=False)

    ' Alle Textbereiche einsammeln: Words PDF-Konverter legt Inhalte vom
    ' oberen Seitenrand (Briefkopf!) teils in die Kopfzeilen-Story, die in
    ' Content fehlt. Kopfzeilen zuerst, dann der Haupttext.
    Dim story As Object, kopfText As String
    On Error Resume Next
    For Each story In doc.StoryRanges
        If story.StoryType <> 1 Then kopfText = kopfText & story.text & vbCr
    Next story
    On Error GoTo 0
    text = kopfText & doc.Content.text
    doc.Close wdDoNotSaveChanges

    r.Lieferant = FindeLieferant(text)
    r.Nummer = RegexErster(text, "Rechnung(?:snummer|s-?Nr\.?|\s+Nr\.?)?\s*:?\s*([A-Za-z]{0,3}-?\d[0-9A-Za-z\-\/]*)")

    Dim s As String
    s = RegexErster(text, "(?:Rechnungsdatum|Datum|vom)\s*:?\s*(\d{2}\.\d{2}\.\d{4})")
    If Len(s) > 0 Then r.Datum = ParseDatumDE(s)

    s = RegexErster(text, "(?:zahlbar\s+bis|f.llig\s+am|Zahlungsziel)\s*:?\s*(\d{2}\.\d{2}\.\d{4})")
    If Len(s) > 0 Then r.Ziel = ParseDatumDE(s)

    s = RegexErster(text, "(?:Nettobetrag|Zwischensumme\s*\(netto\)|Summe\s+netto)\s*:?[\s\S]{0,60}?(\d{1,3}(?:\.\d{3})*,\d{2})")
    If Len(s) > 0 Then r.Netto = ParseBetragDE(s)

    s = RegexErster(text, "(?:MwSt|Umsatzsteuer|USt\s|Mehrwertsteuer)[\s\S]{0,60}?(\d{1,3}(?:\.\d{3})*,\d{2})")
    If Len(s) > 0 Then r.MwSt = ParseBetragDE(s)

    s = RegexErster(text, "(?:Rechnungsbetrag|Gesamtbetrag|Brutto[^\d\r\n]{0,20})\s*:?[\s\S]{0,60}?(\d{1,3}(?:\.\d{3})*,\d{2})")
    If Len(s) > 0 Then r.Brutto = ParseBetragDE(s)

    ExtrahiereRechnung = r
End Function

' Lieferant = erste nicht-leere Textzeile; falls sie keine Rechtsform
' enthaelt, wird in den ersten Zeilen nach einer Firma gesucht
Private Function FindeLieferant(ByVal text As String) As String
    Dim zeilen() As String, i As Long, z As String, erste As String
    text = Replace$(text, vbCr, vbLf)
    zeilen = Split(text, vbLf)

    For i = LBound(zeilen) To UBound(zeilen)
        z = Trim$(zeilen(i))
        If Len(z) > 1 Then
            If Len(erste) = 0 Then erste = z
            If z Like "*GmbH*" Or z Like "*AG*" Or z Like "* KG*" _
               Or z Like "*e.K.*" Or z Like "* UG*" Then
                FindeLieferant = z
                Exit Function
            End If
            If i > 12 Then Exit For
        End If
    Next i
    FindeLieferant = erste
End Function

' Liefert die erste Fanggruppe des Musters oder ""
Private Function RegexErster(ByVal text As String, ByVal muster As String) As String
    Static rx As Object
    If rx Is Nothing Then Set rx = CreateObject("VBScript.RegExp")
    rx.Pattern = muster
    rx.IgnoreCase = True
    rx.Global = False
    If rx.Test(text) Then RegexErster = rx.Execute(text)(0).SubMatches(0)
End Function

' "1.234,56" -> 1234.56 - unabhaengig von den Regionaleinstellungen
Private Function ParseBetragDE(ByVal s As String) As Double
    ParseBetragDE = Val(Replace$(Replace$(s, ".", ""), ",", "."))
End Function

' "24.06.2026" -> Date - ohne CDate, damit die Landeseinstellung egal ist
Private Function ParseDatumDE(ByVal s As String) As Date
    ParseDatumDE = DateSerial(CLng(Mid$(s, 7, 4)), CLng(Mid$(s, 4, 2)), CLng(Mid$(s, 1, 2)))
End Function

' ---------------------------------------------------------------------
' Schreibt die Uebersichtstabelle mit Formatierung und speichert sie
' ---------------------------------------------------------------------
Private Sub SchreibeUebersicht(ByRef rechnungen() As TRechnung, ByVal anzahl As Long, _
                               ByVal zielDatei As String)

    Dim sUeberfaellig As String, sOffen As String, sPruefen As String
    sUeberfaellig = ChrW$(220) & "berf" & ChrW$(228) & "llig"   ' Ueberfaellig
    sOffen = "Offen"
    sPruefen = "Pr" & ChrW$(252) & "fen"                        ' Pruefen

    Dim wb As Workbook, ws As Worksheet
    Set wb = Workbooks.Add
    Set ws = wb.Worksheets(1)
    ws.Name = "Eingangsrechnungen"

    Dim kopf As Variant
    kopf = Array("Datei", "Lieferant", "Rechnungsnr.", "Rechnungsdatum", _
                 "Netto", "MwSt", "Brutto", "Zahlungsziel", "Status")

    Dim c As Long
    For c = 0 To UBound(kopf)
        With ws.Cells(1, c + 1)
            .Value = kopf(c)
            .Font.Bold = True
            .Font.Color = vbWhite
            .Interior.Color = RGB(31, 78, 121)
        End With
    Next c

    Dim i As Long, z As Long, status As String, plausibel As Boolean
    For i = 1 To anzahl
        z = i + 1
        With rechnungen(i)
            ws.Cells(z, 1).Value = .Datei
            ws.Cells(z, 2).Value = .Lieferant
            ws.Cells(z, 3).Value = "'" & .Nummer   ' als Text, damit z.B. 2026/531 kein Datum wird
            If IsDate(.Datum) Then ws.Cells(z, 4).Value = CDate(.Datum)
            If IsNumeric(.Netto) Then ws.Cells(z, 5).Value = CDbl(.Netto)
            If IsNumeric(.MwSt) Then ws.Cells(z, 6).Value = CDbl(.MwSt)
            If IsNumeric(.Brutto) Then ws.Cells(z, 7).Value = CDbl(.Brutto)
            If IsDate(.Ziel) Then ws.Cells(z, 8).Value = CDate(.Ziel)

            ' Plausibilitaet: Netto + MwSt = Brutto (1 Cent Toleranz)
            plausibel = IsNumeric(.Netto) And IsNumeric(.MwSt) And IsNumeric(.Brutto)
            If plausibel Then plausibel = (Abs(CDbl(.Netto) + CDbl(.MwSt) - CDbl(.Brutto)) <= 0.02)

            If Not plausibel Or Len(.Nummer) = 0 Or Not IsDate(.Ziel) Then
                status = sPruefen
                ws.Range(ws.Cells(z, 1), ws.Cells(z, 9)).Interior.Color = RGB(255, 235, 156)
            ElseIf CDate(.Ziel) < Date Then
                status = sUeberfaellig
                ws.Range(ws.Cells(z, 1), ws.Cells(z, 9)).Interior.Color = RGB(255, 199, 206)
                ws.Cells(z, 9).Font.Color = RGB(156, 0, 6)
                ws.Cells(z, 9).Font.Bold = True
            Else
                status = sOffen
            End If
            ws.Cells(z, 9).Value = status
        End With
    Next i

    ' Summenzeile
    Dim sz As Long
    sz = anzahl + 3
    ws.Cells(sz, 2).Value = "Summe (" & anzahl & " Rechnungen)"
    ws.Cells(sz, 2).Font.Bold = True
    For c = 5 To 7
        ws.Cells(sz, c).FormulaR1C1 = "=SUM(R2C:R" & anzahl + 1 & "C)"
        ws.Cells(sz, c).Font.Bold = True
    Next c
    ws.Range(ws.Cells(sz, 2), ws.Cells(sz, 9)).Borders(xlEdgeTop).LineStyle = xlContinuous

    ' Zahlen- und Datumsformate (NumberFormat erwartet US-Formatcodes)
    ws.Range(ws.Cells(2, 5), ws.Cells(sz, 7)).NumberFormat = "#,##0.00 " & ChrW$(8364)
    ws.Range(ws.Cells(2, 4), ws.Cells(anzahl + 1, 4)).NumberFormat = "DD.MM.YYYY"
    ws.Range(ws.Cells(2, 8), ws.Cells(anzahl + 1, 8)).NumberFormat = "DD.MM.YYYY"

    Dim tabelle As Range
    Set tabelle = ws.Range(ws.Cells(1, 1), ws.Cells(anzahl + 1, 9))
    tabelle.Borders.Color = RGB(191, 191, 191)
    tabelle.AutoFilter
    ws.Columns("A:I").AutoFit
    ' Kopfzeile fixieren (in unsichtbarem Excel nicht immer moeglich -> tolerant)
    On Error Resume Next
    ws.Rows(2).Select
    ActiveWindow.FreezePanes = True
    ws.Cells(1, 1).Select
    On Error GoTo 0

    Application.DisplayAlerts = False
    wb.SaveAs zielDatei, xlOpenXMLWorkbook
    wb.Close SaveChanges:=False
    Application.DisplayAlerts = True
End Sub
