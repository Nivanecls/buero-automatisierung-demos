Attribute VB_Name = "modMahnwesen"
Option Explicit

' =====================================================================
' Automatisiertes Mahnwesen
' ---------------------------------------------------------------------
' Liest eine Offene-Posten-Liste (CSV) und erzeugt fuer jede
' ueberfaellige Rechnung automatisch ein Mahnschreiben (Word) mit
' passender Eskalationsstufe:
'   1-7  Tage ueberfaellig -> Stufe 1: freundliche Zahlungserinnerung
'   8-21 Tage ueberfaellig -> Stufe 2: Mahnung mit Fristsetzung
'   22+  Tage ueberfaellig -> Stufe 3: letzte Mahnung vor rechtlichen
'                                       Schritten
' Zusaetzlich entsteht eine Excel-Uebersicht: alle Posten mit Stufe und
' Tagen, farblich markiert, plus Zusammenfassung (Anzahl je Stufe,
' offene Gesamtsumme).
'
' Es wird NICHTS versendet - die Schreiben liegen als Dateien bereit
' und werden bewusst manuell geprueft und verschickt.
'
' CSV-Format (Semikolon, UTF-8): Kunde;Rechnungsnummer;Rechnungsdatum;
'   Betrag;Faelligkeitsdatum;Zahlungsstatus  (Betrag: 1.234,56)
'
' Aufruf ohne Parameter  -> Dialogmodus (Dateiauswahl + Ergebnismeldung)
' Aufruf mit Parametern  -> Automatikmodus (keine Dialoge, fuer Tests)
' =====================================================================

' Absender (fuer die Demo eine fiktive Kanzlei)
Private Const ABSENDER_NAME As String = "Muster & Partner Steuerberatung"
Private Const ABSENDER_ADRESSE As String = "Beispielweg 1, 01069 Dresden"

Private Const wdFormatXMLDocument As Long = 12

Public Sub ErstelleMahnungen(Optional ByVal csvPfad As String = "", _
                             Optional ByVal zielOrdner As String = "")

    Dim automatik As Boolean
    automatik = (Len(csvPfad) > 0)

    If Not automatik Then
        Dim auswahl As Variant
        auswahl = Application.GetOpenFilename( _
            "CSV-Dateien (*.csv), *.csv", , "Offene-Posten-Liste auswaehlen")
        If auswahl = False Then Exit Sub
        csvPfad = CStr(auswahl)

        With Application.FileDialog(msoFileDialogFolderPicker)
            .Title = "Zielordner fuer Mahnschreiben und Uebersicht"
            If .Show = 0 Then Exit Sub
            zielOrdner = .SelectedItems(1)
        End With
    End If
    If Right$(zielOrdner, 1) <> "\" Then zielOrdner = zielOrdner & "\"
    If Dir(zielOrdner, vbDirectory) = "" Then MkDir zielOrdner

    ' ------------------------------------------------------------ CSV lesen
    Dim zeilen() As String, kopf() As String
    zeilen = LiesCsvZeilen(csvPfad)
    If UBound(zeilen) < 1 Then
        If Not automatik Then MsgBox "CSV enthaelt keine Daten.", vbExclamation
        Exit Sub
    End If
    kopf = Split(zeilen(0), ";")

    Dim ixKunde As Long, ixNr As Long, ixRgDat As Long
    Dim ixBetrag As Long, ixFaellig As Long, ixStatus As Long
    ixKunde = SpaltenIndex(kopf, "Kunde")
    ixNr = SpaltenIndex(kopf, "Rechnungsnummer")
    ixRgDat = SpaltenIndex(kopf, "Rechnungsdatum")
    ixBetrag = SpaltenIndex(kopf, "Betrag")
    ixFaellig = SpaltenIndex(kopf, "Faelligkeitsdatum")
    ixStatus = SpaltenIndex(kopf, "Zahlungsstatus")

    ' ------------------------------------------------- Posten verarbeiten
    Dim wordApp As Object
    Dim i As Long, felder() As String
    Dim kunde As String, nr As String, status As String, dateiname As String
    Dim rgDatum As Date, faellig As Date, betrag As Double
    Dim tage As Long, stufe As Long
    Dim anzStufe(0 To 3) As Long, summeOffen As Double, summeUeberfaellig As Double

    ' Ergebnisdaten fuer die Uebersicht (Spalten x Zeilen)
    Dim erg() As Variant
    ReDim erg(1 To 9, 1 To UBound(zeilen))
    Dim anzahl As Long

    On Error GoTo Fehler

    For i = 1 To UBound(zeilen)
        If Len(Trim$(zeilen(i))) = 0 Then GoTo Weiter
        felder = Split(zeilen(i), ";")
        kunde = Trim$(felder(ixKunde))
        nr = Trim$(felder(ixNr))
        rgDatum = ParseDatumDE(Trim$(felder(ixRgDat)))
        betrag = ParseBetragDE(Trim$(felder(ixBetrag)))
        faellig = ParseDatumDE(Trim$(felder(ixFaellig)))
        status = LCase$(Trim$(felder(ixStatus)))

        stufe = 0: tage = 0: dateiname = ""
        If status = "offen" Then
            summeOffen = summeOffen + betrag
            tage = CLng(Date - faellig)
            If tage < 1 Then
                stufe = 0
            ElseIf tage <= 7 Then
                stufe = 1
            ElseIf tage <= 21 Then
                stufe = 2
            Else
                stufe = 3
            End If
        End If

        If stufe > 0 Then
            anzStufe(stufe) = anzStufe(stufe) + 1
            summeUeberfaellig = summeUeberfaellig + betrag
            If wordApp Is Nothing Then
                Set wordApp = CreateObject("Word.Application")
                wordApp.Visible = False
                wordApp.DisplayAlerts = 0
            End If
            dateiname = "Stufe" & stufe & "_" & BereinigeDateiname(kunde) & "_" & _
                        BereinigeDateiname(nr) & ".docx"
            SchreibeMahnbrief wordApp, zielOrdner & dateiname, stufe, kunde, nr, _
                              rgDatum, betrag, faellig, tage
        End If

        anzahl = anzahl + 1
        erg(1, anzahl) = kunde
        erg(2, anzahl) = nr
        erg(3, anzahl) = rgDatum
        erg(4, anzahl) = betrag
        erg(5, anzahl) = faellig
        erg(6, anzahl) = status
        erg(7, anzahl) = IIf(status = "offen" And tage > 0, tage, Empty)
        erg(8, anzahl) = IIf(stufe > 0, "Stufe " & stufe, "-")
        erg(9, anzahl) = dateiname
Weiter:
    Next i

    If Not wordApp Is Nothing Then wordApp.Quit 0
    Set wordApp = Nothing

    ' ------------------------------------------------------------ Uebersicht
    SchreibeUebersicht erg, anzahl, anzStufe, summeOffen, summeUeberfaellig, _
                       zielOrdner & "Mahnuebersicht.xlsx"

    If Not automatik Then
        MsgBox "Fertig!" & vbCrLf & _
               "Stufe 1 (Erinnerung): " & anzStufe(1) & " Schreiben" & vbCrLf & _
               "Stufe 2 (Mahnung mit Frist): " & anzStufe(2) & " Schreiben" & vbCrLf & _
               "Stufe 3 (letzte Mahnung): " & anzStufe(3) & " Schreiben" & vbCrLf & vbCrLf & _
               Uml("~Uberf~allige Summe: ") & FormatEuroDE(summeUeberfaellig) & vbCrLf & _
               Uml("~Ubersicht: Mahnuebersicht.xlsx"), _
               vbInformation, "Mahnwesen"
    End If
    Exit Sub

Fehler:
    On Error Resume Next
    If Not wordApp Is Nothing Then wordApp.Quit 0
    If Not automatik Then
        MsgBox "Fehler: " & Err.Description, vbCritical, "Mahnwesen"
    Else
        Debug.Print "FEHLER: " & Err.Description
    End If
End Sub

' ---------------------------------------------------------------------
' Erzeugt ein einzelnes Mahnschreiben als Word-Dokument
' ---------------------------------------------------------------------
Private Sub SchreibeMahnbrief(ByVal wordApp As Object, ByVal pfad As String, _
                              ByVal stufe As Long, ByVal kunde As String, _
                              ByVal nr As String, ByVal rgDatum As Date, _
                              ByVal betrag As Double, ByVal faellig As Date, _
                              ByVal tage As Long)

    Dim betreff As String, anrede As String, text As String, frist As Date

    anrede = "Sehr geehrte Damen und Herren,"

    Select Case stufe
        Case 1
            frist = Date + 10
            betreff = "Zahlungserinnerung - Rechnung " & nr
            text = Uml("sicherlich ist es Ihrer Aufmerksamkeit entgangen, dass unsere Rechnung " & _
                   nr & " vom " & Format$(rgDatum, "DD.MM.YYYY") & Uml(" ~uber ") & FormatEuroDE(betrag) & _
                   " am " & Format$(faellig, "DD.MM.YYYY") & Uml(" f~allig war.") & vbCr & vbCr & _
                   "Wir bitten Sie, den offenen Betrag bis zum " & Format$(frist, "DD.MM.YYYY") & _
                   " auszugleichen." & vbCr & vbCr & _
                   "Sollte sich Ihre Zahlung mit diesem Schreiben ~uberschnitten haben, " & _
                   "betrachten Sie diese Erinnerung bitte als gegenstandslos.")
        Case 2
            frist = Date + 7
            betreff = "Mahnung - Rechnung " & nr
            text = Uml("trotz unserer Zahlungserinnerung ist die Rechnung " & nr & " vom " & _
                   Format$(rgDatum, "DD.MM.YYYY") & Uml(" ~uber ") & FormatEuroDE(betrag) & _
                   " weiterhin unbeglichen. Sie ist seit " & tage & Uml(" Tagen ~uberf~allig.") & vbCr & vbCr & _
                   "Wir setzen Ihnen hiermit eine Frist bis zum " & Format$(frist, "DD.MM.YYYY") & _
                   ". Bitte ~uberweisen Sie den Betrag unter Angabe der Rechnungsnummer." & vbCr & vbCr & _
                   "Sollten Sie Fragen zur Rechnung haben, sprechen Sie uns bitte umgehend an.")
        Case Else
            frist = Date + 5
            betreff = "Letzte Mahnung - Rechnung " & nr
            text = Uml("trotz mehrfacher Aufforderung ist die Rechnung " & nr & " vom " & _
                   Format$(rgDatum, "DD.MM.YYYY") & Uml(" ~uber ") & FormatEuroDE(betrag) & _
                   " bis heute nicht beglichen (" & tage & Uml(" Tage ~uberf~allig).") & vbCr & vbCr & _
                   "Wir fordern Sie letztmalig auf, den offenen Betrag bis zum " & _
                   Format$(frist, "DD.MM.YYYY") & " zu begleichen." & vbCr & vbCr & _
                   "Nach fruchtlosem Ablauf dieser Frist werden wir ohne weitere " & _
                   Uml("Ank~undigung rechtliche Schritte einleiten und die Forderung an ein ") & _
                   Uml("Inkassob~uro ~ubergeben. Die dadurch entstehenden Kosten gehen zu Ihren Lasten."))
    End Select

    Dim doc As Object, inhalt As String
    Set doc = wordApp.Documents.Add

    inhalt = ABSENDER_NAME & vbCr & _
             ABSENDER_ADRESSE & vbCr & vbCr & _
             kunde & vbCr & vbCr & _
             "Dresden, den " & Format$(Date, "DD.MM.YYYY") & vbCr & vbCr & _
             betreff & vbCr & vbCr & _
             anrede & vbCr & vbCr & _
             text & vbCr & vbCr & _
             Uml("Mit freundlichen Gr~u~sen") & vbCr & vbCr & _
             ABSENDER_NAME

    doc.Content.text = inhalt

    ' Formatierung: Absender klein, Betreff fett
    doc.Content.Font.Name = "Calibri"
    doc.Content.Font.Size = 11
    doc.Paragraphs(1).Range.Font.Bold = True
    doc.Paragraphs(2).Range.Font.Size = 9
    doc.Paragraphs(8).Range.Font.Bold = True   ' Betreffzeile

    doc.SaveAs2 pfad, wdFormatXMLDocument
    doc.Close False
End Sub

' ---------------------------------------------------------------------
' Excel-Uebersicht mit Farbcodierung und Zusammenfassung
' ---------------------------------------------------------------------
Private Sub SchreibeUebersicht(ByRef erg() As Variant, ByVal anzahl As Long, _
                               ByRef anzStufe() As Long, ByVal summeOffen As Double, _
                               ByVal summeUeberfaellig As Double, ByVal zielDatei As String)

    Dim wb As Workbook, ws As Worksheet
    Set wb = Workbooks.Add
    Set ws = wb.Worksheets(1)
    ws.Name = "Mahnliste"

    Dim kopf As Variant, c As Long
    kopf = Array("Kunde", "Rechnungsnr.", "Rechnungsdatum", "Betrag", _
                 Uml("F~alligkeit"), "Status", Uml("Tage ~uberf~allig"), _
                 "Mahnstufe", "Schreiben")
    For c = 0 To UBound(kopf)
        With ws.Cells(1, c + 1)
            .Value = kopf(c)
            .Font.Bold = True
            .Font.Color = vbWhite
            .Interior.Color = RGB(31, 78, 121)
        End With
    Next c

    Dim i As Long, z As Long, stufeText As String
    For i = 1 To anzahl
        z = i + 1
        For c = 1 To 9
            ws.Cells(z, c).Value = erg(c, i)
        Next c
        stufeText = CStr(erg(8, i))
        Select Case stufeText
            Case "Stufe 1"
                ws.Range(ws.Cells(z, 1), ws.Cells(z, 9)).Interior.Color = RGB(255, 242, 204)
            Case "Stufe 2"
                ws.Range(ws.Cells(z, 1), ws.Cells(z, 9)).Interior.Color = RGB(252, 213, 180)
            Case "Stufe 3"
                ws.Range(ws.Cells(z, 1), ws.Cells(z, 9)).Interior.Color = RGB(255, 199, 206)
                ws.Cells(z, 8).Font.Bold = True
        End Select
        If CStr(erg(6, i)) = "bezahlt" Then
            ws.Range(ws.Cells(z, 1), ws.Cells(z, 9)).Font.Color = RGB(128, 128, 128)
        End If
    Next i

    ws.Range(ws.Cells(2, 4), ws.Cells(anzahl + 1, 4)).NumberFormat = "#,##0.00 " & ChrW$(8364)
    ws.Range(ws.Cells(2, 3), ws.Cells(anzahl + 1, 3)).NumberFormat = "DD.MM.YYYY"
    ws.Range(ws.Cells(2, 5), ws.Cells(anzahl + 1, 5)).NumberFormat = "DD.MM.YYYY"

    ' -------------------------------------------------- Zusammenfassung
    Dim sz As Long
    sz = anzahl + 3
    ws.Cells(sz, 1).Value = "Zusammenfassung"
    ws.Cells(sz, 1).Font.Bold = True
    ws.Cells(sz + 1, 1).Value = "Stufe 1 - freundliche Erinnerung:"
    ws.Cells(sz + 1, 3).Value = anzStufe(1) & " Schreiben"
    ws.Cells(sz + 2, 1).Value = "Stufe 2 - Mahnung mit Frist:"
    ws.Cells(sz + 2, 3).Value = anzStufe(2) & " Schreiben"
    ws.Cells(sz + 3, 1).Value = "Stufe 3 - letzte Mahnung:"
    ws.Cells(sz + 3, 3).Value = anzStufe(3) & " Schreiben"
    ws.Cells(sz + 4, 1).Value = "Offene Summe gesamt:"
    ws.Cells(sz + 4, 3).Value = summeOffen
    ws.Cells(sz + 5, 1).Value = Uml("davon ~uberf~allig:")
    ws.Cells(sz + 5, 3).Value = summeUeberfaellig
    ws.Range(ws.Cells(sz + 4, 3), ws.Cells(sz + 5, 3)).NumberFormat = "#,##0.00 " & ChrW$(8364)
    ws.Range(ws.Cells(sz + 4, 1), ws.Cells(sz + 5, 3)).Font.Bold = True

    Dim tabelle As Range
    Set tabelle = ws.Range(ws.Cells(1, 1), ws.Cells(anzahl + 1, 9))
    tabelle.Borders.Color = RGB(191, 191, 191)
    tabelle.AutoFilter
    ws.Columns("A:I").AutoFit

    Application.DisplayAlerts = False
    wb.SaveAs zielDatei, xlOpenXMLWorkbook
    wb.Close SaveChanges:=False
    Application.DisplayAlerts = True
End Sub

' =====================================================================
' Hilfsfunktionen
' =====================================================================

' CSV als UTF-8 lesen (unabhaengig von der Windows-Systemcodepage)
Private Function LiesCsvZeilen(ByVal pfad As String) As String()
    Dim st As Object, inhalt As String
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2                ' Text
    st.Charset = "utf-8"
    st.Open
    st.LoadFromFile pfad
    inhalt = st.ReadText(-1)
    st.Close
    inhalt = Replace$(inhalt, vbCrLf, vbLf)
    LiesCsvZeilen = Split(inhalt, vbLf)
End Function

Private Function SpaltenIndex(ByRef kopf() As String, ByVal name As String) As Long
    Dim i As Long
    For i = LBound(kopf) To UBound(kopf)
        If LCase$(Trim$(kopf(i))) = LCase$(name) Then
            SpaltenIndex = i
            Exit Function
        End If
    Next i
    Err.Raise vbObjectError + 1, , "Spalte '" & name & "' nicht in der CSV gefunden."
End Function

' Platzhalter ~u ~a ~o ~s usw. -> echte Umlaute (Code bleibt reines ASCII)
Private Function Uml(ByVal s As String) As String
    s = Replace$(s, "~u", ChrW$(252)): s = Replace$(s, "~U", ChrW$(220))
    s = Replace$(s, "~a", ChrW$(228)): s = Replace$(s, "~A", ChrW$(196))
    s = Replace$(s, "~o", ChrW$(246)): s = Replace$(s, "~O", ChrW$(214))
    s = Replace$(s, "~s", ChrW$(223))
    Uml = s
End Function

' "1.234,56" -> 1234.56 - unabhaengig von den Regionaleinstellungen
Private Function ParseBetragDE(ByVal s As String) As Double
    ParseBetragDE = Val(Replace$(Replace$(s, ".", ""), ",", "."))
End Function

' "24.06.2026" -> Date - ohne CDate, damit die Landeseinstellung egal ist
Private Function ParseDatumDE(ByVal s As String) As Date
    ParseDatumDE = DateSerial(CLng(Mid$(s, 7, 4)), CLng(Mid$(s, 4, 2)), CLng(Mid$(s, 1, 2)))
End Function

' 1234.56 -> "1.234,56 EUR-Zeichen" - deutsches Format, locale-unabhaengig
Private Function FormatEuroDE(ByVal betrag As Double) As String
    Dim cents As Long, ganz As String, nachkomma As String, erg As String, i As Long
    cents = CLng(betrag * 100 + 0.0000001)
    ganz = CStr(cents \ 100)
    nachkomma = Right$("0" & CStr(cents Mod 100), 2)
    For i = Len(ganz) To 1 Step -3
        If i - 3 >= 1 Then
            erg = "." & Mid$(ganz, i - 2, 3) & erg
        Else
            erg = Mid$(ganz, 1, i) & erg
        End If
    Next i
    If Left$(erg, 1) = "." Then erg = Mid$(erg, 2)
    FormatEuroDE = erg & "," & nachkomma & " " & ChrW$(8364)
End Function

' Entfernt Zeichen, die in Dateinamen nicht erlaubt sind
Private Function BereinigeDateiname(ByVal s As String) As String
    Dim verboten As Variant, v As Variant
    verboten = Array("\", "/", ":", "*", "?", """", "<", ">", "|", "&")
    For Each v In verboten
        s = Replace$(s, CStr(v), "")
    Next v
    BereinigeDateiname = Replace$(BereinigeLeer(s), " ", "_")
End Function

Private Function BereinigeLeer(ByVal s As String) As String
    s = Trim$(s)
    Do While InStr(s, "  ") > 0
        s = Replace$(s, "  ", " ")
    Loop
    BereinigeLeer = s
End Function
