Attribute VB_Name = "modDokumente"
Option Explicit

' =====================================================================
' Automatische Dokumentenerstellung aus Excel
' ---------------------------------------------------------------------
' Liest eine Teilnehmerliste (Excel) und erzeugt fuer jede Zeile eine
' personalisierte Teilnahmebescheinigung auf Basis einer Word-Vorlage
' mit Platzhaltern ({{NAME}}, {{KURS}}, {{DATUM}}, {{STUNDEN}},
' {{AUSSTELLUNGSDATUM}}). Jedes Dokument wird als .docx und .pdf
' gespeichert.
'
' Aufruf ohne Parameter  -> Dialogmodus (Dateiauswahl + Ergebnismeldung)
' Aufruf mit Parametern  -> Automatikmodus (keine Dialoge, fuer Tests)
'
' Voraussetzung: Microsoft Word ist installiert (Word wird unsichtbar
' im Hintergrund gesteuert).
' =====================================================================

' Word-Konstanten (Late Binding, daher hier selbst definiert)
Private Const wdFindContinue As Long = 1
Private Const wdReplaceAll As Long = 2
Private Const wdExportFormatPDF As Long = 17
Private Const wdFormatXMLDocument As Long = 12

Public Sub ErstelleBescheinigungen(Optional ByVal listePfad As String = "", _
                                   Optional ByVal vorlagePfad As String = "", _
                                   Optional ByVal zielOrdner As String = "")

    Dim automatik As Boolean
    automatik = (Len(listePfad) > 0)

    ' ----------------------------------------------------------- Dialogmodus
    If Not automatik Then
        Dim auswahl As Variant

        auswahl = Application.GetOpenFilename( _
            "Excel-Dateien (*.xlsx;*.xlsm), *.xlsx;*.xlsm", , _
            "Teilnehmerliste auswaehlen")
        If auswahl = False Then Exit Sub
        listePfad = CStr(auswahl)

        auswahl = Application.GetOpenFilename( _
            "Word-Vorlagen (*.docx), *.docx", , _
            "Word-Vorlage auswaehlen")
        If auswahl = False Then Exit Sub
        vorlagePfad = CStr(auswahl)

        With Application.FileDialog(msoFileDialogFolderPicker)
            .Title = "Zielordner fuer die Dokumente auswaehlen"
            If .Show = 0 Then Exit Sub
            zielOrdner = .SelectedItems(1)
        End With
    End If

    If Right$(zielOrdner, 1) <> "\" Then zielOrdner = zielOrdner & "\"
    If Dir(zielOrdner, vbDirectory) = "" Then MkDir zielOrdner

    ' ----------------------------------------------------- Teilnehmer einlesen
    Dim wbListe As Workbook, wsListe As Worksheet
    Dim daten As Variant, letzteZeile As Long

    Set wbListe = Workbooks.Open(listePfad, ReadOnly:=True)
    Set wsListe = wbListe.Worksheets(1)
    letzteZeile = wsListe.Cells(wsListe.Rows.Count, 1).End(xlUp).Row
    If letzteZeile < 2 Then
        wbListe.Close SaveChanges:=False
        If Not automatik Then MsgBox "Die Liste enthaelt keine Datenzeilen.", vbExclamation
        Exit Sub
    End If
    daten = wsListe.Range("A2:D" & letzteZeile).Value
    wbListe.Close SaveChanges:=False

    ' ----------------------------------------------------------- Word starten
    Dim wordApp As Object
    Set wordApp = CreateObject("Word.Application")
    wordApp.Visible = False
    wordApp.DisplayAlerts = 0

    Dim i As Long, anzahl As Long
    Dim doc As Object
    Dim nameTn As String, kurs As String, datumStr As String, stunden As String
    Dim dateiBasis As String

    On Error GoTo Fehler

    For i = 1 To UBound(daten, 1)
        nameTn = Trim$(CStr(daten(i, 1)))
        If Len(nameTn) = 0 Then GoTo NaechsteZeile

        kurs = Trim$(CStr(daten(i, 2)))
        ' Datum immer im deutschen Format ausgeben, egal wie es in Excel steht
        If IsDate(daten(i, 3)) Then
            datumStr = Format$(CDate(daten(i, 3)), "DD.MM.YYYY")
        Else
            datumStr = CStr(daten(i, 3))
        End If
        stunden = CStr(daten(i, 4))

        ' Neues Dokument auf Basis der Vorlage (Vorlage bleibt unveraendert)
        Set doc = wordApp.Documents.Add(Template:=vorlagePfad)

        ErsetzePlatzhalter doc, "{{NAME}}", nameTn
        ErsetzePlatzhalter doc, "{{KURS}}", kurs
        ErsetzePlatzhalter doc, "{{DATUM}}", datumStr
        ErsetzePlatzhalter doc, "{{STUNDEN}}", stunden
        ErsetzePlatzhalter doc, "{{AUSSTELLUNGSDATUM}}", Format$(Date, "DD.MM.YYYY")

        ' Dateiname = Name des Teilnehmers (ohne unzulaessige Zeichen)
        dateiBasis = zielOrdner & "Bescheinigung_" & BereinigeDateiname(nameTn)

        doc.SaveAs2 dateiBasis & ".docx", wdFormatXMLDocument
        doc.ExportAsFixedFormat dateiBasis & ".pdf", wdExportFormatPDF
        doc.Close SaveChanges:=False
        Set doc = Nothing

        anzahl = anzahl + 1
NaechsteZeile:
    Next i

    wordApp.Quit
    Set wordApp = Nothing

    If Not automatik Then
        MsgBox "Fertig! " & anzahl & " Bescheinigungen erstellt (jeweils als DOCX und PDF)." _
             & vbCrLf & "Zielordner: " & zielOrdner, vbInformation, "Dokumentenerstellung"
    End If
    Exit Sub

Fehler:
    On Error Resume Next
    If Not doc Is Nothing Then doc.Close SaveChanges:=False
    If Not wordApp Is Nothing Then wordApp.Quit
    If Not automatik Then
        MsgBox "Fehler: " & Err.Description, vbCritical, "Dokumentenerstellung"
    Else
        Debug.Print "FEHLER: " & Err.Description
    End If
End Sub

' Ersetzt einen Platzhalter im gesamten Dokument (inkl. Kopf-/Fusszeilen nicht
' noetig - die Vorlage nutzt nur den Textkoerper)
Private Sub ErsetzePlatzhalter(ByVal doc As Object, ByVal suchen As String, ByVal ersetzen As String)
    With doc.Content.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = suchen
        .Replacement.Text = ersetzen
        .Forward = True
        .Wrap = wdFindContinue
        .MatchCase = True
        .Execute Replace:=wdReplaceAll
    End With
End Sub

' Entfernt Zeichen, die in Dateinamen nicht erlaubt sind
Private Function BereinigeDateiname(ByVal s As String) As String
    Dim verboten As Variant, v As Variant
    verboten = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    For Each v In verboten
        s = Replace$(s, CStr(v), "")
    Next v
    BereinigeDateiname = Replace$(Trim$(s), " ", "_")
End Function
