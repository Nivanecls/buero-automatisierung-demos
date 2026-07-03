Attribute VB_Name = "modStammdaten"
Option Explicit

' =====================================================================
' Bereinigung von Kundenstammdaten
' ---------------------------------------------------------------------
' Liest eine "verschmutzte" Kundenliste (Excel) und erzeugt:
'   Blatt 1 "Kunden bereinigt"     - die bereinigte Liste
'   Blatt 2 "Aenderungsprotokoll"  - jede Aenderung: vorher/nachher/Grund
'
' Bereinigungsschritte:
'   - Leerzeichen (fuehrend/folgend/doppelt) entfernen
'   - GROSSSCHREIBUNG und kleinschreibung von Firmennamen normalisieren
'   - Telefonnummern in einheitliches Format bringen (+49 Vorwahl Nummer)
'   - E-Mail-Adressen korrigieren (Komma->Punkt, Leerzeichen, Kleinbuchst.)
'     und validieren
'   - Pflichtfelder pruefen (Firma, PLZ, Telefon, E-Mail)
'   - Duplikate finden - NICHT nur exakte Treffer, sondern auch
'     abweichende Schreibweisen (Mueller/Mueller GmbH/MUELLER GMBH und
'     Tippfehler) ueber Umlaut-Transliteration + Levenshtein-Distanz.
'     Duplikate werden zusammengefuehrt: fehlende Felder des Hauptsatzes
'     werden aus den Duplikaten uebernommen.
'
' Aufruf ohne Parameter  -> Dialogmodus (Dateiauswahl + Ergebnismeldung)
' Aufruf mit Parametern  -> Automatikmodus (keine Dialoge, fuer Tests)
' =====================================================================

Private Const ANZ_FELDER As Long = 7   ' Firma..E-Mail

' Protokoll waechst dynamisch
Private protokoll() As Variant
Private protAnzahl As Long

Public Sub BereinigeKundendaten(Optional ByVal quellPfad As String = "", _
                                Optional ByVal zielPfad As String = "")

    Dim automatik As Boolean
    automatik = (Len(quellPfad) > 0)

    If Not automatik Then
        Dim auswahl As Variant
        auswahl = Application.GetOpenFilename( _
            "Excel-Dateien (*.xlsx;*.xlsm), *.xlsx;*.xlsm", , _
            "Rohe Kundenliste auswaehlen")
        If auswahl = False Then Exit Sub
        quellPfad = CStr(auswahl)

        auswahl = Application.GetSaveAsFilename( _
            InitialFileName:="Kundenliste_bereinigt.xlsx", _
            FileFilter:="Excel-Datei (*.xlsx), *.xlsx", _
            Title:="Zieldatei fuer das Ergebnis")
        If auswahl = False Then Exit Sub
        zielPfad = CStr(auswahl)
    End If

    ' ------------------------------------------------------------- Einlesen
    Dim wbQ As Workbook, wsQ As Worksheet, daten As Variant, letzte As Long
    Set wbQ = Workbooks.Open(quellPfad, ReadOnly:=True)
    Set wsQ = wbQ.Worksheets(1)
    letzte = wsQ.Cells(wsQ.Rows.Count, 1).End(xlUp).Row
    If letzte < 2 Then
        wbQ.Close False
        If Not automatik Then MsgBox "Keine Datenzeilen gefunden.", vbExclamation
        Exit Sub
    End If
    daten = wsQ.Range(wsQ.Cells(2, 1), wsQ.Cells(letzte, ANZ_FELDER)).Value
    wbQ.Close False

    Dim n As Long
    n = UBound(daten, 1)

    ReDim protokoll(1 To 6, 1 To 1000)
    protAnzahl = 0

    ' --------------------------------------------------- Feld-Bereinigung
    Dim i As Long, f As Long, vorher As String, nachher As String
    Dim feldnamen As Variant
    feldnamen = Array("Firma", "Ansprechpartner", Uml("Stra~se"), "PLZ", _
                      "Ort", "Telefon", "E-Mail")

    Dim hinweise() As String
    ReDim hinweise(1 To n)

    For i = 1 To n
        For f = 1 To ANZ_FELDER
            daten(i, f) = CStr(IIf(IsError(daten(i, f)), "", daten(i, f) & ""))
        Next f

        ' 1) Leerzeichen
        For f = 1 To ANZ_FELDER
            vorher = CStr(daten(i, f))
            nachher = BereinigeLeerzeichen(vorher)
            If nachher <> vorher Then
                daten(i, f) = nachher
                Protokolliere i + 1, CStr(daten(i, 1)), CStr(feldnamen(f - 1)), _
                              vorher, nachher, "Leerzeichen bereinigt"
            End If
        Next f

        ' 2) Schreibweise Firma und Ort
        vorher = CStr(daten(i, 1))
        nachher = NormalisiereFirma(vorher)
        If nachher <> vorher Then
            daten(i, 1) = nachher
            Protokolliere i + 1, nachher, "Firma", vorher, nachher, _
                          "Schreibweise normalisiert"
        End If

        vorher = CStr(daten(i, 5))
        If Len(vorher) > 0 And vorher = UCase$(vorher) And vorher <> StrConv(vorher, vbProperCase) Then
            nachher = StrConv(vorher, vbProperCase)
            daten(i, 5) = nachher
            Protokolliere i + 1, CStr(daten(i, 1)), "Ort", vorher, nachher, _
                          "Schreibweise normalisiert"
        End If

        ' 3) Telefon
        vorher = CStr(daten(i, 6))
        nachher = NormalisiereTelefon(vorher)
        If nachher <> vorher Then
            daten(i, 6) = nachher
            Protokolliere i + 1, CStr(daten(i, 1)), "Telefon", vorher, nachher, _
                          "Telefonformat vereinheitlicht"
        End If

        ' 4) E-Mail
        vorher = CStr(daten(i, 7))
        nachher = NormalisiereEmail(vorher)
        If nachher <> vorher Then
            daten(i, 7) = nachher
            Protokolliere i + 1, CStr(daten(i, 1)), "E-Mail", vorher, nachher, _
                          "E-Mail korrigiert"
        End If
        If Len(nachher) > 0 And Not EmailGueltig(nachher) Then
            HinweisAnfuegen hinweise(i), Uml("E-Mail ung~ultig")
        End If

        ' 5) Pflichtfelder
        If Len(CStr(daten(i, 1))) = 0 Then HinweisAnfuegen hinweise(i), "Firma fehlt"
        If Len(CStr(daten(i, 4))) = 0 Then HinweisAnfuegen hinweise(i), "PLZ fehlt"
        If Len(CStr(daten(i, 6))) = 0 Then HinweisAnfuegen hinweise(i), "Telefon fehlt"
        If Len(CStr(daten(i, 7))) = 0 Then HinweisAnfuegen hinweise(i), "E-Mail fehlt"
    Next i

    ' --------------------------------------------------- Duplikate finden
    ' istDuplikatVon(i) = 0 -> eigenstaendig, sonst Index des Hauptsatzes
    Dim istDuplikatVon() As Long
    ReDim istDuplikatVon(1 To n)

    Dim schluessel() As String
    ReDim schluessel(1 To n)
    For i = 1 To n
        schluessel(i) = VergleichsSchluessel(CStr(daten(i, 1)))
    Next i

    Dim j As Long, distanzMax As Long
    For i = 1 To n
        If istDuplikatVon(i) = 0 And Len(schluessel(i)) > 0 Then
            For j = i + 1 To n
                If istDuplikatVon(j) = 0 And Len(schluessel(j)) > 0 Then
                    distanzMax = IIf(Len(schluessel(i)) < 8, 1, 2)
                    If schluessel(i) = schluessel(j) Or _
                       Levenshtein(schluessel(i), schluessel(j)) <= distanzMax Then
                        istDuplikatVon(j) = i
                    End If
                End If
            Next j
        End If
    Next i

    ' Hauptsatz = vollstaendigster Datensatz der Gruppe; danach werden
    ' fehlende Felder des Hauptsatzes aus den Duplikaten aufgefuellt
    Dim k As Long, best As Long
    For i = 1 To n
        If istDuplikatVon(i) = 0 Then
            best = i
            For j = i + 1 To n
                If istDuplikatVon(j) = i Then
                    If AnzahlGefuellt(daten, j) > AnzahlGefuellt(daten, best) Then best = j
                End If
            Next j
            If best <> i Then
                ' Rollen tauschen: best wird Hauptsatz
                istDuplikatVon(best) = 0
                istDuplikatVon(i) = best
                For j = 1 To n
                    If istDuplikatVon(j) = i Then istDuplikatVon(j) = best
                Next j
                hinweise(best) = hinweise(i)
            End If
        End If
    Next i

    Dim entfernt As Long
    For j = 1 To n
        k = istDuplikatVon(j)
        If k > 0 Then
            For f = 1 To ANZ_FELDER
                If Len(CStr(daten(k, f))) = 0 And Len(CStr(daten(j, f))) > 0 Then
                    daten(k, f) = daten(j, f)
                    Protokolliere k + 1, CStr(daten(k, 1)), CStr(feldnamen(f - 1)), _
                                  "", CStr(daten(j, f)), _
                                  Uml("Wert aus Duplikat (Zeile " & j + 1 & ") ~ubernommen")
                End If
            Next f
            Protokolliere j + 1, CStr(daten(j, 1)), "-", CStr(daten(j, 1)), "", _
                          "Duplikat von Zeile " & k + 1 & " (" & CStr(daten(k, 1)) & ") - entfernt"
            entfernt = entfernt + 1
        End If
    Next j

    ' Hinweise ggf. aktualisieren (Felder koennten aufgefuellt worden sein)
    For i = 1 To n
        If istDuplikatVon(i) = 0 Then
            hinweise(i) = ""
            If Len(CStr(daten(i, 7))) > 0 And Not EmailGueltig(CStr(daten(i, 7))) Then _
                HinweisAnfuegen hinweise(i), Uml("E-Mail ung~ultig")
            If Len(CStr(daten(i, 1))) = 0 Then HinweisAnfuegen hinweise(i), "Firma fehlt"
            If Len(CStr(daten(i, 4))) = 0 Then HinweisAnfuegen hinweise(i), "PLZ fehlt"
            If Len(CStr(daten(i, 6))) = 0 Then HinweisAnfuegen hinweise(i), "Telefon fehlt"
            If Len(CStr(daten(i, 7))) = 0 Then HinweisAnfuegen hinweise(i), "E-Mail fehlt"
        End If
    Next i

    ' --------------------------------------------------- Ergebnis schreiben
    Dim unvollstaendig As Long
    unvollstaendig = SchreibeErgebnis(daten, istDuplikatVon, hinweise, n, feldnamen, zielPfad)

    If Not automatik Then
        MsgBox "Fertig!" & vbCrLf & _
               n & " Zeilen eingelesen" & vbCrLf & _
               entfernt & " Duplikate entfernt" & vbCrLf & _
               protAnzahl & " Aenderungen protokolliert" & vbCrLf & _
               unvollstaendig & Uml(" Datens~atze mit Hinweisen (gelb markiert)"), _
               vbInformation, "Stammdaten-Bereinigung"
    End If
End Sub

' =====================================================================
' Hilfsfunktionen
' =====================================================================

' Platzhalter ~u ~a ~o ~s usw. -> echte Umlaute (Code bleibt reines ASCII)
Private Function Uml(ByVal s As String) As String
    s = Replace$(s, "~u", ChrW$(252)): s = Replace$(s, "~U", ChrW$(220))
    s = Replace$(s, "~a", ChrW$(228)): s = Replace$(s, "~A", ChrW$(196))
    s = Replace$(s, "~o", ChrW$(246)): s = Replace$(s, "~O", ChrW$(214))
    s = Replace$(s, "~s", ChrW$(223))
    Uml = s
End Function

Private Function BereinigeLeerzeichen(ByVal s As String) As String
    s = Trim$(s)
    Do While InStr(s, "  ") > 0
        s = Replace$(s, "  ", " ")
    Loop
    BereinigeLeerzeichen = s
End Function

' GROSSSCHREIBUNG / kleinschreibung -> normale Schreibweise inkl.
' korrekter Rechtsformen (GmbH, KG, AG, UG, e.K.)
Private Function NormalisiereFirma(ByVal s As String) As String
    If Len(s) = 0 Then Exit Function
    If s <> UCase$(s) And s <> LCase$(s) Then
        NormalisiereFirma = s          ' gemischte Schreibweise: nicht anfassen
        Exit Function
    End If
    Dim t As String
    t = StrConv(s, vbProperCase)
    t = Replace$(t, " Gmbh", " GmbH")
    t = Replace$(t, " Kg", " KG")
    t = Replace$(t, " Ag", " AG")
    t = Replace$(t, " Ug", " UG")
    t = Replace$(t, " E.k.", " e.K.")
    t = Replace$(t, " It-", " IT-")
    NormalisiereFirma = t
End Function

' Einheitliches Telefonformat: +49 Vorwahl Nummer
Private Function NormalisiereTelefon(ByVal s As String) As String
    If Len(Trim$(s)) = 0 Then Exit Function

    Dim ziffern As String, i As Long, c As String
    For i = 1 To Len(s)
        c = Mid$(s, i, 1)
        If c >= "0" And c <= "9" Then ziffern = ziffern & c
    Next i
    If Len(ziffern) < 6 Then NormalisiereTelefon = Trim$(s): Exit Function

    Dim national As String
    If Left$(ziffern, 4) = "0049" Then
        national = Mid$(ziffern, 5)
    ElseIf Left$(ziffern, 2) = "49" And Left$(Trim$(s), 1) = "+" Then
        national = Mid$(ziffern, 3)
    ElseIf Left$(ziffern, 1) = "0" Then
        national = Mid$(ziffern, 2)
    Else
        national = ziffern
    End If

    ' bekannte Vorwahlen (laengste zuerst pruefen)
    Dim vorwahlen As Variant, v As Variant
    vorwahlen = Array("3501", "3591", "3521", "3731", "351", "341", "371", _
                      "221", "211", "231", "30", "40", "69", "89")
    For Each v In vorwahlen
        If Left$(national, Len(v)) = v Then
            NormalisiereTelefon = "+49 " & v & " " & Mid$(national, Len(v) + 1)
            Exit Function
        End If
    Next v
    NormalisiereTelefon = "+49 " & national
End Function

Private Function NormalisiereEmail(ByVal s As String) As String
    s = LCase$(Trim$(s))
    s = Replace$(s, " ", "")
    s = Replace$(s, ",", ".")
    Do While InStr(s, "..") > 0
        s = Replace$(s, "..", ".")
    Loop
    NormalisiereEmail = s
End Function

Private Function EmailGueltig(ByVal s As String) As Boolean
    Static rx As Object
    If rx Is Nothing Then
        Set rx = CreateObject("VBScript.RegExp")
        rx.Pattern = "^[a-z0-9._%+\-]+@[a-z0-9.\-]+\.[a-z]{2,}$"
        rx.IgnoreCase = True
    End If
    EmailGueltig = rx.Test(s)
End Function

' Vergleichsschluessel fuer die Duplikatsuche:
' klein, Umlaute transliteriert, Rechtsformen und Sonderzeichen entfernt
Private Function VergleichsSchluessel(ByVal s As String) As String
    Dim t As String
    t = LCase$(s)
    t = Replace$(t, ChrW$(228), "ae"): t = Replace$(t, ChrW$(246), "oe")
    t = Replace$(t, ChrW$(252), "ue"): t = Replace$(t, ChrW$(223), "ss")

    Dim erg As String, i As Long, c As String
    For i = 1 To Len(t)
        c = Mid$(t, i, 1)
        If (c >= "a" And c <= "z") Or (c >= "0" And c <= "9") Or c = " " Then
            erg = erg & c
        Else
            erg = erg & " "
        End If
    Next i

    Dim teile() As String, teil As Variant, kern As String
    teile = Split(BereinigeLeerzeichen(erg), " ")
    For Each teil In teile
        Select Case CStr(teil)
            Case "gmbh", "ag", "kg", "ug", "ek", "e", "k", "co", "mbh"
                ' Rechtsform ignorieren
            Case Else
                kern = kern & CStr(teil)
        End Select
    Next teil
    VergleichsSchluessel = kern
End Function

' klassische Levenshtein-Distanz (zwei Zeilen der DP-Matrix)
Private Function Levenshtein(ByVal a As String, ByVal b As String) As Long
    Dim la As Long, lb As Long, i As Long, j As Long, kosten As Long
    la = Len(a): lb = Len(b)
    If la = 0 Then Levenshtein = lb: Exit Function
    If lb = 0 Then Levenshtein = la: Exit Function

    Dim v0() As Long, v1() As Long
    ReDim v0(0 To lb): ReDim v1(0 To lb)
    For j = 0 To lb: v0(j) = j: Next j

    For i = 1 To la
        v1(0) = i
        For j = 1 To lb
            kosten = IIf(Mid$(a, i, 1) = Mid$(b, j, 1), 0, 1)
            v1(j) = v0(j) + 1
            If v1(j - 1) + 1 < v1(j) Then v1(j) = v1(j - 1) + 1
            If v0(j - 1) + kosten < v1(j) Then v1(j) = v0(j - 1) + kosten
        Next j
        For j = 0 To lb: v0(j) = v1(j): Next j
    Next i
    Levenshtein = v1(lb)
End Function

Private Function AnzahlGefuellt(ByRef daten As Variant, ByVal zeile As Long) As Long
    Dim f As Long
    For f = 1 To ANZ_FELDER
        If Len(CStr(daten(zeile, f))) > 0 Then AnzahlGefuellt = AnzahlGefuellt + 1
    Next f
End Function

Private Sub HinweisAnfuegen(ByRef hinweis As String, ByVal text As String)
    If Len(hinweis) > 0 Then hinweis = hinweis & "; "
    hinweis = hinweis & text
End Sub

Private Sub Protokolliere(ByVal zeile As Long, ByVal firma As String, ByVal feld As String, _
                          ByVal vorher As String, ByVal nachher As String, ByVal grund As String)
    protAnzahl = protAnzahl + 1
    If protAnzahl > UBound(protokoll, 2) Then
        ReDim Preserve protokoll(1 To 6, 1 To UBound(protokoll, 2) + 1000)
    End If
    protokoll(1, protAnzahl) = zeile
    protokoll(2, protAnzahl) = firma
    protokoll(3, protAnzahl) = feld
    protokoll(4, protAnzahl) = vorher
    protokoll(5, protAnzahl) = nachher
    protokoll(6, protAnzahl) = grund
End Sub

' ---------------------------------------------------------------------
' Schreibt bereinigte Liste + Aenderungsprotokoll in eine neue Datei.
' Rueckgabe: Anzahl Datensaetze mit Hinweisen
' ---------------------------------------------------------------------
Private Function SchreibeErgebnis(ByRef daten As Variant, ByRef istDuplikatVon() As Long, _
                                  ByRef hinweise() As String, ByVal n As Long, _
                                  ByVal feldnamen As Variant, ByVal zielPfad As String) As Long

    Dim wb As Workbook, ws As Worksheet, wsP As Worksheet
    Set wb = Workbooks.Add
    Set ws = wb.Worksheets(1)
    ws.Name = "Kunden bereinigt"

    Dim c As Long
    For c = 0 To UBound(feldnamen)
        KopfZelle ws.Cells(1, c + 1), CStr(feldnamen(c))
    Next c
    KopfZelle ws.Cells(1, ANZ_FELDER + 1), "Hinweis"

    Dim i As Long, z As Long, f As Long, mitHinweis As Long
    z = 1
    For i = 1 To n
        If istDuplikatVon(i) = 0 Then
            z = z + 1
            For f = 1 To ANZ_FELDER
                If f = 4 Then
                    ws.Cells(z, f).NumberFormat = "@"   ' PLZ als Text (fuehrende 0)
                End If
                ws.Cells(z, f).Value = CStr(daten(i, f))
            Next f
            If Len(hinweise(i)) > 0 Then
                ws.Cells(z, ANZ_FELDER + 1).Value = hinweise(i)
                ws.Range(ws.Cells(z, 1), ws.Cells(z, ANZ_FELDER + 1)).Interior.Color = RGB(255, 235, 156)
                mitHinweis = mitHinweis + 1
            End If
        End If
    Next i

    With ws.Range(ws.Cells(1, 1), ws.Cells(z, ANZ_FELDER + 1))
        .Borders.Color = RGB(191, 191, 191)
        .AutoFilter
    End With
    ws.Columns("A:H").AutoFit

    ' ------------------------------------------------------ Protokollblatt
    Set wsP = wb.Worksheets.Add(After:=ws)
    wsP.Name = "Aenderungsprotokoll"

    Dim pKopf As Variant
    pKopf = Array("Zeile (Quelle)", "Firma", "Feld", "Vorher", "Nachher", "Grund")
    For c = 0 To 5
        KopfZelle wsP.Cells(1, c + 1), CStr(pKopf(c))
    Next c

    For i = 1 To protAnzahl
        For c = 1 To 6
            wsP.Cells(i + 1, c).Value = protokoll(c, i)
        Next c
        If InStr(CStr(protokoll(6, i)), "Duplikat von") > 0 Then
            wsP.Range(wsP.Cells(i + 1, 1), wsP.Cells(i + 1, 6)).Interior.Color = RGB(255, 199, 206)
        End If
    Next i

    With wsP.Range(wsP.Cells(1, 1), wsP.Cells(protAnzahl + 1, 6))
        .Borders.Color = RGB(191, 191, 191)
        .AutoFilter
    End With
    wsP.Columns("A:F").AutoFit

    Application.DisplayAlerts = False
    wb.SaveAs zielPfad, xlOpenXMLWorkbook
    wb.Close SaveChanges:=False
    Application.DisplayAlerts = True

    SchreibeErgebnis = mitHinweis
End Function

Private Sub KopfZelle(ByVal zelle As Range, ByVal text As String)
    With zelle
        .Value = text
        .Font.Bold = True
        .Font.Color = vbWhite
        .Interior.Color = RGB(31, 78, 121)
    End With
End Sub
