Attribute VB_Name = "ATW_Check"
Dim som12uur2weken As Integer 'max5
Dim som12uur52weken As Integer 'max22
Dim somAuur1week As Integer 'max60 zon00totzat24 o
Dim somAuur4weken As Integer 'max220 o
Dim somAuur16weken As Integer 'met nacht 640 zonder 768 o
Dim somNacht16weken As Integer 'max36 o
Dim somNacht52weken As Integer 'max140 o
Dim somNachtUren2weken As Integer 'max38 00-06 o
Dim som14uur2weken As Integer '1
Dim somRust12uur As Integer 'min 12
Dim somRustNacht As Integer 'min 14
Dim somAEV As Integer 'met nacht max 7(8) (AEV<32uur)
Dim somNachtAEV As Integer '
Dim somRust3Nacht As Integer 'min 46
Dim somRust24uur As Integer 'min 11uur
Dim somRust1weken As Integer 'min 36
Dim somRust2weken As Integer 'min 72
Dim somZon52weken As Integer 'max 13


Function dienstopdatum(datum, persoon)
functie = "OPR"
opnieuw:
With Workbooks("ROOSTERBORD_" & functie & Format(CDate(datum), "yy") & ".xls").Sheets(Left(UCase(Format(CDate(datum), "mmmm")), 3)).Range("B1:B75")
Set c = .Find(persoon, LookAt:=xlValue)
alala = c.Row
If c Is Nothing Then
If functie <> "wacht_kraan" Then
functie = "wacht_kraan"
GoTo opnieuw
Else
dienstopdatum = "Persoon niet gevonden"
End If
Else
If Workbooks("ROOSTERBORD_" & functie & Format(CDate(datum), "yy") & ".xls").Sheets(Left(UCase(Format(CDate(datum), "mmmm")), 3)).Cells(c.Row + 1, 3 + Day(CDate(datum))).Value <> "" Then
If Workbooks("ROOSTERBORD_" & functie & Format(CDate(datum), "yy") & ".xls").Sheets(Left(UCase(Format(CDate(datum), "mmmm")), 3)).Cells(c.Row + 2, 3 + Day(CDate(datum))).Value <> "" Then
dienstopdatum = Workbooks("ROOSTERBORD_" & functie & Format(CDate(datum), "yy") & ".xls").Sheets(Left(UCase(Format(CDate(datum), "mmmm")), 3)).Cells(c.Row + 2, 3 + Day(CDate(datum))).Value
Else
dienstopdatum = Workbooks("ROOSTERBORD_" & functie & Format(CDate(datum), "yy") & ".xls").Sheets(Left(UCase(Format(CDate(datum), "mmmm")), 3)).Cells(c.Row + 1, 3 + Day(CDate(datum))).Value
End If
Else
dienstopdatum = Workbooks("ROOSTERBORD_" & functie & Format(CDate(datum), "yy") & ".xls").Sheets(Left(UCase(Format(CDate(datum), "mmmm")), 3)).Cells(c.Row, 3 + Day(CDate(datum))).Value
End If
If dienstopdatum = "RES" Or dienstopdatum = 0 Or dienstopdatum = "VRIJ" Or dienstopdatum = "BV" Or dienstopdatum = "VAK" Then
dienstopdatum = ""
End If
End If

End With
End Function

Function checkAuur1week(datum, persoon)

datum2 = datum
'Format(CDate(Now), "ddd", vbMonday)
While Format(CDate(datum2), "ddd", vbMonday) <> "zo"
datum2 = datum2 - 1
Wend
If dienstopdatum(datum2 - 1, persoon) = "N" Then
    somAuur1week = somAuur1week + 7
End If
Do
resultD = dienstopdatum(datum2, persoon)

If resultD <> "" Then
somAuur1week = somAuur1week + 8

If resultD = "4" Then
somAuur1week = somAuur1week + 4
ElseIf resultD = "1" Then
somAuur1week = somAuur1week + 1
ElseIf resultD = "+1" Then
somAuur1week = somAuur1week + 1
ElseIf resultD = "-1" Then
somAuur1week = somAuur1week - 1
End If
End If
    
datum2 = datum2 + 1
Loop Until Format(CDate(datum2 - 1), "ddd", vbMonday) = "za"
If resultD = "N" Then
somAuur1week = somAuur1week - 7
End If

checkAuur1week = somAuur1week

somAuur1week = 0
End Function

Function checkAuur4weken(datum, persoon)
While Format(CDate(datum), "ddd", vbMonday) <> "zo"
datum = datum - 1
Wend

datum3 = datum - (3 * 7)

While datum3 <= datum

somAuur4weken = somAuur4weken + checkAuur1week(datum3, persoon)

datum3 = datum3 + 7
Wend
checkAuur4weken = somAuur4weken
somAuur4weken = 0

End Function

Function checkAuur16weken(datum, persoon)
While Format(CDate(datum), "ddd", vbMonday) <> "zo"
datum = datum - 1
Wend

datum3 = datum - (15 * 7)


While datum3 <= datum

somAuur16weken = somAuur16weken + checkAuur1week(datum3, persoon)

datum3 = datum3 + 7
Wend
checkAuur16weken = somAuur16weken
somAuur16weken = 0
End Function

Function checkNin16weken(datum, persoon)
While Format(CDate(datum), "ddd", vbMonday) <> "zo"
datum = datum - 1
Wend

datum3 = datum - (15 * 7)


If dienstopdatum(datum3 - 1, persoon) = "N" Then
somNacht16weken = somNacht16weken + 1
End If
While datum3 <= datum + 6
If dienstopdatum(datum3, persoon) = "N" Then
somNacht16weken = somNacht16weken + 1
End If
datum3 = datum3 + 1
Wend

checkNin16weken = somNacht16weken
somNacht16weken = 0
End Function

Function checkNin52weken(datum, persoon)
While Format(CDate(datum), "ddd", vbMonday) <> "zo"
datum = datum - 1
Wend

datum3 = datum - (51 * 7)

If dienstopdatum(datum3 - 1, persoon) = "N" Then
somNacht52weken = somNacht52weken + 1
End If
While datum3 <= datum + 6
If dienstopdatum(datum3, persoon) = "N" Then
somNacht52weken = somNacht52weken + 1
End If
datum3 = datum3 + 1
Wend

checkNin52weken = somNacht52weken
somNacht52weken = 0
End Function

Function checkNachtUren2weken(datum, persoon)
While Format(CDate(datum), "ddd", vbMonday) <> "zo"
datum = datum - 1
Wend

datum3 = datum - (1 * 7)


While datum3 <= datum + 6
If dienstopdatum(datum3 - 1, persoon) = "N" Then
somNachtUren2weken = somNachtUren2weken + 6
End If
datum3 = datum3 + 1
Wend

checkNachtUren2weken = somNachtUren2weken
somNachtUren2weken = 0
End Function

Function checkRust1week(datum, persoon) '36 aaneengesloten
While Format(CDate(datum), "ddd", vbMonday) <> "zo"
datum = datum - 1
Wend
datumX = datum + 7
temp = 0
somRust1weken = 0
If dienstopdatum(datum - 1, persoon) = "N" And dienstopdatum(datum, persoon) = "" Then
temp = 17
End If
While datum < datumX

Select Case dienstopdatum(datum, persoon)
    Case "V"
        temp = temp + 7
        If temp >= 32 Then
        somRust1weken = somRust1weken + 1
        temp = 0
        End If
    Case "M"
        temp = temp + 15
        If temp >= 32 Then
        somRust1weken = somRust1weken + 1
        temp = 0
        End If
    Case "N"
        temp = temp + 23
        If temp >= 32 Then
        somRust1weken = somRust1weken + 1
        temp = 0
        End If
    Case "D"
        temp = temp + 8
        If temp >= 32 Then
        somRust1weken = somRust1weken + 1
        temp = 0
        End If
    Case ""
        temp = temp + 24
         If temp >= 32 Then
        somRust1weken = somRust1weken + 1
        
        End If
End Select

datum = datum + 1
Wend

checkRust1week = somRust1weken
somRust1weken = 0
temp = 0
End Function

Function checkRust2weken(datum, persoon) '72 aaneengesloten of min 32
While Format(CDate(datum), "ddd", vbMonday) <> "zo"
datum = datum - 1
Wend
datumX = datum + 7
datum = datum - 7
temp = 0
somRust1weken = 0
If dienstopdatum(datum - 1, persoon) = "N" And dienstopdatum(datum, persoon) = "" Then
temp = 17
End If
While datum < datumX

Select Case dienstopdatum(datum, persoon)
    Case "V"
        temp = temp + 7
        If temp >= 72 Then
        somRust2weken = somRust2weken + 1
        temp = 0
        End If
    Case "M"
        temp = temp + 15
        If temp >= 72 Then
        somRust2weken = somRust2weken + 1
        temp = 0
        End If
    Case "N"
        temp = temp + 23
        If temp >= 72 Then
        somRust2weken = somRust2weken + 1
        temp = 0
        End If
    Case "D"
        temp = temp + 8
        If temp >= 72 Then
        somRust2weken = somRust2weken + 1
        temp = 0
        End If
    Case ""
        temp = temp + 24
         If temp >= 72 Then
        somRust2weken = somRust2weken + 1
        
        End If
End Select

datum = datum + 1
Wend

checkRust2weken = somRust2weken
somRust2weken = 0
temp = 0
End Function
