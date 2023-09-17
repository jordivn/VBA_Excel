Attribute VB_Name = "Formules_Algemeen"
Dim getal As Double

Function WillekeurigGetal(Optional Laagste = 0, Optional Hoogste = 1, Optional Complexiteit = 1)
Attribute WillekeurigGetal.VB_Description = "Geeft een willekeurig getal"
Attribute WillekeurigGetal.VB_ProcData.VB_Invoke_Func = " \n23"
getal = ((Hoogste - Laagste) * Rnd + Laagste)

While Complexiteit > 1
getal = (getal + ((Hoogste - Laagste) * Rnd + Laagste)) / 2
Complexiteit = Complexiteit - 1
Wend

WillekeurigGetal = getal


End Function

Function Hex2Dec(HexString As Variant) As Variant
Attribute Hex2Dec.VB_Description = "Zet een Heximaal getal om naar decimaal"
Attribute Hex2Dec.VB_ProcData.VB_Invoke_Func = " \n23"
    Dim X As Integer
    For X = 0 To Len(HexString) - 1
    TASADSA = UCase(Mid(HexString, Len(HexString) - X, 1))
        Select Case UCase(Mid(HexString, Len(HexString) - X, 1))
        Case "A"
            ANumber = 10
        Case "B"
            ANumber = 11
        Case "C"
            ANumber = 12
        Case "D"
        ANumber = 13
        Case "E"
        ANumber = 14
        Case "F"
        ANumber = 15
        Case Else
        ANumber = Val(Mid(HexString, Len(HexString) - X, 1))
        End Select
        Hex2Dec = CDec(Hex2Dec) + ANumber * 16 ^ X
    Next
End Function




Function Dec2Bin(ByVal DecimalIn As Variant, _
              Optional NumberOfBits As Variant) As String
Attribute Dec2Bin.VB_Description = "Zet een decimaal getal om in een binaire notatie"
Attribute Dec2Bin.VB_ProcData.VB_Invoke_Func = " \n23"
    Dec2Bin = ""
    DecimalIn = Int(CDec(DecimalIn))
    Do While DecimalIn <> 0
        Dec2Bin = Format$(DecimalIn - 2 * Int(DecimalIn / 2)) & Dec2Bin
        DecimalIn = Int(DecimalIn / 2)
    Loop
    If Not IsMissing(NumberOfBits) Then
       If Len(Dec2Bin) > NumberOfBits Then
          Dec2Bin = "Error - Number exceeds specified bit size"
       Else
          Dec2Bin = Right$(String$(NumberOfBits, _
                    "0") & Dec2Bin, NumberOfBits)
       End If
    End If
End Function
 
'Binary To Decimal
' =================
Function Bin2Dec(BinaryString As String) As Variant
Attribute Bin2Dec.VB_Description = "Zet een binair getal om naar het decimale stelsel"
Attribute Bin2Dec.VB_ProcData.VB_Invoke_Func = " \n23"
    Dim X As Integer
    For X = 0 To Len(BinaryString) - 1
    
        Bin2Dec = CDec(Bin2Dec) + Val(Mid(BinaryString, Len(BinaryString) - X, 1)) * 2 ^ X
    Next
End Function


