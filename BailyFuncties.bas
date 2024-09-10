Attribute VB_Name = "BailyFuncties"
'===============================================
'@details       Script for simulating Baily Functionblock (ABB Baily Automation DCS)
'@author        Jordi van Nistelrooij @ Webs en Systems
'@email         info@websensystems.nl
'@version       1.0.0
'@date          2024-09-10
'@copyright     Non of these scripts maybe copied or modified without permission of the author
'===============================================

'1
Function Baily_Fx(s1, s2, s3, s4, s5, s6, s7, s8, s9, s10, s11, s12, s13)
    If s1 < s2 Then
        Baily_Fx = s3
    ElseIf s1 > s2 And s1 < s4 Then
        Baily_Fx = (((s1 - s2) / (s4 - s2)) * (s5 - s3)) + s3
    ElseIf s1 > s4 And s1 < s6 Then
        Baily_Fx = (((s1 - s4) / (s6 - s4)) * (s7 - s5)) + s5
    ElseIf s1 > s6 And s1 < s8 Then
        Baily_Fx = (((s1 - s6) / (s8 - s6)) * (s9 - s7)) + s7
    ElseIf s1 > s8 And s1 < s10 Then
        Baily_Fx = (((s1 - s8) / (s10 - s8)) * (s11 - s9)) + s9
    ElseIf s1 > s10 And s1 < s12 Then
        Baily_Fx = (((s1 - s10) / (s12 - s10)) * (s13 - s11)) + s11
    End If
End Function
'2
Function Baily_A(s1)
    Baily_A = s1
End Function
'3
Function Baily_Ft(s1, s2, s3, s4, X, x2)
    If s2 = 0 Then
        Baily_Ft = s1
    Else
        ' hier een of andere manier om lag en lead te intergreren
        ' lag is output wordt input in 5 (seconden) stappen
        ' lead is output is actuele input + verandering vorige tijd
    End If
End Function
'6
Function Baily_HiLowLim(s1, s2, s3)
    If s1 > s2 Then
    Baily_HiLowLim = s2
    ElseIf s1 < s3 Then
    Baily_HiLowLim = s3
    Else
    Baily_HiLowLim = s1
    End If
End Function
'7
Function Baily_SquareRoot(s1, s2)
Baily_SquareRoot = s2 * Sqr(s1)
End Function
'8
Function Baily_RateLimiter(s1, s2, s3, s4)
If s2 = 0 Then
    Baily_RateLimiter = s1
Else
    'verandering is maximaal +s3/s of -s4/s
End If

End Function
'9
Function Baily_T(s1, s2, s3, s4, s5)
    If s3 = 0 Then
        Baily_T = s1
    Else
        Baily_T = s2
    End If
    'when switched 5 time constans * s4,s5
End Function
'10
Function Baily_BIG(s1, s2, s3, s4)
    groot = s1
    
    If groot < s2 Then
        groot = s2
    End If
    
    If groot < s3 Then
        groot = s3
    End If
  
    If groot < s4 Then
        groot = s4
    End If
    Baily_BIG = groot

End Function
'11

Function Baily_SMALL(s1, s2, s3, s4)
    groot = s1
    
    If groot > s2 Then
        groot = s2
    End If
    
    If groot > s3 Then
        groot = s3
    End If
  
    If groot > s4 Then
        groot = s4
    End If
    Baily_SMALL = groot

End Function
'12
Function Baily_HL_N(s1, s2, s3)
    If s1 >= s2 Then
        Baily_HL_N = 1
    End If
End Function

Function Baily_HL_N1(s1, s2, s3)
    If s1 <= s3 Then
        Baily_HL_N1 = 1
    End If
End Function

'14
Function Baily_Sum(s1, s2, s3, s4)
    Baily_Sum = s1 + s2 + s3 + s4
End Function
'15
Function Baily_SumK(s1, s2, s3, s4)
    Baily_SumK = (s1 * s3) + (s2 * s4)
End Function
'16
Function Baily_X(s1, s2, s3)
    Baily_X = s3 * (s1 * s2)
End Function
'17
Function Baily_Divide(s1, s2, s3)
    Baily_D = s3 * (s1 / s2)
End Function
'18
Function Baily_PID(s1, s2, s3, s4, s5, s6, s7, s8, s9, s10)
    If s4 = 0 Then
        Baily_PID = s3
    Else
        getal = s5 * (s1 * (s6 + s7 + s8))
        If getal > s9 Then
            getal = s9
        End If
    
        If getal < s10 Then
            getal = s10
        End If
        Baily_PID = getal
    End If
End Function

Function Baily_APID(s2, s1, s3, s4, s5, s6, s7, s8, s9, s10, s11, s12, s13, s14, s15, s16, s17, s18, s19, s20, s21)
    If s4 = 0 Then
    Baily_APID = s3
    Else
    If s21 = 0 Then
    errorIN = s2 - s1
    Else
    errorIN = s1 - s2
    End If
    
   If s18 = 0 Then
   getal = s11 * s12 * (1 + (s13 / 60) / (1 / 60)) * ((60 * s14 * (1 / 60) + 1) / (((60 * s14) / s15) * (1 / 60) + 1)) * errorIN
   ElseIf s18 = 1 Then
   getal = s11 * (s12 + (s13 / 60) / (1 / 60) + ((60 * s14 * (1 / 60)) / (((60 * s14) / s15) * (1 / 60) + 1))) * errorIN
   End If
   
   If getal > s16 Then
   Baily_APID = s16
   ElseIf getal < s17 Then
   Baily_APID = s17
   Else
   Baily_APID = getal
   End If
   End If
End Function

'33
Function Baily_Not(s1)
If s1 = 1 Then
Baily_Not = 0
Else
Baily_Not = 1
End If
End Function

'36
Function Baily_QOR(s1, s2, s3, s4, s5, s6, s7, s8, s9, s10)
sumo = s1 + s2 + s3 + s4 + s5 + s6 + s7 + s8
If s10 = 0 Then
If sumo >= s9 Then
Baily_QOR = 1
Else
Baily_QOR = 0
End If
Else
If sumo = s9 Then
Baily_QOR = 1
Else
Baily_QOR = 0
End If
End If
End Function

'37
Function Baily_And(s1, s2)
If s1 = 1 And s2 = 1 Then
Baily_And = 1
Else
Baily_And = 0
End If
End Function

'38
Function Baily_And4(s1, s2, s3, s4)
If s1 + s2 + s3 + s4 = 4 Then
Baily_And4 = 1
Else
Baily_And4 = 0
End If
End Function

'39
Function Baily_Or(s1, s2)
If s1 = 1 Or s2 = 1 Then
Baily_Or = 1
Else
Baily_Or = 0
End If
End Function

'40
Function Baily_Or4(s1, s2, s3, s4)
If s1 + s2 + s3 + s4 > 0 Then
Baily_Or4 = 1
Else
Baily_Or4 = 0
End If
End Function

'59
Function Baily_TDIG(s1, s2, s3)
If s3 = 0 Then
Baily_TDIG = s1
Else
Baily_TDIG = s2
End If
End Function

'65
Function Baily_DSUM(s1, s2, s3, s4, s5, s6, s7, s8)
If s5 = "" Then s5 = 1
If s6 = "" Then s6 = 1
If s7 = "" Then s7 = 1
If s8 = "" Then s8 = 1
somu = 0
If s1 <> 0 Then
somu = somu + (s5 * s1)
End If
If s2 <> 0 Then
somu = somu + (s6 * s2)
End If
If s3 <> 0 Then
somu = somu + (s7 * s3)
End If
If s4 <> 0 Then
somu = somu + (s8 * s4)
End If
Baily_DSUM = somu
End Function

'101
Function Baily_XOR(s1, s2)
If s1 + s2 = 1 Then
Baily_XOR = 1
Else
Baily_XOR = 0
End If
End Function

