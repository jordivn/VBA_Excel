Attribute VB_Name = "Tijd_Functie"
Function VolgendeHeleUur(Optional datetime = "", Optional FormatStyle = "")
Attribute VolgendeHeleUur.VB_Description = "Geeft de tijd van het volgende hele uur weer"
Attribute VolgendeHeleUur.VB_ProcData.VB_Invoke_Func = " \n21"
    If datetime = "" Then
        datetime = Now
    End If
   If FormatStyle = "" Then
        FormatStyle = "hh:mm:ss"
   End If
    
    VolgendeHeleUur = Format(Application.WorksheetFunction.RoundDown(CDate(datetime), 0) + (1 / 24 * (Hour(CDate(datetime)) + 1)), FormatStyle)
    
    

End Function

Function LaatsteHeleUur(Optional datetime = "", Optional FormatStyle = "", Optional Move = "")
Attribute LaatsteHeleUur.VB_Description = "Geeft de tijd van het laatste hele uur weer"
Attribute LaatsteHeleUur.VB_ProcData.VB_Invoke_Func = " \n21"
    If datetime = "" Then
        datetime = Now
    End If
   If FormatStyle = "" Then
        FormatStyle = "hh:mm:ss"
   End If
   
   If Move = "" Then
        Move = 0
    End If
   
    
    LaatsteHeleUur = Format(Application.WorksheetFunction.RoundDown(CDate(datetime), 0) + (1 / 24 * (Hour(CDate(datetime)) + CInt(Move))), FormatStyle)
    
    

End Function

Function TijdEindeWacht(Optional datetime = "", Optional FormatStyle = "")
Attribute TijdEindeWacht.VB_Description = "Geeft de tijd van het einde van de wacht weer"
Attribute TijdEindeWacht.VB_ProcData.VB_Invoke_Func = " \n21"
    If datetime = "" Then
        datetime = Now
    End If
   If FormatStyle = "" Then
        FormatStyle = "hh:mm:ss"
   End If
    
    
    
    If Hour(CDate(datetime)) = 23 Then
    tijduitvoer = CDate(Day(datetime) & "-" & Month(datetime) & "-" & Year(datetime) & " 06:59:59") + 1 '7:00 volgende dag
    ElseIf Hour(CDate(datetime)) >= 15 Then
    tijduitvoer = CDate(Day(datetime) & "-" & Month(datetime) & "-" & Year(datetime) & " 22:59:59") '23:00
    ElseIf Hour(CDate(datetime)) >= 7 Then
    tijduitvoer = CDate(Day(datetime) & "-" & Month(datetime) & "-" & Year(datetime) & " 14:59:59") '15:00
    Else
    tijduitvoer = CDate(Day(datetime) & "-" & Month(datetime) & "-" & Year(datetime) & " 06:59:59")  '7:00
    End If
    
    TijdEindeWacht = Format(tijduitvoer, FormatStyle)
    
    

End Function

Function TijdStartWacht(Optional datetime = "", Optional FormatStyle = "")
Attribute TijdStartWacht.VB_Description = "Geeft de tijd van de start van de wacht weer"
Attribute TijdStartWacht.VB_ProcData.VB_Invoke_Func = " \n21"
    If datetime = "" Then
        datetime = Now
    End If
   If FormatStyle = "" Then
        FormatStyle = "hh:mm:ss"
   End If
    
    
    
    If Hour(CDate(datetime)) = 23 Then
    tijduitvoer = CDate(Day(datetime) & "-" & Month(datetime) & "-" & Year(datetime) & " 23:00:00") '7:00 volgende dag
    ElseIf Hour(CDate(datetime)) >= 15 Then
    tijduitvoer = CDate(Day(datetime) & "-" & Month(datetime) & "-" & Year(datetime) & " 15:00:00") '23:00
    ElseIf Hour(CDate(datetime)) >= 7 Then
    tijduitvoer = CDate(Day(datetime) & "-" & Month(datetime) & "-" & Year(datetime) & " 07:00:00") '15:00
    Else
    tijduitvoer = CDate(Day(datetime) & "-" & Month(datetime) & "-" & Year(datetime) & " 23:00:00") - 1 '7:00
    End If
    
    TijdStartWacht = Format(tijduitvoer, FormatStyle)
    
    

End Function

Function StartWeek(Optional datetime = "", Optional FormatStyle = "")
Attribute StartWeek.VB_Description = "Geeft de datum van de start van de week weer"
Attribute StartWeek.VB_ProcData.VB_Invoke_Func = " \n21"
    If datetime = "" Then
        datetime = Now
    End If
   If FormatStyle = "" Then
        FormatStyle = "d-m-yyyy"
   End If
   
   tijduitvoer = CDate(Day(datetime) & "-" & Month(datetime) & "-" & Year(datetime))
   tijduitvoer = tijduitvoer - Weekday(tijduitvoer, vbMonday) + 1
   
  
   StartWeek = Format(tijduitvoer, FormatStyle)

End Function

Function EindWeek(Optional datetime = "", Optional FormatStyle = "")
Attribute EindWeek.VB_Description = "Geeft de datum van het einde van de week"
Attribute EindWeek.VB_ProcData.VB_Invoke_Func = " \n21"
    If datetime = "" Then
        datetime = Now
    End If
   If FormatStyle = "" Then
        FormatStyle = "d-m-yyyy"
   End If
  tijduitvoer = CDate(Day(datetime) & "-" & Month(datetime) & "-" & Year(datetime))
   tijduitvoer = tijduitvoer + Weekday(tijduitvoer, vbMonday) + 2 - (1 / 24 / 60 / 60)
   
  
   EindWeek = Format(tijduitvoer, FormatStyle)
   
End Function

Function StartCyclus(Optional datetime = "", Optional FormatStyle = "")
Attribute StartCyclus.VB_Description = "Geeft de datum van de start van de cyclus weer"
Attribute StartCyclus.VB_ProcData.VB_Invoke_Func = " \n21"


    If datetime = "" Then
        datetime = Now
    End If
   If FormatStyle = "" Then
        FormatStyle = "d-m-yyyy hh:mm:ss"
   End If
   actueleWacht = TijdStartWacht(datetime, "d-m-yyyy hh:mm:ss")
   
   sdate = CDate(actueleWacht)
   X = Hour(CDate(actueleWacht))
   Select Case Hour(CDate(actueleWacht))
        Case 7
            If Weekday(CDate(actueleWacht), vbMonday) = 2 Then
            sdate = sdate - 1
            ElseIf Weekday(CDate(actueleWacht), vbMonday) = 4 Then
            sdate = sdate - 1
            ElseIf Weekday(CDate(actueleWacht), vbMonday) = 6 Then
            sdate = sdate - 1
            ElseIf Weekday(CDate(actueleWacht), vbMonday) = 7 Then
            sdate = sdate - 2
            
            End If
        Case 15
           Select Case Weekday(CDate(actueleWacht), vbMonday)
                Case 1
                    sdate = sdate - 3
                Case 2
                    sdate = sdate - 4
                Case 3
                    sdate = sdate - 2
                Case 4
                    sdate = sdate - 3
                Case 5
                    sdate = sdate - 2
                Case 6
                    sdate = sdate - 3
                Case 7
                    sdate = sdate - 4
           End Select
           
        Case 23
                 Select Case Weekday(CDate(actueleWacht), vbMonday)
                Case 1
                    sdate = sdate - 5
                Case 2
                    sdate = sdate - 6
                Case 3
                    sdate = sdate - 4
                Case 4
                    sdate = sdate - 5
                Case 5
                    sdate = sdate - 4
                Case 6
                    sdate = sdate - 5
                Case 7
                    sdate = sdate - 6
           End Select
   
   
   End Select
       StartCyclus = CDate(Day(sdate) & "-" & Month(sdate) & "-" & Year(sdate) & " 07:00:00")
   
   
End Function
Function EindeCyclus(Optional datetime = "", Optional FormatStyle = "")
Attribute EindeCyclus.VB_Description = "Geeft de datum van het einde van de cyclus weer"
Attribute EindeCyclus.VB_ProcData.VB_Invoke_Func = " \n21"
    If datetime = "" Then
        datetime = Now
    End If
   If FormatStyle = "" Then
        FormatStyle = "d-m-yyyy hh:mm:ss"
   End If
sdate = StartCyclus(datetime) + 7
 EindeCyclus = CDate(Day(sdate) & "-" & Month(sdate) & "-" & Year(sdate) & " 06:59:59")
End Function

