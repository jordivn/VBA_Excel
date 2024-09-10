Attribute VB_Name = "Basis_Functies"
Public RunFirst As Boolean
Public Version As String
Public ConnectionString As String

Sub CheckAndDisplay()
If ActiveWorkbook.Name <> "Jtools Update" Then
    NewVersionWindow.show
End If
End Sub

Sub CheckForUpdate()



    Dim xmlhttp As Object
    Set xmlhttp = CreateObject("MSXML2.serverXMLHTTP")
    
    
    xmlhttp.Open "GET", "https://websensystems.nl/JTools/version.php", False
    
    xmlhttp.Send
    AvailibleVersion = xmlhttp.responseText
        
    
   
        
        'Basis_Functies.Version = textline
        
        
        If AvailibleVersion <> Basis_Functies.Version Then
        
        ' UPDATE FUNCTIE
        ' Moet module (.bas) importeren in bestaande workbook. Zichzelf uitschakelen. Nieuwe copyen. Herladen
        
        
'            szTargetWorkbook = ActiveWorkbook.Name
'           Set wkbTarget = Application.Workbooks(szTargetWorkbook)
            
            
            'Set cmpComponents = wkbTarget.VBProject.VBComponents
            'cmpComponents.import "U:\Excels\Modules\X_steam_Tables.bas"
            
        
        
 '           fs.CopyFile Source:=ThisWorkbook.FullName, Destination:=ThisWorkbook.FullName & ".old"
            
            'fs.CopyFile Source:="P:\Stortkok\Jordi\jtools.xlam", Destination:=ThisWorkbook.FullName
 '           MsgBox ("Jtools updated. Start excel opnieuw op.")
 
            NewVersionWindow.Label2.Caption = "Versie: " & AvailibleVersion
            Application.OnTime Now + TimeValue("00:00:03"), " Basis_Functies.CheckAndDisplay"
            
       
        End If
        
End Sub


Sub doLogging()
    'On Error Resume Next
   '     Open "P:\Stortkok\Jordi\active.logging" For Append As 1
   '     Print #1, Format(Now, "dd-mm-YYYY hh:mm:ss") & ";" & Application.UserName & ";" & ActiveWorkbook.FullName & ";" & ThisWorkbook.FullName
   '     Close #1
    On Error Resume Next
    Dim xmlhttp As Object
    Set xmlhttp = CreateObject("MSXML2.serverXMLHTTP")
    Dim myURL As String
    If Basis_Functies.Version = "" Then
        Call InstelFuncties.GetSettingsOfUser
    End If
    
    Dim strEnviron As String
    Dim I As Long
    For I = 1 To 255
        strEnviron = Environ(I)
        If LenB(strEnviron) = 0& Then Exit For
        EnvString = EnvString & "#" & strEnviron
    Next
    
    strData = "dt=" & Format(Now, "yyyymmddhhmmss")
    strData = strData & ";User=" & Application.UserName
    strData = strData & ";Workbook=" & ActiveWorkbook.FullName
    strData = strData & ";Jtools=" & ThisWorkbook.FullName
    strData = strData & ";Version=" & Basis_Functies.Version
    'strData = strData & ";EnvMent=" & EnvString
    
    myURL = "https://websensystems.nl/JTools/getting.php?action=UserLog&msg=" + Base64EncodeString(strData)
    xmlhttp.Open "GET", myURL, False
    xmlhttp.setRequestHeader "Content-Type", "text/json"
    xmlhttp.Send
    'MsgBox (xmlhttp.responseText)
End Sub

Sub CheckAlarms()
    
    numAlarms = AlertScreen.ListBox1.ListCount
    AlertScreen.UserForm_Activate
    numAlarms2 = AlertScreen.ListBox1.ListCount
    If numAlarms2 > numAlarms And Not RunFirst Then
    Beep
    MsgBox ("Nieuwe Alarmen gedetecteerd")
    End If
    RunFirst = False
       
    TijdVolgendeCheck = CDate(Tijd_Functie.VolgendeHeleUur) + TimeValue("00:00:10")
    
    Application.OnTime TijdVolgendeCheck, "Basis_Functies.CheckAlarms"
End Sub





Sub CreateFunctionsDiscriptions2()
    On Error Resume Next
    
    
    ThisWorkbook.Worksheets("JTools functielijst").Sort.SortFields.Clear
    ThisWorkbook.Worksheets("JTools functielijst").Sort.SortFields.Add2 Key:= _
        Range("C2:C97"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    ThisWorkbook.Worksheets("JTools functielijst").Sort.SortFields.Add2 Key:= _
        Range("A2:A97"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ThisWorkbook.Worksheets("JTools functielijst").Sort
        .SetRange Range("A1:S97")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    
    
    
    
    
    
    
    
    
    
    
    Dim argdesc() As Variant
    
    Dim helpFile As String
    helpFile = "https://websensystems.nl/JTools/Functions.html"
    catName = ""
    
    htmlFile = "<div class=container><div class=row><div class=col><div id=accordion><div><div><div>"
    
  '  If SheetsVBAFunctions.CheckIfSheetExists("JTools functielijst") Then
        rijnum = 2
        While ThisWorkbook.Sheets("JTools functielijst").Range("A" & rijnum).Value <> ""
            If catName <> ThisWorkbook.Sheets("JTools functielijst").Range("C" & rijnum).Value Then
                catName = ThisWorkbook.Sheets("JTools functielijst").Range("C" & rijnum).Value
                catNameStipts = Replace(catName, " ", "")
                htmlFile = htmlFile + "</div></div></div><div class=card><div class=card-header id=""heading" & catNameStipts & """><h5 class=mb-0><button class=""btn btn-link"" data-toggle=collapse data-target=#collapse" & catNameStipts & " aria-expanded=true aria-controls=collapse" & catNameStipts & ">" & catName & "</button></h5></div><div id=collapse" & catNameStipts & " class=collapse aria-labelledby=""heading" & catNameStipts & """ data-parent=#accordion><div class=card-body>"
            End If
            htmlFile = htmlFile + "<table class='table mb-5 table-bordered ' id='" & ThisWorkbook.Sheets("JTools functielijst").Range("A" & rijnum).Value & "'><tr><td colspan=2><h3>" & ThisWorkbook.Sheets("JTools functielijst").Range("A" & rijnum).Value & "<h3></td><tr>"
            htmlFile = htmlFile + "<tr><td colspan=2>" & ThisWorkbook.Sheets("JTools functielijst").Range("B" & rijnum).Value & "</td></tr>"
            htmlFile = htmlFile + "<tr><td colspan=2>=" & ThisWorkbook.Sheets("JTools functielijst").Range("A" & rijnum).Value & "(" & ThisWorkbook.Sheets("JTools functielijst").Range("D" & rijnum).Value & ")</td></tr>"
            If ThisWorkbook.Sheets("JTools functielijst").Range("D" & rijnum).Value <> "" Then
                ArgumentenLijst = Split(ThisWorkbook.Sheets("JTools functielijst").Range("D" & rijnum).Value, ",")
            
            
                ReDim argdesc(0 To UBound(ArgumentenLijst))
                argdesc(0) = ThisWorkbook.Sheets("JTools functielijst").Range("E" & rijnum).Value
                htmlFile = htmlFile + "<tr><td>" & ArgumentenLijst(0) & "</td><td>" & argdesc(0) & "</td></tr>"
                
                
                If UBound(ArgumentenLijst) > 0 Then
                    argdesc(1) = ThisWorkbook.Sheets("JTools functielijst").Range("F" & rijnum).Value
                    htmlFile = htmlFile + "<tr><td>" & ArgumentenLijst(1) & "</td><td>" & argdesc(1) & "</td></tr>"
                End If
                If UBound(ArgumentenLijst) > 1 Then
                    argdesc(2) = ThisWorkbook.Sheets("JTools functielijst").Range("G" & rijnum).Value
                    htmlFile = htmlFile + "<tr><td>" & ArgumentenLijst(2) & "</td><td>" & argdesc(2) & "</td></tr>"
                End If
                If UBound(ArgumentenLijst) > 2 Then
                    argdesc(3) = ThisWorkbook.Sheets("JTools functielijst").Range("H" & rijnum).Value
                    htmlFile = htmlFile + "<tr><td>" & ArgumentenLijst(3) & "</td><td>" & argdesc(3) & "</td></tr>"
                End If
                If UBound(ArgumentenLijst) > 3 Then
                    argdesc(4) = ThisWorkbook.Sheets("JTools functielijst").Range("I" & rijnum).Value
                    htmlFile = htmlFile + "<tr><td>" & ArgumentenLijst(4) & "</td><td>" & argdesc(4) & "</td></tr>"
                End If
                If UBound(ArgumentenLijst) > 4 Then
                    argdesc(5) = ThisWorkbook.Sheets("JTools functielijst").Range("J" & rijnum).Value
                    htmlFile = htmlFile + "<tr><td>" & ArgumentenLijst(5) & "</td><td>" & argdesc(5) & "</td></tr>"
                End If
                If UBound(ArgumentenLijst) > 5 Then
                    argdesc(6) = ThisWorkbook.Sheets("JTools functielijst").Range("K" & rijnum).Value
                    htmlFile = htmlFile + "<tr><td>" & ArgumentenLijst(6) & "</td><td>" & argdesc(6) & "</td></tr>"
                End If
                If UBound(ArgumentenLijst) > 6 Then
                    argdesc(7) = ThisWorkbook.Sheets("JTools functielijst").Range("L" & rijnum).Value
                    htmlFile = htmlFile + "<tr><td>" & ArgumentenLijst(7) & "</td><td>" & argdesc(7) & "</td></tr>"
                End If
                If UBound(ArgumentenLijst) > 7 Then
                    argdesc(8) = ThisWorkbook.Sheets("JTools functielijst").Range("M" & rijnum).Value
                    htmlFile = htmlFile + "<tr><td>" & ArgumentenLijst(8) & "</td><td>" & argdesc(8) & "</td></tr>"
                End If
                If UBound(ArgumentenLijst) > 8 Then
                    argdesc(9) = ThisWorkbook.Sheets("JTools functielijst").Range("N" & rijnum).Value
                    htmlFile = htmlFile + "<tr><td>" & ArgumentenLijst(9) & "</td><td>" & argdesc(9) & "</td></tr>"
                End If
                If UBound(ArgumentenLijst) > 9 Then
                    argdesc(10) = ThisWorkbook.Sheets("JTools functielijst").Range("O" & rijnum).Value
                    htmlFile = htmlFile + "<tr><td>" & ArgumentenLijst(10) & "</td><td>" & argdesc(10) & "</td></tr>"
                End If
                If UBound(ArgumentenLijst) > 10 Then
                    argdesc(11) = ThisWorkbook.Sheets("JTools functielijst").Range("P" & rijnum).Value
                    htmlFile = htmlFile + "<tr><td>" & ArgumentenLijst(11) & "</td><td>" & argdesc(11) & "</td></tr>"
                End If
                If UBound(ArgumentenLijst) > 11 Then
                    argdesc(12) = ThisWorkbook.Sheets("JTools functielijst").Range("Q" & rijnum).Value
                    htmlFile = htmlFile + "<tr><td>" & ArgumentenLijst(12) & "</td><td>" & argdesc(12) & "</td></tr>"
                End If
                If UBound(ArgumentenLijst) > 12 Then
                    argdesc(13) = ThisWorkbook.Sheets("JTools functielijst").Range("R" & rijnum).Value
                    htmlFile = htmlFile + "<tr><td>" & ArgumentenLijst(13) & "</td><td>" & argdesc(13) & "</td></tr>"
                End If
                If UBound(ArgumentenLijst) > 13 Then
                    argdesc(14) = ThisWorkbook.Sheets("JTools functielijst").Range("S" & rijnum).Value
                    htmlFile = htmlFile + "<tr><td>" & ArgumentenLijst(14) & "</td><td>" & argdesc(14) & "</td></tr>"
                End If
                
                
                
            End If
            HelpURL = helpFile ' & "#" & thisworkbook.Sheets("JTools functielijst").Range("A" & rijnum).Value
            Application.MacroOptions ThisWorkbook.Sheets("JTools functielijst").Range("A" & rijnum).Value, ThisWorkbook.Sheets("JTools functielijst").Range("B" & rijnum).Value, Category:="JTools - " & ThisWorkbook.Sheets("JTools functielijst").Range("C" & rijnum).Value, ArgumentDescriptions:=argdesc, StatusBar:=ThisWorkbook.Sheets("JTools functielijst").Range("C" & rijnum).Value, helpFile:=HelpURL
            ReDim argdesc(0 To 1)
            htmlFile = htmlFile + "</table>"
        rijnum = rijnum + 1
        Wend
        Open "U:\Functions.html" For Output As 1
        Print #1, "<!DOCTYPE html><html><head><link rel='stylesheet' href='https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/css/bootstrap.min.css' integrity='sha384-BVYiiSIFeK1dGmJRAkycuHAHRg32OmUcww7on3RYdg4Va+PmSTsz/K68vbdEjh4u' crossorigin='anonymous'><!-- Optional theme --><link rel='stylesheet' href='https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/css/bootstrap-theme.min.css' integrity='sha384-rHyoN1iRsVXV4nD0JutlnGaslCJuC7uwjduW9SVrLvRYooPp2bWYgmgJQIXwl/Sp' crossorigin='anonymous'><script src=https://code.jquery.com/jquery-3.2.1.slim.min.js integrity=sha384-KJ3o2DKtIkvYIK3UENzmM7KCkRr/rE9/Qpg6aAZGJwFDMVNA/GpGFF93hXpG5KkN crossorigin=anonymous></script><script src=https://cdnjs.cloudflare.com/ajax/libs/popper.js/1.12.9/umd/popper.min.js integrity=sha384-ApNbgh9B+Y1QKtv3Rn7W3mgPxhU9K/ScQsAP7hUibX39j7fakFPskvXusvfa0b4Q crossorigin=anonymous></script>"
        Print #1, "<script src=https://maxcdn.bootstrapcdn.com/bootstrap/4.0.0/js/bootstrap.min.js integrity=sha384-JZR6Spejh4U02d8jOt6vLEHfe/JQGiRRSQQxSfFWpi1MquVdAyjUar5+76PVCmYl crossorigin=anonymous></script></head><body>" & htmlFile & "</div></div></div></div></body></html>"

        Close #1
  '  End If
    
End Sub


Sub CreateFunctionsDiscriptions()
    On Error Resume Next
    Dim catagories(1 To 9) As Variant
    
    catagories(1) = "JTools - Rooster"
    catagories(2) = "JTools - Database"
    catagories(3) = "JTools - Omrekenen"
    catagories(4) = "JTools - Feestdagen"
    catagories(5) = "JTools - Tijd&Datum functies"
    catagories(6) = "JTools - KKS functies"
    catagories(7) = "JTools - Weer functies"
    catagories(8) = "JTools - Tennet functies"
    'catagories(6) = "JTools - Baily blokken"
    
    Dim argdesc() As Variant
    
    Dim helpFile As String
    helpFile = "https://websensystems.nl"
    
    
    '================================
    ' Rooster
    '================================
    ReDim argdesc(0 To 1)
    argdesc(0) = "Ploeg. Mogelijkheden zijn A,B,C,D,E,F."
    argdesc(1) = "Optioneel: Datum. Mag een cell verwijzing zijn. Bij geen opgave wordt de huidige dag gebruikt."
    Application.MacroOptions "Dienst_PD", "Functie voor het weer geven van een dienst behorende bij een ploeg en datum.", Category:=catagories(1), ArgumentDescriptions:=argdesc, StatusBar:=catagories(1), helpFile:=helpFile & "#Dienst_PD"


    ReDim argdesc(0 To 1)
    argdesc(0) = "Shift. Een dienst. Mag Enkele letter zijn maar ook uitgescheven voorbeelden: V,v,vroege. "
    argdesc(1) = "Optioneel: Datum. Mag een cell verwijzing zijn. Bij geen opgave wordt de huidige dag gebruikt."
    Application.MacroOptions "Dienst_SD", "Functie voor het weer geven van een ploeg behorende bij een dienst en datum.", Category:=catagories(1), ArgumentDescriptions:=argdesc, StatusBar:=catagories(1), helpFile:=helpFile & "#Dienst_SD"
    
    '================================
    ' Database
    '================================
    
    ReDim argdesc(0 To 6)
    argdesc(0) = "KKS nummer (bv 10hbk01ct044q01 of _1hbk01ct044q01)"
    argdesc(1) = "Optioneel: Bewerking. Kan zijn: SUM (totaal), AVG (Gemiddelde), MAX (Grootste), MIN (Kleinste)"
    argdesc(2) = "Start tijd en datum. Kan eventueel verwijzing zijn naar een cel met =nu() of =vandaag()"
    argdesc(3) = "Stop tijd en datum. Kan eventueel verwijzing zijn naar een cel met =nu() of =vandaag()"
    argdesc(4) = "Het rij nummer van het gegeven. 2 is bijvoorbeeld het 3de gegeven. Standaart is 0"
    argdesc(5) = "Een tabel. Bijvoorbeeld HOUR/DAY/WEEK. Bij geen opgegegeven bewerking wordt avg (gemiddelde) aangehouden. Bij leeg laten wordt de tabel RAW (minuut) gebruikt"
    argdesc(6) = "Comment toevoegen met gegevens van de query True/False Default=true"
    Application.MacroOptions "Get_DB_value", "Functie voor het weergeven van een enkele gegeven uit de database", Category:=catagories(2), ArgumentDescriptions:=argdesc, StatusBar:=catagories(2)
    
    ReDim argdesc(0 To 2)
    argdesc(0) = "KKS nummer. "
    argdesc(1) = "Optioneel: Gegevens bewerking (AVG/SUM/MIN/MAX)."
    argdesc(2) = "Optioneel: Tabel (RAW/HOUR/DAY)."
    Application.MacroOptions "get_DB_table", "Functie voor het weergeven van de mogelijke tabel uit de database", Category:=catagories(2), ArgumentDescriptions:=argdesc, StatusBar:=catagories(2)
    
    ReDim argdesc(0 To 2)
    argdesc(0) = "KKS nummer. "
    argdesc(1) = "Optioneel: Gegevens bewerking (AVG/SUM/MIN/MAX)."
    argdesc(2) = "Optioneel: Tabel (RAW/HOUR/DAY)."
    Application.MacroOptions "get_DB_kks", "Functie voor het weergeven van de juite database kks", Category:=catagories(2), ArgumentDescriptions:=argdesc, StatusBar:=catagories(2)
    
    
    
    '================================
    ' Omrekenen
    '================================
    
    ReDim argdesc(0 To 2)
    argdesc(0) = "Rekenwaarde. "
    argdesc(1) = "Eenheid van het eindproduct."
    argdesc(2) = "Optioneel: Eenheid van het beginproduct."
    Application.MacroOptions "OmrekenenEnergie", "Functie voor het omrekenen van energie.", Category:=catagories(3), ArgumentDescriptions:=argdesc, StatusBar:=catagories(3)
    
    ReDim argdesc(0 To 2)
    argdesc(0) = "Rekenwaarde. "
    argdesc(1) = "Eenheid van het eindproduct."
    argdesc(2) = "Optioneel: Eenheid van het beginproduct."
    Application.MacroOptions "OmrekenenDruk", "Functie voor het omrekenen van drukken.", Category:=catagories(3), ArgumentDescriptions:=argdesc, StatusBar:=catagories(3)
  
    ReDim argdesc(0 To 2)
    argdesc(0) = "Rekenwaarde. "
    argdesc(1) = "Eenheid van het eindproduct."
    argdesc(2) = "Optioneel: Eenheid van het beginproduct."
    Application.MacroOptions "OmrekenenGewicht", "Functie voor het omrekenen van gewichten.", Category:=catagories(3), ArgumentDescriptions:=argdesc, StatusBar:=catagories(3)
  
    ReDim argdesc(0 To 2)
    argdesc(0) = "Rekenwaarde. "
    argdesc(1) = "Eenheid van het eindproduct."
    argdesc(2) = "Optioneel: Eenheid van het beginproduct."
    Application.MacroOptions "OmrekenenInhoud", "Functie voor het omrekenen van inhouden.", Category:=catagories(3), ArgumentDescriptions:=argdesc, StatusBar:=catagories(3)
  
    ReDim argdesc(0 To 2)
    argdesc(0) = "Rekenwaarde. "
    argdesc(1) = "Eenheid van het eindproduct."
    argdesc(2) = "Optioneel: Eenheid van het beginproduct."
    Application.MacroOptions "OmrekenenLengte", "Functie voor het omrekenen van lengtes.", Category:=catagories(3), ArgumentDescriptions:=argdesc, StatusBar:=catagories(3)
  
    ReDim argdesc(0 To 2)
    argdesc(0) = "Rekenwaarde. "
    argdesc(1) = "Eenheid van het eindproduct."
    argdesc(2) = "Optioneel: Eenheid van het beginproduct."
    Application.MacroOptions "OmrekenenSnelheid", "Functie voor het omrekenen van snelheid.", Category:=catagories(3), ArgumentDescriptions:=argdesc, StatusBar:=catagories(3)
    
    ReDim argdesc(0 To 2)
    argdesc(0) = "Flow in m3."
    argdesc(1) = "Druk in Bar"
    argdesc(2) = "Temperatuur in C"
    Application.MacroOptions "OmrekenennaarNm3", "Functie voor het omrekenen van flow van m3 naar Nm3.", Category:=catagories(3), ArgumentDescriptions:=argdesc, StatusBar:=catagories(3)
    
    ReDim argdesc(0 To 2)
    argdesc(0) = "Flow in Nm3."
    argdesc(1) = "Druk in Bar"
    argdesc(2) = "Temperatuur in C"
    Application.MacroOptions "Omrekenennaarm3", "Functie voor het omrekenen van flow van Nm3 naar m3.", Category:=catagories(3), ArgumentDescriptions:=argdesc, StatusBar:=catagories(3)
    
    '================================
    ' Feestdagen
    '================================
    ReDim argdesc(0 To 1)
    argdesc(0) = "Optioneel: Van welk jaar"
    argdesc(1) = "Optioneel: 1 of 2"
    Application.MacroOptions "Pasen", "Functie voor het weergeven van de datum van Pasen.", Category:=catagories(4), ArgumentDescriptions:=argdesc, StatusBar:=catagories(4)
    
    ReDim argdesc(0 To 1)
    argdesc(0) = "Optioneel: Van welk jaar"
    argdesc(1) = "Optioneel: 1 t/m 4"
    Application.MacroOptions "Carnaval", "Functie voor het weergeven van de datum van Carnaval.", Category:=catagories(4), ArgumentDescriptions:=argdesc, StatusBar:=catagories(4)
    
    ReDim argdesc(0 To 0)
    argdesc(0) = "Optioneel: Van welk jaar"
    Application.MacroOptions "GoedeVrijdag", "Functie voor het weergeven van de datum van Goede Vrijdag.", Category:=catagories(4), ArgumentDescriptions:=argdesc, StatusBar:=catagories(4)
    
    ReDim argdesc(0 To 0)
    argdesc(0) = "Optioneel: Van welk jaar"
    Application.MacroOptions "Hemelvaart", "Functie voor het weergeven van de datum van Hemelvaart.", Category:=catagories(4), ArgumentDescriptions:=argdesc, StatusBar:=catagories(4)
    
    ReDim argdesc(0 To 1)
    argdesc(0) = "Optioneel: Van welk jaar"
    argdesc(1) = "Optioneel: 1 of 2"
    Application.MacroOptions "Pinksteren", "Functie voor het weergeven van de datum van Pinksteren.", Category:=catagories(4), ArgumentDescriptions:=argdesc, StatusBar:=catagories(4)
    
    ReDim argdesc(0 To 1)
    argdesc(0) = "Optioneel: Van welk jaar"
    argdesc(1) = "Optioneel: 1 t/m 4"
    Application.MacroOptions "Vierdaagse", "Functie voor het weergeven van de datum van de Vierdaagse.", Category:=catagories(4), ArgumentDescriptions:=argdesc, StatusBar:=catagories(4)
    
    '================================
    ' Datum en Tijd
    '================================

    ReDim argdesc(0 To 0)
    argdesc(0) = "Geboortedatum tussen aanhalingstekens "
    Application.MacroOptions "Leeftijd", "Functie voor het berekenen van een leeftijd.", Category:=catagories(5), ArgumentDescriptions:=argdesc, StatusBar:=catagories(5)

    ReDim argdesc(0 To 0)
    argdesc(0) = "Datum tussen aanhalingstekens"
    Application.MacroOptions "dagenTotDatum", "Functie voor het berekenen van het aantal dagen tot een bepaalde datum.", Category:=catagories(5), ArgumentDescriptions:=argdesc, StatusBar:=catagories(5)
    
    ReDim argdesc(0 To 1)
    argdesc(0) = "Optioneel: Datum en tijd (Default is nu)"
    argdesc(1) = "Optioneel: Format (bv dd-mm-yyyy)"
    Application.MacroOptions "VolgendeHeleUur", "Functie voor het berekenen van het aankomend hele uur.", Category:=catagories(5), ArgumentDescriptions:=argdesc, StatusBar:=catagories(5)
    
    ReDim argdesc(0 To 2)
    argdesc(0) = "Optioneel: Datum en tijd (Default is nu)"
    argdesc(1) = "Optioneel: Format (bv dd-mm-yyyy)"
    argdesc(2) = "Optioneel: Verschuiving in de tijd"
    Application.MacroOptions "LaatsteHeleUur", "Functie voor het berekenen van het laatste hele uur.", Category:=catagories(5), ArgumentDescriptions:=argdesc, StatusBar:=catagories(5)
   
    ReDim argdesc(0 To 1)
    argdesc(0) = "Optioneel: Datum en tijd (Default is nu)"
    argdesc(1) = "Optioneel: Format (bv dd-mm-yyyy)"
    Application.MacroOptions "TijdEindeWacht", "Functie voor het berekenen van het einde van de wacht.", Category:=catagories(5), ArgumentDescriptions:=argdesc, StatusBar:=catagories(5)
   
    ReDim argdesc(0 To 1)
    argdesc(0) = "Optioneel: Datum en tijd (Default is nu)"
    argdesc(1) = "Optioneel: Format (bv dd-mm-yyyy)"
    Application.MacroOptions "TijdStartWacht", "Functie voor het berekenen van het begin van de wacht.", Category:=catagories(5), ArgumentDescriptions:=argdesc, StatusBar:=catagories(5)
   
    ReDim argdesc(0 To 1)
    argdesc(0) = "Optioneel: Datum en tijd (Default is nu)"
    argdesc(1) = "Optioneel: Format (bv dd-mm-yyyy)"
    Application.MacroOptions "StartWeek", "Functie voor het berekenen van het begin van de week.", Category:=catagories(5), ArgumentDescriptions:=argdesc, StatusBar:=catagories(5)
    
    ReDim argdesc(0 To 1)
    argdesc(0) = "Optioneel: Datum en tijd (Default is nu)"
    argdesc(1) = "Optioneel: Format (bv dd-mm-yyyy)"
    Application.MacroOptions "EindWeek", "Functie voor het berekenen van het einde van de week.", Category:=catagories(5), ArgumentDescriptions:=argdesc, StatusBar:=catagories(5)
    
    ReDim argdesc(0 To 1)
    argdesc(0) = "Optioneel: Datum en tijd (Default is nu)"
    argdesc(1) = "Optioneel: Format (bv dd-mm-yyyy)"
    Application.MacroOptions "StartCyclus", "Functie voor het berekenen van het begin van de cycles.", Category:=catagories(5), ArgumentDescriptions:=argdesc, StatusBar:=catagories(5)
    
    ReDim argdesc(0 To 1)
    argdesc(0) = "Optioneel: Datum en tijd (Default is nu)"
    argdesc(1) = "Optioneel: Format (bv dd-mm-yyyy)"
    Application.MacroOptions "EindeCyclus", "Functie voor het berekenen van het einde van de cyclus.", Category:=catagories(5), ArgumentDescriptions:=argdesc, StatusBar:=catagories(5)
    
    '================================
    ' KKS functies
    '================================
    
    ReDim argdesc(0 To 1)
    argdesc(0) = "Kks nummer"
    argdesc(1) = "Optioneel: Lange over korte versie (waar/onwaar)"
    Application.MacroOptions "getDBKKS", "Haalt de bekende db kks format op.", Category:=catagories(6), ArgumentDescriptions:=argdesc, StatusBar:=catagories(6)
    
    ReDim argdesc(0 To 0)
    argdesc(0) = "Kks nummer"
    Application.MacroOptions "getUltimoDiscr", "Geeft de ultimo omschrijving weer.", Category:=catagories(6), ArgumentDescriptions:=argdesc, StatusBar:=catagories(6)
    
    ReDim argdesc(0 To 0)
    argdesc(0) = "Kks nummer"
    Application.MacroOptions "getUltimoKostenplaatsNum", "Geeft het ultimo kostenplaatsnummer weer.", Category:=catagories(6), ArgumentDescriptions:=argdesc, StatusBar:=catagories(6)
    
    ReDim argdesc(0 To 0)
    argdesc(0) = "Kks nummer"
    Application.MacroOptions "getUltimoPIDNum", "Geeft het ultimo pid weer.", Category:=catagories(6), ArgumentDescriptions:=argdesc, StatusBar:=catagories(6)
    
    ReDim argdesc(0 To 0)
    argdesc(0) = "Kks nummer"
    Application.MacroOptions "getUltimoElectrischeVerdelerNum", "Geeft het in ultimo bekende verdeler weer.", Category:=catagories(6), ArgumentDescriptions:=argdesc, StatusBar:=catagories(6)
    
    ReDim argdesc(0 To 0)
    argdesc(0) = "Kks nummer"
    Application.MacroOptions "getUltimoWSinHoofdstroom", "Geeft aan of de werkschakelaar volgens ultimo in de hoofdstroom zit.", Category:=catagories(6), ArgumentDescriptions:=argdesc, StatusBar:=catagories(6)
    
    ReDim argdesc(0 To 0)
    argdesc(0) = "Kks nummer"
    Application.MacroOptions "getUltimoZone", "Geeft het in ultimo bekende flitslicht zone weer.", Category:=catagories(6), ArgumentDescriptions:=argdesc, StatusBar:=catagories(6)
    
    ReDim argdesc(0 To 0)
    argdesc(0) = "Kostenplaatscode"
    Application.MacroOptions "getUltimoKostenplaats", "Geeft de kostenplaats omschrijving weer.", Category:=catagories(6), ArgumentDescriptions:=argdesc, StatusBar:=catagories(6)
    
    ReDim argdesc(0 To 0)
    argdesc(0) = "PID nummer"
    Application.MacroOptions "getUltimoPIDDiscr", "Geeft P&ID omschrijving weer.", Category:=catagories(6), ArgumentDescriptions:=argdesc, StatusBar:=catagories(6)
    
    ReDim argdesc(0 To 0)
    argdesc(0) = "PID nummer"
    Application.MacroOptions "getUltimoPIDVersion", "Geeft P&ID versie weer.", Category:=catagories(6), ArgumentDescriptions:=argdesc, StatusBar:=catagories(6)
    
    ReDim argdesc(0 To 0)
    argdesc(0) = "PID nummer"
    Application.MacroOptions "getUltimoPIDLastChange", "Geeft P&ID versie datum weer.", Category:=catagories(6), ArgumentDescriptions:=argdesc, StatusBar:=catagories(6)
    
    ReDim argdesc(0 To 0)
    argdesc(0) = "PID nummer"
    Application.MacroOptions "getUltimoPIDResponsible", "Geeft P&ID verantwoordelijke weer.", Category:=catagories(6), ArgumentDescriptions:=argdesc, StatusBar:=catagories(6)
    
    ReDim argdesc(0 To 0)
    argdesc(0) = "Electrische verdeler"
    Application.MacroOptions "getUltimoElectrischeVerdeler", "Geeft de omschrijving van de electrische verdeler weer.", Category:=catagories(6), ArgumentDescriptions:=argdesc, StatusBar:=catagories(6)
    
    ReDim argdesc(0 To 0)
    argdesc(0) = "Flitslicht Zone"
    Application.MacroOptions "getUltimoZoneDiscr", "Geeft de flitslicht omschrijving weer.", Category:=catagories(6), ArgumentDescriptions:=argdesc, StatusBar:=catagories(6)
    
    ReDim argdesc(0 To 0)
    argdesc(0) = "Flitslicht Zone"
    Application.MacroOptions "getUltimoZoneLevel", "Geeft de flitslicht zone hoogte/verdieping weer.", Category:=catagories(6), ArgumentDescriptions:=argdesc, StatusBar:=catagories(6)
    
    ReDim argdesc(0 To 0)
    argdesc(0) = "Kks nummer"
    Application.MacroOptions "getUltimoFlitslicht", "Geeft het in ultimo bekende flitslicht nummer weer.", Category:=catagories(6), ArgumentDescriptions:=argdesc, StatusBar:=catagories(6)
    
    ReDim argdesc(0 To 0)
    argdesc(0) = "Flitslicht Zone"
    Application.MacroOptions "getUltimoFlitslichtDiscr", "Geeft de flitslicht omschrijving weer.", Category:=catagories(6), ArgumentDescriptions:=argdesc, StatusBar:=catagories(6)

    '================================
    ' Weer functies
    '================================
    
    ReDim argdesc(0 To 1)
    argdesc(0) = "Optioneel: Stationsnummer, Default = Arnhem"
    Application.MacroOptions "getWeerData_StationNaam", "Haalt de naam van het station op.", Category:=catagories(7), ArgumentDescriptions:=argdesc, StatusBar:=catagories(7)
    ReDim argdesc(0 To 1)
    argdesc(0) = "Optioneel: Stationsnummer, Default = Arnhem"
    Application.MacroOptions "getWeerData_Moment", "Haalt de datum en tijd van de huidige set op.", Category:=catagories(7), ArgumentDescriptions:=argdesc, StatusBar:=catagories(7)
    ReDim argdesc(0 To 1)
    argdesc(0) = "Optioneel: Stationsnummer, Default = Arnhem"
    Application.MacroOptions "getWeerData_Temperatuur", "Haalt de temperatuur van de huidige set op.", Category:=catagories(7), ArgumentDescriptions:=argdesc, StatusBar:=catagories(7)
    ReDim argdesc(0 To 1)
    argdesc(0) = "Optioneel: Stationsnummer, Default = Arnhem"
    Application.MacroOptions "getWeerData_Vochtigheid", "Haalt de vochtigheid van de huidige set op.", Category:=catagories(7), ArgumentDescriptions:=argdesc, StatusBar:=catagories(7)
    ReDim argdesc(0 To 1)
    argdesc(0) = "Optioneel: Stationsnummer, Default = Arnhem"
    Application.MacroOptions "getWeerData_Windsnelheid", "Haalt de Windsnelheid in m/s van de huidige set op.", Category:=catagories(7), ArgumentDescriptions:=argdesc, StatusBar:=catagories(7)
    ReDim argdesc(0 To 1)
    argdesc(0) = "Optioneel: Stationsnummer, Default = Arnhem"
    Application.MacroOptions "getWeerData_WindrichtingGR", "Haalt de windrichting in graden van de huidige set op.", Category:=catagories(7), ArgumentDescriptions:=argdesc, StatusBar:=catagories(7)
    ReDim argdesc(0 To 1)
    argdesc(0) = "Optioneel: Stationsnummer, Default = Arnhem"
    Application.MacroOptions "getWeerData_Windrichting", "Haalt de windrichting in kompassrichtingen van de huidige set op.", Category:=catagories(7), ArgumentDescriptions:=argdesc, StatusBar:=catagories(7)
    ReDim argdesc(0 To 1)
    argdesc(0) = "Optioneel: Stationsnummer, Default = Arnhem"
    Application.MacroOptions "getWeerData_Luchtdruk", "Haalt de luchtdruk van de huidige set op.", Category:=catagories(7), ArgumentDescriptions:=argdesc, StatusBar:=catagories(7)
    ReDim argdesc(0 To 1)
    argdesc(0) = "Optioneel: Stationsnummer, Default = Arnhem"
    Application.MacroOptions "getWeerData_Windstoten", "Haalt de snelheid van windstoten in m/s van de huidige set op.", Category:=catagories(7), ArgumentDescriptions:=argdesc, StatusBar:=catagories(7)
    ReDim argdesc(0 To 1)
    argdesc(0) = "Optioneel: Stationsnummer, Default = Arnhem"
    Application.MacroOptions "getWeerData_Regen", "Haalt de de hoeveelheid regen in mm/h van de huidige set op.", Category:=catagories(7), ArgumentDescriptions:=argdesc, StatusBar:=catagories(7)
    ReDim argdesc(0 To 1)
    argdesc(0) = "Optioneel: Stationsnummer, Default = Arnhem"
    Application.MacroOptions "getWeerData_Zicht", "Haalt de zichtafstand in meters van de huidige set op.", Category:=catagories(7), ArgumentDescriptions:=argdesc, StatusBar:=catagories(7)
    ReDim argdesc(0 To 1)
    argdesc(0) = "Optioneel: Stationsnummer, Default = Arnhem"
    Application.MacroOptions "getWeerData_ZonIntensiteit", "Haalt de zonintensiteit in W/m2 van de huidige set op.", Category:=catagories(7), ArgumentDescriptions:=argdesc, StatusBar:=catagories(7)
    ReDim argdesc(0 To 1)
    argdesc(0) = "Optioneel: Stationsnummer, Default = Arnhem"
    Application.MacroOptions "getWeerData_TemperatuurOp10cm", "Haalt de temperatuur op 10cm vanaf maaiveld van de huidige set op.", Category:=catagories(7), ArgumentDescriptions:=argdesc, StatusBar:=catagories(7)
    
    '================================
    ' Tennet functies
    '================================
    
    ReDim argdesc(0 To 1)
    argdesc(0) = "Optioneel: ItemNum, aantal blokken eerder"
    Application.MacroOptions "getTennetData_Tijd", "Haalt de tijd van de huidige gegevensset op.", Category:=catagories(8), ArgumentDescriptions:=argdesc, StatusBar:=catagories(8)
    
    ReDim argdesc(0 To 1)
    argdesc(0) = "Optioneel: ItemNum, aantal blokken eerder"
    Application.MacroOptions "getTennetData_OpregelVermogen", "Haalt het opgeregelvermogen van de huidige gegevensset op.", Category:=catagories(8), ArgumentDescriptions:=argdesc, StatusBar:=catagories(8)
    
    ReDim argdesc(0 To 1)
    argdesc(0) = "Optioneel: ItemNum, aantal blokken eerder"
    Application.MacroOptions "getTennetData_AfregelVermogen", "Haalt het afgeregelvermogen van de huidige gegevensset op.", Category:=catagories(8), ArgumentDescriptions:=argdesc, StatusBar:=catagories(8)
    
    ReDim argdesc(0 To 1)
    argdesc(0) = "Optioneel: ItemNum, aantal blokken eerder"
    Application.MacroOptions "getTennetData_OpregelVermogenReserve", "Haalt het reserve opgeregelvermogen van de huidige gegevensset op.", Category:=catagories(8), ArgumentDescriptions:=argdesc, StatusBar:=catagories(8)
    
    ReDim argdesc(0 To 1)
    argdesc(0) = "Optioneel: ItemNum, aantal blokken eerder"
    Application.MacroOptions "getTennetData_AfregelVermogenReserve", "Haalt het reserve afgeregelvermogen van de huidige gegevensset op.", Category:=catagories(8), ArgumentDescriptions:=argdesc, StatusBar:=catagories(8)
    
    ReDim argdesc(0 To 1)
    argdesc(0) = "Optioneel: ItemNum, aantal blokken eerder"
    Application.MacroOptions "getTennetData_Noodvermogen", "Haalt op of er noodvermogen is geactiveerd van de huidige gegevensset op.", Category:=catagories(8), ArgumentDescriptions:=argdesc, StatusBar:=catagories(8)
    
    ReDim argdesc(0 To 1)
    argdesc(0) = "Optioneel: ItemNum, aantal blokken eerder"
    Application.MacroOptions "getTennetData_PrijsMin", "Haalt de minimale prijs van de huidige gegevensset op.", Category:=catagories(8), ArgumentDescriptions:=argdesc, StatusBar:=catagories(8)
    
    ReDim argdesc(0 To 1)
    argdesc(0) = "Optioneel: ItemNum, aantal blokken eerder"
    Application.MacroOptions "getTennetData_PrijsMax", "Haalt de maximale prijs van de huidige gegevensset op.", Category:=catagories(8), ArgumentDescriptions:=argdesc, StatusBar:=catagories(8)
    
End Sub
