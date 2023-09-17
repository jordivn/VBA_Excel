Attribute VB_Name = "Celopmaak"
Sub CelOpmakenBedragen(opmaakVorm)
    Select Case opmaakVorm
        Case "E"
            Selection.NumberFormat = "$0.00"
        Case "E/m3"
            Selection.NumberFormat = "$0.00""/m³"""
        Case "E/l"
        Selection.NumberFormat = "$0.00""/l"""
        Case "E/kg"
        Selection.NumberFormat = "$0.00""/kg"""
        Case "E/h"
        Selection.NumberFormat = "$0.00""/h"""
        Case "E/ton"
        Selection.NumberFormat = "$0.00""/ton"""
    End Select
End Sub

Sub CelOpmakenLengteSnelheid(opmaakVorm)
    Select Case opmaakVorm
        Case "mm"
            Selection.NumberFormat = "0 ""mm"""
        Case "mm/s"
            Selection.NumberFormat = "0.0 ""mm/s"""
        Case "mm/min"
            Selection.NumberFormat = "0.0 ""mm/min"""
        Case "mm/h"
            Selection.NumberFormat = "0.0 ""mm/h"""
        
        Case "cm"
            Selection.NumberFormat = "0 ""cm"""
        Case "cm/s"
            Selection.NumberFormat = "0.0 ""cm/s"""
        Case "cm/min"
            Selection.NumberFormat = "0.0 ""cm/min"""
        Case "cm/h"
            Selection.NumberFormat = "0.0 ""cm/h"""
        
        Case "dm"
            Selection.NumberFormat = "0 ""dm"""
        Case "dm/s"
            Selection.NumberFormat = "0.0 ""dm/s"""
        Case "dm/min"
            Selection.NumberFormat = "0.0 ""dm/min"""
        Case "dm/h"
            Selection.NumberFormat = "0.0 ""dm/h"""
        
        Case "m"
            Selection.NumberFormat = "0 ""m"""
        Case "m/s"
            Selection.NumberFormat = "0.0 ""m/s"""
        Case "m/min"
            Selection.NumberFormat = "0.0 ""m/min"""
        Case "m/h"
            Selection.NumberFormat = "0.0 ""m/h"""
        
        Case "km"
            Selection.NumberFormat = "0 ""km"""
        Case "km/s"
            Selection.NumberFormat = "0.0 ""km/s"""
        Case "km/min"
            Selection.NumberFormat = "0.0 ""km/min"""
        Case "km/h"
            Selection.NumberFormat = "0.0 ""km/h"""
       
    End Select
   
End Sub

Sub CelOpmakenDruk(opmaakVorm)
    Select Case opmaakVorm
        Case "bar"
            Selection.NumberFormat = "0 ""bar"""
        Case "bara"
            Selection.NumberFormat = "0.0 ""bara"""
        Case "barg"
            Selection.NumberFormat = "0.0 ""bar(g)"""
        
        Case "mbar"
            Selection.NumberFormat = "0 ""mbar"""
        Case "mbara"
            Selection.NumberFormat = "0.0 ""mbara"""
        Case "mbarg"
            Selection.NumberFormat = "0.0 ""mbar(g)"""
       
        Case "N/mm2"
            Selection.NumberFormat = "0 ""N/mm²"""
        Case "N/cm2"
            Selection.NumberFormat = "0.0 ""N/cm²"""
        Case "N/m2"
            Selection.NumberFormat = "0.0 ""N/m²"""
        Case "KN/m2"
            Selection.NumberFormat = "0.0 ""KN/m²"""
        
        Case "Pa"
            Selection.NumberFormat = "0 ""Pa"""
        Case "KPa"
            Selection.NumberFormat = "0.0 ""KPa"""
        Case "MPa"
            Selection.NumberFormat = "0.0 ""MPa"""
        
        Case "atm"
            Selection.NumberFormat = "0 ""atm"""
        Case "PSI"
            Selection.NumberFormat = "0.0 ""PSI"""
        Case "m H2O"
            Selection.NumberFormat = "0.0 ""m H2O"""
        Case "m wk"
            Selection.NumberFormat = "0.0 ""m wk"""
        
    End Select
   
End Sub
            
Sub CelOpmakenMassa(opmaakVorm)
    Select Case opmaakVorm
        Case "mg"
            Selection.NumberFormat = "0 ""mg"""
        Case "mg/s"
            Selection.NumberFormat = "0.0 ""mg/s"""
        Case "mg/min"
            Selection.NumberFormat = "0.0 ""mg/min"""
        Case "mg/h"
            Selection.NumberFormat = "0.0 ""mg/h"""
        
        Case "g"
            Selection.NumberFormat = "0 ""g"""
        Case "g/s"
            Selection.NumberFormat = "0.0 ""g/s"""
        Case "g/min"
            Selection.NumberFormat = "0.0 ""g/min"""
        Case "g/h"
            Selection.NumberFormat = "0.0 ""g/h"""
        
        Case "Kg"
            Selection.NumberFormat = "0 ""Kg"""
        Case "Kg/s"
            Selection.NumberFormat = "0.0 ""Kg/s"""
        Case "Kg/min"
            Selection.NumberFormat = "0.0 ""Kg/min"""
        Case "Kg/h"
            Selection.NumberFormat = "0.0 ""Kg/h"""
        
        Case "T"
            Selection.NumberFormat = "0 ""T"""
        Case "T/s"
            Selection.NumberFormat = "0.0 ""T/s"""
        Case "T/min"
            Selection.NumberFormat = "0.0 ""T/min"""
        Case "T/h"
            Selection.NumberFormat = "0.0 ""T/h"""
        
        Case "Ton"
            Selection.NumberFormat = "0 ""Ton"""
        Case "gram"
            Selection.NumberFormat = "0.0 ""gram"""
        
    End Select
   
End Sub
            
Sub CelOpmakenEnergie(opmaakVorm)
    Select Case opmaakVorm
        Case "W"
            Selection.NumberFormat = "0 ""W"""
        Case "Wth"
            Selection.NumberFormat = "0.0 ""Wth"""
        Case "W/s"
            Selection.NumberFormat = "0.0 ""W/s"""
        Case "W/h"
            Selection.NumberFormat = "0.0 ""W/h"""
            
       Case "KW"
            Selection.NumberFormat = "0 ""KW"""
        Case "KWth"
            Selection.NumberFormat = "0.0 ""KWth"""
        Case "KW/s"
            Selection.NumberFormat = "0.0 ""KW/s"""
        Case "KW/h"
            Selection.NumberFormat = "0.0 ""KW/h"""
            
        Case "MW"
            Selection.NumberFormat = "0 ""MW"""
        Case "MWth"
            Selection.NumberFormat = "0.0 ""MWth"""
        Case "MW/s"
            Selection.NumberFormat = "0.0 ""MW/s"""
        Case "MW/h"
            Selection.NumberFormat = "0.0 ""MW/h"""
                
        Case "J"
            Selection.NumberFormat = "0 ""J"""
        Case "KJ"
            Selection.NumberFormat = "0.0 ""KJ"""
        Case "MJ"
            Selection.NumberFormat = "0.0 ""MJ"""
        Case "GJ"
            Selection.NumberFormat = "0.0 ""GJ"""
                    
        Case "C"
            Selection.NumberFormat = "0.0 ""°C"""
        Case "F"
            Selection.NumberFormat = "0.0 ""°F"""
        Case "K"
            Selection.NumberFormat = "0.0 ""K"""
            
        Case "kcal"
            Selection.NumberFormat = "0.0 ""kcal"""
        Case "cal"
            Selection.NumberFormat = "0.0 ""cal"""
        Case "PK"
            Selection.NumberFormat = "0.0 ""PK"""
    End Select
   
End Sub

Sub CelOpmakenVolume(opmaakVorm)
    Select Case opmaakVorm
        Case "mm3"
            Selection.NumberFormat = "0.0 ""mm³"""
        Case "mm3/s"
            Selection.NumberFormat = "0.0 ""mm³/s"""
        Case "mm3/min"
            Selection.NumberFormat = "0.0 ""mm³/min"""
        Case "mm3/h"
            Selection.NumberFormat = "0.0 ""mm³/h"""
            
        Case "cm3"
            Selection.NumberFormat = "0.0 ""cm³"""
        Case "cm3/s"
            Selection.NumberFormat = "0.0 ""cm³/s"""
        Case "cm3/min"
            Selection.NumberFormat = "0.0 ""cm³/min"""
        Case "cm3/h"
            Selection.NumberFormat = "0.0 ""cm³/h"""
            
        Case "dm3"
            Selection.NumberFormat = "0.0 ""dm³"""
        Case "dm3/s"
            Selection.NumberFormat = "0.0 ""dm³/s"""
        Case "dm3/min"
            Selection.NumberFormat = "0.0 ""dm³/min"""
        Case "dm3/h"
            Selection.NumberFormat = "0.0 ""dm³/h"""
            
        Case "l"
            Selection.NumberFormat = "0.0 ""l"""
        Case "l/s"
            Selection.NumberFormat = "0.0 ""l/s"""
        Case "l/min"
            Selection.NumberFormat = "0.0 ""l/min"""
        Case "l/h"
            Selection.NumberFormat = "0.0 ""l/h"""
            
        Case "m3"
            Selection.NumberFormat = "0.0 ""m³"""
        Case "m3/s"
            Selection.NumberFormat = "0.0 ""m³/s"""
        Case "m3/min"
            Selection.NumberFormat = "0.0 ""m³/min"""
        Case "m3/h"
            Selection.NumberFormat = "0.0 ""m³/h"""
            
        Case "Nm3"
            Selection.NumberFormat = "0.0 ""Nm³"""
        Case "Nm3/s"
            Selection.NumberFormat = "0.0 ""Nm³/s"""
        Case "Nm3/min"
            Selection.NumberFormat = "0.0 ""Nm³/min"""
        Case "Nm3/h"
            Selection.NumberFormat = "0.0 ""Nm³/h"""
            
        Case "ml"
            Selection.NumberFormat = "0.0 ""ml"""
        Case "cc"
            Selection.NumberFormat = "0.0 ""cc"""
        Case "cl"
            Selection.NumberFormat = "0.0 ""cl"""
        Case "dl"
            Selection.NumberFormat = "0.0 ""dl"""
            
    End Select
End Sub

Sub CelOpmakenOverig(opmaakVorm)
    Select Case opmaakVorm
    Case "perc1"
            Selection.NumberFormat = "0.00 ""%"""
    Case "perc2"
            Selection.NumberFormat = "0.0 ""%"""
    Case "perc3"
            Selection.NumberFormat = "0 ""%"""
    Case "us/cm"
            Selection.NumberFormat = "0 ""µs/cm"""
            
    End Select
End Sub

Sub CelOpmakenDatumTijd(opmaakVorm)
    Select Case opmaakVorm
        Case "ddmmmjj"
            Selection.NumberFormat = "dd mmm yyyy"
        Case "ddmmmmjj"
            Selection.NumberFormat = "dd mmmm yyyy"
        Case "ddmmjj"
            Selection.NumberFormat = "dd-mm-yyyy"
        Case "d"
            Selection.NumberFormat = "dd"
        Case "m"
            Selection.NumberFormat = "mm"
        Case "j"
            Selection.NumberFormat = "yy"
        Case "ddd"
            Selection.NumberFormat = "ddd"
        Case "dddd"
            Selection.NumberFormat = "dddd"
        Case "ddmmmjjuummss"
            Selection.NumberFormat = "dd mmm yyyy hh:mm:ss"
        Case "ddmmmjjuumm"
            Selection.NumberFormat = "dd mmm yyyy hh:mm"
        Case "ddmmjjuummss"
            Selection.NumberFormat = "dd-mm-yyyy hh:mm:ss"
        Case "ddmmjjuumm"
            Selection.NumberFormat = "dd-mm-yyyy hh:mm"
        Case "uumm"
            Selection.NumberFormat = "hh:mm"
        Case "uummap"
            Selection.NumberFormat = "hh:mm AM/PM"
        Case "uummss"
            Selection.NumberFormat = "hh:mm:ss"
    End Select
End Sub
            
Sub CelOpmakenX(opmaakVorm)
    Select Case opmaakVorm
            
        Case "c"
            Selection.NumberFormat = "0.00 ""°C"""
        Case "bara"
            Selection.NumberFormat = "0.0 ""bar"""
        Case "barg"
            Selection.NumberFormat = "0.0 ""bar g"""
        Case "mbar"
            Selection.NumberFormat = "0 ""mbar"""
        Case "m3h"
            Selection.NumberFormat = "0 ""m³/h"""
        Case "Nm3h"
            Selection.NumberFormat = "0 ""Nm³/h"""
        Case "lsec"
            Selection.NumberFormat = "0.0 ""L/s"""
        Case "lmin"
            Selection.NumberFormat = "0.0 ""L/min"""
        Case "lh"
            Selection.NumberFormat = "0.0 ""L/h"""
        Case "MW"
            Selection.NumberFormat = "0.0 ""MW"""
        Case "KW"
            Selection.NumberFormat = "0.0 ""KW"""
        Case "datetime"
            Selection.NumberFormat = "dd-mm-yyyy hh:mm"
        Case "perc"
            Selection.NumberFormat = "0 ""%"""
        Case "T"
            Selection.NumberFormat = "0.0 ""Ton"""
        Case "kg"
            Selection.NumberFormat = "0.00 ""kg"""
        Case "m3"
            Selection.NumberFormat = "0 ""m³"""
        Case "m2" 'Nog knop maken
            Selection.NumberFormat = "0 ""m²"""
        Case "nmm2" 'Nog knop maken
            Selection.NumberFormat = "0 ""N/mm²"""
        Case "j" 'Nog knop maken
            Selection.NumberFormat = "0.00 ""J"""
        Case "w/s" 'Nog knop maken
            Selection.NumberFormat = "0.00 ""W/s"""
        Case "kj" 'Nog knop maken
            Selection.NumberFormat = "0.00 ""kJ"""
        Case "wh" 'Nog knop maken
            Selection.NumberFormat = "0.00 ""Wh"""
         Case "kwh" 'Nog knop maken
            Selection.NumberFormat = "0.00 ""kWh"""
           Case "mwh" 'Nog knop maken
            Selection.NumberFormat = "0.00 ""MWh"""
        Case "uscm" 'Nog knop maken
            Selection.NumberFormat = "0.00 ""µs/cm"""
       
    End Select
   '³²²
End Sub
