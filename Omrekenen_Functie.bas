Attribute VB_Name = "Omrekenen_Functie"
Function OmrekenenEnergie(waarde, eenheidEindproduct, Optional eenheidBeginproduct = "")
Attribute OmrekenenEnergie.VB_Description = "Berekend energie om naar een andere eenheid"
Attribute OmrekenenEnergie.VB_ProcData.VB_Invoke_Func = " \n23"
'alles naar J
If eenheidBeginproduct <> "" Then
        eenheidBeginproduct = LCase(eenheidBeginproduct)
        Select Case eenheidBeginproduct
            Case "kj"
                waarde = waarde * 1000
            Case "th"
                waarde = waarde * 4186800
            Case "mw"
                waarde = waarde * 1000 * 1000
            Case "kw"
                waarde = waarde * 1000
            Case "kwh"
                waarde = waarde * 3600000
            Case "wh"
                waarde = waarde * 3600
            Case "kcal"
                waarde = waarde * 4186.8
            Case "cal"
                waarde = waarde * 4.1868
            Case "hph"
                waarde = waarde * 2684519.5377
            Case "ws"
                waarde = waarde * 1
        End Select
End If
eenheidEindproduct = LCase(eenheidEindproduct)
Select Case eenheidEindproduct
            Case "j", "ws", "j/s"
                waarde = waarde
            Case "kj"
                waarde = waarde / 1000
            Case "th"
                waarde = waarde / 4186800
            Case "mw"
                waarde = waarde / 1000 / 1000
            Case "mwh"
                waarde = waarde / 1000 / 1000 / 3600
            Case "kw"
                waarde = waarde / 1000
            Case "kwh"
                waarde = waarde / 3600000
            Case "wh"
                waarde = waarde / 3600
            Case "kcal"
                waarde = waarde / 4186.8
            Case "cal"
                waarde = waarde / 4.1868
            Case "hph"
                waarde = waarde / 2684519.5377
            
End Select
OmrekenenEnergie = waarde
End Function



Function OmrekenenDruk(waarde, eenheidEindproduct, Optional eenheidBeginproduct = "")
Attribute OmrekenenDruk.VB_Description = "Berekend druk om naar een andere eenheid"
Attribute OmrekenenDruk.VB_ProcData.VB_Invoke_Func = " \n23"

'alles naar Pa
If eenheidBeginproduct <> "" Then
        eenheidBeginproduct = LCase(eenheidBeginproduct)
        Select Case eenheidBeginproduct
            Case "mpa"
                waarde = waarde * 1000000
            Case "kpa"
                waarde = waarde * 1000
            Case "nm/m2"
                waarde = waarde * 1
            Case "kn/m2"
                waarde = waarde * 1000
            Case "n/mm2"
                waarde = waarde * 1000000
            Case "n/cm2"
                waarde = waarde * 10000
            Case "bar"
                waarde = waarde * 100000
            Case "mbar"
                waarde = waarde * 100
            Case "psi"
                waarde = waarde * 6894.75729
            Case "atm"
                waarde = waarde * 101325
            Case "mh2o"
                waarde = waarde * 9806.65
            Case "cmh2o"
                waarde = waarde * 98.0665
            Case "mmh2o"
                waarde = waarde * 9.80665
        End Select
End If
eenheidEindproduct = LCase(eenheidEindproduct)
Select Case eenheidEindproduct
            Case "mpa"
                waarde = waarde / 1000000
            Case "kpa"
                waarde = waarde / 1000
            Case "nm/m2"
                waarde = waarde / 1
            Case "kn/m2"
                waarde = waarde / 1000
            Case "n/mm2"
                waarde = waarde / 1000000
            Case "n/cm2"
                waarde = waarde / 10000
            Case "bar"
                waarde = waarde / 100000
            Case "mbar"
                waarde = waarde / 100
            Case "psi"
                waarde = waarde / 6894.75729
            Case "atm"
                waarde = waarde / 101325
            Case "mh2o"
                waarde = waarde / 9806.65
            Case "cmh2o"
                waarde = waarde / 98.0665
            Case "mmh2o"
                waarde = waarde / 9.80665
            
End Select
OmrekenenDruk = waarde
End Function



Function OmrekenenGewicht(waarde, eenheidEindproduct, Optional eenheidBeginproduct = "")
Attribute OmrekenenGewicht.VB_Description = "Berekend gewicht  om naar een andere eenheid"
Attribute OmrekenenGewicht.VB_ProcData.VB_Invoke_Func = " \n23"

'alles naar g
If eenheidBeginproduct <> "" Then
        eenheidBeginproduct = LCase(eenheidBeginproduct)
        Select Case eenheidBeginproduct
            Case "mg"
                waarde = waarde * 0.001
            Case "cg"
                waarde = waarde * 0.01
            Case "kg"
                waarde = waarde * 1000
            Case "ug"
                waarde = waarde * 0.000001
            Case "ton"
                waarde = waarde * 1000000
                
            Case "karaat"
                waarde = waarde * 0.2
            Case "ounce"
                waarde = waarde * 28.34952
            Case "pound"
                waarde = waarde * 453.59237
            Case "stone"
                waarde = waarde * 6350.29318
            Case "quarter"
                waarde = waarde * 12700
            Case "dram"
                waarde = waarde * 1.77185
                
        
        End Select
End If
eenheidEindproduct = LCase(eenheidEindproduct)
Select Case eenheidEindproduct
            Case "mg"
                waarde = waarde / 0.001
            Case "cg"
                waarde = waarde / 0.01
            Case "kg"
                waarde = waarde / 1000
            Case "ug"
                waarde = waarde / 0.000001
            Case "ton"
                waarde = waarde / 1000000
                
            Case "karaat"
                waarde = waarde / 0.2
            Case "ounce"
                waarde = waarde / 28.34952
            Case "pound"
                waarde = waarde / 453.59237
            Case "stone"
                waarde = waarde / 6350.29318
            Case "quarter"
                waarde = waarde / 12700
            Case "dram"
                waarde = waarde / 1.77185
End Select
OmrekenenGewicht = waarde
End Function


Function OmrekenenInhoud(waarde, eenheidEindproduct, Optional eenheidBeginproduct = "")
Attribute OmrekenenInhoud.VB_Description = "Berekend inhoud om naar een andere eenheid"
Attribute OmrekenenInhoud.VB_ProcData.VB_Invoke_Func = " \n23"

'alles naar m3
If eenheidBeginproduct <> "" Then
        eenheidBeginproduct = LCase(eenheidBeginproduct)
        Select Case eenheidBeginproduct
            Case "mm3"
                waarde = waarde * 0.000000001
            Case "cm3"
                waarde = waarde * 0.000001
            Case "dm3"
                waarde = waarde * 0.001
            Case "cc"
                waarde = waarde * 0.000001
            Case "cl"
                waarde = waarde * 0.00001
            Case "dl"
                waarde = waarde * 0.0001
            Case "ml"
                waarde = waarde * 0.000001
            Case "l"
                waarde = waarde * 0.001
            Case "gill"
                waarde = waarde * 0.00014
            Case "pint"
                waarde = waarde * 0.00057
            Case "quart"
                waarde = waarde * 0.00114
            Case "gallon"
                waarde = waarde * 0.00455
        End Select
End If
    eenheidEindproduct = LCase(eenheidEindproduct)
        Select Case eenheidEindproduct
            Case "mm3"
                waarde = waarde / 0.000000001
            Case "cm3"
                waarde = waarde / 0.000001
            Case "dm3"
                waarde = waarde / 0.001
            Case "cc"
                waarde = waarde / 0.000001
            Case "cl"
                waarde = waarde / 0.00001
            Case "dl"
                waarde = waarde / 0.0001
            Case "ml"
                waarde = waarde / 0.000001
            Case "l"
                waarde = waarde / 0.001
            Case "gill"
                waarde = waarde / 0.00014
            Case "pint"
                waarde = waarde / 0.00057
            Case "quart"
                waarde = waarde / 0.00114
            Case "gallon"
                waarde = waarde / 0.00455
        End Select
    
    
    OmrekenenInhoud = waarde
End Function


Function OmrekenenLengte(waarde, eenheidEindproduct, Optional eenheidBeginproduct = "")
Attribute OmrekenenLengte.VB_Description = "Berekend lengte  om naar een andere eenheid"
Attribute OmrekenenLengte.VB_ProcData.VB_Invoke_Func = " \n23"
    'alles naar m
    If eenheidBeginproduct <> "" Then
         eenheidBeginproduct = LCase(eenheidBeginproduct)
        Select Case eenheidBeginproduct
            Case "dm"
                waarde = waarde * 0.1
            Case "cm"
                waarde = waarde * 0.01
            Case "mm"
                waarde = waarde * 0.001
            Case "dam"
                waarde = waarde * 10
            Case "hm"
                waarde = waarde * 100
            Case "km"
                waarde = waarde * 1000
            Case "el"
                waarde = waarde * 0.697
            Case "voet"
                waarde = waarde * 0.27386
            Case "mi"
                waarde = waarde * 1609.3
            Case "nmi"
                waarde = waarde * 1852
            Case "ft"
                waarde = waarde * 0.3048
            Case "in"
                waarde = waarde * 0.0254
            Case "yd"
                waarde = waarde * 0.9144
            Case "ly"
                waarde = waarde * (9.46073 * (10 ^ 15))

        End Select
    End If
    eenheidEindproduct = LCase(eenheidEindproduct)
     Select Case eenheidEindproduct
            Case "dm"
                waarde = waarde / 0.1
            Case "cm"
                waarde = waarde / 0.01
            Case "mm"
                waarde = waarde / 0.001
            Case "dm"
                waarde = waarde / 10
            Case "hm"
                waarde = waarde / 100
            Case "km"
                waarde = waarde / 1000
            Case "el"
                waarde = waarde / 0.697
            Case "ft"
                waarde = waarde / 0.27386
            Case "mi"
                waarde = waarde / 1609.3
            Case "nmi"
                waarde = waarde / 1852
            Case "ft"
                waarde = waarde / 0.3048
            Case "in"
                waarde = waarde / 0.0254
            Case "yd"
                waarde = waarde / 0.9144
            Case "ly"
                waarde = waarde / (9.46073 * (10 ^ 15))
                
        End Select
    
    
OmrekenenLengte = waarde
End Function

Function OmrekenenSnelheid(waarde, eenheidEindproduct, Optional eenheidBeginproduct = "")
Attribute OmrekenenSnelheid.VB_Description = "Berekend snelheid om naar een andere eenheid"
Attribute OmrekenenSnelheid.VB_ProcData.VB_Invoke_Func = " \n23"


    'alles naar m/s
    If eenheidBeginproduct <> "" Then
     eenheidBeginproduct = LCase(eenheidBeginproduct)
        Select Case eenheidBeginproduct
            Case "km/h"
                waarde = waarde * 0.27778
            Case "m/h"
                waarde = waarde * 0.00028
            Case "km/s"
                waarde = waarde * 1000
            Case "mph"
                waarde = waarde * 0.44704
            Case "mps"
                waarde = waarde * 1609.344
            Case "kn"
                waarde = waarde * 0.51444
            Case "licht"
                waarde = waarde * 299792458
            Case "geluid"
                waarde = waarde * 343.2
                
        End Select
    End If
    eenheidEindproduct = LCase(eenheidEindproduct)
     Select Case eenheidEindproduct
            Case "km/h"
                waarde = waarde / 0.27778
            Case "m/h"
                waarde = waarde / 0.00028
            Case "km/s"
                waarde = waarde / 1000
            Case "mph"
                waarde = waarde / 0.44704
            Case "mps"
                waarde = waarde / 1609.344
            Case "kn"
                waarde = waarde / 0.51444
            Case "licht"
                waarde = waarde / 299792458
            Case "geluid"
                waarde = waarde / 343.2
            
        End Select
    
    
OmrekenenSnelheid = waarde
End Function

Function OmrekenennaarNm3(Flow, Druk, Temperatuur)
Attribute OmrekenennaarNm3.VB_Description = "Berekend normaal kuub om naar kuub"
Attribute OmrekenennaarNm3.VB_ProcData.VB_Invoke_Func = " \n23"
'

OmrekenennaarNm3 = Flow * ((Druk * (10 ^ 3)) / 1013.25) * ((273) / (Temperatuur + 273))
End Function

Function Omrekenennaarm3(Flow, Druk, Temperatuur)
Attribute Omrekenennaarm3.VB_Description = "Berekend kuub om naar normaal kuub"
Attribute Omrekenennaarm3.VB_ProcData.VB_Invoke_Func = " \n23"
'

Omrekenennaarm3 = Flow * (1013.25 / (Druk * (10 ^ 3))) * ((Temperatuur + 273) / (273))
End Function

Function OmrekenenMassaStroom(waarde, eenheidEindproduct, Optional eenheidBeginproduct = "")
Attribute OmrekenenMassaStroom.VB_Description = "Berekend massastroom om naar een andere eenheid"
Attribute OmrekenenMassaStroom.VB_ProcData.VB_Invoke_Func = " \n23"
    'alles naar kg/s
    If eenheidBeginproduct <> "" Then
     eenheidBeginproduct = LCase(eenheidBeginproduct)
        Select Case eenheidBeginproduct
            Case "kg/s"
                waarde = waarde
            Case "t/s"
                waarde = waarde * 1000
            Case "kg/h"
                waarde = waarde / 3600
            Case "t/h"
                waarde = waarde / 3.6
        End Select
    End If
    eenheidEindproduct = LCase(eenheidEindproduct)
     Select Case eenheidEindproduct
           Case "kg/s"
                waarde = waarde
            Case "t/s"
                waarde = waarde / 1000
            Case "kg/h"
                waarde = waarde * 3600
            Case "t/h"
                waarde = waarde * 3.6
            
            
        End Select
    
    
OmrekenenMassaStroom = waarde
End Function

Function OmrekenenVolumeStroom(waarde, eenheidEindproduct, Optional eenheidBeginproduct = "")
Attribute OmrekenenVolumeStroom.VB_Description = "Berekend volumestroom om naar een andere eenheid"
Attribute OmrekenenVolumeStroom.VB_ProcData.VB_Invoke_Func = " \n23"
    'alles naar l/s
    If eenheidBeginproduct <> "" Then
     eenheidBeginproduct = LCase(eenheidBeginproduct)
        Select Case eenheidBeginproduct
            Case "l/s"
                waarde = waarde
            Case "m3/s"
                waarde = waarde * 1000
            Case "l/m"
                waarde = waarde / 60
            Case "m3/m"
                waarde = waarde * 1000 / 60
            Case "l/h"
                waarde = waarde / 3600
            Case "m3/h"
                waarde = waarde / 3.6
        End Select
    End If
    eenheidEindproduct = LCase(eenheidEindproduct)
     Select Case eenheidEindproduct
           Case "l/s"
                waarde = waarde
            Case "m3/s"
                waarde = waarde / 1000
             Case "l/m"
                waarde = waarde * 60
             Case "m3/m"
                waarde = waarde / 1000 * 60
            Case "l/h"
                waarde = waarde * 3600
            Case "m3/h"
                waarde = waarde * 3.6
            
            
        End Select
    
    
OmrekenenVolumeStroom = waarde
End Function

Function OmrekenenTemperatuur(waarde, eenheidEindproduct, Optional eenheidBeginproduct = "")
Attribute OmrekenenTemperatuur.VB_Description = "Berekend temperatuur om naar een andere eenheid"
Attribute OmrekenenTemperatuur.VB_ProcData.VB_Invoke_Func = " \n23"
    'alles naar k
    If eenheidBeginproduct <> "" Then
     eenheidBeginproduct = LCase(eenheidBeginproduct)
        Select Case eenheidBeginproduct
            Case "k"
                waarde = waarde
            Case "c"
                waarde = waarde + 273.15
            Case "f"
                waarde = (waarde - 32) / 1.8 + 273.15
                
            
        End Select
    End If
    eenheidEindproduct = LCase(eenheidEindproduct)
     Select Case eenheidEindproduct
            Case "k"
                waarde = waarde
            Case "c"
                waarde = waarde - 273.15

                
            Case "f"
                waarde = ((waarde - 273.15) * 1.8) + 32
            
            
        End Select
    

OmrekenenTemperatuur = waarde
End Function

Function OmrekenenGetalStelsel(waarde As Variant, eenheidEindproduct, Optional eenheidBeginproduct = "")
Attribute OmrekenenGetalStelsel.VB_Description = "Rekend een getal om naar een andere getalstelsel"
Attribute OmrekenenGetalStelsel.VB_ProcData.VB_Invoke_Func = " \n23"
If eenheidBeginproduct <> "" Then
     eenheidBeginproduct = LCase(eenheidBeginproduct)
        Select Case eenheidBeginproduct
            Case "d"
                waarde = waarde
            Case "h"
                waarde = Formules_Algemeen.Hex2Dec(waarde)
            Case "b"
                waarde = Formules_Algemeen.Bin2Dec(CStr(waarde))
                
            
        End Select
    End If
    eenheidEindproduct = LCase(eenheidEindproduct)
     Select Case eenheidEindproduct
            Case "d"
                waarde = waarde
            Case "h"
                waarde = Hex(waarde)
            Case "b"
                waarde = Formules_Algemeen.Dec2Bin(waarde)
            
            
        End Select
    

OmrekenenGetalStelsel = waarde
End Function


