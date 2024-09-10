VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Weer_Data 
   Caption         =   "Jtools - Weer informatie"
   ClientHeight    =   6330
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16110
   OleObjectBlob   =   "Weer_Data.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Weer_Data"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
Call UserForm_Activate
End Sub
 
 
 
Private Sub UserForm_Activate()
Call Weer.setWeerData

tempText = Weer.Stationnaam
tempText = tempText & vbNewLine & "Moment: " & vbTab & vbTab & Weer.MOMENT
tempText = tempText & vbNewLine & "Temperatuur: " & vbTab & vbTab & vbTab & Weer.TEMP & "°C"
tempText = tempText & vbNewLine & "Vochtigheid: " & vbTab & vbTab & vbTab & Weer.VOCHTIGHEID & "%"
tempText = tempText & vbNewLine & "Windsnelheid (M/s): " & vbTab & vbTab & Weer.WINDSNELHEID & " m/s"
tempText = tempText & vbNewLine & "Windrichting: " & vbTab & vbTab & vbTab & Weer.WINDRICHTINGGR & "° (" & Weer.windrichting & ")"
tempText = tempText & vbNewLine & "Luchtdruk: " & vbTab & vbTab & vbTab & Weer.LUCHTDRUK & " Pa"
tempText = tempText & vbNewLine & "Windstoten (M/s): " & vbTab & vbTab & Weer.WINDSTOTEN & " m/s"
tempText = tempText & vbNewLine & "Regen (mm/u): " & vbTab & vbTab & Weer.REGEN & " mm/h"
tempText = tempText & vbNewLine & "Zichtmeters (M): " & vbTab & vbTab & Weer.ZICHT & " meter"
tempText = tempText & vbNewLine & "Zonintensiteit (W/M2): " & vbTab & Weer.ZON & " W/m²"
tempText = tempText & vbNewLine & "Temperatuur op 10cm: " & vbTab & Weer.TEMP10 & "°C"

Weer_Data.Label1.Caption = tempText
'Weer_Data.WebBrowser1.Navigate Weer.IconActueel

Weer_Data.ZonOp = Format(CDate(Weer.ZonOp), "hh:mm")
Weer_Data.ZonOnder = Format(CDate(Weer.ZonOnder), "hh:mm")

Weer_Data.WeerBerichtTitel.Caption = Weer.WeerBerichtTitel
Weer_Data.WeerBerichtSamengevat.Caption = Weer.WeerBerichtSamengevat
Weer_Data.WeerBerichtTekst.Caption = Replace(Replace(Weer.WeerBerichtTekst, "&nbsp;", vbNewLine), "&agrave;", "á")

tempText = "Dag" & vbNewLine & vbNewLine _
                & "Zon | Regen" & vbNewLine & vbNewLine _
                & "Regen (mm)" & vbNewLine & vbNewLine _
                & "Temperatuur (°C)" & vbNewLine & vbNewLine _
                & "Wind"
Weer_Data.VooruitLegenda.Caption = tempText

tempText = Weer.DagPlus1.dagweek & vbNewLine & vbNewLine _
                & Weer.DagPlus1.kanszon & "% | " & Weer.DagPlus1.kansregen & "%" & vbNewLine & vbNewLine _
                & Weer.DagPlus1.minmmregen & " - " & Weer.DagPlus1.maxmmregen & vbNewLine & vbNewLine _
                & Weer.DagPlus1.mintemp & " - " & Weer.DagPlus1.maxtempmax & vbNewLine & vbNewLine _
                & Weer.DagPlus1.windkracht & " " & Weer.DagPlus1.windrichting
Weer_Data.VooruitDag1.Caption = tempText

tempText = Weer.DagPlus2.dagweek & vbNewLine & vbNewLine _
                & Weer.DagPlus2.kanszon & "% | " & Weer.DagPlus2.kansregen & "%" & vbNewLine & vbNewLine _
                & Weer.DagPlus2.minmmregen & " - " & Weer.DagPlus2.maxmmregen & vbNewLine & vbNewLine _
                & Weer.DagPlus2.mintemp & " - " & Weer.DagPlus2.maxtempmax & vbNewLine & vbNewLine _
                & Weer.DagPlus2.windkracht & " " & Weer.DagPlus2.windrichting
Weer_Data.VooruitDag2.Caption = tempText

tempText = Weer.DagPlus3.dagweek & vbNewLine & vbNewLine _
                & Weer.DagPlus3.kanszon & "% | " & Weer.DagPlus3.kansregen & "%" & vbNewLine & vbNewLine _
                & Weer.DagPlus3.minmmregen & " - " & Weer.DagPlus3.maxmmregen & vbNewLine & vbNewLine _
                & Weer.DagPlus3.mintemp & " - " & Weer.DagPlus3.maxtempmax & vbNewLine & vbNewLine _
                & Weer.DagPlus3.windkracht & " " & Weer.DagPlus3.windrichting
Weer_Data.VooruitDag3.Caption = tempText

tempText = Weer.DagPlus4.dagweek & vbNewLine & vbNewLine _
                & Weer.DagPlus4.kanszon & "% | " & Weer.DagPlus4.kansregen & "%" & vbNewLine & vbNewLine _
                & Weer.DagPlus4.minmmregen & " - " & Weer.DagPlus4.maxmmregen & vbNewLine & vbNewLine _
                & Weer.DagPlus4.mintemp & " - " & Weer.DagPlus4.maxtempmax & vbNewLine & vbNewLine _
                & Weer.DagPlus4.windkracht & " " & Weer.DagPlus4.windrichting
Weer_Data.VooruitDag4.Caption = tempText

tempText = Weer.DagPlus5.dagweek & vbNewLine & vbNewLine _
                & Weer.DagPlus5.kanszon & "% | " & Weer.DagPlus5.kansregen & "%" & vbNewLine & vbNewLine _
                & Weer.DagPlus5.minmmregen & " - " & Weer.DagPlus5.maxmmregen & vbNewLine & vbNewLine _
                & Weer.DagPlus5.mintemp & " - " & Weer.DagPlus5.maxtempmax & vbNewLine & vbNewLine _
                & Weer.DagPlus5.windkracht & " " & Weer.DagPlus5.windrichting
Weer_Data.VooruitDag5.Caption = tempText

tempText = Weer.DagPlus1.dagweek & vbNewLine & vbNewLine _
                & Weer.DagPlus1.kanszon & "% | " & Weer.DagPlus1.kansregen & "%" & vbNewLine & vbNewLine _
                & Weer.DagPlus1.minmmregen & " - " & Weer.DagPlus1.maxmmregen & vbNewLine & vbNewLine _
                & Weer.DagPlus1.mintemp & " - " & Weer.DagPlus1.maxtempmax & vbNewLine & vbNewLine _
                & Weer.DagPlus1.windkracht & " " & Weer.DagPlus1.windrichting
Weer_Data.VooruitDag1.Caption = tempText

End Sub

