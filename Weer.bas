Attribute VB_Name = "Weer"
Public Stationnaam
Public STATIONCODE
Public MOMENT
Public TEMP
Public VOCHTIGHEID
Public WINDSNELHEID
Public WINDRICHTINGGR
Public windrichting
Public LUCHTDRUK
Public WINDSTOTEN
Public REGEN
Public ZICHT
Public ZON
Public TEMP10
Public WeerBerichtTitel
Public WeerBerichtSamengevat
Public WeerBerichtTekst
Public IconActueel
Public ZonOp
Public ZonOnder
Public LastData
Public DagPlus1 As New clsWeer
Public DagPlus2 As New clsWeer
Public DagPlus3 As New clsWeer
Public DagPlus4 As New clsWeer
Public DagPlus5 As New clsWeer

Sub setWeerData(Optional StationNum = "9999")
If StationNum = 9999 Then
StationNum = InstelFuncties.Config("CustomWeerStationCode")
End If
StationNum = CStr(StationNum)
If Weer.LastData < Now - TimeValue("00:05:00") Or Weer.STATIONCODE <> StationNum Then
On Error Resume Next
Set oXMLFile = CreateObject("Microsoft.XMLDOM")
    oXMLFile.async = False
    oXMLFile.Load ("https://data.buienradar.nl/1.0/feed/xml")
    
  
For I = 0 To 300

If oXMLFile.SelectSingleNode("/buienradarnl/weergegevens/actueel_weer/weerstations/weerstation[" & I & "]/stationcode/text()").NodeValue = StationNum Then
Exit For
End If
Next I


Weer.STATIONCODE = oXMLFile.SelectSingleNode("/buienradarnl/weergegevens/actueel_weer/weerstations/weerstation[" & I & "]/stationcode/text()").NodeValue
Weer.Stationnaam = oXMLFile.SelectSingleNode("/buienradarnl/weergegevens/actueel_weer/weerstations/weerstation[" & I & "]/stationnaam/text()").NodeValue
Weer.MOMENT = oXMLFile.SelectSingleNode("/buienradarnl/weergegevens/actueel_weer/weerstations/weerstation[" & I & "]/datum/text()").NodeValue
Weer.TEMP = oXMLFile.SelectSingleNode("/buienradarnl/weergegevens/actueel_weer/weerstations/weerstation[" & I & "]/temperatuurGC/text()").NodeValue
Weer.VOCHTIGHEID = oXMLFile.SelectSingleNode("/buienradarnl/weergegevens/actueel_weer/weerstations/weerstation[" & I & "]/luchtvochtigheid/text()").NodeValue
Weer.WINDSNELHEID = oXMLFile.SelectSingleNode("/buienradarnl/weergegevens/actueel_weer/weerstations/weerstation[" & I & "]/windsnelheidMS/text()").NodeValue
Weer.WINDRICHTINGGR = oXMLFile.SelectSingleNode("/buienradarnl/weergegevens/actueel_weer/weerstations/weerstation[" & I & "]/windrichtingGR/text()").NodeValue
Weer.windrichting = oXMLFile.SelectSingleNode("/buienradarnl/weergegevens/actueel_weer/weerstations/weerstation[" & I & "]/windrichting/text()").NodeValue
Weer.LUCHTDRUK = oXMLFile.SelectSingleNode("/buienradarnl/weergegevens/actueel_weer/weerstations/weerstation[" & I & "]/luchtdruk/text()").NodeValue
Weer.WINDSTOTEN = oXMLFile.SelectSingleNode("/buienradarnl/weergegevens/actueel_weer/weerstations/weerstation[" & I & "]/windstotenMS/text()").NodeValue
Weer.REGEN = oXMLFile.SelectSingleNode("/buienradarnl/weergegevens/actueel_weer/weerstations/weerstation[" & I & "]/regenMMPU/text()").NodeValue
Weer.ZICHT = oXMLFile.SelectSingleNode("/buienradarnl/weergegevens/actueel_weer/weerstations/weerstation[" & I & "]/zichtmeters/text()").NodeValue
Weer.ZON = oXMLFile.SelectSingleNode("/buienradarnl/weergegevens/actueel_weer/weerstations/weerstation[" & I & "]/zonintensiteitWM2/text()").NodeValue
Weer.TEMP10 = oXMLFile.SelectSingleNode("/buienradarnl/weergegevens/actueel_weer/weerstations/weerstation[" & I & "]/temperatuur10cm/text()").NodeValue
Weer.IconActueel = oXMLFile.SelectSingleNode("/buienradarnl/weergegevens/actueel_weer/buienradar/icoonactueel/text()").NodeValue

Weer.ZonOp = oXMLFile.SelectSingleNode("/buienradarnl/weergegevens/actueel_weer/buienradar/zonopkomst/text()").NodeValue
Weer.ZonOnder = oXMLFile.SelectSingleNode("/buienradarnl/weergegevens/actueel_weer/buienradar/zononder/text()").NodeValue


Weer.WeerBerichtTitel = oXMLFile.SelectSingleNode("/buienradarnl/weergegevens/verwachting_vandaag/titel/text()").NodeValue
Weer.WeerBerichtSamengevat = oXMLFile.SelectSingleNode("/buienradarnl/weergegevens/verwachting_vandaag/samenvatting/text()").NodeValue
Weer.WeerBerichtTekst = oXMLFile.SelectSingleNode("/buienradarnl/weergegevens/verwachting_vandaag/formattedtekst/text()").NodeValue

Weer.DagPlus1.dagweek = oXMLFile.SelectSingleNode("/buienradarnl/weergegevens/verwachting_meerdaags/dag-plus1/dagweek/text()").NodeValue
Weer.DagPlus1.datum = oXMLFile.SelectSingleNode("/buienradarnl/weergegevens/verwachting_meerdaags/dag-plus1/datum/text()").NodeValue
Weer.DagPlus1.kanszon = oXMLFile.SelectSingleNode("/buienradarnl/weergegevens/verwachting_meerdaags/dag-plus1/kanszon/text()").NodeValue
Weer.DagPlus1.kansregen = oXMLFile.SelectSingleNode("/buienradarnl/weergegevens/verwachting_meerdaags/dag-plus1/kansregen/text()").NodeValue
Weer.DagPlus1.minmmregen = oXMLFile.SelectSingleNode("/buienradarnl/weergegevens/verwachting_meerdaags/dag-plus1/minmmregen/text()").NodeValue
Weer.DagPlus1.maxmmregen = oXMLFile.SelectSingleNode("/buienradarnl/weergegevens/verwachting_meerdaags/dag-plus1/maxmmregen/text()").NodeValue
Weer.DagPlus1.mintemp = oXMLFile.SelectSingleNode("/buienradarnl/weergegevens/verwachting_meerdaags/dag-plus1/mintemp/text()").NodeValue
Weer.DagPlus1.mintempmax = oXMLFile.SelectSingleNode("/buienradarnl/weergegevens/verwachting_meerdaags/dag-plus1/mintempmax/text()").NodeValue
Weer.DagPlus1.maxtemp = oXMLFile.SelectSingleNode("/buienradarnl/weergegevens/verwachting_meerdaags/dag-plus1/maxtemp/text()").NodeValue
Weer.DagPlus1.maxtempmax = oXMLFile.SelectSingleNode("/buienradarnl/weergegevens/verwachting_meerdaags/dag-plus1/maxtempmax/text()").NodeValue
Weer.DagPlus1.windrichting = oXMLFile.SelectSingleNode("/buienradarnl/weergegevens/verwachting_meerdaags/dag-plus1/windrichting/text()").NodeValue
Weer.DagPlus1.windkracht = oXMLFile.SelectSingleNode("/buienradarnl/weergegevens/verwachting_meerdaags/dag-plus1/windkracht/text()").NodeValue

Weer.DagPlus2.dagweek = oXMLFile.SelectSingleNode("/buienradarnl/weergegevens/verwachting_meerdaags/dag-plus2/dagweek/text()").NodeValue
Weer.DagPlus2.datum = oXMLFile.SelectSingleNode("/buienradarnl/weergegevens/verwachting_meerdaags/dag-plus2/datum/text()").NodeValue
Weer.DagPlus2.kanszon = oXMLFile.SelectSingleNode("/buienradarnl/weergegevens/verwachting_meerdaags/dag-plus2/kanszon/text()").NodeValue
Weer.DagPlus2.kansregen = oXMLFile.SelectSingleNode("/buienradarnl/weergegevens/verwachting_meerdaags/dag-plus2/kansregen/text()").NodeValue
Weer.DagPlus2.minmmregen = oXMLFile.SelectSingleNode("/buienradarnl/weergegevens/verwachting_meerdaags/dag-plus2/minmmregen/text()").NodeValue
Weer.DagPlus2.maxmmregen = oXMLFile.SelectSingleNode("/buienradarnl/weergegevens/verwachting_meerdaags/dag-plus2/maxmmregen/text()").NodeValue
Weer.DagPlus2.mintemp = oXMLFile.SelectSingleNode("/buienradarnl/weergegevens/verwachting_meerdaags/dag-plus2/mintemp/text()").NodeValue
Weer.DagPlus2.mintempmax = oXMLFile.SelectSingleNode("/buienradarnl/weergegevens/verwachting_meerdaags/dag-plus2/mintempmax/text()").NodeValue
Weer.DagPlus2.maxtemp = oXMLFile.SelectSingleNode("/buienradarnl/weergegevens/verwachting_meerdaags/dag-plus2/maxtemp/text()").NodeValue
Weer.DagPlus2.maxtempmax = oXMLFile.SelectSingleNode("/buienradarnl/weergegevens/verwachting_meerdaags/dag-plus2/maxtempmax/text()").NodeValue
Weer.DagPlus2.windrichting = oXMLFile.SelectSingleNode("/buienradarnl/weergegevens/verwachting_meerdaags/dag-plus2/windrichting/text()").NodeValue
Weer.DagPlus2.windkracht = oXMLFile.SelectSingleNode("/buienradarnl/weergegevens/verwachting_meerdaags/dag-plus2/windkracht/text()").NodeValue

Weer.DagPlus3.dagweek = oXMLFile.SelectSingleNode("/buienradarnl/weergegevens/verwachting_meerdaags/dag-plus3/dagweek/text()").NodeValue
Weer.DagPlus3.datum = oXMLFile.SelectSingleNode("/buienradarnl/weergegevens/verwachting_meerdaags/dag-plus3/datum/text()").NodeValue
Weer.DagPlus3.kanszon = oXMLFile.SelectSingleNode("/buienradarnl/weergegevens/verwachting_meerdaags/dag-plus3/kanszon/text()").NodeValue
Weer.DagPlus3.kansregen = oXMLFile.SelectSingleNode("/buienradarnl/weergegevens/verwachting_meerdaags/dag-plus3/kansregen/text()").NodeValue
Weer.DagPlus3.minmmregen = oXMLFile.SelectSingleNode("/buienradarnl/weergegevens/verwachting_meerdaags/dag-plus3/minmmregen/text()").NodeValue
Weer.DagPlus3.maxmmregen = oXMLFile.SelectSingleNode("/buienradarnl/weergegevens/verwachting_meerdaags/dag-plus3/maxmmregen/text()").NodeValue
Weer.DagPlus3.mintemp = oXMLFile.SelectSingleNode("/buienradarnl/weergegevens/verwachting_meerdaags/dag-plus3/mintemp/text()").NodeValue
Weer.DagPlus3.mintempmax = oXMLFile.SelectSingleNode("/buienradarnl/weergegevens/verwachting_meerdaags/dag-plus3/mintempmax/text()").NodeValue
Weer.DagPlus3.maxtemp = oXMLFile.SelectSingleNode("/buienradarnl/weergegevens/verwachting_meerdaags/dag-plus3/maxtemp/text()").NodeValue
Weer.DagPlus3.maxtempmax = oXMLFile.SelectSingleNode("/buienradarnl/weergegevens/verwachting_meerdaags/dag-plus3/maxtempmax/text()").NodeValue
Weer.DagPlus3.windrichting = oXMLFile.SelectSingleNode("/buienradarnl/weergegevens/verwachting_meerdaags/dag-plus3/windrichting/text()").NodeValue
Weer.DagPlus3.windkracht = oXMLFile.SelectSingleNode("/buienradarnl/weergegevens/verwachting_meerdaags/dag-plus3/windkracht/text()").NodeValue

Weer.DagPlus4.dagweek = oXMLFile.SelectSingleNode("/buienradarnl/weergegevens/verwachting_meerdaags/dag-plus4/dagweek/text()").NodeValue
Weer.DagPlus4.datum = oXMLFile.SelectSingleNode("/buienradarnl/weergegevens/verwachting_meerdaags/dag-plus4/datum/text()").NodeValue
Weer.DagPlus4.kanszon = oXMLFile.SelectSingleNode("/buienradarnl/weergegevens/verwachting_meerdaags/dag-plus4/kanszon/text()").NodeValue
Weer.DagPlus4.kansregen = oXMLFile.SelectSingleNode("/buienradarnl/weergegevens/verwachting_meerdaags/dag-plus4/kansregen/text()").NodeValue
Weer.DagPlus4.minmmregen = oXMLFile.SelectSingleNode("/buienradarnl/weergegevens/verwachting_meerdaags/dag-plus4/minmmregen/text()").NodeValue
Weer.DagPlus4.maxmmregen = oXMLFile.SelectSingleNode("/buienradarnl/weergegevens/verwachting_meerdaags/dag-plus4/maxmmregen/text()").NodeValue
Weer.DagPlus4.mintemp = oXMLFile.SelectSingleNode("/buienradarnl/weergegevens/verwachting_meerdaags/dag-plus4/mintemp/text()").NodeValue
Weer.DagPlus4.mintempmax = oXMLFile.SelectSingleNode("/buienradarnl/weergegevens/verwachting_meerdaags/dag-plus4/mintempmax/text()").NodeValue
Weer.DagPlus4.maxtemp = oXMLFile.SelectSingleNode("/buienradarnl/weergegevens/verwachting_meerdaags/dag-plus4/maxtemp/text()").NodeValue
Weer.DagPlus4.maxtempmax = oXMLFile.SelectSingleNode("/buienradarnl/weergegevens/verwachting_meerdaags/dag-plus4/maxtempmax/text()").NodeValue
Weer.DagPlus4.windrichting = oXMLFile.SelectSingleNode("/buienradarnl/weergegevens/verwachting_meerdaags/dag-plus4/windrichting/text()").NodeValue
Weer.DagPlus4.windkracht = oXMLFile.SelectSingleNode("/buienradarnl/weergegevens/verwachting_meerdaags/dag-plus4/windkracht/text()").NodeValue

Weer.DagPlus5.dagweek = oXMLFile.SelectSingleNode("/buienradarnl/weergegevens/verwachting_meerdaags/dag-plus5/dagweek/text()").NodeValue
Weer.DagPlus5.datum = oXMLFile.SelectSingleNode("/buienradarnl/weergegevens/verwachting_meerdaags/dag-plus5/datum/text()").NodeValue
Weer.DagPlus5.kanszon = oXMLFile.SelectSingleNode("/buienradarnl/weergegevens/verwachting_meerdaags/dag-plus5/kanszon/text()").NodeValue
Weer.DagPlus5.kansregen = oXMLFile.SelectSingleNode("/buienradarnl/weergegevens/verwachting_meerdaags/dag-plus5/kansregen/text()").NodeValue
Weer.DagPlus5.minmmregen = oXMLFile.SelectSingleNode("/buienradarnl/weergegevens/verwachting_meerdaags/dag-plus5/minmmregen/text()").NodeValue
Weer.DagPlus5.maxmmregen = oXMLFile.SelectSingleNode("/buienradarnl/weergegevens/verwachting_meerdaags/dag-plus5/maxmmregen/text()").NodeValue
Weer.DagPlus5.mintemp = oXMLFile.SelectSingleNode("/buienradarnl/weergegevens/verwachting_meerdaags/dag-plus5/mintemp/text()").NodeValue
Weer.DagPlus5.mintempmax = oXMLFile.SelectSingleNode("/buienradarnl/weergegevens/verwachting_meerdaags/dag-plus5/mintempmax/text()").NodeValue
Weer.DagPlus5.maxtemp = oXMLFile.SelectSingleNode("/buienradarnl/weergegevens/verwachting_meerdaags/dag-plus5/maxtemp/text()").NodeValue
Weer.DagPlus5.maxtempmax = oXMLFile.SelectSingleNode("/buienradarnl/weergegevens/verwachting_meerdaags/dag-plus5/maxtempmax/text()").NodeValue
Weer.DagPlus5.windrichting = oXMLFile.SelectSingleNode("/buienradarnl/weergegevens/verwachting_meerdaags/dag-plus5/windrichting/text()").NodeValue
Weer.DagPlus5.windkracht = oXMLFile.SelectSingleNode("/buienradarnl/weergegevens/verwachting_meerdaags/dag-plus5/windkracht/text()").NodeValue



Weer.LastData = Now
End If

End Sub

'lijst van weerstations
'weerstation zoeken op plaats

Function getWeerData_StationNaam(Optional StationNum = 9999)
Attribute getWeerData_StationNaam.VB_ProcData.VB_Invoke_Func = " \n27"
If StationNum = 9999 Then
StationNum = InstelFuncties.Config("CustomWeerStationCode")
End If
Call Weer.setWeerData(StationNum)
getWeerData_StationNaam = Weer.Stationnaam
End Function

Function getWeerData_Moment(Optional StationNum = 9999)
Attribute getWeerData_Moment.VB_ProcData.VB_Invoke_Func = " \n27"
If StationNum = 9999 Then
StationNum = InstelFuncties.Config("CustomWeerStationCode")
End If
Call Weer.setWeerData(StationNum)
getWeerData_Moment = Weer.MOMENT
End Function
Function getWeerData_Temperatuur(Optional StationNum = 9999)
Attribute getWeerData_Temperatuur.VB_ProcData.VB_Invoke_Func = " \n27"
If StationNum = 9999 Then
StationNum = InstelFuncties.Config("CustomWeerStationCode")
End If
Call Weer.setWeerData(StationNum)
getWeerData_Temperatuur = Weer.TEMP
End Function
Function getWeerData_Vochtigheid(Optional StationNum = 9999)
Attribute getWeerData_Vochtigheid.VB_ProcData.VB_Invoke_Func = " \n27"
If StationNum = 9999 Then
StationNum = InstelFuncties.Config("CustomWeerStationCode")
End If
Call Weer.setWeerData(StationNum)
getWeerData_Vochtigheid = Weer.VOCHTIGHEID
End Function
Function getWeerData_Windsnelheid(Optional StationNum = 9999)
Attribute getWeerData_Windsnelheid.VB_ProcData.VB_Invoke_Func = " \n27"
If StationNum = 9999 Then
StationNum = InstelFuncties.Config("CustomWeerStationCode")
End If
Call Weer.setWeerData(StationNum)
getWeerData_Windsnelheid = Weer.WINDSNELHEID
End Function
Function getWeerData_WindrichtingGR(Optional StationNum = 9999)
Attribute getWeerData_WindrichtingGR.VB_ProcData.VB_Invoke_Func = " \n27"
If StationNum = 9999 Then
StationNum = InstelFuncties.Config("CustomWeerStationCode")
End If
Call Weer.setWeerData(StationNum)
getWeerData_WindrichtingGR = Weer.WINDRICHTINGGR
End Function
Function getWeerData_Windrichting(Optional StationNum = 9999)
Attribute getWeerData_Windrichting.VB_ProcData.VB_Invoke_Func = " \n27"
If StationNum = 9999 Then
StationNum = InstelFuncties.Config("CustomWeerStationCode")
End If
Call Weer.setWeerData(StationNum)
getWeerData_Windrichting = Weer.windrichting
End Function
Function getWeerData_Luchtdruk(Optional StationNum = 9999)
Attribute getWeerData_Luchtdruk.VB_ProcData.VB_Invoke_Func = " \n27"
If StationNum = 9999 Then
StationNum = InstelFuncties.Config("CustomWeerStationCode")
End If
Call Weer.setWeerData(StationNum)
getWeerData_Luchtdruk = Weer.LUCHTDRUK
End Function
Function getWeerData_Windstoten(Optional StationNum = 9999)
Attribute getWeerData_Windstoten.VB_ProcData.VB_Invoke_Func = " \n27"
If StationNum = 9999 Then
StationNum = InstelFuncties.Config("CustomWeerStationCode")
End If
Call Weer.setWeerData(StationNum)
getWeerData_Windstoten = Weer.WINDSTOTEN
End Function
Function getWeerData_Regen(Optional StationNum = 9999)
Attribute getWeerData_Regen.VB_ProcData.VB_Invoke_Func = " \n27"
If StationNum = 9999 Then
StationNum = InstelFuncties.Config("CustomWeerStationCode")
End If
Call Weer.setWeerData(StationNum)
getWeerData_Regen = Weer.REGEN
End Function
Function getWeerData_Zicht(Optional StationNum = 9999)
Attribute getWeerData_Zicht.VB_ProcData.VB_Invoke_Func = " \n27"
If StationNum = 9999 Then
StationNum = InstelFuncties.Config("CustomWeerStationCode")
End If
Call Weer.setWeerData(StationNum)
getWeerData_Zicht = Weer.ZICHT
End Function
Function getWeerData_ZonIntensiteit(Optional StationNum = 9999)
Attribute getWeerData_ZonIntensiteit.VB_ProcData.VB_Invoke_Func = " \n27"
If StationNum = 9999 Then
StationNum = InstelFuncties.Config("CustomWeerStationCode")
End If
Call Weer.setWeerData(StationNum)
getWeerData_ZonIntensiteit = Weer.ZON
End Function
Function getWeerData_TemperatuurOp10cm(Optional StationNum = 9999)
Attribute getWeerData_TemperatuurOp10cm.VB_ProcData.VB_Invoke_Func = " \n27"
If StationNum = 9999 Then
StationNum = InstelFuncties.Config("CustomWeerStationCode")
End If
Call Weer.setWeerData(StationNum)
getWeerData_TemperatuurOp10cm = Weer.TEMP10
End Function
