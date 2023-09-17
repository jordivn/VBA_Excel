Attribute VB_Name = "Tennet"
Public TENNET_SEQUENCE
Public TENNET_TIME
Public TENNET_UD
Public TENNET_DD
Public TENNET_UR
Public TENNET_DR
Public TENNET_EP
Public TENNET_PRICE
Public TENNET_PRICE2
Public TENNET_LastData
Public TENNET_Item

Sub SetTennetData(Optional ItemNum = 1)

If Tennet.TENNET_LastData < Now - TimeValue("00:01:00") Or ItemNum <> Tennet.TENNET_Item Then
On Error Resume Next
Set oXMLFile = CreateObject("Microsoft.XMLDOM")
    oXMLFile.async = False
    oXMLFile.Load ("https://www.tennet.org/xml/balancedeltaprices/balans-delta.xml")

Tennet.TENNET_Item = ItemNum
Tennet.TENNET_SEQUENCE = oXMLFile.SelectSingleNode("/BALANCE_DELTA/RECORD[" & ItemNum & "]/SEQUENCE_NUMBER/text()").NodeValue
Tennet.TENNET_TIME = oXMLFile.SelectSingleNode("/BALANCE_DELTA/RECORD[" & ItemNum & "]/TIME/text()").NodeValue
Tennet.TENNET_UD = oXMLFile.SelectSingleNode("/BALANCE_DELTA/RECORD[" & ItemNum & "]/UPWARD_DISPATCH/text()").NodeValue
Tennet.TENNET_DD = oXMLFile.SelectSingleNode("/BALANCE_DELTA/RECORD[" & ItemNum & "]/DOWNWARD_DISPATCH/text()").NodeValue
Tennet.TENNET_UR = oXMLFile.SelectSingleNode("/BALANCE_DELTA/RECORD[" & ItemNum & "]/RESERVE_UPWARD_DISPATCH/text()").NodeValue
Tennet.TENNET_DR = oXMLFile.SelectSingleNode("/BALANCE_DELTA/RECORD[" & ItemNum & "]/RESERVE_DOWNWARD_DISPATCH/text()").NodeValue
Tennet.TENNET_EP = oXMLFile.SelectSingleNode("/BALANCE_DELTA/RECORD[" & ItemNum & "]/EMERGENCY_POWER/text()").NodeValue
Tennet.TENNET_PRICE = oXMLFile.SelectSingleNode("/BALANCE_DELTA/RECORD[" & ItemNum & "]/MIN_PRICE/text()").NodeValue
Tennet.TENNET_PRICE2 = oXMLFile.SelectSingleNode("/BALANCE_DELTA/RECORD[" & ItemNum & "]/MAX_PRICE/text()").NodeValue
Tennet.TENNET_LastData = Now
End If
End Sub

Function getTennetData_Tijd(Optional ItemNum = 1)
Attribute getTennetData_Tijd.VB_ProcData.VB_Invoke_Func = " \n26"
Call Tennet.SetTennetData(ItemNum)
getTennetData_Tijd = Tennet.TENNET_TIME
End Function


Function getTennetData_OpregelVermogen(Optional ItemNum = 1)
Attribute getTennetData_OpregelVermogen.VB_ProcData.VB_Invoke_Func = " \n26"
Call Tennet.SetTennetData(ItemNum)
getTennetData_OpregelVermogen = Tennet.TENNET_UD
End Function

Function getTennetData_AfregelVermogen(Optional ItemNum = 1)
Attribute getTennetData_AfregelVermogen.VB_ProcData.VB_Invoke_Func = " \n26"
Call Tennet.SetTennetData(ItemNum)
getTennetData_AfregelVermogen = Tennet.TENNET_DD
End Function

Function getTennetData_OpregelVermogenReserve(Optional ItemNum = 1)
Attribute getTennetData_OpregelVermogenReserve.VB_ProcData.VB_Invoke_Func = " \n26"
Call Tennet.SetTennetData(ItemNum)
getTennetData_OpregelVermogenReserve = Tennet.TENNET_UR
End Function

Function getTennetData_AfregelVermogenReserve(Optional ItemNum = 1)
Attribute getTennetData_AfregelVermogenReserve.VB_ProcData.VB_Invoke_Func = " \n26"
Call Tennet.SetTennetData(ItemNum)
getTennetData_AfregelVermogenReserve = Tennet.TENNET_DR
End Function

Function getTennetData_Noodvermogen(Optional ItemNum = 1)
Attribute getTennetData_Noodvermogen.VB_ProcData.VB_Invoke_Func = " \n26"
Call Tennet.SetTennetData(ItemNum)
getTennetData_Noodvermogen = Tennet.TENNET_EP
End Function

Function getTennetData_PrijsMin(Optional ItemNum = 1)
Attribute getTennetData_PrijsMin.VB_ProcData.VB_Invoke_Func = " \n26"
Call Tennet.SetTennetData(ItemNum)
getTennetData_PrijsMin = Tennet.TENNET_PRICE
End Function

Function getTennetData_PrijsMax(Optional ItemNum = 1)
Attribute getTennetData_PrijsMax.VB_ProcData.VB_Invoke_Func = " \n26"
Call Tennet.SetTennetData(ItemNum)
getTennetData_PrijsMax = Tennet.TENNET_PRICE2
End Function

