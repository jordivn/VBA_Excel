VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Tennet_Data 
   Caption         =   "Jtools - Tennet Data"
   ClientHeight    =   6735
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5445
   OleObjectBlob   =   "Tennet_Data.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Tennet_Data"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub reFresh()
Call Tennet.SetTennetData

If Tennet.TENNET_EP = 1 Then
NV = " GEACTIVEERD"
Else
NV = " niet actief"
End If

Tennet_Data.Label1.Caption = "Tijd: " & vbTab & vbTab & vbTab & vbTab & Tennet.TENNET_TIME & vbNewLine & _
"Opregelen: " & vbTab & vbTab & vbTab & Tennet.TENNET_UD & " MW" & vbNewLine & _
"Afregelen: " & vbTab & vbTab & vbTab & Tennet.TENNET_DD & " MW" & vbNewLine & _
"Reserve Opregelen: " & vbTab & Tennet.TENNET_UR & " MW" & vbNewLine & _
"Reserve Afregelen: " & vbTab & vbTab & Tennet.TENNET_DR & " MW" & vbNewLine & _
"Prijs Af:" & vbTab & vbTab & vbTab & vbTab & " €" & Tennet.TENNET_PRICE & vbNewLine & _
"Prijs Op:" & vbTab & vbTab & vbTab & vbTab & " €" & Tennet.TENNET_PRICE2 & vbNewLine & _
vbNewLine & _
"Noodvermogen is" & NV & vbNewLine & _
""
End Sub

Private Sub CommandButton1_Click()
Tennet_Data.reFresh
End Sub

Private Sub UserForm_Activate()
OnTime_Functie.TENNET_ONTIME = True
OnTime_Functie.OnTime_Tennet_Refresh


End Sub

Private Sub UserForm_Deactivate()
OnTime_Functie.TENNET_ONTIME = False
End Sub
