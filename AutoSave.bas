Attribute VB_Name = "AutoSave"
Public AutoSaveRun As Boolean


Sub AutoSave()
If AutoSaveRun = True Then
Application.Caption = "Microsoft Excel - JTools AutoSave Actief (Saving)"
ActiveWorkbook.Save

Application.OnTime DateAdd("s", InstelFuncties.Config("AutoSave_IntervalTime"), Now), "AutoSave.AutoSave"
Application.Caption = "Microsoft Excel - JTools AutoSave Actief (next AutoSave " & Format(DateAdd("s", InstelFuncties.Config("AutoSave_IntervalTime"), Now), "hh:mm:ss") & ")"
End If
End Sub

Sub Start_AutoSave()
Application.Caption = "Microsoft Excel - JTools AutoSave Actief"
AutoSaveRun = True
Call AutoSave
End Sub

Sub Stop_AutoSave()
Application.Caption = ""
AutoSaveRun = False
End Sub

Sub ToggleAutoSave(control As IRibbonControl, pressed As Boolean)
If pressed Then
Call Start_AutoSave
Else
Call Stop_AutoSave
End If
End Sub

