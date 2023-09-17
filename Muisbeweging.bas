Attribute VB_Name = "Module1"
Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long

Type POINTAPI
 X_Pos As Long
 Y_Pos As Long
End Type


Sub makeamove()
Dim Hold As POINTAPI
GetCursorPos Hold
SetCursorPos Hold.X_Pos + 100, Hold.Y_Pos
SetCursorPos Hold.X_Pos, Hold.Y_Pos
Application.SendKeys "{ESC}"

Application.OnTime Now + TimeValue("00:01:00"), "Module1.makeamove"
Range("A1").Value = Now
End Sub
