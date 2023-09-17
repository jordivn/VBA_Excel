Attribute VB_Name = "SheetsVBAFunctions"
Function CheckIfSheetExists(SheetName As String) As Boolean
Attribute CheckIfSheetExists.VB_Description = "Controleerd of een sheet bestaat in het werkboek"
Attribute CheckIfSheetExists.VB_ProcData.VB_Invoke_Func = " \n33"
      CheckIfSheetExists = False
      For Each WS In Worksheets
        If SheetName = WS.Name Then
          CheckIfSheetExists = True
          Exit Function
        End If
      Next WS
End Function
