Attribute VB_Name = "SheetsVBAFunctions"
'===============================================
'@details       Function for checking if sheet exists
'@author        Jordi van Nistelrooij @ Webs en Systems
'@email         info@websensystems.nl
'@version       1.0.0
'@date          2024-09-10
'@copyright     Non of these scripts maybe copied or modified without permission of the author
'===============================================

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
