Attribute VB_Name = "SearchAndAddSheet"
Sub SearchAndAdd(SheetNaam As String, SheetVisible As Boolean)
    
    Dim SheetBestaat As Boolean
    Dim sh As Object
 
    SheetBestaat = False
    For Each sh In ThisWorkbook.Sheets
        If sh.Name = SheetNaam Then
            SheetBestaat = True
        End If
    Next sh
    If SheetBestaat = False Then
        With ThisWorkbook.Sheets.Add
            .Name = SheetNaam
            .Visible = SheetVisible
        End With
    End If

End Sub
