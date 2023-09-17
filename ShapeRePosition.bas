Attribute VB_Name = "ShapeRePosition"
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

Sub SaveShapesOfWorkbook()
    
    Const SheetNaam As String = "ShapeDB"
    Dim sh As Object
    Dim sa As Object
    Dim rijnummer As Integer
    rijnummer = 1
    
    Call SearchAndAdd(SheetNaam, True)
    ThisWorkbook.Sheets(SheetNaam).Range("A1:F" & ThisWorkbook.Sheets(SheetNaam).UsedRange.Rows.Count).Value = ""
    For Each sh In ThisWorkbook.Sheets
        For Each sa In sh.Shapes
            ThisWorkbook.Sheets(SheetNaam).Range("A" & rijnummer).Value = sa.Name
            ThisWorkbook.Sheets(SheetNaam).Range("B" & rijnummer).Value = sh.Name
            ThisWorkbook.Sheets(SheetNaam).Range("C" & rijnummer).Value = sa.Top
            ThisWorkbook.Sheets(SheetNaam).Range("D" & rijnummer).Value = sa.Left
            ThisWorkbook.Sheets(SheetNaam).Range("E" & rijnummer).Value = sa.Width
            ThisWorkbook.Sheets(SheetNaam).Range("F" & rijnummer).Value = sa.Height
            rijnummer = rijnummer + 1
        Next sa
    Next sh
    
End Sub

Sub RestoreShapePos()

    Const SheetNaam As String = "ShapeDB"
    Dim c As Variant
    Dim sh As Object
    Dim sa As Object
    
    For Each sh In ThisWorkbook.Sheets
        For Each sa In sh.Shapes
            With ThisWorkbook.Sheets(SheetNaam).Range("A1:A" & ThisWorkbook.Sheets(SheetNaam).UsedRange.Rows.Count)
                Set c = .Find(sa.Name, LookAt:=xlWhole)
                If Not c Is Nothing Then
                    sa.Top = ThisWorkbook.Sheets(SheetNaam).Range("C" & c.Row).Value
                    sa.Left = ThisWorkbook.Sheets(SheetNaam).Range("D" & c.Row).Value
                    sa.Width = ThisWorkbook.Sheets(SheetNaam).Range("E" & c.Row).Value
                    sa.Height = ThisWorkbook.Sheets(SheetNaam).Range("F" & c.Row).Value
                End If
            End With
        Next sa
    Next sh
    
End Sub
