Attribute VB_Name = "NewTreat"
Sub OpenWorkbookInNewTreat(FileLink As String)
    Dim obj As Object

    Set obj = CreateObject("Excel.Application")
    obj.Workbooks.Open (FileLink)
    obj.Visible = False
End Sub
