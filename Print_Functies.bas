Attribute VB_Name = "Print_Functies"
Sub SelectionPrintLandA4()
Application.PrintCommunication = False
    With ActiveSheet.PageSetup
        .PaperSize = xlPaperA4
        .Orientation = xlLandscape
        .PrintQuality = 1200
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
        .CenterHorizontally = True
        .CenterVertically = True
        
        
    End With
    Application.PrintCommunication = True
    Selection.PrintOut copies:=1, collate:=True
End Sub

Sub SelectionPrintPortA4()
Application.PrintCommunication = False
    With ActiveSheet.PageSetup
        .PaperSize = xlPaperA4
        .Orientation = xlPortrait
        .PrintQuality = 1200
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
        .CenterHorizontally = True
        .CenterVertically = True
        
        
    End With
    Application.PrintCommunication = True
    Selection.PrintOut copies:=1, collate:=True
End Sub

Sub SelectionPrintLandA3()
Application.PrintCommunication = False
    With ActiveSheet.PageSetup
        .PaperSize = xlPaperA3
        .Orientation = xlLandscape
        .PrintQuality = 1200
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
        .CenterHorizontally = True
        .CenterVertically = True
        
        
    End With
    Application.PrintCommunication = True
    Selection.PrintOut copies:=1, collate:=True
End Sub

Sub SelectionPrintPortA3()
Application.PrintCommunication = False
    With ActiveSheet.PageSetup
        .PaperSize = xlPaperA3
        .Orientation = xlPortrait
        .PrintQuality = 1200
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
        .CenterHorizontally = True
        .CenterVertically = True
        
        
    End With
    Application.PrintCommunication = True
    Selection.PrintOut copies:=1, collate:=True
End Sub


'=====================
'Worksheet
'=====================

Sub WorksheetPrintLandA4()
Application.PrintCommunication = False
    With ActiveSheet.PageSetup
        .PaperSize = xlPaperA4
        .Orientation = xlLandscape
        .PrintQuality = 1200
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
        .CenterHorizontally = True
        .CenterVertically = True
        
        
    End With
    Application.PrintCommunication = True
    ActiveSheet.PrintOut copies:=1, collate:=True
End Sub

Sub WorksheetPrintPortA4()
Application.PrintCommunication = False
    With ActiveSheet.PageSetup
        .PaperSize = xlPaperA4
        .Orientation = xlPortrait
        .PrintQuality = 1200
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
        .CenterHorizontally = True
        .CenterVertically = True
        
        
    End With
    Application.PrintCommunication = True
    ActiveSheet.PrintOut copies:=1, collate:=True
End Sub

Sub WorksheetPrintLandA3()
Application.PrintCommunication = False
    With ActiveSheet.PageSetup
        .PaperSize = xlPaperA3
        .Orientation = xlLandscape
        .PrintQuality = 1200
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
        .CenterHorizontally = True
        .CenterVertically = True
        
        
    End With
    Application.PrintCommunication = True
    ActiveSheet.PrintOut copies:=1, collate:=True
End Sub

Sub WorksheetPrintPortA3()
Application.PrintCommunication = False
    With ActiveSheet.PageSetup
        .PaperSize = xlPaperA3
        .Orientation = xlPortrait
        .PrintQuality = 1200
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
        .CenterHorizontally = True
        .CenterVertically = True
        
        
    End With
    Application.PrintCommunication = True
    ActiveSheet.PrintOut copies:=1, collate:=True
End Sub



