Attribute VB_Name = "Print_Functies"
'===============================================
'@details       Shortcut printing functions
'@author        Jordi van Nistelrooij @ Webs en Systems
'@email         info@websensystems.nl
'@version       1.0.0
'@date          2024-09-10
'@copyright     Non of these scripts maybe copied or modified without permission of the author
'===============================================

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



