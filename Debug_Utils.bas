Attribute VB_Name = "Debug_Utils"
Option Explicit
Option Base 0

'+--------------------+
'| Recovery Functions |
'+--------------------+
Public Sub ResetQt_Hard()
    'Remove and reset
    ActiveWindow.Application.DisplayAlerts = False
    Sheets("Qt").Select
    ActiveWindow.SelectedSheets.Delete
    Sheets("1 Gal").Select
    Sheets.Add before:=Sheets("1 Gal")
    Sheets(1).Select
    Sheets(1).Name = "Qt"
    ActiveWindow.Application.DisplayAlerts = True
    
    Inserted_QT = False
    
    'Set cell dimensions
    'Columns first...
    Range("A1").Select
    Do While (ActiveCell.Column < 249)
        ActiveCell.Columns("A:A").EntireColumn.ColumnWidth = 19.86
        ActiveCell.Offset(0, 1).Columns("A:A").EntireColumn.ColumnWidth = 16.57
        ActiveCell.Offset(0, 2).Columns("A:A").EntireColumn.ColumnWidth = 40.43
        ActiveCell.Offset(0, 3).Columns("A:A").EntireColumn.ColumnWidth = 37.29
        ActiveCell.Select
        ActiveCell.Offset(0, 8).Activate
    Loop
    
    Range("A1").Select
    
    'Then rows...
    Do While (ActiveCell.Row < 863)
        ActiveCell.Rows("1:1").EntireRow.RowHeight = 12.75
        ActiveCell.Offset(1, 0).Rows("1:1").EntireRow.RowHeight = 77.25
        ActiveCell.Offset(2, 0).Rows("1:1").EntireRow.RowHeight = 23.25
        ActiveCell.Offset(3, 0).Rows("1:1").EntireRow.RowHeight = 23.25
        ActiveCell.Offset(4, 0).Rows("1:1").EntireRow.RowHeight = 23.25
        ActiveCell.Offset(5, 0).Rows("1:1").EntireRow.RowHeight = 23.25
        ActiveCell.Offset(6, 0).Rows("1:1").EntireRow.RowHeight = 28.5
        ActiveCell.Offset(7, 0).Rows("1:1").EntireRow.RowHeight = 28.5
        ActiveCell.Offset(8, 0).Rows("1:1").EntireRow.RowHeight = 29.25
        ActiveCell.Offset(8, 0).Rows("1:1").EntireRow.RowHeight = 28.5
        ActiveCell.Offset(9, 0).Rows("1:1").EntireRow.RowHeight = 28.5
        ActiveCell.Offset(10, 0).Rows("1:1").EntireRow.RowHeight = 28.5
        ActiveCell.Offset(11, 0).Rows("1:1").EntireRow.RowHeight = 17.25
        ActiveCell.Offset(12, 0).Rows("1:1").EntireRow.RowHeight = 21
        ActiveCell.Offset(13, 0).Rows("1:1").EntireRow.RowHeight = 15.75
        
        'Second label space
        ActiveCell.Offset(15, 0).Activate
        ActiveCell.Rows("1:1").EntireRow.RowHeight = 12.75
        ActiveCell.Offset(1, 0).Rows("1:1").EntireRow.RowHeight = 77.25
        ActiveCell.Offset(2, 0).Rows("1:1").EntireRow.RowHeight = 23.25
        ActiveCell.Offset(3, 0).Rows("1:1").EntireRow.RowHeight = 23.25
        ActiveCell.Offset(4, 0).Rows("1:1").EntireRow.RowHeight = 23.25
        ActiveCell.Offset(5, 0).Rows("1:1").EntireRow.RowHeight = 23.25
        ActiveCell.Offset(6, 0).Rows("1:1").EntireRow.RowHeight = 28.5
        ActiveCell.Offset(7, 0).Rows("1:1").EntireRow.RowHeight = 28.5
        ActiveCell.Offset(8, 0).Rows("1:1").EntireRow.RowHeight = 29.25
        ActiveCell.Offset(8, 0).Rows("1:1").EntireRow.RowHeight = 28.5
        ActiveCell.Offset(9, 0).Rows("1:1").EntireRow.RowHeight = 28.5
        ActiveCell.Offset(10, 0).Rows("1:1").EntireRow.RowHeight = 28.5
        ActiveCell.Offset(11, 0).Rows("1:1").EntireRow.RowHeight = 17.25
        ActiveCell.Offset(12, 0).Rows("1:1").EntireRow.RowHeight = 21
        ActiveCell.Offset(13, 0).Rows("1:1").EntireRow.RowHeight = 15.75
        
        'Third label space
        ActiveCell.Offset(15, 0).Activate
        ActiveCell.Rows("1:1").EntireRow.RowHeight = 12.75
        ActiveCell.Offset(1, 0).Rows("1:1").EntireRow.RowHeight = 77.25
        ActiveCell.Offset(2, 0).Rows("1:1").EntireRow.RowHeight = 23.25
        ActiveCell.Offset(3, 0).Rows("1:1").EntireRow.RowHeight = 23.25
        ActiveCell.Offset(4, 0).Rows("1:1").EntireRow.RowHeight = 23.25
        ActiveCell.Offset(5, 0).Rows("1:1").EntireRow.RowHeight = 23.25
        ActiveCell.Offset(6, 0).Rows("1:1").EntireRow.RowHeight = 28.5
        ActiveCell.Offset(7, 0).Rows("1:1").EntireRow.RowHeight = 28.5
        ActiveCell.Offset(8, 0).Rows("1:1").EntireRow.RowHeight = 29.25
        ActiveCell.Offset(8, 0).Rows("1:1").EntireRow.RowHeight = 28.5
        ActiveCell.Offset(9, 0).Rows("1:1").EntireRow.RowHeight = 28.5
        ActiveCell.Offset(10, 0).Rows("1:1").EntireRow.RowHeight = 28.5
        ActiveCell.Offset(11, 0).Rows("1:1").EntireRow.RowHeight = 17.25
        ActiveCell.Offset(12, 0).Rows("1:1").EntireRow.RowHeight = 21
        ActiveCell.Offset(13, 0).Rows("1:1").EntireRow.RowHeight = 15.75
        ActiveCell.Offset(-30, 0).Activate
        ActiveCell.Offset(54, 0).Range("A1").Select
        Selection.End(xlToLeft).Select
    Loop
    
    'And ends by placing the active cell at A1, 930 slots for labels
    Range("A1").Select
    
    'Set print settings
    With ActiveSheet.PageSetup
        .LeftHeader = ""
        .CenterHeader = ""
        .RightHeader = ""
        .LeftFooter = ""
        .CenterFooter = ""
        .RightFooter = ""
        .LeftMargin = Application.InchesToPoints(0)
        .RightMargin = Application.InchesToPoints(0)
        .TopMargin = Application.InchesToPoints(0)
        .BottomMargin = Application.InchesToPoints(0)
        .HeaderMargin = Application.InchesToPoints(0)
        .FooterMargin = Application.InchesToPoints(0)
        .PrintHeadings = False
        .PrintGridlines = False
        .PrintComments = xlPrintNoComments
        .PrintQuality = 600
        .CenterHorizontally = True
        .CenterVertically = True
        .Orientation = xlPortrait
        .Draft = False
        .PaperSize = xlPaperLetter
        .FirstPageNumber = xlAutomatic
        .Order = xlOverThenDown
        .BlackAndWhite = False
        .Zoom = 60
    End With
    ActiveWindow.SelectedSheets.PrintPreview
    
    'Supplant Blank
    ActiveWindow.Application.DisplayAlerts = False
    Sheets("Q_Blnk").Select
    ActiveWindow.SelectedSheets.Delete
    Sheets("Qt").Select
    Sheets("Qt").Copy before:=Sheets("1_Blnk")
    Sheets(7).Select
    Sheets(7).Name = "Q_Blnk"
    Sheets("Qt").Select
    ActiveWindow.Application.DisplayAlerts = True
End Sub

Public Sub Reset1Gal_Hard()
    'Remove and reset
    ActiveWindow.Application.DisplayAlerts = False
    Sheets("1 Gal").Select
    ActiveWindow.SelectedSheets.Delete
    Sheets("5 Gal").Select
    Sheets.Add
    Sheets(2).Select
    Sheets(2).Name = "1 Gal"
    ActiveWindow.Application.DisplayAlerts = True
    
    Inserted_1GA = False
    
    'Set cell dimensions
    'Columns first...
    Range("A1").Select
    Do While (ActiveCell.Column < 253)
        ActiveCell.Columns("A:A").EntireColumn.ColumnWidth = 19.86
        ActiveCell.Offset(0, 1).Columns("A:A").EntireColumn.ColumnWidth = 16.71
        ActiveCell.Offset(0, 1).Columns("A:A").EntireColumn.ColumnWidth = 16.43
        ActiveCell.Offset(0, 1).Columns("A:A").EntireColumn.ColumnWidth = 16.57
        ActiveCell.Offset(0, 2).Columns("A:A").EntireColumn.ColumnWidth = 40.43
        ActiveCell.Offset(0, 3).Columns("A:A").EntireColumn.ColumnWidth = 37.29
        ActiveCell.Select
        ActiveCell.Offset(0, 4).Activate
    Loop
    
    Range("A1").Select
    
    'Then rows...
    Do While (ActiveCell.Row < 381)
        ActiveCell.Rows("1:1").EntireRow.RowHeight = 12.75
        ActiveCell.Offset(1, 0).Rows("1:1").EntireRow.RowHeight = 77.25
        ActiveCell.Offset(2, 0).Rows("1:1").EntireRow.RowHeight = 23.25
        ActiveCell.Offset(3, 0).Rows("1:1").EntireRow.RowHeight = 23.25
        ActiveCell.Offset(4, 0).Rows("1:1").EntireRow.RowHeight = 23.25
        ActiveCell.Offset(5, 0).Rows("1:1").EntireRow.RowHeight = 23.25
        ActiveCell.Offset(6, 0).Rows("1:1").EntireRow.RowHeight = 28.5
        ActiveCell.Offset(7, 0).Rows("1:1").EntireRow.RowHeight = 28.5
        ActiveCell.Offset(8, 0).Rows("1:1").EntireRow.RowHeight = 29.25
        ActiveCell.Offset(8, 0).Rows("1:1").EntireRow.RowHeight = 28.5
        ActiveCell.Offset(9, 0).Rows("1:1").EntireRow.RowHeight = 28.5
        ActiveCell.Offset(10, 0).Rows("1:1").EntireRow.RowHeight = 28.5
        ActiveCell.Offset(11, 0).Rows("1:1").EntireRow.RowHeight = 17.25
        ActiveCell.Offset(12, 0).Rows("1:1").EntireRow.RowHeight = 21
        ActiveCell.Offset(13, 0).Rows("1:1").EntireRow.RowHeight = 15.75
        
        'Second label space
        ActiveCell.Offset(15, 0).Range("A1").Activate
        ActiveCell.Rows("1:1").EntireRow.RowHeight = 12.75
        ActiveCell.Offset(1, 0).Rows("1:1").EntireRow.RowHeight = 77.25
        ActiveCell.Offset(2, 0).Rows("1:1").EntireRow.RowHeight = 23.25
        ActiveCell.Offset(3, 0).Rows("1:1").EntireRow.RowHeight = 23.25
        ActiveCell.Offset(4, 0).Rows("1:1").EntireRow.RowHeight = 23.25
        ActiveCell.Offset(5, 0).Rows("1:1").EntireRow.RowHeight = 23.25
        ActiveCell.Offset(6, 0).Rows("1:1").EntireRow.RowHeight = 28.5
        ActiveCell.Offset(7, 0).Rows("1:1").EntireRow.RowHeight = 28.5
        ActiveCell.Offset(8, 0).Rows("1:1").EntireRow.RowHeight = 29.25
        ActiveCell.Offset(8, 0).Rows("1:1").EntireRow.RowHeight = 28.5
        ActiveCell.Offset(9, 0).Rows("1:1").EntireRow.RowHeight = 28.5
        ActiveCell.Offset(10, 0).Rows("1:1").EntireRow.RowHeight = 28.5
        ActiveCell.Offset(11, 0).Rows("1:1").EntireRow.RowHeight = 17.25
        ActiveCell.Offset(12, 0).Rows("1:1").EntireRow.RowHeight = 21
        ActiveCell.Offset(13, 0).Rows("1:1").EntireRow.RowHeight = 15.75
        ActiveCell.Offset(-15, 0).Range("A1").Activate
        ActiveCell.Offset(38, 0).Range("A1").Select
        Selection.End(xlToLeft).Select
    Loop
    
    'And ends by placing the active cell at A1, 1,260 slots for labels
    Range("A1").Select
    
    'Set print settings
    With ActiveSheet.PageSetup
        .PrintTitleRows = ""
        .PrintTitleColumns = ""
    End With
    ActiveSheet.PageSetup.PrintArea = ""
    With ActiveSheet.PageSetup
        .LeftHeader = ""
        .CenterHeader = ""
        .RightHeader = ""
        .LeftFooter = ""
        .CenterFooter = ""
        .RightFooter = ""
        .LeftMargin = Application.InchesToPoints(0)
        .RightMargin = Application.InchesToPoints(0)
        .TopMargin = Application.InchesToPoints(0)
        .BottomMargin = Application.InchesToPoints(0)
        .HeaderMargin = Application.InchesToPoints(0)
        .FooterMargin = Application.InchesToPoints(0)
        .PrintHeadings = False
        .PrintGridlines = False
        .PrintComments = xlPrintNoComments
        .PrintQuality = 600
        .CenterHorizontally = True
        .CenterVertically = True
        .Orientation = xlPortrait
        .Draft = False
        .PaperSize = xlPaperLetter
        .FirstPageNumber = xlAutomatic
        .Order = xlOverThenDown
        .BlackAndWhite = False
        .Zoom = 88
    End With
    ActiveWindow.SelectedSheets.PrintPreview
    
    'Supplant Blank
    ActiveWindow.Application.DisplayAlerts = False
    Sheets("1_Blnk").Select
    ActiveWindow.SelectedSheets.Delete
    Sheets("1 Gal").Select
    Sheets("1 Gal").Copy before:=Sheets("5_Blnk")
    Sheets(8).Select
    Sheets(8).Name = "1_Blnk"
    Sheets("1 Gal").Select
    ActiveWindow.Application.DisplayAlerts = True
End Sub

Public Sub Reset5Gal_Hard()
'
' Reset5Gal Macro
' Macro recorded 6/2/2014 by Doug Saylor
'
    'Remove and reset
    ActiveWindow.Application.DisplayAlerts = False
    Sheets("5 Gal").Select
    Range("A1").Select
    ActiveWindow.SelectedSheets.Delete
    Sheets("Q_Res").Select
    Sheets.Add
    Sheets(3).Select
    Sheets(3).Name = "5 Gal"
    ActiveWindow.Application.DisplayAlerts = True
    
    Inserted_5GA = False
    
    'Set cell dimensions
    'Columns first...
    Range("A1").Select
    Do While (ActiveCell.Column < 253)
        ActiveCell.Columns("A:A").EntireColumn.ColumnWidth = 19.86
        ActiveCell.Offset(0, 1).Columns("A:A").EntireColumn.ColumnWidth = 16.71
        ActiveCell.Offset(0, 1).Columns("A:A").EntireColumn.ColumnWidth = 16.43
        ActiveCell.Offset(0, 1).Columns("A:A").EntireColumn.ColumnWidth = 16.57
        ActiveCell.Offset(0, 2).Columns("A:A").EntireColumn.ColumnWidth = 40.43
        ActiveCell.Offset(0, 3).Columns("A:A").EntireColumn.ColumnWidth = 37.29
        ActiveCell.Select
        ActiveCell.Offset(0, 4).Activate
    Loop
    
    Range("A1").Select
    
    'Then rows...
    Do While (ActiveCell.Row < 481)
        ActiveCell.Rows("1:1").EntireRow.RowHeight = 12.75
        ActiveCell.Offset(1, 0).Rows("1:1").EntireRow.RowHeight = 77.25
        ActiveCell.Offset(2, 0).Rows("1:1").EntireRow.RowHeight = 23.25
        ActiveCell.Offset(3, 0).Rows("1:1").EntireRow.RowHeight = 23.25
        ActiveCell.Offset(4, 0).Rows("1:1").EntireRow.RowHeight = 23.25
        ActiveCell.Offset(5, 0).Rows("1:1").EntireRow.RowHeight = 23.25
        ActiveCell.Offset(6, 0).Rows("1:1").EntireRow.RowHeight = 28.5
        ActiveCell.Offset(7, 0).Rows("1:1").EntireRow.RowHeight = 28.5
        ActiveCell.Offset(8, 0).Rows("1:1").EntireRow.RowHeight = 29.25
        ActiveCell.Offset(8, 0).Rows("1:1").EntireRow.RowHeight = 28.5
        ActiveCell.Offset(9, 0).Rows("1:1").EntireRow.RowHeight = 28.5
        ActiveCell.Offset(10, 0).Rows("1:1").EntireRow.RowHeight = 28.5
        ActiveCell.Offset(11, 0).Rows("1:1").EntireRow.RowHeight = 17.25
        ActiveCell.Offset(12, 0).Rows("1:1").EntireRow.RowHeight = 21
        ActiveCell.Offset(13, 0).Rows("1:1").EntireRow.RowHeight = 15.75
        ActiveCell.Offset(24, 0).Range("A1").Select
        Selection.End(xlToLeft).Select
    Loop
    
    'And ends by placing the active cell at A1, 1,260 slots for labels
    Range("A1").Select
    
    'Set print settings
    With ActiveSheet.PageSetup
        .PrintTitleRows = ""
        .PrintTitleColumns = ""
    End With
    ActiveSheet.PageSetup.PrintArea = ""
    With ActiveSheet.PageSetup
        .LeftHeader = ""
        .CenterHeader = ""
        .RightHeader = ""
        .LeftFooter = ""
        .CenterFooter = ""
        .RightFooter = ""
        .LeftMargin = Application.InchesToPoints(0)
        .RightMargin = Application.InchesToPoints(0)
        .TopMargin = Application.InchesToPoints(0)
        .BottomMargin = Application.InchesToPoints(0)
        .HeaderMargin = Application.InchesToPoints(0)
        .FooterMargin = Application.InchesToPoints(0)
        .PrintHeadings = False
        .PrintGridlines = False
        .PrintComments = xlPrintNoComments
        .PrintQuality = 600
        .CenterHorizontally = True
        .CenterVertically = True
        .Orientation = xlLandscape
        .Draft = False
        .PaperSize = xlPaperLetter
        .FirstPageNumber = xlAutomatic
        .Order = xlOverThenDown
        .BlackAndWhite = False
        .Zoom = 117
    End With
    ActiveWindow.SelectedSheets.PrintPreview
    
    'Supplant Blank
    ActiveWindow.Application.DisplayAlerts = False
    Sheets("5_Blnk").Select
    ActiveWindow.SelectedSheets.Delete
    Sheets("5 Gal").Select
    Sheets("5 Gal").Copy before:=Sheets("FrontPage")
    Sheets(9).Select
    Sheets(9).Name = "5_Blnk"
    Sheets("5 Gal").Select
    ActiveWindow.Application.DisplayAlerts = True
End Sub

'+---------------------+
'| Debugging Functions |
'+---------------------+

Public Sub showAll()
    Sheets(1).Visible = True
    Sheets(2).Visible = True
    Sheets(3).Visible = True
    Sheets(4).Visible = True
    Sheets(5).Visible = True
    Sheets(6).Visible = True
    Sheets(7).Visible = True
    Sheets(8).Visible = True
    Sheets(9).Visible = True
End Sub

Public Sub hideAll()
    Sheets(1).Visible = xlVeryHidden
    Sheets(2).Visible = xlVeryHidden
    Sheets(3).Visible = xlVeryHidden
    Sheets(4).Visible = xlVeryHidden
    Sheets(5).Visible = xlVeryHidden
    Sheets(6).Visible = xlVeryHidden
    Sheets(7).Visible = xlVeryHidden
    Sheets(8).Visible = xlVeryHidden
    Sheets(9).Visible = xlVeryHidden
End Sub

Public Sub getColumn()
    'Does just what it says: get's the column number (long) of the active cell
    MsgBox "Current column is: " & ActiveCell.Column
End Sub
