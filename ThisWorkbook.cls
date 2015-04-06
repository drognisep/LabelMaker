VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Private Sub Workbook_Open()
    Dim response As Integer
    'On Error GoTo Break
    GoodToQuit = False
    Sheets(1).Protect "3141592654", True, True, True, True
    Sheets(1).EnableSelection = xlUnlockedCells
    Sheets(1).Visible = xlVeryHidden
    Sheets(2).Protect "3141592654", True, True, True, True
    Sheets(2).EnableSelection = xlUnlockedCells
    Sheets(2).Visible = xlVeryHidden
    Sheets(3).Protect "3141592654", True, True, True, True
    Sheets(3).EnableSelection = xlUnlockedCells
    Sheets(3).Visible = xlVeryHidden
    Sheets(4).Protect "3141592654", True, True, True, True
    Sheets(4).EnableSelection = xlUnlockedCells
    Sheets(4).Visible = xlVeryHidden
    Sheets(5).Protect "3141592654", True, True, True, True
    Sheets(5).EnableSelection = xlUnlockedCells
    Sheets(5).Visible = xlVeryHidden
    Sheets(6).Protect "3141592654", True, True, True, True
    Sheets(6).EnableSelection = xlUnlockedCells
    Sheets(6).Visible = xlVeryHidden
    Sheets(7).Protect "3141592654", True, True, True, True
    Sheets(7).EnableSelection = xlUnlockedCells
    Sheets(7).Visible = xlVeryHidden
    Sheets(8).Protect "3141592654", True, True, True, True
    Sheets(8).EnableSelection = xlUnlockedCells
    Sheets(8).Visible = xlVeryHidden
    Sheets(9).Protect "3141592654", True, True, True, True
    Sheets(9).EnableSelection = xlUnlockedCells
    Sheets(9).Visible = xlVeryHidden
    Sheets(10).Select
    Sheets(10).PauseProtection
    Range("A1").Select
    Sheets(10).Protect "3141592654", True, True, True, True
    Sheets(10).EnableSelection = xlUnlockedCells
    Sheets(10).SavePos
    Sheets(11).Select
    Sheets(11).PauseProtection
    Range("D1").Select
    Sheets(11).Protect "3141592654", True, True, True, True
    Sheets(11).EnableSelection = xlUnlockedCells
    Sheets(11).SavePos
    Sheets(10).Select
    disallowScroll
    
    On Error GoTo 0
    fLabelCatalog.Show (vbModal)
    Exit Sub
Break:
    response = MsgBox("An error occured during initialization, exit?", _
    vbApplicationModal + vbExclamation, "Startup Error")
    If (response = vbYes) Then
        Application.Quit
        Exit Sub
    ElseIf (response = vbNo) Then
        On Error GoTo Broken
        fDebugConsole.Show
        On Error GoTo 0
        Exit Sub
    Else
        Exit Sub
        GoTo Broken
    End If
Broken:
    MsgBox "Unable to recover, exiting...", vbOKOnly, "Fatal"
    Application.Quit
End Sub
