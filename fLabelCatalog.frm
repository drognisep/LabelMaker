VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} fLabelCatalog 
   Caption         =   "Label Catalog"
   ClientHeight    =   4425
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5265
   OleObjectBlob   =   "fLabelCatalog.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "fLabelCatalog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Base 0
Option Explicit

Dim A_Bnum As String
Dim A_DMan As String
Dim B_Bnum As String
Dim B_DMan As String
Dim Temp_BNum As String
Dim OverrideSize As Boolean

Private Function InsertLabel(CONNUM As String) As Integer
    Dim thisSheet As Worksheet
    Dim srcSheet As Worksheet
    Dim SO As String
    Dim PO As String
    Dim BNUM As String
    Dim DMAN As String
    Dim DEXP As String
    
    On Error GoTo ReturnError
    
    SO = tbSO.Text
    PO = tbPO.Text
    BNUM = tbBNUM.Text
    DMAN = tbDMAN.Text
    If (Exp_Period = Exp_6Mon) Then
        DEXP = Add6Mon(DMAN)
    ElseIf (Exp_Period = Exp_12Mon) Then
        DEXP = Add12Mon(DMAN)
    Else
        InsertLabel = -1
        Exit Function
    End If
    
    Set thisSheet = Application.ActiveSheet
    If (thisSheet.Visible <> xlSheetVisible) Then
        thisSheet.Visible = xlSheetVisible
    End If
    
    Select Case Label_Selected
        Case Label_Silver
            Sheets("5_Res").Visible = xlSheetVisible
            Set srcSheet = Sheets("5_Res")
            srcSheet.Visible = xlSheetVisible
            srcSheet.Range("A1:D14").Copy
        Case Label_White
            Sheets("5_Res").Visible = xlSheetVisible
            Set srcSheet = Sheets("5_Res")
            srcSheet.Visible = xlSheetVisible
            srcSheet.Range("E1:H14").Copy
        Case Label_1336A
            Sheets("5_Res").Visible = xlSheetVisible
            Set srcSheet = Sheets("5_Res")
            srcSheet.Visible = xlSheetVisible
            srcSheet.Range("I1:L14").Copy
        Case Label_1336B
            Sheets("5_Res").Visible = xlSheetVisible
            Set srcSheet = Sheets("5_Res")
            srcSheet.Visible = xlSheetVisible
            srcSheet.Range("M1:P14").Copy
        Case Label_Kit
            Sheets("5_Res").Visible = xlSheetVisible
            Set srcSheet = Sheets("5_Res")
            srcSheet.Visible = xlSheetVisible
            srcSheet.Range("Q1:T14").Copy
        Case Else
            InsertLabel = -1
            Exit Function
    End Select
    thisSheet.Select
    ActiveSheet.Paste
    Selection.Range("B3").Select
    ActiveCell.FormulaR1C1 = SO
    ActiveCell.Offset(1, 0).FormulaR1C1 = PO
    ActiveCell.Offset(4, 0).FormulaR1C1 = CONNUM
    ActiveCell.Offset(5, 0).FormulaR1C1 = BNUM
    ActiveCell.Offset(7, 0).FormulaR1C1 = DMAN
    ActiveCell.Offset(8, 0).FormulaR1C1 = DEXP
    ActiveCell.Offset(-1, -1).Activate
    Application.CutCopyMode = False
    srcSheet.Visible = xlSheetVeryHidden
    
    On Error GoTo 0
    Exit Function
ReturnError:
    InsertLabel = -1
End Function

Private Sub Reset5Gal()
    showAll
    ActiveWindow.Application.DisplayAlerts = False
    Sheets("5 Gal").Select
    ActiveWindow.SelectedSheets.Delete
    Sheets("5_Blnk").Select
    Sheets("5_Blnk").Copy before:=Sheets(3)
    Sheets(3).Select
    Sheets(3).Name = "5 Gal"
    Range("A1").Activate
    Sheets(3).Protect "3141592654", True, True, True, True
    Sheets(3).EnableSelection = xlUnlockedCells
    Inserted_5GA = False
    hideAll
    ActiveWindow.Application.DisplayAlerts = True
End Sub
Private Sub Reset1Gal()
    showAll
    ActiveWindow.Application.DisplayAlerts = False
    Sheets("1 Gal").Select
    ActiveWindow.SelectedSheets.Delete
    Sheets("1_Blnk").Select
    Sheets("1_Blnk").Copy before:=Sheets(2)
    Sheets(2).Select
    Sheets(2).Name = "1 Gal"
    Range("A1").Activate
    Sheets(2).Protect "3141592654", True, True, True, True
    Sheets(2).EnableSelection = xlUnlockedCells
    Inserted_1GA = False
    hideAll
    ActiveWindow.Application.DisplayAlerts = True
End Sub
Private Sub ResetQt()
    showAll
    ActiveWindow.Application.DisplayAlerts = False
    Sheets("Qt").Select
    ActiveWindow.SelectedSheets.Delete
    Sheets("Q_Blnk").Select
    Sheets("Q_Blnk").Copy before:=Sheets(1)
    Sheets(1).Select
    Sheets(1).Name = "Qt"
    Range("A1").Activate
    Sheets(1).Protect "3141592654", True, True, True, True
    Sheets(1).EnableSelection = xlUnlockedCells
    Inserted_QT = False
    hideAll
    ActiveWindow.Application.DisplayAlerts = True
End Sub

Private Function Add6Mon(S As String) As String
    'Need to be one day short of the month length, I haven't figured out why yet
    'Maybe it has something to do with date literals vs. stored date values?
    Dim MonthArray(14) As Date
    MonthArray(1) = #1/30/1900#
    MonthArray(2) = #1/27/1900#
    MonthArray(3) = #1/30/1900#
    MonthArray(4) = #1/29/1900#
    MonthArray(5) = #1/30/1900#
    MonthArray(6) = #1/29/1900#
    MonthArray(7) = #1/30/1900#
    MonthArray(8) = #1/30/1900#
    MonthArray(9) = #1/29/1900#
    MonthArray(10) = #1/30/1900#
    MonthArray(11) = #1/29/1900#
    MonthArray(12) = #1/30/1900#
    
    Dim i, counter As Integer
    Dim d As Date
    
    On Error GoTo ErrorHandler
    d = DateValue(S)
    On Error GoTo 0
    
    i = Month(d)
    counter = 0
    
    Do While counter < 6
        d = d + MonthArray(i)
        
        i = i + 1
        If (i > 12) Then
            i = 1
        End If
        counter = counter + 1
    Loop
    
    Add6Mon = CStr(d)
    Exit Function
    
ErrorHandler:
    Add6Mon = "[BADFORM]"
End Function

Private Function Add12Mon(S As String) As String
    'Need to be one day short of the month length, I haven't figured out why yet
    'Maybe it has something to do with date literals vs. stored date values?
    Dim MonthArray(14) As Date
    MonthArray(1) = #1/30/1900#
    MonthArray(2) = #1/27/1900#
    MonthArray(3) = #1/30/1900#
    MonthArray(4) = #1/29/1900#
    MonthArray(5) = #1/30/1900#
    MonthArray(6) = #1/29/1900#
    MonthArray(7) = #1/30/1900#
    MonthArray(8) = #1/30/1900#
    MonthArray(9) = #1/29/1900#
    MonthArray(10) = #1/30/1900#
    MonthArray(11) = #1/29/1900#
    MonthArray(12) = #1/30/1900#
    
    Dim i, counter As Integer
    Dim d As Date
    
    On Error GoTo ErrorHandler
    d = DateValue(S)
    On Error GoTo 0
    
    i = Month(d)
    counter = 0
    
    Do While counter < 12
        d = d + MonthArray(i)
        
        i = i + 1
        If (i > 12) Then
            i = 1
        End If
        counter = counter + 1
    Loop
    
    Add12Mon = CStr(d)
    Exit Function
    
ErrorHandler:
    Add12Mon = "[BADFORM]"
End Function

Private Sub cbDone_Click()
    Me.Caption = "Label Catalog"
    If (tbStartCONNUM.Text = "SHOWALL") Then
        showAll
        tbStartCONNUM.Text = ""
    ElseIf (tbStartCONNUM.Text = "HARDRESET") Then
        fHardReset.Show
        tbStartCONNUM.Text = ""
    ElseIf (tbStartCONNUM.Text = "HIDEALL") Then
        hideAll
        tbStartCONNUM.Text = ""
    ElseIf (tbStartCONNUM.Text = "RESET") Then
        showAll
        Reset5Gal
        Inserted_5GA = False
        Reset1Gal
        Inserted_1GA = False
        ResetQt
        Inserted_QT = False
        hideAll
        tbStartCONNUM.Text = ""
    ElseIf (tbStartCONNUM.Text = "DEBUG") Then
        Me.Hide
        fDebugConsole.Show (vbModeless)
        tbStartCONNUM.Text = ""
    ElseIf (tbStartCONNUM.Text = "KILLDIALOG") Then
        Unload Me
    Else
        allowScroll
        fLabelCatalog.Hide
        fPlaceholder.Show (vbModeless)
    End If
End Sub

Private Sub cbInsertLabels_Click()
    Dim retval As Integer, amount As Integer
    Dim DEXP As String, temp As String
    
    If (Label_Selected >= Label_None And Label_Selected <= Label_Max) Then
        If (Label_Selected = Label_None) Then
            MsgBox "Select a label type", vbOKOnly, "Incomplete Details"
            Exit Sub
        ElseIf (tbSO.Text = "") Then
            MsgBox "Insert Sales Order Number", vbOKOnly, "Incomplete Details"
            Exit Sub
        ElseIf (tbBNUM.Text = "") Then
            If (Label_Selected = Label_Kit) Then
                tbBNUM.Text = "N/A"
            Else
                MsgBox "Insert Batch Number", vbOKOnly, "Incomplete Details"
                Exit Sub
            End If
        ElseIf (tbPO.Text = "") Then
            MsgBox "Insert Purchase Order Number", vbOKOnly, "Incomplete Details"
            Exit Sub
        ElseIf (tbStartCONNUM.Text = "") Then
            MsgBox "Insert starting container Number", vbOKOnly, "Incomplete Details"
            Exit Sub
        ElseIf (tbAmount.Value = 0 Or tbAmount.Text = "" Or tbAmount.Value < 0) Then
            MsgBox "Insert label amount", vbOKOnly, "Incomplete Details"
            Exit Sub
        ElseIf (tbDMAN.Value = "") Then
            MsgBox "Insert date of manufacture", vbOKOnly, "Incomplete Details"
            Exit Sub
        ElseIf (tbDMAN.TextLength <> 8) Then
            MsgBox "Date must be 8 characters in length (mm/dd/yy)", vbOKOnly, "Incomplete Details"
            Exit Sub
        ElseIf (Exp_Period = Exp_None) Then
            MsgBox "Choose expiration period", vbOKOnly, "Incomplete Details"
            Exit Sub
        Else
            If (Exp_Period = Exp_6Mon) Then
                DEXP = Add6Mon(tbDMAN.Text)
            ElseIf (Exp_Period = Exp_12Mon) Then
                DEXP = Add12Mon(tbDMAN.Text)
            Else
                MsgBox "Invalid Exp_Period: " & Exp_Period & vbNewLine & "SHOW TO DOUG", , "Internal error"
                fDebugConsole.Show
                Exit Sub
            End If
            If (DEXP = "[BADFORM]") Then
                MsgBox "Invalid date format", vbOKOnly, "Try Again"
                Exit Sub
            End If
            If (Label_Selected <> Label_Kit Or OverrideSize = False) Then
                fSizePicker.Show
                If (Label_Size_Selected = Label_Size_None) Then
                    Exit Sub
                End If
            End If
        End If
            
        amount = Int(tbAmount.Value)
        temp = tbStartCONNUM.Text
        'Enhanced select case for Label_Selected
        If (Label_Selected <> Label_Kit) Then
            Select Case Label_Size_Selected
                Case Label_Size_5GA
                    Sheets("5 Gal").Visible = True
                    Sheets("5 Gal").Select
                    fLabelCatalog.Caption = "Label Catalog - Inserting labels, please wait..."
                    Do While (amount > 0)
                        If (ActiveCell.Column >= 253) Then
                            'Progress to beginning of next row
                            ActiveCell.Offset(24, 0).Range("A1").Select
                            Selection.End(xlToLeft).Select
                        End If
                        retval = InsertLabel(temp)
                        If (retval = -1) Then
                            MsgBox "Error inserting label" & vbNewLine & "SHOW THIS TO DOUG", vbOKOnly, "Error Occured"
                            fDebugConsole.Show
                            fLabelCatalog.Caption = "Label Catalog - Error occured"
                            Exit Sub
                        Else
                            'Progress to next 5GA slot
                            Inserted_5GA = True
                            ActiveCell.Offset(0, 3).Activate
                        End If
                        amount = amount - 1
                        temp = CStr(Val(temp) + 1)
                    Loop
                    fLabelCatalog.Caption = "Label Catalog - Ready to print!"
                Case Label_Size_1GA
                    fLabelCatalog.Caption = "Label Catalog - Inserting labels, please wait..."
                    Sheets("1 Gal").Visible = True
                    Sheets("1 Gal").Activate
                    Do While (amount > 0)
                        If (ActiveCell.Column >= 253) Then
                            ActiveCell.Offset(38, 0).Activate
                            Selection.End(xlToLeft).Select
                        End If
                        retval = InsertLabel(temp)
                        If (retval = -1) Then
                            MsgBox "Error inserting label" & vbNewLine & "SHOW THIS TO DOUG", vbOKOnly, "Error Occured"
                            fDebugConsole.Show
                            fLabelCatalog.Caption = "Label Catalog - Error occured"
                            Exit Sub
                        Else
                            Inserted_1GA = True
                            amount = amount - 1
                            temp = CStr(Val(temp) + 1)
                        End If
                        If (amount > 0) Then
                            ActiveCell.Offset(14, 0).Activate
                            retval = InsertLabel(temp)
                            If (retval = -1) Then
                                Exit Sub
                            Else
                                ActiveCell.Offset(-15, 3).Activate
                            End If
                            amount = amount - 1
                            temp = CStr(Val(temp) + 1)
                        Else
                            ActiveCell.Offset(0, 3).Activate
                        End If
                    Loop
                    fLabelCatalog.Caption = "Label Catalog - Ready to print!"
                Case Label_Size_QT
                    Sheets("Qt").Visible = True
                    Sheets("Qt").Select
                    Do While (amount > 0)
                        If (ActiveCell.Column >= 251) Then
                            ActiveCell.Offset(54, 0).Activate
                            Selection.End(xlToLeft).Select
                        End If
                        retval = InsertLabel(temp)
                        If (retval = -1) Then
                            MsgBox "Error inserting label" & vbNewLine & "SHOW THIS TO DOUG", vbOKOnly, "Error Occured"
                            fDebugConsole.Show
                            fLabelCatalog.Caption = "Label Catalog - Error occured"
                            Exit Sub
                        Else
                            Inserted_QT = True
                            amount = amount - 1
                            temp = CStr(Val(temp) + 1)
                        End If
                        If (amount > 0) Then
                            ActiveCell.Offset(14, 0).Activate
                            retval = InsertLabel(temp)
                            If (retval = -1) Then
                                Exit Sub
                            End If
                            amount = amount - 1
                            temp = CStr(Val(temp) + 1)
                            If (amount > 0) Then
                                ActiveCell.Offset(14, 0).Activate
                                retval = InsertLabel(temp)
                                If (retval = -1) Then
                                    Exit Sub
                                End If
                                amount = amount - 1
                                temp = CStr(Val(temp) + 1)
                                ActiveCell.Offset(-30, 7).Activate
                            Else
                                ActiveCell.Offset(-15, 7).Activate
                            End If
                        Else
                            ActiveCell.Offset(0, 7).Activate
                        End If
                    Loop
                    fLabelCatalog.Caption = "Label Catalog - Ready to print!"
                Case Else
                    MsgBox "Unsupported label size used: " & Label_Size_Selected & _
                    vbNewLine & "SHOW THIS TO DOUG!!!", vbOKOnly, "Internal Error"
                    fDebugConsole.Show
                    Exit Sub
            End Select
        Else
            If (tbBNUM.Text = "") Then
                tbBNUM.Text = A_Bnum & "/" & B_Bnum
            End If
            Sheets("5 Gal").Visible = True
            Sheets("5 Gal").Select
            fLabelCatalog.Caption = "Label Catalog - Inserting labels, please wait..."
            Do While (amount > 0)
                If (ActiveCell.Column >= 253) Then
                    'Progress to beginning of next row
                    ActiveCell.Offset(24, 0).Range("A1").Select
                    Selection.End(xlToLeft).Select
                End If
                retval = InsertLabel(temp)
                If (retval = -1) Then
                    MsgBox "Error inserting label" & vbNewLine & "SHOW THIS TO DOUG", vbOKOnly, "Error Occured"
                    fDebugConsole.Show
                    fLabelCatalog.Caption = "Label Catalog - Error occured"
                    Exit Sub
                Else
                    'Progress to next 5GA slot
                    Inserted_5GA = True
                    ActiveCell.Offset(0, 3).Activate
                End If
                amount = amount - 1
                temp = CStr(Val(temp) + 1)
            Loop
            
            'Call recursively for 1336A & B
            tbBNUM.Text = A_Bnum
            tbDMAN.Text = A_DMan
            tbStartCONNUM.Text = temp
            Label_Selected = Label_1336A
            Label_Size_Selected = Label_Size_1GA
            
            cbInsertLabels_Click
            
            tbBNUM.Text = B_Bnum
            tbDMAN.Text = B_DMan
            temp = CStr(Val(temp) + tbAmount.Value)
            tbStartCONNUM.Text = temp
            Label_Selected = Label_1336B
            
            cbInsertLabels_Click
            
            OverrideSize = False
            obKits.Value = False
            Label_Selected = Label_None
            
            fLabelCatalog.Caption = "Label Catalog - Ready to print!"
        End If
    Else
        MsgBox "Label selected out of bounds!!!" & vbNewLine & "Internal error, SHOW THIS TO DOUG", vbOKOnly, "Program error"
        fDebugConsole.Show
        Exit Sub
    End If
End Sub

Private Sub cbReset_Click()
    fLabelCatalog.Caption = "Label Catalog - Resetting file, please wait..."
    
    'Reset all sheets and save
    
    showAll
    If (Inserted_5GA = True) Then
        Reset5Gal
    End If
    If (Inserted_1GA = True) Then
        Reset1Gal
    End If
    If (Inserted_QT = True) Then
        ResetQt
    End If
    'hide everything.
    hideAll
    
    'Save and quit
    ActiveWorkbook.Save
    GoodToQuit = True
    Application.Quit
End Sub

Private Sub ob1336A_Click()
    Label_Selected = Label_1336A
End Sub

Private Sub ob1336B_Click()
    Label_Selected = Label_1336B
End Sub

Private Sub obExp1Year_Click()
    Exp_Period = Exp_12Mon
End Sub

Private Sub obExp6Mon_Click()
    Exp_Period = Exp_6Mon
End Sub

Private Sub obKits_Click()
    Dim temp As String
    Dim dtemp As String
    
    temp = InputBox("Enter 1336A Batch Number", "Kit Info")
    If (temp = "") Then
        obKits.Value = False
        GoTo Canceled
    Else
        A_Bnum = temp
    End If
    temp = InputBox("Enter 1336A Manufacture Date", "Kit Info")
    If (Len(temp) <> 8) Then
        If (temp = "") Then
            obKits.Value = False
            GoTo Canceled
        End If
        MsgBox "Date length must be 8 characters (including '/')", vbOKOnly, "Format error"
        obKits.Value = False
        GoTo Canceled
    Else
        dtemp = Add6Mon(temp)
        If (dtemp = "[BADFORM]") Then
            obKits.Value = False
            GoTo Canceled
        Else
            A_DMan = temp
        End If
    End If
    
    
    temp = InputBox("Enter 1336B Batch Number", "Kit Info")
    If (temp = "") Then
        obKits.Value = False
        GoTo Canceled
    Else
        B_Bnum = temp
    End If
    temp = InputBox("Enter 1336B Manufacture Date", "Kit Info")
    If (Len(temp) <> 8) Then
        If (temp = "") Then
            obKits.Value = False
            GoTo Canceled
        End If
        MsgBox "Date length must be 8 characters (including '/')", vbOKOnly, "Format error"
        obKits.Value = False
        GoTo Canceled
    Else
        dtemp = Add6Mon(temp)
        If (dtemp = "[BADFORM]") Then
            obKits.Value = False
            GoTo Canceled
        Else
            B_DMan = temp
        End If
    End If
    
    Label_Selected = Label_Kit
    OverrideSize = True
    Exit Sub
    
Canceled:
    If (Label_Selected <> Label_None) Then
        Select Case Label_Selected
            Case Label_White
                obWhite.Value = True
            Case Label_Silver
                obSilver.Value = True
            Case Label_1336A
                ob1336A.Value = True
            Case Label_1336B
                ob1336B.Value = True
        End Select
    End If
End Sub

Private Sub obSilver_Click()
    Label_Selected = Label_Silver
End Sub

Private Sub obWhite_Click()
    Label_Selected = Label_White
End Sub

Private Sub UserForm_Initialize()
    If (White_Enabled = True) Then
        obWhite.Enabled = True
    End If
    If (Silver_Enabled = True) Then
        obSilver.Enabled = True
    End If
    If (Kit_Enabled = True) Then
        obKits.Enabled = True
    End If
    If (MCP1336A_Enabled = True) Then
        ob1336A.Enabled = True
    End If
    If (MCP1336B_Enabled = True) Then
        ob1336B.Enabled = True
    End If
    OverrideSize = True
    tbSO.Text = ""
    tbBNUM.Text = ""
    tbPO.Text = ""
    tbStartCONNUM.Text = ""
    tbAmount.Text = ""
    tbDMAN.Text = ""
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    'Prevent user from closing with the Close box in the title bar.
    If CloseMode <> 1 Then Cancel = 1
End Sub

