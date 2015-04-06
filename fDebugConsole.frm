VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} fDebugConsole 
   Caption         =   "DEBUG WINDOW - CONSOLE"
   ClientHeight    =   6030
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4935
   OleObjectBlob   =   "fDebugConsole.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "fDebugConsole"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 0

Private Sub cbDebugClose1_Click()
    fDebugConsole.Hide
    If (fLabelCatalog.Visible <> True) Then
        fLabelCatalog.Show
    End If
End Sub

Private Sub cbDebugClose2_Click()
    cbDebugClose1_Click
End Sub

Private Sub chA1E_Click()
    chA1E.Value = MCP1336A_1GA_Enabled
End Sub

Private Sub chA5E_Click()
    chA5E.Value = MCP1336A_5GA_Enabled
End Sub

Private Sub chAE_Click()
    chAE.Value = MCP1336A_Enabled
End Sub

Private Sub chAQE_Click()
    chAQE.Value = MCP1336A_QT_Enabled
End Sub

Private Sub chB1E_Click()
    chB1E.Value = MCP1336B_1GA_Enabled
End Sub

Private Sub chB5E_Click()
    chB5E.Value = MCP1336B_5GA_Enabled
End Sub

Private Sub chBE_Click()
    chBE.Value = MCP1336B_Enabled
End Sub

Private Sub chBQE_Click()
    chBQE.Value = MCP1336B_QT_Enabled
End Sub

Private Sub chK1E_Click()
    chK1E.Value = Kit_1GA_Enabled
End Sub

Private Sub chK5E_Click()
    chK5E.Value = Kit_5GA_Enabled
End Sub

Private Sub chKE_Click()
    chKE.Value = Kit_Enabled
End Sub

Private Sub chKQE_Click()
    chKQE.Value = Kit_QT_Enabled
End Sub

Private Sub chS1E_Click()
    chS1E.Value = Silver_1GA_Enabled
End Sub

Private Sub chS5E_Click()
    chS5E.Value = Silver_5GA_Enabled
End Sub

Private Sub chSE_Click()
    chSE.Value = Silver_Enabled
End Sub

Private Sub chSQE_Click()
    chSQE.Value = Silver_QT_Enabled
End Sub

Private Sub chW1E_Click()
    chW1E.Value = White_1GA_Enabled
End Sub

Private Sub chW5E_Click()
    chW5E.Value = White_5GA_Enabled
End Sub

Private Sub chWE_Click()
    chWE.Value = White_Enabled
End Sub

Private Sub chWQE_Click()
    chWQE.Value = White_QT_Enabled
End Sub

Private Sub CommandButton1_Click()
    cbDebugClose1_Click
End Sub

Private Sub CommandButton2_Click()
    fHardReset.Show (vbModeless)
End Sub

Private Sub CommandButton3_Click()
    showAll
    ActiveWindow.Application.DisplayAlerts = False
    
    Sheets("5 Gal").Select
    ActiveWindow.SelectedSheets.Delete
    Sheets("5_Blnk").Select
    Sheets("5_Blnk").Copy before:=Sheets(3)
    Sheets(3).Select
    Sheets(3).Name = "5 Gal"
    Range("A1").Activate
    Inserted_5GA = False
    
    Sheets("1 Gal").Select
    ActiveWindow.SelectedSheets.Delete
    Sheets("1_Blnk").Select
    Sheets("1_Blnk").Copy before:=Sheets(2)
    Sheets(2).Select
    Sheets(2).Name = "1 Gal"
    Range("A1").Activate
    Inserted_1GA = False
    
    Sheets("Qt").Select
    ActiveWindow.SelectedSheets.Delete
    Sheets("Q_Blnk").Select
    Sheets("Q_Blnk").Copy before:=Sheets(1)
    Sheets(1).Select
    Sheets(1).Name = "Qt"
    Range("A1").Activate
    Inserted_QT = False
    
    hideAll
    ActiveWindow.Application.DisplayAlerts = True
End Sub

Private Sub CommandButton4_Click()
    showAll
End Sub

Private Sub CommandButton5_Click()
    hideAll
End Sub

Private Sub CommandButton6_Click()

End Sub

Private Sub CommandButton7_Click()
    cbDebugClose1_Click
End Sub

Private Sub UserForm_Initialize()
    SavePos
    UpdateState
    getConsts
    getEnabled
    tbAppName.Text = AppName
    tbAppVersion.Text = AppVersion
    tbAppMaintName.Text = AppMaintName
    tbAppMaintEmail.Text = AppMaintEmail
End Sub

Private Sub UpdateState()
    lblSelected.Caption = Label_Selected
    lblSizeSelected.Caption = Label_Size_Selected
    lblExpPeriod.Caption = Exp_Period
    lblI5.Caption = Inserted_5GA
    lblI1.Caption = Inserted_1GA
    lblIQ.Caption = Inserted_QT
End Sub

Private Sub getConsts()
    'Label type
    lblCLNone.Caption = "Label_None: " & Label_None
    lblCLSilver.Caption = "Label_Silver: " & Label_Silver
    lblCLWhite.Caption = "Label_White: " & Label_White
    lblCLKit.Caption = "Label_Kit: " & Label_Kit
    lblCL1336A.Caption = "Label_1336A: " & Label_1336A
    lblCL1336B.Caption = "Label_1336B: " & Label_1336B
    lblCLMax.Caption = "Label_Max: " & Label_Max
    'Label size
    lblCLSNone.Caption = "Label_Size_None: " & Label_Size_None
    lblCLS5.Caption = "Label_Size_5GA: " & Label_Size_5GA
    lblCLS1.Caption = "Label_Size_1GA: " & Label_Size_1GA
    lblCLSQ.Caption = "Label_Size_QT: " & Label_Size_QT
    lblCLSMax.Caption = "Label_Size_Max: " & Label_Size_Max
    'Expiration type
    lblCX0.Caption = "Exp_None: " & Exp_None
    lblCX6.Caption = "Exp_6Mon: " & Exp_6Mon
    lblCX12.Caption = "Exp_12Mon: " & Exp_12Mon
    lblCXMax.Caption = "Exp_Max: " & Exp_Max
End Sub

Private Sub getEnabled()
    chA1E.Value = MCP1336A_1GA_Enabled
    chA5E.Value = MCP1336A_5GA_Enabled
    chAE.Value = MCP1336A_Enabled
    chAQE.Value = MCP1336A_QT_Enabled
    chB1E.Value = MCP1336B_1GA_Enabled
    chB5E.Value = MCP1336B_5GA_Enabled
    chBE.Value = MCP1336B_Enabled
    chBQE.Value = MCP1336B_QT_Enabled
    chK1E.Value = Kit_1GA_Enabled
    chK5E.Value = Kit_5GA_Enabled
    chKE.Value = Kit_Enabled
    chKQE.Value = Kit_QT_Enabled
    chS1E.Value = Silver_1GA_Enabled
    chS5E.Value = Silver_5GA_Enabled
    chSE.Value = Silver_Enabled
    chSQE.Value = Silver_QT_Enabled
    chW1E.Value = White_1GA_Enabled
    chW5E.Value = White_5GA_Enabled
    chWE.Value = White_Enabled
    chWQE.Value = White_QT_Enabled
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    'Prevent user from closing with the Close box in the title bar.
    If CloseMode <> 1 Then
        Cancel = 1
        LoadPos
        fDebugConsole.Hide
        Unload Me
        fLabelCatalog.Show
    End If
End Sub
