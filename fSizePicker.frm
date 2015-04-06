VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} fSizePicker 
   Caption         =   "Pick Label Size"
   ClientHeight    =   2385
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   1590
   OleObjectBlob   =   "fSizePicker.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "fSizePicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 0

Private Sub cb1Gal_Click()
    Label_Size_Selected = Label_Size_1GA
    Unload Me
End Sub

Private Sub cb5Gal_Click()
    Label_Size_Selected = Label_Size_5GA
    Unload Me
End Sub

Private Sub cbQt_Click()
    Label_Size_Selected = Label_Size_QT
    Unload Me
End Sub

Private Sub cbSizeCancel_Click()
    Label_Size_Selected = Label_Size_None
    Unload Me
End Sub

Private Sub UserForm_Activate()
    Select Case Label_Selected
        Case Label_White
            If (White_Enabled = True) Then
                If (White_5GA_Enabled = True) Then
                    cb5Gal.Enabled = True
                End If
                If (White_1GA_Enabled = True) Then
                    cb1Gal.Enabled = True
                End If
                If (White_QT_Enabled = True) Then
                    cbQt.Enabled = True
                End If
            End If
        Case Label_Silver
            If (Silver_Enabled = True) Then
                If (Silver_5GA_Enabled = True) Then
                    cb5Gal.Enabled = True
                End If
                If (Silver_1GA_Enabled = True) Then
                    cb1Gal.Enabled = True
                End If
                If (Silver_QT_Enabled = True) Then
                    cbQt.Enabled = True
                End If
            End If
        Case Label_Kit
            If (Kit_Enabled = True) Then
                If (Kit_5GA_Enabled = True) Then
                    cb5Gal.Enabled = True
                End If
                If (Kit_1GA_Enabled = True) Then
                    cb1Gal.Enabled = True
                End If
                If (Kit_QT_Enabled = True) Then
                    cbQt.Enabled = True
                End If
            End If
        Case Label_1336A
            If (MCP1336A_Enabled = True) Then
                If (MCP1336A_5GA_Enabled = True) Then
                    cb5Gal.Enabled = True
                End If
                If (MCP1336A_1GA_Enabled = True) Then
                    cb1Gal.Enabled = True
                End If
                If (MCP1336A_QT_Enabled = True) Then
                    cbQt.Enabled = True
                End If
            End If
        Case Label_1336B
            If (MCP1336B_Enabled = True) Then
                If (MCP1336B_5GA_Enabled = True) Then
                    cb5Gal.Enabled = True
                End If
                If (MCP1336B_1GA_Enabled = True) Then
                    cb1Gal.Enabled = True
                End If
                If (MCP1336B_QT_Enabled = True) Then
                    cbQt.Enabled = True
                End If
            End If
        Case Else
            MsgBox "Case else of Label_Selected " & vbNewLine & "SHOW THIS TO DOUG", vbOKOnly, "Internal Error"
    End Select
End Sub
