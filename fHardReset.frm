VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} fHardReset 
   Caption         =   "DEBUG WINDOW - HARD RESET"
   ClientHeight    =   1365
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4710
   OleObjectBlob   =   "fHardReset.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "fHardReset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cbCancelReset_Click()
    Unload Me
End Sub

Private Sub cbConfHardReset_Click()
    fHardReset.Caption = "DEBUG WINDOW - SCORCHING THE EARTH..."
    cbConfHardReset.Enabled = False
    cbCancelReset.Enabled = False
    showAll
    ResetQt_Hard
    Reset1Gal_Hard
    Reset5Gal_Hard
    hideAll
    MsgBox "The Application will now exit", vbOKOnly + vbSystemModal, "Restarting..."
    Application.Quit
    Unload Me
End Sub
