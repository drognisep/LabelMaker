VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} fPlaceholder 
   Caption         =   "Label Catalog"
   ClientHeight    =   570
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2175
   OleObjectBlob   =   "fPlaceholder.frx":0000
End
Attribute VB_Name = "fPlaceholder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 0

Dim StartingCellAddress As String

Private Sub cbUnhide_Click()
    disallowScroll
    LoadPos
    Unload Me
    fLabelCatalog.Show (vbModal)
End Sub

Private Sub UserForm_Initialize()
    SavePos
    Me.Left = Application.ActiveWindow.Width - Me.Width
    Me.Top = Application.ActiveWindow.Height - Me.Height
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    'Prevent user from closing with the Close box in the title bar.
    If CloseMode <> 1 Then Cancel = 1
End Sub
