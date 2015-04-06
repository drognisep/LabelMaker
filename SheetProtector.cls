VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SheetProtector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Option Base 0

Private SavedSheet As Worksheet
Private SavedCell As Range
Private isInit As Boolean

Public Sub Reset()
    isInit = False
End Sub

'Public Property Get Str() As String
    'Str = SavedSheet.Name & "/" & SavedCell.Address
'End Property

Public Sub Save()
    
    If (isInit = True) Then
        If (ActiveSheet.Name <> SavedSheet.Name) Then
            MsgBox "Attempted Location Save on different sheet" & _
            vbNewLine & "Try a reset first"
            Exit Sub
        End If
        Set SavedCell = ActiveCell
    Else
        Class_Initialize
    End If
End Sub

Public Property Get Cell() As Range
    Cell = SavedCell
End Property

Public Property Get Sheet() As Worksheet
    Sheet = SavedSheet
End Property

Public Property Get Location() As String
    Location = SavedSheet.Name & "/" & SavedCell.Address
End Property

Public Sub Load()
    SavedCell.Select
End Sub

Private Sub Class_Initialize()
    Set SavedSheet = ActiveSheet
    Set SavedCell = ActiveCell
End Sub
