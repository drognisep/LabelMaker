Attribute VB_Name = "Globals"
Option Base 0
Option Explicit

Public Const AppName As String = "Gilby's Label Maker"
Public Const AppVersion As String = "1.0"
Public Const AppMaintName As String = "Doug Saylor"
Public Const AppMaintEmail As String = "josephdsaylor@gmail.com"

'Implemented features
Public Const White_Enabled As Boolean = True
Public Const White_5GA_Enabled As Boolean = True
Public Const White_1GA_Enabled As Boolean = True
Public Const White_QT_Enabled As Boolean = True
Public Const Silver_Enabled As Boolean = True
Public Const Silver_5GA_Enabled As Boolean = True
Public Const Silver_1GA_Enabled As Boolean = True
Public Const Silver_QT_Enabled As Boolean = True
Public Const Kit_Enabled As Boolean = True
Public Const Kit_5GA_Enabled As Boolean = True
Public Const Kit_1GA_Enabled As Boolean = True
Public Const Kit_QT_Enabled As Boolean = True
Public Const MCP1336A_Enabled As Boolean = True
Public Const MCP1336A_5GA_Enabled As Boolean = True
Public Const MCP1336A_1GA_Enabled As Boolean = True
Public Const MCP1336A_QT_Enabled As Boolean = True
Public Const MCP1336B_Enabled As Boolean = True
Public Const MCP1336B_5GA_Enabled As Boolean = True
Public Const MCP1336B_1GA_Enabled As Boolean = True
Public Const MCP1336B_QT_Enabled As Boolean = True

'Label constants, no magic numbers!
Public Const Label_None As Integer = 0
Public Const Label_Silver As Integer = 1
Public Const Label_White As Integer = 2
Public Const Label_Kit As Integer = 3
Public Const Label_1336A As Integer = 4
Public Const Label_1336B As Integer = 5
Public Const Label_Max As Integer = Label_1336B

'Label size constants
Public Const Label_Size_None As Integer = 0
Public Const Label_Size_5GA As Integer = 1
Public Const Label_Size_1GA As Integer = 2
Public Const Label_Size_QT As Integer = 3
Public Const Label_Size_Max As Integer = Label_Size_QT

'Expiration date constants
Public Const Exp_None As Integer = 0
Public Const Exp_6Mon As Integer = 6
Public Const Exp_12Mon As Integer = 12
Public Const Exp_Max As Integer = Exp_12Mon

'Global vars, hold current program state
Public Label_Selected As Integer
Public Label_Size_Selected As Integer
Public Exp_Period As Integer
Public Inserted_5GA As Boolean
Public Inserted_1GA As Boolean
Public Inserted_QT As Boolean
Public GoodToQuit As Boolean
Public SavedSheet As String
Public SavedCell As String

'+----------------+
'| Functions/Subs |
'+----------------+

Public Sub SavePos()
    SavedSheet = ActiveSheet.Name
    SavedCell = ActiveCell.Address
End Sub

Public Sub LoadPos()
    Sheets(SavedSheet).Activate
    Range(SavedCell).Activate
End Sub

Public Sub showForm()
    fLabelCatalog.Show
End Sub

Public Sub HideCatalog()
    fLabelCatalog.Hide
    fPlaceholder.Show (vbModeless)
End Sub

Public Sub UnhideCatalog()
    fPlaceholder.Hide
    fLabelCatalog.Show
End Sub

Public Sub allowScroll()
    ActiveWindow.DisplayHorizontalScrollBar = True
    ActiveWindow.DisplayVerticalScrollBar = True
    ActiveWindow.DisplayGridlines = True
End Sub

Public Sub disallowScroll()
    ActiveWindow.DisplayHorizontalScrollBar = False
    ActiveWindow.DisplayVerticalScrollBar = False
    ActiveWindow.DisplayGridlines = False
End Sub

Public Sub DumpGlobals()
    MsgBox "Global variable dump (pt1):" & vbNewLine & _
    "[ENABLED FEATURES]" & vbNewLine & _
    "White_Enabled: " & CStr(White_Enabled) & vbNewLine & _
    "White_5GA_Enabled: " & CStr(White_5GA_Enabled) & vbNewLine & _
    "White_1GA_Enabled: " & CStr(White_1GA_Enabled) & vbNewLine & _
    "White_QT_Enabled: " & CStr(White_QT_Enabled) & vbNewLine & _
    "Silver_Enabled: " & CStr(Silver_Enabled) & vbNewLine & _
    "Silver_5GA_Enabled: " & CStr(Silver_5GA_Enabled) & vbNewLine & _
    "Silver_1GA_Enabled: " & CStr(Silver_1GA_Enabled) & vbNewLine & _
    "Silver_QT_Enabled: " & CStr(Silver_QT_Enabled) & vbNewLine & _
    "Kit_Enabled: " & CStr(Kit_Enabled) & vbNewLine & _
    "Kit_5GA_Enabled: " & CStr(Kit_5GA_Enabled) & vbNewLine & _
    "Kit_1GA_Enabled: " & CStr(Kit_1GA_Enabled) & vbNewLine & _
    "Kit_QT_Enabled: " & CStr(Kit_QT_Enabled), _
    vbOKOnly, "Debug message: SHOW TO DOUG"
    MsgBox "Global variable dump (pt2):" & vbNewLine & _
    "[SYMBOLIC CONSTANTS]" & vbNewLine & _
    "Label_None: " & Label_None & vbNewLine & _
    "Label_Silver: " & Label_Silver & vbNewLine & _
    "Label_White: " & Label_White & vbNewLine & _
    "Label_Kit: " & Label_Kit & vbNewLine & _
    "Label_Max: " & Label_Max & vbNewLine & _
    "Label_Size_None: " & Label_Size_None & vbNewLine & _
    "Label_Size_5GA: " & Label_Size_5GA & vbNewLine & _
    "Label_Size_1GA: " & Label_Size_1GA & vbNewLine & _
    "Label_Size_QT: " & Label_Size_QT & vbNewLine & _
    "Label_Size_Max: " & Label_Size_Max & vbNewLine & _
    "Exp_None: " & Exp_None & vbNewLine & _
    "Exp_6Mon: " & Exp_6Mon & vbNewLine & _
    "Exp_12Mon: " & Exp_12Mon & vbNewLine & _
    "Exp_Max: " & Exp_Max & vbNewLine & _
    "[STATE VARIABLES]" & vbNewLine & _
    "Label_Selected: " & Label_Selected & vbNewLine & _
    "Label_Size_Selected: " & Label_Size_Selected & vbNewLine & _
    "Inserted_5GA: " & Inserted_5GA & vbNewLine & _
    "Inserted_1GA: " & Inserted_1GA & vbNewLine & _
    "Inserted_QT: " & Inserted_QT & vbNewLine & _
    "Exp_Period: " & Exp_Period & vbNewLine, _
    vbExclamation & vbOKOnly, "Debug message: SHOW TO DOUG"
End Sub
