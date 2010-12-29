VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Const cADDIN_CODE_NAME = "VBEHelpers"

' ---------------
' Workbook Events
' ---------------

Private Sub Workbook_Open()
  Build_AddinMenu
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
  Remove_AddinMenu
End Sub

' -------------------
' Addin Specific Code
' -------------------
Private Sub Build_AddinMenu()
  Dim cmdBar As CommandBar
  
  vbeDeleteCustomMenu cADDIN_CODE_NAME
  vbeDeleteEventHandlers
  
  Set cmdBar = Application.VBE.CommandBars.Add(cADDIN_CODE_NAME)
  cmdBar.Position = msoBarTop
  
  vbeAddCommandBarItem _
    "Export Selected File", _
    "vbeExportSelectedCodeModule", _
    cADDIN_CODE_NAME, _
    cmdBar
  
  vbeAddCommandBarItem _
    "Export Active VBProject", _
    "vbeExportActiveVBProject", _
    cADDIN_CODE_NAME, _
    cmdBar
  
  vbeAddCommandBarItem _
    "TEST", _
    "test_TEST", _
    cADDIN_CODE_NAME, _
    cmdBar
    
  cmdBar.RowIndex = 6
  cmdBar.Visible = True
End Sub

Private Sub Remove_AddinMenu()
  vbeDeleteCustomMenu cADDIN_CODE_NAME
End Sub

