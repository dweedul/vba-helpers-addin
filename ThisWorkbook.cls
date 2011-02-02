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
  Dim cmdBar As CommandBar, pop As CommandBarPopup
  
  vbeDeleteCustomMenu cADDIN_CODE_NAME
  vbeDeleteEventHandlers
  
  Set cmdBar = Application.VBE.CommandBars.Add(cADDIN_CODE_NAME)
  cmdBar.Position = msoBarTop
  
  Set pop = vbeAddCommandBarItem( _
    "Import/Export Code", _
    "vbeExportSelectedCodeModule", _
    cADDIN_CODE_NAME, _
    cmdBar, _
    msoControlPopup)
  
  vbeAddCommandBarItem _
    "Export Selected File", _
    "vbeExportSelectedCodeModule", _
    cADDIN_CODE_NAME, _
    pop
  
  vbeAddCommandBarItem _
    "Export Active VBProject", _
    "vbeExportActiveVBProject", _
    cADDIN_CODE_NAME, _
    pop, _
    KeyBoardShortcut:="^%e"
  
  vbeAddCommandBarItem _
    "Reload Selected File", _
    "vbeReloadCodeModule", _
    cADDIN_CODE_NAME, _
    pop
  
  vbeAddCommandBarItem _
    "Reload Active VBProject", _
    "vbeReloadActiveVBProject", _
    cADDIN_CODE_NAME, _
    pop
    
  'cmdBar.RowIndex = 6
  cmdBar.Visible = True
End Sub

Private Sub Remove_AddinMenu()
  vbeDeleteCustomMenu cADDIN_CODE_NAME
End Sub

