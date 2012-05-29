VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'! relative-path vba

Option Explicit

' Initialize
Private Sub Workbook_Open()
  toolbarInit
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
  toolbarDestroy
End Sub

