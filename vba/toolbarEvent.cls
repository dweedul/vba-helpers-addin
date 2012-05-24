VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "toolbarEvent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'! relative-path vba

' This will act a wrapper for the command bar events

Option Explicit

' storage for the command bar event
Public WithEvents eventHandler As VBIDE.CommandBarEvents
Attribute eventHandler.VB_VarHelpID = -1

' Schedule the control's OnAction to occur after the current procedure stack completes
'
' Passes the command bar's name and the control's tag to the handler function
Private Sub eventHandler_Click(ByVal Control As Object, Handled As Boolean, CancelDefault As Boolean)
    'On Error Resume Next
    
    Application.OnTime Now(), "'" & ThisWorkbook.name & "'!'" & Control.OnAction & " """ & _
                              Control.Parent.name & """, """ & _
                              Control.Tag & """'"

    ' Tell VBIDE that we've handled the event.
    Handled = True: CancelDefault = True
End Sub
