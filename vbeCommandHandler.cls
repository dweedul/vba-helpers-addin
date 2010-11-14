VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "vbeCommandHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents EvtHandler As VBIDE.CommandBarEvents
Attribute EvtHandler.VB_VarHelpID = -1

Private Sub EvtHandler_Click(ByVal CommandBarControl As Object, _
    Handled As Boolean, CancelDefault As Boolean)
    '''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' This is called when a item is clicked. Call the
    ' procedure named in the OnAction property of the
    ' CommandBarControl passed into this procedure.
    '''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    On Error Resume Next
    Application.Run CommandBarControl.OnAction
    '''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Tell VBIDE that we've handled the event.
    '''''''''''''''''''''''''''''''''''''''''''''''''''''
    Handled = True
    CancelDefault = True
End Sub


