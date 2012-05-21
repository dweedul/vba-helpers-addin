VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BarCtlEvent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'#RelativePath = vba

' This will act a wrapper for the command bar events

Option Explicit

' storage for the command bar event
Public WithEvents EventHandler As VBIDE.CommandBarEvents
Attribute EventHandler.VB_VarHelpID = -1

' Event code
' This is called when a item is clicked. Call the
' procedure named in the OnAction property of the
' CommandBarControl passed into this procedure.
Private Sub EvtHandler_Click( _
              CommandBarControl As Object, _
              Handled As Boolean, _
              CancelDefault As Boolean)

    On Error Resume Next
    Application.Run CommandBarControl.OnAction

    ' Tell VBIDE that we've handled the event.
    Handled = True: CancelDefault = True
End Sub

