VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ToolbarHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'#RelativePath = vba

' Build the toolbar and handle the events fired.

Option Explicit

Dim m_bar As CommandBar
Dim events As Collection

' ## Constructor/Destructor

Private Sub Class_Initialize()
  Set events = New Collection
End Sub

' Create a new toolbar or attach to an existing toolbar
'
' name - the name of the toolbar
'
' Returns the bar
Public Function newBar( _
                  name As String) _
                  As CommandBar
                    
  On Error GoTo errorHandler
  
  'Delete the bar if it already exists
  On Error Resume Next
  Application.VBE.CommandBars(name).Delete
  On Error GoTo errorHandler
  
  ' Create the new bar
  Set m_bar = Application.VBE.CommandBars.Add
  With m_bar
    .name = name
    .Position = msoBarTop
    .Visible = True
  End With
                    
errorHandler:
  ' support chaining
  Set newBar = m_bar
End Function

' Remove the bar
Public Sub remove()
  On Error Resume Next
  m_bar.Delete
  On Error GoTo 0
End Sub

' Add a button to the bar
'
' Caption    - the label for the button
' OnAction   - the command to call when the button is pressed
' FaceId     - the icon to put on the button. Defaults to no icon
' Style      - style for the button
' BeginGroup - put a separator in front of the button
'
' Returns the button object.
Public Function addButton( _
                  Caption As String, _
                  OnAction As String, _
                  Optional FaceId As Long = 0, _
                  Optional Style As MsoButtonStyle = msoButtonCaption, _
                  Optional BeginGroup As Boolean = False) _
                  As CommandBarButton
  Dim btn As CommandBarButton
  Dim evt As New BarCtlEvent
  
  On Error GoTo errorHandler
  
  Set btn = m_bar.Controls.Add
  
  With btn
    .Caption = Caption
    .OnAction = "'" & ThisWorkbook.name & "'!" & OnAction
  
    .BeginGroup = BeginGroup
    .FaceId = FaceId
  
    .Style = Style
  End With
  
  ' get the event handler for the button
  Set evt.EventHandler = Application.VBE.events.CommandBarEvents(btn)
  
  ' register the event handler in the event store
  events.Add evt
    
errorHandler:
  ' support chaining
  Set addButton = btn
End Function

' ## Helpers

' Check if a toolbar exists
'
' name - the name of the toolbar to check
' bar  - pointer to the toolbar if found
'
' Return true/false existence
Private Function barExists( _
                   name As String, _
                   Optional ByRef bar As CommandBar = Nothing) _
                   As Boolean
  barExists = True

  On Error GoTo errorHandler:
  Set bar = Application.VBE.CommandBars(name)
  Exit Function
  
errorHandler:
  Set bar = Nothing
  barExists = False
End Function
