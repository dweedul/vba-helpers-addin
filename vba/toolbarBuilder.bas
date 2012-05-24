Attribute VB_Name = "toolbarBuilder"
'! relative-path vba

' Wrappers for building toolbars

'! requires toolbarEvent
'! references "Microsoft Visual Basic for Applications Extensibility 5.3"

Option Explicit

' stores the toolbarEvents
Private eventStore As New Collection

Private Const TOOLBAR_NAME As String = "vba-helper"

' ## Public methods to initialize and destroy the toolbar

' Create the toolbar for the add in
Public Sub toolbarInit()
  Dim bar As CommandBar, menu As CommandBarPopup
  
  Set bar = newToolbar(TOOLBAR_NAME)
  
  
  ' ## Export Menu
  Set menu = addMenu(Parent:=bar, Caption:="Export")
    
  ' export currently selected file
  addButton Parent:=menu, Caption:="Export selected module", OnAction:="ExportSelectedModule"
  
  ' export the entire project
  addButton Parent:=menu, Caption:="Export active project", OnAction:="ExportActiveProject"
  
  
  ' ## Reload Menu
  Set menu = addMenu(Parent:=bar, Caption:="Reload")
  
  ' reload the currently selected module from the file path
  addButton Parent:=menu, Caption:="Reload selected module", OnAction:="ReloadSelectedModule"
    
  ' Reload all modules in the active project
  addButton Parent:=menu, Caption:="Reload active project", OnAction:="ReloadActiveProject"
    
  ' ## Other Buttons
    
  ' Add a button that copies the current project's path to the clipboard
  addButton Parent:=bar, Caption:="Copy file-path to clipboard", _
            OnAction:="foo", _
            FaceId:=22, Style:=msoButtonIcon, _
            Tooltip:="Copy the current project's file path to the clipboard."
End Sub

' Remove the addin's toolbar
Public Sub toolbarDestroy()
  removeToolbar TOOLBAR_NAME
End Sub

' ## Toolbar creation methods

' Create a new toolbar
'
' name - the name of the toolbar
'
' Returns the bar
Private Function newToolbar( _
                  name As String) _
                  As CommandBar
  Dim bar As CommandBar
                      
  On Error GoTo errorHandler
  
  'Delete the bar if it already exists
  removeToolbar name
  
  ' Create the new bar
  Set bar = Application.VBE.CommandBars.Add
  With bar
    .name = name
    .Position = msoBarTop
    .Visible = True
  End With
                    
errorHandler:
  ' support chaining
  Set newToolbar = bar
End Function

' Check if the toolbar exists and remove it
Private Sub removeToolbar(name As String)
 On Error Resume Next
 Application.VBE.CommandBars(name).Delete
 clearEvents
End Sub

' Add a button to the parent
'
' Parent     - container for the control
' Caption    - the label for the button
' OnAction   - the command to call when the button is pressed
' FaceId     - the icon to put on the button. Defaults to no icon
' Style      - style for the button
' BeginGroup - put a separator in front of the button
'
' Returns the button object.
Private Function addButton( _
                  Parent As Object, _
                  Caption As String, _
                  OnAction As String, _
                  Optional FaceId As Long = 0, _
                  Optional Style As MsoButtonStyle = msoButtonCaption, _
                  Optional BeginGroup As Boolean = False, _
                  Optional Tooltip As String = vbNullString, _
                  Optional Tag As String = vbNullString) _
                  As CommandBarButton
                    
  Dim btn As CommandBarButton
  
  Set btn = addItem(Parent, Caption, OnAction, msoControlButton)
  With btn
    .Style = Style
    .FaceId = FaceId
    .BeginGroup = BeginGroup
    .TooltipText = Tooltip
    .Tag = Tag
  End With ' btn
                  
errorHandler:
  ' support chaining
  Set addButton = btn
End Function

' Add a menu to the parent
'
' Parent     - container for the control
' Caption    - the label for the button
' BeginGroup - put a separator in front of the button
' Tooltip    - string tooltip to show on hover
'
' Returns the menu.
Private Function addMenu( _
                  Parent As Object, _
                  Caption As String, _
                  Optional BeginGroup As Boolean = False, _
                  Optional Tooltip As String = vbNullString) _
                  As CommandBarPopup
                    
  Dim menu As CommandBarPopup
  
  Set menu = addItem(Parent, Caption, ControlType:=msoControlPopup)
  With menu
    .BeginGroup = BeginGroup
    .TooltipText = Tooltip
  End With ' btn
                  
errorHandler:
  ' support chaining
  Set addMenu = menu
End Function

' Add an item to the parent control container
'
' parent      - container for the control
' caption     - label for the new control
' onAction    - procedure definition for the control
'               this will be stored in the eventStore for later
' controlType - ms enum defining the type of control
'
' Returns the new control.
Private Function addItem( _
                   Parent As Object, _
                   Caption As String, _
                   Optional OnAction As String = vbNullString, _
                   Optional ControlType As MsoControlType = msoControlButton) _
                   As CommandBarControl
  
  Dim evt As New toolbarEvent, ctl As CommandBarControl

  Set ctl = Parent.Controls.Add(Type:=ControlType)
  With ctl
    .Caption = Caption
    .OnAction = OnAction
  End With
  
  If Len(OnAction) > 0 Then
    ' Add the event to the event handler store
    Set evt.eventHandler = Application.VBE.Events.CommandBarEvents(ctl)
    eventStore.Add evt
  End If
  
  ' support chaining
  Set addItem = ctl
End Function

' ## Event code

' Clear all event handlers from the eventStore
' This prevents the event code from firing multiple times.
Private Sub clearEvents()
  Do Until eventStore.Count = 0
    eventStore.remove 1
  Loop
End Sub
