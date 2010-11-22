Attribute VB_Name = "vbeCommandBarHelpers"
Option Explicit
Option Private Module

Private EventHandlers As New Collection

Public Function vbeAddCommandBarItem( _
              Caption As String, _
              ProcName As String, _
              Tag As String, _
              TargetControlCollection As Object, _
              Optional ControlType As MsoControlType = msoControlButton) _
              As CommandBarControl
  
  Dim MenuEvent As vbeCommandHandler
  Dim cmdBarItem As CommandBarControl
  
  Set MenuEvent = New vbeCommandHandler
  
  Set cmdBarItem = TargetControlCollection.Controls.Add(Type:=ControlType)
  With cmdBarItem
    .Caption = Caption
    .OnAction = "'" & ThisWorkbook.Name & "'!" & ProcName
    .Tag = Tag
    If .Type = msoControlButton Then .Style = msoButtonCaption
  End With
  
  Set MenuEvent.EvtHandler = Application.VBE.Events.CommandBarEvents(cmdBarItem)
  
  EventHandlers.Add MenuEvent
End Function

Public Sub vbeDeleteCustomMenu(Name As String)
  If CommandBarExists(Name) Then Application.VBE.CommandBars(Name).Delete
End Sub

Public Sub DeleteMenuItems(Optional Tag As String)
'---------------------------------------------
' Deletes all controls that have a certain tag
'---------------------------------------------
    Dim Ctrl As Office.CommandBarControl
    Set Ctrl = Application.VBE.CommandBars.FindControl(Tag:=Tag)
    Do Until Ctrl Is Nothing
        Ctrl.Delete
        Set Ctrl = Application.VBE.CommandBars.FindControl(Tag:=Tag)
    Loop

End Sub

Public Sub vbeDeleteEventHandlers(Optional HideMe As Boolean)
'-------------------------------------------------
' Delete any existing event handlers.
' HideMe removes this function from the macro menu
'-------------------------------------------------
  Do Until EventHandlers.Count = 0
      EventHandlers.Remove 1
  Loop
End Sub

Private Function CommandBarExists(sName As String) As Boolean
    Dim s As String

    On Error GoTo bWorksheetExistsErr

    s = Application.VBE.CommandBars(sName).Name
    CommandBarExists = True
    Exit Function

bWorksheetExistsErr:
    CommandBarExists = False
End Function

