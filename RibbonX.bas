Attribute VB_Name = "RibbonX"
Option Explicit

Public rbx As IRibbonUI

' ----------------
' Ribbon Callbacks
' ----------------
Public Sub rbx_onLoad(ribbon As IRibbonUI)
  Set rbx = ribbon
End Sub

Public Sub rbxbtnSelectionTypeName_onAction(control As IRibbonControl)
  RefreshRibbonControl control.ID
End Sub

Public Sub rbxbtnSelectionTypeName_getLabel(control As IRibbonControl, ByRef returnedVal)
  On Error GoTo Local_Error
  
  returnedVal = TypeName(Selection)
  
  Exit Sub
  
Local_Error:
  returnedVal = "TypeName(Selection)"
End Sub

' ---------------------
' Ribbon Meta-functions
' ---------------------
Private Function RefreshRibbonControl(ControlId As String) As Boolean
' @return [Boolean] False on error
' @TODO: Add comments here
  
  On Error GoTo Local_Error

  If (Not rbx Is Nothing) Then
    rbx.InvalidateControl ControlId
  End If
  
  RefreshRibbonControl = True
  On Error GoTo 0
  Exit Function

Local_Error:
  RefreshRibbonControl = False
End Function

