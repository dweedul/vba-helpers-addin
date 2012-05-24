Attribute VB_Name = "ribbon"
'! relative-path vba

' Ribbon callbacks and helper functions

Option Explicit

Public ribbon As IRibbonUI

' ## Ribbon callbacks

Public Sub ribbon_onLoad(ribbon As IRibbonUI)
  Set ribbon = ribbon
End Sub

' ## Selection type name button callbacks

' Reload the button on a click
Public Sub rbtnSelectionTypeName_onAction(Control As IRibbonControl)
  reloadRibbonControl Control
End Sub

' Set the label to the current selection's type name
Public Sub rbtnSelectionTypeName_getLabel(Control As IRibbonControl, ByRef returnedVal)
  On Error GoTo Local_Error
  
  returnedVal = typename(Selection)
  
  Exit Sub
  
Local_Error:
  returnedVal = "TypeName(Selection)"
End Sub

' ## Ribbon Helper functions

' Reload a ribbon control.
'
' control - the target control
'
' Returns the control.
Private Function reloadRibbonControl( _
                  Control As IRibbonControl) _
                  As IRibbonControl
  On Error GoTo errorHandler
  
  If Not ribbon Is Nothing Then
    ribbon.InvalidateControl Control.ID
  End If
  
errorHandler:
  ' support chaining
  Set reloadRibbonControl = Control
End Function


