Attribute VB_Name = "toolbarActions"
'#RelativePath = vba
'! relative-path vba

' Handles commands from the toolbars
'
' Note: All toolbar callbacks must end with two strings
'       These will be set to the toolbar's name
'       and the control's tag.

'! requires vbeVBComponent

Option Explicit

' ## Export and Import Handlers

' Exports the selected module based on the options specified within that module
Public Sub ExportSelectedModule(barName As String, ctlTag As String)
  Dim c As New vbeVBComponent
  Set c.baseObject = Application.VBE.SelectedVBComponent
  c.export
End Sub

' Reload the current module
Public Sub ReloadSelectedModule(barName As String, ctlTag As String)
  Dim c As New vbeVBComponent
  Set c.baseObject = Application.VBE.SelectedVBComponent
  c.reload
End Sub

'Application.SendKeys cmdBar.Text
