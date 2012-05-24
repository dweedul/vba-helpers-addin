Attribute VB_Name = "toolbarCallbacks"
'! relative-path vba

' Callbacks from the toolbars
'
' All callbacks should have two string parameters
' at the end of their argument list.
'
' string1 - the name of the calling control's parent
' string2 - the tag of the calling control
'
' Example:
'   Public Sub simpleCallback(barName as String, ctlTag as string)
'     MsgBox barName & "::" & ctlTag
'   End Sub

'! requires vbeVBComponent

Option Explicit

' ## Export and Import Handlers

' Exports the selected module based on the options specified within that module
Public Sub ExportSelectedModule(barName As String, ctlTag As String)
  Dim c As New vbeVBComponent
  Set c.baseObject = Application.VBE.SelectedVBComponent
  c.export
End Sub

' Exports each module in the current project based on the options specified within each module
Public Sub ExportActiveProject(barName As String, ctlTag As String)
  Dim c As vbeVBComponent, temp As VBComponent, proj As VBProject
  
  ' store the proj so that we don't lose it during debugging
  Set proj = Application.VBE.ActiveVBProject
  For Each temp In proj.VBComponents
    Set c = New vbeVBComponent
    Set c.baseObject = temp
    c.export
  Next ' component
End Sub

' Reload the current module
Public Sub ReloadSelectedModule(barName As String, ctlTag As String)
  Dim c As New vbeVBComponent
  
  If warnUser("reload-one") Then
    Set c.baseObject = Application.VBE.SelectedVBComponent
    c.reload
  End If
End Sub

' Reload each module in the current project based on the options specified within each module
Public Sub ReloadActiveProject(barName As String, ctlTag As String)
  Dim c As vbeVBComponent, temp As VBComponent, proj As VBProject
  
  If warnUser("reload-all") Then
    ' store the proj so that we don't lose it during debugging
    Set proj = Application.VBE.ActiveVBProject
    
    For Each temp In proj.VBComponents
      Set c = New vbeVBComponent
      Set c.baseObject = temp
      c.reload
    Next ' component
  End If
  
End Sub

'Application.SendKeys cmdBar.Text
