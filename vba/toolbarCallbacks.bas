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
'! references "MS Forms 2.0 Object Library"
'! references "Microsoft Scripting Runtime"

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

' Import code into the active VB project from a folder of the user's choosing.
Public Sub ImportFolderToActiveProject(barName As String, ctlTag As String)
  Dim proj As VBProject
  
  If warnUser("reload-all") Then
    'store the project
    Set proj = Application.VBE.ActiveVBProject
    
    clearVBProject proj
    importFromFolder proj
  End If
End Sub

' Copy the active project's path to the clipboard
Public Sub CopyPathToClipboard(barName As String, ctlTag As String)
  Dim DataObj As New MSForms.DataObject, s As String, fso As New FileSystemObject
  
  s = fso.GetParentFolderName(Application.VBE.ActiveVBProject.filename)
  
  DataObj.SetText s
  DataObj.PutInClipboard
End Sub

' Copy the active project's path to the clipboard
Public Sub PasteCommandString(barName As String, ctlTag As String)
  Dim txt As String, parser As vbeOptionParser
  
  Set parser = vbeVBComponentOptionParser
  
  txt = parser.optionToken & " " & parser(getControl(barName, ctlTag).Text).optionString
  Application.SendKeys txt
End Sub


' ## Command bar helpers

' Get the control given the barName and tag
'
' barName - the control's parent bar
' ctlTag  - the tag for the control
'
' Return the control.
Private Function getControl(barName As String, ctlTag As String) As CommandBarControl
  Set getControl = Application.VBE.CommandBars(barName).FindControl(Tag:=ctlTag)
End Function

