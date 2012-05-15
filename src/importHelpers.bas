Attribute VB_Name = "importHelpers"
'#RelativePath = src

' The option is the atomic nugget used in the OptionParser.
' This will be where the information about each option is stored.

Option Explicit

' Import a component
'
' project        - name of the project
' component      - name of the component
' path           - file path to the component's file
' shouldActivate - should the component be selected afterward
'                  defaults to TRUE
'
' Note: This should NEVER be called normally.
'       It should ALWAYS be scheduled with `Application.OnTime`
'       The VBE doesn't handle import/exports correctly, otherwise.
Public Sub importFromFile( _
             project As String, _
             component As String, _
             path As String, _
             Optional shouldActivate As Boolean = True)
  
  Dim p As VBProject:  Set p = Application.VBE.VBProjects(project)
  Dim c As VBComponent
  
  ' remove the component from the project if it exists
  If VBComponentExists(component, p) Then
    Set c = p.VBComponents(component)
    p.VBComponents.Remove c
  End If
  
  ' import the component
  Set c = p.VBComponents.import(path)
  
  ' activate as required
  If shouldActivate Then c.Activate
End Sub

' Check for the existence of a vbcomponent
'
' component - name of the component
' project   - either a project name or object
'             defaults to the current VBProject
'
' Returns true/false existence.
Private Function VBComponentExists( _
                  ModuleName As String, _
                  Optional project As Variant) _
                  As Boolean
  
  Dim tmp As Variant, VBProj As Object
  
  On Error GoTo errorHandler
  
  If IsMissing(project) Then
    Set VBProj = ThisWorkbook.VBProject
  ElseIf typename(project) = "VBProject" Then
    Set VBProj = project
  Else
    Set VBProj = Application.VBE.VBProjects(project)
  End If

  Set tmp = VBProj.VBComponents(ModuleName)
  
  VBComponentExists = True
  On Error GoTo 0
  Exit Function
  
errorHandler:
  VBComponentExists = False
End Function

