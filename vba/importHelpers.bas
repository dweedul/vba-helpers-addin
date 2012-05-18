Attribute VB_Name = "importHelpers"
'#RelativePath = vba

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
  Dim c As New vbeVBComponent, tmp As New vbeVBComponent
  
  ' remove the component from the project if it exists
  If VBComponentExists(component, p) Then
    Set c.baseObject = p.VBComponents(component)
    
    If c.baseObject.Type = vbext_ct_Document Then
      c.clear
      
      ' import the component into a new object,
      ' copy the tmp component's code into the old object
      ' and then remove the tmp component
      Set tmp.baseObject = p.VBComponents.import(path)
      c.baseObject.CodeModule.InsertLines 1, tmp.code
      tmp.remove: Set tmp = Nothing
      
    Else
      c.remove
      Set c.baseObject = p.VBComponents.import(path)
    End If
    
  Else
    ' import the component
    Set c.baseObject = p.VBComponents.import(path)
  End If
  
  ' activate as required
  If shouldActivate Then c.activate
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


