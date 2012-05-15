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
  Dim c As VBComponent: Set c = p.VBComponents(component)
  
  ' remove the component from the project
  p.VBComponents.Remove c
  
  ' import the component
  Set c = p.VBComponents.import(path)
  
  ' activate as required
  If shouldActivate Then c.Activate
End Sub
