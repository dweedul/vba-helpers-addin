Attribute VB_Name = "vbeHelpers"
'! relative-path vba

' General VBE helper functions

Option Explicit

' Remove all code from the VBProject
'
' Returns the target project.
Public Function clearVBProject( _
                  project As VBProject) _
                  As VBProject
  Dim c As vbeVBComponent, i As VBComponent
  
  On Error GoTo errorHandler
  
  For Each i In project.VBComponents
    Set c = New vbeVBComponent
    Set c.baseObject = i
    If c.baseObject.Type = vbext_ct_Document Then
      c.clear
    Else
      c.remove
    End If
  Next ' i
  
errorHandler:
  Set clearVBProject = project
End Function

