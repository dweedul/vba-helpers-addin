Attribute VB_Name = "testVBComponent"
'#RelativePath = test

'# no-reload
'# no-export
'# relative-path test

Option Explicit

Public Function testVBComponent() As Boolean
  Dim test As Boolean
  Dim comp As New vbeVBComponent
  
  Set comp.baseObject = Application.VBE.SelectedVBComponent
  
  ' ## test option parsing
  test = comp.options("no-reload") = True
  test = test And comp.options("no-export") = True
  test = test And comp.options("relative-path") = "test"
  
  ' ## test project property
  test = test And comp.project.name = Application.VBE.ActiveVBProject.name
End Function

