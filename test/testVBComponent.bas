Attribute VB_Name = "testVBComponent"
'#RelativePath = test

'# no-reload
'# relative-path test

' These tests will use this code module to test for options.
' If the project or module names are changed, the following
' constants must be changed for tests to work properly

Private Const PROJECT_NAME As String = "VBEHelpersREWRITE"
Private Const MODULE_NAME As String = "testVBComponent"

Option Explicit

Public Function testVBComponent() As Boolean
  Dim test As Boolean
  Dim comp As New vbeVBComponent
  
  Set comp.baseObject = Application.VBE.VBProjects(PROJECT_NAME).VBComponents(MODULE_NAME)
  
  ' ## test option parsing
  test = comp.options("no-reload") = True
  test = test And comp.options("no-export") = True
  test = test And comp.options("relative-path") = "test"
  
  ' ## test properties
  test = test And comp.project.name = PROJECT_NAME
  test = test And comp.filename = MODULE_NAME & ".bas"
  test = test And comp.path = ActiveWorkbook.path & "\test\" & MODULE_NAME & ".bas"
  
  ' ## test export and import
  comp.export
End Function

