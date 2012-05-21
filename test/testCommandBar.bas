Attribute VB_Name = "testCommandBar"
'#RelativePath = test

'! no-reload
'! relative-path test

' These tests are for the toolbar handler.

Option Explicit


Public Function testBar() As Boolean
  Dim bar As New ToolbarHandler
  
  bar.newBar "testBar"
  
  bar.addButton "foo", "foo"
End Function

Public Sub testBar_reset()
  Application.VBE.CommandBars("testBar").Delete
End Sub

Public Sub foo()
  MsgBox "bar"
End Sub
