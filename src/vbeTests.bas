Attribute VB_Name = "vbeTests"
Option Explicit

Private Function testAll()
  Dim test As Boolean: test = True
  
  test = testOptionClass
  
  testAll = test
End Function

Private Function testOptionClass()
  Dim test As Boolean: test = True
  
  Dim opt As New vbeOption
  
  ' # Initializing options
  
  ' ## Test optionString parsing for single options
  opt.optionString = "Option"
  
  test = opt.name = "Option"
  test = test And IsEmpty(opt.expectedArgs)
  
  ' ## Test optionString parsing for options with
  '    expected args
  opt.optionString = "Option arg1 arg2"
  
  test = test And opt.name = "Option"
  test = test And assertArraysEqual(opt.expectedArgs, Array("arg1", "arg2"))
  
  ' ## Test typename validation
  On Error Resume Next
  opt.typename = "garbage"
  If Err.Number <> 0 Then test = test And True
  On Error GoTo 0
  
  ' ## Test missing default and typename
  On Error Resume Next
  opt.default = Empty
  If Err.Number <> 0 Then test = test And True
  On Error GoTo 0
  
  ' ## test missing defaults for each typename
  opt.typename = "bool"
  opt.default = Empty
  test = test And opt.default = False
  
  opt.typename = "string"
  opt.default = Empty
  test = test And opt.default = vbNullString
  
  opt.typename = "num"
  opt.default = Empty
  test = test And opt.default = 0
  
  testOptionClass = test
End Function