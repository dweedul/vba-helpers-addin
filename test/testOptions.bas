Attribute VB_Name = "testOptions"
'#RelativePath = test
'! relative-path test

Option Explicit

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

Private Function testOptionParser()
  Dim test As Boolean: test = True
  Dim parser As New vbeOptionParser, options As Dictionary
  Dim testString As String
  Dim o As Object
  
  testString = "'! boolOption" & vbCrLf & _
               "blah bloo blah" & vbCrLf & _
               "'! strOption ""arg1""    arg2 arg3" & vbCrLf & _
               "hi, i'm football" & vbCrLf & _
               "'! numOption 10" & vbCrLf & _
               "'! pathOption ""C:\path to file"""
  
  ' set up the option parser
  parser.addOption "boolOption", "bool", False
  parser.addOption "strOption <arg1> <arg2> <arg3>", "string"
  parser.addOption "numOption <arg1>", "num"
  parser.addOption "pathOption", "string"
    
  Set options = parser.parse(testString)
  
  ' ## Test the options hash yields correct results
  test = options("boolOption") = True
  test = test And options("strOption") = "arg1"
  test = test And options("strOption").args(2) = "arg2"
  test = test And options("numOption") = 10
  test = test And options("pathOption") = "C:\path to file"
  
  testOptionParser = test
End Function
