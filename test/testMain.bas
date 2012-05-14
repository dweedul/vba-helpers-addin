Attribute VB_Name = "testMain"
'#RelativePath = test

Option Explicit

Public Sub testAll()
  Dim test As Boolean: test = True
  
  test = testOptionClass
  test = test And testOptionParser
  
  Debug.Print test
End Sub
