Attribute VB_Name = "globalsHelpers"
'! relative-path vba

' Saves and loads globals
' All credit to [Tushar Mehta][tm]
'
' [tm]: http://www.tushar-mehta.com/publish_train/xl_vba_cases/1017%20Save%20a%20global%20in%20an%20Excel%20workbook.shtml

Option Explicit

' API definitions
#If VBA7 Then
  Public Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias _
    "RtlMoveMemory" (destination As Any, source As Any, ByVal length As Long)
#Else
  Public Declare Sub CopyMemory Lib "kernel32" Alias _
    "RtlMoveMemory" (destination As Any, source As Any, ByVal length As Long)
#End If

' Save the current object's pointer to a named function in the workbook
Public Sub saveGlobal(Glbl As Object, GlblName As String)
  
  #If VBA7 Then
    Dim lngPtr As LongPtr
  #Else
    Dim lngPtr As Long
  #End If
  
  ' Get the pointer to the object
  lngPtr = ObjPtr(Glbl)
  
  ' Store the pointer in a named function
  With ThisWorkbook
    On Error Resume Next
    .Names(GlblName).Delete
    On Error GoTo 0
    .Names.Add GlblName, lngPtr
    
    ' hide from public view
    .Names(GlblName).Visible = False
    
    ' save the file
    .Saved = True
  End With
End Sub

' Get the current object by name
Function GetGlobal(GlblName As String) As Object
  
  ' Get the pointer from the named function
  #If VBA7 Then
    Dim X As LongPtr
    X = CLngPtr(Mid(ThisWorkbook.Names(GlblName).RefersTo, 2))
  #Else
    Dim X As Long
    X = CLng(Mid(ThisWorkbook.Names(GlblName).RefersTo, 2))
  #End If

  Dim obj As Object
  
  ' Convert the pointer to an object
  CopyMemory obj, X, Len(X)

  ' Return the Object
  Set GetGlobal = obj
End Function


