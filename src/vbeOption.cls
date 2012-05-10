VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "vbeOption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'#RelativePath = src

Option Explicit

Private m_optionString As String
Private m_typename
Private m_default As Variant

' these are derived from the optionString
Private m_name As String
Private m_expectedArgs As Variant

' some important constants
' the values here MUST match those in the vbeOptionParser class
Const cBOOLEAN As String = "bool"
Const cSTRING As String = "string"
Const cNUMBER As String = "num"

Public Property Get optionString() As String: optionString = m_optionString: End Property
Public Property Let optionString(optionString As String)
  parseOptionString optionString
End Property

' Get the default value.  If it is currently empty, set it to the appropriate default.
Public Property Get default() As Variant
  If Not IsEmpty(default) Then m_default = getEmptyDefault
  default = m_default
End Property

' Store the default value. If the supplied value is Empty, then store an appropriate default.
Public Property Let default(default As Variant)
  If Not IsEmpty(default) Then
    m_default = default
  Else
    m_default = getEmptyDefault()
  End If
End Property

Public Property Get typename() As String: typename = m_typename: End Property
Public Property Let typename(typename As String)
  Dim tmp As String
    
  tmp = LCase(typename)
  
  ' Validate the typename before storing
  If tmp <> cBOOLEAN And tmp <> cSTRING And tmp <> cNUMBER Then GoTo ErrorHandler
    
  m_typename = tmp
  
  On Error GoTo 0
  Exit Property
  
ErrorHandler:
  Err.Raise 9999, "vbeOptionParser::getEmptyDefault", _
            "The typename is not defined or is not one of the following: " & cBOOLEAN & ", " & cSTRING & ", or " & cNUMBER
End Property

' The name and the expected args are derived from the optionString
Public Property Get name() As String: name = m_name: End Property
Public Property Get expectedArgs() As Variant: expectedArgs = m_expectedArgs: End Property

' This will parse the option string for name and expected args
'
' Example:
'
'   parseOptionString "Option"
'   ' me.name should be "option"
'   ' me.expectedArgs should be empty
'
'   parseOptionString "Options arg1 arg2"
'   ' me.name should be "option"
'   ' me.expectedArgs should be ["arg1", "arg2"]
Private Sub parseOptionString(str As String)
  Dim pos As Long, args As String
  
  Const cSPACE = " "
  Const cDBL_QUOTES = """"
  
  ' find the position of the first space
  pos = InStr(str, cSPACE)
  
  ' if there is no space, store the whole string as the name
  If pos = 0 Then
    m_name = str
    
  ' otherwise, store the first word as the name
  ' and split the remainder into an array
  Else
    ' grab the first full word from the optionString
    m_name = Trim(Left(str, pos))
    
    ' split the remaining on spaces
    args = Mid(str, pos + 1)
    m_expectedArgs = Split(args, cSPACE)
  End If
End Sub

' Store an appropriate default for each accepted type
Private Function getEmptyDefault() As Variant
  Select Case Me.typename
    Case cSTRING
      getEmptyDefault = vbNullString
    Case cBOOLEAN
      getEmptyDefault = False
    Case cNUMBER
      getEmptyDefault = 0
    Case Else
      GoTo ErrorHandler
  End Select ' Me.typename
  
  On Error GoTo 0
  Exit Function
  
ErrorHandler:
  Err.Raise 9999, "vbeOptionParser::getEmptyDefault", _
            "The typename is not defined or is not one of the following: " & cBOOLEAN & ", " & cSTRING & ", or " & cNUMBER
End Function
