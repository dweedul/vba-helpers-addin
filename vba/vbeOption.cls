VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "vbeOption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'#RelativePath = vba
'! relative-path vba

' The option is the atomic nugget used in the OptionParser.
' This will be where the information about each option is stored.

Option Explicit

' private storage variables
' check out the appropriate get/let|set for more info
Private m_args As Collection
Private m_default As Variant
Private m_expectedArgs As Variant
Private m_name As String
Private m_optionString As String
Private m_typename
Private m_value As Variant

' some important constants
' the values here MUST match those in the vbeOptionParser class
Const cBOOLEAN As String = "bool"
Const cSTRING As String = "string"
Const cNUMBER As String = "num"

' ## Constructor

Private Sub Class_Initialize()
  Set m_args = New Collection
End Sub

' ## Properties

Public Property Get args() As Collection: Set args = m_args: End Property

' store and validate the args provided
Public Property Set args(args As Collection)
  Set m_args = args
  
  ' TODO: validate the number of args matches the expected args
  ' TODO: validate the args based on typename
End Property

' Get the default value.
' If it is currently empty, set it to the appropriate default.
Public Property Get default() As Variant
  If Not IsEmpty(default) Then m_default = getEmptyDefault
  
  default = m_default
End Property

' Store the default value.
' If the supplied value is Empty, then store an appropriate default.
Public Property Let default(default As Variant)
  If Not IsEmpty(default) Then
    m_default = default
  Else
    m_default = getEmptyDefault()
  End If
End Property

' The expected args are derived from the optionString
Public Property Get expectedArgs() As Variant: expectedArgs = m_expectedArgs: End Property

' Get the number of expected arguments for this option
Public Property Get expectedArgsCount() As Long
  expectedArgsCount = UBound(m_expectedArgs) - LBound(m_expectedArgs) + 1
End Property

' The name is derived from the optionString.
Public Property Get name() As String: name = m_name: End Property

' The optionString will be parsed into the command and the expected args.
Public Property Get optionString() As String: optionString = m_optionString: End Property
Public Property Let optionString(optionString As String)
  parseOptionString optionString
End Property

' Typename of the option.
'
' Typename must be one of the defined types.  The letter will validate proper type.
Public Property Get typename() As String: typename = m_typename: End Property
Public Property Let typename(typename As String)
  Dim tmp As String
    
  tmp = LCase(typename)
  
  ' Validate the typename before storing
  If tmp <> cBOOLEAN And tmp <> cSTRING And tmp <> cNUMBER Then GoTo errorHandler
    
  m_typename = tmp
  
  On Error GoTo 0
  Exit Property
  
errorHandler:
  Err.Raise 9999, "vbeOptionParser::getEmptyDefault", _
            "The typename is not defined or is not one of the following: " & cBOOLEAN & ", " & cSTRING & ", or " & cNUMBER
End Property

' Get the value of the option. (DEFAULT MEMBER OF CLASS)
'
' Return the stored value, or the default value.
' NOTE: VB_UserMemId = 0 sets this as the default method of the class
Public Property Get value() As Variant
Attribute value.VB_UserMemId = 0
  If IsEmpty(m_value) Then
    value = Me.default
  Else
    value = m_value
  End If
End Property

' Store the value.
Public Property Let value(value As Variant): m_value = value: End Property

' ## Private Members

' Parse the option string for name and expected args
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

' Get an appropriate default for each accepted type
'
' Returns a type-appropriate default or raises an error.
Private Function getEmptyDefault() As Variant
  Select Case Me.typename
    Case cSTRING
      getEmptyDefault = vbNullString
    Case cBOOLEAN
      getEmptyDefault = False
    Case cNUMBER
      getEmptyDefault = 0
    Case Else
      GoTo errorHandler
  End Select ' Me.typename
  
  On Error GoTo 0
  Exit Function
  
errorHandler:
  Err.Raise 9999, "vbeOptionParser::getEmptyDefault", _
            "The typename is not defined or is not one of the following: " & cBOOLEAN & ", " & cSTRING & ", or " & cNUMBER
End Function
