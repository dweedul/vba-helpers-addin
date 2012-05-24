VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "vbeOptionParser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'#RelativePath = vba
'! relative-path vba

' OptionParsers check a string for specific phrases
' and stores them for use later

'! requires vbeOption
'! references "Microsoft VBScript Regular Expressions 5.5"
'! references "Microsoft Scripting Runtime"

Option Explicit

Private m_Options As Dictionary

' The token to look for within the parsed string
Public optionToken As String

' some important constants
' the values here MUST match those in the vbeOption class
Const cBOOLEAN As String = "bool"
Const cSTRING As String = "string"
Const cNUMBER As String = "num"

' ## Properties

' Expose the options hash
Public Property Get options() As Dictionary: Set options = m_Options: End Property

' Set defaults
Private Sub Class_Initialize()
  Set m_Options = New Dictionary
  
  ' set default optionToken from the config
  optionToken = "'!"
End Sub

' ## Public Members

' Add a new option to the hash
'
' optionString - defines the name and any arguments that you may want to require.
' typename     - type for the option.  Acceptable typenames are "bool", "string", "int"
' default      - a default value. Defaults are set based on the typename
Public Function addOption( _
                optionString As String, _
                Optional typename As String = cBOOLEAN, _
                Optional default As Variant = Empty) _
                As Object
  
  Dim newOption As New vbeOption
  
  newOption.optionString = optionString
  newOption.typename = typename
  newOption.default = default
  
  ' add the new command to the master collection
  m_Options.Add newOption.name, newOption
End Function

' Parse the given string for lines starting with the optionToken
'
' searchString - the string to parse
'
' Returns the options hash for the given string
Public Function parse( _
                  searchString As String) _
                  As Dictionary
  Dim buf As String, optName As String
  Dim opt As vbeOption
  Dim re As New RegExp, optionLines As MatchCollection, optionLine As Match
  Dim args As MatchCollection, arg As Match
  Dim i As Long, REGEX_OPTION_LINE As String
  
  Const cSPACE As String = " "
  REGEX_OPTION_LINE = optionToken & "\s*([\\\.\[\]\(\)\w "":@#$%&-]*)"
  
  ' find all the lines that begin with the optionToken
  re.Pattern = REGEX_OPTION_LINE & "[\n\f\r]+" & "|" & REGEX_OPTION_LINE & "$"
  re.Global = True
  re.IgnoreCase = True
  Set optionLines = re.Execute(searchString)
  
  ' loop through the option lines
  For Each optionLine In optionLines
    buf = Submatch(optionLine)
    
    ' get the first word in the line,
    ' that will be the option name
    optName = Trim(Left(buf, InStr(buf, cSPACE)))
    If optName = vbNullString Then optName = buf
    
    ' look up the option name in the options hash
    ' if it exists, perform the appropriate action
    If m_Options.Exists(optName) Then
      Set opt = m_Options(optName)
      
      ' booleans are flags whose presence toggles the value
      ' to the opposite of the default value.
      If opt.typename = cBOOLEAN Then opt.value = Not opt.default
      
      ' The number and string option types will have at least one argument.
      ' The first argument will be stored as the value for the option.
      ' The argument list is a collection of strings
      If opt.typename = cSTRING Or opt.typename = cNUMBER Then
        Set opt.args = parseArgs(buf)
        opt.value = opt.args(1)
      End If
      
    End If
  Next ' optionLine

  Set parse = m_Options
End Function

' ## Private members

' Parse the given option line for arguments.
'
' optionLine - the String with the argument information
'
' Returns a collection of string arguments
Private Function parseArgs( _
                   optionLine As String) _
                   As Variant
  Dim arr As Collection: Set arr = New Collection
  Dim args As MatchCollection, arg As Match
  
  On Error GoTo errorHandler
  
  Const REGEX_DOUBLE_QUOTED_STRING As String = "\s+""([^""].*)"""
  Const REGEX_SPACE_WRAPPED_STRING As String = "\s+(\w+)\b"
  
  With New RegExp
    .Global = True
    .IgnoreCase = True
    .Pattern = REGEX_DOUBLE_QUOTED_STRING & "|" & REGEX_SPACE_WRAPPED_STRING
    Set args = .Execute(optionLine)
  End With
  
  ' store the args
  For Each arg In args
    arr.Add Submatch(arg)
  Next ' arg
  
  Set parseArgs = arr
  On Error GoTo 0
  Exit Function
  
errorHandler:
  parseArgs = New Collection
End Function
        
' Search the match object for the first non-empty submatch
' This is to correct the bug that occurs when only one submatch should
' be found, but mulitples are returned.
' All but one should really be empty.

' Return the value of the submatch as a string
Private Function Submatch(m As Match) As String
  Dim i As Long
  
  For i = 0 To m.SubMatches.Count - 1
    If Not IsEmpty(m.SubMatches(i)) Then
      Submatch = m.SubMatches(i)
      Exit Function
    End If
  Next ' s
End Function

' Return a vbeOption item from the dictionary __DEFAULT MEMBER OF THE CLASS__
'
' NOTE: VB_UserMemId = 0 sets this as the default method of the class
Function Item(key As Variant) As vbeOption
Attribute Item.VB_UserMemId = 0
  Set Item = m_Options.Item(key)
End Function
