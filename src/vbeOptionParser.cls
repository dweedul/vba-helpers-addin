VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "vbeOptionParser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' OptionParsers check a string for specific phrases
' and stores them for use later

Private mOptions As Dictionary

' some important constants
' the values here MUST match those in the vbeOption class
Const cBOOLEAN As String = "bool"
Const cSTRING As String = "string"
Const cNUMBER As String = "num"

' Expose the options hash
Public Property Get Options() As Dictionary: Set Options = mOptions: End Property

' Set defaults
Private Sub Class_Initialize()
  Set mOptions = New Dictionary
End Sub

' Add a new option to the hash
'
' optionString - defines the name and any arguments that you may want to require.
' typename     - type for the option.  Acceptable typenames are "bool", "string", "int"
' default      - a default value. Defaults are set based on the typename
Public Function addOption( _
                optionString As String, _
                Optional typename As String = "boolean", _
                Optional default As Variant = Empty) _
                As Object
  
  Dim newOption As New vbeOption
  
  newOption.optionString = optionString
  newOption.typename = typename
  newOption.default = default
  
  ' add the new command to the master collection
  mOptions.Add newOption.name, newOption
End Function

Public Function parse()
  Dim opt As vbeOption
  
  ' if the option is a boolean,
  ' treat its presence as a toggle from the default
  If opt.typename = cBOOLEAN Then
    opt.Value = Not opt.default
  End If
End Function
