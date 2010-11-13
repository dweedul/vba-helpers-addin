VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "udtOption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'********************************************************
' Class for encapsulating the parts of the comment string
'********************************************************
Private m_Variable As String
Private m_Value As Variant

'***********
' Properties
'***********
Public Property Get Variable() As String: Variable = m_Variable: End Property
Public Property Let Variable(Variable As String): m_Variable = Variable: End Property

Public Property Get Value() As Variant
Attribute Item.VB_UserMemId = 0
' Default method of the class
  Value = m_Value
End Property
Public Property Let Value(Value As Variant): m_Value = Value: End Property

'***************
' Public methods
'***************
Public Function Parse(OptionLine As String) As udtOption
  Dim equal_pos As Long, sep_pos As Long, off As Long
  Dim var As String, val As String
  
  ' find the position of the separators used
  off = 0
  If Len(OPTIONS_TOKEN) > 1 Then off = Len(OPTIONS_TOKEN)
  sep_pos = InStr(2, OptionLine, OPTIONS_TOKEN, vbTextCompare) + off
  
  off = 0
  If Len(OPTIONS_ASSIGNMENT_TOKEN) > 1 Then off = Len(OPTIONS_ASSIGNMENT_TOKEN)
  equal_pos = InStr(2, OptionLine, OPTIONS_ASSIGNMENT_TOKEN, vbTextCompare) + off
  
  ' get the options and arguments
  If equal_pos < 1 Then
    ' * single word options, assumed to be boolean *
    m_Variable = Trim(Mid(OptionLine, sep_pos + 1))
    m_Value = True
  Else
    ' * multi-word options *
    ' get the option and its value
    m_Variable = Trim(Mid(OptionLine, sep_pos + 1, equal_pos - sep_pos - 1))
    m_Value = Trim(Mid(OptionLine, equal_pos + 1, Len(OptionLine) - equal_pos))
  End If
  
  Set Parse = Me
End Function
