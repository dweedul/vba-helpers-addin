VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "vbeVBComponent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'#RelativePath = src

' This is the base class for the code module manipulations

'# requires vbeOptionParser
'# references "Microsoft Visual Basic for Applications Extensibility 5.3"

Option Explicit

Private m_baseObject As VBComponent
Private m_OptionParser As vbeOptionParser

' ## Properties

' Get the base component stored in this object.
'
' Default member of the class
Public Property Get baseObject() As VBComponent
Attribute baseObject.VB_UserMemId = 0
  Set baseObject = m_baseObject
End Property

' Set the base object
Public Property Set baseObject(baseObject As VBComponent)
  Set m_baseObject = baseObject
  
  ' run the option parser as soon as the baseObject is set
  With m_baseObject.CodeModule
    m_OptionParser.parse .Lines(1, .CountOfLines)
  End With ' codeModule
End Property

' Return the options hash
Public Property Get options() As Dictionary: Set options = m_OptionParser.options: End Property

' Return the module's parent project
Public Property Get project() As VBProject: Set project = m_baseObject.Collection.Parent: End Property

' ## Constructor/Destructor

' Instantiate objects and set defaults
Private Sub Class_Initialize()
  Set m_OptionParser = New vbeOptionParser
  
  buildOptionSet
End Sub

' ## Private Members

' Build the option set for all Components
Private Sub buildOptionSet()
  ' add the no-reload option
  ' this flag, when present, prevents the module from reloading
  m_OptionParser.addOption "no-reload", typename:="bool", default:=False
  
  ' add the no-export option
  ' this flag, when present, prevents the module from exporting
  m_OptionParser.addOption "no-export", typename:="bool", default:=False
  
  ' add the relative-path string option
  ' this tells the addin to save the component to / load the component
  ' from this path, relative to the parent workbook
  m_OptionParser.addOption "relative-path", typename:="string", default:=""
End Sub
