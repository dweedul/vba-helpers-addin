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
'# references "Microsoft Scripting Runtime"

Option Explicit

Private m_baseObject As VBComponent
Private m_OptionParser As vbeOptionParser

Private fso As FileSystemObject

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

' Return the filename based on the component type and name
Public Property Get filename()
  filename = m_baseObject.name & extension
End Property

' Return the options hash
Public Property Get options() As Dictionary: Set options = m_OptionParser.options: End Property

' Return the calculated path to the component
Public Property Get path() As String
  Dim basePath As String
  
  If Len(Me.options("absolute-path")) > 0 Then
    path = Me.options("absolute-path")
  Else
    With fso
      basePath = .BuildPath(.GetParentFolderName(project.filename), Me.options("relative-path"))
      path = .BuildPath(basePath, Me.filename)
    End With ' fso
  End If
End Property

' Return the module's parent project
Public Property Get project() As VBProject: Set project = m_baseObject.Collection.Parent: End Property

' ## Constructor/Destructor

' Instantiate objects and set defaults
Private Sub Class_Initialize()
  Set m_OptionParser = New vbeOptionParser
  Set fso = New FileSystemObject
  
  buildOptionSet
End Sub

' ## Import and Export methods

' Export this code module
'
' overwriteExisting - should the export overwrite an existing copy
'                     defaults to TRUE
' overrideNoExport  - should the export occur in spite of the no-export option
'                     defaults to FALSE
'
' Returns the current module for chaining.
Public Function export( _
                  Optional overwriteExisting As Boolean = True, _
                  Optional overrideNoExport As Boolean = False) _
                  As vbeVBComponent
  
  ' handle the no-export option
  If Me.options("no-export") And Not overrideNoExport Then GoTo errorHandler
  
  ' check for existing file and overwrite as required
  If fso.FileExists(Me.path) And Not overwriteExisting Then GoTo errorHandler
    
  ' create the parent folder if it does not exist
  With fso
    If Not .FolderExists(.GetParentFolderName(Me.path)) Then
      .CreateFolder .GetParentFolderName(Me.path)
    End If
  End With ' fso
  
  ' export the component
  m_baseObject.export Me.path
  
errorHandler:
  ' support chaining by returning Me
  Set export = Me
End Function

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

Private Property Get extension() As String
  Select Case m_baseObject.Type
    Case vbext_ct_ClassModule
        extension = ".cls"
    Case vbext_ct_Document
        extension = ".cls"
    Case vbext_ct_MSForm
        extension = ".frm"
    Case vbext_ct_StdModule
        extension = ".bas"
    Case Else
        extension = ".bas"
  End Select
End Property
