VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "vbeVBComponent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'#RelativePath = vba

' This is the base class for the code module manipulations

'! requires vbeOptionParser
'! references "Microsoft Visual Basic for Applications Extensibility 5.3"
'! references "Microsoft Scripting Runtime"

Option Explicit

Private m_baseObject As VBComponent
Private m_OptionParser As vbeOptionParser

Private fso As FileSystemObject

Const WARNING_IMPORT As String = "We're going to overwrite something!"

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
    m_OptionParser.parse Me.code
  End With ' codeModule
End Property

' Get the code in the component as a string
Public Property Get code() As String
  With m_baseObject.CodeModule
    If .CountOfLines = 0 Then
      code = vbNullString
    Else
      code = .Lines(1, .CountOfLines)
    End If
  End With ' CodeModule
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
  
  ' call the public function where the options are defined.
  Set m_OptionParser = vbeVBComponentOptionParser()
End Sub

' ## Public methods

' Activate this component
'
' Return this
Public Function activate() As vbeVBComponent
  Me.baseObject.activate
  Set activate = Me
End Function

' Clear the code module
'
' Return this
Public Function clear() As vbeVBComponent
  With Me.baseObject.CodeModule
    .DeleteLines 1, .CountOfLines
  End With
  
  Set clear = Me
End Function

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
  
  ' Do not export empty components.  That would be silly
  If Len(Trim(Me.code)) = 0 Then GoTo errorHandler
  
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

' Reload the component from the filesystem
'
' warnUser         - let the user know that there will be an overwrite
'                    defaults to TRUE
' overrideNoImport - reload regardless of the no-reload option
'                    defaluts to FALSE
'
' Returns Me for chaining
Public Function reload( _
                  Optional warnUser As Boolean = True, _
                  Optional overrideNoReload As Boolean = False, _
                  Optional shouldActivate As Boolean = True) _
                  As vbeVBComponent
                    
  Dim resp As VbMsgBoxResult
                    
  ' handle the no-import option
  If Me.options("no-reload") And Not overrideNoReload Then GoTo errorHandler
  
  ' warn the user before importing
  If warnUser Then
    resp = MsgBox(WARNING_IMPORT, vbOKCancel)
  Else
    resp = vbOK
  End If
  
  ' perform the import
  If resp = vbOK Then
    
    ' schedule the import to occur after the current procedure stack completes
    Application.OnTime Now(), ThisWorkbook.name & "!'importFromFile """ & _
                                Me.project.name & """, """ & _
                                Me.baseObject.name & """, """ & _
                                Me.path & """, """ & _
                                shouldActivate & """'"
                                
  End If
                      
errorHandler:
  'support chaining by returning ME
  Set reload = Me
End Function

' Remove this component from it's project
'
' Return this
Public Function remove() As vbeVBComponent
  Me.project.VBComponents.remove m_baseObject
  
  Set m_baseObject = Nothing
  
  ' support chaining
  Set remove = Me
End Function

' ## Private Members
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
