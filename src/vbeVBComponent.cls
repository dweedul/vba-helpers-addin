VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "vbeVBComponent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' #RelativePath = src

Option Explicit

'-----------------------------------------------------------------
' This object gathers the information from the code modules needed
' to perform the actions required by the other modules.
' This object will be a member of a collection.
'-----------------------------------------------------------------

'---------------
' Public members
'---------------
Private m_Options As Collection
Private m_Index As Long ' index in the collection
Private m_ModuleName As String ' name of the Code Module
Private m_Project As VBProject ' name of the Project
Private m_ComponentType As vbext_ComponentType

Private m_Parsed As Boolean ' flag to tell if the CodeModule has been checked for options

'-----------
' Properties
'-----------
Public Property Get Index() As Long: Index = m_Index: End Property
Public Property Let Index(Index As Long): m_Index = Index: End Property

Public Property Get Project() As VBProject: Project = m_Project: End Property
Public Property Set Project(Project As VBProject): m_Project = Project: End Property

Public Property Get ModuleName() As String: ModuleName = m_ModuleName: End Property
Public Property Let ModuleName(ModuleName As String): m_ModuleName = ModuleName: End Property

Public Property Get ComponentType() As vbext_ComponentType: ComponentType = m_ComponentType: End Property
Public Property Let ComponentType(ComponentType As vbext_ComponentType): m_ComponentType = ComponentType: End Property

Public Property Get VBComponent() As Object
  On Error GoTo Local_Error
  Set VBComponent = Project.VBComponents(ModuleName)
  
  On Error GoTo 0
  Exit Property
  
Local_Error:
  Set VBComponent = Nothing
End Property
Public Property Set VBComponent(VBComponent As Object)
  Me.ModuleName = VBComponent.Name
  Me.Project = VBComponent.Collection.Parent
  Me.ComponentType = VBComponent.Type
End Property

Public Property Get IsEmpty() As Boolean
  If Me.VBComponent.CodeModule.CountOfLines < 1 Then
    IsEmpty = True
  Else
    IsEmpty = False
  End If
End Property

Public Property Get Options(Optional key As Variant) As Variant
' will default to a False if the key is not found
' @optparam Key [int|string] The item to search for
' @return [collection|obj|boolean] _
'    default: option collection _
'    key specified and found: the single option _
'    else: FALSE

  Dim out As udtOption
  
  If Not m_Parsed Then ParseCodeModule
  
  If IsMissing(key) Then
    Set Options = m_Options
  ElseIf KeyExists(key, Options) Then
  Else
    ' add an default option to the collection
    Set out = New udtOption
    out.Variable = key ': out.Value = False
    
    Select Case out.OptionType
      Case OptionType_String: out.Value = vbNullString
      Case Else: out.Value = False
    End Select
    
    Set Options = out
  End If
End Property

'-----------------------
' Constructor/Destructor
'-----------------------
Private Sub Class_Initialize()
  Set m_Options = New Collection
End Sub
Private Sub Class_Terminate()
  Set m_Options = Nothing
End Sub

'---------------
' Public methods
'---------------
Public Function OptionList(Optional Delimiter As String = vbCrLf) As String
  Dim o As udtOption, out As String
  For Each o In Me.Options
    out = out & o.Variable & "=" & o.Value & Delimiter
  Next ' o
  
  OptionList = out
End Function

'----------------
' Private methods
'----------------
Private Sub ParseCodeModule()
  Dim i As Long
  Dim opt As udtOption
  Dim tmp As String

  
  With Me.VBComponent.CodeModule
  ' Loop through the lines looking for options to process

    For i = 1 To .CountOfLines
      ' Get the current line
      tmp = .Lines(i, 1)
  
      ' Process all non-blank comment lines
      If Len(Trim(tmp)) > 1 Then
        ' Exit early once the comments are finished
        If Not Left(LTrim(tmp), 1) Like COMMENT_TOKEN Then Exit For
        
        ' save the variables into the collection
        Set opt = New udtOption
        ProcessOption opt.Parse(tmp)
      End If
    Next ' i
  End With ' Me.VBComponent.CodeModule
  
  m_Parsed = True
End Sub

Private Sub ProcessOption(opt As udtOption)
' Adds to the Options collection depending on what is provided.
  
  ' Remove the option if it already exists
  If KeyExists(opt.Variable) Then m_Options.Remove opt.Variable
 
  m_Options.Add opt, opt.Variable
  
End Sub

Private Function KeyExists( _
                  key As Variant, _
                  Optional FoundItem As Variant) _
                  As Boolean
' Check if item exists in the collection by key or by position

  On Error GoTo ErrNotExist
   
  If VarType(key) = vbString Then ' check the string keys
    'Test if item exists
    If VarType(m_Options.Item(CStr(key))) = vbObject Then
      'Return an object
      Set FoundItem = m_Options.Item(CStr(key))
    Else
      'Return an standard variable
      FoundItem = m_Options.Item(CStr(key))
    End If
  Else  ' check the int keys
    'Test if item exists
    If VarType(m_Options.Item(Int(key))) = vbObject Then
      'Return an object
      Set FoundItem = m_Options.Item(Int(key))
    Else
      'Return an standard variable
      FoundItem = m_Options.Item(Int(key))
    End If
  End If
  
  'Return success
  KeyExists = True
  Exit Function
    
ErrNotExist:
  KeyExists = False
  On Error GoTo 0
End Function
