VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "dbCodeModule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*****************************************************************
' This object gathers the information from the code modules needed
' to perform the actions required by the other modules.
' This object will be a member of a collection.
'*****************************************************************

'***************
' Public members
'***************
Private m_Options As Collection
Private m_Index As Long ' index in the collection
Private m_ModuleName As String ' name of the Code Module
Private m_ProjectName As String ' name of the Code Module
Private m_IsEmpty As Boolean ' tells if the code module is empty

Private m_Parsed As Boolean ' flag to tell if the CodeModule has been checked for options

'***********
' Properties
'***********
Public Property Get Index() As Long: Index = m_Index: End Property
Public Property Let Index(Index As Long): m_Index = Index: End Property

Public Property Get ProjectName() As String: ProjectName = m_ProjectName: End Property
Public Property Let ProjectName(ProjectName As String): m_ProjectName = ProjectName: End Property

Public Property Get ModuleName() As String: ModuleName = m_ModuleName: End Property
Public Property Let ModuleName(ModuleName As String): m_ModuleName = ModuleName: End Property

Public Property Get IsEmpty() As Boolean
  If Workbooks(ProjectName).VBProject.VBComponents(ModuleName).CodeModule.CountOfLines < 1 Then
    IsEmpty = True
  Else
    IsEmpty = False
  End If
End Property

Public Property Get Options(Optional Key As Variant) As Variant
' will default to a False if the key is not found

  Dim out As udtOption
  
  If Not m_Parsed Then ParseCodeModule
  
  If IsMissing(Key) Then
    Set Options = m_Options
  ElseIf KeyExists(Key, Options) Then
  Else
    Set out = New udtOption
    out.Variable = Key
    out.Value = False
    Set Options = out
  End If
End Property


' @TODO(DB): Be sure to check that the key exists in the collection when asking for an option

'***********************
' Constructor/Destructor
'***********************
Private Sub Class_Initialize()
  Set m_Options = New Collection
End Sub
Private Sub Class_Terminate()
  Set m_Options = Nothing
End Sub

'****************
' Private methods
'****************
Private Sub ParseCodeModule()
  Dim i As Long
  Dim vbcomp As Object
  Dim opt As udtOption
  Dim tmp As String
  
  If IsEmpty Then Exit Sub
  
  Set opt = New udtOption
  
  Set vbcomp = Workbooks(ProjectName).VBProject.VBComponents(ModuleName)
  
  With vbcomp.CodeModule
  ' Loop through the lines looking for options to process

    For i = 1 To .CountOfLines
      ' Get the current line
      tmp = .Lines(i, 1)
  
      ' Process all non-blank comment lines
      If Len(Trim(tmp)) > 1 Then
        ' Exit early once the comments are finished
        If Not Left(LTrim(tmp), 1) Like COMMENT_TOKEN Then Exit For
        
        ' save the variables into the collection
        ProcessOption ParseOptionLine(tmp)
      End If
    Next ' i
  End With ' vbcomp.CodeModule
  
  m_Parsed = True
End Sub

Private Function ParseOptionLine(OptionLine As String) As udtOption
  Dim equal_pos As Long, sep_pos As Long, off As Long
  Dim var As String, val As String
  Dim opt As udtOption ' output variable
  
  Set opt = New udtOption
  
  ' find the position of the separators used
  off = 0
  If Len(OPTIONS_TOKEN) > 1 Then off = Len(OPTIONS_TOKEN)
  sep_pos = InStr(2, OptionLine, OPTIONS_TOKEN, vbTextCompare) + off
  
  off = 0
  If Len(OPTIONS_ASSIGNMENT_TOKEN) > 1 Then off = Len(OPTIONS_ASSIGNMENT_TOKEN)
  equal_pos = InStr(2, OptionLine, OPTIONS_ASSIGNMENT_TOKEN, vbTextCompare) + off
  
  ' get the options and arguments
  If equal_pos < 1 Then
    ' * single word options *
    opt.Variable = Trim(Mid(OptionLine, sep_pos + 1))
    opt.Value = vbNullString
  Else
    ' * multi-word options *
    ' get the option and its value
    opt.Variable = Trim(Mid(OptionLine, sep_pos + 1, equal_pos - sep_pos - 1))
    opt.Value = Trim(Mid(OptionLine, equal_pos + 1, Len(OptionLine) - equal_pos))
  End If
  
  Set ParseOptionLine = opt
End Function


Private Sub ProcessOption(opt As udtOption)
' Adds to the Options collection depending on what is provided.
  
  ' Remove the option if it already exists
  If KeyExists(opt.Variable) Then m_Options.Remove opt.Variable
  
  If IsBooleanOption(opt.Variable) Then opt.Value = True
  
  m_Options.Add opt, opt.Variable
  
End Sub

Private Function IsBooleanOption(OptionString As String) As Boolean
  If InStr(1, BOOLEAN_OPTIONS, OptionString, vbTextCompare) > 0 Then
    IsBooleanOption = True
  End If
End Function


Private Function KeyExists( _
                  Key As Variant, _
                  Optional FoundItem As Variant) _
                  As Boolean
' Check if item exists in the collection by key or by position

  On Error GoTo ErrNotExist
   
  If VarType(Key) = vbString Then ' check the string keys
    'Test if item exists
    If VarType(m_Options.Item(CStr(Key))) = vbObject Then
      'Return an object
      Set FoundItem = m_Options.Item(CStr(Key))
    Else
      'Return an standard variable
      FoundItem = m_Options.Item(CStr(Key))
    End If
  Else  ' check the int keys
    'Test if item exists
    If VarType(m_Options.Item(Int(Key))) = vbObject Then
      'Return an object
      Set FoundItem = m_Options.Item(Int(Key))
    Else
      'Return an standard variable
      FoundItem = m_Options.Item(Int(Key))
    End If
  End If
  
  'Return success
  KeyExists = True
  Exit Function
    
ErrNotExist:
  KeyExists = False
  On Error GoTo 0
End Function

