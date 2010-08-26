Attribute VB_Name = "VBExportImport"
' _OPTION: EXCLUDE_ME
' _OPTION: NO_REFRESH

Option Explicit

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' This UDT will hold the options that govern Export behavior '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Type ImportExportOptions
  NoExport As Boolean
  NoRefresh As Boolean
  AbsolutePath As String
  RelativePath As String
End Type ' ImportExportOptions

Private Const OPTIONS_TOKEN As String = "_OPTION:"
Private Const OPTIONS_ASSIGNMENT_TOKEN As String = "="
Private Const OPTION_NO_EXPORT As String = "EXCLUDE_ME"
Private Const OPTION_RELATIVE_PATH As String = "RELATIVE_PATH"
Private Const OPTION_ABSOLUTE_PATH As String = "ABSOLUTE_PATH"
Private Const OPTION_NO_REFRESH As String = "NO_REFRESH"

Public Sub ExportAllVBAToWorkingDirectory()
  ExportAllVBA Workbook:=ThisWorkbook, FolderName:=ThisWorkbook.Path
End Sub

Public Sub ImportAllVBAFromWorkingDirectory()
  ImportAllVBAFromFolder ThisWorkbook.Path, True
End Sub

Public Sub OutputVBAModuleListToSelectedCell()
' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Lists the modules in the active workbook's VBA project        '
' to a group of cells starting with the curently selected cell. '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  Dim c As Range, list As Variant
  If TypeName(Selection) <> "Range" Then Exit Sub
  
  Set c = Selection
  list = ListVBAModules(ThisWorkbook.VBProject)
  
  Set c = c.Resize(UBound(list))
  
  c.Value = WorksheetFunction.Transpose(list)
End Sub

Public Function ListVBAModules(VBProject As Object) As Variant()
''''''''''''''''''''''''''''''''''''''''''
' Returns an array of module names       '
' from the current workbooks VBA project '
''''''''''''''''''''''''''''''''''''''''''
  Dim out() As Variant, i As Long
  
  With VBProject.VBComponents
    ReDim out(1 To .Count)
    For i = 1 To .Count
      out(i) = .Item(i).Name
    Next ' i
  End With ' VBProject.VBComponents
  
  ListVBAModules = out
End Function

Public Sub ImportAllVBAFromFolder(Path As String, _
                                  Optional RecurseSubfolders As Boolean = False)
  Dim fso As Object, Folder As Object, f As Object, fol As Object
  
  Set fso = CreateObject("Scripting.FileSystemObject")
  Set Folder = fso.GetFolder(Path)
  
  If RecurseSubfolders Then
    For Each fol In Folder.SubFolders
      If Left(Trim(fol.Name), 1) <> "." Then
        ImportAllVBAFromFolder fol.Path, True
      End If
    Next ' fol
  End If
  
  For Each f In Folder.Files
    If IsValidFileExtension(FileName:=f.Name) Then
      'MsgBox f.Path & vbCrLf & f.Type
      ImportVBComponent ThisWorkbook.VBProject, f.Path
    End If
  Next ' f
End Sub

''''''''''''''''''''''''''''''''''
' Helper Functions for the Above '
''''''''''''''''''''''''''''''''''
Public Sub ExportAllVBA(Workbook As Workbook, _
                        FolderName As String)
  Dim vbcomp As Object

  For Each vbcomp In Workbook.VBProject.VBComponents
    ExportVBComponent vbcomp:=vbcomp, FolderName:=FolderName
  Next ' VBComp
End Sub

Public Sub ExportList(Workbook As Workbook, _
                        FolderName As String, _
                        ModuleList As Variant)
  Dim m As Variant, m_list As Variant
  
  '''''''''''''''''''''''''''''''
  ' Find the type of ModuleList '
  '''''''''''''''''''''''''''''''
  If TypeName(ModuleList) = "String" Then
    m_list = Split(ModuleList)
  ElseIf TypeName(ModuleList) = "Range" Then
    m_list = ModuleList.Value
  ElseIf IsArray(ModuleList) Then
    m_list = ModuleList
  Else
    Exit Sub
  End If
  
  On Error Resume Next
  For Each m In m_list
    ExportVBComponent vbcomp:=ThisWorkbook.VBProject.VBComponents(m), FolderName:=FolderName
  Next ' m
  On Error GoTo 0
End Sub

Private Function ExportVBComponent(vbcomp As Object, _
                ByVal FolderName As String, _
                Optional ByVal FileName As String, _
                Optional OverwriteExisting As Boolean = True) As Boolean
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' This function exports the code module of a VBComponent to a text
' file. If FileName is missing, the code will be exported to
' a file with the same name as the VBComponent followed by the
' appropriate extension.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  Dim Extension As String, FName As String ', Options As ImportExportOptions
  
  ' Don't export empty modules, it is stupid '
  If vbcomp.CodeModule.CountOfLines = 0 Then Exit Function
  
  ''''''''''''''''''''''''''''''''''''
  ' Handle options within the module '
  ''''''''''''''''''''''''''''''''''''
  With ParseOptions(vbcomp)
    
    ' exit early on excluded option '
    If .NoExport Then
      ExportVBComponent = False
      Exit Function
    End If
    
    ' add a relative path if provided
    If .RelativePath <> vbNullString Then
      FolderName = FolderName & "\" & .RelativePath
    End If
  
  End With
  
  Extension = GetFileExtension(vbcomp:=vbcomp)
  If Trim(FileName) = vbNullString Then ' filename != blank
    FName = vbcomp.Name & Extension
  Else
    FName = FileName
    If InStr(1, FName, ".", vbBinaryCompare) = 0 Then ' filename doesn't have an extension
        FName = FName & Extension
    End If
  End If
  
  ' create the directory if it doesn't exist
  If Dir(FolderName, vbDirectory) = vbNullString Then
    MkDir FolderName
  End If
  
  If StrComp(Right(FolderName, 1), "\", vbBinaryCompare) = 0 Then
    FName = FolderName & FName
  Else
    FName = FolderName & "\" & FName
  End If
  
  If Dir(FName, vbNormal + vbHidden + vbSystem) <> vbNullString Then
    If OverwriteExisting = True Then
      Kill FName
    Else
      ExportVBComponent = False
      Exit Function
    End If
  End If
  
  vbcomp.EXPORT FileName:=FName
  ExportVBComponent = True
End Function

Private Function ImportVBComponent(VBProject As Object, _
                  FileName As String, _
                  Optional ModuleName As String, _
                  Optional OverwriteExisting As Boolean = True) _
                  As Boolean
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' This function imports the code module of a VBComponent to a text
' file. If ModuleName is missing, the code will be imported to
' a module with the same name as the filename without the extension
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  Dim vbcomp As Object, TempVBComp As Object, s As String
  Dim SlashPos As Long, ExtPos As Long, opt As ImportExportOptions
  
  On Error Resume Next
  
  ' handle a missing module name
  If ModuleName = vbNullString Then
    SlashPos = InStrRev(FileName, "\")
    ExtPos = InStrRev(FileName, ".")
    ModuleName = Mid(FileName, SlashPos + 1, ExtPos - SlashPos - 1)
  End If
  
  '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  ' check if module exists, then check the import options '
  '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  Set vbcomp = Nothing
  Set vbcomp = VBProject.VBComponents(ModuleName)
  
  If Not vbcomp Is Nothing Then
    With ParseOptions(vbcomp)
      
      ' exit early on NoRefresh
      If .NoRefresh Then
        ImportVBComponent = False
        Exit Function
      End If
      
    End With ' ParseOptions(vbcomp)
  End If
  
  If OverwriteExisting = True Then
    '''''''''''''''''''''''''''''''''''''
    ' If OverwriteExisting is True, Kill
    ' the existing temp file and remove
    ' the existing VBComponent from the
    ' ToVBProject.
    '''''''''''''''''''''''''''''''''''''
    With VBProject.VBComponents
      .Remove .Item(ModuleName)
    End With
  Else
    '''''''''''''''''''''''''''''''''''''''''
    ' OverwriteExisting is False. If there is
    ' already a VBComponent named ModuleName,
    ' exit with a return code of False.
    '''''''''''''''''''''''''''''''''''''''''
    Err.Clear
    Set vbcomp = VBProject.VBComponents(ModuleName)
    If Err.Number <> 0 Then
      If Err.Number = 9 Then
        ' module doesn't exist. ignore error.
      Else
        ' other error. get out with return value of False
        ImportVBComponent = False
        Exit Function
      End If
    End If
  End If
  
  '''''''''''''''''''''''''''''''''''''''''''''''
  ' Document modules (SheetX and ThisWorkbook)
  ' cannot be removed. So, if we are working with
  ' a document object, delete all code in that
  ' component and add the lines of FName
  ' back in to the module.
  '''''''''''''''''''''''''''''''''''''''''''''''
  Set vbcomp = Nothing
  Set vbcomp = VBProject.VBComponents(ModuleName)
  
  If vbcomp Is Nothing Then
    VBProject.VBComponents.Import FileName:=FileName
  Else
    If vbcomp.Type = 100 Then ' 100 = vbext_ct_Document
      ' VBComp is destination module
      Set TempVBComp = VBProject.VBComponents.Import(FileName)
      ' TempVBComp is source module
      With vbcomp.CodeModule
        .DeleteLines 1, .CountOfLines
        s = TempVBComp.CodeModule.Lines(1, TempVBComp.CodeModule.CountOfLines)
        .InsertLines 1, s
      End With
      On Error GoTo 0
      VBProject.VBComponents.Remove TempVBComp
    End If
  End If
  
  ImportVBComponent = True
End Function

Private Function GetFileExtension(vbcomp As Object) As String
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' This returns the appropriate file extension based on the Type of
' the VBComponent.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Select Case vbcomp.Type
        Case 2 ' 2 = vbext_ct_ClassModule
            GetFileExtension = ".cls"
        Case 100 ' 100 = vbext_ct_Document
            GetFileExtension = ".cls"
        Case 3 ' 3 = vbext_ct_MSForm
            GetFileExtension = ".frm"
        Case 1 ' 1 = vbext_ct_StdModule
            GetFileExtension = ".bas"
        Case Else
            GetFileExtension = ".bas"
    End Select
    
End Function

Private Function IsValidFileExtension(FileName As String) As String
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Returns true if the the file extension is bas, cls or frm '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  Dim ExtPos As Long, Ext As String
  
  ExtPos = InStrRev(FileName, ".")
  Ext = Right(FileName, Len(FileName) - ExtPos)
  
  If Ext = "bas" Or Ext = "cls" Or Ext = "frm" Then
    IsValidFileExtension = True
  Else
    IsValidFileExtension = False
  End If
End Function


Private Function ParseOptions(vbcomp As Object) As ImportExportOptions
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Reads through any comments at the top of the code module '
' then parses the options out of the comments.             '
' Returns a type UDT with the options ready for use        '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  Dim i As Long, tmp As String
  Dim equal_pos As Long, sep_pos As Long
  Dim var As String, val As String
  Dim opt As ImportExportOptions
  
  Const comment_string As String = "'"
  
  '''''''''''''''''''''''''''''''''''''''''''
  ' initialize options to default values    '
  ' and prepare for an early exit if needed '
  '''''''''''''''''''''''''''''''''''''''''''
  With opt
    .AbsolutePath = vbNullString
    .NoExport = False
    .RelativePath = vbNullString
    .NoRefresh = False
  End With ' opt
  
  ParseOptions = opt
  
  With vbcomp.CodeModule
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Loop through the lines looking for options to process '
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    For i = 1 To .CountOfLines
      ' Get the current line
      tmp = .Lines(i, 1)
 
      ' Process all non-blank comment lines
      If Len(tmp) > 1 Then
        ' Exit early once the comments are finished
        If Left(LTrim(tmp), 1) <> "'" Then Exit For
        
        ' find the position of the separators used
        sep_pos = InStr(2, tmp, OPTIONS_TOKEN, vbTextCompare) + Len(OPTIONS_TOKEN)
        equal_pos = InStr(2, tmp, OPTIONS_ASSIGNMENT_TOKEN, vbTextCompare)
        
        ' get the options and arguments
        If equal_pos < 1 Then
          ' single word options '
          var = Trim(Mid(tmp, sep_pos + 1))
          val = vbNullString
        Else
          ' multi-word options '
          ' get the option and its value
          var = Trim(Mid(tmp, sep_pos + 1, equal_pos - sep_pos - 1))
          val = Trim(Mid(tmp, equal_pos + 1, Len(tmp) - equal_pos))
        End If
        
        ' save the variables into the UDT
        Select Case UCase(var)
          Case OPTION_NO_EXPORT:
            opt.NoExport = True
          Case OPTION_RELATIVE_PATH:
            opt.RelativePath = val
          Case OPTION_ABSOLUTE_PATH:
            opt.AbsolutePath = val
          Case OPTION_NO_REFRESH:
            opt.NoRefresh = True
        End Select ' var
      End If
    Next ' i
  End With ' vbcomp.CodeModule
  
  ParseOptions = opt
End Function


