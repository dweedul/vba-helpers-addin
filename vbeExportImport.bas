Attribute VB_Name = "vbeExportImport"
Option Explicit
Option Private Module

Private Const cPATH_SEPARATOR = "\"

Public Sub vbeExportActiveVBProject(Optional HideMe As Boolean)
' The HideMe removes this sub from the macros menu
  Dim o As Object
  Set o = Application.VBE.ActiveVBProject
  ExportVBProject o, PathFromFileName(o.FileName)
End Sub

Public Sub vbeExportSelectedCodeModule(Optional HideMe As Boolean)
  Dim o As Object
  Set o = Application.VBE.SelectedVBComponent
  
  ExportVBComponent o, _
                    PathFromFileName(o.Collection.Parent.FileName)
End Sub

Public Sub vbeRefreshSelectedCodeModule(Optional HideMe As Boolean)
  Dim o As Object
  Set o = Application.VBE.SelectedVBComponent
  
  ImportVBComponent o, _
                    PathFromFileName(o.Collection.Parent.FileName)
End Sub

'----------------
' Private methods
'----------------
Private Function ExportVBProject(vbproj As Object, _
                  ByVal FolderName As String, _
                  Optional OverwriteExisting As Boolean = True) _
                  As Boolean
' This function exports all the code modules of a given VBProject
' to text files. Default filenames will be used.

  Dim vbcomp As Object
  For Each vbcomp In vbproj.VBComponents
    ExportVBComponent vbcomp:=vbcomp, _
                      FolderName:=FolderName, _
                      OverwriteExisting:=OverwriteExisting
  Next ' vbcomp
End Function

Private Function ExportVBComponent(vbcomp As Object, _
                  ByVal FolderName As String, _
                  Optional ByVal FileName As String, _
                  Optional OverwriteExisting As Boolean = True) As Boolean
' This function exports the code module of a VBComponent to a text
' file. If FileName is missing, the code will be exported to
' a file with the same name as the VBComponent followed by the
' appropriate extension.

  Dim extension As String, Fname As String
  Dim cm As vbeVBComponent
  
  Set cm = New vbeVBComponent
  Set cm.VBComponent = vbcomp
  
  ' Don't export empty modules, it is stupid '
  If cm.IsEmpty Then Exit Function
  
  '---------------------------------
  ' Handle options within the module
  '---------------------------------
  ' exit early on NoExport option
  If Not IsEmpty(cm.Options(OPTION_NO_EXPORT)) Then
    ExportVBComponent = False
    Exit Function
  End If
  
  ' add a relative path if provided
  If Not IsEmpty(cm.Options(OPTION_RELATIVE_PATH)) Then
    FolderName = FolderName & "\" & cm.Options(OPTION_RELATIVE_PATH)
  End If
  
  Fname = BuildFileName(vbcomp, FileName)
  
  ' create the directory if it doesn't exist
  If Dir(FolderName, vbDirectory) = vbNullString Then
    MkDir FolderName
  End If
  
  If StrComp(Right(FolderName, 1), "\", vbBinaryCompare) = 0 Then
    Fname = FolderName & Fname
  Else
    Fname = FolderName & "\" & Fname
  End If
  
  If Dir(Fname, vbNormal + vbHidden + vbSystem) <> vbNullString Then
    If OverwriteExisting = True Then
      Kill Fname
    Else
      ExportVBComponent = False
      Exit Function
    End If
  End If
  
  vbcomp.Export FileName:=Fname
  ExportVBComponent = True
End Function

Private Function ImportVBComponent(VBProject As Object, _
                  FileName As String, _
                  Optional ModuleName As String, _
                  Optional OverwriteExisting As Boolean = True) _
                  As Boolean
'******************************************************************
' This function imports the code module of a VBComponent to a text
' file. If ModuleName is missing, the code will be imported to
' a module with the same name as the filename without the extension
'******************************************************************
  Dim vbcomp As Object, TempVBComp As Object, s As String
  Dim SlashPos As Long, ExtPos As Long, opt As ImportExportOptions
  
  On Error Resume Next
  
  ' handle a missing module name
  If ModuleName = vbNullString Then
    SlashPos = InStrRev(FileName, "\")
    ExtPos = InStrRev(FileName, ".")
    ModuleName = Mid(FileName, SlashPos + 1, ExtPos - SlashPos - 1)
  End If
  
  '******************************************************
  ' check if module exists, then check the import options
  '******************************************************
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
    '***********************************
    ' If OverwriteExisting is True, Kill
    ' the existing temp file and remove
    ' the existing VBComponent from the
    ' ToVBProject.
    '***********************************
    With VBProject.VBComponents
      .Remove .Item(ModuleName)
    End With
  Else
    '****************************************
    ' OverwriteExisting is False. If there is
    ' already a VBComponent named ModuleName,
    ' exit with a return code of False.
    '****************************************
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
  
  '**********************************************
  ' Document modules (SheetX and ThisWorkbook)
  ' cannot be removed. So, if we are working with
  ' a document object, delete all code in that
  ' component and add the lines of FName
  ' back in to the module.
  '**********************************************
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

Private Function PathFromFileName(FileName As String)
  PathFromFileName = Left(FileName, InStrRev(FileName, cPATH_SEPARATOR))
End Function

Private Function BuildFileName( _
                   Module As Object, _
                   Optional FileName As String) _
                   As String
  Dim extension As String, Fname As String
  
  extension = GetFileExtension(vbcomp:=Module)
  If Trim(FileName) = vbNullString Then ' filename != blank
    Fname = Module.Name & extension
  Else
    Fname = FileName
    If InStr(1, Fname, ".", vbBinaryCompare) = 0 Then ' filename doesn't have an extension
        Fname = Fname & extension
    End If
  End If
  
  BuildFileName = Fname
End Function

Private Function GetFileExtension(vbcomp As Object) As String
' This returns the appropriate file extension based on the Type of
' the VBComponent
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

Private Function IsValidFileExtension(FileName As String) As Boolean
  Dim ExtPos As Long, Ext As String
  
  ExtPos = InStrRev(FileName, ".")
  Ext = Right(FileName, Len(FileName) - ExtPos)
  
  If Ext = "bas" Or Ext = "cls" Or Ext = "frm" Then
    IsValidFileExtension = True
  Else
    IsValidFileExtension = False
  End If
End Function
