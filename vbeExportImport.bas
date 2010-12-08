Attribute VB_Name = "vbeExportImport"
Option Explicit
Option Private Module

Private Const cPATH_SEPARATOR As String = "\"
Private Const cEXTENSION_SEPARATOR As String = "."

' ----------------
' Public functions
' ----------------

Public Sub vbeExportActiveVBProject(Optional HideMe As Boolean)
' @param HideMe [boolean] removes this sub from the macros menu

  Dim o As Object
  Set o = Application.VBE.ActiveVBProject
  ExportVBProject o, ParsePath(o.Filename)
End Sub ' vbeExportActiveVBProject

Public Sub vbeExportSelectedCodeModule(Optional HideMe As Boolean)
' @param HideMe [boolean] removes this sub from the macros menu

  Dim o As Object
  Set o = Application.VBE.SelectedVBComponent
  
  ExportVBComponent o, _
                    ParsePath(o.Collection.Parent.Filename)
End Sub ' vbeExportSelectedCodeModule

'    ' check the options
'    Set cm.VBComponent = VBProject.VBComponents(ModuleName)
'
'    If Not IsError(cm.Options(OPTION_NO_REFRESH)) Then
'      ' @TODO: this is a bad place for this
'      GoTo Local_Error
'    End If

Public Sub vbeRefreshSelectedCodeModule(Optional HideMe As Boolean)
' @param HideMe [boolean] removes this sub from the macros menu

  Dim o As VBComponent, fname As String, cm As vbeVBComponent
  Set o = Application.VBE.SelectedVBComponent
  
  Set cm = New vbeVBComponent
  Set cm.VBComponent = o
  
  If Not cm.Options(OPTION_NO_REFRESH) Then
    fname = StandardizePath(ParsePath(o.Collection.Parent.Filename)) & FileNameFromModule(o)
    
    If Dir(fname, vbNormal + vbHidden + vbSystem) <> vbNullString Then
      ImportVBComponent o.Collection.Parent, _
                        fname
    End If
  End If

End Sub ' vbeRefreshSelectedCodeModule`

' ----------------
' Export functions
' ----------------
Private Function ExportVBProject(vbproj As Object, _
                  ByVal FolderName As String, _
                  Optional OverwriteExisting As Boolean = True) _
                  As Boolean
' This function exports all the code modules of a given VBProject _
' to text files. Default filenames will be used.
' @param FolderName [string]
' @optparam OverwriteExisting [boolean]
' @return [bool] false on error

  Dim vbcomp As Object
  
  On Error GoTo Local_Error
  
  For Each vbcomp In vbproj.VBComponents
    ExportVBComponent vbcomp:=vbcomp, _
                      FolderName:=FolderName, _
                      OverwriteExisting:=OverwriteExisting
  Next ' vbcomp
  
  ExportVBProject = True
  On Error GoTo 0
  Exit Function
  
Local_Error:
  ExportVBProject = False
End Function ' ExportVBProject

Private Function ExportVBComponent(vbcomp As Object, _
                  ByVal FolderName As String, _
                  Optional ByVal Filename As String, _
                  Optional OverwriteExisting As Boolean = True) _
                  As Boolean
' This function exports the code module of a VBComponent to a text _
' file. If FileName is missing, the code will be exported to _
' a file with the same name as the VBComponent followed by the _
' appropriate extension.
' @param FolderName [string]
' @optparam FileName [string]
' @optparam OverwriteExisting [bool]
' @return [bool] False on error

  Dim fname As String
  Dim cm As vbeVBComponent
  
  On Error GoTo Local_Error
  
  Set cm = New vbeVBComponent
  Set cm.VBComponent = vbcomp
  
  ' Don't export empty modules, it is stupid '
  If cm.IsEmpty Then Exit Function
  
  '---------------------------------
  ' Handle options within the module
  '---------------------------------
  ' exit early on NoExport option
  If cm.Options(OPTION_NO_EXPORT) Then
    ExportVBComponent = False
    Exit Function
  End If
  
  ' add a relative path if provided
  If Len(cm.Options(OPTION_RELATIVE_PATH)) > 0 Then
    FolderName = StandardizePath(FolderName) & cm.Options(OPTION_RELATIVE_PATH)
  End If
  
  fname = FileNameFromModule(vbcomp, Filename)
  
  ' create the directory if it doesn't exist
  If Dir(FolderName, vbDirectory) = vbNullString Then
    MkDir FolderName
  End If
  
  fname = StandardizePath(FolderName) & fname
  
  If Dir(fname, vbNormal + vbHidden + vbSystem) <> vbNullString Then
    If OverwriteExisting = True Then
      Kill fname
    Else
      ExportVBComponent = False
      Exit Function
    End If
  End If
  
  vbcomp.Export Filename:=fname
  
  ExportVBComponent = True
  On Error GoTo 0
  Exit Function

Local_Error:
  ExportVBComponent = False
End Function ' ExportVBComponent

' ----------------
' Import Functions
' ----------------
Private Function ImportVBComponent( _
                  VBProject As Object, _
                  Filename As String, _
                  Optional ModuleName As String, _
                  Optional OverwriteExisting As Boolean = True) _
                  As Boolean
' This function imports the code module of a VBComponent from a text _
' file. If ModuleName is missing, the code will be imported to _
' a module with the same name as the filename without the extension _
' @param VBProject [VBProject]
' @param FileName [string]
' @optparam ModuleName [string]
' @optparam OverwriteExisting [bool]
' @return [bool] False on error

  Dim cm As vbeVBComponent
  
  On Error GoTo Local_Error
  
  ' handle a missing module name
  If ModuleName = vbNullString Then
    ModuleName = ParseBaseFilename(Filename)
  End If
  
  If VBComponentExists(ModuleName, VBProject) Then
    If OverwriteExisting Then
      DeleteModule VBProject, ModuleName
    Else
      GoTo Local_Error
    End If
  End If
  
  ImportFromFile VBProject, Filename, ModuleName
  
  ImportVBComponent = True
  On Error GoTo 0
  Exit Function
  
Local_Error:
  ImportVBComponent = False
End Function ' ImportVBComponent

Private Function ImportFromFile( _
                  VBProject As VBProject, _
                  Filename As String, _
                  ModuleName As String) _
                  As Boolean
' @param VBProject [VBProject]
' @param Filename [string]
' @param ModuleName [string]

  Dim vbcomp As VBComponent, tmp_vbcomp As VBComponent
  Dim s As String
  
  On Error GoTo Local_Error
  
  With VBProject.VBComponents
    If VBComponentExists(ModuleName, VBProject) Then
      If .Item(ModuleName).Type = 100 Then ' 100 =vbext_ct_Document
        Set tmp_vbcomp = .Import(Filename)
        s = tmp_vbcomp.CodeModule.Lines(1, tmp_vbcomp.CodeModule.CountOfLines)
        .Item(ModuleName).CodeModule.InsertLines 1, s
        .Remove tmp_vbcomp
      End If
    Else

    End If
    .Import Filename:=Filename
  End With
     
  ImportFromFile = True
  On Error GoTo 0
  Exit Function
  
Local_Error:
  ImportFromFile = False
End Function ' ImportFromFile

' -------------------------------
' Filename manipulation functions
' -------------------------------
' @TODO: Adjust the parse methods to handle .gitignore

Public Function ParsePath(Filename As String) As String
' @param FileName [string]
' @return [string] Path portion of the filename (i.e. everything before the \)
  
  Dim slash_pos As Long
  
  slash_pos = InStrRev(Filename, cPATH_SEPARATOR)
  
  If slash_pos > 0 Then
    ParsePath = Left(Filename, slash_pos)
  Else
    ParsePath = vbNullString
  End If
End Function ' ParsePath

Public Function ParseBaseFilename(Filename As String) As String
' @param FileName [string]
' @return [string] filename without path or extension

  Dim slash_pos As Long, ext_pos As Long
  
  ' handle a missing module name
  slash_pos = InStrRev(Filename, cPATH_SEPARATOR)
  ext_pos = InStrRev(Filename, cEXTENSION_SEPARATOR)
  
  If ext_pos > 1 Then
    ParseBaseFilename = Mid(Filename, slash_pos + 1, ext_pos - slash_pos - 1)
  Else
    ParseBaseFilename = Right(Filename, Len(Filename) - slash_pos)
  End If
End Function ' ParseBaseFilename

Public Function ParseExtension(Filename As String) As String
' @param FileName [string]
' @return [string] file extension (i.e. everything to the right of the last .)
  Dim ext_pos As Long
  ext_pos = InStrRev(Filename, cEXTENSION_SEPARATOR)
  
  If ext_pos > 1 Then
    ParseExtension = Right(Filename, Len(Filename) - ext_pos)
  Else
    ParseExtension = vbNullString
  End If
End Function

Private Function StandardizePath(FolderName As String) As String
' Adjusts the path so that it always ends with a \
' @param FolderName [string]
' @return [string] Path name with a consistent ending caracter

  If StrComp(Right(FolderName, 1), cPATH_SEPARATOR, vbBinaryCompare) = 0 Then
    StandardizePath = FolderName
  Else
    StandardizePath = FolderName & cPATH_SEPARATOR
  End If
End Function ' StandardizePath

Private Function FileNameFromModule( _
                   Module As Object, _
                   Optional Filename As String) _
                   As String
' Builds a filename from a Module object's information
' @param Module [CodeModule]
' @optparam Filename [string] overrides the module's default settings
' @return [string]
' @TODO: give this a more meaningful name

  Dim extension As String, fname As String
  
  extension = GetFileExtension(vbcomp:=Module)
  If Trim(Filename) = vbNullString Then ' filename != blank
    fname = Module.Name & extension
  Else
    fname = Filename
    If InStr(1, fname, ".", vbBinaryCompare) = 0 Then ' filename doesn't have an extension
        fname = fname & extension
    End If
  End If
  
  FileNameFromModule = fname
End Function ' FileNameFromModule


Private Function GetFileExtension(vbcomp As Object) As String
' This returns the appropriate file extension based on the Type of _
' the VBComponent
' @param vbcomp [VBComponent]
' @return [string]

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
    
End Function ' GetFileExtension

Private Function IsValidFileExtension(Filename As String) As Boolean
  Dim ExtPos As Long, Ext As String
  
  ExtPos = InStrRev(Filename, ".")
  Ext = Right(Filename, Len(Filename) - ExtPos)
  
  If Ext = "bas" Or Ext = "cls" Or Ext = "frm" Then
    IsValidFileExtension = True
  Else
    IsValidFileExtension = False
  End If
End Function ' IsValidFileExtension

