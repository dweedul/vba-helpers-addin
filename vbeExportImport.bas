Attribute VB_Name = "vbeExportImport"
Option Explicit
Option Private Module

Private Const cPATH_SEPARATOR = "\"

Public Sub vbeExportActiveVBProject(Optional HideMe As Boolean)
' The HideMe removes this sub from the macros menu
  Dim o As Object
  Set o = Application.VBE.ActiveVBProject
  ExportVBProject o, PathFromFileName(o.Filename)
End Sub

Public Sub vbeExportSelectedCodeModule(Optional HideMe As Boolean)
  Dim o As Object
  Set o = Application.VBE.SelectedVBComponent
  
  ExportVBComponent o, _
                    PathFromFileName(o.Collection.Parent.Filename)
End Sub

Public Sub vbeRefreshSelectedCodeModule(Optional HideMe As Boolean)
  Dim o As VBComponent, fname As String
  Set o = Application.VBE.SelectedVBComponent
  
  ' @TODO: implement some form of option checking for the paths here!
  
  fname = CleanFolderName(PathFromFileName(o.Collection.Parent.Filename)) & BuildFileName(o)
  
  If Dir(fname, vbNormal + vbHidden + vbSystem) <> vbNullString Then
    ImportVBComponent o.Collection.Parent, _
                      fname
  End If
  
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
                  Optional ByVal Filename As String, _
                  Optional OverwriteExisting As Boolean = True) As Boolean
' This function exports the code module of a VBComponent to a text
' file. If FileName is missing, the code will be exported to
' a file with the same name as the VBComponent followed by the
' appropriate extension.

  Dim fname As String
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
    FolderName = CleanFolderName(FolderName) & cm.Options(OPTION_RELATIVE_PATH)
  End If
  
  fname = BuildFileName(vbcomp, Filename)
  
  ' create the directory if it doesn't exist
  If Dir(FolderName, vbDirectory) = vbNullString Then
    MkDir FolderName
  End If
  
  fname = CleanFolderName(FolderName) & fname
  
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
End Function

Private Function ImportVBComponent(VBProject As Object, _
                  Filename As String, _
                  Optional ModuleName As String, _
                  Optional OverwriteExisting As Boolean = True) _
                  As Boolean
' This function imports the code module of a VBComponent to a text
' file. If ModuleName is missing, the code will be imported to
' a module with the same name as the filename without the extension

  Dim cm As vbeVBComponent
  
  On Error GoTo Local_Error
  
  Set cm = New vbeVBComponent
  
  ' handle a missing module name
  If ModuleName = vbNullString Then
    ModuleName = GetModuleNameFromFileName(Filename)
  End If
  
  If VBComponentExists(ModuleName, VBProject) Then
    ' check the options
    Set cm.VBComponent = VBProject.VBComponents(ModuleName)
    
    If Not IsEmpty(cm.Options(OPTION_NO_REFRESH)) Or Not OverwriteExisting Then
      GoTo Local_Error
    End If
  End If
  
  VBProject.VBComponents.Import Filename:=Filename
  
  ImportVBComponent = True
  On Error GoTo 0
  Exit Function
  
Local_Error:
  ImportVBComponent = False
End Function

Private Function PathFromFileName(Filename As String)
  PathFromFileName = Left(Filename, InStrRev(Filename, cPATH_SEPARATOR))
End Function

Private Function CleanFolderName(FolderName As String) As String
                   
  If StrComp(Right(FolderName, 1), "\", vbBinaryCompare) = 0 Then
    CleanFolderName = FolderName
  Else
    CleanFolderName = FolderName & "\"
  End If
End Function

Private Function BuildFileName( _
                   Module As Object, _
                   Optional Filename As String) _
                   As String
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
  
  BuildFileName = fname
End Function

Private Function GetModuleNameFromFileName( _
                   file_name As String) _
                   As String
  Dim slash_pos As Long, ext_pos As Long
  
  ' handle a missing module name
  slash_pos = InStrRev(file_name, "\")
  ext_pos = InStrRev(file_name, ".")
  GetModuleNameFromFileName = Mid(file_name, slash_pos + 1, ext_pos - slash_pos - 1)
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

Private Function IsValidFileExtension(Filename As String) As Boolean
  Dim ExtPos As Long, Ext As String
  
  ExtPos = InStrRev(Filename, ".")
  Ext = Right(Filename, Len(Filename) - ExtPos)
  
  If Ext = "bas" Or Ext = "cls" Or Ext = "frm" Then
    IsValidFileExtension = True
  Else
    IsValidFileExtension = False
  End If
End Function
