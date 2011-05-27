Attribute VB_Name = "vbeCodeExport"
Option Explicit

' ----------------
' Public functions
' ----------------
Public Sub vbeExportActiveVBProject(Optional HideMe As Boolean)
' @param HideMe [boolean] removes this sub from the macros menu

  Dim o As Object
  Set o = Application.VBE.ActiveVBProject
  ExportVBProject o, vbeParsePath(o.Filename)
End Sub ' vbeExportActiveVBProject

Public Sub vbeExportSelectedCodeModule(Optional HideMe As Boolean)
' @param HideMe [boolean] removes this sub from the macros menu

  Dim o As Object
  Set o = Application.VBE.SelectedVBComponent
  
  ExportVBComponent o, _
                    vbeParsePath(o.Collection.Parent.Filename)
End Sub ' vbeExportSelectedCodeModule

' ----------------
' Export functions
' ----------------
Private Function ExportVBProject(VBProj As Object, _
                  ByVal FolderName As String, _
                  Optional OverwriteExisting As Boolean = True) _
                  As Boolean
' This function exports all the code modules of a given VBProject _
' to text files. Default filenames will be used.
' @param FolderName [string]
' @optparam OverwriteExisting [boolean]
' @return [bool] false on error

  Dim VBComp As Object
  
  On Error GoTo Local_Error
  
  For Each VBComp In VBProj.VBComponents
    ExportVBComponent VBComp:=VBComp, _
                      FolderName:=FolderName, _
                      OverwriteExisting:=OverwriteExisting
  Next ' vbcomp
  
  ExportVBProject = True
  On Error GoTo 0
  Exit Function
  
Local_Error:
  ExportVBProject = False
End Function ' ExportVBProject

Private Function ExportVBComponent(VBComp As Object, _
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
  Set cm.VBComponent = VBComp
  
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
    FolderName = vbeStandardizePath(FolderName) & cm.Options(OPTION_RELATIVE_PATH)
  End If
  
  fname = vbeFileNameFromModule(VBComp, Filename)
  
  ' create the directory if it doesn't exist
  If Dir(FolderName, vbDirectory) = vbNullString Then
    MkDir FolderName
  End If
  
  fname = vbeStandardizePath(FolderName) & fname
  
  If Dir(fname, vbNormal + vbHidden + vbSystem) <> vbNullString Then
    If OverwriteExisting = True Then
      Kill fname
    Else
      ExportVBComponent = False
      Exit Function
    End If
  End If
  
  VBComp.Export Filename:=fname
  
  ExportVBComponent = True
  On Error GoTo 0
  Exit Function

Local_Error:
  ExportVBComponent = False
End Function ' ExportVBComponent


