Attribute VB_Name = "vbeExportImport"
Option Explicit

Private Const cPATH_SEPARATOR = "\"

Public Sub vbeExportSelectedCodeModule(Optional HideMe As Boolean)
' The HideMe removes this sub from the macros menu
  
  Dim o As Object
  Set o = Application.VBE.SelectedVBComponent
  
  ExportVBComponent o, _
                    PathFromFileName(o.Collection.Parent.FileName)
End Sub

'****************
' Private methods
'****************

Private Function ExportVBComponent(vbcomp As Object, _
                  ByVal FolderName As String, _
                  Optional ByVal FileName As String, _
                  Optional OverwriteExisting As Boolean = True) As Boolean
'*****************************************************************
' This function exports the code module of a VBComponent to a text
' file. If FileName is missing, the code will be exported to
' a file with the same name as the VBComponent followed by the
' appropriate extension.
'*****************************************************************
  Dim Extension As String, FName As String
  Dim cm As vbeVBComponent
  
  Set cm = New vbeVBComponent
  Set cm.VBComponent = vbcomp
  
  ' Don't export empty modules, it is stupid '
  If cm.IsEmpty Then Exit Function
  
  '*********************************
  ' Handle options within the module
  '*********************************
  ' exit early on NoExport option '
  If Not IsEmpty(cm.Options(OPTION_NO_EXPORT)) Then
    ExportVBComponent = False
    Exit Function
  End If
  
  ' add a relative path if provided
  If Not IsEmpty(cm.Options(OPTION_RELATIVE_PATH)) Then
    FolderName = FolderName & "\" & cm.Options(OPTION_RELATIVE_PATH)
  End If
  
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
  
  vbcomp.Export FileName:=FName
  ExportVBComponent = True
End Function

Private Function PathFromFileName(FileName As String)
  PathFromFileName = Left(FileName, InStrRev(FileName, cPATH_SEPARATOR))
End Function

Private Function GetFileExtension(vbcomp As Object) As String
'*****************************************************************
' This returns the appropriate file extension based on the Type of
' the VBComponent.
'*****************************************************************
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
'**********************************************************
' Returns true if the the file extension is bas, cls or frm
'**********************************************************
  Dim ExtPos As Long, Ext As String
  
  ExtPos = InStrRev(FileName, ".")
  Ext = Right(FileName, Len(FileName) - ExtPos)
  
  If Ext = "bas" Or Ext = "cls" Or Ext = "frm" Then
    IsValidFileExtension = True
  Else
    IsValidFileExtension = False
  End If
End Function
