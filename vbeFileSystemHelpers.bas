Attribute VB_Name = "vbeFileSystemHelpers"
Option Explicit
Option Private Module


Private Const cPATH_SEPARATOR As String = "\"
Private Const cEXTENSION_SEPARATOR As String = "."

' -------------------------------
' Filename manipulation functions
' -------------------------------
' @TODO: Adjust the parse methods to handle .gitignore

Public Function vbeParsePath(Filename As String) As String
' @param FileName [string]
' @return [string] Path portion of the filename (i.e. everything before the \)
  
  Dim slash_pos As Long
  
  slash_pos = InStrRev(Filename, cPATH_SEPARATOR)
  
  If slash_pos > 0 Then
    vbeParsePath = Left(Filename, slash_pos)
  Else
    vbeParsePath = vbNullString
  End If
End Function ' vbeParsePath

Public Function vbeParseBaseFilename(Filename As String) As String
' @param FileName [string]
' @return [string] filename without path or extension

  Dim slash_pos As Long, ext_pos As Long
  
  ' handle a missing module name
  slash_pos = InStrRev(Filename, cPATH_SEPARATOR)
  ext_pos = InStrRev(Filename, cEXTENSION_SEPARATOR)
  
  If ext_pos > 1 Then
    vbeParseBaseFilename = Mid(Filename, slash_pos + 1, ext_pos - slash_pos - 1)
  Else
    vbeParseBaseFilename = Right(Filename, Len(Filename) - slash_pos)
  End If
End Function ' vbeParseBaseFilename

Public Function vbeParseExtension(Filename As String) As String
' @param FileName [string]
' @return [string] file extension (i.e. everything to the right of the last .)
  Dim ext_pos As Long
  ext_pos = InStrRev(Filename, cEXTENSION_SEPARATOR)
  
  If ext_pos > 1 Then
    vbeParseExtension = Right(Filename, Len(Filename) - ext_pos)
  Else
    vbeParseExtension = vbNullString
  End If
End Function

Public Function vbeStandardizePath(FolderName As String) As String
' Adjusts the path so that it always ends with a \
' @param FolderName [string]
' @return [string] Path name with a consistent ending caracter

  If StrComp(Right(FolderName, 1), cPATH_SEPARATOR, vbBinaryCompare) = 0 Then
    vbeStandardizePath = FolderName
  Else
    vbeStandardizePath = FolderName & cPATH_SEPARATOR
  End If
End Function ' vbeStandardizePath

Public Function vbeFileNameFromModule( _
                   Module As Object, _
                   Optional Filename As String) _
                   As String
' Builds a filename from a Module object's information
' @param Module [CodeModule]
' @optparam Filename [string] overrides the module's default settings
' @return [string]
' @TODO: give this a more meaningful name

  Dim extension As String, fname As String
  
  extension = vbeGetFileExtension(vbcomp:=Module)
  If Trim(Filename) = vbNullString Then ' filename != blank
    fname = Module.Name & extension
  Else
    fname = Filename
    If InStr(1, fname, ".", vbBinaryCompare) = 0 Then ' filename doesn't have an extension
        fname = fname & extension
    End If
  End If
  
  vbeFileNameFromModule = fname
End Function ' vbeFileNameFromModule

Public Function vbeGetFileExtension(vbcomp As Object) As String
' This returns the appropriate file extension based on the Type of _
' the VBComponent
' @param vbcomp [VBComponent]
' @return [string]

    Select Case vbcomp.Type
        Case 2 ' 2 = vbext_ct_ClassModule
            vbeGetFileExtension = ".cls"
        Case 100 ' 100 = vbext_ct_Document
            vbeGetFileExtension = ".cls"
        Case 3 ' 3 = vbext_ct_MSForm
            vbeGetFileExtension = ".frm"
        Case 1 ' 1 = vbext_ct_StdModule
            vbeGetFileExtension = ".bas"
        Case Else
            vbeGetFileExtension = ".bas"
    End Select
    
End Function ' vbeGetFileExtension

Private Function vbeIsValidFileExtension(Filename As String) As Boolean
  Dim ExtPos As Long, ext As String
  
  ExtPos = InStrRev(Filename, ".")
  ext = Right(Filename, Len(Filename) - ExtPos)
  
  If ext = "bas" Or ext = "cls" Or ext = "frm" Then
    vbeIsValidFileExtension = True
  Else
    vbeIsValidFileExtension = False
  End If
End Function ' vbeIsValidFileExtension



