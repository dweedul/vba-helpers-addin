Attribute VB_Name = "vbeGeneral"
Option Explicit

Public Function vbeWorkbookFromProject( _
                  VBProject As Variant) As Workbook
' @param VBProject [VBProject]
' @return [workbook|nothing]

  Dim fname As String, wbs As Workbooks
  
  Set wbs = Application.Workbooks
  
  On Error GoTo Local_Error
  
  fname = VBProject.Filename
  fname = vbeParseBaseFilename(fname) & "." & vbeParseExtension(fname)
  
  Set vbeWorkbookFromProject = Application.Workbooks(fname)
  
  On Error GoTo 0
  Exit Function
  
Local_Error:
  Set vbeWorkbookFromProject = Nothing
End Function
