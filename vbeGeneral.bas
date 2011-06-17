Attribute VB_Name = "vbeGeneral"
Option Explicit
Option Private Module

Public Function vbeWorkbookFromProject( _
                  VBProject As Object) _
                  As Workbook
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

Public Sub vbeDeleteVBProject(VBProject As Object)
' Deletes the VBProject by saving as a file that cannot include VBA

  Dim fmt As XlFileFormat
  Dim wb As Workbook
  Dim dsp As Boolean
  
  dsp = Application.DisplayAlerts
  Application.DisplayAlerts = False
  
  On Error GoTo Local_Error

  ' get the currently selected workbook from the vbproject
  Set wb = vbeWorkbookFromProject(VBProject)
  
  If wb Is Nothing Then GoTo Local_Error
  
  ' store the filetype of the current file
  fmt = wb.FileFormat
  fname = wb.FullName
  
  ' save the file as an xslx file
  wb.SaveAs fname & vbeSTRIPPED_FILE_SUFFIX, xlOpenXMLWorkbook
  
  ' close and reopen to ensure that the code is gone
  wb.Close
  Set wb = Workbooks.Open(fname & vbeSTRIPPED_FILE_SUFFIX)
 
  ' save again as the original
  wb.SaveAs fname, fmt
  
  ' delete the temp file

Local_Error:
  If Err.Number <> 0 Then Debug.Print Err.Description
  Application.DisplayAlerts = dsp
End Sub

