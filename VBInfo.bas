Attribute VB_Name = "VBInfo"
' #NoExport
' #NoRefresh
' #NoList

Option Explicit

'**************************************************************
' This UDT will hold the options that govern Export behavior
' Options must be in a comment and before any non-comment code.
' The whole option line will be read and processed accordingly.
'**************************************************************
Private Type ImportExportOptions
  NoList As Boolean
End Type ' ImportExportOptions

'*****************************************
' These constants hold the option strings.
'*****************************************
Private Const OPTIONS_TOKEN As String = "#"
Private Const OPTIONS_ASSIGNMENT_TOKEN As String = "="
Private Const OPTION_NO_LIST As String = "NoList"

Public Sub OutputVBAModuleListToSelectedCell()
'**************************************************************
' Lists the modules in the active workbook's VBA project
' to a group of cells starting with the curently selected cell.
'**************************************************************
  Dim c As Range, list As Variant
  If TypeName(Selection) <> "Range" Then Exit Sub
  
  Set c = Selection
  list = ListVBAModules(ThisWorkbook.VBProject, False)
  
  Set c = c.Resize(UBound(list))
  
  c.Value = WorksheetFunction.Transpose(list)
End Sub

Public Function ListVBAModules(VBProject As Object, _
                               Optional ListEmpty As Boolean = True) As Variant()
'**************************************
' Returns an array of module names from
' the current workbook's VBA project
'**************************************
  Dim out() As Variant, i As Long
  Dim col As Collection: Set col = New Collection
  
  ' add the components to the collection
  With VBProject.VBComponents
    For i = 1 To .Count
      If ListEmpty Or .Item(i).CodeModule.CountOfLines < 1 Or ParseOptions(.Item(i)).NoList Then
      Else
        col.Add .Item(i)
      End If
    Next ' i
  End With ' VBProject.VBComponents
  
  ReDim out(1 To col.Count)
  
  ' output the names to an array
  For i = 1 To col.Count
    out(i) = col(i).Name
  Next ' i
  
  ListVBAModules = out
  
  Set col = Nothing
End Function

'********************
'* Helper functions *
'********************

Private Function ParseOptions(vbcomp As Object) As ImportExportOptions
'*******************************************************
' Reads through any comments at the top of the code
' module, then parses the options out of those comments.
' Returns a UDT with the options ready for use
'*******************************************************
  Dim i As Long, tmp As String
  Dim equal_pos As Long, sep_pos As Long, off As Long
  Dim Var As String, Val As String
  Dim opt As ImportExportOptions
  
  Const comment_string As String = "'"
  
  ' initialize options to default values
  ' and prepare for an early exit if needed
  ProcessOption opt
  ParseOptions = opt
  
  With vbcomp.CodeModule
    '******************************************************
    ' Loop through the lines looking for options to process
    '******************************************************
    For i = 1 To .CountOfLines
      ' Get the current line
      tmp = .Lines(i, 1)
 
      ' Process all non-blank comment lines
      If Len(Trim(tmp)) > 1 Then
        ' Exit early once the comments are finished
        If Left(LTrim(tmp), 1) <> "'" Then Exit For
        
        ' find the position of the separators used
        off = 0
        If Len(OPTIONS_TOKEN) > 1 Then off = Len(OPTIONS_TOKEN)
        sep_pos = InStr(2, tmp, OPTIONS_TOKEN, vbTextCompare) + off
        
        off = 0
        If Len(OPTIONS_ASSIGNMENT_TOKEN) > 1 Then off = Len(OPTIONS_ASSIGNMENT_TOKEN)
        equal_pos = InStr(2, tmp, OPTIONS_ASSIGNMENT_TOKEN, vbTextCompare) + off
        
        ' get the options and arguments
        If equal_pos < 1 Then
          ' * single word options *
          Var = Trim(Mid(tmp, sep_pos + 1))
          Val = vbNullString
        Else
          ' * multi-word options *
          ' get the option and its value
          Var = Trim(Mid(tmp, sep_pos + 1, equal_pos - sep_pos - 1))
          Val = Trim(Mid(tmp, equal_pos + 1, Len(tmp) - equal_pos))
        End If
        
        ' save the variables into the UDT
        ProcessOption opt, Var, Val
      End If
    Next ' i
  End With ' vbcomp.CodeModule
  
  ParseOptions = opt
End Function

Private Sub ProcessOption( _
             ByRef Options As ImportExportOptions, _
             Optional Var As String = vbNullString, _
             Optional Val As String = vbNullString)
'* Modifies the Options object depending on what is provided.
'* This function is meant to be changed when options change.

  ' Initialize the settings if no inputs given
  If Var = vbNullString Then
    With Options
      .NoList = False
    End With ' out
    Exit Sub
  End If
  
  ' Handle the inputs
  Select Case LCase(Var)
    Case LCase(OPTION_NO_LIST):
      Options.NoList = True
  End Select ' var
End Sub

