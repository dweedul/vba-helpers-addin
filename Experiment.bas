Attribute VB_Name = "Experiment"
Public Sub vbeExperimental()
  ' Clear all the code from the active vb project
  
  Dim VBComp As VBComponent
  Dim cm As New vbeVBComponent

  'For Each vbcomp In Application.VBE.ActiveVBProject.VBComponents
    Set cm.VBComponent = Application.VBE.SelectedVBComponent
    
    If cm.Options.Count > 0 Then
      resp = MsgBox("The following options were found in the selected file:" & _
                    vbCrLf & _
                    vbCrLf & _
                    cm.OptionList & _
                    vbCrLf & _
                    "Would you like to proceed?", vbYesNoCancel, "VBEHelper Addin")
      If resp = vbYes Then
        MsgBox "DELETED"
      End If
    Else
      MsgBox "No options found in " & cm.VBComponent.Name
    End If
  'Next ' vbcomp
End Sub

Sub DeleteAllVBACode()
    Dim VBProj As VBIDE.VBProject
    Dim VBComp As VBIDE.VBComponent
    Dim CodeMod As VBIDE.CodeModule
    
    Set VBProj = ActiveWorkbook.VBProject
    
    For Each VBComp In VBProj.VBComponents
        If VBComp.Type = vbext_ct_Document Then
            Set CodeMod = VBComp.CodeModule
            With CodeMod
                .DeleteLines 1, .CountOfLines
            End With
        Else
            VBProj.VBComponents.Remove VBComp
        End If
    Next VBComp
End Sub

