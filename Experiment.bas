Attribute VB_Name = "Experiment"
Public Sub vbeExperimental()
  ' Clear all the code from the active vb project
  
  Dim vbcomp As VBComponent
  Dim cm As New vbeVBComponent
  
  Debug.Print "beginning expt"
  'For Each vbcomp In Application.VBE.ActiveVBProject.VBComponents
    Set cm.VBComponent = Application.VBE.SelectedVBComponent
    
    If cm.Options.Count > 0 Then MsgBox cm.OptionList
  'Next ' vbcomp
End Sub
