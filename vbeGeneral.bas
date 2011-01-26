Attribute VB_Name = "vbeGeneral"
Option Explicit
'Option Private Module

Public Function VBComponentExists( _
                  ModuleName As String, _
                  Optional VBProject As Variant) _
                  As Boolean
  
  Dim tmp As Variant, vbproj As Object
  
  On Error GoTo Local_Error
  
  If IsMissing(VBProject) Then
    Set VBProject = ThisWorkbook.VBProject
  Else
    Set vbproj = VBProject
  End If

  Set tmp = vbproj.VBComponents(ModuleName)
  
  VBComponentExists = True
  On Error GoTo 0
  Exit Function
  
Local_Error:
  VBComponentExists = False
End Function

Public Function DeleteModule( _
                  VBProject As VBProject, _
                  ModuleName As String) _
                  As Boolean
' This does not work properly on any code module that is used actively at the time of running.
' Any bars that depend on the deleted module will stop working properly and need to be re-built
' e.g. CommandBar1 has button that is linked to some code that is Reloaded.  The bar will stop working after
' the code is run.

  Dim vbcomp As VBComponent
  
  On Error GoTo Local_Error
  
  With VBProject.VBComponents
    If .Item(ModuleName).Type = vbext_ct_Document Then
      ClearCodeModule .Item(ModuleName)
    Else
      Set vbcomp = .Item(ModuleName)
      .Remove vbcomp
    End If
  End With
  
  DeleteModule = True
  On Error GoTo 0
  Exit Function
  
Local_Error:
  DeleteModule = False
End Function

Public Sub ClearCodeModule(vbcomp As VBComponent)
  
  With vbcomp.CodeModule
    .DeleteLines 1, .CountOfLines
  End With ' VBComp.CodeModule
End Sub

Public Function CopyCodeModule( _
                  Source As VBComponent, _
                  Destination As VBComponent) _
                  As Boolean
  
  On Error GoTo Local_Error
  
  Dim tmp As String
  With Source.CodeModule
    tmp = .Lines(1, .CountOfLines)
  End With
  
   Destination.InsertLines 1, tmp
Local_Error:
  If Err.Number <> 0 Then Debug.Print Err.Description
End Function

