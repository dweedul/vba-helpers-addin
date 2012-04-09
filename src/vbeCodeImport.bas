Attribute VB_Name = "vbeCodeImport"
' #RelativePath = src

Option Explicit
Option Private Module

' ----------------
' Public functions
' ----------------
Public Sub vbeReloadActiveVBProject(Optional HideMe As Boolean)
  Dim vbproj As VBProject, VBComp As VBComponent
  Dim cm As vbeVBComponent, col As Collection, fn As Variant
  Dim msg As VbMsgBoxResult
   
  On Error GoTo Local_Error
  
  ' get the active vb project
  Set vbproj = Application.VBE.ActiveVBProject
  
  ' confirm this action
  msg = MsgBox("Are you sure that you want to reload the code in " & vbproj.Name, vbYesNo)
  If msg = vbNo Then GoTo Local_Error
  
  ' build a collection of filenames
  Set col = New Collection
  For Each VBComp In vbproj.VBComponents
    Set cm = New vbeVBComponent
    Set cm.VBComponent = VBComp
    If Not cm.IsEmpty And Not cm.Options(OPTION_NO_RELOAD) Then
      col.Add vbeParsePath(vbproj.Filename) & vbeFileNameFromModule(VBComp)
    End If
   Next ' vbcomp
  
  ' delete the current code project
  Set vbproj = vbeDeleteVBProject(vbproj)
  
  ' import the files back into the project
  For Each fn In col
    ImportVBComponent vbproj, CStr(fn)
  Next ' fn
  
Local_Error:
End Sub

Public Sub vbeReloadCodeModule( _
             Optional VBComp As VBComponent)
' @optparam vbcomp [VBComponent]

  Dim o As VBComponent, fname As String, cm As vbeVBComponent
  
  If VBComp Is Nothing Then
    Set o = Application.VBE.SelectedVBComponent
  Else
    Set o = VBComp
  End If
  
  Set cm = New vbeVBComponent
  Set cm.VBComponent = o
  
  If Not cm.Options(OPTION_NO_RELOAD) Then
    fname = vbeStandardizePath(vbeParsePath(o.Collection.Parent.Filename)) & vbeFileNameFromModule(o)
    
    If Dir(fname, vbNormal + vbHidden + vbSystem) <> vbNullString Then
      Set ThisWorkbook.VBProjCache = o.Collection.Parent
      
      Application.OnTime Now(), ThisWorkbook.Name & "!'ImportVBComponent ThisWorkbook.VBProjCache, """ & fname & """'"
      
      'ImportVBComponent o.Collection.Parent, _
      '                  fname
    End If
  End If

End Sub ' vbeReloadCodeModule

' ----------------
' Import Functions
' ----------------
Public Function ImportVBComponent( _
                  VBProject As Object, _
                  Filename As String, _
                  Optional ModuleName As String, _
                  Optional OverwriteExisting As Boolean = True, _
                  Optional SelectOnDone As Boolean = True) _
                  As Boolean
' This function imports the code module of a VBComponent from a text _
' file. If ModuleName is missing, the code will be imported to _
' a module with the same name as the filename without the extension _
' @param VBProject [VBProject]
' @param FileName [string]
' @optparam ModuleName [string]
' @optparam OverwriteExisting [bool]
' @return [bool] False on error
  
  On Error GoTo Local_Error
  
  ' handle a missing module name
  If ModuleName = vbNullString Then
    ModuleName = vbeParseBaseFilename(Filename)
  End If
  
  If vbeVBComponentExists(ModuleName, VBProject) Then
    If OverwriteExisting Then
      vbeDeleteModule VBProject, ModuleName
    Else
      GoTo Local_Error
    End If
  End If
  
  ImportFromFile VBProject, Filename, ModuleName

  If SelectOnDone Then VBProject.VBComponents(ModuleName).Activate
  
  ImportVBComponent = True
  On Error GoTo 0
  Exit Function
  
Local_Error:
  ImportVBComponent = False
End Function ' ImportVBComponent

Private Function ImportFromFile( _
                  VBProject As VBProject, _
                  Filename As String, _
                  ModuleName As String) _
                  As Boolean
' @param VBProject [VBProject]
' @param Filename [string]
' @param ModuleName [string]

  Dim tmp_vbcomp As VBComponent
  Dim s As String
  
  On Error GoTo Local_Error
  
  With VBProject.VBComponents
    If vbeVBComponentExists(ModuleName, VBProject) Then
      If .Item(ModuleName).Type = 100 Then ' 100 =vbext_ct_Document
        Set tmp_vbcomp = .Import(Filename)
        s = tmp_vbcomp.CodeModule.Lines(1, tmp_vbcomp.CodeModule.CountOfLines)
        .Item(ModuleName).CodeModule.InsertLines 1, s
        .Remove tmp_vbcomp
      End If
    Else
      .Import Filename:=Filename
    End If
  End With
     
  ImportFromFile = True
  On Error GoTo 0
  Exit Function
  
Local_Error:
  ImportFromFile = False
End Function ' ImportFromFile
