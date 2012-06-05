Attribute VB_Name = "importHelpers"
'! relative-path vba

' This module contains all the code needed to make an import happen
' for a given VBComponent.

'! requires vbeVBComponent
'! references "Microsoft Visual Basic for Applications Extensibility 5.3"

Option Explicit

' Import a component
'
' project        - name of the project that contains the component
' component      - name of the component
' path           - file path to the component's file
' shouldActivate - should the component be selected afterward
'                  defaults to TRUE
'
' Note: This should NEVER be called normally.
'       It should ALWAYS be scheduled with `Application.OnTime`
'       The VBE doesn't handle imports correctly, otherwise.
Public Sub importFromFile( _
             project As String, _
             component As String, _
             path As String, _
             Optional shouldActivate As Boolean = True)
  
  Dim p As VBProject:  Set p = Application.VBE.VBProjects(project)
  Dim c As New vbeVBComponent, tmp As New vbeVBComponent
  Dim fso As New FileSystemObject
  
  ' check if the file exists at the indicated path
  If Not fso.FileExists(path) Then
    MsgBox "The file was not found at " & vbCrLf & path
    Exit Sub
  End If
  
  ' remove the component from the project if it exists
  If VBComponentExists(component, p) Then
    Set c.baseObject = p.VBComponents(component)
    
    If c.baseObject.Type = vbext_ct_Document Then
      c.clear
      
      ' import the component into a new object,
      ' copy the tmp component's code into the old object
      ' and then remove the tmp component
      Set tmp.baseObject = p.VBComponents.import(path)
      c.baseObject.CodeModule.InsertLines 1, tmp.code
      tmp.remove: Set tmp = Nothing
      
    Else
      c.remove
      Set c.baseObject = p.VBComponents.import(path)
    End If
    
  Else
    ' import the component
    Set c.baseObject = p.VBComponents.import(path)
  End If
  
  ' activate as required
  If shouldActivate Then c.activate
End Sub

' Import a folder of code into a VBProject.
'
' Returns the target project.
Public Function importFromFolder( _
                  vbProj As VBProject) _
                  As VBProject
  Dim path As String, fso As New FileSystemObject, f As File
  
  If Len(vbProj.filename) = 0 Then
    path = pickFolder("%DESKTOP%")
  Else
    path = pickFolder(fso.GetParentFolderName(vbProj.filename))
  End If
  
  ' Validate the file, then import
  For Each f In fso.getFolder(path).Files
    If isValidExtension(fso.GetExtensionName(f.path)) Then
      Debug.Print fso.GetBaseName(f.path)
    Application.OnTime Now(), ThisWorkbook.name & "!'importFromFile """ & _
                              vbProj.name & """, """ & _
                              fso.GetBaseName(f.path) & """, """ & _
                              f.path & """'"
    End If
  Next ' f
  
End Function

' Check for the existence of a vbcomponent
'
' component - name of the component
' project   - either a project name or object
'             defaults to the current VBProject
'
' Returns true/false existence.
Private Function VBComponentExists( _
                  ModuleName As String, _
                  Optional project As Variant) _
                  As Boolean
  
  Dim tmp As Variant, vbProj As Object
  
  On Error GoTo errorHandler
  
  ' If the project reference is missing,
  ' set it to the current project
  If IsMissing(project) Then
    Set vbProj = ThisWorkbook.VBProject
    
  ' If an object was passed, use that object.
  ElseIf typename(project) = "VBProject" Then
    Set vbProj = project
  
  ' Otherwise, assume a string and use that.
  Else
    Set vbProj = Application.VBE.VBProjects(project)
    
  End If
  
  ' Try to set the temp object to the component
  ' if there is an error, jump to errorHandler
  ' and return false. Otherwise, return true.
  Set tmp = vbProj.VBComponents(ModuleName)
  
  VBComponentExists = True
  On Error GoTo 0
  Exit Function
  
errorHandler:
  VBComponentExists = False
End Function

' Get a folder from a file dialog
'
' path - the path to start from
'
' Returns the selected folder path.
Private Function pickFolder( _
                   path As String) _
                   As String
  Dim fol As String
    
  With Application.FileDialog(msoFileDialogFolderPicker)
    .Title = "Select a folder"
    .AllowMultiSelect = False
    .InitialFileName = path
    If .Show <> -1 Then GoTo errorHandler
    fol = .SelectedItems(1)
  End With ' folderPicker

errorHandler:
  pickFolder = fol
End Function


' Check the extension is valid
Private Function isValidExtension( _
                   ext As String) _
                   As Boolean
  If ext = "bas" Or ext = "cls" Or ext = "frm" Then isValidExtension = True
End Function

