Attribute VB_Name = "importHelpers"
Public Const IMPORT_DELIMITER As String = "|"

' Import the queued components
Public Sub importQueuedComponents()
             
  Dim q As Collection: Set q = ThisWorkbook.importQueue
  Dim c As vbeVBComponent, arr As Variant
  
  ' check that there's something in the queue
  If q.Count < 1 Then GoTo errorHandler
  
  ' split the queued string
  arr = Split(q(q.Count), IMPORT_DELIMITER)
  
  ' import the first item
  Set proj = Application.VBE.VBProjects(arr(0))
  Set comp = proj.VBComponents(arr(1))
  path = arr(2)
  
  proj.VBComponents.Remove comp
  proj.VBComponents.import path
  
  ' remove that item from the queue
  q.Remove q.Count
  
  ' if there are more in the queue, schedule the queue again
  If q.Count > 0 Then
    Application.OnTime Now(), ThisWorkbook.name & "!importQueuedComponents"
  End If
  
errorHandler:
End Sub


