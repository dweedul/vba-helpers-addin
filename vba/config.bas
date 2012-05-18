Attribute VB_Name = "config"
'#RelativePath = vba

' Configuration for the addin

'! requires vbeOptionParser

Option Explicit

' Option token - marks the beginning of the option to parse from the modules.
' This should be some form of comment to prevent debugging errors.
Public Const OPTION_TOKEN As String = "'!"

' Set up the options for the the vbeVBComponent Class.
'
' Note: This will be called by the vbeVBComponent class upon initialization.
'
' Returns an option parser for use in the calling class.
Public Function vbeVBComponentOptionParser() As vbeOptionParser
  Dim optParse As New vbeOptionParser
  
  ' add the no-export option
  ' this flag, when present, prevents the module from exporting
  optParse.addOption "no-export", typename:="bool", default:=False
  
  ' add the no-reload option
  ' this flag, when present, prevents the module from reloading
  optParse.addOption "no-reload", typename:="bool", default:=False
  
  ' add the relative-path string option
  ' this tells the addin to save the component to / load the component
  ' from this path, relative to the parent workbook
  optParse.addOption "relative-path", typename:="string", default:=""
  
  Set vbeVBComponentOptionParser = optParse
End Function
