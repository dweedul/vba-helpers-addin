Attribute VB_Name = "config"
'#RelativePath = vba
'! relative-path vba

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
  
  ' Add the no-export option.
  ' This flag, when present, prevents the module from exporting.
  optParse.addOption "no-export", typename:="bool", default:=False
  
  ' Add the no-reload option.
  ' This flag, when present, prevents the module from reloading
  optParse.addOption "no-reload", typename:="bool", default:=False
  
  ' Add the absolute-path string option.
  ' This sets the path for exporting/reloading.
  ' This will override the normal filename pattern of _<moduleName>.<typedExtension>_
  ' with the contents of the option.
  optParse.addOption "absolute-path", typename:="string", default:=""
  
  ' Add the relative-path string option.
  ' This sets the path for exporting/reloading relative to the base workbook.
  optParse.addOption "relative-path", typename:="string", default:=""
  
  Set vbeVBComponentOptionParser = optParse
End Function
