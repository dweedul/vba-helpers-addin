Attribute VB_Name = "config"
'! relative-path vba

' Configuration for the addin
' Note: Changing anything here requires that the addon be reloaded.

'! requires vbeOptionParser
'! requires vbeOption

Option Explicit

' storage for all the warnings
Private warnings As Dictionary



' ## Warning flags

' Throw a warning when a single module will be reloaded.
Public Const WARN_ON_RELOAD_SINGLE As Boolean = False

' Throw an error when a whole project will be reloaded.
Public Const WARN_ON_RELOAD_ALL As Boolean = True



' ## Component Options

' Set up the options for the the vbeVBComponent Class.
'
' Note: This will be called by the vbeVBComponent class upon initialization.
'
' Returns an option parser for use in the calling class.
Public Function vbeVBComponentOptionParser() As vbeOptionParser
  Dim optParse As New vbeOptionParser
  
  ' set the default option token within the vbComponent
  ' This should be some form of comment to prevent debugging errors.
  optParse.optionToken = "'!"
  
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
  optParse.addOption "absolute-path <path>", typename:="string", default:=""
  
  ' Add the relative-path string option.
  ' This sets the path for exporting/reloading relative to the base workbook.
  optParse.addOption "relative-path <path>", typename:="string", default:=""
  
  Set vbeVBComponentOptionParser = optParse
End Function



' ## Helper functions for the warnings.

' configure all warnings
'
' hideMe - exclude this from the macro window
Public Sub configWarnings(Optional hideMe As Byte)
  Dim o As vbeOption
  
  ' reset the warnings collection
  Set warnings = New Dictionary
  
  ' store the reload-all warning info
  Set o = New vbeOption
  o.value = WARN_ON_RELOAD_ALL:  o.args.Add WARNING_OVERWRITE
  warnings.Add "reload-all", o
  
  ' store the reload-one warning info
  Set o = New vbeOption
  o.value = WARN_ON_RELOAD_SINGLE:  o.args.Add WARNING_OVERWRITE
  warnings.Add "reload-one", o

End Sub

' Warn the user about something.
'
' prompt - a string warning
'
' Returns user's decision. Proceed = true. Defaults to false
Public Function warnUser(warningName As String) As Boolean
  Dim resp As VbMsgBoxResult
  
  ' default to false
  warnUser = False
  
  ' load the warnings
  configWarnings
  
  With warnings(warningName)
    If .value Then
      resp = MsgBox(.args(1), vbOKCancel, "WARNING!!!")
    Else
      resp = vbOK
    End If
  End With
  
  If resp = vbOK Then
    warnUser = True
  Else
    warnUser = False
  End If

End Function
