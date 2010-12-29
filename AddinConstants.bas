Attribute VB_Name = "AddinConstants"
' #NoRefresh

Option Explicit
Option Private Module

Public Enum OptionTypesEnum
  OptionType_Error = -1
  [_FirstOption] = 0
  OptionType_Boolean = 0
  OptionType_String = 1
  [_LastOption] = 1
End Enum ' OptionTypesEnum

' General option tokens
Public Const OPTIONS_TOKEN As String = "#"
Public Const OPTIONS_ASSIGNMENT_TOKEN As String = "="
Public Const OPTION_DELIMITER As String = "|"

Public Const COMMENT_TOKEN = "'"

' Lists with the option types
Public Const OPTION_STRINGS = "RelativePath|AbsolutePath"

' Information Options
Public Const OPTION_NO_LIST As String = "NoList"

' Import/Export Options
Public Const OPTION_NO_EXPORT As String = "NoExport"
Public Const OPTION_RELATIVE_PATH As String = "RelativePath"
Public Const OPTION_ABSOLUTE_PATH As String = "AbsolutePath" ' *NYI*
Public Const OPTION_NO_REFRESH As String = "NoRefresh"

Public Const cDELETED_MODULE_NAME_APPENDIX As String = "__d3l"
