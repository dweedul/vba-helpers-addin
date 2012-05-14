VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "vbeVBComponent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'#RelativePath = src

' This is the base class for the code module manipulations

'# references "Microsoft Visual Basic for Applications Extensibility 5.3"

Option Explicit

Private m_baseObject As VBComponent

' Get the base component stored in this object.
'
' Default member of the class
Public Property Get baseObject() As VBComponent
Attribute baseObject.VB_UserMemId = 0
  Set baseObject = m_baseObject
End Property

' Set the base object
Public Property Set baseObject(baseObject As VBComponent)
  Set m_baseObject = baseObject
End Property
