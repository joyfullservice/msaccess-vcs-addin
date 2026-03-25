VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsLayoutNode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : clsLayoutNode
' Author    : Adam Waller
' Date      : 03/19/2026
' Purpose   : Lightweight node class for representing parsed form/report layout elements.
'           : Each node stores its control type, a Dictionary of properties, and an
'           : ordered Collection of child nodes for tree traversal.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
'@Folder("Core")

Private Type udtThis
    strControlType As String
    dProps As Dictionary
    colChildren As Collection
End Type
Private this As udtThis


'---------------------------------------------------------------------------------------
' Procedure : Class_Initialize
' Author    : Adam Waller
' Date      : 03/19/2026
' Purpose   : Set up property dictionary and children collection.
'---------------------------------------------------------------------------------------
'
Private Sub Class_Initialize()
    Set this.dProps = New Dictionary
    this.dProps.CompareMode = TextCompare
    Set this.colChildren = New Collection
End Sub


'---------------------------------------------------------------------------------------
' Procedure : ControlType
' Author    : Adam Waller
' Date      : 03/19/2026
' Purpose   : The type of control (e.g. "Form", "Section", "Label", "CommandButton").
'---------------------------------------------------------------------------------------
'
Public Property Get ControlType() As String
    ControlType = this.strControlType
End Property

Public Property Let ControlType(strValue As String)
    this.strControlType = strValue
End Property


'---------------------------------------------------------------------------------------
' Procedure : Props
' Author    : Adam Waller
' Date      : 03/19/2026
' Purpose   : Direct access to the properties dictionary.
'---------------------------------------------------------------------------------------
'
Public Property Get Props() As Dictionary
    Set Props = this.dProps
End Property


'---------------------------------------------------------------------------------------
' Procedure : Children
' Author    : Adam Waller
' Date      : 03/19/2026
' Purpose   : Ordered collection of child clsLayoutNode objects.
'---------------------------------------------------------------------------------------
'
Public Property Get Children() As Collection
    Set Children = this.colChildren
End Property


'---------------------------------------------------------------------------------------
' Procedure : AddChild
' Author    : Adam Waller
' Date      : 03/19/2026
' Purpose   : Add a child node and return it for chaining.
'---------------------------------------------------------------------------------------
'
Public Function AddChild(cNode As clsLayoutNode) As clsLayoutNode
    this.colChildren.Add cNode
    Set AddChild = cNode
End Function


'---------------------------------------------------------------------------------------
' Procedure : HasProp
' Author    : Adam Waller
' Date      : 03/19/2026
' Purpose   : Returns True if the named property exists.
'---------------------------------------------------------------------------------------
'
Public Function HasProp(strKey As String) As Boolean
    HasProp = this.dProps.Exists(strKey)
End Function


'---------------------------------------------------------------------------------------
' Procedure : PropVal
' Author    : Adam Waller
' Date      : 03/19/2026
' Purpose   : Returns the property value, or a default if not present.
'---------------------------------------------------------------------------------------
'
Public Function PropVal(strKey As String, Optional varDefault As Variant = vbNullString) As Variant
    If this.dProps.Exists(strKey) Then
        PropVal = this.dProps(strKey)
    Else
        PropVal = varDefault
    End If
End Function


'---------------------------------------------------------------------------------------
' Procedure : TwipLeft / TwipTop / TwipWidth / TwipHeight
' Author    : Adam Waller
' Date      : 03/19/2026
' Purpose   : Convenience accessors for geometry properties (in twips).
'---------------------------------------------------------------------------------------
'
Public Property Get TwipLeft() As Long
    TwipLeft = CLng(PropVal("Left", 0))
End Property

Public Property Get TwipTop() As Long
    TwipTop = CLng(PropVal("Top", 0))
End Property

Public Property Get TwipWidth() As Long
    TwipWidth = CLng(PropVal("Width", 0))
End Property

Public Property Get TwipHeight() As Long
    TwipHeight = CLng(PropVal("Height", 0))
End Property
