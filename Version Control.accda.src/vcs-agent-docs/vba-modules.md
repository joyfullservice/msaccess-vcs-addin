# VBA Module Source Files

Covers `.bas` and `.cls` files in `modules/`, and the `.cls` code-behind files in
`forms/` and `reports/`.

## Standard modules (`.bas`)

```
Attribute VB_Name = "ModuleName"        <- Required: must match the filename
'---------------------------------------------------------------------------------------
' Module    : ModuleName                <- Optional comment header
' Purpose   : Description here
'---------------------------------------------------------------------------------------
Option Compare Database                  <- Optional
Option Explicit                          <- Optional

Public Sub MySub()
    ' Code starts after any Option statements
End Sub
```

Add and edit procedures, functions, and declarations freely. Leave the
`Attribute VB_Name` line alone; it must equal the filename without its extension.

`Option` statements are not required. Some projects omit them entirely, so match
whatever the surrounding modules do rather than adding them.

## Class modules (`.cls`)

```
VERSION 1.0 CLASS                        <- Required header block
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassName"          <- Required: must match the filename
Attribute VB_GlobalNameSpace = False     <- Required (legacy, but must be present)
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False       <- True gives the class a default instance
Attribute VB_Exposed = False             <- True exposes it to other projects
Option Compare Database                  <- Optional
Option Explicit                          <- Optional

Private m_Value As String

Public Property Get Value() As String
Attribute Value.VB_Description = "Returns the value"   <- Member attribute
    Value = m_Value
End Property
```

Edit the code and comment headers. Leave the `VERSION` block and the module-level
`Attribute VB_*` lines exactly as they are.

## Member attributes

Attributes can also appear **inside** procedures and properties, carrying metadata
that the VBA editor hides but the Object Browser displays. They go on the line
immediately after the `Sub`, `Function`, or `Property` declaration:

```vba
Public Function GetItem(Index As Long) As Variant
Attribute GetItem.VB_Description = "Returns item at specified index"
Attribute GetItem.VB_UserMemId = 0
    GetItem = m_Items(Index)
End Function
```

| Attribute | Effect |
|-----------|--------|
| `Attribute [Member].VB_Description = "text"` | Description shown in the Object Browser |
| `Attribute [Member].VB_UserMemId = 0` | Makes this the class's default member |
| `Attribute [Member].VB_UserMemId = -4` | Returns an enumerator, enabling `For Each` |

Reword the description text if it is wrong. Do not change the attribute syntax or
the member name it refers to; the name must match the procedure it sits under.

## Form and report code-behind

When the **Split Layout from VBA** option is on, a form's or report's code lives in
its own `.cls` beside the layout file:

```
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmMyForm"     <- Must be "Form_" (or "Report_") + object name
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_Load()
    ' Event handler code
End Sub
```

Edit event handlers and procedures here as you would in any class module. When the
option is off, the code is instead inside the `.form` or `.report` file under a
`CodeBehindForm` line.
