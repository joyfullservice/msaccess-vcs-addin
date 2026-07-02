VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IOptionsSection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : IOptionsSection
' Author    : Adam Waller
' Date      : 3/17/2026
' Purpose   : Interface for options form section subforms. Each section form implements
'           : this interface to provide a consistent load/save contract with the main
'           : options form.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
'@Folder("Interfaces")

'---------------------------------------------------------------------------------------
' Procedure : LoadOptions
' Author    : Adam Waller
' Date      : 3/17/2026
' Purpose   : Load option values from the options class instance into form controls.
'---------------------------------------------------------------------------------------
'
Public Sub LoadOptions()
End Sub


'---------------------------------------------------------------------------------------
' Procedure : SaveOptions
' Author    : Adam Waller
' Date      : 3/17/2026
' Purpose   : Save form control values back to the options class instance.
'---------------------------------------------------------------------------------------
'
Public Sub SaveOptions()
End Sub
