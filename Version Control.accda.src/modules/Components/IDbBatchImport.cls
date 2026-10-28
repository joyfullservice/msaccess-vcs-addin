VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IDbBatchImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : IDbBatchImport (Interface)
' Author    : Adam Waller
' Date      : 9/15/2026
' Purpose   : Optional capability for component classes whose full-build imports can
'           : create every object first, refresh the DAO Documents collection once,
'           : then apply metadata and update the index in a second pass.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
'@Folder("Components")


Public Sub ImportFast(strFile As String)
End Sub


Public Sub FinalizeImports()
End Sub
