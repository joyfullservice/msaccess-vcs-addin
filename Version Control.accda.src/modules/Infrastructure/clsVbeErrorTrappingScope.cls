VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsVbeErrorTrappingScope"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : clsVbeErrorTrappingScope
' Author    : Adam Waller
' Date      : 9/7/2026
' Purpose   : Restorable VBE Error Trapping override. Saves the caller's mode, raises
'           : it to a required floor, and restores the original on explicit Restore or
'           : Class_Terminate so early returns cannot leak the override.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
'@Folder("Infrastructure")


Private m_intSaved As eVbeErrorTrapping
Private m_blnActive As Boolean


'---------------------------------------------------------------------------------------
' Procedure : Init
' Author    : Adam Waller
' Date      : 9/7/2026
' Purpose   : Save the current mode and raise it to intRequired when the current value
'           : is below that floor. A caller already at a more permissive mode is left
'           : alone.
'---------------------------------------------------------------------------------------
'
Public Sub Init(intRequired As eVbeErrorTrapping)
    m_intSaved = SaveUserErrorTrapping()
    ApplyUserErrorTrapping EffectiveVbeErrorTrapping(m_intSaved, intRequired)
    m_blnActive = True
End Sub


'---------------------------------------------------------------------------------------
' Procedure : Restore
' Author    : Adam Waller
' Date      : 9/7/2026
' Purpose   : Put the saved mode back exactly once.
'---------------------------------------------------------------------------------------
'
Public Sub Restore()
    If Not m_blnActive Then Exit Sub
    RestoreUserErrorTrapping m_intSaved
    m_blnActive = False
End Sub


Private Sub Class_Terminate()
    On Error Resume Next
    If m_blnActive Then Restore
End Sub
