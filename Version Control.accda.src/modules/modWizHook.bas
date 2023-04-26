Attribute VB_Name = "modWizHook"
'---------------------------------------------------------------------------------------
' Module    : modWizHook
' Author    : Adam Waller
' Date      : 5/4/2022
' Purpose   : Expose some WizHook functions utilized by this add-in.
'           : Documentation links: http://accessblog.net/2016/06/access-wizhook-library.html
'---------------------------------------------------------------------------------------
Option Compare Database
Option Private Module
Option Explicit


'---------------------------------------------------------------------------------------
' Procedure : CloseCurrentDatabase2
' Author    : Adam Waller
' Date      : 5/4/2022
' Purpose   : Unlike the Application method, this WizHook version does not stop all
'           : running code. This allows you to automate the closing of the current
'           : database while still continuing the add-in code.
'---------------------------------------------------------------------------------------
'
Public Sub CloseCurrentDatabase2()
    CheckKey
    WizHook.CloseCurrentDatabase
    DoEvents
End Sub


'---------------------------------------------------------------------------------------
' Procedure : CurrentVBProject
' Author    : Adam Waller
' Date      : 3/10/2023
' Purpose   : Get a reference to the VB Project for the current database
'---------------------------------------------------------------------------------------
'
Public Function CurrentVBProject() As VBProject
    CheckKey
    Set CurrentVBProject = WizHook.DbcVbProject
End Function


'---------------------------------------------------------------------------------------
' Procedure : GlobalProcExists
' Author    : Adam Waller
' Date      : 3/10/2023
' Purpose   : Returns true if the procedure exists (Public proc in standard module)
'---------------------------------------------------------------------------------------
'
Public Function GlobalProcExists(strName) As Boolean
    CheckKey
    GlobalProcExists = WizHook.GlobalProcExists(strName)
End Function


'---------------------------------------------------------------------------------------
' Procedure : CheckKey
' Author    : Adam Waller
' Date      : 5/4/2022
' Purpose   : Make sure we have set the WizHook key before using commands that require
'           : it to be set. (Caches value since it only needs to be set once per
'           : session.)
'---------------------------------------------------------------------------------------
'
Private Function CheckKey()
    Static lngKey As Long
    If lngKey = 0 Then
        lngKey = 51488399
        WizHook.Key = lngKey
    End If
End Function
