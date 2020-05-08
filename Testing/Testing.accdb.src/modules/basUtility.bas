Option Compare Database
Option Explicit

' Keep a persistent reference to file system object after initializing version control.
' This way we don't have to recreate this object dozens of times while using VCS.
Private m_FSO As Scripting.FileSystemObject


'---------------------------------------------------------------------------------------
' Procedure : FSO
' Author    : Adam Waller
' Date      : 1/18/2019
' Purpose   : Wrapper for file system object. A property allows us to clear the object
'           : reference when we have completed an export or import operation.
'---------------------------------------------------------------------------------------
'
Public Property Get FSO() As Scripting.FileSystemObject
    If m_FSO Is Nothing Then Set m_FSO = New Scripting.FileSystemObject
    Set FSO = m_FSO
End Property
Public Property Set FSO(ByVal RHS As Scripting.FileSystemObject)
    Set m_FSO = RHS
End Property



'---------------------------------------------------------------------------------------
' Procedure : RunAfterBuild
' Author    : Adam Waller
' Date      : 5/1/2020
' Purpose   : Compile and run testing after build.
'---------------------------------------------------------------------------------------
'
Public Sub RunAfterBuild()

    ' Compile and save VBA code. Should prompt for any errors here.
    DoCmd.RunCommand acCmdCompileAndSaveAllModules
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestRegEx
' Author    : Adam Waller
' Date      : 5/1/2020
' Purpose   : This tests VBA code that uses early binding with non-built-in
'           : VBE References. (Add references before importing code)
'---------------------------------------------------------------------------------------
'
Public Sub TestRegEx()
    Dim regEx As VBScript_RegExp_55.RegExp
End Sub


Public Sub TestProperties()

    Dim dbs As DAO.Database
    Dim prp As AccessObjectProperty
    Dim proj As CurrentProject
    
    Set dbs = CurrentDb
    Set proj = CurrentProject
    
    For Each prp In proj.AllModules(0).Properties
    'For Each prp In dbs.TableDefs("tblLinkedCSV").Properties
        Debug.Print prp.Name & ": " & prp.Value
    Next prp
    
End Sub