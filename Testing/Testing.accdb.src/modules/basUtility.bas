Attribute VB_Name = "basUtility"
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

    ' Run startup macro to execute tests.
    DoCmd.RunMacro "AutoExec"

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


'---------------------------------------------------------------------------------------
' Procedure : RunTests
' Author    : Adam Waller
' Date      : 5/13/2020
' Purpose   : Run the tests (Fired from autoexec macro)
'---------------------------------------------------------------------------------------
'
Public Function RunTests()
    Form_frmMain.cmdRunTests_Click
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetVBProjectForCurrentDB
' Author    : Adam Waller
' Date      : 7/25/2017
' Purpose   : Get the actual VBE project for the current top-level database.
'           : (This is harder than you would think!)
'---------------------------------------------------------------------------------------
'
Public Function GetVBProjectForCurrentDB() As VBProject

    Dim objProj As Object
    Dim strPath As String

    strPath = CurrentProject.FullName
    If VBE.ActiveVBProject.FileName = strPath Then
        ' Use currently active project
        Set GetVBProjectForCurrentDB = VBE.ActiveVBProject
    Else
        ' Search for project with matching filename.
        For Each objProj In VBE.VBProjects
            If objProj.FileName = strPath Then
                Set GetVBProjectForCurrentDB = objProj
                Exit For
            End If
        Next objProj
    End If

End Function


'---------------------------------------------------------------------------------------
' Procedure : SetClassInstancing
' Author    : Adam Waller
' Date      : 11/25/2023
' Purpose   : Set the clsPublic class to use Public Creatable instancing.
'---------------------------------------------------------------------------------------
'
Public Sub SetClassInstancing()
    Const PublicCreatable = 5
    GetVBProjectForCurrentDB.VBComponents("clsPublic").Properties("Instancing") = PublicCreatable
End Sub


'---------------------------------------------------------------------------------------
' Procedure : Menu
' Author    : Adam Waller
' Date      : 1/16/2024
' Purpose   : Wrapper for persistent class instance
'---------------------------------------------------------------------------------------
'
Public Function Menu() As clsPopupMenu
    Static cMenu As clsPopupMenu
    If cMenu Is Nothing Then Set cMenu = New clsPopupMenu
    Set Menu = cMenu
End Function


'---------------------------------------------------------------------------------------
' Procedure : MenuHandler
' Author    : Adam Waller
' Date      : 1/16/2024
' Purpose   : Global handler for managing command bar events. The single parameter
'           : is the name of the command bar, followed by the caption of the item.
'           : I.e. "MyCommandBar.Print Preview"
'---------------------------------------------------------------------------------------
'
Public Function MenuHandler(strSource As String)

    Select Case strSource
        Case "CRM Navigation General.Contact": MsgBox "Contact"
        Case Else
            Debug.Print "Source not defined in MenuHandler: Case """ & strSource & """"
    End Select

End Function


'---------------------------------------------------------------------------------------
' Procedure : NZ2
' Author    : Adam Waller
' Date      : 2/18/2021
' Purpose   : Extend the NZ function to also include 0 or empty string.
'---------------------------------------------------------------------------------------
'
Public Function NZ2(varValue, Optional varIfNull) As Variant
    Select Case varValue
        Case vbNullString, 0
            If IsMissing(varIfNull) Then
                NZ2 = vbNullString
            Else
                NZ2 = varIfNull
            End If
        Case Else
            If IsNull(varValue) Then
                NZ2 = varIfNull
            Else
                NZ2 = varValue
            End If
    End Select
End Function


'---------------------------------------------------------------------------------------
' Procedure : MultiReplace
' Author    : Adam Waller
' Date      : 1/18/2019
' Purpose   : Does a string replacement of multiple items in one call.
'---------------------------------------------------------------------------------------
'
Public Function MultiReplace(ByVal strText As String, ParamArray varPairs()) As String
    Dim intPair As Integer
    For intPair = 0 To UBound(varPairs) Step 2
        strText = Replace(strText, varPairs(intPair), varPairs(intPair + 1))
    Next intPair
    MultiReplace = strText
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetDatabaseFileName
' Author    : Adam Waller
' Date      : 2/8/2025
' Purpose   : Function for use in queries
'---------------------------------------------------------------------------------------
'
Public Function GetDatabaseFileName() As String
    GetDatabaseFileName = CurrentDb.Name
End Function
