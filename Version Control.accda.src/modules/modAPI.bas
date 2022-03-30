Attribute VB_Name = "modAPI"
'---------------------------------------------------------------------------------------
' Module    : modAPI
' Author    : Adam Waller
' Date      : 1/13/2021
' Purpose   : This module exposes a set of VCS tools to other projects.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit

' Note, some enums are listed here when they are directly exposed
' through the Options class. (Allowing them to be used externally)

' Formats used when exporting table data.
Public Enum eTableDataExportFormat
    etdNoData = 0
    etdTabDelimited = 1
    etdXML = 2
    [_Last] = 2
End Enum

Public Enum eSanitizeLevel
    eslNone = 0     ' Sanitize only items which won't build correctly unless you sanitize them.
    eslBasic        ' Strip out excess items (like GUIDs) that are just noise and no effect can be found.
    eslAggressive    ' Strip out anything that can be reliably rebuilt by Access during Build (themed control colors).

    ' WARNING: AdvancedBeta introduces sanitzation that may or may not work in all environments, and has known
    '          (or highly suspected) edge cases where it does not always operate correctly. Do not use this level in
    '          production databases.
    eslAdvancedBeta ' Remove all excess noise. Try out new sanitize features that still have ragged edges.
    [_Last]         ' DO NOT REMOVE: This is a "Fake" level, and must be at the end.
End Enum


'---------------------------------------------------------------------------------------
' Procedure : HandleRibbonCommand
' Author    : Adam Waller
' Date      : 2/28/2022
' Purpose   : Handle an incoming command from the TwinBasic ribbon COM add-in.
'           : (Allows us to keep all the logic between the XML ribbon file and the
'           :  Access add-in.)
'---------------------------------------------------------------------------------------
'
Public Function HandleRibbonCommand(strCommand As String) As Boolean

    ' If a function is not found, this will throw an error. It is up to the ribbon
    ' designer to ensure that the control IDs match public procedures in the VCS
    ' (clsVersionControl) class module. Additional parameters are not supported.
    ' For example, to run VCS.Export, the ribbon button ID should be named "btnExport"
    
    ' Trim off control ID prefix when calling command
    CallByName VCS, Mid(strCommand, 4), VbMethod

End Function


'---------------------------------------------------------------------------------------
' Procedure : VCS
' Author    : Adam Waller
' Date      : 3/28/2022
' Purpose   : Wrapper for the VCS class, providing easy API access to VCS functions.
'           : *NOTE* that this class is not persisted. This allows us to wrap up and
'           : remove any object references after the call completes.
'---------------------------------------------------------------------------------------
'
Public Function VCS() As clsVersionControl
    Set VCS = New clsVersionControl
End Function


'---------------------------------------------------------------------------------------
' Procedure : ExampleLoadAddInAndRunExport
' Author    : Adam Waller
' Date      : 11/13/2020
' Purpose   : This function can be copied to a local database and triggered with a
'           : command line argument or other automation technique to load the VCS
'           : add-in file and initiate an export.
'           : NOTE: This expects the add-in to be installed in the default location
'           : and using the default file name.
'---------------------------------------------------------------------------------------
'
Public Function ExampleLoadAddInAndRunExport()

    Dim strAddInPath As String
    Dim proj As Object      ' VBProject
    Dim objAddIn As Object  ' VBProject
    
    ' Build default add-in path
    strAddInPath = Environ$("AppData") & "\MSAccessVCS\Version Control.accda"

    ' See if add-in project is already loaded.
    For Each proj In VBE.VBProjects
        If StrComp(proj.FileName, strAddInPath, vbTextCompare) = 0 Then
            Set objAddIn = proj
        End If
    Next proj
    
    ' If not loaded, then attempt to load the add-in.
    If objAddIn Is Nothing Then
        
        ' The following lines will load the add-in at the application level,
        ' but will not actually call the function. Ignore the error of function not found.
        ' https://stackoverflow.com/questions/62270088/how-can-i-launch-an-access-add-in-not-com-add-in-from-vba-code
        On Error Resume Next
        Application.Run strAddInPath & "!DummyFunction"
        On Error GoTo 0
    
        ' See if it is loaded now...
        For Each proj In VBE.VBProjects
            If StrComp(proj.FileName, strAddInPath, vbTextCompare) = 0 Then
                Set objAddIn = proj
            End If
        Next proj
    End If

    If objAddIn Is Nothing Then
        MsgBox "Unable to load Version Control add-in. Please ensure that it has been installed" & vbCrLf & _
            "and is functioning correctly. (It should be available in the Add-ins menu.)", vbExclamation
    Else
        ' Launch add-in export for current database.
        Application.Run "MSAccessVCS.ExportSource", True
    End If
    
End Function


