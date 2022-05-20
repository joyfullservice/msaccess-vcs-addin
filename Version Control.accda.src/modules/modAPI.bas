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
' through the Options and VCS classes. (Allowing them to be used externally)

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

    ' Make sure we are not attempting to run this from the current database when making
    ' changes to the add-in itself. (It will re-run the command through the add-in.)
    If RunningOnLocal(strCommand) Then Exit Function
    
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
' Procedure : RunningOnLocal
' Author    : Adam Waller
' Date      : 5/18/2022
' Purpose   : Returns true if the code is running in the current database instead of
'           : the add-in database.
'---------------------------------------------------------------------------------------
'
Private Function RunningOnLocal(strCommand As String) As Boolean
    
    Dim projAddIn As VBProject
    
    ' Make sure the add-in is loaded.
    If Not AddinLoaded Then LoadVCSAddIn

    ' When running code from the add-in project itself, it gets a little
    ' tricky because both the add-in and the currentdb have the same VBProject name.
    ' This means we can't just call `Run "MSAccessVCS.*" because it will run in
    ' the local project instead of the add-in. To pull this off, we will temporarily
    ' change the project name of the add-in so we can call it as distinct from the
    ' current project.
    Set projAddIn = GetAddInProject
    If StrComp(CurrentProject.FullName, CodeProject.FullName, vbTextCompare) = 0 Then
        ' When this is run from the CurrentDB, we should rename the add-in project,
        ' then call it again using the renamed project to ensure we are running it
        ' from the add-in.
        projAddIn.Name = "MSAccessVCS-Lib"
        RunningOnLocal = True
        Run "MSAccessVCS-Lib.HandleRibbonCommand", strCommand
    Else
        ' Reset project name if needed
        With projAddIn
            ' Technically, changes in the add-in will not be saved anyway, so this
            ' may not be needed, but just in case we refer to this project by name
            ' anywhere else in the code, we will restore the original name before
            ' passing in the
            If .Name = "MSAccessVCS-Lib" Then
                .Name = PROJECT_NAME
                ' User a timer to initiate the command again in the add-in database,
                ' giving the calling code a moment to close and release references.
                RunningOnLocal = True
                SetTimer roRibbonCommand, strCommand
            End If
        End With
    End If
    
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


