'---------------------------------------------------------------------------------------
' Module    : modAddIn
' Author    : Adam Waller
' Date      : 12/4/2020
' Purpose   : Manages the Microsoft Access Add-in, including menu items and VCS version.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit


Public Enum eReleaseType
    Major_Vxx = 0
    Minor_xVx = 1
    Build_xxV = 2
    Same_Version = 3
End Enum

Public Enum eRibbonLaunch
    erlVCSOpen
    erlVCSOptions
    erlExportAllRibbon
    erlExportFormsReportsRibbon
    erlExportFormsRibbon
    erlExportReportsRibbon
    erlExportModulesMacros
    erlExportTablesQueries
    erlExportTables
    erlExportQueries
End Enum


'---------------------------------------------------------------------------------------
' Procedure : AddInLaunch
' Author    : hecon5
' Date      : 2/05/2020
' Purpose   : Launch the main add-in form.
'---------------------------------------------------------------------------------------
'
Private Function AddInLaunch(RibbonCmdIn As Long) As Boolean
    PreloadVBE
    Form_frmVCSMain.HandleCmd RibbonCmdIn
    AddInLaunch = True
End Function

'---------------------------------------------------------------------------------------
' Procedure : AddInMenuItemLaunch
' Author    : Adam Waller
' Date      : 1/14/2020
' Purpose   : Launch the main add-in form.
'---------------------------------------------------------------------------------------
'
Public Function AddInMenuItemLaunch() As Boolean
    AddInLaunch erlVCSOpen
    AddInMenuItemLaunch = True
End Function


'---------------------------------------------------------------------------------------
' Procedure : AddInOptionsLaunch
' Author    : Hecon5
' Date      : 2/05/2020
' Purpose   : Launch the main add-in form.
'---------------------------------------------------------------------------------------
'
Public Function AddInOptionsLaunch() As Boolean
    AddInLaunch erlVCSOptions
    AddInOptionsLaunch = True
End Function


'---------------------------------------------------------------------------------------
' Procedure : AddInMenuItemExport
' Author    : Adam Waller
' Date      : 4/15/2020
' Purpose   : Open main form and start export immediately. (Save users a click)
'---------------------------------------------------------------------------------------
'
Public Function AddInMenuItemExport() As Boolean
    AddInLaunch erlExportAllRibbon
    AddInMenuItemExport = True
End Function


'---------------------------------------------------------------------------------------
' Procedure : RunExportForCurrentDB
' Author    : Adam Waller
' Date      : 11/10/2020
' Purpose   : The primary purpose of this function is to be able to use VBA code to
'           : initiate a source code export, without currupting the current DB. This
'           : would typically be used in a build automation environment, or when
'           : exporting code from the add-in itself.
'           : To avoid causing file corruption issues, we need to run the export using
'           : the installed add-in, not the local MSAccessVCS project. In order to do
'           : this, we need to load the VCS add-in at the application level, then
'           : make it the active VB Project, then call the export function. When the
'           : export function is called, we need to complete any running code in the
'           : current database before export, so we will use a timer callback to
'           : launch the export cleanly from the installed add-in.
'           : This sounds complicated, but it is critical that we don't attempt to
'           : export code from a module that is currently running, or it may corrupt
'           : the file and cause Access to crash the next time the file is opened.
'           : (This can be repaired by rebuilding from source, but let's work to
'           :  prevent the problem in the first place.)
'---------------------------------------------------------------------------------------
'
Public Function RunExportForCurrentDB()

    Dim projAddIn As VBProject
    
    ' Make sure the add-in is loaded.
    If Not AddinLoaded Then LoadVCSAddIn

    ' When exporting code from the add-in project itself, it gets a little
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
        Run "MSAccessVCS-Lib.RunExportForCurrentDB"
    Else
        ' Reset project name if needed
        With projAddIn
            ' Technically, changes in the add-in will not be saved anyway, so this
            ' may not be needed, but just in case we refer to this project by name
            ' anywhere else in the code, we will restore the original name before
            ' moving on to the actual export.
            If .Name = "MSAccessVCS-Lib" Then .Name = "MSAccessVCS"
        End With
        ' Call export function with an API callback.
        modTimer.LaunchExportAfterTimer
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


'---------------------------------------------------------------------------------------
' Procedure : GetAddInProject
' Author    : Adam Waller
' Date      : 11/10/2020
' Purpose   : Return the VBProject of the MSAccessVCS add-in.
'---------------------------------------------------------------------------------------
'
Public Function GetAddInProject() As VBProject
    Dim oProj As VBProject
    For Each oProj In VBE.VBProjects
        If StrComp(oProj.FileName, GetAddinFileName, vbTextCompare) = 0 Then
            Set GetAddInProject = oProj
            Exit For
        End If
    Next oProj
End Function


'---------------------------------------------------------------------------------------
' Procedure : LoadVCSAddIn
' Author    : Adam Waller
' Date      : 11/10/2020
' Purpose   : Load the add-in at the application level so it can stay active
'           : even if the current database is closed.
'           : https://stackoverflow.com/questions/62270088/how-can-i-launch-an-access-add-in-not-com-add-in-from-vba-code
'---------------------------------------------------------------------------------------
'
Private Sub LoadVCSAddIn()
    ' The following lines will load the add-in at the application level,
    ' but will not actually call the function. Ignore the error of function not found.
    If DebugMode(True) Then On Error Resume Next Else On Error Resume Next
    Application.Run GetAddinFileName & "!DummyFunction"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : IncrementAppVersion
' Author    : Adam Waller
' Date      : 1/6/2017
' Purpose   : Increments the build version (1.0.12)
'---------------------------------------------------------------------------------------
'
Public Sub IncrementAppVersion(ReleaseType As eReleaseType)
    
    Dim varParts As Variant
    Dim strFrom As String
    
    If ReleaseType = Same_Version Then Exit Sub
    strFrom = AppVersion
    varParts = Split(AppVersion, ".")
    varParts(ReleaseType) = varParts(ReleaseType) + 1
    If ReleaseType < Minor_xVx Then varParts(Minor_xVx) = 0
    If ReleaseType < Build_xxV Then varParts(Build_xxV) = 0
    AppVersion = Join(varParts, ".")

    ' Display old and new versions
    Debug.Print "Updated from " & strFrom & " to " & AppVersion

End Sub


'---------------------------------------------------------------------------------------
' Procedure : AppVersion
' Author    : Adam Waller
' Date      : 1/5/2017
' Purpose   : Get the version from the database property.
'---------------------------------------------------------------------------------------
'
Public Property Get AppVersion() As String
    Dim strVersion As String
    strVersion = GetDBProperty("AppVersion", CodeDb)
    If strVersion = vbNullString Then strVersion = "1.0.0"
    AppVersion = strVersion
End Property


'---------------------------------------------------------------------------------------
' Procedure : AppVersion
' Author    : Adam Waller
' Date      : 1/5/2017
' Purpose   : Set version property in current database.
'---------------------------------------------------------------------------------------
'
Public Property Let AppVersion(strVersion As String)
    SetDBProperty "AppVersion", strVersion, , CodeDb
End Property


'---------------------------------------------------------------------------------------
' Procedure : InstalledVersion
' Author    : Adam Waller
' Date      : 4/21/2020
' Purpose   : Returns the installed version of the add-in from the registry.
'           : (We are saving this in the user hive, since it requires admin rights
'           :  to change the keys actually used by Access to register the add-in)
'---------------------------------------------------------------------------------------
'
Public Property Let InstalledVersion(strVersion As String)
    SaveSetting GetCodeVBProject.Name, "Add-in", "Installed Version", strVersion
End Property
Public Property Get InstalledVersion() As String
    InstalledVersion = GetSetting(GetCodeVBProject.Name, "Add-in", "Installed Version", vbNullString)
End Property


'---------------------------------------------------------------------------------------
' Procedure : PreloadVBE
' Author    : Adam Waller
' Date      : 5/25/2020
' Purpose   : Force Access to load the VBE project. (This can help prevent crashes
'           : when code is run before the VB Project is fully loaded.)
'---------------------------------------------------------------------------------------
'
Public Sub PreloadVBE()
    Dim strName As String
    DoCmd.Hourglass True
    strName = VBE.ActiveVBProject.Name
    DoCmd.Hourglass False
End Sub