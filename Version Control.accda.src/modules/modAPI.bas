Option Compare Database
Option Explicit

'Wrapper module for exposing selected routines to the referencing database.
'Required especially for when users cannot load the addin as an application
'level addin.

'VCS Functions (Addin specific)
'These load / unload the VCS Addin, return the installed version, etc.
Public Function LaunchVCS() As Boolean
    LaunchVCS = AddInMenuItemLaunch
End Function

Public Property Get VCSVersion() As String
    VCSVersion = InstalledVersion
End Property

'Project Functions (these export/import code for the referencing database
' or build it, etc.)
Public Sub ExportSourceAPI()
    RunExportForCurrentDB
End Sub
