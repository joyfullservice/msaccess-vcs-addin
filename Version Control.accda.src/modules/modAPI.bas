Option Compare Database
Option Explicit

'Exposes the release types to referencing files.
Public Enum eReleaseType
    Major_Vxx = 0
    Minor_xVx = 1
    Build_xxV = 2
    Same_Version = 3
End Enum

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
Public Property Get myAppVersion() As String
    myAppVersion = AppVersion
End Property

Public Property Let myAppVersion(strVersion As String)
    AppVersion = strVersion
End Property

Public Sub incrementMyAppVersion(ReleaseType As eReleaseType)
    IncrementAppVersion (ReleaseType)
End Sub

Public Sub ExportSourceAPI()
    RunExportForCurrentDB
End Sub

Public Sub DeployMyApp(Optional ReleaseType As eReleaseType = Build_xxV)
    Deploy (ReleaseType)
End Sub
