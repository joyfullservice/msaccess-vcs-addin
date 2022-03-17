VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsVersionControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'---------------------------------------------------------------------------------------
' Module    : clsVersionControl
' Author    : Adam Waller
' Date      : 1/13/2021
' Purpose   : This class is publicly exposed through the VCS object in modAPI and can
'           : be used to automate add-in functionality from other systems.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit


'---------------------------------------------------------------------------------------
' Procedure : Show
' Author    : Adam Waller
' Date      : 1/13/2021
' Purpose   : Show the Version Control System main form
'---------------------------------------------------------------------------------------
'
Public Sub Show()
    AddInMenuItemLaunch
End Sub


'---------------------------------------------------------------------------------------
' Procedure : ShowOptions
' Author    : Adam Waller
' Date      : 3/4/2022
' Purpose   : Show the options form
'---------------------------------------------------------------------------------------
'
Public Sub ShowOptions()
    ' Force reload of options from current project before opening the form.
    Set modObjects.Options = Nothing
    Form_frmVCSOptions.Visible = True
End Sub


'---------------------------------------------------------------------------------------
' Procedure : Export
' Author    : Adam Waller
' Date      : 1/13/2021
' Purpose   : Export the source code for the current database
'---------------------------------------------------------------------------------------
'
Public Sub Export()
    AddInMenuItemExport
End Sub


'---------------------------------------------------------------------------------------
' Procedure : GetOptions
' Author    : Adam Waller
' Date      : 1/13/2021
' Purpose   : Access the options
'---------------------------------------------------------------------------------------
'
Public Function Options() As clsOptions
    Set Options = modObjects.Options
End Function


'---------------------------------------------------------------------------------------
' Procedure : Version
' Author    : Adam Waller
' Date      : 1/13/2021
' Purpose   : Return the version of this (CodeDB) instance of VCS.
'---------------------------------------------------------------------------------------
'
Public Property Get Version() As String
    Version = GetVCSVersion
End Property



'---------------------------------------------------------------------------------------
' Procedure : GoToLink...
' Author    : Adam Waller
' Date      : 3/4/2022
' Purpose   : Wrapper functions to visit online links (called from Ribbon menu)
'---------------------------------------------------------------------------------------
'
Public Sub GoToLinkHome()
    FollowHyperlink "https://github.com/joyfullservice/msaccess-vcs-integration"
End Sub
Public Sub GoToLinkDocumentation()
    FollowHyperlink "https://github.com/joyfullservice/msaccess-vcs-integration/wiki"
End Sub
Public Sub GoToLinkSupport()
    FollowHyperlink "https://github.com/joyfullservice/msaccess-vcs-integration/issues"
End Sub
Public Sub GoToLinkDownload()
    FollowHyperlink "https://github.com/joyfullservice/msaccess-vcs-integration/releases"
End Sub

