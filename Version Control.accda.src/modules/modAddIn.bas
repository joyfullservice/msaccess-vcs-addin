Option Compare Database
Option Explicit

' Used to determine if Access is running as administrator. (Required for installing the add-in)
Private Declare PtrSafe Function IsUserAnAdmin Lib "shell32" () As Long


'---------------------------------------------------------------------------------------
' Procedure : AddInMenuItemLaunch
' Author    : Adam Waller
' Date      : 1/14/2020
' Purpose   : Launch the main add-in form.
'---------------------------------------------------------------------------------------
'
Public Function AddInMenuItemLaunch()
    DoCmd.OpenForm "frmMain"
End Function


'---------------------------------------------------------------------------------------
' Procedure : InstallVCSAddin
' Author    : Adam Waller
' Date      : 1/14/2020
' Purpose   : Installs the add-in for the current user.
'---------------------------------------------------------------------------------------
'
Public Sub InstallVCSAddin()
    
    Dim strSource As String
    Dim strDest As String

    Dim blnExists As Boolean
    
    strSource = CodeProject.FullName
    strDest = Environ("AppData") & "\Microsoft\AddIns\" & CodeProject.Name
    
    ' Copy the file, overwriting any existing file.
    ' Requires FSO to copy open database files. (VBA.FileCopy give a permission denied error.)
    FSO.CopyFile strSource, strDest, True
    
    
End Sub