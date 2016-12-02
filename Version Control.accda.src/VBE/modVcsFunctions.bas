Attribute VB_Name = "modVcsFunctions"
'---------------------------------------------------------------------------------------
' Functions used by the Version Control System menu handlers.
' Added here to reduce the memory overhead if we are only loading the menu.
' (VBA loads code modules on the first call to a function in the module)
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
Option Private Module


'---------------------------------------------------------------------------------------
' Procedure : ShowIDE
' Author    : Adam Waller
' Date      : 5/18/2015
' Purpose   : Show the VBA code editor (used in autoexec macro)
'---------------------------------------------------------------------------------------
'
Public Function ShowIDE()
    DoCmd.RunCommand acCmdVisualBasicEditor
End Function


'---------------------------------------------------------------------------------------
' Procedure : ProgramFilesFolder
' Author    : Adam Waller
' Date      : 5/15/2015
' Purpose   : Returns the program files folder on the OS. (32 or 64 bit)
'---------------------------------------------------------------------------------------
'
Public Function ProgramFilesFolder() As String
    Dim strFolder As String
    strFolder = Environ$("PROGRAMFILES")
    ' Should always work, but just in case!
    If strFolder = "" Then strFolder = "C:\Program Files (x86)"
    ProgramFilesFolder = strFolder & "\"
End Function


'---------------------------------------------------------------------------------------
' Procedure : ProjectIsSelected
' Author    : Adam Waller
' Date      : 5/15/2015
' Purpose   : Returns true if the base project is selected in the VBE
'---------------------------------------------------------------------------------------
'
Public Function ProjectIsSelected() As Boolean
    ProjectIsSelected = (Application.VBE.SelectedVBComponent Is Nothing)
End Function


'---------------------------------------------------------------------------------------
' Procedure : SelectionInActiveProject
' Author    : Adam Waller
' Date      : 5/15/2015
' Purpose   : Returns true if the current selection is in the active project
'---------------------------------------------------------------------------------------
'
Public Function SelectionInActiveProject() As Boolean
    SelectionInActiveProject = (Application.VBE.ActiveVBProject.fileName = UncPath(CurrentProject.FullName))
End Function


'---------------------------------------------------------------------------------------
' Procedure : UncPath
' Author    : Adam Waller
' Date      : 5/18/2015
' Purpose   : Returns the UNC path of a mapped network drive, if applicable
'---------------------------------------------------------------------------------------
'
Public Function UncPath(strPath As String) As String
    
    Dim objFSO As Object ' FileSystemObject
    Dim strDrive As String
    Dim strShare As String
    
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    
    ' Identify drive letter and share name
    With objFSO
        strDrive = .GetDriveName(.GetAbsolutePathName(strPath))
        strShare = .GetDrive(strDrive).ShareName
    End With
    
    If strShare <> "" Then
        ' Replace drive with UNC path
        UncPath = strShare & Mid(strPath, Len(strDrive) + 1)
    Else
        ' Return unmodified path
        UncPath = strPath
    End If
    
    ' Clean up references
    Set objFSO = Nothing
    
End Function



