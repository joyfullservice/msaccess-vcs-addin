'---------------------------------------------------------------------------------------
' Functions used by the Version Control System menu handlers.
' Added here to reduce the memory overhead if we are only loading the menu.
' (VBA loads code modules on the first call to a function in the module)
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
Option Private Module


'---------------------------------------------------------------------------------------
' Procedure : StartUp
' Author    : Adam Waller
' Date      : 5/15/2015
' Purpose   : Auto-load version control when opening the library file.
'---------------------------------------------------------------------------------------
'
Public Function StartUp()
    LoadVBEMenuForVCS
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
    SelectionInActiveProject = (Application.VBE.ActiveVBProject.fileName = CurrentProject.FullName)
End Function