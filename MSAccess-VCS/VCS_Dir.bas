Attribute VB_Name = "VCS_Dir"
Option Compare Database

Option Private Module
Option Explicit


' Path/Directory of the current database file.
Public Function VCS_ProjectPath() As String
    VCS_ProjectPath = CurrentProject.Path
    If Right$(VCS_ProjectPath, 1) <> "\" Then VCS_ProjectPath = VCS_ProjectPath & "\"
End Function

' Create folder `Path`. Silently do nothing if it already exists.
Public Sub VCS_MkDirIfNotExist(ByVal Path As String)
    On Error GoTo MkDirIfNotexist_noop
    MkDir Path
MkDirIfNotexist_noop:
    On Error GoTo 0
End Sub

' Delete a file if it exists.
Public Sub VCS_DelIfExist(ByVal Path As String)
    On Error GoTo DelIfNotExist_Noop
    Kill Path
DelIfNotExist_Noop:
    On Error GoTo 0
End Sub

' Erase all *.`ext` files in `Path`.
Public Sub VCS_ClearTextFilesFromDir(ByVal Path As String, ByVal Ext As String)
    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    If Not FSO.FolderExists(Path) Then Exit Sub

    On Error GoTo VCS_ClearTextFilesFromDir_noop
    If Dir$(Path & "*." & Ext) <> vbNullString Then
        FSO.DeleteFile Path & "*." & Ext
    End If
    
VCS_ClearTextFilesFromDir_noop:
    On Error GoTo 0
End Sub

Public Function VCS_FileExists(ByVal strPath As String) As Boolean
    On Error Resume Next
    VCS_FileExists = False
    VCS_FileExists = ((GetAttr(strPath) And vbDirectory) <> vbDirectory)
End Function