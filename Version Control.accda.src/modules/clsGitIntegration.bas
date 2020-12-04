Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


' Enum for commands we can run with Git.
Private Enum eGitCommand
    egcGetVersion
    egcGetHeadCommitDate
    egcGetCommittedFiles
    egcGetAllChangedFiles
    egcGetUntrackedFiles
    egcGetHeadCommit
    egcSetTaggedCommit
End Enum


' The structure of this dictionary is very similar to the VCS Index of components.
Private m_dChangedItems As Dictionary


' Peforms operations related to interrogating the status of Git
' Note: All of these operations make certain assumptions:
' 1) The database is in the root of the git repository.
' 2) Source code is in the source\ directory.



'---------------------------------------------------------------------------------------
' Procedure : RunGitCommand
' Author    : Adam Waller
' Date      : 11/24/2020
' Purpose   : Run a git command, and return the result.
'           : (Define the specific git commands in this function)
'---------------------------------------------------------------------------------------
'
Private Function RunGitCommand(intCmd As eGitCommand) As String

    Dim strCmd As String
    Dim strResult As String
    
    ' Translate enum to command
    Select Case intCmd
        Case egcGetHeadCommitDate:      strCmd = "git show -s --format=%ci HEAD"
        Case egcGetCommittedFiles:      strCmd = "git diff --name-status access-vcs-last-imported-commit..HEAD"
        Case egcGetUntrackedFiles:      strCmd = "git ls-files . --exclude-standard --others"
        Case egcGetVersion:             strCmd = "git version"
        Case egcSetTaggedCommit:        strCmd = "git tag access-vcs-last-imported-commit HEAD -f"
        Case egcGetAllChangedFiles:     strCmd = "git diff --name-status access-vcs-last-imported-commit"
        Case egcGetHeadCommit:          strCmd = "git show -s --format=%h HEAD"
    End Select

    ' Run command, and get result
    Perf.OperationStart "Git Commands"
    strResult = ShellRun(strCmd)
    Perf.OperationEnd
    
    ' Trim any trailing vbLf
    If Right$(strResult, 1) = vbLf Then strResult = Left$(strResult, Len(strResult) - 1)
    RunGitCommand = strResult
    
End Function



' Return the datestamp of the current head commit
Public Function GetHeadCommitDate() As Date

    Dim strDate As String
    Dim varParts As Variant
    
    ' Returns something like "2020-11-23 16:08:47 -0600"
    strDate = RunGitCommand(egcGetHeadCommitDate)
    
    ' convert the result from ISO 8601 to Access,
    ' trimming off the timezone at the end (should be local)
    ' see StackOverflow #38751429
    varParts = Split(strDate, " -")
    If IsDate(varParts(0)) Then GetHeadCommitDate = CDate(varParts(0))

End Function


'---------------------------------------------------------------------------------------
' Procedure : GetHeadCommitHash
' Author    : Adam Waller
' Date      : 11/24/2020
' Purpose   : Return the 7-character hash of the head commit.
'---------------------------------------------------------------------------------------
'
Public Function GetHeadCommitHash() As String
    GetHeadCommitHash = RunGitCommand(egcGetHeadCommit)
End Function


'---------------------------------------------------------------------------------------
' Procedure : ShellRun
' Author    : Adam Waller
' Date      : 11/24/2020
' Purpose   : Pass a git command to this function to return the result as a string.
'---------------------------------------------------------------------------------------
'
Private Function ShellRun(strCmd As String) As String
    
    Dim oShell As WshShell
    Dim strFile As String
    
    ' Get path to temp file
    strFile = GetTempFile
    
    ' Build command line string
    With New clsConcat
        ' Open command prompt in export folder
        .Add "cmd.exe /c cd ", Options.GetExportFolder
        ' Run git command
        .Add " & ", strCmd
        ' Output to temp file
        .Add " > """, strFile, """"
        ' Execute command
        Set oShell = New WshShell
        oShell.Run .GetStr, WshHide, True
        'Debug.Print .GetStr    ' To debug
    End With
    
    ' Read from temp file
    ShellRun = ReadFile(strFile)
    
    ' Remove temp file
    FSO.DeleteFile strFile

End Function


'
'' Returns a collcetion containing two lists:
'' first, of all the objects to modify or re-import based on the state of the git repo
'' second, of all the objects to delete based on the same
'' if getUncommittedFiles is false, files list is all files between the current HEAD
'' and the commit carrying the last-imported-commit tag that are in the
'' /source directory. if it is true, file list includes any uncommitted changes
'' Note: Last entries in file arrays will be empty.
'Public Function GetSourceFilesSinceLastImport(getUncommittedFiles As Boolean) As Variant
'    Dim FileListString As String
'    Dim AllFilesArray As Variant
'    Dim SourceFilesToImportCollection As Collection
'    Dim SourceFilesToRemoveCollection As Collection
'    Set SourceFilesToImportCollection = New Collection
'    Set SourceFilesToRemoveCollection = New Collection
'    Dim FileStatus As Variant
'    Dim CommandToRun As String
'    Dim File As Variant
'    Dim Status As String
'    Dim FileStatusSplit As Variant
'    Dim ReturnArray(2) As Variant
'
'    If getUncommittedFiles = True Then
'        CommandToRun = GetAllChangedFilesCommand
'    Else
'        CommandToRun = GetCommittedFilesCommand
'    End If
'
'    ' get files already committed (and staged, if flag passed)
'    FileListString = ShellRun(CommandToRun)
'
'    ' sanitize paths, determine the operation type, and add to relevant collection
'    For Each FileStatus In Split(FileListString, vbLf)
'        If FileStatus = "" Then Exit For
'
'        FileStatusSplit = Split(FileStatus, vbTab)
'        Status = Left(FileStatusSplit(0), 1) ' only first character actually indicates status; the rest is "score"
'        File = FileStatusSplit(1)
'
'        If File <> "" And File Like "source/*" Then
'            File = Replace(File, "/", "\")
'
'            ' overwrite/add modified, copied, added
'            If Status = "M" Or Status = "A" Or Status = "U" Then
'                SourceFilesToImportCollection.Add File
'            End If
'
'            ' overwrite result of rename or copy
'            If Status = "R" Or Status = "C" Then
'                ' add the result to the collection of import files
'                SourceFilesToImportCollection.Add Replace(FileStatusSplit(2), "/", "\")
'            End If
'
'            ' remove deleted objects and original renamed files
'            If Status = "D" Or Status = "R" Then
'                SourceFilesToRemoveCollection.Add File
'            End If
'        End If
'    Next
'
'    ' get and add untracked files
'    If getUncommittedFiles = True Then
'        FileListString = ShellRun(GetUntrackedFilesCommand)
'        For Each File In Split(FileListString, vbLf)
'            If File <> "" And File Like "source/*" Then
'                File = Replace(File, "/", "\")
'                SourceFilesToImportCollection.Add File
'            End If
'        Next
'    End If
'
'    Set ReturnArray(0) = SourceFilesToImportCollection
'    Set ReturnArray(1) = SourceFilesToRemoveCollection
'    GetSourceFilesSinceLastImport = ReturnArray
'End Function
'
'Public Sub SetLastImportedCommitToCurrent()
'    ShellRun SetTaggedCommitCommand
'End Sub


'---------------------------------------------------------------------------------------
' Procedure : GetChangedFileList
' Author    : Adam Waller
' Date      : 11/25/2020
' Purpose   : Returns a collection of the files that have been changed. Only includes
'           : source files used by VCS.
'---------------------------------------------------------------------------------------
'
Public Function GetChangedFileList() As Dictionary

    Dim varItems As Variant
    Dim strBasePath As String
    
    Set GetChangedFileList = New Dictionary
    
'    varitems = split(RunGitCommand(egcGetAllChangedFiles
    
    

End Function


'---------------------------------------------------------------------------------------
' Procedure : GitInstalled
' Author    : Adam Waller
' Date      : 11/24/2020
' Purpose   : Returns true if git is installed.
'---------------------------------------------------------------------------------------
'
Public Function GitInstalled() As Boolean
    ' Expecting something like "git version 2.29.2.windows.2"
    GitInstalled = InStr(1, RunGitCommand(egcGetVersion), "git version ") = 1
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetModifiedSourceFiles
' Author    : Adam Waller
' Date      : 11/21/2020
' Purpose   : Return the modified source file paths for this component type.
'---------------------------------------------------------------------------------------
'
Public Function GetModifiedSourceFiles(cCategory As IDbComponent) As Collection

    ' Make sure the changes are loaded from Git
    If m_dChangedItems Is Nothing Then LoadChangesFromGit
    
End Function


'---------------------------------------------------------------------------------------
' Procedure : LoadChangesFromGit
' Author    : Adam Waller
' Date      : 12/4/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Function LoadChangesFromGit()

    Dim dCategory As IDbComponent
    Dim dFolders As Dictionary
    
    Set m_dChangedItems = New Dictionary
    
    ' Build a list of source path folders by category.
    ' For single file items, it will be the full path.
    For Each dCategory In GetAllContainers
        With dCategory
            If .SingleFile Then
                dFolders.Add .BaseFolder & FSO.GetFileName(.SourceFile), .Category
            Else
                dFolders.Add .BaseFolder, .Category
            End If
            Set m_dChangedItems(.Category) = New Dictionary
        End With
    Next dCategory
    Debug.Print ConvertToJson(dFolders, "  ")
    
    ' Check for match on entire file, then on the folder to determine
    ' the type of source file.
    
    
End Function