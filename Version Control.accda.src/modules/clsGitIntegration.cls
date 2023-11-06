VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsGitIntegration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : clsGitIntegration
' Author    : Adam Waller
' Date      : 3/10/2023
' Purpose   : Commands and functions relating specifically to Git repositories.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit

' Hash for revision we are diffing from.
Public FromRevision As String

' Set this to true to output git commands and return to debug window
Public ShowDebug As Boolean
Public LogCommands As Boolean

' Working folder for git commands. (Used when operating outside the database project)
Public WorkingFolder As String

' Enum for commands we can run with Git.
Private Enum eGitCommand
    egcGetVersion
    egcGetHeadCommitDate
    egcGetCommittedFiles
    egcGetAllChangedFiles
    egcGetUntrackedFiles
    egcGetHeadCommit
    egcGetBranchName
    egcGetReproPath
    egcGetRevision
    egcGetStatusPorcelain
    egcIsInsideTree
    ' Action commands
    egcSetTaggedCommit
    egcInitialize
    egcAddAll
    egcCommit
    egcResetHard
    egcMerge
    egcMergeNoFastFwd
    egcCheckoutBranch
    egcCheckoutNewBranch
    egcCheckoutHeadToCurrent
    egcDeleteBranch
End Enum


' The structure of this dictionary is very similar to the VCS Index of components.
Private m_dChangedItems As Dictionary
Private m_strRepositoryRoot As String


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
Private Function RunGitCommand(intCmd As eGitCommand, Optional strArgument As String) As String

    Dim strCmd As String
    Dim strResult As String

    ' Translate enum to command
    Select Case intCmd
        Case egcGetHeadCommitDate:      strCmd = "git show -s --format=%ci HEAD"
        Case egcGetCommittedFiles:      strCmd = "git diff --name-status {MyArg}..HEAD"
        Case egcGetUntrackedFiles:      strCmd = "git ls-files . --exclude-standard --others"
        Case egcGetVersion:             strCmd = "git version"
        Case egcSetTaggedCommit:        strCmd = "git tag {MyArg} HEAD -f"
        Case egcGetAllChangedFiles:     strCmd = "git diff --name-status {MyArg}"
        Case egcGetBranchName:          strCmd = "git rev-parse --abbrev-ref HEAD"
        Case egcGetHeadCommit:          strCmd = "git show -s --format=%h HEAD"
        Case egcGetReproPath:           strCmd = "git rev-parse --show-toplevel"
        Case egcGetRevision:            strCmd = "git rev-parse --verify {MyArg}"
        Case egcGetStatusPorcelain:     strCmd = "git status --porcelain"
        Case egcInitialize:             strCmd = "git init"
        Case egcAddAll:                 strCmd = "git add --all"
        Case egcCommit:                 strCmd = "git commit --all --message ""{MyArg}"""
        Case egcResetHard:              strCmd = "git reset --hard HEAD^"
        Case egcMerge:                  strCmd = "git merge {MyArg}"
        Case egcMergeNoFastFwd:         strCmd = "git merge --no-ff {MyArg}"
        Case egcCheckoutBranch:         strCmd = "git checkout {MyArg}"
        Case egcCheckoutNewBranch:      strCmd = "git checkout -b {MyArg}"
        Case egcCheckoutHeadToCurrent:  strCmd = "git checkout HEAD~ ."
        Case egcDeleteBranch:           strCmd = "git branch --delete {MyArg}"
        Case egcIsInsideTree:           strCmd = "git rev-parse --is-inside-work-tree"
        Case Else
            Log.Error eelError, "Unrecognized Git Command Enum: " & intCmd
            Stop
    End Select

    ' Add argument, if supplied
    strCmd = Replace(strCmd, "{MyArg}", strArgument)

    ' Run command, and get result
    Perf.OperationStart "Git Command (id:" & intCmd & ")"
    strResult = ShellRun(strCmd, intCmd)
    Perf.OperationEnd

    ' Trim any trailing vbLf
    If Right$(strResult, 1) = vbLf Then strResult = Left$(strResult, Len(strResult) - 1)
    RunGitCommand = strResult

    ' Show debug info, if flag is on
    With New clsConcat
        .Add "Sent Git Command: ", strCmd
        If strResult <> vbNullString Then
            .Add "Returned:"
            .Add strResult
        End If
        If Me.LogCommands Then Log.Add .GetStr, False
        If Me.ShowDebug Then Debug.Print .GetStr
    End With

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
' Procedure : GetRepositoryRoot
' Author    : Adam Waller
' Date      : 5/8/2023
' Purpose   : Returns the path to the root of the repository.
'---------------------------------------------------------------------------------------
'
Public Function GetRepositoryRoot(Optional blnFallBackToWorking As Boolean = True) As String

    Static strLastFolder As String  ' Working folder
    Static strLastRoot As String    ' Repository Root

    Dim strWorking As String

    ' Determine the current working folder
    strWorking = GetWorkingFolder

    ' Make sure git is actually installed
    If Not Me.GitInstalled Then
        If blnFallBackToWorking Then GetRepositoryRoot = strWorking
        Exit Function
    End If

    ' On first call, we will attempt to get the repository root from the working
    ' folder, or the export folder if a working folder is not specified.
    If strLastRoot = vbNullString Or (strLastFolder <> strWorking) Then
        ' Use working folder, if specified, otherwise go with export folder.
        If strLastFolder <> strWorking Then
            ' We will verify the repository path on the next call.
            strLastFolder = strWorking
            strLastRoot = vbNullString
            ' Recursively call this function to verify the path with git
            GetRepositoryRoot = GetRepositoryRoot(blnFallBackToWorking)
        Else
            ' Run git command from last folder
            strLastRoot = strLastFolder
            ' Use Git to look up root folder in repository.
            strLastRoot = Replace(RunGitCommand(egcGetReproPath), "/", PathSep) & PathSep
            If strLastRoot = PathSep Then
                If blnFallBackToWorking Then
                    ' Might not be in a git repository. Fall back to working folder.
                    GetRepositoryRoot = strWorking
                    strLastRoot = strWorking
                Else
                    GetRepositoryRoot = vbNullString
                End If
            Else
                ' Found the root folder. Return to caller.
                GetRepositoryRoot = strLastRoot
            End If
        End If
    Else
        ' Return cached root folder
        GetRepositoryRoot = strLastRoot
    End If

End Function


'---------------------------------------------------------------------------------------
' Procedure : GetWorkingFolder
' Author    : Adam Waller
' Date      : 5/8/2023
' Purpose   : Return the working folder from either the specified folder, or fall back
'           : to the export path.
'---------------------------------------------------------------------------------------
'
Private Function GetWorkingFolder() As String

    Dim strWorking As String

    ' Avoid calling Options if the working folder is already defined to prevent
    ' a possible stack overflow. (That's why we don't use Nz2() here)
    If Len(Me.WorkingFolder) Then
        strWorking = Me.WorkingFolder
    Else
        strWorking = Options.GetExportFolder
    End If

    ' Return path in consistent format
    GetWorkingFolder = StripSlash(strWorking) & PathSep

End Function


'---------------------------------------------------------------------------------------
' Procedure : Version
' Author    : Adam Waller
' Date      : 3/10/2023
' Purpose   : Return git version (Cached between calls)
'---------------------------------------------------------------------------------------
'
Public Function Version() As String
    Static strVersion As String
    If strVersion = vbNullString Then strVersion = Replace(RunGitCommand(egcGetVersion), "git version ", vbNullString)
    Version = strVersion
End Function


'---------------------------------------------------------------------------------------
' Procedure : IsInsideRepository
' Author    : Adam Waller
' Date      : 11/6/2023
' Purpose   : Returns true if the current working folder is inside a git repository.
'---------------------------------------------------------------------------------------
'
Public Function IsInsideRepository() As Boolean
    Dim strResult As String
    If Me.GitInstalled Then
        IsInsideRepository = (RunGitCommand(egcIsInsideTree) = "true")
    End If
End Function


'---------------------------------------------------------------------------------------
' Procedure : IsCleanBranch
' Author    : Adam Waller
' Date      : 3/10/2023
' Purpose   : Returns true if the branch has no changes or untracked files.
'           : See discussion on StackOverflow on pros and cons of different approaches
'           : for this. I went with simple for our purposes.
'           : See: https://stackoverflow.com/questions/2657935
'---------------------------------------------------------------------------------------
'
Public Function IsCleanBranch() As Boolean
    IsCleanBranch = (RunGitCommand(egcGetStatusPorcelain) = vbNullString)
End Function


'---------------------------------------------------------------------------------------
' Procedure : BranchName
' Author    : Adam Waller
' Date      : 3/21/2023
' Purpose   : Return the name of the current branch
'---------------------------------------------------------------------------------------
'
Public Function BranchName() As String
    BranchName = RunGitCommand(egcGetBranchName)
End Function


'---------------------------------------------------------------------------------------
' Procedure : SplitFilesWithHistory
' Author    : Adam Waller
' Date      : 3/10/2023
' Purpose   : Accepts two arrays of file paths. The first array represents the existing
'           : files, and the second represents the new files that will take on the
'           : content and history of the first files.
'           : Based on https://devblogs.microsoft.com/oldnewthing/20190919-00/?p=102904
'---------------------------------------------------------------------------------------
'
Public Sub SplitFilesWithHistory(strFilePathsArray() As String, strNewPaths() As String, _
    strCommitMessage As String, Optional strRunInFolder As String = vbNullString)

    Dim lngCnt As Long
    Dim strOrig As String
    Dim strNew As String
    Dim strBranch As String
    Dim blnLog As Boolean
    Dim strBaseFolder As String

    ' Save current value and turn on logging.
    blnLog = Me.LogCommands
    strBaseFolder = m_strRepositoryRoot
    m_strRepositoryRoot = strRunInFolder
    Me.LogCommands = True

    ' We should be starting with a clean slate (no uncommitted changes)
    If Not Me.IsCleanBranch Then
        Log.Error eelCritical, "Cannot split files in Git when changes are present in the branch", _
            ModuleName(Me) & ".SplitFilesWithHistory"
        GoTo CleanUp
    End If

    ' Initialize the repository just to make sure everthing is in order
    RunGitCommand egcInitialize

    ' Get the current branch name, so we can switch back to it later
    strBranch = Me.BranchName

    ' Create a new branch to use when splitting the files
    RunGitCommand egcCheckoutNewBranch, "split-files"

    ' We could use git to move the files individually, but this would be much slower.
    ' For our purposes, the files should be pretty unique in content, so we should be
    ' fine to move them in batches. (As passed to this function)
    For lngCnt = 0 To UBound(strFilePathsArray)
        strOrig = strFilePathsArray(lngCnt)
        strNew = strNewPaths(lngCnt)
        If FSO.FileExists(strOrig) Then
            If FSO.FileExists(strNew) Then DeleteFile strNew
            FSO.MoveFile strOrig, strNew
        End If
    Next lngCnt

    ' Update the index, and commit the move
    RunGitCommand egcAddAll
    RunGitCommand egcCommit, "Rename as new files"

    ' Restore the original files
    ' (TODO: see if we can do this from a named branch)
    RunGitCommand egcCheckoutHeadToCurrent

    ' Commit these files (with history) to the temporary branch
    RunGitCommand egcCommit, "Restore original files"

    ' Move back to original branch
    RunGitCommand egcCheckoutBranch, strBranch

    ' Merge the temporary branch into the original branch
    RunGitCommand egcMergeNoFastFwd, "split-files" & " -m """ & strCommitMessage & """"

    ' Delete the temporary branch now that we are finished using it
    RunGitCommand egcDeleteBranch, "split-files"

    ' Restore original logging
    Me.LogCommands = blnLog

CleanUp:

    ' Restore base folder
    m_strRepositoryRoot = strBaseFolder

End Sub


'---------------------------------------------------------------------------------------
' Procedure : ShellRun
' Author    : Adam Waller
' Date      : 11/24/2020
' Purpose   : Pass a git command to this function to return the result as a string.
'---------------------------------------------------------------------------------------
'
Private Function ShellRun(strCmd As String, intCmd As eGitCommand) As String

    Dim oShell As WshShell
    Dim strFile As String

    ' Get path to temp file
    strFile = GetTempFile

    ' Build command line string
    With New clsConcat
        Select Case intCmd
            Case egcGetVersion
                ' Run independent of repository
                .Add "cmd.exe /c ", strCmd
            Case Else
                ' Open command prompt in repository folder
                .Add "cmd.exe /c cd ", GetRepositoryRoot
                ' Run git command
                .Add " & ", strCmd
        End Select
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
' Procedure : GitInstalled
' Author    : Adam Waller
' Date      : 11/24/2020
' Purpose   : Returns true if git is installed.
'---------------------------------------------------------------------------------------
'
Public Function GitInstalled() As Boolean
    GitInstalled = (Len(Me.Version))
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetModifiedSourceFiles
' Author    : Adam Waller
' Date      : 11/21/2020
' Purpose   : Return the modified source file paths for this component type.
'---------------------------------------------------------------------------------------
'
Public Function GetModifiedSourceFiles(cCategory As IDbComponent) As Collection

    Dim varKey As Variant

    ' Make sure the changes are loaded from Git
    If m_dChangedItems Is Nothing Then Set m_dChangedItems = GetChangedFileIndex(Me.FromRevision)

    ' Check for any matching changes.
    Set GetModifiedSourceFiles = New Collection
    With m_dChangedItems
        If .Exists(cCategory.Category) Then
            For Each varKey In .Item(cCategory.Category).Keys
                ' Add source file
                GetModifiedSourceFiles.Add CStr(varKey)
            Next varKey
        End If
    End With

End Function


'---------------------------------------------------------------------------------------
' Procedure : RevisionExists
' Author    : Adam Waller
' Date      : 1/19/2021
' Purpose   : Returns true if the revision exists on Git.
'---------------------------------------------------------------------------------------
'
Public Function RevisionExists(strHash As String) As Boolean
    RevisionExists = (RunGitCommand(egcGetRevision, strHash) <> vbNullString)
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetChangedFileList
' Author    : Adam Waller
' Date      : 11/25/2020
' Purpose   : Returns a collection of the files that have been changed. Only includes
'           : source files used by VCS.
'---------------------------------------------------------------------------------------
'
Public Function GetChangedFileIndex(strFromCommit As String) As Dictionary

    Dim varItems As Variant
    Dim varFile As Variant
    Dim strPath As String
    Dim strBasePath As String
    Dim varParts As Variant
    Dim strExportFolder As String
    Dim strCategory As String
    Dim dIndex As Dictionary
    Dim dFolders As Dictionary
    Dim dCategory As Dictionary
    Dim cComp As IDbComponent
    Dim strSourceFile As String
    Dim strFlag As String
    Dim strRootPath As String
    Dim strResponse As String

    ' Get the base export folder
    strExportFolder = Options.GetExportFolder
    varParts = Split(strExportFolder, PathSep)
    strBasePath = varParts(UBound(varParts) - 1)
    strRootPath = GetRepositoryRoot

    ' Get base folder list from component types.
    ' (Used to organize the changed files by type)
    Set dFolders = New Dictionary
    For Each cComp In GetContainers
        strCategory = StripSlash(Mid$(cComp.BaseFolder, Len(strRootPath) + 1))
        If strCategory = strBasePath Then
            ' Include file name in category
            strCategory = Mid$(cComp.SourceFile, Len(strRootPath) + 1)
        End If
        ' Replace backslashes with forward slashes to match git output
        strCategory = Replace(strCategory, PathSep, "/")
        dFolders.Add strCategory, cComp.Category
    Next cComp

    ' Windows 10 can optionally support case-sensitive file names, but for
    ' now we will go with case insensitive names for the purpose of the index.
    Set dIndex = New Dictionary
    dIndex.CompareMode = TextCompare

    ' Return a list of changed and new files from git.
    strResponse = RunGitCommand(egcGetAllChangedFiles, strFromCommit) & vbLf & _
        RunGitCommand(egcGetUntrackedFiles)

    ' Check for errors such as invalid commit
    If InStr(1, strResponse, ": unknown revision") > 0 Then
        Log.Error eelError, "Unknown git revision: " & strFromCommit, "clsGitIntegration.GetChangedFileIndex"
        Log.Spacer False
        Log.Add strResponse, False
        Log.Spacer
    Else
        ' Convert to list of items
        varItems = Split(strResponse, vbLf)

        ' Loop through list of changed files
        For Each varFile In varItems

            ' Check for flag from changed files.
            If Mid(varFile, 2, 1) = vbTab Then
                strFlag = Mid(varFile, 1, 1)
                strPath = Mid(varFile, 3)
            Else
                strFlag = "U" ' Unversioned file.
                strPath = varFile
            End If

            ' Skip any blank lines
            If strPath <> vbNullString Then

                ' Check for match on entire file name. (For single file items
                ' in the root export folder.)
                If dFolders.Exists(strPath) Then
                    ' Use this component type.
                    strCategory = dFolders(strPath)
                Else
                    ' Use the folder name to look up component type.
                    strCategory = dNZ(dFolders, FSO.GetParentFolderName(strPath))
                End If

                ' Ignore files outside standard VCS source folders.
                If strCategory <> vbNullString Then

                    ' Add to index of changed files.
                    With dIndex

                        ' Add category if it does not exist.
                        If Not .Exists(strCategory) Then
                            Set dCategory = New Dictionary
                            .Add strCategory, dCategory
                        End If

                        ' Build full path to source file, and add to index.
                        strSourceFile = strRootPath & Replace(strPath, "/", PathSep)

                        ' Add full file path to category, including flag with change type.
                        .Item(strCategory).Add strSourceFile, strFlag
                    End With
                End If
            End If
        Next varFile

    End If

    ' Return dictionary of file paths.
    Set GetChangedFileIndex = dIndex

End Function
