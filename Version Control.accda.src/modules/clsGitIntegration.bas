Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' Hash for revision we are diffing from.
Public FromRevision As String

' Enum for commands we can run with Git.
Private Enum eGitCommand
    egcGetVersion
    egcGetHeadCommitDate
    egcGetCommittedFiles
    egcGetAllChangedFiles
    egcGetUntrackedFiles
    egcGetHeadCommit
    egcSetTaggedCommit
    egcGetReproPath
    egcGetRevision
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
        Case egcGetHeadCommit:          strCmd = "git show -s --format=%h HEAD"
        Case egcGetReproPath:           strCmd = "git rev-parse --show-toplevel"
        Case egcGetRevision:            strCmd = "git rev-parse --verify {MyArg}"
    End Select

    ' Add argument, if supplied
    strCmd = Replace(strCmd, "{MyArg}", strArgument)

    ' Run command, and get result
    Perf.OperationStart "Git Command (id:" & intCmd & ")"
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
' Procedure : GetRepositoryPath
' Author    : Adam Waller
' Date      : 1/19/2021
' Purpose   : Returns the path to the root of the repository.
'---------------------------------------------------------------------------------------
'
Public Function GetRepositoryPath() As String
    If m_strRepositoryRoot = vbNullString Then
        m_strRepositoryRoot = Replace(RunGitCommand(egcGetReproPath), "/", PathSep) & PathSep
    End If
    GetRepositoryPath = m_strRepositoryRoot
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
    Dim strWorking As String
    
    ' Get working path for command. (Prefer repository root, if available)
    If m_strRepositoryRoot = vbNullString Then
        strWorking = Options.GetExportFolder
    Else
        strWorking = m_strRepositoryRoot
    End If
    
    ' Get path to temp file
    strFile = GetTempFile
    
    ' Build command line string
    With New clsConcat
        ' Open command prompt in export folder
        .Add "cmd.exe /c cd ", strWorking
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
    strRootPath = GetRepositoryPath

    ' Get base folder list from component types.
    ' (Used to organize the changed files by type)
    Set dFolders = New Dictionary
    For Each cComp In GetAllContainers
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


'---------------------------------------------------------------------------------------
' Procedure : Class_Initialize
' Author    : Adam Waller
' Date      : 1/19/2021
' Purpose   : Load path to root repository
'---------------------------------------------------------------------------------------
'
Private Sub Class_Initialize()
    GetRepositoryPath
End Sub