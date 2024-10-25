Attribute VB_Name = "modFileAccess"
'---------------------------------------------------------------------------------------
' Module    : modFileAccess
' Author    : Adam Waller
' Date      : 12/4/2020
' Purpose   : General functions for reading and writing files, building and verifying
'           : paths, and parsing file names.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Private Module
Option Explicit

Private Const ModuleName As String = "modFileAccess"

Private Declare PtrSafe Function getTempPath Lib "kernel32" Alias "GetTempPathA" ( _
    ByVal nBufferLength As Long, _
    ByVal lpBuffer As String) As Long

Private Declare PtrSafe Function getTempFileName Lib "kernel32" Alias "GetTempFileNameA" ( _
    ByVal lpszPath As String, _
    ByVal lpPrefixString As String, _
    ByVal wUnique As Long, _
    ByVal lpTempFileName As String) As Long

Private Declare PtrSafe Function SHCreateDirectoryEx Lib "shell32" Alias "SHCreateDirectoryExW" ( _
    ByVal hwnd As LongPtr _
    , ByVal pszPath As LongPtr _
    , ByVal psa As Any) As Long

'---------------------------------------------------------------------------------------
' Procedure : GetTempFile
' Author    : Adapted by Adam Waller
' Date      : 1/23/2019
' Purpose   : Generate Random / Unique temporary file name. (Also creates the file)
'---------------------------------------------------------------------------------------
'
Public Function GetTempFile(Optional strPrefix As String = "VBA") As String

    Dim strPath As String * 512
    Dim strName As String * 576
    Dim lngReturn As Long

    lngReturn = getTempPath(512, strPath)
    lngReturn = getTempFileName(strPath, strPrefix, 0, strName)
    If lngReturn <> 0 Then GetTempFile = Left$(strName, InStr(strName, vbNullChar) - 1)

End Function


'---------------------------------------------------------------------------------------
' Procedure : GetTempFolder
' Author    : Adam Waller
' Date      : 9/24/2021
' Purpose   : Get a random unique folder name and create the folder.
'---------------------------------------------------------------------------------------
'
Public Function GetTempFolder(Optional strPrefix As String = "VBA") As String

    Dim strPath As String
    Dim strFile As String
    Dim strFolder As String

    ' Generate a random temporary file name, and delete the temp file
    strPath = GetTempFile(strPrefix)
    DeleteFile strPath

    ' Change path to use underscore instead of period.
    strFile = PathSep & FSO.GetFileName(strPath)
    strFolder = Replace(strFile, ".", "_")
    strPath = Replace(strPath, strFile, strFolder)

    If FSO.FolderExists(strPath) Then
        ' Oops, this folder already exists. Try again.
        GetTempFolder = GetTempFolder(strPrefix)
    Else
        ' Create folder and return path
        FSO.CreateFolder strPath
        GetTempFolder = strPath
    End If

End Function


'---------------------------------------------------------------------------------------
' Procedure : ReadFile
' Author    : Adam Waller / Indigo
' Date      : 11/4/2020
' Purpose   : Read text file.
'           : Read in UTF-8 encoding, removing a BOM if found at start of file.
'---------------------------------------------------------------------------------------
'
Public Function ReadFile(strPath As String, Optional strCharset As String = "utf-8") As String

    Dim cData As clsConcat

    Set cData = New clsConcat

    If FSO.FileExists(strPath) Then
        Perf.OperationStart "Read File"
        With New ADODB.Stream
            .Charset = strCharset
            .Open
            .LoadFromFile strPath
            ' Read chunks of text, rather than the whole thing at once for massive
            ' performance gains when reading large files.
            ' See https://docs.microsoft.com/is-is/sql/ado/reference/ado-api/readtext-method
            Do While Not .EOS
                cData.Add .ReadText(CHUNK_SIZE) ' 128K
            Loop
            .Close
        End With
        Perf.OperationEnd
    End If

    ' Return text contents of file.
    ReadFile = cData.GetStr

End Function


'---------------------------------------------------------------------------------------
' Procedure : WriteFile
' Author    : Adam Waller
' Date      : 1/23/2019
' Purpose   : Save string variable to text file. (Building the folder path if needed)
'           : Saves in UTF-8 encoding, adding a BOM if extended or unicode content
'           : is found in the file. https://stackoverflow.com/a/53036838/4121863
'---------------------------------------------------------------------------------------
'
Public Sub WriteFile(strText As String, strPath As String, Optional strEncoding As String = "utf-8")

    Perf.OperationStart "Write File"

    ' Write to a UTF-8 eoncoded file
    With New ADODB.Stream
        .Type = adTypeText
        .Open
        .Charset = strEncoding
        .WriteText strText
        ' Ensure that we are ending the content with a vbcrlf
        If Right(strText, 2) <> vbCrLf Then .WriteText vbCrLf
        ' Write to disk
        VerifyPath strPath
        ' Watch out for possible write error
        LogUnhandledErrors
        On Error Resume Next
        .SaveToFile strPath, adSaveCreateOverWrite
        If Catch(3004) Then
            ' File is locked. Try again after 1 second, just in case something
            ' like Google Drive momentarily locked the file.
            Err.Clear
            Pause 1
            .SaveToFile strPath, adSaveCreateOverWrite
        End If
        CatchAny eelError, "Error writing file: " & strPath, ModuleName & ".WriteFile"
        .Close
    End With

    Perf.OperationEnd

End Sub


'---------------------------------------------------------------------------------------
' Procedure : GetFileBytes
' Author    : Adam Waller
' Date      : 7/31/2020
' Purpose   : Returns a byte array of the file contents.
'           : This function supports Unicode paths, unlike VBA's Open statement.
'---------------------------------------------------------------------------------------
'
Public Function GetFileBytes(strPath As String, Optional lngBytes As Long = adReadAll) As Byte()
    Perf.OperationStart "Read File Bytes"
    With New ADODB.Stream
        .Type = adTypeBinary
        .Open
        .LoadFromFile strPath
        GetFileBytes = .Read(lngBytes)
        .Close
    End With
    Perf.OperationEnd
End Function


'---------------------------------------------------------------------------------------
' Procedure : WriteBinaryFile
' Author    : Adam Waller
' Date      : 7/9/2021
' Purpose   : Writes the file bytes to a file (with Unicode path support)
'---------------------------------------------------------------------------------------
'
Public Function WriteBinaryFile(strPath As String, bteArray() As Byte)
    Perf.OperationStart "Write Binary File"
    With New ADODB.Stream
        .Type = adTypeBinary
        .Open
        .Write bteArray
        VerifyPath strPath
        .SaveToFile strPath, adSaveCreateOverWrite
        .Close
    End With
    Perf.OperationEnd
End Function


'---------------------------------------------------------------------------------------
' Procedure : DeleteFile
' Author    : Adam Waller
' Date      : 11/5/2020
' Purpose   : Wrapper to delete file while monitoring performance.
'---------------------------------------------------------------------------------------
'
Public Sub DeleteFile(strFile As String, Optional blnForce As Boolean = True)
    Perf.OperationStart "Delete File"
    FSO.DeleteFile strFile, blnForce
    Perf.OperationEnd
End Sub


'---------------------------------------------------------------------------------------
' Procedure : MoveFileIfExists
' Author    : Adam Waller
' Date      : 9/10/2022
' Purpose   : Moves a file to a specified destination folder, creating the destination
'           : folder if it does not exist.
'---------------------------------------------------------------------------------------
'
Public Sub MoveFileIfExists(strFilePath As String, strToFolder As String)
    Dim strNewPath As String
    If FSO.FileExists(strFilePath) Then
        Perf.OperationStart "Move File"
        VerifyPath strToFolder
        strNewPath = StripSlash(strToFolder) & PathSep & FSO.GetFileName(strFilePath)
        If FSO.FileExists(strNewPath) Then DeleteFile strNewPath
        FSO.MoveFile strFilePath, strNewPath
        Perf.OperationEnd
    End If
End Sub


'---------------------------------------------------------------------------------------
' Procedure : MoveFolderIfExists
' Author    : Adam Waller
' Date      : 9/10/2022
' Purpose   : Move a folder to a new location, replacing any existing folder.
'---------------------------------------------------------------------------------------
'
Public Sub MoveFolderIfExists(strFolderPath As String, strToParentFolder As String)
    Dim strNewPath As String
    If FSO.FolderExists(strFolderPath) Then
        Perf.OperationStart "Move Folder"
        VerifyPath strToParentFolder
        strNewPath = StripSlash(strToParentFolder) & PathSep & FSO.GetFolder(strFolderPath).Name
        If FSO.FolderExists(strNewPath) Then FSO.DeleteFolder strNewPath, True
        FSO.MoveFolder strFolderPath, strNewPath
        Perf.OperationEnd
    End If
End Sub


'---------------------------------------------------------------------------------------
' Procedure : clearfilesbyextension
' Author    : Adam Waller
' Date      : 1/25/2019
' Purpose   : Erase all *.`ext` files in `Path`.
'---------------------------------------------------------------------------------------
'
Public Sub ClearFilesByExtension(ByVal strFolder As String, strExt As String)

    Dim oFile As Scripting.File
    Dim strFolderNoSlash As String

    ' While the Dir() function would be simpler, it does not support Unicode.
    Perf.OperationStart "Clear Files by Ext"
    strFolderNoSlash = StripSlash(strFolder)
    If FSO.FolderExists(strFolderNoSlash) Then
        For Each oFile In FSO.GetFolder(strFolderNoSlash).Files
            If StrComp(FSO.GetExtensionName(oFile.Name), strExt, vbTextCompare) = 0 Then
                ' Found at least one matching file. Use the wildcard delete.
                DeleteFile FSO.BuildPath(strFolderNoSlash, "*." & strExt)
                Exit Sub
            End If
        Next
    End If
    Perf.OperationEnd

End Sub


' ----------------------------------------------------------------
' Procedure : VerifyPath (Renamed from EnsurePathExists)
' DateTime  : 8/15/2022, 10/24/2024
' Author    : Mike Wolfe, hecon5
' Source    : https://nolongerset.com/ensurepathexists/
' Purpose   : Unicode-safe method to ensure a folder exists and
'           : create the folder (and all subfolders) if it does not.
'           : Added in additional error handling and logging.
' ----------------------------------------------------------------
Public Function VerifyPath(ByRef PathToCheck As String _
                        , Optional ByVal EnableLongPath As Boolean = True) As Boolean

    Const FunctionName As String = ModuleName & ".VerifyPath"

    Const ERROR_SUCCESS As Long = &H0
    Const ERROR_ACCESS_DENIED As Long = &H5         'Could not create directory; access denied.
    Const ERROR_BAD_PATHNAME As Long = &HA1         'The pszPath parameter was set to a relative path.
    Const ERROR_FILENAME_EXCED_RANGE As Long = &HCE 'The path pointed to by pszPath is too long.
    Const ERROR_FILE_EXISTS As Long = &H50          'The directory exists.
    Const ERROR_ALREADY_EXISTS As Long = &HB7       'The directory exists.
    Const ERROR_CANCELLED As Long = &H4C7           'The user canceled the operation.
    Const ERROR_INVALID_NAME As Long = &H7B         'Unicode path passed when SHCreateDirectoryEx passes PathToCheck as string.

    Const LONG_PATH_PREFIX As String = "\\?\"

    Dim ReturnCode As Long
    Dim strFolder As String

    LogUnhandledErrors FunctionName
    On Error Resume Next

    Perf.OperationStart FunctionName

    If PathToCheck = vbNullString Then GoTo Exit_Here

    If Right$(PathToCheck, 1) = PathSep Then
        ' Folder name. (Folder names can contain periods)
        strFolder = Left$(PathToCheck, Len(PathToCheck) - 1)
    Else
        ' File name
        strFolder = FSO.GetParentFolderName(PathToCheck)
    End If

    If EnableLongPath And Not StartsWith(strFolder, ".") Then ' Can't use relative paths for LongPaths.
        ReturnCode = SHCreateDirectoryEx(ByVal 0&, StrPtr(LONG_PATH_PREFIX & strFolder), ByVal 0&)
    Else
        ReturnCode = SHCreateDirectoryEx(ByVal 0&, StrPtr(strFolder), ByVal 0&)
    End If

    Select Case ReturnCode
    Case ERROR_SUCCESS, _
         ERROR_FILE_EXISTS, _
         ERROR_ALREADY_EXISTS
        VerifyPath = True
    Case ERROR_ACCESS_DENIED: Log.Error eelError, "Could not create path: Access denied. Path: " & PathToCheck
    Case ERROR_BAD_PATHNAME: Log.Error eelError, "Cannot use relative path. Path: " & PathToCheck, FunctionName
    Case ERROR_FILENAME_EXCED_RANGE: Log.Error eelError, "Path too long. Path: " & PathToCheck, FunctionName
    Case ERROR_CANCELLED: Log.Error eelError, "User cancelled CreateDirectory operation. Path: " & PathToCheck, FunctionName
    Case ERROR_INVALID_NAME: Log.Error eelError, "Invalid path name. Path: " & PathToCheck, FunctionName
    Case Else: Log.Error eelError, "Unexpected error verifying path. Return Code: " & CStr(ReturnCode) & vbNewLine & vbNewLine & "Path: " & PathToCheck, FunctionName
    End Select

Exit_Here:
    CatchAny eelError, "Unexpected Error verifying path: " & vbNewLine & vbNewLine & PathToCheck, FunctionName
    Perf.OperationEnd

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
    If strFolder = vbNullString Then strFolder = "C:\Program Files (x86)"
    ProgramFilesFolder = strFolder & PathSep
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetFilePathsInFolder
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Returns a collection containing the full paths of files in a folder.
'           : Wildcards are supported.
'---------------------------------------------------------------------------------------
'
Public Function GetFilePathsInFolder(strFolder As String, Optional strFilePattern As String = "*.*") As Dictionary

    Dim oFile As Scripting.File
    Dim strBaseFolder As String

    strBaseFolder = StripSlash(strFolder)
    Set GetFilePathsInFolder = New Dictionary

    Perf.OperationStart "Get File List"
    If FSO.FolderExists(strBaseFolder) Then
        For Each oFile In FSO.GetFolder(strBaseFolder).Files
            ' Add files that match the pattern.
            If oFile.Name Like strFilePattern Then GetFilePathsInFolder.Add oFile.Path, vbNullString
        Next oFile
    End If
    Perf.OperationEnd

End Function


'---------------------------------------------------------------------------------------
' Procedure : GetSubfolderPaths
' Author    : Adam Waller
' Date      : 7/30/2020
' Purpose   : Return a collection of subfolders inside a folder.
'---------------------------------------------------------------------------------------
'
Public Function GetSubfolderPaths(strPath As String) As Dictionary

    Dim strBase As String
    Dim oFolder As Scripting.Folder

    Set GetSubfolderPaths = New Dictionary

    strBase = StripSlash(strPath)
    If FSO.FolderExists(strBase) Then
        For Each oFolder In FSO.GetFolder(strBase).SubFolders
            GetSubfolderPaths.Add oFolder.Path, vbNullString
        Next oFolder
    End If

End Function


'---------------------------------------------------------------------------------------
' Procedure : ReadJsonFile
' Author    : Adam Waller
' Date      : 5/5/2020
' Purpose   : Reads a Json file into a dictionary object
'---------------------------------------------------------------------------------------
'
Public Function ReadJsonFile(strPath As String) As Dictionary

    Dim strText As String
    strText = ReadFile(strPath)

    ' If it looks like json content, then parse into a dictionary object.
    If Left$(strText, 1) = "{" Then
        Set ReadJsonFile = ParseJson(strText)
    End If

End Function


'---------------------------------------------------------------------------------------
' Procedure : BuildPath2
' Author    : Adam Waller
' Date      : 3/3/2021
' Purpose   : Like FSO.BuildPath, but with unlimited arguments)
'---------------------------------------------------------------------------------------
'
Public Function BuildPath2(ParamArray Segments())
    Dim lngPart As Long
    With New clsConcat
        For lngPart = LBound(Segments) To UBound(Segments)
            .Add CStr(Segments(lngPart))
            If lngPart < UBound(Segments) Then .Add PathSep
        Next lngPart
    BuildPath2 = .GetStr
    End With
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetRelativePath
' Author    : Adam Waller
' Date      : 5/11/2020
' Purpose   : Returns a path relative to current database.
'           : If a relative path is not possible, it returns the original full path.
'---------------------------------------------------------------------------------------
'
Public Function GetRelativePath(strPath As String) As String

    Dim strFolder As String
    Dim strUncPath As String
    Dim strUncTest As String
    Dim strRelative As String

    ' Check for matching parent folder as relative to the project path.
    strFolder = GetUncPath(CurrentProject.Path) & PathSep

    ' Default to original path if no relative path could be resolved.
    strRelative = strPath

    ' Compare strPath to the current project path
    If InStr(1, strPath, strFolder, vbTextCompare) = 1 Then
        ' In export folder or subfolder. Simple replacement
        strRelative = "rel:" & Mid$(strPath, Len(strFolder) + 1)
    Else
        ' Make sure we have a path, not just a file name.
        If InStr(1, strRelative, PathSep) > 0 Then
            ' Check UNC path for network drives
            strUncPath = GetUncPath(strPath)
            If StrComp(strUncPath, strPath, vbTextCompare) <> 0 Then
                ' We are dealing with a network drive
                strUncTest = GetRelativePath(strUncPath)
                If StrComp(strUncPath, strUncTest, vbTextCompare) <> 0 Then
                    ' Resolved to relative UNC path
                    strRelative = strUncTest
                End If
            End If
        End If
    End If

    ' Return relative (or original) path
    GetRelativePath = strRelative

End Function


'---------------------------------------------------------------------------------------
' Procedure : GetPathFromRelative
' Author    : Adam Waller
' Date      : 5/11/2020
' Purpose   : Expands a relative path out to the full path.
'---------------------------------------------------------------------------------------
'
Public Function GetPathFromRelative(strPath As String) As String
    If IsRelativePath(strPath) Then
        GetPathFromRelative = FSO.BuildPath(CurrentProject.Path, Mid$(strPath, 5))
    Else
        ' No relative path used.
        GetPathFromRelative = strPath
    End If
End Function


'---------------------------------------------------------------------------------------
' Procedure : IsRelativePath
' Author    : Adam Waller
' Date      : 10/29/2021
' Purpose   : Returns true if the specified path is stored as relative.
'---------------------------------------------------------------------------------------
'
Public Function IsRelativePath(strPath As String) As Boolean
    IsRelativePath = (Left$(strPath, 4) = "rel:")
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetUncPath
' Author    : Adam Waller, hecon5
' Date      : 7/14/2020, 2022 Sept 27
' Purpose   : Returns the UNC path for a network location (if applicable)
'---------------------------------------------------------------------------------------
'
Public Function GetUNCPath(ByRef PathIn As String)

    Const FunctionName As String = ModuleName & ".GetUNCPath"
    Dim DriveLetter As String
    Dim UNCPath As String

    LogUnhandledErrors FunctionName
    On Error Resume Next

    Perf.OperationStart FunctionName
    UNCPath = PathIn

Retry:
    DriveLetter = FSO.GetDriveName(PathIn)
    If Catch(68) Then GoTo HandleDriveLoss
    CatchAny eelError, "Issue getting drive paths.", FunctionName
    With FSO.GetDrive(DriveLetter)
        If Catch(68) Then GoTo HandleDriveLoss
        If .DriveType = Remote Then
            If .IsReady Then
                UNCPath = Replace(PathIn, DriveLetter, .ShareName, , 1, vbTextCompare)
            Else
                GoTo HandleDriveLoss
            End If
        End If
    End With
    GetUNCPath = UNCPath

Exit_Here:
    Perf.OperationEnd
    CatchAny eelError, "Issue getting drive paths.", FunctionName
    Exit Function

HandleDriveLoss:
    ' This was borrowed from our applicaion, which has more error handling, so we're doing this in two steps now.
    Log.Error eelError, "Your drive isn't ready! Reconnect " & DriveLetter & " to continue." _
                        , FunctionName

    Select Case MsgBox2("Click [Retry] AFTER reconnecting drive " & DriveLetter & " to continue." _
                        , "This usually just means you need to simply open the drive in File Explorer. " _
                        , "Click Cancel to stop operation." _
                        , vbRetryCancel + vbDefaultButton1 + vbExclamation _
                        , "Drive not ready.")

        Case vbRetry
            GoTo Retry

        Case Else
            ' Log error, quit operation.
            GoTo Exit_Here
    End Select

End Function


'---------------------------------------------------------------------------------------
' Procedure : GetLastModifiedDate
' Author    : Adam Waller
' Date      : 7/30/2020
' Purpose   : Get the last modified date on a folder or file with Unicode support.
'           : Returns 0 (blank date) if the file is not found.
'---------------------------------------------------------------------------------------
'
Public Function GetLastModifiedDate(strPath As String) As Date

    Dim oFile As Scripting.File
    Dim oFolder As Scripting.Folder

    Perf.OperationStart "Get Modified Date"
    If FSO.FileExists(strPath) Then
        Set oFile = FSO.GetFile(strPath)
        GetLastModifiedDate = oFile.DateLastModified
    ElseIf FSO.FolderExists(strPath) Then
        Set oFolder = FSO.GetFolder(strPath)
        GetLastModifiedDate = oFolder.DateLastModified
    End If
    Perf.OperationEnd

End Function


'---------------------------------------------------------------------------------------
' Procedure : StripSlash
' Author    : Adam Waller
' Date      : 1/25/2019
' Purpose   : Strip the trailing slash (or other path separator)
'---------------------------------------------------------------------------------------
'
Public Function StripSlash(strText As String) As String
    If Right$(strText, 1) = PathSep Then
        StripSlash = Left$(strText, Len(strText) - 1)
    Else
        StripSlash = strText
    End If
End Function


'---------------------------------------------------------------------------------------
' Procedure : AddSlash
' Author    : Adam Waller
' Date      : 7/28/2023
' Purpose   : Ensure that the string or path ends with a slash (or path separator)
'---------------------------------------------------------------------------------------
'
Public Function AddSlash(strText As String) As String
    If Right$(strText, 1) = PathSep Then
        AddSlash = strText
    Else
        AddSlash = strText & PathSep
    End If
End Function


'---------------------------------------------------------------------------------------
' Procedure : PathSep
' Author    : Adam Waller
' Date      : 3/3/2021
' Purpose   : Return the current path separator, based on language settings.
'           : Caches value to avoid extra calls to FSO object.
'---------------------------------------------------------------------------------------
'
Public Function PathSep() As String
    Static strSeparator As String
    If strSeparator = vbNullString Then strSeparator = Mid$(FSO.BuildPath("a", "b"), 2, 1)
    PathSep = strSeparator
End Function
