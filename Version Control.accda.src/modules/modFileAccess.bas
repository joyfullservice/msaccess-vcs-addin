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


Private Declare PtrSafe Function getTempPath Lib "kernel32" Alias "GetTempPathA" ( _
    ByVal nBufferLength As Long, _
    ByVal lpBuffer As String) As Long
    
Private Declare PtrSafe Function getTempFileName Lib "kernel32" Alias "GetTempFileNameA" ( _
    ByVal lpszPath As String, _
    ByVal lpPrefixString As String, _
    ByVal wUnique As Long, _
    ByVal lpTempFileName As String) As Long
    
    
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

    Dim strContent As String
    Dim dblPos As Double
    
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
        .SaveToFile strPath, adSaveCreateOverWrite
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
' Procedure : MkDirIfNotExist
' Author    : Adam Waller
' Date      : 1/25/2019
' Purpose   : Create folder `Path`. Silently do nothing if it already exists.
'---------------------------------------------------------------------------------------
'
Public Sub MkDirIfNotExist(strPath As String)
    If Not FSO.FolderExists(StripSlash(strPath)) Then
        Perf.OperationStart "Create Folder"
        FSO.CreateFolder StripSlash(strPath)
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
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : VerifyPath
' Author    : Adam Waller
' Date      : 8/3/2020
' Purpose   : Verifies that the folder path to a folder or file exists.
'           : Use this to verify the folder path before attempting to write a file.
'---------------------------------------------------------------------------------------
'
Public Sub VerifyPath(strPath As String)
    
    Dim strFolder As String
    Dim varParts As Variant
    Dim intPart As Integer
    Dim strVerified As String
    
    If strPath = vbNullString Then Exit Sub
    
    Perf.OperationStart "Verify Path"
    
    ' Determine if the path is a file or folder
    If Right$(strPath, 1) = PathSep Then
        ' Folder name. (Folder names can contain periods)
        strFolder = Left$(strPath, Len(strPath) - 1)
    Else
        ' File name
        strFolder = FSO.GetParentFolderName(strPath)
    End If
    
    ' Check if full path exists.
    If Not FSO.FolderExists(strFolder) Then
        ' Start from the root, and build out full path, creating folders as needed.
        ' UNC path? change 3 "\" into 3 "@"
        If strFolder Like PathSep & PathSep & "*" & PathSep & "*" Then
            strFolder = Replace(strFolder, PathSep, "@", 1, 3)
        End If

        ' Separate folders from server name
        varParts = Split(strFolder, PathSep)
        ' Get the slashes back
        varParts(0) = Replace(varParts(0), "@", PathSep, 1, 3)
                
        ' Make sure the root folder exists. If it doesn't we probably have some other issue.
        If Not FSO.FolderExists(varParts(0)) Then
            MsgBox2 "Path Not Found", "Could not find the path '" & varParts(0) & "' on this system.", _
                    "While trying to verify this path: " & strFolder, vbExclamation
        Else
            ' Loop through folder structure, creating as needed.
            strVerified = varParts(0) & PathSep
            For intPart = 1 To UBound(varParts)
                strVerified = FSO.BuildPath(strVerified, varParts(intPart))
                MkDirIfNotExist strVerified

            Next intPart
        End If
    End If
    
    ' End timing of operation
    Perf.OperationEnd
    
End Sub


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
    
    If FSO.FolderExists(strBaseFolder) Then
        For Each oFile In FSO.GetFolder(strBaseFolder).Files
            ' Add files that match the pattern.
            If oFile.Name Like strFilePattern Then GetFilePathsInFolder.Add oFile.Path, vbNullString
        Next oFile
    End If
    
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
    If Left$(strPath, 4) = "rel:" Then
        GetPathFromRelative = FSO.BuildPath(CurrentProject.Path, Mid$(strPath, 5))
    Else
        ' No relative path used.
        GetPathFromRelative = strPath
    End If
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetUncPath
' Author    : Adam Waller
' Date      : 7/14/2020
' Purpose   : Returns the UNC path for a network location (if applicable)
'---------------------------------------------------------------------------------------
'
Public Function GetUncPath(strPath As String)

    Dim strDrive As String
    Dim strUNC As String
    
    strUNC = strPath
    strDrive = FSO.GetDriveName(strPath)
    With FSO.GetDrive(strDrive)
        If .DriveType = Remote Then
            strUNC = Replace(strPath, strDrive, .ShareName, , 1, vbTextCompare)
        End If
    End With
    GetUncPath = strUNC
    
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetLastModifiedDate
' Author    : Adam Waller
' Date      : 7/30/2020
' Purpose   : Get the last modified date on a folder or file with Unicode support.
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
' Purpose   : Strip the trailing slash
'---------------------------------------------------------------------------------------
'
Public Function StripSlash(strText As String) As String
    If Right$(strText, 1) = PathSep Then
        StripSlash = Left$(strText, Len(strText) - 1)
    Else
        StripSlash = strText
    End If
End Function

