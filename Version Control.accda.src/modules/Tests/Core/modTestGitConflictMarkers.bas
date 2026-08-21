Attribute VB_Name = "modTestGitConflictMarkers"
'---------------------------------------------------------------------------------------
' Module    : modTestGitConflictMarkers
' Author    : Adam Waller
' Date      : 8/20/2026
' Purpose   : Tests for detecting unresolved Git conflict markers in source files
'           : before import. Marker text is built at runtime so this module never
'           : contains column-0 markers in its own source.
'           : Run: ?VCS.RunTests("modTestGitConflictMarkers")
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
Option Private Module
'@Folder("Tests.Core")
'@Tag("unit")


Private Const ModuleName As String = "modTestGitConflictMarkers"


'---------------------------------------------------------------------------------------
' Procedure : TestGitConflictMarkerLine_CleanContent
' Author    : Adam Waller
' Date      : 8/20/2026
' Purpose   : Clean text has no conflict markers.
'---------------------------------------------------------------------------------------
'
Public Sub TestGitConflictMarkerLine_CleanContent()

    TestAssert GitConflictMarkerLine("Option Explicit" & vbCrLf & "Public Sub Foo()") = 0, _
        "clean content returns 0"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestGitConflictMarkerLine_FirstLine
' Author    : Adam Waller
' Date      : 8/20/2026
' Purpose   : A marker at the start of the file is on line 1.
'---------------------------------------------------------------------------------------
'
Public Sub TestGitConflictMarkerLine_FirstLine()

    TestAssert GitConflictMarkerLine(MarkerOpenLine() & vbCrLf & "content") = 1, _
        "marker on first line returns 1"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestGitConflictMarkerLine_MidFile
' Author    : Adam Waller
' Date      : 8/20/2026
' Purpose   : Line numbering matches a marker deep in a form-like source file.
'---------------------------------------------------------------------------------------
'
Public Sub TestGitConflictMarkerLine_MidFile()

    Dim strContent As String
    Dim lngLine As Long
    Dim lngExpected As Long

    strContent = RepeatLine("Begin Label", 94) & MarkerOpenLine()
    lngExpected = 95
    lngLine = GitConflictMarkerLine(strContent)
    TestAssert lngLine = lngExpected, "marker on line " & lngExpected & ", got " & lngLine

End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestGitConflictMarkerLine_IndentedAndCommentMarkersIgnored
' Author    : Adam Waller
' Date      : 8/20/2026
' Purpose   : Markers not at column 0, including inside comments, are ignored.
'---------------------------------------------------------------------------------------
'
Public Sub TestGitConflictMarkerLine_IndentedAndCommentMarkersIgnored()

    TestAssert GitConflictMarkerLine(Space$(4) & MarkerOpenLine()) = 0, _
        "indented marker returns 0"
    TestAssert GitConflictMarkerLine("' " & MarkerOpenLine()) = 0, _
        "comment marker returns 0"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestGitConflictMarkerLine_EmbeddedThenRealMarker
' Author    : Adam Waller
' Date      : 8/20/2026
' Purpose   : An embedded marker does not stop a real column-0 marker later in the file.
'---------------------------------------------------------------------------------------
'
Public Sub TestGitConflictMarkerLine_EmbeddedThenRealMarker()

    Dim strContent As String

    strContent = "Caption =" & Quote(MarkerOpenLine()) & vbCrLf & MarkerCloseLine()
    TestAssert GitConflictMarkerLine(strContent) = 2, "real closing marker on line 2"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestGitConflictMarkerLine_ClosingMarkerOnly
' Author    : Adam Waller
' Date      : 8/20/2026
' Purpose   : A partially resolved file with only the closing marker is caught.
'---------------------------------------------------------------------------------------
'
Public Sub TestGitConflictMarkerLine_ClosingMarkerOnly()

    TestAssert GitConflictMarkerLine("resolved content" & vbCrLf & MarkerCloseLine()) = 2, _
        "closing marker only returns line 2"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestGitConflictMarkerLine_EqualsSeparatorIgnored
' Author    : Adam Waller
' Date      : 8/20/2026
' Purpose   : A row of equals signs is not treated as a conflict marker.
'---------------------------------------------------------------------------------------
'
Public Sub TestGitConflictMarkerLine_EqualsSeparatorIgnored()

    TestAssert GitConflictMarkerLine(String(80, 61)) = 0, "equals separator returns 0"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestFileHasGitConflictMarkers_LogsAndReturnsTrue
' Author    : Adam Waller
' Date      : 8/20/2026
' Purpose   : File-level scan logs an error and refuses marker-corrupted files.
'---------------------------------------------------------------------------------------
'
Public Sub TestFileHasGitConflictMarkers_LogsAndReturnsTrue()

    Dim strFile As String
    Dim lngErrBefore As Long

    strFile = WriteTempTextFile(MarkerOpenLine() & vbCrLf & "{""Items"":{}}")
    lngErrBefore = Log.ErrorCount

    TestAssert FileHasGitConflictMarkers(strFile, ModuleName & ".TestFileHasGitConflictMarkers_LogsAndReturnsTrue"), _
        "marker file returns True"
    TestAssert Log.ErrorCount > lngErrBefore, "marker file logged an error"

    DeleteFile strFile

End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestGitConflictMarkerLineInFile_PastFirstChunk
' Author    : Adam Waller
' Date      : 8/20/2026
' Purpose   : Chunked scanning finds a marker beyond the first read chunk.
'---------------------------------------------------------------------------------------
'
Public Sub TestGitConflictMarkerLineInFile_PastFirstChunk()

    Dim strFile As String
    Dim strPad As String
    Dim lngLine As Long

    strPad = String(CHUNK_SIZE + 1000, 65) & vbCrLf
    strFile = WriteTempTextFile(strPad & MarkerOpenLine())
    lngLine = GitConflictMarkerLineInFile(strFile)

    TestAssert lngLine = 2, "marker past first chunk is on line 2, got " & lngLine

    DeleteFile strFile

End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestGitConflictMarkerLineInFile_StraddlesChunkBoundary
' Author    : Adam Waller
' Date      : 8/20/2026
' Purpose   : Overlap handling finds a marker split across chunk boundaries.
'---------------------------------------------------------------------------------------
'
Public Sub TestGitConflictMarkerLineInFile_StraddlesChunkBoundary()

    Dim strFile As String
    Dim strPad As String
    Dim lngLine As Long

    ' Place the column-0 marker exactly at the start of the second chunk.
    strPad = String(CHUNK_SIZE - 2, 65) & vbCrLf
    strFile = WriteTempTextFile(strPad & MarkerOpenLine())
    lngLine = GitConflictMarkerLineInFile(strFile)

    TestAssert lngLine = 2, "straddling marker is on line 2, got " & lngLine

    DeleteFile strFile

End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestReadJsonFile_RejectsConflictMarkers
' Author    : Adam Waller
' Date      : 8/20/2026
' Purpose   : ReadJsonFile refuses JSON with unresolved Git conflict markers.
'---------------------------------------------------------------------------------------
'
Public Sub TestReadJsonFile_RejectsConflictMarkers()
    '@Tag("integration")

    Dim dFile As Dictionary
    Dim strFile As String
    Dim lngErrBefore As Long

    strFile = WriteTempTextFile(MarkerOpenLine() & vbCrLf & "{""Items"":{}}")
    lngErrBefore = Log.ErrorCount
    Set dFile = ReadJsonFile(strFile)

    TestAssert dFile Is Nothing, "ReadJsonFile returns Nothing for marker file"
    TestAssert Log.ErrorCount > lngErrBefore, "ReadJsonFile logged an error"

    DeleteFile strFile

End Sub


Private Function MarkerOpenLine() As String

    MarkerOpenLine = RepeatChar(60, 7) & " Updated upstream"

End Function


Private Function MarkerCloseLine() As String

    MarkerCloseLine = RepeatChar(62, 7) & " Stashed changes"

End Function


Private Function RepeatChar(ByVal intChar As Integer, ByVal lngCount As Long) As String

    Dim lngIndex As Long
    Dim strResult As String

    For lngIndex = 1 To lngCount
        strResult = strResult & Chr$(intChar)
    Next lngIndex
    RepeatChar = strResult

End Function


Private Function RepeatLine(strLine As String, lngCount As Long) As String

    Dim lngIndex As Long
    Dim cData As clsConcat

    Set cData = New clsConcat
    cData.AppendOnAdd = vbCrLf
    For lngIndex = 1 To lngCount
        cData.Add strLine
    Next lngIndex
    RepeatLine = cData.GetStr

End Function


Private Function Quote(strText As String) As String

    Quote = """" & strText & """"

End Function


Private Function WriteTempTextFile(strContent As String) As String

    WriteTempTextFile = GetTempFolder("vcs_git_marker") & PathSep & "marker.txt"
    WriteFile strContent, WriteTempTextFile

End Function
