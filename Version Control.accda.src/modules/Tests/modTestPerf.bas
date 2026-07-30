Attribute VB_Name = "modTestPerf"
'---------------------------------------------------------------------------------------
' Module    : modTestPerf
' Author    : Adam Waller
' Date      : 7/29/2026
' Purpose   : Micro-benchmarks for the source file scanning path used by merge builds.
'           : These are measurement tools rather than assertions, so every public entry
'           : point takes a parameter. (The test runner only discovers parameterless
'           : Public Sub procedures, so these stay out of the test tree.)
'           :
'           : Run from the Immediate Window in the add-in's own VBE:
'           :   ?modTestPerf.BenchmarkHashPrimitives()
'           :
'           : The end-to-end measurement remains the PERFORMANCE REPORTS section of an
'           : actual Export/Merge log. This module isolates the individual primitives so
'           : a regression can be attributed to one of them.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
Option Private Module
'@Folder("Tests")

Private Const ModuleName As String = "modTestPerf"

' Report layout
Private Const clngLabelWidth As Long = 46
Private Const clngLineWidth As Long = 78

' Repetition counts for the cheap primitives, so a single measurement is not
' dominated by timer granularity.
Private Const clngStringHashReps As Long = 1000
Private Const clngFileHashReps As Long = 200
Private Const clngComponentHashReps As Long = 200


'---------------------------------------------------------------------------------------
' Procedure : BenchmarkHashPrimitives
' Author    : Adam Waller
' Date      : 7/29/2026
' Purpose   : Time each primitive used by merge change detection and return a report.
'           : lngIterations applies to the expensive whole-folder/whole-category
'           : operations; the cheap per-file primitives use their own repeat counts.
'---------------------------------------------------------------------------------------
'
Public Function BenchmarkHashPrimitives(Optional lngIterations As Long = 3) As String

    Dim cOut As clsConcat
    Dim cModules As IDbComponent
    Dim dMeta As Dictionary
    Dim dModuleMeta As Dictionary
    Dim strRoot As String
    Dim strSmallFile As String
    Dim strLargeFile As String
    Dim strTestFile As String
    Dim lngIdx As Long
    Dim dblStart As Double
    Dim strJunk As String

    If lngIterations < 1 Then lngIterations = 1

    strRoot = Options.GetExportFolder
    If Not FSO.FolderExists(strRoot) Then
        BenchmarkHashPrimitives = "Export folder does not exist: " & strRoot
        Exit Function
    End If

    Set cModules = GetCategoryContainer("Modules")
    If cModules Is Nothing Then
        BenchmarkHashPrimitives = "Could not resolve the Modules container."
        Exit Function
    End If

    Set cOut = New clsConcat
    cOut.Add String$(clngLineWidth, "-"), vbCrLf
    cOut.Add "SCAN PATH BENCHMARK", vbCrLf
    cOut.Add "  Project:   ", CurrentProject.Name, vbCrLf
    cOut.Add "  Source:    ", strRoot, vbCrLf
    ' Parentheses force ByVal so these Variant results coerce to clsConcat.Add's
    ' ByRef String parameters.
    cOut.Add "  Algorithm: ", (Nz2(Options.HashAlgorithm, DefaultHashAlgorithm)), _
        (IIf(Options.UseShortHash, " (short)", vbNullString)), vbCrLf
    cOut.Add String$(clngLineWidth, "-"), vbCrLf
    cOut.Add PadRight("Operation", clngLabelWidth), PadLeft("Calls", 8), _
        PadLeft("Seconds", 10), PadLeft("ms/call", 12), vbCrLf
    cOut.Add String$(clngLineWidth, "-"), vbCrLf

    ' --- Folder metadata scans -------------------------------------------------------
    ' The recursive root scan is what a merge now performs once for the whole phase.
    dblStart = MicroSeconds
    For lngIdx = 1 To lngIterations
        Set dMeta = ScanFolderMetadata(strRoot)
    Next lngIdx
    AddResult cOut, "ScanFolderMetadata (export root, recursive)", lngIterations, dblStart
    cOut.Add PadRight("  files found", clngLabelWidth), PadLeft(CStr(dMeta.Count), 8), vbCrLf

    dblStart = MicroSeconds
    For lngIdx = 1 To lngIterations
        Set dModuleMeta = ScanFolderMetadata(cModules.BaseFolder)
    Next lngIdx
    AddResult cOut, "ScanFolderMetadata (modules folder)", lngIterations, dblStart

    ' --- Raw hashing primitives ------------------------------------------------------
    dblStart = MicroSeconds
    For lngIdx = 1 To clngStringHashReps
        strJunk = GetStringHash("benchmark content for a short string")
    Next lngIdx
    AddResult cOut, "GetStringHash (short string)", clngStringHashReps, dblStart

    ' Pick the smallest and largest source file in the modules folder
    GetSizeExtremes dModuleMeta, strSmallFile, strLargeFile

    If Len(strSmallFile) Then
        dblStart = MicroSeconds
        For lngIdx = 1 To clngFileHashReps
            strJunk = GetFileHash(strSmallFile)
        Next lngIdx
        AddResult cOut, "GetFileHash (smallest module file)", clngFileHashReps, dblStart
    End If

    If Len(strLargeFile) Then
        dblStart = MicroSeconds
        For lngIdx = 1 To clngFileHashReps
            strJunk = GetFileHash(strLargeFile)
        Next lngIdx
        AddResult cOut, "GetFileHash (largest module file)", clngFileHashReps, dblStart
    End If

    ' --- Per-component hash helpers --------------------------------------------------
    strTestFile = GetFirstSourceFile(cModules)
    If Len(strTestFile) Then

        dblStart = MicroSeconds
        For lngIdx = 1 To clngComponentHashReps
            strJunk = GetSourceFilesPropertyHash(cModules, strTestFile)
        Next lngIdx
        AddResult cOut, "GetSourceFilesPropertyHash (FSO)", clngComponentHashReps, dblStart

        dblStart = MicroSeconds
        For lngIdx = 1 To clngComponentHashReps
            strJunk = GetSourceFilesPropertyHash(cModules, strTestFile, dModuleMeta)
        Next lngIdx
        AddResult cOut, "GetSourceFilesPropertyHash (scan map)", clngComponentHashReps, dblStart

        dblStart = MicroSeconds
        For lngIdx = 1 To clngComponentHashReps
            strJunk = GetSourceFilesContentHash(cModules, strTestFile)
        Next lngIdx
        AddResult cOut, "GetSourceFilesContentHash (FSO)", clngComponentHashReps, dblStart

        dblStart = MicroSeconds
        For lngIdx = 1 To clngComponentHashReps
            strJunk = GetSourceFilesContentHash(cModules, strTestFile, dModuleMeta)
        Next lngIdx
        AddResult cOut, "GetSourceFilesContentHash (scan map)", clngComponentHashReps, dblStart

    End If

    ' --- End to end change detection per category ------------------------------------
    cOut.Add String$(clngLineWidth, "-"), vbCrLf
    AddCategoryBenchmark cOut, "Modules", dMeta, lngIterations
    AddCategoryBenchmark cOut, "Forms", dMeta, lngIterations
    AddCategoryBenchmark cOut, "Queries", dMeta, lngIterations

    cOut.Add String$(clngLineWidth, "-"), vbCrLf
    BenchmarkHashPrimitives = cOut.GetStr

End Function


'---------------------------------------------------------------------------------------
' Procedure : AddCategoryBenchmark
' Author    : Adam Waller
' Date      : 7/29/2026
' Purpose   : Time GetModifiedSourceFiles for one category.
'           :
'           : Reported in three rows, because a container caches GetFileList and
'           : GetAllFromDB on first use and a naive back-to-back comparison therefore
'           : charges the whole cache fill to whichever variant runs first:
'           :   cold        - a fresh container, one call. This is what a merge actually
'           :                 pays: the source file enumeration and the database object
'           :                 scan, plus the per-file change detection loop.
'           :   shared map  - warm, using the caller-supplied folder metadata map.
'           :   own scan    - warm, letting the category scan its own folder. The gap
'           :                 against "shared map" is the per-category folder walk that
'           :                 a merge build now avoids.
'---------------------------------------------------------------------------------------
'
Private Sub AddCategoryBenchmark(cOut As clsConcat, ByVal strCategory As String, _
    dMeta As Dictionary, ByVal lngIterations As Long)

    Dim cCategory As IDbComponent
    Dim lngIdx As Long
    Dim dblStart As Double
    Dim dResult As Dictionary

    ' Cold: a container that has never enumerated source files or database objects
    Set cCategory = GetCategoryContainer(strCategory)
    If cCategory Is Nothing Then Exit Sub

    dblStart = MicroSeconds
    Set dResult = VCSIndex.GetModifiedSourceFiles(cCategory, dMeta)
    AddResult cOut, strCategory & ": change scan (cold, shared map)", 1, dblStart

    ' Warm: the container's file list and object list are now cached, so these rows
    ' isolate the per-file detection loop and the folder metadata source.
    dblStart = MicroSeconds
    For lngIdx = 1 To lngIterations
        Set dResult = VCSIndex.GetModifiedSourceFiles(cCategory, dMeta)
    Next lngIdx
    AddResult cOut, strCategory & ": change scan (warm, shared map)", lngIterations, dblStart

    dblStart = MicroSeconds
    For lngIdx = 1 To lngIterations
        Set dResult = VCSIndex.GetModifiedSourceFiles(cCategory)
    Next lngIdx
    AddResult cOut, strCategory & ": change scan (warm, own scan)", lngIterations, dblStart

    cOut.Add PadRight("  source files / reported modified", clngLabelWidth), _
        PadLeft(CStr(cCategory.GetFileList.Count) & " / " & CStr(dResult.Count), 14), vbCrLf

End Sub


'---------------------------------------------------------------------------------------
' Procedure : AddResult
' Author    : Adam Waller
' Date      : 7/29/2026
' Purpose   : Append one measured row to the report.
'---------------------------------------------------------------------------------------
'
Private Sub AddResult(cOut As clsConcat, ByVal strLabel As String, _
    ByVal lngCalls As Long, ByVal dblStart As Double)

    Dim dblElapsed As Double

    dblElapsed = MicroSeconds - dblStart
    cOut.Add PadRight(strLabel, clngLabelWidth), _
        PadLeft(CStr(lngCalls), 8), _
        PadLeft(Format$(dblElapsed, "0.000"), 10), _
        PadLeft(Format$((dblElapsed / lngCalls) * 1000, "0.0000"), 12), vbCrLf

End Sub


'---------------------------------------------------------------------------------------
' Procedure : MicroSeconds
' Author    : Adam Waller
' Date      : 7/29/2026
' Purpose   : High resolution timer value in seconds. (Perf.MicroTimer returns Currency,
'           : which would truncate the arithmetic used to average per-call times.)
'---------------------------------------------------------------------------------------
'
Private Function MicroSeconds() As Double
    MicroSeconds = CDbl(Perf.MicroTimer)
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetCategoryContainer
' Author    : Adam Waller
' Date      : 7/29/2026
' Purpose   : Return the container class for a category name, or Nothing if this project
'           : does not include that category.
'---------------------------------------------------------------------------------------
'
Private Function GetCategoryContainer(ByVal strCategory As String) As IDbComponent

    Dim cCategory As IDbComponent

    For Each cCategory In GetContainers()
        If StrComp(cCategory.Category, strCategory, vbTextCompare) = 0 Then
            Set GetCategoryContainer = cCategory
            Exit Function
        End If
    Next cCategory

End Function


'---------------------------------------------------------------------------------------
' Procedure : GetFirstSourceFile
' Author    : Adam Waller
' Date      : 7/29/2026
' Purpose   : Return the first existing source file path for a category.
'---------------------------------------------------------------------------------------
'
Private Function GetFirstSourceFile(cCategory As IDbComponent) As String

    Dim varFile As Variant

    For Each varFile In cCategory.GetFileList
        If FSO.FileExists(CStr(varFile)) Then
            GetFirstSourceFile = CStr(varFile)
            Exit Function
        End If
    Next varFile

End Function


'---------------------------------------------------------------------------------------
' Procedure : GetSizeExtremes
' Author    : Adam Waller
' Date      : 7/29/2026
' Purpose   : Return the smallest and largest file in a ScanFolderMetadata map, so the
'           : file hashing measurement covers both ends of the size range.
'---------------------------------------------------------------------------------------
'
Private Sub GetSizeExtremes(dMeta As Dictionary, ByRef strSmallest As String, ByRef strLargest As String)

    Dim varKey As Variant
    Dim varMeta As Variant
    Dim dblSize As Double
    Dim dblMin As Double
    Dim dblMax As Double

    If dMeta Is Nothing Then Exit Sub

    dblMin = -1
    For Each varKey In dMeta.Keys
        varMeta = dMeta(varKey)
        dblSize = CDbl(varMeta(1))
        If dblMin < 0 Or dblSize < dblMin Then
            dblMin = dblSize
            strSmallest = CStr(varKey)
        End If
        If dblSize > dblMax Then
            dblMax = dblSize
            strLargest = CStr(varKey)
        End If
    Next varKey

End Sub


'---------------------------------------------------------------------------------------
' Procedure : PadRight
' Author    : Adam Waller
' Date      : 7/29/2026
' Purpose   : Left justify text in a fixed width column.
'---------------------------------------------------------------------------------------
'
Private Function PadRight(ByVal strText As String, ByVal lngLen As Long) As String

    Dim strResult As String
    Dim strTrimmed As String

    strResult = Space$(lngLen)
    strTrimmed = Left$(strText, lngLen - 1)
    Mid$(strResult, 1, Len(strTrimmed)) = strTrimmed
    PadRight = strResult

End Function


'---------------------------------------------------------------------------------------
' Procedure : PadLeft
' Author    : Adam Waller
' Date      : 7/29/2026
' Purpose   : Right justify text in a fixed width column.
'---------------------------------------------------------------------------------------
'
Private Function PadLeft(ByVal strText As String, ByVal lngLen As Long) As String

    Dim strResult As String
    Dim strTrimmed As String

    strResult = Space$(lngLen)
    strTrimmed = Left$(strText, lngLen - 1)
    Mid$(strResult, lngLen - Len(strTrimmed) + 1, Len(strTrimmed)) = strTrimmed
    PadLeft = strResult

End Function
