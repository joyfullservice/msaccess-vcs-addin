VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPerformance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : clsPerformance
' Author    : Adam Waller
' Date      : 12/4/2020
' Purpose   : Measure the performance of the export/import process. Since different
'           : users have different needs and work with sometimes very different
'           : databases, this tool will help identify potential bottlenecks in the
'           : performance of the add-in in real-life scenarios. The results are
'           : typically added to the log files.
'           : Note: This class has been updated to use API calls for timing to the
'           : microsecond level. For additional details, see the following link:
'           : http://www.mendipdatasystems.co.uk/timer-comparison-tests/4594552971
'---------------------------------------------------------------------------------------

Option Compare Database
Option Explicit
'@Folder("Infrastructure")

' Enable or disable timing operations
Public Enabled As Boolean

' Record time against the full operation call path, not just the operation name.
' Off by default: it costs a path string per start/stop on a hot path, and nothing in
' the human-readable report consumes it. Turned on by Options.ExportPerfJson, which is
' how an agent asks for the machine-readable breakdown before running an operation.
Public TrackCallPaths As Boolean

' Use type for private class variables
Private Type udtPerformance
    Overall As clsPerformanceItem
    CategoryName As String
    Categories As Dictionary
    OperationName As String
    Operations As Dictionary
    Frequency As Currency
    intDigitsAfterDecimal As Integer
    ' Exclusive time keyed by call path (segments joined by PathSep), so the same
    ' operation is measured separately for each place it is called from.
    Paths As Dictionary
    PathSep As String
    ' Manage a type of call stack to track nested operations.
    ' When an operation finishes, it goes back to timing the
    ' previous operation.
    CallStackItems As Collection
    ' Generic footnote annotations for categories.
    FootnoteMarks As Dictionary     ' mark char -> description
    CategoryFootnotes As Dictionary ' category name -> accumulated marks string
End Type
Private this As udtPerformance

' API calls to get more precise time than Timer function
Private Declare PtrSafe Function GetFrequencyAPI Lib "kernel32" Alias "QueryPerformanceFrequency" (ByRef Frequency As Currency) As Long
Private Declare PtrSafe Function GetTimeAPI Lib "kernel32" Alias "QueryPerformanceCounter" (ByRef Counter As Currency) As Long


'---------------------------------------------------------------------------------------
' Procedure : StartTiming
' Author    : Adam Waller
' Date      : 11/3/2020
' Purpose   : Start the overall timing.
'---------------------------------------------------------------------------------------
'
Public Sub StartTiming()
    If Not Me.Enabled Then Exit Sub
    Me.Reset
    ' Guarded like DebugMode: Options may not be loaded this early, and asking for it
    ' here would recurse through the option loader's own timing calls.
    If OptionsLoaded Then Me.TrackCallPaths = Options.ExportPerfJson
    Set this.Overall = New clsPerformanceItem
    this.Overall.Start = MicroTimer
End Sub


'---------------------------------------------------------------------------------------
' Procedure : CallStack
' Author    : hecon5
' Date      : 01/26/2022; 04/28/2022
' Purpose   : Return the call stack in operation.
'           : Adds major element "Component"(now "Category") if in use, and the trace of all minor
'           : elements (Operations) affixed with a period. This facilitates use of
'           : performance and because the VBA error may be returned before you expect
'           : (or after...VBA can be fun!).
'---------------------------------------------------------------------------------------
'
Public Property Get CallStack() As String

    Dim lngCallStackPosition As Long

    If Not Me.Enabled Then Exit Property

    ' Build out breadcrumb type call stack string
    With New clsConcat
        .AppendOnAdd = "."
        If this.CategoryName <> vbNullString Then .Add this.CategoryName
        If Not this.CallStackItems Is Nothing Then
            For lngCallStackPosition = 1 To this.CallStackItems.Count
                .Add CStr(this.CallStackItems(lngCallStackPosition))
            Next lngCallStackPosition
        End If
        If this.OperationName <> vbNullString Then .Add this.OperationName
        If .Length > 0 Then .Remove 1 ' Remove trailing delimiter only if there's something to remove.
        CallStack = .GetStr
    End With

End Property


'---------------------------------------------------------------------------------------
' Procedure : CurrentCategoryName
' Author    : hecon5
' Date      : 10/3/2023
' Purpose   : Return the current category name.
'---------------------------------------------------------------------------------------
'
Public Property Get CurrentCategoryName() As String
    CurrentCategoryName = this.CategoryName
End Property


'---------------------------------------------------------------------------------------
' Procedure : AddCategoryNote
' Author    : Adam Waller
' Date      : 3/12/2026
' Purpose   : Add a footnote annotation to a category. The mark character (default "*")
'           : is appended to the category name in the performance report, and the note
'           : text is rendered as a footnote below the category table.
'---------------------------------------------------------------------------------------
'
Public Sub AddCategoryNote(strCategory As String, strNote As String, Optional strMark As String = "*")

    If Len(strCategory) = 0 Then Exit Sub
    If Not Me.Enabled Then Exit Sub

    ' Initialize dictionaries on first use
    If this.FootnoteMarks Is Nothing Then Set this.FootnoteMarks = New Dictionary
    If this.CategoryFootnotes Is Nothing Then Set this.CategoryFootnotes = New Dictionary

    ' Register the mark -> description (first registration wins)
    If Not this.FootnoteMarks.Exists(strMark) Then
        this.FootnoteMarks.Add strMark, strNote
    End If

    ' Append mark to the category's footnote string
    If this.CategoryFootnotes.Exists(strCategory) Then
        ' Only append if this mark isn't already on this category
        If InStr(1, this.CategoryFootnotes(strCategory), strMark) = 0 Then
            this.CategoryFootnotes(strCategory) = this.CategoryFootnotes(strCategory) & strMark
        End If
    Else
        this.CategoryFootnotes.Add strCategory, strMark
    End If

End Sub


'---------------------------------------------------------------------------------------
' Procedure : CategoryStart
' Author    : Adam Waller
' Date      : 11/3/2020
' Purpose   : Start timing a component type.
'---------------------------------------------------------------------------------------
'
Public Sub CategoryStart(strName As String)
    If Not Me.Enabled Then Exit Sub
    If this.Categories Is Nothing Then Set this.Categories = New Dictionary
    If this.CategoryName <> vbNullString Then CategoryEnd
    StartTimer this.Categories, strName
    this.CategoryName = strName
End Sub


'---------------------------------------------------------------------------------------
' Procedure : CategoryEnd
' Author    : Adam Waller
' Date      : 11/3/2020
' Purpose   : End the timing of the active component
'---------------------------------------------------------------------------------------
'
Public Sub CategoryEnd(Optional lngCount As Long = 1)
    If Not Me.Enabled Then Exit Sub
    If this.CategoryName <> vbNullString And Not this.Categories Is Nothing Then
        LapTimer this.Categories(this.CategoryName), lngCount
        this.CategoryName = vbNullString
    End If
End Sub


'---------------------------------------------------------------------------------------
' Procedure : OperationStart
' Author    : Adam Waller
' Date      : 11/3/2020
' Purpose   : Start timing a named operation. (i.e. Sanitize Files)
'           : Note: This does a type of "call stack" function, where nested operations
'           : are recorded exclusive of the parent operations.
'---------------------------------------------------------------------------------------
'
Public Sub OperationStart(strName As String)

    If Not Me.Enabled Then Exit Sub

    ' Ensure that we have created the operations dictionary
    If this.Operations Is Nothing Then Set this.Operations = New Dictionary

    ' See if we are already timing something
    If this.OperationName <> vbNullString Then

        ' We are already timing something else right now.
        ' Save the current process to the call stack before switching
        ' to the new operation.
        LapTimer this.Operations(this.OperationName), 0
        If Me.TrackCallPaths Then LapPath CurrentPath, 0
        If this.CallStackItems Is Nothing Then Set this.CallStackItems = New Collection
        With this.CallStackItems
            ' Safety check!
            If .Count < 100 Then .Add this.OperationName
        End With
    End If

    ' Start the timer for this operation.
    StartTimer this.Operations, strName
    this.OperationName = strName

    ' Path timer starts after OperationName is set, since the path includes it.
    If Me.TrackCallPaths Then
        If this.Paths Is Nothing Then Set this.Paths = New Dictionary
        StartTimer this.Paths, CurrentPath
    End If

End Sub


'---------------------------------------------------------------------------------------
' Procedure : OperationEnd
' Author    : Adam Waller
' Date      : 11/3/2020
' Purpose   : Stop the timing of the active operation.
'---------------------------------------------------------------------------------------
'
Public Sub OperationEnd(Optional lngCount As Long = 1)

    Dim strLastOperation As String

    If Not Me.Enabled Then Exit Sub

    ' Verify that we are timing something, and record the elapsed time.
    If this.OperationName <> vbNullString And Not this.Operations Is Nothing Then

        ' Record the elapsed time.
        LapTimer this.Operations(this.OperationName), lngCount
        If Me.TrackCallPaths Then LapPath CurrentPath, lngCount

        ' Check the call stack to see if we need to move back to the previous process.
        If Not this.CallStackItems Is Nothing Then
            With this.CallStackItems
                If .Count > 0 Then
                    ' Resume previous activity. The parent comes off the stack before it
                    ' is restarted, because OperationStart derives the call path from the
                    ' stack plus the active name -- leave it on and its own segment
                    ' repeats, recording the parent as a child of itself.
                    strLastOperation = .Item(.Count)
                    .Remove .Count
                    this.OperationName = vbNullString
                    OperationStart strLastOperation
                Else
                    ' No longer timing any operations.
                    this.OperationName = vbNullString
                End If
            End With
        Else
            this.OperationName = vbNullString
        End If
    End If

End Sub


'---------------------------------------------------------------------------------------
' Procedure : CurrentPath
' Author    : Adam Waller
' Date      : 8/27/2026
' Purpose   : The active operation's full call path, segments joined by PathSep. Built
'           : with plain concatenation rather than clsConcat because this runs on every
'           : OperationStart and OperationEnd when path tracking is on.
'---------------------------------------------------------------------------------------
'
Private Function CurrentPath() As String

    Dim lngItem As Long
    Dim strPath As String

    If Not this.CallStackItems Is Nothing Then
        For lngItem = 1 To this.CallStackItems.Count
            If Len(strPath) > 0 Then strPath = strPath & this.PathSep
            strPath = strPath & CStr(this.CallStackItems(lngItem))
        Next lngItem
    End If

    If this.OperationName <> vbNullString Then
        If Len(strPath) > 0 Then strPath = strPath & this.PathSep
        strPath = strPath & this.OperationName
    End If

    CurrentPath = strPath

End Function


'---------------------------------------------------------------------------------------
' Procedure : LapPath
' Author    : Adam Waller
' Date      : 8/27/2026
' Purpose   : Add elapsed time to a call path timer. Missing keys are ignored rather
'           : than created, so toggling TrackCallPaths mid-operation cannot fault on a
'           : path whose start was never recorded.
'---------------------------------------------------------------------------------------
'
Private Sub LapPath(strPath As String, lngCount As Long)
    If this.Paths Is Nothing Then Exit Sub
    If Not this.Paths.Exists(strPath) Then Exit Sub
    LapTimer this.Paths(strPath), lngCount
End Sub


'---------------------------------------------------------------------------------------
' Procedure : CurrentOperationName
' Author    : hecon5
' Date      : 10/3/2023
' Purpose   : Return the current operation's name.
'---------------------------------------------------------------------------------------
'
Public Property Get CurrentOperationName() As String
    CurrentOperationName = this.OperationName
End Property

'---------------------------------------------------------------------------------------
' Procedure : DigitsAfterDecimal
' Author    : Eugen Albiker
' Date      : 16/3/2021
' Purpose   : Set Number of Digits after Decimal for the Elapsed Time
'---------------------------------------------------------------------------------------
'
Public Property Get DigitsAfterDecimal() As Integer
    DigitsAfterDecimal = this.intDigitsAfterDecimal
End Property

Public Property Let DigitsAfterDecimal(intDigitsAfterDecimal As Integer)
    If intDigitsAfterDecimal > 4 Then intDigitsAfterDecimal = 4
    this.intDigitsAfterDecimal = intDigitsAfterDecimal
End Property


'---------------------------------------------------------------------------------------
' Procedure : EndTiming
' Author    : Adam Waller
' Date      : 11/3/2020
' Purpose   : End the overall timing, adding to total. (Allows you to start and stop
'           : during the instance of the class.)
'---------------------------------------------------------------------------------------
'
Public Sub EndTiming()
    If Not Me.Enabled Then Exit Sub
    LapTimer this.Overall, 1
End Sub


'---------------------------------------------------------------------------------------
' Procedure : MicroTimer
' Author    : Adam Waller
' Date      : 12/4/2020
' Purpose   : Return time in seconds with microsecond precision
'---------------------------------------------------------------------------------------
'
Public Function MicroTimer() As Currency

    Dim curTime As Currency

    ' Call API to get current time
    GetTimeAPI curTime

    ' Convert to seconds
    MicroTimer = (curTime / this.Frequency)

End Function


'---------------------------------------------------------------------------------------
' Procedure : StartTimer
' Author    : Adam Waller
' Date      : 11/3/2020
' Purpose   : Add the item if it doesn't exist, then set the start time.
'---------------------------------------------------------------------------------------
'
Private Sub StartTimer(dItems As Dictionary, strName As String)

    Dim cItem As clsPerformanceItem

    If Not dItems.Exists(strName) Then
        Set cItem = New clsPerformanceItem
        dItems.Add strName, cItem
    End If
    dItems(strName).Start = MicroTimer
End Sub


'---------------------------------------------------------------------------------------
' Procedure : LapTimer
' Author    : Adam Waller
' Date      : 11/3/2020
' Purpose   : Adds the elapsed time to the timer.
'---------------------------------------------------------------------------------------
'
Private Sub LapTimer(cItem As clsPerformanceItem, lngCount As Long)
    If Not cItem Is Nothing Then
        With cItem
            If .Start > 0 Then
                .Total = .Total + GetElapsed(.Start)
                .Start = 0
                .Count = .Count + lngCount
            End If
        End With
    End If
End Sub


'---------------------------------------------------------------------------------------
' Procedure : PauseTiming
' Author    : Adam Waller
' Date      : 10/20/2021
' Purpose   : Lap all timers to capture elapsed time.
'---------------------------------------------------------------------------------------
'
Public Sub PauseTiming()

    If Not Me.Enabled Then Exit Sub

    ' Lap overall time
    LapTimer this.Overall, 0

    ' Lap any active timing operation without incrementing
    If this.OperationName <> vbNullString And Not this.Operations Is Nothing Then
        LapTimer this.Operations(this.OperationName), 0
    End If

    If this.CategoryName <> vbNullString And Not this.Categories Is Nothing Then
        LapTimer this.Categories(this.CategoryName), 0
    End If
End Sub


'---------------------------------------------------------------------------------------
' Procedure : ResumeTiming
' Author    : Adam Waller
' Date      : 10/20/2021
' Purpose   : Resume the timing of all timers by resetting the start time.
'---------------------------------------------------------------------------------------
'
Public Sub ResumeTiming()

    If Not Me.Enabled Then Exit Sub

    ' Resume overall time
    If Not this.Overall Is Nothing Then this.Overall.Start = MicroTimer

    ' Resume current operation
    If this.OperationName <> vbNullString Then StartTimer this.Operations, this.OperationName

    ' Resume current Category
    If this.CategoryName <> vbNullString Then StartTimer this.Categories, this.CategoryName

End Sub


'---------------------------------------------------------------------------------------
' Procedure : GetElapsed
' Author    : Adam Waller
' Date      : 11/3/2020
' Purpose   : Add current timer to sngStart to get elapsed seconds.
'---------------------------------------------------------------------------------------
'
Private Function GetElapsed(curStart As Currency) As Currency

    Dim curNow As Currency

    ' Only return a value if we have a starting time.
    If curStart > 0 Then
        curNow = MicroTimer
        If curStart <= curNow Then
            GetElapsed = curNow - curStart
        Else
            ' Just in case someone was up really late, and crossed midnight...
            GetElapsed = curStart + ((24# * 60 * 60) - curStart)
        End If
    End If

End Function


'---------------------------------------------------------------------------------------
' Procedure : TotalTime
' Author    : Adam Waller
' Date      : 4/29/2021
' Purpose   : Return total time from start to end, or to now if end not specified.
'---------------------------------------------------------------------------------------
'
Public Property Get TotalTime() As Currency
    If Not Me.Enabled Then Exit Function
    If this.Overall Is Nothing Then
        TotalTime = 0
    Else
        LapTimer this.Overall, 0
        TotalTime = this.Overall.Total
    End If
End Property


'---------------------------------------------------------------------------------------
' Procedure : GetCategoryCounts
' Author    : Adam Waller
' Date      : 1/26/2026
' Purpose   : Return a dictionary of category names and their counts.
'           : Used by MCP callbacks to report actual export/import counts.
'---------------------------------------------------------------------------------------
'
Public Function GetCategoryCounts() As Dictionary
    Dim dCounts As Dictionary
    Dim varKey As Variant

    Set dCounts = New Dictionary

    If Not this.Categories Is Nothing Then
        For Each varKey In this.Categories.Keys
            dCounts.Add CStr(varKey), this.Categories(varKey).Count
        Next varKey
    End If

    Set GetCategoryCounts = dCounts
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetDiagnosticData
' Author    : Adam Waller
' Date      : 8/27/2026
' Purpose   : Return the same measurements the text report shows, plus the call-path
'           : breakdown, in a structure meant to be serialized to JSON and read by a
'           : tool rather than a person.
'           : The text report has to pick one view; this returns several, because the
'           : useful question changes. "Which operation costs the most in total, from
'           : wherever it is called" is answered by `operations`. "What does this feature
'           : cost including everything beneath it" is answered by `callPaths`, whose
'           : inclusive seconds roll up descendants. "Who is driving this shared
'           : operation's count" is answered by `callers`.
'           : `callPaths` and `callers` are empty unless TrackCallPaths was on for the
'           : run. Exclusive seconds still sum to the accounted total in every view, so
'           : inclusive figures overlap by design and must not be added together.
'---------------------------------------------------------------------------------------
'
Public Function GetDiagnosticData() As Dictionary

    Dim dData As Dictionary
    Dim dTotals As Dictionary
    Dim dEntry As Dictionary
    Dim dCaller As Dictionary
    Dim colRows As Collection
    Dim colInner As Collection
    Dim dByName As Dictionary
    Dim dNameTotals As Dictionary
    Dim dParents As Dictionary
    Dim dSorted As Dictionary
    Dim dSortedParents As Dictionary
    Dim cItem As clsPerformanceItem
    Dim varKey As Variant
    Dim varInner As Variant
    Dim curAccounted As Currency
    Dim strName As String
    Dim strCaller As String

    Set dData = New Dictionary
    dData.Add "schemaVersion", 1
    dData.Add "generated", Format$(Now, "yyyy-mm-dd hh:nn:ss")
    dData.Add "trackedCallPaths", Me.TrackCallPaths
    dData.Add "stream", Log.GetStreamDiagnosticData
    dData.Add "mcp", MCP.GetDiagnosticData

    ' Categories (component types)
    Set colRows = New Collection
    If Not this.Categories Is Nothing Then
        Set dSorted = SortItemsByTime(this.Categories)
        For Each varKey In dSorted.Keys
            Set dEntry = New Dictionary
            dEntry.Add "name", CStr(varKey)
            dEntry.Add "count", dSorted(varKey).Count
            dEntry.Add "seconds", CDbl(dSorted(varKey).Total)
            colRows.Add dEntry
        Next varKey
    End If
    dData.Add "categories", colRows

    ' Operations, exclusive of nested operations and aggregated by name. This is the
    ' view that answers "many small calls adding up", so it counts every call site.
    Set colRows = New Collection
    If Not this.Operations Is Nothing Then
        Set dSorted = SortItemsByTime(this.Operations)
        For Each varKey In dSorted.Keys
            Set dEntry = New Dictionary
            dEntry.Add "name", CStr(varKey)
            dEntry.Add "count", dSorted(varKey).Count
            dEntry.Add "seconds", CDbl(dSorted(varKey).Total)
            colRows.Add dEntry
            curAccounted = curAccounted + dSorted(varKey).Total
        Next varKey
    End If
    dData.Add "operations", colRows

    ' Totals. "other" is the unattributed remainder the text report shows.
    Set dTotals = New Dictionary
    If Not this.Overall Is Nothing Then
        dTotals.Add "runtimeSeconds", CDbl(this.Overall.Total)
        dTotals.Add "otherSeconds", CDbl(this.Overall.Total - curAccounted)
    End If
    dTotals.Add "accountedSeconds", CDbl(curAccounted)
    dData.Add "totals", dTotals

    ' Call paths, with inclusive rollups
    Set colRows = New Collection
    If Not this.Paths Is Nothing Then
        Set dSorted = SortItemsByTime(this.Paths)
        For Each varKey In dSorted.Keys
            Set dEntry = New Dictionary
            Set colInner = New Collection
            For Each varInner In Split(CStr(varKey), this.PathSep)
                colInner.Add CStr(varInner)
            Next varInner
            dEntry.Add "path", colInner
            dEntry.Add "depth", colInner.Count
            dEntry.Add "count", dSorted(varKey).Count
            dEntry.Add "exclusiveSeconds", CDbl(dSorted(varKey).Total)
            dEntry.Add "inclusiveSeconds", CDbl(InclusiveTotal(CStr(varKey)))
            colRows.Add dEntry
        Next varKey
    End If
    dData.Add "callPaths", colRows

    ' Callers: for each operation, which parent drove it, and how often. Built from the
    ' same path data, grouped the other way round.
    Set colRows = New Collection
    If Not this.Paths Is Nothing Then
        Set dByName = New Dictionary
        Set dNameTotals = New Dictionary
        For Each varKey In this.Paths.Keys
            strName = LastSegment(CStr(varKey))
            strCaller = ParentSegment(CStr(varKey))
            If Not dByName.Exists(strName) Then dByName.Add strName, New Dictionary
            Set dParents = dByName(strName)
            If Not dParents.Exists(strCaller) Then dParents.Add strCaller, New clsPerformanceItem
            Set cItem = dParents(strCaller)
            cItem.Total = cItem.Total + this.Paths(varKey).Total
            cItem.Count = cItem.Count + this.Paths(varKey).Count
            If Not dNameTotals.Exists(strName) Then dNameTotals.Add strName, New clsPerformanceItem
            Set cItem = dNameTotals(strName)
            cItem.Total = cItem.Total + this.Paths(varKey).Total
            cItem.Count = cItem.Count + this.Paths(varKey).Count
        Next varKey

        Set dSorted = SortItemsByTime(dNameTotals)
        For Each varKey In dSorted.Keys
            strName = CStr(varKey)
            Set dParents = dByName(strName)
            Set dEntry = New Dictionary
            dEntry.Add "operation", strName
            dEntry.Add "count", dSorted(varKey).Count
            dEntry.Add "exclusiveSeconds", CDbl(dSorted(varKey).Total)
            dEntry.Add "callerCount", dParents.Count
            Set colInner = New Collection
            Set dSortedParents = SortItemsByTime(dParents)
            For Each varInner In dSortedParents.Keys
                Set dCaller = New Dictionary
                dCaller.Add "caller", CStr(varInner)
                dCaller.Add "count", dSortedParents(varInner).Count
                dCaller.Add "exclusiveSeconds", CDbl(dSortedParents(varInner).Total)
                colInner.Add dCaller
            Next varInner
            dEntry.Add "calledFrom", colInner
            colRows.Add dEntry
        Next varKey
    End If
    dData.Add "callers", colRows

    ' Operations left open, which make the numbers above suspect
    Set colRows = New Collection
    If Not this.CallStackItems Is Nothing Then
        For Each varKey In this.CallStackItems
            colRows.Add CStr(varKey)
        Next varKey
    End If
    dData.Add "unfinishedOperations", colRows

    Set GetDiagnosticData = dData

End Function


'---------------------------------------------------------------------------------------
' Procedure : InclusiveTotal
' Author    : Adam Waller
' Date      : 8/27/2026
' Purpose   : Exclusive time for a call path plus everything recorded beneath it.
'---------------------------------------------------------------------------------------
'
Private Function InclusiveTotal(strPath As String) As Currency

    Dim varKey As Variant
    Dim strPrefix As String
    Dim curTotal As Currency

    If this.Paths Is Nothing Then Exit Function

    strPrefix = strPath & this.PathSep
    For Each varKey In this.Paths.Keys
        If CStr(varKey) = strPath Then
            curTotal = curTotal + this.Paths(varKey).Total
        ElseIf Left$(CStr(varKey), Len(strPrefix)) = strPrefix Then
            curTotal = curTotal + this.Paths(varKey).Total
        End If
    Next varKey

    InclusiveTotal = curTotal

End Function


'---------------------------------------------------------------------------------------
' Procedure : LastSegment
' Author    : Adam Waller
' Date      : 8/27/2026
' Purpose   : The operation name at the end of a call path.
'---------------------------------------------------------------------------------------
'
Private Function LastSegment(strPath As String) As String
    Dim lngPos As Long
    lngPos = InStrRev(strPath, this.PathSep)
    If lngPos = 0 Then
        LastSegment = strPath
    Else
        LastSegment = Mid$(strPath, lngPos + Len(this.PathSep))
    End If
End Function


'---------------------------------------------------------------------------------------
' Procedure : ParentSegment
' Author    : Adam Waller
' Date      : 8/27/2026
' Purpose   : The operation that called the one at the end of a call path, or a marker
'           : for paths that start at the top level.
'---------------------------------------------------------------------------------------
'
Private Function ParentSegment(strPath As String) As String
    Dim lngPos As Long
    lngPos = InStrRev(strPath, this.PathSep)
    If lngPos = 0 Then
        ParentSegment = "(top level)"
    Else
        ParentSegment = LastSegment(Left$(strPath, lngPos - 1))
    End If
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetReports
' Author    : Adam Waller
' Date      : 11/3/2020
' Purpose   : Return report text
'---------------------------------------------------------------------------------------
'
Public Function GetReports() As String

    Const cstrTitle As String = "PERFORMANCE REPORTS"

    Dim varKey As Variant
    Dim curTotal As Currency
    Dim dblCount As Double
    Dim lngCol(0 To 2) As Long
    Dim strSpacer As String

    If Not Me.Enabled Then Exit Function

    ' Set up column sizes
    lngCol(0) = 30
    lngCol(1) = 10
    lngCol(2) = 10
    strSpacer = Space(lngCol(0) + lngCol(1) + lngCol(2))
    strSpacer = Replace(strSpacer, " ", "-")

    With New clsConcat
        .AppendOnAdd = vbCrLf
        .Add strSpacer
        .Add Space((Len(strSpacer) - Len(cstrTitle)) / 2) & cstrTitle

        ' Table for object types
        If Not this.Categories Is Nothing Then
            Dim strCatName As String
            Set this.Categories = SortItemsByTime(this.Categories)
            .Add strSpacer
            .Add ListResult("Category", "Count", "Seconds", lngCol), vbCrLf, strSpacer
            For Each varKey In this.Categories.Keys
                strCatName = CStr(varKey)
                ' Append footnote marks if this category has annotations
                If Not this.CategoryFootnotes Is Nothing Then
                    If this.CategoryFootnotes.Exists(strCatName) Then
                        strCatName = strCatName & " " & this.CategoryFootnotes(strCatName)
                    End If
                End If
                .Add ListResult(strCatName, CStr(this.Categories(varKey).Count), _
                    Format$(this.Categories(varKey).Total, "0." & String$(this.intDigitsAfterDecimal, "0")), lngCol)
                ' Add to totals
                dblCount = dblCount + this.Categories(varKey).Count
                curTotal = curTotal + this.Categories(varKey).Total
            Next varKey
            .Add strSpacer
            .Add ListResult("TOTALS:", CStr(dblCount), _
                Format$(curTotal, "0." & String$(this.intDigitsAfterDecimal, "0")), lngCol)
            ' Render footnote descriptions
            If Not this.FootnoteMarks Is Nothing Then
                If this.FootnoteMarks.Count > 0 Then
                    For Each varKey In this.FootnoteMarks.Keys
                        .Add CStr(varKey) & " " & CStr(this.FootnoteMarks(varKey))
                    Next varKey
                End If
            End If
            .Add strSpacer
            .Add vbNullString
        End If

        ' Table for operations
        If Not this.Operations Is Nothing Then
            Set this.Operations = SortItemsByTime(this.Operations)
            curTotal = 0
            .Add strSpacer
            .Add ListResult("Operations", "Count", "Seconds", lngCol), vbCrLf, strSpacer
            For Each varKey In this.Operations.Keys
                .Add ListResult(CStr(varKey), CStr(this.Operations(varKey).Count), _
                    Format$(this.Operations(varKey).Total, "0." & String$(this.intDigitsAfterDecimal, "0")), lngCol)
                curTotal = curTotal + this.Operations(varKey).Total
            Next varKey
            .Add strSpacer
            If Not this.Overall Is Nothing Then
                .Add ListResult("Other Operations", vbNullString, _
                    Format$(this.Overall.Total - curTotal, "0." & String$(this.intDigitsAfterDecimal, "0")), lngCol)
                .Add ListResult("TOTAL RUNTIME:", vbNullString, _
                    Format$(this.Overall.Total, "0." & String$(this.intDigitsAfterDecimal, "0")), lngCol)
                .Add strSpacer
            End If
            .Add vbNullString
        End If

        ' Check for unfinished operations
        If Not this.CallStackItems Is Nothing Then
            If this.CallStackItems.Count > 0 Then
                .Add vbNullString
                .Add "WARNING: The performance monitoring for operations still"
                .Add "had items in the call stack. This typically happens when"
                .Add "performance logging is started for an operation, but not"
                .Add "closed at the conclusion of the operation."
                .Add "The call stack currently contains the following ", this.CallStackItems.Count, " items:"
                For Each varKey In this.CallStackItems
                    .Add " - ", CStr(varKey)
                Next varKey
            End If
        End If

        ' Return report section
        GetReports = .GetStr
    End With

End Function


'---------------------------------------------------------------------------------------
' Procedure : ListResult
' Author    : Adam Waller
' Date      : 11/3/2020
' Purpose   : List the result of a test in a fixed width format. The result strings
'           : are positioned at the number of characters specified. If the heading size
'           : exceeds the width of the column, the text will be wrapped.
'           : I.e:
'           : MyFancyTest      23     2.45
'           : My very long nam
'           : e that I probabl
'           : y should condens
'           : e                12     3.23
'---------------------------------------------------------------------------------------
'
Private Function ListResult(strHeading As String, _
                            strResult1 As String, _
                            strResult2 As String, _
                            lngCol() As Long) As String

    Dim strRowHeading As String
    Dim lngPos As Long
    Dim intMax As Integer

    ' Wrap at one character less than the column width
    intMax = lngCol(0) - 1

    ' Use concatenation class in case we need to deal with line wrapping
    With New clsConcat

        ' Check for size overflow on heading. (Wrap on multiple lines)
        strRowHeading = strHeading
        If Len(strRowHeading) > intMax Then
            lngPos = 1
            Do While lngPos + intMax <= Len(strHeading)
                ' Add segment and linebreak
                .Add Mid$(strHeading, lngPos, intMax), " ", vbCrLf
                lngPos = lngPos + intMax
            Loop
            ' Get last heading line to use with results
            strRowHeading = Mid$(strHeading, lngPos)
        End If

        ' Display heading and amounts
        .Add PadRight(strRowHeading, lngCol(0))
        .Add PadRight(strResult1, lngCol(1))
        .Add strResult2
        ListResult = .GetStr
    End With

End Function


'---------------------------------------------------------------------------------------
' Procedure : PadRight
' Author    : Adam Waller
' Date      : 11/3/2020
' Purpose   : Pads a string
'---------------------------------------------------------------------------------------
'
Private Function PadRight(strText As String, _
                        lngLen As Long, _
                        Optional lngMinTrailingSpaces As Long = 1) As String

    Dim strResult As String
    Dim strTrimmed As String

    strResult = Space$(lngLen)
    strTrimmed = Left$(strText, lngLen - lngMinTrailingSpaces)

    ' Use mid function to write over existing string of spaces.
    Mid$(strResult, 1, Len(strTrimmed)) = strTrimmed
    PadRight = strResult

End Function


'---------------------------------------------------------------------------------------
' Procedure : SortItemsByTime
' Author    : Adam Waller
' Date      : 4/4/2023
' Purpose   : Sort the items by total time in descending order to put the slowest
'           : items at the top of the list.
'---------------------------------------------------------------------------------------
'
Private Function SortItemsByTime(dItems As Dictionary) As Dictionary

    Dim varItems() As Variant
    Dim varKey As Variant
    Dim lngCnt As Long
    Dim strRecord As String
    Dim cItem As clsPerformanceItem
    Dim dSorted As Dictionary

    ' Nothing to sort, and ReDim to an upper bound of -1 would fail
    If dItems.Count = 0 Then
        Set SortItemsByTime = dItems
        Exit Function
    End If

    ' Create an array to build our items
    ReDim varItems(0 To dItems.Count - 1)

    ' The idea here is that we want to build a fixed width record with the time on
    ' the left, and the key on the right. We will then sort this list of "records"
    ' to get them sorted by time. However, since the sorting is in ascending order,
    ' we will loop backwards through the resulting array to rebuild a dictionary
    ' of items in descending order. We will split out the key and use this to get
    ' a reference back to the original item in the source dictionary.

    ' Build our list of records
    For Each varKey In dItems.Keys
        ' Create a record like this: "00062840.170000|Export Form Objects"
        strRecord = Format$(dItems(varKey).Total, "00000000.000000") & "|" & CStr(varKey)
        ' Add to array.
        varItems(lngCnt) = strRecord
        ' Increment counter for array
        lngCnt = lngCnt + 1
    Next varKey

    ' Sort the array in ascending order
    QuickSort varItems

    ' Now, build a new dictionary with the sorted items
    ' (We are walking backwards through the array to flip the sort to descending)
    Set dSorted = New Dictionary
    For lngCnt = dItems.Count - 1 To 0 Step -1
        ' Parse key from record (text after first pipe character)
        strRecord = varItems(lngCnt)
        varKey = Mid$(strRecord, InStr(1, strRecord, "|") + 1)
        ' Reference performance item class
        Set cItem = dItems(varKey)
        ' Add to dictionary of resorted items
        dSorted.Add varKey, cItem
    Next lngCnt

    ' Return the sorted dictionary
    Set SortItemsByTime = dSorted

End Function


'---------------------------------------------------------------------------------------
' Procedure : Reset
' Author    : Adam Waller
' Date      : 11/3/2020
' Purpose   : Reset all class values to start timing again.
'---------------------------------------------------------------------------------------
'
Public Sub Reset()
    Set this.Overall = Nothing
    Set this.CallStackItems = Nothing
    this.CategoryName = vbNullString
    Set this.Categories = Nothing
    this.OperationName = vbNullString
    Set this.Operations = Nothing
    Set this.Paths = Nothing
    Set this.FootnoteMarks = Nothing
    Set this.CategoryFootnotes = Nothing
    Me.Enabled = True
End Sub


'---------------------------------------------------------------------------------------
' Procedure : Class_Initialize
' Author    : Adam Waller
' Date      : 11/5/2020
' Purpose   : Initialize objects for timing.
'---------------------------------------------------------------------------------------
'
Private Sub Class_Initialize()
    ' this.curFrequency need only be queried once
    ' https://docs.microsoft.com/en-us/windows/win32/api/profileapi/nf-profileapi-queryperformancefrequency
    GetFrequencyAPI this.Frequency
    this.intDigitsAfterDecimal = 2
    ' A control character cannot occur in an operation name, so it can never be mistaken
    ' for part of one when a path is split back into segments.
    this.PathSep = ChrW$(1)
    Me.Enabled = True
End Sub
