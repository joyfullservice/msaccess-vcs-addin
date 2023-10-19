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

' Enable or disable timing operations
Public Enabled As Boolean

' Use type for private class variables
Private Type udtPerformance
    Overall As clsPerformanceItem
    CategoryName As String
    Categories As Dictionary
    OperationName As String
    Operations As Dictionary
    Frequency As Currency
    intDigitsAfterDecimal As Integer
    ' Manage a type of call stack to track nested operations.
    ' When an operation finishes, it goes back to timing the
    ' previous operation.
    CallStackItems As Collection
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
    Reset
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
        If this.CallStackItems Is Nothing Then Set this.CallStackItems = New Collection
        With this.CallStackItems
            ' Safety check!
            If .Count < 100 Then .Add this.OperationName
        End With
    End If

    ' Start the timer for this operation.
    StartTimer this.Operations, strName
    this.OperationName = strName

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

        ' Check the call stack to see if we need to move back to the previous process.
        If Not this.CallStackItems Is Nothing Then
            With this.CallStackItems
                If .Count > 0 Then
                    ' Resume previous activity
                    strLastOperation = .Item(.Count)
                    this.OperationName = vbNullString
                    OperationStart strLastOperation
                    ' Remove last item from call stack
                    .Remove .Count
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
            Set this.Categories = SortItemsByTime(this.Categories)
            .Add strSpacer
            .Add ListResult("Category", "Count", "Seconds", lngCol), vbCrLf, strSpacer
            For Each varKey In this.Categories.Keys
                .Add ListResult(CStr(varKey), CStr(this.Categories(varKey).Count), _
                    Format(this.Categories(varKey).Total, "0." & String$(this.intDigitsAfterDecimal, "0")), lngCol)
                ' Add to totals
                dblCount = dblCount + this.Categories(varKey).Count
                curTotal = curTotal + this.Categories(varKey).Total
            Next varKey
            .Add strSpacer
            .Add ListResult("TOTALS:", CStr(dblCount), _
                Format(curTotal, "0." & String$(this.intDigitsAfterDecimal, "0")), lngCol)
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
                    Format(this.Operations(varKey).Total, "0." & String$(this.intDigitsAfterDecimal, "0")), lngCol)
                curTotal = curTotal + this.Operations(varKey).Total
            Next varKey
            .Add strSpacer
            If Not this.Overall Is Nothing Then
                .Add ListResult("Other Operations", vbNullString, _
                    Format(this.Overall.Total - curTotal, "0." & String$(this.intDigitsAfterDecimal, "0")), lngCol)
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
Private Function ListResult(strHeading As String _
                            , strResult1 As String _
                            , strResult2 As String _
                            , lngCol() As Long) As String

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
Private Function PadRight(strText As String _
                        , lngLen As Long _ 
                        , Optional lngMinTrailingSpaces As Long = 1) As String

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
        ' Create a record like this: "00062840.170000|Export Form Objects       ..."
        strRecord = Format(dItems(varKey).Total, "00000000.000000") & "|" & PadRight(CStr(varKey), 255)
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
        ' Parse key from record
        varKey = Trim(Split(varItems(lngCnt), "|")(1))
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
    Enabled = True
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
    Enabled = True
End Sub
