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


Private m_Overall As clsPerformanceItem
Private m_strComponent As String
Private m_dComponents As Scripting.Dictionary
Private m_strOperation As String
Private m_dOperations As Scripting.Dictionary
Private m_curFrequency As Currency
Private m_intDigitsAfterDecimal As Integer

' API calls to get more precise time than Timer function
Private Declare PtrSafe Function GetFrequencyAPI Lib "kernel32" Alias "QueryPerformanceFrequency" (ByRef Frequency As Currency) As Long
Private Declare PtrSafe Function GetTimeAPI Lib "kernel32" Alias "QueryPerformanceCounter" (ByRef Counter As Currency) As Long

' Manage a type of call stack to track nested operations.
' When an operation finishes, it goes back to timing the
' previous operation.
Private m_colOpsCallStack As VBA.Collection


'---------------------------------------------------------------------------------------
' Procedure : StartTiming
' Author    : Adam Waller
' Date      : 11/3/2020
' Purpose   : Start the overall timing.
'---------------------------------------------------------------------------------------
'
Public Sub StartTiming()
    ResetAll
    m_Overall.Start = MicroTimer
End Sub


'---------------------------------------------------------------------------------------
' Procedure : ComponentStart
' Author    : Adam Waller
' Date      : 11/3/2020
' Purpose   : Start timing a component type.
'---------------------------------------------------------------------------------------
'
Public Sub ComponentStart(strName As String)
    StartTimer m_dComponents, strName
    m_strComponent = strName
End Sub


'---------------------------------------------------------------------------------------
' Procedure : ComponentEnd
' Author    : Adam Waller
' Date      : 11/3/2020
' Purpose   : End the timing of the active component
'---------------------------------------------------------------------------------------
'
Public Sub ComponentEnd(Optional lngCount As Long = 1)
    If m_strComponent <> vbNullString Then
        LapTimer m_dComponents(m_strComponent), lngCount
        m_strComponent = vbNullString
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
    
    ' See if we are already timing something
    If m_strOperation <> vbNullString Then
    
        ' We are already timing something else right now.
        ' Save the current process to the call stack before switching
        ' to the new operation.
        LapTimer m_dOperations(m_strOperation), 0
        With m_colOpsCallStack
            ' Safety check!
            If .Count < 100 Then .Add m_strOperation
        End With
    End If
    
    ' Start the timer for this operation.
    StartTimer m_dOperations, strName
    m_strOperation = strName
    
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

    ' Verify that we are timing something, and record the elapsed time.
    If m_strOperation <> vbNullString Then

        ' Record the elapsed time.
        LapTimer m_dOperations(m_strOperation), lngCount
        
        ' Check the call stack to see if we need to move back to the previous process.
        With m_colOpsCallStack
            If .Count > 0 Then
                ' Resume previous activity
                strLastOperation = .Item(.Count)
                m_strOperation = vbNullString
                OperationStart strLastOperation
                ' Remove last item from call stack
                .Remove .Count
            Else
                ' No longer timing any operations.
                m_strOperation = vbNullString
            End If
        End With
    End If
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : DigitsAfterDecimal
' Author    : Eugen Albiker
' Date      : 16/3/2021
' Purpose   : Set Number of Digits after Decimal for the Elapsed Time
'---------------------------------------------------------------------------------------
'
Public Property Let DigitsAfterDecimal(intDigitsAfterDecimal As Integer)
    If intDigitsAfterDecimal > 4 Then intDigitsAfterDecimal = 4
    m_intDigitsAfterDecimal = intDigitsAfterDecimal
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
    LapTimer m_Overall, 1
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
    MicroTimer = (curTime / m_curFrequency)
    
End Function


'---------------------------------------------------------------------------------------
' Procedure : StartTimer
' Author    : Adam Waller
' Date      : 11/3/2020
' Purpose   : Add the item if it doesn't exist, then set the start time.
'---------------------------------------------------------------------------------------
'
Private Sub StartTimer(dItems As Scripting.Dictionary, strName As String)
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
    With cItem
        If .Start > 0 Then
            .Total = .Total + GetElapsed(.Start)
            .Start = 0
            .Count = .Count + lngCount
        End If
    End With
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
    TotalTime = GetElapsed(m_Overall.Start)
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
        If m_dComponents.Count > 0 Then
            .Add strSpacer
            .Add ListResult("Object Type", "Count", "Seconds", lngCol), vbCrLf, strSpacer
            For Each varKey In m_dComponents.Keys
                .Add ListResult(CStr(varKey), CStr(m_dComponents(varKey).Count), _
                    Format(m_dComponents(varKey).Total, "0." & String$(m_intDigitsAfterDecimal, "0")), lngCol)
                ' Add to totals
                dblCount = dblCount + m_dComponents(varKey).Count
                curTotal = curTotal + m_dComponents(varKey).Total
            Next varKey
            .Add strSpacer
            .Add ListResult("TOTALS:", CStr(dblCount), _
                Format(curTotal, "0." & String$(m_intDigitsAfterDecimal, "0")), lngCol)
            .Add strSpacer
            .Add vbNullString
        End If
        ' Table for operations
        curTotal = 0
        .Add strSpacer
        .Add ListResult("Operations", "Count", "Seconds", lngCol), vbCrLf, strSpacer
        For Each varKey In m_dOperations.Keys
            .Add ListResult(CStr(varKey), CStr(m_dOperations(varKey).Count), _
                Format(m_dOperations(varKey).Total, "0." & String$(m_intDigitsAfterDecimal, "0")), lngCol)
            curTotal = curTotal + m_dOperations(varKey).Total
        Next varKey
        .Add strSpacer
        .Add ListResult("Other Operations", vbNullString, _
            Format(m_Overall.Total - curTotal, "0." & String$(m_intDigitsAfterDecimal, "0")), lngCol)
        .Add strSpacer
        .Add vbNullString
        
        ' Check for unfinished operations
        If m_colOpsCallStack.Count > 0 Then
            .Add vbNullString
            .Add "WARNING: The performance monitoring for operations still"
            .Add "had items in the call stack. This typically happens when"
            .Add "performance logging is started for an operation, but not"
            .Add "closed at the conclusion of the operation."
            .Add "The call stack currently contains the following ", m_colOpsCallStack.Count, " items:"
            For Each varKey In m_colOpsCallStack
                .Add " - ", CStr(varKey)
            Next varKey
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
'           : are positioned at the number of characters specified.
'           : I.e:
'           : MyFancyTest      23     2.45
'---------------------------------------------------------------------------------------
'
Private Function ListResult(strHeading As String, strResult1 As String, strResult2 As String, _
    lngCol() As Long) As String
    ListResult = PadRight(strHeading, lngCol(0)) & _
        PadRight(strResult1, lngCol(1)) & strResult2
End Function


'---------------------------------------------------------------------------------------
' Procedure : PadRight
' Author    : Adam Waller
' Date      : 11/3/2020
' Purpose   : Pads a string
'---------------------------------------------------------------------------------------
'
Private Function PadRight(strText As String, lngLen As Long, Optional lngMinTrailingSpaces As Long = 1) As String

    Dim strResult As String
    Dim strTrimmed As String
    
    strResult = Space$(lngLen)
    strTrimmed = Left$(strText, lngLen - lngMinTrailingSpaces)
    
    ' Use mid function to write over existing string of spaces.
    Mid$(strResult, 1, Len(strTrimmed)) = strTrimmed
    PadRight = strResult
    
End Function


'---------------------------------------------------------------------------------------
' Procedure : ResetAll
' Author    : Adam Waller
' Date      : 11/3/2020
' Purpose   : Reset all class values
'---------------------------------------------------------------------------------------
'
Private Sub ResetAll()
    Class_Initialize
    m_strComponent = vbNullString
    m_strOperation = vbNullString
    m_intDigitsAfterDecimal = 2
End Sub


'---------------------------------------------------------------------------------------
' Procedure : Class_Initialize
' Author    : Adam Waller
' Date      : 11/5/2020
' Purpose   : Initialize objects for timing.
'---------------------------------------------------------------------------------------
'
Private Sub Class_Initialize()
    Set m_Overall = New clsPerformanceItem
    Set m_dComponents = New Scripting.Dictionary
    Set m_dOperations = New Scripting.Dictionary
    Set m_colOpsCallStack = New VBA.Collection

    ' m_curFrequency need only be queried once
    ' https://docs.microsoft.com/en-us/windows/win32/api/profileapi/nf-profileapi-queryperformancefrequency
    GetFrequencyAPI m_curFrequency
End Sub