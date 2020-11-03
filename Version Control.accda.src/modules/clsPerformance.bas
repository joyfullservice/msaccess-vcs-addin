Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : clsPerformance
' Author    : Adam Waller
' Date      : 11/3/2020
' Purpose   : Measure the performance of the export/import process. Since different
'           : users have different needs and work with sometimes very different
'           : databases, this tool will help identify potential bottlenecks in the
'           : performance of the add-in in real-life scenarios. The results are
'           : typically added to the log files.
'           : Note: This class uses the Timer() function for basic timing operations.
'---------------------------------------------------------------------------------------

Option Compare Database
Option Explicit


Private m_Overall As clsPerformanceItem
Private m_strComponent As String
Private m_dComponents As Dictionary
Private m_strOperation As String
Private m_dOperations As Dictionary

' Manage a type of call stack to track nested operations.
' When an operation finishes, it goes back to timing the
' previous operation.
Private m_colOpsCallStack As Collection


'---------------------------------------------------------------------------------------
' Procedure : StartTiming
' Author    : Adam Waller
' Date      : 11/3/2020
' Purpose   : Start the overall timing.
'---------------------------------------------------------------------------------------
'
Public Sub StartTiming()
    ResetAll
    m_Overall.sngStart = Timer
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
    dItems(strName).sngStart = Timer
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
        If .sngStart > 0 Then
            .sngTotal = .sngTotal + GetElapsed(.sngStart)
            .sngStart = 0
            .lngCount = .lngCount + lngCount
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
Private Function GetElapsed(sngStart As Single) As Single

    Dim sngNow As Single
    Dim sngTotal As Single
    
    ' Only return a value if we have a starting time.
    If sngStart > 0 Then
        sngNow = Timer
        If sngStart <= sngNow Then
            GetElapsed = sngNow - sngStart
        Else
            ' Just in case someone was up really late, and crossed midnight...
            GetElapsed = sngStart + ((24# * 60 * 60) - sngStart)
        End If
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

    Const cstrSpacer As String = "-------------------------------------"
    
    Dim varKey As Variant

    With New clsConcat
        .AppendOnAdd = vbCrLf
        .Add cstrSpacer
        .Add "        PERFORMANCE REPORTS"
        .Add cstrSpacer, vbCrLf
        
        ' Table for object types
        .Add ListResult("Object Type", "Count", "Seconds", 20, 30), vbCrLf, cstrSpacer
        For Each varKey In m_dComponents.Keys
            .Add ListResult(CStr(varKey), CStr(m_dComponents(varKey).lngCount), _
                Format(m_dComponents(varKey).sngTotal, "0.00"), 20, 30)
        Next varKey
        .Add vbNullString
        
        ' Table for operations
        .Add ListResult("Operations", "Count", "Seconds", 20, 30), vbCrLf, cstrSpacer
        For Each varKey In m_dOperations.Keys
            .Add ListResult(CStr(varKey), CStr(m_dOperations(varKey).lngCount), _
                Format(m_dOperations(varKey).sngTotal, "0.00"), 20, 30)
        Next varKey
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
    intResultPos1 As Integer, intResultPos2 As Integer) As String
    ListResult = PadRight(strHeading, intResultPos1) & _
        PadRight(strResult1, intResultPos2 - intResultPos1) & strResult2
End Function


'---------------------------------------------------------------------------------------
' Procedure : PadRight
' Author    : Adam Waller
' Date      : 11/3/2020
' Purpose   : Pads a string
'---------------------------------------------------------------------------------------
'
Private Function PadRight(strText As String, intLen As Integer, Optional intMinTrailingSpaces As Integer = 1) As String

    Dim strResult As String
    Dim strTrimmed As String
    
    strResult = Space$(intLen)
    strTrimmed = Left$(strText, intLen - intMinTrailingSpaces)
    
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
    Set m_Overall = New clsPerformanceItem
    Set m_dComponents = New Dictionary
    Set m_dOperations = New Dictionary
    m_strComponent = vbNullString
    m_strOperation = vbNullString
    Set m_colOpsCallStack = New Collection
End Sub