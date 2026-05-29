VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
'@Folder("Infrastructure")

Public PadLength As Integer
Public SourcePath As String

' Set this to true when logging an operation such as an export or build
' then set back to false after writing the log file. This affects
' how error messages are reported to the user outside of operations.
Public Active As Boolean

Private Const cstrSpacer As String = "-------------------------------------"

' Throttle interval for Log.Increment. The visible bar work and DoEvents are
' throttled separately by clsLblProg, so 10 updates/sec here is plenty while
' eliminating the Perf wrap cost on the rest of the calls.
Private Const INCREMENT_THROTTLE_SEC As Single = 0.1

Private m_Log As clsConcat      ' Log file output
Private m_Console As clsConcat  ' Console output
Private m_RichText As TextBox   ' Text box to display HTML
Private m_Prog As clsLblProg    ' Progress bar
Private m_blnProgressActive As Boolean
Private m_blnKeepProgressVisible As Boolean
Private m_blnConsoleDirty As Boolean  ' True when m_Console has new content not yet assigned to txtLog
Private m_sngLastUpdate As Single
Private m_sngLastMCPProgress As Single  ' Throttle MCP progress updates
Private m_dblPendingIncrements As Double  ' Accumulated increments awaiting throttled flush
Private m_sngLastIncrement As Single    ' Timer value of the last applied Increment
Private m_OperationId As String ' Timestamp with milliseconds: "20260126_153042_847"
Private m_lngErrCount(eelNoError To eelCritical) As Long
Private m_colErrorJournal As Collection


'---------------------------------------------------------------------------------------
' Procedure : OperationId
' Author    : Adam Waller
' Date      : 1/26/2026
' Purpose   : Return the unique operation ID (timestamp with milliseconds)
'---------------------------------------------------------------------------------------
'
Public Property Get OperationId() As String
    OperationId = m_OperationId
End Property


'---------------------------------------------------------------------------------------
' Procedure : WarningCount
' Author    : Adam Waller
' Date      : 04/09/2026
' Purpose   : Return the number of warnings logged during this operation.
'---------------------------------------------------------------------------------------
'
Public Property Get WarningCount() As Long
    WarningCount = m_lngErrCount(eelWarning)
End Property


'---------------------------------------------------------------------------------------
' Procedure : ErrorCount
' Author    : Adam Waller
' Date      : 04/09/2026
' Purpose   : Return the number of errors (eelError + eelCritical) logged.
'---------------------------------------------------------------------------------------
'
Public Property Get ErrorCount() As Long
    ErrorCount = m_lngErrCount(eelError) + m_lngErrCount(eelCritical)
End Property


'---------------------------------------------------------------------------------------
' Procedure : ErrorJournalCount
' Author    : Adam Waller
' Date      : 5/29/2026
' Purpose   : Return the number of Log.Error entries recorded in the error journal.
'---------------------------------------------------------------------------------------
'
Public Function ErrorJournalCount() As Long
    ErrorJournalCount = m_colErrorJournal.Count
End Function


'---------------------------------------------------------------------------------------
' Procedure : ClearErrorJournal
' Author    : Adam Waller
' Date      : 5/29/2026
' Purpose   : Clear the error journal (call at the start of a test run).
'---------------------------------------------------------------------------------------
'
Public Sub ClearErrorJournal()
    Set m_colErrorJournal = New Collection
End Sub


'---------------------------------------------------------------------------------------
' Procedure : GetErrorJournalSince
' Author    : Adam Waller
' Date      : 5/29/2026
' Purpose   : Return journal entries logged after the given index (for per-test capture).
'---------------------------------------------------------------------------------------
'
Public Function GetErrorJournalSince(lngStart As Long) As Collection

    Dim colResult As Collection
    Dim i As Long

    Set colResult = New Collection
    For i = lngStart + 1 To m_colErrorJournal.Count
        colResult.Add m_colErrorJournal(i)
    Next i
    Set GetErrorJournalSince = colResult

End Function


'---------------------------------------------------------------------------------------
' Procedure : Clear
' Author    : Adam Waller
' Date      : 4/16/2020
' Purpose   : Clear the log buffers and release object references
'---------------------------------------------------------------------------------------
'
Public Sub Clear()
    Class_Initialize
End Sub


'---------------------------------------------------------------------------------------
' Procedure : Spacer
' Author    : Adam Waller
' Date      : 4/28/2020
' Purpose   : Add a spacer to the log
'---------------------------------------------------------------------------------------
'
Public Sub Spacer(Optional blnPrint As Boolean = True, Optional lngLength As Long = 0)
    If lngLength > 0 Then
        Me.Add String$(lngLength, "-"), blnPrint
    Else
        Me.Add cstrSpacer, blnPrint
    End If
End Sub


'---------------------------------------------------------------------------------------
' Procedure : Add
' Author    : Adam Waller
' Date      : 1/18/2019
' Purpose   : Add a log file entry.
'---------------------------------------------------------------------------------------
'
Public Sub Add(strText As String, Optional blnPrint As Boolean = True, _
    Optional blnNextOutputOnNewLine As Boolean = True, _
    Optional strColor As String = vbNullString, _
    Optional blnBold As Boolean = False, _
    Optional blnItalic As Boolean = False)

    ' Add to log file output (respect inline continuation)
    m_Log.Add strText
    If blnNextOutputOnNewLine Then m_Log.Add vbCrLf

    ' See if we want to print the output of this message.
    If blnPrint Then
        ' Remove existing progress indicator if in use.
        If m_blnProgressActive And Not m_blnKeepProgressVisible Then
            m_blnProgressActive = False
            m_Prog.Hide
        End If

        ' Build HTML output for console
        With m_Console

            ' Opening tags
            If blnBold Then .Add "<strong>"
            If blnItalic Then .Add "<i>"
            If strColor <> vbNullString Then .Add "<font color=", strColor, ">"

            ' Content
            .Add MultiReplace(strText, _
                " ", "&nbsp;", _
                vbCrLf, "<br>")

            ' Closing tags
            If strColor <> vbNullString Then .Add "</font>"
            If blnItalic Then .Add "</i>"
            If blnBold Then .Add "</strong>"

            ' Add line break for HTML
            If blnNextOutputOnNewLine Then m_Console.Add "<br>"
        End With
        m_blnConsoleDirty = True

        ' Run debug output
        If m_RichText Is Nothing Then
            ' Only print debug output if not running from the GUI.
            If blnNextOutputOnNewLine Then
                ' Create new line
                Debug.Print strText
            Else
                ' Continue next printout on this line.
                Debug.Print strText;
            End If
        End If

        ' Allow an update to the screen every second.
        ' (This keeps the aplication from an apparent hang while
        '  running intensive export processes.)
        If m_sngLastUpdate + 1 < Timer Then
            Me.Flush
            m_sngLastUpdate = Timer
        End If

        ' Send to MCP if active (plain text, no HTML)
        If MCP.IsActive Then
            MCP.PostCallback "log", -1, -1, strText
        End If
    End If

End Sub


'---------------------------------------------------------------------------------------
' Procedure : Flush
' Author    : Adam Waller
' Date      : 4/29/2020
' Purpose   : Flushes the buffer to the console
'---------------------------------------------------------------------------------------
'
Public Sub Flush()

    ' Apply any throttled-but-not-yet-shown increments before refreshing the UI.
    ApplyPendingIncrements

    ' See if the GUI form is loaded.
    Perf.OperationStart "Console Updates"
    If Not m_RichText Is Nothing Then
        If Not m_blnKeepProgressVisible Then
            m_blnProgressActive = False
            If Not m_Prog Is Nothing Then m_Prog.Hide
        End If
        If m_blnConsoleDirty Then
            With m_RichText
                Echo False
                .Value = ConsoleHtml
                If VerifyFocus(m_RichText) Then
                    .SelStart = ClampInt(Len(Nz(.Value, vbNullString)))
                End If
                Echo True
            End With
            m_blnConsoleDirty = False
        End If
    Else
        ' Update display in immediate window if not using console
        DoEvents
    End If
    Perf.OperationEnd

End Sub


'---------------------------------------------------------------------------------------
' Procedure : Error
' Author    : Adam Waller
' Date      : 12/4/2020
' Purpose   : Log an error, and update error level if needed. Reads Err object values.
'           : A critical error will also present a message box with the details.
'---------------------------------------------------------------------------------------
'
Public Sub Error(eLevel As eErrorLevel, strDescription As String, Optional strSource As String, Optional blnIncludeErrorMessage As Boolean = False)

    Dim strPrefix As String
    Dim strDisplay As String
    Dim lngErrNum As Long
    Dim strErrDesc As String

    lngErrNum = Err.Number
    strErrDesc = Err.Description
    m_lngErrCount(eLevel) = m_lngErrCount(eLevel) + 1
    Select Case eLevel
        Case eelWarning:    strPrefix = "WARNING: "
        Case eelError:      strPrefix = "ERROR: "
        Case eelCritical:   strPrefix = "CRITICAL: "
    End Select

    ' Build the error message string.
    With New clsConcat

        ' Sometimes the error description is sufficient for the description
        If strDescription = vbNullString And lngErrNum <> 0 Then
            strDisplay = strPrefix & lngErrNum & " - " & strErrDesc
        ElseIf blnIncludeErrorMessage Then
            strDisplay = strPrefix & strDescription & "; Error " & lngErrNum & " - " & strErrDesc
        Else
            strDisplay = strPrefix & strDescription
        End If

        ' Log all errors, and display on the output screen anything higher than a warning
        Me.Add vbNullString, (eLevel > eelWarning)
        Me.Spacer (eLevel > eelWarning)
        Me.Add strDisplay, (eLevel > eelWarning), , "red"

        ' Log the additional error and source details to the log file
        If lngErrNum <> 0 Then .Add "Error ", CStr(lngErrNum), ": ", strErrDesc, " "
        If strSource <> vbNullString Then .Add "Source: ", strSource
        Me.Add .GetStr, False
        Me.Spacer (eLevel > eelWarning)
        Me.Flush ' Force a screen update.
        AppendErrorJournal eLevel, strDisplay, strSource, lngErrNum, strErrDesc
        ' Show message boxes only when no console is attached to receive the output.
        ' (Test runs and other console-driven operations log here instead of blocking
        ' with a dialog when Log.Active is false.)
        If m_RichText Is Nothing Then
            If Log.Active Then
                If eLevel = eelCritical Then
                    MsgBox2 "Unable to Continue", .GetStr, _
                        "Please review the log file for additional details.", vbCritical
                End If
            Else
                Select Case eLevel
                    Case eelNoError:    ' Do nothing
                    Case eelWarning:    MsgBox2 "Warning", strDisplay, .GetStr, vbInformation
                    Case eelError:      MsgBox2 "Error", strDisplay, .GetStr, vbExclamation
                    Case eelCritical:   MsgBox2 "Critical", strDisplay, .GetStr, vbCritical
                End Select
            End If
        End If
    End With

    ' Update error level if higher.
    If Operation.ErrorLevel < eLevel Then Operation.ErrorLevel = eLevel

End Sub


'---------------------------------------------------------------------------------------
' Procedure : SetConsole
' Author    : Adam Waller
' Date      : 4/28/2020
' Purpose   : Set a reference to the rich text box for output display, and the
'           : associated progress bar for logging progress.
'---------------------------------------------------------------------------------------
'
Public Sub SetConsole(txtRichText As TextBox, cProg As clsLblProg)
    Set m_Prog = cProg
    Set m_RichText = txtRichText
    If Not m_RichText Is Nothing Then
        m_RichText.AllowAutoCorrect = False
        m_blnConsoleDirty = (m_Console.Length > 0)
    End If
End Sub


'---------------------------------------------------------------------------------------
' Procedure : ReleaseConsole
' Author    : Adam Waller
' Date      : 5/5/2022
' Purpose   : Releases object references to the console form that were created with
'           : the SetConsole sub.
'---------------------------------------------------------------------------------------
'
Public Sub ReleaseConsole()
    Set m_RichText = Nothing
    Set m_Prog = Nothing
End Sub


'---------------------------------------------------------------------------------------
' Procedure : ProgressBar
' Author    : Adam Waller
' Date      : 11/6/2020
' Purpose   : Pass the Progress Bar reference to this class.
'---------------------------------------------------------------------------------------
'
Public Property Set ProgressBar(cProg As clsLblProg)
    Set m_Prog = cProg
End Property
Public Property Get ProgressBar() As clsLblProg
    Set ProgressBar = m_Prog
End Property


'---------------------------------------------------------------------------------------
' Procedure : KeepProgressVisible
' Author    : VCS contributors
' Date      : 4/29/2026
' Purpose   : Keep the form progress bar visible while console output is refreshed.
'---------------------------------------------------------------------------------------
'
Public Property Let KeepProgressVisible(ByVal blnKeepVisible As Boolean)
    m_blnKeepProgressVisible = blnKeepVisible
End Property

Public Property Get KeepProgressVisible() As Boolean
    KeepProgressVisible = m_blnKeepProgressVisible
End Property


'---------------------------------------------------------------------------------------
' Procedure : ProgMax
' Author    : Adam Waller
' Date      : 11/6/2020
' Purpose   : Wrapper to set max value for progress bar.
'---------------------------------------------------------------------------------------
'
Public Property Let ProgMax(lngMaxValue As Long)
    If Not m_Prog Is Nothing Then m_Prog.Max = lngMaxValue
End Property


'---------------------------------------------------------------------------------------
' Procedure : SaveFile
' Author    : Adam Waller
' Date      : 1/18/2019
' Purpose   : Saves the current log data to a file.
'---------------------------------------------------------------------------------------
'
Public Sub SaveFile(Optional strAlternatePath As String)
    Dim strPath As String
    Dim strBase As String
    Dim strLogsFolder As String

    strPath = Nz2(strAlternatePath, Me.LogFilePath)
    strLogsFolder = FSO.GetParentFolderName(strPath)

    ' Migrate legacy logs on first use of logs folder
    If Not FSO.FolderExists(strLogsFolder) Then
        MigrateLegacyLogs FSO.GetParentFolderName(strLogsFolder), strLogsFolder
    End If

    ' Ensure logs folder exists
    VerifyPath strPath

    ' Write log file (unique filename per operation)
    WriteFile m_Log.GetStr, strPath

    ' Cleanup old logs beyond MaxLogFiles
    Select Case Operation.OperationType
        Case eotExport:   strBase = "Export"
        Case eotBuild:    strBase = "Build"
        Case eotMerge:    strBase = "Merge"
        Case eotTestRun:  strBase = "TestRun"
        Case Else:        strBase = "Other"
    End Select
    CleanupOldLogs strLogsFolder, strBase, Options.MaxLogFiles
End Sub


'---------------------------------------------------------------------------------------
' Procedure : LogFilePath
' Author    : Adam Waller
' Date      : 5/13/2023
' Purpose   : Return the (default) full path to the current log file.
'---------------------------------------------------------------------------------------
'
Public Property Get LogFilePath() As String
    Dim strBase As String
    Dim strFolder As String
    Select Case Operation.OperationType
        Case eotExport:   strBase = "Export"
        Case eotBuild:    strBase = "Build"
        Case eotMerge:    strBase = "Merge"
        Case eotTestRun:  strBase = "TestRun"
        Case Else:        strBase = "Other"
    End Select
    ' Use logs subfolder with unique operation ID in filename
    ' e.g., Export_20260126_153042_847.log
    strFolder = FSO.BuildPath(Nz2(Me.SourcePath, Options.GetExportFolder), "logs")
    LogFilePath = FSO.BuildPath(strFolder, strBase & "_" & m_OperationId & ".log")
End Property


'---------------------------------------------------------------------------------------
' Procedure : Class_Initialize
' Author    : Adam Waller
' Date      : 4/28/2020
' Purpose   : Set initial options
'---------------------------------------------------------------------------------------
'
Private Sub Class_Initialize()
    Set m_Console = New clsConcat
    Set m_Log = New clsConcat
    Set m_colErrorJournal = New Collection
    m_blnProgressActive = False
    m_blnKeepProgressVisible = False
    m_sngLastUpdate = 0
    m_dblPendingIncrements = 0
    m_sngLastIncrement = 0
    Erase m_lngErrCount
    Me.PadLength = 30
    Me.Active = False
    ' Generate timestamp-based operation ID with millisecond precision
    m_OperationId = Format$(Now, "yyyymmdd\_hhnnss") & "_" & Format$((Timer * 1000) Mod 1000, "000")
End Sub


'---------------------------------------------------------------------------------------
' Procedure : PadRight
' Author    : Adam Waller
' Date      : 1/25/2019
' Purpose   : Pad a string on the right to make it `count` characters long.
'---------------------------------------------------------------------------------------
'
Public Sub PadRight(strText As String, Optional blnPrint As Boolean = True, Optional blnNextOutputOnNewLine As Boolean = False, Optional ByVal intCharacters As Integer)
    If intCharacters = 0 Then intCharacters = Me.PadLength
    If Len(strText) < intCharacters Then
        Me.Add strText & Space$(intCharacters - Len(strText)), blnPrint, blnNextOutputOnNewLine
    Else
        Me.Add strText, blnPrint, blnNextOutputOnNewLine
    End If
End Sub


'---------------------------------------------------------------------------------------
' Procedure : Increment
' Author    : Adam Waller
' Date      : 4/28/2020
' Purpose   : Increment the progress bar/output display.
'           : Calls are accumulated and only applied to the bar at most every
'           : INCREMENT_THROTTLE_SEC, so tight inner loops skip both the Perf
'           : wrap and the underlying ProgressBar.Increment cost on most calls.
'           : Any remaining pending increments are flushed by Log.Flush.
'---------------------------------------------------------------------------------------
'
Public Sub Increment()

    ' Ignore if we are not using the form
    If m_Prog Is Nothing Then Exit Sub

    ' Accumulate so we can apply the batched count when the throttle fires.
    m_dblPendingIncrements = m_dblPendingIncrements + 1

    ' Skip everything (including Perf wrapping) until enough time has passed.
    If Timer - m_sngLastIncrement < INCREMENT_THROTTLE_SEC Then Exit Sub

    ApplyPendingIncrements

End Sub


'---------------------------------------------------------------------------------------
' Procedure : ApplyPendingIncrements
' Author    : Adam Waller
' Date      : 4/24/2026
' Purpose   : Apply any accumulated increments to the progress bar in a single batched
'           : update. Called by the throttled Increment path and by Flush, so the bar
'           : still reaches 100% even when the final increments fall inside the
'           : throttle window.
'---------------------------------------------------------------------------------------
'
Private Sub ApplyPendingIncrements()

    If m_Prog Is Nothing Then Exit Sub
    If m_dblPendingIncrements = 0 Then Exit Sub

    Perf.OperationStart "Increment Progress"
    If Not m_blnProgressActive And m_blnConsoleDirty Then
        ' Flush any pending output
        With m_RichText
            Echo False
            .Value = ConsoleHtml
            If VerifyFocus(m_RichText) Then
                .SelStart = ClampInt(Len(Nz(.Value, vbNullString)))
            End If
            Echo True
        End With
        m_blnConsoleDirty = False
    End If
    Me.ProgressBar.Increment m_dblPendingIncrements

    ' Reset throttle state
    m_dblPendingIncrements = 0
    m_sngLastIncrement = Timer

    ' Status is now active
    m_blnProgressActive = True
    Perf.OperationEnd

End Sub


'---------------------------------------------------------------------------------------
' Procedure : Progress
' Author    : Adam Waller
' Date      : 1/26/2026
' Purpose   : Send progress update to console progress bar and MCP.
'---------------------------------------------------------------------------------------
'
Public Sub Progress(lngCurrent As Long, lngTotal As Long, _
    Optional strMessage As String = vbNullString)

    ' Update console progress bar (always)
    If Not m_Prog Is Nothing Then
        m_Prog.Max = lngTotal
        m_Prog.Value = lngCurrent
        m_Prog.UpdateDisplay (Not m_blnProgressActive)
        m_blnProgressActive = True
    End If

    ' Send to MCP with throttling (max every 250ms)
    If MCP.IsActive Then
        If Timer - m_sngLastMCPProgress > 0.25 Then
            MCP.PostCallback "progress", lngCurrent, lngTotal, strMessage
            m_sngLastMCPProgress = Timer
        End If
    End If

End Sub


'---------------------------------------------------------------------------------------
' Procedure : CleanupOldLogs
' Author    : Adam Waller
' Date      : 1/26/2026
' Purpose   : Delete oldest log files when count exceeds maximum.
'---------------------------------------------------------------------------------------
'
Private Sub CleanupOldLogs(strFolder As String, strBase As String, lngMaxFiles As Long)

    Dim oFolder As Object
    Dim oFile As Object
    Dim colFiles As Collection

    ' 0 = keep all logs (no cleanup)
    If lngMaxFiles <= 0 Then Exit Sub

    Set colFiles = New Collection
    Set oFolder = FSO.GetFolder(strFolder)

    ' Collect matching archived log files
    ' (Timestamp naming means alphabetical sort = chronological sort)
    For Each oFile In oFolder.Files
        If oFile.Name Like strBase & "_*.log" Then
            colFiles.Add oFile.Path
        End If
    Next oFile

    ' Delete oldest files if over limit
    Do While colFiles.Count > lngMaxFiles
        FSO.DeleteFile colFiles(1)
        colFiles.Remove 1
    Loop

End Sub


'---------------------------------------------------------------------------------------
' Procedure : ConsoleHtml
' Author    : Adam Waller
' Date      : 5/21/2026
' Purpose   : Return the console buffer trimmed to a safe size, ensuring the result
'           : starts at a clean <br> boundary. A raw Right$ cut can split an HTML tag,
'           : producing malformed markup that causes the Rich Text control to silently
'           : stop rendering content before the end.
'---------------------------------------------------------------------------------------
'
Private Function ConsoleHtml() As String

    Const CONSOLE_LIMIT As Long = 64000
    Dim strRaw As String
    Dim lngPos As Long

    strRaw = m_Console.RightStr(CONSOLE_LIMIT)
    If m_Console.Length > CONSOLE_LIMIT Then
        ' Truncation occurred; find the first line break and start after it.
        lngPos = InStr(1, strRaw, "<br>")
        If lngPos > 0 Then
            ConsoleHtml = Mid$(strRaw, lngPos + 4)
        Else
            ConsoleHtml = strRaw
        End If
    Else
        ConsoleHtml = strRaw
    End If

End Function


'---------------------------------------------------------------------------------------
' Procedure : ClampInt
' Author    : Adam Waller
' Date      : 5/21/2026
' Purpose   : Cap a Long value at 32000 to avoid Overflow when assigning to Integer
'           : properties like TextBox.SelStart.
'---------------------------------------------------------------------------------------
'
Private Function ClampInt(lngValue As Long) As Long
    If lngValue > 32000 Then
        ClampInt = 32000
    Else
        ClampInt = lngValue
    End If
End Function


'---------------------------------------------------------------------------------------
' Procedure : AppendErrorJournal
' Author    : Adam Waller
' Date      : 5/29/2026
' Purpose   : Record a Log.Error call for per-test JSON export (TestResults loggedErrors).
'---------------------------------------------------------------------------------------
'
Private Sub AppendErrorJournal(eLevel As eErrorLevel, strMessage As String, _
    strSource As String, lngErrNum As Long, strErrDesc As String)

    Dim dEntry As Dictionary

    Set dEntry = New Dictionary
    dEntry.Add "level", ErrorLevelToJournalString(eLevel)
    dEntry.Add "message", strMessage
    dEntry.Add "source", Nz2(strSource, vbNullString)
    dEntry.Add "errNumber", lngErrNum
    dEntry.Add "errDescription", Nz2(strErrDesc, vbNullString)
    m_colErrorJournal.Add dEntry

End Sub


'---------------------------------------------------------------------------------------
' Procedure : ErrorLevelToJournalString
' Author    : Adam Waller
' Date      : 5/29/2026
' Purpose   : Serialize eErrorLevel for the error journal JSON field.
'---------------------------------------------------------------------------------------
'
Private Function ErrorLevelToJournalString(eLevel As eErrorLevel) As String
    Select Case eLevel
        Case eelWarning:  ErrorLevelToJournalString = "warning"
        Case eelError:    ErrorLevelToJournalString = "error"
        Case eelCritical: ErrorLevelToJournalString = "critical"
        Case Else:        ErrorLevelToJournalString = "unknown"
    End Select
End Function


'---------------------------------------------------------------------------------------
' Procedure : MigrateLegacyLogs
' Author    : Adam Waller
' Date      : 1/26/2026
' Purpose   : Migrate old-format log files (Export.log, Build.log, etc.) from the
'           : source folder root to the new logs/ subfolder with timestamp naming.
'---------------------------------------------------------------------------------------
'
Private Sub MigrateLegacyLogs(strSourceFolder As String, strLogsFolder As String)

    Dim varName As Variant
    Dim strOldPath As String
    Dim strNewPath As String
    Dim strTimestamp As String
    Dim oFile As Object

    ' Check for each legacy log file type
    For Each varName In Array("Export", "Build", "Merge", "Other")
        strOldPath = FSO.BuildPath(strSourceFolder, varName & ".log")
        If FSO.FileExists(strOldPath) Then
            ' Create logs folder if needed
            If Not FSO.FolderExists(strLogsFolder) Then
                FSO.CreateFolder strLogsFolder
            End If
            ' Get timestamp from file's last modified date
            Set oFile = FSO.GetFile(strOldPath)
            strTimestamp = Format$(oFile.DateLastModified, "yyyymmdd\_hhnnss") & "_000"
            ' Build new path: Export.log -> Export_20260115_093042_000.log
            strNewPath = FSO.BuildPath(strLogsFolder, varName & "_" & strTimestamp & ".log")
            ' Move file to new location
            FSO.MoveFile strOldPath, strNewPath
        End If
    Next varName

End Sub
