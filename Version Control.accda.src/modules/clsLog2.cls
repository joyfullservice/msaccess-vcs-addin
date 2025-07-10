VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsLog2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
'
' Note: much of this module was originally borrowed from the MSAccessVCS-Addin
' tool we also use (and contribute to).
' It has subsequently recieved numerous makeovers and changes to integrate
' with our use cases.
'
' Original Source: https://github.com/joyfullservice/msaccess-vcs-integration
' This has been heavily modified for our uses, and was re-contributed
' back into the upstream source (as accepted).
'
Private Const ModuleName As String = "clsLog2"

Private Const DefaultLogTableName As String = "tblErrorLog"

Private Const MAX_ENTRY_COUNT_PROMPT As Long = 15000
Private Const WIDTH_MAX_DEFAULT As Long = 200
Private Const SCREEN_ROWS_MAX_DEFAULT As Long = 30 ' Assumes a msgbox half screen height at 1920x1080 resolution, allowing for Source String to be added.

Private Const WIDTH_MAX_MSGBOX As Long = 350
Private Const WIDTH_DESIRED_MSGBOX As Long = 200
Private Const MSGBOX_SECTION_ADDER As Long = 1
Private Const MSGBOX_BOLD_ROWS As Long = 20
Private Const MSGBOX_LINE1_ROWS As Long = 10
Private Const MSGBOX_LINE2_ROWS As Long = 10
Private Const MSGBOX_SOURCE_ROWS As Long = 5

Private Const cstrSpacer As String = "-------------------------------------"

Private m_Console As clsConcat                   ' Console output
Private m_RichText As TextBox                    ' Text box to display HTML
Private m_Prog As clsLblProg                     ' Progress bar

Private m_LogTableName As String

Private m_LogadoStoredConnection As ADODB.Connection
Private m_adoStoredCMDAddEntry As ADODB.Command
Private m_sngLastUpdate As Single

' Log level Everything below this (except for "no error")
' is not logged; for now, this will be determined by the devs
' but it is a preference. It's read only because it's read during
' init and stored for faster processing; this will avoid delays
' if user preferences expire, but the user hasn't refreshed the application.
Private m_LogLevel As eErrorLevel

Private Type udtThis
    SourcePath As String

End Type
Private this As udtThis

'---------------------------------------------------------------------------------------
' Procedure : Class_Initialize
' Author    : Adam Waller
' Date      : 4/28/2020
' Purpose   : Set initial options
'---------------------------------------------------------------------------------------
'
Private Sub Class_Initialize()

    Set m_Console = New clsConcat
    m_sngLastUpdate = 0
    'SetupLogTable

End Sub


'---------------------------------------------------------------------------------------
' Procedure : Clear
' Author    : Adam Waller, hecon5
' Date      : 4/16/2020
' Purpose   : Clear the log buffers
'---------------------------------------------------------------------------------------
'
Public Sub Clear(Optional ByVal CheckFirst As Boolean = True)

    Const FunctionName As String = ModuleName & ".Clear"

    If Not CheckFirst Then GoTo ClearLog
    Select Case Me.Prompt(strSource:=FunctionName _
                        , strBold:="Clear the log? There are " & Me.EntryCount & " entries in the log." _
                        , strLine1:="Save and Export log before clearing?" _
                        , strLine2:="If you have not experienced any issues, you should not export the log." & vbNewLine & _
                                    "[YES] = Save before clearing Log." & vbNewLine & _
                                    "[NO] = Do not save before clearing Log." & vbNewLine & _
                                    "[Cancel] = Cancel, Log will not be cleared" _
                        , vbButtons:=vbYesNoCancel + vbQuestion + vbDefaultButton2 _
                        , strTitle:="Save before clearing?" _
                        , ErrorLevel:=eErrorLevel.eelNoError)

        Case vbYes
            ' Save then Clear log.
            Me.SaveFile , CheckFirst

        Case vbNo
            ' Don't save, but still delete

        Case Else
            Exit Sub
    End Select

ClearLog:
    ExecuteLocalSQL "DELETE FROM " & Me.LogTableName & ";"
    Class_Initialize

End Sub


Public Sub ClearConsole()
    m_Console.Clear
    Me.Flush
End Sub


Public Function IsEntryLogged(ByRef ErrorLevel As eErrorLevel) As Boolean

'    If (EventErrorLevel And Me.LogLevel) = EventErrorLevel _
'        And (EventErrorLevel >= Me.LogLevel) _
'        or  (me.LogLevel <>
'        Then IsEntryLogged = True
    If (ErrorLevel >= Me.LogLevel) Or ErrorLevel = eelNoError Then IsEntryLogged = True

End Function


'---------------------------------------------------------------------------------------
' Procedure : Spacer
' Author    : Adam Waller
' Date      : 4/28/2020
' Purpose   : Add a spacer to the log
'---------------------------------------------------------------------------------------
'
Public Sub Spacer(Optional ByVal blnPrint As Boolean = True)
    AddEntry vbNullString, cstrSpacer, eelError, blnPrint, True
End Sub


Public Sub AddEntry(ByRef strMessage As String _
                    , Optional ByRef strSource As String _
                    , Optional ByVal ErrorLevel As eErrorLevel = eelNoError _
                    , Optional ByVal PrintEvent As Boolean = True _
                    , Optional ByVal NextOutputOnNewLine As Boolean = True _
                    , Optional ByRef OutputColor As String = vbNullString _
                    , Optional ByVal PrintBold As Boolean = False _
                    , Optional ByVal PrintItalic As Boolean = False)

    Const FunctionName As String = ModuleName & ".AddEntry"

    Dim adoCMD As ADODB.Command
    Dim CachedErrorNumber As Long

    CachedErrorNumber = VBA.Err.Number
    On Error Resume Next
    Perf.OperationStart FunctionName

    If Not IsEntryLogged(ErrorLevel) Then GoTo Exit_Here
    If PrintEvent Then
        m_Console.Add FormatEntry(strMessage, NextOutputOnNewLine, OutputColor, PrintBold, PrintItalic)

        ' Allow an update to the screen every second.
        ' (This keeps the aplication from an apparent hang while
        '  running intensive export processes.)
        If m_sngLastUpdate + 0.1 < Timer Then Me.Flush
    End If

    Set adoCMD = GetLogStoredCommand
    With adoCMD
        If (.ActiveConnection.State And adStateOpen) <> adStateOpen Then .ActiveConnection.Open
        .Parameters("[paramEventLogTime]") = ISO8601TimeStamp
        .Parameters("[paramUserID]") = GetUserID
        .Parameters("[paramComputerName]") = GetUserComputerName
        .Parameters("[paramErrorNumber]") = CachedErrorNumber
        .Parameters("[paramEventSource]") = strSource
        .Parameters("[paramEventMessage]") = strMessage
        .Parameters("[paramPrinted]") = Not PrintEvent
        .Parameters("[paramErrorLevel]") = ErrorLevel

Retry:
        .Execute
        If Err Then
        ' Because we're in the log class, we can't rely on the catch
        ' or other tools because everything logs to here...
            If Err.Number = -2147217865 Then
                ' The log table got deleted somehow.
                Err.Clear
                'SetupLogTable
                GoTo Retry
            End If
        End If
    End With

Exit_Here:
    Perf.OperationEnd

End Sub


Private Function GetUserID() As String
    ' Gets the user's ID from the local machine.
    Static UserIDCache As String
    Dim objNetwork As Object

    If UserIDCache = vbNullString Then
        Set objNetwork = CreateObject("WScript.Network")
        UserIDCache = objNetwork.UserName
        Set objNetwork = Nothing
    End If
    GetUserID = UserIDCache

End Function


Private Function GetUserComputerName() As String

    Static UserComputerNameCache As String
    Dim objNetwork As Object

    If UserComputerNameCache = vbNullString Then
        Set objNetwork = CreateObject("WScript.Network")
        UserComputerNameCache = objNetwork.ComputerName
        Set objNetwork = Nothing
    End If
    GetUserComputerName = UserComputerNameCache

End Function


Private Function GetLogStoredCommand() As ADODB.Command

    Dim SQLString As clsConcat

    If m_adoStoredCMDAddEntry Is Nothing Then
        Set SQLString = New clsConcat
        With SQLString
            .Add "PARAMETERS "
            .Add "[paramEventLogTime] VARCHAR(40)"
            .Add ",[paramUserID] VARCHAR(40)"
            .Add ",[paramComputerName] VARCHAR(255)"
            .Add ",[paramErrorNumber] LONG"
            .Add ",[paramEventSource] LONGTEXT"
            .Add ",[paramEventMessage] LONGTEXT"
            .Add ",[paramPrinted] BIT"
            .Add ",[paramErrorLevel] LONG"
            .Add "; "
            .Add "INSERT INTO ", Me.LogTableName, " ("
            .Add "[EventLogTime]"
            .Add ",[UserID]"
            .Add ",[ComputerName]"
            .Add ",[ErrorNumber]"
            .Add ",[EventSource]"
            .Add ",[EventMessage]"
            .Add ",[Printed]"
            .Add ",[ErrorLevel]"
            .Add ") VALUES ("
            .Add "[paramEventLogTime]"
            .Add ",[paramUserID]"
            .Add ",[paramComputerName]"
            .Add ",[paramErrorNumber]"
            .Add ",[paramEventSource]"
            .Add ",[paramEventMessage]"
            .Add ",[paramPrinted]"
            .Add ",[paramErrorLevel]"
            .Add ");"
        End With

        Set m_adoStoredCMDAddEntry = New ADODB.Command
        With m_adoStoredCMDAddEntry
            Set .ActiveConnection = Me.LogConnectionADO
            .CommandText = SQLString.GetStr
            .CommandType = adCmdText
            .Prepared = True
            .NamedParameters = True
            .Parameters.Append .CreateParameter("[paramEventLogTime]", adVarChar, adParamInput, 40)
            .Parameters.Append .CreateParameter("[paramUserID]", adVarChar, adParamInput, 40)
            .Parameters.Append .CreateParameter("[paramComputerName]", adVarChar, adParamInput, 255)
            .Parameters.Append .CreateParameter("[paramErrorNumber]", adInteger, adParamInput)
            .Parameters.Append .CreateParameter("[paramEventSource]", adLongVarChar, adParamInput, 5000)
            .Parameters.Append .CreateParameter("[paramEventMessage]", adLongVarChar, adParamInput, 63999)
            .Parameters.Append .CreateParameter("[paramPrinted]", adBoolean, adParamInput)
            .Parameters.Append .CreateParameter("[paramErrorLevel]", adInteger, adParamInput)
        End With
    End If

    With m_adoStoredCMDAddEntry
        If .ActiveConnection Is Nothing Then Set .ActiveConnection = Me.LogConnectionADO
        If (.ActiveConnection.State And adStateOpen) <> adStateOpen Then .ActiveConnection.Open
    End With

    Set GetLogStoredCommand = m_adoStoredCMDAddEntry

End Function


Private Function FormatEntry(ByRef MessageToFormat As String _
                            , Optional ByVal NextOutputOnNewLine As Boolean = True _
                            , Optional ByVal OutputColor As String = vbNullString _
                            , Optional ByVal PrintBold As Boolean = False _
                            , Optional ByVal PrintItalic As Boolean) As String

    With New clsConcat
        ' Opening tags
        If PrintBold Then .Add "<strong>"
        If PrintItalic Then .Add "<i>"
        If OutputColor <> vbNullString Then .Add "<font color=", OutputColor, ">"

        ' Content
        .Add MultiReplace(MessageToFormat, _
                          " ", "&nbsp;", _
                          vbNewLine, "<br>")
        ' Closing tags
        If OutputColor <> vbNullString Then .Add "</font>"
        If PrintItalic Then .Add "</i>"
        If PrintBold Then .Add "</strong>"

        ' Add line break for HTML
        If NextOutputOnNewLine Then .Add "<br>"
        ' Now build string
        FormatEntry = .GetStr
    End With

End Function


'---------------------------------------------------------------------------------------
' Procedure : Flush
' Author    : Adam Waller, hecon5
' Date      : 4/29/2020
' Purpose   : Flushes the buffer to the console
'---------------------------------------------------------------------------------------
'
Public Sub Flush()

    ' See if the GUI form is loaded.
    Perf.OperationStart "Console Updates"
    If Not m_RichText Is Nothing Then
        VerifyFocus m_RichText
        With m_RichText
            ' Set value, not text to avoid errors with large text strings.
            Echo False
            ' Show the last 20K characters so
            ' we don't hit the Integer limit
            ' on the SelStart property.
            .Value = m_Console.RightStr(20000)
            If VerifyFocus(m_RichText, True) Then .SelStart = Len(.Text & vbNullString)
            Echo True
        End With
    End If
    m_sngLastUpdate = Timer
    ' Update the display (especially for immediate window)
    DoEvents
    Perf.OperationEnd

End Sub


'---------------------------------------------------------------------------------------
' Procedure : Error
' Author    : Adam Waller, hecon5, heavily modified from original
' Date      : 12/4/2020
' Purpose   : Log an error, and update error level if needed. Reads Err object values.
'           : A critical error will also present a message box with the details.
'---------------------------------------------------------------------------------------
'
Public Function Error(ByRef strBold As String _
                    , ByRef strSource As String _
                    , ByRef ErrorLevel As eErrorLevel _
                    , Optional ByRef strErrDescription As String _
                    , Optional ByRef vbButtons As VbMsgBoxStyle = vbOKOnly _
                    , Optional ByRef strTitle As String = vbNullString _
                    , Optional ByRef moreDetails As String = vbNullString) As VbMsgBoxResult

    Const FunctionName As String = ModuleName & ".Error"

    Dim ErrorResponse As VbMsgBoxResult
    Dim errorBold As New clsConcat
    Dim errorFullDetails As New clsConcat
    Dim errorLevelInfo As New clsErrorInfo
    Dim sourceFormatted As String
    Dim intButtons As VbMsgBoxStyle
    Dim strWorking As New clsConcat

    Error = vbCancel

    Perf.OperationStart FunctionName
    errorLevelInfo.ErrorLevel = ErrorLevel
    sourceFormatted = strSource & ".Error"

    With strWorking
        .Add errorLevelInfo.Prefix
        If Err Then .Add "Number: ", Err.Number, " ", Err.Description, " ", Err.Source
        .Add vbNewLine, strErrDescription
    End With

    With errorBold
        .Add cstrSpacer, vbNewLine
        .Add strWorking.GetStr, vbNewLine
        .Add cstrSpacer
    End With

    strWorking.Clear

    With errorFullDetails
        If Err Then
            .Add vbNewLine, "Error Number: ", Err.Number, " ", Err.Description
            .Add vbNewLine, "Help Context: ", Err.HelpContext
            .Add vbNewLine, "HelpFile: ", Err.HelpFile
            .Add vbNewLine, "LastDllError: ", Err.LastDllError
            .Add vbNewLine, "Source: ", Err.Source
        End If
        .Add vbNewLine, "CallStack: ", Perf.CallStack
    End With

    ' Or performs a bitwise operation here
    intButtons = vbButtons Or errorLevelInfo.MsgBoxStyle
    ' Build the error message string.

    ' Show message on any error level when we are not logging to a file.
    Select Case ErrorLevel
        Case eelWarning, eelError, eelCritical
            ErrorResponse = Me.Prompt(strSource:=sourceFormatted _
                                    , strBold:=errorBold.GetStr _
                                    , strLine1:=errorFullDetails.GetStr _
                                    , strLine2:=moreDetails _
                                    , ErrorLevel:=ErrorLevel _
                                    , vbButtons:=vbButtons _
                                    , strTitle:=strTitle)

            If ErrorLevel = eelCritical Then
                Me.SaveFile ' Performance reports are logged during save routines.
            Else
                LogPerformanceReports
            End If

    Case Else
        ' Log the error only if user isn't prompted; the prompt logs it for you
        With strWorking
            .Add FormatEntry(errorBold.GetStr, True, errorLevelInfo.ErrorColor, True)
            If errorFullDetails.Length Then .Add vbNewLine, errorFullDetails.GetStr, vbNewLine
            If moreDetails <> vbNullString Then .Add vbNewLine, "More Details: ", moreDetails
            If intButtons <> vbOKOnly Then .Add vbNewLine, "Buttons: ", CStr(intButtons)
            If strTitle <> vbNullString Then .Add vbNewLine, "Title: ", strTitle
        End With

        Me.AddEntry strSource:=sourceFormatted _
                    , strMessage:=strWorking.GetStr _
                    , ErrorLevel:=ErrorLevel _
                    , PrintEvent:=True _
                    , NextOutputOnNewLine:=True _
                    , PrintBold:=True
    End Select
    'Err.Clear ' Clear the error, it's been logged and dealt with now.
    Perf.OperationEnd
    Error = ErrorResponse

End Function


Public Function Prompt(ByRef strBold As String _
                    , ByRef strSource As String _
                    , Optional ByRef ErrorLevel As eErrorLevel = eErrorLevel.eelNoError _
                    , Optional ByRef strLine1 As String = vbNullString _
                    , Optional ByRef strLine2 As String = vbNullString _
                    , Optional ByRef vbButtons As VbMsgBoxStyle = vbOKOnly _
                    , Optional ByRef strTitle As String = vbNullString _
                    , Optional ByVal AddToLog As Boolean = True _
                    , Optional ByVal vbDefaultResult As VbMsgBoxResult = vbOK) As VbMsgBoxResult

    ' This function will prompt the user to do something, or display a message, and return the value to the calling function.
    ' In addition, it will log that the prompt happened, and the result of the user action.

    Const FunctionName As String = ModuleName & ".Prompt"

    Dim intCursor As Integer
    Dim PromptResult As VbMsgBoxResult
    Dim sourceFormatted As String
    Dim titleFormatted As String
    Dim errorLevelInfo As New clsErrorInfo

    Dim strBoldLines As String
    Dim strBoxLines1 As String
    Dim msgBoxStr2 As New clsConcat

    Dim logEntryDetails As New clsConcat
    Dim LineCount As Long

    Perf.OperationStart FunctionName

    ' Turn off any hourglass
    intCursor = Screen.MousePointer

    sourceFormatted = strSource & ".Prompt"
    titleFormatted = Nz2(strTitle, sourceFormatted)
    errorLevelInfo.ErrorLevel = ErrorLevel

    LineCount = SCREEN_ROWS_MAX_DEFAULT
    strBoldLines = FitStringToWidth(strBold, WIDTH_MAX_MSGBOX, WIDTH_DESIRED_MSGBOX, LineCount)

    If strLine1 <> vbNullString Then
        LineCount = LineCount + MSGBOX_LINE1_ROWS
        strBoxLines1 = FitStringToWidth(strLine1, WIDTH_MAX_MSGBOX, WIDTH_DESIRED_MSGBOX, LineCount)
    End If

    With msgBoxStr2
        If strLine2 <> vbNullString Then .Add FitStringToWidth(strLine2, WIDTH_MAX_MSGBOX, WIDTH_DESIRED_MSGBOX, LineCount + MSGBOX_LINE2_ROWS)
        If strSource <> vbNullString Then
            If .Length Then .Add vbNewLine, vbNewLine
            .Add "Source: ", FitStringToWidth(sourceFormatted, WIDTH_MAX_MSGBOX - 8, WIDTH_DESIRED_MSGBOX - 8, MSGBOX_SOURCE_ROWS)
        End If
    End With

    Perf.PauseTiming

    ' Check interaction mode
    If InteractionMode = eimNormal Then
        PromptResult = MsgBox3(strBoldLines _
                            , strBoxLines1 _
                            , msgBoxStr2.GetStr _
                            , vbButtons, titleFormatted)
    Else
        ' Just log the output
        ' https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/msgbox-function
        ' The enum for VbMsgBoxResult does not have a zero value and we need a way to tell the system user hasn't said anything.
        PromptResult = 0
    End If

    Perf.ResumeTiming
    If AddToLog Then
        With logEntryDetails
            .AppendOnAdd = "<br>"
            .Add cstrSpacer
            If InteractionMode = eimSilent Then .Add "[**MessageBox Not Displayed**]"
            .Add "Prompt:"
            .Add FormatEntry(strBold, True, errorLevelInfo.ErrorColor, True)
            .Add strLine1
            .Add strLine2
            .Add "Source: ", sourceFormatted
            If vbButtons <> vbOKOnly Then .Add "Buttons: ", CStr(vbButtons)
            If strTitle <> vbNullString Then .Add "Title: ", titleFormatted
            .Add FormatEntry("User Response: " & PromptResult, True, , True)
            .Add cstrSpacer
            .Remove 4 ' Remove the last line feed.
        End With
        Me.AddEntry strSource:=sourceFormatted, strMessage:=logEntryDetails.GetStr, ErrorLevel:=ErrorLevel
    End If

    Prompt = PromptResult

    ' Restore MousePointer (if needed)
    If intCursor > 0 Then Screen.MousePointer = intCursor

    Perf.OperationEnd

End Function


'---------------------------------------------------------------------------------------
' Procedure : MsgBox3
' Author    : Adam Waller
' Date      : 1/27/2017
' Purpose   : Alternate message box with bold prompt on first line. Renamed so MSGBOX2
'           : can be wrapped and logged.
'---------------------------------------------------------------------------------------
'
Private Function MsgBox3(ByRef strBold As String _
                        , Optional ByRef strLine1 As String = vbNullString _
                        , Optional ByRef strLine2 As String = vbNullString _
                        , Optional ByRef intButtons As VbMsgBoxStyle = vbOKOnly _
                        , Optional ByRef strTitle As String = vbNullString) As VbMsgBoxResult

    Dim strMsg As String
    Dim varLines(0 To 3) As String

    ' Escape single quotes by doubling them.
    varLines(0) = Replace(strBold, "'", "''")
    varLines(1) = Replace(strLine1, "'", "''")
    varLines(2) = Replace(strLine2, "'", "''")
    varLines(3) = Replace(strTitle, "'", "''")

    If varLines(3) = vbNullString Then varLines(3) = Application.VBE.ActiveVBProject.Name
    strMsg = "MsgBox('" & varLines(0) & "@" & varLines(1) & "@" & varLines(2) & "@'," & intButtons & ",'" & varLines(3) & "')"
    MsgBox3 = Eval(strMsg)

End Function


'---------------------------------------------------------------------------------------
' Procedure : SetConsole
' Author    : Adam Waller
' Date      : 4/28/2020
' Purpose   : Set a reference to the text box.
'---------------------------------------------------------------------------------------
'
Public Sub SetConsole(ByRef txtRichText As TextBox)
    Set m_RichText = txtRichText
    If Not m_RichText Is Nothing Then
        m_RichText.AllowAutoCorrect = False
    End If
End Sub


'---------------------------------------------------------------------------------------
' Procedure : LogFilePath
' Author    : Adam Waller
' Date      : 5/13/2023
' Purpose   : Return the (default) full path to the current log file.
'---------------------------------------------------------------------------------------
'
Public Property Get LogFilePath() As String
    Dim strFile As String
'    Select Case Operation.OperationType
'        Case eotExport: strFile = "Export.log"
'        Case eotBuild:  strFile = "Build.log"
'        Case eotMerge:  strFile = "Merge.log"
'        Case Else:      strFile = "Other.log"
'    End Select
    LogFilePath = FSO.BuildPath(Nz2(this.SourcePath, Options.GetExportFolder), strFile)
End Property


Public Property Get LogTableName() As String

    If m_LogTableName = vbNullString Then m_LogTableName = DefaultLogTableName
    LogTableName = m_LogTableName

End Property
Public Property Let LogTableName(ByRef NewLogTableName As String)

    If Not TableExists(NewLogTableName) Then
        m_LogTableName = NewLogTableName
        Class_Initialize
   End If

End Property


Public Sub CheckSize()
    ' See if we should clear out the old log entries.
    If Me.EntryCount > MAX_ENTRY_COUNT_PROMPT Then Me.Clear True
End Sub


Public Property Get EntryCount() As Long

    Perf.OperationStart ModuleName & ".EntryCount"
    EntryCount = DCount("[EntryId]", Me.LogTableName)
    Perf.OperationEnd

End Property


'---------------------------------------------------------------------------------------
' Procedure : ProgressBar
' Author    : Adam Waller
' Date      : 11/6/2020
' Purpose   : Pass the Progress Bar reference to this class.
'---------------------------------------------------------------------------------------
'
Public Property Set ProgressBar(ByRef cProg As clsLblProg)
    Set m_Prog = cProg
End Property
Public Property Get ProgressBar() As clsLblProg
    If m_Prog Is Nothing Then Set m_Prog = New clsLblProg
    Set ProgressBar = m_Prog
End Property


Public Sub LogPerformanceReports(Optional ByRef PrintEvent As Boolean = True)
    Me.AddEntry strSource:="Performance Reports" _
                , strMessage:=FormatEntry(Perf.GetReports) _
                , ErrorLevel:=eelDebugInfo _
                , PrintEvent:=PrintEvent _
                , PrintBold:=True
End Sub


Public Sub SaveFile(Optional strAlternateFileName As String _
                    , Optional ByRef NotQuiet As Boolean = False)

    Const FunctionName As String = ModuleName & ".SaveFile"
    Dim OutputFilePath As String
    Dim ReportTimeStamp As String
    Dim SanitizedFileName As String
    Dim ShellCmdString As String

Retry:
    If (Me.Prompt(strSource:=FunctionName _
                    , strBold:="Saving and exporting the log." _
                    , strLine1:="Press [OK] to continue." & vbNewLine & _
                                "Press [Cancel] to abort export." _
                    , vbButtons:=vbOKCancel + vbDefaultButton1) <> vbOK) Then Exit Sub

    ' Before exporting, insert entries for performance, because that might be useful.
    ' You are, after all, exporting the log to troubleshoot...

    LogPerformanceReports False

    ReportTimeStamp = ISO8601TimeStamp(False, False)
    SanitizedFileName = Nz2(strAlternateFileName, "ErrorLog") & "." & ReportTimeStamp & ".log.xml"
    SanitizedFileName = MultiReplace(SanitizedFileName, ":", "-") ' Replace Timestamp charachters with valid charachters.
    OutputFilePath = FSO.BuildPath(Me.LogFilePath, SanitizedFileName)
    VerifyPath OutputFilePath

    Application.ExportXML acExportTable, Me.LogTableName, OutputFilePath



    With New clsSourceParser
        .LoadSourceFile OutputFilePath, edbTableData
        WriteFile .Sanitize(ectXML), OutputFilePath
    End With

    If Err.Number <> 0 Then
        If Me.Prompt(strSource:=FunctionName, strBold:="Error saving file." _
                    , strLine1:="Retry?" _
                    , vbButtons:=vbYesNo, ErrorLevel:=eelError) = vbYes Then GoTo Retry

    Else
        If NotQuiet Then
            Me.Prompt strSource:=FunctionName, strBold:="Successfully saved log!" _
                    , strLine1:="Saved to: " & Me.LogFilePath _
                    , strLine2:="File Name: " & SanitizedFileName _
                    , strTitle:="Successfully saved log!"

            If FSO.FileExists(OutputFilePath) Then
                ShellCmdString = "explorer.exe /select, """ & OutputFilePath & """"
                Shell ShellCmdString, vbNormalFocus
            End If
        End If
    End If

End Sub


'Private Sub SetupLogTable()
'
'ReRun:
'    SetupLocalTable LogTableName, eFieldTypeNotMatchDelete _
'                    , "EntryID", "AUTOINCREMENT" _
'                    , "EventLogTime", "VARCHAR(40)" _
'                    , "UserID", "VARCHAR(40)" _
'                    , "ComputerName", "VARCHAR(255)" _
'                    , "ErrorNumber", "Long" _
'                    , "EventSource", "LongText" _
'                    , "EventMessage", "LongText" _
'                    , "Printed", "Bit" _
'                    , "ErrorLevel", "Long"
'                   '
''' Note: the later portions of this were originally to change the EventMessage field to RichText.
''' I'm leaving it here because it might be useful later.
''    Dim Db As DAO.Database
''    Dim LogTable As DAO.TableDef
''    Dim fld As DAO.field
''    Dim logProperty As Property
'
''    'Now set MessageText to rich text, because you can't do it in SQL...
''    Set Db = CurrentDb
''    Set LogTable = Db.TableDefs(Me.LogTableName)
''    Set fld = LogTable.Fields("EventMessage")
''Recheck:
''    With LogTable.Fields("EventMessage").Properties("TextFormat")
''        If Catch(3270) Then GoTo AddProperty
''        If .Value = acTextFormatPlain Then .Value = acTextFormatHTMLRichText
''    End With
''    'Normally , we 'd log this. but since we're trying to setup the log, don't.
''    If CatchAny(eelerror, "Setting up Log table issue.", ModuleName & ".SetupLogTable", False, False) Then
''        MsgBox3 "ERROR: Setting up Log table issue.", Err.Number, ModuleName & ".SetupLogTable"
''        Err.Clear
''    End If
'
'    Exit Sub
'
''AddProperty:
''    Set logProperty = fld.CreateProperty("TextFormat", dbInteger, acTextFormatHTMLRichText)
''    fld.Properties.Append logProperty
''    GoTo Recheck
'
'End Sub

'Public Function SetupLocalTable(ByRef TableName As String _
'                                , ByRef DeleteExist As eTableRefreshMethod _
'                                , ParamArray cNameTypePairs() As Variant) As String
'
'    Const FunctionName As String = ModuleName & ".SetupLocalTable"
'    Dim SQLString As String
'    Dim intPairs As Long
'    Dim FieldList As New clsConcat
'
'    'LogUnhandledErrors FunctionName
'   ' On Error GoTo ERR_Handle
'
''    If TableExists(TableName) Then
''        Select Case DeleteExist
''        Case eNoDelete
''           Exit Function
''        Case eFieldNotMatchDelete
''            If FieldsMatch(TableName, False, CVar(cNameTypePairs)) Then Exit Function
''        Case eFieldTypeNotMatchDelete
''            If FieldsMatch(TableName, True, CVar(cNameTypePairs)) Then Exit Function
''        Case eDeleteData
''            'Nothing extra
''        End Select
''        'Close the table and delete it.
''        DoCmd.Close acTable, TableName, acSaveNo
''        DoCmd.DeleteObject acTable, TableName
''    End If
'
'    For intPairs = 0 To UBound(cNameTypePairs) Step 2
'        FieldList.Add "[", CStr(cNameTypePairs(intPairs)), "] ", CStr(cNameTypePairs(intPairs + 1)), vbNewLine, ", "
'    Next intPairs
'    ' Trim last comma and space
'    FieldList.Remove 2
'
'    SQLString = "CREATE TABLE " & TableName & " (" & FieldList.GetStr & "); "
'    SQLString = MultiReplace(SQLString, _
'                                        vbNewLine, vbNullString)
'    ExecuteLocalSQL SQLString
'    SetupLocalTable = TableName
'    Exit Function
'
'ERR_Handle:
'    CatchAny eelError, "Cannot create table: " & TableName & ". SQL: " & SQLString, FunctionName
'    Resume Next
'End Function


Public Property Let LogLevel(ByRef NewLogLevel As eErrorLevel)
    m_LogLevel = NewLogLevel
End Property
Public Property Get LogLevel() As eErrorLevel
    LogLevel = m_LogLevel
End Property


'---------------------------------------------------------------------------------------
' Procedure : Increment
' Author    : Adam Waller, hecon5
' Date      : 4/28/2020
' Purpose   : Increment the progress bar
'---------------------------------------------------------------------------------------
'
Public Sub Increment()

    ' Converted to wrapper
    If Me.ProgressBar Is Nothing Then Exit Sub
    Me.ProgressBar.Increment

End Sub


' This is only here because we want to remove external dependancies.
Private Sub ExecuteLocalSQL(ByRef SQLString As String)

    Dim Db As DAO.Database

    Set Db = CurrentDb
    Db.Execute SQLString
    Set Db = Nothing

End Sub


Public Property Get LogConnectionADO() As ADODB.Connection

    If m_LogadoStoredConnection Is Nothing Then Set m_LogadoStoredConnection = GetDefaultLogConnectionADO
    Set LogConnectionADO = m_LogadoStoredConnection

End Property
Public Property Let LogConnectionADO(NewVal As ADODB.Connection)

    If Not NewVal Is Nothing Then
        Set m_LogadoStoredConnection = NewVal
    Else
        Set m_LogadoStoredConnection = GetDefaultLogConnectionADO
    End If

End Property


Private Function GetDefaultLogConnectionADO() As ADODB.Connection

        Set GetDefaultLogConnectionADO = New ADODB.Connection
        Set GetDefaultLogConnectionADO = CurrentProject.Connection ' Or CurrentDB.Connection, depending on what your goal is here

End Function


Private Function FitStringToWidth(ByRef LongString As String _
                                , Optional ByVal MaxWidth As Long = 200 _
                                , Optional ByVal DesiredWidth As Long = 75 _
                                , Optional ByRef MaxLines As Long = 0 _
                                , Optional ByVal RemoveBlankLines As Boolean = True) As String

    Const FunctionName As String = ModuleName & ".FitStringToWidth"

    ' Fits a string to a message box if it's wider than MaxWidth and ensures the strings won't
    ' extend past the bottom of the screen.
    ' MaxLines is returned with the remaining lines
    '     Example: MaxLines starts at 10, output string is 3 lines long, MaxLines is set to 7 when finished.

    Dim OutputConcat As clsConcat
    Dim StrPosition As Long
    Dim StrLen As Long ' Length of total string
    Dim NewLineCount As Long ' Number of newlines
    Dim OriginalLineCount As Long ' Number of lines now.
    Dim ArrPosition As Long
    Dim StrArrLen As Long ' Length of substring
    Dim StringArr() As String

    Perf.OperationStart FunctionName
    StrLen = Len(LongString)
    StringArr = Split(LongString, vbNewLine, , vbTextCompare)
    OriginalLineCount = UBound(StringArr) - LBound(StringArr)

    If StrLen > MaxWidth Or OriginalLineCount > MaxLines Then
        Perf.OperationStart "FitStringToWidth.Resize"
        Set OutputConcat = New clsConcat
        For ArrPosition = 0 To OriginalLineCount
            StrPosition = 1

            StrArrLen = Len(StringArr(ArrPosition))
            If ArrPosition > 0 And (StrArrLen > 0 And Not RemoveBlankLines) Then
                NewLineCount = NewLineCount + 1
                OutputConcat.Add vbNewLine
            End If
            Do While StrPosition < StrArrLen And NewLineCount <= MaxLines And MaxLines > 0
                If StrPosition > 1 Then OutputConcat.Add vbNewLine
                OutputConcat.Add Mid$(StringArr(ArrPosition), StrPosition, DesiredWidth)
                StrPosition = StrPosition + DesiredWidth
                NewLineCount = NewLineCount + 1
                If NewLineCount >= MaxLines And MaxLines > 0 Then GoTo StringLengthMaxedOut
            Loop
            If NewLineCount >= MaxLines And MaxLines > 0 Then GoTo StringLengthMaxedOut
        Next ArrPosition

StringLengthMaxedOut:
        MaxLines = MaxLines - NewLineCount ' Return the number of lines remaining.
        FitStringToWidth = OutputConcat.GetStr
        Perf.OperationEnd
    Else
        MaxLines = MaxLines - OriginalLineCount - 1 ' Return the number of lines remaining. Subtract one because OriginalLineCount is zero based.
        FitStringToWidth = LongString
    End If
    Perf.OperationEnd

End Function
