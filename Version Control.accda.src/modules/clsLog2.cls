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

' Because VBA doesn't allow use of private types in constants, we need to declare these as long.
Private Const DEFAULT_LogLevel As Long = eErrorLevel.eelDebugTrace
Private Const DEFAULT_DisplayLogLevel As Long = eErrorLevel.eelNoError

Private Const DEFAULT_PROMPT_VBRESULT As Long = vbOK

Private Const DefaultLogTableName As String = "tblLog"

Private Const MAX_ENTRY_COUNT_PROMPT As Long = 15000
Private Const SCREEN_ROWS_MAX_DEFAULT As Long = 30 ' Assumes a msgbox half screen height at 1920x1080 resolution, allowing for Source String to be added.

Private Const WIDTH_MAX_MSGBOX As Long = 350
Private Const WIDTH_DESIRED_MSGBOX As Long = 200
Private Const MSGBOX_LINE1_ROWS As Long = 10        ' If strBold exceeds SCREEN_ROWS_MAX_DEFAULT then we still want to show a minimum of context to user.
Private Const MSGBOX_LINE2_ROWS As Long = 10
Private Const MSGBOX_SOURCE_ROWS As Long = 2        ' We want to show the source and ensure they get to see something at a minimum.
Private Const cstrSpacer As String = "-------------------------------------"

Public SourcePath As String         ' Externally defined, used to tell the log where to store the log.

Private Type udtThis
    ' Log level: log everything at and below set level (except for "no error"
    ' which is logged unless log is set to off).
    LogLevel As eErrorLevel
    DisplayLogLevel As eErrorLevel  ' The level at which the console is displayed. Works the same as LogLevel.
    LogTableName As String          ' Name of the table used by the log class.

    sngLastUpdate As Single         ' Timer value at last console update.
    Console As clsConcat            ' Console output
    RichText As TextBox             ' Text box to display HTML
    Prog As clsLblProg              ' Progress bar

    LogadoStoredConnection As ADODB.Connection
    adoStoredCMDAddEntry As ADODB.Command

    StorePrivateDetails As Boolean  ' Store potentially private details such as User and Computer name
                                    ' NOTE: These can be very helpful in troubleshooting, but may reveal
                                    '       information which could be used nefariously.

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

    Dim emptyThis As udtThis

    Set this.adoStoredCMDAddEntry = Nothing
    Set this.LogadoStoredConnection = Nothing

    LSet this = emptyThis
    Set this.Console = New clsConcat
    this.sngLastUpdate = 0

End Sub


Private Sub Class_Terminate()

    ' close the connection to the external objects if not closed
    If Not this.adoStoredCMDAddEntry Is Nothing Then
        If this.adoStoredCMDAddEntry.State <> adStateClosed Then this.adoStoredCMDAddEntry.ActiveConnection.Close
        Set this.adoStoredCMDAddEntry = Nothing
    End If

    If Not this.LogadoStoredConnection Is Nothing Then
        If this.LogadoStoredConnection.State <> adStateClosed Then this.LogadoStoredConnection.Close
        Set this.LogadoStoredConnection = Nothing
    End If

End Sub


Public Property Let LogLevel(NewLogLevel As eErrorLevel)
    this.LogLevel = NewLogLevel
    If this.LogLevel = eelUninitialized Then this.LogLevel = DEFAULT_LogLevel
End Property
Public Property Get LogLevel() As eErrorLevel
    LogLevel = this.LogLevel
End Property


Public Property Let DisplayLogLevel(NewLogLevel As eErrorLevel)
    this.DisplayLogLevel = NewLogLevel
    If this.DisplayLogLevel = eelUninitialized Then this.DisplayLogLevel = DEFAULT_DisplayLogLevel
End Property
Public Property Get DisplayLogLevel() As eErrorLevel
    If this.DisplayLogLevel = eelUninitialized Then this.DisplayLogLevel = DEFAULT_DisplayLogLevel
    DisplayLogLevel = this.DisplayLogLevel
End Property


Public Property Let StorePrivateDetails(NewPropBool As Boolean)
    this.StorePrivateDetails = NewPropBool
End Property
Public Property Get StorePrivateDetails() As Boolean
    StorePrivateDetails = this.StorePrivateDetails
End Property


Public Function IsEntryLogged(ByVal ErrorLevelIn As eErrorLevel) As Boolean

    If ErrorLevelIn >= Me.LogLevel And Not Me.LogLevel = eelLoggingOff Then IsEntryLogged = True

End Function


Public Function IsEntryDisplayed(ByVal ErrorLevelIn As eErrorLevel) As Boolean

    If ErrorLevelIn >= Me.DisplayLogLevel And Not Me.DisplayLogLevel = eelLoggingOff Then IsEntryDisplayed = True

End Function

'---------------------------------------------------------------------------------------
' Procedure : Clear
' Author    : Adam Waller, hecon5
' Date      : 4/16/2020
' Purpose   : Clear the log buffers
'---------------------------------------------------------------------------------------
'
Public Sub Clear(Optional ByVal CheckFirst As Boolean = False _
                , Optional ByVal ClearConsoleAlso As Boolean = True)

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
                        , ErrorLevelIn:=eErrorLevel.eelNoError)

        Case vbYes
            ' Save then Clear log.
            Me.SaveFile , CheckFirst

        Case vbNo
            ' Don't save, but still delete

        Case Else
            Exit Sub
    End Select

ClearLog:
    With New ADODB.Command
        '.ActiveConnection = Me.LogConnectionADO
        '.CommandText = "DELETE FROM " & Me.LogTableName & ";"
        '.Execute
    End With
    If ClearConsoleAlso Then
        this.Console.Clear
        this.sngLastUpdate = 0
    End If

End Sub


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
    If Not this.RichText Is Nothing Then
        VerifyFocus this.RichText
        With this.RichText
            ' Set value, not text to avoid errors with large text strings.
            Echo False
            ' Show the last 20K characters so
            ' we don't hit the Integer limit
            ' on the SelStart property.
            .Value = this.Console.RightStr(20000)
            If VerifyFocus(this.RichText, True) Then .SelStart = Len(.Text & vbNullString)
            Echo True
        End With
    End If
    this.sngLastUpdate = Timer
    ' Update the display (especially for immediate window)
    DoEvents
    Perf.OperationEnd

End Sub


'---------------------------------------------------------------------------------------
' Procedure : Spacer
' Author    : Adam Waller
' Date      : 4/28/2020
' Purpose   : Add a spacer to the log
'---------------------------------------------------------------------------------------
'
Public Sub Spacer(Optional ByVal blnPrint As Boolean = True)
    AddEntry strMessage:=cstrSpacer, strSource:="Log.Spacer", ErrorLevelIn:=IIf(blnPrint, Me.DisplayLogLevel, Operation.ErrorLevel), NextOutputOnNewLine:=True
End Sub

'
'---------------------------------------------------------------------------------------
' Procedure : Add
' Author    : hecon5
' Date      : 4/28/2020
' Purpose   : Wrapper to ensure backward compatibility with prior log with minimal refactoring
'---------------------------------------------------------------------------------------
'
Public Sub Add(strText As String _
    , Optional blnPrint As Boolean = True _
    , Optional blnNextOutputOnNewLine As Boolean = True _
    , Optional strColor As String = vbNullString _
    , Optional blnBold As Boolean = False _
    , Optional blnItalic As Boolean = False)

    Dim f_ErrorInfo As clsErrorInfo

    Set f_ErrorInfo = New clsErrorInfo
    With f_ErrorInfo
        .ErrorLevel = IIf(blnPrint, eelNoError, eelDebugTrace)
        If strColor <> vbNullString Then .ErrorColor = strColor
        .FunctionSource = "Log.Add"
    End With

    SaveEntry strMessage:=strText _
            , ErrorInfoIn:=f_ErrorInfo _
            , NextOutputOnNewLine:=blnNextOutputOnNewLine _
            , PrintBold:=blnBold _
            , PrintItalic:=blnItalic

End Sub


'---------------------------------------------------------------------------------------
' Procedure : AddEntry
' Author    : hecon5
' Date      : 4/28/2020
' Purpose   : Prepares a log entry to be saved to the log.
'---------------------------------------------------------------------------------------
'
Public Sub AddEntry(ByRef strMessage As String _
                    , ByRef strSource As String _
                    , Optional ByVal ErrorLevelIn As eErrorLevel = eelNoError _
                    , Optional ByVal NextOutputOnNewLine As Boolean = True _
                    , Optional ByRef OutputColor As String = vbNullString _
                    , Optional ByVal PrintBold As Boolean = False _
                    , Optional ByVal PrintItalic As Boolean = False _
                    , Optional ByRef ErrorInfoIn As clsErrorInfo = Nothing)

    Const FunctionName As String = ModuleName & ".AddEntry"

    Dim adoCMD As ADODB.Command
    Dim fWasDisplayed As Boolean
    Dim f_ErrorInfo As clsErrorInfo

    Perf.OperationStart FunctionName
    If ErrorInfoIn Is Nothing Then
        Set f_ErrorInfo = New clsErrorInfo
    Else
        Set f_ErrorInfo = ErrorInfoIn
    End If

    With f_ErrorInfo
        .ErrorLevel = ErrorLevelIn
        .ErrorColor = OutputColor
        .FunctionSource = strSource
    End With

    On Error Resume Next


    SaveEntry strMessage:=strMessage _
            , ErrorInfoIn:=f_ErrorInfo _
            , NextOutputOnNewLine:=NextOutputOnNewLine _
            , PrintBold:=PrintBold _
            , PrintItalic:=PrintItalic

Exit_Here:
    Perf.OperationEnd

End Sub


Private Sub SaveEntry(ByRef strMessage As String _
                    , ByRef ErrorInfoIn As clsErrorInfo _
            , Optional ByVal NextOutputOnNewLine As Boolean = True _
            , Optional ByVal PrintBold As Boolean = False _
            , Optional ByVal PrintItalic As Boolean = False)

    Const FunctionName As String = ModuleName & ".SaveEntry"

    Dim adoCMD As ADODB.Command
    Dim fWasDisplayed As Boolean

    On Error Resume Next
    Perf.OperationStart FunctionName

    fWasDisplayed = IsEntryDisplayed(ErrorInfoIn.ErrorLevel)
    ' There may be cases where we want to print but not log the action.
    If fWasDisplayed Then

        If this.RichText Is Nothing Then
            ' Only print debug output if not running from the GUI.
            If NextOutputOnNewLine Then
                ' Create new line
                Debug.Print CleanHTMLFromStr(strMessage)
            Else
                ' Continue next printout on this line.
                Debug.Print CleanHTMLFromStr(strMessage);
            End If
        End If

        this.Console.Add FormatEntry(strMessage, NextOutputOnNewLine _
                                    , ErrorInfoIn.ErrorColor, PrintBold, PrintItalic)

        ' Allow an update to the screen every second.
        ' (This keeps the aplication from an apparent hang while
        ' running intensive export processes.)
        If this.sngLastUpdate + 0.1 < Timer Then Me.Flush
    End If

    If Not IsEntryLogged(ErrorInfoIn.ErrorLevel) Then GoTo Exit_Here

    Set adoCMD = GetLogStoredCommand
    With adoCMD
        If (.ActiveConnection.State And adStateOpen) <> adStateOpen Then .ActiveConnection.Open
        .Parameters("[paramEventLogTime]") = ErrorInfoIn.TimeStamp
        If this.StorePrivateDetails Then
            .Parameters("[paramUserID]") = GetUserID
            .Parameters("[paramComputerName]") = GetUserComputerName
        Else
            .Parameters("[paramUserID]") = Null
            .Parameters("[paramComputerName]") = Null
        End If
        .Parameters("[paramErrorNumber]") = ErrorInfoIn.ErrNumber
        .Parameters("[paramEventSource]") = ErrorInfoIn.FunctionSource
        .Parameters("[paramEventMessage]") = strMessage
        .Parameters("[paramPrinted]") = fWasDisplayed
        .Parameters("[paramErrorLevel]") = ErrorInfoIn.ErrorLevel

Retry:
        .Execute
        If Err Then
        ' Because we're in the log class, we can't rely on the catch
        ' or other tools because everything logs to here...
            If Err.Number = -2147217865 Then
                ' The log table got deleted somehow.
                Err.Clear
                SetupLogTable this.LogadoStoredConnection
                .Execute
                If Err Then this.Console.Add FormatEntry("COULD NOT SAVE LOG! " & strMessage, True, "RED", True, True)
            Else
                this.Console.Add FormatEntry("ERROR During Log Save: " & Err.Number & " : " & Err.Description, True, "RED", True, True)
            End If
            Err.Clear
        End If
    End With

Exit_Here:
    Perf.OperationEnd

End Sub


Private Function GetUserID() As String
    ' Gets the user's ID from the local machine.
    Static UserIDCache As String
    Dim objNetwork As Object

    If UserIDCache = vbNullString And this.StorePrivateDetails Then
        Set objNetwork = CreateObject("WScript.Network")
        UserIDCache = objNetwork.UserName
        Set objNetwork = Nothing
    End If
    GetUserID = UserIDCache

End Function


Private Function GetUserComputerName() As String

    Static UserComputerNameCache As String
    Dim objNetwork As Object

    If UserComputerNameCache = vbNullString And this.StorePrivateDetails Then
        Set objNetwork = CreateObject("WScript.Network")
        UserComputerNameCache = objNetwork.ComputerName
        Set objNetwork = Nothing
    End If
    GetUserComputerName = UserComputerNameCache

End Function


Private Function GetLogStoredCommand() As ADODB.Command

    Dim SQLString As clsConcat

    If this.adoStoredCMDAddEntry Is Nothing Then
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

        Set this.adoStoredCMDAddEntry = New ADODB.Command
        With this.adoStoredCMDAddEntry
            Set .ActiveConnection = Me.LogConnectionADO
            .CommandText = SQLString.GetStr
            .CommandType = adCmdText
            .Prepared = True
            .NamedParameters = True
            .Parameters.Append .CreateParameter("[paramEventLogTime]", adVarChar, adParamInput, 40)
            .Parameters.Append .CreateParameter("[paramUserID]", adVarChar, adParamInput, 40, Null)
            .Parameters.Append .CreateParameter("[paramComputerName]", adVarChar, adParamInput, 255, Null)
            .Parameters.Append .CreateParameter("[paramErrorNumber]", adInteger, adParamInput)
            .Parameters.Append .CreateParameter("[paramEventSource]", adLongVarChar, adParamInput, 5000)
            .Parameters.Append .CreateParameter("[paramEventMessage]", adLongVarChar, adParamInput, 63999)
            .Parameters.Append .CreateParameter("[paramPrinted]", adBoolean, adParamInput)
            .Parameters.Append .CreateParameter("[paramErrorLevel]", adInteger, adParamInput)
        End With
    End If

    With this.adoStoredCMDAddEntry
        If .ActiveConnection Is Nothing Then Set .ActiveConnection = Me.LogConnectionADO
        If (.ActiveConnection.State And adStateOpen) <> adStateOpen Then .ActiveConnection.Open
    End With

    Set GetLogStoredCommand = this.adoStoredCMDAddEntry

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
' Procedure : Error
' Author    : Adam Waller, hecon5, heavily modified from original
' Date      : 12/4/2020
' Purpose   : Log an error, and update error level if needed. Reads Err object values.
'           : A critical error will also present a message box with the details.
'---------------------------------------------------------------------------------------
'
Public Function Error(ByVal ErrorLevelIn As eErrorLevel _
                    , ByRef strBold As String _
                    , Optional ByRef strSource As String _
                    , Optional ByRef strErrDescription As String _
                    , Optional ByVal vbButtons As VbMsgBoxStyle = vbOKOnly _
                    , Optional ByRef strTitle As String = vbNullString _
                    , Optional ByRef moreDetails As String = vbNullString _
                    , Optional ByRef ErrorInfoIn As clsErrorInfo = Nothing) As VbMsgBoxResult

    Const FunctionName As String = ModuleName & ".Error"

    Dim ErrorResponse As VbMsgBoxResult
    Dim errorBold As New clsConcat
    Dim f_ErrorInfo As clsErrorInfo
    Dim f_vbButtons As VbMsgBoxStyle
    Dim strWorking As New clsConcat

    Error = vbCancel    ' Set default output.

    Perf.OperationStart FunctionName
    If ErrorInfoIn Is Nothing Then
        Set f_ErrorInfo = New clsErrorInfo
    Else
        Set f_ErrorInfo = ErrorInfoIn
    End If

    With f_ErrorInfo
        .ErrorLevel = ErrorLevelIn
        If strSource <> vbNullString Then .FunctionSource = f_ErrorInfo.FunctionSource & _
                            IIf(f_ErrorInfo.FunctionSource = vbNullString, strSource, "." & strSource)
        ' Now append remaining bits.
        .FunctionSource = f_ErrorInfo.FunctionSource & _
                            IIf(f_ErrorInfo.FunctionSource = vbNullString, "Log.", ".") & "Error"
    End With

    Set ErrorInfoIn = f_ErrorInfo ' Hand out the updated class details.

    With errorBold
        .Add cstrSpacer, vbNewLine
        .Add f_ErrorInfo.ErrorPrefix, vbNewLine
        .Add strBold, vbNewLine
        .Add f_ErrorInfo.ErrorDetailsBrief
        .Add vbNewLine, strErrDescription, vbNewLine
        .Add cstrSpacer
    End With

    ' Or performs a bitwise operation here
    f_vbButtons = vbButtons Or f_ErrorInfo.ErrorMsgBoxStyle
    ' Build the error message string.

    ' Show message on any error level when we are not logging to a file.
    Select Case ErrorLevelIn
        Case eelWarning, eelError, eelCritical
            ErrorResponse = Me.Prompt(strSource:=vbNullString _
                                    , strBold:=errorBold.GetStr _
                                    , strLine1:=ErrorInfoIn.ErrorDetailsFull _
                                    , strLine2:=moreDetails _
                                    , ErrorLevelIn:=ErrorLevelIn _
                                    , vbButtons:=f_vbButtons _
                                    , strTitle:=strTitle _
                                    , vbDefaultResult:=DEFAULT_PROMPT_VBRESULT _
                                    , ErrorInfoIn:=ErrorInfoIn)

            If ErrorLevelIn = eelCritical Then
                Me.SaveFile ' Performance reports are logged during save routines.
            Else
                LogPerformanceReports
            End If

    Case Else
        ' Log the error only if user isn't prompted; the prompt logs it for you
        With strWorking
            .Add FormatEntry(errorBold.GetStr, True, f_ErrorInfo.ErrorColor, True)
            If Len(f_ErrorInfo.ErrorDetailsFull) > 0 Then .Add vbNewLine, f_ErrorInfo.ErrorDetailsFull
            If moreDetails <> vbNullString Then .Add vbNewLine, "More Details: ", moreDetails
            If f_vbButtons <> vbOKOnly Then .Add vbNewLine, "Buttons: ", CStr(f_vbButtons)
            If strTitle <> vbNullString Then .Add vbNewLine, "Title: ", strTitle
        End With

        SaveEntry strMessage:=strWorking.GetStr _
                    , ErrorInfoIn:=ErrorInfoIn _
                    , NextOutputOnNewLine:=True _
                    , PrintBold:=True
    End Select

    ' Update error level if higher.
    If Operation.ErrorLevel < ErrorLevelIn Then Operation.ErrorLevel = ErrorLevelIn
    Perf.OperationEnd
    Error = ErrorResponse

End Function


Public Function Prompt(ByRef strBold As String _
                    , Optional ByRef strSource As String = vbNullString _
                    , Optional ByRef ErrorLevelIn As eErrorLevel = eErrorLevel.eelNoError _
                    , Optional ByRef strLine1 As String = vbNullString _
                    , Optional ByRef strLine2 As String = vbNullString _
                    , Optional ByRef vbButtons As VbMsgBoxStyle = vbOKOnly _
                    , Optional ByRef strTitle As String = vbNullString _
                    , Optional ByVal vbDefaultResult As VbMsgBoxResult = DEFAULT_PROMPT_VBRESULT _
                    , Optional ByRef ErrorInfoIn As clsErrorInfo = Nothing) As VbMsgBoxResult

    ' This function will prompt the user to do something, or display a message, and return the value to the calling function.
    ' In addition, it will log that the prompt happened, and the result of the user action.

    Const FunctionName As String = ModuleName & ".Prompt"

    Dim intCursor As Integer
    Dim PromptResult As VbMsgBoxResult
    Dim titleFormatted As String
    Dim f_ErrorInfo As clsErrorInfo

    Dim strBoldLines As String
    Dim strBoxLines1 As String
    Dim msgBoxStr2 As New clsConcat

    Dim logEntryDetails As New clsConcat
    Dim LineCount As Long

    Perf.OperationStart FunctionName

    If ErrorInfoIn Is Nothing Then
        Set f_ErrorInfo = New clsErrorInfo
    Else
        Set f_ErrorInfo = ErrorInfoIn
    End If

    With f_ErrorInfo
        .ErrorLevel = ErrorLevelIn
        If strSource <> vbNullString Then .FunctionSource = f_ErrorInfo.FunctionSource & _
                            IIf(f_ErrorInfo.FunctionSource = vbNullString, strSource, "." & strSource)
        ' Now append remaining bits.
        .FunctionSource = f_ErrorInfo.FunctionSource & _
                            IIf(f_ErrorInfo.FunctionSource = vbNullString, "Log.", ".") & "Prompt"
    End With

    Set ErrorInfoIn = f_ErrorInfo ' Hand out the updated class details.

    ' Turn off any hourglass
    intCursor = Screen.MousePointer

    titleFormatted = Nz2(strTitle, ErrorInfoIn.FunctionSource)

    LineCount = SCREEN_ROWS_MAX_DEFAULT
    strBoldLines = FitStringToWidth(strBold, WIDTH_MAX_MSGBOX, WIDTH_DESIRED_MSGBOX, LineCount)

    If strLine1 <> vbNullString Then
        LineCount = LineCount + MSGBOX_LINE1_ROWS
        strBoxLines1 = FitStringToWidth(strLine1, WIDTH_MAX_MSGBOX, WIDTH_DESIRED_MSGBOX, LineCount)
    End If

    With msgBoxStr2
        If strLine2 <> vbNullString Then .Add FitStringToWidth(strLine2, WIDTH_MAX_MSGBOX, WIDTH_DESIRED_MSGBOX, LineCount + MSGBOX_LINE2_ROWS)
        If ErrorInfoIn.FunctionSource <> vbNullString Then
            If .Length Then .Add vbNewLine, vbNewLine
            .Add "Function Source: ", FitStringToWidth(ErrorInfoIn.FunctionSource, WIDTH_MAX_MSGBOX - 8, WIDTH_DESIRED_MSGBOX - 8, MSGBOX_SOURCE_ROWS)
        End If
    End With

    Perf.PauseTiming

    ' Check interaction mode
    If InteractionMode = eimNormal Then
        PromptResult = MsgBoxLog(strBold:=strBoldLines _
                            , strLine1:=strBoxLines1 _
                            , strLine2:=msgBoxStr2.GetStr _
                            , vbButtons:=vbButtons, strTitle:=titleFormatted)
    Else
        ' Just log the output
        ' https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/msgbox-function
        ' The enum for VbMsgBoxResult does not have a zero value and we need a way to tell the system user hasn't said anything.
        PromptResult = 0
    End If

    Perf.ResumeTiming
    If IsEntryLogged(ErrorLevelIn) Then
        With logEntryDetails
            .AppendOnAdd = "<br>"
            .Add cstrSpacer
            If InteractionMode = eimSilent Then .Add "[**MessageBox Not Displayed**]"
            .Add "Prompt:"
            .Add FormatEntry(strBold, True, ErrorInfoIn.ErrorColor, True)
            .Add strLine1
            .Add strLine2
            .Add "Function Source: ", ErrorInfoIn.FunctionSource
            If vbButtons <> vbOKOnly Then .Add "Buttons: ", CStr(vbButtons)
            If strTitle <> vbNullString Then .Add "Title: ", titleFormatted
            .Add FormatEntry("User Response: " & PromptResult, True, , True)
            .Add cstrSpacer
            .Remove 4 ' Remove the last line feed.
        End With
        SaveEntry strMessage:=logEntryDetails.GetStr, ErrorInfoIn:=ErrorInfoIn
    End If

    Prompt = PromptResult

    ' Restore MousePointer (if needed)
    If intCursor > 0 Then Screen.MousePointer = intCursor

    Perf.OperationEnd

End Function


Public Function PromptInput(ByRef strMainLine As String _
                    , Optional ByRef strSource As String = vbNullString _
                    , Optional ByRef ErrorLevelIn As eErrorLevel = eErrorLevel.eelNoError _
                    , Optional ByRef strLine1 As String = vbNullString _
                    , Optional ByRef strLine2 As String = vbNullString _
                    , Optional ByRef DefaultInput As String = vbNullString _
                    , Optional ByRef strTitle As String = vbNullString _
                    , Optional ByVal LogInputReturn As Boolean = False _
                    , Optional ByRef ErrorInfoIn As clsErrorInfo = Nothing) As String

    ' This function will prompt the user to do something, or display a message, and return the value to the calling function.
    ' In addition, it will log that the prompt happened, and the result of the user action.

    Const FunctionName As String = ModuleName & ".PromptInput"

    Dim intCursor As Integer
    Dim PromptResult As String
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

    sourceFormatted = Nz2(strSource, "Log") & ".PromptInput"
    titleFormatted = Nz2(strTitle, sourceFormatted)
    errorLevelInfo.ErrorLevel = ErrorLevelIn

    LineCount = SCREEN_ROWS_MAX_DEFAULT
    strBoldLines = FitStringToWidth(strMainLine, WIDTH_MAX_MSGBOX, WIDTH_DESIRED_MSGBOX, LineCount)

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
        PromptResult = InputBoxLog(strMainLine:=strBoldLines _
                            , strLine1:=strBoxLines1 _
                            , strLine2:=msgBoxStr2.GetStr _
                            , DefaultInput:=DefaultInput, strTitle:=titleFormatted)
    Else
        ' Just log the output
        ' https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/msgbox-function
        ' The enum for VbMsgBoxResult does not have a zero value and we need a way to tell the system user hasn't said anything.
        PromptResult = vbNullString
    End If

    Perf.ResumeTiming
    If IsEntryLogged(ErrorLevelIn) Then
        With logEntryDetails
            .AppendOnAdd = "<br>"
            .Add cstrSpacer
            If InteractionMode = eimSilent Then .Add "[**MessageBox Not Displayed**]"
            .Add "Prompt:"
            .Add FormatEntry(strMainLine, True, errorLevelInfo.ErrorColor, True)
            .Add strLine1
            .Add strLine2
            .Add "Source: ", sourceFormatted
            If strTitle <> vbNullString Then .Add "Title: ", titleFormatted
            .Add FormatEntry("User Response: " & PromptResult, True, , True)
            If LogInputReturn Then
                .Add "User Input Response: ", PromptResult
            Else
                .Add "User Input Response: {REDACTED}"
            End If
            .Add cstrSpacer
            .Remove 4 ' Remove the last line feed.
        End With
        Me.AddEntry strSource:=sourceFormatted, strMessage:=logEntryDetails.GetStr, ErrorLevelIn:=ErrorLevelIn
    End If

    PromptInput = PromptResult

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
Private Function MsgBoxLog(ByRef strBold As String _
                        , Optional ByRef strLine1 As String = vbNullString _
                        , Optional ByRef strLine2 As String = vbNullString _
                        , Optional ByRef vbButtons As VbMsgBoxStyle = vbOKOnly _
                        , Optional ByRef strTitle As String = vbNullString) As VbMsgBoxResult

    Dim strMsg As String
    Dim varLines(0 To 3) As String

    ' Escape single quotes by doubling them.
    varLines(0) = Replace(strBold, "'", "''")
    varLines(1) = Replace(strLine1, "'", "''")
    varLines(2) = Replace(strLine2, "'", "''")
    varLines(3) = Replace(strTitle, "'", "''")

    If varLines(3) = vbNullString Then varLines(3) = Application.VBE.ActiveVBProject.Name
    strMsg = "MsgBox('" & varLines(0) & "@" & varLines(1) & "@" & varLines(2) & "@'," & vbButtons & ",'" & varLines(3) & "')"
    MsgBoxLog = Eval(strMsg)

End Function


'---------------------------------------------------------------------------------------
' Procedure : InputBox3
' Author    : hecon5
' Date      : 2025-07-03
' Purpose   : Alternate facilitate inputbox functionality and ability to log it.
'---------------------------------------------------------------------------------------
'
Private Function InputBoxLog(ByRef strMainLine As String _
                        , Optional ByRef strLine1 As String = vbNullString _
                        , Optional ByRef strLine2 As String = vbNullString _
                        , Optional ByRef DefaultInput As String = vbNullString _
                        , Optional ByRef strTitle As String = vbNullString) As String

    Dim strMsg As String
    Dim varLines(0 To 4) As String

    ' Escape single quotes by doubling them.
    varLines(0) = Replace(strMainLine, "'", "''")
    varLines(1) = Replace(strLine1, "'", "''")
    varLines(2) = Replace(strLine2, "'", "''")
    varLines(3) = Replace(strTitle, "'", "''")
    varLines(4) = Replace(DefaultInput, "'", "''")

    If varLines(3) = vbNullString Then varLines(3) = Application.VBE.ActiveVBProject.Name
    strMsg = "InputBox('" & varLines(0) & vbNewLine & vbNewLine & varLines(1) & vbNewLine & vbNewLine & varLines(2) & "'" & _
                    ",'" & varLines(3) & "','" & varLines(4) & "')"
    InputBoxLog = Eval(strMsg)

End Function


'---------------------------------------------------------------------------------------
' Procedure : SetConsole
' Author    : Adam Waller
' Date      : 4/28/2020
' Purpose   : Set a reference to the text box.
'---------------------------------------------------------------------------------------
'
Public Sub SetConsole(ByRef RichTextControlIn As TextBox _
                    , ByRef ProgressBarIn As clsLblProg)

    Set this.RichText = RichTextControlIn
    Set this.Prog = ProgressBarIn
    If Not this.RichText Is Nothing Then
        this.RichText.AllowAutoCorrect = False
    End If

End Sub


' Clear the console box of other text.
Public Sub ClearConsole()
    this.Console.Clear
    Me.Flush
End Sub


'---------------------------------------------------------------------------------------
' Procedure : ReleaseConsole
' Author    : Adam Waller, hecon5
' Date      : 5/5/2022
' Purpose   : Releases object references to the console form that were created with
'           : the SetConsole sub.
'---------------------------------------------------------------------------------------
Public Sub ReleaseConsole()
    Set this.RichText = Nothing
    Set this.Prog = Nothing
End Sub


'---------------------------------------------------------------------------------------
' Procedure : ProgressBar
' Author    : Adam Waller
' Date      : 11/6/2020
' Purpose   : Pass the Progress Bar reference to this class.
'---------------------------------------------------------------------------------------
'
Public Property Set ProgressBar(cProg As clsLblProg)
    Set this.Prog = cProg
End Property
Public Property Get ProgressBar() As clsLblProg
    If this.Prog Is Nothing Then Set this.Prog = New clsLblProg
    Set ProgressBar = this.Prog
End Property


'---------------------------------------------------------------------------------------
' Procedure : ProgMax
' Author    : Adam Waller
' Date      : 11/6/2020
' Purpose   : Wrapper to set max value for progress bar.
'---------------------------------------------------------------------------------------
'
Public Property Let ProgMax(NewMaxValue As Long)
    If Not this.Prog Is Nothing Then this.Prog.Max = NewMaxValue
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


Public Sub LogPerformanceReports(Optional ByVal EventLevelIn As eErrorLevel = eelDebugInfo)
    Me.AddEntry strSource:="Performance Reports" _
                , strMessage:=Perf.GetReports _
                , ErrorLevelIn:=EventLevelIn _
                , OutputColor:="black" _
                , PrintBold:=False
End Sub


'---------------------------------------------------------------------------------------
' Procedure : LogFilePath
' Author    : Adam Waller
' Date      : 5/13/2023
' Purpose   : Return the (default) full path to the current log file.
'---------------------------------------------------------------------------------------
'
Public Property Get LogFilePath() As String

    Dim SanitizedFileName As String

    Select Case Operation.OperationType
        Case eotExport: SanitizedFileName = "Export"
        Case eotBuild:  SanitizedFileName = "Build"
        Case eotMerge:  SanitizedFileName = "Merge"
        Case Else:      SanitizedFileName = "Other"
    End Select

    SanitizedFileName = SanitizedFileName & ".Log"
    'SanitizedFileName = SanitizedFileName & ISO8601TimeStamp(False, False) & ".xml"
    'SanitizedFileName = MultiReplace(SanitizedFileName, ":", "-") ' Replace Timestamp charachters with valid charachters.

    LogFilePath = FSO.BuildPath(Nz2(SourcePath, Options.GetExportFolder), SanitizedFileName)

End Property


Public Sub SaveFile(Optional ByRef strAlternateFilePath As String _
                    , Optional ByVal NotQuiet As Boolean = False _
                    , Optional ByVal ClearConsoleAlso As Boolean = False)

    Const FunctionName As String = ModuleName & ".SaveFile"

    Dim OutputFilePath As String
    Dim SanitizedFileName As String

Retry:
    ' Before exporting, insert entries for performance, because that might be useful.
    LogPerformanceReports

    OutputFilePath = Nz2(strAlternateFilePath, Me.LogFilePath)
    VerifyPath OutputFilePath
    SanitizedFileName = FSO.GetFileName(OutputFilePath)

    With New ADODB.Recordset
        .Open Me.LogTableName, Me.LogConnectionADO, adOpenDynamic
        If FSO.FileExists(OutputFilePath) Then DeleteFile OutputFilePath, True
        .Save OutputFilePath, adPersistXML
    End With

    With New clsSourceParser
        .LoadSourceFile OutputFilePath, edbTableData
        WriteFile .Sanitize(ectXML), OutputFilePath
    End With

    If Err.Number <> 0 Then
        If Me.Prompt(strSource:=FunctionName, strBold:="Error saving file." _
                    , strLine1:="Retry?" _
                    , vbButtons:=vbYesNo, ErrorLevelIn:=eelError) = vbYes Then GoTo Retry

    Else
        If NotQuiet Then
            Me.Prompt strSource:=FunctionName, strBold:="Successfully saved log!" _
                    , strLine1:="Saved to: " & Me.LogFilePath _
                    , strLine2:="File Name: " & SanitizedFileName _
                    , strTitle:="Successfully saved log!"
        End If
        Me.Clear False, ClearConsoleAlso
    End If

End Sub


Public Sub CheckSize()
    ' See if we should clear out the old log entries.
    If Me.EntryCount > MAX_ENTRY_COUNT_PROMPT Then Me.Clear True
End Sub


Public Property Get EntryCount() As Long

    Perf.OperationStart ModuleName & ".EntryCount"
    EntryCount = DCount("[EntryId]", Me.LogTableName)
    Perf.OperationEnd

End Property


Public Property Get LogTableName() As String

    If this.LogTableName = vbNullString Then this.LogTableName = DefaultLogTableName
    LogTableName = this.LogTableName

End Property


Public Property Get LogConnectionADO() As ADODB.Connection

    If this.LogadoStoredConnection Is Nothing Then Set this.LogadoStoredConnection = GetDefaultLogConnectionADO
    Set LogConnectionADO = this.LogadoStoredConnection

End Property
Public Property Let LogConnectionADO(NewVal As ADODB.Connection)

    If Not NewVal Is Nothing Then
        Set this.LogadoStoredConnection = NewVal
    Else
        Set this.LogadoStoredConnection = GetDefaultLogConnectionADO
    End If

End Property


Private Function GetDefaultLogConnectionADO() As ADODB.Connection

    Dim strPath As String
    Dim NeedToSetupDB As Boolean
    Dim NeedToSetupTable As Boolean

    Dim NewDbFile As DAO.Database

    strPath = CodeProject.Path & PathSep & "Log.accdu"
    NeedToSetupDB = Not FSO.FileExists(strPath)

    If NeedToSetupDB Then
        Set NewDbFile = DBEngine.CreateDatabase(strPath, dbLangGeneral)

        NeedToSetupTable = True
    Else
        Set NewDbFile = DBEngine.OpenDatabase(strPath)
        NeedToSetupTable = TableExists(Me.LogTableName, NewDbFile)
    End If
    NewDbFile.Close
    Set GetDefaultLogConnectionADO = New ADODB.Connection
    With GetDefaultLogConnectionADO
        .ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & strPath & ";Persist Security Info=False;"
        .Open
    End With
    If NeedToSetupDB Then SetupLogTable GetDefaultLogConnectionADO

End Function


Private Function SetupLogTable(ByRef adoConIn As ADODB.Connection) As String

    Dim SQLString As String
    Dim strPath As String
    Dim adoCmdRun As ADODB.Command
    Dim NewDbFile As DAO.Database
    Dim lTable As DAO.TableDef
    Dim fld As DAO.Field
    Dim logProperty As Property

    SQLString = "CREATE TABLE " & this.LogTableName & " (" & _
                        "[EntryID] AUTOINCREMENT, " & _
                        "[EventLogTime] VARCHAR(40), " & _
                        "[UserID] VARCHAR(40), " & _
                        "[ComputerName] VARCHAR(255), " & _
                        "[ErrorNumber] Long, " & _
                        "[EventSource] LongText, " & _
                        "[EventMessage] LongText, " & _
                        "[Printed] Bit, " & _
                        "[ErrorLevel] Long" & _
                        ");"

    With New ADODB.Command
        .CommandText = SQLString
        .ActiveConnection = adoConIn
        If (.ActiveConnection.State And adStateOpen) <> adStateOpen Then .ActiveConnection.Open
        .Execute
    End With

    strPath = CodeProject.Path & PathSep & "Log.accdu"

    Set NewDbFile = DBEngine.OpenDatabase(strPath)

    ' Now set EventMessage to rich text, because you can't do it in SQL...
    Set lTable = NewDbFile.TableDefs(this.LogTableName)
    Set fld = lTable.Fields("EventMessage")

    On Error Resume Next
    With lTable.Fields("EventMessage").Properties("TextFormat")
        If Catch(3270) Then
            Set logProperty = fld.CreateProperty("TextFormat", dbInteger, acTextFormatHTMLRichText)
            fld.Properties.Append logProperty
        End If
        If .Value = acTextFormatPlain Then .Value = acTextFormatHTMLRichText
    End With
    On Error GoTo 0
    SetupLogTable = lTable.Name
    Exit Function

End Function


Private Function CleanHTMLFromStr(ByRef StringToClean As String) As String

    Dim StrCleaned As String

    StrCleaned = MultiReplace(StringToClean, _
                                "<br>", vbNewLine, _
                                "&nbsp;", " ")

    'CleanHTMLFromStr = CleanStr(StrCleaned, "<", ">")
    CleanHTMLFromStr = StrCleaned

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
            If ArrPosition > 0 And (StrArrLen > 0 Or Not RemoveBlankLines) Then
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
