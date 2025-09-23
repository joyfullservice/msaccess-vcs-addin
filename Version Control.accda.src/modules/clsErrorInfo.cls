VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsErrorInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Text
Option Explicit

Private Const ModuleName As String = "clsErrorInfo"


Private Type udtThis
    ' Input Data
    ErrorLevel As eErrorLevel       ' The level of Error
    FunctionSource As String        ' The calling function's source name

    ' Cached Error Info (grabbed when the class is instantiated).
    ' If you simply create an Err object, you will just get a pointer
    ' to the main Err ojbect, and this isn't what we want.
    ' Err Object variables
    ErrDescription As String
    ErrHelpContext As Long
    ErrHelpFile As String
    ErrLastDllError As Long
    ErrNumber As Long
    ErrSource As String

    ' Details saved when the class is initialized.
    TimeStamp As String
    CallStack As String

    ' Output Data
    ErrorLevelName As String
    ErrorPrefix As String
    ErrorOptionName As String
    ErrorColor As String
    ErrorMsgBoxStyle As VbMsgBoxStyle

    ErrorDetailsBrief As String     ' oneliner error info.
    ErrorDetailsFull As String      ' Standard formatted detail block.

End Type
Private this As udtThis


Private Sub Class_Initialize()

    ' When the class is initialized, automatically grab the Error Information
    ' We can set this manually later, but having this done here speeds things
    ' along and simplifies error caching.
    ErrStateCache

End Sub


Public Sub Init(ByVal ErrorLevelIn As eErrorLevel _
                , Optional ByRef SourceIn As String = vbNullString)

    With this
        .ErrorLevel = ErrorLevelIn
        .FunctionSource = SourceIn
    End With

End Sub


Public Sub ErrStateCache(Optional ByVal Refresh As Boolean = False)

    If this.ErrorLevel = eelUninitialized Or Refresh Then
        With this
            .TimeStamp = ISO8601TimeStamp   ' Get the time of event.
            .CallStack = Perf.CallStack     ' Get CallStack at time of creation

            ' Get Err details.
            .ErrDescription = Err.Description
            .ErrHelpContext = Err.HelpContext
            .ErrHelpFile = Err.HelpFile
            .ErrLastDllError = Err.LastDllError
            .ErrNumber = Err.Number
            .ErrSource = Err.Source

            If .ErrNumber <> 0 Then
                .ErrorLevel = eelError      ' Raise an error if there is one.
            Else
                .ErrorLevel = eelNoError    ' Set to eelNoError; this can be set elsewhere.
            End If
        End With

        ' We've logged the error. The purpose of this class is to cache this info, so we
        ' do not need to keep the Err state, as we can get that from this class.
        Err.Clear

        SetErrorLevelInfoData
    End If

End Sub


Public Sub ErrClear()

    With this
        .ErrDescription = vbNullString
        .ErrHelpContext = 0
        .ErrHelpFile = vbNullString
        .ErrLastDllError = 0
        .ErrNumber = 0
    End With

End Sub


Public Sub ErrRaise()

    With this
        Err.Raise .ErrNumber, .ErrSource, .ErrDescription, .ErrHelpFile, .ErrHelpContext
    End With

End Sub


Public Property Get ErrDescription() As String
    ErrDescription = this.ErrDescription
End Property

Public Property Get ErrHelpContext() As Long
    ErrHelpContext = this.ErrHelpContext
End Property

Public Property Get ErrHelpFile() As String
    ErrHelpFile = this.ErrHelpFile
End Property

Public Property Get ErrLastDllError() As Long
    ErrLastDllError = this.ErrLastDllError
End Property

Public Property Get ErrNumber() As Long
    ErrNumber = this.ErrNumber
End Property

Public Property Get ErrSource() As String
    ErrSource = this.ErrSource
End Property


Public Property Get TimeStamp() As String
    If this.TimeStamp = vbNullString Then this.TimeStamp = ISO8601TimeStamp
    TimeStamp = this.TimeStamp
End Property


Public Property Get CallStack() As String
    If this.CallStack = vbNullString Then this.CallStack = Perf.CallStack
    CallStack = this.CallStack
End Property


Public Property Get ErrorLevel() As eErrorLevel
    ErrorLevel = this.ErrorLevel
End Property
Public Property Let ErrorLevel(newErrorLevel As eErrorLevel)
    If newErrorLevel = eelUninitialized Then Exit Property ' Don't save resets.
    this.ErrorLevel = newErrorLevel
    SetErrorLevelInfoData
End Property


Public Property Get FunctionSource() As String
    FunctionSource = this.FunctionSource
End Property
Public Property Let FunctionSource(ByRef newFunctionSource As String)
    this.FunctionSource = newFunctionSource
End Property


Public Property Get ErrorLevelName() As String
    ErrorLevelName = this.ErrorLevelName
End Property


Public Property Get ErrorOptionName() As String
    ErrorOptionName = this.ErrorOptionName
End Property


Public Property Get ErrorPrefix() As String
    ErrorPrefix = this.ErrorPrefix
End Property


Public Property Get ErrorColor() As String
    If this.ErrorColor = vbNullString Then this.ErrorColor = getErrorColor(this.ErrorLevel)
    ErrorColor = this.ErrorColor
End Property
Public Property Let ErrorColor(newColor As String)
    this.ErrorColor = ErrorColor
End Property


Public Property Get ErrorMsgBoxStyle() As VbMsgBoxStyle
    ErrorMsgBoxStyle = this.ErrorMsgBoxStyle
End Property


Public Property Get ErrorDetailsBrief() As String

    If this.ErrorDetailsBrief = vbNullString And this.ErrNumber <> 0 Then
        this.ErrorDetailsBrief = "Number: " & this.ErrNumber & " : " & this.ErrDescription & " Source: " & this.ErrSource
    ElseIf this.ErrNumber = 0 Then
        ' make sure details are cleared.
        this.ErrorDetailsBrief = vbNullString
    End If

    ErrorDetailsBrief = this.ErrorDetailsBrief

End Property


Public Property Get ErrorDetailsFull() As String

    If this.ErrorDetailsFull = vbNullString Then
        With New clsConcat
            If this.ErrNumber <> 0 Then
                .Add "Error Number: ", CStr(this.ErrNumber), " ", this.ErrDescription
                .Add vbNewLine, "Help Context: ", CStr(this.ErrHelpContext)
                .Add vbNewLine, "HelpFile: ", this.ErrHelpFile
                .Add vbNewLine, "LastDllError: ", CStr(this.ErrLastDllError)
                .Add vbNewLine, "Source: ", this.ErrSource
            End If
            .Add vbNewLine, "TimeStamp: ", this.TimeStamp
            .Add vbNewLine, "CallStack: ", this.CallStack
            this.ErrorDetailsFull = .GetStr
        End With
    End If

    ErrorDetailsFull = this.ErrorDetailsFull

End Property


Private Function SetErrorLevelInfoData()

    With this
        .ErrorLevelName = GetErrorLevelName(this.ErrorLevel)
        .ErrorColor = getErrorColor(this.ErrorLevel)
        .ErrorPrefix = .ErrorLevelName & ": "
        .ErrorOptionName = GetErrorOptionName(this.ErrorLevel)
        .ErrorMsgBoxStyle = GetMsgBoxStyle(this.ErrorLevel)
    End With

End Function


Private Function GetMsgBoxStyle(ByVal ErrorLevelIn As eErrorLevel) As VbMsgBoxStyle

    Select Case ErrorLevelIn
'    Case eelNoError:        GetMsgBoxStyle = vbOKOnly
'    Case eelDebugTrace:     GetMsgBoxStyle = vbOKOnly
'    Case eelDebugEvent:     GetMsgBoxStyle = vbOKOnly
    Case eelDebugInfo:      GetMsgBoxStyle = vbQuestion
    Case eelWarning:        GetMsgBoxStyle = vbInformation
    Case eelError:          GetMsgBoxStyle = vbExclamation
    Case eelCritical:       GetMsgBoxStyle = vbCritical
    End Select

End Function


Private Function GetErrorLevelName(ByVal ErrorLevelIn As eErrorLevel) As String

    Select Case ErrorLevelIn
    Case eelUninitialized:  GetErrorLevelName = "Uninitialized Error Level"
    Case eelLoggingOff:     GetErrorLevelName = "LOGGING OFF"
    Case eelNoError:        GetErrorLevelName = "NO ERROR"
    Case eelDebugTrace:     GetErrorLevelName = "TRACE"
    Case eelDebugEvent:     GetErrorLevelName = "EVENT"
    Case eelDebugInfo:      GetErrorLevelName = "INFO"
    Case eelWarning:        GetErrorLevelName = "WARNING"
    Case eelError:          GetErrorLevelName = "ERROR"
    Case eelCritical:       GetErrorLevelName = "CRITICAL"
    End Select

End Function


Public Function GetErrorOptionName(ByRef ErrorLevel As eErrorLevel) As String

    Select Case ErrorLevel
    Case eelUninitialized:  GetErrorOptionName = "Uninitialized Error Level"
    Case eelLoggingOff:     GetErrorOptionName = "Off"
    Case eelNoError:        GetErrorOptionName = "Note"
    Case eelDebugTrace:     GetErrorOptionName = "Debug Trace"
    Case eelDebugEvent:     GetErrorOptionName = "Debug Event"
    Case eelDebugInfo:      GetErrorOptionName = "Debug Info"
    Case eelWarning:        GetErrorOptionName = "Warning"
    Case eelError:          GetErrorOptionName = "Error"
    Case eelCritical:       GetErrorOptionName = "Critical Error"
    End Select

End Function


' Checks the pretty version and the formal name of the error level and sets the error level
' based on the input.
Public Function SetErrorLevelFromName(ByRef strErrorName As String) As eErrorLevel

    Dim f_ErrorLevel As eErrorLevel

    For f_ErrorLevel = eelNoError To eErrorLevel.[_Last]
        If strErrorName = GetErrorLevelName(f_ErrorLevel) Then
            SetErrorLevelFromName = f_ErrorLevel
            this.ErrorLevel = f_ErrorLevel
            Exit For
        End If
        ' Also check the "Pretty" version as well.
        If strErrorName = GetErrorOptionName(f_ErrorLevel) Then
            SetErrorLevelFromName = f_ErrorLevel
            this.ErrorLevel = f_ErrorLevel
            Exit For
        End If
    Next f_ErrorLevel

End Function


Private Function getErrorColor(ByRef ErrorLevel As eErrorLevel) As String

    Select Case ErrorLevel
    Case eelNoError:        getErrorColor = "black"
    Case eelDebugTrace:     getErrorColor = "black"
    Case eelDebugEvent:     getErrorColor = "blue"
    Case eelDebugInfo:      getErrorColor = "green"
    Case eelWarning:        getErrorColor = "yellow"
    Case eelError:          getErrorColor = "orange"
    Case eelCritical:       getErrorColor = "red"
    End Select

End Function
