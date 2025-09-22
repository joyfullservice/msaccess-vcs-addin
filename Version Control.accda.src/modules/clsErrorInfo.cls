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

    IsSetup As Boolean

    ' Input Data
    ErrorLevel As eErrorLevel
    FunctionSource As String

    ' Cached Error Info (grabbed when the class is instantiated).
    ' If you simply create an Err object, you will just get a pointer
    ' to the main Err ojbect, and this isn't what we want.
    Description As String
    HelpContext As Long
    HelpFile As String
    LastDllError As Long
    Number As Long
    Source As String

    TimeStamp As String
    CallStack As String

    ' Output Data
    ErrorLevelName As String
    ErrorPrefix As String
    OptionName As String
    ErrorColor As String

    ErrorDetailsBrief As String     ' oneliner error info.
    ErrorDetailsFull As String      ' Standard formatted detail block.

    MsgBoxStyle As VbMsgBoxStyle

End Type
Private this As udtThis


Private Sub Class_Initialize()

    ' When the class is initialized, automatically grab the Error Information
    ' We can set this manually later, but having this done here speeds things
    ' along and simplifies error caching.
    ErrStateCache

End Sub


Public Sub ErrStateCache(Optional ByVal Reset As Boolean = False)

    If this.TimeStamp = vbNullString Or Reset Then
        With this
            .Description = Err.Description
            .HelpContext = Err.HelpContext
            .HelpFile = Err.HelpFile
            .LastDllError = Err.LastDllError
            .Number = Err.Number
            .Source = Err.Source

            .TimeStamp = ISO8601TimeStamp   ' Get the time of event.
            .CallStack = Perf.CallStack     ' Get CallStack at time of creation
        End With
        this.ErrorLevel = eelNoError        ' Set to eelNoError; this can be set elsewhere.
    End If

End Sub


Public Sub ErrCacheClear()

    Dim eThis As udtThis

    this = eThis

    With this
        .Description = Err.Description
        .HelpContext = 0
        .HelpFile = vbNullString
        .LastDllError = 0
        .Number = 0
        .Source = vbNullString
    End With

End Sub


Public Property Get TimeStamp() As String
    If this.TimeStamp = vbNullString Then this.TimeStamp = ISO8601TimeStamp
    TimeStamp = this.TimeStamp
End Property


Public Sub Raise()

    With this
        Err.Raise .Number, .Source, .Description, .HelpFile, .HelpContext
    End With

End Sub


Public Property Get Description() As String
    Description = this.Description
End Property
Public Property Let Description(newDescription As String)
    this.Description = newDescription
End Property


Public Property Get HelpContext() As Long
    HelpContext = this.Number
End Property
Public Property Let HelpContext(newHelpContext As Long)
    this.HelpContext = newHelpContext
End Property


Public Property Get HelpFile() As String
    HelpFile = this.HelpFile
End Property
Public Property Let HelpFile(newHelpFile As String)
    this.HelpFile = newHelpFile
End Property


Public Property Get LastDllError() As Long
    LastDllError = this.LastDllError
End Property
Public Property Let LastDllError(newLastDllError As Long)
    this.LastDllError = newLastDllError
End Property


Public Property Get Number() As Long
    Number = this.Number
End Property
Public Property Let Number(newNumber As Long)
    this.Number = newNumber
    If this.Number = 0 Then Me.ErrCacheClear
End Property


Public Property Get Source() As String
    Source = this.Source
End Property
Public Property Let Source(newSource As String)
    this.Source = newSource
End Property


Public Property Get ErrorLevel() As eErrorLevel
    ErrorLevel = this.ErrorLevel
End Property
Public Property Let ErrorLevel(newErrorLevel As eErrorLevel)
    this.ErrorLevel = newErrorLevel
    SetErrorLevelInfoData
End Property


Public Property Get ErrorLevelName() As String
    ErrorLevelName = this.ErrorLevelName
End Property


Public Property Get ErrorOptionName() As String
    ErrorOptionName = this.OptionName
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




Public Property Get MsgBoxStyle() As VbMsgBoxStyle
    MsgBoxStyle = this.MsgBoxStyle
End Property


Public Property Get ErrorDetailsBrief() As String

    If this.ErrorDetailsBrief = vbNullString And this.Number <> 0 Then
        this.ErrorDetailsBrief = "Number: " & this.Number & " : " & this.Description & " Source: " & this.Source
    ElseIf this.Number = 0 Then
        ' make sure details are cleared.
        this.ErrorDetailsBrief = vbNullString
    End If

    ErrorDetailsBrief = this.ErrorDetailsBrief

End Property


Public Property Get ErrorDetailsFull() As String

    If this.ErrorDetailsFull = vbNullString Then
        With New clsConcat
            If this.Number <> 0 Then
                .Add "Error Number: ", CStr(this.Number), " ", this.Description
                .Add vbNewLine, "Help Context: ", CStr(this.HelpContext)
                .Add vbNewLine, "HelpFile: ", this.HelpFile
                .Add vbNewLine, "LastDllError: ", CStr(this.LastDllError)
                .Add vbNewLine, "Source: ", this.Source
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
        .OptionName = GetErrorNameOption(this.ErrorLevel)
        .MsgBoxStyle = GetMsgBoxStyle(this.ErrorLevel)
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


Private Function GetErrorNameOption(ByRef ErrorLevel As eErrorLevel) As String

    Select Case ErrorLevel
    Case eelUninitialized:  GetErrorNameOption = "Uninitialized Error Level"
    Case eelLoggingOff:     GetErrorNameOption = "Off"
    Case eelNoError:        GetErrorNameOption = "Note"
    Case eelDebugTrace:     GetErrorNameOption = "Debug Trace"
    Case eelDebugEvent:     GetErrorNameOption = "Debug Event"
    Case eelDebugInfo:      GetErrorNameOption = "Debug Info"
    Case eelWarning:        GetErrorNameOption = "Warning"
    Case eelError:          GetErrorNameOption = "Error"
    Case eelCritical:       GetErrorNameOption = "Critical Error"
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
        If strErrorName = GetErrorNameOption(f_ErrorLevel) Then
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
