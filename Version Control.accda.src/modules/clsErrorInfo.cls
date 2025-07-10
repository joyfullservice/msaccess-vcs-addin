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
    ErrorLevel As eErrorLevel

    ' Output Data
    ErrorName As String
    Prefix As String
    PrettyName As String
    ErrorColor As String
    MsgBoxStyle As VbMsgBoxStyle

End Type
Private this As udtThis


Public Property Get ErrorLevel() As eErrorLevel
    ErrorLevel = this.ErrorLevel
End Property
Public Property Let ErrorLevel(newErrorLevel As eErrorLevel)
    this.ErrorLevel = newErrorLevel
    SetErrorLevelInfoData
End Property


Public Property Get ErrorName() As String
    ErrorName = this.ErrorName
End Property


Public Property Get Prefix() As String
    Prefix = this.Prefix
End Property


Public Property Get ErrorColor() As String
    ErrorColor = this.ErrorColor
End Property


Public Property Get MsgBoxStyle() As VbMsgBoxStyle
    MsgBoxStyle = this.MsgBoxStyle
End Property


Private Function SetErrorLevelInfoData()

    With this
        .ErrorName = GetErrorName(this.ErrorLevel)
        .ErrorColor = getErrorColor(this.ErrorLevel)
        .Prefix = .ErrorName & ": "
        .PrettyName = GetErrorNamePretty(this.ErrorLevel)
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


Private Function GetErrorName(ByVal ErrorLevelIn As eErrorLevel) As String

    Select Case ErrorLevelIn
    Case eelUninitialized:  GetErrorName = "ERROR Uninitialized"
    Case eelLoggingOff:     GetErrorName = "LOGGING OFF"
    Case eelNoError:        GetErrorName = "NO ERROR"
    Case eelDebugTrace:     GetErrorName = "TRACE"
    Case eelDebugEvent:     GetErrorName = "EVENT"
    Case eelDebugInfo:      GetErrorName = "INFO"
    Case eelWarning:        GetErrorName = "WARNING"
    Case eelError:          GetErrorName = "ERROR"
    Case eelCritical:       GetErrorName = "CRITICAL"
    End Select

End Function


Private Function GetErrorNamePretty(ByRef ErrorLevel As eErrorLevel) As String

    Select Case ErrorLevel
    Case eelUninitialized:  GetErrorNamePretty = "ERROR Uninitialized"
    Case eelLoggingOff:     GetErrorNamePretty = "Off"
    Case eelNoError:        GetErrorNamePretty = "Note"
    Case eelDebugTrace:     GetErrorNamePretty = "Debug Trace"
    Case eelDebugEvent:     GetErrorNamePretty = "Debug Event"
    Case eelDebugInfo:      GetErrorNamePretty = "Debug Info"
    Case eelWarning:        GetErrorNamePretty = "Warning"
    Case eelError:          GetErrorNamePretty = "Error"
    Case eelCritical:       GetErrorNamePretty = "Critical Error"
    End Select

End Function


' Checks the pretty version and the formal name of the error level and sets the error level
' based on the input.
Public Function SetErrorLevelFromName(ByRef strErrorName As String) As eErrorLevel

    Dim f_ErrorLevel As eErrorLevel

    For f_ErrorLevel = eelNoError To eErrorLevel.[_Last]
        If strErrorName = GetErrorName(f_ErrorLevel) Then
            SetErrorLevelFromName = f_ErrorLevel
            this.ErrorLevel = f_ErrorLevel
            Exit For
        End If
        ' Also check the "Pretty" version as well.
        If strErrorName = GetErrorNamePretty(f_ErrorLevel) Then
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
