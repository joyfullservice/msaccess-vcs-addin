VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsErrorInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : clsErrorInfo
' Author    : Josef Pötzl
' Date      : 08/25/2026
' Purpose   : Encapsulates error information and provides standardized
'           : error-level-specific display data for logging and user notifications.
'---------------------------------------------------------------------------------------
Option Compare Text
Option Explicit
'@Folder("Infrastructure")

' Error levels used for logging and monitoring the status
' of the current operation.
Public Enum eErrorLevel
    eelNoError
    eelWarning      ' Logged to file
    eelError        ' Displayed and logged
    eelCritical     ' Cancel operation
End Enum

Private Const ModuleName As String = "clsErrorInfo"

Private m_ErrorDataInitialized As Boolean

' Input data
Private m_ErrorLevel As eErrorLevel
Private m_Description As String
Private m_Source As String
Private m_IncludeErrorMessage As Boolean

'Error data
Private m_ErrNum As Long
Private m_ErrDesc As String

' Output data
Private m_Prefix As String
Private m_ErrorLevelText As String
Private m_MsgBoxStyle As VbMsgBoxStyle

Public Sub Init(eLevel As eErrorLevel, strDescription As String, Optional strSource As String, Optional blnIncludeErrorMessage As Boolean = False)

    m_ErrorLevel = eLevel
    m_Description = strDescription
    m_Source = strSource
    m_IncludeErrorMessage = blnIncludeErrorMessage

    m_ErrNum = Err.Number
    m_ErrDesc = Err.Description

    SetErrorLevelInfoData
    m_ErrorDataInitialized = True

End Sub

Public Property Get ErrorNumber() As Long
    ErrorNumber = m_ErrNum
End Property

Public Property Get ErrorDescription() As String
    ErrorDescription = m_ErrDesc
End Property

Public Property Get ErrorLevel() As eErrorLevel
    If Not m_ErrorDataInitialized Then
        Err.Raise vbObjectError, ModuleName & ".MsgBoxStyle", "ErrorLevel was not set."
    End If
    ErrorLevel = m_ErrorLevel
End Property

Public Property Get Description() As String
    Description = m_Description
End Property

Public Property Get Source() As String
    Source = m_Source
End Property

Public Property Get DisplayText() As String

    If Not m_ErrorDataInitialized Then
        Err.Raise vbObjectError, ModuleName & ".MsgBoxStyle", "ErrorLevel was not set."
    End If

    ' Sometimes the error description is sufficient for the description
    If m_Description = vbNullString And m_ErrNum <> 0 Then
        DisplayText = m_Prefix & Err.Number & " - " & Err.Description
    ElseIf m_IncludeErrorMessage Then
        DisplayText = m_Prefix & m_Description & "; Error " & m_ErrNum & " - " & m_ErrDesc
    Else
        DisplayText = m_Prefix & m_Description
    End If

End Property

Public Property Get ErrorLevelText() As String
    If Not m_ErrorDataInitialized Then
        Err.Raise vbObjectError, ModuleName & ".MsgBoxStyle", "ErrorLevel was not set."
    End If
    ErrorLevelText = m_ErrorLevelText
End Property

Public Property Get MsgBoxStyle() As VbMsgBoxStyle
    If Not m_ErrorDataInitialized Then
        Err.Raise vbObjectError, ModuleName & ".MsgBoxStyle", "ErrorLevel was not set."
    End If
    MsgBoxStyle = m_MsgBoxStyle
End Property

'---------------------------------------------------------------------------------------
' Procedure : SetErrorLevelInfoData
' Author    : Josef Pötzl
' Date      : 08/25/2026
' Purpose   : Sets the prefix, error level text, and message box style
'           : based on the specified error level.
'---------------------------------------------------------------------------------------
'
Private Function SetErrorLevelInfoData()

    Select Case m_ErrorLevel

        Case eErrorLevel.eelWarning
            m_Prefix = "WARNING: "
            m_ErrorLevelText = "Warning"
            m_MsgBoxStyle = vbInformation

        Case eErrorLevel.eelError
            m_Prefix = "ERROR: "
            m_ErrorLevelText = "Error"
            m_MsgBoxStyle = vbExclamation

        Case eErrorLevel.eelCritical
            m_Prefix = "CRITICAL: "
            m_ErrorLevelText = "Critical"
            m_MsgBoxStyle = vbCritical

        Case eErrorLevel.eelNoError:
            m_Prefix = vbNullString
            m_ErrorLevelText = vbNullString
            m_MsgBoxStyle = 0

        Case Else
            ' Raise error: "ErrorLevel {0} not supported"?
            m_Prefix = vbNullString
            m_ErrorLevelText = vbNullString
            m_MsgBoxStyle = vbExclamation

    End Select

End Function
