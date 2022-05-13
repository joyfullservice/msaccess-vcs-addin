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

Public PadLength As Integer
Public LogFilePath As String
Public ErrorLevel As eErrorLevel
Public OperationType As eOperationType

' Set this to true when logging an operation such as an export or build
' then set back to false after writing the log file. This affects
' how error messages are reported to the user outside of operations.
Public Active As Boolean

Private Const cstrSpacer As String = "-------------------------------------"

Private m_Log As clsConcat      ' Log file output
Private m_Console As clsConcat  ' Console output
Private m_RichText As TextBox   ' Text box to display HTML
Private m_Prog As clsLblProg    ' Progress bar
Private m_blnProgressActive As Boolean
Private m_sngLastUpdate As Single


'---------------------------------------------------------------------------------------
' Procedure : Clear
' Author    : Adam Waller
' Date      : 4/16/2020
' Purpose   : Clear the log buffers
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
Public Sub Spacer(Optional blnPrint As Boolean = True)
    Me.Add cstrSpacer, blnPrint
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
    Optional blnBold As Boolean = False)
    
    ' Add to log file output
    m_Log.Add strText, vbCrLf
    
    ' See if we want to print the output of this message.
    If blnPrint Then
        ' Remove existing progress indicator if in use.
        If m_blnProgressActive Then
            m_blnProgressActive = False
            m_Prog.Hide
        End If
    
        ' Build HTML output for console
        With m_Console
        
            ' Opening tags
            If blnBold Then .Add "<strong>"
            If strColor <> vbNullString Then .Add "<font color=", strColor, ">"
            
            ' Content
            .Add MultiReplace(strText, _
                " ", "&nbsp;", _
                vbCrLf, "<br>")
            
            ' Closing tags
            If strColor <> vbNullString Then .Add "</font>"
            If blnBold Then m_Console.Add "</strong>"
            
            ' Add line break for HTML
            If blnNextOutputOnNewLine Then m_Console.Add "<br>"
        End With
        
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

    ' See if the GUI form is loaded.
    Perf.OperationStart "Console Updates"
    If Not m_RichText Is Nothing Then
        With m_RichText
            m_blnProgressActive = False
            If Not m_Prog Is Nothing Then m_Prog.Hide
            ' Set value, not text to avoid errors with large text strings.
            Echo False
            '.SelStart = Len(.Text & vbNullString)
            ' Show the last 20K characters so
            ' we don't hit the Integer limit
            ' on the SelStart property.
            .Value = m_Console.RightStr(20000)
            If VerifyFocus(m_RichText) Then .SelStart = 20000
            Echo True
            'Form_frmVCSMain.Repaint
        End With
    End If
    
    ' Update the display (especially for immediate window)
    DoEvents
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
Public Sub Error(eLevel As eErrorLevel, strDescription As String, Optional strSource As String)

    Dim strPrefix As String
    Dim strDisplay As String
    
    Select Case eLevel
        Case eelWarning:    strPrefix = "WARNING: "
        Case eelError:      strPrefix = "ERROR: "
        Case eelCritical:   strPrefix = "CRITICAL: "
    End Select
    
    ' Build the error message string.
    With New clsConcat
    
        ' Sometimes the error description is sufficient for the description
        If strDescription = vbNullString And Err.Number <> 0 Then
            strDisplay = strPrefix & Err.Description
        Else
            strDisplay = strPrefix & strDescription
        End If
        
        ' Display on the output screen anything higher than a warning
        If eLevel > eelWarning Then
            Me.Add vbNullString
            Me.Spacer
            Me.Add strDisplay, , , "red"
            Me.Spacer
        End If
        
        ' Log the full detail to the log file
        If Err Then .Add "Error ", Err.Number, ": ", Err.Description, " "
        If strSource <> vbNullString Then .Add "Source: ", strSource
        Me.Add .GetStr, False
        
        ' See if we are actively logging an operation
        If Log.Active Then
            ' Show message box for fatal error.
            If eLevel = eelCritical Then
                MsgBox2 "Unable to Continue", .GetStr, _
                    "Please review the log file for additional details.", vbCritical
            End If
        Else
            ' Show message on any error level when we are not logging to a file.
            Select Case eLevel
                Case eelNoError:    ' Do nothing
                Case eelWarning:    MsgBox2 "Warning", strDisplay, .GetStr, vbInformation
                Case eelError:      MsgBox2 "Error", strDisplay, .GetStr, vbExclamation
                Case eelCritical:   MsgBox2 "Critical", strDisplay, .GetStr, vbCritical
            End Select
        End If
    End With
    
    ' Update error level if higher.
    If Me.ErrorLevel < eLevel Then Me.ErrorLevel = eLevel
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : SetConsole
' Author    : Adam Waller
' Date      : 4/28/2020
' Purpose   : Set a reference to the text box.
'---------------------------------------------------------------------------------------
'
Public Sub SetConsole(txtRichText As TextBox)
    Set m_RichText = txtRichText
    If Not m_RichText Is Nothing Then
        m_RichText.AllowAutoCorrect = False
    End If
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
' Purpose   : Saves the log data to a file, and resets the log buffer.
'---------------------------------------------------------------------------------------
'
Public Sub SaveFile(strPath As String)
    WriteFile m_Log.GetStr, strPath
    Set m_Log = New clsConcat
    LogFilePath = strPath
End Sub


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
    m_blnProgressActive = False
    m_sngLastUpdate = 0
    Me.PadLength = 30
    Me.ErrorLevel = eelNoError
    Me.Active = False
    Me.LogFilePath = vbNullString
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
' Purpose   : Increment the progress bar
'---------------------------------------------------------------------------------------
'
Public Sub Increment()

    ' Ignore if we are not using the form
    If m_Prog Is Nothing Then Exit Sub
    ' Check the current status.
    Perf.OperationStart "Increment Progress"
    If Not m_blnProgressActive Then
        ' Flush any pending output
        With m_RichText
            Echo False
            ' Show the last 20K characters so
            ' we don't hit the Integer limit
            ' on the SelStart property.
            .Value = m_Console.RightStr(20000)
            If VerifyFocus(m_RichText) Then .SelStart = 20000
            Echo True
        End With
    End If
    Me.ProgressBar.Increment

    ' Status is now active
    m_blnProgressActive = True
    Perf.OperationEnd
End Sub
