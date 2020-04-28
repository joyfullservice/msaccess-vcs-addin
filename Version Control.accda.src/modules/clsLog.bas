Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Public PadLength As Integer

Private Const cstrSpacer As String = "---------------------------------------"

Private m_Log As clsConcat      ' Log file output
Private m_Console As clsConcat  ' Console output
Private m_RichText As TextBox   ' Text box to display HTML


'---------------------------------------------------------------------------------------
' Procedure : Clear
' Author    : Adam Waller
' Date      : 4/16/2020
' Purpose   : Clear the log buffers
'---------------------------------------------------------------------------------------
'
Public Sub Clear()
    Set m_Console = New clsConcat
    Set m_Log = New clsConcat
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
Public Sub Add(strText As String, Optional blnPrint As Boolean = True, Optional blnNextOutputOnNewLine As Boolean = True)

    Static dblLastLog As Double
    Dim strLine As String
    
    m_Log.Add strText
    If blnPrint Then
        ' Use bold/green text for completion line.
        strLine = strText
        If InStr(1, strText, "Done. ") = 1 Then
            strLine = "<font color=green><strong>" & strText & "</strong></font>"
        End If
        m_Console.Add strLine
        If blnNextOutputOnNewLine Then m_Console.Add "<br>"
        ' Only print debug output if not running from the GUI.
        If Not IsLoaded(acForm, "frmMain") Then
            If blnNextOutputOnNewLine Then
                ' Create new line
                Debug.Print strText
                
            Else
                ' Continue next printout on this line.
                Debug.Print strText;
            End If
        End If
    End If
    
    ' Add carriage return to log file if specified
    If blnNextOutputOnNewLine Then m_Log.Add vbCrLf
    
    ' Allow an update to the screen every second.
    ' (This keeps the aplication from an apparent hang while
    '  running intensive export processes.)
    If dblLastLog + 1 < Timer Then
        DoEvents
        dblLastLog = Timer
    End If
    
    ' Update log display on form if open.
    If blnPrint And IsLoaded(acForm, "frmMain") Then
        With Form_frmMain.txtLog
            .Text = m_Console.GetStr
            ' Move cursor to end of log for scroll effect.
            .SelStart = Len(.Text)
            .SelLength = 0
        End With
    End If
    
End Sub


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
End Sub


'---------------------------------------------------------------------------------------
' Procedure : Class_Initialize
' Author    : Adam Waller
' Date      : 4/28/2020
' Purpose   : Set initial options
'---------------------------------------------------------------------------------------
'
Private Sub Class_Initialize()
    Me.PadLength = 30
    Set m_Log = New clsConcat
    Set m_Console = New clsConcat
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