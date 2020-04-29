Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Public PadLength As Integer

Private Const cstrSpacer As String = "-------------------------------------"

Private m_Log As clsConcat      ' Log file output
Private m_Console As clsConcat  ' Console output
Private m_RichText As TextBox   ' Text box to display HTML
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
    Set m_Console = New clsConcat
    Set m_Log = New clsConcat
    m_blnProgressActive = False
    m_sngLastUpdate = 0
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

    Dim strLine As String
    Dim strHtml As String
    
    ' Add to log file output
    m_Log.Add strText
    
    ' See if we want to print the output of this message.
    If blnPrint Then
        ' Remove existing progress indicator if in use.
        If m_blnProgressActive Then RemoveProgressIndicator
    
        ' Use bold/green text for completion line.
        strHtml = Replace(strText, " ", "&nbsp;")
        If InStr(1, strText, "Done. ") = 1 Then
            strHtml = "<font color=green><strong>" & strText & "</strong></font>"
        End If
        m_Console.Add strHtml
        ' Add line break for HTML
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
        
        ' Allow an update to the screen every second.
        ' (This keeps the aplication from an apparent hang while
        '  running intensive export processes.)
        If m_sngLastUpdate + 1 < Timer Then
            DoEvents
            m_sngLastUpdate = Timer
            Debug.Print Timer
        End If
    End If
    
    ' Add carriage return to log file if specified
    If blnNextOutputOnNewLine Then m_Log.Add vbCrLf
    
    ' Update log display on form if open.
    If blnPrint And IsLoaded(acForm, "frmMain") Then
        With Form_frmMain.txtLog
            m_blnProgressActive = False
            ' Set value, not text to avoid errors with large text strings.
            .SelStart = Len(.Text & vbNullString)
            Echo False
            .Value = m_Console.GetStr
            .SelStart = Len(.Text & vbNullString)
            Echo True
        End With
    End If
    
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
    m_RichText.AllowAutoCorrect = False
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
' Procedure : RemoveProgressIndicator
' Author    : Adam Waller
' Date      : 4/28/2020
' Purpose   : Remove the progress indicator if found at the end of the console output.
'---------------------------------------------------------------------------------------
'
Private Sub RemoveProgressIndicator()
    m_Console.Remove 2 + 4 ' (For unicode) plus <br>
    m_blnProgressActive = False
End Sub


'---------------------------------------------------------------------------------------
' Procedure : Increment
' Author    : Adam Waller
' Date      : 4/28/2020
' Purpose   : Increment the clock icon
'---------------------------------------------------------------------------------------
'
Public Sub Increment()

    ' Ongoing progress of clock
    Static intProgress As Integer
    
    ' Track the last time we did an increment
    Static sngLastIncrement As Single
    
    Dim strClock As String
    
    ' Ignore if we are not using the form
    If m_RichText Is Nothing Then Exit Sub
    
    ' Don't run the incrementer unless it has been 1
    ' second since the last displayed output refresh.
    If m_sngLastUpdate > Timer - 1 Then Exit Sub
    
    ' Allow an update to the screen every x seconds.
    ' Find the balance between good progress feedback
    ' without slowing down the overall export time.
    If sngLastIncrement > Timer - 0.2 Then Exit Sub

    ' Check the current status.
    If m_blnProgressActive Then
        ' Remove any existing character
        RemoveProgressIndicator
    Else
        ' Restart progress indicator when activating.
        intProgress = 11
    End If
        
    ' Rotate through the hours
    intProgress = intProgress + 1
    If intProgress = 13 Then intProgress = 1
    
    ' Status is now active
    m_blnProgressActive = True
    
    ' Set clock characters 1-12
    ' https://www.fileformat.info/info/unicode/char/1f552/index.htm
    strClock = ChrW(55357) & ChrW(56655 + intProgress)
    m_Console.Add strClock
    m_Console.Add "<br>"
    
    ' Update the log display
    With m_RichText
        Echo False
        .Value = m_Console.GetStr
        .SelStart = Len(.Text & vbNullString)
        Echo True
    End With
    sngLastIncrement = Timer
    DoEvents
    
End Sub