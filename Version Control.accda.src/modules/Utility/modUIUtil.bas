Attribute VB_Name = "modUIUtil"
'---------------------------------------------------------------------------------------
' Module    : modUIUtil
' Author    : Adam Waller
' Date      : 12/4/2020
' Purpose   : UI and dialog helpers: enhanced message box, IDE display, resizable
'           : dialogs, and datasheet column scaling.
' Layer     : Utility
' Depends on: modObjects, clsOperation (InteractionMode)
'---------------------------------------------------------------------------------------
Option Compare Database
Option Private Module
Option Explicit
'@Folder("Utility")

' API calls to change window style
Private Const GWL_STYLE = -16
Private Const WS_SIZEBOX = &H40000
Private Declare PtrSafe Function IsWindowUnicode Lib "user32" (ByVal hwnd As LongPtr) As Long
#If Win64 Then
    ' 64-bit versions of Access
    Private Declare PtrSafe Function GetWindowLongPtr Lib "user32" Alias "GetWindowLongPtrW" (ByVal hwnd As LongPtr, ByVal nIndex As Long) As LongPtr
    Private Declare PtrSafe Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongPtrW" (ByVal hwnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
#Else
    ' 32-bit versions of Access
    Private Declare PtrSafe Function GetWindowLongPtr Lib "user32" Alias "GetWindowLongW" (ByVal hwnd As LongPtr, ByVal nIndex As Long) As LongPtr
    Private Declare PtrSafe Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongW" (ByVal hwnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
#End If


'---------------------------------------------------------------------------------------
' Procedure : ShowIDE
' Author    : Adam Waller
' Date      : 5/18/2015
' Purpose   : Show the VBA code editor (used in autoexec macro)
'---------------------------------------------------------------------------------------
'
Public Function ShowIDE() As Boolean
    DoCmd.RunCommand acCmdVisualBasicEditor
    DoEvents
    ShowIDE = True
End Function


'---------------------------------------------------------------------------------------
' Procedure : MsgBox2
' Author    : Adam Waller
' Date      : 1/27/2017
' Purpose   : Alternate message box with bold prompt on first line.
'---------------------------------------------------------------------------------------
'
Public Function MsgBox2(strBold As String, Optional strLine1 As String, Optional strLine2 As String, _
    Optional intButtons As VbMsgBoxStyle = vbOKOnly, Optional strTitle As String, Optional intDefaultResult As VbMsgBoxResult = vbOK) As VbMsgBoxResult

    Dim strMsg As String
    Dim varLines(0 To 3) As String
    Dim intCursor As Integer

    ' Turn off any hourglass
    intCursor = Screen.MousePointer
    If intCursor > 0 Then Screen.MousePointer = 0

    ' Escape single quotes by doubling them.
    varLines(0) = Replace(strBold, "'", "''")
    varLines(1) = Replace(strLine1, "'", "''")
    varLines(2) = Replace(strLine2, "'", "''")
    varLines(3) = Replace(strTitle, "'", "''")

    ' Check interaction mode (lives on the Operation singleton)
    If Operation.InteractionMode = eimNormal Then
        ' Normal user interaction with MsgBox
        If varLines(3) = vbNullString Then varLines(3) = T("Version Control Add-in")
        strMsg = "MsgBox('" & varLines(0) & "@" & varLines(1) & "@" & varLines(2) & "@'," & intButtons & ",'" & varLines(3) & "','',0)"
        Perf.PauseTiming
        MsgBox2 = Eval(strMsg)
        Perf.ResumeTiming
    Else
        ' Silent mode. Don't display any message, but log it instead.
        With New clsConcat
            .AppendOnAdd = vbCrLf
            .Add "[**MessageBox Not Displayed**]"
            If Len(strTitle) Then .Add "Title: " & strTitle
            If Len(strBold) Then .Add strBold
            If Len(strLine1) Then .Add strLine1
            If Len(strLine2) Then .Add strLine2
            If intButtons <> vbOKOnly Then .Add "Buttons Flag: " & intButtons
            Log.Add .GetStr
        End With
        ' Return default (unattended) result
        MsgBox2 = intDefaultResult
    End If

    ' Restore MousePointer (if needed)
    If intCursor > 0 Then Screen.MousePointer = intCursor

End Function


'---------------------------------------------------------------------------------------
' Procedure : MakeDialogResizable
' Author    : Adam Waller
' Date      : 5/16/2023
' Purpose   : Change the window style of an existing dialog window to make it resizable.
'           : (This allows you to use the acDialog argument when opening a form, but
'           :  still have the form resizable by the user.)
'           : NOTE: As of 5/6/2026, this no longer works on newer versions of Windows/Access.
'           : The Win32 style change is applied but the resize behavior is not honored.
'           : Leaving in place for older platform compatibility.
'---------------------------------------------------------------------------------------
'
Public Sub MakeDialogResizable(frmMe As Form)

    Dim lngHwnd As LongPtr
    Dim lngFlags As LongPtr
    Dim lngResult As LongPtr

    ' Get handle for form
    lngHwnd = frmMe.hwnd

    ' Debug.Print IsWindowUnicode(lngHwnd) - Testing indicates that the windows are
    ' Unicode, so we are using the Unicode versions of the GetWindowLong functions.

    ' Get the current window style
    lngFlags = GetWindowLongPtr(lngHwnd, GWL_STYLE)

    ' Set resizable flag and apply updated style
    lngFlags = lngFlags Or WS_SIZEBOX
    lngResult = SetWindowLongPtr(lngHwnd, GWL_STYLE, lngFlags)

End Sub


'---------------------------------------------------------------------------------------
' Procedure : ScaleColumns
' Author    : Adam Waller
' Date      : 5/16/2023
' Purpose   : Size the datasheet columns evenly to fill the available width, minus an
'           : allotment for the width of the vertical scroll bar.
'---------------------------------------------------------------------------------------
'
Public Sub ScaleColumns(frmDatasheet As Form, Optional lngScrollWidthTwips As Long = 300, _
    Optional varFixedControlNameArray As Variant)

    Dim lngTotal As Long
    Dim lngCurrent As Long
    Dim lngSizeable As Long
    Dim lngFixed As Long
    Dim lngWidth As Long
    Dim dblRatio As Double
    Dim ctl As Control
    Dim colResize As Collection

    lngTotal = frmDatasheet.InsideWidth - lngScrollWidthTwips
    Set colResize = New Collection

    ' Loop through the columns twice, once to get the current widths, then to set them.
    For Each ctl In frmDatasheet.Controls
        Select Case ctl.ControlType
            Case acTextBox, acComboBox
                If ctl.Visible Then
                    ' Get column width
                    lngWidth = ctl.ColumnWidth
                    If lngWidth < 0 Then
                        ' Set to not hidden to get the actual width of the column
                        ' -1 = Default Width
                        ' -2 = Fit to Text
                        ctl.ColumnHidden = False
                        lngWidth = ctl.ColumnWidth
                    End If
                    lngCurrent = lngCurrent + lngWidth
                    If Not InArray(varFixedControlNameArray, ctl.Name, vbTextCompare) Then
                        lngSizeable = lngSizeable + lngWidth
                        colResize.Add ctl
                    End If
                End If
        End Select
    Next ctl

    ' Exit if we have no sizable controls
    If lngSizeable = 0 Then Exit Sub

    ' Get ratio for new sizes (Scales resizable controls proportionately)
    lngFixed = lngCurrent - lngSizeable
    dblRatio = (lngTotal - lngFixed) / lngSizeable

    ' Resize each control
    For Each ctl In colResize
        ctl.ColumnWidth = ctl.ColumnWidth * dblRatio
    Next ctl

End Sub
