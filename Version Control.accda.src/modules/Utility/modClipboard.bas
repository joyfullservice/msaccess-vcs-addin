Attribute VB_Name = "modClipboard"
'---------------------------------------------------------------------------------------
' Module    : modClipboard
' Author    : Adam Waller
' Date      : 11/20/2026
' Purpose   : Read and write text on the Windows clipboard through the Win32 API.
'           : The MSForms DataObject is not used because it renders its data on
'           : demand from a COM object released when the call returns, so readers
'           : racing that release see stale or empty text.
' Layer     : Utility
' Depends on: modFunctions (Pause)
'---------------------------------------------------------------------------------------
Option Compare Database
Option Private Module
Option Explicit
'@Folder("Utility")


Private Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hWndNewOwner As LongPtr) As Long
Private Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
Private Declare PtrSafe Function EmptyClipboard Lib "user32" () As Long
Private Declare PtrSafe Function GetClipboardData Lib "user32" (ByVal uFormat As Long) As LongPtr
Private Declare PtrSafe Function SetClipboardData Lib "user32" (ByVal uFormat As Long, ByVal hMem As LongPtr) As LongPtr
Private Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal dwBytes As LongPtr) As LongPtr
Private Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
Private Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As LongPtr) As Long
Private Declare PtrSafe Function GlobalFree Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
Private Declare PtrSafe Function lstrlenW Lib "kernel32" (ByVal lpString As LongPtr) As Long
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
    (ByVal Destination As LongPtr, ByVal Source As LongPtr, ByVal Length As LongPtr)

Private Const CF_UNICODETEXT As Long = 13
Private Const GHND As Long = &H42   ' GMEM_MOVEABLE Or GMEM_ZEROINIT

' Clipboard managers and clipboard history open the clipboard right after every
' change, so OpenClipboard briefly fails for everyone else.
Private Const CLIPBOARD_ATTEMPTS As Integer = 10
Private Const CLIPBOARD_RETRY_SECONDS As Single = 0.02


'---------------------------------------------------------------------------------------
' Procedure : SetClipboardText
' Author    : Adam Waller
' Date      : 5/18/2026
' Purpose   : Copy a text string to the Windows clipboard. Returns False if the
'           : clipboard stayed locked by another process.
'---------------------------------------------------------------------------------------
'
Public Function SetClipboardText(strText As String) As Boolean

    Dim lngBytes As Long
    Dim hMem As LongPtr
    Dim lngPtr As LongPtr

    ' Zero-initialized, so the terminating null is already in place.
    lngBytes = Len(strText) * 2
    hMem = GlobalAlloc(GHND, lngBytes + 2)
    If hMem = 0 Then Exit Function
    lngPtr = GlobalLock(hMem)
    If lngPtr = 0 Then
        GlobalFree hMem
        Exit Function
    End If
    ' StrPtr is 0 for vbNullString, so never copy from an empty string.
    If lngBytes > 0 Then CopyMemory lngPtr, StrPtr(strText), lngBytes
    GlobalUnlock hMem

    If OpenClipboardWithRetry Then
        If EmptyClipboard <> 0 Then
            ' On success the system owns hMem.
            SetClipboardText = (SetClipboardData(CF_UNICODETEXT, hMem) <> 0)
        End If
        CloseClipboard
    End If
    If Not SetClipboardText Then GlobalFree hMem

End Function


'---------------------------------------------------------------------------------------
' Procedure : GetClipboardText
' Author    : Adam Waller
' Date      : 5/18/2026
' Purpose   : Return the current clipboard text, or vbNullString if the clipboard
'           : is empty, locked, or contains non-text content (image, file, etc.).
'---------------------------------------------------------------------------------------
'
Public Function GetClipboardText() As String

    Dim hMem As LongPtr
    Dim lngPtr As LongPtr
    Dim lngLen As Long
    Dim strText As String

    If Not OpenClipboardWithRetry Then Exit Function
    hMem = GetClipboardData(CF_UNICODETEXT)
    If hMem <> 0 Then
        lngPtr = GlobalLock(hMem)
        If lngPtr <> 0 Then
            lngLen = lstrlenW(lngPtr)
            If lngLen > 0 Then
                strText = String$(lngLen, vbNullChar)
                CopyMemory StrPtr(strText), lngPtr, lngLen * 2
            End If
            GlobalUnlock hMem
        End If
    End If
    CloseClipboard
    GetClipboardText = strText

End Function


'---------------------------------------------------------------------------------------
' Procedure : OpenClipboardWithRetry
' Author    : Adam Waller
' Date      : 11/20/2026
' Purpose   : Open the clipboard, waiting briefly while another process holds it.
'           : The owner window must not be null, or SetClipboardData fails.
'---------------------------------------------------------------------------------------
'
Private Function OpenClipboardWithRetry() As Boolean
    Dim intAttempt As Integer
    For intAttempt = 1 To CLIPBOARD_ATTEMPTS
        If OpenClipboard(Application.hWndAccessApp) <> 0 Then
            OpenClipboardWithRetry = True
            Exit Function
        End If
        Pause CLIPBOARD_RETRY_SECONDS
    Next intAttempt
End Function
