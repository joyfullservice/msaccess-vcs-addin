Attribute VB_Name = "modFileWinAPI"
'---------------------------------------------------------------------------------------
' Module    : modFileScan
' Author    : Adam Waller
' Date      : 7/25/2023
' Purpose   : Functions for extremely fast file system scan utilizing the Windows API.
'           : Other functions to accurately return/set file modified dates with dates
'           : that correctly convert for time zone/daylight savings time for historical
'           : files in other years.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Private Module
Option Explicit


Private Declare PtrSafe Function FindFirstFileW Lib "kernel32" (ByVal lpFileName As LongPtr, ByVal lpFindFileData As LongPtr) As LongPtr
Private Declare PtrSafe Function FindNextFileW Lib "kernel32" (ByVal hFindFile As LongPtr, ByVal lpFindFileData As LongPtr) As LongPtr
Private Declare PtrSafe Function FindClose Lib "kernel32" (ByVal hFindFile As LongPtr) As Long
Private Declare PtrSafe Function FileTimeToLocalFileTime Lib "kernel32" (lpFileTime As FILETIME, lpLocalFileTime As FILETIME) As LongPtr
Private Declare PtrSafe Sub GetSystemTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)
Private Declare PtrSafe Sub GetLocalTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)

' Time zone conversions
Private Declare PtrSafe Function GetTimeZoneInformation Lib "kernel32" (lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long
Private Declare PtrSafe Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long
Private Declare PtrSafe Function SystemTimeToFileTime Lib "kernel32" (lpSystemTime As SYSTEMTIME, lpFileTime As FILETIME) As Long
Private Declare PtrSafe Function TzSpecificLocalTimeToSystemTime Lib "kernel32" (lpTimeZoneInformation As TIME_ZONE_INFORMATION, lpLocalTime As SYSTEMTIME, lpUniversalTime As SYSTEMTIME) As LongPtr
Private Declare PtrSafe Function SystemTimeToTzSpecificLocalTime Lib "kernel32" (lpTimeZoneInformation As TIME_ZONE_INFORMATION, lpUniversalTime As SYSTEMTIME, lpLocalTime As SYSTEMTIME) As LongPtr

' Set file time
Private Declare PtrSafe Function GetFileTime Lib "kernel32" (ByVal hFile As LongPtr, lpCreationTime As FILETIME, lpLastAccessTime As FILETIME, lpLastWriteTime As FILETIME) As Long
Private Declare PtrSafe Function SetFileTime Lib "kernel32" (ByVal hFile As LongPtr, lpCreationTime As FILETIME, lpLastAccessTime As FILETIME, lpLastWriteTime As FILETIME) As Long
Declare PtrSafe Function CloseHandle Lib "kernel32" (ByVal hObject As LongPtr) As Long

'lpSecurityAttributes As SECURITY_ATTRIBUTES,
Private Declare PtrSafe Function CreateFile Lib "kernel32" Alias "CreateFileA" ( _
    ByVal lpFileName As String, _
    ByVal dwDesiredAccess As Long, _
    ByVal dwShareMode As Long, _
    lpSecurityAttributes As Any, _
    ByVal dwCreationDisposition As Long, _
    ByVal dwFlagsAndAttributes As Long, _
    ByVal hTemplateFile As LongPtr) As LongPtr

' Constants for CreateFile (used when changing modified date)
Private Const OPEN_EXISTING = &H3
Private Const FILE_SHARE_READ = &H1
Private Const FILE_SHARE_WRITE = &H2
Private Const CREATE_ALWAYS = &H2
Private Const OPEN_ALWAYS = &H4
Private Const INVALID_HANDLE_VALUE = -1
Private Const ERROR_ALREADY_EXISTS = &HB7
Private Const GENERIC_ALL = &H10000000
Private Const GENERIC_EXECUTE = &H20000000
Private Const GENERIC_READ = &H80000000
Private Const GENERIC_WRITE = &H40000000

' Other constants
Private Const MAX_PATH  As Long = 260
Private Const ALTERNATE As Long = 14
Private Const FILE_ATTRIBUTE_DIRECTORY As Long = 16 '0x10

Private Type FILETIME
    dwLowDateTime  As Long
    dwHighDateTime As Long
End Type

Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As LongPtr
    bInheritHandle As Long
End Type

Private Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

Private Type TIME_ZONE_INFORMATION
    Bias As Long
    StandardName(0 To 31) As Integer
    StandardDate As SYSTEMTIME
    StandardBias As Long
    DaylightName(0 To 31) As Integer
    DaylightDate As SYSTEMTIME
    DaylightBias As Long
End Type

Private Enum TIME_ZONE
    TIME_ZONE_ID_INVALID = 0
    TIME_ZONE_STANDARD = 1
    TIME_ZONE_DAYLIGHT = 2
End Enum

' Can be used with either W or A functions
' Pass VarPtr(wfd) to W or simply wfd to A
Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime   As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime  As FILETIME
    nFileSizeHigh    As Long
    nFileSizeLow     As Long
    dwReserved0      As Long
    dwReserved1      As Long
    cFileName        As String * MAX_PATH
    cAlternate       As String * ALTERNATE
End Type


'---------------------------------------------------------------------------------------
' Procedure : GetFileList
' Author    : Adam Waller
' Date      : 7/25/2023
' Purpose   : Return a list of files from the specified folder. Returns a dictionary
'           : with the filename as the key, and the modified date as the value.
'           : (Could be extended in the future to return other values)
'---------------------------------------------------------------------------------------
'
Public Function GetFileList(strFolder As String, Optional strPattern As String = "*.*", Optional blnAsLocalTime As Boolean = True) As Dictionary

    Dim dList As Dictionary
    Dim pFileHandle As LongPtr
    Dim strSearchPath As String
    Dim tFileData As WIN32_FIND_DATA
    Dim strName As String

    Perf.OperationStart "Get File Listing (API)"
    Set dList = New Dictionary

    ' Build full search path
    strSearchPath = AddSlash(strFolder) & strPattern

    ' Attempt to find first file
    pFileHandle = FindFirstFileW(StrPtr(strSearchPath), VarPtr(tFileData))
    If pFileHandle <> INVALID_HANDLE_VALUE Then
        Do
            ' Get file name from API call
            strName = Left$(tFileData.cFileName, InStr(tFileData.cFileName, vbNullChar) - 1)
            If strName = "." Or strName = ".." Then
                ' Skip meta directories
            ElseIf tFileData.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY Then
                ' Skip subfolders
            Else
                ' Save file to list
                dList.Add strName, FileTimeToDate(tFileData.ftLastWriteTime, blnAsLocalTime)
            End If
        Loop While FindNextFileW(pFileHandle, VarPtr(tFileData))
    End If

    ' Close handle
    FindClose pFileHandle
    Perf.OperationEnd

    ' Return listing of files
    Set GetFileList = dList

End Function


'---------------------------------------------------------------------------------------
' Procedure : SetFileDate
' Author    : Adam Waller
' Date      : 7/28/2023
' Purpose   : This is WAY more complicated than it might first appear. In Windows 7 and
'           : newer, Windows Explorer attempts to display file modified dates as
'           : relative to the Daylight Savings Time offset in effect at the time the
'           : file was modified. Setting a file date to match what you see in Windows
'           : explorer requires converting the local date/time to a UTC time using the
'           : same DST rules used by Windows. (Hence the additional API calls required
'           : to make this conversion.)
'           : Further Reading: https://stackoverflow.com/q/66615978/4121863
'---------------------------------------------------------------------------------------
'
Public Sub SetFileDate(strFile As String, dteDate As Date, blnAsLocalTime As Boolean)

    Dim lngHandle As LongPtr
    Dim stNewDate As SYSTEMTIME
    Dim stUtc As SYSTEMTIME
    Dim ftUtc As FILETIME
    Dim ftBlank As FILETIME
    Dim lngResult As LongPtr
    Dim strFullPath As String

    Perf.OperationStart "Set file modified date"

    ' Support long paths
    strFullPath = "\\?\" & strFile

    ' Don't attempt this if the file does not exist
    If Not FSO.FileExists(strFile) Then Exit Sub

    ' Open a handle to the existing file with write access
    lngHandle = CreateFile(strFullPath, GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0&, OPEN_EXISTING, 0&, 0&)
    If lngHandle = INVALID_HANDLE_VALUE Then
        'Debug.Print GetSystemErrorMessageText(Err.LastDllError)
        'Log.Error eelError, "Unable to set file date for " & strFile & ". (Unable to write to file)", ModuleName & "SetFileDate"
        Exit Sub
    End If

    ' Convert the date to a SYSTEMTIME
    stNewDate = DateToSystemTime(dteDate)

    ' See if we are converting this from a local time
    If blnAsLocalTime Then
        ' Convert to UTC using an API that is able to translate Timezone/DST to UTC
        ' This is SUPPOSED to default to the local TZ if null is provided, but this was not the case
        ' in my testing, so we are passing the current time zone information just to be safe.
        lngResult = TzSpecificLocalTimeToSystemTime(GetLocalTimeZoneInfo, stNewDate, stUtc)
        stNewDate = stUtc
    End If

    ' Convert the UTC system time to a FILETIME
    lngResult = SystemTimeToFileTime(stNewDate, ftUtc)

    ' Set the file date using the converted UTC time
    lngResult = SetFileTime(lngHandle, ftBlank, ftBlank, ftUtc)

    ' Close the file handle
    CloseHandle lngHandle

    Perf.OperationEnd

End Sub


'---------------------------------------------------------------------------------------
' Procedure : GetFileDateEx
' Author    : Adam Waller
' Date      : 7/28/2023
' Purpose   : Return the actual date displayed in Windows Explorer (DST aware for
'           : historical dates), not just the FSO LastModified date, which may not be
'           : accurate for dates outside the current DST settings.
'---------------------------------------------------------------------------------------
'
Public Function GetFileModifiedDateEx(strFile As String) As Date

End Function


'---------------------------------------------------------------------------------------
' Procedure : FileTimeToDate
' Author    : Adam Waller
' Date      : 7/25/2023
' Purpose   : Convert a Win32 API FileTime to a VBA Datetime value
'---------------------------------------------------------------------------------------
'
Public Function FileTimeToDate(tFileTime As FILETIME, blnAsLocalTime As Boolean) As Date

    Dim tReturnTime As SYSTEMTIME
    Dim tUtcTime As SYSTEMTIME
    Dim lngResult As LongPtr

    ' Get UTC file time
    FileTimeToSystemTime tFileTime, tUtcTime

    ' Perform local time conversion, if requested
    If blnAsLocalTime Then
        lngResult = SystemTimeToTzSpecificLocalTime(GetLocalTimeZoneInfo, tUtcTime, tReturnTime)
    Else
        tReturnTime = tUtcTime
    End If

    ' Convert to a VBA date value
    With tReturnTime
        FileTimeToDate = DateSerial(.wYear, .wMonth, .wDay) + TimeSerial(.wHour, .wMinute, .wSecond)
    End With

End Function


'---------------------------------------------------------------------------------------
' Procedure : DateToSystemTime
' Author    : Adam Waller
' Date      : 7/28/2023
' Purpose   : Convert a VBA date to a systemtime structure
'---------------------------------------------------------------------------------------
'
Private Function DateToSystemTime(dteDate) As SYSTEMTIME
    With DateToSystemTime
        .wYear = Year(dteDate)
        .wMonth = Month(dteDate)
        .wDay = Day(dteDate)
        .wDayOfWeek = Weekday(dteDate) - 1 ' Adjust to expected format
        .wHour = Hour(dteDate)
        .wMinute = Minute(dteDate)
        .wSecond = Second(dteDate)
        .wMilliseconds = 0  ' Not used with VBA dates
    End With
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetLocalTimeZoneInfo
' Author    : Adam Waller
' Date      : 7/28/2023
' Purpose   : A function to return a copy of the current time zone information
'           : (Cached for performance reasons)
'---------------------------------------------------------------------------------------
'
Private Function GetLocalTimeZoneInfo() As TIME_ZONE_INFORMATION
    Static blnCached As Boolean
    Static tzLocal As TIME_ZONE_INFORMATION
    If Not blnCached Then
        GetTimeZoneInformation tzLocal
        blnCached = True
    End If
    GetLocalTimeZoneInfo = tzLocal
End Function
