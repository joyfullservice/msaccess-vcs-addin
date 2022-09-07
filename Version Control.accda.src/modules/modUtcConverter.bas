Attribute VB_Name = "modUtcConverter"
Private Const ModuleName As String = "modUtcConverter"

''
' VBA-UTC v2.0.1
' (c) Tim Hall - https://github.com/VBA-tools/VBA-UtcConverter
' (c) hecon5 - 2022-08-30T16:00:20.540Z rewrites and updates.
' UTC/ISO 8601 Converter for VBA
'
' Errors:
' 10011 - UTC parsing error
' 10012 - UTC conversion error
' 10013 - ISO 8601 parsing error
' 10014 - ISO 8601 conversion error
'
' @module UtcConverter
' @author tim.hall.engr@gmail.com, hecon5
' @license MIT (http://www.opensource.org/licenses/mit-license.php)
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
Option Compare Text
Option Explicit

' Spec details which make parsing easier, instead of calling and / or doing math every time.
Private Const TotalHoursInDay As Double = 24
Private Const TotalMinutesInDay As Double = TotalHoursInDay * 60
Private Const TotalSecondsInDay As Double = TotalMinutesInDay * 60
Private Const TotalMillisecondsInDay As Double = TotalSecondsInDay * 1000


Private Const DecimalSeparator As String = "."
Private Const ISO8601DateDelimiter As String = "-"
Private Const ISO8601DateTimeSeparator As String = "T"
Private Const ISO8601TimeDelimiter As String = ":"
Private Const ISO8601UTCTimeZone As String = "Z"

#If Mac Then
#If VBA7 Then
' 64-bit Mac (2016)
Private Declare PtrSafe Function utc_popen Lib "/usr/lib/libc.dylib" Alias "popen" _
    (ByVal utc_Command As String, ByVal utc_Mode As String) As LongPtr
Private Declare PtrSafe Function utc_pclose Lib "/usr/lib/libc.dylib" Alias "pclose" _
    (ByVal utc_File As LongPtr) As LongPtr
Private Declare PtrSafe Function utc_fread Lib "/usr/lib/libc.dylib" Alias "fread" _
    (ByVal utc_Buffer As String, ByVal utc_Size As LongPtr, ByVal utc_Number As LongPtr, ByVal utc_File As LongPtr) As LongPtr
Private Declare PtrSafe Function utc_feof Lib "/usr/lib/libc.dylib" Alias "feof" _
    (ByVal utc_File As LongPtr) As LongPtr

#Else
' 32-bit Mac
Private Declare Function utc_popen Lib "libc.dylib" Alias "popen" _
    (ByVal utc_Command As String, ByVal utc_Mode As String) As Long
Private Declare Function utc_pclose Lib "libc.dylib" Alias "pclose" _
    (ByVal utc_File As Long) As Long
Private Declare Function utc_fread Lib "libc.dylib" Alias "fread" _
    (ByVal utc_Buffer As String, ByVal utc_Size As Long, ByVal utc_Number As Long, ByVal utc_File As Long) As Long
Private Declare Function utc_feof Lib "libc.dylib" Alias "feof" _
    (ByVal utc_File As Long) As Long

#End If
' End of Mac
#ElseIf VBA7 Then
' Windows VBA7

Private Declare PtrSafe Sub GetSystemTime Lib "kernel32" (lpSystemTime As utc_SYSTEMTIME)
Private Declare PtrSafe Sub GetLocalTime Lib "kernel32" (lpSystemTime As utc_SYSTEMTIME)

' http://msdn.microsoft.com/en-us/library/windows/desktop/ms724421.aspx
' http://msdn.microsoft.com/en-us/library/windows/desktop/ms724949.aspx
' http://msdn.microsoft.com/en-us/library/windows/desktop/ms725485.aspx
Private Declare PtrSafe Function utc_GetTimeZoneInformation Lib "kernel32" Alias "GetTimeZoneInformation" _
    (utc_lpTimeZoneInformation As utc_TIME_ZONE_INFORMATION) As Long
    
Private Declare PtrSafe Function utc_SystemTimeToTzSpecificLocalTime Lib "kernel32" Alias "SystemTimeToTzSpecificLocalTime" _
    (utc_lpTimeZoneInformation As utc_TIME_ZONE_INFORMATION, utc_lpUniversalTime As utc_SYSTEMTIME, utc_lpLocalTime As utc_SYSTEMTIME) As Long
Private Declare PtrSafe Function utc_TzSpecificLocalTimeToSystemTime Lib "kernel32" Alias "TzSpecificLocalTimeToSystemTime" _
    (utc_lpTimeZoneInformation As utc_TIME_ZONE_INFORMATION, utc_lpLocalTime As utc_SYSTEMTIME, utc_lpUniversalTime As utc_SYSTEMTIME) As Long

' Dynamic Functions allow for past Time Zones to be accounted for. Above will work for "now".
' https://docs.microsoft.com/en-us/windows/win32/api/timezoneapi/nf-timezoneapi-gettimezoneinformationforyear
' From docs: the wYear is LOCAL time, so if the year converts over, you need to check the following (or prior) year.
' to ensure you get the correct time zone detail.
' Word of warning: https://devblogs.microsoft.com/oldnewthing/20110311-00/?p=11243
Private Declare PtrSafe Function GetTimeZoneInformationForYear Lib "kernel32" ( _
    wYear As Integer _
    , ByRef lpDynamicTimeZoneInformation As DYNAMIC_TIME_ZONE_INFORMATION _
    , ByRef lpTimeZoneInformation As utc_TIME_ZONE_INFORMATION) As Long

Private Declare PtrSafe Function GetDynamicTimeZoneInformation Lib "kernel32" ( _
    ByRef pTimeZoneInformation As DYNAMIC_TIME_ZONE_INFORMATION) As Long

Private Declare PtrSafe Function SystemTimeToTzSpecificLocalTimeEx Lib "kernel32" ( _
    ByRef lpDynamicTimeZoneInformation As DYNAMIC_TIME_ZONE_INFORMATION _
    , ByRef lpUniversalTime As utc_SYSTEMTIME _
    , ByRef lpLocalTime As utc_SYSTEMTIME) As Long

Private Declare PtrSafe Function TzSpecificLocalTimeToSystemTimeEx Lib "kernel32" ( _
    ByRef lpDynamicTimeZoneInformation As DYNAMIC_TIME_ZONE_INFORMATION _
    , ByRef lpLocalTime As utc_SYSTEMTIME _
    , ByRef lpUniversalTime As utc_SYSTEMTIME) As Long
    
#Else
' VBA 6 or less.

Private Declare Function GetTimeZoneInformationForYear Lib "kernel32" ( _
    wYear As Integer, _
    lpDynamicTimeZoneInformation As DYNAMIC_TIME_ZONE_INFORMATION, _
    lpTimeZoneInformation As utc_TIME_ZONE_INFORMATION _
) As Long

Private Declare Function GetDynamicTimeZoneInformation Lib "kernel32" ( _
    pTimeZoneInformation As DYNAMIC_TIME_ZONE_INFORMATION _
) As Long
Private Declare Function SystemTimeToTzSpecificLocalTimeEx Lib "kernel32" ( _
    ByRef lpDynamicTimeZoneInformation As DYNAMIC_TIME_ZONE_INFORMATION _
    , ByRef lpUniversalTime As utc_SYSTEMTIME _
    , ByRef lpLocalTime As utc_SYSTEMTIME) As Long

Private Declare Function TzSpecificLocalTimeToSystemTimeEx Lib "kernel32" ( _
    lpDynamicTimeZoneInformation As DYNAMIC_TIME_ZONE_INFORMATION, _
    lpLocalTime As utc_SYSTEMTIME, _
    lpUniversalTime As utc_SYSTEMTIME _
) As Long

Private Declare Sub GetSystemTime Lib "kernel32" (lpSystemTime As utc_SYSTEMTIME)
Private Declare Sub GetLocalTime Lib "kernel32" (lpSystemTime As utc_SYSTEMTIME)

Private Declare Function utc_GetTimeZoneInformation Lib "kernel32" Alias "GetTimeZoneInformation" _
    (utc_lpTimeZoneInformation As utc_TIME_ZONE_INFORMATION) As Long
Private Declare Function utc_SystemTimeToTzSpecificLocalTime Lib "kernel32" Alias "SystemTimeToTzSpecificLocalTime" _
    (utc_lpTimeZoneInformation As utc_TIME_ZONE_INFORMATION, utc_lpUniversalTime As utc_SYSTEMTIME, utc_lpLocalTime As utc_SYSTEMTIME) As Long
Private Declare Function utc_TzSpecificLocalTimeToSystemTime Lib "kernel32" Alias "TzSpecificLocalTimeToSystemTime" _
    (utc_lpTimeZoneInformation As utc_TIME_ZONE_INFORMATION, utc_lpLocalTime As utc_SYSTEMTIME, utc_lpUniversalTime As utc_SYSTEMTIME) As Long
#End If


' ============================================= '
' Required types
' ============================================= '

#If Mac Then
#If VBA7 Then
Private Type utc_ShellResult
    utc_Output As String
    utc_ExitCode As LongPtr
End Type

#Else
Private Type utc_ShellResult
    utc_Output As String
    utc_ExitCode As Long
End Type

#End If
#Else
' Windows time structures.
Public Enum TIME_ZONE
    TIME_ZONE_ID_INVALID = 0
    TIME_ZONE_STANDARD = 1
    TIME_ZONE_DAYLIGHT = 2
End Enum

Public Type utc_SYSTEMTIME
    utc_wYear As Integer
    utc_wMonth As Integer
    utc_wDayOfWeek As Integer
    utc_wDay As Integer
    utc_wHour As Integer
    utc_wMinute As Integer
    utc_wSecond As Integer
    utc_wMilliseconds As Integer
End Type

Private Type utc_TIME_ZONE_INFORMATION
    utc_Bias As Long
    utc_StandardName(0 To 31) As Integer
    utc_StandardDate As utc_SYSTEMTIME
    utc_StandardBias As Long
    utc_DaylightName(0 To 31) As Integer
    utc_DaylightDate As utc_SYSTEMTIME
    utc_DaylightBias As Long
End Type

Private Type DYNAMIC_TIME_ZONE_INFORMATION
    Bias As Long
    StandardName(0 To 31) As Integer
    StandardDate As utc_SYSTEMTIME
    StandardBias As Long
    DaylightName(0 To 31) As Integer
    DaylightDate As utc_SYSTEMTIME
    DaylightBias As Long
    TimeZoneKeyName(0 To 127) As Integer
    DynamicDaylightTimeDisabled As Integer
End Type
#End If


Private Type StringBufferCache
    String_Buffer As String
    string_BufferPosition As Long
    string_BufferLength As Long
End Type


' ============================================= '
' Public Methods
' ============================================= '

''
' Parse UTC date to local date
'
' @method ParseUtc
' @param {Date} UtcDate
' @return {Date} Local date
' @throws 10011 - UTC parsing error

' NOTE: Mac functions may or may not return the millisecond portion of the value; they're untested.
'       Windows time parsing has been extensively tested to return the correct value.
''
Public Function ParseUtc(utc_UtcDate As Date) As Date
    ParseUtc = ConvertToLocalDate(utc_UtcDate)
End Function


Public Function ConvertToLocalDate(ByVal utc_UtcDate As Date) As Date
    On Error GoTo utc_ErrorHandling

#If Mac Then
    ConvertToLocalDate = utc_ConvertDate(utc_UtcDate)
#Else
    Dim utc_DynamicTimeZoneInfo As DYNAMIC_TIME_ZONE_INFORMATION
    Dim UTCDateYear As Integer ' The year of UTC date.
    
    Dim utc_UtcDateSysTime As utc_SYSTEMTIME ' Gets the year and month to compare.
    Dim utc_TimeZoneInfo As utc_TIME_ZONE_INFORMATION
    
    Dim utc_LocalDateSysTime As utc_SYSTEMTIME

    ' Convert to SystemTime to facilitate more accurate date checking.
    utc_UtcDateSysTime = utc_DateToSystemTime(utc_UtcDate)
    
    UTCDateYear = utc_UtcDateSysTime.utc_wYear

Recheck_Year:
    ' Get the timezone data for that year.
    GetDynamicTimeZoneInformation utc_DynamicTimeZoneInfo
    GetTimeZoneInformationForYear UTCDateYear, utc_DynamicTimeZoneInfo, utc_TimeZoneInfo
    SystemTimeToTzSpecificLocalTimeEx utc_DynamicTimeZoneInfo, utc_UtcDateSysTime, utc_LocalDateSysTime
    
    If UTCDateYear <> utc_LocalDateSysTime.utc_wYear Then
        UTCDateYear = utc_LocalDateSysTime.utc_wYear
        GoTo Recheck_Year
    End If
    
    ConvertToLocalDate = utc_SystemTimeToDate(utc_LocalDateSysTime)
#End If
 
    Exit Function

utc_ErrorHandling:
    Err.Raise 10011, "UtcConverter.ConvertToLocalDate", "UTC parsing error: " & Err.Number & " - " & Err.Description
End Function

''
' Convert local date to UTC date
'
' @method ConvertToUrc
' @param {Date} utc_LocalDate
' @return {Date} UTC date
' @throws 10012 - UTC conversion error
''
'Public Function LocalToUTC(utc_LocalDate As Date) As Date
'    LocalToUTC = ConvertToUtc(utc_LocalDate)
'End Function

Public Function ConvertToUtc(utc_LocalDate As Date) As Date
    On Error GoTo utc_ErrorHandling
    
#If Mac Then
    ConvertToUtc = utc_ConvertDate(utc_LocalDate, utc_ConvertToUtc:=True)
#Else
    Dim utc_DynamicTimeZoneInfo As DYNAMIC_TIME_ZONE_INFORMATION
    Dim utc_TimeZoneInfo As utc_TIME_ZONE_INFORMATION
    Dim utc_UtcDate As utc_SYSTEMTIME
    Dim utc_LocalSystemTime As utc_SYSTEMTIME
    
    utc_LocalSystemTime = utc_DateToSystemTime(utc_LocalDate)
    GetDynamicTimeZoneInformation utc_DynamicTimeZoneInfo
    GetTimeZoneInformationForYear utc_LocalSystemTime.utc_wYear, utc_DynamicTimeZoneInfo, utc_TimeZoneInfo
    TzSpecificLocalTimeToSystemTimeEx utc_DynamicTimeZoneInfo, utc_LocalSystemTime, utc_UtcDate

    ConvertToUtc = utc_SystemTimeToDate(utc_UtcDate)
#End If

    Exit Function

utc_ErrorHandling:
    Err.Raise 10012, "UtcConverter.ConvertToUtc", "UTC conversion error: " & Err.Number & " - " & Err.Description
End Function

' NOTE: As of now, "LocalTimeStamp" does nothing on a Mac; need to build "getTimeZoneOffset" for Mac, and I don't have one.
'       It will, however, output a UTC string that is correct for local time (eg, in the correct UTC for the given local time)
'       I also don't know how to get millisecond values out of a Mac, so that'll return zero, as well.
Public Function ISO8601TimeStamp(Optional IncludeMilliseconds As Boolean = True _
                                , Optional LocalTimeStamp As Boolean = False) As String
    Dim CurrentTimeVB As Date
    
    Dim tString_Buffer As StringBufferCache
' Note: This varies slightly from ConvertToISO8601Time because it's faster to do on Windows if you have SYSTEMTIME
#If Mac Then
    ' I'm sure there's a way to do this better, but this works for now.
    CurrentTimeVB = ConvertToUtc(VBA.Now())

    String_BufferAppend tString_Buffer, VBA.Format(CurrentTimeVB, ISOTimeFormatStr)
    If IncludeMilliseconds Then String_BufferAppend tString_Buffer, "." & VBA.Format(GetMilliseconds(CurrentTimeVB), "000")

#Else
    Dim tSysTime As utc_SYSTEMTIME

    If Not LocalTimeStamp Then
        GetSystemTime tSysTime
        CurrentTimeVB = utc_SystemTimeToDate(tSysTime)
    Else
        GetLocalTime tSysTime
        CurrentTimeVB = utc_SystemTimeToDate(tSysTime)
    End If

    String_BufferAppend tString_Buffer, VBA.Format(CurrentTimeVB, ISOTimeFormatStr)
    If IncludeMilliseconds Then String_BufferAppend tString_Buffer, "." & VBA.Format(tSysTime.utc_wMilliseconds, "000")

    If LocalTimeStamp Then
        String_BufferAppend tString_Buffer, CurrentISOTimezoneOffset
    Else
        String_BufferAppend tString_Buffer, ISO8601UTCTimeZone
    End If
#End If

    ISO8601TimeStamp = String_BufferToString(tString_Buffer)
End Function

' Wrappers to make it easier to use the below.
Public Function ParseISOTimeStampToUTC(utc_IsoString As String) As Date
    ParseISOTimeStampToUTC = ParseIso(utc_IsoString, True)
End Function

Public Function ParseISOTimeStampToLocal(utc_IsoString As String) As Date
    ParseISOTimeStampToLocal = ParseIso(utc_IsoString)
End Function

' While this function may look silly, it is useful when converting disparate time zone stamps in a log to a common one when aligning user input data.
Public Function ParseISOTimeStampToISO8601TimeStamp(ByRef InVal As String _
                                                , Optional LocalOut As Boolean = False) As String
    Dim tDateTime As Date
    
    tDateTime = ParseIso(InVal, True)
    ParseISOTimeStampToISO8601TimeStamp = ConvertToISO8601Time(tDateTime, True, LocalOut, True)
End Function

''
' Parse ISO 8601 date string to local date
'
' @method ParseIso
' @param {Date} utc_IsoString
' @return {Date} Local date
' @throws 10013 - ISO 8601 parsing error
'
Public Function ParseIso(utc_IsoString As String _
                        , Optional ByVal OutputUTCDate As Boolean = False) As Date
    On Error GoTo utc_ErrorHandling
    Dim utc_Parts() As String
    Dim utc_DateTimeOut As Date

    If utc_IsoString = vbNullString Then Exit Function
    utc_Parts = VBA.Split(utc_IsoString, ISO8601DateTimeSeparator)

#If Mac Then
' Mac doesn't have RegEx, so we can't map all of the dates, only date numbers, unlike RegEx which can support date names and most of the suite of
' ISO8601 Date formatting.
    Dim utc_DateParts() As String
    Dim utc_TimeParts() As String
    Dim utc_OffsetIndex As Long
    Dim utc_HasOffset As Boolean
    Dim utc_NegativeOffset As Boolean
    Dim utc_OffsetParts() As String
    Dim utc_Offset As Date

    utc_DateParts = VBA.Split(utc_Parts(0), ISO8601DateDelimiter)
    utc_DateTimeOut = VBA.DateSerial(VBA.CInt(utc_DateParts(0)), VBA.CInt(utc_DateParts(1)), VBA.CInt(utc_DateParts(2)))
'TimeSerialDbl
    If UBound(utc_Parts) > 0 Then
        If VBA.InStr(utc_Parts(1), ISO8601UTCTimeZone) Then
            utc_TimeParts = VBA.Split(VBA.Replace(utc_Parts(1), ISO8601UTCTimeZone, vbNullString), ISO8601TimeDelimiter)
        Else
            utc_OffsetIndex = VBA.InStr(1, utc_Parts(1), "+")
            If utc_OffsetIndex = 0 Then
                utc_NegativeOffset = True
                utc_OffsetIndex = VBA.InStr(1, utc_Parts(1), "-")
            End If

            If utc_OffsetIndex > 0 Then
                utc_HasOffset = True
                utc_TimeParts = VBA.Split(VBA.Left$(utc_Parts(1), utc_OffsetIndex - 1), ISO8601TimeDelimiter)
                utc_OffsetParts = VBA.Split(VBA.Right$(utc_Parts(1), Len(utc_Parts(1)) - utc_OffsetIndex), ISO8601TimeDelimiter)

                Select Case UBound(utc_OffsetParts)
                Case 0
                    utc_Offset = TimeSerialDbl(VBA.CDbl(utc_OffsetParts(0)), 0, 0)
                Case 1
                    utc_Offset = TimeSerialDbl(VBA.CDbl(utc_OffsetParts(0)), VBA.CDbl(utc_OffsetParts(1)), 0)
                Case 2
                    ' VBA.Val does not use regional settings, use for seconds to avoid decimal/comma issues
                    utc_Offset = TimeSerialDbl(VBA.CDbl(utc_OffsetParts(0)), VBA.CDbl(utc_OffsetParts(1)), VBA.CDbl(VBA.Val(utc_OffsetParts(2))))
                End Select

                If utc_NegativeOffset Then: utc_Offset = -utc_Offset
            Else
                utc_TimeParts = VBA.Split(utc_Parts(1), ISO8601TimeDelimiter)
            End If
        End If

        Select Case UBound(utc_TimeParts)
        Case 0
            utc_DateTimeOut = utc_DateTimeOut + TimeSerialDbl(VBA.CInt(utc_TimeParts(0)), 0, 0)
        Case 1
            utc_DateTimeOut = utc_DateTimeOut + TimeSerialDbl(VBA.CInt(utc_TimeParts(0)), VBA.CInt(utc_TimeParts(1)), 0)
        Case 2
            ' VBA.Val does not use regional settings, use for seconds to avoid decimal/comma issues
            utc_DateTimeOut = utc_DateTimeOut + TimeSerialDbl(VBA.CInt(utc_TimeParts(0)), VBA.CInt(utc_TimeParts(1)), Int(VBA.Val(utc_TimeParts(2))))
        End Select

        If OutputUTCDate Then utc_DateTimeOut = ConvertToLocalDate(utc_DateTimeOut)

        If utc_HasOffset Then
            ParseIso = utc_DateTimeOut - utc_Offset
        End If
    End If

    Exit Function
#Else
    If UBound(utc_Parts) > 0 Then
        utc_DateTimeOut = ConvDateUTC(utc_Parts(0)) + ConvTimeUTC(utc_Parts(1))
        If Not OutputUTCDate Then
            ParseIso = ConvertToLocalDate(utc_DateTimeOut)
        Else
            ParseIso = utc_DateTimeOut
        End If
    Else ' Assume any "Date Only" Text doesn't have a timezone (they aren't converted the other way, either)
        ParseIso = ConvDateUTC(utc_Parts(0))
    End If
    Exit Function
#End If
utc_ErrorHandling:
    Err.Raise 10013, "UtcConverter.ParseIso", "ISO 8601 parsing error for " & utc_IsoString & ": " & Err.Number & " - " & Err.Description
End Function

Public Function ConvertToUTCISO8601TimeStamp(ByVal LocalDateIn As Date) As String
    ConvertToUTCISO8601TimeStamp = ConvertToISO8601Time(LocalDateIn, False, False, True)
End Function

Public Function ConvertToLocalISO8601TimeStamp(ByVal UTCDateIn As Date) As String
    ConvertToLocalISO8601TimeStamp = ConvertToISO8601Time(UTCDateIn, True, True, True)
End Function

''
' Convert local date to ISO 8601 string
'
' @method ConvertToIso
' @param {Date} utc_LocalDate
' @return {Date} ISO 8601 string
' @throws 10014 - ISO 8601 conversion error
''
Public Function ConvertToIsoTime(utc_LocalDate As Date _
                            , Optional OutputAsLocalDate As Boolean = False) As String
                            
    On Error GoTo utc_ErrorHandling
    ConvertToIsoTime = ConvertToISO8601Time(utc_LocalDate, False, False, True)
    Exit Function

utc_ErrorHandling:
    Err.Raise 10014, "UtcConverter.ConvertToIso", "ISO 8601 conversion error: " & Err.Number & " - " & Err.Description
End Function


' Convert to ISOTimeStamp
' Converts a provided date into an ISO8601 formatted string.
' By default, assumes you pass in a local date and outputs a UTC date string.
' Set isUTC to True if you already have the UTC date.
' Set OutputLocalString to true if you want to output a localized timestamp string.
' This would be useful for instance if you want to know the geographic region an
' action was performed by a user.
' Prior versions of this function did not convert if it was a date only.
' This is no longer true, all dates and times are always localaized.
' To revert back to that behavior, set ConvertDateOnly to False
Public Function ConvertToISO8601Time(ByVal DateIn As Date _
                                    , Optional isUTC As Boolean = False _
                                    , Optional OutputLocalString As Boolean = False _
                                    , Optional IncludeMilliseconds As Boolean = True) As String

    Dim fStringBuffer As StringBufferCache
  
    Dim tBias As Long
    Dim OutputDate As Date
    Dim MSCount As Long

    If (isUTC And Not OutputLocalString) Then
        tBias = 0
        ' Don't need to convert.
        OutputDate = DateIn
    ElseIf (isUTC And OutputLocalString) Then
        ' Convert UTC to local
        OutputDate = ConvertToLocalDate(DateIn)
        tBias = VBA.DateDiff("n", OutputDate, DateIn)
    ElseIf OutputLocalString Then
        ' No conversi on needed; get bias.
        OutputDate = DateIn
        tBias = GetBiasForGivenLocalDate(OutputDate)
    Else
        OutputDate = ConvertToUtc(DateIn)
        tBias = GetBiasForGivenLocalDate(OutputDate)
    End If
    
    Dim tString_Buffer As StringBufferCache

    String_BufferAppend tString_Buffer, VBA.Format(OutputDate, ISOTimeFormatStr)
    
    If IncludeMilliseconds Then
        MSCount = GetMilliseconds(OutputDate)
        String_BufferAppend tString_Buffer, "." & VBA.Format(MSCount, "000")
    End If
    
    If OutputLocalString Then
        String_BufferAppend tString_Buffer, ISOTimezoneOffset(tBias)
    Else
        String_BufferAppend tString_Buffer, ISO8601UTCTimeZone
    End If

    ConvertToISO8601Time = String_BufferToString(tString_Buffer)
End Function


' Provides a format string to other functions that complies with ISO8601
Private Function ISOTimeFormatStr(Optional IncludeMilliseconds As Boolean = False _
                                , Optional includeTimeZone As Boolean = False) As String
    Dim tString_Buffer As StringBufferCache

    String_BufferAppend tString_Buffer, "yyyy-mm-ddTHH:mm:ss"
    If IncludeMilliseconds Then String_BufferAppend tString_Buffer, ".000"
    If includeTimeZone Then String_BufferAppend tString_Buffer, ISOTimezoneOffset
    ISOTimeFormatStr = String_BufferToString(tString_Buffer)
End Function


Private Function RoundUp(ByVal Value As Double) As Long
    Dim lngVal As Long
    Dim deltaValue As Double
    
    lngVal = VBA.CLng(Value)
    deltaValue = lngVal - Value
        
    If deltaValue < 0 Then
        RoundUp = lngVal + 1
    Else
        RoundUp = lngVal
    End If
End Function
Private Function RoundDown(ByVal Value As Double) As Long
    Dim lngVal As Long
    Dim deltaValue As Double
    
    lngVal = VBA.CLng(Value)
    deltaValue = lngVal - Value
        
    If deltaValue <= 0 Then
        RoundDown = lngVal
    Else
        RoundDown = lngVal - 1
    End If
End Function


' ============================================= '
' Private Functions
' ============================================= '

#If Mac Then

Private Function utc_ConvertDate(utc_Value As Double _
                                , Optional utc_ConvertToUtc As Boolean = False) As Date
    Dim utc_ShellCommand As String
    Dim utc_Result As utc_ShellResult
    Dim utc_Parts() As String
    Dim utc_DateParts() As String
    Dim utc_TimeParts() As String

    If utc_ConvertToUtc Then
        utc_ShellCommand = "date -ur `date -jf '%Y-%m-%d %H:%M:%S' " & _
            "'" & VBA.Format$(utc_Value, "yyyy-mm-dd HH:mm:ss") & "' " & _
            " +'%s'` +'%Y-%m-%d %H:%M:%S'"
    Else
        utc_ShellCommand = "date -jf '%Y-%m-%d %H:%M:%S %z' " & _
            "'" & VBA.Format$(utc_Value, "yyyy-mm-dd HH:mm:ss") & " +0000' " & _
            "+'%Y-%m-%d %H:%M:%S'"
    End If

    utc_Result = utc_ExecuteInShell(utc_ShellCommand)

    If utc_Result.utc_Output = "" Then
        Err.Raise 10015, "UtcConverter.utc_ConvertDate", "'date' command failed"
    Else
        utc_Parts = Split(utc_Result.utc_Output, " ")
        utc_DateParts = Split(utc_Parts(0), "-")
        utc_TimeParts = Split(utc_Parts(1), ":")

        utc_ConvertDate = DateSerial(utc_DateParts(0), utc_DateParts(1), utc_DateParts(2)) + _
            TimeSerial(utc_TimeParts(0), utc_TimeParts(1), utc_TimeParts(2))
    End If
End Function

Private Function utc_ExecuteInShell(utc_ShellCommand As String) As utc_ShellResult
#If VBA7 Then
    ' 64bit Mac
    Dim utc_File As LongPtr
    Dim utc_Read As LongPtr
#Else
    Dim utc_File As Long
    Dim utc_Read As Long
#End If

    Dim utc_Chunk As String

    On Error GoTo utc_ErrorHandling
    utc_File = utc_popen(utc_ShellCommand, "r")

    If utc_File = 0 Then: Exit Function

    Do While utc_feof(utc_File) = 0
        utc_Chunk = VBA.Space$(50)
        utc_Read = VBA.CLng(utc_fread(utc_Chunk, 1, VBA.Len(utc_Chunk) - 1, utc_File))
        If utc_Read > 0 Then
            utc_Chunk = VBA.Left$(utc_Chunk, VBA.CLng(utc_Read))
            utc_ExecuteInShell.utc_Output = utc_ExecuteInShell.utc_Output & utc_Chunk
        End If
    Loop

utc_ErrorHandling:
    utc_ExecuteInShell.utc_ExitCode = CLng(utc_pclose(utc_File))
End Function

#Else
' Windows

' Pass in a date, this will return a Windows SystemTime structure with millisecond accuracy.
Private Function utc_DateToSystemTime(ByRef utc_Value As Date) As utc_SYSTEMTIME ' "Helper Functions
    With utc_DateToSystemTime
        .utc_wYear = VBA.Year(utc_Value)
        .utc_wMonth = VBA.Month(utc_Value)
        .utc_wDay = VBA.Day(utc_Value)
        .utc_wHour = VBA.Hour(utc_Value)
        .utc_wMinute = VBA.Minute(utc_Value)
        .utc_wMilliseconds = GetMilliseconds(utc_Value)
        If .utc_wMilliseconds >= 500 Then
            .utc_wSecond = VBA.Second(utc_Value) - 1
        Else
            .utc_wSecond = VBA.Second(utc_Value)
        End If
    End With
End Function


Private Function utc_SystemTimeToDate(ByRef utc_Value As utc_SYSTEMTIME) As Date ' "Helper Function" for Public Functions (below)
    utc_SystemTimeToDate = DateSerial(utc_Value.utc_wYear _
                                    , utc_Value.utc_wMonth _
                                    , utc_Value.utc_wDay) + _
                            TimeSerialDbl(utc_Value.utc_wHour _
                                        , utc_Value.utc_wMinute _
                                        , utc_Value.utc_wSecond _
                                        , utc_Value.utc_wMilliseconds)
End Function


Private Function ConvDateUTC(ByVal InVal As String) As Date
    Dim RetVal As Variant
    
'    Dim RegEx As Object
'    Set RegEx = CreateObject("VBScript.RegExp")
    Dim RegEx As New RegExp
    With RegEx
        .Global = True
        .Multiline = True
        .IgnoreCase = False
    End With
    
    RegEx.Pattern = "^(\d{4})-?(\d{2})?-?(\d{1,2})?$|^(\d{4})-?W(\d{2})?-?(\d)?$|^(\d{4})-?(\d{3})$"
    Dim Match As Object
    Set Match = RegEx.Execute(InVal)
    
    If Match.Count <> 1 Then Exit Function
    With Match(0)
        If Not IsEmpty(.SubMatches(0)) Then
            'YYYY-MM-DD
            If IsEmpty(.SubMatches(1)) Then  'YYYY
                RetVal = DateSerial(CInt(.SubMatches(0)), 1, 1)
            ElseIf IsEmpty(.SubMatches(2)) Then 'YYYY-MM
                RetVal = DateSerial(CInt(.SubMatches(0)), CInt(.SubMatches(1)), 1)
            Else 'YYYY-MM-DD or YYYY-MM-D
                RetVal = DateSerial(CInt(.SubMatches(0)), CInt(.SubMatches(1)), CInt(.SubMatches(2)))
            End If
        ElseIf Not IsEmpty(.SubMatches(3)) Then
            'YYYY-Www-D
            RetVal = DateSerial(CInt(.SubMatches(3)), 1, 4) '4th of jan is always week 1
            RetVal = RetVal - Weekday(RetVal, 2) 'subtract the weekday number of 4th of jan
            RetVal = RetVal + 7 * (CInt(.SubMatches(4)) - 1) 'add 7 times the (weeknumber - 1)
            
            If IsEmpty(.SubMatches(5)) Then 'YYYY-Www
                RetVal = RetVal + 1 'choose monday of that week
            Else 'YYYY-Www-D
                RetVal = RetVal + CInt(.SubMatches(5)) 'choose day of that week 1-7 monday to sunday
            End If
        Else
            'YYYY-DDD
            RetVal = DateSerial(CInt(.SubMatches(6)), 1, 1) + CInt(.SubMatches(7)) - 1
        End If
    End With
    
    ConvDateUTC = RetVal
End Function

Private Function ConvTimeUTC(ByRef InVal As String) As Date

    Dim dblHours As Double
    Dim dblMinutes As Double
    Dim dblSeconds As Double
    Dim dblMilliseconds As Double
        
    Dim RegEx As New RegExp ' Object
    'Set RegEx = CreateObject("VBScript.RegExp")
    
    With RegEx
        .Global = True
        .Multiline = False
        .IgnoreCase = False
    End With

    ' Allowing for hours,minutes, and seconds to have partial amounts per ISO8601 standard.
    RegEx.Pattern = "^(\d{0,2}[\.\,]?\d*(?=[\+\-Z :]|$)):?(\d{0,2}[\.\,]?\d*(?=[\+\-Z :]|$))?:?(\d{0,2}[\.\,]?\d*(?=[\+\-Z :]|$))?(\+|\-|Z)?(\d{1,2})?:?(\d{1,2})?$"

    Dim Match As Object
    Set Match = RegEx.Execute(InVal)
    
    If Match.Count <> 1 Then Exit Function

    With Match(0)
        'hh:mm:ss.nnn detection
        ' Load hours in, then detect if there's more to do.
        dblHours = CDbl(NzEmpty(.SubMatches(0), 0))

        If Not (IsEmpty(.SubMatches(3)) Or IsEmpty(.SubMatches(4)) Or NzEmpty(.SubMatches(3), ISO8601UTCTimeZone) = ISO8601UTCTimeZone) Then _
            dblHours = dblHours - CDbl(NzEmpty(.SubMatches(3) & .SubMatches(4), vbNullString))
        
        dblMinutes = CDbl(NzEmpty(.SubMatches(1), vbNullString))
        
        If Not (IsEmpty(.SubMatches(3)) Or IsEmpty(.SubMatches(5)) Or NzEmpty(.SubMatches(3), ISO8601UTCTimeZone) = ISO8601UTCTimeZone) Then _
            dblMinutes = dblMinutes - CDbl(NzEmpty(.SubMatches(3), vbNullString) & NzEmpty(.SubMatches(5), vbNullString))
        
        dblSeconds = CDbl(NzEmpty(.SubMatches(2), vbNullString))
    End With
    
    ConvTimeUTC = TimeSerialDbl(dblHours, dblMinutes, dblSeconds)

End Function

Private Function NzEmpty(ByVal Value As Variant, Optional ByVal value_when_null As Variant = 0) As Variant

    Dim return_value As Variant
    On Error Resume Next 'supress error handling

    If IsEmpty(Value) Or IsNull(Value) Or (VarType(Value) = vbString And Value = vbNullString) Then
        return_value = value_when_null
    Else
        return_value = Value
    End If

    Err.Clear 'clear any errors that might have occurred
    On Error GoTo 0 'reinstate error handling

    NzEmpty = return_value

End Function
#End If


' Will return a Date type Double (specified as Double because it makes VBA less likely to "help")
Public Function TimeSerialDbl(ByVal HoursIn As Double _
                            , ByVal MinutesIn As Double _
                            , ByVal SecondsIn As Double _
                            , Optional ByVal MillisecondsIn As Double = 0) As Double
    Dim tMS As Double
    Dim tSec As Double
    Dim tSecTemp As Double
    tSec = VBA.CDbl(RoundDown(SecondsIn))
    tSecTemp = SecondsIn - tSec
    tMS = (tSecTemp * (TotalMillisecondsInDay / TotalSecondsInDay)) \ 1
    tMS = tMS + MillisecondsIn
    If (tSecTemp > 0.5) Then tSec = tSec - 1
    If tMS = 500 Then tMS = tMS - 0.001 ' Shave a hair, because otherwise it'll round up too much.
    TimeSerialDbl = (HoursIn / TotalHoursInDay) + (MinutesIn / TotalMinutesInDay) + CDbl((tSec / TotalSecondsInDay)) + (tMS / TotalMillisecondsInDay)
End Function

' If given a time double, will return the millisecond portion of the time.
Private Function GetMilliseconds(ByVal TimeIn As Double) As Variant
    Dim IntDatePart As Long
    Dim DblTimePart As Double
    Dim LngSeconds As Long ' Used to remove whole seconds.
    Dim DblSecondsPart As Double
    
    Dim DblMS As Double
    Dim MSCount As Double
        
    ' Get rid of the date portion
    ' There is an annoying bug where VBA rounds up in certain cases when
    ' using the \ operator and dividing by 1. So, divide by 2 and double it.
    ' this side steps the bug and ensures it always rounds down.
    IntDatePart = RoundDown(TimeIn)
    DblTimePart = TimeIn - IntDatePart
    
    LngSeconds = RoundDown(TotalSecondsInDay * DblTimePart)
    DblSecondsPart = LngSeconds / TotalSecondsInDay
    DblMS = DblTimePart - DblSecondsPart
    MSCount = ((DblMS * (TotalMillisecondsInDay))) \ 1
    If MSCount >= 1000 Then MSCount = 0
    GetMilliseconds = MSCount
End Function


Public Function CurrentLocalBiasFromUTC(Optional ByVal OutputAsHours As Boolean = False) As Long
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' This returns the CURRENT amount of time in minutes (if OutputAsHours is omitted or
' false) or hours (if OutputAsHours is True) that should be added (or subtracted) to the
' local time to get UTC. It should (untested on Mac as of yet) return the value
' adjusted for DST if active.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim tBias As Long

#If Mac Then
    ' While we could do this for the Windows version, too, the Windows functions are rock solid and
    ' work (these work, too), and are certain to get the correct data. I'm reasonably sure these
    ' work now, but don't have a Mac to test.
    tBias = GetBiasForGivenLocalDate(VBA.Now(), OutputAsHours)
#Else
    Dim TZI As utc_TIME_ZONE_INFORMATION
    Dim DST As TIME_ZONE
    DST = utc_GetTimeZoneInformation(TZI)

    If DST = TIME_ZONE_DAYLIGHT Then
        tBias = TZI.utc_Bias + TZI.utc_DaylightBias
    Else
        tBias = TZI.utc_Bias + TZI.utc_StandardBias
    End If
    
    If OutputAsHours Then tBias = tBias / 60 ' This is already done in GetBiasForGivenLocalDate for Mac.
#End If

    CurrentLocalBiasFromUTC = tBias

End Function

Public Function CurrentISOTimezoneOffset() As String
    CurrentISOTimezoneOffset = ISOTimezoneOffset(CurrentLocalBiasFromUTC)
End Function


Public Function GetBiasForGivenLocalDate(ByVal LocalDateIn As Date _
                                        , Optional ByVal OutputAsHours As Boolean = False) As Long
    Dim DateUTCNow As Date
    
    DateUTCNow = ConvertToUtc(LocalDateIn)

    ' I tried to get fancy here and retrieve the bias from the OS, but that turned into a huge amount of work.
    ' unless your time zone is defined by change on a specific day, this is far simpler and easier
    ' than chasing week numbers around.
    If Not OutputAsHours Then
        GetBiasForGivenLocalDate = VBA.DateDiff("n", LocalDateIn, DateUTCNow)
    Else
        GetBiasForGivenLocalDate = VBA.DateDiff("h", LocalDateIn, DateUTCNow)
    End If
End Function

Public Function ISOTimezoneOffsetOnDate(ByVal LocalDateIn As Date) As String
    ISOTimezoneOffsetOnDate = ISOTimezoneOffset(GetBiasForGivenLocalDate(LocalDateIn))
End Function


' Provides the ISO Offset time from an input (or current offset if none is passed in) to build an ISO8601 output String
Private Function ISOTimezoneOffset(Optional TimeBias As Long = 0) As String

    Dim strOffsetOut As String

    Dim tString_Buffer As StringBufferCache

    Dim OffsetLong As Long
    Dim hourOffset As Long
    Dim minOffset As Long
    
    ' Counterintuitively, the Bias is postive (time ahead), the offset is the negative value of bias.
    OffsetLong = TimeBias * -1
    
    hourOffset = OffsetLong \ 60
    minOffset = OffsetLong Mod 60
    
    If OffsetLong = 0 Then
        ISOTimezoneOffset = ISO8601UTCTimeZone
    Else
        If OffsetLong > 0 Then String_BufferAppend tString_Buffer, "+"
        String_BufferAppend tString_Buffer, VBA.CStr(VBA.Format(hourOffset, "00"))
        String_BufferAppend tString_Buffer, ISO8601TimeDelimiter
        String_BufferAppend tString_Buffer, VBA.CStr(VBA.Format(minOffset, "00"))
        
        ISOTimezoneOffset = String_BufferToString(tString_Buffer)
    End If
End Function


' String_BufferAppend
' Based on VBA-Tools\Jsonconverter's "json_BufferAppend" functions
' To use, your calling routine needs to store the input variables to be handed back.
Private Sub String_BufferAppend(ByRef StringBufferIn As StringBufferCache _
                                , ByRef String_Append As Variant)
    ' VBA can be slow to append strings due to allocating a new string for each append
    ' Instead of using the traditional append, allocate a large empty string and then copy string at append position
    '
    ' Example:
    ' Buffer: "abc  "
    ' Append: "def"
    ' Buffer Position: 3
    ' Buffer Length: 5
    '
    ' Buffer position + Append length > Buffer length -> Append chunk of blank space to buffer
    ' Buffer: "abc       "
    ' Buffer Length: 10
    '
    ' Put "def" into buffer at position 3 (0-based)
    ' Buffer: "abcdef    "
    '
    ' Approach based on cStringBuilder from vbAccelerator
    ' http://www.vbaccelerator.com/home/VB/Code/Techniques/RunTime_Debug_Tracing/VB6_Tracer_Utility_zip_cStringBuilder_cls.asp
    '
    ' and clsStringAppend from Philip Swannell
    ' https://github.com/VBA-tools/VBA-JSON/pull/82

    Dim String_AppendLength As Long
    Dim String_LengthPlusPosition As Long

    String_AppendLength = VBA.Len(String_Append)
    String_LengthPlusPosition = String_AppendLength + StringBufferIn.string_BufferPosition

    If String_LengthPlusPosition > StringBufferIn.string_BufferLength Then
        ' Appending would overflow buffer, add chunk
        ' (double buffer length or append length, whichever is bigger)
        Dim string_AddedLength As Long
        string_AddedLength = IIf(String_AppendLength > StringBufferIn.string_BufferLength, String_AppendLength, StringBufferIn.string_BufferLength)

        StringBufferIn.String_Buffer = StringBufferIn.String_Buffer & VBA.Space$(string_AddedLength)
        StringBufferIn.string_BufferLength = StringBufferIn.string_BufferLength + string_AddedLength
    End If

    ' Note: Namespacing with VBA.Mid$ doesn't work properly here, throwing compile error:
    ' Function call on left-hand side of assignment must return Variant or Object
    If String_AppendLength > 0 Then
        Mid$(StringBufferIn.String_Buffer, StringBufferIn.string_BufferPosition + 1, String_AppendLength) = CStr(String_Append)
    End If
    StringBufferIn.string_BufferPosition = StringBufferIn.string_BufferPosition + String_AppendLength
End Sub

Private Function String_BufferToString(ByRef StringBufferIn As StringBufferCache) As String
    If StringBufferIn.string_BufferPosition > 0 Then
        String_BufferToString = VBA.Left$(StringBufferIn.String_Buffer, StringBufferIn.string_BufferPosition)
    End If
End Function



