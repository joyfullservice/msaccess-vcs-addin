Attribute VB_Name = "modTestUtcConverter"
'---------------------------------------------------------------------------------------
' Module    : modTestUtcConverter
' Author    : Adam Waller
' Date      : 5/12/2026
' Purpose   : Unit tests for modUtcConverter date/time functions.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
Option Private Module
'@Folder("Tests.Utility")


Public Sub TestParseIso_BasicDate()
    Dim dte As Date
    dte = ParseIso("2024-06-15T12:30:00Z")
    TestAssert Year(dte) = 2024, "year"
    TestAssert Month(dte) = 6, "month"
    TestAssert Day(dte) = 15, "day"
End Sub


Public Sub TestParseIso_WithTimezone()
    Dim dteZ As Date
    Dim dtePlus As Date
    dteZ = ParseIso("2024-01-01T12:00:00Z")
    dtePlus = ParseIso("2024-01-01T12:00:00+00:00")
    TestAssert dteZ = dtePlus, "Z and +00:00 produce same result"
End Sub


Public Sub TestConvertToUtcAndBack()
    Dim dteLocal As Date
    Dim dteUtc As Date
    Dim dteBack As Date
    dteLocal = #6/15/2024 2:30:00 PM#
    dteUtc = ConvertToUtc(dteLocal)
    dteBack = ConvertToLocalDate(dteUtc)
    TestAssert Abs(dteLocal - dteBack) < 1 / 86400, "round trip within 1 second"
End Sub


Public Sub TestISO8601TimeStamp()
    Dim strTs As String
    strTs = ISO8601TimeStamp()
    TestAssert Len(strTs) > 0, "returns non-empty"
    TestAssert InStr(strTs, "T") > 0, "contains T separator"
End Sub


Public Sub TestConvertToIsoTime()
    Dim strIso As String
    strIso = ConvertToIsoTime(#6/15/2024 2:30:00 PM#)
    TestAssert InStr(strIso, "2024") > 0, "contains year"
    TestAssert InStr(strIso, "T") > 0, "contains T separator"
End Sub


Public Sub TestTimeSerialDbl()
    Dim dblResult As Double
    dblResult = TimeSerialDbl(1, 30, 0)
    TestAssert dblResult > 0, "returns positive value for 1h 30m"
    Dim dblResult2 As Double
    dblResult2 = TimeSerialDbl(0, 0, 45)
    TestAssert dblResult2 > 0, "returns positive value for 45 seconds"
    TestAssert dblResult > dblResult2, "1h 30m is greater than 45 seconds"
End Sub


Public Sub TestCurrentLocalBiasFromUTC()
    Dim lngBias As Long
    lngBias = CurrentLocalBiasFromUTC()
    ' Bias should be between -12 and +14 hours (in minutes)
    TestAssert lngBias >= -720 And lngBias <= 840, "bias in valid range (minutes)"

    Dim lngHours As Long
    lngHours = CurrentLocalBiasFromUTC(True)
    TestAssert lngHours >= -12 And lngHours <= 14, "bias in valid range (hours)"
End Sub
