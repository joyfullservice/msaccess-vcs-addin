Attribute VB_Name = "modStringUtil"
'---------------------------------------------------------------------------------------
' Module    : modStringUtil
' Author    : Adam Waller
' Date      : 12/4/2020
' Purpose   : String manipulation utilities: replace, repeat, match, deduplicate,
'           : quote escaping, and coalesce.
' Layer     : Utility
' Depends on: (none)
'---------------------------------------------------------------------------------------
Option Compare Database
Option Private Module
Option Explicit
'@Folder("Utility")


'---------------------------------------------------------------------------------------
' Procedure : MultiReplace
' Author    : Adam Waller
' Date      : 1/18/2019
' Purpose   : Does a string replacement of multiple items in one call.
'           : IMPORTANT NOTE: When no `compare` option is specified, the Replace
'           : function will use the module's Option Compare value, not the default
'           : parameter of vbBinaryCompare.
'---------------------------------------------------------------------------------------
'
Public Function MultiReplace(ByVal strText As String, ParamArray varPairs()) As String
    Dim intPair As Integer
    For intPair = 0 To UBound(varPairs) Step 2
        strText = Replace(strText, varPairs(intPair), varPairs(intPair + 1))
    Next intPair
    MultiReplace = strText
End Function


'---------------------------------------------------------------------------------------
' Procedure : Coalesce
' Author    : Adam Waller
' Date      : 5/15/2021
' Purpose   : Return the first non-empty string from an array of string values
'---------------------------------------------------------------------------------------
'
Public Function Coalesce(ParamArray varStrings()) As String
    Dim intString As Integer
    For intString = 0 To UBound(varStrings)
        If Nz(varStrings(intString)) <> vbNullString Then
            Coalesce = varStrings(intString)
            Exit For
        End If
    Next intString
End Function


'---------------------------------------------------------------------------------------
' Procedure : DblQ
' Author    : Adam Waller
' Date      : 9/10/2022
' Purpose   : Double any single or double quotes in the string (Used for SQL output)
'---------------------------------------------------------------------------------------
'
Public Function DblQ(strText As String) As String
    DblQ = MultiReplace(strText, "'", "''", """", """""")
End Function


'---------------------------------------------------------------------------------------
' Procedure : DeDupString
' Author    : Adam Waller
' Date      : 9/10/2022
' Purpose   : Removes consecutive duplicates of a string within a string.
'           : (Some logic built in for efficiency and to prevent infinite loops)
'---------------------------------------------------------------------------------------
'
Public Function DeDupString(strText As String, strDuplicated As String) As String

    Dim lngCnt As Long
    Dim strNew As String

    strNew = strText

    ' See if the searched string exists before attempting to replace
    If InStr(1, strText, strDuplicated) > 0 Then
        For lngCnt = 10 To 2 Step -1
            strNew = Replace(strNew, Repeat(strDuplicated, lngCnt), strDuplicated)
        Next lngCnt
    End If

    ' Return deduplicated string
    DeDupString = strNew

End Function


'---------------------------------------------------------------------------------------
' Procedure : StartsWith
' Author    : Adam Waller
' Date      : 11/5/2020
' Purpose   : See if a string begins with a specified string.
'---------------------------------------------------------------------------------------
'
Public Function StartsWith(strText As String, strStartsWith As String, Optional Compare As VbCompareMethod = vbBinaryCompare) As Boolean
    If Len(strStartsWith) = 0 Then
        StartsWith = True
    Else
        StartsWith = (InStr(1, strText, strStartsWith, Compare) = 1)
    End If
End Function


'---------------------------------------------------------------------------------------
' Procedure : EndsWith
' Author    : Adam Waller
' Date      : 4/29/2021
' Purpose   : See if a string ends with a specified string.
'---------------------------------------------------------------------------------------
'
Public Function EndsWith(strText As String, strEndsWith As String, Optional Compare As VbCompareMethod = vbBinaryCompare) As Boolean
    EndsWith = (StrComp(Right$(strText, Len(strEndsWith)), strEndsWith, Compare) = 0)
    'EndsWith = (InStr(1, strText, strEndsWith, Compare) = len(strtext len(strendswith) 1)
End Function


'---------------------------------------------------------------------------------------
' Procedure : Repeat
' Author    : Adam Waller
' Date      : 4/29/2021
' Purpose   : Repeat a string a specified number of times
'---------------------------------------------------------------------------------------
'
Public Function Repeat(strText As String, lngTimes As Long) As String
    Repeat = Replace$(Space$(lngTimes), " ", strText)
End Function


'---------------------------------------------------------------------------------------
' Procedure : LikeAny
' Author    : Adam Waller
' Date      : 3/14/2023
' Purpose   : Returns true if strTest is LIKE any of the array of expressions.
'           : (Short ciruits to first matching expression)
'---------------------------------------------------------------------------------------
'
Public Function LikeAny(strTest As String, ParamArray varLikeThis()) As Boolean
    Dim lngCnt As Long
    For lngCnt = 0 To UBound(varLikeThis)
        If strTest Like varLikeThis(lngCnt) Then
            LikeAny = True
            Exit For
        End If
    Next lngCnt
End Function
