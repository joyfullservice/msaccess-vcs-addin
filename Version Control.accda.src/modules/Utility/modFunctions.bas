Attribute VB_Name = "modFunctions"
'---------------------------------------------------------------------------------------
' Module    : modFunctions
' Author    : Adam Waller
' Date      : 12/4/2020
' Purpose   : General-purpose utility functions: VBE type helpers, file name encoding,
'           : date comparison, sorting, null handling, array helpers, and environment.
' Layer     : Utility
' Depends on: modObjects (FSO, Perf), modErrorHandling
'---------------------------------------------------------------------------------------
Option Compare Database
Option Private Module
Option Explicit
'@Folder("Utility")


' API function to pause processing
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)


'---------------------------------------------------------------------------------------
' Procedure : GetVBEExtByType
' Author    : Adam Waller
' Date      : 6/2/2015
' Purpose   : Return a standardized VBE component extension by type
'---------------------------------------------------------------------------------------
'
Public Function GetVBEExtByType(cmp As VBComponent) As String
    Dim strExt As String
    Select Case cmp.Type
        Case vbext_ct_StdModule:    strExt = ".bas"
        Case vbext_ct_MSForm:       strExt = ".frm" ' (not used in Microsoft Access)
        Case Else ' vbext_ct_Document, vbext_ct_ActiveXDesigner, vbext_ct_ClassModule
            strExt = ".cls"
    End Select
    GetVBEExtByType = strExt
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetSafeFileName
' Author    : Adam Waller
' Date      : 1/14/2019
' Purpose   : Replace illegal filename characters with URL encoded substitutes
'           : Sources: http://stackoverflow.com/questions/1976007/what-characters-are-forbidden-in-windows-and-linux-directory-names
'---------------------------------------------------------------------------------------
'
Public Function GetSafeFileName(strName As String) As String

    Dim strSafe As String

    ' Use URL encoding for these characters
    ' https://www.w3schools.com/tags/ref_urlencode.asp
    strSafe = Replace(strName, "%", "%25")  ' This should be done first.
    strSafe = Replace(strSafe, "<", "%3C")
    strSafe = Replace(strSafe, ">", "%3E")
    strSafe = Replace(strSafe, ":", "%3A")
    strSafe = Replace(strSafe, """", "%22")
    strSafe = Replace(strSafe, "/", "%2F")
    strSafe = Replace(strSafe, "\", "%5C")
    strSafe = Replace(strSafe, "|", "%7C")
    strSafe = Replace(strSafe, "?", "%3F")
    strSafe = Replace(strSafe, "*", "%2A")

    ' Return the sanitized file name.
    GetSafeFileName = strSafe

End Function


'---------------------------------------------------------------------------------------
' Procedure : GetOriginalFromSafeName
' Author    : Adam Waller
' Date      : 2/27/2025
' Purpose   : Return the name or path after translating the HTML encoding back to normal
'           : file name characters.
'---------------------------------------------------------------------------------------
'
Public Function GetOriginalFromSafeName(strSafeName As String) As String

    Dim strName As String

    strName = strSafeName
    ' Make sure the following list matches the one above.
    strName = Replace(strName, "%3C", "<")
    strName = Replace(strName, "%3E", ">")
    strName = Replace(strName, "%3A", ":")
    strName = Replace(strName, "%22", """")
    strName = Replace(strName, "%2F", "/")
    strName = Replace(strName, "%5C", "\")
    strName = Replace(strName, "%7C", "|")
    strName = Replace(strName, "%3F", "?")
    strName = Replace(strName, "%2A", "*")
    strName = Replace(strName, "%25", "%")  ' This should be done last.

    ' Return the object name
    GetOriginalFromSafeName = strName

End Function


'---------------------------------------------------------------------------------------
' Procedure : GetObjectNameFromFileName
' Author    : Adam Waller
' Date      : 5/6/2020
' Purpose   : Return the object name after translating the HTML encoding back to normal
'           : file name characters.
'---------------------------------------------------------------------------------------
'
Public Function GetObjectNameFromFileName(strFile As String) As String
    GetObjectNameFromFileName = GetOriginalFromSafeName(FSO.GetBaseName(strFile))
End Function


'---------------------------------------------------------------------------------------
' Procedure : DatesClose
' Author    : Adam Waller
' Date      : 10/13/2017
' Purpose   : Returns true if the dates are within the threshhold.
'           : (Used when dates are very similar, but not exact)
'---------------------------------------------------------------------------------------
'
Public Function DatesClose(dte1 As Date, dte2 As Date, Optional lngMaxDiffSeconds As Long = 3) As Boolean
    DatesClose = (Abs(DateDiff("s", dte1, dte2)) < lngMaxDiffSeconds)
End Function


'---------------------------------------------------------------------------------------
' Procedure : QuickSort
' Author    : Stack Overflow
' Date      : 5/8/2020
' Purpose   : Adapted from https://stackoverflow.com/a/152325/4121863
' Usage     : QuickSort MyArray
'---------------------------------------------------------------------------------------
'
Public Sub QuickSort(ByRef vArray As Variant, Optional ByVal inLow, Optional ByVal inHi)

    Dim pivot   As Variant
    Dim tmpSwap As Variant
    Dim tmpLow  As Long
    Dim tmpHi   As Long

    If IsMissing(inLow) Then inLow = LBound(vArray)
    If IsMissing(inHi) Then inHi = UBound(vArray)

    tmpLow = inLow
    tmpHi = inHi

    pivot = vArray((inLow + inHi) \ 2)

    While (tmpLow <= tmpHi)
        While (vArray(tmpLow) < pivot And tmpLow < inHi)
            tmpLow = tmpLow + 1
        Wend

        While (pivot < vArray(tmpHi) And tmpHi > inLow)
            tmpHi = tmpHi - 1
        Wend

        If (tmpLow <= tmpHi) Then
            tmpSwap = vArray(tmpLow)
            vArray(tmpLow) = vArray(tmpHi)
            vArray(tmpHi) = tmpSwap
            tmpLow = tmpLow + 1
            tmpHi = tmpHi - 1
        End If
    Wend

    If (inLow < tmpHi) Then QuickSort vArray, inLow, tmpHi
    If (tmpLow < inHi) Then QuickSort vArray, tmpLow, inHi

End Sub


'---------------------------------------------------------------------------------------
' Procedure : Pause
' Author    : Adam Waller
' Date      : 6/3/2020
' Purpose   : Pause the code execution for x seconds.
'---------------------------------------------------------------------------------------
'
Public Sub Pause(sngSeconds As Single)
    If sngSeconds > 0.1 Then Perf.OperationStart "Pause execution"
    Sleep sngSeconds * 1000
    If sngSeconds > 0.1 Then Perf.OperationEnd
End Sub


'---------------------------------------------------------------------------------------
' Procedure : Largest
' Author    : Adam Waller
' Date      : 12/2/2020
' Purpose   : Return the largest of an array of values
'---------------------------------------------------------------------------------------
'
Public Function Largest(ParamArray varValues()) As Variant

    Dim varLargest As Variant
    Dim intCnt As Integer

    For intCnt = LBound(varValues) To UBound(varValues)
        If varLargest < varValues(intCnt) Then varLargest = varValues(intCnt)
    Next intCnt
    Largest = varLargest

End Function


'---------------------------------------------------------------------------------------
' Procedure : ZN
' Author    : Adam Waller
' Date      : 12/2/2020
' Purpose   : Opposite of the NZ function, where we convert an empty string or 0 to null.
'---------------------------------------------------------------------------------------
'
Public Function ZN(varValue As Variant) As Variant
    If varValue = vbNullString Or varValue = 0 Then
        ZN = Null
    Else
        ZN = varValue
    End If
End Function


'---------------------------------------------------------------------------------------
' Procedure : DateTruncToSeconds
' Author    : Adam Waller
' Date      : 3/9/2026
' Purpose   : Truncate a Date value to whole seconds, removing any sub-second
'           : precision from the underlying Double. Access DateModified values
'           : may carry fractional seconds that don't survive serialization.
'---------------------------------------------------------------------------------------
'
Public Function DateTruncToSeconds(dteValue As Date) As Date
    If dteValue = 0 Then Exit Function
    DateTruncToSeconds = CDate(Fix(CDbl(dteValue) * 86400#) / 86400#)
End Function


'---------------------------------------------------------------------------------------
' Procedure : ZNDate
' Author    : Adam Waller
' Date      : 12/4/2020
' Purpose   : Return null for an empty date value
'---------------------------------------------------------------------------------------
'
Public Function ZNDate(varValue As Variant) As Variant
    Dim blnDateValue As Boolean
    If IsDate(varValue) Then blnDateValue = (CDate(varValue) <> 0)
    If blnDateValue Then
        ZNDate = varValue
    Else
        ZNDate = Null
    End If
End Function


'---------------------------------------------------------------------------------------
' Procedure : Nz2
' Author    : Adam Waller
' Date      : 2/18/2021
' Purpose   : Extend the NZ function to also include 0 or empty string.
'---------------------------------------------------------------------------------------
'
Public Function Nz2(varValue, Optional varIfNull) As Variant
    Select Case varValue
        Case vbNullString, 0
            If IsMissing(varIfNull) Then
                Nz2 = vbNullString
            Else
                Nz2 = varIfNull
            End If
        Case Else
            If IsNull(varValue) Then
                Nz2 = varIfNull
            Else
                Nz2 = varValue
            End If
    End Select
End Function


'---------------------------------------------------------------------------------------
' Procedure : InArray
' Author    : Adam Waller
' Date      : 5/16/2023
' Purpose   : Returns true if the matching value is found in the array.
'---------------------------------------------------------------------------------------
'
Public Function InArray(varArray, varValue, Optional intCompare As VbCompareMethod = vbBinaryCompare) As Boolean

    Dim lngCnt As Long

    ' Guard clauses
    If Not IsArray(varArray) Then Exit Function
    If IsEmptyArray(varArray) Then Exit Function

    ' Loop through array items, looking for a match
    For lngCnt = LBound(varArray) To UBound(varArray)
        If TypeName(varValue) = "String" Then
            ' Use specified compare method
            If StrComp(varArray(lngCnt), varValue, intCompare) = 0 Then
                InArray = True
                Exit For
            End If
        Else
            ' Compare non-string value
            If varValue = varArray(lngCnt) Then
                ' Found exact match
                InArray = True
                Exit For
            End If
        End If
    Next lngCnt

End Function


'---------------------------------------------------------------------------------------
' Procedure : AddToArray
' Author    : Adam Waller
' Date      : 5/8/2023
' Purpose   : Extends the array by one, and adds the new element to the last segment.
'---------------------------------------------------------------------------------------
'
Public Function AddToArray(ByRef varArray As Variant, varNewElement As Variant)
    ' See if we have defined an index yet
    If IsEmptyArray(varArray) Then
        ' Add first index to array
        ReDim varArray(0)
    Else
        ' Expand array by one element while preserving existing values
        ReDim Preserve varArray(LBound(varArray) To UBound(varArray) + 1)
    End If
    varArray(UBound(varArray)) = varNewElement
End Function


'---------------------------------------------------------------------------------------
' Procedure : IsEmptyArray
' Author    : Adam Waller
' Date      : 3/25/2025
' Purpose   : Return true if the passed array is empty, meaning it does not have any
'           : indexes defined. (Unfortunately we have to use on error resume next to
'           : trap the error when accessing the index.)
'           : This also has the side affect that it creates an array index for object
'           : arrays. See issue #610 for details.
'---------------------------------------------------------------------------------------
'
Public Function IsEmptyArray(varArray As Variant) As Boolean

    Dim lngLowerBound As Long
    Dim lngUpperBound As Long

    ' Exit (returning False) if we are not dealing with an array variable
    If Not IsArray(varArray) Then Exit Function

    LogUnhandledErrors
    On Error Resume Next

    ' Attempt to read the lower bound of the array
    lngLowerBound = 0
    lngUpperBound = -1
    lngLowerBound = LBound(varArray)
    lngUpperBound = UBound(varArray)

    ' Clear any error thrown while attempting to read LBound()
    If Err Then Err.Clear

    ' If the above assignment fails, we have an empty array
    ' (In the case of empty object arrays, this may be reflected
    '  as Ubound = -1, LBound = 0
    IsEmptyArray = (lngUpperBound < lngLowerBound)

End Function


'---------------------------------------------------------------------------------------
' Procedure : BitSet
' Author    : Adam Waller
' Date      : 5/19/2020
' Purpose   : Returns true if the flag is set.
'---------------------------------------------------------------------------------------
'
Public Function BitSet(lngFlags As Long, lngValue As Long) As Boolean
    BitSet = CBool((lngFlags And lngValue) = lngValue)
End Function


'---------------------------------------------------------------------------------------
' Procedure : SwapExtension
' Author    : Adam Waller
' Date      : 8/9/2023
' Purpose   : Return the file path with a different file extension.
'           : I.e.  c:\test.bas > c:\test.cls
'---------------------------------------------------------------------------------------
'
Public Function SwapExtension(strFilePath As String, strNewExtensionWithoutDelimiter As String) As String
    Dim strCurrentExt As String
    strCurrentExt = FSO.GetExtensionName(strFilePath)
    SwapExtension = Left(strFilePath, Len(strFilePath) - Len(strCurrentExt)) & strNewExtensionWithoutDelimiter
End Function


'---------------------------------------------------------------------------------------
' Procedure : ModuleName
' Author    : Adam Waller
' Date      : 3/2/2023
' Purpose   : Dynamically return the class name from a class object.
'           : (This way we don't have to maintain name constants in class modules.)
'---------------------------------------------------------------------------------------
'
Public Function ModuleName(clsMe As Object) As String
    ModuleName = TypeName(clsMe)
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetOfficeBitness
' Author    : Adam Waller
' Date      : 3/5/2022
' Purpose   : Returns "32" or "64" as the bitness of Microsoft Office (not Windows)
'---------------------------------------------------------------------------------------
'
Public Function GetOfficeBitness() As String
    #If Win64 Then
        ' 64-bit add-in (Office x64)
        GetOfficeBitness = "64"
    #Else
        ' 32-bit add-in
        GetOfficeBitness = "32"
    #End If
End Function


'---------------------------------------------------------------------------------------
' Procedure : ExpandEnvironmentVariables
' Author    : Adam Waller
' Date      : 2/12/2024
' Purpose   : Expand out environment variables in a string.
'---------------------------------------------------------------------------------------
'
Public Function ExpandEnvironmentVariables(strString) As String

    Dim lngPos As Long
    Dim lngEnd As Long
    Dim strVariable As String
    Dim strNew As String
    Dim strValue As String

    ' Prepare return value
    strNew = strString

    ' Find pairs of % characters
    Do
        lngPos = InStr(lngPos + 1, strString, "%")
        If lngPos = 0 Then
            Exit Do
        Else
            lngEnd = InStr(lngPos + 2, strString, "%")
            If lngEnd > 0 Then
                ' Found a pair of delimiters. Check the value
                strVariable = Mid$(strString, lngPos + 1, (lngEnd - lngPos) - 1)
                strValue = Environ$(strVariable)
                If Len(strValue) Then
                    ' Replace with expanded value
                    strNew = Replace(strNew, "%" & strVariable & "%", strValue)
                End If
            Else
                lngEnd = lngPos
            End If
        End If
        lngPos = lngEnd
    Loop

    ' Return string with any changes
    ExpandEnvironmentVariables = strNew

End Function
