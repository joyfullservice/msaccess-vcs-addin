Attribute VB_Name = "modFunctions"
'---------------------------------------------------------------------------------------
' Module    : modFunctions
' Author    : Adam Waller
' Date      : 12/4/2020
' Purpose   : General functions that don't fit more specifically into another module.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Private Module
Option Explicit


' API function to pause processing
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)

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
' Procedure : InCollection
' Author    : Adam Waller
' Date      : 6/2/2015
' Purpose   : Returns true if the item value is found in the collection
'---------------------------------------------------------------------------------------
'
Public Function InCollection(ByVal MyCol As Collection, ByVal MyValue As Variant) As Boolean
    Dim intCnt As Integer
    For intCnt = 1 To MyCol.Count
        If MyCol(intCnt) = MyValue Then
            InCollection = True
            Exit For
        End If
    Next intCnt
End Function


'---------------------------------------------------------------------------------------
' Procedure : MergeCollection
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Adds a collection into another collection.
'---------------------------------------------------------------------------------------
'
Public Sub MergeCollection(ByRef colOriginal As Collection, ByVal colToAdd As Collection)
    Dim varItem As Variant
    For Each varItem In colToAdd
        colOriginal.Add varItem
    Next varItem
End Sub


'---------------------------------------------------------------------------------------
' Procedure : MergeDictionary
' Author    : Adam Waller
' Date      : 7/31/2023
' Purpose   : Merge a dictionary into another dictionary. Existing values are replaced
'           : from the incoming dictionary.
'---------------------------------------------------------------------------------------
'
Public Sub MergeDictionary(ByRef dOriginal As Dictionary, ByVal dToAdd As Dictionary)
    Dim varKey As Variant
    For Each varKey In dToAdd.Keys
        If IsObject(dToAdd(varKey)) Then
            Set dOriginal(varKey) = dToAdd(varKey)
        Else
            dOriginal(varKey) = dToAdd(varKey)
        End If
    Next varKey
End Sub


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
' Procedure : GetObjectNameFromFileName
' Author    : Adam Waller
' Date      : 5/6/2020
' Purpose   : Return the object name after translating the HTML encoding back to normal
'           : file name characters.
'---------------------------------------------------------------------------------------
'
Public Function GetObjectNameFromFileName(strFile As String) As String

    Dim strName As String

    strName = FSO.GetBaseName(strFile)
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
    GetObjectNameFromFileName = strName

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
' Procedure : MultiReplace
' Author    : Adam Waller
' Date      : 1/18/2019
' Purpose   : Does a string replacement of multiple items in one call.
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

    ' Check interaction mode
    If InteractionMode = eimNormal Then
        ' Normal user interaction with MsgBox
        If varLines(3) = vbNullString Then varLines(3) = Application.VBE.ActiveVBProject.Name
        strMsg = "MsgBox('" & varLines(0) & "@" & varLines(1) & "@" & varLines(2) & "@'," & intButtons & ",'" & varLines(3) & "')"
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
' Procedure : dNZ
' Author    : Adam Waller
' Date      : 3/23/2020
' Purpose   : Like the NZ function but for dictionary elements
'---------------------------------------------------------------------------------------
'
Public Function dNZ(dObject As Dictionary, strPath As String, Optional strDelimiter As String = "\") As String

    Dim varPath As Variant
    Dim intCnt As Integer
    Dim dblVal As Double
    Dim strKey As String
    Dim varSegment As Variant

    ' Split path into parts
    varPath = Split(strPath, strDelimiter)
    Set varSegment = dObject

    For intCnt = LBound(varPath) To UBound(varPath)

        strKey = varPath(intCnt)
        If dObject Is Nothing Then
            ' No object found
            Exit For
        ElseIf TypeOf varSegment Is Collection Then
            ' Expect index (integer)
            If IsNumeric(strKey) Then
                ' Looks like an array index
                dblVal = CDbl(strKey)
                ' Do a couple more checks to see if this looks like a valid index
                If dblVal < 1 Or dblVal > 32000 Or dblVal <> CInt(dblVal) Then Exit For
                ' See if this is the last segment
                If intCnt = UBound(varPath) Then
                    If TypeOf varSegment(dblVal) Is Dictionary Then
                        ' Need a named key
                        Exit For
                    Else
                        ' Could be an array of values
                        dNZ = Nz(varSegment(dblVal))
                    End If
                Else
                    ' Move out to next segment
                    Set varSegment = varSegment(dblVal)
                End If
            End If
        ElseIf TypeOf varSegment Is Dictionary Then
            ' Expect key (string)
            If intCnt = UBound(varPath) Then
                ' Reached last segment
                If varSegment.Exists(strKey) Then
                    If TypeOf varSegment Is Dictionary Then
                        dNZ = Nz(varSegment(strKey))
                    Else
                        ' Might be array
                        Exit For
                    End If
                End If
            Else
                ' Move out to next segment
                If varSegment.Exists(strKey) Then
                    If Not IsEmpty(varSegment(strKey)) Then
                        Set varSegment = varSegment(strKey)
                    Else
                        ' Empty value
                        Exit For
                    End If
                Else
                    ' Path not found
                    Exit For
                End If
            End If
        End If
    Next intCnt

End Function


'---------------------------------------------------------------------------------------
' Procedure : KeyExists
' Author    : Adam Waller
' Date      : 5/28/2021
' Purpose   : Returns true if the specified nested segment is found to exist.
'           : Note that this currently only supports nested child dictionary objects,
'           : not nested collections.
'---------------------------------------------------------------------------------------
'
Public Function KeyExists(dDictionary As Dictionary, ParamArray varSegmentKeys()) As Boolean

    Dim intSegment As Integer
    Dim dBase As Dictionary

    ' Bail out if no valid dictionary passed
    If dDictionary Is Nothing Then Exit Function

    ' Start with based dictionary
    Set dBase = dDictionary
    KeyExists = True

    ' Loop through segments, confirming that each one exists
    For intSegment = 0 To UBound(varSegmentKeys)
        If dBase.Exists(varSegmentKeys(intSegment)) Then
            If intSegment < UBound(varSegmentKeys) Then
                Set dBase = dBase(varSegmentKeys(intSegment))
            End If
        Else
            KeyExists = False
            Exit For
        End If
    Next intSegment

End Function


'---------------------------------------------------------------------------------------
' Procedure : SortCollection
' Author    : Adam Waller
' Date      : 11/14/2020
' Purpose   : Sort a collection of items by value. (Returns a new sorted collection)
'---------------------------------------------------------------------------------------
'
Public Function SortCollectionByValue(colSource As Collection) As Collection

    Dim colSorted As Collection
    Dim varItems() As Variant
    Dim lngCnt As Long

    ' Don't need to sort empty collection or single item
    If colSource.Count < 2 Then
        Set SortCollectionByValue = colSource
        Exit Function
    End If

    ' Build and sort array of keys
    ReDim varItems(0 To colSource.Count - 1)
    For lngCnt = 0 To UBound(varItems)
        varItems(lngCnt) = colSource(lngCnt + 1)
    Next lngCnt
    QuickSort varItems

    ' Build and return new collection using sorted values
    Set colSorted = New Collection
    For lngCnt = 0 To UBound(varItems)
        colSorted.Add varItems(lngCnt)
    Next lngCnt
    Set SortCollectionByValue = colSorted

End Function


'---------------------------------------------------------------------------------------
' Procedure : SortDictionaryByKeys
' Author    : Adam Waller
' Date      : 5/8/2020
' Purpose   : Rebuilds a dictionary object by adding all the items to a new dictionary
'           : sorted by keys.
'---------------------------------------------------------------------------------------
'
Public Function SortDictionaryByKeys(dSource As Dictionary) As Dictionary

    Dim dSorted As Dictionary
    Dim varKeys() As Variant
    Dim varKey As Variant
    Dim lngCnt As Long

    ' Don't need to sort empty dictionary or single item
    If dSource.Count < 2 Then
        Set SortDictionaryByKeys = dSource
        Exit Function
    End If

    ' Build and sort array of keys
    ReDim varKeys(0 To dSource.Count - 1)
    For Each varKey In dSource.Keys
        varKeys(lngCnt) = varKey
        lngCnt = lngCnt + 1
    Next varKey

    QuickSort varKeys

    ' Build and return new dictionary using sorted keys
    Set dSorted = New Dictionary
    dSorted.CompareMode = dSource.CompareMode
    For lngCnt = 0 To dSource.Count - 1
        dSorted.Add varKeys(lngCnt), dSource(varKeys(lngCnt))
    Next lngCnt

    Set SortDictionaryByKeys = dSorted

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
' Procedure : DictionaryEqual
' Author    : Adam Waller
' Date      : 6/2/2020
' Purpose   : Returns true if the two dictionary objects are equal in values to each
'           : other, including nested values. Testing the quickest comparisons first
'           : to make the function as performant as possible.
'---------------------------------------------------------------------------------------
'
Public Function DictionaryEqual(dOne As Dictionary, dTwo As Dictionary) As Boolean

    Dim strOne As String
    Dim strTwo As String
    Dim blnEqual As Boolean

    Perf.OperationStart "Compare Dictionary"
    If dOne Is Nothing And dTwo Is Nothing Then
        ' Neither object set.
        blnEqual = True
    ElseIf Not dOne Is Nothing And Not dTwo Is Nothing Then
        ' Both are objects. Check count property.
        If dOne.Count = dTwo.Count Then
            strOne = ConvertToJson(dOne)
            strTwo = ConvertToJson(dTwo)
            ' Compare string length
            If Len(strOne) = Len(strTwo) Then
                ' Perform a binary (case-sensitive) comparison of strings.
                blnEqual = (StrComp(strOne, strTwo, vbBinaryCompare) = 0)
            End If
        End If
    End If
    Perf.OperationEnd

    ' Return comparison result
    DictionaryEqual = blnEqual

End Function


'---------------------------------------------------------------------------------------
' Procedure : CloneDictionary
' Author    : Adam Waller
' Date      : 7/21/2023
' Purpose   : Recursive function to deep-clone a dictionary object, including nested
'           : dictionaries.
'           : NOTE: All other object types are cloned as a reference to the same object
'           : referenced by the original dictionary, not a new object.
'---------------------------------------------------------------------------------------
'
Public Function CloneDictionary(dSource As Dictionary, _
    Optional Compare As eCompareMethod2 = ecmSourceMethod) As Dictionary

    Dim dNew As Dictionary
    Dim dChild As Dictionary
    Dim varKey As Variant

    ' No object returned if source is nothing
    If dSource Is Nothing Then Exit Function

    ' Create new dictionary object and set compare mode
    Set dNew = New Dictionary
    If Compare = ecmSourceMethod Then
        ' Use the same compare mode as the original dictionary.
        dNew.CompareMode = dSource.CompareMode
    Else
        dNew.CompareMode = Compare
    End If

    ' Loop through keys
    For Each varKey In dSource.Keys
        If TypeOf dSource(varKey) Is Dictionary Then
            ' Call this function recursively to add nested dictionary
            Set dChild = dSource(varKey)
            dNew.Add varKey, CloneDictionary(dChild, Compare)
        Else
            ' Add key to dictionary
            dNew.Add varKey, dSource(varKey)
        End If
    Next varKey

    ' Return new dictionary
    Set CloneDictionary = dNew

End Function


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
' Procedure : PathSep
' Author    : Adam Waller
' Date      : 3/3/2021
' Purpose   : Return the current path separator, based on language settings.
'           : Caches value to avoid extra calls to FSO object.
'---------------------------------------------------------------------------------------
'
Public Function PathSep() As String
    Static strSeparator As String
    If strSeparator = vbNullString Then strSeparator = Mid$(FSO.BuildPath("a", "b"), 2, 1)
    PathSep = strSeparator
End Function


'---------------------------------------------------------------------------------------
' Procedure : BuildPath2
' Author    : Adam Waller
' Date      : 3/3/2021
' Purpose   : Like FSO.BuildPath, but with unlimited arguments)
'---------------------------------------------------------------------------------------
'
Public Function BuildPath2(ParamArray Segments())
    Dim lngPart As Long
    With New clsConcat
        For lngPart = LBound(Segments) To UBound(Segments)
            .Add CStr(Segments(lngPart))
            If lngPart < UBound(Segments) Then .Add PathSep
        Next lngPart
    BuildPath2 = .GetStr
    End With
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
' Date      : 5/13/2023
' Purpose   : Return true if the passed array is empty, meaning it does not have any
'           : indexes defined. (Unfortunately we have to use on error resume next to
'           : trap the error when accessing the index.)
'---------------------------------------------------------------------------------------
'
Public Function IsEmptyArray(varArray As Variant) As Boolean

    ' Use an arbitrary number extremly unlikely to collide with an existing index
    Const clngTest As Long = -2147483646

    Dim lngLowBound As Long

    ' Exit (returning False) if we are not dealing with an array variable
    If Not IsArray(varArray) Then Exit Function

    LogUnhandledErrors
    On Error Resume Next

    ' Attempt to read the lower bound of the array
    lngLowBound = clngTest
    lngLowBound = LBound(varArray)

    ' Clear any error thrown while attempting to read LBound()
    If Err Then Err.Clear

    ' If the above assignment fails, we have an empty array
    IsEmptyArray = (lngLowBound = clngTest)

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
    StartsWith = (InStr(1, strText, strStartsWith, Compare) = 1)
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
' Procedure : MakeDialogResizable
' Author    : Adam Waller
' Date      : 5/16/2023
' Purpose   : Change the window style of an existing dialog window to make it resizable.
'           : (This allows you to use the acDialog argument when opening a form, but
'           :  still have the form resizable by the user.)
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
