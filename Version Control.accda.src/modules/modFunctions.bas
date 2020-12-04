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
' Procedure : Shell2
' Author    : Adam Waller
' Date      : 6/3/2015
' Purpose   : Alternative to VBA Shell command, to work around issues with the
'           : TortoiseSVN command line for commits.
'---------------------------------------------------------------------------------------
'
Public Sub Shell2(strCmd As String)
    Dim objShell As WshShell
    Set objShell = New WshShell
    objShell.Exec strCmd
    Set objShell = Nothing
End Sub


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
Public Function MsgBox2(strBold As String, Optional strLine1 As String, Optional strLine2 As String, Optional intButtons As VbMsgBoxStyle = vbOKOnly, Optional strTitle As String) As VbMsgBoxResult
    
    Dim strMsg As String
    Dim varLines(0 To 3) As String
    
    ' Escape single quotes by doubling them.
    varLines(0) = Replace(strBold, "'", "''")
    varLines(1) = Replace(strLine1, "'", "''")
    varLines(2) = Replace(strLine2, "'", "''")
    varLines(3) = Replace(strTitle, "'", "''")
    
    If varLines(3) = vbNullString Then varLines(3) = Application.VBE.ActiveVBProject.Name
    strMsg = "MsgBox('" & varLines(0) & "@" & varLines(1) & "@" & varLines(2) & "@'," & intButtons & ",'" & varLines(3) & "')"
    MsgBox2 = Eval(strMsg)
    
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
                If varSegment.Exists(strKey) And Not IsEmpty(varSegment(strKey)) Then
                    Set varSegment = varSegment(strKey)
                Else
                    ' Path not found
                    Exit For
                End If
            End If
        End If
    Next intCnt

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
    Dim varItem As Variant
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
    
    Dim sourceCount As Long
    sourceCount = dSource.Count - 1
    
    ' Build and sort array of keys
    ReDim varKeys(0 To sourceCount)
    For Each varKey In dSource.Keys
        varKeys(lngCnt) = varKey
        lngCnt = lngCnt + 1
    Next varKey
    
    QuickSort varKeys, 0, sourceCount
    
    ' Build and return new dictionary using sorted keys
    Set dSorted = New Dictionary
    For lngCnt = 0 To sourceCount
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
Public Sub QuickSort(ByRef vArray As Variant, Optional ByVal inLow As Long, Optional ByVal inHi As Long)

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
    
    ' Return comparison result
    DictionaryEqual = blnEqual
    
End Function


'---------------------------------------------------------------------------------------
' Procedure : Pause
' Author    : Adam Waller
' Date      : 6/3/2020
' Purpose   : Pause the code execution for x seconds.
'---------------------------------------------------------------------------------------
'
Public Sub Pause(sngSeconds As Single)
    Sleep sngSeconds * 1000
End Sub


'---------------------------------------------------------------------------------------
' Procedure : Catch
' Author    : Adam Waller
' Date      : 11/23/2020
' Purpose   : Returns true if the last error matches any of the passed error numbers,
'           : and clears the error object.
'---------------------------------------------------------------------------------------
'
Public Function Catch(ParamArray lngErrorNumbers()) As Boolean
    Dim intCnt As Integer
    For intCnt = LBound(lngErrorNumbers) To UBound(lngErrorNumbers)
        If lngErrorNumbers(intCnt) = Err.Number Then
            Err.Clear
            Catch = True
            Exit For
        End If
    Next intCnt
End Function


'---------------------------------------------------------------------------------------
' Procedure : CatchAny
' Author    : Adam Waller
' Date      : 12/3/2020
' Purpose   : Generic error handler with logging.
'---------------------------------------------------------------------------------------
'
Public Function CatchAny(eLevel As eErrorLevel, strDescription As String, Optional strSource As String, _
    Optional blnLogError As Boolean = True, Optional blnClearError As Boolean = True) As Boolean
    If Err Then
        If blnLogError Then Log.Error eLevel, strDescription, strSource
        If blnClearError Then Err.Clear
        CatchAny = True
    End If
End Function


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