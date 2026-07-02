Attribute VB_Name = "modCollectionUtil"
'---------------------------------------------------------------------------------------
' Module    : modCollectionUtil
' Author    : Adam Waller
' Date      : 12/4/2020
' Purpose   : Helper functions for Collection and Dictionary objects: lookup, merge,
'           : sort, compare, clone, and nested path access.
' Layer     : Utility
' Depends on: modConstants, modObjects (Perf)
'---------------------------------------------------------------------------------------
Option Compare Database
Option Private Module
Option Explicit
'@Folder("Utility")


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
' Procedure : DictionaryEqual
' Author    : Adam Waller
' Date      : 6/2/2020
' Purpose   : Returns true if the two dictionary objects are equal in values to each
'           : other, including nested values. Testing the quickest comparisons first
'           : to make the function as performant as possible.
'---------------------------------------------------------------------------------------
'
Public Function DictionaryEqual(dSource As Dictionary, dCompare As Dictionary) As Boolean

    Dim strSource As String
    Dim strCompare As String
    Dim blnEqual As Boolean

    Perf.OperationStart "Compare Dictionary"
    If dSource Is Nothing And dCompare Is Nothing Then
        ' Neither object set.
        blnEqual = True
    ElseIf Not dSource Is Nothing And Not dCompare Is Nothing Then
        ' Both are objects. Check count property.
        If dSource.Count = dCompare.Count Then
            strSource = ConvertToJson(dSource)
            strCompare = ConvertToJson(dCompare)
            ' Compare string length
            If Len(strSource) = Len(strCompare) Then
                ' Perform a binary (case-sensitive) comparison of strings.
                blnEqual = (StrComp(strSource, strCompare, vbBinaryCompare) = 0)
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
