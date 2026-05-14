Attribute VB_Name = "modTestCollectionUtil"
'---------------------------------------------------------------------------------------
' Module    : modTestCollectionUtil
' Author    : Adam Waller
' Date      : 5/12/2026
' Purpose   : Unit tests for modCollectionUtil dictionary/collection helpers.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
Option Private Module
'@Folder("Tests.Utility")


Public Sub TestSortDictionaryByKeys()
    Dim dItems As Dictionary
    Set dItems = New Dictionary
    dItems.Add "C", "C"
    dItems.Add "A", "A"
    dItems.Add "B", "B"
    Set dItems = SortDictionaryByKeys(dItems)
    TestAssert dItems.Keys()(0) = "A", "first key is A"
    TestAssert dItems.Keys()(1) = "B", "second key is B"
    TestAssert dItems.Keys()(2) = "C", "third key is C"
End Sub


Public Sub TestSortDictionaryByKeys_SingleItem()
    Dim dItems As Dictionary
    Set dItems = New Dictionary
    dItems.Add "only", "value"
    Set dItems = SortDictionaryByKeys(dItems)
    TestAssert dItems.Count = 1, "single item unchanged"
    TestAssert dItems.Keys()(0) = "only", "key preserved"
End Sub


Public Sub TestCloneDictionary()
    Dim dOrig As Dictionary
    Dim dChild As Dictionary
    Dim dClone As Dictionary

    Set dChild = New Dictionary
    dChild.CompareMode = TextCompare
    dChild.Add "key1", "value1"

    Set dOrig = New Dictionary
    dOrig.Add "nested", dChild
    dOrig.Add "flat", "hello"

    Set dClone = CloneDictionary(dOrig)

    TestAssert dClone("flat") = "hello", "flat value cloned"
    TestAssert dClone("nested")("key1") = "value1", "nested value cloned"

    ' Modify clone and verify original is unaffected
    dClone("nested")("key1") = "modified"
    TestAssert dOrig("nested")("key1") = "value1", "original unaffected by clone mutation"
End Sub


Public Sub TestCloneDictionary_PreservesCompareMode()
    Dim dOrig As Dictionary
    Set dOrig = New Dictionary
    dOrig.CompareMode = TextCompare
    dOrig.Add "Key", "val"

    Dim dClone As Dictionary
    Set dClone = CloneDictionary(dOrig, ecmSourceMethod)
    TestAssert dClone.CompareMode = TextCompare, "compare mode preserved"
    TestAssert dClone.Exists("KEY"), "text compare mode active"
End Sub


Public Sub TestMergeDictionary()
    Dim d1 As Dictionary
    Dim d2 As Dictionary
    Set d1 = New Dictionary
    Set d2 = New Dictionary
    d1.Add "a", 1
    d1.Add "b", 2
    d2.Add "b", 99
    d2.Add "c", 3
    MergeDictionary d1, d2
    TestAssert d1("a") = 1, "existing key preserved"
    TestAssert d1("b") = 99, "overlapping key replaced"
    TestAssert d1("c") = 3, "new key added"
End Sub


Public Sub TestDictionaryEqual()
    Dim d1 As Dictionary
    Dim d2 As Dictionary
    Set d1 = New Dictionary
    Set d2 = New Dictionary
    d1.Add "x", 1
    d2.Add "x", 1
    TestAssert DictionaryEqual(d1, d2), "identical dictionaries"

    d2("x") = 2
    TestAssert Not DictionaryEqual(d1, d2), "different values"

    TestAssert DictionaryEqual(Nothing, Nothing), "both nothing"
    TestAssert Not DictionaryEqual(d1, Nothing), "one nothing"
End Sub


Public Sub TestInCollection()
    Dim col As New Collection
    col.Add "apple"
    col.Add "banana"
    TestAssert InCollection(col, "apple"), "found"
    TestAssert Not InCollection(col, "cherry"), "not found"
End Sub


Public Sub TestKeyExists()
    Dim dOuter As Dictionary
    Dim dInner As Dictionary
    Set dOuter = New Dictionary
    Set dInner = New Dictionary
    dInner.Add "deep", "value"
    dOuter.Add "level1", dInner
    TestAssert KeyExists(dOuter, "level1", "deep"), "nested key exists"
    TestAssert Not KeyExists(dOuter, "level1", "missing"), "nested key missing"
    TestAssert Not KeyExists(dOuter, "nope"), "top-level key missing"
    TestAssert Not KeyExists(Nothing, "any"), "nothing dictionary"
End Sub


Public Sub TestdNZ()
    Dim dTest As Dictionary
    Dim dNested As Dictionary
    Set dTest = New Dictionary
    Set dNested = New Dictionary
    dNested.Add "name", "Alice"
    dTest.Add "user", dNested
    dTest.Add "flat", "value"

    TestAssert dNZ(dTest, "flat") = "value", "flat key"
    TestAssert dNZ(dTest, "user\name") = "Alice", "nested key"
    TestAssert dNZ(dTest, "missing") = "", "missing returns empty"
    TestAssert dNZ(dTest, "user\missing") = "", "missing nested returns empty"
End Sub
