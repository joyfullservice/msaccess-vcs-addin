Attribute VB_Name = "modUnitTesting"
Option Compare Database
Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests")

Private Assert As Object
Private Fakes As Object

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = CreateObject("Rubberduck.AssertClass")
    Set Fakes = CreateObject("Rubberduck.FakesProvider")
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    Set Fakes = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    'this method runs before every test in the module.
End Sub

'@TestCleanup
Private Sub TestCleanup()
    'this method runs after every test in the module.
End Sub

' Test shows that UCS-2 files exported by Access make round trip through our conversions.
'@TestMethod("TextConversions")
Public Sub TestUCS2toUTF8RoundTrip()
    On Error GoTo TestFail
    
    'Arrange:
    Dim queryName As String
    queryName = "Temp_Test_Query_Delete_Me_Æ_ø_Å"
    Dim tempFileName As String
    tempFileName = GetTempFile()
    
    Dim UCStoUCS As String
    Dim UCStoUTF As String
    Dim UTFtoUTF As String
    Dim UTFtoUCS As String
    UCStoUCS = tempFileName & "UCS-2toUCS-2"
    UCStoUTF = tempFileName & "UCS-2toUTF-8"
    UTFtoUTF = tempFileName & "UTF-8toUTF-8"
    UTFtoUCS = tempFileName & "UTF-8toUCS-2"
    
    ' Use temporary query to export example file
    CurrentDb.CreateQueryDef queryName, "SELECT * FROM TEST WHERE TESTING='ÆØÅ'"
    Application.SaveAsText acQuery, queryName, tempFileName
    CurrentDb.QueryDefs.Delete queryName
        
    ' Read original export
    Dim originalExport As String
    With FSO.OpenTextFile(tempFileName, ForReading, False, TristateTrue)
        originalExport = .ReadAll
        .Close
    End With
            
    'Act:
    ConvertUtf8Ucs2 tempFileName, UCStoUCS
    ConvertUcs2Utf8 UCStoUCS, UCStoUTF
    ConvertUcs2Utf8 UCStoUTF, UTFtoUTF
    ConvertUtf8Ucs2 UTFtoUTF, UTFtoUCS
    
    ' Read final file that went through all permutations of conversion
    Dim finalFile As String
    With FSO.OpenTextFile(UTFtoUCS, ForReading, False, TristateTrue)
        finalFile = .ReadAll
        .Close
    End With
    
    ' Cleanup temp files
    'deletefile tempFileName
    'deletefile UTFtoUCS
    
    'Assert:
    Assert.AreEqual originalExport, finalFile
    
    GoTo TestExit
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description

TestExit:
    
End Sub

'@TestMethod("TextConversion")
Private Sub TestParseSpecialCharsInJson()
    On Error GoTo TestFail
       
    'Arrange:
    Dim strPath As String
    Dim dict As Dictionary
    Dim FSO As Object
    
    strPath = GetTempFile
        
    Set FSO = CreateObject("Scripting.FileSystemObject")
    With FSO.CreateTextFile(strPath, True)
        .WriteLine "{""Test"":""ÆØÅ are special?""}"
        .Close
    End With
    
    Debug.Print strPath
    
    'Act:
    Set dict = modFileAccess.ReadJsonFile(strPath)
    
    'Assert:
    If dict Is Nothing Then
        Assert.Fail "Empty dictionary returned"
    Else
        Debug.Print dict("Test")
        Assert.Succeed
    End If
    

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Sorting")
Private Sub TestSortDictionaryByKeys()
    On Error GoTo TestFail
    
    'Arrange:
    Dim dItems As Dictionary
    
    Set dItems = New Dictionary
    dItems.Add "C", "C"
    dItems.Add "A", "A"
    dItems.Add "B", "B"
    
    'Act:
    Set dItems = SortDictionaryByKeys(dItems)
    
    'Assert:
    Assert.AreEqual dItems.Items(0), "A"
    Assert.AreEqual dItems.Items(1), "B"
    Assert.AreEqual dItems.Items(2), "C"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("QuickSort")
Private Sub TestQuickSort()
    
    Dim arr() As String
    Dim result As String
    
    arr = Split("u i a")
    
    QuickSort arr
    result = Join(arr, " ")
    Assert.AreEqual result, "a i u"
    
End Sub


'@TestMethod("Concat")
Private Sub TestConcat()
    
    With New clsConcat
        .SelfTest
    End With
    
End Sub


'@TestMethod("SanitizeConnectionString")
Private Sub TestSanitizeConnectionString()

    ' Verify semicolon placement matches original
    Debug.Assert SanitizeConnectionString(";test;test;") = ";test;test;"
    Debug.Assert SanitizeConnectionString("test;test") = "test;test"
    Debug.Assert SanitizeConnectionString(";test;test") = ";test;test"
    Debug.Assert SanitizeConnectionString("test;test;") = "test;test;"
    Debug.Assert SanitizeConnectionString("test;test;") = "test;test;"
    Debug.Assert SanitizeConnectionString("test") = "test"
    Debug.Assert SanitizeConnectionString(vbNullString) = vbNullString

End Sub


'@TestMethod("CloneDictionary")
Private Sub TestCloneDictionary()

    Dim dFruit As Dictionary
    Dim dApple As Dictionary
    Dim dClone As Dictionary
    
    Set dFruit = New Dictionary
    Set dApple = New Dictionary
    
    ' Create text compare dictionary
    With dApple
        .CompareMode = TextCompare
        .Add "SEED1", "Apple Seed"
        .Add "seed2", "Apple Seed"
    End With
    
    ' Create binary compare dictionary with nested dictionary
    With dFruit
        .CompareMode = BinaryCompare
        .Add "Apple", dApple
        .Add "Orange", "Orange"
        .Add "Pear", "Pear"
    End With
    
    ' Clone the dictionary
    Set dClone = CloneDictionary(dFruit, ecmSourceMethod)
    
    ' Test the results to make sure it cloned correctly.
    Debug.Assert dClone.Exists("APPLE") = False
    Debug.Assert dClone.Exists("Apple") = True
    Debug.Assert dClone.Exists("ORANGE") = False
    Debug.Assert dClone.Exists("Orange") = True
    Debug.Assert dClone("Apple").CompareMode = Scripting.CompareMethod.TextCompare
    Debug.Assert dClone("Apple").Exists("seed1") = True
    Debug.Assert dClone("Apple").Exists("SEED1") = True
    Debug.Assert dClone("Apple").Exists("Seed3") = False
    
End Sub
