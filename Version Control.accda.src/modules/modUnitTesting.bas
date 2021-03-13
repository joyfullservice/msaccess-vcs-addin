﻿Option Compare Database
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


'@TestMethod("Encryption")
Private Sub TestSecureBetween()
    On Error GoTo TestFail
    
    'Arrange:
    Dim expNonEncrypted As String
    Dim actnonEncrypted As String
    Dim expRemove As String
    Dim actRemove As String
    Dim expEncrypted As String
    Dim actEncrypted As String
    
    expNonEncrypted = "<firsttag>this should be not be encrypted</firsttag>"
    expRemove = "<firsttag></firsttag>"
    expEncrypted = "<firsttag>@{*"
    
    'Act:
    Options.Security = esNone
    actnonEncrypted = SecureBetween(expNonEncrypted, "<firsttag>", "</firsttag>")
    
    Options.Security = esRemove
    actRemove = SecureBetween(expNonEncrypted, "<firsttag>", "</firsttag>")
    
    Options.Security = esEncrypt
    actEncrypted = SecureBetween(expNonEncrypted, "<firsttag>", "</firsttag>")
    
    Debug.Print actnonEncrypted
    Debug.Print actRemove
    Debug.Print actEncrypted
    
    'Assert:
    Assert.AreEqual expNonEncrypted, actnonEncrypted
    Assert.AreEqual expRemove, actRemove
    Assert.IsTrue actEncrypted Like expEncrypted
    

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Concat")
Private Sub TestConcat()
    
    Dim objConcat As clsConcat
    
    Set objConcat = New clsConcat
    objConcat.SelfTest
    Set objConcat = Nothing
End Sub