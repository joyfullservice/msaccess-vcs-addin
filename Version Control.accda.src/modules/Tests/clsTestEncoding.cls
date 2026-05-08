VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTestEncoding"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : clsTestEncoding
' Author    : Adam Waller
' Date      : 5/8/2026
' Purpose   : Encoding and file-hashing tests extracted into a class to exercise the
'           : class-based test runner and provide natural setup/teardown for temp files.
'           :
'           : Class_Initialize creates a temp folder for test artifacts.
'           : Class_Terminate removes it, ensuring no temp file leaks.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
'@Folder("Tests")

Private m_strTempFolder As String


Private Sub Class_Initialize()
    m_strTempFolder = GetTempFile() & "_enc_tests\"
    On Error Resume Next
    If FSO.FolderExists(m_strTempFolder) Then
        FSO.DeleteFolder StripSlash(m_strTempFolder), True
    End If
    On Error GoTo 0
    VerifyPath m_strTempFolder
End Sub


Private Sub Class_Terminate()
    On Error Resume Next
    If FSO.FolderExists(m_strTempFolder) Then
        FSO.DeleteFolder StripSlash(m_strTempFolder), True
    End If
    On Error GoTo 0
End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestUCS2toUTF8RoundTrip
' Author    : Adam Waller
' Date      : 5/8/2026
' Purpose   : UCS-2 files exported by Access survive round trip through our conversions.
'---------------------------------------------------------------------------------------
'
Public Sub TestUCS2toUTF8RoundTrip()

    Dim strQueryName As String
    Dim strTempFile As String
    Dim strUCStoUCS As String
    Dim strUCStoUTF As String
    Dim strUTFtoUTF As String
    Dim strUTFtoUCS As String
    Dim strOriginal As String
    Dim strFinal As String

    strQueryName = "Temp_Test_Query_Delete_Me_Æ_ø_Å"
    strTempFile = m_strTempFolder & "export.txt"
    strUCStoUCS = m_strTempFolder & "UCS-2toUCS-2"
    strUCStoUTF = m_strTempFolder & "UCS-2toUTF-8"
    strUTFtoUTF = m_strTempFolder & "UTF-8toUTF-8"
    strUTFtoUCS = m_strTempFolder & "UTF-8toUCS-2"

    ' Use temporary query to export example file
    CurrentDb.CreateQueryDef strQueryName, "SELECT * FROM TEST WHERE TESTING='ÆØÅ'"
    Application.SaveAsText acQuery, strQueryName, strTempFile
    CurrentDb.QueryDefs.Delete strQueryName

    ' Read original export
    With FSO.OpenTextFile(strTempFile, ForReading, False, TristateTrue)
        strOriginal = .ReadAll
        .Close
    End With

    ConvertUtf8Ucs2 strTempFile, strUCStoUCS
    ConvertUcs2Utf8 strUCStoUCS, strUCStoUTF
    ConvertUcs2Utf8 strUCStoUTF, strUTFtoUTF
    ConvertUtf8Ucs2 strUTFtoUTF, strUTFtoUCS

    ' Read final file that went through all permutations of conversion
    With FSO.OpenTextFile(strUTFtoUCS, ForReading, False, TristateTrue)
        strFinal = .ReadAll
        .Close
    End With

    TestAssert strOriginal = strFinal, "UCS-2 round trip should produce identical output"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestStringFileHash
' Author    : Adam Waller
' Date      : 5/8/2026
' Purpose   : Verify consistent results between hashing a string and hashing a file.
'---------------------------------------------------------------------------------------
'
Public Sub TestStringFileHash()

    Const cstrText As String = "This is my text content."
    Dim strTempFile As String

    strTempFile = m_strTempFolder & "hash_test.txt"
    WriteFile cstrText, strTempFile

    TestAssert GetStringHash(cstrText) = "f80a555", "String hash without BOM"
    TestAssert GetStringHash(cstrText, True) = "b628391", "String hash with UTF-8 BOM"
    TestAssert GetFileHash(strTempFile) = GetStringHash(cstrText, True), "File hash matches string hash with BOM"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestParseSpecialCharsInJson
' Author    : Adam Waller
' Date      : 5/8/2026
' Purpose   : Verify that special characters (ÆØÅ) survive JSON round trip.
'---------------------------------------------------------------------------------------
'
Public Sub TestParseSpecialCharsInJson()

    Dim strPath As String
    Dim dict As Dictionary

    strPath = m_strTempFolder & "special_chars.json"

    With FSO.CreateTextFile(strPath, True)
        .WriteLine "{""Test"":""ÆØÅ are special?""}"
        .Close
    End With

    Set dict = modFileAccess.ReadJsonFile(strPath)

    TestAssert Not dict Is Nothing, "JSON file parsed successfully"
    TestAssert dict("Test") = "ÆØÅ are special?", "Special characters preserved"

End Sub
