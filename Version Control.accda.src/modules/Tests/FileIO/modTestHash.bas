Attribute VB_Name = "modTestHash"
'---------------------------------------------------------------------------------------
' Module    : modTestHash
' Author    : Adam Waller
' Date      : 5/12/2026
' Purpose   : Unit tests for modHash hashing functions.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
Option Private Module
'@Folder("Tests.FileIO")
'@Tag("io")


Public Sub TestGetStringHash()
    Dim strHash1 As String
    Dim strHash2 As String
    strHash1 = GetStringHash("test content")
    strHash2 = GetStringHash("test content")
    TestAssert Len(strHash1) > 0, "returns non-empty hash"
    TestAssert strHash1 = strHash2, "deterministic (same input = same output)"
End Sub


Public Sub TestGetStringHash_DifferentInputs()
    Dim strHash1 As String
    Dim strHash2 As String
    strHash1 = GetStringHash("input A")
    strHash2 = GetStringHash("input B")
    TestAssert strHash1 <> strHash2, "different inputs produce different hashes"
End Sub


Public Sub TestGetDictionaryHash()
    Dim d1 As Dictionary
    Dim d2 As Dictionary
    Set d1 = New Dictionary
    Set d2 = New Dictionary
    d1.Add "key", "value"
    d2.Add "key", "value"
    TestAssert Len(GetDictionaryHash(d1)) > 0, "returns non-empty hash"
    TestAssert GetDictionaryHash(d1) = GetDictionaryHash(d2), "identical dictionaries same hash"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestKnownSha256Digest
' Author    : Adam Waller
' Date      : 7/29/2026
' Purpose   : Anchor the hex formatting and digest bytes against a published SHA-256
'           : value, so an optimization to either cannot quietly change stored hashes.
'           : (Every index entry and every conflict check depends on these staying
'           : byte-for-byte stable across releases.)
'---------------------------------------------------------------------------------------
'
Public Sub TestKnownSha256Digest()

    Const cstrExpected As String = _
        "9f86d081884c7d659a2feaa0c55ad015a3bf4f1b2b0b822cd15d6c15b0f00a08"

    Dim strHash As String

    If StrComp(Nz2(Options.HashAlgorithm, DefaultHashAlgorithm), "SHA256", vbTextCompare) <> 0 Then
        TestAssert True, "SKIP: project is not configured for SHA256"
        Exit Sub
    End If

    ' UTF-8 bytes of "test", no BOM
    strHash = GetStringHash("test")

    If Options.UseShortHash Then
        TestAssert strHash = Left$(cstrExpected, 7), "short hash matches known SHA-256 prefix"
    Else
        TestAssert strHash = cstrExpected, "matches known SHA-256 of 'test'"
    End If
    TestAssert strHash = LCase$(strHash), "digest is lowercase hex"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestHashProviderCacheStability
' Author    : Adam Waller
' Date      : 7/29/2026
' Purpose   : The CNG algorithm provider handle and its hash object buffer are cached and
'           : reused across calls rather than opened per hash. Verify that reuse produces
'           : identical digests, that switching algorithms mid-stream re-opens the right
'           : provider, and that switching back is not contaminated by the detour.
'---------------------------------------------------------------------------------------
'
Public Sub TestHashProviderCacheStability()

    Dim strFirst As String
    ' Not named strRepeat: clsSqlFormatter has a private StrRepeat function, and VBA
    ' normalizes identifier casing project-wide, so the export kept rewriting the
    ' declaration to StrRepeat and breaking the str prefix convention.
    Dim strRepeatHash As String
    Dim strOther As String
    Dim strBack As String
    Dim lngIdx As Long

    strFirst = GetStringHash("provider cache probe")

    ' Repeated hashing must reuse the cached provider without drift
    For lngIdx = 1 To 50
        strRepeatHash = GetStringHash("provider cache probe")
        TestAssert strRepeatHash = strFirst, "stable across repeated calls"
    Next lngIdx

    ' Interleave a different algorithm, which forces the cached provider to be replaced
    strOther = HashWithAlgorithm("provider cache probe", "SHA1")
    TestAssert Len(strOther) > 0, "alternate algorithm returns a hash"
    TestAssert strOther <> strFirst, "alternate algorithm produces a different digest"

    ' Switching back must land on the original digest, not a stale buffer
    strBack = GetStringHash("provider cache probe")
    TestAssert strBack = strFirst, "digest unchanged after switching algorithms and back"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : HashWithAlgorithm
' Author    : Adam Waller
' Date      : 7/29/2026
' Purpose   : Hash a string under a specific algorithm by temporarily overriding the
'           : option, restoring it afterwards even if hashing fails.
'---------------------------------------------------------------------------------------
'
Private Function HashWithAlgorithm(ByVal strText As String, ByVal strAlgorithm As String) As String

    Dim strOriginal As String

    strOriginal = Options.HashAlgorithm
    Options.HashAlgorithm = strAlgorithm
    LogUnhandledErrors
    On Error Resume Next
    HashWithAlgorithm = GetStringHash(strText)
    On Error GoTo 0
    Options.HashAlgorithm = strOriginal

End Function


Public Sub TestUniqueHashSuffix()
    Dim strSuffix1 As String
    Dim strSuffix2 As String
    strSuffix1 = UniqueHashSuffix("same input")
    strSuffix2 = UniqueHashSuffix("same input")
    TestAssert Len(strSuffix1) = 7, "returns 7-character suffix"
    TestAssert strSuffix1 <> strSuffix2, "non-deterministic (same input = different output)"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestUtf8EncodingKnownDigests
' Author    : Adam Waller
' Date      : 7/30/2026
' Purpose   : GetUTF8Bytes reuses a cached pair of ADODB.Stream objects rather than
'           : constructing them per call. Anchor the encoding against externally computed
'           : SHA-256 digests of known UTF-8 byte sequences so a regression in the reuse
'           : logic cannot silently change what gets hashed. Covers multi-byte accented
'           : and CJK characters plus a surrogate pair, which must encode to four bytes
'           : rather than two three-byte lone surrogates.
'---------------------------------------------------------------------------------------
'
Public Sub TestUtf8EncodingKnownDigests()

    ' Skip unless the project is hashing with SHA256, since the expected values are fixed
    If StrComp(Nz2(Options.HashAlgorithm, DefaultHashAlgorithm), "SHA256", vbTextCompare) <> 0 Then
        TestAssert True, "SKIP: project is not configured for SHA256"
        Exit Sub
    End If

    ' "caf" & U+00E9  ->  63 61 66 C3 A9
    AssertKnownDigest GetStringHash("caf" & ChrW$(&HE9)), _
        "850f7dc43910ff890f8879c0ed26fe697c93a067ad93a7d50f466a7028a9bf4e", _
        "accented character encodes as two-byte UTF-8"

    ' U+4E2D U+6587  ->  E4 B8 AD E6 96 87
    AssertKnownDigest GetStringHash(ChrW$(&H4E2D) & ChrW$(&H6587)), _
        "72726d8818f693066ceb69afa364218b692e62ea92b385782363780f47529c21", _
        "CJK characters encode as three-byte UTF-8"

    ' U+1F600 as a surrogate pair  ->  F0 9F 98 80
    AssertKnownDigest GetStringHash(ChrW$(&HD83D) & ChrW$(&HDE00)), _
        "f0443a342c5ef54783a111b51ba56c938e474c32324d90c3a60c9c8e3a37e2d9", _
        "surrogate pair encodes as a single four-byte sequence"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : AssertKnownDigest
' Author    : Adam Waller
' Date      : 7/30/2026
' Purpose   : Compare a digest against a published full-length SHA-256 value, honoring
'           : Options.UseShortHash (GetHash truncates to 7 characters when it is set, so
'           : a test comparing against the full value fails on any project using short
'           : hashes). Reports both values, since a bare pass/fail on a digest comparison
'           : gives nothing to diagnose from.
'---------------------------------------------------------------------------------------
'
Private Sub AssertKnownDigest(ByVal strActual As String, ByVal strExpectedFull As String, _
    ByVal strContext As String)

    Dim strExpected As String

    strExpected = strExpectedFull
    If Options.UseShortHash Then strExpected = Left$(strExpectedFull, 7)

    TestAssert strActual = strExpected, _
        strContext & " (expected " & strExpected & ", got " & strActual & ")"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestUtf8StreamReuseLeavesNoResidue
' Author    : Adam Waller
' Date      : 7/30/2026
' Purpose   : The cached UTF-8 streams are rewound and truncated between calls. Without
'           : the SetEOS truncation a short input would inherit the tail of a longer
'           : preceding one, so hash a long string between two hashes of a short one and
'           : require the short digests to agree. Also covers the empty-string and
'           : BOM variants, whose lengths differ most from a preceding long input.
'---------------------------------------------------------------------------------------
'
Public Sub TestUtf8StreamReuseLeavesNoResidue()

    Dim strLong As String
    Dim strShortBefore As String
    Dim strShortAfter As String
    Dim strEmptyBefore As String
    Dim strEmptyAfter As String
    Dim strBomBefore As String
    Dim strBomAfter As String

    strLong = String$(4000, ChrW$(&H4E2D))

    ' Short input, then a much longer one, then the short input again
    strShortBefore = GetStringHash("x")
    GetStringHash strLong
    strShortAfter = GetStringHash("x")
    TestAssert strShortAfter = strShortBefore, "short digest unaffected by a preceding long input"

    ' The BOM variant reads from position zero, so residue would show up there too
    strBomBefore = GetStringHash("x", True)
    GetStringHash strLong, True
    strBomAfter = GetStringHash("x", True)
    TestAssert strBomAfter = strBomBefore, "BOM digest unaffected by a preceding long input"
    TestAssert strBomBefore <> strShortBefore, "BOM and non-BOM digests differ for the same text"

    ' Shortest payload the BOM path can produce (GetStringHash appends the trailing CRLF)
    strEmptyBefore = GetStringHash(vbNullString, True)
    GetStringHash strLong, True
    strEmptyAfter = GetStringHash(vbNullString, True)
    TestAssert strEmptyAfter = strEmptyBefore, "empty BOM digest unaffected by a preceding long input"

    ' Long inputs must still round-trip identically after the short ones above
    TestAssert GetStringHash(strLong) = GetStringHash(strLong), "long digest is repeatable"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestUtf8StreamsSurviveRelease
' Author    : Adam Waller
' Date      : 7/30/2026
' Purpose   : ReleaseUtf8Streams is called during teardown and from the GetUTF8Bytes
'           : error handler. Releasing mid-session must not break later hashing: the
'           : next call has to reopen the pair and produce the same digest.
'---------------------------------------------------------------------------------------
'
Public Sub TestUtf8StreamsSurviveRelease()

    Dim strBefore As String
    Dim strAfter As String

    strBefore = GetStringHash("release probe " & ChrW$(&H4E2D))
    ReleaseUtf8Streams
    strAfter = GetStringHash("release probe " & ChrW$(&H4E2D))
    TestAssert strAfter = strBefore, "digest unchanged after releasing the cached streams"

    ' A second release with nothing cached must be harmless
    ReleaseUtf8Streams
    ReleaseUtf8Streams
    TestAssert GetStringHash("release probe " & ChrW$(&H4E2D)) = strBefore, _
        "repeated release is safe"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : AssertNormalized
' Author    : Adam Waller
' Date      : 9/29/2026
' Purpose   : Check the output of NormalizeVbaCodeCasing, comparing in binary mode since
'           : the whole point is the letter case.
'---------------------------------------------------------------------------------------
'
Private Sub AssertNormalized(strInput As String, strExpected As String, strDescription As String)

    Dim blnMatch As Boolean

    blnMatch = (StrComp(NormalizeVbaCodeCasing(strInput), strExpected, vbBinaryCompare) = 0)
    TestAssert blnMatch, strDescription

End Sub


Public Sub TestNormalizeCasing_01_Basic()
    AssertNormalized "Const B = 1", "const b = 1", "code is lower case"
End Sub


Public Sub TestNormalizeCasing_02_StringUntouched()
    AssertNormalized "s = ""Hola Mundo""", "s = ""Hola Mundo""", "string literal keeps its case"
End Sub


Public Sub TestNormalizeCasing_03_EscapedQuote()
    AssertNormalized "s = ""A""""B"" & Foo", "s = ""A""""B"" & foo", "doubled quote does not close the string"
End Sub


Public Sub TestNormalizeCasing_04_ApostropheComment()
    AssertNormalized "x = 1 ' Nota Con Caja", "x = 1 ' Nota Con Caja", "apostrophe comment keeps its case"
End Sub


Public Sub TestNormalizeCasing_05_ApostropheInString()
    AssertNormalized "MsgBox ""It's"" & Nombre", "msgbox ""It's"" & nombre", "apostrophe inside a string is not a comment"
End Sub


Public Sub TestNormalizeCasing_06_QuoteInComment()
    AssertNormalized "s = ""a"" ' c ""B""", "s = ""a"" ' c ""B""", "quote inside a comment is not a string"
End Sub


Public Sub TestNormalizeCasing_07_RemAtLineStart()
    AssertNormalized "Rem Nota X", "Rem Nota X", "Rem at the start of a line is a comment"
End Sub


Public Sub TestNormalizeCasing_08_RemAfterColon()
    AssertNormalized "x = 1: Rem Nota X", "x = 1: Rem Nota X", "Rem after a colon is a comment"
End Sub


Public Sub TestNormalizeCasing_09_RemAfterLineNumber()
    AssertNormalized "10 Rem Nota X", "10 Rem Nota X", "Rem after a line number is a comment"
End Sub


Public Sub TestNormalizeCasing_10_RemAsIdentifierPrefix()
    AssertNormalized "Call RemoveItem(X)", "call removeitem(x)", "RemoveItem is not a comment"
End Sub


Public Sub TestNormalizeCasing_11_RemInsideStatement()
    AssertNormalized "x = Remanente", "x = remanente", "Rem not at the start of a statement is not a comment"
End Sub


Public Sub TestNormalizeCasing_12_CommentContinuation()
    AssertNormalized "' Sigue _" & vbCrLf & "Linea Dos", "' Sigue _" & vbCrLf & "Linea Dos", _
        "a comment ending in space-underscore continues on the next line"
End Sub


Public Sub TestNormalizeCasing_13_UnderscoreWithoutSpace()
    AssertNormalized "' Nota_" & vbCrLf & "Dim X", "' Nota_" & vbCrLf & "dim x", _
        "an underscore attached to the text does not continue the comment"
End Sub


Public Sub TestNormalizeCasing_14_CodeContinuation()
    AssertNormalized "x = Foo _" & vbCrLf & "  + Bar", "x = foo _" & vbCrLf & "  + bar", _
        "continued code is all lower case"
End Sub


Public Sub TestNormalizeCasing_15_UnclosedString()
    AssertNormalized "x = ""Sin Cerrar" & vbCrLf & "Dim Y", "x = ""Sin Cerrar" & vbCrLf & "dim y", _
        "an unclosed string ends with its line"
End Sub


Public Sub TestNormalizeCasing_16_HashSign()
    AssertNormalized "#If VBA7 Then" & vbCrLf & "Print #1, Dato", "#if vba7 then" & vbCrLf & "print #1, dato", _
        "the # sign delimits nothing"
End Sub


Public Sub TestNormalizeCasing_17_AccentedIdentifier()
    AssertNormalized "Dim Descripci" & ChrW$(243) & "n", "dim descripci" & ChrW$(243) & "n", _
        "identifier with an accent"
End Sub


Public Sub TestNormalizeCasing_18_Empty()
    AssertNormalized vbNullString, vbNullString, "empty text stays empty"
End Sub


Public Sub TestNormalizeCasing_19_MixedLines()
    AssertNormalized "Const B = 1" & vbCrLf & "' B" & vbCrLf & "s = ""B""", _
        "const b = 1" & vbCrLf & "' B" & vbCrLf & "s = ""B""", "code, comment and string on separate lines"
End Sub


Public Sub TestNormalizeCasing_20_UpperRemOnLaterLine()
    AssertNormalized "Dim A" & vbCrLf & "REM Nota X" & vbCrLf & "Call REMOVE(A)", _
        "dim a" & vbCrLf & "REM Nota X" & vbCrLf & "call remove(a)", _
        "REM in upper case on a later line is a comment, REMOVE is not"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestNormalizeCasing_21_LargeModule
' Author    : Adam Waller
' Date      : 9/29/2026
' Purpose   : A module of about 500 KB, with strings, comments and Rem on every few
'           : lines, must come out right and fast. A text search would cost the length
'           : of the whole module on every call, and this would take over a minute.
'---------------------------------------------------------------------------------------
'
Public Sub TestNormalizeCasing_21_LargeModule()

    Const clngBlocks As Long = 4000

    Dim strBlock As String
    Dim strExpected As String
    Dim strInput As String
    Dim strOutput As String
    Dim sngStart As Single
    Dim sngSeconds As Single

    strBlock = "    s = ""Texto Con Caja"" ' Comentario Con Caja" & vbCrLf & _
        "    Call RemoveItem(X): Rem Nota Con Caja" & vbCrLf & _
        "    Dim Remanente As Long" & vbCrLf
    strExpected = "    s = ""Texto Con Caja"" ' Comentario Con Caja" & vbCrLf & _
        "    call removeitem(x): Rem Nota Con Caja" & vbCrLf & _
        "    dim remanente as long" & vbCrLf
    strInput = Replace(Space$(clngBlocks), " ", strBlock)
    strExpected = Replace(Space$(clngBlocks), " ", strExpected)

    sngStart = Timer
    strOutput = NormalizeVbaCodeCasing(strInput)
    sngSeconds = Timer - sngStart

    TestAssert Len(strInput) > 450000, "the module is large"
    TestAssert StrComp(strOutput, strExpected, vbBinaryCompare) = 0, "large module normalized as expected"
    TestAssert sngSeconds < 2, "large module normalized in under 2 seconds (took " & sngSeconds & " s)"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestCodeTextHash_IgnoresCodeCasing
' Author    : Adam Waller
' Date      : 9/29/2026
' Purpose   : The code hash must not change when the VBE re-cases an identifier, but it
'           : must change when the case changes inside a string or a comment.
'---------------------------------------------------------------------------------------
'
Public Sub TestCodeTextHash_IgnoresCodeCasing()

    Dim blnSame As Boolean

    blnSame = (GetCodeTextHash("Const B = 1", vbNullString) = GetCodeTextHash("Const b = 1", vbNullString))
    TestAssert blnSame, "identifier case does not change the hash"

    blnSame = (GetCodeTextHash("x = ""B""", vbNullString) = GetCodeTextHash("x = ""b""", vbNullString))
    TestAssert Not blnSame, "case inside a string changes the hash"

    blnSame = (GetCodeTextHash("x = 1 ' B", vbNullString) = GetCodeTextHash("x = 1 ' b", vbNullString))
    TestAssert Not blnSame, "case inside a comment changes the hash"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestCodeTextHash_PrefixAndLegacy
' Author    : Adam Waller
' Date      : 9/29/2026
' Purpose   : The new hash carries the version prefix, empty code gives an empty hash,
'           : and the legacy hash is exactly the hash of the text and the flag, which
'           : the upgrade of existing index entries depends on.
'---------------------------------------------------------------------------------------
'
Public Sub TestCodeTextHash_PrefixAndLegacy()

    Dim strLegacy As String
    Dim blnMatch As Boolean

    TestAssert Left$(GetCodeTextHash("Const B = 1", vbNullString), Len(cstrCodeHashPrefix)) = cstrCodeHashPrefix, _
        "hash starts with the version prefix"
    TestAssert cstrCodeHashPrefix = "ci1:", "prefix is ci1:"
    TestAssert Len(GetCodeTextHash(vbNullString, vbNullString)) = 0, "empty code gives an empty hash"
    TestAssert GetCodeTextHash(vbNullString, vbNullString, True) = GetStringHash(vbNullString), "empty code keeps the old legacy hash"

    strLegacy = GetCodeTextHash("Const B = 1", "1", True)
    blnMatch = (strLegacy = GetStringHash("Const B = 1" & "1"))
    TestAssert blnMatch, "legacy hash is the hash of the text and the flag"
    TestAssert Left$(strLegacy, Len(cstrCodeHashPrefix)) <> cstrCodeHashPrefix, "legacy hash has no prefix"

    blnMatch = (GetCodeTextHash("Const B = 1", "1", True) = GetCodeTextHash("Const b = 1", "1", True))
    TestAssert Not blnMatch, "legacy hash is case sensitive"

End Sub
