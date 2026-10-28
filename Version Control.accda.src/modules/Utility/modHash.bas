Attribute VB_Name = "modHash"
'---------------------------------------------------------------------------------------
' Module    : modHash
' Author    : Adam Waller, Erik A, 2019; hecon5, 2021
' Date      : 12/4/2020, 4/9/2020; Revised and adapted Jan. 21, 2021
' Purpose   : Build hashes for content comparison.
'           :
'           : Adapted from: https://stackoverflow.com/questions/61929229/creating-secure-password-hash-in-php-but-checking-access-vba
'           :
'           : Removes dependancy on .NET 3.5 and others for hashing and securing data.
'           : This also has the ancilliary benefit of being able to use OS-level optimizations
'           : and hardware accelerators (if present).
'           :
'           : References: https://docs.microsoft.com/en-us/windows/win32/seccng/cng-algorithm-identifiers
'           : https://docs.microsoft.com/en-us/windows/win32/seccng/cng-portal
'           :
'           : See also: https://github.com/joyfullservice/msaccess-vcs-addin/wiki/Encryption
'---------------------------------------------------------------------------------------
Option Compare Database
Option Private Module
Option Explicit
'@Folder("Utility")


Private Declare PtrSafe Function BCryptOpenAlgorithmProvider Lib "BCrypt.dll" ( _
                            ByRef phAlgorithm As LongPtr, _
                            ByVal pszAlgId As LongPtr, _
                            ByVal pszImplementation As LongPtr, _
                            ByVal dwFlags As Long) As Long

Private Declare PtrSafe Function BCryptCloseAlgorithmProvider Lib "BCrypt.dll" ( _
                            ByVal hAlgorithm As LongPtr, _
                            ByVal dwFlags As Long) As Long

Private Declare PtrSafe Function BCryptCreateHash Lib "BCrypt.dll" ( _
                            ByVal hAlgorithm As LongPtr, _
                            ByRef phHash As LongPtr, pbHashObject As Any, _
                            ByVal cbHashObject As Long, _
                            ByVal pbSecret As LongPtr, _
                            ByVal cbSecret As Long, _
                            ByVal dwFlags As Long) As Long

Private Declare PtrSafe Function BCryptHashData Lib "BCrypt.dll" ( _
                            ByVal hHash As LongPtr, _
                            pbInput As Any, _
                            ByVal cbInput As Long, _
                            Optional ByVal dwFlags As Long = 0) As Long

Private Declare PtrSafe Function BCryptFinishHash Lib "BCrypt.dll" ( _
                            ByVal hHash As LongPtr, _
                            pbOutput As Any, _
                            ByVal cbOutput As Long, _
                            ByVal dwFlags As Long) As Long

Private Declare PtrSafe Function BCryptDestroyHash Lib "BCrypt.dll" (ByVal hHash As LongPtr) As Long

Private Declare PtrSafe Function BCryptGetProperty Lib "BCrypt.dll" ( _
                            ByVal hObject As LongPtr, _
                            ByVal pszProperty As LongPtr, _
                            ByRef pbOutput As Any, _
                            ByVal cbOutput As Long, _
                            ByRef pcbResult As Long, _
                            ByVal dfFlags As Long) As Long

Private Const ModuleName As String = "modHash"

' Cached CNG algorithm provider. Opening a provider and querying its two size
' properties costs more than hashing the small inputs this add-in mostly deals with
' (file property strings, combined hash strings, VBA module text), and an export or
' merge on a large project performs tens of thousands of hashes. The handle is opened
' on first use and reused until the algorithm changes or ReleaseHashProvider is called.
Private m_hAlg As LongPtr
Private m_strCachedAlg As String
Private m_lngObjectLength As Long
Private m_lngHashLength As Long
Private m_bteHashObject() As Byte

' Lookup table of two-character lowercase hex values for 0-255, built on first use.
' Formatting a 32-byte digest one byte at a time through Hex/Right/LCase is several
' string allocations per byte, which is significant at hashing volumes.
Private m_strHexBytes() As String
Private m_blnHexTableReady As Boolean

' Cached UTF-8 conversion streams. Constructing and opening the two ADODB.Stream
' objects dominates GetUTF8Bytes: the COM creation costs roughly ten times the actual
' encoding for the short inputs this add-in mostly hashes. Since every property hash,
' content hash, and metadata hash routes through here, the pair is opened once and
' rewound between calls instead. See ReleaseUtf8Streams for the teardown, and
' modTestHash for the tests pinning byte-for-byte equality with the fresh-stream form.
Private m_stmUtf8Text As ADODB.Stream
Private m_stmUtf8Binary As ADODB.Stream


Private Function NGHash(pData As LongPtr, lenData As Long, Optional HashingAlgorithm As String = DefaultHashAlgorithm) As Byte()

    'Erik A, 2019, adapted by Adam Waller
    'Hash data by using the Next Generation Cryptography API
    'Loosely based on https://docs.microsoft.com/en-us/windows/desktop/SecCNG/creating-a-hash-with-cng
    'Allowed algorithms:  https://docs.microsoft.com/en-us/windows/desktop/SecCNG/cng-algorithm-identifiers. Note: only hash algorithms, check OS support
    'Error messages not implemented
    '
    LogUnhandledErrors
    On Error GoTo VBErrHandler
    Dim errorMessage As String
    Dim hHash As LongPtr
    Dim bHash() As Byte

    ' Open (or reuse) the algorithm provider and its cached size properties
    If Not EnsureHashProvider(HashingAlgorithm) Then GoTo ErrHandler

    ' Allocate the digest buffer
    ReDim bHash(0 To m_lngHashLength - 1)

    'Create hash object
    If BCryptCreateHash(m_hAlg, hHash, m_bteHashObject(0), m_lngObjectLength, 0, 0, 0) <> 0 Then GoTo ErrHandler

    'Hash data
    If BCryptHashData(hHash, ByVal pData, lenData) <> 0 Then GoTo ErrHandler
    If BCryptFinishHash(hHash, bHash(0), m_lngHashLength, 0) <> 0 Then GoTo ErrHandler

    'Return result
    NGHash = bHash

ExitHandler:
    'Cleanup. (The algorithm provider is cached; only the per-call hash object is freed.)
    If hHash <> 0 Then BCryptDestroyHash hHash
    Exit Function

VBErrHandler:
    errorMessage = "VB Error " & Err.Number & ": " & Err.Description

ErrHandler:
    ' Free the hash object before discarding the provider that owns it, then drop the
    ' cached provider so a transient failure cannot poison every later call.
    If hHash <> 0 Then
        BCryptDestroyHash hHash
        hHash = 0
    End If
    ReleaseHashProvider
    CatchAny eelCritical, "Error hashing! " & errorMessage & ". Algorithm: " & HashingAlgorithm, ModuleName & ".NGHash", True, True
    GoTo ExitHandler

End Function


'---------------------------------------------------------------------------------------
' Procedure : EnsureHashProvider
' Author    : Adam Waller
' Date      : 7/29/2026
' Purpose   : Open the CNG algorithm provider for the requested algorithm and cache it
'           : along with its hash object length and digest length. Returns True when a
'           : usable provider is available. Re-opens when the algorithm changes.
'---------------------------------------------------------------------------------------
'
Private Function EnsureHashProvider(strAlgorithm As String) As Boolean

    Dim strAlgId As String
    Dim strProperty As String

    ' Reuse the cached provider when the algorithm has not changed
    If m_hAlg <> 0 And StrComp(m_strCachedAlg, strAlgorithm, vbBinaryCompare) = 0 Then
        EnsureHashProvider = True
        Exit Function
    End If

    ' Different algorithm requested (or first call). Discard any existing provider.
    ReleaseHashProvider

    strAlgId = strAlgorithm & vbNullChar
    If BCryptOpenAlgorithmProvider(m_hAlg, StrPtr(strAlgId), 0, 0) Then
        m_hAlg = 0
        Exit Function
    End If

    ' Hash object size (buffer BCryptCreateHash writes its state into)
    strProperty = "ObjectLength" & vbNullChar
    If BCryptGetProperty(m_hAlg, StrPtr(strProperty), m_lngObjectLength, LenB(m_lngObjectLength), 0, 0) <> 0 Then
        ReleaseHashProvider
        Exit Function
    End If
    ReDim m_bteHashObject(0 To m_lngObjectLength - 1)

    ' Digest size
    strProperty = "HashDigestLength" & vbNullChar
    If BCryptGetProperty(m_hAlg, StrPtr(strProperty), m_lngHashLength, LenB(m_lngHashLength), 0, 0) <> 0 Then
        ReleaseHashProvider
        Exit Function
    End If

    m_strCachedAlg = strAlgorithm
    EnsureHashProvider = True

End Function


'---------------------------------------------------------------------------------------
' Procedure : ReleaseHashProvider
' Author    : Adam Waller
' Date      : 7/29/2026
' Purpose   : Close the cached CNG algorithm provider. Called on a clean exit (see
'           : modObjects.ReleaseObjects) and from the hashing error handler so a failure
'           : does not leave a bad handle cached. Safe to call when nothing is cached.
'---------------------------------------------------------------------------------------
'
Public Sub ReleaseHashProvider()

    If m_hAlg <> 0 Then BCryptCloseAlgorithmProvider m_hAlg, 0
    m_hAlg = 0
    m_strCachedAlg = vbNullString
    m_lngObjectLength = 0
    m_lngHashLength = 0
    Erase m_bteHashObject

End Sub


'---------------------------------------------------------------------------------------
' Procedure : HashBytes
' Author    : Adam Waller
' Date      : 1/21/2021
' Purpose   : Wrappers for NGHash functions
'---------------------------------------------------------------------------------------
'
Private Function HashBytes(Data() As Byte, Optional HashingAlgorithm As String = DefaultHashAlgorithm) As Byte()
    LogUnhandledErrors
    On Error Resume Next
    HashBytes = NGHash(VarPtr(Data(LBound(Data))), UBound(Data) - LBound(Data) + 1, HashingAlgorithm)
    If Catch(9) Then HashBytes = NGHash(VarPtr(Null), UBound(Data) - LBound(Data) + 1, HashingAlgorithm)
    CatchAny eelCritical, "Error hashing data!", ModuleName & ".HashBytes", True, True
End Function

Private Function HashString(str As String, Optional HashingAlgorithm As String = DefaultHashAlgorithm) As Byte()
    LogUnhandledErrors
    On Error Resume Next
    HashString = NGHash(StrPtr(str), Len(str) * 2, HashingAlgorithm)
    If Catch(9) Then HashString = NGHash(StrPtr(vbNullString), Len(str) * 2, HashingAlgorithm)
    CatchAny eelCritical, "Error hashing string!", ModuleName & ".HashString", True, True
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetStringHash
' Author    : Adam Waller
' Date      : 11/30/2020
' Purpose   : Convert string to byte array, and return a hash. Optionally include the
'           : UTF-8 BOM. (Useful when comparing to a file hash)
'---------------------------------------------------------------------------------------
'
Public Function GetStringHash(ByVal strText As String, Optional blnWithBom As Boolean = False) As String
    If blnWithBom Then
        ' Ensure that we are ending the content with a vbCrLf
        ' (To match the behavior of the WriteFile function)
        If Right(strText, 2) <> vbCrLf Then strText = strText & vbCrLf
    End If
    GetStringHash = GetHash(GetUTF8Bytes(strText, blnWithBom))
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetFileHash
' Author    : Adam Waller
' Date      : 11/30/2020
' Purpose   : Return a hash from a file
'---------------------------------------------------------------------------------------
'
Public Function GetFileHash(strPath As String) As String
    GetFileHash = GetHash(GetFileBytes(strPath))
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetBytesHash
' Author    : Adam Waller
' Date      : 11/1/2021
' Purpose   : Return hash from byte array
'---------------------------------------------------------------------------------------
'
Public Function GetBytesHash(bteData() As Byte) As String
    GetBytesHash = GetHash(bteData())
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetDictionaryHash
' Author    : Adam Waller
' Date      : 12/1/2020
' Purpose   : Wrapper to get a hash from a dictionary object (converted to json)
'---------------------------------------------------------------------------------------
'
Public Function GetDictionaryHash(dSource As Dictionary) As String
    GetDictionaryHash = GetStringHash(ConvertToJson(dSource))
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetHash
' Author    : Adam Waller
' Date      : 11/30/2020
' Purpose   : Create a hash from the byte array
'---------------------------------------------------------------------------------------
'
Private Function GetHash(bteContent() As Byte) As String

    ' Variant rather than Byte(): HashBytes returns an unallocated array for empty input
    ' (GetStringHash("") is reachable, e.g. a component with no source files on disk).
    ' A Variant absorbs that as Empty, where LenB reports 0 and the loop below is skipped,
    ' preserving the empty-string result callers and stored hashes already depend on.
    Dim bteHash As Variant
    Dim strHash As String
    Dim lngPos As Long
    Dim lngBytes As Long
    Dim intLength As Integer
    Dim strAlgorithm As String

    ' Get hashing options
    strAlgorithm = Nz2(Options.HashAlgorithm, DefaultHashAlgorithm)
    If Options.UseShortHash Then intLength = 7

    ' Start performance timer and compute the hash
    Perf.OperationStart "Compute " & strAlgorithm
    bteHash = HashBytes(bteContent, strAlgorithm)

    ' Convert the digest to a hexadecimal string. Writing into a preallocated buffer
    ' with a byte -> hex lookup table avoids the string allocations that Hex/Right/LCase
    ' cost per byte, which add up over the tens of thousands of hashes an export or
    ' merge performs. (NGHash returns a zero-based array; LenB gives its length.)
    EnsureHexTable
    lngBytes = LenB(bteHash)
    strHash = Space$(lngBytes * 2)
    For lngPos = 0 To lngBytes - 1
        Mid$(strHash, (lngPos * 2) + 1, 2) = m_strHexBytes(bteHash(lngPos))
    Next lngPos

    ' Return hash, truncating if needed.
    If intLength > 0 And intLength < Len(strHash) Then
        GetHash = Left$(strHash, intLength)
    Else
        GetHash = strHash
    End If
    Perf.OperationEnd

End Function


'---------------------------------------------------------------------------------------
' Procedure : EnsureHexTable
' Author    : Adam Waller
' Date      : 7/29/2026
' Purpose   : Build the byte -> two-character lowercase hex lookup table on first use.
'---------------------------------------------------------------------------------------
'
Private Sub EnsureHexTable()

    Dim lngByte As Long

    If m_blnHexTableReady Then Exit Sub

    ReDim m_strHexBytes(0 To 255)
    For lngByte = 0 To 255
        m_strHexBytes(lngByte) = LCase$(Right$("0" & Hex$(lngByte), 2))
    Next lngByte
    m_blnHexTableReady = True

End Sub


'---------------------------------------------------------------------------------------
' Procedure : GetCodeModuleHash
' Author    : Adam Waller
' Date      : 11/30/2020
' Purpose   : Return a hash from the VBA code module behind an object.
'---------------------------------------------------------------------------------------
'
Public Function GetCodeModuleHash(intType As eDatabaseComponentType, strName As String) As String

    Dim strHash As String
    Dim cmpItem As VBComponent
    Dim strPrefix As String
    Dim proj As VBProject
    Dim blnNoCode As Boolean
    Dim strInstancingFlag As String

    Perf.OperationStart "Get VBA Hash"
    Select Case intType
        Case edbForm:   strPrefix = "Form_"
        Case edbReport: strPrefix = "Report_"
        Case edbModule, edbVbeForm
        Case Else
            ' No code module
            blnNoCode = True
    End Select

    ' Get the hash from the VBA code module content.
    If Not blnNoCode Then

        ' Get a reference for the VBProject in the current (not code) database.
        Set proj = CurrentVBProject

        ' Attempt to locate the object in the VBComponents collection
        LogUnhandledErrors
        On Error Resume Next
        Set cmpItem = proj.VBComponents(strPrefix & strName)
        Catch 9 ' Component not found. (Could be an object with no code module)
        CatchAny eelError, "Error accessing VBComponent for '" & strPrefix & strName & "'", ModuleName & ".GetCodeModuleHash"
        If DebugMode(True) Then On Error GoTo 0 Else On Error Resume Next

        ' Output the hash
        If Not cmpItem Is Nothing Then
            With cmpItem
                ' Check for class module
                If .Type = vbext_ct_ClassModule Then
                    ' Save instancing property as a flag to include with hash
                    strInstancingFlag = CStr(.Properties("Instancing"))
                End If
                ' Generate hash from code and instancing flag (if applicable)
                strHash = GetStringHash(.CodeModule.Lines(1, 999999) & strInstancingFlag)
            End With
        End If

    End If

    ' Return hash (if any)
    GetCodeModuleHash = strHash
    Perf.OperationEnd

End Function


'---------------------------------------------------------------------------------------
' Procedure : GetUTF8Bytes
' Author    : Adam Waller
' Date      : 11/2/2021
' Purpose   : Return a UTF-8 (wide) byte array from a string. Optionally include the
'           : UTF-8 BOM. (Useful when comparing to a file hash)
'---------------------------------------------------------------------------------------
'
Private Function GetUTF8Bytes(strText As String, Optional blnWithBom As Boolean = False) As Byte()

    ' Check for empty string
    If (Len(strText) = 0) And Not blnWithBom Then
        GetUTF8Bytes = vbNullString
        Exit Function
    End If

    LogUnhandledErrors
    On Error GoTo ErrHandler

    If Not EnsureUtf8Streams Then Exit Function

    ' Rewind both streams and truncate whatever the previous call left behind. SetEOS
    ' at position zero is what makes reuse safe: without it a short input would inherit
    ' the tail of a longer one.
    With m_stmUtf8Text
        .Position = 0
        .SetEOS
        .WriteText strText
        .Position = 0
    End With
    With m_stmUtf8Binary
        .Position = 0
        .SetEOS
    End With

    m_stmUtf8Text.CopyTo m_stmUtf8Binary, adReadAll

    ' Include the BOM, or step over it
    If blnWithBom Then
        m_stmUtf8Binary.Position = 0
    Else
        m_stmUtf8Binary.Position = 3
    End If
    GetUTF8Bytes = m_stmUtf8Binary.Read(adReadAll)

    Exit Function

ErrHandler:
    ' Drop the cached pair so a transient stream failure cannot poison every later
    ' call, then let the caller's handler see the original error.
    ReleaseUtf8Streams
    CatchAny eelCritical, "Error converting text to UTF-8 bytes", _
        ModuleName & ".GetUTF8Bytes", True, True

End Function


'---------------------------------------------------------------------------------------
' Procedure : EnsureUtf8Streams
' Author    : Adam Waller
' Date      : 7/30/2026
' Purpose   : Open the cached UTF-8 text and binary streams if they are not already
'           : available. Returns True when both are usable.
'---------------------------------------------------------------------------------------
'
Private Function EnsureUtf8Streams() As Boolean

    If m_stmUtf8Text Is Nothing Then
        Set m_stmUtf8Text = New ADODB.Stream
        m_stmUtf8Text.Open
        m_stmUtf8Text.Charset = "utf-8"
        m_stmUtf8Text.Type = adTypeText
    End If

    If m_stmUtf8Binary Is Nothing Then
        Set m_stmUtf8Binary = New ADODB.Stream
        m_stmUtf8Binary.Open
        m_stmUtf8Binary.Charset = "utf-8"
        m_stmUtf8Binary.Type = adTypeBinary
    End If

    EnsureUtf8Streams = Not (m_stmUtf8Text Is Nothing Or m_stmUtf8Binary Is Nothing)

End Function


'---------------------------------------------------------------------------------------
' Procedure : ReleaseUtf8Streams
' Author    : Adam Waller
' Date      : 7/30/2026
' Purpose   : Close the cached UTF-8 conversion streams. Called on a clean exit (see
'           : modObjects.ReleaseObjects) and from the GetUTF8Bytes error handler so a
'           : failure does not leave an unusable stream cached. Safe to call when
'           : nothing is cached.
'---------------------------------------------------------------------------------------
'
Public Sub ReleaseUtf8Streams()

    ' Closing a stream that is already closed raises, and this runs on teardown and from
    ' an error handler, so failures here are ignored by design.
    LogUnhandledErrors
    On Error Resume Next
    If Not m_stmUtf8Text Is Nothing Then m_stmUtf8Text.Close
    If Not m_stmUtf8Binary Is Nothing Then m_stmUtf8Binary.Close
    Set m_stmUtf8Text = Nothing
    Set m_stmUtf8Binary = Nothing

End Sub


'---------------------------------------------------------------------------------------
' Procedure : UniqueHashSuffix
' Author    : Adam Waller
' Date      : 4/24/2026
' Purpose   : 7-character hex suffix for an arbitrary string. Used to make
'           : sandbox object names unique across concurrent / repeated runs.
'           : Uniqueness comes from Perf.MicroTimer plus a per-call static
'           : counter (so two calls landing in the same timer tick still differ),
'           : not from the hash width; 7 chars matches the short-hash convention
'           : used elsewhere. The counter wraps just below Long.MaxValue to avoid
'           : overflow; its period far exceeds the calls possible per tick.
'---------------------------------------------------------------------------------------
'
Public Function UniqueHashSuffix(ByVal s As String) As String
    Static lngCounter As Long
    Dim strFull As String
    lngCounter = (lngCounter Mod 2147483646) + 1
    strFull = GetStringHash(s & ":" & CStr(Perf.MicroTimer) & ":" & CStr(lngCounter))
    UniqueHashSuffix = LCase$(Left$(strFull, 7))
End Function
