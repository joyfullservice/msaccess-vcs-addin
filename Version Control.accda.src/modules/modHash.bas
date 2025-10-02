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

Private Declare PtrSafe Function CoCreateGuid Lib "ole32" (ID As Any) As Long
'Private Declare PtrSafe Function StringFromGUID2 Lib "ole32" (rguid As Any _
                                                            , ByVal lpsz As LongPtr _
                                                            , ByVal cchMax As Long _
                                                            ) As Long

Public Const SHA256_HASH_LENGTH As Long = 64

Private Const ModuleName As String = "modHash"
Private Const DEFAULT_SHORT_HASH_LENGTH As Long = 7
Private Const GUID_BRACE As String = "}"


Private Function NGHash(pData As LongPtr, lenData As Long, Optional HashingAlgorithm As String = DefaultHashAlgorithm) As Byte()

    'Erik A, 2019, adapted by Adam Waller
    'Hash data by using the Next Generation Cryptography API
    'Loosely based on https://docs.microsoft.com/en-us/windows/desktop/SecCNG/creating-a-hash-with-cng
    'Allowed algorithms:  https://docs.microsoft.com/en-us/windows/desktop/SecCNG/cng-algorithm-identifiers. Note: only hash algorithms, check OS support
    'Error messages not implemented
    '
    On Error GoTo VBErrHandler
    Dim errorMessage As String

    Dim hAlg As LongPtr
    Dim algId As String

    'Open crypto provider
    algId = HashingAlgorithm & vbNullChar
    If BCryptOpenAlgorithmProvider(hAlg, StrPtr(algId), 0, 0) Then GoTo ErrHandler

    'Determine hash object size, allocate memory
    Dim bHashObject() As Byte
    Dim cmd As String
    cmd = "ObjectLength" & vbNullString
    Dim Length As Long
    If BCryptGetProperty(hAlg, StrPtr(cmd), Length, LenB(Length), 0, 0) <> 0 Then GoTo ErrHandler
    ReDim bHashObject(0 To Length - 1)

    'Determine digest size, allocate memory
    Dim HashLength As Long
    cmd = "HashDigestLength" & vbNullChar
    If BCryptGetProperty(hAlg, StrPtr(cmd), HashLength, LenB(HashLength), 0, 0) <> 0 Then GoTo ErrHandler
    Dim bHash() As Byte
    ReDim bHash(0 To HashLength - 1)

    'Create hash object
    Dim hHash As LongPtr
    If BCryptCreateHash(hAlg, hHash, bHashObject(0), Length, 0, 0, 0) <> 0 Then GoTo ErrHandler

    'Hash data
    If BCryptHashData(hHash, ByVal pData, lenData) <> 0 Then GoTo ErrHandler
    If BCryptFinishHash(hHash, bHash(0), HashLength, 0) <> 0 Then GoTo ErrHandler

    'Return result
    NGHash = bHash

ExitHandler:
    'Cleanup
    If hAlg <> 0 Then BCryptCloseAlgorithmProvider hAlg, 0
    If hHash <> 0 Then BCryptDestroyHash hHash
    Exit Function

VBErrHandler:
    errorMessage = "VB Error " & Err.Number & ": " & Err.Description

ErrHandler:
    CatchAny eelCritical, "Error hashing! " & errorMessage & ". Algorithm: " & HashingAlgorithm, ModuleName & ".NGHash", True, True
    GoTo ExitHandler

End Function


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
Public Function GetStringHash(ByVal strText As String _
                            , Optional blnWithBom As Boolean = False _
                            , Optional ByVal intHashLen As Integer = 0) As String
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
Public Function GetFileHash(strPath As String _
                            , Optional ByVal intHashLen = 0) As String

    GetFileHash = GetHash(GetFileBytes(strPath), intHashLen)
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetBytesHash
' Author    : Adam Waller
' Date      : 11/1/2021
' Purpose   : Return hash from byte array
'---------------------------------------------------------------------------------------
'
Public Function GetBytesHash(bteData() As Byte _
                            , Optional ByVal intHashLen = 0) As String

    GetBytesHash = GetHash(bteData(), intHashLen)
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetDictionaryHash
' Author    : Adam Waller
' Date      : 12/1/2020
' Purpose   : Wrapper to get a hash from a dictionary object (converted to json)
'---------------------------------------------------------------------------------------
'
Public Function GetDictionaryHash(dSource As Dictionary _
                                , Optional ByVal intHashLen = 0) As String

    GetDictionaryHash = GetStringHash(ConvertToJson(dSource), , intHashLen)
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetHash
' Author    : Adam Waller
' Date      : 11/30/2020
' Purpose   : Create a hash from the byte array
'---------------------------------------------------------------------------------------
'
Public Function GetHash(bteContent() As Byte _
                        , Optional ByVal intHashLen As Integer = 0) As String

    Dim bteHash As Variant
    Dim strHash As String
    Dim intPos As Integer
    Dim intLength As Integer
    Dim strAlgorithm As String

    ' Get hashing options
    strAlgorithm = Nz2(Options.HashAlgorithm, DefaultHashAlgorithm)
    If intHashLen <> 0 Then
        intLength = intHashLen

    ElseIf Options.UseShortHash Then
        intLength = DEFAULT_SHORT_HASH_LENGTH
    End If

    ' Start performance timer and compute the hash
    Perf.OperationStart "Compute " & strAlgorithm
    bteHash = HashBytes(bteContent, strAlgorithm)

    ' Create string buffer to avoid concatenation
    strHash = Space(LenB(bteHash) * 2)

    ' Convert full hash to hexidecimal string
    For intPos = 1 To LenB(bteHash)
        Mid$(strHash, (intPos * 2) - 1, 2) = LCase(Right("0" & Hex(AscB(MidB(bteHash, intPos, 1))), 2))
    Next

    ' Return hash, truncating if needed.
    If intLength > 0 And intLength < Len(strHash) Then
        GetHash = Left$(strHash, intLength)
    Else
        GetHash = strHash
    End If
    Perf.OperationEnd

End Function


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

    Dim stmBinary As ADODB.Stream

    ' Check for empty string
    If (Len(strText) = 0) And Not blnWithBom Then
        GetUTF8Bytes = vbNullString
        Exit Function
    End If

    ' Set up binary stream
    Set stmBinary = New ADODB.Stream
    stmBinary.Open
    stmBinary.Charset = "utf-8"
    stmBinary.Type = adTypeBinary

    ' Load text into text stream
    With New ADODB.Stream
        .Open
        .Charset = "utf-8"
        .Type = adTypeText
        .WriteText strText
        .Position = 0
        ' Copy to binary stream
        .CopyTo stmBinary, adReadAll
        If blnWithBom Then
            ' Include BOM
            stmBinary.Position = 0
        Else
            ' Move past BOM
            stmBinary.Position = 3
        End If
        ' Return binary stream
        GetUTF8Bytes = stmBinary.Read(adReadAll)
    End With

End Function


'---------------------------------------------------------------------------------------
' Procedure : SimpleHash
' Author    : Adam Waller
' Date      : 7/24/2023
' Purpose   : Return a simple SHA256 hash from a file without any Windows API calls.
'           : (This function can be ported to VBScript as a worker process)
'           : Adapted from https://en.wikibooks.org/wiki/Visual_Basic_for_Applications/String_Hashing_in_VBA
'---------------------------------------------------------------------------------------
'
Public Function GetSimpleHash(strText As String) As String

    Dim objEnc As Object
    Dim objSHA256 As Object
    Dim objDoc As Object
    Dim bteText() As Byte
    Dim bteHash() As Byte
    Dim strHash As String

    On Error Resume Next

    ' Create objects
    Set objEnc = CreateObject("System.Text.UTF8Encoding")
    Set objSHA256 = CreateObject("System.Security.Cryptography.SHA256Managed")

    ' Compute hash
    bteText = objEnc.GetBytes_4(strText)
    bteHash = objSHA256.ComputeHash_2((bteText))

    ' Convert to hex string
    Set objDoc = CreateObject("MSXML2.DOMDocument")
    objDoc.LoadXML "<root />"
    With objDoc.DocumentElement
        .DataType = "bin.Hex"
        .nodeTypedValue = bteHash
        strHash = Replace(.Text, vbLf, vbNullString)
    End With

    ' Return short hash
    GetSimpleHash = Left(strHash, 7)

    ' Clear any errors
    If Err Then Err.Clear

End Function


Public Function GetHashValue(ByRef InputValue As Variant _
                            , Optional ByVal blnWithBom As Boolean = False _
                            , Optional ByVal intHashLen As Integer = DEFAULT_SHORT_HASH_LENGTH) As String

    Const FunctionName As String = ModuleName & ".GetHashValue"

    LogUnhandledErrors FunctionName
    On Error Resume Next

    If IsNull(InputValue) Then
        GetHashValue = GetStringHash(vbNullString, blnWithBom, intHashLen)

    ElseIf TypeName(InputValue) = "Byte()" Then
        Dim tByte() As Byte
        tByte = InputValue

        GetHashValue = GetHash(tByte, intHashLen)

    ElseIf TypeOf InputValue Is Dictionary Or TypeOf InputValue Is Collection Then
        Dim dInDict As Dictionary
        Set dInDict = InputValue
        GetHashValue = GetDictionaryHash(dInDict, intHashLen)

    Else
        GetHashValue = GetStringHash(CStr(InputValue), blnWithBom, intHashLen)

    End If

    CatchAny eelError, "Error detecting input type.", FunctionName, True, True

End Function


Public Function HexStringToByteArray(hexString As String) As Byte()

    Const FunctionName As String = ModuleName & ".HexStringToByteArray"

    Dim i As Long
    Dim tByteArr() As Byte

    Perf.OperationStart FunctionName

    ReDim tByteArr(1 To Len(hexString) / 2) As Byte

    For i = 2 To Len(hexString) Step 2
        tByteArr(i / 2) = "&H" & Mid(hexString, i - 1, 2)
    Next

    HexStringToByteArray = tByteArr

    Perf.OperationEnd

End Function


Public Function ByteArraytoHexString(ByteArrIn As Variant) As String

    Const FunctionName As String = ModuleName & ".ByteArraytoHexString"
    Dim tStrArr As String
    Dim tArrLen As Long
    Dim tPosition As Long
    Dim tByte As Variant

    ' Create string buffer to avoid concatenation
    tArrLen = (UBound(ByteArrIn) - LBound(ByteArrIn) + 1) * 2
    tStrArr = Space(tArrLen)
    tByte = ByteArrIn
    ' Convert full hash to hexidecimal string
    For tPosition = 1 To tArrLen / 2
        Mid$(tStrArr, (tPosition * 2) - 1, 2) = LCase(Right("0" & Hex(AscB(MidB(tByte, tPosition, 1))), 2))
    Next

    ByteArraytoHexString = tStrArr

End Function


Public Function GetGUID() As Variant
    Dim ID(0 To 15) As Byte
    CoCreateGuid ID(0)
    GetGUID = ID
End Function


Public Function IsGUID(ByRef GUIDIn As Variant) As Boolean

    Const FunctionName As String = ModuleName & ".IsGUID"

    Dim f_GUIDSubstring As String
    Dim f_BracePosition As String
    Dim f_HashIn As String
    Dim f_HashReturn As String
    Dim f_HalfWayStr As String
    Dim f_GUIDIn As Variant
    Dim f_tempByte() As Byte

    LogUnhandledErrors FunctionName
    On Error GoTo Exit_Error

    If IsEmpty(GUIDIn) Then Exit Function
    If IsNull(GUIDIn) Then Exit Function

    If TypeName(GUIDIn) = "Byte()" Then
        If Not (UBound(GUIDIn) - LBound(GUIDIn) + 1) = 16 Then Exit Function
        f_GUIDIn = GUIDIn

    ElseIf TypeName(GUIDIn) = "String" Then
        If Nz2(GUIDIn, vbNullString) = vbNullString Then Exit Function
        ' Detect if this is actually a GUID Byte() arr but parsed into a string:
        If Len(GUIDIn) = 8 Then ' This is likely a GUID
            f_tempByte = GUIDIn
            f_GUIDIn = f_tempByte
            GoTo TestGUID
        End If
        f_GUIDIn = getGUIDFromString(CStr(GUIDIn))
    End If

TestGUID:
    f_HashIn = GetHashValue(f_GUIDIn, , SHA256_HASH_LENGTH)
    f_HalfWayStr = Mid$(StringFromGUID(f_GUIDIn), 7, 38)
    f_HashReturn = GetHashValue(getGUIDFromString(f_HalfWayStr), , SHA256_HASH_LENGTH)
    If f_HashIn = f_HashReturn Then IsGUID = True

Exit_Here:
    Exit Function

Exit_Error:
    Catch 5, 13
    Resume Exit_Here

End Function


Public Function GetStringFromGUID(Optional ByRef GUIDIn As Variant = vbNullString) As String

    Const FunctionName As String = ModuleName & ".GetStringFromGUID"

    Static fNullHash As String

    Dim tGuIn() As Byte
    Dim tHashGUID As String

    LogUnhandledErrors FunctionName
    On Error Resume Next

    ' Check if the GUID in was a GUID or not.
    If Not IsGUID(GUIDIn) Then
        tGuIn = GetGUID
    Else
        tGuIn = GUIDIn
    End If

    'GetStringFromGUID = StringFromGUID(tGuIn) ' This adds a "{GUID { }" to the GUID.

    ' If you don't want the GUID tag, use this instead:
    GetStringFromGUID = Mid$(StringFromGUID(tGuIn), 7, 38)

End Function


Public Function getGUIDFromString(ByRef GUIDStringIn As String) As Variant

    Const FunctionName As String = ModuleName & ".getGUIDfromstring"

    Dim tBracePos As Long
    Dim tByteTemp() As Byte

    LogUnhandledErrors FunctionName
    On Error Resume Next

    tBracePos = InStrRev(GUIDStringIn, GUID_BRACE)
    If Len(GUIDStringIn) = 8 Then ' Somehow, a GUID got converted to a string so we need to just spit it back out again.
        tByteTemp = GUIDStringIn
        getGUIDFromString = tByteTemp

    ElseIf tBracePos > 0 Then
        getGUIDFromString = GUIDFromString(Mid$(GUIDStringIn, tBracePos - 38 + 1, 38))

    Else
        ' Brace not found, need to include braces.
        getGUIDFromString = GUIDFromString("{" & GUIDStringIn & "}")
    End If

    If Catch(5) Then getGUIDFromString = 0


End Function


Public Sub TestCreateGuid(Optional ByVal Iterations As Long = 10000)

    Dim GuidDict As Object 'Scripting.Dictionary
    Set GuidDict = CreateObject("Scripting.Dictionary")

    Dim i As Long

    Perf.Reset
    Perf.CategoryStart "test TestCreateGuid"

    For i = 1 To Iterations
        Dim GUID As String
        Perf.OperationStart "Add GUID to Dict"
        GUID = GetStringFromGUID

        If GuidDict.Exists(GUID) Then
            Debug.Print "Duplicate GUID created: " & GUID
            Stop
        Else
            GuidDict.Add GUID, vbNullString
        End If
        Perf.OperationEnd
    Next i

    Debug.Print GuidDict.Count; " unique GUIDs generated out of "; Iterations; " attempts"

    Perf.CategoryEnd
    Debug.Print Perf.GetReports

End Sub


Public Sub TestHashDict(Optional ByVal DictEntries As Long = 20)

    Dim GuidDict As New Scripting.Dictionary
    Dim GUID As String
    Dim i As Long

    Perf.Reset
    Perf.CategoryStart "test TestHashDict"

    For i = 1 To DictEntries
        Perf.OperationStart "Add GUID to Dict"
        GUID = GetStringFromGUID

        If Not GuidDict.Exists(GUID) Then
            GuidDict.Add GUID, vbNullString
        End If
        Perf.OperationEnd
    Next i

    Debug.Print GetHashValue(GuidDict)

    Perf.CategoryEnd
    Debug.Print Perf.GetReports

End Sub
