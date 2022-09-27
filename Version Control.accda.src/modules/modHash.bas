﻿Attribute VB_Name = "modHash"
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
'           : See also: https://github.com/joyfullservice/msaccess-vcs-integration/wiki/Encryption
'---------------------------------------------------------------------------------------
Option Compare Database
Option Private Module
Option Explicit

Private Const DefaultShortHashLen As Integer = 7

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

Private m_UseShortHash As Boolean

'---------------------------------------------------------------------------------------
' Procedure : UseShortHash
' Author    : hecon5
' Date      : 2022 SEPT 27
' Purpose   : Property for Hashing; allows for moving to class or easier management outside MSAccessVCS (if used).
'---------------------------------------------------------------------------------------
'
Public Property Get UseShortHash() As Boolean
'    UseShortHash = m_UseShortHash ' Comment out if using within MSAccessVCS (property is managed in the Options class).
    UseShortHash = Options.UseShortHash ' If using outside MSAccessVCS, then comment this out.
End Property
Public Property Let UseShortHash(NewVal As Boolean)
'    m_UseShortHash = NewVal ' Comment out if using within MSAccessVCS (property is managed in the Options class).
    Options.UseShortHash = NewVal ' If using outside MSAccessVCS, then comment this out.
End Property


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
    If DebugMode(True) Then On Error Resume Next Else On Error Resume Next
    HashBytes = NGHash(VarPtr(Data(LBound(Data))), UBound(Data) - LBound(Data) + 1, HashingAlgorithm)
    If Catch(9) Then HashBytes = NGHash(VarPtr(Null), UBound(Data) - LBound(Data) + 1, HashingAlgorithm)
    CatchAny eelCritical, "Error hashing data!", ModuleName & ".HashBytes", True, True
End Function

Private Function HashString(str As String, Optional HashingAlgorithm As String = DefaultHashAlgorithm) As Byte()
    If DebugMode(True) Then On Error Resume Next Else On Error Resume Next
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
Public Function GetStringHash(strText As String _
                            , Optional blnWithBom As Boolean = False _
                            , Optional ByVal intHashLen As Integer = DefaultShortHashLen) As String
    GetStringHash = GetHash(GetUTF8Bytes(strText, blnWithBom), intHashLen)
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetFileHash
' Author    : Adam Waller
' Date      : 11/30/2020
' Purpose   : Return a hash from a file
'---------------------------------------------------------------------------------------
'
Public Function GetFileHash(strPath As String _
                            , Optional ByVal intHashLen = DefaultShortHashLen) As String
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
                            , Optional ByVal intHashLen = DefaultShortHashLen) As String
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
                                , Optional ByVal intHashLen = DefaultShortHashLen) As String
    GetDictionaryHash = GetStringHash(ConvertToJson(dSource), , intHashLen)
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetHash
' Author    : Adam Waller
' Date      : 11/30/2020
' Purpose   : Create a hash from the byte array
'---------------------------------------------------------------------------------------
'
Private Function GetHash(bteContent() As Byte _
                        , Optional ByVal intHashLen As Integer = DefaultShortHashLen) As String
    
    Dim objEnc As Object
    Dim bteHash As Variant
    Dim strHash As String
    Dim intPos As Integer
    Dim intLength As Integer
    Dim strAlgorithm As String
    
    ' Get hashing options
    strAlgorithm = Nz2(Options.HashAlgorithm, DefaultHashAlgorithm)
    If UseShortHash Or (intHashLen <> DefaultShortHashLen) Then intLength = intHashLen
    
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
Public Function GetCodeModuleHash(intType As eDatabaseComponentType _
                                , strName As String _
                                , Optional intHashLen As Integer = DefaultShortHashLen) As String

    Dim strHash As String
    Dim cmpItem As VBComponent
    Dim strPrefix As String
    Dim proj As VBProject
    Dim blnNoCode As Boolean
    
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
        Set proj = GetVBProjectForCurrentDB
        
        ' Attempt to locate the object in the VBComponents collection
        If DebugMode(True) Then On Error Resume Next Else On Error Resume Next
        Set cmpItem = proj.VBComponents(strPrefix & strName)
        Catch 9 ' Component not found. (Could be an object with no code module)
        CatchAny eelError, "Error accessing VBComponent for '" & strPrefix & strName & "'", ModuleName & ".GetCodeModuleHash"
        If DebugMode(True) Then On Error GoTo 0 Else On Error Resume Next
                
        ' Output the hash
        If Not cmpItem Is Nothing Then
            With cmpItem.CodeModule
                strHash = GetStringHash(.Lines(1, 999999), , intHashLen)
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

