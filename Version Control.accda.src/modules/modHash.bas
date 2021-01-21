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




#If Win64 And VBA7 Then

Public Declare PtrSafe Function BCryptOpenAlgorithmProvider Lib "BCrypt.dll" ( _
                            ByRef phAlgorithm As LongPtr, _
                            ByVal pszAlgId As LongPtr, _
                            ByVal pszImplementation As LongPtr, _
                            ByVal dwFlags As Long) As Long

Public Declare PtrSafe Function BCryptCloseAlgorithmProvider Lib "BCrypt.dll" ( _
                            ByVal hAlgorithm As LongPtr, _
                            ByVal dwFlags As Long) As Long

Public Declare PtrSafe Function BCryptCreateHash Lib "BCrypt.dll" ( _
                            ByVal hAlgorithm As LongPtr, _
                            ByRef phHash As LongPtr, pbHashObject As Any, _
                            ByVal cbHashObject As Long, _
                            ByVal pbSecret As LongPtr, _
                            ByVal cbSecret As Long, _
                            ByVal dwFlags As Long) As Long

Public Declare PtrSafe Function BCryptHashData Lib "BCrypt.dll" ( _
                            ByVal hHash As LongPtr, _
                            pbInput As Any, _
                            ByVal cbInput As Long, _
                            Optional ByVal dwFlags As Long = 0) As Long

Public Declare PtrSafe Function BCryptFinishHash Lib "BCrypt.dll" ( _
                            ByVal hHash As LongPtr, _
                            pbOutput As Any, _
                            ByVal cbOutput As Long, _
                            ByVal dwFlags As Long) As Long

Public Declare PtrSafe Function BCryptDestroyHash Lib "BCrypt.dll" (ByVal hHash As LongPtr) As Long

Public Declare PtrSafe Function BCryptGetProperty Lib "BCrypt.dll" ( _
                            ByVal hObject As LongPtr, _
                            ByVal pszProperty As LongPtr, _
                            ByRef pbOutput As Any, _
                            ByVal cbOutput As Long, _
                            ByRef pcbResult As Long, _
                            ByVal dfFlags As Long) As Long


#Else

Private Declare Function BCryptOpenAlgorithmProvider Lib "BCrypt.dll" ( _
                            ByRef phAlgorithm As Long, _
                            ByVal pszAlgId As Long
                            ByVal pszImplementation As Long
                            Optional ByVal dwFlags As Long) As Long

Private Declare Function BCryptCloseAlgorithmProvider Lib "BCrypt.dll" ( _
                            ByVal hAlgorithm As Long, _
                            Optional ByVal dwFlags As Long) As Long

Public Declare Function BCryptCreateHash Lib "BCrypt.dll" ( _
                            ByVal hAlgorithm As Long, _
                            ByRef phHash As Long, _
                            pbHashObject As Any, _
                            ByVal cbHashObject As Long, _
                            ByVal pbSecret As Long, _
                            ByVal cbSecret As Long, _
                            ByVal dwFlags As Long) As Long

Public Declare Function BCryptHashData Lib "BCrypt.dll" ( _
                            ByVal hHash As LongPtr, _
                            pbInput As Any, _
                            ByVal cbInput As Long, _
                            Optional ByVal dwFlags As Long = 0) As Long

Public Declare Function BCryptFinishHash Lib "BCrypt.dll" ( _
                            ByVal hHash As Long, _
                            pbOutput As Any, _
                            ByVal cbOutput As Long, _
                            ByVal dwFlags As Long) As Long

Public Declare Function BCryptDestroyHash Lib "BCrypt.dll" (ByVal hHash As Long) As Long

Public Declare Function BCryptGetProperty Lib "BCrypt.dll" ( _
                            ByVal hObject As Long, _
                            ByVal pszProperty As LongPtr, _
                            ByRef pbOutput As Any, _
                            ByVal cbOutput As Long, _
                            ByRef pcbResult As Long, _
                            ByVal dfFlags As Long) As Long

Private Declare Function BCryptEnumAlgorithms Lib "Brypt.dll" ( _
                            ByVal dwAlgOperations As Long, _
                            ByRef pAlgCount As Long, _
                            ByRef ppAlgList As Any, _
                            ByVal dwFlags As Long) As Long



#End If


Public Function NGHash(pData As LongPtr, lenData As Long, Optional HashingAlgorithm As String = "SHA1") As Byte()
    
    
    'Erik A, 2019
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
    Dim hashLength As Long
    cmd = "HashDigestLength" & vbNullChar
    If BCryptGetProperty(hAlg, StrPtr(cmd), hashLength, LenB(hashLength), 0, 0) <> 0 Then GoTo ErrHandler
    Dim bHash() As Byte
    ReDim bHash(0 To hashLength - 1)

    'Create hash object
    Dim hHash As LongPtr
    If BCryptCreateHash(hAlg, hHash, bHashObject(0), Length, 0, 0, 0) <> 0 Then GoTo ErrHandler

    'Hash data
    If BCryptHashData(hHash, ByVal pData, lenData) <> 0 Then GoTo ErrHandler
    If BCryptFinishHash(hHash, bHash(0), hashLength, 0) <> 0 Then GoTo ErrHandler

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
    If errorMessage <> "" Then MsgBox errorMessage
    Resume ExitHandler
End Function


Public Function HashBytes(Data() As Byte, Optional HashingAlgorithm As String = "SHA512") As Byte()
    HashBytes = NGHash(VarPtr(Data(LBound(Data))), UBound(Data) - LBound(Data) + 1, HashingAlgorithm)
End Function

Public Function hashString(str As String, Optional HashingAlgorithm As String = "SHA512") As Byte()
    
    hashString = NGHash(StrPtr(str), Len(str) * 2, HashingAlgorithm)
End Function

'---------------------------------------------------------------------------------------
' Procedure : GetStringHash
' Author    : Adam Waller
' Date      : 11/30/2020
' Purpose   : Convert string to byte array, and return a Sha1 hash.
'---------------------------------------------------------------------------------------
'
Public Function GetStringHash(strText As String, Optional intLength As Integer = 7) As String
    Dim bteText() As Byte
    bteText = strText
    GetStringHash = Sha1(bteText, intLength)
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetFileHash
' Author    : Adam Waller
' Date      : 11/30/2020
' Purpose   : Return a Sha1 hash from a file
'---------------------------------------------------------------------------------------
'
Public Function GetFileHash(strPath As String, Optional intLength As Integer = 7) As String
    GetFileHash = Sha1(GetFileBytes(strPath), intLength)
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetDictionaryHash
' Author    : Adam Waller
' Date      : 12/1/2020
' Purpose   : Wrapper to get a hash from a dictionary object (converted to json)
'---------------------------------------------------------------------------------------
'
Public Function GetDictionaryHash(dSource As Dictionary, Optional intLength As Integer = 7) As String
    GetDictionaryHash = GetStringHash(ConvertToJson(dSource), intLength)
End Function


'---------------------------------------------------------------------------------------
' Procedure : Sha1
' Author    : Adam Waller
' Date      : 11/30/2020
' Purpose   : Create a Sha1 hash of the byte array
'---------------------------------------------------------------------------------------
'
Private Function Sha1(bteContent() As Byte, Optional intLength As Integer) As String
    
    Dim objEnc As Object
    Dim bteSha1 As Variant
    Dim strSha1 As String
    Dim intPos As Integer
    
    Perf.OperationStart "Compute SHA1"
    'Set objEnc = CreateObject("System.Security.Cryptography.SHA1CryptoServiceProvider")
    'bteSha1 = objEnc.ComputeHash_2(bteContent)
    'Set objEnc = Nothing
    bteSha1 = HashBytes(bteContent, "SHA1")
    ' Create string buffer to avoid concatenation
    strSha1 = Space(LenB(bteSha1) * 2)
    
    ' Convert full sha1 to hexidecimal string
    For intPos = 1 To LenB(bteSha1)
        Mid$(strSha1, (intPos * 2) - 1, 2) = LCase(Right("0" & Hex(AscB(MidB(bteSha1, intPos, 1))), 2))
    Next
    
    ' Return hash, truncating if needed.
    If intLength > 0 And intLength < Len(strSha1) Then
        Sha1 = Left$(strSha1, intLength)
    Else
        Sha1 = strSha1
    End If
    Perf.OperationEnd
    
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetCodeModuleHash
' Author    : Adam Waller
' Date      : 11/30/2020
' Purpose   : Return a Sha1 hash of the VBA code module behind an object.
'---------------------------------------------------------------------------------------
'
Public Function GetCodeModuleHash(intType As eDatabaseComponentType, strName As String, Optional intLength As Integer = 7) As String

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
        On Error Resume Next
        Set cmpItem = proj.VBComponents(strPrefix & strName)
        Catch 9 ' Component not found. (Could be an object with no code module)
        If Err Then Log.Add "WARNING: Error accessing VBComponent for '" & strPrefix & strName & _
            "'. Error " & Err.Number & ": " & Err.Description, False
        On Error GoTo 0
        
        ' Output the hash
        If Not cmpItem Is Nothing Then
            With cmpItem.CodeModule
                strHash = GetStringHash(.Lines(1, 99999), intLength)
            End With
        End If
    
    End If
    
    ' Return hash (if any)
    GetCodeModuleHash = strHash
    Perf.OperationEnd
    
End Function