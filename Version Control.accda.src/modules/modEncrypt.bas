'---------------------------------------------------------------------------------------
' Module    : basEncrypt
' Author    : Adam Waller
' Date      : 4/9/2020
' Purpose   : Adapted from: https://stackoverflow.com/questions/7025644/vb6-encrypt-text-using-password
'           :
'           : *** IMPORTANT!! ***
'           : This is not considered a secure encryption algorithm for sensitive data.
'           : If you need something more secure, please utilize actual cryptography
'           : API calls or functions. This is intended simply as a basic way of masking
'           : semi-secure data in source code.
'           : See: https://github.com/joyfullservice/msaccess-vcs-integration/wiki/Encryption
'---------------------------------------------------------------------------------------
Option Compare Database
Option Private Module
Option Explicit


Public Const DefaultKeyName = "MSAccessVCS"

Private m_Name As String
Private m_Key As String


'---------------------------------------------------------------------------------------
' Procedure : Secure
' Author    : Adam Waller
' Date      : 6/1/2020
' Purpose   : Secure the text based on the loaded option.
'---------------------------------------------------------------------------------------
'
Public Function Secure(strText As String) As String
    Select Case Options.Security
        Case esEncrypt: Secure = Encrypt(strText)
        Case esRemove:  Secure = vbNullString
        Case esNone:    Secure = strText
    End Select
End Function


'---------------------------------------------------------------------------------------
' Procedure : SecureBetween
' Author    : Casper Englund
' Date      : 2020-06-03
' Purpose   : Secures content between two strings.
'---------------------------------------------------------------------------------------
'
Public Function SecureBetween(strText As String, strStartAfter As String, strEndBefore As String, Optional Compare As VbCompareMethod) As String
        
        If strText = vbNullString Or Options.Security = esNone Then
            SecureBetween = strText
        Else
            If Options.Security = esEncrypt Then
                SecureBetween = EncryptBetween(strText, strStartAfter, strEndBefore, Compare)
            ElseIf Options.Security = esRemove Then
                Dim lngPos As Long
                Dim lngStart As Long
                Dim lngLen As Long
                
                lngPos = InStr(1, strText, strStartAfter, Compare)
                If lngPos > 0 Then
                    lngStart = lngPos + Len(strStartAfter) - 1
                    lngPos = InStr(lngStart + 1, strText, strEndBefore)
                    If lngPos > 0 Then
                        lngLen = lngPos - lngStart
                    End If
                End If
                
                If lngLen = 0 Then
                    ' No tags found. Return original string
                    SecureBetween = strText
                Else
                    SecureBetween = Left$(strText, lngStart) & Mid$(strText, lngStart + lngLen)
                End If
    
            End If
        End If
        
End Function


'---------------------------------------------------------------------------------------
' Procedure : SecurePath
' Author    : Adam Waller
' Date      : 6/1/2020
' Purpose   : Secures just the folder path, not the filename.
'---------------------------------------------------------------------------------------
'
Public Function SecurePath(strPath As String) As String

    Dim strParent As String

    strParent = FSO.GetParentFolderName(strPath)
    If strParent = vbNullString Then
        ' Could be relative path or just a filename.
        SecurePath = strPath
    Else
        If Options.Security = esRemove Then
            SecurePath = FSO.GetFileName(strPath)
        Else
            ' Could be encrypted or plain text, depending on options.
            SecurePath = FSO.BuildPath(Secure(strParent), FSO.GetFileName(strPath))
        End If
    End If

End Function


'---------------------------------------------------------------------------------------
' Procedure : IsEncrypted
' Author    : Adam Waller
' Date      : 4/28/2020
' Purpose   : Returns true if the value appears to be encrypted.
'---------------------------------------------------------------------------------------
'
Public Function IsEncrypted(strText As String) As Boolean
    IsEncrypted = (Left$(strText, 2) = "@{" And Right$(strText, 1) = "}")
End Function


'---------------------------------------------------------------------------------------
' Procedure : EncryptionKeySet
' Author    : Adam Waller
' Date      : 4/24/2020
' Purpose   : Returns true if the encryption key has been set.
'---------------------------------------------------------------------------------------
'
Public Function EncryptionKeySet() As Boolean
    EncryptionKeySet = (GetKey <> CodeProject.Name)
End Function


'---------------------------------------------------------------------------------------
' Procedure : Encrypt
' Author    : Adam Waller
' Date      : 4/24/2020
' Purpose   : Encrypt a string using the saved key. Uses random key if none set.
'---------------------------------------------------------------------------------------
'
Public Function Encrypt(strText As String) As String
    If strText <> vbNullString Then Encrypt = "@{" & LCase$(EncryptRC4("RC4" & strText, GetKey)) & "}"
End Function


'---------------------------------------------------------------------------------------
' Procedure : EncryptBetween
' Author    : Adam Waller
' Date      : 5/6/2020
' Purpose   : Encrypt an embedded string. (Such as the path in an XML document)
'---------------------------------------------------------------------------------------
'
Public Function EncryptBetween(strText As String, strStartAfter As String, strEndBefore As String, Optional Compare As VbCompareMethod) As String
    
    Dim lngPos As Long
    Dim lngStart As Long
    Dim lngLen As Long
    
    lngPos = InStr(1, strText, strStartAfter, Compare)
    If lngPos > 0 Then
        lngStart = lngPos + Len(strStartAfter) - 1
        lngPos = InStr(lngStart + 1, strText, strEndBefore)
        If lngPos > 0 Then
            lngLen = lngPos - lngStart
        End If
    End If
    
    If lngLen = 0 Then
        ' No tags found. Return original string
        EncryptBetween = strText
    Else
        EncryptBetween = Left$(strText, lngStart) & _
            Encrypt(Mid$(strText, lngStart, lngLen)) & _
            Mid$(strText, lngStart + lngLen)
    End If
    
End Function


'---------------------------------------------------------------------------------------
' Procedure : Decrypt
' Author    : Adam Waller
' Date      : 4/24/2020
' Purpose   : Decrypt the string using the saved key. (Keep in mind that only part(s) of
'           : the string may be encrypted.)
'---------------------------------------------------------------------------------------
'
Public Function Decrypt(strToDecrypt As String) As String

    Dim strSegment As String
    Dim strTest As String
    Dim strDecrypted As String
    Dim lngStart As Long
    Dim lngEnd As Long
    
    ' Start search at first character in string
    lngStart = 1
    lngEnd = 1
    
    ' Loop through each encrypted part of the string
    Do
        ' Identify encrypted portion of the string.
        lngStart = InStr(lngStart, strToDecrypt, "@{")
    
        ' Any more tags found?
        If lngStart < 1 Then
            If lngEnd < 1 Then
                ' Might not have been anything to decrypt
                strDecrypted = strToDecrypt
            Else
                ' Add any remaining portion of the string
                strDecrypted = strDecrypted & Mid$(strToDecrypt, lngEnd)
            End If
            Exit Do
        End If
    
        ' Add any intermediate text
        If lngStart > lngEnd Then
            strDecrypted = strDecrypted & Mid$(strToDecrypt, lngEnd, lngStart - lngEnd)
        End If
        
        ' Look for ending termination
        lngEnd = InStr(lngStart + 3, strToDecrypt, "}") + 1
        If lngEnd > 1 Then
            ' Get full encrypted segment
            strSegment = Mid$(strToDecrypt, lngStart, lngEnd - lngStart)
            ' Decrypt this segment.
            strTest = DecryptRC4(Mid$(UCase$(strSegment), 3, lngEnd - lngStart - 3), GetKey)
            If Left$(strTest, 3) = "RC4" Then
                ' Successfully decrypted.
                strDecrypted = strDecrypted & Mid$(strTest, 4)
            Else
                ' Leave encrypted string
                strDecrypted = strDecrypted & strSegment
            End If
            ' Move to next position
            lngStart = lngEnd
        End If
    Loop
    
    ' Return decrypted value
    Decrypt = strDecrypted
    
End Function


'---------------------------------------------------------------------------------------
' Procedure : CanDecrypt
' Author    : Adam Waller
' Date      : 7/29/2020
' Purpose   : Returns true of the encrypted string can be successfully decrypted.
'---------------------------------------------------------------------------------------
'
Public Function CanDecrypt(strEncrypted As String) As Boolean
    CanDecrypt = (strEncrypted <> Decrypt(strEncrypted))
End Function


'---------------------------------------------------------------------------------------
' Procedure : SetEncryptionKey
' Author    : Adam Waller
' Date      : 4/24/2020
' Purpose   : Sets the encryption key in the current user's registry.
'---------------------------------------------------------------------------------------
'
Public Sub SetEncryptionKey(strName As String, strKey As String)
    SaveSetting GetCodeVBProject.Name, "Private Keys", strName, strKey
    m_Name = strName
    m_Key = strKey
End Sub


'---------------------------------------------------------------------------------------
' Procedure : GetKey
' Author    : Adam Waller
' Date      : 4/24/2020
' Purpose   : Get the current key from the registry. (Or the CodeProject name if no key set.)
'---------------------------------------------------------------------------------------
'
Private Function GetKey() As String
    If m_Name = vbNullString Then m_Name = Options.KeyName
    If m_Name = vbNullString Then m_Name = DefaultKeyName
    If m_Key = vbNullString Then m_Key = GetSetting(GetCodeVBProject.Name, "Private Keys", m_Name, CodeProject.Name)
    ' Return cached key name, rather than looking it up from the registry each time.
    GetKey = m_Key
End Function


'---------------------------------------------------------------------------------------
' Procedure : EncryptRC4
' Author    : Adam Waller
' Date      : 4/9/2020
' Purpose   : Encrypt some text with a key. (Reversible Encryption)
'---------------------------------------------------------------------------------------
'
Private Function EncryptRC4(strText As String, strKey As String) As String
    EncryptRC4 = ToHexDump(CryptRC4(strText, strKey))
End Function


'---------------------------------------------------------------------------------------
' Procedure : DecryptRC4
' Author    : Adam Waller
' Date      : 4/9/2020
' Purpose   : Decrypt the text using a key.
'---------------------------------------------------------------------------------------
'
Private Function DecryptRC4(strEncrypted As String, strKey As String) As String
    DecryptRC4 = CryptRC4(FromHexDump(strEncrypted), strKey)
End Function


' The following code is credited to https://stackoverflow.com/questions/7025644/vb6-encrypt-text-using-password
Private Function CryptRC4(sText As String, sKey As String) As String
    Dim baS(0 To 255) As Byte
    Dim baK(0 To 255) As Byte
    Dim bytSwap     As Byte
    Dim lI          As Long
    Dim lJ          As Long
    Dim lIdx        As Long

    For lIdx = 0 To 255
        baS(lIdx) = lIdx
        baK(lIdx) = Asc(Mid$(sKey, 1 + (lIdx Mod Len(sKey)), 1))
    Next
    For lI = 0 To 255
        lJ = (lJ + baS(lI) + baK(lI)) Mod 256
        bytSwap = baS(lI)
        baS(lI) = baS(lJ)
        baS(lJ) = bytSwap
    Next
    lI = 0
    lJ = 0
    For lIdx = 1 To Len(sText)
        lI = (lI + 1) Mod 256
        lJ = (lJ + baS(lI)) Mod 256
        bytSwap = baS(lI)
        baS(lI) = baS(lJ)
        baS(lJ) = bytSwap
        CryptRC4 = CryptRC4 & Chr$((pvCryptXor(baS((CLng(baS(lI)) + baS(lJ)) Mod 256), Asc(Mid$(sText, lIdx, 1)))))
    Next
End Function

Private Function pvCryptXor(ByVal lI As Long, ByVal lJ As Long) As Long
    If lI = lJ Then
        pvCryptXor = lJ
    Else
        pvCryptXor = lI Xor lJ
    End If
End Function

Private Function ToHexDump(sText As String) As String
    Dim lIdx As Long
    For lIdx = 1 To Len(sText)
        ToHexDump = ToHexDump & Right$("0" & Hex$(Asc(Mid$(sText, lIdx, 1))), 2)
    Next
End Function

Private Function FromHexDump(sText As String) As String
    Dim lIdx As Long
    For lIdx = 1 To Len(sText) Step 2
        FromHexDump = FromHexDump & Chr$(CLng("&H" & Mid$(sText, lIdx, 2)))
    Next
End Function