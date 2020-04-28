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
'---------------------------------------------------------------------------------------
Option Compare Database
Option Private Module
Option Explicit


'---------------------------------------------------------------------------------------
' Procedure : SetEncryptionKey
' Author    : Adam Waller
' Date      : 4/24/2020
' Purpose   : Sets the encryption key in the current user's registry.
'---------------------------------------------------------------------------------------
'
Public Sub SetEncryptionKey(strKey As String)
    SaveSetting GetCodeVBProject.Name, "Add-in", "Encryption Key", strKey
End Sub


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
    Encrypt = "@{" & LCase(EncryptRC4("RC4" & strText, GetKey)) & "}"
End Function


'---------------------------------------------------------------------------------------
' Procedure : Decrypt
' Author    : Adam Waller
' Date      : 4/24/2020
' Purpose   : Decrypt the string using the saved key.
'---------------------------------------------------------------------------------------
'
Public Function Decrypt(ByRef strToDecrypt As String) As Boolean
    Dim strDecrypted As String
    strDecrypted = DecryptRC4(Mid(UCase(strToDecrypt), 3, Len(strToDecrypt) - 4), GetKey)
    If Left$(strDecrypted, 3) = "RC4" Then
        ' Successfully decrypted.
        strToDecrypt = Mid$(strDecrypted, 4)
        Decrypt = True
    End If
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetKey
' Author    : Adam Waller
' Date      : 4/24/2020
' Purpose   : Get the current key from the registry. (Or the CodeProject name if no key set.)
'---------------------------------------------------------------------------------------
'
Private Function GetKey() As String
    GetKey = GetSetting(GetCodeVBProject.Name, "Add-in", "Encryption Key", CodeProject.Name)
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
        ToHexDump = ToHexDump & Right$("0" & Hex(Asc(Mid(sText, lIdx, 1))), 2)
    Next
End Function

Private Function FromHexDump(sText As String) As String
    Dim lIdx As Long
    For lIdx = 1 To Len(sText) Step 2
        FromHexDump = FromHexDump & Chr$(CLng("&H" & Mid(sText, lIdx, 2)))
    Next
End Function