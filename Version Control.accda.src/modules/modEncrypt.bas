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
' Procedure : EncryptionKeySet
' Author    : Adam Waller
' Date      : 4/24/2020
' Purpose   : Returns true if the encryption key has been set.
'---------------------------------------------------------------------------------------
'
Public Function EncryptionKeySet() As Boolean
    EncryptionKeySet = (GetKey <> CodeProject.Name)
End Function


Private Function GetKey() As String
    If m_Name = vbNullString Then m_Name = Options.KeyName
    If m_Name = vbNullString Then m_Name = DefaultKeyName
    If m_Key = vbNullString Then m_Key = GetSetting(GetCodeVBProject.Name, "Private Keys", m_Name, CodeProject.Name)
    ' Return cached key name, rather than looking it up from the registry each time.
    GetKey = m_Key
End Function