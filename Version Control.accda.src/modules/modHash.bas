'---------------------------------------------------------------------------------------
' Module    : modHash
' Author    : Adam Waller
' Date      : 12/4/2020
' Purpose   : Build SHA-1 hashes for content comparison.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Private Module
Option Explicit


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
    Set objEnc = CreateObject("System.Security.Cryptography.SHA1CryptoServiceProvider")
    bteSha1 = objEnc.ComputeHash_2(bteContent)
    Set objEnc = Nothing
    
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