VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDotEnv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : clsDotEnv
' Author    : Adam Waller
' Date      : 7/20/2023
' Purpose   : Simple wrapper class for working with .env files
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit


' Dictionary of lines in the .env file
Public Lines As Dictionary


'---------------------------------------------------------------------------------------
' Procedure : LoadFromFile
' Author    : Adam Waller
' Date      : 7/20/2023
' Purpose   : Load .env file contents into dictionary of lines.
'           : (Comments will have incremented keys representing the current line number
'           : like: "COMMENT_001 = # My Comment")
'---------------------------------------------------------------------------------------
'
Public Function LoadFromFile(strFilePath As String)

    Dim varLines As Variant
    Dim strLine As String
    Dim lngLine As Long
    Dim lngPos As Long
    Dim strKey As String
    Dim strValue As String

    ' Reset Lines dictionary
    Class_Initialize

    ' Attempt to read file to array of lines
    varLines = Split(ReadFile(strFilePath), vbCrLf)
    For lngLine = 0 To UBound(varLines)
        strLine = varLines(lngLine)
        If (Left(Trim(strLine), 1) = "#") _
            Or (Trim(strLine) = vbNullString) Then
            ' Add comment line or blank line
            strKey = "COMMENT_" & Format(lngLine, "###")
            Lines(strKey) = strLine
        Else
            ' Find first equals sign
            lngPos = InStr(2, strLine, "=")
            If lngPos > 0 Then
                strKey = Trim(Left(strLine, lngPos - 1))
                strValue = Mid(strLine, lngPos + 1)
                Lines(strKey) = strValue
            Else
                ' Neither comment or key/value pair. Discard line
            End If
        End If
    Next lngLine

End Function


'---------------------------------------------------------------------------------------
' Procedure : SaveToFile
' Author    : Adam Waller
' Date      : 7/20/2023
' Purpose   : Save lines dictionary to file
'---------------------------------------------------------------------------------------
'
Public Function SaveToFile(strFilePath As String)

    Dim varKey As Variant
    Dim strKey As String

    ' Rebuild content in .env file format
    With New clsConcat
        .AppendOnAdd = vbCrLf
        For Each varKey In Lines.Keys
            strKey = varKey
            If strKey Like "COMMENT_*" Then
                ' Comment line
                .Add Lines(strKey)
            Else
                ' Key=value pair
                .Add strKey, "=", Lines(strKey)
            End If
        Next varKey

        ' Remove any trailing blank lines
        Do While .RightStr(2) = vbCrLf
            .Remove 2
        Loop

        ' Write to file
        WriteFile .GetStr, strFilePath
    End With

End Function


'---------------------------------------------------------------------------------------
' Procedure : GetVar
' Author    : Adam Waller
' Date      : 7/20/2023
' Purpose   : Return the variable from the key value
'---------------------------------------------------------------------------------------
'
Public Function GetVar(strKey As String, Optional blnUseEnviron As Boolean = True) As String
    Dim strEnv As String
    If blnUseEnviron Then strEnv = Environ(strKey)
    If Len(strEnv) Then
        ' Use environment variable as override
        GetVar = strEnv
    Else
        If Lines.Exists(strKey) Then GetVar = Lines(strKey)
    End If
End Function


'---------------------------------------------------------------------------------------
' Procedure : SetVar
' Author    : Adam Waller
' Date      : 7/20/2023
' Purpose   : Save a key=value pair, updating if it already exists.
'---------------------------------------------------------------------------------------
'
Public Sub SetVar(strKey As String, strValue As String)
    Lines(strKey) = strValue
End Sub


'---------------------------------------------------------------------------------------
' Procedure : MergeIntoDictionary
' Author    : Adam Waller
' Date      : 7/20/2023
' Purpose   : Merge the .env keys into the destination dictionary. (Excluding comments)
'---------------------------------------------------------------------------------------
'
Public Sub MergeIntoDictionary(ByRef dDestination As Dictionary, Optional blnUseEnviron As Boolean = True)
    Dim varKey As Variant
    For Each varKey In Lines.Keys
        If varKey Like "COMMENT_*" Then
            ' Skip comment lines
        Else
            ' Add or replace existing values
            dDestination(varKey) = GetVar(CStr(varKey), blnUseEnviron)
        End If
    Next varKey
End Sub


'---------------------------------------------------------------------------------------
' Procedure : Class_Initialize
' Author    : Adam Waller
' Date      : 7/20/2023
' Purpose   : Initialize the lines dictionary
'---------------------------------------------------------------------------------------
'
Private Sub Class_Initialize()
    Set Lines = New Dictionary
    Lines.CompareMode = TextCompare ' Case insensitive for key lookups
End Sub
