Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Author    : hecon5
' Date      : 2/28/2021
' Purpose   : This class extends the IDbComponent class to perform the specific
'           : operations required by this particular object type.
'           : (I.e. The specific way you export or import this component.)
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit

'Private m_Type as "Qry/Table/LinkedTable/etc"
Private m_StrSQL As String
Private m_StrConnection As String
Private m_Dbs As Database

Private Const ModuleName = "clsSQLDataTools"

Private m_Query As AccessObject
Private m_AllItems As Collection
Private m_blnModifiedOnly As Boolean




'---------------------------------------------------------------------------------------
' Procedure : SanitizeConnection
' Author    : hecon5 / Adam Waller
' Date      : 02/26/2021
' Purpose   : Sanitizes a connection string.
'---------------------------------------------------------------------------------------
Public Function SanitizeConnectionString(strConnection As String) As String
    Dim strPart as String

    ' Apply additional sanitizing when turned on (Default)
    If Options.AggressiveSanitize Then
    
        ' Remove UID and PWD values if using a trusted connection.
        ' (These are not required or needed in this context, and create differences
        '  in exported source files when the project is built by different users.)
        With New RegExp
            .Global = True
            .IgnoreCase = True
            .Pattern = "(.*)(;UID=)[^ \r\n]*(?=;Trusted_Connection=Yes)|(;PWD=)[^ \r\n]*(?=;Trusted_Connection=Yes)|(DATABASE=)([^; \r\n]*)(;)"
            SanitizeConnectionString = .Replace(strConnection, "$1$4") & GetRelativeConnect(.Replace(strConnection, "$5")) & .Replace(strConnection, "$6")
        End With
    Else
        ' Not using the AggressiveSanitize option
        SanitizeConnectionString = strConnection
    End If
    
End Function

'---------------------------------------------------------------------------------------
' Procedure : GetRelativeConnect
' Author    : Adam Waller
' Date      : 2/22/2021
' Purpose   : Convert the connection string to a relative path if possible
'---------------------------------------------------------------------------------------
'
Private Function GetRelativeConnect(strConnect As String) As String
    
    Dim strPath As String
    Dim strAdd As String
    
    ' Extract the file path from the connection string
    strPath = GetConnectPath(strConnect)
    
    ' With linked text files, it may just be the folder path. If so, include
    ' a final slash so it can correctly resolve to a relative path.
    If strPath = GetUncPath(CurrentProject.Path) Then strAdd = PathSep
    
    ' Prefix with the database property name, just in case the same string is used in
    ' another part of the connection string.
    GetRelativeConnect = Replace(strConnect, "DATABASE=" & strPath, _
        "DATABASE=" & GetRelativePath(strPath & strAdd), , , vbTextCompare)

End Function


'---------------------------------------------------------------------------------------
' Procedure : GetFullConnect
' Author    : Adam Waller
' Date      : 2/22/2021
' Purpose   : Return the full path expanded from a relative one.
'---------------------------------------------------------------------------------------
'
Private Function GetFullConnect(strRelative As String) As String

    Dim strPath As String
    Dim strFull As String
    
    ' Extract the file path from the connection string
    strPath = GetConnectPath(strRelative)
    
    ' Convert to full path
    strFull = GetPathFromRelative(strPath)
    
    ' Remove any trailing slash after conversion from relative.
    If strFull <> strPath And Right$(strFull, 1) = PathSep Then
        strFull = Left$(strFull, Len(strFull) - 1)
    End If
    
    ' Prefix with the database property name, just in case the same string is used in
    ' another part of the connection string.
    GetFullConnect = Replace(strRelative, "DATABASE=" & strPath, _
        "DATABASE=" & strFull, , , vbTextCompare)

End Function


'---------------------------------------------------------------------------------------
' Procedure : GetConnectPath
' Author    : Adam Waller
' Date      : 2/22/2021
' Purpose   : Return embedded connect path
'---------------------------------------------------------------------------------------
'
Private Function GetConnectPath(strConnect As String) As String

    Dim lngPos As Integer
    Dim lngStart As String
    Dim strPath As String
    
    ' Look for path after "Database" connection property
    lngPos = InStr(1, strConnect, "Database=", vbTextCompare)
    If lngPos > 0 Then
        lngStart = lngPos + Len("Database=")
        lngPos = InStr(lngStart, strConnect, ";")
        If lngPos > 0 Then
            ' Stop at semi-colon
            strPath = Mid$(strConnect, lngStart, lngPos - lngStart)
        Else
            ' Grab the rest of the string
            strPath = Mid$(strConnect, lngStart)
        End If
    End If
    
    ' Return path, if any
    GetConnectPath = strPath
    
End Function