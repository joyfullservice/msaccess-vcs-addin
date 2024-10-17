VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsConnectionODBC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Author    : bclothier
' Date      : 4/1/2023
' Purpose   : This class provides the parsing of ODBC connection string into key-value
'           : pairs, conforming to the grammar defined by MS-ODBCSTR open specification.
'           : This class can be also used to aid in comparing whether 2 ODBC connection
'           : strings are considered equivalent by Access via the property named
'           : `SanitizedConnectionString`.
'           :
'           : For the complete speccification of MS-ODBCSTR, see:
'           : https://learn.microsoft.com/en-us/openspecs/sql_server_protocols/ms-odbcstr/13b4e848-b36c-4b11-acce-d6bf199d5391
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit

Private Type udtData
    ' 2.2 Generic keys
    Driver As String
    DSN As String
    FileDSN As String
    UID As String
    PWD As String
    SaveFile As String

    ' Those technically are not generic keys but they are very
    ' common among a variety of ODBC drivers and Access do treat
    ' two connection strings with a different values for those
    ' keys differently.
    Server As String
    Port As String ' Should be either an integer between 0 to 65535 or empty
    Database As String
    Encrypt As String

    ' The rest of members below are for class' use and is not related to
    ' the ODBC connection string specifications.
    OriginalConnectionString As String
    SanitizedConnectionString As String

    KeyValues As Dictionary
End Type
Private this As udtData


'---------------------------------------------------------------------------------------
' Procedure : Driver
' Author    : bclothier
' Date      : 4/1/2023
' Purpose   : Returns the value for the driver key in the connection string if present.
'---------------------------------------------------------------------------------------
'
Public Property Get Driver() As String
    Driver = this.Driver
End Property


'---------------------------------------------------------------------------------------
' Procedure : DSN
' Author    : bclothier
' Date      : 4/1/2023
' Purpose   : Returns the value for the DSN key in the connection string if present.
'---------------------------------------------------------------------------------------
'
Public Property Get DSN() As String
    DSN = this.DSN
End Property


'---------------------------------------------------------------------------------------
' Procedure : FileDSN
' Author    : bclothier
' Date      : 4/1/2023
' Purpose   : Returns the value for the FileDSN key in the connection string if present.
'---------------------------------------------------------------------------------------
'
Public Property Get FileDSN() As String
    FileDSN = this.FileDSN
End Property


'---------------------------------------------------------------------------------------
' Procedure : UID
' Author    : bclothier
' Date      : 4/1/2023
' Purpose   : Returns the value for the UID key in the connection string if present.
'---------------------------------------------------------------------------------------
'
Public Property Get UID() As String
    UID = this.UID
End Property


'---------------------------------------------------------------------------------------
' Procedure : PWD
' Author    : bclothier
' Date      : 4/1/2023
' Purpose   : Returns the value for the PWD key in the connection string if present.
'---------------------------------------------------------------------------------------
'
Public Property Get PWD() As String
    PWD = this.PWD
End Property


'---------------------------------------------------------------------------------------
' Procedure : SaveFile
' Author    : bclothier
' Date      : 4/1/2023
' Purpose   : Returns the value for the SaveFile key in the connection string if present.
'---------------------------------------------------------------------------------------
'
Public Property Get SaveFile() As String
    SaveFile = this.SaveFile
End Property


'---------------------------------------------------------------------------------------
' Procedure : Server
' Author    : bclothier
' Date      : 4/1/2023
' Purpose   : Returns the value for the Server key in the connection string if present.
'           : This is not a generic key but Access does uses it for comparison.
'---------------------------------------------------------------------------------------
'
Public Property Get Server() As String
    Server = this.Server
End Property


'---------------------------------------------------------------------------------------
' Procedure : Port
' Author    : bclothier
' Date      : 4/1/2023
' Purpose   : Returns the value for the Port key in the connection string if present.
'           : This is not a generic key but Access does uses it for comparison.
'---------------------------------------------------------------------------------------
'
Public Property Get Port() As String
    Port = this.Port
End Property


'---------------------------------------------------------------------------------------
' Procedure : Database
' Author    : bclothier
' Date      : 4/1/2023
' Purpose   : Returns the value for the Database key in the connection string if present.
'           : This is not a generic key but Access does uses it for comparison.
'---------------------------------------------------------------------------------------
'
Public Property Get Database() As String
    Database = this.Database
End Property


'---------------------------------------------------------------------------------------
' Procedure : Database
' Author    : bclothier
' Date      : 4/1/2023
' Purpose   : Returns the value for the Database key in the connection string if present.
'           : This is not a generic key but Access does uses it for comparison.
'---------------------------------------------------------------------------------------
'
Public Property Get Encrypt() As String
    Encrypt = this.Encrypt
End Property


'---------------------------------------------------------------------------------------
' Procedure : OriginalConnectionString
' Author    : bclothier
' Date      : 4/1/2023
' Purpose   : Returns the original connection string that was used in the call to
'           : ParseOdbcConnectionString.
'---------------------------------------------------------------------------------------
'
Public Property Get OriginalConnectionString() As String
    OriginalConnectionString = this.OriginalConnectionString
End Property


'---------------------------------------------------------------------------------------
' Procedure : SanitizedConnectionString
' Author    : bclothier
' Date      : 4/1/2023
' Purpose   : Returns the sanitized connection string that was returned from the call to
'           : ParseOdbcConnectionString. This is useful for comparing with other
'           : connection strings to determine if they are equivalent.
'---------------------------------------------------------------------------------------
'
Public Property Get SanitizedConnectionString() As String
    SanitizedConnectionString = this.SanitizedConnectionString
End Property


'---------------------------------------------------------------------------------------
' Procedure : GetKeyValues
' Author    : bclothier
' Date      : 4/1/2023
' Purpose   : Returns a dictionary of key/values. If a key is non-generic and has
'           : multiple values, the value are split with vbNullChar character. The
'           : generic key retains only the last value.
'---------------------------------------------------------------------------------------
'
Public Function GetKeyValues() As Dictionary
    Set GetKeyValues = New Dictionary
    GetKeyValues.CompareMode = TextCompare

    Dim varKey As Variant
    For Each varKey In this.KeyValues
        GetKeyValues.Add varKey, this.KeyValues(varKey)
    Next
End Function

Public Sub ParseOdbcConnectionString(strConnect As String)
    If StartsWith(strConnect, "ODBC;", vbTextCompare) = False Then
        Exit Sub
    End If

    ResetData

    ' 2.1.1 Common ABNF Rules
    ' SC           = %x3B         ; Semicolon
    ' LCB          = %x7B         ; Left curly brackets
    ' RCB          = %x7D         ; Right curly brackets
    ' EQ           = %x3D         ; Equal sign
    ' ESCAPEDRCB   = 2RCB         ; Double right curly brackets
    ' SpaceStr     = *(SP)        ; Any number of spaces (including 0 spaces)

    ' 2.1.2 ODBC Connection String Format
    ' ODBCConnectionString =  *(KeyValuePair SC) KeyValuePair [SC]
    ' KeyValuePair = (Key EQ Value / SpaceStr)
    ' Key = SpaceStr KeyName
    ' KeyName = (nonSP-SC-EQ *nonEQ)
    ' Value = (SpaceStr ValueFormat1 SpaceStr) / (ValueContent2)
    ' ValueFormat1 = LCB ValueContent1 RCB
    ' ValueContent1 = *(nonRCB / ESCAPEDRCB)
    ' ValueContent2 = SpaceStr / SpaceStr (nonSP-LCB-SC) *nonSC
    ' nonRCB = %x01-7C / %x7E-FFFF                                 ; not "}"
    ' nonSP-LCB-SC = %x01-1F / %x21-3A / %x3C-7A / %x7C-FFFF       ; not space, "{" or ";"
    ' nonSP-SC-EQ = %x01-1F / %x21-3A / %x3C / %x3E-FFFF           ; not space, ";" or "="
    ' nonEQ = %x01-3C / %x3E-FFFF                                  ; not "="
    ' nonSC = %x01-003A / %x3C-FFFF                                ; not ";"

    ' Use https://regex101.com/ to provide a detailed explanation of the pattern
    ' It should conform to the rules defined by the MS-ODBCSTR and extract the
    ' key as the first match and the value as either 2nd or 3rd match. The 2nd
    ' match represents the ValueFormat1 specification and 3rd match, the
    ' ValueContent2 specification.
    Const RegExpPattern As String = "\s?([^ ;=][^=]*?)=(?:\s?(\{(?:[^}]|\}\})*?\})\s?|\s?([^ ;{][^;]*)|\s?)(?:;|$)"

    Dim objRegExp As VBScript_RegExp_55.RegExp
    Dim objMatches As VBScript_RegExp_55.MatchCollection
    Dim objMatch As VBScript_RegExp_55.Match
    Dim strKey As String
    Dim strValue As String

    Set objRegExp = New VBScript_RegExp_55.RegExp
    With objRegExp
        .Global = True
        .IgnoreCase = True
        .Multiline = True
        .Pattern = RegExpPattern
        ' Test only the substring without the `ODBC;` prefix which technically is not
        ' a part of the ODBC connection string but rather a protocol specifier used by
        ' Access itself.
        Set objMatches = .Execute(Mid$(strConnect, 6))
    End With

    this.OriginalConnectionString = strConnect
    Set this.KeyValues = New Dictionary
    this.KeyValues.CompareMode = TextCompare

    For Each objMatch In objMatches
        With objMatch
            strKey = .SubMatches(0)
            If Len(.SubMatches(1)) Then
                strValue = .SubMatches(1) ' ValueFormat1
            ElseIf Len(.SubMatches(2)) Then
                strValue = .SubMatches(2) ' ValueContent2
            Else
                strValue = vbNullString ' No matches; assume empty string
            End If

            If this.KeyValues.Exists(strKey) Then
                ' 2.2.3 indicates generic key takes the last value. Driver specific
                ' key however is driver-defined. We'll use vbNullChar as delimiter
                ' to support weird oddball drivers that likes having multiple values.
                '
                ' NOTE: according to 3.10, SQL Server drivers will take the first value
                ' and ignore subsequent values.
                If IsGenericKey(strKey) Then
                    this.KeyValues(strKey) = strValue
                Else
                    this.KeyValues(strKey) = this.KeyValues(strKey) & vbNullChar & strValue
                End If
            Else
                this.KeyValues.Add strKey, strValue
            End If

            Select Case strKey
                Case "Driver"
                    this.Driver = strValue
                Case "DSN"
                    this.DSN = strValue
                Case "FileDSN"
                    this.FileDSN = strValue
                Case "SaveFile"
                    this.SaveFile = strValue
                Case "UID"
                    this.UID = strValue
                Case "PWD"
                    this.PWD = strValue

                ' The following keys are non-generic so take the first value
                ' for the purpose of sanitized connection string comparison
                Case "Server"
                    If Len(this.Server) = 0 Then
                        this.Server = strValue
                    End If
                Case "Database"
                    If Len(this.Database) = 0 Then
                        this.Database = strValue
                    End If
                Case "Port"
                    If Len(this.Port) = 0 Then
                        this.Port = strValue
                    End If
                Case "Encrypt"

                    this.Encrypt = CBool(strValue)
            End Select
        End With
    Next

    ' Build the sanitized connection string based on key parameters that Access will use
    ' to differeniate one connection string from another. This is based on experience
    ' because there is no formal documentation for how Access treats one connection string.
    ' WizHook provides IsMatchToDbcConnectString but that does not seem to suit our needs.
    ' What is known:
    '   Access pays attention to the following keywords: Driver, Server, Database and Port
    '   Access does not consider UID or PWD
    '   Other keys are ignored.
    '   Order of keys appearing in the connection string seems to have a factor.
    '
    ' As a matter of best practice, it's best that we avoid having multiple variants of
    ' essentially same connection string and user probably should consider multiple
    ' occurrences of same connection strings differently formatted as a potential problem
    ' to be fixed to avoid runtime problems such as Access prompting for credentials
    ' in unexpected places. So for this purpose, we will allow multiple variants and leave it
    ' up to the user to make corrections to their project to minimize the variants.
    '
    ' According to 2.3.2 & 3.8, generally the key value pairs in the connection string
    ' takes precedence so if we find `Driver` key, we always use that, followed by `DSN`
    ' then finally `FileDSN`. The specification also mentions FileDSN needing to precede
    ' DSN but we don't handle this.

    Dim objConcat As clsConcat
    Set objConcat = New clsConcat

    If Len(this.Driver) Then
        objConcat.Add ";Driver=", this.Driver
    ElseIf Len(this.DSN) Then
        objConcat.Add ";DSN=", this.DSN
    ElseIf Len(this.FileDSN) Then
        objConcat.Add ";FileDSN=", this.FileDSN
    Else
        ' Apparently this is an incomplete ODBC connection string. Access would always prompt to fill in.... Masochistic?
    End If

    If Len(this.SaveFile) Then
        objConcat.Add ";SaveFile=", this.SaveFile
    End If

    If Len(this.Server) Then
        objConcat.Add ";SERVER=", this.Server
    End If

    If Len(this.Port) Then
        objConcat.Add ";PORT=", this.Port
    End If

    If Len(this.Database) Then
        objConcat.Add ";DATABASE=", this.Database
    End If

    If Len(this.UID) Then
        objConcat.Add ";UID=", this.UID
    End If

    If Len(this.PWD) Then
        objConcat.Add ";PWD=", this.PWD
    End If

    this.SanitizedConnectionString = "ODBC" & objConcat.GetStr
End Sub


'---------------------------------------------------------------------------------------
' Procedure : IsGenericKey
' Author    : bclothier
' Date      : 4/1/2023
' Purpose   : Indicates if a given key name is generic or not, as specified in
'           : MS-ODBCSTR 2.2 Generic Keys
'---------------------------------------------------------------------------------------
'
Public Function IsGenericKey(strKey As String) As Boolean
    Select Case strKey
        Case "Driver", _
             "DSN", _
             "FileDSN", _
             "UID", _
             "PWD", _
             "SaveFile"
            IsGenericKey = True
    End Select
End Function


'---------------------------------------------------------------------------------------
' Procedure : ResetData
' Author    : bclothier
' Date      : 4/1/2023
' Purpose   : Resets the cached values
'---------------------------------------------------------------------------------------
'
Private Sub ResetData()
    Dim blank As udtData

    ' Not necessary but best to ensure proper de-referencing
    Set this.KeyValues = Nothing

    LSet this = blank
End Sub
