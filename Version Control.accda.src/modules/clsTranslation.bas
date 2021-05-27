VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTranslation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : clsTranslation
' Author    : Adam Waller
' Date      : 5/15/2021
' Purpose   : Used for the translation of strings to different languages, similar to
'           : the gettext toolset.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit

Private Const en_US As String = "en_US"

' Cache strings to dictionary objects so we don't have to do database lookups
' each time we need to return translated strings
Private dStrings As Dictionary
Private dEnglish As Dictionary
Private dTranslation As Dictionary
Private m_strCurrentLanguage As String

' A private type to work with file entries
Private Type tEntry
    strHeader As String
    strContext As String
    strMsgID As String
    strTrans As String
End Type


'---------------------------------------------------------------------------------------
' Procedure : GetLanguageName
' Author    : Adam Waller
' Date      : 5/26/2021
' Purpose   : Return the display name for the language. (Add new languages here)
'---------------------------------------------------------------------------------------
'
Private Function GetLanguageName(strIdentifier As String) As String
    Select Case strIdentifier
        Case en_US:     GetLanguageName = "English"
        Case "pt_BR":   GetLanguageName = "Brazilian Portuguese"
        ' For undefined languages, use the identifier for now.
        Case Else:      GetLanguageName = strIdentifier
    End Select
End Function


'---------------------------------------------------------------------------------------
' Procedure : T
' Author    : Adam Waller
' Date      : 5/15/2021
' Purpose   : Return the translated version of the string.
'---------------------------------------------------------------------------------------
'
Public Function T(strText As String, Optional strContext As String, Optional strComments As String) As String

    Dim strNew As String
    Dim strKey As String
    
    ' Skip processing if no value is passed
    If strText = vbNullString Then Exit Function
    
    ' Make sure the language has been initialized
    CheckInit
    
    ' Check for the master string
    strKey = BuildKey(strText, strContext)
    If dStrings.Exists(strKey) Then
        If dTranslation.Exists(dStrings(strKey)) Then
            ' Use translated string
            strNew = dTranslation(dStrings(strKey))
        End If
    Else
        ' Add to master list of strings (no translation exists)
        dStrings.Add strKey, strKey
        ' Add to strings table
        SaveString strText, strContext, strComments
    End If
    
    ' Return translated string
    T = Coalesce(strNew, strText)

End Function


'---------------------------------------------------------------------------------------
' Procedure : ApplyTo
' Author    : Adam Waller
' Date      : 5/17/2021
' Purpose   : Apply language translation to a form object (From English values)
'---------------------------------------------------------------------------------------
'
Public Sub ApplyTo(frmObject As Form)
   
    Dim ctl As Control
    Dim ctl2 As Control
    Dim ctlAsc As Control
    Dim strContext As String
    Dim strName As String
    
    ' No translation needed for English
    'If m_strCurrentLanguage = en_US Then Exit Sub
    
    ' Loop through all controls
    For Each ctl In frmObject.Controls
        
        ' Only check certain types of controls
        Select Case TypeName(ctl)
            Case "Label"
            
                ' Build base context
                strContext = frmObject.Name & "." & ctl.Name
                
                ' Check for associated control
                ' (It is easier to go from the object to the label, but not
                '  all labels may have objects, so we loop through other controls
                On Error Resume Next
                For Each ctl2 In frmObject.Controls
                    strName = vbNullString
                    strName = ctl2.Controls(0).Name
                    If strName = ctl.Name Then
                        ' Found associated label
                        ' Add extended context
                        strContext = strContext & "(" & ctl2.Name & ")"
                        Exit For
                    End If
                Next ctl2
                If DebugMode(False) Then On Error GoTo 0 Else On Error Resume Next
                
                ' Translation caption
                ctl.Caption = T(ctl.Caption, strContext)
                
            Case "TextBox"
                ' Nothing to translate
                
            Case "Page"
                ' Tab control page caption
                strContext = frmObject.Name & "." & ctl.Parent.Name & "." & ctl.Name
                ctl.Caption = T(ctl.Caption, strContext)
        
        End Select
        
    Next ctl
    
    ' Other properties
    frmObject.Caption = T(frmObject.Caption, frmObject.Name & ".Caption")

End Sub


'---------------------------------------------------------------------------------------
' Procedure : ExportTranslations
' Author    : Adam Waller
' Date      : 5/19/2021
' Purpose   : Export translations to files
'---------------------------------------------------------------------------------------
'
Public Sub ExportTranslations()

    Dim dbs As Database
    Dim rst As Recordset
    Dim strLanguage As String
    Dim strFolder As String
    Dim strFile As String
    
    strFolder = TranslationsPath
    If strFolder = vbNullString Then Exit Sub
    
    Set dbs = CodeDb
    Set rst = dbs.OpenRecordset("tblLanguages", dbOpenSnapshot)
    With rst
        Do While Not .EOF
            strLanguage = Nz(!ID)
            If strLanguage <> vbNullString Then
                If strLanguage = en_US Then
                    ' Template file (master list of strings)
                    strFile = FSO.BuildPath(strFolder, GetVBProjectForCurrentDB.Name & ".pot")
                Else
                    ' Translation work file
                    strFile = FSO.BuildPath(strFolder, strLanguage & ".po")
                End If
                WriteFile BuildFileContent(strLanguage), strFile
            End If
            .MoveNext
        Loop
        .Close
    End With

End Sub


'---------------------------------------------------------------------------------------
' Procedure : LoadTranslations
' Author    : Adam Waller
' Date      : 5/19/2021
' Purpose   : Load translation data from translation files
'---------------------------------------------------------------------------------------
'
Public Sub LoadTranslations()

    ' Load English translation
    

End Sub


'---------------------------------------------------------------------------------------
' Procedure : TranslationsPath
' Author    : Adam Waller
' Date      : 5/19/2021
' Purpose   : Translation path saved in registry
'---------------------------------------------------------------------------------------
'
Public Property Get TranslationsPath() As String
    TranslationsPath = GetSetting(GetCodeVBProject.Name, "Language", "Translation Path", vbNullString)
End Property
Public Property Let TranslationsPath(strPath As String)
    SaveSetting GetCodeVBProject.Name, "Language", "Translation Path", strPath
End Property


'---------------------------------------------------------------------------------------
' Procedure : Contribute
' Author    : Adam Waller
' Date      : 5/19/2021
' Purpose   : Whether the user desires to contribute to translations
'---------------------------------------------------------------------------------------
'
Public Property Get Contribute() As Boolean
    Contribute = GetSetting(GetCodeVBProject.Name, "Language", "Contribute To Translations", False)
End Property
Public Property Let Contribute(blnContributeToTranslations As Boolean)
    SaveSetting GetCodeVBProject.Name, "Language", "Contribute To Translations", blnContributeToTranslations
End Property


'---------------------------------------------------------------------------------------
' Procedure : Language
' Author    : Adam Waller
' Date      : 5/19/2021
' Purpose   : Selected language
'---------------------------------------------------------------------------------------
'
Public Property Get Language() As String
    Language = GetSetting(GetCodeVBProject.Name, "Language", "Language", en_US)
End Property
Public Property Let Language(strLanguage As String)
    SaveSetting GetCodeVBProject.Name, "Language", "Language", strLanguage
End Property


'---------------------------------------------------------------------------------------
' Procedure : BuildKey
' Author    : Adam Waller
' Date      : 5/15/2021
' Purpose   : Build a dictionary key from the values, joined by pipe character
'---------------------------------------------------------------------------------------
'
Private Function BuildKey(ParamArray varParts()) As String
    BuildKey = Join(varParts, "|")
End Function


'---------------------------------------------------------------------------------------
' Procedure : SaveString
' Author    : Adam Waller
' Date      : 5/15/2021
' Purpose   : Save the string to the database table
'---------------------------------------------------------------------------------------
'
Private Sub SaveString(strText As String, strContext As String, strComments As String, ParamArray varParams() As Variant)
    
    Dim dbs As Database
    Dim rst As Recordset
    
    Set dbs = CodeDb
    Set rst = dbs.OpenRecordset("tblStrings")
    
    With rst
        .AddNew
            !msgid = strText
            !Context = ZN(strContext)
            !Comments = ZN(strComments)
        .Update
        .Close
    End With
    
End Sub



Private Sub LoadStrings()
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : SetLanguage
' Author    : Adam Waller
' Date      : 5/15/2021
' Purpose   : Set the current language
'---------------------------------------------------------------------------------------
'
Public Sub SetLanguage(strLanguage As String)
    LoadLanguage strLanguage
End Sub


'---------------------------------------------------------------------------------------
' Procedure : CheckInit
' Author    : Adam Waller
' Date      : 5/15/2021
' Purpose   : Ensure that the language strings have been loaded
'---------------------------------------------------------------------------------------
'
Private Sub CheckInit()
    If m_strCurrentLanguage = vbNullString Then
        Set dStrings = New Dictionary
        Set dTranslation = New Dictionary
        LoadLanguage GetCurrentLanguage
    End If
End Sub


'---------------------------------------------------------------------------------------
' Procedure : LoadLanguage
' Author    : Adam Waller
' Date      : 5/15/2021
' Purpose   : Loads the language entries into the dictionary objects.
'---------------------------------------------------------------------------------------
'
Private Sub LoadLanguage(strLanguage As String)
    
    Dim dbs As Database
    Dim rst As Recordset
    Dim strSql As String
        
    m_strCurrentLanguage = strLanguage
    Set dStrings = New Dictionary
    Set dTranslation = New Dictionary
    
    ' Load strings and translations
    Set dbs = CodeDb
    strSql = "select * from qryStrings where (Language = '" & strLanguage & "' or Language Is Null)"
    Set rst = dbs.OpenRecordset("qryStrings", dbOpenDynaset)
    With rst
        Do While Not .EOF
            If Not dStrings.Exists(!Key) Then dStrings.Add !Key, !ID
            If Not dEnglish.Exists(!ID) Then dEnglish.Add !ID, !msgid
            If Nz(!Translation) <> vbNullString Then dTranslation.Add !ID, !Translation
            .MoveNext
        Loop
        .Close
    End With
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : CurrentLanguage
' Author    : Adam Waller
' Date      : 5/15/2021
' Purpose   : Return the currently selected language, falling back to operating system
'           : UI language, then to US English.
'---------------------------------------------------------------------------------------
'
Public Function GetCurrentLanguage() As String
    GetCurrentLanguage = Coalesce(m_strCurrentLanguage, GetSavedLanguage, GetOsLanguage, en_US)
End Function


Private Function GetSavedLanguage() As String

End Function


Private Function GetOsLanguage() As String

End Function


'---------------------------------------------------------------------------------------
' Procedure : Export
' Author    : Adam Waller
' Date      : 5/15/2021
' Purpose   : Export current language to .po file for translation.
'---------------------------------------------------------------------------------------
'
Public Sub ExportTranslation(Optional strPath As String)

End Sub


'---------------------------------------------------------------------------------------
' Procedure : ImportTranslation
' Author    : Adam Waller
' Date      : 5/15/2021
' Purpose   : Import a translation file. (*.po/*.pot)
'---------------------------------------------------------------------------------------
'
Private Sub ImportTranslation(strFile As String)

    Dim strName As String
    Dim strLanguage As String
    Dim strContent As String
    Dim varLines As Variant
    Dim lngLine As Long
    Dim strLine As String
    Dim tStr As tEntry
    Dim tBlank As tEntry
    Dim cHeader As clsConcat
    Dim blnInHeader As Boolean
    Dim dbs As Database
    
    ' Read file contents
    strContent = ReadFile(strFile)
    If strContent = vbNullString Then Exit Sub
    
    ' Get language from file name, and remove existing entries
    Set dbs = CodeDb
    strName = FSO.GetFileName(strFile)
    If strName = GetVBProjectForCurrentDB.Name & ".pot" Then
        ' Template file (English strings)
        strLanguage = en_US
        dbs.Execute "delete from tblStrings", dbFailOnError
        ' Remove all translations, since we are resetting IDs
        dbs.Execute "delete from tblTranslation", dbFailOnError
    Else
        ' Other language
        strLanguage = FSO.GetBaseName(strName)
        dbs.Execute "delete from tblTranslation where LanguageID='" & strLanguage & "'", dbFailOnError
    End If
    
    ' Split into lines
    varLines = Split(strContent, vbCrLf)
    
    ' Prepare header class
    cHeader.AppendOnAdd = vbCrLf
    blnInHeader = True
    
    ' Loop through lines
    For lngLine = 0 To UBound(varLines)
        strLine = Trim(varLines(lngLine))
        
        ' Add header till we reach first entry
        If blnInHeader And strLine = vbNullString Then
            blnInHeader = False
            AddLanguage strLanguage, cHeader.GetStr
        End If
        If blnInHeader Then cHeader.Add strLine
        
        ' Watch for new section
        If strLine = vbNullString Then
            cHeader.Clear
            tStr = tBlank
        End If
        
        ' Look for headings
        If StartsWith(strLine, "msgctxt ") Then
            'tstr.strContext=
        
        End If
        
    Next lngLine
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : VerifyString
' Author    : Adam Waller
' Date      : 5/26/2021
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub VerifyString(strLanguage As String, tString As tEntry)

End Sub


'---------------------------------------------------------------------------------------
' Procedure : AddLanguage
' Author    : Adam Waller
' Date      : 5/26/2021
' Purpose   : Add a language entry with a header section.
'---------------------------------------------------------------------------------------
'
Private Sub AddLanguage(strLanguage As String, strHeader As String)
    
    Dim dbs As Database
    Dim rst As Recordset
    
    Set dbs = CodeDb
    
    ' Clear any existing record
    dbs.Execute "delete from tblLanguages where ID='" & strLanguage & "'", dbFailOnError
    
    ' Add new/replacement record
    Set rst = dbs.OpenRecordset("tblLanguages")
    With rst
        .AddNew
            !ID = strLanguage
            !DisplayName = GetLanguageName(strLanguage)
            !Header = strHeader
        .Update
        .Close
    End With
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : SaveTemplate
' Author    : Adam Waller
' Date      : 5/17/2021
' Purpose   : Save the translation template file (projectname.pot)
'---------------------------------------------------------------------------------------
'
Private Sub SaveTemplate()

End Sub


'---------------------------------------------------------------------------------------
' Procedure : BuildFileContent
' Author    : Adam Waller
' Date      : 5/19/2021
' Purpose   : Creates the .po/.pot file. (en_US will be treated as a template)
'---------------------------------------------------------------------------------------
'
Private Function BuildFileContent(strLanguage As String) As String

    Dim dbs As Database
    Dim rst As Recordset
    Dim strHeader As String
    
    With New clsConcat
        .AppendOnAdd = vbCrLf
    
        ' Add header section
        strHeader = Nz(DLookup("Header", "tblLanguages", "ID='" & strLanguage & "'"))
        .Add strHeader
        
        ' Load strings from database
        Set dbs = CodeDb
        Set rst = dbs.OpenRecordset( _
            "select * from qryStrings where language='" & strLanguage & "' or language is null", _
            dbOpenSnapshot)
    
        ' Loop through strings
        Do While Not rst.EOF
            .Add vbNullString ' (blank line)
            If Nz(rst!Comments) <> vbNullString Then
                ' Include additional comments for translators
                .Add "# TRANSLATORS: ", rst!Comments
            End If
            .Add "msgctxt ", Q(Nz(rst!Context))
            .Add "msgid ", Q(Nz(rst!msgid))
            .Add "msgstr ", Q(Nz(rst!Translation))
            rst.MoveNext
        Loop
        rst.Close
        
        ' Return assembled content
        BuildFileContent = .GetStr
    End With

End Function


'---------------------------------------------------------------------------------------
' Procedure : Q
' Author    : Adam Waller
' Date      : 5/19/2021
' Purpose   : Quotes the string, and escapes any embedded quotes. Also breaks long
'           : strings into multiple lines and replaces vbCrLf with \n.
'---------------------------------------------------------------------------------------
'
Private Function Q(strText As String) As String

    ' Maximum line length
    Const MAX_LEN As Integer = 70

    Dim strNew As String
    Dim intPos As Integer
    Dim intStart As Integer
    
    ' Replace newlines and quotes with placeholder
    strNew = Replace(strText, vbCrLf, "\n")
    strNew = Replace(strNew, """", "\""")
    
    ' Add line breaks for over 70 characters.
    ' (80 characters is standard for PO files)
    If Len(strNew) > 70 Then
        
        With New clsConcat
            
            ' Start with blank string
            .Add """"""
            
            ' Begin at first character
            intStart = 1
            
            ' Continue while
            Do While intStart < Len(strNew)
                intPos = MAX_LEN
                ' Walk backwards through the string, looking for spaces
                ' where we can break the line.
                For intPos = (intStart + MAX_LEN) To intStart Step -1
                    If Mid$(strNew, intPos, 1) = " " Then
                        ' Break here after space
                        intPos = intPos + 1
                        Exit For
                    End If
                Next intPos
                ' Use full max length if we don't find a space
                If intPos = intStart - 1 Then intPos = intStart + MAX_LEN
                ' Break string here, and move start
                .Add vbCrLf, """", Mid$(strNew, intStart, intPos - intStart), """"
                intStart = intPos
                ' Add final partial string
                If Len(strNew) - intStart < MAX_LEN Then
                    .Add vbCrLf, """", Mid$(strNew, intStart), """"
                    Exit Do
                End If
                
                ' for debugging
                DoEvents
            Loop
            
            ' Return multi-line result
            Q = .GetStr
        End With
    Else
        ' Return single line
        Q = """" & strNew & """"
    End If
    
End Function


Private Function UnQ(strLine As String, strHeading As String)

End Function


