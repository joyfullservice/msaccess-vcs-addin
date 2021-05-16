'---------------------------------------------------------------------------------------
' Module    : modTranslation
' Author    : Adam Waller
' Date      : 5/15/2021
' Purpose   : Used for the translation of strings to different languages, similar to
'           : the gettext toolset.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit


' Cache strings to dictionary objects so we don't have to do database lookups
' each time we need to return translated strings
Private dStrings As Dictionary
Private dTranslation As Dictionary
Private m_strCurrentLanguage As String


'---------------------------------------------------------------------------------------
' Procedure : T
' Author    : Adam Waller
' Date      : 5/15/2021
' Purpose   : Return the translated version of the string.
'---------------------------------------------------------------------------------------
'
Public Function T(strText As String, Optional strContext As String) As String

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
        SaveString strText, strContext
    End If
    
    ' Return translated string
    T = Coalesce(strNew, strText)

End Function


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
Private Sub SaveString(strText As String, strContext As String, ParamArray varParams() As Variant)
    
    Dim dbs As Database
    Dim rst As Recordset
    
    Set dbs = CodeDb
    Set rst = dbs.OpenRecordset("tblStrings")
    
    With rst
        .AddNew
            !msgid = Left$(strText, 255)
            If Len(strText) > 255 Then !FullString = strText
            !Context = strContext
            !AddDate = Now()
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
        
    m_strCurrentLanguage = strLanguage
    Set dStrings = New Dictionary
    Set dTranslation = New Dictionary
    
    ' Load strings and translations
    Set dbs = CodeDb
    Set rst = dbs.OpenRecordset("qryStrings", dbOpenDynaset)
    With rst
        Do While Not .EOF
            If Not dStrings.Exists(!Key) Then dStrings.Add !Key, !ID
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
    GetCurrentLanguage = Coalesce(m_strCurrentLanguage, GetSavedLanguage, GetOsLanguage, "en-US")
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
' Purpose   : Import a translation file. (*.po)
'---------------------------------------------------------------------------------------
'
Private Sub ImportTranslation(strFile As String)

End Sub