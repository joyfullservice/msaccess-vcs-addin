Attribute VB_Name = "modTestExportFormat"
'---------------------------------------------------------------------------------------
' Module    : modTestExportFormat
' Author    : Adam Waller
' Date      : 7/31/2026
' Purpose   : Guard the export format version list against drift. VBA cannot enumerate
'           : the members of an Enum at runtime, so GetExportFormatVersions repeats the
'           : members of eExportFormatVersion by hand. A missing entry there is invisible
'           : in normal use: the new format gates work everywhere in code, but the format
'           : never appears in the options combo. These tests parse the enum out of the
'           : add-in's own source and fail when the two disagree.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
Option Private Module
'@Folder("Tests.Infrastructure")

Private Const ENUM_NAME As String = "eExportFormatVersion"
Private Const MEMBER_PREFIX As String = "EFV_"
Private Const SOURCE_MODULE As String = "modConstants"


'---------------------------------------------------------------------------------------
' Procedure : TestExportFormatListMatchesEnum
' Author    : Adam Waller
' Date      : 7/31/2026
' Purpose   : Every enum member is selectable in the options form, and the list names
'           : no version the enum does not define.
'---------------------------------------------------------------------------------------
'
Public Sub TestExportFormatListMatchesEnum()

    Dim dMembers As Dictionary
    Dim dList As Dictionary
    Dim varFormat As Variant
    Dim varName As Variant

    Set dMembers = GetEnumMembers
    If Not SourceWasRead(dMembers) Then Exit Sub

    TestAssert dMembers.Count > 0, "parsed at least one " & ENUM_NAME & " member"

    Set dList = New Dictionary
    For Each varFormat In GetExportFormatVersions
        dList(CLng(varFormat)) = True
    Next varFormat

    For Each varName In dMembers.Keys
        TestAssert dList.Exists(dMembers(varName)), _
            varName & " is missing from GetExportFormatVersions"
    Next varName

    TestAssert dList.Count = dMembers.Count, _
        "GetExportFormatVersions has " & dList.Count & " entries for " & _
        dMembers.Count & " enum members"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestExportFormatListIsAscending
' Author    : Adam Waller
' Date      : 7/31/2026
' Purpose   : LatestExportFormat returns the last entry, so the list has to be sorted.
'---------------------------------------------------------------------------------------
'
Public Sub TestExportFormatListIsAscending()

    Dim varFormat As Variant
    Dim lngPrior As Long
    Dim lngLast As Long

    For Each varFormat In GetExportFormatVersions
        TestAssert CLng(varFormat) > lngPrior, _
            ExportFormatToVersion(CLng(varFormat)) & " sorts after the entry before it"
        lngPrior = CLng(varFormat)
        lngLast = lngPrior
    Next varFormat

    TestAssert LatestExportFormat = lngLast, "LatestExportFormat returns the newest entry"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestExportFormatNamesMatchValues
' Author    : Adam Waller
' Date      : 7/31/2026
' Purpose   : EFV_5_1_0 must be 50100. Catches a copied line where the name was updated
'           : but the packed value was not (or the reverse).
'---------------------------------------------------------------------------------------
'
Public Sub TestExportFormatNamesMatchValues()

    Dim dMembers As Dictionary
    Dim varName As Variant
    Dim strVersion As String

    Set dMembers = GetEnumMembers
    If Not SourceWasRead(dMembers) Then Exit Sub

    For Each varName In dMembers.Keys
        strVersion = Replace(Mid$(CStr(varName), Len(MEMBER_PREFIX) + 1), "_", ".")
        TestAssert VersionToExportFormat(strVersion) = dMembers(varName), _
            varName & " should be " & VersionToExportFormat(strVersion) & _
            ", not " & dMembers(varName)
    Next varName

End Sub


'---------------------------------------------------------------------------------------
' Procedure : SourceWasRead
' Author    : Adam Waller
' Date      : 7/31/2026
' Purpose   : Report a readable failure when the add-in's own source is unavailable,
'           : rather than letting the test pass without checking anything.
'---------------------------------------------------------------------------------------
'
Private Function SourceWasRead(dMembers As Dictionary) As Boolean
    SourceWasRead = Not (dMembers Is Nothing)
    If Not SourceWasRead Then
        TestAssert False, "unable to read " & SOURCE_MODULE & " source " & _
            "(enable Trust access to the VBA project object model)"
    End If
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetEnumMembers
' Author    : Adam Waller
' Date      : 7/31/2026
' Purpose   : Parse the eExportFormatVersion block out of the running add-in's source,
'           : returning member name -> packed value. Returns Nothing when the code
'           : module cannot be read.
'---------------------------------------------------------------------------------------
'
Private Function GetEnumMembers() As Dictionary

    Dim cmp As VBIDE.VBComponent
    Dim varLines As Variant
    Dim lngLine As Long
    Dim lngPos As Long
    Dim strLine As String
    Dim strName As String
    Dim blnInEnum As Boolean
    Dim d As Dictionary

    LogUnhandledErrors
    On Error Resume Next
    Set cmp = GetCodeVBProject.VBComponents(SOURCE_MODULE)
    If cmp Is Nothing Then Exit Function
    varLines = Split(cmp.CodeModule.Lines(1, cmp.CodeModule.CountOfLines), vbCrLf)
    On Error GoTo 0
    If Not IsArray(varLines) Then Exit Function

    Set d = New Dictionary
    d.CompareMode = TextCompare

    For lngLine = 0 To UBound(varLines)
        strLine = Trim$(varLines(lngLine))
        If Not blnInEnum Then
            blnInEnum = (strLine Like "*Enum " & ENUM_NAME)
        ElseIf strLine Like "End Enum*" Then
            Exit For
        Else
            ' Drop any trailing comment, then split on the assignment
            lngPos = InStr(strLine, "'")
            If lngPos > 0 Then strLine = Trim$(Left$(strLine, lngPos - 1))
            lngPos = InStr(strLine, "=")
            If lngPos > 0 Then
                strName = Trim$(Left$(strLine, lngPos - 1))
                If Left$(strName, Len(MEMBER_PREFIX)) = MEMBER_PREFIX Then
                    d.Add strName, CLng(Trim$(Mid$(strLine, lngPos + 1)))
                End If
            End If
        End If
    Next lngLine

    Set GetEnumMembers = d

End Function
