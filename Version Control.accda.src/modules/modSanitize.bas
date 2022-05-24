Attribute VB_Name = "modSanitize"
'---------------------------------------------------------------------------------------
' Module    : modSanitize
' Author    : Adam Waller
' Date      : 12/4/2020
' Purpose   : Functions to sanitize files to remove non-essential metadata
'---------------------------------------------------------------------------------------
Option Compare Database
Option Private Module
Option Explicit

Private Const ModuleName = "modSanitize"

' Array of lines to skip
Private m_SkipLines() As Long
Private m_lngSkipIndex As Long
Private m_colBlocks As Collection


'---------------------------------------------------------------------------------------
' Procedure : SanitizeFile
' Author    : Adam Waller
' Date      : 11/4/2020
' Purpose   : Rewritten version of sanitize function. Returns hash of content as well
'           : as saving to the specified path.
'---------------------------------------------------------------------------------------
'
Public Function SanitizeFile(strPath As String, blnReturnHash As Boolean) As String

    Dim strFile As String
    Dim varLines As Variant
    Dim lngLine As Long
    Dim cData As clsConcat
    Dim strLine As String
    Dim strTLine As String
    Dim blnInsideIgnoredBlock As Boolean
    Dim intIndent As Integer
    Dim blnIsReport As Boolean
    Dim blnIsPassThroughQuery As Boolean
    Dim curStart As Currency
    Dim strTempFile As String
    Dim strContent As String
    
    If DebugMode(True) Then On Error GoTo 0 Else On Error Resume Next

    ' Read text from file, and split into lines
    If HasUcs2Bom(strPath) Then
        strFile = ReadFile(strPath, "Unicode")
    Else
        ' ADP projects may contain mixed Unicode content
        If CurrentProject.ProjectType = acADP Then
            strTempFile = GetTempFile
            ConvertUcs2Utf8 strPath, strTempFile, False
            strFile = ReadFile(strTempFile)
            DeleteFile strTempFile
        Else
            If DbVersion <= 4 Then
                ' Access 2000 format exports using system codepage
                ' See issue #217
                strFile = ReadFile(strPath, GetSystemEncoding)
            Else
                ' Newer versions export as UTF-8
                strFile = ReadFile(strPath)
            End If
        End If
    End If
    
    Perf.OperationStart "Sanitize File"
    varLines = Split(strFile, vbCrLf)
    
    If Options.SanitizeLevel = eslNone Then GoTo Build_Output

    ' Set up index of lines to skip
    ReDim m_SkipLines(0 To UBound(varLines)) As Long
    m_lngSkipIndex = 0
    Set m_colBlocks = New Collection

    ' Initialize concatenation class to include line breaks
    ' after each line that we add when building new file text.
    curStart = Perf.MicroTimer

    ' Using a do loop since we may adjust the line counter
    ' during a loop iteration.
    Do While lngLine <= UBound(varLines)
        
        ' Get unmodified and trimmed line
        strLine = varLines(lngLine)
        strTLine = Trim$(strLine)
        
        ' Improve performance by reducing comparisons
        If Len(strTLine) > 3 And blnInsideIgnoredBlock Then
            SkipLine lngLine
        ElseIf Len(strTLine) > 60 And StartsWith(strTLine, "0x") Then
            ' Add binary data line. No need to test this line further.
        Else
            ' Run the rest of the tests
            Select Case strTLine
            
                ' File version
                Case "Version =21"
                    ' Change version down to 20 to allow import into Access 2010.
                    ' (Haven't seen any significant issues with this.)
                    varLines(lngLine) = "Version =20"
                
                ' Print settings blocks to ignore
                Case "PrtMip = Begin", _
                    "PrtDevMode = Begin", _
                    "PrtDevModeW = Begin", _
                    "PrtDevNames = Begin", _
                    "PrtDevNamesW = Begin"
                    ' Set flag to ignore lines inside this block.
                    blnInsideIgnoredBlock = True
                    SkipLine lngLine
        
                ' Aggressive sanitize blocks
                Case "GUID = Begin", _
                    "NameMap = Begin", _
                    "dbLongBinary ""DOL"" = Begin", _
                    "dbBinary ""GUID"" = Begin"
                    If Options.AggressiveSanitize Then
                        blnInsideIgnoredBlock = True
                        SkipLine lngLine
                    End If
                    
                ' Single lines to ignore (#249)
                Case "NoSaveCTIWhenDisabled =1", _
                    "AllowPivotTableView =0", _
                    "AllowPivotChartView =0"
                    SkipLine lngLine
        
                ' Publish option (used in Queries)
                Case "dbByte ""PublishToWeb"" =""1""", _
                    "PublishOption =1"
                    If Options.StripPublishOption Then SkipLine lngLine
                
                ' End of block section
                Case "End"
                    If blnInsideIgnoredBlock Then
                        ' Reached the end of the ignored block.
                        blnInsideIgnoredBlock = False
                        SkipLine lngLine
                    Else
                        ' Check for theme color index
                        CloseBlock
                    End If
                
                ' See if this file is from a report object
                Case "Begin Report"
                    ' Turn flag on to ignore Right and Bottom lines
                    blnIsReport = True
                    BeginBlock
                
                ' Beginning of main section
                Case "Begin"
                    BeginBlock
                    If blnIsPassThroughQuery And Options.AggressiveSanitize Then
                        ' Ignore remaining content. (See Issue #182)
                        Do While lngLine < UBound(varLines)
                            SkipLine lngLine
                            lngLine = lngLine + 1
                        Loop
                        Exit Do
                    End If
                
                ' Code section behind form or report object
                Case "CodeBehindForm"
                    ' Keep everything from this point on
                    Exit Do
                
                Case Else
                    If blnInsideIgnoredBlock Then
                        ' Skip content inside ignored blocks.
                        SkipLine lngLine
                    ElseIf StartsWith(strTLine, "Checksum =") Then
                        ' Ignore Checksum lines, since they will change.
                        SkipLine lngLine
                    ElseIf StartsWith(strTLine, "BaseInfo =") Then
                        ' BaseInfo is used with combo boxes, similar to RowSource.
                        ' Since the value could span multiple lines, we need to
                        ' check the indent level of the following lines to see how
                        ' many lines to skip.
                        SkipLine lngLine
                        intIndent = GetIndent(strLine)
                        ' Preview the next line, and check the indent level
                        Do While GetIndent(varLines(lngLine + 1)) > intIndent
                            ' Skip previewed line and move to next line
                            SkipLine lngLine + 1
                            lngLine = lngLine + 1
                        Loop
                    ElseIf blnIsReport And StartsWith(strLine, "    Right =") Then
                        ' Ignore this line. (Not important, and frequently changes.)
                        SkipLine lngLine
                    ElseIf blnIsReport And StartsWith(strLine, "    Bottom =") Then
                        ' Turn flag back off now that we have ignored these two lines.
                        SkipLine lngLine
                        blnIsReport = False
                    ElseIf StartsWith(strTLine, "Begin ") Then
                        ' Include block type name for controls
                        BeginBlock Mid$(strTLine, 7)
                    ElseIf EndsWith(strTLine, " = Begin") Then
                        BeginBlock
                    Else
                        ' All other lines will be added.
                        
                        ' Check for color properties
                        If InStr(1, strTLine, " =") > 1 Then CheckColorProperties strTLine, lngLine
                        
                        ' Check for pass-through query connection string
                        If StartsWith(strLine, "dbMemo ""Connect"" =""") Then
                            blnIsPassThroughQuery = True
                        End If
                    End If
            
            End Select
        End If
    
        ' Increment counter to next line
        lngLine = lngLine + 1
    Loop
    
    ' Ensure that we correctly processed the nested block sequence.
    If m_colBlocks.Count > 0 Then Log.Error eelWarning, Replace(Replace( _
        "Found ${BlockCount} unclosed blocks after sanitizing ${File}.", _
        "${BlockCount}", m_colBlocks.Count), _
        "${File}", strPath), ModuleName & ".SanitizeFile"
    
Build_Output:
    ' Build the final output
    strContent = BuildOutput(varLines)
    WriteFile strContent, strPath
    
    ' Return hash of content
    If blnReturnHash Then SanitizeFile = GetStringHash(strContent, True)

    ' Log performance
    Set m_colBlocks = Nothing
    Perf.OperationEnd
    Log.Add "    Sanitized in " & Format$(Perf.MicroTimer - curStart, "0.000") & " seconds.", Options.ShowDebug
    
    ' Log any errors
    CatchAny eelError, "Error sanitizing file " & FSO.GetFileName(strPath), ModuleName & ".SanitizeFile"
    
End Function


'---------------------------------------------------------------------------------------
' Procedure : BuildOutput
' Author    : Adam Waller
' Date      : 6/4/2021
' Purpose   : Splitting this out into its own sub to reduce complexity.
'---------------------------------------------------------------------------------------
'
Private Function BuildOutput(varLines As Variant) As String

    Dim cData As clsConcat
    Dim lngSkip As Long
    Dim lngLine As Long
    
    ' Check index of skipped lines
    If m_lngSkipIndex = 0 Then
        ' No lines to skip
        ReDim m_SkipLines(0 To 0)
        m_SkipLines(0) = UBound(varLines) + 1
    Else
        ' Trim and sort index array
        ReDim Preserve m_SkipLines(0 To m_lngSkipIndex - 1)
        QuickSort m_SkipLines
    End If
    
    ' Use concatenation class to maximize performance
    Set cData = New clsConcat
    With cData
        .AppendOnAdd = vbCrLf
        
        ' Loop through array of lines in source file
        For lngLine = 0 To UBound(varLines)
            
            ' Iterate the sorted skipped lines index to keep up with main loop
            ' (Using parallel loops to optimize performance)
            If m_SkipLines(lngSkip) < lngLine Then
                If lngSkip < UBound(m_SkipLines) Then lngSkip = lngSkip + 1
            End If
            
            ' Add content, unless the line is flagged to skip
            If m_SkipLines(lngSkip) <> lngLine Then .Add CStr(varLines(lngLine))
        
        Next lngLine
        
        ' Remove last vbcrlf
        cData.Remove Len(vbCrLf)
    
        ' Return assembled output
        BuildOutput = .GetStr
    End With

End Function


'---------------------------------------------------------------------------------------
' Procedure : SkipLine
' Author    : Adam Waller
' Date      : 6/4/2021
' Purpose   : Skip this line in the final output file
'---------------------------------------------------------------------------------------
'
Private Function SkipLine(lngLine As Long)
    m_SkipLines(m_lngSkipIndex) = lngLine
    m_lngSkipIndex = m_lngSkipIndex + 1
End Function


'---------------------------------------------------------------------------------------
' Procedure : BeginBlock
' Author    : Adam Waller
' Date      : 6/4/2021
' Purpose   : Add a dictionary object to represent the block
'---------------------------------------------------------------------------------------
'
Private Sub BeginBlock(Optional strType As String)
    Dim dBlock As Dictionary
    If m_colBlocks Is Nothing Then Set m_colBlocks = New Collection
    Set dBlock = New Dictionary
    If strType <> vbNullString Then dBlock.Add "Type", strType
    m_colBlocks.Add dBlock
End Sub


'---------------------------------------------------------------------------------------
' Procedure : CloseBlock
' Author    : Adam Waller
' Date      : 6/4/2021
' Purpose   : Determine if the block used any theme-based dynamic colors that should
'           : be skipped in the output file. (See issue #183)
'---------------------------------------------------------------------------------------
'
Private Sub CloseBlock()
    
    Dim varBase As Variant
    Dim intCnt As Integer
    Dim dBlock As Dictionary
    Dim strKey As String
        
    ' Skip if we are not using aggressive color sanitize
    If Options.SanitizeColors <= eslNone Then Exit Sub
    
    ' Bail out if we don't have a block to review
    If m_colBlocks.Count = 0 Then Exit Sub
    Set dBlock = m_colBlocks(m_colBlocks.Count)
    
    ' Skip if we are not using themes for this control (UseTheme=0)
    ' (Applies to "CommandButton", "Tab", "ToggleButton")
    If dBlock.Exists("UseTheme") Then Exit Sub
    
    ' Build array of base properties
    varBase = Array("Back", "AlternateBack", "Border", _
            "Fore", "Gridline", "HoverFore", _
            "Hover", "PressedFore", "Pressed", _
            "DatasheetFore", "DatasheetBack", "DatasheetGridlines")
    
    ' Loop through properties, checking for index
    For intCnt = 0 To UBound(varBase)
        strKey = varBase(intCnt) & "ThemeColorIndex"
        If dBlock.Exists(strKey) Then
            If dBlock(strKey) <> NO_THEME_INDEX Then
                ' Check for corresponding color property
                strKey = varBase(intCnt) & "Color"
                If dBlock.Exists(strKey) Then
                    ' Skip the dynamic color line
                    SkipLine dBlock(strKey)
                End If
            End If
        Else
            Select Case dBlock("Type")
                Case "Section", "FormHeader", "FormFooter"
                    ' Some controls like form sections don't use color values
                    ' if a theme index is specified. If a color value exists,
                    ' we should preserve it.
                Case Else
                    ' Most controls automatically use theme indexes
                    ' unless otherwise specified.
                    ' As discussed in #183, this can be affected by incomplete
                    ' component definition blocks.
                    If Options.SanitizeColors = eslAdvancedBeta Then
                        strKey = varBase(intCnt) & "Color"
                        If dBlock.Exists(strKey) Then
                            ' Skip the dynamic color line
                            SkipLine dBlock(strKey)
                        End If
                    End If
            End Select
        End If
    Next intCnt
    
    ' Remove this block
    m_colBlocks.Remove m_colBlocks.Count
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : CheckColorProperties
' Author    : Adam Waller
' Date      : 6/4/2021
' Purpose   : Use an index to reference color properties so we can determine any lines
'           : that we need to discard after finishing the block.
'---------------------------------------------------------------------------------------
'
Private Sub CheckColorProperties(strTLine As String, lngLine As Long)

    Dim dBlock As Dictionary
    Dim varParts As Variant
    Dim lngCnt As Long
    Dim lngID As Long
    Dim strID As String
    Dim lngValue As Long
    Dim lngColor As Long
    
    ' Skip if not using this option
    If Options.SanitizeColors <= eslNone Then Exit Sub
    
    ' Exit if we are not inside a block
    If Not m_colBlocks Is Nothing Then lngCnt = m_colBlocks.Count
    If lngCnt = 0 Then Exit Sub
    Set dBlock = m_colBlocks(m_colBlocks.Count)
    
    ' Split on property/value
    varParts = Split(strTLine, " =")
    Select Case varParts(0)
        
        ' Theme color index properties
        Case "BackThemeColorIndex", "AlternateBackThemeColorIndex", "BorderThemeColorIndex", _
            "ForeThemeColorIndex", "GridlineThemeColorIndex", "HoverForeThemeColorIndex", _
            "HoverThemeColorIndex", "PressedForeThemeColorIndex", "PressedThemeColorIndex", _
            "DatasheetBackThemeColorIndex", "DatasheetForeThemeColorIndex", "DatasheetGridlinesThemeColorIndex"
            ' Save to dictionary if using a theme index color
            dBlock.Add varParts(0), varParts(1)
    
        ' Matching color properties
        Case "BackColor", "AlternateBackColor", "BorderColor", _
            "ForeColor", "GridlineColor", "HoverForeColor", _
            "HoverColor", "PressedForeColor", "PressedColor", _
            "DatasheetBackColor", "DatasheetForeColor", "DatasheetGridlinesColor"
            
            ' Check for system color constants
            If IsNumeric(varParts(1)) Then lngColor = varParts(1)
            If lngColor < 0 Then
                ' Using a system color constant or other Access constant value.
                ' https://stackoverflow.com/a/30396550/4121863
                ' Leave this color value intact.
            Else
                ' Save line of color property
                dBlock.Add varParts(0), lngLine
            End If
        
        Case "UseTheme"
            ' You can tell certain controls to not use the theme. (Buttons, Tabs, Toggles)
            If varParts(1) = 0 Then dBlock.Add varParts(0), 0
        
        Case Else
            ' Check for other related dynamic color properties/indexes
            If StartsWith(strTLine, "DatasheetGridlinesColor") Then
                ' May include the index number in the property name. (I.e. DatasheetGridlinesColor12 =0)
                ' Convert to a more consistent identifier, using the index suffix as the value.
                dBlock.Add "DatasheetGridlinesThemeColorIndex", Mid$(varParts(0), 24)
            End If
    
    End Select
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : SanitizeXML
' Author    : Adam Waller
' Date      : 4/29/2021
' Purpose   : Remove non-essential data that changes every time the file is exported.
'           : Optionally returns a hash of the file content. (To save reading the file
'           : back again afterwards to compute the hash.)
'---------------------------------------------------------------------------------------
'
Public Function SanitizeXML(strPath As String, blnReturnHash As Boolean) As String

    Dim curStart As Currency
    Dim cData As clsConcat
    Dim strFile As String
    Dim strText As String
    Dim strTLine As String
    Dim strLine As String
    Dim lngLine As Long
    Dim rxLine As VBScript_RegExp_55.RegExp
    Dim objMatches As VBScript_RegExp_55.MatchCollection
    Dim varLines As Variant
    
    If DebugMode(True) Then On Error GoTo 0 Else On Error Resume Next
    
    Set cData = New clsConcat
    cData.AppendOnAdd = vbCrLf
    Set rxLine = New VBScript_RegExp_55.RegExp

    ' Read text from file
    strFile = ReadFile(strPath)
    Perf.OperationStart "Sanitize XML"
    curStart = Perf.MicroTimer
    
    ' Exporting Table Def as XML does not properly encode ampersand character (See #314)
    ' Most likely if any ampersands are encoded correctly, all of them will be.
    With New VBScript_RegExp_55.RegExp
        .Multiline = True
        .Global = True
        ' Match &amp; &quot; &gt; &lt; etc...
        .Pattern = "&[A-z]{2,6};"
        If Not .Test(strFile) Then
            ' Properly encode any embedded ampersand characters to make valid XML
            strFile = Replace(strFile, "&", "&amp;")
        End If
    End With
    
    ' Split into array of lines
    varLines = Split(FormatXML(strFile), vbCrLf)

    ' Using a do loop since we may adjust the line counter
    ' during a loop iteration.
    Do While lngLine <= UBound(varLines)
    
        ' Get unmodified and trimmed line
        strLine = varLines(lngLine)
        strTLine = TrimTabs(Trim$(strLine))
        
        ' Look for specific lines
        Select Case True
            
            ' Discard blank lines
            Case strTLine = vbNullString
            
            ' Remove generated timestamp in header
            Case StartsWith(strTLine, "<dataroot ")
                '<dataroot xmlns:od="urn:schemas-microsoft-com:officedata" generated="2020-04-27T10:28:32">
                '<dataroot generated="2021-04-29T17:27:33" xmlns:od="urn:schemas-microsoft-com:officedata">
                With rxLine
                    .Pattern = "( generated="".+?"")"
                    If .Test(strLine) Then
                        ' Replace timestamp with empty string.
                        Set objMatches = .Execute(strLine)
                        strText = Replace(strLine, objMatches(0).SubMatches(0), vbNullString, , 1)
                        cData.Add strText
                    Else
                        ' Did not contain a timestamp. Keep the whole line
                        cData.Add strLine
                    End If
                End With
            
            ' Remove non-critical single lines that are not consistent between systems
            'Case StartsWith(strTLine, "<od:tableProperty name=""NameMap""")
            '    If Not Options.AggressiveSanitize Then cData.Add strLine
                
            ' Remove multi-line sections
            Case StartsWith(strTLine, "<od:tableProperty name=""NameMap"""), _
                StartsWith(strTLine, "<od:tableProperty name=""GUID"""), _
                StartsWith(strTLine, "<od:fieldProperty name=""GUID""")
                If Options.AggressiveSanitize Then
                    Do While Not EndsWith(strTLine, "/>")
                        lngLine = lngLine + 1
                        strTLine = TrimTabs(Trim$(varLines(lngLine)))
                    Loop
                Else
                    ' Keep line and continue
                    cData.Add strLine
                End If
            
            ' Publish to web sections
            Case StartsWith(strTLine, "<od:tableProperty name=""PublishToWeb""")
                If Not Options.StripPublishOption Then cData.Add strLine
            
            ' Keep everything else
            Case Else
                cData.Add strLine
            
        End Select
        
        ' Move to next line
        lngLine = lngLine + 1
    Loop
    
    Perf.OperationEnd
    
    ' Write out sanitized XML file
    WriteFile cData.GetStr, strPath

    ' Return hash, if requested
    If blnReturnHash Then SanitizeXML = GetStringHash(cData.GetStr, True)
    
    ' Show stats if debug turned on.
    Log.Add "    Sanitized in " & Format$(Perf.MicroTimer - curStart, "0.000") & " seconds.", Options.ShowDebug

    ' Log any errors
    CatchAny eelError, "Error sanitizing XML file " & FSO.GetFileName(strPath), ModuleName & ".SanitizeXML"

End Function


'---------------------------------------------------------------------------------------
' Procedure : TrimTabs
' Author    : Adam Waller
' Date      : 4/29/2021
' Purpose   : Trim off tabs from beginning and end of string
'---------------------------------------------------------------------------------------
'
Public Function TrimTabs(strText As String) As String

    Dim dblStart As Double
    Dim dblEnd As Double
    Dim dblPos As Double
    
    ' Look for leading tabs
    dblStart = 1
    For dblPos = 1 To Len(strText)
        If Mid$(strText, dblPos, 1) <> vbTab Then
            dblStart = dblPos
            Exit For
        End If
    Next dblPos
    
    ' Look for trailing tabs
    dblEnd = 1
    If Right$(strText, 1) = vbTab Then
        For dblPos = Len(strText) To 1 Step -1
            If Mid$(strText, dblPos, 1) <> vbTab Then
                dblEnd = dblPos + 1
                Exit For
            End If
        Next dblPos
    Else
        ' No trailing tabs
        dblEnd = Len(strText) + 1
    End If
    
    ' Return string
    TrimTabs = Mid$(strText, dblStart, dblEnd - dblStart)
    
End Function


'---------------------------------------------------------------------------------------
' Procedure : StartsWith
' Author    : Adam Waller
' Date      : 11/5/2020
' Purpose   : See if a string begins with a specified string.
'---------------------------------------------------------------------------------------
'
Public Function StartsWith(strText As String, strStartsWith As String, Optional Compare As VbCompareMethod = vbBinaryCompare) As Boolean
    StartsWith = (InStr(1, strText, strStartsWith, Compare) = 1)
End Function


'---------------------------------------------------------------------------------------
' Procedure : EndsWith
' Author    : Adam Waller
' Date      : 4/29/2021
' Purpose   : See if a string ends with a specified string.
'---------------------------------------------------------------------------------------
'
Public Function EndsWith(strText As String, strEndsWith As String, Optional Compare As VbCompareMethod = vbBinaryCompare) As Boolean
    EndsWith = (StrComp(Right$(strText, Len(strEndsWith)), strEndsWith, Compare) = 0)
    'EndsWith = (InStr(1, strText, strEndsWith, Compare) = len(strtext len(strendswith) 1)
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetIndent
' Author    : Adam Waller
' Date      : 11/5/2020
' Purpose   : Returns the number of spaces until the first non-space character.
'---------------------------------------------------------------------------------------
'
Public Function GetIndent(strLine As Variant) As Integer
    Dim strChar As String
    strChar = Left$(Trim(strLine), 1)
    If strLine <> vbNullString Then GetIndent = InStr(1, strLine, strChar) - 1
End Function


'---------------------------------------------------------------------------------------
' Procedure : FormatXML
' Author    : Adam Waller
' Date      : 4/22/2021
' Purpose   : Format XML content for consistent and readable output.
'---------------------------------------------------------------------------------------
'
Private Function FormatXML(strSourceXML As String, _
    Optional blnOmitDeclaration As Boolean) As String

    Dim objReader As SAXXMLReader60
    Dim objWriter As MXHTMLWriter60
    Dim strOutput As String

    ' Skip processing if no content to format
    If strSourceXML = vbNullString Then Exit Function
    
    Perf.OperationStart "Format XML"
    
    ' Trap any errors with parsing or formatting the XML
    If DebugMode(True) Then On Error GoTo 0 Else On Error Resume Next
    
    Set objWriter = New MXHTMLWriter60
    Set objReader = New SAXXMLReader60
    
    ' Set up writer
    With objWriter
        .indent = True
        .omitXMLDeclaration = Not blnOmitDeclaration
        Set objReader.contentHandler = objWriter
    End With
    
    ' Prepare reader
    With objReader
        Set .contentHandler = objWriter
        .parse strSourceXML
    End With

    ' Apply custom indent to output
    strOutput = CustomIndent(objWriter.output)
    
    ' Check for any errors parsing the XML
    If CatchAny(eelError, "Error parsing XML content", ModuleName & ".FormatXML") Then
        ' Fall back to input XML
        strOutput = strSourceXML
        ' Output XML to log file
        Log.Spacer False
        Log.Add strSourceXML, False
        Log.Spacer False
    End If
    
    ' Return formatted output
    Perf.OperationEnd
    FormatXML = strOutput
    
End Function


'---------------------------------------------------------------------------------------
' Procedure : CustomIndent
' Author    : Adam Waller
' Date      : 2/16/2022
' Purpose   : Convert tabbed indents to two spaces for more compact display
'---------------------------------------------------------------------------------------
'
Private Function CustomIndent(strText As String, Optional intSpaces As Integer = 2) As String

    Dim lngLine As Long
    Dim varLines As Variant
    Dim strLine As String
    Dim lngPos As Long
    
    ' Split content into lines
    varLines = Split(strText, vbCrLf)
    
    ' Rebuild while converting tabs to
    With New clsConcat
        .AppendOnAdd = vbCrLf
        
        ' Loop through lines
        For lngLine = 0 To UBound(varLines)
            strLine = varLines(lngLine)
            For lngPos = 1 To Len(strLine)
                If Mid$(strLine, lngPos, 1) <> vbTab Then
                    ' Replace any leading tabs with space indent
                    .Add Space$(intSpaces * (lngPos - 1)), Mid$(strLine, lngPos)
                    Exit For
                End If
            Next lngPos
        Next lngLine
        
        ' Return result after trimming off last return
        If lngLine > 0 Then .Remove 2
        CustomIndent = .GetStr
    End With
    
End Function

