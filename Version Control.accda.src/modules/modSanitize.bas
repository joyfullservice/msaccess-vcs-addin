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


'---------------------------------------------------------------------------------------
' Procedure : SanitizeFile
' Author    : Adam Waller
' Date      : 11/4/2020
' Purpose   : Rewritten version of sanitize function
'---------------------------------------------------------------------------------------
'
Public Sub SanitizeFile(strPath As String)

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
            strFile = ReadFile(strPath)
        End If
    End If
    Perf.OperationStart "Sanitize File"
    varLines = Split(strFile, vbCrLf)

    ' Initialize concatenation class to include line breaks
    ' after each line that we add when building new file text.
    curStart = Perf.MicroTimer
    Set cData = New clsConcat
    cData.AppendOnAdd = vbCrLf

    ' Using a do loop since we may adjust the line counter
    ' during a loop iteration.
    Do While lngLine <= UBound(varLines)
        
        ' Get unmodified and trimmed line
        strLine = varLines(lngLine)
        strTLine = Trim$(strLine)
        
        ' Improve performance by reducing comparisons
        If Len(strTLine) > 3 And blnInsideIgnoredBlock Then
            ' Ignore this line
        ElseIf Len(strTLine) > 60 And StartsWith(strTLine, "0x") Then
            ' Add binary data line. No need to test this line further.
            cData.Add strLine
        Else
            ' Run the rest of the tests
            Select Case strTLine
            
                ' File version
                Case "Version =21"
                    ' Change version down to 20 to allow import into Access 2010.
                    ' (Haven't seen any significant issues with this.)
                    cData.Add "Version =20"
                
                ' Print settings blocks to ignore
                Case "PrtMip = Begin", _
                    "PrtDevMode = Begin", _
                    "PrtDevModeW = Begin", _
                    "PrtDevNames = Begin", _
                    "PrtDevNamesW = Begin"
                    ' Set flag to ignore lines inside this block.
                    blnInsideIgnoredBlock = True
        
                ' Aggressive sanitize blocks
                Case "GUID = Begin", _
                    "NameMap = Begin", _
                    "dbLongBinary ""DOL"" = Begin", _
                    "dbBinary ""GUID"" = Begin"
                    If Options.AggressiveSanitize Then
                        blnInsideIgnoredBlock = True
                    Else
                        ' Include these sections
                        cData.Add strLine
                    End If
                    
                ' Single lines to ignore
                Case "NoSaveCTIWhenDisabled =1"
        
                ' Publish option (used in Queries)
                Case "dbByte ""PublishToWeb"" =""1""", _
                    "PublishOption =1"
                    If Not Options.StripPublishOption Then cData.Add strLine
                
                ' End of block section
                Case "End"
                    If blnInsideIgnoredBlock Then
                        ' Reached the end of the ignored block.
                        blnInsideIgnoredBlock = False
                    Else
                        ' End of included block
                        cData.Add strLine
                    End If
                
                ' See if this file is from a report object
                Case "Begin Report"
                    ' Turn flag on to ignore Right and Bottom lines
                    blnIsReport = True
                    cData.Add strLine
                
                ' Beginning of main section
                Case "Begin"
                    If blnIsPassThroughQuery And Options.AggressiveSanitize Then
                        ' Ignore remaining content. (See Issue #182)
                        Exit Do
                    Else
                        cData.Add strLine
                    End If
                
                Case Else
                    If blnInsideIgnoredBlock Then
                        ' Skip if we are in an ignored block
                    ElseIf StartsWith(strTLine, "Checksum =") Then
                        ' Ignore Checksum lines, since they will change.
                    ElseIf StartsWith(strTLine, "BaseInfo =") Then
                        ' BaseInfo is used with combo boxes, similar to RowSource.
                        ' Since the value could span multiple lines, we need to
                        ' check the indent level of the following lines to see how
                        ' many lines to skip.
                        intIndent = GetIndent(strLine)
                        ' Preview the next line, and check the indent level
                        Do While GetIndent(varLines(lngLine + 1)) > intIndent
                            ' Move
                            lngLine = lngLine + 1
                        Loop
                    ElseIf blnIsReport And StartsWith(strLine, "    Right =") Then
                        ' Ignore this line. (Not important, and frequently changes.)
                    ElseIf blnIsReport And StartsWith(strLine, "    Bottom =") Then
                        ' Turn flag back off now that we have ignored these two lines.
                        blnIsReport = False
                    Else
                        ' All other lines will be added.
                        cData.Add strLine
                        
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
    
    ' Remove last vbcrlf
    cData.Remove Len(vbCrLf)

    ' Log performance
    Perf.OperationEnd
    Log.Add "    Sanitized in " & Format$(Perf.MicroTimer - curStart, "0.000") & " seconds.", Options.ShowDebug
    
    ' Replace original file with sanitized version
    WriteFile cData.GetStr, strPath
    
    ' Log any errors
    CatchAny eelError, "Error sanitizing file " & FSO.GetFileName(strPath), ModuleName & ".SanitizeFile"
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : SanitizeXML
' Author    : Adam Waller
' Date      : 4/29/2021
' Purpose   : Remove non-essential data that changes every time the file is exported.
'---------------------------------------------------------------------------------------
'
Public Sub SanitizeXML(strPath As String)

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

    ' Show stats if debug turned on.
    Log.Add "    Sanitized in " & Format$(Perf.MicroTimer - curStart, "0.000") & " seconds.", Options.ShowDebug

    ' Log any errors
    CatchAny eelError, "Error sanitizing XML file " & FSO.GetFileName(strPath), ModuleName & ".SanitizeXML"

End Sub


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
    Dim objWriter As MXXMLWriter60
    
    Perf.OperationStart "Format XML"
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

    ' Return formatted output
    FormatXML = objWriter.output
    Perf.OperationEnd
    
End Function