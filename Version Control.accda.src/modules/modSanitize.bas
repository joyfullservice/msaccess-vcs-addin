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
    Dim sngStartTime As Single
    Dim strTempFile As String
    
    If DebugMode Then On Error GoTo 0 Else On Error Resume Next

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
    
    ' Delete original file now so we can write it immediately
    ' when the new data has been constructed.
    DeleteFile strPath

    ' Initialize concatenation class to include line breaks
    ' after each line that we add when building new file text.
    sngStartTime = Timer
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
    Log.Add "    Sanitized in " & Format$(Timer - sngStartTime, "0.00") & " seconds.", Options.ShowDebug
    
    ' Replace original file with sanitized version
    WriteFile cData.GetStr, strPath
    
    ' Log any errors
    CatchAny eelError, "Error sanitizing file " & FSO.GetFileName(strPath), ModuleName & ".SanitizeFile"
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : SanitizeXML
' Author    : Adam Waller
' Date      : 4/27/2020
' Purpose   : Remove non-essential data that changes every time the file is exported.
'---------------------------------------------------------------------------------------
'
Public Sub SanitizeXML(strPath As String, Options As clsOptions)

    Dim sngOverall As Single
    Dim sngTimer As Single
    Dim cData As clsConcat
    Dim strText As String
    Dim rxLine As VBScript_RegExp_55.RegExp
    Dim objMatches As VBScript_RegExp_55.MatchCollection
    Dim stmInFile As ADODB.Stream
    Dim blnFound As Boolean
    
    If DebugMode Then On Error GoTo 0 Else On Error Resume Next
    
    Set cData = New clsConcat
    Set rxLine = New VBScript_RegExp_55.RegExp
    
    ' Timers to monitor performance
    sngTimer = Timer
    sngOverall = sngTimer
    
    ' Set line search pattern (To remove generated timestamp)
    '<dataroot xmlns:od="urn:schemas-microsoft-com:officedata" generated="2020-04-27T10:28:32">
    rxLine.Pattern = "^\s*(?:<dataroot xmlns:(.+))( generated="".+"")"
    
    ' Open file to read contents line by line.
    Set stmInFile = New ADODB.Stream
    stmInFile.Charset = "utf-8"
    stmInFile.Open
    stmInFile.LoadFromFile strPath
    strText = stmInFile.ReadText(adReadLine)
    
    
    ' Loop through all the lines in the file
    Do Until stmInFile.EOS
        
        ' Read line from file
        strText = stmInFile.ReadText(adReadLine)
        If Left$(strText, 3) = UTF8_BOM Then strText = Mid$(strText, 4)
        ' Just looking for the first match.
        If Not blnFound Then
        
            ' Check for matching pattern
            If rxLine.Test(strText) Then
                
                ' Return actual matches
                Set objMatches = rxLine.Execute(strText)
                
                ' Replace with empty string
                strText = Replace(strText, objMatches(0).SubMatches(1), vbNullString, , 1)
                blnFound = True
            End If
        End If
        
        ' Add to return string
        cData.Add strText
        cData.Add vbCrLf
    Loop
    
    ' Close and delete original file
    stmInFile.Close
    DeleteFile strPath
    
    ' Write file all at once, rather than line by line.
    ' (Otherwise the code can bog down with tens of thousands of write operations)
    WriteFile cData.GetStr, strPath

    ' Show stats if debug turned on.
    Log.Add "    Sanitized in " & Format$(Timer - sngOverall, "0.00") & " seconds.", Options.ShowDebug

End Sub


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