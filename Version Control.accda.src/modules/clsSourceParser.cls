VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSourceParser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : clsSourceParser
' Author    : Adam Waller
' Date      : 12/4/2020
' Purpose   : Functions to sanitize files to remove non-essential metadata
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit

Public ObjectName As String

Public Enum eContentType
    ectObjectDefinition
    ectXML
    ectVBA
End Enum

' Private type to handle internal variables
Private Type udtThis
    lngSkipLines() As Long  ' Array of lines to skip
    lngSkipIndex As Long
    strFilePath As String   ' Path to loaded file
    blnOutputModified As Boolean
    strInput As String
    strOutput As String
    strVBA As String
    colBlocks As Collection
End Type
Private this As udtThis


'---------------------------------------------------------------------------------------
' Procedure : LoadFile
' Author    : Adam Waller
' Date      : 10/25/2023
' Purpose   : Load a source file to sanitize
'---------------------------------------------------------------------------------------
'
Public Function LoadSourceFile(strPath As String)
    ResetContent
    this.strInput = ReadSourceFile(strPath)
End Function


'---------------------------------------------------------------------------------------
' Procedure : LoadString
' Author    : Adam Waller
' Date      : 10/27/2023
' Purpose   : Load input from a string
'---------------------------------------------------------------------------------------
'
Public Function LoadString(ByVal strContent As String)
    ResetContent
    this.strInput = strContent
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetOutput
' Author    : Adam Waller
' Date      : 11/8/2023
' Purpose   : Wrapper to return output string
'---------------------------------------------------------------------------------------
'
Public Function GetOutput() As String
    GetOutput = this.strOutput
End Function


'---------------------------------------------------------------------------------------
' Procedure : OutputModified
' Author    : Adam Waller
' Date      : 11/8/2023
' Purpose   : Return true if the output has been modified
'---------------------------------------------------------------------------------------
'
Public Property Get OutputModified() As Boolean
    OutputModified = this.blnOutputModified
End Property


'---------------------------------------------------------------------------------------
' Procedure : ResetContent
' Author    : Adam Waller
' Date      : 10/31/2023
' Purpose   : Resets the local variables when data is loaded from a new source.
'---------------------------------------------------------------------------------------
'
Private Sub ResetContent()
    With this
        .strFilePath = vbNullString
        .strInput = vbNullString
        .strOutput = vbNullString
        .strVBA = vbNullString
        .blnOutputModified = False
    End With
    Me.ObjectName = vbNullString
End Sub


'---------------------------------------------------------------------------------------
' Procedure : Sanitize
' Author    : Adam Waller
' Date      : 10/25/2023
' Purpose   : Public wrapper for sanitize functions
'---------------------------------------------------------------------------------------
'
Public Function Sanitize(intContentType As eContentType) As String
    Select Case intContentType
        Case ectObjectDefinition:   SanitizeObject
        Case ectVBA:                SanitizeVBA
        Case ectXML:                SanitizeXML
    End Select
    Sanitize = this.strOutput
End Function


'---------------------------------------------------------------------------------------
' Procedure : MergeVBA
' Author    : Adam Waller
' Date      : 10/27/2023
' Purpose   : Merge VBA into the output content.
'---------------------------------------------------------------------------------------
'
Public Sub MergeVBA(strVbaCode As String)

    Dim varLines As Variant
    Dim lngLine As Long

    ' Make sure we have some output
    If Not Len(this.strOutput) Then this.strOutput = this.strInput

    ' Rebuild output using provided VBA code
    With New clsConcat
        .AppendOnAdd = vbCrLf

        ' Load in existing sanitized content
        varLines = Split(this.strOutput, vbCrLf)
        For lngLine = 0 To UBound(varLines)
            ' Note that the same heading name is used in both forms and reports
            If varLines(lngLine) = "CodeBehindForm" Then
                ' Allow merge of empty string to remove VBA code module
                If Len(strVbaCode) Then .Add CStr(varLines(lngLine))
                Exit For
            Else
                ' Add all other lines
                .Add CStr(varLines(lngLine))
            End If
        Next lngLine

        ' Add the VBA code here, and remove extra vbCrLf
        .Add strVbaCode
        .Remove 2

        ' Update output with combined content
        this.strOutput = .GetStr
        this.strVBA = strVbaCode
        this.blnOutputModified = True
    End With

End Sub


'---------------------------------------------------------------------------------------
' Procedure : MergePrintSettings
' Author    : Adam Waller
' Date      : 11/8/2023
' Purpose   : Merge print settings into the current source file.
'---------------------------------------------------------------------------------------
'
Public Sub MergePrintSettings(strJson As String)

    Dim dSettings As Dictionary

    ' Make sure we have some output
    If Not Len(this.strOutput) Then this.strOutput = this.strInput

    ' Don't try to parse an empty string
    If strJson = vbNullString Then Exit Sub

    ' Read settings from JSON
    Set dSettings = ParseJson(strJson)
    If dSettings.Exists("Items") Then
        With New clsDevMode
            ' Load default printer settings, then overlay
            ' settings saved with report.
            .ApplySettings dSettings("Items")
            ' Write the printer settings to the output content
            this.strOutput = .AddToExportFile(this.strOutput)
            this.blnOutputModified = True
        End With
    End If

End Sub


'---------------------------------------------------------------------------------------
' Procedure : SaveObjectVBA
' Author    : Adam Waller
' Date      : 10/31/2023
' Purpose   : Return the extracted VBA code
'---------------------------------------------------------------------------------------
'
Public Function GetObjectVBA() As String
    GetObjectVBA = this.strVBA
End Function


'---------------------------------------------------------------------------------------
' Procedure : Hash
' Author    : Adam Waller
' Date      : 10/25/2023
' Purpose   : Return a hash of the sanitized content
'---------------------------------------------------------------------------------------
'
Public Function Hash(Optional blnWithBOM As Boolean = True) As String
    Hash = GetStringHash(this.strOutput, blnWithBOM)
End Function


'---------------------------------------------------------------------------------------
' Procedure : SanitizeFile
' Author    : Adam Waller
' Date      : 11/4/2020
' Purpose   : Rewritten version of sanitize function. Returns hash of content as well
'           : as saving to the specified path.
'---------------------------------------------------------------------------------------
'
Private Function SanitizeObject() As String 'strPath As String, blnReturnHash As Boolean, Optional strObjectName As String) As String

    Dim varLines As Variant
    Dim lngLine As Long
    Dim strLine As String
    Dim strTLine As String
    Dim blnInsideIgnoredBlock As Boolean
    Dim intIndent As Integer
    Dim blnIsReport As Boolean
    Dim blnIsPassThroughQuery As Boolean
    Dim curStart As Currency
    Dim strTempFile As String
    Dim strContent As String
    Dim cVBA As clsConcat

    If DebugMode(True) Then On Error GoTo 0 Else On Error Resume Next

    ' If not sanitizing, then return output
    If Options.SanitizeLevel = eslNone Then
        this.strOutput = this.strInput
        SanitizeObject = this.strOutput
        Exit Function
    End If

    Perf.OperationStart "Sanitize File"
    varLines = Split(this.strInput, vbCrLf)

    ' Set up index of lines to skip
    ReDim this.lngSkipLines(0 To UBound(varLines)) As Long
    this.lngSkipIndex = 0
    Set this.colBlocks = New Collection

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
                    If Options.SanitizeLevel >= eslStandard Then
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
                    If blnIsPassThroughQuery Then
                        ' Ignore remaining content. (See Issue #182)
                        Do While lngLine < UBound(varLines)
                            SkipLine lngLine, eslStandard
                            lngLine = lngLine + 1
                        Loop
                        Exit Do
                    Else
                        BeginBlock
                    End If

                ' Code section behind form or report object
                Case "CodeBehindForm"
                    If Options.SplitLayoutFromVBA Then
                        ' Remove the VBA code from the layout file, but add a placeholder
                        ' comment just in case the user wonders what happened to the VBA
                        ' source code.
                        lngLine = lngLine + 1
                        Set cVBA = New clsConcat
                        cVBA.AppendOnAdd = vbCrLf
                        cVBA.Add CStr(varLines(lngLine))
                        varLines(lngLine) = "' See """ & Nz2(Me.ObjectName, "FileName") & ".cls"""
                        ' Skip remaining lines
                        Do While lngLine < UBound(varLines)
                            lngLine = lngLine + 1
                            cVBA.Add CStr(varLines(lngLine))
                            SkipLine lngLine, eslStandard
                        Loop
                        Exit Do
                    Else
                        ' Apply sanitize rules to VBA code
                        SanitizeCodeLines lngLine, varLines
                        ' Keep everything from this point on
                        Exit Do
                    End If

                Case Else
                    If blnInsideIgnoredBlock Then
                        ' Skip content inside ignored blocks.
                        SkipLine lngLine
                    ElseIf StartsWith(strTLine, "Checksum =") Then
                        ' Ignore Checksum lines, since they will change.
                        SkipLine lngLine, eslMinimal
                    ElseIf StartsWith(strTLine, "ColumnInfo =") _
                        Or StartsWith(strTLine, "BaseInfo =") Then
                        ' [ColumnInfo] contains some cached info from the record source
                        ' and will be regenerated when the form is imported. See #412
                        ' [BaseInfo] is used with combo boxes, similar to RowSource.
                        ' Since the value could span multiple lines, we need to
                        ' check the indent level of the following lines to see how
                        ' many lines to skip.
                        SkipLine lngLine, eslStandard
                        intIndent = GetIndent(strLine)
                        ' Preview the next line, and check the indent level
                        Do While GetIndent(varLines(lngLine + 1)) > intIndent
                            ' Skip previewed line and move to next line
                            SkipLine lngLine + 1, eslStandard
                            lngLine = lngLine + 1
                        Loop
                    ElseIf blnIsReport And StartsWith(strLine, "    Right =") Then
                        ' Ignore this line. (Not important, and frequently changes.)
                        SkipLine lngLine, eslStandard
                    ElseIf blnIsReport And StartsWith(strLine, "    Bottom =") Then
                        ' Turn flag back off now that we have ignored these two lines.
                        SkipLine lngLine, eslStandard
                        blnIsReport = False
                    ElseIf StartsWith(strTLine, "WebImagePadding") Then
                        ' These values tend to drift between builds. See #423
                        SkipLine lngLine, eslStandard
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
                            ' Not just an empty value (See issue #337)
                            If Len(strLine) > 20 Then
                                blnIsPassThroughQuery = True
                            End If
                        End If
                    End If

            End Select
        End If

        ' Increment counter to next line
        lngLine = lngLine + 1
    Loop

    ' Ensure that we correctly processed the nested block sequence.
    If this.colBlocks.Count > 0 Then
        Log.Error eelWarning, MultiReplace( _
            "Found ${BlockCount} unclosed blocks after sanitizing ${File}.", _
            "${BlockCount}", this.colBlocks.Count, _
            "${File}", Nz2(Me.ObjectName, this.strFilePath)), _
            ModuleName(Me) & ".SanitizeFile"
    End If

    ' Prepare primary output
    this.strOutput = BuildOutput(varLines)
    SanitizeObject = this.strOutput

    ' Prepare VBA output (if used)
    If Not cVBA Is Nothing Then
        ' Build sanitized VBA string
        With New clsSourceParser
            .LoadString cVBA.GetStr
            this.strVBA = .Sanitize(ectVBA)
        End With
    End If

    ' Log performance
    Set this.colBlocks = Nothing
    Perf.OperationEnd
    Log.Add "    Sanitized in " & Format$(Perf.MicroTimer - curStart, "0.000") & " seconds.", Options.ShowDebug

    ' Log any errors
    CatchAny eelError, "Error sanitizing " & Nz2(Me.ObjectName, FSO.GetFileName(this.strFilePath)), ModuleName(Me) & ".SanitizeFile"

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
    If this.lngSkipIndex = 0 Then
        ' No lines to skip
        ReDim this.lngSkipLines(0 To 0)
        this.lngSkipLines(0) = UBound(varLines) + 1
    Else
        ' Trim and sort index array
        ReDim Preserve this.lngSkipLines(0 To this.lngSkipIndex - 1)
        QuickSort this.lngSkipLines
    End If

    ' Use concatenation class to maximize performance
    Set cData = New clsConcat
    With cData
        .AppendOnAdd = vbCrLf

        ' Loop through array of lines in source file
        For lngLine = 0 To UBound(varLines)

            ' Iterate the sorted skipped lines index to keep up with main loop
            ' (Using parallel loops to optimize performance)
            If this.lngSkipLines(lngSkip) < lngLine Then
                If lngSkip < UBound(this.lngSkipLines) Then lngSkip = lngSkip + 1
            End If

            ' Add content, unless the line is flagged to skip
            If this.lngSkipLines(lngSkip) <> lngLine Then .Add CStr(varLines(lngLine))

        Next lngLine

        ' Remove last vbcrlf
        cData.Remove Len(vbCrLf)

        ' Return assembled output
        BuildOutput = .GetStr
    End With

End Function


'---------------------------------------------------------------------------------------
' Procedure : StripClassHeader
' Author    : Adam Waller
' Date      : 10/24/2023
' Purpose   : Strip the class header section from the VBA content. (Remove the version
'           : and VBE attributes lines that come before the actual VBA code.)
'---------------------------------------------------------------------------------------
'
Public Function StripClassHeader(strContent As String, blnStripNameOnly As Boolean) As String

    Dim lngLine As Long
    Dim varLines As Variant
    Dim strLine As String
    Dim blnPastHeader As Boolean

    ' Split code into lines
    varLines = Split(strContent, vbCrLf)

    With New clsConcat
        .AppendOnAdd = vbCrLf

        ' Skip the header information saved in the VBA class
        For lngLine = 0 To UBound(varLines)
            If Not blnPastHeader Then
                strLine = varLines(lngLine)
                If blnStripNameOnly Then
                    If StartsWith(strLine, "Attribute VB_Name = ") Then
                        ' Just skip that line. Keep everything else.
                        blnPastHeader = True
                    Else
                        .Add CStr(varLines(lngLine))
                    End If
                Else
                    Select Case True
                        Case (strLine = "VERSION 1.0 CLASS")
                        Case (strLine = "BEGIN")
                        Case (strLine = "  MultiUse = -1  'True")
                        Case (strLine = "END")
                        Case StartsWith(strLine, "Attribute VB_")
                        Case Else
                            blnPastHeader = True
                            .Add CStr(varLines(lngLine))
                    End Select
                End If
            Else
                ' Add remaining lines
                .Add CStr(varLines(lngLine))
            End If
        Next lngLine

        ' remove trailing CrLf and return result
        .Remove 2
        StripClassHeader = .GetStr
    End With

End Function


'---------------------------------------------------------------------------------------
' Procedure : SkipLine
' Author    : Adam Waller
' Date      : 6/4/2021
' Purpose   : Skip this line in the final output file. Optionally include a minimum
'           : sanitize level to skip this line.
'---------------------------------------------------------------------------------------
'
Private Function SkipLine(lngLine As Long, Optional intMinSanitizeLevel As eSanitizeLevel)
    If Options.SanitizeLevel >= intMinSanitizeLevel Then
        this.lngSkipLines(this.lngSkipIndex) = lngLine
        this.lngSkipIndex = this.lngSkipIndex + 1
    End If
End Function


'---------------------------------------------------------------------------------------
' Procedure : SanitizeCodeLines
' Author    : Adam Waller
' Date      : 7/11/2023
' Purpose   : Perform any sanitizing of code lines.
'---------------------------------------------------------------------------------------
'
Private Function SanitizeCodeLines(lngLineStart As Long, ByRef varLines As Variant)

    Dim lngLine As Long

    Perf.OperationStart "Sanitize Code Lines"
    For lngLine = lngLineStart To UBound(varLines)
        ' Check for lines that include only space padding
        ' added by the IDE automatic indenting. (The padding is removed if you
        ' comment out a block, then uncomment the same block, causing unwanted
        ' noise in version control.)
        If Len(varLines(lngLine)) > 0 Then
            If Trim(varLines(lngLine)) = vbNullString Then
                ' Remove the space padding
                varLines(lngLine) = vbNullString
            End If
        End If
    Next lngLine
    Perf.OperationEnd

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
    If this.colBlocks Is Nothing Then Set this.colBlocks = New Collection
    Set dBlock = New Dictionary
    If strType <> vbNullString Then dBlock.Add "Type", strType
    this.colBlocks.Add dBlock
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
    If this.colBlocks.Count = 0 Then Exit Sub
    Set dBlock = this.colBlocks(this.colBlocks.Count)

    ' Skip if we are not using themes for this control (UseTheme=0)
    ' (Applies to "CommandButton", "Tab", "ToggleButton")
    If dBlock.Exists("UseTheme") Then
        ' Remove this block
        this.colBlocks.Remove this.colBlocks.Count
        Exit Sub
    End If

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
                    If Options.SanitizeColors = eslExtended Then
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
    this.colBlocks.Remove this.colBlocks.Count

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
    Dim lngColor As Long

    ' Skip if not using this option
    If Options.SanitizeColors <= eslNone Then Exit Sub

    ' Exit if we are not inside a block
    If Not this.colBlocks Is Nothing Then lngCnt = this.colBlocks.Count
    If lngCnt = 0 Then Exit Sub
    Set dBlock = this.colBlocks(this.colBlocks.Count)

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
' Procedure : SanitizeVBA
' Author    : Adam Waller
' Date      : 7/12/2023
' Purpose   : Standardizes blank line padding and trailing lines in VBA code content.
'           : NOTE: This does not make any code changes, but only modifies padding
'           : and blank lines to generate standardized output for version control.
'           :
'           : OPERATIONS:
'           :   - Trim blank lines that only include space padding.
'           :   - Remove extra trailing lines from the end of the module.
'---------------------------------------------------------------------------------------
'
Private Function SanitizeVBA() As String

    Dim lngLine As Long
    Dim varLines As Variant
    Dim lngLastLine As Long

    ' Skip sanitizing if not using that option.
    If Options.SanitizeLevel < eslStandard Then
        SanitizeVBA = this.strInput
        Exit Function
    End If

    Perf.OperationStart "Sanitize VBA Code"

    ' Split code into lines
    varLines = Split(this.strInput, vbCrLf)

    ' Build sanitized content
    With New clsConcat
        .AppendOnAdd = vbCrLf

        ' Find the last non-blank line
        For lngLine = UBound(varLines) To 0 Step -1
            If Len(Trim(varLines(lngLine))) > 0 Then
                lngLastLine = lngLine
                Exit For
            End If
        Next lngLine

        ' Loop through lines
        For lngLine = 0 To lngLastLine
            .Add RTrim(varLines(lngLine))
        Next lngLine

        ' Return standardized code block
        this.strOutput = .GetStr
        this.strVBA = this.strOutput
        SanitizeVBA = this.strOutput
        Perf.OperationEnd
    End With

End Function


'---------------------------------------------------------------------------------------
' Procedure : SanitizeXML
' Author    : Adam Waller
' Date      : 4/29/2021
' Purpose   : Remove non-essential data that changes every time the file is exported.
'           : Optionally returns a hash of the file content. (To save reading the file
'           : back again afterwards to compute the hash.)
'---------------------------------------------------------------------------------------
'
Private Function SanitizeXML() As String

    Dim curStart As Currency
    Dim cData As clsConcat
    Dim strXML As String
    Dim rxLine As VBScript_RegExp_55.RegExp

    If DebugMode(True) Then On Error GoTo 0 Else On Error Resume Next

    Set cData = New clsConcat
    cData.AppendOnAdd = vbCrLf
    Set rxLine = New VBScript_RegExp_55.RegExp

    Perf.OperationStart "Sanitize XML"
    curStart = Perf.MicroTimer
    strXML = this.strInput

    ' Exporting Table Def as XML does not properly encode ampersand character (See #314)
    ' Most likely if any ampersands are encoded correctly, all of them will be.
    With New VBScript_RegExp_55.RegExp
        .Multiline = True
        .Global = True
        ' Match &amp; &quot; &gt; &lt; etc...
        .Pattern = "&[A-z]{2,6};"
        If Not .Test(strXML) Then
            ' Properly encode any embedded ampersand characters to make valid XML
            strXML = Replace(strXML, "&", "&amp;")
        End If
    End With

    Static objXml As MSXML2.DOMDocument60
    Dim objNode As MSXML2.IXMLDOMNode

    If objXml Is Nothing Then
        Set objXml = New MSXML2.DOMDocument60
    End If

    If objXml.LoadXML(strXML) = False Then
        Log.Error eelError, _
            "Unable to parse the XML for file '" & Nz2(this.strFilePath, Me.ObjectName) & _
            "'. This may be due to containing malformed XML. Check the source XML document for validity. " & _
            "In some cases, this may be due to table data containing characters not allowed in XML documents.", _
            ModuleName(Me) & ".SanitizeXML"
        Exit Function
    End If

    ' Determine if it's a table data with schema
    For Each objNode In objXml.SelectNodes("/root/dataroot")
        ' Remove the generated timestamp attribute to reduce noise
        '   <dataroot xmlns:od="urn:schemas-microsoft-com:officedata" generated="2020-04-27T10:28:32">
        '   <dataroot generated="2021-04-29T17:27:33" xmlns:od="urn:schemas-microsoft-com:officedata">
        objNode.Attributes.removeNamedItem "generated"

        ' Determine whether the schema is required for import. If the schema contains elements:
        '   <xsd:element od:expression ...>
        '   <xsd:element od:jetType="complex" ...>
        '   <xsd:element od:jetType="oleobject" ...>
        ' Then the schema must be retained. Otherwise, discard the schema and retain only the data.
        If objXml.SelectNodes("//*[(namespace-uri()='http://www.w3.org/2001/XMLSchema' and local-name()='element' and @*[namespace-uri()='urn:schemas-microsoft-com:officedata' and ((local-name()='jetType' and (string()='complex' or string()='oleobject')) or (local-name()='expression'))])]").Length = 0 Then
            objXml.replaceChild objXml.SelectSingleNode("/root/dataroot"), objXml.SelectSingleNode("/root")
        End If
    Next

    ' Remove generated timestamp from tables exported without schema (such as linked tables)
    For Each objNode In objXml.SelectNodes("/dataroot")
        objNode.Attributes.removeNamedItem "generated"
    Next

    ' Remove all nodes that are meaningless noise:
    '   <od:tableProperty name="NameMap" ...>
    '   <od:tableProperty name="GUID" ...>
    '   <od:fieldProperty name="GUID" ...>
    For Each objNode In objXml.SelectNodes("//*[(namespace-uri()='urn:schemas-microsoft-com:officedata' and local-name()='tableProperty' and (@name='NameMap' or @name='GUID')) or (namespace-uri()='urn:schemas-microsoft-com:officedata' and local-name()='fieldProperty' and @name='GUID')]")
        objNode.ParentNode.RemoveChild objNode
    Next

    If Options.StripPublishOption Then
        ' Remove all web publish options:
        '   <od:tableProperty name="PublishToWeb" ...>
        For Each objNode In objXml.SelectNodes("//*[(namespace-uri()='urn:schemas-microsoft-com:officedata' and local-name()='tableProperty' and @name='PublishToWeb')]")
            objNode.ParentNode.RemoveChild objNode
        Next
    End If

    Perf.OperationEnd

    ' Save the output
    this.strOutput = FormatXML(objXml)
    SanitizeXML = this.strOutput

    ' Show stats if debug turned on.
    Log.Add "    Sanitized in " & Format$(Perf.MicroTimer - curStart, "0.000") & " seconds.", Options.ShowDebug

    ' Log any errors
    CatchAny eelError, "Error sanitizing XML for " & Nz2(FSO.GetFileName(this.strFilePath), Me.ObjectName), ModuleName(Me) & ".SanitizeXML"

End Function


'---------------------------------------------------------------------------------------
' Procedure : TrimTabs
' Author    : Adam Waller
' Date      : 4/29/2021
' Purpose   : Trim off tabs from beginning and end of string
'---------------------------------------------------------------------------------------
'
Private Function TrimTabs(strText As String) As String

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
' Procedure : GetIndent
' Author    : Adam Waller
' Date      : 11/5/2020
' Purpose   : Returns the number of spaces until the first non-space character.
'---------------------------------------------------------------------------------------
'
Private Function GetIndent(strLine As Variant) As Integer
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
Private Function FormatXML(objInput As MSXML2.DOMDocument60, _
                            Optional blnOmitDeclaration As Boolean) As String

    ' XSLT stylesheet that allow us to control indenting and also get a better indent result.
    ' For testing and adjusting, you can use https://www.online-toolz.com/tools/xslt-validator-tester-online.php
    'Const strIndentXslt As String = "<xsl:stylesheet xmlns:xsl=""http://www.w3.org/1999/XSL/Transform"" version=""1.0""><xsl:output method=""xml""/><xsl:template match=""@*""><xsl:copy/></xsl:template><xsl:template match=""text()""><xsl:value-of select=""normalize-space(.)""/></xsl:template><xsl:template match=""*""><xsl:param name=""indent"" select=""''""/><xsl:text>&#xA;</xsl:text><xsl:value-of select=""$indent""/><xsl:copy><xsl:apply-templates select=""@*|*|text()""><xsl:with-param name=""indent"" select=""concat($indent, '  ')""/></xsl:apply-templates></xsl:copy><xsl:if test=""count(../*)&gt;0 and ../*[last()]=. and not(following-sibling::*)""><xsl:text>&#xA;</xsl:text><xsl:value-of select=""substring($indent,3)""/></xsl:if></xsl:template></xsl:stylesheet>"
    Const strIndentXslt As String = "<xsl:stylesheet xmlns:xsl=""http://www.w3.org/1999/XSL/Transform"" version=""1.0""><xsl:output method=""xml""/><xsl:template match=""@*""><xsl:copy/></xsl:template><xsl:template match=""*""><xsl:param name=""indent"" select=""''""/><xsl:text>&#xA;</xsl:text><xsl:value-of select=""$indent""/><xsl:copy><xsl:apply-templates select=""@*|*|text()""><xsl:with-param name=""indent"" select=""concat($indent, '  ')""/></xsl:apply-templates></xsl:copy><xsl:if test=""count(../*)&gt;0 and ../*[last()]=. and not(following-sibling::*)""><xsl:text>&#xA;</xsl:text><xsl:value-of select=""substring($indent,3)""/></xsl:if></xsl:template></xsl:stylesheet>"
    ' This constant has the `omit-xml-declaration="yes"` added to remove XML declarations.
    'Const strIndentXsltNoDeclarations As String = "<xsl:stylesheet xmlns:xsl=""http://www.w3.org/1999/XSL/Transform"" version=""1.0""><xsl:output method=""xml"" omit-xml-declaration=""yes""/><xsl:template match=""@*""><xsl:copy/></xsl:template><xsl:template match=""text()""><xsl:value-of select=""normalize-space(.)""/></xsl:template><xsl:template match=""*""><xsl:param name=""indent"" select=""''""/><xsl:text>&#xA;</xsl:text><xsl:value-of select=""$indent""/><xsl:copy><xsl:apply-templates select=""@*|*|text()""><xsl:with-param name=""indent"" select=""concat($indent, '  ')""/></xsl:apply-templates></xsl:copy><xsl:if test=""count(../*)&gt;0 and ../*[last()]=. and not(following-sibling::*)""><xsl:text>&#xA;</xsl:text><xsl:value-of select=""substring($indent,3)""/></xsl:if></xsl:template></xsl:stylesheet>"
    Const strIndentXsltNoDeclarations As String = "<xsl:stylesheet xmlns:xsl=""http://www.w3.org/1999/XSL/Transform"" version=""1.0""><xsl:output method=""xml"" omit-xml-declaration=""yes""/><xsl:template match=""@*""><xsl:copy/></xsl:template><xsl:template match=""*""><xsl:param name=""indent"" select=""''""/><xsl:text>&#xA;</xsl:text><xsl:value-of select=""$indent""/><xsl:copy><xsl:apply-templates select=""@*|*|text()""><xsl:with-param name=""indent"" select=""concat($indent, '  ')""/></xsl:apply-templates></xsl:copy><xsl:if test=""count(../*)&gt;0 and ../*[last()]=. and not(following-sibling::*)""><xsl:text>&#xA;</xsl:text><xsl:value-of select=""substring($indent,3)""/></xsl:if></xsl:template></xsl:stylesheet>"

    Static objTransform As MSXML2.DOMDocument60
    Static objTransformNoDeclaration As MSXML2.DOMDocument60

    Dim strOutput As String

    ' Skip processing if no content to format
    If objInput.ChildNodes.Length = 0 Then Exit Function

    Perf.OperationStart "Format XML"

    ' Trap any errors with parsing or formatting the XML
    If DebugMode(True) Then On Error GoTo 0 Else On Error Resume Next

    ' Transform the input; we don't want to use transformNodeToObject
    ' because that would be defeated by MSXML reformatting when reading
    ' from the XML property. We also cache the MSXML2.DOMDocument to
    ' avoid paying the cost of loading the XSLT repeatedly.
    If blnOmitDeclaration Then
        If objTransformNoDeclaration Is Nothing Then
            Set objTransformNoDeclaration = New MSXML2.DOMDocument60
            objTransformNoDeclaration.LoadXML strIndentXsltNoDeclarations
        End If
        strOutput = objInput.transformNode(objTransformNoDeclaration)
    Else
        If objTransform Is Nothing Then
            Set objTransform = New MSXML2.DOMDocument60
            objTransform.LoadXML strIndentXslt
        End If
        strOutput = objInput.transformNode(objTransform)
    End If

    ' Check for any errors parsing the XML
    If CatchAny(eelError, "Error parsing XML content", ModuleName(Me) & ".FormatXML") Then
        ' Fall back to input XML
        strOutput = objInput.XML
        ' Output XML to log file
        Log.Spacer False
        Log.Add objInput.XML, False
        Log.Spacer False
    End If

    ' Return formatted output
    Perf.OperationEnd
    FormatXML = strOutput

End Function


'---------------------------------------------------------------------------------------
' Procedure : Class_Terminate
' Author    : Adam Waller
' Date      : 10/25/2023
' Purpose   : Release private objects when terminating
'---------------------------------------------------------------------------------------
'
Private Sub Class_Terminate()
    Set this.colBlocks = Nothing
End Sub
