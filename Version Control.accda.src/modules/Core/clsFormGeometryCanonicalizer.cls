VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFormGeometryCanonicalizer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : clsFormGeometryCanonicalizer
' Author    : Adam Waller
' Date      : 9/7/2026
' Purpose   : Structure-aware 60-twip canonicalizer for form layout geometry. Port of
'           : tools/dpi-layout-probe/form_geometry.py. Rewrites DPI-derived layout
'           : values in place; never adds or removes a geometry property.
'           : See docs\access-form-geometry.md.
'           :
'           : Performance note: parsed blocks and geometry properties live in a typed
'           : UDT array. The parser classifies structural and property lines by length
'           : before binary comparisons. Replacing either path with generic Dictionary
'           : lookups or locale-aware comparisons materially slows large forms.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
'@Folder("Core")

Private Const GRID_TWIPS As Long = 60
Private Const ModuleNameConst As String = "clsFormGeometryCanonicalizer"

Private Enum eGeometryAxis
    egaHorizontal
    egaVertical
End Enum

Private Enum eGeometryProperty
    egpNone
    egpLeft
    egpTop
    egpWidth
    egpHeight
    egpColumnStart
    egpColumnEnd
    egpRowStart
    egpRowEnd
    egpLayoutGroup
    egpGroupTable
    egpCachedLeft
    egpCachedTop
    egpCachedWidth
    egpCachedHeight
End Enum

' Parsed geometry for one Begin/End block. Long indices link blocks to their section
' and tab ancestors without creating object reference cycles.
Private Type udtBlock
    strType As String
    lngSection As Long
    lngTab As Long

    lngLeft As Long
    lngLeftLine As Long
    lngTop As Long
    lngTopLine As Long
    lngWidth As Long
    lngWidthLine As Long
    lngHeight As Long
    lngHeightLine As Long

    blnColumnStart As Boolean
    lngColumnStart As Long
    blnColumnEnd As Boolean
    lngColumnEnd As Long
    blnRowStart As Boolean
    lngRowStart As Long
    blnRowEnd As Boolean
    lngRowEnd As Long
    blnLayoutGroup As Boolean
    lngLayoutGroup As Long
    blnGroupTable As Boolean
    lngGroupTable As Long
    blnCachedLeft As Boolean
    lngCachedLeft As Long
    blnCachedTop As Boolean
    lngCachedTop As Long
    blnCachedWidth As Boolean
    lngCachedWidth As Long
    blnCachedHeight As Boolean
    lngCachedHeight As Long
End Type

' A single planned geometry rewrite. Held in an array so the planning passes do not
' allocate an object per change.
Private Type udtChange
    lngBlock As Long
    eProperty As eGeometryProperty
    lngValue As Long
End Type

Private Type udtThis
    lngRoot As Long
    strFormName As String
    lngChanged As Long
End Type
Private this As udtThis

' Typed storage replaces three Dictionaries and one Collection previously allocated
' for every parsed block.
Private m_aBlocks() As udtBlock
Private m_lngBlockCount As Long
Private m_blnHasLayout As Boolean

' The line buffer is read once per line by the parse and written by each rewrite. It is
' held directly rather than as a UDT member, and its bound cached, because both indirect
' access and re-evaluating UBound measured significantly slower on large forms.
Private m_strLines() As String
Private m_lngUpper As Long

' Phase timings in seconds from the most recent Canonicalize call, used by
' modTestPerf.BenchmarkFormGeometry to attribute cost across the passes.
Public PhaseCopyIn As Double
Public PhaseParse As Double
Public PhaseGroups As Double
Public PhaseTabs As Double
Public PhaseEnvelopes As Double
Public PhaseCopyOut As Double
Public BlockCount As Long
Public GroupCount As Long
Public PlanAxisCalls As Long
Public RewriteCalls As Long

' Number of lines this pass actually rewrote. Zero means the input was already
' canonical, which is the expected result for a file that has been exported before.
Public ChangedLines As Long


'---------------------------------------------------------------------------------------
' Procedure : Canonicalize
' Author    : Adam Waller
' Date      : 9/7/2026
' Purpose   : Rewrite layout-derived geometry in the supplied line array so the same
'           : form exports identically regardless of the DPI Access last solved it at.
'---------------------------------------------------------------------------------------
'
Public Sub Canonicalize(ByRef varLines As Variant, Optional strFormName As String)

    If DebugMode(True) Then On Error GoTo 0 Else On Error Resume Next

    Dim curPhase As Currency
    Dim lngLine As Long

    Perf.OperationStart "Canonicalize Form Geometry"
    curPhase = Perf.MicroTimer

    ' Work from a String array with a cached bound. Reading lines through a Variant
    ' array, and re-evaluating UBound on every loop pass, measured materially slower.
    m_lngUpper = UBound(varLines)
    ReDim m_strLines(0 To m_lngUpper)
    For lngLine = 0 To m_lngUpper
        m_strLines(lngLine) = CStr(varLines(lngLine))
    Next lngLine
    this.strFormName = strFormName
    this.lngChanged = 0
    this.lngRoot = -1
    m_lngBlockCount = 0
    m_blnHasLayout = False
    ReDim m_aBlocks(0 To 255)
    Me.BlockCount = 0
    Me.GroupCount = 0
    Me.PlanAxisCalls = 0
    Me.RewriteCalls = 0
    Me.ChangedLines = 0
    Me.PhaseCopyIn = CDbl(Perf.MicroTimer - curPhase)

    curPhase = Perf.MicroTimer
    ParseDocument
    Me.PhaseParse = CDbl(Perf.MicroTimer - curPhase)

    curPhase = Perf.MicroTimer
    CanonicalizeGroups
    Me.PhaseGroups = CDbl(Perf.MicroTimer - curPhase)

    curPhase = Perf.MicroTimer
    CanonicalizeTabs
    Me.PhaseTabs = CDbl(Perf.MicroTimer - curPhase)

    curPhase = Perf.MicroTimer
    GrowEnvelopes
    Me.PhaseEnvelopes = CDbl(Perf.MicroTimer - curPhase)

    curPhase = Perf.MicroTimer
    Me.BlockCount = m_lngBlockCount
    Me.ChangedLines = this.lngChanged
    If this.lngChanged > 0 Then
        For lngLine = 0 To m_lngUpper
            varLines(lngLine) = m_strLines(lngLine)
        Next lngLine
    End If
    Erase m_aBlocks
    Erase m_strLines
    Me.PhaseCopyOut = CDbl(Perf.MicroTimer - curPhase)
    Perf.OperationEnd

    CatchAny eelError, T("Error canonicalizing form geometry in {0}", _
        var0:=Nz2(strFormName, "form")), ModuleNameConst & ".Canonicalize", True, True

End Sub


'---------------------------------------------------------------------------------------
' Procedure : ParseDocument
'---------------------------------------------------------------------------------------
'
Private Sub ParseDocument()

    Dim lngLine As Long
    Dim lngBlock As Long
    Dim strTrimmed As String

    lngLine = 0
    Do While lngLine <= m_lngUpper
        strTrimmed = Trim$(m_strLines(lngLine))
        If Len(strTrimmed) = 14 Then
            If strTrimmed = "CodeBehindForm" Then Exit Do
        End If
        If StartsWith(strTrimmed, "Begin") Then
            lngLine = ParseBlock(lngLine + 1, Mid$(strTrimmed, 6), -1, lngBlock)
            this.lngRoot = lngBlock
        Else
            lngLine = lngLine + 1
        End If
    Loop

End Sub


'---------------------------------------------------------------------------------------
' Procedure : ParseBlock
'---------------------------------------------------------------------------------------
'
Private Function ParseBlock(ByVal lngLine As Long, strType As String, _
    lngParent As Long, ByRef lngBlock As Long) As Long

    Dim strTrimmed As String
    Dim strName As String
    Dim strValue As String
    Dim lngPos As Long
    Dim lngLen As Long
    Dim lngUpper As Long
    Dim blnIsBegin As Boolean
    Dim blnIsInlineBinary As Boolean
    Dim lngChild As Long
    Dim eProperty As eGeometryProperty
    Dim lngValue As Long

    lngUpper = m_lngUpper
    lngBlock = NewBlock(Trim$(strType), lngParent)

    Do While lngLine <= lngUpper
        strTrimmed = Trim$(m_strLines(lngLine))
        lngLen = Len(strTrimmed)

        ' Classify structural lines by length and first character before comparing
        ' contents. VBA does not short-circuit And, so the gates have to nest.
        If lngLen = 3 Then
            If StrComp(strTrimmed, "End", vbBinaryCompare) = 0 Then
                ParseBlock = lngLine + 1
                Exit Function
            End If
        ElseIf lngLen = 14 Then
            If StrComp(strTrimmed, "CodeBehindForm", vbBinaryCompare) = 0 Then
                ParseBlock = lngUpper + 1
                Exit Function
            End If
        End If

        blnIsBegin = False
        If lngLen >= 5 Then
            If AscW(strTrimmed) = 66 Then
                blnIsBegin = (StrComp(Left$(strTrimmed, 5), "Begin", vbBinaryCompare) = 0)
            End If
        End If
        If blnIsBegin Then
            lngLine = ParseBlock(lngLine + 1, Mid$(strTrimmed, 6), lngBlock, lngChild)
        Else
            lngPos = InStr(strTrimmed, " =")
            If lngPos > 1 Then
                ' Inline binary blocks are properties whose value is Begin. Detect
                ' them here instead of calling EndsWith/StartsWith for every line.
                blnIsInlineBinary = False
                If lngLen - lngPos = 7 Then
                    blnIsInlineBinary = (StrComp(Right$(strTrimmed, 7), _
                        "= Begin", vbBinaryCompare) = 0)
                End If
                If blnIsInlineBinary Then
                    lngLine = lngLine + 1
                    Do While lngLine <= lngUpper
                        strTrimmed = Trim$(m_strLines(lngLine))
                        If Len(strTrimmed) = 3 Then
                            If StrComp(strTrimmed, "End", vbBinaryCompare) = 0 Then Exit Do
                        End If
                        lngLine = lngLine + 1
                    Loop
                Else
                    ' Dispatch first on property-name length. This avoids locale-aware
                    ' comparisons and value parsing for unrelated form properties.
                    eProperty = egpNone
                    Select Case lngPos
                    Case 4
                        If StrComp(Left$(strTrimmed, 3), "Top", vbBinaryCompare) = 0 Then _
                            eProperty = egpTop
                    Case 5
                        If StrComp(Left$(strTrimmed, 4), "Left", vbBinaryCompare) = 0 Then _
                            eProperty = egpLeft
                    Case 6
                        If StrComp(Left$(strTrimmed, 5), "Width", vbBinaryCompare) = 0 Then _
                            eProperty = egpWidth
                    Case 7
                        strName = Left$(strTrimmed, 6)
                        If StrComp(strName, "Height", vbBinaryCompare) = 0 Then
                            eProperty = egpHeight
                        ElseIf StrComp(strName, "RowEnd", vbBinaryCompare) = 0 Then
                            eProperty = egpRowEnd
                        End If
                    Case 9
                        If StrComp(Left$(strTrimmed, 8), "RowStart", vbBinaryCompare) = 0 Then _
                            eProperty = egpRowStart
                    Case 10
                        If StrComp(Left$(strTrimmed, 9), "ColumnEnd", vbBinaryCompare) = 0 Then _
                            eProperty = egpColumnEnd
                    Case 11
                        If StrComp(Left$(strTrimmed, 10), "GroupTable", vbBinaryCompare) = 0 Then _
                            eProperty = egpGroupTable
                    Case 12
                        strName = Left$(strTrimmed, 11)
                        If StrComp(strName, "ColumnStart", vbBinaryCompare) = 0 Then
                            eProperty = egpColumnStart
                        ElseIf StrComp(strName, "LayoutGroup", vbBinaryCompare) = 0 Then
                            eProperty = egpLayoutGroup
                        End If
                    Case 16
                        If StrComp(Left$(strTrimmed, 15), "LayoutCachedTop", _
                            vbBinaryCompare) = 0 Then eProperty = egpCachedTop
                    Case 17
                        If StrComp(Left$(strTrimmed, 16), "LayoutCachedLeft", _
                            vbBinaryCompare) = 0 Then eProperty = egpCachedLeft
                    Case 18
                        If StrComp(Left$(strTrimmed, 17), "LayoutCachedWidth", _
                            vbBinaryCompare) = 0 Then eProperty = egpCachedWidth
                    Case 19
                        If StrComp(Left$(strTrimmed, 18), "LayoutCachedHeight", _
                            vbBinaryCompare) = 0 Then eProperty = egpCachedHeight
                    End Select

                    If eProperty <> egpNone Then
                        strValue = Mid$(strTrimmed, lngPos + 2)
                        If IsIntegerToken(strValue) Then
                            lngValue = CLng(strValue)
                            With m_aBlocks(lngBlock)
                                Select Case eProperty
                                Case egpLeft
                                    .lngLeft = lngValue
                                    .lngLeftLine = lngLine
                                Case egpTop
                                    .lngTop = lngValue
                                    .lngTopLine = lngLine
                                Case egpWidth
                                    .lngWidth = lngValue
                                    .lngWidthLine = lngLine
                                Case egpHeight
                                    .lngHeight = lngValue
                                    .lngHeightLine = lngLine
                                Case egpColumnStart
                                    .blnColumnStart = True
                                    .lngColumnStart = lngValue
                                Case egpColumnEnd
                                    .blnColumnEnd = True
                                    .lngColumnEnd = lngValue
                                Case egpRowStart
                                    .blnRowStart = True
                                    .lngRowStart = lngValue
                                Case egpRowEnd
                                    .blnRowEnd = True
                                    .lngRowEnd = lngValue
                                Case egpLayoutGroup
                                    .blnLayoutGroup = True
                                    .lngLayoutGroup = lngValue
                                    m_blnHasLayout = True
                                Case egpGroupTable
                                    .blnGroupTable = True
                                    .lngGroupTable = lngValue
                                    m_blnHasLayout = True
                                Case egpCachedLeft
                                    .blnCachedLeft = True
                                    .lngCachedLeft = lngValue
                                Case egpCachedTop
                                    .blnCachedTop = True
                                    .lngCachedTop = lngValue
                                Case egpCachedWidth
                                    .blnCachedWidth = True
                                    .lngCachedWidth = lngValue
                                Case egpCachedHeight
                                    .blnCachedHeight = True
                                    .lngCachedHeight = lngValue
                                End Select
                            End With
                        End If
                    End If
                End If
            End If
            lngLine = lngLine + 1
        End If
    Loop

    ParseBlock = lngLine

End Function


'---------------------------------------------------------------------------------------
' Procedure : NewBlock
'---------------------------------------------------------------------------------------
'
Private Function NewBlock(strType As String, lngParent As Long) As Long

    Dim lngBlock As Long

    If m_lngBlockCount > UBound(m_aBlocks) Then
        ReDim Preserve m_aBlocks(0 To (UBound(m_aBlocks) + 1) * 2 - 1)
    End If
    lngBlock = m_lngBlockCount
    m_lngBlockCount = m_lngBlockCount + 1

    With m_aBlocks(lngBlock)
        .strType = strType
        .lngSection = -1
        .lngTab = -1
        .lngLeftLine = -1
        .lngTopLine = -1
        .lngWidthLine = -1
        .lngHeightLine = -1
    End With

    If lngParent >= 0 Then
        With m_aBlocks(lngParent)
            If IsSectionType(.strType) Then
                m_aBlocks(lngBlock).lngSection = lngParent
            Else
                m_aBlocks(lngBlock).lngSection = .lngSection
            End If
            If .strType = "Tab" Then
                m_aBlocks(lngBlock).lngTab = lngParent
            Else
                m_aBlocks(lngBlock).lngTab = .lngTab
            End If
        End With
    End If
    NewBlock = lngBlock

End Function


'---------------------------------------------------------------------------------------
' Procedure : CanonicalizeGroups
'---------------------------------------------------------------------------------------
'
Private Sub CanonicalizeGroups()

    Dim dGroups As Dictionary
    Dim colGroup As Collection
    Dim strKey As String
    Dim varKey As Variant
    Dim lngBlock As Long

    Set dGroups = New Dictionary
    For lngBlock = 0 To m_lngBlockCount - 1
        With m_aBlocks(lngBlock)
            If .blnLayoutGroup Or .blnGroupTable Then
                If .blnGroupTable Then
                    strKey = CStr(.lngGroupTable)
                Else
                    strKey = "-1"
                End If
                strKey = strKey & "/"
                If .blnLayoutGroup Then
                    strKey = strKey & .lngLayoutGroup
                Else
                    strKey = strKey & "-1"
                End If
                If dGroups.Exists(strKey) Then
                    Set colGroup = dGroups(strKey)
                Else
                    Set colGroup = New Collection
                    dGroups.Add strKey, colGroup
                End If
                colGroup.Add lngBlock
            End If
        End With
    Next lngBlock

    Me.GroupCount = dGroups.Count
    For Each varKey In dGroups.Keys
        Set colGroup = dGroups(varKey)
        CanonicalizeGroup colGroup, CStr(varKey)
    Next varKey

End Sub


'---------------------------------------------------------------------------------------
' Procedure : CanonicalizeGroup
'---------------------------------------------------------------------------------------
'
Private Sub CanonicalizeGroup(colCells As Collection, strLabel As String)

    Dim aChanges() As udtChange
    Dim lngChanges As Long
    Dim dBySection As Dictionary
    Dim colSection As Collection
    Dim strSection As String
    Dim varBlock As Variant
    Dim varSection As Variant
    Dim lngI As Long
    Dim lngBlock As Long

    ReDim aChanges(0 To 63)
    lngChanges = 0

    If Not PlanAxis(colCells, egaHorizontal, aChanges, lngChanges) Then
        WarnSkip strLabel & " (horizontal)"
        Exit Sub
    End If

    ' Vertical geometry is planned per section, since rows restart in each one.
    Set dBySection = New Dictionary
    For Each varBlock In colCells
        lngBlock = CLng(varBlock)
        strSection = CStr(m_aBlocks(lngBlock).lngSection)
        If dBySection.Exists(strSection) Then
            Set colSection = dBySection(strSection)
        Else
            Set colSection = New Collection
            dBySection.Add strSection, colSection
        End If
        colSection.Add lngBlock
    Next varBlock

    For Each varSection In dBySection.Keys
        Set colSection = dBySection(varSection)
        If Not PlanAxis(colSection, egaVertical, aChanges, lngChanges) Then
            WarnSkip strLabel & " (vertical)"
            Exit Sub
        End If
    Next varSection

    For lngI = 0 To lngChanges - 1
        RewriteProp aChanges(lngI).lngBlock, aChanges(lngI).eProperty, _
            aChanges(lngI).lngValue
    Next lngI

End Sub


'---------------------------------------------------------------------------------------
' Procedure : PlanAxis
' Purpose   : Append the planned rewrites for one axis to aChanges. Returns False when
'           : the group cannot be solved, in which case no change is applied at all.
'---------------------------------------------------------------------------------------
'
Private Function PlanAxis(colCells As Collection, eAxis As eGeometryAxis, _
    ByRef aChanges() As udtChange, _
    ByRef lngChanges As Long) As Boolean

    Dim aBlocks() As Long
    Dim lngCount As Long
    Dim dSizes As Dictionary
    Dim dPositions As Dictionary
    Dim lngStart As Long
    Dim lngEnd As Long
    Dim lngI As Long
    Dim varDerived As Variant
    Dim varBlock As Variant

    Me.PlanAxisCalls = Me.PlanAxisCalls + 1

    lngCount = colCells.Count
    If lngCount = 0 Then
        PlanAxis = True
        Exit Function
    End If
    ReDim aBlocks(0 To lngCount - 1)
    lngI = 0
    For Each varBlock In colCells
        aBlocks(lngI) = CLng(varBlock)
        lngI = lngI + 1
    Next varBlock

    Set dSizes = TrackSizes(aBlocks, lngCount, eAxis)
    Set dPositions = BoundaryPositions(aBlocks, lngCount, eAxis)
    If dPositions Is Nothing Then Exit Function
    InferEndSizes aBlocks, lngCount, eAxis, dPositions, dSizes

    For lngI = 0 To lngCount - 1
        lngStart = GetTrackStart(aBlocks(lngI), eAxis)
        lngEnd = GetTrackEnd(aBlocks(lngI), eAxis, lngStart)
        If HasPosition(aBlocks(lngI), eAxis) Then
            If Not dPositions.Exists(lngStart) Then Exit Function
            If eAxis = egaHorizontal Then
                AddChange aChanges, lngChanges, aBlocks(lngI), egpLeft, _
                    CLng(dPositions(lngStart))
            Else
                AddChange aChanges, lngChanges, aBlocks(lngI), egpTop, _
                    CLng(dPositions(lngStart))
            End If
        End If
        If HasSize(aBlocks(lngI), eAxis) Then
            If lngStart = lngEnd Then
                If dSizes.Exists(lngStart) Then
                    varDerived = CLng(dSizes(lngStart))
                Else
                    varDerived = Null
                End If
            Else
                varDerived = DeriveSpan(lngStart, lngEnd, dPositions, dSizes)
            End If
            If IsNull(varDerived) Then Exit Function
            If eAxis = egaHorizontal Then
                AddChange aChanges, lngChanges, aBlocks(lngI), egpWidth, CLng(varDerived)
            Else
                AddChange aChanges, lngChanges, aBlocks(lngI), egpHeight, CLng(varDerived)
            End If
        End If
    Next lngI

    PlanAxis = True

End Function


'---------------------------------------------------------------------------------------
' Procedure : AddChange
'---------------------------------------------------------------------------------------
'
Private Sub AddChange(ByRef aChanges() As udtChange, ByRef lngChanges As Long, _
    lngBlock As Long, eProperty As eGeometryProperty, lngValue As Long)

    If lngChanges > UBound(aChanges) Then
        ReDim Preserve aChanges(0 To (UBound(aChanges) + 1) * 2 - 1)
    End If
    aChanges(lngChanges).lngBlock = lngBlock
    aChanges(lngChanges).eProperty = eProperty
    aChanges(lngChanges).lngValue = lngValue
    lngChanges = lngChanges + 1

End Sub


'---------------------------------------------------------------------------------------
' Procedure : TrackSizes
'---------------------------------------------------------------------------------------
'
Private Function TrackSizes(ByRef aBlocks() As Long, lngCount As Long, _
    eAxis As eGeometryAxis) As Dictionary

    Dim dObserved As Dictionary
    Dim dSizes As Dictionary
    Dim colObs As Collection
    Dim lngStart As Long
    Dim lngEnd As Long
    Dim lngSize As Long
    Dim lngI As Long
    Dim varIndex As Variant

    Set dObserved = New Dictionary
    For lngI = 0 To lngCount - 1
        lngStart = GetTrackStart(aBlocks(lngI), eAxis)
        lngEnd = GetTrackEnd(aBlocks(lngI), eAxis, lngStart)
        If lngStart = lngEnd Then
            If TryGetSize(aBlocks(lngI), eAxis, lngSize) Then
                If dObserved.Exists(lngStart) Then
                    Set colObs = dObserved(lngStart)
                Else
                    Set colObs = New Collection
                    dObserved.Add lngStart, colObs
                End If
                colObs.Add lngSize
            End If
        End If
    Next lngI

    Set dSizes = New Dictionary
    For Each varIndex In dObserved.Keys
        Set colObs = dObserved(varIndex)
        dSizes(CLng(varIndex)) = SnapTwips(MedianLong(colObs))
    Next varIndex
    Set TrackSizes = dSizes

End Function


'---------------------------------------------------------------------------------------
' Procedure : BoundaryPositions
'---------------------------------------------------------------------------------------
'
Private Function BoundaryPositions(ByRef aBlocks() As Long, lngCount As Long, _
    eAxis As eGeometryAxis) As Dictionary

    Dim dObserved As Dictionary
    Dim dRaw As Dictionary
    Dim dPositions As Dictionary
    Dim colObs As Collection
    Dim lngPosition As Long
    Dim lngStarts() As Long
    Dim lngPrev As Long
    Dim lngCurrent As Long
    Dim lngI As Long

    Set dObserved = New Dictionary
    For lngI = 0 To lngCount - 1
        If TryGetPos(aBlocks(lngI), eAxis, lngPosition) Then
            lngCurrent = GetTrackStart(aBlocks(lngI), eAxis)
            If dObserved.Exists(lngCurrent) Then
                Set colObs = dObserved(lngCurrent)
            Else
                Set colObs = New Collection
                dObserved.Add lngCurrent, colObs
            End If
            colObs.Add lngPosition
        End If
    Next lngI
    If dObserved.Count = 0 Then Exit Function

    lngStarts = SortedKeys(dObserved)
    Set dRaw = New Dictionary
    For lngI = 0 To UBound(lngStarts)
        Set colObs = dObserved(lngStarts(lngI))
        dRaw(lngStarts(lngI)) = MedianLong(colObs)
    Next lngI

    Set dPositions = New Dictionary
    dPositions(lngStarts(0)) = SnapTwips(CLng(dRaw(lngStarts(0))))
    For lngI = 0 To UBound(lngStarts) - 1
        lngPrev = lngStarts(lngI)
        lngCurrent = lngStarts(lngI + 1)
        dPositions(lngCurrent) = CLng(dPositions(lngPrev)) + _
            SnapTwips(CLng(dRaw(lngCurrent)) - CLng(dRaw(lngPrev)))
    Next lngI
    Set BoundaryPositions = dPositions

End Function


'---------------------------------------------------------------------------------------
' Procedure : InferEndSizes
'---------------------------------------------------------------------------------------
'
Private Sub InferEndSizes(ByRef aBlocks() As Long, lngCount As Long, _
    eAxis As eGeometryAxis, dPositions As Dictionary, dSizes As Dictionary)

    Dim dInferred As Dictionary
    Dim colObs As Collection
    Dim lngStart As Long
    Dim lngEnd As Long
    Dim lngSize As Long
    Dim lngI As Long
    Dim varIndex As Variant

    Set dInferred = New Dictionary
    For lngI = 0 To lngCount - 1
        lngStart = GetTrackStart(aBlocks(lngI), eAxis)
        lngEnd = GetTrackEnd(aBlocks(lngI), eAxis, lngStart)
        If lngStart <> lngEnd Then
            If TryGetSize(aBlocks(lngI), eAxis, lngSize) Then
                If dPositions.Exists(lngStart) And dPositions.Exists(lngEnd) Then
                    If dInferred.Exists(lngEnd) Then
                        Set colObs = dInferred(lngEnd)
                    Else
                        Set colObs = New Collection
                        dInferred.Add lngEnd, colObs
                    End If
                    colObs.Add lngSize - _
                        (CLng(dPositions(lngEnd)) - CLng(dPositions(lngStart)))
                End If
            End If
        End If
    Next lngI

    For Each varIndex In dInferred.Keys
        If Not dSizes.Exists(varIndex) Then
            Set colObs = dInferred(varIndex)
            dSizes(CLng(varIndex)) = SnapTwips(MedianLong(colObs))
        End If
    Next varIndex

End Sub


'---------------------------------------------------------------------------------------
' Procedure : DeriveSpan
'---------------------------------------------------------------------------------------
'
Private Function DeriveSpan(lngStart As Long, lngEnd As Long, _
    dPositions As Dictionary, dSizes As Dictionary) As Variant

    If dPositions.Exists(lngStart) And dPositions.Exists(lngEnd) And _
        dSizes.Exists(lngEnd) Then

        DeriveSpan = CLng(dPositions(lngEnd)) + CLng(dSizes(lngEnd)) - _
            CLng(dPositions(lngStart))
    Else
        DeriveSpan = Null
    End If

End Function


'---------------------------------------------------------------------------------------
' Procedure : CanonicalizeTabs
'---------------------------------------------------------------------------------------
'
Private Sub CanonicalizeTabs()

    Dim lngBlock As Long

    For lngBlock = 0 To m_lngBlockCount - 1
        With m_aBlocks(lngBlock)
            If .strType = "Tab" Then
                If Not (.blnLayoutGroup Or .blnGroupTable) Then
                    If .lngWidthLine >= 0 Then
                        RewriteProp lngBlock, egpWidth, SnapTwips(.lngWidth)
                    End If
                    If .lngHeightLine >= 0 Then
                        RewriteProp lngBlock, egpHeight, SnapTwips(.lngHeight)
                    End If
                End If
            End If
        End With
    Next lngBlock

End Sub


'---------------------------------------------------------------------------------------
' Procedure : GrowEnvelopes
'---------------------------------------------------------------------------------------
'
Private Sub GrowEnvelopes()

    Dim aSectionBottom() As Long
    Dim lngBlock As Long
    Dim lngSection As Long
    Dim lngExtent As Long
    Dim lngMaxRight As Long
    Dim lngGrown As Long

    If this.lngRoot < 0 Or Not m_blnHasLayout Then Exit Sub
    ReDim aSectionBottom(0 To m_lngBlockCount - 1)

    ' Gather form and section extents in one pass. Section indices map directly into
    ' the typed block array, so no child index Dictionary or per-section Collection
    ' is needed.
    For lngBlock = 0 To m_lngBlockCount - 1
        With m_aBlocks(lngBlock)
            If Not IsEnvelopeContainer(.strType) Then
                lngMaxRight = MaxLong(lngMaxRight, ChildExtent(lngBlock, egaHorizontal))
            End If
            lngSection = .lngSection
            If lngSection >= 0 And .strType <> "Tab" And .strType <> "Page" _
                And .lngTab < 0 Then

                lngExtent = ChildExtent(lngBlock, egaVertical)
                If lngExtent > aSectionBottom(lngSection) Then
                    aSectionBottom(lngSection) = lngExtent
                End If
            End If
        End With
    Next lngBlock

    With m_aBlocks(this.lngRoot)
        If .lngWidthLine >= 0 Then
            lngGrown = MaxLong(SnapTwips(.lngWidth), SnapUp(lngMaxRight))
            RewriteProp this.lngRoot, egpWidth, lngGrown
        End If
    End With

    For lngBlock = 0 To m_lngBlockCount - 1
        With m_aBlocks(lngBlock)
            If IsSectionType(.strType) And .lngHeightLine >= 0 Then
                If aSectionBottom(lngBlock) > 0 Then
                    lngGrown = MaxLong(SnapTwips(.lngHeight), _
                        SnapUp(aSectionBottom(lngBlock)))
                Else
                    lngGrown = SnapTwips(.lngHeight)
                End If
                RewriteProp lngBlock, egpHeight, lngGrown
            End If
        End With
    Next lngBlock

End Sub


'=======================================================================================
' Helpers
'=======================================================================================

Private Function SnapTwips(ByVal lngValue As Long) As Long
    SnapTwips = CLng(Round(CDbl(lngValue) / GRID_TWIPS, 0)) * GRID_TWIPS
End Function

Private Function SnapUp(ByVal lngValue As Long) As Long
    Dim lngSnapped As Long
    lngSnapped = SnapTwips(lngValue)
    If lngSnapped >= lngValue Then
        SnapUp = lngSnapped
    Else
        SnapUp = lngSnapped + GRID_TWIPS
    End If
End Function

Private Function IsSectionType(strType As String) As Boolean
    Select Case strType
        Case "Section", "FormHeader", "FormFooter", "PageHeader", "PageFooter"
            IsSectionType = True
    End Select
End Function

Private Function IsEnvelopeContainer(strType As String) As Boolean
    Select Case strType
        Case "Form", "Report", "Tab", "Page"
            IsEnvelopeContainer = True
    End Select
End Function

Private Function IsInlineBinary(strTrimmed As String) As Boolean
    IsInlineBinary = EndsWith(strTrimmed, " = Begin") And Not StartsWith(strTrimmed, "Begin")
End Function

Private Function IsIntegerToken(strValue As String) As Boolean
    If Len(strValue) = 0 Then Exit Function
    If InStr(strValue, ".") > 0 Then Exit Function
    If InStr(strValue, """") > 0 Then Exit Function
    IsIntegerToken = IsNumeric(strValue)
End Function

Private Function GetTrackStart(lngBlock As Long, eAxis As eGeometryAxis) As Long
    With m_aBlocks(lngBlock)
        If eAxis = egaHorizontal Then
            If .blnColumnStart Then GetTrackStart = .lngColumnStart
        Else
            If .blnRowStart Then GetTrackStart = .lngRowStart
        End If
    End With
End Function

Private Function GetTrackEnd(lngBlock As Long, eAxis As eGeometryAxis, _
    lngDefault As Long) As Long

    With m_aBlocks(lngBlock)
        If eAxis = egaHorizontal Then
            If .blnColumnEnd Then
                GetTrackEnd = .lngColumnEnd
            Else
                GetTrackEnd = lngDefault
            End If
        ElseIf .blnRowEnd Then
            GetTrackEnd = .lngRowEnd
        Else
            GetTrackEnd = lngDefault
        End If
    End With

End Function

Private Function HasPosition(lngBlock As Long, eAxis As eGeometryAxis) As Boolean
    If eAxis = egaHorizontal Then
        HasPosition = (m_aBlocks(lngBlock).lngLeftLine >= 0)
    Else
        HasPosition = (m_aBlocks(lngBlock).lngTopLine >= 0)
    End If
End Function

Private Function HasSize(lngBlock As Long, eAxis As eGeometryAxis) As Boolean
    If eAxis = egaHorizontal Then
        HasSize = (m_aBlocks(lngBlock).lngWidthLine >= 0)
    Else
        HasSize = (m_aBlocks(lngBlock).lngHeightLine >= 0)
    End If
End Function

Private Function TryGetSize(lngBlock As Long, eAxis As eGeometryAxis, _
    ByRef lngSize As Long) As Boolean

    With m_aBlocks(lngBlock)
        If eAxis = egaHorizontal Then
            If .lngWidthLine >= 0 Then
                lngSize = .lngWidth
                TryGetSize = True
            ElseIf .blnCachedWidth And .lngLeftLine >= 0 Then
                lngSize = .lngCachedWidth - .lngLeft
                TryGetSize = True
            End If
        Else
            If .lngHeightLine >= 0 Then
                lngSize = .lngHeight
                TryGetSize = True
            ElseIf .blnCachedHeight And .lngTopLine >= 0 Then
                lngSize = .lngCachedHeight - .lngTop
                TryGetSize = True
            End If
        End If
    End With
End Function

Private Function TryGetPos(lngBlock As Long, eAxis As eGeometryAxis, _
    ByRef lngPosition As Long) As Boolean

    With m_aBlocks(lngBlock)
        If eAxis = egaHorizontal Then
            If .lngLeftLine >= 0 Then
                lngPosition = .lngLeft
                TryGetPos = True
            ElseIf .blnCachedLeft Then
                lngPosition = .lngCachedLeft
                TryGetPos = True
            End If
        Else
            If .lngTopLine >= 0 Then
                lngPosition = .lngTop
                TryGetPos = True
            ElseIf .blnCachedTop Then
                lngPosition = .lngCachedTop
                TryGetPos = True
            End If
        End If
    End With
End Function

Private Function ChildExtent(lngBlock As Long, eAxis As eGeometryAxis) As Long
    With m_aBlocks(lngBlock)
        If eAxis = egaHorizontal Then
            If .lngLeftLine >= 0 And .lngWidthLine >= 0 Then
                ChildExtent = .lngLeft + .lngWidth
            End If
        ElseIf .lngTopLine >= 0 And .lngHeightLine >= 0 Then
            ChildExtent = .lngTop + .lngHeight
        End If
    End With
End Function

Private Sub RewriteProp(lngBlock As Long, eProperty As eGeometryProperty, lngValue As Long)

    Dim lngLine As Long
    Dim strLine As String
    Dim strName As String
    Dim lngCurrent As Long
    Dim intIndent As Integer

    With m_aBlocks(lngBlock)
        Select Case eProperty
            Case egpLeft
                strName = "Left"
                lngLine = .lngLeftLine
                lngCurrent = .lngLeft
            Case egpTop
                strName = "Top"
                lngLine = .lngTopLine
                lngCurrent = .lngTop
            Case egpWidth
                strName = "Width"
                lngLine = .lngWidthLine
                lngCurrent = .lngWidth
            Case egpHeight
                strName = "Height"
                lngLine = .lngHeightLine
                lngCurrent = .lngHeight
        End Select
    End With
    If lngLine < 0 Or lngCurrent = lngValue Then Exit Sub

    strLine = m_strLines(lngLine)
    intIndent = Len(strLine) - Len(LTrim$(strLine))
    m_strLines(lngLine) = Space$(intIndent) & strName & " =" & lngValue
    With m_aBlocks(lngBlock)
        Select Case eProperty
            Case egpLeft:   .lngLeft = lngValue
            Case egpTop:    .lngTop = lngValue
            Case egpWidth:  .lngWidth = lngValue
            Case egpHeight: .lngHeight = lngValue
        End Select
    End With
    this.lngChanged = this.lngChanged + 1
    Me.RewriteCalls = Me.RewriteCalls + 1

End Sub

Private Function MedianLong(colValues As Collection) As Long
    Dim lngValues() As Long
    Dim lngCount As Long
    Dim lngI As Long
    Dim lngJ As Long
    Dim lngSwap As Long
    Dim dblMid As Double
    Dim varItem As Variant

    lngCount = colValues.Count
    If lngCount = 0 Then Exit Function
    ReDim lngValues(1 To lngCount)
    lngI = 1
    For Each varItem In colValues
        lngValues(lngI) = CLng(varItem)
        lngI = lngI + 1
    Next varItem
    For lngI = 1 To lngCount - 1
        For lngJ = lngI + 1 To lngCount
            If lngValues(lngJ) < lngValues(lngI) Then
                lngSwap = lngValues(lngI)
                lngValues(lngI) = lngValues(lngJ)
                lngValues(lngJ) = lngSwap
            End If
        Next lngJ
    Next lngI
    If lngCount Mod 2 = 1 Then
        MedianLong = lngValues((lngCount + 1) \ 2)
    Else
        dblMid = (CDbl(lngValues(lngCount \ 2)) + CDbl(lngValues(lngCount \ 2 + 1))) / 2
        MedianLong = CLng(Round(dblMid, 0))
    End If
End Function

Private Function SortedKeys(dMap As Dictionary) As Long()
    Dim varKey As Variant
    Dim lngI As Long
    Dim lngJ As Long
    Dim lngSwap As Long
    Dim lngKeys() As Long

    ReDim lngKeys(0 To dMap.Count - 1)
    lngI = 0
    For Each varKey In dMap.Keys
        lngKeys(lngI) = CLng(varKey)
        lngI = lngI + 1
    Next varKey
    For lngI = 0 To dMap.Count - 2
        For lngJ = lngI + 1 To dMap.Count - 1
            If lngKeys(lngJ) < lngKeys(lngI) Then
                lngSwap = lngKeys(lngI)
                lngKeys(lngI) = lngKeys(lngJ)
                lngKeys(lngJ) = lngSwap
            End If
        Next lngJ
    Next lngI
    SortedKeys = lngKeys
End Function

Private Function MaxLong(lngA As Long, lngB As Long) As Long
    If lngA > lngB Then MaxLong = lngA Else MaxLong = lngB
End Function

Private Sub WarnSkip(strReason As String)
    Log.Error eelWarning, T("Form geometry: skipped {0} on {1}", _
        var0:=strReason, var1:=Nz2(this.strFormName, "form")), _
        ModuleNameConst & ".Canonicalize"
End Sub
