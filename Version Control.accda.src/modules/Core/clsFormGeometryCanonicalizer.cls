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
'           : Performance note: block properties are reached through typed Dictionary
'           : locals rather than chained expressions like dBlock("Props")(strName).
'           : Dictionary.Item returns a Variant, so a chained access dispatches late on
'           : every call, which measured roughly seven times slower across the parse.
'           : Keep the hoisted locals when editing these routines.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
'@Folder("Core")

Private Const GRID_TWIPS As Long = 60
Private Const ModuleNameConst As String = "clsFormGeometryCanonicalizer"

' A single planned geometry rewrite. Held in an array rather than a Dictionary so the
' planning passes do not allocate an object per change.
Private Type udtChange
    dBlock As Dictionary
    dProps As Dictionary
    strName As String
    lngValue As Long
End Type

Private Type udtThis
    colBlocks As Collection
    dRoot As Dictionary
    strFormName As String
    lngChanged As Long
End Type
Private this As udtThis

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
    Set this.colBlocks = New Collection
    Set this.dRoot = Nothing
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
    Me.BlockCount = this.colBlocks.Count
    Me.ChangedLines = this.lngChanged
    If this.lngChanged > 0 Then
        For lngLine = 0 To m_lngUpper
            varLines(lngLine) = m_strLines(lngLine)
        Next lngLine
    End If
    ReleaseBlocks
    Set this.colBlocks = Nothing
    Set this.dRoot = Nothing
    Erase m_strLines
    Me.PhaseCopyOut = CDbl(Perf.MicroTimer - curPhase)
    Perf.OperationEnd

    CatchAny eelError, T("Error canonicalizing form geometry in {0}", _
        var0:=Nz2(strFormName, "form")), ModuleNameConst & ".Canonicalize", True, True

End Sub


'---------------------------------------------------------------------------------------
' Procedure : ReleaseBlocks
' Author    : Adam Waller
' Date      : 9/9/2026
' Purpose   : Break the parent/child reference cycles in the block graph. Dictionary and
'           : Collection are reference counted, so a parent holding its children while
'           : each child holds its parent would never be reclaimed, leaking the whole
'           : graph for every form exported.
'---------------------------------------------------------------------------------------
'
Private Sub ReleaseBlocks()

    Dim dBlock As Dictionary

    If this.colBlocks Is Nothing Then Exit Sub
    For Each dBlock In this.colBlocks
        If dBlock.Exists("Parent") Then dBlock.Remove "Parent"
        If dBlock.Exists("Children") Then dBlock.Remove "Children"
        If dBlock.Exists("Section") Then dBlock.Remove "Section"
        If dBlock.Exists("Tab") Then dBlock.Remove "Tab"
    Next dBlock

End Sub


'---------------------------------------------------------------------------------------
' Procedure : ParseDocument
'---------------------------------------------------------------------------------------
'
Private Sub ParseDocument()

    Dim lngLine As Long
    Dim strTrimmed As String
    Dim dBlock As Dictionary

    lngLine = 0
    Do While lngLine <= m_lngUpper
        strTrimmed = Trim$(m_strLines(lngLine))
        If Len(strTrimmed) = 14 Then
            If strTrimmed = "CodeBehindForm" Then Exit Do
        End If
        If StartsWith(strTrimmed, "Begin") Then
            Set dBlock = ParseBlock(lngLine + 1, Mid$(strTrimmed, 6), lngLine, Nothing)
            Set this.dRoot = dBlock
            lngLine = dBlock("NextLine")
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
    lngStart As Long, dParent As Dictionary) As Dictionary

    Dim dBlock As Dictionary
    Dim dProps As Dictionary
    Dim dPropLines As Dictionary
    Dim strTrimmed As String
    Dim strName As String
    Dim strValue As String
    Dim lngPos As Long
    Dim lngLen As Long
    Dim lngUpper As Long
    Dim blnIsEnd As Boolean
    Dim dChild As Dictionary

    lngUpper = m_lngUpper
    Set dBlock = NewBlock(Trim$(strType), lngStart, dParent)
    Set dProps = dBlock("Props")
    Set dPropLines = dBlock("PropLines")

    Do While lngLine <= lngUpper
        strTrimmed = Trim$(m_strLines(lngLine))
        lngLen = Len(strTrimmed)

        ' Compare lengths before contents. Option Compare Database makes every string
        ' comparison locale aware, and an inline binary payload such as an embedded
        ' PictureData image can run to tens of thousands of long hex lines. Collating
        ' each of those against a short keyword dominated the parse on such forms.
        ' VBA does not short-circuit And, so these gates have to nest.
        If lngLen = 14 Then
            If strTrimmed = "CodeBehindForm" Then
                dBlock("EndLine") = lngLine
                dBlock("NextLine") = lngUpper + 1
                Set ParseBlock = dBlock
                Exit Function
            End If
        End If
        blnIsEnd = False
        If lngLen = 3 Then blnIsEnd = (strTrimmed = "End")

        If IsInlineBinary(strTrimmed) Then
            lngLine = lngLine + 1
            Do While lngLine <= lngUpper
                strTrimmed = Trim$(m_strLines(lngLine))
                If Len(strTrimmed) = 3 Then
                    If strTrimmed = "End" Then Exit Do
                End If
                lngLine = lngLine + 1
            Loop
            lngLine = lngLine + 1
        ElseIf StartsWith(strTrimmed, "Begin") Then
            Set dChild = ParseBlock(lngLine + 1, Mid$(strTrimmed, 6), lngLine, dBlock)
            lngLine = dChild("NextLine")
        ElseIf blnIsEnd Then
            dBlock("EndLine") = lngLine
            dBlock("NextLine") = lngLine + 1
            Set ParseBlock = dBlock
            Exit Function
        Else
            lngPos = InStr(strTrimmed, " =")
            If lngPos > 1 Then
                strName = Left$(strTrimmed, lngPos - 1)
                strValue = Mid$(strTrimmed, lngPos + 2)
                If strName = "Name" And StartsWith(strValue, """") Then
                    dBlock("Name") = Mid$(strValue, 2, Len(strValue) - 2)
                ElseIf IsIntegerToken(strValue) Then
                    dProps(strName) = CLng(strValue)
                    dPropLines(strName) = lngLine
                End If
            End If
            lngLine = lngLine + 1
        End If
    Loop

    dBlock("EndLine") = lngLine
    dBlock("NextLine") = lngLine
    Set ParseBlock = dBlock

End Function


'---------------------------------------------------------------------------------------
' Procedure : NewBlock
'---------------------------------------------------------------------------------------
'
Private Function NewBlock(strType As String, lngStart As Long, dParent As Dictionary) As Dictionary

    Dim dBlock As Dictionary
    Dim colChildren As Collection
    Dim strParentType As String

    Set dBlock = New Dictionary
    dBlock("Type") = strType
    dBlock("Start") = lngStart
    dBlock("Name") = vbNullString
    Set dBlock("Props") = New Dictionary
    Set dBlock("PropLines") = New Dictionary
    Set dBlock("Children") = New Collection
    Set dBlock("Parent") = dParent
    If Not dParent Is Nothing Then
        Set colChildren = dParent("Children")
        colChildren.Add dBlock
        strParentType = dParent("Type")
        If IsSectionType(strParentType) Then
            Set dBlock("Section") = dParent
        ElseIf dParent.Exists("Section") Then
            Set dBlock("Section") = dParent("Section")
        End If
        If strParentType = "Tab" Then
            Set dBlock("Tab") = dParent
        ElseIf dParent.Exists("Tab") Then
            Set dBlock("Tab") = dParent("Tab")
        End If
    End If
    this.colBlocks.Add dBlock
    Set NewBlock = dBlock

End Function


'---------------------------------------------------------------------------------------
' Procedure : CanonicalizeGroups
'---------------------------------------------------------------------------------------
'
Private Sub CanonicalizeGroups()

    Dim dGroups As Dictionary
    Dim dBlock As Dictionary
    Dim dProps As Dictionary
    Dim colGroup As Collection
    Dim strKey As String
    Dim varKey As Variant

    Set dGroups = New Dictionary
    For Each dBlock In this.colBlocks
        Set dProps = dBlock("Props")
        If dProps.Exists("LayoutGroup") Or dProps.Exists("GroupTable") Then
            strKey = PropOrDefault(dProps, "GroupTable", -1) & "/" & _
                PropOrDefault(dProps, "LayoutGroup", -1)
            If dGroups.Exists(strKey) Then
                Set colGroup = dGroups(strKey)
            Else
                Set colGroup = New Collection
                dGroups.Add strKey, colGroup
            End If
            colGroup.Add dBlock
        End If
    Next dBlock

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
    Dim dBlock As Dictionary
    Dim dSectionRef As Dictionary
    Dim colSection As Collection
    Dim strSection As String
    Dim varSection As Variant
    Dim lngI As Long

    ReDim aChanges(0 To 63)
    lngChanges = 0

    If Not PlanAxis(colCells, "x", "ColumnStart", "ColumnEnd", "Left", "Width", _
        aChanges, lngChanges) Then
        WarnSkip strLabel & " (horizontal)"
        Exit Sub
    End If

    ' Vertical geometry is planned per section, since rows restart in each one.
    Set dBySection = New Dictionary
    For Each dBlock In colCells
        If dBlock.Exists("Section") Then
            Set dSectionRef = dBlock("Section")
            strSection = CStr(ObjPtr(dSectionRef))
        Else
            strSection = "none"
        End If
        If dBySection.Exists(strSection) Then
            Set colSection = dBySection(strSection)
        Else
            Set colSection = New Collection
            dBySection.Add strSection, colSection
        End If
        colSection.Add dBlock
    Next dBlock

    For Each varSection In dBySection.Keys
        Set colSection = dBySection(varSection)
        If Not PlanAxis(colSection, "y", "RowStart", "RowEnd", "Top", "Height", _
            aChanges, lngChanges) Then
            WarnSkip strLabel & " (vertical)"
            Exit Sub
        End If
    Next varSection

    For lngI = 0 To lngChanges - 1
        RewriteProp aChanges(lngI).dBlock, aChanges(lngI).dProps, _
            aChanges(lngI).strName, aChanges(lngI).lngValue
    Next lngI

End Sub


'---------------------------------------------------------------------------------------
' Procedure : PlanAxis
' Purpose   : Append the planned rewrites for one axis to aChanges. Returns False when
'           : the group cannot be solved, in which case no change is applied at all.
'---------------------------------------------------------------------------------------
'
Private Function PlanAxis(colCells As Collection, strAxis As String, _
    strStartName As String, strEndName As String, strPosName As String, _
    strSizeName As String, ByRef aChanges() As udtChange, _
    ByRef lngChanges As Long) As Boolean

    Dim aProps() As Dictionary
    Dim aBlocks() As Dictionary
    Dim lngCount As Long
    Dim dSizes As Dictionary
    Dim dPositions As Dictionary
    Dim dProps As Dictionary
    Dim dBlock As Dictionary
    Dim lngStart As Long
    Dim lngEnd As Long
    Dim lngI As Long
    Dim varDerived As Variant

    Me.PlanAxisCalls = Me.PlanAxisCalls + 1

    ' Hoist each cell's Props once, so the passes below never chain through a Variant.
    lngCount = colCells.Count
    If lngCount = 0 Then
        PlanAxis = True
        Exit Function
    End If
    ReDim aBlocks(0 To lngCount - 1)
    ReDim aProps(0 To lngCount - 1)
    lngI = 0
    For Each dBlock In colCells
        Set aBlocks(lngI) = dBlock
        Set aProps(lngI) = dBlock("Props")
        lngI = lngI + 1
    Next dBlock

    Set dSizes = TrackSizes(aProps, lngCount, strAxis, strStartName, strEndName)
    Set dPositions = BoundaryPositions(aProps, lngCount, strAxis, strStartName)
    If dPositions Is Nothing Then Exit Function
    InferEndSizes aProps, lngCount, strAxis, strStartName, strEndName, dPositions, dSizes

    For lngI = 0 To lngCount - 1
        Set dProps = aProps(lngI)
        lngStart = PropOrDefault(dProps, strStartName, 0)
        lngEnd = PropOrDefault(dProps, strEndName, lngStart)
        If dProps.Exists(strPosName) Then
            If Not dPositions.Exists(lngStart) Then Exit Function
            AddChange aChanges, lngChanges, aBlocks(lngI), dProps, strPosName, _
                CLng(dPositions(lngStart))
        End If
        If dProps.Exists(strSizeName) Then
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
            AddChange aChanges, lngChanges, aBlocks(lngI), dProps, strSizeName, _
                CLng(varDerived)
        End If
    Next lngI

    PlanAxis = True

End Function


'---------------------------------------------------------------------------------------
' Procedure : AddChange
'---------------------------------------------------------------------------------------
'
Private Sub AddChange(ByRef aChanges() As udtChange, ByRef lngChanges As Long, _
    dBlock As Dictionary, dProps As Dictionary, strName As String, lngValue As Long)

    If lngChanges > UBound(aChanges) Then
        ReDim Preserve aChanges(0 To (UBound(aChanges) + 1) * 2 - 1)
    End If
    Set aChanges(lngChanges).dBlock = dBlock
    Set aChanges(lngChanges).dProps = dProps
    aChanges(lngChanges).strName = strName
    aChanges(lngChanges).lngValue = lngValue
    lngChanges = lngChanges + 1

End Sub


'---------------------------------------------------------------------------------------
' Procedure : TrackSizes
'---------------------------------------------------------------------------------------
'
Private Function TrackSizes(ByRef aProps() As Dictionary, lngCount As Long, _
    strAxis As String, strStartName As String, strEndName As String) As Dictionary

    Dim dObserved As Dictionary
    Dim dSizes As Dictionary
    Dim dProps As Dictionary
    Dim colObs As Collection
    Dim lngStart As Long
    Dim lngEnd As Long
    Dim lngI As Long
    Dim varSize As Variant
    Dim varIndex As Variant

    Set dObserved = New Dictionary
    For lngI = 0 To lngCount - 1
        Set dProps = aProps(lngI)
        lngStart = PropOrDefault(dProps, strStartName, 0)
        lngEnd = PropOrDefault(dProps, strEndName, lngStart)
        varSize = GetSize(dProps, strAxis)
        If lngStart = lngEnd And Not IsNull(varSize) Then
            If dObserved.Exists(lngStart) Then
                Set colObs = dObserved(lngStart)
            Else
                Set colObs = New Collection
                dObserved.Add lngStart, colObs
            End If
            colObs.Add CLng(varSize)
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
Private Function BoundaryPositions(ByRef aProps() As Dictionary, lngCount As Long, _
    strAxis As String, strStartName As String) As Dictionary

    Dim dObserved As Dictionary
    Dim dRaw As Dictionary
    Dim dPositions As Dictionary
    Dim dProps As Dictionary
    Dim colObs As Collection
    Dim varPos As Variant
    Dim lngStarts() As Long
    Dim lngPrev As Long
    Dim lngCurrent As Long
    Dim lngI As Long

    Set dObserved = New Dictionary
    For lngI = 0 To lngCount - 1
        Set dProps = aProps(lngI)
        varPos = GetPos(dProps, strAxis)
        If Not IsNull(varPos) Then
            lngCurrent = PropOrDefault(dProps, strStartName, 0)
            If dObserved.Exists(lngCurrent) Then
                Set colObs = dObserved(lngCurrent)
            Else
                Set colObs = New Collection
                dObserved.Add lngCurrent, colObs
            End If
            colObs.Add CLng(varPos)
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
Private Sub InferEndSizes(ByRef aProps() As Dictionary, lngCount As Long, _
    strAxis As String, strStartName As String, strEndName As String, _
    dPositions As Dictionary, dSizes As Dictionary)

    Dim dInferred As Dictionary
    Dim dProps As Dictionary
    Dim colObs As Collection
    Dim lngStart As Long
    Dim lngEnd As Long
    Dim lngI As Long
    Dim varSize As Variant
    Dim varIndex As Variant

    Set dInferred = New Dictionary
    For lngI = 0 To lngCount - 1
        Set dProps = aProps(lngI)
        lngStart = PropOrDefault(dProps, strStartName, 0)
        lngEnd = PropOrDefault(dProps, strEndName, lngStart)
        varSize = GetSize(dProps, strAxis)
        If (Not IsNull(varSize)) And lngStart <> lngEnd Then
            If dPositions.Exists(lngStart) And dPositions.Exists(lngEnd) Then
                If dInferred.Exists(lngEnd) Then
                    Set colObs = dInferred(lngEnd)
                Else
                    Set colObs = New Collection
                    dInferred.Add lngEnd, colObs
                End If
                colObs.Add CLng(varSize) - (CLng(dPositions(lngEnd)) - CLng(dPositions(lngStart)))
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

    If dPositions.Exists(lngStart) And dPositions.Exists(lngEnd) And dSizes.Exists(lngEnd) Then
        DeriveSpan = CLng(dPositions(lngEnd)) + CLng(dSizes(lngEnd)) - CLng(dPositions(lngStart))
    Else
        DeriveSpan = Null
    End If

End Function


'---------------------------------------------------------------------------------------
' Procedure : CanonicalizeTabs
'---------------------------------------------------------------------------------------
'
Private Sub CanonicalizeTabs()

    Dim dBlock As Dictionary
    Dim dProps As Dictionary

    For Each dBlock In this.colBlocks
        If dBlock("Type") = "Tab" Then
            Set dProps = dBlock("Props")
            If Not (dProps.Exists("LayoutGroup") Or dProps.Exists("GroupTable")) Then
                If dProps.Exists("Width") Then
                    RewriteProp dBlock, dProps, "Width", SnapTwips(CLng(dProps("Width")))
                End If
                If dProps.Exists("Height") Then
                    RewriteProp dBlock, dProps, "Height", SnapTwips(CLng(dProps("Height")))
                End If
            End If
        End If
    Next dBlock

End Sub


'---------------------------------------------------------------------------------------
' Procedure : GrowEnvelopes
'---------------------------------------------------------------------------------------
'
Private Sub GrowEnvelopes()

    Dim dBlock As Dictionary
    Dim dProps As Dictionary
    Dim dSection As Dictionary
    Dim dSectionRef As Dictionary
    Dim dRootProps As Dictionary
    Dim dChildren As Dictionary
    Dim colChildren As Collection
    Dim strSection As String
    Dim lngMaxRight As Long
    Dim lngMaxBottom As Long
    Dim lngGrown As Long
    Dim blnHasLayout As Boolean

    If this.dRoot Is Nothing Then Exit Sub

    For Each dBlock In this.colBlocks
        Set dProps = dBlock("Props")
        If dProps.Exists("LayoutGroup") Or dProps.Exists("GroupTable") Then
            blnHasLayout = True
            Exit For
        End If
    Next dBlock
    If Not blnHasLayout Then Exit Sub

    ' Index children by section in one pass, rather than rescanning every block for
    ' each section.
    Set dChildren = New Dictionary
    For Each dBlock In this.colBlocks
        If dBlock.Exists("Section") Then
            Set dSectionRef = dBlock("Section")
            strSection = CStr(ObjPtr(dSectionRef))
            If dChildren.Exists(strSection) Then
                Set colChildren = dChildren(strSection)
            Else
                Set colChildren = New Collection
                dChildren.Add strSection, colChildren
            End If
            colChildren.Add dBlock
        End If
    Next dBlock

    For Each dBlock In this.colBlocks
        If Not IsEnvelopeContainer(dBlock("Type")) Then
            Set dProps = dBlock("Props")
            lngMaxRight = MaxLong(lngMaxRight, ChildExtent(dProps, "Left", "Width"))
        End If
    Next dBlock
    Set dRootProps = this.dRoot("Props")
    If dRootProps.Exists("Width") Then
        lngGrown = MaxLong(SnapTwips(CLng(dRootProps("Width"))), SnapUp(lngMaxRight))
        RewriteProp this.dRoot, dRootProps, "Width", lngGrown
    End If

    For Each dSection In this.colBlocks
        If IsSectionType(dSection("Type")) Then
            Set dProps = dSection("Props")
            If dProps.Exists("Height") Then
                lngMaxBottom = 0
                strSection = CStr(ObjPtr(dSection))
                If dChildren.Exists(strSection) Then
                    Set colChildren = dChildren(strSection)
                    For Each dBlock In colChildren
                        If dBlock("Type") <> "Tab" And dBlock("Type") <> "Page" _
                            And Not dBlock.Exists("Tab") Then
                            lngMaxBottom = MaxLong(lngMaxBottom, _
                                ChildExtent(dBlock("Props"), "Top", "Height"))
                        End If
                    Next dBlock
                End If
                If lngMaxBottom > 0 Then
                    lngGrown = MaxLong(SnapTwips(CLng(dProps("Height"))), SnapUp(lngMaxBottom))
                Else
                    lngGrown = SnapTwips(CLng(dProps("Height")))
                End If
                RewriteProp dSection, dProps, "Height", lngGrown
            End If
        End If
    Next dSection

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

Private Function PropOrDefault(dProps As Dictionary, strName As String, lngDefault As Long) As Long
    If dProps.Exists(strName) Then
        PropOrDefault = CLng(dProps(strName))
    Else
        PropOrDefault = lngDefault
    End If
End Function

Private Function GetSize(dProps As Dictionary, strAxis As String) As Variant
    If strAxis = "x" Then
        If dProps.Exists("Width") Then
            GetSize = CLng(dProps("Width"))
        ElseIf dProps.Exists("LayoutCachedWidth") And dProps.Exists("Left") Then
            GetSize = CLng(dProps("LayoutCachedWidth")) - CLng(dProps("Left"))
        Else
            GetSize = Null
        End If
    Else
        If dProps.Exists("Height") Then
            GetSize = CLng(dProps("Height"))
        ElseIf dProps.Exists("LayoutCachedHeight") And dProps.Exists("Top") Then
            GetSize = CLng(dProps("LayoutCachedHeight")) - CLng(dProps("Top"))
        Else
            GetSize = Null
        End If
    End If
End Function

Private Function GetPos(dProps As Dictionary, strAxis As String) As Variant
    Dim strKey As String
    Dim strCache As String
    If strAxis = "x" Then
        strKey = "Left"
        strCache = "LayoutCachedLeft"
    Else
        strKey = "Top"
        strCache = "LayoutCachedTop"
    End If
    If dProps.Exists(strKey) Then
        GetPos = CLng(dProps(strKey))
    ElseIf dProps.Exists(strCache) Then
        GetPos = CLng(dProps(strCache))
    Else
        GetPos = Null
    End If
End Function

Private Function ChildExtent(dProps As Dictionary, strPos As String, strSize As String) As Long
    If dProps.Exists(strPos) And dProps.Exists(strSize) Then
        ChildExtent = CLng(dProps(strPos)) + CLng(dProps(strSize))
    End If
End Function

Private Sub RewriteProp(dBlock As Dictionary, dProps As Dictionary, _
    strName As String, lngValue As Long)

    Dim dPropLines As Dictionary
    Dim lngLine As Long
    Dim strLine As String
    Dim intIndent As Integer

    ' Order matters: reading a missing key from a Dictionary would add it.
    Set dPropLines = dBlock("PropLines")
    If Not dPropLines.Exists(strName) Then Exit Sub
    If CLng(dProps(strName)) = lngValue Then Exit Sub
    lngLine = dPropLines(strName)
    strLine = m_strLines(lngLine)
    intIndent = Len(strLine) - Len(LTrim$(strLine))
    m_strLines(lngLine) = Space$(intIndent) & strName & " =" & lngValue
    dProps(strName) = lngValue
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
