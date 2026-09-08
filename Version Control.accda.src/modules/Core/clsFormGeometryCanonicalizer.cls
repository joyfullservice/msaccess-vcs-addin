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
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
'@Folder("Core")

Private Const GRID_TWIPS As Long = 60
Private Const ModuleNameConst As String = "clsFormGeometryCanonicalizer"

Private Type udtThis
    varLines As Variant
    colBlocks As Collection
    dRoot As Dictionary
    strFormName As String
    lngChanged As Long
End Type
Private this As udtThis


'---------------------------------------------------------------------------------------
' Procedure : Canonicalize
' Author    : Adam Waller
' Date      : 9/7/2026
' Purpose   : Normalize layout-group geometry in an already-split SaveAsText line array.
'---------------------------------------------------------------------------------------
'
Public Sub Canonicalize(ByRef varLines As Variant, Optional strFormName As String)

    If DebugMode(True) Then On Error GoTo 0 Else On Error Resume Next

    Perf.OperationStart "Canonicalize Form Geometry"
    this.varLines = varLines
    this.strFormName = strFormName
    this.lngChanged = 0
    Set this.colBlocks = New Collection
    Set this.dRoot = Nothing

    ParseDocument
    CanonicalizeGroups
    CanonicalizeTabs
    GrowEnvelopes

    varLines = this.varLines
    Set this.colBlocks = Nothing
    Set this.dRoot = Nothing
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
    Dim strTrimmed As String
    Dim dBlock As Dictionary

    lngLine = 0
    Do While lngLine <= UBound(this.varLines)
        strTrimmed = Trim$(CStr(this.varLines(lngLine)))
        If strTrimmed = "CodeBehindForm" Then Exit Do
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
    Dim strTrimmed As String
    Dim strName As String
    Dim strValue As String
    Dim lngPos As Long
    Dim dChild As Dictionary

    Set dBlock = NewBlock(Trim$(strType), lngStart, dParent)

    Do While lngLine <= UBound(this.varLines)
        strTrimmed = Trim$(CStr(this.varLines(lngLine)))
        If strTrimmed = "CodeBehindForm" Then
            dBlock("EndLine") = lngLine
            dBlock("NextLine") = UBound(this.varLines) + 1
            Set ParseBlock = dBlock
            Exit Function
        End If

        If IsInlineBinary(strTrimmed) Then
            lngLine = lngLine + 1
            Do While lngLine <= UBound(this.varLines)
                If Trim$(CStr(this.varLines(lngLine))) = "End" Then Exit Do
                lngLine = lngLine + 1
            Loop
            lngLine = lngLine + 1
        ElseIf StartsWith(strTrimmed, "Begin") Then
            Set dChild = ParseBlock(lngLine + 1, Mid$(strTrimmed, 6), lngLine, dBlock)
            lngLine = dChild("NextLine")
        ElseIf strTrimmed = "End" Then
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
                    dBlock("Props")(strName) = CLng(strValue)
                    dBlock("PropLines")(strName) = lngLine
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

    Set dBlock = New Dictionary
    dBlock("Type") = strType
    dBlock("Start") = lngStart
    dBlock("Name") = vbNullString
    Set dBlock("Props") = New Dictionary
    Set dBlock("PropLines") = New Dictionary
    Set dBlock("Children") = New Collection
    Set dBlock("Parent") = dParent
    If Not dParent Is Nothing Then
        dParent("Children").Add dBlock
        If IsSectionType(dParent("Type")) Then
            Set dBlock("Section") = dParent
        ElseIf dParent.Exists("Section") Then
            Set dBlock("Section") = dParent("Section")
        End If
        If dParent("Type") = "Tab" Then
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
    Dim strKey As String
    Dim varKey As Variant

    Set dGroups = New Dictionary
    For Each dBlock In this.colBlocks
        If IsLayout(dBlock) Then
            strKey = GroupKey(dBlock)
            If Not dGroups.Exists(strKey) Then dGroups.Add strKey, New Collection
            dGroups(strKey).Add dBlock
        End If
    Next dBlock

    For Each varKey In dGroups.Keys
        CanonicalizeGroup dGroups(varKey), CStr(varKey)
    Next varKey

End Sub


'---------------------------------------------------------------------------------------
' Procedure : CanonicalizeGroup
'---------------------------------------------------------------------------------------
'
Private Sub CanonicalizeGroup(colCells As Collection, strLabel As String)

    Dim colPlanned As Collection
    Dim dBySection As Dictionary
    Dim dBlock As Dictionary
    Dim strSection As String
    Dim varSection As Variant
    Dim varItem As Variant

    Set colPlanned = PlanAxis(colCells, "x", "ColumnStart", "ColumnEnd", "Left", "Width")
    If colPlanned Is Nothing Then
        WarnSkip strLabel & " (horizontal)"
        Exit Sub
    End If

    Set dBySection = New Dictionary
    For Each dBlock In colCells
        If dBlock.Exists("Section") Then
            strSection = CStr(ObjPtr(dBlock("Section")))
        Else
            strSection = "none"
        End If
        If Not dBySection.Exists(strSection) Then dBySection.Add strSection, New Collection
        dBySection(strSection).Add dBlock
    Next dBlock

    For Each varSection In dBySection.Keys
        Set varItem = PlanAxis(dBySection(varSection), "y", "RowStart", "RowEnd", "Top", "Height")
        If varItem Is Nothing Then
            WarnSkip strLabel & " (vertical)"
            Exit Sub
        End If
        For Each dBlock In varItem
            colPlanned.Add dBlock
        Next dBlock
    Next varSection

    For Each dBlock In colPlanned
        RewriteProp dBlock("Block"), CStr(dBlock("Name")), CLng(dBlock("After"))
    Next dBlock

End Sub


'---------------------------------------------------------------------------------------
' Procedure : PlanAxis
'---------------------------------------------------------------------------------------
'
Private Function PlanAxis(colCells As Collection, strAxis As String, _
    strStartName As String, strEndName As String, strPosName As String, _
    strSizeName As String) As Collection

    Dim dSizes As Dictionary
    Dim dPositions As Dictionary
    Dim colPlanned As Collection
    Dim dBlock As Dictionary
    Dim dChange As Dictionary
    Dim lngStart As Long
    Dim lngEnd As Long
    Dim varDerived As Variant

    Set dSizes = TrackSizes(colCells, strAxis, strStartName, strEndName)
    Set dPositions = BoundaryPositions(colCells, strAxis, strStartName)
    If dPositions Is Nothing Then
        Set PlanAxis = Nothing
        Exit Function
    End If
    InferEndSizes colCells, strAxis, strStartName, strEndName, dPositions, dSizes

    Set colPlanned = New Collection
    For Each dBlock In colCells
        lngStart = PropOrDefault(dBlock, strStartName, 0)
        lngEnd = PropOrDefault(dBlock, strEndName, lngStart)
        If dBlock("Props").Exists(strPosName) Then
            If Not dPositions.Exists(lngStart) Then
                Set PlanAxis = Nothing
                Exit Function
            End If
            Set dChange = New Dictionary
            Set dChange("Block") = dBlock
            dChange("Name") = strPosName
            dChange("After") = CLng(dPositions(lngStart))
            colPlanned.Add dChange
        End If
        If dBlock("Props").Exists(strSizeName) Then
            If lngStart = lngEnd Then
                If dSizes.Exists(lngStart) Then
                    varDerived = CLng(dSizes(lngStart))
                Else
                    varDerived = Null
                End If
            Else
                varDerived = DeriveSpan(lngStart, lngEnd, dPositions, dSizes)
            End If
            If IsNull(varDerived) Then
                Set PlanAxis = Nothing
                Exit Function
            End If
            Set dChange = New Dictionary
            Set dChange("Block") = dBlock
            dChange("Name") = strSizeName
            dChange("After") = CLng(varDerived)
            colPlanned.Add dChange
        End If
    Next dBlock

    Set PlanAxis = colPlanned

End Function


'---------------------------------------------------------------------------------------
' Procedure : TrackSizes
'---------------------------------------------------------------------------------------
'
Private Function TrackSizes(colCells As Collection, strAxis As String, _
    strStartName As String, strEndName As String) As Dictionary

    Dim dObserved As Dictionary
    Dim dSizes As Dictionary
    Dim dBlock As Dictionary
    Dim lngStart As Long
    Dim lngEnd As Long
    Dim varSize As Variant
    Dim varIndex As Variant

    Set dObserved = New Dictionary
    For Each dBlock In colCells
        lngStart = PropOrDefault(dBlock, strStartName, 0)
        lngEnd = PropOrDefault(dBlock, strEndName, lngStart)
        varSize = GetSize(dBlock, strAxis)
        If lngStart = lngEnd And Not IsNull(varSize) Then
            If Not dObserved.Exists(lngStart) Then dObserved.Add lngStart, New Collection
            dObserved(lngStart).Add CLng(varSize)
        End If
    Next dBlock

    Set dSizes = New Dictionary
    For Each varIndex In dObserved.Keys
        dSizes(CLng(varIndex)) = SnapTwips(MedianLong(dObserved(varIndex)))
    Next varIndex
    Set TrackSizes = dSizes

End Function


'---------------------------------------------------------------------------------------
' Procedure : BoundaryPositions
'---------------------------------------------------------------------------------------
'
Private Function BoundaryPositions(colCells As Collection, strAxis As String, _
    strStartName As String) As Dictionary

    Dim dObserved As Dictionary
    Dim dRaw As Dictionary
    Dim dPositions As Dictionary
    Dim dBlock As Dictionary
    Dim varPos As Variant
    Dim colStarts As Collection
    Dim lngPrev As Long
    Dim lngCurrent As Long
    Dim lngI As Long

    Set dObserved = New Dictionary
    For Each dBlock In colCells
        varPos = GetPos(dBlock, strAxis)
        If Not IsNull(varPos) Then
            lngCurrent = PropOrDefault(dBlock, strStartName, 0)
            If Not dObserved.Exists(lngCurrent) Then dObserved.Add lngCurrent, New Collection
            dObserved(lngCurrent).Add CLng(varPos)
        End If
    Next dBlock
    If dObserved.Count = 0 Then
        Set BoundaryPositions = Nothing
        Exit Function
    End If

    Set colStarts = SortedKeys(dObserved)
    Set dRaw = New Dictionary
    For lngI = 1 To colStarts.Count
        dRaw(CLng(colStarts(lngI))) = MedianLong(dObserved(colStarts(lngI)))
    Next lngI

    Set dPositions = New Dictionary
    dPositions(CLng(colStarts(1))) = SnapTwips(CLng(dRaw(colStarts(1))))
    For lngI = 1 To colStarts.Count - 1
        lngPrev = CLng(colStarts(lngI))
        lngCurrent = CLng(colStarts(lngI + 1))
        dPositions(lngCurrent) = CLng(dPositions(lngPrev)) + SnapTwips(CLng(dRaw(lngCurrent)) - CLng(dRaw(lngPrev)))
    Next lngI
    Set BoundaryPositions = dPositions

End Function


'---------------------------------------------------------------------------------------
' Procedure : InferEndSizes
'---------------------------------------------------------------------------------------
'
Private Sub InferEndSizes(colCells As Collection, strAxis As String, _
    strStartName As String, strEndName As String, dPositions As Dictionary, _
    dSizes As Dictionary)

    Dim dInferred As Dictionary
    Dim dBlock As Dictionary
    Dim lngStart As Long
    Dim lngEnd As Long
    Dim varSize As Variant
    Dim varIndex As Variant

    Set dInferred = New Dictionary
    For Each dBlock In colCells
        lngStart = PropOrDefault(dBlock, strStartName, 0)
        lngEnd = PropOrDefault(dBlock, strEndName, lngStart)
        varSize = GetSize(dBlock, strAxis)
        If (Not IsNull(varSize)) And lngStart <> lngEnd Then
            If dPositions.Exists(lngStart) And dPositions.Exists(lngEnd) Then
                If Not dInferred.Exists(lngEnd) Then dInferred.Add lngEnd, New Collection
                dInferred(lngEnd).Add CLng(varSize) - (CLng(dPositions(lngEnd)) - CLng(dPositions(lngStart)))
            End If
        End If
    Next dBlock

    For Each varIndex In dInferred.Keys
        If Not dSizes.Exists(varIndex) Then
            dSizes(CLng(varIndex)) = SnapTwips(MedianLong(dInferred(varIndex)))
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
    Dim varName As Variant

    For Each dBlock In this.colBlocks
        If dBlock("Type") = "Tab" And Not IsLayout(dBlock) Then
            For Each varName In Array("Width", "Height")
                If dBlock("Props").Exists(CStr(varName)) Then
                    RewriteProp dBlock, CStr(varName), SnapTwips(CLng(dBlock("Props")(varName)))
                End If
            Next varName
        End If
    Next dBlock

End Sub


'---------------------------------------------------------------------------------------
' Procedure : GrowEnvelopes
'---------------------------------------------------------------------------------------
'
Private Sub GrowEnvelopes()

    Dim dBlock As Dictionary
    Dim dSection As Dictionary
    Dim lngMaxRight As Long
    Dim lngMaxBottom As Long
    Dim lngGrown As Long
    Dim blnHasLayout As Boolean

    If this.dRoot Is Nothing Then Exit Sub
    For Each dBlock In this.colBlocks
        If IsLayout(dBlock) Then
            blnHasLayout = True
            Exit For
        End If
    Next dBlock
    If Not blnHasLayout Then Exit Sub

    For Each dBlock In this.colBlocks
        If Not IsEnvelopeContainer(dBlock("Type")) Then
            lngMaxRight = MaxLong(lngMaxRight, ChildExtent(dBlock, "Left", "Width"))
        End If
    Next dBlock
    If this.dRoot("Props").Exists("Width") Then
        lngGrown = MaxLong(SnapTwips(CLng(this.dRoot("Props")("Width"))), SnapUp(lngMaxRight))
        RewriteProp this.dRoot, "Width", lngGrown
    End If

    For Each dSection In this.colBlocks
        If IsSectionType(dSection("Type")) And dSection("Props").Exists("Height") Then
            lngMaxBottom = 0
            For Each dBlock In this.colBlocks
                If dBlock.Exists("Section") Then
                    If ObjPtr(dBlock("Section")) = ObjPtr(dSection) Then
                        If dBlock("Type") <> "Tab" And dBlock("Type") <> "Page" _
                            And Not dBlock.Exists("Tab") Then
                            lngMaxBottom = MaxLong( _
                                lngMaxBottom, ChildExtent(dBlock, "Top", "Height"))
                        End If
                    End If
                End If
            Next dBlock
            If lngMaxBottom > 0 Then
                lngGrown = MaxLong(SnapTwips(CLng(dSection("Props")("Height"))), SnapUp(lngMaxBottom))
            Else
                lngGrown = SnapTwips(CLng(dSection("Props")("Height")))
            End If
            RewriteProp dSection, "Height", lngGrown
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

Private Function IsLayout(dBlock As Dictionary) As Boolean
    IsLayout = dBlock("Props").Exists("LayoutGroup") Or dBlock("Props").Exists("GroupTable")
End Function

Private Function GroupKey(dBlock As Dictionary) As String
    GroupKey = PropOrDefault(dBlock, "GroupTable", -1) & "/" & PropOrDefault(dBlock, "LayoutGroup", -1)
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

Private Function PropOrDefault(dBlock As Dictionary, strName As String, lngDefault As Long) As Long
    If dBlock("Props").Exists(strName) Then
        PropOrDefault = CLng(dBlock("Props")(strName))
    Else
        PropOrDefault = lngDefault
    End If
End Function

Private Function GetSize(dBlock As Dictionary, strAxis As String) As Variant
    If strAxis = "x" Then
        If dBlock("Props").Exists("Width") Then
            GetSize = CLng(dBlock("Props")("Width"))
        ElseIf dBlock("Props").Exists("LayoutCachedWidth") And dBlock("Props").Exists("Left") Then
            GetSize = CLng(dBlock("Props")("LayoutCachedWidth")) - CLng(dBlock("Props")("Left"))
        Else
            GetSize = Null
        End If
    Else
        If dBlock("Props").Exists("Height") Then
            GetSize = CLng(dBlock("Props")("Height"))
        ElseIf dBlock("Props").Exists("LayoutCachedHeight") And dBlock("Props").Exists("Top") Then
            GetSize = CLng(dBlock("Props")("LayoutCachedHeight")) - CLng(dBlock("Props")("Top"))
        Else
            GetSize = Null
        End If
    End If
End Function

Private Function GetPos(dBlock As Dictionary, strAxis As String) As Variant
    Dim strKey As String
    Dim strCache As String
    If strAxis = "x" Then
        strKey = "Left"
        strCache = "LayoutCachedLeft"
    Else
        strKey = "Top"
        strCache = "LayoutCachedTop"
    End If
    If dBlock("Props").Exists(strKey) Then
        GetPos = CLng(dBlock("Props")(strKey))
    ElseIf dBlock("Props").Exists(strCache) Then
        GetPos = CLng(dBlock("Props")(strCache))
    Else
        GetPos = Null
    End If
End Function

Private Function ChildExtent(dBlock As Dictionary, strPos As String, strSize As String) As Long
    If dBlock("Props").Exists(strPos) And dBlock("Props").Exists(strSize) Then
        ChildExtent = CLng(dBlock("Props")(strPos)) + CLng(dBlock("Props")(strSize))
    End If
End Function

Private Sub RewriteProp(dBlock As Dictionary, strName As String, lngValue As Long)
    Dim lngLine As Long
    Dim strLine As String
    Dim intIndent As Integer
    If Not dBlock("PropLines").Exists(strName) Then Exit Sub
    If CLng(dBlock("Props")(strName)) = lngValue Then Exit Sub
    lngLine = dBlock("PropLines")(strName)
    strLine = CStr(this.varLines(lngLine))
    intIndent = Len(strLine) - Len(LTrim$(strLine))
    this.varLines(lngLine) = Space$(intIndent) & strName & " =" & lngValue
    dBlock("Props")(strName) = lngValue
    this.lngChanged = this.lngChanged + 1
End Sub

Private Function MedianLong(colValues As Collection) As Long
    Dim lngValues() As Long
    Dim lngI As Long
    Dim lngJ As Long
    Dim lngSwap As Long
    Dim dblMid As Double

    If colValues.Count = 0 Then Exit Function
    ReDim lngValues(1 To colValues.Count)
    For lngI = 1 To colValues.Count
        lngValues(lngI) = CLng(colValues(lngI))
    Next lngI
    For lngI = 1 To colValues.Count - 1
        For lngJ = lngI + 1 To colValues.Count
            If lngValues(lngJ) < lngValues(lngI) Then
                lngSwap = lngValues(lngI)
                lngValues(lngI) = lngValues(lngJ)
                lngValues(lngJ) = lngSwap
            End If
        Next lngJ
    Next lngI
    If colValues.Count Mod 2 = 1 Then
        MedianLong = lngValues((colValues.Count + 1) \ 2)
    Else
        dblMid = (CDbl(lngValues(colValues.Count \ 2)) + CDbl(lngValues(colValues.Count \ 2 + 1))) / 2
        MedianLong = CLng(Round(dblMid, 0))
    End If
End Function

Private Function SortedKeys(dMap As Dictionary) As Collection
    Dim colKeys As Collection
    Dim varKey As Variant
    Dim lngI As Long
    Dim lngJ As Long
    Dim lngSwap As Long
    Dim lngKeys() As Long

    Set colKeys = New Collection
    If dMap.Count = 0 Then
        Set SortedKeys = colKeys
        Exit Function
    End If
    ReDim lngKeys(1 To dMap.Count)
    lngI = 1
    For Each varKey In dMap.Keys
        lngKeys(lngI) = CLng(varKey)
        lngI = lngI + 1
    Next varKey
    For lngI = 1 To dMap.Count - 1
        For lngJ = lngI + 1 To dMap.Count
            If lngKeys(lngJ) < lngKeys(lngI) Then
                lngSwap = lngKeys(lngI)
                lngKeys(lngI) = lngKeys(lngJ)
                lngKeys(lngJ) = lngSwap
            End If
        Next lngJ
    Next lngI
    For lngI = 1 To dMap.Count
        colKeys.Add lngKeys(lngI)
    Next lngI
    Set SortedKeys = colKeys
End Function

Private Function MaxLong(lngA As Long, lngB As Long) As Long
    If lngA > lngB Then MaxLong = lngA Else MaxLong = lngB
End Function

Private Sub WarnSkip(strReason As String)
    Log.Error eelWarning, T("Form geometry: skipped {0} on {1}", _
        var0:=strReason, var1:=Nz2(this.strFormName, "form")), _
        ModuleNameConst & ".Canonicalize"
End Sub
