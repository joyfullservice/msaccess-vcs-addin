Attribute VB_Name = "modLetterCasing"
Option Compare Database
Option Private Module
Option Explicit
'@Folder("Core")


' ----------------------------------------------------------------
' Procedure : StandardizeLetterCasing
' Date      : 5/7/2025
' Author    : Mike Wolfe
' Source    : https://nolongerset.com/standardizelettercasing/
' Purpose   : Enforce consistent letter-casing for VBA identifiers
'               as a workaround for VBA's case-changing "feature."
'  5/23/25  : Add "PtrSafe" to avoid compile errors in 64-bit VBA.
' 10/29/25  : Implement in VCS Add-In (Adam Waller)
'  7/06/26  : Return Collection of before/after change strings (Adam Waller)
' ----------------------------------------------------------------
Public Function StandardizeLetterCasing() As Collection

    Const StandardLetterCasingModuleName As String = "clsStandardLetterCasing"

    'Get the Standard Letter Casing class module
    Dim Comp As VBIDE.VBComponent
    Dim cm As VBIDE.CodeModule
    For Each Comp In CurrentVBProject.VBComponents
        Set cm = Comp.CodeModule
        If cm.Name = StandardLetterCasingModuleName Then Exit For
    Next Comp
    If cm Is Nothing Then Exit Function
    If cm.Name <> StandardLetterCasingModuleName Then Exit Function

    Dim colChanges As New Collection

    'Loop through each line of code and replace the identifier name with its
    '   canonical form in the trailing comment if casing is different
    Dim i As Long
    For i = 1 To cm.CountOfLines
        Dim OrigLine As String
        Dim LineOfCode As String
        OrigLine = cm.Lines(i, 1)
        LineOfCode = Trim$(OrigLine)
        Dim CurrentCasing As String, CanonicalCasing As String
        Dim NamesMatch As String, CasingDiffers As Boolean
        If Left(LineOfCode, 3) = "Dim" Then
            CurrentCasing = Trim$(Mid$(LineOfCode, 5, InStr(5, LineOfCode, " ") - 5))
            CanonicalCasing = Trim$(Mid$(LineOfCode, InStr(1, LineOfCode, "'") + 1))
            NamesMatch = (UCase$(CurrentCasing) = UCase$(CanonicalCasing))
            If NamesMatch Then
                'Perform a case-sensitive text comparison between the comment and its identifier counterpart
                CasingDiffers = (InStr(1, CurrentCasing, CanonicalCasing, vbBinaryCompare) = 0)

                If CasingDiffers Then
                    ' Patch the identifier in place so the user's indentation and any
                    ' custom whitespace between the Dim declaration and the trailing
                    ' comment are preserved. UCase equality above guarantees both
                    ' strings are the same length, so Mid$ left-side assignment is safe.
                    Dim posIdent As Long
                    posIdent = InStr(4, OrigLine, CurrentCasing, vbBinaryCompare)
                    If posIdent > 0 Then
                        Mid$(OrigLine, posIdent, Len(CurrentCasing)) = CanonicalCasing
                        cm.ReplaceLine i, OrigLine
                        colChanges.Add CurrentCasing & " -> " & CanonicalCasing
                    End If
                End If
            Else
                Debug.Print "Identifier mismatch on line " & i & " of " & _
                            StandardLetterCasingModuleName & " module: " & _
                            LineOfCode
            End If
        ElseIf Left(LineOfCode, 32) = "Private Declare PtrSafe Function" Then
            Dim StartPos As Long, EndPos As Long
            StartPos = InStr(1, LineOfCode, """") + 1
            EndPos = InStr(StartPos, LineOfCode, """") - 1
            CurrentCasing = Mid(LineOfCode, StartPos, EndPos - StartPos + 1)

            CanonicalCasing = Trim$(Mid$(LineOfCode, InStr(1, LineOfCode, "'") + 1))
            NamesMatch = (UCase$(CurrentCasing) = UCase$(CanonicalCasing))
            If NamesMatch Then
                CasingDiffers = (InStr(1, CurrentCasing, CanonicalCasing, vbBinaryCompare) = 0)

                If CasingDiffers Then
                    cm.ReplaceLine i, "Private Declare PtrSafe Function zzz_" & Replace(CanonicalCasing, ".", "_") & _
                                      " Lib """ & CanonicalCasing & """ '" & CanonicalCasing
                    colChanges.Add CurrentCasing & " -> " & CanonicalCasing
                End If
            Else
                Debug.Print "Identifier mismatch on line " & i & " of " & _
                            StandardLetterCasingModuleName & " module: " & _
                            LineOfCode
            End If
        End If
    Next i

    Set StandardizeLetterCasing = colChanges

    ' Persist the corrections so the VBA project isn't left with unsaved
    ' changes that prompt a save when Access is closed. Correcting the casing
    ' in clsStandardLetterCasing propagates canonical casing project-wide via
    ' VBA (dirtying other modules too); saving one module saves the whole project.
    If colChanges.Count > 0 Then
        LogUnhandledErrors
        On Error Resume Next
        DoCmd.Save acModule, StandardLetterCasingModuleName
        If Err Then Err.Clear
        On Error GoTo 0
    End If

End Function
