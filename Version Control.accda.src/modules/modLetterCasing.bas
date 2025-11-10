Attribute VB_Name = "modLetterCasing"
Option Compare Database
Option Private Module
Option Explicit


' ----------------------------------------------------------------
' Procedure : StandardizeLetterCasing
' Date      : 5/7/2025
' Author    : Mike Wolfe
' Source    : https://nolongerset.com/standardizelettercasing/
' Purpose   : Enforce consistent letter-casing for VBA identifiers
'               as a workaround for VBA's case-changing "feature."
'  5/23/25  : Add "PtrSafe" to avoid compile errors in 64-bit VBA.
' 10/29/25  : Implement in VCS Add-In (Adam Waller)
' ----------------------------------------------------------------
Sub StandardizeLetterCasing()

    Const StandardLetterCasingModuleName As String = "clsStandardLetterCasing"

    'Get the Standard Letter Casing class module
    Dim Comp As VBIDE.VBComponent
    Dim cm As VBIDE.CodeModule
    For Each Comp In CurrentVBProject.VBComponents
        Set cm = Comp.CodeModule
        If cm.Name = StandardLetterCasingModuleName Then Exit For
    Next Comp
    If cm Is Nothing Then Exit Sub
    If cm.Name <> StandardLetterCasingModuleName Then
        'Debug.Print "Could not find '" & StandardLetterCasingModuleName & "' code module"
        Exit Sub
    End If

    'Loop through each line of code and replace the identifier name with its
    '   canonical form in the trailing comment if casing is different
    Dim i As Long
    For i = 1 To cm.CountOfLines
        Dim LineOfCode As String
        LineOfCode = Trim$(cm.Lines(i, 1))
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
                    cm.ReplaceLine i, "Dim " & CanonicalCasing & " '" & CanonicalCasing
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
                End If
            Else
                Debug.Print "Identifier mismatch on line " & i & " of " & _
                            StandardLetterCasingModuleName & " module: " & _
                            LineOfCode
            End If
        End If
    Next i

End Sub
