Attribute VB_Name = "modLetterCasing"
'---------------------------------------------------------------------------------------
' Module    : modLetterCasing
' Author    : Adam Waller
' Date      : 10/29/2025
' Purpose   : Enforce consistent letter-casing for VBA identifiers as a workaround for
'           : VBA's case-changing behavior, using rules from a clsStandardLetterCasing
'           : module in the current project.
' Source    : Based on Mike Wolfe's technique - https://nolongerset.com/standardizelettercasing/
' Layer     : Core Logic
' Depends on: modObjects, modConstants, modErrorHandling
'---------------------------------------------------------------------------------------
Option Compare Database
Option Private Module
Option Explicit
'@Folder("Core")

Private Const ModuleName As String = "modLetterCasing"


'---------------------------------------------------------------------------------------
' Procedure : StandardizeLetterCasing
' Author    : Adam Waller
' Date      : 10/29/2025
' Purpose   : Enforce consistent letter-casing for VBA identifiers as a workaround for
'           : VBA's case-changing behavior, using rules from clsStandardLetterCasing.
' Source    : Based on Mike Wolfe's technique - https://nolongerset.com/standardizelettercasing/
'  5/23/25  : Add "PtrSafe" to avoid compile errors in 64-bit VBA. (Mike Wolfe)
' 10/29/25  : Implement in VCS Add-In. (Adam Waller)
'  7/06/26  : Return Collection of before/after change strings. (Adam Waller)
'  7/09/26  : Validate entries; log and skip invalid lines. (Adam Waller)
'---------------------------------------------------------------------------------------
'
Public Function StandardizeLetterCasing(Optional ByRef colIssues As Collection) As Collection

    Const FunctionName As String = ModuleName & ".StandardizeLetterCasing"
    Const StandardLetterCasingModuleName As String = "clsStandardLetterCasing"

    Dim cmp As VBIDE.VBComponent
    Dim cm As VBIDE.CodeModule
    Dim colChanges As New Collection
    Dim lngLine As Long
    Dim strOrigLine As String
    Dim strLine As String
    Dim strCurrent As String
    Dim strCanonical As String
    Dim blnNamesMatch As Boolean
    Dim blnCasingDiffers As Boolean
    Dim lngComment As Long
    Dim lngIdent As Long
    Dim lngStart As Long

    LogUnhandledErrors FunctionName
    On Error GoTo ErrHandler

    ' Find the standard letter casing class module
    For Each cmp In CurrentVBProject.VBComponents
        Set cm = cmp.CodeModule
        If cm.Name = StandardLetterCasingModuleName Then Exit For
    Next cmp
    If cm Is Nothing Then Exit Function
    If cm.Name <> StandardLetterCasingModuleName Then Exit Function

    ' Loop through each line and replace the identifier with its canonical form
    ' in the trailing comment when casing differs.
    For lngLine = 1 To cm.CountOfLines
        strOrigLine = cm.Lines(lngLine, 1)
        strLine = Trim$(strOrigLine)

        If Left$(strLine, 3) = "Dim" Then
            lngComment = InStr(1, strLine, "'")
            If lngComment = 0 Then
                RecordLetterCasingIssue colIssues, lngLine, _
                    T("Missing canonical comment. Expected: Dim Name 'Name")
            Else
                ' Position 5 is the first character after "Dim "
                strCurrent = FirstToken(Trim$(Mid$(strLine, 5, lngComment - 5)))
                strCanonical = Trim$(Mid$(strLine, lngComment + 1))
                If Len(strCurrent) = 0 Or Len(strCanonical) = 0 Then
                    RecordLetterCasingIssue colIssues, lngLine, _
                        T("Missing identifier or canonical name. Expected: Dim Name 'Name")
                Else
                    blnNamesMatch = (UCase$(strCurrent) = UCase$(strCanonical))
                    If blnNamesMatch Then
                        blnCasingDiffers = (InStr(1, strCurrent, strCanonical, vbBinaryCompare) = 0)

                        If blnCasingDiffers Then
                            ' Patch the identifier in place so the user's indentation and any
                            ' custom whitespace between the Dim declaration and the trailing
                            ' comment are preserved. UCase equality above guarantees both
                            ' strings are the same length, so Mid$ left-side assignment is safe.
                            lngIdent = InStr(4, strOrigLine, strCurrent, vbBinaryCompare)
                            If lngIdent > 0 Then
                                Mid$(strOrigLine, lngIdent, Len(strCurrent)) = strCanonical
                                cm.ReplaceLine lngLine, strOrigLine
                                colChanges.Add strCurrent & " -> " & strCanonical
                            End If
                        End If
                    Else
                        RecordLetterCasingIssue colIssues, lngLine, _
                            T("Identifier mismatch: '{0}' does not match '{1}'.", _
                                var0:=strCurrent, var1:=strCanonical)
                    End If
                End If
            End If
        ElseIf Left$(strLine, 32) = "Private Declare PtrSafe Function" Then
            lngComment = InStr(1, strLine, "'")
            If lngComment = 0 Then
                RecordLetterCasingIssue colIssues, lngLine, _
                    T("Missing canonical comment on API declare line.")
            Else
                lngStart = InStr(1, strLine, """")
                If lngStart = 0 Then
                    RecordLetterCasingIssue colIssues, lngLine, _
                        T("Missing DLL name in Lib clause.")
                Else
                    strCurrent = ExtractDeclareLibName(strLine)
                    If Len(strCurrent) = 0 Then
                        RecordLetterCasingIssue colIssues, lngLine, _
                            T("Malformed Lib clause.")
                    Else
                        strCanonical = Trim$(Mid$(strLine, lngComment + 1))
                        If Len(strCanonical) = 0 Then
                            RecordLetterCasingIssue colIssues, lngLine, _
                                T("Missing DLL name or canonical name on API declare line.")
                        Else
                            blnNamesMatch = (UCase$(strCurrent) = UCase$(strCanonical))
                            If blnNamesMatch Then
                                blnCasingDiffers = (InStr(1, strCurrent, strCanonical, vbBinaryCompare) = 0)

                                If blnCasingDiffers Then
                                    cm.ReplaceLine lngLine, "Private Declare PtrSafe Function zzz_" & Replace(strCanonical, ".", "_") & _
                                                      " Lib """ & strCanonical & """ '" & strCanonical
                                    colChanges.Add strCurrent & " -> " & strCanonical
                                End If
                            Else
                                RecordLetterCasingIssue colIssues, lngLine, _
                                    T("Identifier mismatch: '{0}' does not match '{1}'.", _
                                        var0:=strCurrent, var1:=strCanonical)
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Next lngLine

    Set StandardizeLetterCasing = colChanges

    ' Persist the corrections so the VBA project isn't left with unsaved changes that
    ' prompt a save when Access is closed. Correcting the casing in
    ' clsStandardLetterCasing propagates canonical casing project-wide via VBA, dirtying
    ' form and report class modules too -- which is why the single-module save this used to
    ' do never finished the job, and why the warning below fired run after run without the
    ' corrections ever converging.
    If colChanges.Count > 0 Then
        ' SaveCurrentVBProject reports the project's real state and handles its own errors,
        ' so its return value is the signal here rather than Err.
        If Not SaveCurrentVBProject Then
            Log.Error eelWarning, T("VBA project still has unsaved changes after letter casing corrections"), FunctionName
        End If
    End If

    Exit Function

ErrHandler:
    RecordLetterCasingIssue colIssues, lngLine, T("Error standardizing letter casing")
    CatchAny eelError, T("Error standardizing letter casing"), FunctionName
    Set StandardizeLetterCasing = colChanges

End Function


'---------------------------------------------------------------------------------------
' Procedure : RecordLetterCasingIssue
' Author    : Adam Waller
' Date      : 7/09/2026
' Purpose   : Collect, log, and debug-print an invalid clsStandardLetterCasing entry.
'---------------------------------------------------------------------------------------
'
Private Sub RecordLetterCasingIssue(ByRef colIssues As Collection, ByVal lngLine As Long, ByVal strMessage As String)

    Dim strFull As String

    strFull = T("Line {0} in clsStandardLetterCasing: {1}", var0:=lngLine, var1:=strMessage)
    If Not colIssues Is Nothing Then colIssues.Add strFull
    Log.Error eelWarning, strFull, ModuleName & ".StandardizeLetterCasing"
    Debug.Print strFull

End Sub


'---------------------------------------------------------------------------------------
' Procedure : FirstToken
' Author    : Adam Waller
' Date      : 7/09/2026
' Purpose   : Return the first space-delimited token from strText.
'---------------------------------------------------------------------------------------
'
Private Function FirstToken(ByVal strText As String) As String
    FirstToken = Split(Trim$(strText))(0)
End Function


'---------------------------------------------------------------------------------------
' Procedure : ExtractDeclareLibName
' Author    : Adam Waller
' Date      : 8/17/2026
' Purpose   : Extract the DLL name from the Lib clause of an API declare line.
'           : Input:  Private Declare PtrSafe Function zzz_kernel32 Lib "kernel32" () 'kernel32
'           : Output: kernel32
'---------------------------------------------------------------------------------------
'
Public Function ExtractDeclareLibName(ByVal strLine As String) As String

    Dim lngStart As Long
    Dim lngEnd As Long

    lngStart = InStr(1, strLine, """")
    If lngStart = 0 Then Exit Function
    lngStart = lngStart + 1
    lngEnd = InStr(lngStart, strLine, """")
    If lngEnd <= lngStart Then Exit Function
    ExtractDeclareLibName = Mid$(strLine, lngStart, lngEnd - lngStart)

End Function
