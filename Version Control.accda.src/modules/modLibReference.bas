Attribute VB_Name = "modLibReference"
'---------------------------------------------------------------------------------------
' Module    : modLibReference
' Author    : Adam Waller
' Date      : 2/17/2022
' Purpose   : Module for handling functions related to Access Library references
'---------------------------------------------------------------------------------------
Option Compare Database
Option Private Module
Option Explicit


'---------------------------------------------------------------------------------------
' Procedure : LocalizeLibraryReferences
' Author    : Adam Waller
' Date      : 2/17/2022
' Purpose   : Ensures that referenced Microsoft Access Library databases are linked to
'           : the version in the same folder as the parent database if a copy of the
'           : file exists in that location. (Used in some deployment scenarios)
'           :
'           : Background:
'           : One database can reference another database as a library reference. If
'           : the library database is in the same folder as the parent database, it can
'           : automatically find it when the parent database is first opened on a client
'           : computer. This gives the parent database the ability to open without
'           : compile errors.
'---------------------------------------------------------------------------------------
'
Public Sub LocalizeLibraryReferences(Optional blnAlwaysShowGUI As Boolean)

    Dim strPath As String
    Dim strFolder As String
    Dim dRefs As Dictionary
    Dim varKey As Variant
    Dim frm As Form_frmVCSMain
    Dim oApp As Access.Application

    ' Look up the references from the current database
    Perf.StartTiming
    Perf.OperationStart "Scan references"
    Set dRefs = GetReferencesDictionary
    Perf.OperationEnd

    ' We may not need to show the GUI if no problems are found.
    If dRefs("ProjCount") = 0 And Not blnAlwaysShowGUI Then
        Debug.Print "Verified local library references"
        Exit Sub
    End If

    ' Reset the log file
    Log.Clear
    strPath = CurrentProject.FullName
    strFolder = CurrentProject.Path & PathSep

    ' Use the main form to display progress
    DoCmd.OpenForm "frmVCSMain", , , , , acHidden
    Set frm = Form_frmVCSMain   ' Connect to hidden instance
    With frm

        ' Prepare the UI screen
        .cmdClose.SetFocus
        .HideActionButtons
        DoEvents
        With .txtLog
            .ScrollBars = 0
            .Visible = True
            .SetFocus
        End With
        Log.SetConsole .txtLog, .GetProgressBar

        ' Show the status
        .SetStatusText "Running...", "Localizing References", _
            "Linking Access database library references to files in the current folder."
        Log.Add "Fixing " & dRefs("RefCount") & " references in " & dRefs("ProjCount") & " databases. " & _
            "This may take several seconds to complete, so please be patient."
        Log.Spacer
        .Visible = True
    End With

    ' Loop through databases
    For Each varKey In dRefs.Keys
        Select Case varKey
            Case "ProjCount", "RefCount"
            Case Else
                Log.Add FSO.GetFileName(varKey)
                ShiftOpenDatabase strFolder & varKey, True, frm
                FixReferences dRefs(varKey)
        End Select
    Next varKey
    Log.Spacer

    ' Reopen the original database, if it is not already open
    If CurrentProject.FullName <> strPath Then
        Log.Add "Opening original database..."
        Log.Flush
        ShiftOpenDatabase strPath, False, frm
        DoEvents
    End If

    Log.Spacer
    Log.Add "Operation Complete", , , , True

    ' Save the log file
    Perf.EndTiming
    With Log
        .Flush
        .Add vbCrLf & Perf.GetReports, False
        .SaveFile FSO.BuildPath(Options.GetExportFolder, "References.log")
        .Active = False
        .Clear
    End With

End Sub


'---------------------------------------------------------------------------------------
' Procedure : GetReferencesDictionary
' Author    : Adam Waller
' Date      : 2/17/2022
' Purpose   : Return a dictionary object containing the databases and references
'           : that should be localized. (Returns no items if all Access libraries are
'           : already localized.)
'           : TIP: Use `?ConvertToJson(GetReferencesDictionary,2)` to inspect dictionary
'---------------------------------------------------------------------------------------
'
Public Function GetReferencesDictionary() As Dictionary

    Dim proj As VBProject
    Dim ref As VBIDE.Reference
    Dim strFolder As String
    Dim strFile As String
    Dim strPath As String
    Dim dProjects As Dictionary
    Dim dRefs As Dictionary
    Dim varKey As Variant
    Dim blnAdd As Boolean
    Dim strRefPath As String

    ' Create dictionary and header info
    Set dProjects = New Dictionary
    dProjects.CompareMode = TextCompare
    dProjects("ProjCount") = 0
    dProjects("RefCount") = 0

    ' Activate the VB Project for the current database
    Set proj = CurrentVBProject
    strFolder = FSO.GetParentFolderName(proj.FileName) & PathSep

    ' Loop through all projects
    For Each proj In VBE.VBProjects

        ' Get name of project file, and see if it exists in the parent folder
        'Debug.Print proj.Name & " (" & proj.FileName & ")"
        strFile = FSO.GetFileName(proj.FileName)
        strPath = strFolder & strFile
        If FSO.FileExists(strPath) Then

            ' Set up dictionary of refs for this file
            Set dRefs = New Dictionary

            ' Loop through references, looking for projects (libraries)
            For Each ref In proj.References
                If ref.Type = vbext_rk_Project Then
                    'Debug.Print " - " & ref.Name & " (" & ref.FullPath & ")"
                    strRefPath = strFolder & FSO.GetFileName(ref.FullPath)
                    If FSO.FileExists(strRefPath) Then
                        ' Target file exists in the parent folder. Check ref path.
                        If StrComp(strRefPath, ref.FullPath, vbTextCompare) <> 0 Then
                            ' We need to relink this reference
                            dRefs.Add ref.Name, strRefPath
                        End If
                    End If
                End If
            Next ref

            ' Add to list of databases if we found at least one reference to fix
            If dRefs.Count > 0 Then

                ' We might have multiple VB projects with the same name, but
                ' pointing to different locations. Add ALL broken references
                ' since we don't know which file has the broken reference.
                If dProjects.Exists(strFile) Then
                    For Each varKey In dRefs.Keys
                        If Not dProjects(strFile).Exists(varKey) Then
                            dProjects(strPath).Add varKey, dRefs(varKey)
                            dProjects("RefCount") = dProjects("RefCount") + 1
                        End If
                    Next varKey
                Else
                    ' Add database and broken references
                    dProjects.Add strFile, dRefs
                    dProjects("ProjCount") = dProjects("ProjCount") + 1
                    dProjects("RefCount") = dProjects("RefCount") + dRefs.Count
                End If
            End If
        End If
    Next proj

    ' Return dictionary
    Set GetReferencesDictionary = dProjects

End Function


'---------------------------------------------------------------------------------------
' Procedure : FixReferences
' Author    : Adam Waller
' Date      : 2/25/2022
' Purpose   : Work through the dictionary of references to fix
'---------------------------------------------------------------------------------------
'
Private Sub FixReferences(dProject As Dictionary)

    Dim proj As VBProject
    Dim colExisting As Collection
    Dim varItem As Variant
    Dim ref As VBIDE.Reference

    Set colExisting = New Collection
    Set proj = CurrentVBProject

    ' Build a collection of the existing library references
    ' (We can't change the order, other than by removing and adding
    '  back in, so we need to work from this list to preserve the
    '  order of the references.)
    ' This also assumes that the the library references are AFTER
    ' any type libraries. (Which should generally be the case)
    For Each ref In proj.References
        If ref.Type = vbext_rk_Project Then
            If colExisting.Count = 0 Then
                ' Only need to remove and add the ones at and after
                ' the point where we find the first broken one.
                If dProject.Exists(ref.Name) Then colExisting.Add Array(ref.Name, ref.FullPath)
            Else
                ' Add all library references from this point on.
                colExisting.Add Array(ref.Name, ref.FullPath)
            End If
        End If
    Next ref
    Set ref = Nothing

    ' Now go through the list of saved references, and remove and add them back in.
    For Each varItem In colExisting
        proj.References.Remove proj.References(varItem(0))
        If dProject.Exists(varItem(0)) Then
            Log.Add " - Fixing: " & varItem(0), , , "Blue"
            proj.References.AddFromFile dProject(varItem(0))
        Else
            Log.Add " - Linking: " & varItem(0)
            proj.References.AddFromFile varItem(1)
        End If
    Next varItem

End Sub
