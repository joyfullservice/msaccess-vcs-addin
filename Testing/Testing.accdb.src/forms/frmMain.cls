Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


' Keep track of total results
Private m_Totals(True To False) As Integer


'---------------------------------------------------------------------------------------
' Procedure : cmdRunTests_Click
' Author    : Adam Waller
' Date      : 5/1/2020
' Purpose   : Trying to keep things simple here... Verify that the object exists in the
'           : correct format. (Adjust as needed)
'---------------------------------------------------------------------------------------
'
Public Sub cmdRunTests_Click()

    Dim strTest As String
    Dim intTest As Integer
    Dim dbs As DAO.Database
    Dim rsc As SharedResource

    Set dbs = CurrentDb

    ' Clear list and totals
    lstResults.RowSource = ""
    m_Totals(True) = 0
    m_Totals(False) = 0

    ' Ignore any errors.
    ' NOTE: don't include the test result on a line that may throw an error.
    On Error Resume Next

    ' Update linked tables/CSV to use the current directory
    dbs.TableDefs("tblLinkedAccess").Connect = ";DATABASE=" & Application.CurrentProject.Path & "\Testing.accdb"
    dbs.TableDefs("tblLinkedAccess").RefreshLink
    dbs.TableDefs("tblLinkedCSV").Connect = "Text;DSN=Linked Link Specification;FMT=Delimited;HDR=NO;IMEX=2;CharacterSet=437;ACCDB=YES;DATABASE=" & Application.CurrentProject.Path
    dbs.TableDefs("tblLinkedCSV").RefreshLink

    '========================
    '   BEGIN TESTS
    '========================

    ' Tables
    strTest = dbs.TableDefs("tblInternal").Name
    ShowResult "Access Table exists", (strTest = "tblInternal")

    intTest = 0
    intTest = DCount("*", "tblInternal")
    ShowResult "tblInternal has data", (intTest > 0)

    strTest = dbs.TableDefs("tblLinkedCSV").Name
    ShowResult "Linked Table exists", (strTest = "tblLinkedCSV")

    intTest = 0
    intTest = DCount("*", "tblLinkedCSV")
    ShowResult "tblLinkedCSV has data", (intTest > 0)

    ShowResult "Saved Table Data (TDF)", FSO.FileExists(ExportFolder & "tables\tblInternal.txt")

    ShowResult "Saved Table Data (XML)", FSO.FileExists(ExportFolder & "tables\tblSaveXML.xml")

    ShowResult "Table SQL", FSO.FileExists(ExportFolder & "tbldefs\tblInternal.sql")

    ShowResult "Linked Table JSON", FSO.FileExists(ExportFolder & "tbldefs\tblLinkedCSV.json")

    ShowResult "Linked Table structure", FSO.FileExists(ExportFolder & "tbldefs\tblLinkedCSV.sql")

    intTest = 0
    intTest = dbs.Relations("tblInternaltblSaveXML").Fields.Count
    ShowResult "Table Relationship", (intTest = 1)

    intTest = 0
    intTest = DCount("*", "MSysObjects", "Not IsNull(LvExtra) and Type = 1 and [Name] = 'tblSaveXML'")
    ShowResult "Table Data Macro Exists", (intTest > 0)


    ' Queries
    strTest = dbs.QueryDefs("qryNavigationPaneGroups").Name
    ShowResult "Query exists", (strTest = "qryNavigationPaneGroups")


    ' Forms
    strTest = CurrentProject.AllForms("frmMain").Name
    ShowResult "Form exists", (strTest = "frmMain")


    ' Reports
    strTest = CurrentProject.AllReports("rptNavigationPaneGroups").Name
    ShowResult "Report exists", (strTest = "rptNavigationPaneGroups")
    ShowResult "Landscape Orientation", (Report_rptNonDefaultPaperSize.Printer.Orientation = acPRORLandscape)
    ShowResult "A4 Paper Size", (Report_rptNonDefaultPaperSize.Printer.PaperSize = acPRPSA4)


    ' Macros
    strTest = CurrentProject.AllMacros("AutoExec").Name
    ShowResult "Macro exists", (strTest = "AutoExec")


    ' Modules
    strTest = CurrentProject.AllModules("basUtility").Name
    ShowResult "Standard Module exists", (strTest = "basUtility")
    strTest = GetVBProjectForCurrentDB.VBComponents("basExtendedChars").CodeModule.Lines(6, 1)
    ShowResult "Extended ASCII text in VBA", (Mid$(strTest, 10, 1) = Chr(151))

    strTest = CurrentProject.AllModules("clsPerson").Name
    ShowResult "Class Module exists", (strTest = "clsPerson")


    ' Database properties
    strTest = ""
    strTest = dbs.Properties("AppIcon")
    ShowResult "Application Icon is set", (Len(strTest) > 5)

    strTest = dbs.Properties("DAOProperty").Value
    ShowResult "Custom Database (DAO) property", (strTest = "DAO")

    strTest = CurrentProject.Properties("ProjectProperty").Value
    ShowResult "Custom Project Property", (strTest = "TestValue")

    strTest = dbs.Containers("Databases").Documents("SummaryInfo").Properties("Title")
    ShowResult "Database Summary Property (Title)", (strTest = "VCS Testing")

    strTest = dbs.Containers("Tables").Documents("tblSaveXML").Properties("Description")
    ShowResult "Navigation pane object description", (strTest = "Saved description in XML table.")

    strTest = dbs.Containers("Modules").Documents("basUtility").Properties("Description")
    ShowResult "Module description", (strTest = "My special description on the code module.")

    ShowResult "Saved shared images", (CurrentProject.Resources.Count > 2)

    ShowResult "Saved import/export specs (XML)", (CurrentProject.ImportExportSpecifications.Count > 0)

    strTest = CurrentProject.ImportExportSpecifications(0).Name
    ShowResult "Name of saved specification", (strTest = "Export-MSysIMEXColumns")

    strTest = Nz(DLookup("SpecName", "MSysIMEXSpecs", "SpecName=""Test 2"""))
    ShowResult "Saved IMEX spec (Table based)", (strTest = "Test 2")

    strTest = Nz(DLookup("Name", "MSysNavPaneGroups", "Name=""My Modules"""))
    ShowResult "Custom navigation pane group", (strTest = "My Modules")

    ' VBE Project
    With GetVBProjectForCurrentDB

        ShowResult "VBE project name", (.Name = "VCS Testing")
        ShowResult "VBE project description", (.Description = "For automated testing of Version Control")
        ShowResult "Help context id", (.HelpContextId = 123456)

        strTest = .References("Scripting").Name
        ShowResult "GUID reference (scripting)", (strTest = "Scripting")

        strTest = .References("MSForms").Name
        ShowResult "MS Forms 2.0 reference", (strTest = "MSForms")

    End With

    ' Theme
    strTest = CurrentDb.Properties("Theme Resource Name")
    ShowResult "Active theme = Angles", (strTest = "Angles")

    strTest = vbNullString
    For Each rsc In CurrentProject.Resources
        If rsc.Type = acResourceTheme Then
            strTest = rsc.Name
            If strTest = "Angles" Then Exit For
        End If
    Next rsc
    ShowResult "Theme resource exists", (strTest = "Angles")

    ' Other
    ShowResult "VCS Options file exists", FSO.FileExists(ExportFolder & "vcs-options.json")


    '========================
    '   END TESTS
    '========================

    ' Display results
    lblResults.Caption = _
        m_Totals(True) & " tests passed" & vbCrLf & _
        m_Totals(False) & " tests failed"

    If m_Totals(False) = 0 Then
        imgResult.Picture = "button_ok"
    Else
        imgResult.Picture = "button_error"
    End If

    If Err Then Err.Clear

End Sub




'---------------------------------------------------------------------------------------
' Procedure : ShowResult
' Author    : Adam Waller
' Date      : 5/1/2020
' Purpose   : Add the result to the list.
'---------------------------------------------------------------------------------------
'
Private Function ShowResult(strText As String, blnPassed As Boolean)
    Dim strIcon As String
    If blnPassed Then
        strIcon = ChrW(10004)   ' Check
    Else
        strIcon = ChrW(10060)   ' X
    End If
    lstResults.AddItem strIcon & ";" & strText
    m_Totals(blnPassed) = m_Totals(blnPassed) + 1
    DoEvents
End Function


'---------------------------------------------------------------------------------------
' Procedure : cmdEditTests_Click
' Author    : Adam Waller
' Date      : 5/1/2020
' Purpose   : Go to the code where you can edit the tests.
'---------------------------------------------------------------------------------------
'
Private Sub cmdEditTests_Click()
    Dim intLine As Integer
    VBE.MainWindow.Visible = True
    With VBE.VBProjects("VCS Testing").VBComponents("Form_frmMain")
        .Activate
        With .CodeModule
            intLine = 27 + .ProcStartLine("cmdRunTests_Click", vbext_pk_Proc)
            .CodePane.SetSelection intLine, 1, intLine, 1
            .CodePane.Show
        End With
    End With
    AppActivate VBE.MainWindow.Caption
End Sub


'---------------------------------------------------------------------------------------
' Procedure : ExportFolder
' Author    : Adam Waller
' Date      : 5/7/2020
' Purpose   : Return base export folder for testing for source files.
'---------------------------------------------------------------------------------------
'
Private Function ExportFolder() As String
    ExportFolder = CurrentProject.FullName & ".src\"
End Function


Private Sub Form_Load()
    imgResult.Picture = vbNullString
End Sub
