Attribute VB_Name = "modDatabase"
'---------------------------------------------------------------------------------------
' Module    : modDatabase
' Author    : Adam Waller
' Date      : 12/4/2020
' Purpose   : General functions for interacting with the current database.
'           : (See modVCSUtility for other functions more specific to this add-in.)
'---------------------------------------------------------------------------------------
Option Compare Database
Option Private Module
Option Explicit

Private Const ModuleName As String = "modDatabase"


'---------------------------------------------------------------------------------------
' Procedure : ProjectPath
' Author    : Adam Waller
' Date      : 1/25/2019
' Purpose   : Path/Directory of the current database file.
'---------------------------------------------------------------------------------------
'
Public Function ProjectPath() As String
    ProjectPath = CurrentProject.Path
    If Right$(ProjectPath, 1) <> PathSep Then ProjectPath = ProjectPath & PathSep
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetDBProperty
' Author    : Adam Waller
' Date      : 5/6/2021
' Purpose   : Get a database property (Default to MDB version)
'---------------------------------------------------------------------------------------
'
Public Function GetDBProperty(strName As String, Optional dbs As DAO.Database) As Variant

    Dim prp As Object ' DAO.Property
    Dim oParent As Object
    
    ' Check for database reference
    If Not dbs Is Nothing Then
        Set oParent = dbs.Properties
    Else
        If DatabaseFileOpen Then
            ' Get parent container for properties
            If CurrentProject.ProjectType = acADP Then
                Set oParent = CurrentProject.Properties
            Else
                If dbs Is Nothing Then Set dbs = CurrentDb
                Set oParent = dbs.Properties
            End If
        Else
            ' No database open
            GetDBProperty = vbNullString
            Exit Function
        End If
    End If
    
    ' Look for property by name
    For Each prp In oParent
        If prp.Name = strName Then
            GetDBProperty = prp.Value
            Exit For
        End If
    Next prp
    Set prp = Nothing
    
End Function


'---------------------------------------------------------------------------------------
' Procedure : SetDBProperty
' Author    : Adam Waller
' Date      : 9/1/2017
' Purpose   : Set a database property
'---------------------------------------------------------------------------------------
'
Public Sub SetDBProperty(ByVal strName As String, ByVal varValue As Variant, Optional ByVal prpType As Long = dbText, Optional dbs As DAO.Database)

    Dim prp As Object ' DAO.Property
    Dim blnFound As Boolean
    Dim oParent As Object
    
    ' Properties set differently for databases and ADP projects
    If CurrentProject.ProjectType = acADP Then
        Set oParent = CurrentProject.Properties
    Else
        If dbs Is Nothing Then Set dbs = CurrentDb
        Set oParent = dbs.Properties
    End If
    
    ' Look for property in collection
    For Each prp In oParent
        If prp.Name = strName Then
            ' Check for matching type
            If Not dbs Is Nothing Then
                If prp.Type <> prpType Then
                    ' Remove so we can add it back in with the correct type.
                    dbs.Properties.Delete strName
                    Exit For
                End If
            End If
            blnFound = True
            ' Skip set on matching value
            If prp.Value = varValue Then
                Set dbs = Nothing
            Else
                ' Update value
                prp.Value = varValue
            End If
            Exit Sub
        End If
    Next prp
    
    ' Add new property
    If Not blnFound Then
        If CurrentProject.ProjectType = acADP Then
            CurrentProject.Properties.Add strName, varValue
        Else
            Set prp = dbs.CreateProperty(strName, prpType, varValue)
            dbs.Properties.Append prp
            Set dbs = Nothing
        End If
    End If
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : SetDAOProperty
' Author    : Adam Waller
' Date      : 5/8/2020
' Purpose   : Updates a DAO property, adding if it does not exist or is the wrong type.
'---------------------------------------------------------------------------------------
'
Public Sub SetDAOProperty(objParent As Object, intType As Integer, strName As String, varValue As Variant)

    Dim prp As DAO.Property
    Dim blnFound As Boolean
    
    ' Look through existing properties.
    For Each prp In objParent.Properties
        If prp.Name = strName Then
            blnFound = True
            Exit For
        End If
    Next prp

    ' Verify type, and update value if found.
    If blnFound Then
        If prp.Type <> intType Then
            objParent.Properties.Delete strName
            blnFound = False
        Else
            If objParent.Properties(strName).Value <> varValue Then
                objParent.Properties(strName).Value = varValue
            End If
        End If
    End If
        
    ' Add new property if needed
    If Not blnFound Then
        ' Create property, then append to collection
        Set prp = objParent.CreateProperty(strName, intType, varValue)
        objParent.Properties.Append prp
    End If

End Sub


'---------------------------------------------------------------------------------------
' Procedure : SetAccessObjectProperty
' Author    : Adam Waller
' Date      : 10/13/2017
' Purpose   : Sets a custom access object property.
'---------------------------------------------------------------------------------------
'
Public Sub SetAccessObjectProperty(objItem As AccessObject, strProperty As String, strValue As String)
    Dim prp As AccessObjectProperty
    For Each prp In objItem.Properties
        If StrComp(prp.Name, strProperty, vbTextCompare) = 0 Then
            ' Update value of property.
            prp.Value = strValue
            Exit Sub
        End If
    Next prp
    ' Property not found. Create it.
    objItem.Properties.Add strProperty, strValue
End Sub


'---------------------------------------------------------------------------------------
' Procedure : GetAccessObjectProperty
' Author    : Adam Waller
' Date      : 10/13/2017
' Purpose   : Get the value of a custom access property
'---------------------------------------------------------------------------------------
'
Public Function GetAccessObjectProperty(objItem As AccessObject, strProperty As String, Optional strDefault As String) As Variant
    Dim prp As AccessObjectProperty
    For Each prp In objItem.Properties
        If StrComp(prp.Name, strProperty, vbTextCompare) = 0 Then
            GetAccessObjectProperty = prp.Value
            Exit Function
        End If
    Next prp
    ' Nothing found. Return default
    GetAccessObjectProperty = strDefault
End Function


'---------------------------------------------------------------------------------------
' Procedure : IsLoaded
' Author    : Adam Waller
' Date      : 9/22/2017
' Purpose   : Returns true if the object is loaded and not in design view.
'---------------------------------------------------------------------------------------
'
Public Function IsLoaded(intType As AcObjectType, strName As String, Optional blnAllowDesignView As Boolean = False) As Boolean

    Dim frm As Form
    Dim rpt As Report
    Dim ctl As Control
    
    If SysCmd(acSysCmdGetObjectState, intType, strName) <> adStateClosed Then
        If blnAllowDesignView Then
            IsLoaded = True
        Else
            Select Case intType
                Case acForm
                    ' Loop through forms collection, since this includes instances
                    ' of add-in forms that cannot be referenced directly by name.
                    For Each frm In Forms
                        If StrComp(frm.Name, strName, vbTextCompare) = 0 Then
                            IsLoaded = frm.CurrentView <> acCurViewDesign
                            Exit For
                        End If
                    Next frm
                Case acReport
                    ' Loop through reports, looking for matching name.
                    For Each rpt In Reports
                        If StrComp(rpt.Name, strName, vbTextCompare) = 0 Then
                            IsLoaded = rpt.CurrentView <> acCurViewDesign
                            Exit For
                        End If
                    Next rpt
                Case acServerView
                    IsLoaded = CurrentData.AllViews(strName).CurrentView <> acCurViewDesign
                Case acStoredProcedure
                    IsLoaded = CurrentData.AllStoredProcedures(strName).CurrentView <> acCurViewDesign
                Case Else
                    ' Other unsupported object
                    IsLoaded = True
            End Select
        End If
    Else
        ' Could be loaded as subform
        If intType = acForm Then
            For Each frm In Forms
                For Each ctl In frm.Controls
                    If TypeOf ctl Is SubForm Then
                        If ctl.SourceObject = strName Then
                            IsLoaded = True
                            Exit For
                        End If
                    End If
                Next ctl
                If IsLoaded Then Exit For
            Next frm
        End If
    End If
    
End Function


'---------------------------------------------------------------------------------------
' Procedure : CloseDatabaseObjects
' Author    : Adam Waller
' Date      : 4/1/2022
' Purpose   : Close any open database objects, returns true if no errors were
'           : encountered. (This is run before a major operation like an export or
'           : a merge.) ADP-specific items aren't particularly supported here.
'---------------------------------------------------------------------------------------
'
Public Function CloseDatabaseObjects() As Boolean

    Dim blnSuccess As Boolean
    Dim objItem As AccessObject
    
    If DebugMode(True) Then On Error GoTo ErrHandler Else On Error GoTo ErrHandler
    
    Perf.OperationStart "Close Open Objects"
    
    ' Check forms and reports
    blnSuccess = CloseAllFormsReports
    
    ' If all forms and reports are closed, proceed with other object types.
    If blnSuccess Then
        
        ' Macros
        For Each objItem In CurrentProject.AllMacros
            If IsLoaded(acMacro, objItem.Name) Then DoCmd.Close acMacro, objItem.Name
        Next objItem
        
        ' Tables
        For Each objItem In CurrentData.AllTables
            If IsLoaded(acTable, objItem.Name) Then DoCmd.Close acTable, objItem.Name
        Next objItem
        
        ' Queries
        For Each objItem In CurrentData.AllQueries
            If IsLoaded(acQuery, objItem.Name) Then DoCmd.Close acQuery, objItem.Name
        Next objItem
        
    End If

    Perf.OperationEnd
    CloseDatabaseObjects = blnSuccess
    
    Exit Function
    
ErrHandler:
    
    blnSuccess = False
    
    ' Handle any error message in calling function
    CatchAny eelNoError, "Unable to close database object", ModuleName & ".CloseDatabaseObjects", False
    
End Function


'---------------------------------------------------------------------------------------
' Procedure : CloseAllFormsReports
' Author    : Adam Waller
' Date      : 1/25/2019
' Purpose   : Close all open forms and reports. Returns true if successful.
'---------------------------------------------------------------------------------------
'
Private Function CloseAllFormsReports() As Boolean

    Dim strName As String
    Dim intOpened As Integer
    Dim intItem As Integer
    
    ' Get count of opened objects
    intOpened = Forms.Count + Reports.Count
    If intOpened > 0 Then
        On Error GoTo ErrorHandler
        ' Loop through forms
        For intItem = Forms.Count - 1 To 0 Step -1
            If Forms(intItem).Caption <> PROJECT_NAME Then
                DoCmd.Close acForm, Forms(intItem).Name
                DoEvents
            End If
            intOpened = intOpened - 1
        Next intItem
        ' Loop through reports
        Do While Reports.Count > 0
            strName = Reports(0).Name
            DoCmd.Close acReport, strName
            DoEvents
            intOpened = intOpened - 1
        Loop
        If intOpened = 0 Then CloseAllFormsReports = True
    Else
        ' No forms or reports currently open.
        CloseAllFormsReports = True
    End If
    
    Exit Function

ErrorHandler:
    Debug.Print "Error closing " & strName & ": " & Err.Number & vbCrLf & Err.Description
End Function


'---------------------------------------------------------------------------------------
' Procedure : ProjectIsSelected
' Author    : Adam Waller
' Date      : 5/15/2015
' Purpose   : Returns true if the base project is selected in the VBE
'---------------------------------------------------------------------------------------
'
Public Function ProjectIsSelected() As Boolean
    ProjectIsSelected = (Application.VBE.SelectedVBComponent Is Nothing)
End Function


'---------------------------------------------------------------------------------------
' Procedure : SelectionInActiveProject
' Author    : Adam Waller
' Date      : 5/15/2015
' Purpose   : Returns true if the current selection is in the active project
'---------------------------------------------------------------------------------------
'
Public Function SelectionInActiveProject() As Boolean
    SelectionInActiveProject = (Application.VBE.ActiveVBProject.FileName = GetUncPath(CurrentProject.FullName))
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetVBProjectForCurrentDB
' Author    : Adam Waller
' Date      : 7/25/2017
' Purpose   : Get the actual VBE project for the current top-level database.
'           : (This is harder than you would think!)
'---------------------------------------------------------------------------------------
'
Public Function GetVBProjectForCurrentDB() As VBProject
    Set GetVBProjectForCurrentDB = GetProjectByName(CurrentProject.FullName)
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetCodeVBProject
' Author    : Adam Waller
' Date      : 4/24/2020
' Purpose   : Get a reference to the VB Project for the running code.
'---------------------------------------------------------------------------------------
'
Public Function GetCodeVBProject() As VBProject
    Set GetCodeVBProject = GetProjectByName(CodeProject.FullName)
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetProjectByName
' Author    : Adam Waller
' Date      : 5/26/2020
' Purpose   : Return the VBProject by file path. (Also supports network drives)
'---------------------------------------------------------------------------------------
'
Private Function GetProjectByName(ByVal strPath As String) As VBProject

    Dim objProj As VBIDE.VBProject
    Dim strUncPath As String
    
    ' Use currently active project by default
    Set GetProjectByName = VBE.ActiveVBProject
    
    ' VBProject filenames are UNC paths
    strUncPath = GetUncPath(strPath)
    
    If VBE.ActiveVBProject.FileName <> strUncPath Then
        ' Search for project with matching filename.
        For Each objProj In VBE.VBProjects
            If objProj.FileName = strUncPath Then
                Set GetProjectByName = objProj
                Exit For
            End If
        Next objProj
    End If
    
End Function


'---------------------------------------------------------------------------------------
' Procedure : RunInCurrentProject
' Author    : Adam Waller
' Date      : 4/22/2020
' Purpose   : Use the Run command but make sure it is running in the context of the
'           : current project, not the add-in file.
'---------------------------------------------------------------------------------------
'
Public Sub RunSubInCurrentProject(strSubName As String)

    Dim strCmd As String
    
    ' Don't need the parentheses after the sub name
    strCmd = Replace(strSubName, "()", vbNullString)
    
    ' Make sure we are not trying to run a function with arguments
    If InStr(strCmd, "(") > 0 Then
        MsgBox2 "Unable to Run Command", _
            "Parameters are not supported for this command.", _
            "If you need to use parameters, please create a wrapper sub or function with" & vbCrLf & _
            "no parameters that you can call instead of " & strSubName & ".", vbExclamation
        Exit Sub
    End If
    
    ' Add project name so we can run it from the current datbase
    strCmd = "[" & GetVBProjectForCurrentDB.Name & "]." & strCmd
    
    ' Run the sub
    Application.Run strCmd

End Sub


'---------------------------------------------------------------------------------------
' Procedure : DatabaseFileOpen
' Author    : Adam Waller
' Date      : 7/14/2020
' Purpose   : Returns true if a database (or ADP project) is currently open.
'---------------------------------------------------------------------------------------
'
Public Function DatabaseFileOpen() As Boolean

    Dim strTest As String
    
    ' See if we have a reference to a CurrentProject object
    If CurrentProject Is Nothing Then
        DatabaseFileOpen = False
    Else
        ' For ADP projects, CurrentProject may be an invalid object reference
        ' after the database file (adp) is closed.
        If DebugMode(True) Then On Error Resume Next Else On Error Resume Next
        strTest = CurrentProject.FullName
        CatchAny eelNoError, vbNullString
        DatabaseFileOpen = (strTest <> vbNullString)
    End If
    
    'DatabaseOpen = Workspaces(0).Databases.Count > 0   ' Another approach (Not ADP compatible)
    
End Function


'---------------------------------------------------------------------------------------
' Procedure : TableExists
' Author    : Adam Waller
' Date      : 5/7/2020
' Purpose   : Returns true if the table object is found in the dabase.
'---------------------------------------------------------------------------------------
'
Public Function TableExists(strName As String) As Boolean
    TableExists = Not (DCount("*", "MSysObjects", "Name=""" & strName & """ AND Type=1") = 0)
End Function


'---------------------------------------------------------------------------------------
' Procedure : DeleteObject
' Author    : Adam Waller
' Date      : 11/23/2020
' Purpose   : Deletes the object if it exists. (Surpresses error)
'---------------------------------------------------------------------------------------
'
Public Sub DeleteObjectIfExists(intType As AcObjectType, strName As String)

    Dim proj As VBProject
    Dim cmpItem As VBComponent
    
    ' Delete the database object
    If DebugMode(True) Then On Error Resume Next Else On Error Resume Next
    DoCmd.DeleteObject intType, strName
    If Not Catch(7874) Then DoEvents  ' Object not found
    
    ' Delete any associated VBE object
    Select Case intType
        Case acForm, acReport, acModule
            Set proj = GetVBProjectForCurrentDB
            With proj
                Set cmpItem = .VBComponents(strName)
                If Not Catch(9) Then
                    .VBComponents.Remove cmpItem
                    Set cmpItem = Nothing
                    DoEvents
                End If
            End With
    End Select
    
    ' Trap another other unexpected errors
    CatchAny eelError, "Deleting object " & strName
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : DbVersion
' Author    : Adam Waller
' Date      : 5/4/2021
' Purpose   : Return the database version as an integer. Works in non-English locales
'           : where CInt(CurrentDb.Version) doesn't work correctly.
'---------------------------------------------------------------------------------------
'
Public Function DbVersion() As Integer
    DbVersion = CInt(Split(CurrentDb.Version, ".")(0))
End Function


'---------------------------------------------------------------------------------------
' Procedure : FormLoaded
' Author    : Adam Waller
' Date      : 7/8/2021
' Purpose   : Helps identify if a form has been closed, but is still running code
'           : after the close event.
'---------------------------------------------------------------------------------------
'
Public Function FormLoaded(frmMe As Form) As Boolean
    Dim strName As String
    ' If no forms are open, we already have our answer.  :-)
    If Forms.Count > 0 Then
        ' We will throw an error accessing the name property if the form is closed
        If DebugMode(True) Then On Error Resume Next Else On Error Resume Next
        strName = frmMe.Name
        ' Return true if we were able to read the name property
        FormLoaded = strName <> vbNullString
    End If
End Function


'---------------------------------------------------------------------------------------
' Procedure : VerifyFocus
' Author    : Adam Waller
' Date      : 7/8/2021
' Purpose   : Verify that a control currently has the focus. (Is the active control)
'---------------------------------------------------------------------------------------
'
Public Function VerifyFocus(ctlWithFocus As Control) As Boolean

    Dim frmParent As Form
    Dim objParent As Object
    Dim ctlCurrentFocus As Control
    
    ' Determine parent form for control
    Set objParent = ctlWithFocus
    Do While Not TypeOf objParent Is Form
        Set objParent = objParent.Parent
    Loop
    Set frmParent = objParent
    
    ' Ignore any errors with Screen.* functions
    If DebugMode(True) Then On Error Resume Next Else On Error Resume Next
    
    ' Verify focus of parent form
    Set frmParent = Screen.ActiveForm
    If Not frmParent Is objParent Then
        Set frmParent = objParent
        frmParent.SetFocus
        DoEvents
    End If
    
    ' Verify focus of control on form
    Set ctlCurrentFocus = frmParent.ActiveControl
    If Not ctlCurrentFocus Is ctlWithFocus Then
        ctlWithFocus.SetFocus
        DoEvents
    End If
    
    ' Return true if the control currently has the focus
    VerifyFocus = frmParent.ActiveControl Is ctlWithFocus
    
    ' Discard any errors
    CatchAny eelNoError, vbNullString, , False
    
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetAvailableConnectionCount
' Author    : Adam Waller
' Date      : 6/8/2022
' Purpose   : Returns the number of available connections in the current database.
'           : (Access has a maximum of 255 connections that can be made to the current
'           :  database object.) See issue #338
'---------------------------------------------------------------------------------------
'
Public Function GetAvailableConnectionCount()

    Dim colDbs As Collection
    Dim dbs As DAO.Database
    Dim intCnt As Integer
    
    Set colDbs = New Collection
    
    On Error Resume Next
    For intCnt = 1 To 300
        Set dbs = CurrentDb
        If Err Then
            ' Probably cannot open any more databases
            Err.Clear
            Exit For
        End If
        colDbs.Add dbs
    Next intCnt
    
    Set dbs = Nothing

    ' Return count of how many connections we were
    ' able to create before hitting an error
    GetAvailableConnectionCount = intCnt
    
End Function
