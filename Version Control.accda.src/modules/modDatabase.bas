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
' Procedure : ObjectExists
' Author    : Adam Waller
' Date      : 3/3/2023
' Purpose   : See if the object exists in the current database/project
'---------------------------------------------------------------------------------------
'
Public Function ObjectExists(intType As AcObjectType, strName As String, Optional blnInCodeDb As Boolean = False) As Boolean

    Dim objTest As Object
    Dim objContainer As Object

    Set objContainer = GetParentContainer(intType, blnInCodeDb)
    If objContainer Is Nothing Then
        Log.Error eelError, "Parent container not supported for this object type: " & intType, ModuleName & ".ObjectExists"
    Else
        ' Attempt to reference the object by name
        If DebugMode(True) Then On Error Resume Next Else On Error Resume Next
        Set objTest = objContainer(strName)
        ObjectExists = Not Catch(2467)
    End If

End Function


'---------------------------------------------------------------------------------------
' Procedure : GetParentContainer
' Author    : Adam Waller
' Date      : 3/3/2023
' Purpose   : Get the parent container collection for the object type. (Not all types
'           : are supported.)
'           : Also allows you to specify whether to use the CodeDb or CurrentDb
'---------------------------------------------------------------------------------------
'
Public Function GetParentContainer(intType As AcObjectType, Optional blnInCodeDb As Boolean = False) As Object

    Dim objHostData As CurrentData
    Dim objHostProject As CurrentProject

    ' Set the host objects
    If blnInCodeDb Then
        Set objHostData = CurrentData
        Set objHostProject = CurrentProject
    Else
        Set objHostData = CodeData
        Set objHostProject = CodeProject
    End If

    ' Return the associated parent container
    Select Case intType

        ' ADP Specific
        Case acDiagram:             Set GetParentContainer = objHostData.AllDatabaseDiagrams
        Case acFunction:            Set GetParentContainer = objHostData.AllFunctions
        Case acServerView:          Set GetParentContainer = objHostData.AllViews
        Case acStoredProcedure:     Set GetParentContainer = objHostData.AllStoredProcedures

        ' Database objects
        Case acForm:                Set GetParentContainer = objHostProject.AllForms
        Case acMacro:               Set GetParentContainer = objHostProject.AllMacros
        Case acModule:              Set GetParentContainer = objHostProject.AllModules
        Case acQuery:               Set GetParentContainer = objHostData.AllQueries
        Case acReport:              Set GetParentContainer = objHostProject.AllReports
        Case acTable:               Set GetParentContainer = objHostData.AllTables

        ' Unsupported
        Case acTableDataMacro, acDatabaseProperties
        Case Else

    End Select

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
    strUncPath = UCase(GetUncPath(strPath))

    If UCase(VBE.ActiveVBProject.FileName) <> strUncPath Then
        ' Search for project with matching filename.
        For Each objProj In VBE.VBProjects
            If UCase(objProj.FileName) = strUncPath Then
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

    ' Make sure procedure exists in current database
    If Not GlobalProcExists(strSubName) Then
        Log.Error eelError, "The procedure """ & strSubName & """ not found.", ModuleName & ".RunSubInCurrentProject"
        Log.Add "The procedure must be declared as public in a standard module.", False
        Exit Sub
    End If
    
    ' Add project name so we can run it from the current datbase
    strCmd = "[" & CurrentVBProject.Name & "]." & strCmd

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
' Purpose   : Returns true if the table object is found in the dabase. (SQL version)
'---------------------------------------------------------------------------------------
'
Public Function TableExists(strName As String) As Boolean
    TableExists = Not (DCount("*", "MSysObjects", "Name=""" & strName & """ AND Type=1") = 0)
End Function


'---------------------------------------------------------------------------------------
' Procedure : DeleteObjectIfExists
' Author    : Adam Waller
' Date      : 3/3/2023
' Purpose   : Deletes the object if it exists. This gets really tricky if the object
'           : exists in both the add-in database and the current database. Thankfully
'           : we have a way to work around this for most objects by renaming to a
'           : unique name that doesn't exist in the add-in, and deleting the renamed
'           : object.
'---------------------------------------------------------------------------------------
'
Public Function DeleteObjectIfExists(intType As AcObjectType, strName As String) As Boolean

    Dim blnExistsInAddIn As Boolean
    Dim strTempName As String

    ' If object does not exist in the current database, no need to go further
    If Not ObjectExists(intType, strName) Then Exit Function

    ' Check to see if the object exists in the add-in database. (See note above)
    Select Case intType
        ' Object types used in the add-in
        Case acForm, acMacro, acModule, acQuery, acTable
            blnExistsInAddIn = ObjectExists(intType, strName, True)
    End Select

    ' Trap errors when attempting to delete the object
    If DebugMode(True) Then On Error Resume Next Else On Error Resume Next

    If Not blnExistsInAddIn Then
        ' Nice! We can use a simple call to delete the object
        DoCmd.DeleteObject intType, strName
    Else
        ' This is where it gets fun... If you attempt to delete an object from the
        ' VBA code in the add-in, it will default to operating on the add-in object
        ' first, before looking in the current database. This can cause corruption
        ' in the add-in, in addition to failing to delete the object from the
        ' current database. To work around this, we will rename the object to something
        ' random first, then delete it based on the randomized (non-matching) name.

        ' Build a random temp name that will not collide with the add-in or any existing
        ' object in the current database. (But leave a meaningful clue in the name, in
        ' case something goes wrong and it ends up staying in the database.)
        strTempName = strName & "_DELETE_" & GetStringHash(Perf.MicroTimer)

        ' We need to avoid using DoCmd.Rename for the same reasons
        Select Case intType
            Case acForm
                Log.Error eelError, "Cannot delete a form with the same name as an add-in form.", _
                    ModuleName & ".DeleteObjectIfExists"
                    Exit Function   ' (Returns False)
            Case acMacro
                ' The rename command seems to work on this object type... (At least
                ' in Access 2010)
                DoCmd.Rename strTempName, intType, strName
            Case acModule
                ' Rename the VBE object
                CurrentVBProject.VBComponents(strName).Name = strTempName
            Case acQuery
                ' Rename the Query Definition object
                CurrentDb.QueryDefs(strName).Name = strTempName
            Case acTable
                ' Rename the Table Definition object
                CurrentDb.TableDefs(strName).Name = strTempName
        End Select

        ' Trap any errors involved in renaming the object
        If Not CatchAny(eelError, "Error renaming object: " & strName, ModuleName & ".DeleteObjectIfExists") Then
            ' Delete object using the temp name
            DoCmd.DeleteObject intType, strTempName
        End If
    End If

    ' Catch any errors with deleting the object
    CatchAny eelError, "Error deleting object: " & strName, ModuleName & ".DeleteObjectIfExists"

    ' Return success if the object no longer exists
    DeleteObjectIfExists = Not ObjectExists(intType, strName)

End Function


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
