Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : This class extends the IDbComponent class to perform the specific
'           : operations required by this particular object type.
'           : (I.e. The specific way you export or import this component.)
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit


Private m_Report As AccessObject
Private m_AllItems As Collection

' This requires us to use all the public methods and properties of the implemented class
' which keeps all the component classes consistent in how they are used in the export
' and import process. The implemented functions should be kept private as they are called
' from the implementing class, not this class.
Implements IDbComponent


'---------------------------------------------------------------------------------------
' Procedure : Export
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Export the individual database component (table, form, query, etc...)
'---------------------------------------------------------------------------------------
'
Private Sub IDbComponent_Export()

    Dim cDevMode As clsDevMode
    Dim strTempFile As String

    ' Check Save Print Vars settings
    If Options.SavePrintVars Then

        ' Take a little more manual approach on the export so we can grab the
        ' printer settings before sanitizing the file.
        Set cDevMode = New clsDevMode
    
        ' Export to temporary file
        strTempFile = GetTempFile
        
        ' Save as text, then grab and save printer info.
        Application.SaveAsText acReport, m_Report.Name, strTempFile
        cDevMode.LoadFromExportFile strTempFile
        WriteJsonFile Me, cDevMode.GetDictionary, _
            GetPrintVarsFileName(m_Report.Name), "Report Print Settings"
        
        ' Handle UCS conversion if needed
        ConvertUcs2Utf8 strTempFile, IDbComponent_SourceFile
        
        ' Sanitize source file
        SanitizeFile IDbComponent_SourceFile
        
    Else
        ' Simple export of report object
        SaveComponentAsText acReport, m_Report.Name, IDbComponent_SourceFile
    End If

End Sub


'---------------------------------------------------------------------------------------
' Procedure : ImportPrintVars
' Author    : Adam Waller
' Date      : 5/7/2020
' Purpose   : Import the print vars back into the report.
'---------------------------------------------------------------------------------------
'
Public Sub ImportPrintVars(strFile As String)

'    Dim DevModeString As str_DEVMODE
'    Dim tDevMode As type_DEVMODE
'    Dim DevModeExtra As String
'    Dim dFile As Dictionary
'    Dim strReport As String
'
'    Set dFile = ReadJsonFile(strFile)
'    If Not dFile Is Nothing Then
'
'        ' Prepare data structures
'        tDevMode = DictionaryToDevMode(dFile("Items"))
'        LSet DevModeString = tDevMode
'        Mid(DevModeExtra, 1, 94) = DevModeString.RGB
'
'        ' Apply to report
'        strReport = GetObjectNameFromFileName(strFile)
'        DoCmd.Echo False
'        DoCmd.OpenReport strReport, acViewDesign
'        Reports(strReport).PrtDevMode = DevModeExtra
'        DoCmd.Close acReport, strReport, acSaveYes
'    End If

End Sub


'---------------------------------------------------------------------------------------
' Procedure : Import
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Import the individual database component from a file.
'---------------------------------------------------------------------------------------
'
Private Sub IDbComponent_Import(strFile As String)

    Dim strReport As String

    ' Import the report object
    strReport = GetObjectNameFromFileName(strFile)
    LoadComponentFromText acReport, strReport, strFile

    ' Import the print vars if specified
    If Options.SavePrintVars Then
        ImportPrintVars GetPrintVarsFileName(strReport)
    End If

End Sub


'---------------------------------------------------------------------------------------
' Procedure : GetAllFromDB
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Return a collection of class objects represented by this component type.
'---------------------------------------------------------------------------------------
'
Private Function IDbComponent_GetAllFromDB() As Collection

    Dim rpt As AccessObject
    Dim cReport As IDbComponent

    ' Build collection if not already cached
    If m_AllItems Is Nothing Then
        Set m_AllItems = New Collection
        For Each rpt In CurrentProject.AllReports
            Set cReport = New clsDbReport
            Set cReport.DbObject = rpt
            m_AllItems.Add cReport, rpt.Name
        Next rpt
    End If

    ' Return cached collection
    Set IDbComponent_GetAllFromDB = m_AllItems

End Function


'---------------------------------------------------------------------------------------
' Procedure : GetPrintVarsFileName
' Author    : Adam Waller
' Date      : 5/7/2020
' Purpose   : Return the file name used to export/import print vars
'---------------------------------------------------------------------------------------
'
Private Function GetPrintVarsFileName(strReport As String) As String
    GetPrintVarsFileName = IDbComponent_BaseFolder & GetSafeFileName(strReport) & ".json"
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetFileList
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Return a list of file names to import for this component type.
'---------------------------------------------------------------------------------------
'
Private Function IDbComponent_GetFileList() As Collection
    Set IDbComponent_GetFileList = GetFilePathsInFolder(IDbComponent_BaseFolder & "*.bas")
End Function


'---------------------------------------------------------------------------------------
' Procedure : ClearOrphanedSourceFiles
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Remove any source files for objects not in the current database.
'---------------------------------------------------------------------------------------
'
Private Sub IDbComponent_ClearOrphanedSourceFiles()
    ClearFilesByExtension IDbComponent_BaseFolder, "pv"
    If Not Options.SavePrintVars Then ClearFilesByExtension IDbComponent_BaseFolder, "json"
    ClearOrphanedSourceFiles Me, "bas", "json"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : DateModified
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : The date/time the object was modified. (If possible to retrieve)
'           : If the modified date cannot be determined (such as application
'           : properties) then this function will return 0.
'---------------------------------------------------------------------------------------
'
Private Function IDbComponent_DateModified() As Date
    IDbComponent_DateModified = m_Report.DateModified
End Function


'---------------------------------------------------------------------------------------
' Procedure : SourceModified
' Author    : Adam Waller
' Date      : 4/27/2020
' Purpose   : The date/time the source object was modified. In most cases, this would
'           : be the date/time of the source file, but it some cases like SQL objects
'           : the date can be determined through other means, so this function
'           : allows either approach to be taken.
'---------------------------------------------------------------------------------------
'
Private Function IDbComponent_SourceModified() As Date
    If FSO.FileExists(IDbComponent_SourceFile) Then IDbComponent_SourceModified = FileDateTime(IDbComponent_SourceFile)
End Function


'---------------------------------------------------------------------------------------
' Procedure : Category
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Return a category name for this type. (I.e. forms, queries, macros)
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_Category() As String
    IDbComponent_Category = "reports"
End Property


'---------------------------------------------------------------------------------------
' Procedure : BaseFolder
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Return the base folder for import/export of this component.
'---------------------------------------------------------------------------------------
Private Property Get IDbComponent_BaseFolder() As String
    IDbComponent_BaseFolder = Options.GetExportFolder & "reports\"
End Property


'---------------------------------------------------------------------------------------
' Procedure : Name
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Return a name to reference the object for use in logs and screen output.
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_Name() As String
    IDbComponent_Name = m_Report.Name
End Property


'---------------------------------------------------------------------------------------
' Procedure : SourceFile
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Return the full path of the source file for the current object.
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_SourceFile() As String
    IDbComponent_SourceFile = IDbComponent_BaseFolder & GetSafeFileName(m_Report.Name) & ".bas"
End Property


'---------------------------------------------------------------------------------------
' Procedure : Count
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Return a count of how many items are in this category.
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_Count() As Long
    IDbComponent_Count = IDbComponent_GetAllFromDB.Count
End Property


'---------------------------------------------------------------------------------------
' Procedure : ComponentType
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : The type of component represented by this class.
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_ComponentType() As eDatabaseComponentType
    IDbComponent_ComponentType = edbReport
End Property


'---------------------------------------------------------------------------------------
' Procedure : Upgrade
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Run any version specific upgrade processes before importing.
'---------------------------------------------------------------------------------------
'
Private Sub IDbComponent_Upgrade()
    ' No upgrade needed.
End Sub


'---------------------------------------------------------------------------------------
' Procedure : DbObject
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : This represents the database object we are dealing with.
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_DbObject() As Object
    Set IDbComponent_DbObject = m_Report
End Property
Private Property Set IDbComponent_DbObject(ByVal RHS As Object)
    Set m_Report = RHS
End Property


'---------------------------------------------------------------------------------------
' Procedure : SingleFile
' Author    : Adam Waller
' Date      : 4/24/2020
' Purpose   : Returns true if the export of all items is done as a single file instead
'           : of individual files for each component. (I.e. properties, references)
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_SingleFile() As Boolean
    IDbComponent_SingleFile = False
End Property


'---------------------------------------------------------------------------------------
' Procedure : Parent
' Author    : Adam Waller
' Date      : 4/24/2020
' Purpose   : Return a reference to this class as an IDbComponent. This allows you
'           : to reference the public methods of the parent class without needing
'           : to create a new class object.
'---------------------------------------------------------------------------------------
'
Public Property Get Parent() As IDbComponent
    Set Parent = Me
End Property