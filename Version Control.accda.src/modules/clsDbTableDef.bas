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

Private m_Table As AccessObject
Private m_Options As clsOptions
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
    
    Dim strFile As String
    Dim strTempFile As String
    Dim dbs As Database

    ' Check for existing file
    strFile = IDbComponent_SourceFile
    If FSO.FileExists(strFile) Then Kill strFile

    ' Save structure in XML format
    Application.ExportXML acExportTable, m_Table.Name, , strFile

    ' Optionally save in SQL format
    If IDbComponent_Options.SaveTableSQL Then
        Set dbs = CurrentDb
        Log.Add "  " & m_Table.Name & " (SQL)", IDbComponent_Options.ShowDebug
        SaveTableSqlDef dbs, m_Table.Name, IDbComponent_BaseFolder, IDbComponent_Options
    End If

End Sub


'---------------------------------------------------------------------------------------
' Procedure : SaveTableSqlDef
' Author    : Adam Waller
' Date      : 1/28/2019
' Purpose   : Save a version of the table formatted as a SQL statement.
'           : (Makes it easier to see table changes in version control systems.)
'---------------------------------------------------------------------------------------
'
Public Sub SaveTableSqlDef(dbs As DAO.Database, strTable As String, strFolder As String, cOptions As clsOptions)

    Dim cData As New clsConcat
    Dim cAttr As New clsConcat
    Dim idx As DAO.Index
    Dim fld As DAO.Field
    Dim strFile As String
    Dim tdf As DAO.TableDef

    Set tdf = dbs.TableDefs(strTable)

    With cData
        .Add "CREATE TABLE ["
        .Add strTable
        .Add "] ("
        .Add vbCrLf

        ' Loop through fields
        For Each fld In tdf.Fields
            .Add "  ["
            .Add fld.Name
            .Add "] "
            If (fld.Attributes And dbAutoIncrField) Then
                .Add "AUTOINCREMENT"
            Else
                .Add GetTypeString(fld.Type)
                .Add " "
            End If
            Select Case fld.Type
                Case dbText, dbVarBinary
                    .Add "("
                    .Add fld.Size
                    .Add ")"
            End Select

            ' Indexes
            For Each idx In tdf.Indexes
                Set cAttr = New clsConcat
                If idx.Fields.Count = 1 And idx.Fields(0).Name = fld.Name Then
                    If idx.Primary Then cAttr.Add " PRIMARY KEY"
                    If idx.Unique Then cAttr.Add " UNIQUE"
                    If idx.Required Then cAttr.Add " NOT NULL"
                    If idx.Foreign Then AddFieldReferences dbs, idx.Fields, strTable, cAttr
                    If Len(cAttr.GetStr) > 0 Then
                        .Add " CONSTRAINT ["
                        .Add idx.Name
                        .Add "]"
                    End If
                End If
                .Add cAttr.GetStr
            Next
            .Add ","
            .Add vbCrLf
        Next fld
        .Remove 3   ' strip off last comma and crlf

        ' Constraints
        Set cAttr = New clsConcat
        For Each idx In tdf.Indexes
            If idx.Fields.Count > 1 Then
                If Len(cAttr.GetStr) = 0 Then cAttr.Add " CONSTRAINT "
                If idx.Primary Then
                    cAttr.Add "["
                    cAttr.Add idx.Name
                    cAttr.Add "] PRIMARY KEY ("
                    For Each fld In idx.Fields
                        cAttr.Add fld.Name
                        cAttr.Add ", "
                    Next fld
                    cAttr.Remove 2
                    cAttr.Add ")"
                End If
                If Not idx.Foreign Then
                    If Len(cAttr.GetStr) > 0 Then
                        .Add ","
                        .Add vbCrLf
                        .Add "  "
                        .Add cAttr.GetStr
                        AddFieldReferences dbs, idx.Fields, strTable, cData
                    End If
                End If
            End If
        Next
        .Add vbCrLf
        .Add ")"

        ' Build file name and create file.
        strFile = strFolder & GetSafeFileName(strTable) & ".sql"
        WriteFile .GetStr, strFile

    End With

End Sub


'---------------------------------------------------------------------------------------
' Procedure : AddFieldReferences
' Author    : Adam Waller
' Date      : 1/18/2019
' Purpose   : Add references to other fields in table definition.
'---------------------------------------------------------------------------------------
'
Private Sub AddFieldReferences(dbs As Database, fld As Object, strTable As String, cData As clsConcat)

    Dim rel As DAO.Relation
    Dim fld2 As DAO.Field

    For Each rel In dbs.Relations
        If (rel.ForeignTable = strTable) Then
            If FieldsIdentical(fld, rel.Fields) Then

                ' References
                cData.Add " REFERENCES "
                cData.Add rel.Table
                cData.Add " ("
                For Each fld2 In rel.Fields
                    cData.Add fld2.Name
                    cData.Add ","
                Next fld2
                ' Remove trailing comma
                If rel.Fields.Count > 0 Then cData.Remove 1
                cData.Add ")"

                ' Attributes for cascade update or delete
                If rel.Attributes And dbRelationUpdateCascade Then cData.Add " ON UPDATE CASCADE "
                If rel.Attributes And dbRelationDeleteCascade Then cData.Add " ON DELETE CASCADE "

                ' Exit now that we have found the matching relationship.
                Exit For

            End If
        End If
    Next rel

End Sub


'---------------------------------------------------------------------------------------
' Procedure : FieldsIdentical
' Author    : Adam Waller
' Date      : 1/21/2019
' Purpose   : Return true if the two collections of fields have the same field names.
'           : (Even if the order of the fields is different.)
'---------------------------------------------------------------------------------------
'
Private Function FieldsIdentical(oFields1 As Object, oFields2 As Object) As Boolean

    Dim fld As Object
    Dim fld2 As Object
    Dim blnMismatch As Boolean
    Dim blnFound As Boolean

    If oFields1.Count <> oFields2.Count Then
        blnMismatch = True
    Else
        ' Set this flag to false after going through each field.
        For Each fld In oFields1
            blnFound = False
            For Each fld2 In oFields2
                If fld.Name = fld2.Name Then
                    blnFound = True
                    Exit For
                End If
            Next fld2
            If Not blnFound Then
                blnMismatch = True
                Exit For
            End If
        Next
    End If

    ' Return result
    FieldsIdentical = Not blnMismatch

End Function


'---------------------------------------------------------------------------------------
' Procedure : GetTypeString
' Author    : Adam Waller
' Date      : 1/18/2019
' Purpose   : Get the type string used by Access SQL
'---------------------------------------------------------------------------------------
'
Private Function GetTypeString(intType As DAO.DataTypeEnum) As String
    Select Case intType
        Case dbLongBinary:      GetTypeString = "LONGBINARY"
        Case dbBinary:          GetTypeString = "BINARY"
        Case dbBoolean:         GetTypeString = "BIT"
        Case dbAutoIncrField:   GetTypeString = "COUNTER"
        Case dbCurrency:        GetTypeString = "CURRENCY"
        Case dbDate, dbTime:    GetTypeString = "DATETIME"
        Case dbGUID:            GetTypeString = "GUID"
        Case dbMemo:            GetTypeString = "LONGTEXT"
        Case dbDouble:          GetTypeString = "DOUBLE"
        Case dbSingle:          GetTypeString = "SINGLE"
        Case dbByte:            GetTypeString = "UNSIGNED BYTE"
        Case dbInteger:         GetTypeString = "SHORT"
        Case dbLong:            GetTypeString = "LONG"
        Case dbNumeric:         GetTypeString = "NUMERIC"
        Case dbText:            GetTypeString = "VARCHAR"
        Case Else:              GetTypeString = "VARCHAR"
    End Select
End Function


'---------------------------------------------------------------------------------------
' Procedure : Import
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Import the individual database component from a file.
'---------------------------------------------------------------------------------------
'
Private Sub IDbComponent_Import(strFile As String)

End Sub


'---------------------------------------------------------------------------------------
' Procedure : GetAllFromDB
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Return a collection of class objects represented by this component type.
'---------------------------------------------------------------------------------------
'
Private Function IDbComponent_GetAllFromDB(Optional cOptions As clsOptions) As Collection
    
    Dim dbs As Database
    Dim tdf As TableDef
    Dim cTable As IDbComponent

    ' Build collection if not already cached
    If m_AllItems Is Nothing Then

        ' Use parameter options if provided.
        If Not cOptions Is Nothing Then Set IDbComponent_Options = cOptions
    
        Set m_AllItems = New Collection
        Set dbs = CurrentDb
            
        For Each tdf In dbs.TableDefs
            ' Skip system, temp, and linked tables
            If Left$(tdf.Name, 4) <> "MSys" Then
                If Left$(tdf.Name, 1) <> "~" Then
                    If Len(tdf.connect) = 0 Then
                        Set cTable = New clsDbTableDef
                        Set cTable.DbObject = CurrentData.AllTables(tdf.Name)
                        Set cTable.Options = IDbComponent_Options
                        m_AllItems.Add cTable, tdf.Name
                    End If
                End If
            End If
        Next tdf
        
        Set tdf = Nothing
        Set dbs = Nothing

    End If

    ' Return cached collection
    Set IDbComponent_GetAllFromDB = m_AllItems
        
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetFileList
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Return a list of file names to import for this component type.
'---------------------------------------------------------------------------------------
'
Private Function IDbComponent_GetFileList() As Collection
    Set IDbComponent_GetFileList = GetFilePathsInFolder(IDbComponent_BaseFolder & "*.xml")
End Function


'---------------------------------------------------------------------------------------
' Procedure : ClearOrphanedSourceFiles
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Remove any source files for objects not in the current database.
'---------------------------------------------------------------------------------------
'
Private Function IDbComponent_ClearOrphanedSourceFiles() As Variant
    ClearOrphanedSourceFiles Me, "LNKD", "bas", "sql", "xml", "tdf", "json"
End Function


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
    IDbComponent_DateModified = m_Table.DateModified
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
Public Function IDbComponent_SourceModified() As Date
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
    IDbComponent_Category = "tables"
End Property


'---------------------------------------------------------------------------------------
' Procedure : BaseFolder
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Return the base folder for import/export of this component.
'---------------------------------------------------------------------------------------
Private Property Get IDbComponent_BaseFolder() As String
    IDbComponent_BaseFolder = IDbComponent_Options.GetExportFolder & "tbldefs\"
End Property


'---------------------------------------------------------------------------------------
' Procedure : Name
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Return a name to reference the object for use in logs and screen output.
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_Name() As String
    IDbComponent_Name = m_Table.Name
End Property


'---------------------------------------------------------------------------------------
' Procedure : SourceFile
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Return the full path of the source file for the current object.
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_SourceFile() As String
    IDbComponent_SourceFile = IDbComponent_BaseFolder & GetSafeFileName(m_Table.Name) & ".xml"
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
    IDbComponent_ComponentType = edbTableDef
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
' Procedure : Options
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Return or set the options being used in this context.
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_Options() As clsOptions
    If m_Options Is Nothing Then Set m_Options = LoadOptions
    Set IDbComponent_Options = m_Options
End Property
Private Property Set IDbComponent_Options(ByVal RHS As clsOptions)
    Set m_Options = RHS
End Property


'---------------------------------------------------------------------------------------
' Procedure : DbObject
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : This represents the database object we are dealing with.
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_DbObject() As Object
    Set IDbComponent_DbObject = m_Table
End Property
Private Property Set IDbComponent_DbObject(ByVal RHS As Object)
    Set m_Table = RHS
End Property


'---------------------------------------------------------------------------------------
' Procedure : SingleFile
' Author    : Adam Waller
' Date      : 4/24/2020
' Purpose   : Returns true if the export of all items is done as a single file instead
'           : of individual files for each component. (I.e. properties, references)
'---------------------------------------------------------------------------------------
'
Public Property Get IDbComponent_SingleFile() As Boolean
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