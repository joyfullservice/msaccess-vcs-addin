# Compatibility
This page outlines the types of objects and information that can be exported from a Microsoft Access Database and which items are supported when building a database from source files. If you find that you are missing something in the export or import, please open an issue and link it to the appropriate row in the table.

Legend: ✔️ Fully supported ⚠️ Partially supported ❌ Not supported

## Summary
Most types of objects can be exported and imported using this tool.

|Type     |Export|Import|
|---------|:----:|:----:|
|Tables   |✔️|✔️|
|Queries  |✔️|✔️|
|Forms    |✔️|✔️|
|Reports  |✔️|✔️|
|Macros   |✔️|✔️|
|Modules  |✔️|✔️|
|Database Settings|✔️|✔️|

## Detailed List

If you are looking for a specific type of object or property that you want to export or import, you can refer to the following more comprehensive list. (Updates and additions welcome.)

The *Testing* column indicates whether a test item and testing code has been created in the *Testing.accdb* database to verify that particular item after import from source code. *Test location* indicates where to find the object in the database.

|Type                     |Export|Import|VBA Class|Testing|Test Location|
|-------------------------|:----:|:----:|---------|:-----:|-------------|
|💼 **TABLES**
|Access Table             |✔️|✔️|clsDbTableDef|✔️|tblInternal|
|Extended Properties      |✔️|✔️|clsDbTableDef|✔️|tblInternal|
|Table SQL                |✔️|✔️|clsDbTableDef|✔️|tblInternal|
|Linked Table             |✔️|✔️|clsDbTableDef|✔️|tblLinkedAccess|
|Linked PrimaryKey        |✔️|✔️|clsDbTableDef|✔️|tblLinkedAccess|
|ODBC Table               |✔️|✔️|clsDbTableDef||
|Linked Structure         |✔️|✔️|clsDbTableDef|✔️|tblLinkedAccess|
|Table Data (TDF)         |✔️|✔️|clsDbTableData|✔️|tblInternal.txt|
|Table Data (XML)         |✔️|✔️|clsDbTableData|✔️|tblSaveXML.xml|
|Table Data Macros        |✔️|✔️|clsDbTableDataMacro|✔️|tblSaveXML|
|💼 **QUERIES**
|Designer Layout          |✔️|✔️|clsDbQuery||
|SQL Output               |✔️|✔️|clsDbQuery||
|Pass Through Queries     |✔️|✔️|clsDbQuery||
|💼 **FORMS**
|Form objects             |✔️|✔️|clsDbForm||
|Saved print settings     |✔️|✔️|clsDbReport|Optional
|💼 **REPORTS**
|Report objects           |✔️|✔️|clsDbReport
|Saved print settings     |✔️|✔️|clsDbReport|Optional
|💼 **MACROS**
|Macro objects            |✔️|✔️|clsDbMacro
|💼 **MODULES**
|Standard Modules         |✔️|✔️|clsDbModule
|Class Modules            |✔️|✔️|clsDbModule
|Object Modules           |✔️|✔️|clsDbModule
|Hidden VBE Attributes    |✔️|✔️|clsDbModule
|💼 **DATABASE**
|DAO Properties           |✔️|✔️|clsDbProperty
|Project Properties       |✔️|✔️|clsDbProjProperty
|Object Descriptions      |✔️|✔️|clsDbDocument
|Hidden Attribute         |✔️|✔️|clsDbHiddenAttribute
|Remove Personal Info     |✔️|✔️|clsDbProject
|Application Icon         |✔️|✔️|clsDbProperty
|Embedded Images          |✔️|✔️|clsDbSharedImage
|Saved Imp/Exp Specs      |✔️|✔️|clsDbSavedSpec
|System Imp/Exp Specs     |✔️|✔️|clsDbImexSpec
|Summary Properties       |✔️|✔️|clsDbDocument
|Relationships            |✔️|✔️|clsDbRelation
|Nav. Pane Groups         |✔️|✔️|clsDbNavPaneGroup
|Embedded Office Theme(s) |✔️|✔️|clsDbTheme
|CommandBars              |❌<p>(yet)|❌<p>(yet)|See [Discussion 478](https://github.com/joyfullservice/msaccess-vcs-addin/discussions/478) for more details.
|💼 **VBE PROJECT**
|Project Properties       |✔️|✔️|clsDbVbeProject
|Compilation Arguments    |✔️|✔️|clsDbVbeProject
|GUID References          |✔️|✔️|clsDbVbeReference
|File/Lib References      |✔️|✔️|clsDbVbeReference
|Forms 2.0 documents      |✔️|✔️|clsDbVbeForm
|💼 **ADP PROJECTS**
|Connection Settings      |||clsDbProjProperty
|SQL Functions            |✔️|n/a|clsAdpFunction
|SQL Views                |✔️|n/a|clsAdpServerView
|SQL Stored Procedures    |✔️|n/a|clsAdpProcedure
|SQL Tables               |✔️|n/a|clsAdpTable
|SQL Triggers             |✔️|n/a|clsAdpTrigger
|💼 **OTHER**
|Saved VCS Options        |✔️|✔️|clsOptions


## ADP Projects

Working with *.adp files is very similar to working with regular (MDB) Microsoft Access databases. All of the main database objects can be exported and imported just like MDB files. When it comes to SQL server objects, the object definitions are exported for tracking in source control, but this system does not attempt to modify any SQL server objects. (Hence the n/a note on importing SQL objects.)

## Supported Versions
This add-in is designed to work in Microsoft Access 2010 and newer. Access 2007 users, please see [this issue](https://github.com/joyfullservice/msaccess-vcs-addin/issues/464).

Missing something? Create an issue or a pull request.
