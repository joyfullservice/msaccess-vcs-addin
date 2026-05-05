Note: This page is a work in progress. Feel free to contribute!

# File Structure Overview
The following outline illustrates the folder and file structure used by the add-in when exporting source files for a project. (For simplicity, ADP project folders are not shown here.)

Note that in your export folder you will only see the folders and files for the components actually used in your database. See [Supported Objects](Supported-Objects) for a comprehensive list of object types and the related code sections in this add-in.

    .
    ├── ...
    ├── .gitattributes                  # (Recommended for git) Configure git attributes
    ├── .gitignore                      # (Recommended for git) Ignore specific files in git
    ├── Database.accdb                  # Main database file
    └── Database.accdb.src              # Folder for source file export (default path)
        ├── forms                       # Access forms
        │   ├── MyForm.form             # Object definition (legacy: .bas)
        │   └── MyForm.json             # Custom print settings (if used)
        ├── images                      # Linked images
        │   ├── MyPicture.png           # Image file
        │   └── MyPicture.json          # Image metadata
        ├── imexspecs                   # Legacy table-based import/export specifications
        │   └── LinkedTableSpec.json    # Json representation of specification
        ├── macros                      # Macros
        │   └── AutoExec.macro          # Object definition (legacy: .bas)
        ├── modules                     # VBA modules (standard and class)
        │   ├── MyModule.bas            # Standard module
        │   └── MyClass.cls             # Class module
        ├── queries                     # Queries
        │   ├── MyQuery.sql             # SQL statement (source of truth)
        │   └── MyQuery.json            # Query metadata and design layout
        ├── relations                   # Database relationships
        │   └── Table1Table2.json       # Json representation of a single relationship
        ├── reports                     # Reports
        │   ├── MyReport.report         # Object definition (legacy: .bas)
        │   └── MyReport.json           # Print settings
        ├── savedspecs                  # XML-based import/export specifications
        │   └── MyExport.json           # Metadata with embedded XML
        ├── tables                      # Table data
        │   ├── MyTable.txt             # Tab-delimited text format (more readable)
        │   └── MyTable.xml             # XML table data (better data preservation)
        ├── tbldefs                     # Table definitions (structure, columns)
        │   ├── MyTable.sql             # A SQL rendition of the table structure (more readable)
        │   └── MyTable.xml             # XML object definition (for export/import) 
        ├── tdmacros                    # Table Data Macros
        │   └── MyTable.xml             # Object definition
        ├── themes                      # Database themes (visual appearance)
        │   ├── Angles.thmx             # Self-contained zip file
        │   └── Executive               # (Alternate) Extracted contents folder
        │       ├── _rels               # Recommended when customizing the theme
        │       ├── theme               # files in a database project
        │       └── ...                 # 
        ├── vbeforms                    # MSForms 2.0 Forms (Traditional VBA Forms)
        │   ├── MyForm.frm              # Form definition stub
        │   ├── MyForm.frx              # Binary OLE Blob content and structure
        │   └── MyForm.json             # Json representation of content and structure
        ├── Build.log                   # VCS build log
        ├── dbs-properties.json         # DAO database properties
        ├── documents.json              # Document container metadata
        ├── Export.log                  # VCS export log
        ├── hidden-attributes.json      # Hidden flag attributes
        ├── nav-pane-groups.json        # Custom navigation pane groups
        ├── project.json                # File format and Remove Personal Info flag
        ├── proj-properties.json        # CurrentProject properties
        ├── vbe-project.json            # VBA project properties
        ├── vbe-references.json         # VBA project references
        ├── vcs-index.idx               # Binary index to track changes and versions for VCS
        └── vcs-options.json            # VCS options for this project


# Main Folder
These files are present in the MSAccessVCS-Addin project; they may or may not be present in your repository depending on how you have yours set up.

|File |Present When |Details & Description
|:-|:-:|:-
|`.gitattributes`|Always|Tells `git` and `github` what each file type does, and how to handle line endings, etc. See [git-scm gitattributes](https://git-scm.com/docs/gitattributes) for more information.
|`.gitignore`|Always|Tells `git` about any specific files, file types, or folders that should be ignored (ensures you don't commit stuff that shouldn't be there). See [git-scm gitignore](https://git-scm.com/docs/gitignore) for more information
|Contributing.md|Always (Advised)|Front page information about how to contribute to the MSAccessVCS-Addin project!
|License.txt|Always (Advised)|Outlines the requirements to share this code. Include in your repository if you share it.
|README.md|Always (MSAccessVCS-Addin)|Landing page for MSAccessVCS-Addin on github.
|`Version Control.accda`|MSAccessVCS-Addin |This is the versioned tool; your file will be located in here as well. 
|`Version Control_VCSBackup[n].accda`|MSAccessVCS-Addin|When you build from source, the existing file is backedup. The first one is simply named `[Your Database]_VCSBackup.[extension]`. Subsequent ones are numbered `[Your Database]_VCSBackup[1,2,3,...,n].[extension]` and so forth.

## [YourSource.extension].src
The source code for the exported database is (by default) stored in the same folder as the database, with the name `[YourSource.extension].src`. See Options for more information on this if you want to move it elsewhere.

The below are the default folders for export/import. There may be future work to change this; see [Issue 222 "Support Rubberduck @Folder annotations in modules"](https://github.com/joyfullservice/msaccess-vcs-integration/issues/222) if you want to help with this.


### [YourSource.extension].src\forms
Forms are exported into the `forms` folder by default. Each form will be exported as `[FormName].form`.

|File or Type|Present When |Details & Description
|:-|:-:|:-
|`*.form`|Form Export |Object definition for each form present in the database. (Legacy projects may use `*.bas` instead; both are accepted on import.)
|`*.cls`|Code Behind Form Export |Exported for each form present in the database. **Only exported when [Options/Export/ "Split Layout from VBA"](Options#export-tab) is _ON_.**
|`*.json`|Form Print settings |Only exported if the form has print settings (rare). Contains printer formatting and page settings.


### [YourSource.extension].src\images
Images stored in `MSysResources` are exported into the `images` folder by default. 

|File or Type|Present When |Details & Description
|:-|:-:|:-
|`*.[imgext]`|Image Export |Exported for each Image present in the database. Exported with file name and type.
|`*.json`|Image details and attributes |Exported for each image. Contains meta data such as the image mouseover name, internal name of the file, extension, and handling class (export/import handler).


### [YourSource.extension].src\macros
Macros are exported into the `macros` folder by default.

|File or Type|Present When |Details & Description
|:-|:-:|:-
|`*.macro`|Macro Export |Object definition for each macro present in the database. (Legacy projects may use `*.bas` instead; both are accepted on import.)


### [YourSource.extension].src\modules
VBA modules are exported into the `modules` folder by default.

|File or Type|Present When |Details & Description
|:-|:-:|:-
|`*.bas`|Code Module Export |Exported for each standard module present in the database. Contains source code and other attributes.
|`*.cls`|Code Class Module Export |Exported for each class module present in the database. Contains source code and other attributes such as default member, internal VBA name, etc.


### [YourSource.extension].src\queries
Queries are exported into the `queries` folder by default. Each query is exported as a pair of files: a `.sql` file containing the SQL statement and a `.json` file containing metadata.

|File or Type|Present When |Details & Description
|:-|:-:|:-
|`*.sql`|Each Query |The SQL statement, formatted for readability. This is the **source of truth** for the query's SQL text. Edit this file directly to change query logic.
|`*.json`|Each Query |Companion metadata that cannot be expressed in SQL: query type, query properties, column metadata, and Design View layout (table positions, window dimensions). If the SQL is designer-compatible and layout data is present, the query is imported in Design View format, preserving table positions. Otherwise it uses SQL View format.

Legacy `*.qdef` and `*.bas` query files from older export formats are still accepted on import. On the next export they are automatically replaced with the `.sql` + `.json` pair.


### [YourSource.extension].src\reports
Reports are exported into the `reports` folder by default.

|File or Type|Present When |Details & Description
|:-|:-:|:-
|`*.report`|Reports Export |Object definition for each report present in the database. (Legacy projects may use `*.bas` instead; both are accepted on import.)
|`*.cls`|Code Behind Report Export |Exported for each report present in the database. **Only exported when [Options / Export: "Split Layout from VBA"](Options#export-tab) is _ON_.**
|`*.json`|Report Print settings |Only exported if the report has print settings (nearly always for reports). Contains printer formatting and page settings.


### [YourSource.extension].src\tables
Table data is exported into `tables` folder by for each table you setup to export data. See [Options](Options#table-data) for what you can choose.

|File or Type|Present When |Details & Description
|:-|:-:|:-
|`*.xml`|XML Data Export |There will be one `.XML` file for each table setup to have the data exported as XML.
|`*.TDF`|TDF Data Export |There will be one `.TDF` file for each table setup to have the data exported as TDF. This is more human readable, but may lose some information. **Not reccomended if "Special Characters" are present.**


### [YourSource.extension].src\tbldefs
Table definitions are exported into the `tbldefs` folder by default. Note table data is not part of the table definition. Table data is stored elsewhere.

|File or Type|Present When |Details & Description
|:-|:-:|:-
|`JSON` |Linked Tables Only|Contains additional table attributes (should be all, but I'm sure there's probably an edge case that's not been caught yet), such as connection string, the Access table name, the linked table name (on the server), a list of attributes (it's a `long` type that's used in `bitwise` fashion), and a few other bits to cue MSAccessVCS how to import the table.
|`XML` |Local Tables Only|Table format and extra information (such as row display height, column display width, required/not, column formatting information (Plain Text/RTF, Date Time formatting, etc.) and so on). It is the most 'reliable' and full export of the local table as it (usually) contains all the meta data about the table. `XML` ensures pretty much all the format and data will be the same when you rebuild. If you export the table data, too (it's stored in a separate folder), then `XML` is the way to go (usually). You can also export Table data via `TDF` (Tab Delimited Format) if you want something more human-readable. `XML` by its nature is not as row or character dependent. Access also isn't careful at exporting the same components in the same way every time, so the export is mildly prone to code noise during export (AKA: `git` will think there's changes, when there aren't).
|`SQL` |Local Tables Only|If you were to generate a table via SQL, this would be the SQL you'd use. Because Access SQL (SQL Jet) doesn't care about formatting/display width/display height, if you use this, the table won't "look" the same when you open it in Access.


### [YourSource.extension].src\themes
Themes are exported into the `themes` folder by default. See [Options](Options#export-tab) for what you can choose. 

|File or Type|Present When |Details & Description
|:-|:-:|:-
|`*.thmx`|Extract theme files is *OFF*|If you don't export the Themes, each theme in your database will be extracted separately in its own `.thmx` file. This is simply a `.zip` file with some special handling.
|`\ThemeFolder`|Extract theme files is *ON*|If you extract theme files, each theme will be in its own subfolder; there will be `.XML` files inside as well for the vairious theme components.
