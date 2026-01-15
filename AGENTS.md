# AGENTS.md - AI Agent Guide to MSAccess VCS Add-in Development

## Project Overview

This repository contains the **MSAccess Version Control System (VCS) Add-in** - a Microsoft Access add-in that enables version control for Access database projects. The add-in exports database objects (forms, reports, queries, modules, etc.) to text-based source files suitable for Git or other version control systems, and can rebuild databases entirely from these source files.

**Key capabilities:**
- Export Access database objects to text-based source files
- Build/rebuild a database entirely from source files
- Merge source file changes into an existing database
- Track changes via an index system for fast incremental exports
- Detect and resolve conflicts between database and source file changes
- Support for ADP projects and external SQL database schema export

---

## Repository Structure

| Folder | Purpose |
|--------|---------|
| `Version Control.accda.src/` | **Core add-in source code** - Exported VBA modules, classes, forms, and queries |
| `Ribbon/` | **COM Add-in for ribbon UI** - twinBASIC project providing 64-bit ribbon toolbar support |
| `Hook/` | **Export-on-save hook DLLs** - External library for automatic export when saving objects |
| `Testing/` | **Test database** - Sample database (`Testing.accdb.src`) for testing import/export functionality |
| `Translation/` | **Localization files** - `.pot` and `.po` files for UI translation support |
| `Wiki/` | **Documentation** - Markdown files synced to GitHub Wiki |
| `Template/` | **Database template** - Binary template used when creating new databases |
| `img/` | **README images** - Screenshots and demos for documentation |

---

## Architecture Overview

### Component Diagram

```
┌────────────────────────────────────────────────────────────────────┐
│                        Microsoft Access                            │
├────────────────────────────────────────────────────────────────────┤
│  ┌─────────────────────┐    ┌──────────────────────────────────┐   │
│  │  COM Ribbon Add-in  │───▶│  Version Control.accda (Add-in)  │   │
│  │  (twinBASIC DLLs)   │    │  ┌────────────────────────────┐  │   │
│  │  - MSAccessVCSLib   │    │  │ clsVersionControl (API)    │  │   │
│  └─────────────────────┘    │  │ modImportExport (Core)     │  │   │
│                             │  │ IDbComponent (Interface)   │  │   │
│  ┌─────────────────────┐    │  │ clsDb* (Component Classes) │  │   │
│  │  Hook DLLs          │    │  │ clsOptions, clsVCSIndex    │  │   │
│  │  - Export on Save   │───▶│  └────────────────────────────┘  │   │
│  └─────────────────────┘    └──────────────────────────────────┘   │
└────────────────────────────────────────────────────────────────────┘
                                       │
                                       ▼
                          ┌────────────────────────┐
                          │   Source Files (.src)  │
                          │   - forms/*.bas, *.cls │
                          │   - modules/*.bas,*.cls│
                          │   - queries/*.bas,*.sql│
                          │   - vcs-options.json   │
                          │   - vcs-index.json     │
                          └────────────────────────┘
```

### Key Architectural Patterns

1. **Interface-Based Component System**: All database object types implement `IDbComponent`, providing a consistent API for export, import, merge, and metadata operations.

2. **Singleton Pattern for Global State**: Key objects (`Options`, `VCSIndex`, `Log`, `Perf`, `Operation`) are accessed via `modObjects` module-level functions.

3. **Two Types of Build**: Full builds create a new database from source; merge builds update existing databases with changed files only.

4. **Index-Based Change Detection**: `vcs-index.json` tracks file hashes and timestamps to detect changes and enable "fast save" exports.

---

## Core Components

### Public API (`clsVersionControl`)

The primary entry point for external automation. Exposed via the `VCS` object in `modAPI`.

```vba
' Key public methods:
VCS.Export              ' Export all source
VCS.ExportVBA           ' Export VBA components only
VCS.Build strFolder     ' Full build from source
VCS.MergeBuild          ' Merge changes into existing database
VCS.Options             ' Access project options
```

### Component Interface (`IDbComponent`)

Every exportable object type implements this interface:

| Method/Property | Purpose |
|-----------------|---------|
| `Export()` | Export object to source file(s) |
| `Import(strFile)` | Import object from source file |
| `Merge(strFile)` | Update or replace existing object |
| `GetAllFromDB()` | Return dictionary of all objects of this type |
| `IsModified()` | Check if object changed since last export |
| `SourceFile` | Path to primary source file |
| `BaseFolder` | Export folder for this component type |
| `Category` | Display name (e.g., "Forms", "Queries") |
| `ComponentType` | Enum value from `eDatabaseComponentType` |

### Component Classes (`clsDb*`)

Each database object type has a dedicated class implementing `IDbComponent`:

| Class | Object Type |
|-------|-------------|
| `clsDbForm` | Forms |
| `clsDbReport` | Reports |
| `clsDbQuery` | Queries |
| `clsDbModule` | VBA Modules (standard and class) |
| `clsDbTableDef` | Table definitions |
| `clsDbTableData` | Table data export |
| `clsDbTableDataMacro` | Table data macros |
| `clsDbRelation` | Table relationships |
| `clsDbProperty` | Database properties |
| `clsDbVbeReference` | VBA library references |
| `clsDbTheme` | Office themes |
| `clsDbSharedImage` | Embedded images |
| `clsDbCommandBar` | Menus and toolbars |
| ... | (and more) |

### Core Modules

| Module | Purpose |
|--------|---------|
| `modImportExport` | Main export/import/build logic |
| `modObjects` | Global singleton accessors (`Options`, `Log`, `VCSIndex`, etc.) |
| `modConstants` | Shared constants and enums |
| `modDatabase` | Database utility functions |
| `modFileAccess` | File I/O operations |
| `modEncoding` | UTF-8/BOM encoding handling |
| `modErrorHandling` | Error trapping and logging |
| `modLoadFromText` | Access `LoadFromText`/`SaveAsText` wrappers |
| `modHash` | Hashing functions for change detection |

---

## Coding Standards

### File Headers

All modules and classes should include a standard header block:

```vba
'---------------------------------------------------------------------------------------
' Module    : ModuleName
' Author    : Author Name
' Date      : MM/DD/YYYY
' Purpose   : Brief description of the module's purpose
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
```

### Procedure Headers

Document all public procedures and significant private procedures:

```vba
'---------------------------------------------------------------------------------------
' Procedure : ProcedureName
' Author    : Author Name
' Date      : MM/DD/YYYY
' Purpose   : What this procedure does
'---------------------------------------------------------------------------------------
'
Public Sub ProcedureName()
```

### Naming Conventions

| Element                  | Convention                     | Example                                 |
| ------------------------ | ------------------------------ | --------------------------------------- |
| Standard modules         | `mod` prefix                   | `modImportExport`                       |
| Class modules            | `cls` prefix                   | `clsDbForm`                             |
| Interface classes        | `I` prefix                     | `IDbComponent`, `IDbSchema`             |
| Forms                    | `frm` prefix                   | `frmVCSMain`                            |
| Private module variables | `m_` prefix                    | `m_Items`, `m_FileList`                 |
| UDT instance variables   | `this`                         | `Private this As udtThis`               |
| Constants                | `UPPER_CASE` or `PascalCase`   | `CHUNK_SIZE`, `ModuleName`              |
| Enums                    | `e` prefix, PascalCase members | `eErrorLevel`, `eDatabaseComponentType` |
| Boolean parameters       | `bln` prefix                   | `blnModifiedOnly`                       |
| String parameters        | `str` prefix                   | `strFile`                               |
| Long/Integer             | `lng`/`int` prefix             | `lngCount`, `intCnt`                    |
| Dictionary               | `d` prefix                     | `dCategories`, `dFiles`                 |
| Collection               | `col` prefix                   | `colCategories`                         |
| Class objects            | `c` prefix                     | `cCategory`, `cDbObject`                |

### Error Handling

The add-in uses a structured error handling approach:

```vba
Public Sub SomeOperation()
    ' Use inline error handling with debug mode check
    If DebugMode(True) Then On Error GoTo 0 Else On Error Resume Next
    
    ' ... operation code ...
    
    ' Catch and log errors inline
    CatchAny eelError, "Error description", ModuleName & ".SomeOperation", True, True
    
    ' For critical errors that should stop the operation
    If Operation.ErrorLevel = eelCritical Then GoTo CleanUp
    
CleanUp:
    ' Cleanup code
End Sub
```

**Key functions:**
- `DebugMode(True)` - Returns true if debug mode enabled; also logs any unhandled errors
- `LogUnhandledErrors` - Call before any `On Error` to catch silent errors
- `CatchAny()` - Log error if one exists, optionally clear it
- `Catch()` - Check for specific error numbers

### Option Statements

Always include at the top of modules:

```vba
Option Compare Database  ' Use database collation for string comparison
Option Explicit          ' Require variable declaration
Option Private Module    ' For internal modules (not exposed via add-in API)
```

### Translation Support

All user-facing strings should use the `T()` function for translation:

```vba
' Simple text
Log.Add T("Beginning Export of Source Files")

' With variable substitution (use {0}, {1}, etc.)
Log.Add T("Error in file: {0}", var0:=strFileName)
MsgBox2 T("VCS Version {0}", var0:=GetVCSVersion), ...
```

---

## Key Enums

### `eDatabaseComponentType` (modConstants)

Defines all exportable object types. Maps to Access object types where applicable:

```vba
edbForm = acForm          ' Forms
edbModule = acModule      ' VBA modules
edbQuery = acQuery        ' Queries
edbReport = acReport      ' Reports
edbTableDef = acTable     ' Table definitions
edbTableData              ' Table data (custom)
edbVbeReference           ' VBA references (custom)
' ... etc.
```

### `eErrorLevel` (modConstants)

```vba
eelNoError   ' No error
eelWarning   ' Logged to file
eelError     ' Displayed and logged
eelCritical  ' Cancels current operation
```

### `eOperationType` (modConstants)

```vba
eotExport = 1  ' Exporting source files
eotBuild = 2   ' Full build from source
eotMerge = 3   ' Merge build
eotOther = 4   ' Other operations
```

---

## Working with the Codebase

### Adding a New Component Type

1. Create a new class `clsDbNewType.cls` implementing `IDbComponent`
2. Add a new enum value to `eDatabaseComponentType` in `modConstants`
3. Add the class to `GetContainers()` function in `modVCSUtility`
4. Implement all interface methods (Export, Import, Merge, GetAllFromDB, etc.)

### Modifying Export/Import Behavior

- Export logic: `modImportExport.ExportSource()`
- Import logic: `modImportExport.Build()`
- Single object operations: Individual `clsDb*.Export()` and `clsDb*.Import()` methods

### Adding Options

1. Add public property to `clsOptions`
2. Add default value in `clsOptions.LoadDefaults()`
3. Update `GetOptionsDictionary()` and loading code
4. Update `frmVCSOptions` form if user-configurable

### Testing Changes

1. Use `Testing.accdb.src` to test import/export
2. Run `Deploy` in the immediate window to increment version and export
3. Test with a variety of database objects

---

## COM Ribbon Add-in (Ribbon/)

The ribbon toolbar is implemented as a COM add-in using **twinBASIC**, enabling 64-bit compatibility.

| File | Purpose |
|------|---------|
| `MSAccessVCS_Ribbon.twinproj` | twinBASIC project file |
| `AddInRibbon.twin` | Main class implementing `IDTExtensibility2` and `IRibbonExtensibility` |
| `Ribbon.xml` | Ribbon UI definition |
| `Build/*.dll` | Compiled 32-bit and 64-bit DLLs |

The ribbon add-in acts as a thin wrapper that:
1. Loads ribbon UI from `Ribbon.xml`
2. Relays button clicks to `Version Control.accda` via `Application.Run`
3. Loads localized strings from `Ribbon.json`

---

## Export-on-Save Hook (Hook/)

Optional DLLs that hook into Access to automatically export objects when saved:
- Source: https://github.com/bclothier/AccessAppHook
- Licensed under LGPL-2.1

---

## Exported Database Source Files

The `Version Control.accda.src/` folder contains the add-in's own exported source. For detailed information about working with exported Access database source files, see:

**[Version Control.accda.src/AGENTS.md](Version%20Control.accda.src/AGENTS.md)**

This companion file explains:
- Source file formats (`.bas`, `.cls`, `.sql`, etc.)
- UTF-8 BOM encoding requirements
- VBA file structure and attributes
- Safe editing guidelines
- Import/export workflow

---

## Development Workflow

### Building from Source

1. Install a recent version of the add-in
2. Clone the repository
3. In Access, use the add-in to **Build From Source** selecting the `Version Control.accda.src` folder
4. The newly built `Version Control.accda` is ready to use

### Making Changes

1. Make modifications in the running `Version Control.accda`
2. Test thoroughly
3. Run `Deploy` in the VBA Immediate Window to:
   - Increment version number
   - Export to source files
   - Install the development version
4. Commit changes to source files (not auto-generated files)
5. Create pull request targeting `dev` branch

### Deployment

Releases are created from the `master` branch. The add-in is self-installing - users run the `.accda` file to install or update.

---

## Key Files Reference

| File | Purpose |
|------|---------|
| `vcs-options.json` | Per-project configuration (export folder, options, etc.) |
| `vcs-index.json` | Change tracking index (do not edit manually) |
| `project.json` | Database file format version |
| `vbe-references.json` | VBA library references |
| `dbs-properties.json` | Database properties |

---

## Resources

- **GitHub Repository**: https://github.com/joyfullservice/msaccess-vcs-addin
- **Wiki Documentation**: https://github.com/joyfullservice/msaccess-vcs-addin/wiki
- **Issue Tracker**: https://github.com/joyfullservice/msaccess-vcs-addin/issues
- **Releases**: https://github.com/joyfullservice/msaccess-vcs-addin/releases

---

*This file helps AI agents understand and work with the MSAccess VCS Add-in codebase. For exported database source file formats, see the AGENTS.md in the source folder.*
