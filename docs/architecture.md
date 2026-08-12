# Architecture reference

How the MSAccess VCS Add-in is put together: what lives where, the interfaces
that hold it together, and the conventions every module follows. Read this when
orienting in the codebase, adding a component type, or trying to find the class
responsible for a given object type.

For the day-to-day workflow and the invariants that matter on every change, see
[AGENTS.md](../AGENTS.md).

---

## Repository structure

| Folder | Purpose |
|--------|---------|
| `Version Control.accda.src/` | **Core add-in source code** — exported VBA modules, classes, forms, and queries |
| `Ribbon/` | **COM add-in for ribbon UI** — twinBASIC project providing 64-bit ribbon toolbar support |
| `Hook/` | **Export-on-save hook DLLs** — external library for automatic export when saving objects |
| `TestRunner/` | **Web test runner HTML** — packaging assets embedded at build (e.g. `runner.html`); not part of the Access export tree |
| `Testing/` | **Test database and fixtures** — sample database (`Testing.accdb.src`) plus the round-trip fixture corpus (`Fixtures/`) |
| `Translation/` | **Localization files** — `.pot` and `.po` files for UI translation support |
| `Wiki/` | **Documentation** — Markdown files synced to the GitHub Wiki |
| `Template/` | **Database template** — binary template used when creating new databases |
| `img/` | **README images** — screenshots and demos for documentation |

---

## Component diagram

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
                          │   - forms/*.form,*.cls │
                          │   - modules/*.bas,*.cls│
                          │   - queries/*.sql,*.json│
                          │   - vcs-options.json   │
                          │   - vcs-index.idx      │
                          └────────────────────────┘
```

## Key architectural patterns

1. **Interface-based component system.** All database object types implement
   `IDbComponent`, providing a consistent API for export, import, merge, and
   metadata operations.
2. **Singleton pattern for global state.** Key objects (`Options`, `VCSIndex`,
   `Log`, `Perf`, `Operation`) are accessed via `modObjects` module-level
   functions.
3. **Two types of build.** Full builds create a new database from source; merge
   builds update existing databases with changed files only.
4. **Index-based change detection.** `vcs-index.idx` (binary format) tracks file
   hashes and timestamps to detect changes and enable "fast save" exports.

---

## Public API (`clsVersionControl`)

The primary entry point for external automation. Exposed via the `VCS` object in
`modAPI`.

```vba
' Key public methods:
VCS.Export              ' Export all source (fast save)
VCS.FullExport          ' Export all source (full)
VCS.ExportVBA           ' Export VBA components only
VCS.ExportByType types  ' Export one or more categories (category-scoped sync)
VCS.Build strFolder     ' Full build from source
VCS.MergeBuild          ' Merge changes into existing database
VCS.ImportByType types  ' Import one or more categories (category-scoped sync)
VCS.Options             ' Access project options
```

### Category-scoped sync (`ExportByType` / `ImportByType`)

A middle tier between the single-object API (`ExportObject` / `ImportObject`)
and full Export/Build:

| Tier | Entry point | Scope | Deletions | Backup | Conflicts |
|---|---|---|---|---|---|
| Surgical | `ExportObject` / `ImportObject` | one named object | no | no | mode-driven |
| **Category sync** | **`ExportByType` / `ImportByType`** | **whole category(ies)** | **yes, within category** | **no** | **mode-driven** |
| Comprehensive | `Export` / `Build` / `MergeBuild` | entire project | yes | yes (merge) | mode-driven |

```vba
' varTypes accepts a single type or Array of types (enum or string alias):
?VCS.ExportByType("menus")                        ' Changed command bars only
?VCS.ExportByType("menus", True)                  ' All command bars (full re-export)
?VCS.ExportByType(Array(edbCommandBar, edbQuery)) ' Multiple categories
?VCS.ImportByType("command_bars")                 ' Merge changed command bars from source
?VCS.ImportByType("menus", True)                  ' Merge all command bars from source
```

Accepted aliases include `menu`/`menus` and `command_bar`/`command_bars`
(on-disk folder is `menus/`). Duplicate types in an array are collapsed. Types
not supported in the current database format (for example `connection` on an ADP
project) are rejected before the operation starts.

**Import restrictions.** `ImportByType` rejects categories whose merge path is
unsupported, including `table_data` and ADP schema objects. Use `Build` /
`MergeBuild` for those. Table data *is* merged by a normal merge build —
reconciled row by row against the primary key, gated on
`Options.MergeTableData` (see `Wiki/Merge-Build.md` and the 2026-07-28
DECISIONS.md entry). It is only category-scoped sync that rejects it, because
scoped sync takes no database backup.

**Export restrictions.** When global export options have changed (export format
version, Access version), run a normal full `Export` or `FullExport` to migrate
the project before relying on fast category export. Category-scoped export
updates index metadata only for the categories processed; it does not update the
project-wide full-export timestamp or untouched option hashes.

These methods treat each named category **as a whole**: they reconcile deletions
within the category (export removes orphaned source files; import removes
orphaned DB objects). Conflicts are auto-resolved under MCP/API and prompted
interactively. No database backup is taken. Open UI objects of targeted
categories are closed before export/import (save behavior follows interaction
mode); module code is flushed via VBA project save rather than closing module
windows.

Cross-project access uses
`Application.Run(Environ$("AppData") & "\MSAccessVCS\Version Control.API", "ExportByType", "menus")`
— pass one type per call when using `Application.Run` (VBA arrays do not marshal
reliably across projects). Qualify the call with the **full path** to the add-in,
which also loads it on demand. The bare file name (`"Version Control.API"`) does
not resolve, because `Application.Run` matches on the VBA project name rather
than the file name; `"MSAccessVCS.API"` does work, but only once the add-in is
already loaded.

---

## Component interface (`IDbComponent`)

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

### Component classes (`clsDb*`)

Each database object type has a dedicated class implementing `IDbComponent`.
The authoritative list is the contents of
`Version Control.accda.src/modules/Database/` plus the registrations in
`GetContainers()` in `modVCSUtility`; representative members include `clsDbForm`,
`clsDbReport`, `clsDbQuery`, `clsDbModule`, `clsDbTableDef`, `clsDbTableData`,
`clsDbTableDataMacro`, `clsDbRelation`, `clsDbProperty`, `clsDbVbeReference`,
`clsDbTheme`, `clsDbSharedImage`, and `clsDbCommandBar`.

### Core modules

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

## Key enums (`modConstants`)

### `eDatabaseComponentType`

Defines all exportable object types. Maps to Access object types where
applicable:

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

### `eErrorLevel`

```vba
eelNoError   ' No error
eelWarning   ' Logged to file
eelError     ' Displayed and logged
eelCritical  ' Cancels current operation
```

### `eOperationType`

```vba
eotExport = 1  ' Exporting source files
eotBuild = 2   ' Full build from source
eotMerge = 3   ' Merge build
eotTestRun = 4 ' VCS.RunTests / clsTestRunner suite
eotOther = 9   ' Other / catch-all operations
```

---

## Modifying the query parser

The query parser (`clsQueryComposer.cls` + `clsDbQuery.cls`) carries hard-won
decisions in places that are not always obvious from a casual read. Before
modifying either class, read these in order:

- **[access-query-storage.md](access-query-storage.md)** — how Access stores
  queries, what shapes our parser handles (with the canonical fixture for each),
  known gaps where behaviour is unverified, and findings unique to our pipeline
  (`Application.LoadFromText` / `Application.SaveAsText` asymmetries).
- **[DECISIONS.md](../DECISIONS.md)** — search for entries mentioning
  `clsQueryComposer` or `clsDbQuery` (e.g. `rg "clsQueryComposer" DECISIONS.md -A 30`).
  Captures the rationale and rejected alternatives behind each choice.
- **`Testing/Fixtures/queries/regression/*.notes.md`** — each one pins a specific
  SQL shape and explains what would re-break if a careful decision were reverted.
- **Procedure-header comments** on the functions you're modifying —
  `RequiresDesignView`, `IsDesignerCompatible`, `HasTopLevelBoolean`,
  `ParseJoinExpression`, `SafeBreak`, and `EmitDbMemoSql` carry constraints in
  their headers that the body alone does not convey.

Do not look in `Testing.accdb.src` for query regression fixtures; the round-trip
corpus is `Testing/Fixtures/queries/`. When you discover a new invariant or edge
case worth preserving, follow the four-layer documentation pattern at
[Testing/Fixtures/README.md](../Testing/Fixtures/README.md).

---

## Header and option conventions

Every module and class opens with a header block:

```vba
'---------------------------------------------------------------------------------------
' Module    : ModuleName
' Author    : Author Name
' Date      : MM/DD/YYYY
' Purpose   : Brief description of the module's purpose
'---------------------------------------------------------------------------------------
Option Compare Database  ' Use database collation for string comparison
Option Explicit          ' Require variable declaration
Option Private Module    ' For internal modules (not exposed via add-in API)
```

Public procedures and significant private procedures carry the same treatment:

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

---

## COM ribbon add-in (`Ribbon/`)

The ribbon toolbar is implemented as a COM add-in using **twinBASIC**, enabling
64-bit compatibility.

| File | Purpose |
|------|---------|
| `MSAccessVCS_Ribbon.twinproj` | twinBASIC project file |
| `AddInRibbon.twin` | Main class implementing `IDTExtensibility2` and `IRibbonExtensibility` |
| `Ribbon.xml` | Ribbon UI definition |
| `Build/*.dll` | Compiled 32-bit and 64-bit DLLs |

The ribbon add-in is a thin wrapper: it loads the UI from `Ribbon.xml`, relays
button clicks to `Version Control.accda` via `Application.Run`, and loads
localized strings from `Ribbon.json`.

## Helper script (`clsWorker`)

A handful of jobs cannot run in the process that is running the add-in.
`clsWorker` extracts the VBScript below its `' *** BEGIN WORKER SCRIPT ***` marker
into `Worker.vbs` in the install folder and launches it with `wscript`. Anything
added to the class **above** that marker stays VBA and is not part of the script.
Results come back through `modAPI.WorkerCallback` → `Worker.ReturnWorker`.

| Consumer | Why out of process |
|---|---|
| `Run_SaveVbaProject` (via `modVbeUtility.SaveCurrentVBProject`) | The VBE Save command saves nothing while the caller's own VBA is on the stack. No in-process substitute works — that procedure's header lists four that were tried. |
| `IsDatabaseAccessible` | The engine does not report its lock state to same-process callers. |
| `Run_UninstallAddin` | Deletes the add-in file, which Access holds open until it exits. |
| `Run_BuildAndInstall` (`VCS.RebuildAddIn`) | The add-in cannot rebuild and reinstall itself while loaded. |

Endpoint protection in some managed environments blocks Access from launching a
freshly written script, so `modInstall.UseWorkerScript` (per-user registry,
default on, set on `frmVCSInstall`) turns the whole mechanism off. `CallWorker`
checks it and no-ops, which makes a call site that forgets to branch degrade
rather than launch `wscript` — and is the single seam where a different
out-of-process backend would attach. Each consumer has a script-free fallback;
the one with correctness stakes is the VBA project save, which warns the user to
save in the VBE rather than letting an export silently omit unsaved class-module
edits. See the 2026-08-07 `DECISIONS.md` entry.

## Export-on-save hook (`Hook/`)

Optional DLLs that hook into Access to automatically export objects when saved.
Source: https://github.com/bclothier/AccessAppHook, licensed under LGPL-2.1.
