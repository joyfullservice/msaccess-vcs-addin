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
| `TestRunner/` | **Web test runner HTML** - Packaging assets embedded at build (e.g. `runner.html`); not part of the Access export tree |
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
                          │   - forms/*.form,*.cls │
                          │   - modules/*.bas,*.cls│
                          │   - queries/*.sql,*.json│
                          │   - vcs-options.json   │
                          │   - vcs-index.idx      │
                          └────────────────────────┘
```

### Key Architectural Patterns

1. **Interface-Based Component System**: All database object types implement `IDbComponent`, providing a consistent API for export, import, merge, and metadata operations.

2. **Singleton Pattern for Global State**: Key objects (`Options`, `VCSIndex`, `Log`, `Perf`, `Operation`) are accessed via `modObjects` module-level functions.

3. **Two Types of Build**: Full builds create a new database from source; merge builds update existing databases with changed files only.

4. **Index-Based Change Detection**: `vcs-index.idx` (binary format) tracks file hashes and timestamps to detect changes and enable "fast save" exports.

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

### Understanding `LogUnhandledErrors` in Log Files

VBA's `On Error` statements silently clear the current `Err` object. To avoid losing error information, `LogUnhandledErrors` is called *just before* an `On Error` directive to capture any leftover error before it gets wiped. `DebugMode(True)` calls `LogUnhandledErrors` internally, so the same behavior applies at the top of functions that use the `DebugMode` pattern.

**The error did NOT originate where it was logged.** When you see a log entry like:

```
ERROR: Unhandled error, likely before `On Error` directive
```

This entry means the exact origin is not known — `LogUnhandledErrors` detected a leftover error but has no information about which function raised it. The error came from whatever code ran immediately *before* the `LogUnhandledErrors` call. To find the real source, look at the surrounding log context (the operation in progress, the objects being processed) and find the `LogUnhandledErrors` call site in the source code, then look at what executed before it.

Some call sites pass a `CallingFunction` parameter, which narrows the search to a specific function (e.g., `Source: modBuild.Build.Unknown.LogUnhandledErrors`). Even then, the error did not originate in that function — it came from code that ran before the call. For example:

```vba
Public Sub Build()
    ' ... earlier code calls helper functions ...
    SomeHelperFunction   ' <-- If this raises an error internally and doesn't handle it,
                         '     the error persists in the Err object after it returns.

    LogUnhandledErrors   ' <-- Catches the leftover error from SomeHelperFunction
    On Error Resume Next ' <-- Would have silently cleared it without the line above
    ' ... more code ...
End Sub
```

In this example, the actual source of the error is `SomeHelperFunction`, not `Build`.

### Option Statements

Always include at the top of modules:

```vba
Option Compare Database  ' Use database collation for string comparison
Option Explicit          ' Require variable declaration
Option Private Module    ' For internal modules (not exposed via add-in API)
```

### Library Constants

All modules in the VBA project share the same library references (DAO, VBE, Scripting, etc.). Use the library-defined constants (e.g., `dbQSQLPassThrough`, `acQuery`, `vbTextCompare`) rather than hard-coding their numeric values. Magic numbers obscure intent and bypass compile-time checking.

### Translation Support

All user-facing strings should use the `T()` function for translation:

```vba
' Simple text
Log.Add T("Beginning Export of Source Files")

' With variable substitution (use {0}, {1}, etc.)
Log.Add T("Error in file: {0}", var0:=strFileName)
MsgBox2 T("VCS Version {0}", var0:=GetVCSVersion), ...
```

### File System Operations

**Never use the VBA `Dir()` function.** `Dir()` does not support Unicode filenames and will silently skip or fail on paths containing non-ASCII characters. Access databases frequently contain objects with Unicode names (accented characters, CJK, etc.).

Instead, use:
- **`Scripting.FileSystemObject`** (FSO) for general file operations — supports Unicode natively
- **Win32 API** (`FindFirstFileW`/`FindNextFileW` in `modFileWinAPI`) for performance-critical scans or existence checks
- **`FilePatternExists()`** in `modFileWinAPI` for quick wildcard existence checks (O(1) when no match)
- **`ScanFolderContents()`** in `modFileWinAPI` for enumerating files and subfolders in a single pass

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
eotTestRun = 4 ' VCS.RunTests / clsTestRunner suite
eotOther = 9   ' Other / catch-all operations
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

**Export Format Versioning:** Any change that alters the content or structure of exported source files (sanitization rules, property stripping, file layout, JSON structure, etc.) **must** be gated behind an export format version check. This allows users to upgrade the add-in without being forced to adopt new export formatting until they choose to.

How to gate a new export behavior change:

1. Add a new member to the `eExportFormatVersion` enum in `modConstants.bas` (e.g., `EFV_5_1_0 = 50100`) and update `[_Last]`
2. Wrap the new behavior: `If Options.ExportFormatVersion >= EFV_5_1_0 Then`

`LATEST_EXPORT_FORMAT` is derived automatically from `eExportFormatVersion.[_Last]`.

Import logic does **not** need gating — it must remain backwards compatible with all prior export formats.

### Modifying the Query Parser

The query parser (`clsQueryComposer.cls` + `clsDbQuery.cls`) carries hard-won decisions in places that are not always obvious from a casual read. Before modifying either class, read these in order:

Do not look in `Testing.accdb.src` for query regression fixtures; the shipped round-trip corpus is `Testing/Fixtures/queries/`.

- **[docs/access-query-storage.md](docs/access-query-storage.md)** — in-repo reference for how Access stores queries, what shapes our parser handles (with the canonical fixture for each), known gaps where behaviour is unverified, and findings unique to our pipeline (`Application.LoadFromText` / `Application.SaveAsText` asymmetries).
- **[DECISIONS.md](DECISIONS.md)** — search for entries mentioning `clsQueryComposer` or `clsDbQuery` (e.g. `rg "clsQueryComposer" DECISIONS.md -A 30`). Captures the rationale and rejected alternatives behind each choice.
- **`Testing/Fixtures/queries/regression/*.notes.md`** — each one pins a specific SQL shape and explains what would re-break if a careful decision were reverted.
- **Procedure-header comments** on the functions you're modifying — `RequiresDesignView`, `IsDesignerCompatible`, `HasTopLevelBoolean`, `ParseJoinExpression`, `SafeBreak`, `EmitDbMemoSql` carry constraints in their headers that the body alone does not convey.

When you discover a new invariant or edge case worth preserving, follow the four-layer documentation pattern at [Testing/Fixtures/README.md § Documenting parser invariants and edge cases](Testing/Fixtures/README.md).

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

## Debugging RunVBA Failures

`clsVersionControl.RunVBA` (exposed to agents via the `vcs_run_vba` MCP tool) wraps caller-supplied VBA code in a temporary module, compiles it, runs it, and returns a JSON result. Two debugging affordances make iteration on agent-authored snippets much faster:

### 1. Auto-injected line numbers and `errorLine`

Before the wrapper is built, `RunVBA` runs the submitted code through the private helper `AddVbaLineNumbers` (in the same class). That helper prepends sequential 1-based VBA line numbers to every executable statement. The number value equals the 1-based ordinal of the line within the original `code` string (the counter advances on every physical input line, blanks/comments/continuations included), so when a runtime error fires, the JSON result contains an `errorLine` field that maps directly back to the caller's source.

```json
{
  "success": false,
  "error": "Type mismatch",
  "errorNumber": 13,
  "errorLine": 7
}
```

`errorLine: 7` literally means "line 7 of what I submitted" — no offset math required. The field is omitted when no `Erl` value is available (e.g., the wrapper itself failed to compile, or the error fired before any numbered line ran).

Lines that cannot legally carry a VBA line number — blank lines, pure comments, continuations of a prior `_`-terminated line, and lines the caller already pre-numbered — are passed through unchanged. The counter still advances over them so the line numbers remain in sync with the original text.

### 2. Concise multi-error test procedures

The default wrapper uses `On Error Resume Next` and reports the **last** runtime error. When you want a test to keep going past the first failure and report every problem in a single round-trip, write your own handler that exploits the auto-injected line numbers:

```vba
Dim col As New Collection
On Error GoTo H
CurrentDb.Execute "DELETE * FROM tblA"
CurrentDb.Execute "INSERT INTO tblB SELECT * FROM nope"
CurrentDb.Execute "UPDATE tblC SET x = 1"
MCP_TempFunction = "errors=" & col.Count & " | " & Join(CollectionToArray(col), "; ")
Exit Function
H: col.Add Erl & ": " & Err.Number & " " & Err.Description
Resume Next
```

Each `Erl` value collected inside the `H:` label is meaningful because the wrapper auto-numbered every line for you. You don't need to write `10`, `20`, `30` yourself — the line numbers are already there by the time your code runs.

The decision between the default single-error capture and an explicit multi-error handler is per-test: pick whichever shape best matches what the test is trying to verify.

### VBA error-handler state: `Err.Clear` is not enough

When execution is inside an active `On Error GoTo Handler` block, `Err.Clear`
clears the error object but does **not** reset the active exception/handler
state. Expected cleanup errors inside that handler can still break or poison
the wrapper if you only write `On Error Resume Next`.

Use this pattern before any expected-error cleanup inside a handler:

```vba
Handler:
    strMsg = Err.Description
    Err.Clear
    On Error GoTo -1      ' clear active handler state
    On Error Resume Next  ' now expected cleanup errors are safe
    CurrentDb.QueryDefs.Delete "__temp_query__"
    Err.Clear
    On Error GoTo 0
    GoTo ContinueAfterHandler
```

Do not use `Resume` after `On Error GoTo -1`; jump to a continuation label
instead. Prefer explicit existence checks over expected-error cleanup when the
code is simple enough.

---

## Testing Strategy

The add-in benefits from three distinct testing layers, each catching a different bug class. Keeping them as separate layers (rather than one giant test database) is deliberate — each layer trades scope for fidelity.

| Layer | What it tests | Lives in |
|---|---|---|
| **1. VBA logic tests** | "Given inputs, does this function return the right output?" | `modTestSuite.bas` (formerly `modUnitTesting.bas`) |
| **2. Object round-trip tests** | "Does this database object survive a serialize/deserialize cycle unchanged?" | `modTestRoundtrip.bas` + `Testing/Fixtures/` |
| **3. Whole-database integration** | "Does building an entire database from source produce a working database?" | `Testing/Testing.accdb.src` |

**Log files are gitignored.** All `logs/` directories and `*.log` files are excluded by `.gitignore`. Agent tools that respect `.gitignore` (Glob, Grep, semantic search) will silently skip them. Use shell commands to find and read log files:

```powershell
# Find log files (run from repository root)
Get-ChildItem -Recurse -Include "*.log","*.json" | Where-Object { $_.DirectoryName -like "*logs*" }
```

Key log locations:
- `Version Control.accda.src/logs/` — build, export, merge, and **ephemeral test run** logs (`TestResults_*.json`, `TestRun_*.log`)
- `Version Control.accda.src/test-results/` — **durable test state** (`test-state.json`), **JUnit XML** (`test-results.xml`), and **HTML report** (`test-results.html`); gitignored
- `Testing/Fixtures/logs/` — object round-trip test logs (`ObjectRoundtrip_*.log`)

Important location distinction: the canonical object round-trip fixture corpus
is `Testing/Fixtures/` as plain text files. Query fixtures live under
`Testing/Fixtures/queries/` as `.sql` + `.json` pairs, with optional
`.notes.md` files for regression context. `Testing/Testing.accdb.src` is the
sample/integration database used for full build/export flows; do not look there
for the primary `VCS.RunRoundtripTests` fixture corpus. `MSysQueriesExamples`
and `db-analysis-tools` are useful sources or validation projects for query
shapes, but they are not the add-in's canonical regression fixture store.

### `modTest*` naming convention

All test-infrastructure modules use the `modTest*` family prefix. This matches the existing family-grouping conventions already used in the codebase (`clsDb*` for component classes, `clsLv*` for ListView property parsers) and gives "Test" maximum prominence for developer and agent discoverability.

- `modTestSuite` — heterogeneous unit tests (encoding, JSON, sanitization, formatter, hashing, IDbComponent invariants).
- `modTestRoundtrip` — generalized object round-trip regression harness.
- Future siblings (e.g., `modTestPerf`, `modTestFixtures`, `modTestEncoding`) should adopt the same prefix automatically.

### Object round-trip harness (`modTestRoundtrip`)

The Layer 2 harness is generic over `IDbComponent`. v1 ships query support; forms, reports, modules, and table data follow the same pattern by adding a per-type helper.

When adding regression fixtures from a user or production database, sanitize
the fixture and its `.notes.md` file. Do not include source database names,
source query names, table/field names, business concepts, file paths, customer
names, or server names. Use generic parser-shape language such as "production
validation exposed a cross-subtree join predicate placement bug."

For each fixture under `Testing/Fixtures/`, the harness:

1. Imports the fixture into the running database under a sandboxed name (`vcs_test_<basename>_<hash>`).
2. Validates the emitter's `.qdef` output:
   - `qdef_joins` — structural check: each join row's `LeftTable`/`RightTable` matches its `Expression` (Design View only).
   - `qdef_vs_fixture` — drift check: compares generated `.qdef` against stored `.qdef` baseline (if present).
3. Exports it twice (Pass 1 and Pass 2), into a per-run scratch folder.
4. Asserts Pass 2 == Pass 1 (idempotency, hard requirement).
5. Asserts Pass 1 == fixture (drift check, soft requirement).
6. Drops the sandboxed object and moves on.

Output goes to three coordinated channels:

- **`frmVCSMain` console** — live progress (one line per fixture).
- **Per-session log file** — `Testing/Fixtures/logs/ObjectRoundtrip_<opId>.log`, with full unified diffs for any failures.
- **JSON return value** — machine-parseable summary suitable for `vcs_run_vba` callers and CI.

All external invocations go through the public API method `VCS.RunRoundtripTests`. The implementation in `modTestRoundtrip.bas` uses `Option Private Module` so test internals stay hidden from cross-project `Application.Run` lookups, matching the rest of the add-in.

Run it from the VBA Immediate Window:

```vba
?VCS.RunRoundtripTests
```

Or via MCP (requires `McpAllowRunVBA` to be enabled, same as any other agent-driven code execution):

```
vcs_run_vba(<addin-path>, "MCP_TempFunction = VCS.RunRoundtripTests()")
```

End users can point the same harness at their own fixture corpus:

```vba
?VCS.RunRoundtripTests("C:\path\to\my-fixtures\")
```

Pass `True` as the second argument to rebaseline mismatched fixtures (review the resulting git diff before committing). When working inside the add-in's own VBE — e.g., debugging the harness itself — the in-project entry point `?modTestRoundtrip.RunObjectRoundtripTests()` is also available, since `Option Private Module` only blocks cross-project callers, not in-project ones.

### Bug-as-fixture: the contribution workflow

The harness was designed to support a specific contribution workflow that is uniquely enabled by the add-in's text-source format. When a user hits an edge case where an object fails to round-trip:

1. They reproduce the bug in their own database.
2. They sanitize the failing object's `.sql` + `.json` pair (strip business-sensitive names, replace `Connect` strings with `env:` references, remove embedded data).
3. They drop the pair into `Testing/Fixtures/queries/regression/` (or the appropriate category) on a branch.
4. They optionally add a sibling `<name>.notes.md` describing what the bug was and linking to the issue.
5. They open a PR against the add-in.

The fixture becomes a permanent regression test against every future change. The user's bug report and the regression test are literally the same artifact.

See `Testing/Fixtures/README.md` for the full workflow, the `_scaffold/` convention (for fixtures with shared dependencies), and a sanitization checklist.

### Adding new test modules

When adding a new test module, follow the `modTest*` prefix convention and place it under `Version Control.accda.src/modules/Tests/`. If the module wraps an existing concept (encoding, hashing, sanitization), prefer extracting it from `modTestSuite` into a focused `modTest<Topic>` module rather than letting `modTestSuite` grow indefinitely.

### Web test runner UI

On Microsoft 365 builds with the Edge browser control (file build ≥ 16327) and
with the `UseWebTestRunner` option enabled (default **on**), `VCS.RunTests` /
ribbon **Run Tests** open `frmVCSTestRunner`, merge-scan for tests, and publish the
test tree to `TestRunner/runner.html` via `modTestRunnerUI` (`ExecuteJavascript` →
`window.TestUI.*`). Tests are **not** auto-run; the user clicks the primary **Run**
button (label reflects scope: all, folder, suite, filter, or failed) or a per-suite/per-test
▶ when ready. When the option is off (or the build is older), it falls back to
the `frmVCSMain` console unchanged.

- **HTML**: [`TestRunner/runner.html`](TestRunner/runner.html) (repo-root packaging
  asset, embedded at build like `Ribbon.xml`; extracted to a temp cache at runtime)
- **Entry points**: `VCS.RunTests` (show + deferred scan; user clicks Run),
  `VCS.OpenTestRunner` (open to view last results — rehydrates from singleton or
  `test-state.json` on standalone opens)
- **Option**: `Options.UseWebTestRunner` (Advanced options → Automated Testing;
  default True) gates the whole routing in `clsVersionControl.ExecuteTests`.
- **Inbound bridge**: outbox **polling** — JS enqueues commands in
  `window.__vbaOutbox`, the form timer drains them via `RetrieveJavascriptValue`
  (no navigation; see the `frmVCSTestRunner` header and DECISIONS.md 2026-07-08).
  Allowlisted callbacks: `RunAll,RunSelected,RunFailed,Cancel,OpenTestSource,RefreshTests,OpenResultsReport`.
- **Form lifecycle**: opens as a **pop-up** window (`PopUp=1`). Closing via the
  X button or Escape **hides** the form (timer disabled, WebView2 stays warm);
  re-opening **reuses** the hidden instance without reloading the page when healthy
  (see `open.reuse.warm` in the diag log). A forced reload replays completed results.
  Programmatic `CloseWebTestRunner` sets `AllowClose` and issues `DoCmd.Close` for
  a real unload.
- **Tree refresh**: after show (or on **Refresh** in the toolbar), VBA runs
  `ScanMergingPriorResults` — rediscovers tests while preserving pass/fail for
  unchanged `Module.Proc` keys — and republishes the tree only.
- **Quiet mode**: while the web runner hosts a run, `Log.SuppressDebugOutput` is
  set so per-test results are not echoed to the Immediate window (the UI shows
  them; the log file is unaffected).
- **State persistence**: durable merged state in `<export-folder>/test-results/test-state.json`
  (survives Access restarts; partial runs update only executed tests and mark others
  `stale`). The in-memory `clsTestRunner` singleton is still used within a session;
  `VCS.OpenTestRunner` reloads from `test-state.json` when the singleton is empty.
- **JUnit export**: `Options.ExportTestResultsJUnit` (default **on**) writes
  `<export-folder>/test-results/test-results.xml` after each run as a projection of
  the durable state. `VCS.ExportTestResultsJUnit` regenerates it on demand without
  re-running tests.
- **HTML report**: `Options.ExportTestResultsHtml` (default **on**) writes
  `<export-folder>/test-results/test-results.html` after each run — a self-contained
  dashboard with inlined `test-state.json` (opens offline via double-click,
  **Open report** in the web runner toolbar, or **Open Test Results...** on
  `frmVCSMain` after a console test run). `VCS.ExportTestResultsHtml`
  regenerates it on demand. The console logs the report file path (plain text).
- **UI affordances**: sidebar has **All tests** / **Failed tests (N)** focus
  entries; nested **@Folder** tree with folder select (click name) and ▶ run;
  **Tags** section with include/exclude cycle; a single filter box using
  `VCS.RunTests` token syntax (`SQL -slow`) that scopes the test list and Run
  (sidebar tree is navigation-only, not a second filter); **Recent** stores full
  `{folder, suite, filterText}` snapshots so combinations restore on click; the
  stats bar shows PHPUnit-style **tests** and **assertions** totals; primary Run
  executes the **visible scope** (composed navigation + filter); per-test/per-suite/
  folder ▶ run buttons; clicking a location opens the VBE at the proc.
  `VCS.RunTests(...)` / ribbon `DefaultTestFilter` prefill the filter box when the
  web runner opens (no auto-run). Opening the web runner does not begin an
  `Operation` — bridge Run callbacks do.
- **Late binding**: the Edge control is referenced as `As Object` only so the
  add-in still compiles on older Access.

See `DECISIONS.md` (2026-07-07 through 2026-07-09 — Web test runner) for rationale.

#### Diagnostic trace log (debugging the bridge)

`modTestRunnerDiag` writes a single agent-readable trace of the real
bridge/lifecycle flow to **`<ExportFolder>\logs\TestRunnerDiag.log`** (the
resolved path is printed in the log header; falls back to a temp folder when
Options aren't loaded). It is truncated at the start of each session
(`DiagStart`, called from `OpenWebTestRunner`). Each line is
`[+elapsed ms] TAG | detail`.

Key tags: `form.load` / `form.unload` / `form.hide` / `form.show` (form
lifecycle), `navigate.url` / `navigate.call` (what URL the control was given),
`documentcomplete` (page finished loading — the gap from `navigate.call` is the
WebView2 load/cold-start time), `wait.ready` / `wait.timeout` (readiness-wait
outcome), `beforenavigate` (proves the JS→VBA event bridge fired, with fn+id),
`defer.exec` → `dispatch.begin`/`dispatch.end` (deferred command execution),
`resolve` / `reject`, `push` / `push.dropped` (VBA→JS streaming), and `js.*`
(JS-side breadcrumbs drained from `window.__diag`: `js.call`, `js.signal`,
`js.resolve`, `js.timeout`, `js.onReady`, `js.onTestComplete`, etc.). Read this
file first when the page doesn't load or a call times out — it shows exactly
where the flow diverged.

### Running and filtering tests

`VCS.RunTests` accepts an optional `ParamArray` of filter arguments. Each argument is resolved in priority order:

1. **Module name** — exact match on the module/class name
2. **Suite/folder** — match against `@Folder` annotation values (exact or final-segment, e.g., `"SQL"` matches `"Tests.SQL"`)
3. **Procedure name** — match on procedure name or full `Module.Procedure` key
4. **Tag** — match against `'@Tag("...")` annotations

Prefix any argument with `-` to **exclude** it. Inclusions combine with OR; exclusions combine with AND. If only exclusions are specified, the base set is all tests.

```vba
?VCS.RunTests                                    ' Run all tests
?VCS.RunTests("modTestEncoding")                 ' Run one module
?VCS.RunTests("-slow")                           ' Run all except slow-tagged tests
?VCS.RunTests("SQL", "-slow")                    ' Run SQL suite, skip slow tests
?VCS.RunTests("TestParseJoinExpression")         ' Run one specific procedure
?VCS.RunTests("-modTestConnect", "-slow")        ' Exclude a module and a tag
```

### Global suite hooks

Optional once-per-run `GlobalTestSetup` / `GlobalTestTeardown` in `modTestAssert`. See [`.cursor/rules/testing.mdc`](.cursor/rules/testing.mdc) (Global Suite Hooks) for the full contract when working in this repository.

### Tagging tests

Use `'@Tag("name")` annotations to categorize tests. Tags are case-insensitive.

**Module-level tags** (in the first ~30 lines, before any procedure) apply to all tests in the module:

```vba
Option Private Module
'@Folder("Tests")
'@Tag("slow")
'@Tag("database")
```

**Procedure-level tags** go inside the procedure body (first lines, before any executable code):

```vba
Public Sub TestExpensiveQuery()
    '@Tag("slow")
    '@Tag("regression")
    TestAssert RunExpensiveCheck(), "check passes"
End Sub
```

Tags must appear at the very top of the body (before `Dim` statements or code). The scanner stops at the first non-comment, non-blank line. Module-level and procedure-level tags are merged — a test inherits all module-level tags plus its own.

Tags are included in the JSON test results under a `"tags"` array per test entry.

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
- Source file formats (`.bas`, `.cls`, `.form`, `.report`, `.macro`, `.sql`, `.json`, etc.)
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
| `vcs-index.idx` | Change tracking index, binary format (do not edit manually). Use `VCSIndex.DumpToJson` for a human-readable dump. |
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
