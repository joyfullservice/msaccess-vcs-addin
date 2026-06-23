# AGENTS.md - AI Agent Guide to MSAccess VCS Export Structure

## Context: What Is This Folder?

This folder contains **exported source files from a Microsoft Access database**. The MSAccess VCS (Version Control System) add-in generated these files by converting binary Access database objects (forms, reports, queries, modules, etc.) into text-based source files suitable for version control with Git.

**Key points for AI agents:**
- This is ONE exported Access database, not the VCS add-in itself
- The user exported their `.accdb`/`.mdb` database using the add-in
- Changes made here must be imported back using the add-in in Microsoft Access
- The add-in is required to build a working database from these source files
- You cannot run or test this code directly - it must be imported into Access first

---

## Critical Rules for AI Agents

### 1. UTF-8 BOM Encoding

All text files in this exported database use **UTF-8 with BOM** encoding (bytes `EF BB BF` at file start). This is enforced by `.editorconfig` and is mandatory for successful import into Access.

### 2. CRLF Line Endings

All text files use **CRLF** (`\r\n`) line endings. This is enforced by `.editorconfig` and `.gitattributes`, and is mandatory for Access `LoadFromText` import.

### 3. Preserve VBA File Structure

VBA files (`.bas`, `.cls`) have required headers that must not be modified. See "VBA File Format" section below.

### 4. Do Not Edit These Files

- **`vcs-index.idx`** - Binary index managed automatically by the add-in (tracks file hashes). Use `VCSIndex.DumpToJson` to export a human-readable JSON copy for debugging.
- **`.frx` files** - Binary OLE data, not editable as text
- **`.thmx` files** - ZIP archives, edit extracted contents instead

---

## VBA File Format Details

### Standard Modules (`.bas` files in `modules/`)

```
Attribute VB_Name = "ModuleName"        <- Required: Module name (must match filename)
'---------------------------------------------------------------------------------------
' Module    : ModuleName                <- Optional: Comment header
' Purpose   : Description here
'---------------------------------------------------------------------------------------
Option Compare Database                  <- Optional (coding style varies)
Option Explicit                          <- Optional (coding style varies)

Public Sub MySub()
    ' Code starts after any Option statements (if present)
End Sub
```

**Safe to modify:** VBA code (procedures, functions, declarations)
**Do not modify:** The `Attribute VB_Name` line (must match filename without extension)

**Note:** `Option` statements are not required - some codebases omit them entirely.

**`@Folder` annotations:** Starting with export format 5.0.0, the add-in reads Rubberduck-style `'@Folder("Category")` annotations and exports modules into matching subdirectories under `modules/` (e.g., `'@Folder("Core")` exports to `modules/Core/`). Dots become path separators (`'@Folder("Components.ADP")` exports to `modules/Components/ADP/`). During import, subdirectories are scanned recursively. Modules without an `@Folder` annotation remain directly in `modules/`. The same subfolder layout applies to forms (`forms/`) and reports (`reports/`).

**Critical for agents — file location follows the annotation, not the folder you are editing in:**

- The `'@Folder("X.Y")` comment inside the file decides where it belongs on disk (`modules/X/Y/<name>.bas`, `forms/X/Y/<name>.form`, etc.).
- For **modules**, the annotation is in the `.bas` or `.cls` file itself. For **forms and reports**, the annotation lives in the `.cls` code-behind file (not in `.form`/`.report`, `.json`, or `.svg`).
- Before creating or moving a component, search the **entire** `modules/`, `forms/`, or `reports/` tree for an existing file with the same basename (`Attribute VB_Name` for modules). Edit that file in place; never create a second copy elsewhere.
- If two copies exist (e.g. `modules/modFoo.bas` and `modules/Tests/modFoo.bas`), build/import processes **both** and the last import silently wins — no error is shown. This is a common source of drift when agents write to the wrong folder.
- Build now auto-removes misplaced duplicates (modules, forms, and reports) when exactly one copy sits in its annotation-derived folder; ambiguous cases are left in place with a warning. Deleting a duplicate form/report removes the whole source group (`.form`/`.report` + `.cls` + `.json` + `.svg`) together.

### Class Modules (`.cls` files in `modules/`)

```
VERSION 1.0 CLASS                        <- Required header block
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassName"          <- Required: Class name
Attribute VB_GlobalNameSpace = False     <- Required attributes (legacy, but must exist)
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False       <- True = has default instance
Attribute VB_Exposed = False             <- True = accessible from other projects
Option Compare Database                  <- Optional
Option Explicit                          <- Optional

Private m_Value As String

Public Property Get Value() As String
Attribute Value.VB_Description = "Returns the value"   <- Member attribute (see below)
    Value = m_Value
End Property
```

**Safe to modify:** VBA code (procedures, properties, declarations), comment headers
**Do not modify:** VERSION block, module-level `Attribute VB_*` lines

### VBA Member Attributes

Attributes can also appear **inside procedures and properties** to add metadata not visible in the VBA IDE but shown in the Object Browser:

| Attribute | Purpose |
|-----------|---------|
| `Attribute [Member].VB_Description = "text"` | Adds description visible in Object Browser |
| `Attribute [Member].VB_UserMemId = 0` | Makes this the default member of the class |
| `Attribute [Member].VB_UserMemId = -4` | Makes this return an enumerator for `For Each` |

These appear on the line immediately after the `Sub`, `Function`, or `Property` declaration:

```vba
Public Function GetItem(Index As Long) As Variant
Attribute GetItem.VB_Description = "Returns item at specified index"
Attribute GetItem.VB_UserMemId = 0
    GetItem = m_Items(Index)
End Function
```

**Safe to modify:** The description text in quotes
**Do not modify:** The attribute syntax or member name reference

### Form/Report Code-Behind (`.cls` files in `forms/` or `reports/`)

When "Split Layout from VBA" is enabled, the VBA code is in a separate `.cls` file:

```
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmMyForm"     <- Must be "Form_" + form name
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database                  <- Optional

Private Sub Form_Load()
    ' Event handler code
End Sub
```

---

## File Extension Note

Starting with export format version 5.0.0, forms, reports, queries, and macros use descriptive file extensions (`.form`, `.report`, `.qdef`, `.macro`) instead of the overloaded `.bas` extension. Older projects may still use `.bas` for these object types. When working with source files, check which extensions are actually present in the folder. The add-in accepts both old and new extensions during import.

| Object Type | New Extension | Legacy Extension |
|-------------|---------------|------------------|
| Forms       | `.form`       | `.bas`           |
| Reports     | `.report`     | `.bas`           |
| Queries     | `.sql` + `.json` | `.qdef`, `.bas`  |
| Macros      | `.macro`      | `.bas`           |

VBA modules continue to use `.bas` and `.cls` (unchanged).

---

## Form and Report Object Files

### Object Definition Files (`.form` / `.report`, or legacy `.bas`, in `forms/` and `reports/`)

These files use Access's `SaveAsText` format - a custom text format that is NOT VBA code:

```
Version =21                              <- Access version marker
VersionRequired =20
Begin Form                               <- Object type (Form or Report)
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    Caption ="My Form Title"
    Begin
        Begin Label                      <- Control definition
            Left =1234
            Top =567
            Width =2000
            Height =300
            Name ="lblTitle"
            Caption ="Welcome"
        End
        Begin CommandButton
            Left =1234
            Top =1000
            Name ="cmdOK"
            Caption ="OK"
            OnClick ="[Event Procedure]"
        End
    End
End
CodeBehindForm                           <- VBA code section (if not split)
Attribute VB_Name = "Form_frmMyForm"
...VBA code here...
```

**Can be modified with great care:**
- Caption text values (in quotes)
- Simple property values like Width, Height, Left, Top
- Color values, font sizes, and similar visual properties
- Property values where you understand the expected format

**Do NOT modify:**
- Version lines at the top
- Begin/End structure (must remain balanced)
- Control names referenced in VBA code
- Index numbers or internal ordering of controls
- Properties you don't understand (may break layout or cause import errors)

**Important:** The format is largely undocumented. Small, careful changes (like adjusting a text box width) are feasible, but avoid changes that affect control indexing or structural relationships. When in doubt, make layout changes in Access directly.

**Conditional formatting:** When the **Decode conditional formatting to JSON** option is enabled (export format version 5.0.0+), the opaque `ConditionalFormat` / `ConditionalFormat14` binary hex blocks are stripped from each control and stored as decoded, human-readable rules in the companion `.json` file (under `Items.ConditionalFormatting`, keyed by control name). When disabled, the binary blocks remain inline in the source file. Either way the blocks are rebuilt on import. Edit the rules in the `.json` rather than the hex. On import the JSON is authoritative: if a control has a JSON entry, any stale inline block for that control is stripped before the rebuilt block is injected (so JSON wins and re-imports stay idempotent); a control with an inline block but no JSON entry is left untouched. See [docs/access-conditional-format.md](../docs/access-conditional-format.md) for the binary format and the decode/rebuild fidelity boundary (`clsConditionalFormat`).

---

## Query Files

### Source Format: `.sql` + `.json`

Each query is exported as two files:

- **`.sql`** — The SQL statement, formatted for readability. This is the **sole source of truth** for the query's SQL text.
- **`.json`** — Companion metadata: query properties, column metadata, design layout, description, and hidden status.

On import, the add-in reads `.sql` + `.json` and generates a temporary `.qdef` on-the-fly for `LoadFromText`. If the SQL is designer-compatible and layout data exists in the `.json`, the query is imported in Design View format (preserving table positions and designer layout). Otherwise, it uses SQL View format.

### `.sql` File

Contains the Access SQL statement, formatted for readability:

```sql
SELECT
    CustomerID,
    CustomerName,
    Email
FROM tblCustomers
WHERE Active = True
ORDER BY CustomerName;
```

**Safe to modify.** Changes to the `.sql` file are the primary way to edit query logic. On import, the SQL is decomposed and translated into the appropriate format automatically.

### `.json` File

Contains metadata that cannot be expressed in SQL:

- **`QueryType`**: Integer (0=Select, 16=Crosstab, 32=Delete, 48=Update, 64=Append)
- **`Connect`**: Connection string for pass-through queries (may use `env:` reference)
- **`QueryProperties`**: Query properties from the LvProp binary blob (only non-default values)
- **`Columns`**: Column metadata sorted alphabetically for deterministic output
- **`DesignLayout`**: Design View layout (table positions, window dimensions). Only present for queries last saved in Design View.
- **`Description`**: Object description (from DAO document properties)
- **`Hidden`**: Boolean, true if the query is marked as hidden in the Navigation Pane

### Legacy Format (Backward Compatible)

Legacy `.qdef` and `.bas` files are still supported for import. On the next export, they are automatically replaced with `.sql` + `.json` format.

### Recommended Approach for Query Changes

Edit the `.sql` file directly. The companion `.json` preserves design layout and properties automatically. On import, the add-in generates the correct `.qdef` and calls `LoadFromText`.

If the SQL becomes incompatible with Design View (e.g., UNION, subqueries), the layout data from `.json` is ignored and the query is imported as SQL View with a log warning.

---

## Linked Table Connections (`.env` / `env:`)

Linked tables and pass-through queries store `env:conn_*` references (not raw ODBC
strings) in `tbldefs/*.json`, `queries/*.json` (`Connect`), and `db-connection.json`.
At build/import, the add-in resolves each reference to a connection string in the
`.env` file at the export-folder root (per-developer, gitignored).

`db-connection.json` only **primes** Access's ODBC cache during import (one temp
query per connection so later linked-table imports reuse cached credentials). It
holds `env:` keys, not connection strings — do not treat it as a connection registry.

Per-environment `.env` layering via an `APP_ENV` selector is supported; see
[Wiki/Connections.md](../Wiki/Connections.md).

If build prompts for SQL authentication, the referenced key is missing from the
resolved `.env` config or the server is unreachable. Check the build log for
`Connection key not found in .env file: ...`.

---

## Safe Editing Guidelines

### What You CAN Safely Modify

| File Type | Safe Modifications |
|-----------|-------------------|
| `.bas` (modules) | VBA code (after any `Attribute` and `Option` lines) |
| `.cls` (modules) | VBA code, member attribute descriptions |
| `.cls` (form/report code) | Event handlers and procedures |
| `.sql` (queries) | SQL text (source of truth for query logic) |
| `.form` / `.report` / `.bas` (forms/reports) | Simple property values with great care |
| `.json` (most) | Configuration values (maintain valid JSON) |

### What You Must PRESERVE

| Element                            | Reason                          |
| ---------------------------------- | ------------------------------- |
| UTF-8 BOM encoding                 | Required for import             |
| `Attribute VB_Name` lines          | Links code to correct object    |
| VERSION/BEGIN/END blocks in `.cls` | Required file structure         |
| Module-level `Attribute VB_*` lines | Required class metadata        |
| XML structure in `.xml` files      | Must remain valid XML           |
| File names                         | Must match object names exactly |
| Line endings (CRLF)                | Required for import (see Rule 2)|
| Control index/ordering in forms    | Affects layout and tab order    |

---

## Common Tasks

- **Editing VBA code:** Modify `.bas` or `.cls` files in `modules/` (or subdirectories). Add functions after existing code, preserving all `Attribute` and `Option` lines.
- **Editing queries:** Edit the `.sql` file directly in `queries/`. The companion `.json` preserves design layout and properties automatically on import.
- **Editing form/report code:** If "Split Layout from VBA" is enabled, edit `forms/FormName.cls`. Otherwise, find the `CodeBehindForm` section in the `.form` / `.report` file.
- **Testing changes:** Have the user perform a merge build in Access to import and verify changes. See the next section for automated testing.

---

## Automated Testing

The VCS add-in includes a test runner that discovers and executes tests in the current Access database. Tests use `TestAssert` (a drop-in replacement for `Debug.Assert`) for assertions. No compile-time dependency on the add-in is required.

### Setup

Run `VCS.InstallTestAssertModule` from the Immediate Window to inject `modTestAssert` into the project. If the project already uses `Debug.Assert` in test modules, `VCS.MigrateDebugAssert` can batch-convert them to `TestAssert`.

### Writing Tests

Create a standard module with parameterless `Public Sub` procedures that call `TestAssert`:

```vba
Option Compare Database
Option Explicit
Option Private Module
'@Folder("Tests")

Public Sub TestDoubleInput()
    TestAssert MyFunction(42) = 84, "MyFunction should double input"
    TestAssert MyFunction(0) = 0, "Zero input returns zero"
End Sub
```

**Discovery rules:**
- A module is a test module if it has `'@Folder("...Tests...")` (when the project uses `@Folder` annotations) OR its name contains `Test`
- Only parameterless `Public Sub` procedures are discovered as tests
- The module must contain at least one `TestAssert` call, otherwise no tests are registered from it

To exclude helpers from discovery, make them `Private` or give them a parameter.

**Class modules** work the same way but provide built-in setup/teardown: each test method gets a fresh instance, so `Class_Initialize` runs before and `Class_Terminate` runs after every test. Use `Public Sub` or `Public Function` (parameterless, excluding `Class_Initialize`/`Class_Terminate`).

### TestAssert

```vba
TestAssert condition, "optional context for disambiguation"
```

When the VCS test runner is active, `TestAssert` reports each result to the add-in for streaming display. When called outside a test run (or when the add-in is not installed), it falls back to `Debug.Assert`.

### Running Tests

- **Immediate Window:** `?VCS.RunTests`
- **Filtered:** `?VCS.RunTests("modTestEncoding", "-slow")`
- **Ribbon:** Tools > Run Tests
- **Re-run failures only:** The runner supports `RunFailed` after a completed run

`RunTests` accepts an optional `ParamArray` of filter arguments. Each argument is resolved in priority order:

1. **Module name** — exact match on the module/class name
2. **Suite/folder** — match against `@Folder` values (exact or final-segment, e.g., `"SQL"` matches `"Tests.SQL"`)
3. **Procedure name** — match on procedure name or full `Module.Procedure` key
4. **Tag** — match against `'@Tag("...")` annotations

Prefix any argument with `-` to exclude. Inclusions combine with OR; exclusions combine with AND. If only exclusions are specified, the base set is all tests.

```vba
?VCS.RunTests("-slow")                   ' Run all except slow-tagged
?VCS.RunTests("SQL", "-slow")            ' Run SQL suite, skip slow
?VCS.RunTests("TestParseJoinExpression") ' Run one specific procedure
```

### Tagging Tests

Use `'@Tag("name")` annotations (case-insensitive) to categorize tests:

- **Module-level** (first ~30 lines, before any procedure) — applies to all tests in the module
- **Procedure-level** (inside the procedure body, first lines before any code) — applies to that test only

```vba
'@Tag("slow")           ' Module-level: all tests inherit this tag

Public Sub TestExpensiveQuery()
    '@Tag("database")   ' Procedure-level: only this test gets this tag
    TestAssert RunCheck(), "check passes"
End Sub
```

Module-level and procedure-level tags are merged. `RunTests` returns a JSON summary string with per-test status, assertion details, and a `"tags"` array per test.

### Global suite hooks

Optional once-per-run setup/teardown in `modTestAssert` (same module as `TestAssert`):

```vba
Public Sub GlobalTestSetup()    ' Before first test when ≥1 test is selected
Public Sub GlobalTestTeardown() ' After all tests; results JSON already built
```

`VCS.InstallTestAssertModule` includes empty stubs with inline guidance. If absent, the runner skips silently. Hooks do not run when no tests are discovered or a filter matches nothing (@BeforeAll semantics). Hook errors are non-fatal (console only); the run continues and teardown still runs. Per-test `Class_Initialize` / `Class_Terminate` are unchanged and nest inside these hooks.

### Test Logs

After a test run, two log files are written to the `logs/` subfolder inside the export folder:

| File | Contents |
|------|----------|
| `TestResults_<timestamp>.json` | Machine-readable results with per-assertion detail |
| `TestRun_<timestamp>.log` | Full console output including timing |

**Log files are gitignored.** Agent search tools (Glob, Grep, semantic search) will not find them. Use shell commands instead:

```powershell
Get-ChildItem -Path "<export-folder>\logs" -Filter "TestR*" | Sort-Object LastWriteTime -Descending | Select-Object -First 2
```

### Conventions

- Use `'@Folder("Tests")` and `Option Private Module` on test modules
- Name test modules `modTest*` (standard) or `clsTest*` (class)

---

## Build and Merge Workflow

### How Changes Get Back Into Access

1. **User opens Access** with the VCS add-in installed
2. **User initiates build/merge** from the add-in ribbon
3. **Add-in reads source files** and imports them into the database
4. **Conflicts are detected** if both source and database changed

### Build Types

| Type | Description | Use Case |
|------|-------------|----------|
| **Full Build** | Creates new database from source | Fresh start, new developer setup |
| **Merge Build** | Updates existing database with changed files | Incremental updates, pulling changes |
| **Export** | Writes database objects to source files | After making changes in Access |

### Conflict Resolution

When the source file and database object have both changed:
- The add-in detects the conflict
- User is prompted to choose which version to keep
- Conflicts must be resolved before proceeding

---

## Directory Structure Reference

### Root Level Files

| File | Purpose |
|------|---------|
| `vcs-options.json` | Export/import configuration for this project |
| `db-connection.json` | ODBC connection priming map for import (`env:` keys only) |
| `.env` | Per-developer connection strings (gitignored; resolves `env:conn_*` refs) |
| `vcs-index.idx` | Change tracking, binary format (do not edit) |
| `dbs-properties.json` | DAO database properties |
| `documents.json` | DAO database document properties (Description, Hidden for containers) |
| `nav-pane-groups.json` | Navigation Pane custom groups |
| `project.json` | File format version |
| `vbe-project.json` | VBA project properties |
| `vbe-references.json` | VBA library references |

### Logs Folder

Operation logs are stored in the `logs/` subfolder with timestamped filenames:

| Pattern | Purpose |
|---------|---------|
| `Export_YYYYMMDD_HHMMSS_mmm.log` | Export operation logs |
| `Build_YYYYMMDD_HHMMSS_mmm.log` | Build/merge operation logs |
| `TestResults_YYYYMMDD_HHMMSS_mmm.json` | Test runner results (JSON) |
| `TestRun_YYYYMMDD_HHMMSS_mmm.log` | Test runner console output |

**Important: Log files are gitignored.** The `.gitignore` excludes `logs/` directories and `*.log` files. This means agent tools that respect `.gitignore` (such as Glob, Grep, and semantic search) will **not** find these files. To locate and read log files, use shell commands instead:

```powershell
# List log files (run from repository root or source folder)
Get-ChildItem -Recurse -Include "*.log","*.json" | Where-Object { $_.DirectoryName -like "*logs*" }
```

The same applies to `Testing/Fixtures/logs/` (round-trip test logs).

**When troubleshooting:** Always check the most recent log file in `logs/` for detailed operation output, timing information, and any error messages.

### Object Folders

| Folder | Contents | Primary Extension |
|--------|----------|-------------------|
| `forms/` | Access forms | `.form` (or `.bas`), `.cls`, `.json` |
| `reports/` | Access reports | `.report` (or `.bas`), `.cls`, `.json` |
| `modules/` | VBA modules | `.bas`, `.cls` |
| `queries/` | Queries | `.sql`, `.json` (legacy: `.qdef`, `.bas`) |
| `macros/` | Access macros | `.macro` (or `.bas`) |
| `tables/` | Table data (if exported) | `.txt`, `.xml` |
| `tbldefs/` | Table definitions (`Connect` may use `env:conn_*`) | `.sql`, `.xml`, `.json` |
| `relations/` | Table relationships | `.json` |
| `images/` | Shared images | various + `.json` |
| `themes/` | Visual themes | `.thmx` or folder |
| `vbeforms/` | VBA UserForms | `.frm`, `.frx`, `.json` |
| `menus/` | Command bars | `.json` |
| `tdmacros/` | Table data macros | `.xml` |

---

## Troubleshooting

### Import Fails With No Clear Error

**Most likely cause:** Encoding changed from UTF-8 with BOM
**Solution:** Verify file starts with bytes `EF BB BF`, re-save with correct encoding

### Import Fails or Objects Corrupted After Edit (Line Endings)

**Most likely cause:** Line endings changed from CRLF to LF
**Solution:** Verify file contains `\r\n` (hex `0D 0A`) line endings, not just `\n`. Revert the file from Git or re-save with CRLF line endings.

### "Object not found" After Import

**Possible causes:**
- `Attribute VB_Name` doesn't match filename
- File was renamed but internal name wasn't updated
**Solution:** Ensure `VB_Name` matches the filename (without extension)

### VBA Compile Error After Import

**Possible causes:**
- Syntax error introduced in edit
- Missing variable declaration
- Reference to undefined object

**If you are using MCP** (`vcs_compile_vba` returned `success: false`, or
`vcs_check_vba_compiled` returned `compiled: false`):

1. **Stop** — do not edit more source files or guess at fixes.
2. Ask the user to open the database in Access, open the VBE, and choose
   **Debug → Compile**. Access navigates to the first error line.
3. Ask the user to paste the code snippet around that line (a few lines above
   and below). The error message text is optional — the snippet is usually enough.
4. Propose a targeted fix, then re-import and re-compile.

MCP cannot report the failing module or line programmatically.

**If the user is in Access/VBE directly:** use the line Access highlights after
**Debug → Compile** to locate and fix the error.

### Form/Report Looks Wrong After Import

**Possible causes:**
- Control positions or properties were incorrectly modified
- Begin/End blocks became unbalanced
**Solution:** Compare with a known good version, restore the `.form`/`.report` (or legacy `.bas`) file from git

### XML Parse Error

**Possible causes:**
- Invalid XML syntax (unclosed tags, invalid characters)
- Encoding issues with special characters
**Solution:** Validate XML structure, check for unescaped special characters

---

## References

- [GitHub: MSAccess VCS Add-in](https://github.com/joyfullservice/msaccess-vcs-addin)
- [Wiki: Export-Import File Types](https://github.com/joyfullservice/msaccess-vcs-addin/wiki/Export-Import-File-Types)
- [Wiki: Supported Objects](https://github.com/joyfullservice/msaccess-vcs-addin/wiki/Supported-Objects)
- [Wiki: Options](https://github.com/joyfullservice/msaccess-vcs-addin/wiki/Options)
- [VBA Attributes Guide](https://www.thevbahelp.com/post/vba-attributes) - Detailed explanation of VBA module and member attributes

---

*This file is provided by the MSAccess VCS Add-in and placed in exported source folders to assist AI agents working with Access database source files.*
