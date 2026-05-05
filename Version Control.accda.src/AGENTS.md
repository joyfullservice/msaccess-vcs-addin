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
- **`Properties`**: Query properties from the LvProp binary blob (only non-default values)
- **`Columns`**: Column metadata sorted alphabetically for deterministic output
- **`DesignLayout`**: Design View layout (table positions, window dimensions). Only present for queries last saved in Design View.
- **`Properties`** / **`Hidden`**: Document metadata (Description, Hidden attribute)

### Legacy Format (Backward Compatible)

Legacy `.qdef` and `.bas` files are still supported for import. On the next export, they are automatically replaced with `.sql` + `.json` format.

### Recommended Approach for Query Changes

Edit the `.sql` file directly. The companion `.json` preserves design layout and properties automatically. On import, the add-in generates the correct `.qdef` and calls `LoadFromText`.

If the SQL becomes incompatible with Design View (e.g., UNION, subqueries), the layout data from `.json` is ignored and the query is imported as SQL View with a log warning.

### Before changing the query parser

The query parser (`clsQueryComposer.cls` + `clsDbQuery.cls`) carries hard-won decisions in places that are not always obvious from a casual read. Before modifying either class, read these in order:

Do not look in `Testing.accdb.src` for query regression fixtures; the shipped
round-trip corpus is `../Testing/Fixtures/queries/`.

- **[../docs/access-query-storage.md](../docs/access-query-storage.md)** — in-repo reference for how Access stores queries, what shapes our parser handles (with the canonical fixture for each), known gaps where behaviour is unverified, and findings unique to our pipeline (`Application.LoadFromText` / `Application.SaveAsText` asymmetries).
- **[../DECISIONS.md](../DECISIONS.md)** — search for entries mentioning `clsQueryComposer` or `clsDbQuery` (e.g. `rg "clsQueryComposer" DECISIONS.md -A 30`). Captures the rationale and rejected alternatives behind each choice.
- **`../Testing/Fixtures/queries/regression/*.notes.md`** — each one pins a specific SQL shape and explains what would re-break if a careful decision were reverted.
- **Procedure-header comments** on the functions you're modifying — `RequiresDesignView`, `IsDesignerCompatible`, `HasTopLevelBoolean`, `ParseJoinExpression`, `SafeBreak`, `EmitDbMemoSql` carry constraints in their headers that the body alone does not convey.

When you discover a new invariant or edge case worth preserving, follow the four-layer documentation pattern at [../Testing/Fixtures/README.md § Documenting parser invariants and edge cases](../Testing/Fixtures/README.md).

### VBA error-handler cleanup

Inside an active `On Error GoTo Handler` block, `Err.Clear` clears the error
object but does not reset the active handler state. If cleanup code in that
handler may raise expected errors, call `On Error GoTo -1` before switching to
`On Error Resume Next`, and use `GoTo` rather than `Resume` afterward:

```vba
Handler:
    Err.Clear
    On Error GoTo -1
    On Error Resume Next
    CurrentDb.QueryDefs.Delete "__temp_query__"
    Err.Clear
    On Error GoTo 0
    GoTo ContinueAfterHandler
```

Prefer checking for object existence before deletion when practical.

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

### Adding a Function to a Module

1. Open the `.bas` file in `modules/`
2. Add your function after existing code:
```vba
Public Function CalculateTotal(dblPrice As Double, intQty As Integer) As Double
    CalculateTotal = dblPrice * intQty
End Function
```

### Modifying a SQL Query

Edit the `.sql` file directly in `queries/`:

```sql
SELECT CustomerID, CustomerName, Email, Phone
FROM tblCustomers
WHERE Active = True
ORDER BY CustomerName;
```

The companion `.json` preserves design layout and properties. On import, the add-in generates the appropriate `.qdef` format automatically. Design View layout is preserved when the SQL remains designer-compatible.

### Fixing a Bug in Form Code

If "Split Layout from VBA" is enabled:
1. Open `forms/FormName.cls`
2. Find and modify the relevant procedure

If code is embedded in the layout file (`.form` or legacy `.bas`):
1. Open `forms/FormName.form` (or `FormName.bas` in older projects)
2. Find the `CodeBehindForm` section near the end
3. Modify code in that section

### Bulk Find and Replace

When making changes across multiple files:
1. Search within the appropriate folder (`modules/`, `forms/`, etc.)
2. Test by having the user perform a merge build in Access

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
| `vcs-index.idx` | Change tracking, binary format (do not edit) |
| `dbs-properties.json` | DAO database properties |
| `project.json` | File format version |
| `vbe-project.json` | VBA project properties |
| `vbe-references.json` | VBA library references |

### Logs Folder

Operation logs are stored in the `logs/` subfolder with timestamped filenames:

| Pattern | Purpose |
|---------|---------|
| `Export_YYYYMMDD_HHMMSS_mmm.log` | Export operation logs |
| `Build_YYYYMMDD_HHMMSS_mmm.log` | Build/merge operation logs |

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
| `tbldefs/` | Table definitions | `.sql`, `.xml`, `.json` |
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
**Solution:** Check the specific line mentioned in the error

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

*This file is automatically generated by MSAccess VCS Add-in and placed in exported source folders to assist AI agents working with Access database source files.*
