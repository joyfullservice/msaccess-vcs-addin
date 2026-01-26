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

### 1. PRESERVE UTF-8 BOM ENCODING (Most Important)

**ALL text files** in this exported database use **UTF-8 with BOM** encoding (bytes `EF BB BF` at file start). This includes:
- `.bas` and `.cls` files (VBA modules and classes)
- `.sql` files (query SQL)
- `.json` files (configuration and metadata)
- `.xml` files (table definitions, etc.)
- `.txt` files (table data exports)
- Any other text-based source files

This encoding is **mandatory** for successful import. **Encoding errors cause import failures with no clear error message.**

**You MUST:**
- Preserve UTF-8 BOM encoding on ALL text files you edit
- Be aware that many editing tools and file operations may strip the BOM
- Verify encoding after edits if you suspect it may have been changed

**You MUST NOT:**
- Remove the BOM from any file
- Change encoding to UTF-8 without BOM, ASCII, ANSI, or any other encoding
- Add BOM to files that don't have one (check first)
- Assume editing tools preserve encoding - they often don't

**Verification and Restoration:**
If you need to verify or restore BOM encoding on any file:

```powershell
# Check if file has BOM
$file = "path\to\file.ext"
$bytes = [System.IO.File]::ReadAllBytes($file)
$hasBOM = ($bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF)
Write-Host "BOM present: $hasBOM"

# Restore BOM if missing
if (-not $hasBOM) {
    $bom = [byte[]](0xEF, 0xBB, 0xBF)
    $newBytes = $bom + $bytes
    [System.IO.File]::WriteAllBytes($file, $newBytes)
    Write-Host "BOM restored"
}
```

**Note:** The add-in's export process automatically ensures all files have UTF-8 BOM. The risk occurs when files are edited outside of Access. Always verify encoding if you're unsure.

### 2. Preserve VBA File Structure

VBA files (`.bas`, `.cls`) have required headers that must not be modified. See "VBA File Format" section below.

### 3. Do Not Edit These Files

- **`vcs-index.json`** - Managed automatically by the add-in (tracks file hashes)
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

## Form and Report Object Files

### Object Definition Files (`.bas` in `forms/` and `reports/`)

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

### Important: Which File Is Used for Import?

**By default, the `.bas` file is used for import, NOT the `.sql` file.**

The `.sql` file exists for easier reading and cleaner diffs, but the add-in imports from the `.bas` file unless the option `ForceImportOriginalQuerySQL` is enabled in `vcs-options.json`.

**Why not always use `.sql`?** The `.bas` file preserves query designer properties (like field descriptions, text format settings, column widths) that are lost when importing from pure SQL.

### Query Definition (`.bas`)

The `.bas` file uses Access's `SaveAsText` format. There are **two different formats** depending on how the query was last saved:

**Format 1: Saved from SQL View (simpler)**
```
Operation =1
Option =0
dbMemo "SQL" ="SELECT CustomerID, CustomerName FROM tblCustomers WHERE Active = True"
```

**Format 2: Saved from Query Designer (complex)**
```
Operation =1
Option =0
Begin InputTables
    Name ="tblCustomers"
End
Begin OutputColumns
    Expression ="CustomerID"
    Name ="CustomerID"
    ...
End
dbMemo "SQL" ="SELECT tblCustomers.CustomerID..."
```

**Warning:** The `.bas` format is largely undocumented and easy to corrupt. Modifications are risky.

### SQL File (`.sql`)

Contains just the SQL statement, formatted for readability:

```sql
SELECT
    CustomerID,
    CustomerName,
    Email
FROM tblCustomers
WHERE Active = True
ORDER BY CustomerName;
```

### Recommended Approaches for Query Changes

**Option 1: Provide SQL to user (safest)**
- Show the user the desired SQL changes
- Let them make the modification directly in Access
- Have them re-export to update source files

**Option 2: Edit `.sql` and enable import option**
1. Modify the `.sql` file with your changes
2. Set `"ForceImportOriginalQuerySQL": true` in `vcs-options.json`
3. User rebuilds from source
4. **Caveat:** Query designer properties (descriptions, text formats) will be lost

**Option 3: Edit `.bas` directly (risky)**
- Only attempt for simple changes to the `dbMemo "SQL"` line
- Do not modify the structural elements (InputTables, OutputColumns, etc.)
- Test thoroughly after import

---

## Safe Editing Guidelines

### What You CAN Safely Modify

| File Type | Safe Modifications |
|-----------|-------------------|
| `.bas` (modules) | VBA code (after any `Attribute` and `Option` lines) |
| `.cls` (modules) | VBA code, member attribute descriptions |
| `.cls` (form/report code) | Event handlers and procedures |
| `.sql` (queries) | SQL text (but see Query Files section for import caveats) |
| `.bas` (forms/reports) | Simple property values with great care |
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
| Line endings (CRLF)                | Windows standard for VBA        |
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
3. Save with UTF-8 BOM encoding

### Modifying a SQL Query

**Recommended approach:** Provide the user with the desired SQL and let them make the change in Access, then re-export.

**Alternative (if user enables SQL import):**
1. Edit the `.sql` file in `queries/`:
```sql
SELECT CustomerID, CustomerName, Email, Phone  -- Added Phone
FROM tblCustomers
WHERE Active = True
ORDER BY CustomerName;
```
2. Set `"ForceImportOriginalQuerySQL": true` in `vcs-options.json`
3. Save with UTF-8 BOM encoding
4. User rebuilds from source

**Note:** Importing from `.sql` loses query designer properties (field descriptions, text formats, etc.)

### Fixing a Bug in Form Code

If "Split Layout from VBA" is enabled:
1. Open `forms/FormName.cls`
2. Find and modify the relevant procedure
3. Save with UTF-8 BOM encoding

If code is embedded in `.bas`:
1. Open `forms/FormName.bas`
2. Find the `CodeBehindForm` section near the end
3. Modify code in that section
4. Save with UTF-8 BOM encoding

### Bulk Find and Replace

When making changes across multiple files:
1. Use an editor that preserves UTF-8 BOM (VS Code, Notepad++, etc.)
2. Search within the appropriate folder (`modules/`, `forms/`, etc.)
3. Verify encoding is preserved after save
4. Test by having the user perform a merge build in Access

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
| `vcs-index.json` | Change tracking (do not edit) |
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
| `forms/` | Access forms | `.bas`, `.cls`, `.json` |
| `reports/` | Access reports | `.bas`, `.cls`, `.json` |
| `modules/` | VBA modules | `.bas`, `.cls` |
| `queries/` | Queries | `.bas`, `.sql` |
| `macros/` | Access macros | `.bas` |
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
**Solution:** Compare with a known good version, restore `.bas` file from git

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
