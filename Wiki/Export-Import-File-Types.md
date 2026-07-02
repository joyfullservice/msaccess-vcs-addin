# Export / Import File Types

Reference for the folder and file layout created when you export source. Your `.src` folder only contains directories for object types that exist in your database. See [Supported Objects](Supported-Objects) for the full matrix.

**Export format 5.0** uses descriptive extensions (`.form`, `.report`, `.sql`, `.macro`). Older projects may still show `.bas` until you upgrade — see [Version 5 Overview](Version-5-Overview).

**Secrets:** Connection strings are usually stored in a gitignored `.env` at the export root with `env:conn_*` placeholders in JSON. See [Connections](Connections).

> **Is this list exhaustive?** No — this page is a *representative* reference covering the folders and file types you will encounter in a typical project. There is no single document that enumerates every possible file name, because the layout is defined per object type in code (each `clsDb*` component sets its own folder and extension). The authoritative, complete picture is simply your own `.src` folder after an **Export**: folders are only created for object types that exist in your database. For the full object-type matrix (and which `clsDb*` class handles each), see [Supported Objects](Supported-Objects).

---

## Folder tree (typical)

```
.
├── .gitattributes
├── .gitignore                      # Excludes .env, vcs-index.idx, database binary, logs/
├── .env                            # Gitignored — real connection strings (optional)
├── MyApp.accdb                     # Database (usually gitignored)
└── MyApp.accdb.src/
    ├── forms/
    │   ├── MyForm.form             # Layout (legacy: .bas)
    │   ├── MyForm.cls              # Code-behind if Split Layout is On
    │   └── MyForm.json             # Print settings (if any)
    ├── reports/
    │   ├── MyReport.report
    │   ├── MyReport.cls
    │   └── MyReport.json
    ├── queries/
    │   ├── MyQuery.sql             # SQL text (source of truth)
    │   └── MyQuery.json            # Metadata + Design View layout
    ├── modules/
    │   ├── MyModule.bas
    │   └── Area/SubArea/           # @Folder subfolders (export format 5.0+)
    │       └── NestedModule.bas
    ├── macros/
    │   └── AutoExec.macro
    ├── tbldefs/
    ├── tables/                     # Optional table *data* only
    ├── relations/
    ├── images/
    ├── themes/
    ├── databases/                  # External SQL schema export (optional)
    ├── dbs-properties.json
    ├── db-connection.json          # Named connection metadata (env: placeholders)
    ├── vcs-options.json
    ├── vcs-index.idx               # Binary change index — pair with DB, gitignore
    ├── project.json
    └── ...
```

---

## Root-level project files

| File | Purpose |
|------|---------|
| `vcs-options.json` | Per-project export/build/MCP options |
| `vcs-index.idx` | Binary fast-save index; **do not commit** to Git |
| `project.json` | Access file format version |
| `db-connection.json` | Connection definitions with `env:conn_*` values when using `.env` |
| `dbs-properties.json` | DAO database properties |
| `vbe-references.json` | VBA references |
| `Export.log` / `Build.log` | Operation logs (often gitignored) |

---

## Forms (`forms/`)

| File | When |
|------|------|
| `*.form` | Form layout (export format 5.0; legacy `*.bas` still imports) |
| `*.cls` | Code-behind when **Split Layout from VBA** is On |
| `*.json` | Optional print settings |

See [Split Files](Split-Files).

---

## Reports (`reports/`)

Same pattern as forms: `*.report`, optional `*.cls`, optional `*.json` for printer settings.

---

## Queries (`queries/`)

| File | When |
|------|------|
| `*.sql` | Query SQL text — edit this in Git for logic changes |
| `*.json` | Type, columns, properties, Design View layout |

Legacy `*.qdef` / `*.bas` import cleanly; the next export replaces them with `.sql` + `.json` when deterministic export is enabled.

See [Query Source Files](Query-Source-Files).

---

## Modules (`modules/`)

| File | When |
|------|------|
| `*.bas` | Standard modules |
| `*.cls` | Class modules |

Rubberduck `@Folder("Area.SubArea")` annotations create subfolders under `modules/` when using export format 5.0.0+.

---

## Table definitions and data

| Folder | Contents |
|--------|----------|
| `tbldefs/` | Table structure — `.xml` (local), `.json` (linked), optional `.sql` |
| `tables/` | Table **data** only for tables selected in Options → Table Data |
| `tdmacros/` | Table data macros |

---

## External SQL schema (`databases/`)

When configured under **Options** → **Databases**, the add-in exports read-only snapshots of linked SQL Server or MySQL schema objects. **Export only** — it does not apply changes to the server. See [Options](Options#databases-external).

---

## Legacy extensions

Import accepts older layouts. After you set **Export Format** to 5.0 and run **Export**, files are renamed/migrated to the new extensions (Git may show renames as delete+add).

| Legacy | Current (5.0) |
|--------|----------------|
| `forms/*.bas` | `forms/*.form` |
| `reports/*.bas` | `reports/*.report` |
| `macros/*.bas` | `macros/*.macro` |
| `queries/*.qdef` | `queries/*.sql` + `*.json` |

---

## Related pages

- [Connections](Connections) — `.env` and `env:conn_*`
- [Options](Options) — control what gets exported
- [Version 5 Overview](Version-5-Overview) — migration checklist
