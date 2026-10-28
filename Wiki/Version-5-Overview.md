# Version 5 Overview

Version **5.0** of the MSAccess VCS Add-in introduces a new **export format** (5.0.0), a redesigned **Options** dialog, deterministic **query** export, a binary change **index**, and optional **MCP** agent automation. Import remains **backward compatible** with projects exported by older add-in versions.

---

## What changed

| Area | Version 5 behavior |
|------|-------------------|
| **File extensions** | `.form`, `.report`, `.macro`, `.sql` + `.json` for queries (when format 5.0 is selected) |
| **Queries** | Primary export is `.sql` + `.json` via MSysQueries reconstruction (**Use Deterministic Query Export**, default On) |
| **Index** | `vcs-index.idx` binary file replaces JSON index — faster scans, smaller footprint |
| **Options UI** | Left navigation: General, Export, Build, Table Data, Databases, MCP, Translation, Advanced, Defaults |
| **Modules** | Rubberduck `@Folder` paths create subfolders under `modules/` |
| **Connections** | `env:conn_*` in exported JSON; secrets in gitignored `.env` ([Connections](Connections)) |
| **MCP** | Opt-in agent permissions ([MCP and Automation](MCP-and-Automation)) |

---

## Export format 5.1.0

Add-in version **5.1** adds export format **5.1.0**. New projects use it automatically; existing projects stay on their current format until you select 5.1.0 under **Options** → **Export**. It changes:

- **Form layout geometry:** controls in Tabular and Stacked layouts are written on a fixed grid, so an untouched form exports the same at any display scaling. See [Form Layout Geometry](Form-Layout-Geometry).
- **Table definitions:** properties in `tbldefs/*.xml` are written in a fixed order. This lets a merge build create a changed local table directly through DAO instead of `ImportXML`, which is dramatically faster on large databases.
- **Queries:** the `OptionFlag` in a query's `.json` is derived from its `.sql` (modifiers such as `DISTINCT` or `WITH OWNERACCESS OPTION`), so the two files can no longer disagree.
- **Import/export specifications:** an unnamed specification keeps its ID, so rebuilding no longer renames its file.
- **Companion `.json` files:** the `Info.Class` header names the owning object type.

Moving from 5.0.0 to 5.1.0 is a smaller step than moving from 4.1.2: run a full export, review the diff (mostly form geometry and table property order), and commit.

---

## Should I upgrade export format?

Upgrade when you are ready to accept a **one-time large Git diff** (extension renames and query file pairs). Benefits:

- Clearer file types in Git and code review
- More stable query exports
- Faster incremental export with the binary index

Stay on **4.1.2** export format until a maintenance window if you have strict change-control on source layout.

---

## Migration checklist

1. **Back up** your `.accdb` / `.mdb` and `.src` folder.
2. Install the [latest add-in release](https://github.com/joyfullservice/msaccess-vcs-addin/releases/latest).
3. Open the database → **Options** → **Export** → set **Export Format** to **5.1.0 (latest)**.
4. Review other export settings (deterministic queries, split layout, `.env` connections).
5. Run a full **Export** (not only Fast Save) so every object migrates.
6. Review the Git diff:
   - Extension renames (`*.bas` → `*.form`, etc.)
   - Query files split into `.sql` + `.json`
   - New or moved `modules/` subfolders from `@Folder`
7. Run a **full build** on a copy of the project to validate round-trip.
8. Commit with a clear message, for example: `chore: migrate VCS export format to 5.1.0`.

**Do not commit** `vcs-index.idx` or `.env` — use the default `.gitignore` template.

---

## Import without migrating export format

You can run add-in **5.x** against a **4.x** source tree. Build and merge still work. The next export while format remains 4.1.2 keeps the old layout; switching format to 5.0 triggers migration on export.

---

## Related pages

- [Query Source Files](Query-Source-Files)
- [Export / Import File Types](Export-Import-File-Types)
- [Options](Options)
- [FAQs](FAQs)
