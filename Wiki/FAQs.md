# Frequently Asked Questions

- [Is there a ribbon toolbar?](#is-there-a-ribbon-toolbar)
- [Why are issues sometimes out of scope?](#why-are-some-issuesideas-considered-out-of-scope-for-this-project)
- [Why do many files show as changed after a build?](#why-am-i-seeing-a-large-number-of-changed-files-after-building-my-project-from-source)
- [How do I export data from all tables?](#how-do-i-also-export-data-from-all-the-tables-in-my-database)
- [My queries changed to .sql and .json — is that normal?](#my-queries-changed-to-sql-and-json--is-that-normal)
- [How do I upgrade to export format 5.0?](#how-do-i-upgrade-to-export-format-50)
- [Merge build vs full build — when to use which?](#merge-build-vs-full-build--when-to-use-which)
- [False conflicts after a full build?](#false-conflicts-after-a-full-build)
- [Where do connection strings go?](#where-do-connection-strings-go)
- [What are @Folder subfolders under modules?](#what-are-folder-subfolders-under-modules)

---

## Is there a ribbon toolbar?

Yes. Version 4 and later ship a **twinBASIC COM ribbon add-in** (32- and 64-bit) installed with the add-in. It provides Export, Build, Merge, Options, Run Tests, and related commands.

If you do not see it, see [Installation](Installation) (COM add-ins, trust, **Use Ribbon Addin**). You can still run the add-in from **Database Tools** → **Add-Ins**.

---

## Why are some issues/ideas considered out of scope for this project?

See [Project Scope](Project-Scope). Examples often declined: full CI/CD pipelines inside the add-in, replacing dedicated SQL tools, or features that only help a single exotic environment without broad benefit.

---

## Why am I seeing a large number of "changed" files after building my project from source?

In normal use, a second export after build should show few or no unintended changes. If Git shows many diffs, consider the cases below.

<details>
<summary><b>Form source files are showing changes in color values</b></summary>

Often caused by different monitors or color profiles. Access stores colors in a way that can shift between machines.

```diff
-     BackColor =11830108
+     BackColor =12874308
```

Try **Sanitize Colors** on the Export options. See [Options](Options).
</details>

<details>
<summary><b>Changes in form dimension values</b></summary>

Common with different screen DPI or monitor layouts. Often safe to ignore. Sanitization removes some report dimension noise; forms may still drift slightly.
</details>

<details>
<summary><b>Query files look completely different (.sql / .json)</b></summary>

Version 5 uses [deterministic query export](Query-Source-Files). The first export after upgrading changes layout from legacy `.qdef` or single `.bas` files. Review the diff once, commit the new shape, then expect stable exports.

If drift continues, check **Use Deterministic Query Export** or temporarily use legacy behavior — see [Query Source Files](Query-Source-Files).
</details>

<details>
<summary><b>Case changes in VBA code</b></summary>

The VBA editor may normalize identifier casing. Tips: Pascal case for procedures; Hungarian-style prefixes for variables (`lngTotal`, `strCaption`).

```diff
-    cancel = True
+    Cancel = True
```
</details>

<details>
<summary><b>Upgraded export format 5.0 (extensions renamed)</b></summary>

Setting **Export Format** to 5.0 renames extensions (`.form`, `.report`, `.sql`, `.macro`). Git sees this as delete+add. Use a dedicated migration commit. See [Version 5 Overview](Version-5-Overview).
</details>

---

## How do I also export data from all the tables in my database?

This tool targets **schema and application design**, not production data backups. Exporting all table data risks committing PII or sensitive records to Git.

For lookup/config tables, select them individually under **Options** → **Table Data**. There is no "select all tables" button by design.

To add many tables manually, edit `vcs-options.json` using the same JSON shape as entries created through the UI (add one table via the UI first as a template).

---

## My queries changed to .sql and .json — is that normal?

Yes for export format 5.0 with **Use Deterministic Query Export** on (default). The `.sql` file holds the SQL text; the `.json` file holds metadata and Design View layout. See [Query Source Files](Query-Source-Files).

---

## How do I upgrade to export format 5.0?

1. Back up your database and `.src` folder.
2. Install the latest add-in release.
3. Open **Options** → **Export** → set **Export Format** to **5.0.0**.
4. Run a full **Export** and review the Git diff.
5. Commit in a single "migrate to VCS export format 5" changeset.

Details: [Version 5 Overview](Version-5-Overview).

---

## Merge build vs full build — when to use which?

| Situation | Recommendation |
|-----------|----------------|
| Daily sync after `git pull` | [Merge build](Merge-Build) |
| Release candidate / clean-room verify | Full build |
| First time using VCS on a database | Full export, then full build on another machine to validate |
| Suspect index or conflict weirdness | Full build, then export |

---

## False conflicts after a full build?

Usually caused by an out-of-sync `vcs-index.idx` or exporting before the database matches source. After a full build, run **Export** once so the index and source align. Do not commit `vcs-index.idx` to Git (default `.gitignore` excludes it).

---

## Where do connection strings go?

In a gitignored `.env` file at the export root, referenced as `env:conn_*` in exported JSON. See [Connections](Connections).

---

## What are @Folder subfolders under modules?

Rubberduck `@Folder("Area.SubArea")` annotations can place modules in subfolders under `modules/` (export format 5.0+). Folders appear only when modules use those annotations.
