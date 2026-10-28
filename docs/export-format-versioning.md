# Changing what gets exported

Any change that alters the content or structure of exported source files
(sanitization rules, property stripping, file layout, JSON structure) must be
gated so users can upgrade the add-in without being forced to adopt new export
formatting until they choose to. This document covers the three mechanisms for
that, plus the checklists for adding a component type and adding an option.

Import logic does **not** need gating — it must remain backwards compatible with
all prior export formats.

---

## Where the logic lives

- Export: `modImportExport.ExportSource()`
- Import: `modImportExport.Build()`
- Single-object operations: individual `clsDb*.Export()` and `clsDb*.Import()`

## Choosing a mechanism

| Situation | Mechanism |
|-----------|-----------|
| Opt-in output change; old and new formats must coexist | `eExportFormatVersion` gate (`If Options.ExportFormatVersion >= EFV_...`) |
| Bug fix to blind-spot output (sidecars, date-fast-path) | Bump `GetExporterRevisions` for that `IDbComponent.Category` |
| Bug fix to content-hashed primary output | Nothing — `IsModified` self-heals |
| Change to external database schema DDL output | Bump `SCHEMA_EXPORTER_REVISION_*` in `modConstants.bas` |

---

## Export format versions

How to gate a new export behavior change:

1. Add a new member to the `eExportFormatVersion` enum in `modConstants.bas`
   (e.g., `EFV_5_1_0 = 50100`).
2. Add a matching `col.Add EFV_5_1_0` line to `GetExportFormatVersions()`,
   directly below the enum.
3. Wrap the new behavior: `If Options.ExportFormatVersion >= EFV_5_1_0 Then`.
4. Run the gate in [agent-docs-maintenance.md](agent-docs-maintenance.md). A
   format change that alters what a user's source files look like, or which file
   is authoritative for an object type, needs a matching edit to the shipped
   agent docs; one that only changes internals does not.

`GetExportFormatVersions()` is the single list of selectable formats.
`LatestExportFormat()` returns its last entry, and the Options > Export combo box
is populated from it, so neither needs updating by hand. VBA cannot enumerate
enum members at runtime, so the list does repeat the enum — `modTestExportFormat`
parses the enum out of the add-in's own source and fails if the two drift, if the
list is out of order, or if a member's name and packed value disagree.

## Exporter revisions (cache-bust for bug fixes)

When a bug fix changes exported output in a way the change index cannot detect
(sidecar/companion files such as command-bar `_Images`, or components whose
`IsModified` relies on `DateModified`), bump the category's revision in
`GetExporterRevisions()` in `modConstants.bas` instead of adding an export format
version gate. The revision is folded into that category's `CategoryHashes` entry;
on the user's next export, the existing stale-category path re-exports that
category once and persists the new hash.

1. Add or increment the category entry in `GetExporterRevisions()` (key =
   `IDbComponent.Category` string, e.g. `"CommandBars"`).
2. Add a history comment line documenting the fix.
3. Ship the actual exporter fix in the same release.

## External database schemas

Schema exports (`IDbSchema`, `clsSchemaMsSql`, `clsSchemaMySql`) do not
participate in `CategoryHashes` or the component index, so `GetExporterRevisions`
does not reach them. Their only change signal is timestamp equality:
`ExportObject` stamps each exported `.sql` file with the server object's
`last_modified` date, and the next export compares the two. Nothing in that
comparison reflects *how* the DDL was generated.

Two things can change the generated DDL without touching any object's
`last_modified` date:

- **Exporter revision** — a DDL-shape fix in the exporter class. Bump
  `SCHEMA_EXPORTER_REVISION_MSSQL` / `SCHEMA_EXPORTER_REVISION_MYSQL` in
  `modConstants.bas` and add a history comment line.
- **Runtime server capability** — on SQL Server, whether `sp_GetDDL` is
  installed. Installing or removing it switches every object between rich
  SP-generated DDL and the built-in `object_definition()` fallback. This is
  probed per connection and needs no code change to take effect.

Both are folded into a per-schema fingerprint stored in the index's
`SchemaState` section (`VCSIndex.SchemaState(<schema name>)`). When the recorded
fingerprint no longer matches, the exporter treats every object in that schema as
modified for one export, then records the new value. A schema with no recorded
fingerprint re-exports once to establish a baseline, so the first export after
upgrading re-exports schema objects (content is normally identical, so git shows
no diff).

`IDbSchema.Export(blnFullExport)` is also honored by both exporters, so a full
export re-exports all schema objects regardless of dates. When `VCSIndex.Disabled`
is set there is nowhere to record the fingerprint, so the capability check is
skipped rather than forcing a re-export on every run.

---

## Adding a new component type

1. Create a new class `clsDbNewType.cls` implementing `IDbComponent`.
2. Add a new enum value to `eDatabaseComponentType` in `modConstants`.
3. Add the class to `GetContainers()` in `modVCSUtility`.
4. Implement all interface methods (Export, Import, Merge, GetAllFromDB, etc.).
5. Declare file extensions in `FileExtensions`:
   - **`efesIndexed` (default):** authoritative files used for hash/change
     detection, conflict resolution, and merge detection (`FilePropertiesHash` +
     `AllFilesHash`). Form/report companion `.json` (metadata, conditional
     formatting sections) belongs here.
   - **`efesAll`:** indexed set plus derived sidecar files only (e.g. form/report
     `.svg` previews). Orphan cleanup and `MoveComponentSource` read `efesAll`.
     Indexed companion files **must** be produced on the alternate/temp export
     path so `GetDifferingFiles` file counts stay balanced.
   - Sidecar cleanup for derived-only files is automatic via
     `modOrphaned.ClearOrphanedComponentArtifacts` (`efesAll − efesIndexed`).
   - Per-object **folders** (command-bar `_Images`, extracted theme folders) are
     not flat extensions. Add a branch to `ClearOrphanedComponentFolders` in
     `modOrphaned.bas` and move the folder in `MoveSource`.
6. Run the gate in [agent-docs-maintenance.md](agent-docs-maintenance.md). A new
   component type adds a row to the "Which file do I edit?" table in the shipped
   `AGENTS.md` only if users will hand-edit its source files; if they will not,
   ship nothing.

## Adding an option

1. Add a public property to `clsOptions`.
2. Add the default value in `clsOptions.LoadDefaults()`.
3. Update `GetOptionsDictionary()` and the loading code.
4. Update the `frmVCSOptions` form if user-configurable.
