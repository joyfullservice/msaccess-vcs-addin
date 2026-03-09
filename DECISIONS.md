<!-- BEGIN HEADER -->

# Decision Log

A reverse-chronological journal of architectural and strategic decisions.
Maintained by AI coding agents (and human developers) at the end of working
sessions. Each entry captures what was decided, what alternatives were
considered, and why — so future contributors never revisit dead ends or lose
context on trade-offs already evaluated.

Agents: read this file before working on any module referenced here.

### When to log

Log decisions that constrain future design, involved genuine alternatives,
or would be non-obvious to a future contributor. A good litmus test: does
the "What this rules out" section have something meaningful to say?

Do NOT log: bug fixes with obvious solutions, test-only refactors,
documentation updates, or minor config tweaks that don't affect
architecture.

### Entry format

Insert new entries directly below this header, newest first. Do not modify
or reorder existing entries except to add supersession notes (see below).
If a session produced multiple independent decisions, create a separate
entry for each.

**Year-end summaries:** When the log rolls into a new calendar year, add
a summary entry titled "Summary of [previous year] decisions" that
briefly describes each decision from that year in one line. This gives
agents scanning forward a checkpoint before older entries.

```
---

## YYYY-MM-DD — [Short descriptive title]

**Trigger**: What problem, requirement, or situation prompted this work.

**Options explored**:
- For each option, name the approach, its strengths, and why it was or
  wasn't chosen. Include options that were tried and reverted.

**Decision**: What was chosen and the core trade-off.

**What this rules out**: Future directions now constrained or foreclosed.
What would trigger revisiting this decision.

**Relevant files**: Key files created or modified.
```

### Guidelines

- Focus on **why**, not what. The diff shows what changed; this log
  explains the reasoning.
- Capture rejected alternatives with equal care. Future agents need to
  know what was already tried.
- Be specific — name libraries, files, config choices, error messages.
- Aim for 10–50 lines per entry. Reference document, not narrative.
- Plain language. No jargon, no editorializing, no padding.

### Superseded entries

When a new decision invalidates, corrects, or replaces guidance in an older
entry, add a blockquote annotation to the affected older entry — do not
rewrite or delete its original text. Place the note immediately after the
entry's heading or after the paragraph containing the superseded claim.

> **⚠ Superseded** (YYYY-MM-DD): [Brief explanation of what changed and
> why.] See "[title of newer entry]" above.

Use **⚠ Partially superseded** when only specific claims are affected, and
**⚠ Superseded** when the entire entry's premise or decision has been
overturned. Always scan older entries for claims that conflict with the new
decision — agents reading the log linearly will otherwise encounter
contradictory guidance.

<!-- END HEADER -->

---

## 2026-03-06 — Export format versioning system

**Trigger**: When users updated the add-in, export format changes (sanitization adjustments, structural tweaks to forms/reports/command bars) would produce hundreds of source file diffs unrelated to the user's actual work. Users couldn't distinguish their five real changes from hundreds of format-upgrade changes, especially mid-feature when the working tree was dirty.

**Options explored**:

- **String-based version constants with helper function**: Constants like `EFV_NORMALIZE_FORM_VIEWPORT = "5.0.0"` with a `ExportFormatAtLeast(strMinVersion)` helper that builds padded comparison strings. Clear and self-documenting per feature, but slower (string comparison at every gate point) and adds an unnecessary helper function. Rejected in favor of enums.
- **Feature-flag booleans derived from format version**: A module that sets `m_blnNormalizeViewport = True` etc. based on the selected version. Single definition point, but adds indirection and a parallel set of variables to maintain. Rejected as over-engineered.
- **Packed-integer enum with native comparison**: `eExportFormatVersion` enum using `Major * 10000 + Minor * 100 + Patch` (e.g., `EFV_4_1_2 = 40102`, `EFV_5_0_0 = 50000`). Gate points use native `>=` comparison: `If Options.ExportFormatVersion >= EFV_5_0_0 Then`. No helper function needed. Chosen.

**Decision**: Introduced `eExportFormatVersion` enum in `modConstants.bas` with packed-integer values, a `LATEST_EXPORT_FORMAT` constant, and an `ExportFormatVersion` Long property on `clsOptions`. The property participates in the existing `m_colOptions`/`CallByName` serialization loop, storing as an integer in `vcs-options.json`. The `Upgrade` method in `clsOptions` converts the loaded `Info.AddinVersion` string to a packed integer via `VersionToExportFormat()` so existing projects default to whatever format they were last exported with (e.g., 4.1.2 projects stay on 40102). New projects default to `LATEST_EXPORT_FORMAT`. Two behaviors are gated behind `>= EFV_5_0_0`: form viewport normalization in `clsSourceParser` and command bar position sanitization in `clsDbCommandBar`. Import remains fully backwards compatible — no gating needed on the import side.

For the UI notification, the main form (`frmVCSMain`) shows a clickable `lblFormatUpdate` label when `ExportFormatVersion < LATEST_EXPORT_FORMAT`, and the export log prints a blue note with the same message. No message boxes — the user upgrades at their convenience via the Options form. Form layout files (`.bas`) are not modified by the agent; controls are added manually in Access to avoid corrupting the binary form structure.

**What this rules out**: Export format changes can no longer be introduced without gating. Every future sanitization or structural change to exported source files must: (1) add an enum member like `EFV_5_1_0 = 50100`, (2) update `[_Last]`, (3) wrap the new behavior in `If Options.ExportFormatVersion >= EFV_5_1_0`. `LATEST_EXPORT_FORMAT` is derived automatically from `eExportFormatVersion.[_Last]`. This is the intended maintenance pattern. The `dblExportFormatVersion` parameter on `BuildJsonFile` in `modVCSUtility` is a separate, older concept for JSON schema versioning and is unrelated to this system. If the packed-integer scheme ever runs out of range (99 minor or 99 patch versions per major), the packing formula would need adjustment, but this is unlikely.

**Relevant files**:

- `Version Control.accda.src/modules/modConstants.bas` — `eExportFormatVersion` enum, `LATEST_EXPORT_FORMAT`
- `Version Control.accda.src/modules/clsOptions.cls` — `ExportFormatVersion` property, default, `Upgrade` migration
- `Version Control.accda.src/modules/modVCSUtility.bas` — `VersionToExportFormat()`, `ExportFormatToVersion()`
- `Version Control.accda.src/modules/clsSourceParser.cls` — viewport normalization gated
- `Version Control.accda.src/modules/clsDbCommandBar.cls` — position sanitization gated
- `Version Control.accda.src/forms/frmVCSOptions.cls` — combo box population logic
- `Version Control.accda.src/forms/frmVCSMain.cls` — format update notification
- `Version Control.accda.src/modules/modImportExport.bas` — export log format version + upgrade note

---

## 2026-03-06 — ObjectDate caching for fast-save change detection

**Trigger**: After building a database from source, a subsequent "fast save" export re-exported every single object (e.g., all 3,673 queries in `sec.accdb`, taking ~1,600s). The existing `IsModified` logic compared `DateModified > ExportDate`, but every object received a fresh `DateModified` from Access during import, making all objects appear modified.

**Options explored**:

- **Keep `DateModified > ExportDate` and fix by updating `ExportDate` after build**: Would require a post-build export pass or index manipulation. Fragile — still uses a directional comparison that can't detect objects restored to an earlier date. Rejected.
- **Content hash comparison for all components**: Would catch every change accurately but is expensive — requires a full export (SaveAsText) of each object just to check, defeating the performance goal. Rejected for date-trackable components; already used by 14 other component types that lack reliable DateModified.
- **Store `ObjectDate` (the object's `DateModified` at export/import time) and compare with exact match (`<>`)**: Records the actual timestamp Access assigned. After a build, the stored ObjectDate matches the current DateModified for unmodified objects. Uses `<>` instead of `>` to also detect objects restored to earlier dates. Chosen.

**Decision**: Added `ObjectDate` field to `clsVCSIndexItem`, stored it in `clsVCSIndex.Update` from `cItem.DateModified`, loaded it in `clsVCSIndex.LoadItem`, and switched all 6 DateModified-based `IsModified` implementations to compare against `ObjectDate` instead of `ExportDate`. Forms and reports retain their secondary `OtherHash` (VBA code module hash) check since VBA edits don't always update `DateModified`. Backward compatible: missing `ObjectDate` in existing index entries defaults to `0`, which never matches a real `DateModified`, so objects are conservatively treated as modified until the first export stores the value.

**What this rules out**: The `ExportDate` field is no longer used for change detection in any component class (though it's still stored and used elsewhere, e.g., conflict detection in `IsExportConflict`). Future component classes that track `DateModified` should use `ObjectDate` for their `IsModified` logic, not `ExportDate`. If Access ever changes how `DateModified` behaves (e.g., sub-second precision, or changing it on compact/repair), the exact-match comparison may need revisiting.

**Relevant files**:

- `Version Control.accda.src/modules/clsVCSIndexItem.cls` — new `ObjectDate` field
- `Version Control.accda.src/modules/clsVCSIndex.cls` — load/save ObjectDate
- `Version Control.accda.src/modules/clsDbQuery.cls` — IsModified updated
- `Version Control.accda.src/modules/clsDbMacro.cls` — IsModified updated
- `Version Control.accda.src/modules/clsDbTableDef.cls` — IsModified updated
- `Version Control.accda.src/modules/clsDbTableDataMacro.cls` — IsModified updated
- `Version Control.accda.src/modules/clsDbForm.cls` — IsModified updated (keeps OtherHash)
- `Version Control.accda.src/modules/clsDbReport.cls` — IsModified updated (keeps OtherHash)
- `Version Control.accda.src/modules/modImportExport.bas` — skip-count logging during fast save

---
