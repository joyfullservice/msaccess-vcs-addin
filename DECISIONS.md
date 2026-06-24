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

## 2026-06-23 — Full-build module import: two-pass ImportFast + FinalizeImports

**Trigger**: Full builds on module-heavy projects spend ~85% of the `Modules`
category time in the per-file tail (save, `DoEvents`, `AllModules` retry,
metadata, hash, index) rather than the VBE `.Import` itself (~0.4 s for 135
modules per the 2026-05-29 measurement). The 2026-05-29 decision added
per-module synchronization for correctness; this entry revisits the deferred
batch path that decision's "Revisit if" clause anticipated.

**Options explored**:
- **Keep per-module tail in `IDbComponent_Import` (status quo)** — robust;
  O(N) interleaved `DoEvents`/`AllModules` retries and per-module
  `Documents.Refresh` inside `ImportObjectMetadata` limit throughput on large
  projects.
- **Add `AfterCategoryImport` to `IDbComponent`** — reusable hook, but forces
  ~28 empty stubs on every component class for a problem unique to modules.
- **Add a second interface (`IDbBatchImport`) on `clsDbModule`** — opt-in
  without universal stubs, but adds a second `Implements` contract to the
  class for a single consumer.
- **Public `ImportFast` / `FinalizeImports` on `clsDbModule` + component-type
  branch in `modBuild` (chosen)** — same pattern as existing special cases
  (`InitializeForms`, merge skip for `edbTableData`). Full builds call
  `ImportFast` per file, then `FinalizeImports` once; merge and single-object
  import keep `IDbComponent_Import` unchanged.

**Decision**: Pass 1 (`ImportFast`) parses and loads via VBE only, recording
`{sourceFile, moduleName, blnPublicCreatable}` in a batch collection.
Pass 2 (`FinalizeImports`) runs `DoCmd.Save acModule` for each batched
module (one `DoEvents` after the loop), one `Documents.Refresh`, then per-file
resolve (`GetAccessModuleObject` with bounded retry), metadata
(`ImportObjectMetadata` with optional skip-refresh), and `VCSIndex.Update`.
`acCmdCompileAndSaveAllModules` was tried first but does not reliably publish
unsaved VBE imports to `AllModules` when the project does not compile yet.
`m_strSourceFile` is set explicitly per file in pass 2 so the shared instance
never indexes under a stale path.

**What this rules out**: Removing bounded `AllModules` retry entirely.
Deferring metadata/index without pass-2 re-resolve. Applying the batch path
to merge builds without revisiting export-after-merge and conflict semantics.

**Relevant files**: `clsDbModule.cls` (`ImportFast`, `FinalizeImports`,
`FinalizeOneModule`), `modBuild.bas` (full-build `edbModule` branch),
`modLoadSaveText.bas` (`ImportObjectMetadata` optional skip-refresh),
`modTestConflicts.bas` (`TestModuleImportFast_IndexesEachFileOnSharedInstance`).
Supersedes the "Revisit if" clause on 2026-05-29 module import sync.

---

## 2026-06-19 — Gracefully skip engine-managed DAO properties on import (error 3916)

**Trigger**: Building a database with linked tables that use a newer data type (e.g.
DateTime2) failed three table imports with `Error 3916: The property 'FCMinWriteVer' can
only be set or changed by the Microsoft Access database engine`. Access stamps the
`FCMin*` family (`FCMinDesignVer`/`FCMinReadVer`/`FCMinWriteVer` -- "Feature Compatibility
minimum version") on objects that use such features. These had been captured into the
linked-table `.json` (`TableProperties`), and replaying them on import via `SetDAOProperty`
raised 3916, which aborted the table import and inflated the build's error count -- even
though the table itself had already been linked successfully.

**Are these properties worth preserving?** No. `FCMin*` are derived, engine-managed version
stamps: Access regenerates them automatically from the object's actual structure when it is
recreated, they cannot be set by code at all (that is what 3916 means), and their values are
build/machine-specific (`16.0.12600.10000`), so storing them only produces noisy,
non-portable diffs.

**Options explored**:
- *Name-based skip* (initial fix) — skip the known `FCMin*` names on import and strip them
  on export. Works, but brittle: any future engine-managed property would reintroduce the
  same hard failure until its name was added.
- *Generic tolerance in `SetDAOProperty` (chosen)* — catch error 3916 when applying any
  property, skip it with a debug note, and re-raise every other error so real failures still
  surface. This is inherently safe: it only ever skips a property the engine refuses to let
  us set (i.e. an engine-managed/derived one); every property we can legitimately set is
  still applied, so nothing meaningful is lost.

**Decision**: `modDatabase.SetDAOProperty` now wraps the property mutation, swallows error
3916 (debug-note only), and re-raises anything else. This makes all property importers
resilient -- linked tables (`clsDbTableDef`), document properties (`clsDbDocument`,
`modLoadSaveText`) -- and removes the need for the name-based import skip, which was reverted.
As a separate source-cleanliness measure (not a correctness requirement), export still strips
the `FCMin*` family from linked-table `TableProperties`, gated at `EFV_5_0_0`, via the now-
public `modDatabase.IsEngineManagedProperty` / `FilterEngineManagedProps` helpers. Import is
not gated and stays backward compatible with older source that still contains these stamps.

**What this rules out**: We no longer fail an import when a property is engine-managed; the
trade-off is that a genuinely unsettable property is silently skipped (visible only with
ShowDebug). Other property-set errors are still surfaced unchanged.

**Relevant files**: `modDatabase.bas` (`SetDAOProperty`, `IsEngineManagedProperty`,
`FilterEngineManagedProps`), `clsDbTableDef.cls` (export filter; reverted import skips),
`modConstants.bas` (`EFV_5_0_0` comment), `modTestDatabase.bas` (helper tests).

---

## 2026-06-19 — Never write raw passwords to source files (any mode)

**Trigger**: The 2026-03-17 `.env` design left `UseEnvForConnections = Never` defined as
"keep complete connection strings in source," which means a SQL-auth password is written
verbatim into committed source files. A user can pick `Never` (or hit a non-externalized
path) without realizing credentials will land in a public repo. The risk is amplified as
AI agents author/edit database projects with less human review in the loop.

**Options explored**:
- *Warn only* — log a warning when a password is written to source, leave behavior as-is.
  Surfaces the issue but still ships the secret.
- *Strip only PWD, keep UID* — removes the secret but leaves the username; simpler but
  inconsistent with the existing UID/PWD pairing in `SanitizeConnectionString`.
- *Redefine `Never` to strip credentials; gate at export format 5.0.0 (chosen)* — passwords
  are never written to source in any mode; `Never` means "connection strings in source,
  minus credentials." Users who want self-contained source must manage credentials
  themselves (runtime prompt or their own priming).

**Decision**: New `modConnect.GetSourceSafeConnect` (gated by `EFV_5_0_0`) strips `UID`/`PWD`
from any connection string written to source when it is not externalized to `.env`, and
logs one `eelWarning` per distinct connection. It only acts when an actual `PWD` value is
present (so passwordless AD/integrated connections, which may carry an empty `PWD=`, do not
trip a false warning). Applied uniformly at all three connection-bearing exporters:
`clsDbTableDef` (linked tables), `clsDbQuery` (pass-through queries), and `clsDbConnection`
(`db-connection.json`, inner + outer keys). `clsDbQuery` previously emitted an `env:`
reference without calling `SaveConnectionToEnv`; that gap is now fixed so the `.env` is
populated for pass-through queries too. Import is unchanged and remains backward compatible
with older source that still contains credentials.

**What this rules out**: Self-contained source files with embedded passwords are no longer
supported at export format 5.0.0+. Existing repos keep the old behavior until they bump
their `ExportFormatVersion`, so the secret-leak window persists for un-migrated projects
(mitigated by the warning when stripping occurs). Stripping covers `UID`/`PWD` only — if a
driver carries a secret under a different key, `StripConnectionCredentials` would need
extending.

**Testing & accepted risk**: Locked in by unit tests on the single chokepoint —
`modTestConnect.TestStripConnectionCredentials` (the strip logic across SQL-auth, Access
back-end, lower-case keys, passwordless AD, and no-credential shapes) and
`TestGetSourceSafeConnectGating` (the `EFV_5_0_0` gate: passthrough below 5.0.0, strip at/above,
no-op for empty `PWD=` and credential-free strings). A full end-to-end test (link a
password-protected back-end table, run the component export to file, grep the output for `PWD=`)
was considered and deliberately *not* implemented: driving the real export path mutates shared
state (the live `VCSIndex`, the project export folder, the log) and is flaky inside the unit
suite, while a temp-linked table without the export-to-file step only re-tests
`GetSourceSafeConnect` with a live string. Consequence: the unit tests guard the strip/gate
logic but do **not** catch a refactor that removes a `GetSourceSafeConnect` *call site*. That
gap is mitigated by an explicit SECURITY reminder comment at each of the three call sites
(`clsDbTableDef`, `clsDbQuery`, `clsDbConnection`) and is accepted for now.

**Relevant files**: `modConnect.bas` (`GetSourceSafeConnect`, `StripConnectionCredentials`,
`m_dStrippedConnWarn`), `clsDbTableDef.cls`, `clsDbQuery.cls`, `clsDbConnection.cls`,
`modConstants.bas` (`EFV_5_0_0` comment), `frmVCSOptionsExport.cls` (help text),
`modTestConnect.bas` (regression tests).

---
## 2026-06-20 — Fold unreleased 5.1.0 export gates into format 5.0.0

**Trigger**: Several v5 behaviors — conditional formatting decode-to-JSON, source-safe
connection strings (no raw passwords in source), and linked-table `FCMin*` export
filtering — were initially gated behind unreleased `EFV_5_1_0`, but v5 has not shipped
to the general public yet (only a handful of beta users). Keeping a separate 5.1.0 format
version would make 5.0.0 an incomplete "first release" snapshot.

**Options explored**:
- **Keep `EFV_5_1_0` for these behaviors**: Clean separation, but forces the first general
  release to advertise two format versions when only one meaningful baseline is needed.
- **Fold into `EFV_5_0_0` (chosen)**: Same precedent as file extension migration
  (2026-03-10). All unreleased v5 behaviors ship as part of the v5 baseline.
- **Auto-migrate beta `"5.1.0"` in `clsOptions.Upgrade()`**: Rejected — only one known
  beta user; manual `vcs-options.json` edit is sufficient.

**Decision**: Remove `EFV_5_1_0` from `eExportFormatVersion`, set `[_Last] = 50000`, and
retarget all gate sites from `>= EFV_5_1_0` to `>= EFV_5_0_0`:
- CF decode: `clsSourceParser`, `modLoadSaveText` (the `DecodeConditionalFormatting`
  option gate is unchanged)
- Source-safe connections: `modConnect.GetSourceSafeConnect` and its three call sites
- `FCMin*` export filtering: `clsDbTableDef` via `FilterEngineManagedProps`

No runtime migration for stale `"5.1.0"` values in `vcs-options.json` (50100 still
satisfies `>= 50000` if left untouched).

**What this rules out**: These behaviors are no longer post-5.0.0 format bumps; they are
part of the v5 baseline. The `EFV_5_1_0 = 50100` slot is free again for the first
*post-release* export format change.

**Relevant files**: `modules/Infrastructure/modConstants.bas`, `modules/Core/clsSourceParser.cls`,
`modules/Core/modLoadSaveText.bas`, `modules/Utility/modConnect.bas`,
`modules/Components/clsDbTableDef.cls`, `modules/Components/clsDbQuery.cls`,
`modules/Components/clsDbConnection.cls`, `forms/frmVCSOptionsExport.cls`,
`modules/Tests/Connect/modTestConnect.bas`, `docs/access-conditional-format.md`.

## 2026-06-18 — Build-time cleanup for duplicate `@Folder` source files

**Trigger**: AI agents repeatedly created a second copy of a VBA module in the wrong
folder (e.g. `modules/modTestRoundtrip.bas` alongside `modules/Tests/modTestRoundtrip.bas`)
because file placement is driven by the `'@Folder` comment inside the file, not the folder
being edited. Build/import scanned both copies recursively and silently last-one-wins; orphan
cleanup did not remove them because the DB object still existed. Export already deleted
stale copies per module via `CleanupDuplicateSourceFiles`, but build had no equivalent.

**Options explored**:
- *Agent guidance only* — document the rule in AGENTS.md and `.cursor/rules`. Cheap but
  agents still miss it; duplicates persist until someone exports from Access.
- *Import warning only* — detect duplicates and warn without deleting. Surfaces the problem
  but still requires manual cleanup and leaves merge-index false positives.
- *Build-time auto-cleanup + guidance + export warning (chosen)* — before build/merge scan,
  group module files by basename; parse each file's `@Folder` from text; when exactly one copy
  sits in its annotation-derived folder, delete the others; ambiguous groups warn and are
  left alone.

**Decision**: Add `GetFolderAnnotationFromText` (shared with live VBE reader) and
`RemoveDuplicateComponentFiles` (with module/form/report wrappers), called at the start of
`modBuild.Build` for the `modules/`, `forms/`, and `reports/` base folders. Duplicate
detection keys on **distinct folders** per basename (not raw file count), so a form's
`.form` + `.cls` + `.json` in one folder is not treated as duplicates. For forms/reports,
`@Folder` is read from the `.cls` code-behind when present. `WarnDuplicate*Basenames`
runs after export as a safety net. Agent docs updated to require searching the full
component tree before creating a source file.

**What this rules out**: We do not auto-delete when zero or multiple copies match their
annotation path (divergent edits or two agents writing different folders). Those cases log
a warning and keep current last-one-wins import behavior until a human resolves them.
We do not relocate a lone misplaced instance with no duplicate to compare against — export
handles moves via `MoveSource` + `CleanupDuplicateSourceFiles`.

**Relevant files**: `modules/Core/modVbeUtility.bas` (`GetFolderAnnotationFromText`,
`RemoveDuplicateModuleFiles`, `WarnDuplicateModuleBasenames`), `modules/Core/modBuild.bas`,
`modules/Core/modExport.bas`, `modules/Tests/Core/modTestFolderPlacement.bas`,
`.cursor/rules/vba-source-files.mdc`, `Version Control.accda.src/AGENTS.md`.

---

## 2026-06-17 — Conditional formatting blocks decoded to companion JSON

> **⚠ Partially superseded** (2026-06-20): Export format gating moved from `EFV_5_1_0`
> to `EFV_5_0_0` before v5 shipped. Decode/rebuild behavior and the
> `DecodeConditionalFormatting` option are unchanged. See "Fold unreleased 5.1.0 export
> gates into format 5.0.0" above.

**Trigger**: The per-control `ConditionalFormat` / `ConditionalFormat14` properties on form
and report controls export as opaque binary hex blocks. Any formatting change produces a
large, meaningless hex diff. We wanted the same clean-diff treatment we already give print
settings (`PrtMip`): strip the binary from source and store decoded, human-readable rules.

**Options explored**:
- *Raw hex in JSON* — store the hex blocks verbatim in the `.json`. Lossless and trivially
  byte-exact, but no more readable than leaving them inline. Rejected (defeats the purpose).
- *Hybrid (decode + raw hex fallback)* — decode for readability, keep raw hex for blocks we
  can't byte-rebuild. Safe but reintroduces hex noise. Rejected by the maintainer.
- *Full decode + rebuild (chosen)* — decode both blocks to a rule model, rebuild both on
  import. Cleanest JSON; relies on rebuild fidelity.

**Decision**: Full decode + rebuild via `clsConditionalFormat`. The **CF14** block is the
authoritative source and rebuilds **byte-for-byte** for every rule shape (expression,
field-value/between, focus, data bar), validated by formulas derived from the fixtures
(non-data-bar body length = `37 + 2·exprUnits`; data bar length = `P + 13`). The **legacy**
block is single-type and rebuilds byte-exact for single-rule controls (the common case);
its multi-rule per-rule layout is undocumented, so multi-rule legacy is rebuilt best-effort
(correct header/flags/colors/expressions). Both blocks are always emitted to stay consistent
with Access's precedence (legacy wins for overlapping rules). Gated behind export format
version `EFV_5_1_0` and the `DecodeConditionalFormatting` option (default on); import is
unconditional and backward compatible.

**What this rules out**: We do not store raw hex, so a control whose CF14 cannot be decoded
would lose its formatting on rebuild — acceptable because CF14 is the complete, verified
copy. Multi-rule legacy blocks are not guaranteed byte-identical to Access's original; if a
future Access version rejects our best-effort legacy layout, revisit by reverse-engineering
the multi-rule legacy per-rule descriptor bytes (offsets 40–55 in the Text11 fixture) or by
falling back to the hybrid raw-hex approach. Byte-exactness is enforced by
`modTestConditionalFormat` (CF14 all shapes; legacy single-rule shapes).

**Relevant files**: `modules/Core/clsConditionalFormat.cls` (new),
`modules/Core/clsSourceParser.cls` (capture/strip + `MergeConditionalFormat`),
`modules/Core/modLoadSaveText.bas` (`WriteConditionalFormatting` + pipeline),
`modules/Infrastructure/modConstants.bas` (`EFV_5_1_0`),
`modules/Infrastructure/clsOptions.cls` + `forms/frmVCSOptionsExport`
(`DecodeConditionalFormatting`), `modules/Tests/Core/modTestConditionalFormat.bas`,
`docs/access-conditional-format.md`.
## 2026-06-09 — Batch file metadata (date+size) for source property hashing

**Trigger**: Merge-build change detection on a large project (~7,300-file `queries`
folder) spent ~7.4s in "Get File Property Hash". `GetModifiedSourceFiles` already only
hashes files that have an index entry, so the cost was not redundant hashing — it was the
per-file `FSO.GetFile` (DateLastModified + Size) inside `GetSourceFilesPropertyHash`,
called once per source file/extension.

**Options explored**:
- **Per-file Win32 stat** (`FindFirstFileW` per file): measured ~400ms vs ~745ms FSO for
  3,659 files (~1.9x). Rejected — still per-file, ~12x slower than a batch scan.
- **Batch Win32 scan** (one directory walk capturing date+size): measured ~35ms (~22x
  faster than FSO). Chosen.
- **Capture date+size during the existing enumeration walk**: architecturally ideal (zero
  extra passes) but requires threading metadata through the cached, component-specific
  `GetFileList` (especially `clsDbQuery`). Deferred — a dedicated metadata walk is ~35ms
  and far less invasive.
- **Switch the stored property hash to Win32 UTC ticks** (DST-immune): rejected — changes
  the hash format, forcing a one-time content re-hash for every existing index.

**Decision**: Add `ScanFolderMetadata` (modFileWinAPI): one Win32 pass returning
`fullPath -> Array(date, size)` with case-insensitive keys, using the same `FileTimeToDate`
local conversion FSO uses. `GetModifiedSourceFiles` builds this map once per category and
passes it to `GetSourceFilesPropertyHash` via a new optional `dMeta` parameter; when `dMeta`
is omitted (the export-write path, where files are changing) it falls back to per-file FSO.
Verified empirically on 3,659 files: Win32 dates equal FSO dates (0 mismatches) and the
resulting property hashes are byte-identical (0 mismatches). Variant array elements must be
passed to `clsConcat.Add` wrapped in parentheses — `(varMeta(0)), (varMeta(1))` — to force
ByVal coercion into its `ByRef ... As String` parameters (a bare `varMeta(0)` raises
"ByRef argument type mismatch").

**What this rules out**: Using `dMeta` on the export path (files change during writes — the
cache would be stale). Assuming Win32==FSO date equality universally — a file modified
across a DST boundary on another machine may differ; the existing content-hash fallback in
`GetModifiedSourceFiles` keeps this safe (a one-time, self-healing re-hash) but not free.
Revisit the "capture during enumeration" single-walk design only if profiling shows the
extra metadata walk matters.

**Measured**: Real merge before/after on the ~7,300-file project: `Get File Property Hash`
7.44s -> 1.01s (the per-file `FSO.GetFile` is gone), at the cost of a single
`Scan Folder Metadata` batch walk of 3.07s — a net ~3.4s reduction on this run, with
`Compute SHA256` also dropping 2.84s -> 1.96s. The 3.07s metadata walk (date+size + Array
allocation over the full tree) is the remaining cost; the "capture during enumeration"
design would remove even that.

**Relevant files**: `modFileWinAPI.bas` (`ScanFolderMetadata`, `ScanMetadataRecurse`),
`modContainers.bas` (`GetSourceFilesPropertyHash`), `clsVCSIndex.cls`
(`GetModifiedSourceFiles`).

---

## 2026-06-09 — Win32 multi-pattern folder enumeration

**Trigger**: Folder scans used `FileSystemObject` `.Files`/`.SubFolders` iteration, whose
per-item COM overhead dominated "Get File List" (~20s in a merge log on a ~7,300-file
`queries` folder). Multi-format component types compounded it by scanning the same folder
once per extension (`clsDbQuery` scanned three times, plus ~3,659 per-file
`FSO.FileExists` calls to pair `.sql` with `.json`).

**Options explored**:
- **Push the extension into `FindFirstFileW`** (kernel filter): rejected as the primary
  mechanism — Win32 masks also match 8.3 short names (`*.sql` would hit a `.sqlite`) and
  `*.*` matches extension-less files, diverging from VBA `Like`. Kept VBA `Like` on file
  names for exact, 8.3-safe semantics.
- **N filtered calls vs one unfiltered scan + classify**: measured — two filtered calls
  62.8ms vs one scan+classify 33.1ms; three-pattern query case 40.7ms vs 35.3ms. One
  unfiltered scan wins (fewer directory traversals; per-entry marshaling dominates).
- **Full enumerate-once refactor of every multi-format component**: rejected — post-Win32
  the simple components (form/report/module/macro/tabledef) gain ~nothing; only
  `clsDbQuery` had meaningful cost.

**Decision**: Route `GetFilePathsInFolder`/`GetFilePathsInFolderRecursive` through
`ScanFolderContents` (Win32) and filter names with VBA `Like`. Extend both to accept a
`ParamArray` of patterns matched in a single pass; an empty `ParamArray` defaults to `*.*`
so existing single-pattern/no-pattern callers are unchanged. A `ParamArray` cannot be
forwarded directly to another procedure ("Invalid ParamArray use") — it is copied to a
`Variant` first. Collapse the simple components' `Set + MergeDictionary` pairs into one
multi-pattern call. Refactor `clsDbQuery` to one combined `.qdef/.bas/.sql/.json` scan with
an in-memory `.json` sibling lookup (a `TextCompare` set matching `FSO.FileExists`'s
case-insensitivity), eliminating the ~3,659 per-file `FSO.FileExists` calls (~2.3x faster;
identical file set verified, 0 mismatches).

**What this rules out**: Passing patterns to the Win32 mask (8.3/`*.*` semantics differ
from VBA `Like`). Two scanning idioms — all multi-extension components now share one
multi-pattern primitive. `clsDbQuery` keeps a bespoke post-classification block because its
legacy-priority/`.json`-pairing rules are irreducible.

**Measured**: Real merge before/after on the ~7,300-file project: `Get File List`
20.04s -> 0.02s and `Get File List Recursive` 3.46s -> 0.00s (both apples-to-apples — the
full tree is enumerated regardless of whether anything changed).

**Relevant files**: `modFileAccess.bas` (`GetFilePathsInFolder`,
`GetFilePathsInFolderRecursive`, `GetMatchingFilePaths`, `NormalizePatterns`),
`modFileWinAPI.bas` (`ScanFolderContents`), `clsDbForm/Report/Module/Macro/TableData/TableDef`
(`GetFileList`), `clsDbQuery.cls` (`GetFileList`).

---

## 2026-06-09 — Defer pre-merge database reopen until changes are confirmed

**Trigger**: Every merge build unconditionally closed and shift-reopened the current
database before scanning source files (to unload objects ahead of the destructive merge),
costing ~23s even when no source files had changed — the common "pull / switch-branch"
case.

**Options explored**:
- **Lightweight pre-scan to decide, then the existing flow**: rejected — re-scans on the
  change path (double the scan cost).
- **Scan first, reopen later, reuse the scan's component classes**: unsafe —
  `ShiftOpenDatabase` invalidates the cached database object references held by the scan's
  `IDbComponent` instances.
- **Scan first, reopen only when changes exist, rebuild component classes**: chosen.

**Decision**: Run the read-only scan + conflict resolution before the reopen. Only
close/shift-open when `dCategories.Count > 0` (real changes to merge). After reopening,
`RefreshContainerClasses` rebuilds the component instances against the reopened database
while preserving the already-computed file-path lists (plain strings, reopen-safe) and the
resolved conflicts. `ReleaseDbReferences` is called before the deferred reopen because the
scan now caches `CurrentDb` (the old pre-scan reopen did not need this). Conflict-detection
temp-exports run before the reopen — safe, since a normal export already temp-exports
without a reopen. Full builds are unchanged.

**What this rules out**: Reusing scan-built component-class instances across a reopen (they
hold stale object references and must be rebuilt). Does not address the post-merge
shared-mode reopen (~32s), which is dominated by Access re-opening a large database and is a
separate, still-open question.

**Measured**: On a ~7,300-file project, a no-change merge that previously logged
`Reopen DB before Merge` = 23.01s now skips it entirely (the deferred reopen never fires
when `dCategories.Count = 0`). Combined with the enumeration and metadata changes below,
total no-change merge time fell from 96.3s to 11.8s. The separate ~32s post-merge
shared-mode reopen did not occur on this run because nothing was imported — it still fires
on merges that import objects, confirming it as the next (Phase 4) target.

**Relevant files**: `modBuild.bas` (`Build`, `RefreshContainerClasses`).

---

## 2026-06-02 — Global suite hooks in VCS test runner

**Trigger**: Consumer projects need once-per-run setup/teardown (suite fixtures) around
`VCS.RunTests`, distinct from per-test `Class_Initialize` / `Class_Terminate`. Example
use: sweep leftover temp objects from a prior test run before executing the suite.

**Options explored**:
- **Module-qualified `Application.Run`** (`modTestAssert.GlobalTestSetup`) — rejected;
  fails for `Option Private Module` and conflicts with existing cross-project run pattern
  (see 2026-05-08 entry).
- **Catch error 2517 around `Application.Run`** for missing procedures — rejected in
  favor of **`GlobalProcExists`** pre-check (same as module test discovery).
- **Run hooks when zero tests selected** — rejected; standard @BeforeAll / pytest session
  semantics skip fixtures when nothing is selected.
- **Include hook status in `TestResults_*.json`** — deferred; teardown runs after JSON is
  written, so a JSON block would be asymmetric. v1 logs hook errors to the console only.

**Decision**: Add optional parameterless `GlobalTestSetup` / `GlobalTestTeardown` public
subs in the target project's `modTestAssert`. `ExecuteTests` calls setup immediately before
`RunAll`/`RunSelected` and teardown after `GetResultsAsJson`, only when ≥1 test will run.
Missing procedures skip silently. Hook errors use `Log.Add` (never `Log.Error`) and do not
fail the run. Fresh `InstallTestAssertModule` installs include empty stubs with inline
comments; existing projects are not auto-upgraded.

**What this rules out**: Auto-migrating existing `modTestAssert` modules to add hook stubs.
Global hooks on `RunFailed` (not routed through `ExecuteTests` today). Parameterized hook
signatures in v1.

**Relevant files**: `clsTestRunner.cls` (`InvokeGlobalTestSetup`, `InvokeGlobalTestTeardown`,
`InvokeOptionalGlobalHook`), `clsVersionControl.cls` (`ExecuteTests`, `InstallTestAssertModule`).

---

## 2026-05-29 — Layered `.env` resolution via `APP_ENV`

**Trigger**: Projects with live/offline (or dev/staging/production) backends need
the same exported source tree to target different ODBC servers without editing
connection strings in source files. A prototype in a consumer project used a
selector `.env` plus `.env.{environment}` files; the add-in previously resolved
all `env:conn_*` references from a single flat `.env`.

**Options explored**:
- **Runtime public API** (`VCS.GetEnv`) for ADODB code — rejected for this change;
  scope limited to build/import resolution only.
- **Replace semantics** (environment file fully replaces base) — rejected; layered
  merge lets shared keys live in base `.env` with environment-specific overrides.
- **Configurable selector key in `vcs-options.json`** — rejected; fixed `APP_ENV`
  matches common dotenv-flow conventions and keeps config surface small.

**Decision**: At import/build, merge `.env` files in dotenv-flow order: `.env` →
`.env.local` → `.env.{APP_ENV}` → `.env.{APP_ENV}.local`. `APP_ENV` comes from
the OS environment first, then the merged base level. Export writes remain on base
`.env` only; reads use the merged config. No export-format-version gate — exported
source content is unchanged.

**What this rules out**: Automatic relinking when `APP_ENV` changes without a
rebuild/merge. Runtime VBA in the user's database still needs its own `.env` reader
if it opens ADODB connections outside the add-in's import path.

**Secrets safety**: The auto-`.gitignore` only excluded `*.env`, which (by
gitignore glob rules) does not match `.env.local`, `.env.<APP_ENV>`, or
`.env.<APP_ENV>.local`. Extended the default template and `EnsureGitignore` logic
to also exclude `.env.*` with a `!.env*.example` negation so layered credential
files are ignored while `*.example` templates stay committed.

**Relevant files**: `clsDotEnv.cls` (`LoadFromFileIfExists`, merge flag),
`modConnect.bas` (`BuildResolvedEnv`, split read/write caches),
`modVCSUtility.bas` (gitignore `.env.*` / `!.env*.example`), `.gitignore.default`,
`modTestConnect.bas`, `Version Control.accda.src/AGENTS.md`, `Wiki/Connections.md`

---

## 2026-05-29 — Module full-build import: sync VBE with AllModules before index/metadata

**Trigger**: Two related bugs during full builds of ~135 modules. (1) After
`VBComponents.Remove` / `.Import`, `CurrentProject.AllModules` can lag behind
VBE, causing intermittent error 2467 on the immediate `AllModules(strName)`
lookup. This produced critical "Imported module not found after import" failures
and skipped `ImportObjectMetadata` / `VCSIndex.Update` for the affected module.
(2) Full builds reuse a single `clsDbModule` instance across all files, but
`m_strSourceFile` was cached from the prior import, causing `VCSIndex.Update` to
index the new module under the previous file's path. Both bugs surfaced as false
export conflict prompts after a full build. Commit `2e3b6abd`.

**Options explored**:
- **Immediate `AllModules(strName)` with no retry (status quo)** — simple; fails
  intermittently when 135 modules are imported in a tight loop and Access's
  navigation catalog lags behind VBE.
- **Batch verification after entire module category** — fewer `DoEvents` in the
  hot loop; rejected because `ImportObjectMetadata` needs the DAO document per
  file, and `VCSIndex.Update` reads `m_Module.DateModified`. Deferring these
  requires either passing explicit per-file keys (API change) or accepting
  wrong/stale index entries during the loop. Fail-late also wastes work when an
  early import is broken.
- **`Sleep` between retries** — does not pump Access's message queue the way
  `DoEvents` does; same finding as the worker-queue decision (2026-04-03).
- **`DoEvents` after save + `AllModules` retry loop (chosen)** — per-module
  message-pump cost so each import leaves metadata and index in a correct state
  before the next file starts.

**Decision**: After `DoCmd.Save`, call `DoEvents` once to let Access publish the
module. Resolve `m_Module` via `GetAccessModuleObject` (up to 3 tries with
`DoEvents` between failures); fail critical if still missing. Clear
`m_strSourceFile` at the top of each `Import` call so the shared instance never
indexes under a stale path. Add `VbeModuleExists` check (VBE-side only, no pump)
inside `LoadVbeModuleFromFile` for early detection of bad `Attribute VB_Name`.

**Performance**: On the 135-module add-in build, `Import VBE Module` stays flat
at ~0.37–0.40 s total. The `Modules` category rises from ~3.1 s (old code) to
~3.3–3.5 s typical, with occasional spikes to ~4.5 s. The extra cost sits in
`DoEvents` and `AllModules` retries between `Import VBE Module` and
`VCSIndex.Update` — work not captured under any named `Perf.OperationStart`.

**What this rules out**: Deferring `AllModules` verification to end-of-category
without also deferring metadata/index or supplying explicit per-file keys to
`VCSIndex.Update`. Removing all `DoEvents` without an alternative queue pump.
Treating `VBComponents` existence alone as proof the module is ready for DAO
document property writes.

**Revisit if**: Access offers a reliable "module published to navigation
container" event or callback; or `VCSIndex.Update` is refactored to accept an
explicit file key and timestamp so it no longer depends on `m_Module` /
`DateModified` during import. *(Superseded for full builds by 2026-06-23
two-pass `ImportFast`/`FinalizeImports`; per-file path retained for merge and
single-object import.)*

**Relevant files**: `clsDbModule.cls` (`IDbComponent_Import`, `GetAccessModuleObject`,
`VbeModuleExists`, `LoadVbeModuleFromFile`), `modTestConflicts.bas`
(`TestModuleImport_IndexesEachFileOnSharedInstance`), `clsVCSIndex.cls` (module
`Update` / `VBAProjectDate`). Related: "VBProject.Saved + DateModified fast path"
(2026-05-05) for `AllModules` semantics on the export side.

---

## 2026-05-29 — Test runs: dedicated eotTestRun operation, TestRun_ log path, loggedErrors in JSON

**Trigger**: Test runs already wrapped `Operation.Begin`, but used `eotOther` with a hard-coded `TestRun_*.log` alternate path in `ExecuteTests`. Import/build failures logged via `Log.Error` during tests (e.g. `clsDbModule.Import` critical errors) appeared in the console and log file but not in `TestResults_*.json` — agents and MCP tooling parse JSON first and only saw a generic "Logged error(s) during test" message.

**Options explored**:
- **`Test_` log prefix with `eotTest`** — shorter filename, but too generic; breaks continuity with existing `TestRun_*.log` files and weakens the pairing with `TestResults_*.json`.
- **`TestRun_` prefix via hard-coded alternate path (status quo)** — worked, but bypassed `LogFilePath` and did not exercise the same save/cleanup path as Export/Build.
- **`eotTestRun` + `LogFilePath` base `TestRun` (chosen)** — dedicated operation type maps to `TestRun_{OperationId}.log` through the normal `Log.SaveFile` path, including `CleanupOldLogs`. Enum name and log prefix align for clarity.

**Decision**: Add `eotTestRun = 4` to `eOperationType` and move `eotOther = 9` to the end of the enum (values 5–8 reserved for future dedicated operation types). `ExecuteTests` calls `Operation.Begin(eotTestRun)`, sets `Log.Active = True` and `InteractionMode = eimSilent`, clears an error journal at run start, and saves via `Log.SaveFile` (no alternate path). `clsLog` maintains an error journal on each `Log.Error` call; `clsTestRunner` snapshots the journal per test and exports a `loggedErrors` array (level, message, source, errNumber, errDescription) in `TestResults_*.json`, with `errorMessage` set to the first logged error text.

**What this rules out**: Using `eotOther` for the main test suite (round-trip and other harnesses may still use `eotOther` with custom prefixes). Relying on agents to open `TestRun_*.log` for operation-level failure details when JSON is available.

**Relevant files**: `modConstants.bas` (`eotTestRun`), `clsLog.cls` (`LogFilePath`, error journal), `clsVersionControl.cls` (`ExecuteTests`), `clsTestRunner.cls` (`AttachLoggedErrors`, `GetResultsAsJson`).

---

## 2026-05-21 — Rich Text console truncation: boundary-aware HTML truncation

**Trigger**: Console output in `frmVCSMain.txtLog` was visibly truncated — the test summary and final lines never appeared on screen, even though `txtLog.Value` contained the complete HTML. The problem was intermittent and sometimes occurred with minimal content. Previous attempts (reducing buffer from 10K to 8K, replacing `&nbsp;` entities with `ChrW$(160)`) did not resolve it.

**Options explored**:
- **Reduce buffer limit (10K → 8K → smaller)** — tried and reverted. Empirical probing showed the Rich Text control renders at least 256KB of well-formed HTML without issue. The character limit was a red herring.
- **Replace `&nbsp;` with `ChrW$(160)` to shrink HTML source** — tried and reverted. Reduced source size ~5× per space, but had no effect on rendering because the actual limit was not size-related.
- **Add `DoEvents` after `Echo True`** — probed empirically. Made no difference; the control updates correctly without it.
- **Trim `RightStr` output to the first `<br>` boundary (chosen)** — root cause: `m_Console.RightStr(N)` cuts at an arbitrary character position, often splitting an HTML tag (e.g., producing `olor=gray>Text...</font>`). The Access Rich Text control accepts malformed HTML into `.Value` but its renderer silently stops partway through, truncating the visual display. Trimming to the first `<br>` after the cut ensures the HTML always starts at a clean line boundary.

**Decision**: Added `ConsoleHtml()` private function in `clsLog.cls` that (1) fetches the last 64K characters via `RightStr`, (2) if truncation occurred, finds the first `<br>` and discards everything before it. Buffer limit raised from 8K to 64K since the control has no meaningful rendering limit for well-formed HTML. Also added `ClampInt()` to cap `.SelStart` at 32000 (the property is Integer-typed and overflows above 32,767).

**What this rules out**: Any future assumption that the Rich Text control has a ~10K rendering capacity. It does not — the limit is at least 256K. What *does* break rendering is malformed HTML at the start of the assigned string. If `ConsoleHtml` is ever bypassed or a different truncation method is used, it must guarantee valid HTML at the start. Revisit if Access gains a different Rich Text implementation or if performance degrades with very large console buffers.

**Relevant files**: `Version Control.accda.src/modules/Infrastructure/clsLog.cls` (`ConsoleHtml`, `ClampInt`, `Flush`, `ApplyPendingIncrements`).

---

## 2026-05-14 — Keep SELECT/UPDATE modifiers (DISTINCT, TOP N) on the same line

**Trigger**: After switching to the MSysQueries-based `.sql` + `.json` export format, users noticed that `SELECT TOP N` was being split across two lines: `SELECT` alone on the first line, then `TOP N` indented with the first column on the next. The same issue affected `UPDATE DISTINCTROW`. The formatter had always done this, but it was only visible now that formatted `.sql` files became the primary source.

**Options explored**:
- **Add TOP/DISTINCT/DISTINCTROW to `cstrReservedToplevel`** — rejected: these are not clause-level keywords. Making them top-level would force line breaks *before* them too, creating `SELECT\nTOP\n  ID` rather than `SELECT TOP N\n  ID`.
- **Suppress `blnNewline` for SELECT and re-enable after modifiers** — rejected: requires threading state through several iterations of the main loop; fragile.
- **Post-emit look-ahead after SELECT and UPDATE (chosen)** — after emitting `SELECT`, a small loop peeks ahead and consumes `DISTINCT`/`DISTINCTROW`, then `TOP` + number + optional `PERCENT`, emitting them inline. After emitting `UPDATE`, a simpler check consumes `DISTINCTROW`. `blnNewline` (already set by the `ttReservedTopLevel` block) takes effect for the next token. `DELETE DISTINCTROW` was already correct — `DELETE` alone is not top-level (only `DELETE FROM` is), so no newline is forced.

**Decision**: Inline modifier consumption after SELECT and UPDATE. Matches the convention used by SQLFluff (rule LT10: "SELECT clause modifiers must be on the same line as SELECT") and the expected output of Poor Man's T-SQL Formatter. No export format version gate — the formatter is stateless and the change is cosmetic whitespace only.

**What this rules out**: Formatters that place each modifier on its own line (`SELECT\n  DISTINCT\n  TOP 3\n  Column`). If someone wanted that style, the look-ahead would need to be made conditional. Revisit if a formatting-options system is ever added to `clsSqlFormatter`.

**Relevant files**: `clsSqlFormatter.cls` (modifier look-ahead in `FormatSQL`, four new `SelfTest` cases), `Testing/Fixtures/queries/` (updated `.sql` and `.qdef` baselines for TOP, DISTINCT, and DISTINCTROW fixtures).

---

## 2026-05-08 — Class-based test suites via TestClassFactory dispatcher

**Trigger**: The original test runner (2026-05-08) explicitly ruled out class-based test suites. As the test suite grew, the limitation became painful: standard module tests pollute the global public namespace, and there is no built-in setup/teardown mechanism. Class modules naturally solve both problems via `Class_Initialize`/`Class_Terminate` and encapsulated scope.

**Options explored**:
- **Temporary factory module (inject/remove per run)** — rejected: VBE `CodeModule` manipulations are expensive, risk recompile between tests, and leave orphan modules on crash.
- **One factory function per class (N separate `Public Function` declarations)** — rejected: clutters the module, harder to read, generates more VBE code churn during reconciliation.
- **`PredeclaredId = True` on test classes** — rejected: requires modifying every test class's attributes, non-standard for user code, confuses developers unfamiliar with default instances.
- **Single `TestClassFactory` dispatcher with `Select Case` (chosen)** — one function in `modTestAssert`, `Select Case` entries auto-maintained by the runner. Minimal code surface, single `GlobalProcExists` check, and the pattern is self-documenting.

**Decision**: `clsTestRunner.Scan` now discovers class modules (alongside standard modules) using the same `@Folder("Tests")` or `*Test*` naming rules. After discovery, `SyncFactoryEntries` reconciles the `Select Case` block inside `TestClassFactory` (in `modTestAssert`) — only writing if entries are stale. At execution time, `RunSelected` calls `Application.Run(BuildRunCmd("TestClassFactory"), className)` to get a fresh instance per test method, then `CallByName obj, procName, VbMethod`. `Set obj = Nothing` triggers `Class_Terminate` (teardown). A compile check (`acCmdCompileAllModules` + `Application.IsCompiled`) gates test execution — the run aborts if the project has compile errors.

**What this rules out**: Shared state across test methods within a class (each method gets its own instance). Parameterized test classes (the factory takes only a class name string). Custom setup/teardown method names — only `Class_Initialize` and `Class_Terminate` serve this role. The `TestClassFactory` function must remain in a non-`Option Private Module` standard module for `GlobalProcExists` to work.

**Relevant files**: `clsTestRunner.cls` (`IsTestModule`, `ScanModuleForTests`, `SyncFactoryEntries`, `RunSelected`), `modTestAssert.bas` (`TestClassFactory` template), `.cursor/rules/testing.mdc` (agent documentation).

---

## 2026-05-08 — BreakOnError: read live from Options instead of caching

**Trigger**: `clsTestRunner.RunSelected` sets `Options.BreakOnError = False` during test execution so errors don't break into the debugger. But `DebugMode()` in `modErrorHandling` was reading a stale cached copy (`this.blnBreakOnError`) that was only updated by the `ConfigureErrorHandling` push-function. Setting the public field had no effect on `DebugMode()`.

**Options explored**:
- **Make `BreakOnError` a `Property Let` that calls `ConfigureErrorHandling`** — rejected: every other option in `clsOptions` is a plain public field. Adding a setter to one field breaks the pattern and creates a maintenance trap (future option fields would need the same treatment).
- **Have `DebugMode()` read `Options.BreakOnError` directly (chosen)** — guarded by `OptionsLoaded` (already exists in `modObjects`) to prevent circular initialization during the Options load sequence.

**Decision**: `DebugMode()` and `LogUnhandledErrors` now read `Options.BreakOnError` directly. The `ConfigureErrorHandling` sub and the `blnBreakOnError` UDT field in `modErrorHandling` are deleted. During early initialization (before Options loads), `OptionsLoaded` returns `False` and `DebugMode` returns `False` — the same safe default the cache used.

**What this rules out**: The push-cache pattern for error handling configuration. Any future setting that `modErrorHandling` needs must either be read through `Options` with an `OptionsLoaded` guard, or use a different mechanism. Revisit if `OptionsLoaded` ever becomes unreliable or if `modErrorHandling` needs settings before Options initialization.

**Relevant files**: `modErrorHandling.bas` (live read, removed cache), `modObjects.bas` (removed `ConfigureErrorHandling` calls), `clsOptions.cls` (removed `ConfigureErrorHandling` call in `LoadProjectOptions`).

---

## 2026-05-08 — Cross-project test execution: unqualified Application.Run with GlobalProcExists guard

**Trigger**: `Application.Run "ModuleName.ProcName"` fails with Error 28 (out of stack space) or Error 2517 when the target module uses `Option Private Module`. Module-qualified names also don't resolve correctly across library references. The test runner needs to call test procedures in the user's `CurrentVBProject` from the add-in.

**Options explored**:
- **Module-qualified `Application.Run` calls** (`"modTests.TestFoo"`) — rejected: fails for `Option Private Module`, and the qualification is unnecessary when procedure names are globally unique (which they should be in a well-structured project).
- **Unqualified `Application.Run` with no pre-check** — rejected: produces confusing stack overflow errors when a procedure is uncallable.
- **Unqualified `Application.Run` with `GlobalProcExists` pre-check (chosen)** — `GlobalProcExists` (already exists in the codebase) verifies the procedure is callable before attempting `Application.Run`. Uncallable procedures are logged as SKIP with a clear message.

**Decision**: `clsTestRunner.RunSelected` passes only the procedure name (no module qualifier) to `Application.Run`. Before each call, `GlobalProcExists` checks callability. Procedures in `Option Private Module` are skipped and logged as SKIP rather than producing runtime errors.

**What this rules out**: Testing procedures inside `Option Private Module` via the add-in toolbar. Those modules can still be tested directly via F5 or Immediate Window (where `TestAssert` falls back to `Debug.Assert`). Revisit if Access adds a way to call private-module procedures cross-project.

**Relevant files**: `clsTestRunner.cls` (`RunSelected`, `BuildRunCmd`).

---

## 2026-05-08 — Test UI: reuse frmVCSMain console instead of a dedicated form

**Trigger**: The test runner needs to display real-time progress and results. A dedicated `frmVCSTestRunner` was prototyped but added unnecessary complexity — another form to maintain, no consistency with existing UI patterns.

**Options explored**:
- **New native Access form (`frmVCSTestRunner`)** — tried and reverted. Added a form file, a class module, and UI layout work for something that duplicated what `frmVCSMain` already does.
- **EdgeBrowserControl web UI** — deferred as Plan A for a future enhancement. Requires Access versions that ship the control and adds HTML/JS asset management complexity.
- **Stream results through `frmVCSMain`'s rich-text console (chosen)** — matches the existing query validation pattern. Right-aligned color-coded status (green PASS, red FAIL/ERROR, gray EMPTY, orange SKIP) using the `Log.Add` HTML formatting already available.

**Decision**: `clsVersionControl.RunTests` opens `frmVCSMain` via `PrepareTestConsole` (sets `InsideWidth=12000`, `InsideHeight=9000`), streams test lines via `Log.Add`, and finalizes with `FinalizeTestConsole`. No new form files. Individual test results are formatted as `TestName` + right-aligned `STATUS` with color coding.

**What this rules out**: Interactive test selection, tree-view grouping, or re-run buttons in the current UI. These would require a dedicated form or the EdgeBrowserControl plan. Revisit when the web UI plan is implemented.

**Relevant files**: `clsVersionControl.cls` (`RunTests`, `PrepareTestConsole`, `FinalizeTestConsole`), `clsTestRunner.cls` (`LogTestResult`).

---

## 2026-05-08 — Test discovery: any parameterless Public Sub in a test module

**Trigger**: Designing the test discovery rules for the TestAssert framework. Needed a convention that requires zero boilerplate but doesn't accidentally pick up non-test code.

**Options explored**:
- **Rubberduck-style `'@Test` annotations** — rejected: adds a dependency on a specific comment convention that most Access developers don't use. Scanning for magic comments is fragile.
- **`Test_` prefix requirement** — rejected: too restrictive. Existing tests in the project use `Test` without underscore, and the prefix convention varies across developers.
- **`*Test*` in procedure name** — rejected as too narrow: would miss legitimate test subs like `VerifyHashConsistency` or `CheckEncodingRoundtrip`.
- **Any parameterless `Public Sub` in a designated test module (chosen)** — the module-level designation (`@Folder("Tests")` or `*Test*` in the module name) scopes what counts as a test module. Within a test module, every parameterless `Public Sub` is a test. Simple, zero-boilerplate, matches how most VBA developers already write tests.

**Decision**: Test module identification: standard modules only (not class modules) with either `@Folder("Tests")` annotation in the first 30 lines or `*Test*` anywhere in the module name. Test procedure identification: any `Public Sub` with zero parameters. No tags, no naming conventions beyond the module-level designation.

**What this rules out**: Parameterized tests, class-based test suites, and selective test tagging within a module. Helper subs in test modules must be `Private` or take parameters to avoid being treated as tests. Revisit if parameterized test support is needed (would require a `'@TestCase` annotation or similar).

**Relevant files**: `clsTestRunner.cls` (`IsTestModule`, `ScanModuleForTests`, `ProjectHasFolderAnnotations`).

---

## 2026-05-08 — TestAssert framework: dual execution model with Application.Run callback

**Trigger**: The add-in needed a built-in test runner. Existing tests used `Debug.Assert` which provides no result capture, no progress display, and halts execution on failure in break mode.

**Options explored**:
- **Full Rubberduck-style framework** (class-based, annotation-driven) — rejected: too heavy for the Access VBA ecosystem, requires significant boilerplate, and Rubberduck itself provides this if users want it.
- **Add-in-only test execution** (tests only work when the add-in is loaded) — rejected: developers need to run individual tests via F5 or Immediate Window during active development without the add-in toolbar.
- **`TestAssert` as a thin wrapper with dual paths (chosen)** — `modTestAssert.bas` is installed in the user's project. `TestAssert condition` calls `Application.Run` to notify the add-in's `HandleTestAssertion` function. If the add-in isn't loaded or the runner isn't active, it falls back to `Debug.Assert condition`. Same test code works in both contexts.

**Decision**: `modTestAssert.bas` ships as a standalone file, offered for installation on first "Run Tests" click (similar to the letter casing template). It resolves the add-in path by scanning `VBE.VBProjects` for the `MSAccessVCS` project name, with a `CurrentProject` fallback for self-testing. `HandleTestAssertion` lives in `modAPI.bas` (not a class) so it's callable via `Application.Run`. The runner (`clsTestRunner`) is a singleton accessed via `modObjects.TestRunner`. Results are persisted as JSON in the project's `logs/` folder.

**What this rules out**: Assertion variants beyond pass/fail (no `AssertEqual`, `AssertThrows`, etc. in v1). The `TestAssert` sub takes a boolean condition and an optional context variant — richer assertion types would require additional subs in `modTestAssert.bas`. Also rules out automatic `modTestAssert` updates — once installed, the user's copy is independent. Revisit assertion API if users request structured matchers.

**Relevant files**: `modTestAssert.bas` (user-installed helper), `modAPI.bas` (`HandleTestAssertion`), `clsTestRunner.cls` (singleton engine), `clsVersionControl.cls` (`RunTests`, `RunFilteredTests`, `InstallTestAssertModule`, `MigrateDebugAssert`), `modObjects.bas` (`TestRunner` accessor), `Ribbon/Ribbon.xml` (`btnRunFilteredTests` in Tools group, `MacroPlay` icon).

---

## 2026-05-05 — Multi-file conflict detection: per-file diff with per-component resolution

**Trigger**: On first export (empty index) all table definitions showed as export conflicts even though the XML files were byte-identical. Root cause: `SourceMatches` compared all `FileExtensions` across source and temp directories, but companion files (`.json` metadata, `.sql` DDL) were never produced during temp/alternate-path exports — they were gated behind `If strAlternatePath = vbNullString`. The file-count mismatch caused every multi-file component to report a false conflict.

**Options explored**:
- **Relax `SourceMatches` to intersection-based comparison** (only compare files present in both directories): rejected. This masks the root cause — companion files simply aren't being exported. It also prevents the conflict dialog from ever diffing companion files, since they don't exist in the temp folder.
- **Export all companion files during temp exports** (chosen): Fix the component `Export` methods to produce all files regardless of `strAlternatePath`. Destructive operations (stale file deletion, format switching) remain gated to real exports. This makes `SourceMatches` correct again and provides temp copies of every file for per-file diffs.

**Decision**: Component `Export` methods (`clsDbTableDef`, `clsDbModule`) now produce companion files during alternate-path exports. `SourceMatches` was replaced with `GetDifferingFiles` which returns a `Collection` of file names that differ (or `Nothing` when all match). The conflict dialog writes one `tblConflicts` row per differing file (all sharing the same `ItemKey`), and a `cboResolution_AfterUpdate` handler propagates resolution to all sibling rows — keeping resolution atomic at the component level while allowing per-file diffs. Forms and reports are unaffected because their `FileExtensions` do not include `json` or `svg`.

**What this rules out**: Per-file resolution (skip one file but overwrite another within the same component) — export/import operates atomically on whole components, so partial resolution would require fundamentally different import/export logic. If per-file resolution is ever needed, it would require splitting components into independently importable sub-units. Adding new file extensions to a component's `FileExtensions` now requires ensuring those files are also produced during alternate-path exports, or the strict count comparison in `GetDifferingFiles` will flag false conflicts.

**Relevant files**: `clsDbTableDef.cls` (companion file export), `clsDbModule.cls` (metadata export), `clsVCSIndex.cls` (`GetDifferingFiles`, `GetExportConflictFiles`, `IsMergeConflict`), `clsConflictItem.cls` (`DifferingFiles` property), `clsConflicts.cls` (multi-row `SaveToTable`), `frmVCSConflictList.cls` (resolution propagation), `frmVCSConflictList.form` (`AfterUpdate` event wiring).

---

## 2026-05-04 — Gate deterministic query export behind `UseDeterministicQueryExport` option

**Trigger**: The new MSysQueries-based export path (`clsQueryComposer` + `clsDbQuery.ExportNewFormat`) is a large architectural change covering SQL reconstruction, Design View vs SQL View arbitration, `LvExtra`/`LvProp` layout handling, and qdef generation. Despite a 40+ fixture regression corpus, undiscovered edge cases are likely in real-world databases with thousands of queries. Users need a fallback to continue development while parser bugs are resolved.

**Options explored**:
- **Always-on behind export-format-version only** (`EFV_5_0_0`): rejected. Format version gating prevents format-version downgrade but offers no escape hatch if the new code path has a runtime bug on a specific query. The user is stuck until a fix ships.
- **Per-query toggle** (e.g. a list of query names that use legacy export): rejected. Too granular — the user would have to identify each failing query individually, and the option surface is unwieldy.
- **Ship as beta/preview flag** (hidden, not in the UI): rejected. No existing flag infrastructure in the add-in; a hidden option is easily forgotten and hard to document.
- **User-visible boolean option** (chosen): `UseDeterministicQueryExport` in `clsOptions`, default `True`, exposed as a checkbox on the Export Options form. When `False`, `clsDbQuery` routes to the legacy `SaveAsText`-based `.qdef` export. Simple, discoverable, one click to revert.

**Decision**: Add `UseDeterministicQueryExport` as a user-visible boolean option (default `True`). The export path in `clsDbQuery` checks this option: when enabled, queries export as `.sql` + `.json` via `clsQueryComposer`; when disabled, queries export as `.qdef` via `SaveAsText`. Import remains extension-based regardless of this setting — `.sql` files always use the new import path, `.qdef` files always use the legacy path. This decouples the export rollout from the import path, ensuring users can always build from source regardless of which format was used to export.

**What this rules out**: Removing the option without a follow-up decision — the escape hatch is a shipped contract until the new path has proven stable across a broad user base. Making the option affect import behavior — import must always handle both formats since a repository may contain a mix of `.sql` and `.qdef` files from different contributors or time periods.

**Relevant files**: `Version Control.accda.src/modules/Infrastructure/clsOptions.cls`, `Version Control.accda.src/forms/frmVCSOptionsExport.cls`, `Version Control.accda.src/forms/frmVCSOptionsExport.form`, `Version Control.accda.src/vcs-options.json`, `Version Control.accda.src/modules/Components/clsDbQuery.cls` (gate check in export path).

---

## 2026-05-01 — Pass-through queries bypass SQL formatter and composer entirely

**Trigger**: Exporting a database containing `dbQSQLPassThrough` queries crashed `clsSqlFormatter` with "Unable to parse SQL after position N" — the formatter's tokenizer is designed for Access SQL syntax and cannot handle T-SQL, PL/SQL, or other server-side dialects that pass-through queries may contain.

**Options explored**:
- **Teach the formatter about T-SQL/PL-SQL**: rejected. Scope explosion — every server dialect has its own syntax, reserved words, quoting rules, and comment styles. The formatter would become a multi-dialect parser with no clear boundary.
- **Format only the SELECT-like subset** (heuristic detection of "looks like Access SQL"): rejected. Fragile — any heuristic would produce false positives on server SQL that happens to resemble Access SQL, silently corrupting the stored query text.
- **Detect and bypass entirely** (chosen): Check `QueryDef.Type` for `dbQSQLPassThrough` (112) and `dbQSPTBulk` (144) early in `clsDbQuery.ExportNewFormat`. Store the SQL verbatim — no formatting, no decomposition through `clsQueryComposer`, no MSysQueries reconstruction. The `Connect` string is captured via `QueryDef.Connect` with an `ODBC;` placeholder fallback.

**Decision**: Pass-through query types are detected at the top of the export path and routed to a verbatim-storage branch. SQL is written as-is to the `.sql` file. The `.json` metadata includes `QueryType` so the import path knows to skip `LoadFromText` qdef generation and use `CreateQueryDef` directly. `clsSqlFormatter` and `clsQueryComposer` are never invoked for these query types.

**What this rules out**: Future attempts to "fix" the formatter or composer for non-Access SQL dialects — pass-through SQL must always be stored verbatim. If a future need arises to pretty-print server SQL (e.g. for diff readability), it must be a separate, opt-in formatter that does not share code paths with the Access SQL formatter.

**Relevant files**: `Version Control.accda.src/modules/Components/clsDbQuery.cls` (type detection and verbatim export/import), `Version Control.accda.src/modules/Utility/clsQueryComposer.cls` (`ConnectString` property for SQL View qdef emission), `Testing/Fixtures/queries/passthrough/qryPassThroughNoConnect.*` (regression fixture).

---

## 2026-04-29 — Replace `Dir()` and FSO folder scanning with Win32 API (`FindFirstFileW`)

> **⚠ Partially superseded** (2026-06-09): This entry's "rules out" note that "FSO remains
> acceptable for targeted single-file operations (`FSO.FileExists`, `FSO.GetFile`) where COM
> overhead is negligible" holds for one-off calls but not for hot per-file loops.
> `GetFilePathsInFolder(Recursive)` now use `ScanFolderContents` (the recursive-glob gap
> this entry left open is closed), and per-file `FSO.GetFile` date/size lookups during
> change detection were replaced by a single batched `ScanFolderMetadata` pass. See "Win32
> multi-pattern folder enumeration" and "Batch file metadata (date+size) for source property
> hashing" above.

**Trigger**: Export profiling on a large database (`db-sec`, ~3,500 components) showed orphan scanning and file-extension migration checks dominated the "no changes" export time. Two separate problems: (1) `Dir()` does not support Unicode filenames — it silently skips or fails on paths containing non-ASCII characters, which Access databases frequently contain (accented characters, CJK object names). (2) `Scripting.FileSystemObject` (FSO) folder enumeration is correct but slow — each `oFolder.Files` / `oFolder.SubFolders` iteration creates COM proxy objects with per-item round-trip overhead.

**Options explored**:
- **FSO-only** (drop `Dir()`, keep FSO for all scanning): rejected. Correct for Unicode but too slow — FSO `GetFolder().Files` on a 500-file export folder added measurable latency per component type during orphan cleanup.
- **`Dir()` with Unicode workarounds** (e.g. short 8.3 names, `Dir$` variants): rejected. Fragile — 8.3 name generation is optional on NTFS and disabled by default on modern Windows; `Dir$` has the same Unicode limitation as `Dir`.
- **Shell out to PowerShell** (`Get-ChildItem`): rejected. Process startup overhead per invocation; unsuitable for hot paths called hundreds of times per export.
- **Win32 API via `FindFirstFileW` / `FindNextFileW`** (chosen): Single kernel call enumerates all entries in one pass with full Unicode support. Wrapped in `modFileWinAPI.bas` as `ScanFolderContents` (returns files + subfolders in one pass) and `FilePatternExists` (O(1) early-exit check for wildcard matches).

**Decision**: Blanket prohibition on `Dir()` in all add-in code — documented in `AGENTS.md` under "File System Operations". All folder scanning converted to Win32 API wrappers in `modFileWinAPI.bas`. `modOrphaned.ScanFolderForOrphans` now takes a `String` path instead of a `Scripting.Folder` object. `modFileAccess.ClearFilesByExtension` and `modSourceUpgrade.RenameFilesInFolder` use `FilePatternExists` for early exit before attempting FSO operations.

**What this rules out**: Any future use of `Dir()` without an explicit follow-up decision overriding this one. New file-scanning code must use the `modFileWinAPI` wrappers or FSO (for cases where the API wrappers don't yet cover the need, e.g. recursive glob patterns). FSO remains acceptable for targeted single-file operations (`FSO.FileExists`, `FSO.GetFile`) where COM overhead is negligible.

**Relevant files**: `Version Control.accda.src/modules/Utility/modFileWinAPI.bas` (new wrappers), `Version Control.accda.src/modules/Core/modOrphaned.bas` (orphan scan converted), `Version Control.accda.src/modules/Utility/modFileAccess.bas` (`ClearFilesByExtension` converted), `Version Control.accda.src/modules/Core/modSourceUpgrade.bas` (`FilePatternExists` guard), `AGENTS.md` ("File System Operations" section).

---

## 2026-04-29 — Reconstruct stored query attributes instead of normalizing them away

**Trigger**: The SEC `ValidateQuerySqlBuilder` run flagged 398 queries for
review. Most were harmless formatting or commutative join predicates, but a
small set showed stored MSysQueries attributes being dropped: external
make-table targets (`Attribute 1 Name2`), parameter declarations
(`Attribute 2 Name1/Flag`), action-query `DISTINCTROW` (`Attribute 3 bit 8`),
and UNION `ORDER BY` rows (`Attribute 11`).

**Options explored**:
- **Broaden the validation canonicalizer**: rejected for these cases. It would
  hide real reconstruction loss, especially external destination paths and
  UNION ordering.
- **Treat action-query `DISTINCTROW` as semantic noise**: rejected. It may be
  benign for many queries, but Access stores it explicitly and users expect
  export/import fidelity.
- **Preserve the stored attributes in `clsQueryComposer`**: chosen. The builder
  already reads the MSysQueries row stream; the missing behavior belongs at
  reconstruction and parsing boundaries, not in downstream formatter rules.

**Decision**: `clsQueryComposer` reconstructs the stored attributes directly:
external targets emit `IN '<path>'`, Attribute 2 rows become `PARAMETERS`
clauses, UPDATE/DELETE emit `DISTINCTROW` when bit 8 is set, UNION appends
stored `ORDER BY`, and aliases use a wider Access reserved/contextual keyword
set for bracketing.

**What this rules out**: Do not classify these as formatting-only review
cases, and do not solve them by changing only `modTestQuerySqlBuilder`.
Canonical comparison is allowed to ignore presentation drift, but not loss of
stored query attributes.

**Relevant files**: `Version Control.accda.src/modules/Utility/clsQueryComposer.cls`,
`Testing/Fixtures/queries/regression/qryRegressionExternalMakeTable.*`,
`qryRegressionParameterizedCrosstab.*`, `qryRegressionUnionOrderBy.*`,
`qryRegressionDeleteDistinctRow.*`, `qryRegressionUpdateDistinctRow.*`,
`qryRegressionReservedAlias.*`, `docs/access-query-storage.md`.

---

## 2026-04-28 — Replace JSON index with binary `.idx` format and promote `clsVCSIndexItem` to persistent storage

**Trigger**: On a large database (db-sec: ~3,500 component entries), the `vcs-index.json` file grew to 1.5MB / 40K lines. Parsing it via `modJsonConverter.ParseJson` consumed 1.5-2.2s per export — nearly half the total runtime for a no-changes export. The bottleneck was threefold: three `Replace()` calls stripping whitespace from a 1.5MB string, character-by-character recursive descent creating ~10,000 `Scripting.Dictionary` COM objects, and ~3,500 ISO 8601 date string parses.

**Options explored**:
- **SQLite sidecar database**: Maximum query flexibility and proven binary format. Rejected: requires distributing and maintaining an external DLL dependency (`sqlite3.dll`), version management across 32/64-bit Access, and introduces a non-VBA dependency for a core infrastructure component.
- **ACE/DAO sidecar `.accdb`**: Zero-dependency since the Jet/ACE engine is always present. Considered seriously, but rejected: adds file locking complexity, requires schema migrations for index structure changes, and the overhead of opening a second database connection on every export.
- **Optimized JSON (minified, pre-sorted)**: Marginal improvement. The fundamental bottleneck is the recursive descent parser and COM object creation, not whitespace. Would not change the O(n) string manipulation cost.
- **Custom binary flat file** (chosen): A length-prefixed binary format using VBA's native UTF-16LE string encoding and raw `Double` dates. Eliminates all JSON parsing overhead. Dates stored as UTC for cross-timezone portability. File size drops ~73% (1.5MB to ~400KB). Load time drops ~90-95% (1.5s to ~0.05-0.15s).

**Decision**: Two coordinated changes in `clsVCSIndex.cls`:

1. **Binary format**: `vcs-index.json` replaced by `vcs-index.idx`. Format is: 4-byte magic (`VCSI`), 2-byte version (UInt16 LE), global date properties (UTC doubles), length-prefixed category hashes, then per-category component entries with a flags byte controlling which optional hash strings are present. Strings are length-prefixed UTF-16LE (VBA native, zero conversion cost). Dates use `LSet` UDT punning for `Double` <-> `Byte()` conversion (no `CopyMemory` dependency). UTC conversion uses existing `ConvertToUtc`/`ConvertToLocalDate` from `modUtcConverter.bas`.

2. **Entry storage refactoring**: `clsVCSIndexItem` promoted from a throwaway view object (created fresh on every `Item()` call, linked to a per-entry `Dictionary` via `dParent`) to persistent storage (stored directly in `m_dIndex("Components")(category)(filename)`). Eliminates ~3,500 per-entry `Dictionary` objects. The `dParent` property was removed from `clsVCSIndexItem`. The public API (`Item`, `Update`, `Remove`, `Exists`) is unchanged — callers still receive `clsVCSIndexItem` objects.

No backward compatibility: if `vcs-index.idx` is missing and `vcs-index.json` is found, the legacy file is deleted. The next full export regenerates the binary index. No export format version gating since the index is gitignored and local.

A `DumpToJson` method is available for troubleshooting — it reconstructs a temporary `Dictionary` tree and serializes it through the existing JSON pipeline.

**What this rules out**: The index can no longer be inspected with a text editor or `jq`. Use `VCSIndex.DumpToJson` (from the Immediate Window or via `vcs_run_vba`) to generate a human-readable JSON snapshot. Any future index schema changes must bump `IDX_FORMAT_VERSION` and handle the version mismatch in `LoadFromFile` (currently treats unknown versions as corrupt, triggering a full re-export). Adding new fields to `clsVCSIndexItem` requires updating both `Save` (write the field) and `LoadFromFile` (read it), plus bumping the format version.

**Relevant files**: `clsVCSIndex.cls` (binary I/O, entry storage refactoring), `clsVCSIndexItem.cls` (removed `dParent`), `.gitignore` / `.gitignore.default` (changed `vcs-index.json` to `vcs-index.*`).

---

## 2026-04-28 — Handle FROM-clause subqueries at the emitter (`BuildFromClause`), not upstream in `ReconstructSQL` Case 5

**Trigger**: A user reported that `clsQueryComposer.ReconstructSQL` was emitting `FROM   AS % $ ##@_Alias;` for queries with a derived table in the FROM clause (subquery), losing the entire subquery body. Two coordinated bugs: the FROM emitter read MSysQueries `Name1` (NULL for derived tables) instead of `Expression` (the inner SELECT), and `BracketIfNeeded` did not bracket the `%$##@_Alias` placeholder, so `clsSqlFormatter` then tokenized `%`, `$`, `#`, `@` as separate operators. See [docs/access-query-storage.md § 6](docs/access-query-storage.md) for the empirical evidence and [regression/qryRegressionFromSubquery](Testing/Fixtures/queries/regression/qryRegressionFromSubquery.sql) for the pinned shape.

**Options explored**:
- **Detect derived tables in `ReconstructSQL` Case 5 and pre-populate `name = "(" & expression & ")"`**: rejected. Attribute 5 has the same shape (Name1 empty, Name2 = alias / segment id, Expression contains a SELECT) for both derived tables *and* UNION segments. UNION segments go through the dedicated case 9 branch (line 321) which reads `expression` directly and never reaches the FROM emitter. Mutating `name` upstream couples the derived-table fix to UNION processing -- benign today (case 9 doesn't read `name`), but every future maintainer of the UNION branch would have to remember the upstream rewrite. Fragile.
- **Read `Expression` directly inside `BuildFromClause` whenever `Name1` is empty**: rejected for code-clarity reasons. The check would have to repeat at three call sites (the no-joins branch line 672, the join-chain `dTableLookup` line 697, and the cartesian fallback line 727), and each site would conflate "is this a derived table" with "format this as a FROM operand." Hard to grep for, easy to miss when adding a fourth FROM emission path.
- **Force Design View qdef for FROM-subquery shapes (mirroring the multi-cond `ON` workaround)**: rejected. `IsDesignerCompatible` already returns False for `HasSubqueries`, so the importer correctly emits SQL View qdef -- which `LoadFromText` accepts. Forcing Design View would re-introduce the legacy 4.x `InputTables.Name = "<entire SELECT>"` / `Alias = "%$##@_Alias"` shape that `LoadFromText` rejects with "Resource failure" (the original user bug). The export reconstruction is the right fix layer.
- **Centralize at the emitter via a `FormatInputTableName` helper**: chosen. Single function captures the "render an input table for a FROM clause" rule (derived-table → `(<expr>)`, normal → `BracketIfNeeded(name)`); all three FROM emission sites now route through it. UNION processing is unaffected because case 9 never calls the helper.

**Decision**: Handle FROM-clause derived tables centrally at the emitter (`BuildFromClause` via the new private helper `FormatInputTableName`), and broaden `BracketIfNeeded` via a new `HasNonIdentChars` predicate to bracket any identifier with characters outside `[A-Za-z0-9_]`. The two fixes are coordinated -- without the bracketing fix the formatter still mangles the alias even with the subquery correctly emitted; without the emitter fix the alias is correct but references nothing.

**What this rules out**: Refactoring the derived-table handling back into `ReconstructSQL` Case 5 (the more "intuitive" location) -- doing so re-couples the fix to UNION processing and makes future UNION changes risky. Loosening `HasNonIdentChars` to accept additional characters (e.g. `?`, `!`, `#` in pre-bracketed contexts) -- the simpler "alphanumeric + underscore only" rule covers all known Access auto-generated alias shapes (`%$##@_Alias`, `~sq_*`, `~TMPCLP*`) and matches what `[...]` brackets already escape, so any expansion would need a fixture proving the looser rule is needed. Reverting to single-call-site bracketing in `BracketIfNeeded` (e.g. only checking spaces) -- the formatter's tokenization of `%`, `$`, `#`, `@` is the actual constraint and is independent of identifier choice.

**Relevant files**: Modified: `Version Control.accda.src/modules/Utility/clsQueryComposer.cls` (added `FormatInputTableName`, `HasNonIdentChars`; extended `BracketIfNeeded`; routed three FROM emission sites through `FormatInputTableName`). New: `Testing/Fixtures/queries/regression/qryRegressionFromSubquery.{sql,json,notes.md}` (regression pin). Updated: `docs/access-query-storage.md` §§ 4 and 6 (added "Derived table in FROM" row to handled-shapes table, new finding subsection documenting the bug + fix).

---

## 2026-04-27 — Adopt top-level `docs/` folder for internal reference documentation (separate from public-facing `Wiki/`)

**Trigger**: Drafting the first long-form internal reference doc (`docs/access-query-storage.md`, ~28 KB synthesizing MSysQueries field semantics, Design View vs SQL View arbitration, the `LoadFromText` / `SaveAsText` asymmetries the round-trip harness exposed, and parser-handled-vs-known-gaps tables) raised the question of where this kind of content belongs. None of the existing venues fit cleanly: `Wiki/` is user-facing how-to that syncs to the public GitHub Wiki, `AGENTS.md` is workflow + standards, `DECISIONS.md` is the why-journal, and per-fixture `.notes.md` files are bug-specific. A long reference about how a third-party system (Access query storage) works and what the add-in depends on from it doesn't match any of those audiences.

**Options explored**:
- **Put the new doc in `Wiki/`**: rejected. Wiki pages sync to the public GitHub Wiki and are written for end users learning to use the add-in. A long internal reference about MSysQueries field bits, `Lv*` binary blobs, and `LoadFromText` rejection asymmetries dilutes the wiki for that audience and pulls maintenance attention away from the user-facing pages already there (`Options.md`, `FAQs.md`, `Supported-Objects.md`, etc.).
- **Co-locate with the artifacts** (e.g. `Testing/Fixtures/queries/REFERENCE.md`): rejected. The query doc covers parser logic in `clsQueryComposer.cls` and `clsDbQuery.cls`, which live under `Version Control.accda.src/modules/`, so co-location with fixtures is a partial fit at best. More importantly, the same problem repeats for plausible future siblings (form storage, report storage, COM ribbon DLL, hook DLL): each would need its own scattered home, defeating consolidation. Per-artifact `.notes.md` for narrow bug-specific context is still the right pattern at that scope; long-form reference about a *family* of artifacts is a different shape.
- **Embed the content into `AGENTS.md`**: rejected. `AGENTS.md` is already a long workflow/standards guide; absorbing multiple 20–30 KB references would bury the workflow guidance under reference material. `AGENTS.md` should *point at* `docs/` references (it now does, in the new "Before changing the query parser" subsection), not contain them.
- **Top-level `docs/` folder**: chosen. Conventional OSS layout — a separate venue for developer/maintainer reference, distinct from user-facing wiki content. Future siblings (`access-form-storage.md`, `access-binary-formats.md`, `com-ribbon-addin.md`, `hook-dll-architecture.md`, etc.) cluster naturally without needing per-doc location decisions.

**Decision**: Top-level `docs/` is the home for internal/agent-facing reference documentation about underlying systems and what the add-in depends on (Access internals, binary blob formats, COM ribbon architecture, hook DLL architecture, etc.). `Wiki/` continues to hold user-facing how-to material. A small `docs/README.md` index file is added now so the folder's intent is visible at the folder level and future contributors/agents don't have to infer it from a single existing entry.

**What this rules out**: Putting future internal/maintainer reference material into `Wiki/` — the user/internal split is now load-bearing. Litmus test: if a doc's primary audience is end users learning the product, `Wiki/`; if it's a developer/agent reference about how something works internally or what we depend on, `docs/`. Co-locating long-form reference docs with their artifacts (per-artifact `.notes.md` companions for narrow bug-specific context still belong with the artifact; long-form references about a family of artifacts go in `docs/`). Treating `docs/` as a dumping ground for one-shot or session-scoped notes — entries here are sustained reference material, edited as understanding evolves; one-shot architectural rationale belongs in `DECISIONS.md`, and bug-specific context belongs in a `.notes.md`.

**Relevant files**: New: `docs/access-query-storage.md` (first reference doc, seed of the family), `docs/README.md` (folder index). Cross-references already in place: `Testing/Fixtures/README.md` ("Documenting parser invariants and edge cases" section links to `docs/access-query-storage.md`), `Version Control.accda.src/AGENTS.md` ("Before changing the query parser" subsection links the same doc).

---

## 2026-04-24 — Object round-trip regression harness lives inside the add-in, fixtures are versioned text files, queries pilot the IDbComponent abstraction, and the public surface routes through `clsVersionControl`

**Trigger**: Post-`clsQueryComposer` work on the SQL/JSON query format surfaced ~723 affected queries in `db-sec` from a single self-join alias bug (`qryCurrencyCrossRates` archetype). Manual repro-and-fix is unsustainable as more edge cases land. Traditional VBA unit testing (Rubberduck-style or hand-rolled) would require hundreds of fixture queries hard-coded into the add-in — thousands of lines of VBA permanently loaded into memory in every running instance, for code paths that are only exercised during development. A different shape was needed.

**Options explored**:
- **Per-query VBA unit tests with hard-coded SQL strings**: rejected. Bloats the add-in's `.accda` permanently for a dev-only feature; every new edge case requires editing VBA and redeploying; no easy way to inspect the input/output of a specific failing case.
- **External test harness in the existing `Testing.accdb` database that calls the add-in via Automation**: rejected. Splits the test code from the add-in code that produces the export/import logic; loses access to internal helpers (`modFileAccess`, `modHash`, `clsLog`, `Operation`, `VCSIndex`); developers would have to context-switch between two databases mid-debug; agents using `vcs_*` MCP tools would have to coordinate across two `.accdb` files.
- **Harness inside the add-in, fixtures stored *inside* the test database (sample queries baked into `.accda`)**: rejected. Same bloat problem at smaller scale; queries can't be diff-reviewed in PRs; rebaselining requires re-exporting a binary database.
- **Harness inside the add-in, fixtures as text files in the repo**: chosen. The harness has full access to internal helpers; fixtures are diffable in PRs; new fixtures cost only two text files (`.sql` + `.json`); the bloat from sandbox queries created during a run is addressed structurally (see below) and the worst case is a `compact-and-repair` or rebuild-from-source — acceptable for a dev/CI-only operation.
- **One-pass round-trip (import → export → diff against fixture)**: rejected. Misses non-deterministic export bugs where Pass 1 happens to match the fixture but Pass 2 (re-importing the Pass 1 output) produces a different export. The two-pass design (Pass 1 vs. fixture *and* Pass 2 vs. Pass 1) catches both regressions and idempotency failures with the same fixture corpus.
- **Query-only harness with the abstraction left for "later"**: rejected. The dispatch layer (`Run<Type>Fixtures` per component, category subfolders, `_scaffold/` for shared dependencies) costs almost nothing now, but retrofitting it after queries are entrenched would force a breaking reorganization of every existing fixture path. Building on `IDbComponent` from day one means future component types (forms, reports, modules) plug in without touching any existing code.
- **JSON name-rewriting for comparison** (rewrite `Info.Description` from sandbox name → original name in the Pass 1 output before diffing): rejected. Brittle — every name-bearing field needs explicit handling, easy to miss future fields. The cleaner answer is to drop the entire `Info` block: it's purely descriptive metadata for human readers and is *not* consumed by `clsDbQuery.ImportNewFormat` (which reads the query name from the filename, not the JSON). Stripping `Info` wholesale is name-agnostic, format-agnostic, and degrades gracefully if new descriptive fields are added later.
- **Expose `RunObjectRoundtripTests` / `RunOurFixtures` as `Public` functions in `modTestRoundtrip.bas` *without* `Option Private Module`**: rejected. Reachable from cross-project `Application.Run` without going through the documented API surface, and — worse — any future helper in the same module that's added without the `Private` keyword would silently leak. Inconsistent with the rest of the add-in, where every implementation module hides behind `Option Private Module` and is reached only through `clsVersionControl`.
- **Expose `RunObjectRoundtripTests` directly via `vcs_call_vba`** (which uses `Application.Run` and doesn't require `McpAllowRunVBA`): rejected as the *primary* path. The agent-friendliness gain isn't worth either keeping the module publicly exposed or carving out a private-module exception for `Application.Run` lookup. The harness *is* arbitrary code execution from the user's perspective (it imports/exports/deletes objects), so gating it behind the same `McpAllowRunVBA` opt-in that already governs `vcs_run_vba` is the correct security model — not a friction worth designing around.
- **Single delegate method on `clsVersionControl` (`VCS.RunRoundtripTests`) with `Option Private Module` on `modTestRoundtrip.bas`**: chosen. Matches the established add-in pattern exactly (everything user-visible lives on `clsVersionControl`; implementation modules are private). One curated public symbol instead of N. Future helpers added to the test module are automatically blocked from external callers — no future-leak hazard. Immediate-Window dev access from inside the add-in's own VBE still works (`?modTestRoundtrip.RunObjectRoundtripTests()`) because `Option Private Module` only blocks cross-project lookups, not in-project ones. `RunOurFixtures` is dropped as redundant — `RunRoundtripTests()` with no args produces the identical zero-arg-shipped-corpus behavior.

**Decision**: Implement `modTestRoundtrip.bas` inside the add-in with `Option Private Module` and `RunObjectRoundtripTests(Optional strFixtureFolder, Optional blnRebaseline)` as its single in-project entry point. Expose this externally through one public delegate, `clsVersionControl.RunRoundtripTests`, alongside the other dev/agent tools (`RunVBA`, `ExecuteSQL`, `CompileVBA`). External invocation: Immediate Window uses `?VCS.RunRoundtripTests`; MCP/CI uses `vcs_run_vba` with `MCP_TempFunction = VCS.RunRoundtripTests()` (gated by `McpAllowRunVBA`). Fixtures live in `Testing/Fixtures/<component>/<category>/` as plain text (`.sql` + `.json` for queries today; the slot is reserved for `forms/`, `reports/`, etc.) with a `_scaffold/` sibling folder for shared supporting objects loaded once per session. Each fixture runs through a two-pass round trip (import to `vcs_test_<name>_<hash>` sandbox, export, re-import, re-export) with three independent SHA-256 comparisons: Pass 1 vs. fixture, Pass 1 vs. Pass 2 (idempotency), JSON-with-`Info`-stripped both directions. Bloat is addressed structurally: random-suffix sandbox names allow parallel runs and unambiguous leftover detection, every fixture cleans up via `DoCmd.DeleteObject` + `DBEngine.Idle dbRefreshCache`, the run starts with a `CleanupStaleObjects` sweep over any `vcs_test_*` survivors from a crashed prior run, and `VCSIndex.Disabled = True` for the entire run prevents test operations from polluting `vcs-index.json`. Output flows through the existing `Log` singleton (live console in `frmVCSMain` + per-session `ObjectRoundtrip_<opId>.log` with full inline diffs) and a structured JSON return for programmatic parsing. Bug-as-fixture is the canonical contribution path: real-world failures from `db-sec` or user reports are distilled into a fixture under `regression/` with a `.notes.md` companion documenting the failure mode and resolution status — `qryCurrencyCrossRates` is the seed entry, currently failing as expected.

**What this rules out**: Storing test fixtures inside any `.accdb` (must remain text files in the repo). Per-component bespoke comparison logic — new component types must conform to the import-export-compare shape and use the shared `Run<Type>Fixtures` dispatch. Loading fixture corpora that exceed sandbox-name uniqueness guarantees (the 7-hex-char suffix gives ~268M combinations per fixture name; collision-handling beyond that is not designed for). Adding *additional* test entry points to the add-in's external API surface without an explicit follow-up decision — `VCS.RunRoundtripTests` is the single sanctioned public method; future test categories (perf, validation, etc.) should add new module(s) under the `modTest*` convention with their own delegate methods on `clsVersionControl` rather than expanding the test modules' own public surface. Reaching the harness via `vcs_call_vba` (the lower-friction MCP path that doesn't require `McpAllowRunVBA`) — agents must use `vcs_run_vba` with the security gate enabled, by design. JSON comparison schemes that depend on specific field names (the `Info`-stripping strategy assumes the import path will continue to ignore `Info`; if a future format change makes `Info` semantically load-bearing, the comparator must change in lockstep). Combining the harness with operations that want to own the global `Operation` state — `RunObjectRoundtripTests` calls `Operation.Begin(eotOther)` and refuses to run if another operation is in flight, so it cannot be invoked from inside an active export/build/merge.

**Relevant files**: New: `Version Control.accda.src/modules/Tests/modTestRoundtrip.bas` (harness, `Option Private Module`), `Testing/Fixtures/README.md`, `Testing/Fixtures/.gitignore`, `Testing/Fixtures/_scaffold/.gitkeep`, `Testing/Fixtures/queries/{select,crosstab,append,update,delete,regression,passthrough,union,ddl}/` (15 seeded fixtures with `.sql`+`.json` pairs and four `regression/*.notes.md` files), `Wiki/Regression-Testing.md`. Updated: `Version Control.accda.src/modules/API/clsVersionControl.cls` (added `RunRoundtripTests` delegate method), `AGENTS.md` (Testing Strategy section + `modTest*` convention), `Wiki/Home.md` (link to new page). Consumed but unchanged: `Version Control.accda.src/modules/Components/clsDbQuery.cls` (export/import path being verified), `Version Control.accda.src/modules/Utility/clsQueryComposer.cls` (subject of the regression harness).

---

## 2026-04-24 — Adopt `modTest*` family-prefix convention for test modules; rename `modUnitTesting.bas` → `modTestSuite.bas`

**Trigger**: Adding a new test module (`modTestRoundtrip.bas`) for query round-trip regression testing prompted thinking about a forward-looking naming convention for test-infrastructure modules. The existing `modUnitTesting.bas` name described the *style* (Rubberduck unit-testing) rather than its actual *contents*, and after the earlier 2026-04-24 Rubberduck-removal decision the file isn't unit-test-framework-style in any meaningful sense anyway — it's a heterogeneous Debug.Assert-based catch-all (encoding, JSON, sanitization, formatter, hashing, IDbComponent invariants, path utilities, etc.). The codebase already uses family-prefix grouping for related types (`clsDb*` for IDbComponent implementations, `clsLv*` for ListView property parsers); test modules should follow the same established pattern.

**Options explored**:
- **Status quo: leave `modUnitTesting`, name the new module `modRoundtripTests` or similar**: rejected. Names diverge in style; new contributors have no convention to follow; future test modules will keep accumulating ad-hoc names with no shared discoverability hook.
- **Family suffix `mod*Testing`**: e.g., `modUnitTesting` (existing), `modRoundtripTesting` (new): rejected. Less prominent placement of "Test" in alphabetical sort; doesn't group test modules together when scanning a flat module list in the VBE Project pane or in grep output.
- **Family prefix `modTest*` with rename**: `modTestSuite` (renamed) + `modTestRoundtrip` (new) and any future siblings: chosen. Matches existing `clsDb*`/`clsLv*` family-grouping convention; "Test" appears at the front for maximum discoverability; alphabetical sort groups all test modules together; new test modules conform automatically without needing per-contributor reminders.
- **Defer rename and grandfather `modUnitTesting`**: rejected. Permanent inconsistency for a one-time cost. The rename is cheap — single module with no external API consumers (Rubberduck reads the `@TestModule` attribute, not the module name; after the Rubberduck-removal decision even that no longer applies). The convention should be in place *before* the second test module ships, not after.
- **Aggressively split `modTestSuite` into focused modules by topic** (`modTestEncoding`, `modTestSqlFormatter`, `modTestHashing`, etc.): deferred. Worth doing if the suite grows further; today the rename alone establishes the convention without forcing a content reorganization that would expand the diff and risk introducing test regressions.

**Decision**: Adopt the `modTest*` family-prefix convention for all test-infrastructure modules. Rename `modUnitTesting.bas` → `modTestSuite.bas` (better describes the heterogeneous catch-all contents than "unit testing"). Two changes inside the file: `Attribute VB_Name` and the `Private Const ModuleName` constant. The convention is documented in `AGENTS.md` so future test modules conform. Future siblings (`modTestRoundtrip`, `modTestPerf`, `modTestFixtures`, etc.) will conform automatically.

**What this rules out**: Mixed naming conventions within `modules/Tests/`. Naming new test modules without the `modTest*` prefix without an explicit decision overriding this one. Folder reorganization that would break the family-prefix grouping (e.g., moving test modules out of `Tests/` into per-topic folders that also contain non-test code). If the test suite grows large enough to warrant a split into focused modules, those splits must also use the `modTest*` prefix (`modTestEncoding`, `modTestSqlFormatter`, `modTestHashing`, etc.) rather than reverting to topic-only names.

**Relevant files**: Renamed: `Version Control.accda.src/modules/Tests/modUnitTesting.bas` → `Version Control.accda.src/modules/Tests/modTestSuite.bas`. Convention documented in: `AGENTS.md`. Coming next (separate decision entry): `Version Control.accda.src/modules/Tests/modTestRoundtrip.bas`.

---

## 2026-04-24 — Drop Rubberduck testing-framework dependency from `modUnitTesting.bas`

> **⚠ Partially superseded** (2026-04-24): The decision content still applies in full, but the file `modUnitTesting.bas` was subsequently renamed to `modTestSuite.bas` to fit the `modTest*` family-prefix convention. See "Adopt `modTest*` family-prefix convention for test modules" above.

**Trigger**: The unit-test module created `Rubberduck.AssertClass` and `Rubberduck.FakesProvider` COM objects in its `ModuleInitialize`, so the entire test suite failed to even initialize unless Rubberduck was installed and registered. The user reported that the Rubberduck Test Explorer is virtually unusable in their larger production databases, that `Rubberduck.FakesProvider` was never actually used anywhere in the file (only initialized and torn down), and that Rubberduck itself is shifting direction — making it a poor long-term peg to hang the add-in's tests on. Of the ~20 tests in the module, only three (`TestUCS2toUTF8RoundTrip`, `TestParseSpecialCharsInJson`, `TestSortDictionaryByKeys`) actually called `Assert.AreEqual` / `Assert.Fail` / `Assert.Succeed`; the rest already used native `Debug.Assert`.

**Options explored**:
- **Status quo — keep the Rubberduck dependency**: rejected. The framework is a hard runtime requirement (`CreateObject` fails if not registered) for a feature that ~85% of the existing tests don't actually use. With Rubberduck's own roadmap in flux, betting future tests on its annotations and `Assert.*` API is increasing risk for no current benefit.
- **Remove the testing framework AND the `PreserveRubberDuckID` option AND the `@Folder` annotation support**: rejected. Despite their names, neither of the latter two requires Rubberduck to be *installed* — `PreserveRubberDuckID` (in `clsOptions` / `clsDbVbeProject`, see issue #197) only decides whether to preserve a numeric ID Rubberduck happens to stash in the VBE project's `HelpFile` field, and `@Folder` annotation parsing (in `modVbeUtility.bas`, gated behind `EFV_5_0_0` per the 2026-03-10 decision) is a self-contained subfolder-organization feature that just borrows Rubberduck's annotation syntax. Removing them would punish users who rely on those interop features without any payoff in dependency reduction.
- **Add a `RunAllTests` orchestrator that loops every test sub and prints a pass/fail summary**: rejected for now. Each test is already a self-contained `Sub` callable from the Immediate Window, the suite is small enough that batch invocation isn't a real friction, and adding an orchestrator would require either reflection (no clean VBA story) or a hand-maintained list that drifts from the actual test set. The user explicitly chose individual invocation.
- **Keep the inert Rubberduck annotations** (`'@TestModule`, `'@TestMethod("...")`, etc.) as comments since they don't break anything when Rubberduck isn't installed: rejected. They suggest a framework dependency that no longer exists and would mislead future contributors into thinking the Test Explorer integration is supported. Strip them all *except* `'@Folder("Tests")`, which is still a live, actively-used feature of this codebase.
- **Wrap each converted test in `On Error GoTo TestFail` scaffolding so unexpected runtime errors get logged** (mirroring the original three Rubberduck tests): rejected. The other ~17 native-`Debug.Assert` tests already let runtime errors propagate naturally, which is fine for tests run individually from the IDE — the dev sees the error dialog with line context. Adding scaffolding would make the converted three inconsistent with the rest of the file.

**Decision**: Strip the Rubberduck *testing* dependency only. In `modUnitTesting.bas`: delete `Private Assert As Object` / `Private Fakes As Object` and the four lifecycle subs (`ModuleInitialize`, `ModuleCleanup`, `TestInitialize`, `TestCleanup`); remove `'@TestModule` and all 14 `'@TestMethod("...")` annotations; convert the three `Assert.*`-using tests to plain `Debug.Assert` (promoted from `Private` to `Public Sub` for Immediate-Window invocation, scaffolding dropped); preserve `'@Folder("Tests")`. The unrelated `PreserveRubberDuckID` option and `@Folder` annotation feature stay untouched.

**What this rules out**: Test Explorer-style GUI test running (no `'@TestMethod` discovery). Adding new tests with the Rubberduck `Assert.*` API — future tests must use `Debug.Assert` or roll their own assertion helpers. Reviving a `RunAllTests` orchestrator without a deliberate follow-up decision (the file no longer has a discoverable list of test names; an orchestrator would need a hand-maintained registry). If Rubberduck's annotation syntax for `@Folder` ever changes incompatibly, the `@Folder("Tests")` line in this file becomes an inert comment alongside everything else — but the underlying subfolder-organization feature lives in `modVbeUtility.bas` and would need its own decision (see 2026-03-10 entry).

**Relevant files**: `Version Control.accda.src/modules/Tests/modUnitTesting.bas` (sole edit). Untouched but in scope of the discussion: `Version Control.accda.src/modules/Infrastructure/clsOptions.cls`, `Version Control.accda.src/modules/Components/clsDbVbeProject.cls`, `Version Control.accda.src/forms/frmVCSOptionsAdvanced.cls/.form`, `Version Control.accda.src/vcs-options.json`, `Testing/Testing.accdb.src/vcs-options.json`, `Version Control.accda.src/modules/Core/modVbeUtility.bas`.

---

## 2026-04-24 — Auto-inject VBA line numbers in `RunVBA` wrapper for `Err.Erl` reporting

**Trigger**: When agents call `vcs_run_vba` (which routes to `clsVersionControl.RunVBA`), the wrapper used `On Error Resume Next` and only surfaced `Err.Number` / `Err.Description` from whichever statement errored last. There was no indication of which line of the agent's submitted `code` actually failed, so debugging multi-statement test snippets meant guessing from the description alone or chopping the code into single-statement calls. The user proposed leveraging VBA's `Err.Erl` intrinsic (which returns the most recently executed labeled line at the time of the error) by adding line numbers to the dynamically generated test procedure.

**Options explored**:
- **Documentation only**: tell agents to hand-number their own snippets and use `Erl` if they want line tracking. Rejected — every test would carry boilerplate the wrapper could trivially generate, and most agents would skip it for one-off probes, losing the diagnostic for free wins. The user explicitly asked which would be better for agents; auto-injection was the answer.
- **Auto-inject only when the agent opts in** via a flag arg on the MCP tool. Rejected as needless ceremony — line numbers are cheap, harmless to correctly-written code, and `Erl` is `0` when no error fires so success paths see no change.
- **Step size 10/20/30** (traditional QBASIC convention, leaves gaps for hand-edits). Considered briefly; rejected because this is throwaway generated code that nobody hand-edits between generation and execution. Step of 1 makes `errorLine` equal the 1-based ordinal of the line within the agent's submitted `code` string, which is the most intuitive thing for the agent to interpret — no offset math to map a reported line back to source.
- **Capture only the FIRST error** using `On Error GoTo H` + `Resume Next` from a real handler (records `Erl` exactly when it's still pointing at the failing line, then continues). Rejected for now because it would change the long-standing observable semantic of "what error gets reported" from "last" to "first" without user consensus, and the user clarified that multi-error capture is a per-test agent decision rather than a default. The single-line capture-at-end pattern still works correctly: with `On Error Resume Next`, `Erl` is updated each time an error is raised, so `m_ErrLine = Erl` after the user code reflects the *last* error's line (consistent with the existing `m_ErrNum` / `m_ErrDesc` capture).
- **Number every line including blanks/comments** (so output line offset literally equals input line offset). Adopted only for the *counter*, not the prepended digits — VBA rejects line numbers on blank or pure-comment lines, and continuation lines (those following a `_`-terminated parent) cannot carry their own number. Final design: counter advances on every physical input line (so `errorLine` matches the agent's source), but the digits are only prepended to lines that can legally hold one. Pre-numbered lines (caller already wrote `5 Foo`) are detected by leading-digit and passed through, letting agents override.

**Decision**: New private helper `AddVbaLineNumbers` in `clsVersionControl` walks the submitted code and prepends 1-based line numbers to each executable statement; the wrapper template gains `m_ErrLine As Long` plus a `MCP_GetErrLine` accessor, captures `Erl` immediately after `Err.Description`, and the `RunVBA` JSON result gains an `errorLine` field that is omitted when `Erl` is `0`. The default capture remains last-error-wins (no behavior change for callers that don't read the new field). Agents who need richer per-error reporting are documented to write their own `On Error GoTo H` / `Resume Next` handler that reads `Erl` into a collection — the auto-injected numbers make this work without the agent having to write any line numbers themselves.

**Continuation-line detection gotcha**: First pass detected continuations by `Right$(strTrimmed, 1) = "_"`. That misfires on identifiers ending in underscore (`Dim Foo_`, `Set rs_ = ...`) — a common VBA naming pattern that would have caused the next line to be treated as a continuation and lose its number. Fixed by additionally requiring the character before the trailing `_` to be a space or tab (which is what VBA's actual continuation marker requires). `Trim$` strips trailing whitespace so the post-trim string ends literally with `... _` for genuine continuations.

**What this rules out**: Switching to first-error-wins capture without a deliberate follow-up decision (the wrapper now exposes `errorLine` for the last error; flipping to first-error would change which `errorLine` value a given test reports). Removing line-number injection without breaking the documented `errorLine` contract. Agents writing tests that assume line numbers are *not* present in the executed code (e.g., parsing the `code` string back from `generatedSource` in compile-error responses) — `generatedSource` now contains numbered lines.

**Relevant files**: `Version Control.accda.src/modules/API/clsVersionControl.cls` (added `AddVbaLineNumbers`, modified `RunVBA` wrapper template and JSON result construction); `C:\Repos\msaccess-vcs-mcp\src\msaccess_vcs_mcp\tools.py` (extended `vcs_run_vba` docstring with line-number behavior and multi-error pattern); cached MCP descriptor `mcps/user-msaccess-vcs-mcp/tools/vcs_run_vba.json` (mirrored docstring update); `AGENTS.md` (new "Debugging RunVBA Failures" section).

---

## 2026-04-20 — Wrap query composer pipeline in CatchAny error handling

**Trigger**: The new `clsQueryComposer` (introduced as part of the 5.0 deterministic-query format, see entry "Replace SaveAsText with MSysQueries-based query export") had no error handling on any of its parsers, emitters, or helpers. An unexpected VBA error inside a single query during a full export or build would drop into break mode in debug builds, and in release builds would either bubble up to a parent's handler with no useful context or crash the entire batch. Same risk existed in `clsDbQuery.ExportNewFormat` / `ImportNewFormat`, which had no top-level error guards at all.

**Options explored**:
- **Wrap only the four public methods** (`ReconstructSQL`, `DecomposeSQL`, `IsDesignerCompatible`, `GenerateQdef`). Minimal boilerplate; errors in private helpers still propagate up to the public method's `On Error Resume Next` and get logged once. Rejected as the sole scope because the resulting log line only identifies which public method failed, not which parser stage — a `ParseJoinExpression` failure is indistinguishable from an `EmitDesignLayout` failure in the log.
- **Wrap every helper, including leaves** (`BracketIfNeeded`, `IsAccessReservedWord`, `FindMatchingParen`, etc., ~50 functions). Maximum log granularity but adds ~5 lines of identical boilerplate to every trivial string helper. The error would already be logged by the wrapped parent — no incremental information. Rejected as bloat.
- **Wrap composer publics + ~13 major top-level private helpers (chosen)**, plus `clsDbQuery.{IDbComponent_Export, IDbComponent_Import, ExportNewFormat, ImportNewFormat, ExportLegacy}`. Each wrapped composer helper logs its procedure name and a 200-char SQL snippet so a failure in `EmitColumnMetadata` is distinguishable from one in `BuildJoinChain`. Leaf helpers stay unwrapped — their errors still bubble up to the nearest wrapped parent.
- **Add a `Name` property to `clsQueryComposer`** so error messages could include the query name. Rejected — the calling `clsDbQuery` already prints the query name in surrounding `Log.Add` / `Perf.OperationStart` lines, and adding mutable state to the composer just for logging context would be a regression. A new private `SqlSnippet()` helper truncates `m_strRawSql` to 200 chars, which gives enough context to identify the query when scrolling logs.

**Decision**: Two-layer protection: every public method and every "stage-level" private helper in `clsQueryComposer` uses the standard `If DebugMode(True) Then On Error GoTo 0 Else On Error Resume Next` / `CatchAny` pattern, with `CleanUp:` labels routed through `GoTo CleanUp` (replacing 4 internal `Exit Sub` / `Exit Function` early-exits) so the `CatchAny` always runs. The five `clsDbQuery` entry points carry the same wrap so errors from anywhere in the composer or its callers are converted to log entries instead of break-mode entries. `CatchAny` calls in the two file-writing paths (`ExportNewFormat`, `ImportNewFormat`) pass `blnIncludeErrorWithDescription:=True` so the underlying VBA error number/description appears in the log.

**DAO recordset cleanup gotcha**: First implementation pattern in `ExportNewFormat`'s `CleanUp:` block was `If rst.State <> 0 Then rst.Close` — copied from ADO recordset cleanup elsewhere in the codebase. **`DAO.Recordset` has no `.State` property** (that's an ADO concept); accessing it raises 438. More damaging: when the body completes normally, it has already called `rst.Close` but left the reference set; the cleanup's second `rst.Close` raises 3420 ("Object invalid or no longer set"). Even with `On Error Resume Next` silencing the error, `Err.Number` remains set when control reaches `CatchAny`, which dutifully logs a phantom failure for every successful export. First test run logged 15 false errors out of 15 exports. Final pattern (now used in both export and import cleanup blocks):

```vba
CleanUp:
    Dim lngOrigErr As Long, strOrigDesc As String
    lngOrigErr = Err.Number
    strOrigDesc = Err.Description
    On Error Resume Next
    If Not rst Is Nothing Then rst.Close
    Set rst = Nothing
    Err.Clear
    If lngOrigErr <> 0 Then Err.Raise lngOrigErr, , strOrigDesc
    CatchAny eelError, "...", ModuleName(Me) & ".ExportNewFormat", True, True, True
```

**ExportLegacy inner-handler conflict**: The existing inner `Catch(3258)` block ended with `On Error GoTo 0`, which would defeat the new outer `On Error Resume Next` in release builds (errors after the inner block would propagate to break mode). Replaced the inner reset with `If DebugMode(False) Then On Error GoTo 0 Else On Error Resume Next` so it restores the outer mode rather than forcing GoTo 0. Future code that adds inner error scopes inside an outer-wrapped procedure should use the same conditional restore.

**What this rules out**:
- Any DAO recordset cleanup pattern that checks `.State`. Always use the cache-Err / Close / Err.Clear / re-raise idiom above. The same caveat applies to other DAO collection cleanup (`db.Close`, `qdf.Close`).
- Adding a `Name` / `SourceContext` property to `clsQueryComposer` solely for logging — `SqlSnippet()` plus the caller's existing log line is the agreed source of context.
- Wrapping every leaf parser helper. New leaf helpers added to `clsQueryComposer` should remain unwrapped; only new top-level stages (Parse*, Emit*, Build*Clause) need the pattern.
- Inner `On Error GoTo 0` resets inside any procedure that has the outer `DebugMode(True)` wrap. Use the conditional restore.

What would trigger revisiting: if a future composer rewrite collapses multiple stages into a single function, the per-stage logging granularity would be lost; the wrap pattern would need to be re-evaluated for that consolidated function.

**Relevant files**:
- `Version Control.accda.src/modules/Utility/clsQueryComposer.cls` — added `SqlSnippet()` helper; wrapped `ReconstructSQL`, `DecomposeSQL`, `IsDesignerCompatible`, `GenerateQdef`, `BuildFromClause`, `ConsolidateJoins`, `BuildJoinChain`, `ParseSelectQuery`, `ParseInsertQuery`, `ParseUpdateQuery`, `ParseDeleteQuery`, `ParseFromAndClauses`, `ParseFromExpression`, `ParseJoinExpression`, `EmitDesignViewQdef`, `EmitSqlViewQdef`, `EmitDbMemoSql`, `EmitAllProperties`, `EmitColumnMetadata`, `EmitDesignLayout`.
- `Version Control.accda.src/modules/Components/clsDbQuery.cls` — wrapped `IDbComponent_Export`, `IDbComponent_Import`, `ExportNewFormat`, `ImportNewFormat`, `ExportLegacy`; added cache-Err recordset cleanup in `ExportNewFormat`; added `Err.Clear` to `ImportNewFormat` cleanup so file-op errors don't leak to caller; replaced inner `On Error GoTo 0` in `ExportLegacy` with `DebugMode(False)` conditional restore.

---

## 2026-04-15 — Skip VCS index for MCP/API single-object imports (agent-as-user)

**Trigger**: Importing one query via `ImportObject` (MCP) in a large database with thousands of objects takes 7+ seconds, of which only 0.33s is the actual `LoadFromText`. The rest is JSON overhead: 5.19s parsing three JSON files (dominated by the 3.7 MB `vcs-index.json` with 21,747 ISO date entries), plus ~3-5s of hidden time re-serializing and saving the full index after `Perf.EndTiming`. The index exists for conflict detection — comparing source file state against database state across sessions — but for an MCP agent that just wrote the source file and is deliberately importing it, this check is meaningless.

**Options explored**:
- **Text-level parse-and-patch**: Read the index file as raw text, use `InStr` + brace-matching to locate the single relevant entry, parse/update only that ~200-byte snippet, splice it back. Would preserve full index consistency and conflict detection. Rejected as brittle — edge cases multiply (missing category sections, object deletions, new objects, comma handling, braces in strings). ~150 lines of new utility code for a narrow use case.
- **Lazy category-level parsing**: Parse only the index categories actually accessed during an operation (e.g., only "Queries" for a query import). Would benefit all operations. Rejected because the operations that use the full index (merge builds, full exports) must scan every category to determine what changed — lazy parsing provides no benefit. The only operation that touches a single entry is single-object import, which is better served by skipping the index entirely.
- **Disable index for MCP single-object imports (chosen)**: Treat the agent as a user making a direct edit. When a user modifies a query in the Access designer, there is no confirmation dialog — they save and that's the new state. The agent writing a source file and calling `ImportObject` is the same kind of deliberate action.

**Decision**: Added `Optional blnNoIndex As Boolean = False` parameter to both `LoadSingleObject` (imports) and `ExportSingleObject` (exports). When `True`, the VCS index is disabled (`VCSIndex.Disabled = True`) for the duration of the call — all index operations (`Update`, `Save`, `Item`, `CheckMergeConflicts`) become no-ops via existing guard clauses. Also skips the `Set VCSIndex = Nothing` / `Set Options = Nothing` reset and conflict detection block. `ImportObject` and `ExportObject` pass `blnNoIndex:=True` when `Operation.Source` is `eosMCPTool` or `eosExternalAPI`. Expected time drops from 7-12s (actual wall clock) to ~0.5s.

**What this rules out**: The index won't reflect MCP-imported objects until the next full export or merge build, which rebuilds index entries for all processed objects. A subsequent manual merge may see the imported object as "potentially modified" (stale/missing index data), but the content comparison during conflict detection will show source matches database, resolving without data loss. If index consistency for MCP operations becomes important, the text-level patching approach remains available as a future enhancement.

**Relevant files**:
- `Version Control.accda.src/modules/Core/modBuild.bas` — `LoadSingleObject`: new `blnNoIndex` parameter with conditional index/options/conflict skip
- `Version Control.accda.src/modules/Core/modExport.bas` — `ExportSingleObject`: same `blnNoIndex` pattern for single-object exports
- `Version Control.accda.src/modules/API/clsVersionControl.cls` — `ImportObject` and `ExportObject`: pass `blnNoIndex:=True` for MCP/API callers

---

## 2026-04-15 — Auto-resolve conflicts for agent/API operations

**Trigger**: When an MCP or API caller triggers a merge build or export that encounters conflicts (source file and database object both changed since last sync), the add-in opens the modal `frmVCSConflict` dialog and blocks indefinitely. Agents have no programmatic way to dismiss or respond to this dialog, causing the operation to hang.

**Options explored**:
- **Add a `McpConflictMode` option with three modes (fail, overwrite, prompt)**: Gives users fine-grained control. Initially planned but rejected as over-engineered — agents are automated callers by definition, and source files are already in Git, which provides the safety net for reviewing and reverting changes.
- **Return an error to the agent with conflict details**: Considered as the safest default. Rejected because it forces agents to handle a failure case that has no good resolution path — the agent would need to tell the user to open Access and resolve manually, defeating the purpose of automation.
- **Auto-resolve all conflicts (chosen)**: Treat agent operations the same way full builds are treated — just proceed. For imports, source is truth (overwrite DB objects). For exports, database is truth (overwrite source files). For deletes, proceed with the delete. The `ActionType` already set by `IsMergeConflict`/`IsExportConflict` carries the correct resolution.

**Decision**: Added `ResolveOrPrompt` method to `clsConflicts` that checks `Operation.Source`. For `eosMCPTool` or `eosExternalAPI`, it auto-resolves all conflicts using each item's `ActionType` (equivalent to clicking "Overwrite All" in the dialog). For user-initiated operations, it delegates to `ShowDialog` unchanged. All five call sites in `modBuild.bas` and `modExport.bas` now call `ResolveOrPrompt` instead of `ShowDialog` directly.

**What this rules out**: Agents cannot selectively skip individual conflicts — it's all-or-nothing overwrite. If a workflow emerges where agents need per-object conflict control, the `ResolveOrPrompt` method is the natural extension point. The three-mode option approach could be revisited if users report unwanted overwrites in practice.

**Relevant files**:
- `Version Control.accda.src/modules/Core/clsConflicts.cls` — new `ResolveOrPrompt` method
- `Version Control.accda.src/modules/Core/modBuild.bas` — 2 call sites updated
- `Version Control.accda.src/modules/Core/modExport.bas` — 3 call sites updated

---

## 2026-04-15 — Skip/auto-close UI for API and MCP-initiated operations

**Trigger**: When the MCP server calls `ImportObject` to merge a single component from source, `frmVCSMain` opens, becomes visible, and stays open — adding unnecessary overhead and requiring manual dismissal. Full builds and exports initiated by an agent also leave the form open after completion. The build confirmation dialog (`vbDefaultButton3` = Cancel) could block API callers entirely.

**Options explored**:
- **Set `InteractionMode = eimSilent` from MCP/API layer**: Leverages existing silent mode infrastructure. Rejected because (a) silent mode also suppresses `MsgBox2` dialogs with default button values, and the build confirmation dialog defaults to Cancel — which would cause API-initiated full builds to cancel themselves; (b) managing `InteractionMode` state across sync and async (timer-based) operations adds complexity; (c) the semantic of "silent" is about dialog suppression, not form visibility.
- **Add a new `ShowUI` / `Headless` flag**: Would work but adds new state to manage. Rejected because `Operation.Source` already distinguishes callers (`eosUserInterface`, `eosExternalAPI`, `eosMCPTool`) and is set reliably before any operation runs.
- **Use `Operation.Source` to control UI behavior**: Reuses existing infrastructure with no new state. Each UI decision point checks whether the caller is interactive. Chosen for simplicity and consistency.

**Decision**: Use `Operation.Source` checks at four specific points: (1) `ImportObject` skips `frmVCSMain` entirely for single-object imports — `LoadSingleObject` doesn't depend on the form; (2) `FinishBuild` auto-closes the form when source is API/MCP; (3) build confirmation dialog is skipped for API/MCP; (4) "Build Complete" MsgBox is suppressed for non-UI callers. Full builds/exports still show the form for progress visibility but auto-close on completion.

**What this rules out**: API/MCP callers cannot keep the form open after an operation to let the user review the log. If that's needed in the future, it would require a new parameter (e.g., `blnKeepOpen`) on the API methods. The `InteractionMode` mechanism remains available for other uses (e.g., truly silent batch processing from VBA scripts) and is unaffected by this change.

**Relevant files**:
- `Version Control.accda.src/modules/API/clsVersionControl.cls` — `ImportObject`: removed form open/close and redundant `Operation.Finish`
- `Version Control.accda.src/forms/frmVCSMain.cls` — `FinishBuild`: auto-close for API/MCP; `GetSourceFolder`: skip confirmation dialog for API/MCP
- `Version Control.accda.src/modules/Core/modBuild.bas` — `Build`: suppress "Build Complete" MsgBox for API/MCP

---

## 2026-04-15 — Use ByVal on clsVersionControl public API parameters for CallByName compatibility

**Trigger**: Calling `API("ExportObject", "query", "qryTest")` via `Application.Run` failed with a type mismatch. The `API` function receives arguments as `Variant` and forwards them through `CallByName`. VBA's default `ByRef` parameter passing requires an exact type match at the call site — a `Variant` cannot bind to a `ByRef String` parameter. The COM dispatch layer used by `CallByName` cannot coerce the type in-place.

**Options explored**:
- **Change method parameters to `Variant`**: Works, but loses type safety and makes the API less self-documenting. Callers inspecting the method signature can no longer see what type is expected.
- **Coerce arguments in `API()` with `CStr()`/`CLng()`**: Doesn't generalize — `API` is a generic dispatcher and different methods expect different types (String, Long, Boolean). Would require a mapping of method names to parameter types.
- **Replace `CallByName` with a `Select Case` dispatch**: Gives full control over coercion per method, but creates a maintenance burden — every new method on `clsVersionControl` requires a new `Case` branch.
- **Add `ByVal` to method parameters**: When a parameter is `ByVal`, VBA creates a local copy and performs implicit type coercion (Variant → String, etc.) automatically. No changes needed to the `API` function or `CallByName` call sites.

**Decision**: Add `ByVal` to all typed parameters on `clsVersionControl` public methods that are callable through `CallByName`. This is semantically correct (none of these methods modify their input parameters), backward compatible (existing direct callers are unaffected), and requires no changes to the dispatch infrastructure.

**What this rules out**: Methods on `clsVersionControl` that need to modify caller variables via `ByRef` would not work through the `CallByName` path. This is not a practical constraint — the API methods are input-only by design. If a future method genuinely needed `ByRef` semantics, it would need a different dispatch mechanism.

**Relevant files**: `Version Control.accda.src/modules/API/clsVersionControl.cls` — 9 methods updated, 12 parameters changed to `ByVal`.

---

## 2026-04-14 — Architectural principle: all external automation goes through the public API

**Trigger**: Adding 7 new methods to `clsVersionControl` for MCP tool support raised the question of where the boundary sits between the add-in's internal logic and what external consumers can reach. The MCP server, PowerShell scripts, and other VBA projects all need to call add-in functionality — should they use different entry points?

**Options explored**:
- **Let external tools call internal modules directly** (e.g., `Application.Run("Version Control.modExport.ExportSingleObject", ...)`). Gives maximum flexibility but exposes internal structure. Refactoring internal modules would break external callers. Internal functions often take object parameters (`AccessObject`, `IDbComponent`) that can't cross the COM boundary.
- **Separate API layer for MCP vs. ribbon vs. other callers**. Each consumer gets its own entry point optimized for its needs. Duplicates logic and creates divergent behavior.
- **Single public API on `clsVersionControl`, all consumers equal (chosen)**. Every external capability is a public method on `clsVersionControl`, callable via the `API()` dispatcher in `modAPI.bas` using `CallByName`. The MCP server, PowerShell, other VBA projects, and the ribbon all use the same methods. Internal modules remain `Option Private Module` and can be refactored freely.

**Decision**: All new external capabilities (`ExportObject`, `ImportObject`, `ExecuteSQL`, `RunVBA`, `GetOption`, `SetOption`, `GetLogContent`) are public methods on `clsVersionControl`. They accept only string/numeric parameters (COM-boundary safe) and return JSON strings for structured results. The `API()` function in `modAPI.bas` dispatches via `CallByName`, so new methods are automatically callable without modifying the dispatcher. `APIAsync` routes them through the sync `Case Else` branch since single-object operations don't need async/callback infrastructure.

**What this rules out**: No external tool can call internal functions like `ExportSingleObject` or `LoadSingleObject` directly — they must go through the public API wrappers that handle parameter resolution (string → `AccessObject`) and result serialization (errors → JSON). If a future capability requires passing objects across the COM boundary, it cannot use this pattern and would need a different approach (e.g., serializing the object identity as a string, which is what `ExportObject` already does with type + name).

**Relevant files**: `clsVersionControl.cls` (all public methods), `modAPI.bas` (`API()`, `APIAsync()`).

---

## 2026-04-14 — GetOption/SetOption: open dynamic dispatch via CallByName

**Trigger**: Agents need to read and write add-in options at runtime (e.g., `ShowDebug`, `McpAllowRunVBA`, `FormatSQL`) to control behavior during a session without modifying `vcs-options.json`. The question was how to expose option access through the API.

**Options explored**:
- **Switch statement whitelist**. `GetOption` maps known option names to specific property reads. Safe — only explicitly listed options are accessible. But requires updating the switch every time an option is added, and the add-in already has 30+ options with more planned.
- **Dictionary-based property bag**. Store options in a `Dictionary` instead of typed public properties. Makes dynamic access trivial but loses compile-time type safety, IntelliSense, and the established `clsOptions` pattern.
- **`CallByName` dynamic dispatch (chosen)**. `GetOption` calls `CallByName(Options, strName, VbGet)` and `SetOption` calls `CallByName(Options, strName, VbLet, varValue)`. Any public property on `clsOptions` is automatically accessible. Zero maintenance when adding new options.

**Decision**: `GetOption(strName)` and `SetOption(strName, varValue)` use `CallByName` for fully dynamic property access. Changes via `SetOption` are session-level — they take effect immediately but are not persisted to `vcs-options.json` until the user (or agent) explicitly saves. This lets agents freely adjust behavior (e.g., `ShowDebug`, `MaxLogFiles`) without risking permanent changes to project configuration.

**What this rules out**: There is no per-property access control. Every public property on `clsOptions` is readable and writable via the API, including the MCP security options themselves. An agent can call `SetOption("McpAllowRunVBA", True)` to enable `RunVBA` for the session, even if the persisted setting is False. This is acceptable because: (1) the change is session-level and doesn't persist, (2) the user has already opted into MCP access by configuring the server, and (3) adding a property-level ACL would require maintaining a second list of "allowed" properties alongside `m_colOptions`. If a property is ever added that should be truly non-modifiable via API (e.g., a licensing key), it should be implemented as a private property with a read-only public wrapper, which `CallByName` on `VbLet` would fail to set.

**Relevant files**: `clsVersionControl.cls` (`GetOption`, `SetOption`), `clsOptions.cls` (all public properties).

---

## 2026-04-14 — ExecuteSQL: add-in as a data access layer

**Trigger**: Agents frequently need to inspect database contents — `MSysObjects` for object inventory, `MSysQueries` for raw query definitions, table data for validation. The `db-inspector-mcp` server can do this via a separate ODBC/COM connection, but that requires a second MCP configured and opens a second connection to the same `.accdb` file, risking file-locking conflicts. Usage logs showed 67% of db-inspector calls were just SELECT queries.

**Options explored**:
- **Keep data access in db-inspector-mcp only**. Clean separation of concerns, but requires both MCPs configured for the common case. Two connections to the same file can conflict.
- **Add ODBC/pyodbc query execution in the VCS MCP server (Python-side)**. Avoids VBA roundtrip but creates a second database connection from Python. Would need to handle Access SQL dialect quirks in Python.
- **Route through the add-in's existing DAO connection (chosen)**. `ExecuteSQL` on `clsVersionControl` uses `CurrentDb.OpenRecordset` — the same connection the add-in already holds. No file-locking conflict. Access SQL handled natively. Results serialized as JSON with field names and values.

**Decision**: `ExecuteSQL(strSQL, lngMaxRows)` opens a read-only snapshot recordset, iterates rows up to the limit, serializes each row as a `Dictionary` (field name → value), collects into a `Collection`, and returns the whole result as JSON via `ConvertToJson`. Non-SELECT statements are rejected by checking the first token. The `McpAllowExecuteSQL` option (default: True) gates access. This expands the add-in's scope from "export/import engine" to "export/import engine + data inspection" — the first time the API returns raw query results rather than operating on database objects.

**What this rules out**: No write operations through `ExecuteSQL` — only SELECT. Agents needing INSERT/UPDATE/DELETE must use `RunVBA` or `CallVBA` with appropriate VBA code (gated by separate permissions). The SQL validation is intentionally simple (prefix check for `SELECT`) rather than parsing the full statement. A sufficiently creative agent could construct a SELECT with side effects (e.g., calling a VBA function from a query expression), but this is no different from the existing `RunVBA` risk and is gated by `McpAllowExecuteSQL`.

**Relevant files**: `clsVersionControl.cls` (`ExecuteSQL`), `clsOptions.cls` (`McpAllowExecuteSQL`).

---

## 2026-04-14 — Per-object API for MCP-driven development

**Trigger**: The VCS add-in's public API only supported whole-database operations (Export, Build, MergeBuild). AI agents using the MCP server had no way to export or import a single named object, which forced a full database export/import cycle for every iteration during development. Testing the query export refactoring against a ~3K query corpus required a tighter loop.

**Options explored**:
- **Add object_types filter to bulk export**. Would let agents export "just queries" but not a single named query. Still exports hundreds or thousands of objects per call. Not granular enough for the edit-import-compile-test loop.
- **Expose `ExportSingleObject` directly via Application.Run**. Not possible — it takes an `AccessObject` parameter, which can't be passed through `Application.Run` (only strings and numbers cross the COM boundary).
- **New `ExportObject`/`ImportObject` methods on clsVersionControl (chosen)**. Accept type string + name string, resolve to `AccessObject` or `IDbComponent` internally, delegate to existing `ExportSingleObject`/`LoadSingleObject`. Returns structured JSON with success/error status and log path. Works through the existing `API()` dispatcher via `CallByName`.

**Decision**: Added `ExportObject(strObjectType, strObjectName)` and `ImportObject(strObjectType, strObjectName)` to `clsVersionControl`. Both accept string parameters ("query", "form", "report", "module", "table", "macro") and return JSON results. They use the synchronous `API()` path since single-object operations are fast. The existing `APIAsync` `Case Else` branch routes them correctly without modification. A private `FindSourceFile` helper resolves source files for objects not yet in the database (new objects being imported for the first time).

**What this rules out**: These methods don't support bulk filtering (e.g., "all queries matching a pattern"). That would require a different approach — likely iterating through `GetAllFromDB` with a filter. For now, bulk operations use the existing `Export`/`MergeBuild` commands. If per-object operations prove too slow for large batches, the agent can fall back to bulk export and read the results from disk.

**Relevant files**: `clsVersionControl.cls` (`ExportObject`, `ImportObject`, `FindSourceFile`).

---

## 2026-04-14 — RunVBA: agent-generated code execution in temporary modules

**Trigger**: Agents testing VBA code (e.g., `clsQueryComposer` pipeline stages) needed a way to execute arbitrary VBA snippets and get results back without manually creating modules. The closed-loop debugging pattern — write code, run it, read result, iterate — required an API endpoint for ad-hoc VBA execution.

**Options explored**:
- **`Application.Eval` for expression evaluation**. Only handles single-line expressions, not multi-line statements. Can't declare variables, call methods with side effects, or build complex inspection logic.
- **Python-side VBE manipulation via COM**. The MCP server (Python) creates the temp module, compiles, runs, and cleans up using the VBE COM object model. Gives the Python layer full control but duplicates logic better handled in VBA, creates tight coupling to VBE internals, and can't leverage the add-in's error handling patterns.
- **Add-in manages the full lifecycle (chosen)**. A `RunVBA` method on `clsVersionControl` creates a temp module, wraps agent code in a function with error capture, compiles, executes, retrieves errors via accessor functions, removes the module, and returns JSON. The Python MCP layer just passes the code string through.

**Decision**: `RunVBA(strCode)` creates a temp standard module with three generated functions: `MCP_TempFunction` (wraps agent code with `On Error Resume Next`), `MCP_GetErrNum` and `MCP_GetErrDesc` (return captured error info via module-level variables). The error capture via accessor functions avoids the fragile alternative of embedding JSON string construction inside generated VBA code. Gated by `McpAllowRunVBA` option (default: False) — arbitrary code execution requires explicit user opt-in.

**2026-04-20 follow-up**: The original implementation never actually worked end-to-end. Several bugs had to be peeled apart in order:

1. **Identifier syntax**: The wrappers were declared as `_MCP_TempFunction`, `_MCP_GetErrNum`, `_MCP_GetErrDesc` and the temp module as `_MCP_Temp_<n>`. VBA's lexer rejects identifiers with a leading underscore in normal (unbracketed) declaration form, so every `vcs_run_vba` call returned a generic "VBA compilation failed" with no usable detail. Renamed all four identifiers to drop the leading underscore. No back-compat shim — the add-in is internal/pre-release, so callers were updated to the new names directly.
2. **`Option` statement collision**: New modules in Access auto-populate `Option Compare Database` (and sometimes `Option Explicit` depending on VBE settings). `InsertLines 1, ...` prepended our wrapper, leaving duplicate `Option` statements that triggered "Multiple Option Compare statements are not allowed". Now `DeleteLines 1, CountOfLines` clears the module before insertion.
3. **Wrong VBE active project for compile**: `acCmdCompileAllModules` only compiles the project currently active in the VBE, which (when the call originates from add-in COM) is the add-in itself. Our just-inserted temp module was in the host, never got compiled, and `Application.IsCompiled` returned False with no actionable error. Now we explicitly `Set VBE.ActiveVBProject = CurrentVBProject` before compiling. The "compileError" response also distinguishes between "host fails on its own" vs "wrapper itself fails" by re-checking compile state after removing the temp module.
4. **`Application.Run` qualifier syntax**: The qualified syntax for `Application.Run` is `[ProjectName].[FunctionName]` — module name is **not** supported, and unqualified calls from add-in code resolve against the add-in's own project (which doesn't have our temp module). The path-without-extension qualifier convention used in the reverse direction (`modAPI.GetRunCmdAddInFullLibName`) only works for *loaded library references* — the host project is not a loaded library from the add-in's perspective and Application.Run returns error 2517 ("cannot find the procedure") for the path form. The working qualifier is `CurrentVBProject.Name & "." & FunctionName`. From add-in code, `CurrentVBProject` points at the host project, and Access resolves the host's "Database" project ahead of the add-in's identically-named project (the inverse of the host→add-in collision documented in #593).
5. **Err leak through `Application.Run`**: The wrapper's `On Error Resume Next` captured the user-code error into `m_ErrNum`/`m_ErrDesc` but did not clear `Err` before returning. The raised error then propagated up through `Application.Run` to the caller, where it was indistinguishable from a real Application.Run failure (e.g., 2517 for an unresolved qualifier). Added `Err.Clear` at the end of `MCP_TempFunction` so that any error visible to the calling code is genuinely from `Application.Run` itself, while user-code errors are surfaced exclusively through `MCP_GetErrNum`/`MCP_GetErrDesc`.

The "compileError" return value now also includes the full generated wrapper source under `generatedSource` so future compile failures can be diagnosed directly from the tool result.

**What this rules out**: The agent's code runs with `On Error Resume Next` — it cannot use its own `On Error GoTo` handlers. If agent code needs structured error handling, it should use `CatchAny` or return error info through the function return value. The temp module is always removed, even on errors, so agent code cannot persist state between `RunVBA` calls (use `SetOption` or database tables for that).

**Relevant files**: `clsVersionControl.cls` (`RunVBA`), `clsOptions.cls` (`McpAllowRunVBA`).

---

## 2026-04-14 — MCP security options in clsOptions

**Trigger**: The new MCP tools (`RunVBA`, `ExecuteSQL`, `CallVBA`, `ImportObject`) have different risk profiles. Executing arbitrary agent-generated VBA code is fundamentally different from reading an option value. Users need granular control over what agents can do via MCP tools, and those controls should be discoverable in the existing options UI, not hidden in environment variables.

**Options explored**:
- **Environment variables only** (e.g., `ACCESS_VCS_ALLOW_VBA_EXEC=true`). Easy for CI/dev scenarios but invisible to users. Doesn't travel with the project. No UI discoverability.
- **Per-tool parameters** (e.g., `vcs_run_vba(db, code, allow=True)`). Agents would have to pass permission flags on every call, which is noisy and easily forgotten. Also pushes the security decision to the agent rather than the user.
- **Properties on `clsOptions` with defaults, serialized in `vcs-options.json` (chosen)**. Follows the same pattern as every other VCS option. Lives in the project's options file, visible in the options form. Environment variables can override for development scenarios.

**Decision**: Four boolean properties added to `clsOptions`: `McpAllowRunVBA` (default: False), `McpAllowExecuteSQL` (default: True), `McpAllowCallVBA` (default: True), `McpAllowImport` (default: True). Defaults follow least-privilege: read-like operations are on by default; arbitrary code execution is off. Properties are registered in `m_colOptions` for JSON serialization and excluded from `GetCategoryHashes` (they don't affect export output). The UI sub-form (`frmVCSOptionsMCP`) is deferred — the options are fully functional via `GetOption`/`SetOption` API.

**What this rules out**: Security is per-project, not per-session or per-agent. An agent connecting to a database with `McpAllowRunVBA = False` cannot escalate by setting it via `SetOption` because the check happens before the option can be changed in the same call. However, an agent *can* call `SetOption("McpAllowRunVBA", True)` to enable it for a session if `SetOption` itself isn't gated. This is acceptable because `SetOption` changes are session-level (not persisted) and the user has already opted into MCP tool access by configuring the MCP server.

**Relevant files**: `clsOptions.cls` (properties, defaults, `LoadDefaults`, `m_colOptions`, `GetCategoryHashes`).

---

## 2026-04-14 — SQL reconstruction fidelity: JOIN chain ordering and UNION handling

**Trigger**: After implementing the MSysQueries-based export (see "Deterministic query export with performance optimization" below), round-trip testing against real databases (MSysQueriesExamples, db-analysis-tools/sec) revealed that the reconstructed SQL differed from the COM `QueryDefs.SQL` property in JOIN nesting order and failed entirely for UNION queries.

**Options explored**:

- **Simple sequential JOIN emission (original)**: Emit joins in MSysQueries row order. Produced valid SQL but with different nesting than Access's own output. Differences in nesting can affect query plan and caused `.sql` vs `.com.sql` mismatches, making fidelity verification impossible.
- **Graph-based JOIN chain with DFS traversal** (chosen): Treat joins as a directed graph (leftTable → rightTable). Find the root table (appears only as leftTable, never as rightTable). DFS from root with deterministic sorting (INNER before LEFT/RIGHT, then alphabetical by rightTable) produces the same nesting order as Access's COM property. Handles star joins (multiple joins from same hub), self-joins (via aliases), and Cartesian products (no joins → comma-separated table list).
- **RIGHT JOIN normalization**: RIGHT JOINs are temporarily flipped to LEFT JOINs during graph traversal (so the hub table becomes the graph root), then restored to RIGHT JOIN syntax during emission. This avoids special-casing RIGHT JOINs in the graph algorithm.

**Decision**: `BuildJoinChain` uses DFS from the root table with `InsertJoinSorted` for deterministic child ordering. `ConsolidateJoins` merges multi-condition ON clauses (Access stores each condition as a separate Attribute 7 row) before traversal. For UNION queries, each segment is identified by its Attribute 5 `Name2` identifier (e.g. `X7YZ_____1`, `X7YZ_____2`); the SQL for each segment is reconstructed independently and joined with `UNION` or `UNION ALL` based on the Attribute 3 flag.

**What this rules out**: The reconstructed SQL must match Access's COM `QueryDefs.SQL` output in structure (not just semantics). Any future changes to `BuildJoinChain` or `ReconstructSQL` should be validated using `SqlBuilderValidation` (which writes diff artifacts under `logs/`). If Access changes its internal JOIN ordering algorithm, `BuildJoinChain` will need to be updated to match. *(The `.com.sql` per-query sidecar originally described here was removed — see 2026-05-19 entry.)*

**Relevant files**: `clsQueryComposer.cls` (`BuildJoinChain`, `BuildFromClause`, `ConsolidateJoins`, `DFSTraverse`, `InsertJoinSorted`).

---

## 2026-04-14 — Round-trip import with Design View / SQL View fallback

**Trigger**: Building databases from source with the new `.sql` + `.json` format revealed that some queries failed to import in Design View format (e.g. complex join topologies, non-equi-joins, subqueries). Additionally, alternate-path exports (used for merge conflict detection) were still using legacy `SaveAsText`, creating format mismatches.

**Options explored**:

- **Always import as SQL View**: Simple and reliable, but loses Design View layout (table positions, window dimensions) for queries that were saved in Design View. Users lose the visual layout they had before export.
- **Always import as Design View**: Fails for SQL-only query types (UNION, DDL, pass-through) and for queries with complex syntax that the designer cannot represent.
- **Attempt Design View, fall back to SQL View** (chosen): When layout data exists in the `.json` and `IsDesignerCompatible` returns True, generate a Design View `.qdef` and attempt `LoadFromText`. If import fails, regenerate as SQL View `.qdef` and retry. Log a warning so the user knows layout was lost. This preserves layout for the majority of queries while never failing outright.

**Decision**: `ImportNewFormat` attempts Design View first when conditions are met, then falls back to SQL View on failure. Alternate-path exports now route through `ExportNewFormat` when format version >= 5.0, producing `.sql` + `.json` instead of legacy `.qdef`. The `VBA Dim As New` anti-pattern (which caused "key already exists" errors in the column property loop because VBA scopes `Dim` to the procedure, not the block) was replaced with explicit `Set = New Dictionary` at the top of each loop iteration throughout all new code. *(The `.tmp`, `.failed.tmp`, and `.qdf` debug sidecar files originally described here were removed — see 2026-05-19 entry.)*

**What this rules out**: Queries imported via SQL View fallback lose their Design View layout permanently — the next export will have no `DesignLayout` in the `.json`. This is acceptable because the SQL itself is preserved faithfully. If a future Access update improves the designer's tolerance for complex SQL, the `IsDesignerCompatible` check could be relaxed to attempt Design View for more query types. The `ForceImportOriginalQuerySQL` option is only relevant to legacy `.qdef` imports and has no effect on the new format.

**Relevant files**: `clsDbQuery.cls` (`ImportNewFormat`, `IDbComponent_Export`), `clsQueryComposer.cls` (`IsDesignerCompatible`, `GenerateQdef`).

---

## 2026-05-19 — Remove query debug sidecar files from export/import

**Trigger**: When `ShowDebug` ("Show Detailed Output") was enabled, `ExportNewFormat` wrote `.qdf` and `.com.sql` sidecar files alongside each query's `.sql`, and `ImportNewFormat` wrote `.tmp` and `.failed.tmp` files preserving the generated `.qdef`. These files were intended for ad-hoc developer debugging during the early development of the deterministic query export pipeline. Turning `ShowDebug` off did not reliably clean them up (Fast Save skips unchanged queries, and `ShowDebug` is a non-export option that doesn't trigger category re-export).

Rather than building cleanup infrastructure for a feature that had outlived its purpose, the sidecar-writing code was removed entirely.

**Why removal instead of cleanup**: The dedicated testing tools — `SqlBuilderValidation` (writes artifacts under `logs/SqlBuilderValidation_*/`) and the round-trip harness (`modTestRoundtrip`, writes to `Testing/Fixtures/logs/`) — already produce their own diagnostic artifacts in gitignored locations. Per-query sidecars in the source tree were redundant with these tools and created a cleanup problem that no approach could solve cheaply without re-exporting all queries.

**Decision**: All `ShowDebug`-gated sidecar-writing code was removed from `ExportNewFormat` and `ImportNewFormat`. The `ShowDebug` option itself remains — it still controls verbose per-object logging throughout the codebase.

**Relevant files**: `clsDbQuery.cls` (`ExportNewFormat`, `ImportNewFormat`).

---

## 2026-04-14 — Column metadata and property serialization strategy

**Trigger**: The `.json` companion file needed a strategy for storing column-level metadata (AggregateType, ColumnWidth, ColumnHidden, Caption, etc.) parsed from the `MSysObjects.LvProp` binary blob. The format had to be deterministic for version control, compact for readability, and round-trippable back to `.qdef` format on import.

**Options explored**:

- **Store all properties with explicit type tags**: Every property gets a `{"Type": "dbLong", "Value": 123}` wrapper. Consistent but verbose — the majority of column properties are well-known types that don't need explicit tagging.
- **Store all properties as bare values**: Compact but loses type information for custom or unusual properties. On import, the code would have to guess the DAO data type, risking incorrect `.qdef` generation.
- **Known properties bare, unknown properties typed** (chosen): Properties with well-known names (`AggregateType`, `ColumnWidth`, `ColumnHidden`, `ColumnOrder`, `Caption`, `Description`, `TextAlign`, `DisplayControl`, `ResultType`, `CurrencyLCID`, `ShowDatePicker`, `IMEMode`, `IMESentenceMode`) are stored as bare values since their DAO types can be inferred from the name. Unknown or custom properties include an explicit type tag (e.g. `{"Type": "dbText", "Value": "..."}`). This keeps the common case compact while preserving full fidelity for edge cases.

**Decision**: `IsKnownColumnProperty` maps property names to the bare-value path; everything else goes through `DaoTypeToQdefPrefix` for explicit typing. `AggregateType = -1` is always emitted as a sentinel default (Access requires this property on every column in Design View `.qdef` files, even when no aggregation is used). Column metadata is sorted alphabetically by field name (`SortDictionaryByKeys`) for deterministic JSON output. The `clsLvPropParser` class (originally written for linked table LvProp blobs) was verified to work unchanged on query LvProp blobs — both use the same MR2 binary format with table-level and field-level property sections.

**What this rules out**: Adding a new known column property requires updating `IsKnownColumnProperty` in `clsDbQuery.cls` (and the corresponding import logic in `clsQueryComposer.GenerateQdef`). If a property name is ambiguous (same name, different types in different contexts), it must use the typed format. The alphabetical sort of columns means field rename operations will change the key ordering in the `.json`, producing a larger diff than strictly necessary — but this is acceptable for determinism.

**Relevant files**: `clsDbQuery.cls` (`IsKnownColumnProperty`, `DaoTypeToQdefPrefix`, column metadata loop in `ExportNewFormat`), `clsLvPropParser.cls` (shared MR2 parser), `clsQueryComposer.cls` (`GenerateQdef` column property emission).

---

## 2026-04-14 — Relax merge build gate to accept full export as baseline

**Trigger**: Merge builds were gated on `VCSIndex.FullBuildDate <> 0`, requiring a full build from source before merge was available. After index refactoring, full exports now populate the same per-component hashes (`FileHash`, `OtherHash`, `MetaHash`, `FilePropertiesHash`) that merge relies on. The `FullBuildDate` gate blocked a natural workflow: export from an existing database, pull source changes from Git, then merge those changes back in.

**Options explored**:
- **Keep the full-build-only gate**: Safe but overly restrictive. Forces users to do a throwaway full build before they can merge, even when they already have a working database with a complete index from export.
- **Remove the gate entirely (check only for non-empty index)**: Too permissive. A user who has never run the add-in would have no index at all, and merge would process every file as "modified" without proper dependency resolution.
- **Accept either `FullBuildDate` or `FullExportDate`**: Chosen. Both operations produce a complete index baseline. A full export from the existing database means the index and database are already in sync — exactly the state needed for merge to work correctly.

**Decision**: Changed the gate condition from `VCSIndex.FullBuildDate = 0` to `VCSIndex.FullBuildDate = 0 And VCSIndex.FullExportDate = 0`. The merge engine itself (`GetModifiedSourceFiles`) never checked `FullBuildDate` — it only needs index entries with `FilePropertiesHash` to diff against. This was purely a UI/API gate that no longer reflected a technical requirement.

**What this rules out**: The assumption that merge requires a prior full build is no longer valid. Future code should not re-introduce a `FullBuildDate`-only check. If a new component type is added that requires special handling on first import (like table data), it should be handled in the merge path's category filtering, not by gating on build history. Revisit if a scenario is found where export-generated index entries are insufficient for accurate merge detection.

**Relevant files**: `Version Control.accda.src/forms/frmVCSMain.cls` (gate condition and comment), `Version Control.accda.src/modules/API/clsVersionControl.cls` (user-facing message, added `T()` wrapping).

---

## 2026-04-09 — Stable, readable .env connection keys with named connection overrides

**Trigger**: Auto-generated `.env` keys for linked table connection strings used a hash of the full connection string (`conn_<hash>`). When developers worked across environments (e.g., local SQL dev vs. production server), different SERVER=, DRIVER=, or credential values produced different hashes — breaking the key mapping. Source files exported on one machine wouldn't resolve on another because the `env:conn_<hash>` reference pointed to a key that didn't exist in the other developer's `.env`.

**Options explored**:
- **Hash of full connection string (original)**: Simple and unique. Failed across environments because volatile parts (SERVER=, DRIVER=, UID=, PWD=) changed the hash. This was the behavior being replaced.
- **Hash of stable parts only**: Strip volatile parts before hashing. Still produces opaque keys. Considered but rejected in favor of readable keys.
- **Readable key from database identity**: Use the DATABASE= value (for ODBC) or the Access filename (for linked `.accdb`/`.mdb`) or DSN= as the key basis. Produces `conn_myappdb` instead of `conn_a3f72b1`. Chosen as the Tier 1 default — human-readable, stable across environments, and only falls back to hash when no identity can be extracted.
- **Include server/driver in key**: Would make keys environment-specific again. Rejected — the whole point is cross-environment stability.
- **User-configurable key composition**: Let users pick which parts (driver, server, database, table) form the key. Overcomplicated for minimal benefit. Rejected in favor of Tier 2 named connections.

**Decision**: Two-tier approach for `.env` connection key generation, implemented within the existing (unreleased) `EFV_5_0_0` gate:

**Tier 1 — Auto-generated readable keys** (`GetConnectionEnvKey`): Extract the database identity from the connection string — `DATABASE=` value for ODBC, `FSO.GetBaseName` for Access file paths, or `DSN=` as fallback. Run through `SanitizeKeyName` (lowercase, replace non-alphanumeric with underscores). Result: `conn_myappdb`. Falls back to `conn_<hash>` only when no identity is extractable.

**Tier 2 — User-defined named connections** (`EnvConnectionNames` in `vcs-options.json`): Users list key names (e.g., `["conn_production", "conn_warehouse"]`) in the shared options file. The actual connection strings live in each developer's `.env`. On export, `FindNamedConnectionKey` compares the live table's connection string against each named key's `.env` value using order-independent, case-insensitive parameter matching (`ConnectionParamsMatch`). Named connections are checked first in `ShouldUseEnvForConnection` and `SaveConnectionToEnv` tracks them but does not overwrite user-maintained `.env` values.

Key design choices:
- Auto-generated keys are always lowercase. User-defined keys preserve the user's original casing.
- `SanitizeKeyName` lowercases first, then replaces non-`[a-z0-9_]` characters with underscores.
- `ParseConnectionParams` splits connection strings into dictionaries with `TextCompare` for case-insensitive key lookup.
- No new export format version — changes are within the unreleased `EFV_5_0_0`.

**What this rules out**: Connection keys based on server name or driver version are intentionally excluded — the key must be the same regardless of where the database is hosted. If two different databases on the same server have the same DATABASE= name, they'll get the same auto-generated key and Tier 2 named connections must be used to disambiguate. Revisit if users report frequent collisions with common database names.

**Relevant files**:
- `Version Control.accda.src/modules/Utility/modConnect.bas` — `GetConnectionEnvKey` (rewritten), `SanitizeKeyName` (new), `ShouldUseEnvForConnection` (updated for Tier 2), `SaveConnectionToEnv` (updated to skip named connections), `FindNamedConnectionKey` (new), `IsDefinedConnectionName` (new), `ConnectionParamsMatch` (new), `ParseConnectionParams` (new)
- `Version Control.accda.src/modules/Infrastructure/clsOptions.cls` — `EnvConnectionNames` property (new Collection), serialization, loading, category hash

---

## 2026-04-09 — Filter auto-determined linked table properties at Standard sanitize level

**Trigger**: After implementing LvProp parsing, the exported JSON contained
significant noise from properties that Access auto-determines when linking a
table: `UnicodeCompression` (set per column type — True for nvarchar, False
for varchar/memo), `AppendOnly` (always False), and `TextFormat` (always 0 =
plain text). These appeared on nearly every text field but were never manually
customized.

**Options explored**:
- **Always include**: Safe but verbose. Every text field gets 1-3 extra
  properties that convey no user intent. Makes diffs noisy.
- **Always skip**: Cleanest output but removes information even when a user
  explicitly chose a non-default setting (rare but possible).
- **Skip at Standard sanitize level or above**: Matches the existing pattern
  for form/report sanitization. Users who set sanitize level below Standard
  retain full fidelity.

**Decision**: Gate these filters on `m_intSanitizeLevel >= eslStandard`.
`ParseLvProp` now accepts the sanitize level as a parameter. The three
properties are skipped only when at their default values (UnicodeCompression
is always skipped since its value is fully determined by the back-end column
type and cannot be predicted without schema knowledge). A block comment
explains the rationale so the filter can be revisited if a real use case for
preserving these emerges.

**What this rules out**: Users at Standard (the default) will not see these
properties in JSON. If someone discovers a scenario where manually overriding
UnicodeCompression on a linked table is meaningful, the filter should be
changed to skip only at the default value rather than unconditionally. Revisit
if bug reports mention missing UnicodeCompression after round-trip.

**Relevant files**:
- `modules/Utility/clsLvPropParser.cls` — `ShouldSkipFieldProperty`, new
  `m_intSanitizeLevel` member, `ParseLvProp` signature change
- `modules/Components/clsDbTableDef.cls` — passes `Options.SanitizeLevel`

---

## 2026-04-09 — Parse LvProp binary blob for linked table property export

**Trigger**: Issue #691 — linked table JSON files were missing front-end
display properties (column widths, lookup combos, captions, descriptions,
custom properties). The initial DAO property iteration approach worked but
was extremely slow: ~14ms per property read due to COM overhead, producing
2.65s+ per table and 15-28 minute exports for databases with 350+ linked
tables. ExportXML was tested but did not capture lookup/display properties
for linked tables.

**Options explored**:
- **DAO property iteration with blacklist filtering**: Worked correctly but
  inherently slow (~14ms per `.Value` access). Each property read triggers
  COM overhead and, for Access-linked tables, a round-trip to the back-end
  that causes a visible screen flash. Fails when back-end is offline.
  Implemented first, then abandoned.
- **DAO whitelist/direct-access**: Only read known properties by name.
  Benchmarks showed it was not reliably faster due to cold-cache effects on
  first `.Value` access. Loses unknown/custom properties.
- **Application.ExportXML**: Tested with `acExportAllTableAndFieldProperties`.
  Does not capture DisplayControl, RowSource, or other lookup properties for
  linked tables. Also fails when back-end is offline. Eliminated.
- **LvProp binary blob parsing**: The `LvProp` column in `MSysObjects` stores
  all locally-overridden properties in a binary TLV format. Sub-millisecond
  SQL read, works offline, captures everything including custom properties.
  Requires reverse-engineering an undocumented binary format.
- **Optional toggle (`SaveLinkedFieldProperties`)**: Added as a stopgap for
  the DAO approach to let users skip the slow export. Removed after LvProp
  eliminated the performance concern.

**Decision**: Read `LvProp` blob via SQL for export (sub-millisecond), parse
with `clsLvPropParser` (pure VBA byte math, no API calls). Write properties
via `SetDAOProperty` on import (safe, documented API). This asymmetry is
intentional: the undocumented blob format is safe to read but risky to write.

Key properties of the LvProp blob:
- Header: `MR2\0` magic + 4-byte dictionary size
- Dictionary section: property name table (2-byte len + UTF-16LE entries)
- Data section: field blocks (flag=1) and table block (flag=0)
- Each entry: `[2:size][1:flags][1:type][2:nameIndex][2:valLen][value]`
- ODBC-linked tables store ALL display/lookup properties locally
- Access-linked tables store only overrides (layout, custom); lookup defs
  are inherited from the back-end

The `SaveLinkedFieldProperties` option was removed since performance is no
longer a concern. The feature is always-on, gated only by
`ExportFormatVersion >= EFV_5_0_0`.

**What this rules out**: Any future change to the LvProp binary format by
Microsoft would break the parser. This is low risk — the format has been
stable across Access 2007-2021+. If it changes, the parser will fail
gracefully (MR2 magic check) and produce empty property sets rather than
corrupt data. Writing LvProp directly is explicitly ruled out in favor of
the DAO import path.

**Relevant files**:
- `modules/Utility/clsLvPropParser.cls` — new binary parser class
- `modules/Components/clsDbTableDef.cls` — export uses parser, import
  unchanged (DAO `SetDAOProperty`)
- `modules/Utility/modDatabase.bas` — `LongToSingle` helper + UDTs for
  IEEE float conversion (BackTint/BackShade properties)
- `modules/Infrastructure/clsOptions.cls` — `SaveLinkedFieldProperties`
  removed
- `forms/frmVCSOptionsAdvanced.cls` and `.form` — checkbox removed
- `vcs-options.json` — option entry removed
- `Issues/691.md` (msaccess-vcs-mgmt repo) — updated with LvProp findings

---

## 2026-04-03 — Template command bar unavailability is expected during consecutive add-in builds

**Trigger**: Running two consecutive "Build from Source" operations on the
add-in itself caused `Error 5: Invalid procedure call or argument` on the
second build. The error originated in `clsDbCommandBar.Class_Initialize`
at `Set m_TemplateCommandBar = Application.CommandBars(strTemplateCommandBarName)`.
The existing `On Error Resume Next` suppressed the runtime failure but
never cleared `Err`, so `LogUnhandledErrors` surfaced it later as an
unhandled error from an unknown source.

**Options explored**:

- **Restore the add-in's template after `ImportCommandBarsTemplate`**:
  After importing the template into the newly-built database, call
  `WizCopyCmdbars CodeProject.FullName` to reload the add-in's bars.
  Rejected: `WizCopyCmdbars` always imports into the *current* database,
  not the library database. The restored bar would still be associated
  with the current database and lost when it closes on the next build.
  This just repeats the delete/reimport cycle without fixing the root cause.

- **Try `WizCopyCmdbars` without pre-deleting**: Skip the delete loop and
  attempt import first; only delete-and-retry if it fails. Speculative:
  the existing comment says `WizCopyCmdbars` won't import when the name
  exists, and there's no API to distinguish which database owns a bar in
  `Application.CommandBars`, so selective deletion isn't possible.

- **On-demand recovery in `BuildControls`**: When the template is actually
  needed (custom built-in controls), attempt to reload from
  `CodeProject.FullName`. Would work within a single build but imports
  the add-in's bars into the user's database as a side effect. Also only
  needed for custom built-in controls, which the add-in itself doesn't use.

- **Consumer-side resilience** (chosen): Clear the error with `CatchAny`
  and log a diagnostic message. The original developer already anticipated
  this scenario (comment on lines 763-767) and used `On Error Resume Next`
  — the only bug was the missing error clear.

**Decision**: Catch and clear the expected error in `Class_Initialize`
using `CatchAny(eelNoError, vbNullString)` with a log-only diagnostic
message. The template command bar is only needed for importing custom
built-in controls, which the add-in itself doesn't use. Normal database
projects are unaffected because the add-in's template persists in
`Application.CommandBars` until `ImportCommandBarsTemplate` runs (an
add-in-specific `AfterBuild` hook that deletes all instances before
reimporting via `WizCopyCmdbars`).

**What this rules out**: This does not fix the underlying limitation that
`ImportCommandBarsTemplate` permanently removes the add-in's template
from `Application.CommandBars` for the rest of the session. If a user
database with custom built-in command bar controls is built immediately
after building the add-in (without restarting Access), those controls
would fail to import. Revisit if that scenario is reported, or if Access
exposes an API to reload a library database's command bars without
closing/reopening it.

**Relevant files**: `Version Control.accda.src/modules/Components/clsDbCommandBar.cls`

---

## 2026-04-03 — CloseCurrentDatabase2 retries internally; ReleaseDbReferences for shared mode reopen

**Trigger**: After the shared mode reopen at the end of a build, the
navigation pane was missing and consecutive build operations triggered
VBA errors. Diagnostic logging revealed `DatabaseFileOpen=True` after a
single `CloseCurrentDatabase2`, indicating the database was not fully
closing (same pattern as the full build's exclusive-mode close).

**Options explored**:

- **Caller-side retry** (initially chosen, then improved): Each call site
  checks `If DatabaseFileOpen Then CloseCurrentDatabase2` after the first
  call. This worked but was error-prone — forgetting the check at any call
  site would leave the database open. The full build and shared mode
  reopen both had this pattern, but merge reopen, theme reopen, and
  `ShiftOpenDatabase` did not.

- **Full `ReleaseObjects` teardown**: Clear all singletons before close.
  Too aggressive: destroys `Log`, `Perf`, `Options` and other state
  needed for the remainder of the build.

- **Internal retry in `CloseCurrentDatabase2`** (chosen): Move the
  `If DatabaseFileOpen Then` retry into the function itself. All call
  sites benefit automatically with no code duplication.

**Decision**: `CloseCurrentDatabase2` in `modWizHook.bas` now checks
`DatabaseFileOpen` after the first close and retries if needed. Removed
redundant retry checks from `modBuild.bas` (full build and shared mode
reopen blocks). Also added `ReleaseDbReferences` to `modObjects.bas`
(clears only `this.dbs`) called before the shared mode reopen close to
prevent stale cached `CurrentDb` references.

**What this rules out**: The consecutive-build VBA errors are a separate
issue (template command bar lifecycle, see entry above) and not caused by
dangling `SharedDb` references. `ReleaseDbReferences` is narrowly scoped
to database-bound singletons; expanding it to clear FSO or other
non-database singletons is unnecessary.

**Relevant files**: `Version Control.accda.src/modules/Utility/modWizHook.bas`,
`Version Control.accda.src/modules/Infrastructure/modObjects.bas`,
`Version Control.accda.src/modules/Core/modBuild.bas`

---

## 2026-04-03 — Worker WaitForQueue must use tight DoEvents loop, not Sleep

**Trigger**: The `DoEvents` polling loop in `clsWorker.WaitForQueue` spins
thousands of iterations per second, raising concerns about CPU churn and
reentrancy. Adding `Sleep 100` (kernel32) between `DoEvents` calls seemed
like a safe way to yield CPU while still allowing queued COM callbacks to be
dispatched on the next `DoEvents`.

**Options explored**:
- **Sleep 100ms between DoEvents calls**: Tested in practice. Reduced CPU
  usage but increased `Wait for Job Queue` time from ~0.8s to 5.3s (6–7x
  slower). The root cause: the worker VBScript makes many individual COM
  calls back into Access during execution (property access on `objApp.VBE`,
  iterating `VBProjects`, `GetObject`, etc.), not just one final callback.
  Each COM call is marshaled through the STA message queue and blocks until
  Access processes it via `DoEvents`. With `Sleep 100` between pumps, every
  round-trip adds ~100ms latency. With 40–50 round-trips, this compounds to
  4–5 seconds of added wait time.
- **`MsgWaitForMultipleObjects` API**: Would yield CPU like `Sleep` but wake
  on incoming messages. Complex to wire up and would need careful testing
  with VBA's message pump. Not attempted — tight `DoEvents` loop is already
  fast enough for the sub-second operations involved.
- **Tight `DoEvents` loop (original design)**: Keeps the message pump
  responsive to all inbound COM calls with near-zero latency. Higher CPU
  usage but the total wait is typically under 1 second, so the window of
  elevated CPU is brief.

**Decision**: Keep the tight `DoEvents` loop. The worker VBScript's many
COM round-trips into the host application make any message pump delay
multiplicative. Updated the loop comment to document this constraint so
future contributors don't re-attempt the Sleep approach.

**What this rules out**: Adding `Sleep` or any blocking wait inside the
`WaitForQueue` polling loop. This could be revisited if the worker script
were restructured to batch COM calls or minimize round-trips (e.g.,
collecting all needed data in a single `objApp.Run` call and doing work
locally in VBScript). That would reduce the number of COM calls that need
pump dispatch.

**Relevant files**: `Version Control.accda.src/modules/Integration/clsWorker.cls`

---

## 2026-04-02 — Out-of-process worker probe for post-build database lock

**Trigger**: After a build or merge, external clients (MCP tools, ODBC connections) receive JET/ACE error 3734: "The database has been placed in a state by user 'Admin' on machine '...' that prevents it from being opened or locked." The database is unusable to other clients until manually closed and reopened in Access. This blocks automated workflows that query the database immediately after a build.

**Options explored**:
- **Win32 file lock check (`IsFileOpenExclusive` via `CreateFileW`)**: Tried first. The OS-level file is not exclusively locked — the issue is an engine-internal state flag set during DDL/schema operations (importing forms, tables, queries). This check always reported the file as accessible even when external clients were blocked. Removed.
- **In-process DAO probe (`DBEngine.OpenDatabase` from the add-in)**: Tried second. The same JET/ACE engine instance allows intra-process connections even when blocking external clients. Confirmed by running an identical DAO test from Excel VBA (out-of-process), which correctly detected the block while the in-process check passed. Removed from `modDatabase.bas`.
- **Always close/reopen unconditionally**: Simple and guaranteed to work, but pays the time cost (several seconds) on every build/merge even when the database is already accessible.
- **Out-of-process worker probe via `clsWorker` VBScript**: Launches the existing worker script infrastructure, which runs as a separate process with its own `DAO.DBEngine.120` instance. Accurately detects the engine-level lock state. Only triggers the close/reopen when actually needed.

**Decision**: Use the out-of-process worker probe. Added a `CheckDatabaseAccessible` action to the worker script that creates an independent `DAO.DBEngine.120` via `CreateObject` and attempts `OpenDatabase(path, False, True)`. The add-in calls `Worker.IsDatabaseAccessible` which launches the worker, waits for the callback via `WaitForQueue`, and reads the result from `m_varLastResult`. If inaccessible, the existing `StageMainForm`/`CloseCurrentDatabase2`/`ShiftOpenDatabase`/`RestoreMainForm` pattern reopens the database in shared mode. The trade-off is a brief VBScript launch overhead (~1s) on every build/merge to run the probe, but this avoids the heavier close/reopen cycle when it isn't needed.

**What this rules out**: In-process detection of this engine-level lock state — the JET/ACE engine does not expose the DDL state flag to same-process callers. Any future attempt to detect this condition must use an out-of-process mechanism. If the worker script infrastructure is ever removed, this check would need to fall back to always closing/reopening unconditionally.

**Relevant files**:
- `Version Control.accda.src/modules/Integration/clsWorker.cls` — `IsDatabaseAccessible` method (add-in side), `CheckDatabaseAccessible` function (worker script side), `m_varLastResult` for callback return values, updated `ReturnWorker` to store results
- `Version Control.accda.src/modules/Core/modBuild.bas` — post-build/merge reopen block uses `Worker.IsDatabaseAccessible`

---

## 2026-03-27 — Enforce canonical add-in filename and fix .accde path bugs

**Trigger**: Issue #693 reported that renaming `Version Control.accda` to a different filename causes error 2517 at runtime. Investigation revealed two problems: (1) `GetAddInFileName` dynamically derived the installed filename from `CodeProject.Name`, so a renamed file would install under the wrong name and break the COM ribbon DLL's hardcoded `Application.Run` calls; (2) several comparison and loading spots always assumed the `.accda` extension, silently failing when the compiled `.accde` version was installed.

**Options explored**:
- **Make the ribbon DLL discover the .accda name dynamically** (e.g., from a registry key or by scanning the install folder). Would support arbitrary filenames, but adds complexity for no compelling use case — users who want to test different versions can build/install from different branches.
- **Keep `GetAddInFileName` dynamic but add runtime validation**. Would catch the mismatch later. Rejected because the root issue is that the filename is a contract between three components (VBA add-in, COM ribbon DLL, worker scripts), and allowing divergence invites breakage.
- **Replace dynamic derivation with a constant, block renamed files at install time (chosen)**. New `ADDIN_BASENAME` constant in `modConstants.bas`. `GetAddInFileName` uses it instead of `CodeProject.Name`. Installer checks the filename up front and shows a clear error. Simple, explicit, and aligns all components on the same name.
- **For the .accde bug: change `GetAddInFileName`'s default to respect `blnUseCompiledAddIn`**. Would fix comparisons but break `UpdateAddInFile`, which uses explicit `.accda`/`.accde` paths for cleanup during install transitions.
- **Add `GetInstalledAddInFileName` helper (chosen)**. Delegates to `GetAddInFileName(GetInstallSettings.blnUseCompiledAddIn)`. Non-install callers use this; install logic continues using `GetAddInFileName` with explicit extension control. Clean separation.

**Decision**: `ADDIN_BASENAME` constant enforces the canonical name. `GetInstalledAddInFileName` returns the correct `.accda`/`.accde` path based on persisted install settings. All comparison/loading spots (`AutoRun`, `GetAddInProject`, `LoadVCSAddIn`, `RegisterMenuItem`, `RelaunchAsAdmin`, `Run_UninstallAddin`, `frmVCSInstall`, `frmVCSOptionsTranslation`) use the new helper. `clsWorker.GetAddInVBProject` compares by base name only (no extension) since it runs in VBScript without access to VBA constants. The add-in filename is now a fixed contract — renaming it requires changing one constant plus rebuilding the twinBASIC ribbon DLL.

**What this rules out**: The add-in filename can no longer be set dynamically by renaming the `.accda` file. If the project ever renames the add-in (e.g., from "Version Control" to "MSAccessVCS" for v5), only the `ADDIN_BASENAME` constant and the ribbon DLL's `strAddInLib` need updating. A `RunUpgrades` migration step would handle the transition for existing installs. The naming discussion is open but deferred — v5 would be the appropriate time.

**Relevant files**: `modConstants.bas` (new `ADDIN_BASENAME`), `modInstall.bas` (install guard, `GetAddInFileName` rewrite, `GetInstalledAddInFileName`, 7 caller updates, `RunUpgrades` legacy path fix), `modVbeUtility.bas` (`GetAddInProject`, `LoadVCSAddIn`), `modAPI.bas` (`GetRunCmdAddInFullLibName` rewrite, example functions), `clsWorker.cls` (`GetAddInVBProject`, `Run_UninstallAddin`), `frmVCSInstall.cls`, `frmVCSOptionsTranslation.cls`.

---

## 2026-03-19 — Layout SVG: subform, tab control, and hidden control rendering strategies

**Trigger**: When generating SVG from form source files, three control types require non-obvious rendering decisions because they involve content that may not be visible, may live in separate source files, or may vary at runtime. Each choice affects what an AI agent can "see" in the SVG and how closely the SVG matches a screenshot.

**Options explored**:

*Subforms:*
- **Embed subform SVG inline**: Would give agents a complete picture in one file, but subform source objects are often swapped at runtime, and embedding creates coupling between independently versioned files. A change to the subform would require regenerating the parent SVG.
- **Render as labeled placeholder box** (chosen): Dashed border with `[Subform: Name]` label. Each subform generates its own independent `.svg` alongside its own `.form` file. Agents can cross-reference by name. This matches the existing component model where subforms are independent `IDbComponent` objects.
- **Link via SVG `<use>` or `<image>` reference**: Would allow lazy composition but adds fragile path dependencies and complicates standalone viewing.

*Tab controls:*
- **Render all pages stacked vertically**: Would show all content but produces an SVG that doesn't match any real visual state of the form — confusing for screenshot comparison and spatially misleading since controls on different pages occupy the same coordinates.
- **Generate multiple SVGs per form** (one per tab page): Comprehensive but multiplies output files, complicates the file naming convention, and doesn't reflect what a user actually sees.
- **Render only the first visible/default page** (chosen): Matches the most common runtime state. Controls on other pages are omitted. This is the simplest approach and produces an SVG that corresponds to what a user sees when opening the form. If hidden-page content becomes important, a future option could render all pages as separate SVGs.

*Hidden controls (Visible = NotDefault):*
- **Omit entirely**: Cleanest SVG but loses structural information — an agent wouldn't know the control exists, which matters for layout analysis (e.g., controls that toggle visibility at runtime still occupy design-time space).
- **Render at reduced opacity** (chosen, opacity 0.3): Preserves positional information while visually distinguishing hidden controls. Agents can see where hidden controls sit relative to visible ones. A future option could toggle between omit/transparent/full rendering.
- **Render normally with a metadata attribute**: Would require agents to parse SVG attributes rather than relying on visual inspection, which defeats the purpose of a visual representation.

**Decision**: Subforms as independent placeholders, first tab page only, hidden controls at 30% opacity. All three choices prioritize a clean visual that matches the default runtime appearance while preserving enough structural information for layout analysis.

**What this rules out**: Agents cannot see controls on non-default tab pages or the actual content of subforms from the parent SVG alone. Revisit if agents frequently need cross-page or cross-subform layout analysis — the most likely extension would be an option to render all tab pages as separate named SVG groups or files.

**Relevant files**:
- `Version Control.accda.src/modules/Core/clsFormLayoutSvgWriter.cls` — `RenderTabControl` (first page only), `RenderSubform` (placeholder), `RenderControl` (opacity check)

---

## 2026-03-19 — Form/report layout SVG export from SaveAsText source files

**Trigger**: AI agents can perform major code refactors but struggle with Access form layout design because `.form` files are hard to reason about structurally. An SVG representation of the layout — generated deterministically from exported source files — gives agents a visual artifact they can interpret, enabling them to identify and suggest layout improvements. Future work will pair this with an MCP server to apply layout changes via VBA scripts in design view.

**Options explored**:
- **MSXML2.DOMDocument60 for SVG output**: DOM provides structural correctness guarantees but has per-element COM overhead. Since SVG generation is write-only (no querying or transforming), DOM's overhead provides no benefit. Not chosen.
- **clsConcat (paged Mid$ buffer)**: O(n) string assembly, already proven fast in the codebase. Chosen for SVG output with a small `EscapeXml()` helper for text content.
- **Single monolithic class vs pipeline of specialized classes**: A pipeline (parser → theme resolver → SVG writer) was chosen for separation of concerns and independent testability. Each class has a clear responsibility and can be extended without touching the others.
- **Call site in SaveComponentAsText (DRY) vs component Export methods (contextual clarity)**: Hybrid chosen — shared implementation in `modFormLayoutSvg.TryExportLayoutSvg`, called from `clsDbForm.IDbComponent_Export` and `clsDbReport.IDbComponent_Export` after `SaveComponentAsText` returns.
- **Theme color extraction via ExtractFromZip**: Initial implementation used the existing `ExtractFromZip` function, which has a broken exit condition when the destination folder is non-empty (it polls for 60 seconds until timeout). Replaced with a targeted `Shell.Application.CopyHere` of just the `theme` folder, polling for the specific output file with 0.1s intervals and a 10s timeout. Extracted files are cached in a stable temp folder keyed by theme name, so subsequent exports skip extraction entirely.

**Decision**: Four new VBA classes (`clsLayoutNode`, `clsFormLayoutParser`, `clsFormLayoutThemeColors`, `clsFormLayoutSvgWriter`) plus an orchestrator module (`modFormLayoutSvg`). Gated by `Options.ExportLayoutSvg` (default False). SVG is indented for version-control-friendly diffs. Coordinates use twips-to-CSS-px at 96 DPI ("Universal" mode). `LAYOUT_SVG_GENERATOR_VERSION` constant enables cache invalidation when the generator changes.

Key learnings from initial testing:
- SaveAsText nests sections and controls inside anonymous `Begin`/`End` ("Defaults") blocks — tree traversal must recurse through these to find sections and their child controls.
- Control-associated labels (e.g. checkbox labels) are children of the parent control node, not siblings at the section level — rendering must descend into control children after drawing the control itself.
- `Dir$` is unsafe with Unicode paths in this project; all file/folder iteration must use FSO (`Folder.SubFolders`, `Folder.Files`).
- Disabling the option cleans up existing `.svg` files on next export rather than leaving stale artifacts.

**What this rules out**: SVG generation is purely from exported text files — it does not open the `.accdb` at runtime, so it cannot capture runtime-only visual state (conditional formatting, VBA-driven visibility). Revisit if screenshot-based validation shows major gaps that can only be resolved with runtime data. The `"Screenshot"` scale mode option is stubbed but not yet differentiated from `"Universal"`.

**Relevant files**:
- `Version Control.accda.src/modules/Core/clsLayoutNode.cls` — tree node with ControlType, Props dictionary, Children collection
- `Version Control.accda.src/modules/Core/clsFormLayoutParser.cls` — line scanner producing node tree from `.form`/`.report` files
- `Version Control.accda.src/modules/Core/clsFormLayoutThemeColors.cls` — resolves theme color indices to RGB hex via `.thmx` extraction and HSL tint/shade math
- `Version Control.accda.src/modules/Core/clsFormLayoutSvgWriter.cls` — depth-first tree walk emitting SVG via clsConcat
- `Version Control.accda.src/modules/Core/modFormLayoutSvg.bas` — orchestrator: TryExportLayoutSvg, theme cache management
- `Version Control.accda.src/modules/Infrastructure/clsOptions.cls` — ExportLayoutSvg, LayoutSvgImageEmbed, LayoutSvgScaleMode options
- `Version Control.accda.src/modules/Infrastructure/modConstants.bas` — LAYOUT_SVG_GENERATOR_VERSION constant
- `Version Control.accda.src/modules/Components/clsDbForm.cls` — SVG call site and .svg cleanup
- `Version Control.accda.src/modules/Components/clsDbReport.cls` — SVG call site and .svg cleanup

---

## 2026-03-19 — Options form redesign: tabbed interface → left-nav with subform-per-section

**Trigger**: The existing options form used a tabbed interface (`pagGeneral`, `pagExport`, etc.) with some pages hidden. This constrained screen real estate, made it difficult to add descriptive text alongside options, and required users to discover hidden pages. A left-navigation + scrollable detail section is the standard pattern in modern applications.

**Options explored**:
- **Single scrollable form with show/hide frames**: One subform containing all options, with frames toggled visible/hidden based on navigation selection. Simplest code, but Access has no way to limit scrolling to only the visible section — the user would scroll past large hidden gaps. Rejected.
- **Subform-per-section with dynamic SourceObject**: Each section is a separate form loaded into a single subform control on the main form. True scroll containment per section, independent layout, and modular code. Higher initial cost (8 subforms + interface), but better long-term maintainability. Chosen.

**Decision**: Main form (`frmVCSOptions`) holds a private `m_Options As clsOptions` working copy, an option group (`fraNav`) with toggle buttons (`tglGeneral`, `tglExport`, etc.), and a subform control (`subOptionsDetail`). Navigation derives the target form name by stripping the `tgl` prefix from the selected toggle button's name (translation-safe — not dependent on display text). `IOptionsSection` interface enforces `LoadOptions`/`SaveOptions` on all 8 section forms. Each subform's `Form_Load` calls `LoadOptions`; `SaveCurrentSubform` calls `SaveOptions` via the interface before switching sections. Changes are committed only on "Save & Close" (`Set Options = m_Options` + `Options.SaveOptionsForProject`); Cancel discards `m_Options` by closing.

The subform control's `SourceObject` is left blank at design time. The main form's `Form_Load` initializes `m_Options` first, then sets `SourceObject` via `fraNav_AfterUpdate`, avoiding the chicken-and-egg problem where a subform's `Form_Load` fires before `m_Options` is ready.

Registry-based settings (Diff Tool, Open Repository) use deferred save via public properties on the main form (`DiffTool`, `OpenRepository`). The General subform reads/writes these properties; the main form commits them to the registry in `cmdSaveAndClose_Click`. This keeps registry settings consistent with the deferred-save pattern of `clsOptions` properties.

External database schemas use a shared dictionary bridge: `frmVCSOptionsDatabases.LoadOptions` clones schemas into `Form_frmVCSOptions.DatabaseSchemas` and points its private `m_dSchemas` at the same object. This allows `frmVCSDatabase` (the add/edit popup) to write directly to the dictionary that `RefreshSchemaList` reads from.

**Sections**: General (export folder, tools, language), Export (source files, sanitization, content, printer settings, hooks), Tables & Data (table data export selection), External Databases (schema connections), Build (build/merge behavior, hooks), Translation (contribute, path, sync), Defaults (project defaults, read-only install settings), Advanced (debugging, hashing, export tweaks, logging).

**What this rules out**: The tabbed interface pattern is retired for the options form. All new options must be added to the appropriate section subform's `LoadOptions`/`SaveOptions` and the corresponding form layout. Adding a new section requires: (1) create `frmVCSOptionsXxx.cls` implementing `IOptionsSection`, (2) create `frmVCSOptionsXxx.form`, (3) add `tglXxx` toggle button to `fraNav` on the main form. The toggle button naming convention (`tgl` prefix mapping to `frmVCSOptions` + suffix) is load-bearing — changing it breaks navigation.

**Relevant files**:
- `Version Control.accda.src/forms/frmVCSOptions.cls` — main form orchestrator
- `Version Control.accda.src/modules/Interfaces/IOptionsSection.cls` — LoadOptions/SaveOptions interface
- `Version Control.accda.src/forms/frmVCSOptionsGeneral.cls` — General section
- `Version Control.accda.src/forms/frmVCSOptionsExport.cls` — Export section
- `Version Control.accda.src/forms/frmVCSOptionsTableData.cls` — Tables & Data section
- `Version Control.accda.src/forms/frmVCSOptionsDatabases.cls` — External Databases section
- `Version Control.accda.src/forms/frmVCSOptionsBuild.cls` — Build section
- `Version Control.accda.src/forms/frmVCSOptionsTranslation.cls` — Translation section
- `Version Control.accda.src/forms/frmVCSOptionsDefaults.cls` — Defaults section
- `Version Control.accda.src/forms/frmVCSOptionsAdvanced.cls` — Advanced section

---

## 2026-03-19 — Install settings displayed as read-only on options form

**Trigger**: The Defaults section of the new options form displays installation settings (install folder, trust folder, use ribbon, compile accde, open after install). These are registry values set during the `InstallVCSAddin` process. The question was whether to make them editable from the options form.

**Options explored**:
- **Editable with deferred registry save**: Let users change values, save to registry on "Save & Close." Problem: the settings only take effect during installation (file copy, COM registration, trust location setup). Saving registry values without applying them would mislead users into thinking the change took effect. Rejected.
- **Editable with immediate apply (trigger reinstall)**: Apply changes by invoking `InstallVCSAddin`. Problem: the add-in cannot reinstall itself while loaded — it would require a VBScript worker process to close Access, copy files, and reopen. Over-engineered for a rarely-needed operation. Rejected.
- **Read-only display with guidance to reinstall**: Show current values as locked/disabled controls with a label explaining these are set during installation. Users see their current configuration without confusion. The dedicated `frmVCSInstall` form handles changes through the proper install flow. Chosen.

**Decision**: Controls are displayed read-only (locked/disabled at the form layout level). `SaveOptions` is intentionally empty — these settings are not part of the deferred-save flow. If reinstalling from the options form becomes a frequent user need, a VBScript-based reinstall mechanism could be added, but this is deferred until there's evidence of demand.

**What this rules out**: Install settings cannot be changed from the options form. The `frmVCSInstall` form remains the only supported path for changing install configuration. If a future version adds a "Reinstall" button, it would need to handle the add-in-loaded constraint (likely via an external VBScript worker that closes Access, copies files, and reopens).

**Relevant files**:
- `Version Control.accda.src/forms/frmVCSOptionsDefaults.cls` — read-only load, empty SaveOptions

---

## 2026-03-18 — Standardize Letter Casing ribbon command with user feedback and template creation

**Trigger**: The `StandardizeLetterCasing` feature (Mike Wolfe's technique, integrated in the add-in) ran silently during export and build with no way for a user to invoke it on demand. Users who didn't already have a `clsStandardLetterCasing` module in their project had no discoverability path to the feature.

**Options explored**:
- **Boolean return from StandardizeLetterCasing**: Function returns True (found) / False (not found). Simple, but doesn't tell the user whether corrections were actually made or casing was already consistent. Rejected.
- **Long return with sentinel (-1 = not found, 0 = no corrections, 1+ = count)**: Single return value conveys both status and count. Existing callers that ignore the return value are unaffected (VBA ignores function return values when called as a Sub). Chosen.
- **Separate ByRef parameter for count**: Cleaner separation of concerns but more complex call site and requires all callers to pass a variable even if they don't care. Rejected.

**Decision**: Changed `StandardizeLetterCasing` from `Sub` to `Function ... As Long` returning -1 (module not found), 0 (already consistent), or the correction count. Added a `lngCorrections` counter incremented at both `cm.ReplaceLine` call sites (Dim lines and API declares). The ribbon command in `clsVersionControl` uses a `Select Case` on the return value to show three distinct `MsgBox2` messages. When the module is not found, the user is prompted (Yes/No) to create a starter template. If they accept, `CreateLetterCasingTemplate` creates the class module via `CurrentVBProject.VBComponents.Add(vbext_ct_ClassModule)`, inserts a header and example Dim lines via `CodeModule.InsertLines`, shows a confirmation message, and opens the module in the VBE with `DoCmd.OpenModule`. No second prompt before opening — the user just opted in, so navigating directly is the natural next step.

The ribbon button (`btnStandardizeLetterCasing`) is placed in the Advanced Tools menu before Reload Ribbon, using the `ChangeCaseDialogClassic` imageMso icon. Wiring is automatic via the existing `CallByName VCS, Mid(strCommand, 4)` routing in `modAPI.HandleRibbonCommand`.

**What this rules out**: The `-1` sentinel means future callers must not use negative counts for other purposes. If more granular status is needed (e.g., distinguishing "module exists but empty" from "module exists with rules"), the return value scheme would need rethinking — but the current three states cover all practical scenarios. The template content is hardcoded in `CreateLetterCasingTemplate`; if the canonical template format changes, this code must be updated manually.

**Relevant files**:
- `Version Control.accda.src/modules/Core/modLetterCasing.bas` — `Sub` → `Function As Long`, counter, sentinel return
- `Version Control.accda.src/modules/API/clsVersionControl.cls` — `StandardizeLetterCasing` with `Select Case` feedback, `CreateLetterCasingTemplate` private helper
- `Version Control.accda.src/modules/Install/modRibbonStrings.bas` — label and description for `btnStandardizeLetterCasing`
- `Ribbon/Ribbon.xml` — button definition in `mnuAdvancedTools` menu

---

## 2026-03-17 — Secure connection string storage via .env file references

> **⚠ Partially superseded** (2026-04-09): The key generation algorithm (`GetConnectionEnvKey`) was rewritten to produce readable, environment-stable keys instead of hashes. A second tier of user-defined named connections was added. See "Stable, readable .env connection keys with named connection overrides" above.

**Trigger**: Exported source files contained plaintext passwords in linked table connection strings, pass-through query definitions, and `db-connection.json`. When committed to version control, credentials were exposed to anyone with repository access (GitHub issue #670, #476).

**Options explored**:
- **Hash the full connection string as the .env key**: User's initial proposal. Brute-forceable — an attacker with the hash and knowledge of the server/driver could try password combinations to reproduce the hash. Rejected.
- **Hash with salt**: Adds security but makes keys non-deterministic across machines — different developers would generate different keys for the same connection, breaking shared source files. Rejected.
- **Hash only non-sensitive parts (DRIVER, SERVER, DATABASE, DSN)**: Deterministic across machines (same connection = same key regardless of credentials). Immune to brute-force since the hashed components are already visible in source files. Keys remain stable when passwords change. Chosen.
- **Descriptive prefix for keys** (`sql_myserver_mydb` vs `conn_a3f72b1`): Considered human-readable prefixes derived from connection components. Compact hash is more uniform, avoids special character issues, and the auto-generated comment above each entry provides the human context. Chose `conn_` prefix with 7-char SHA-256 hash.

**Decision**: Connection strings with credentials are replaced by `env:conn_<hash>` references in exported source files. The full connection string is stored in `{ExportFolder}/.env`, which is excluded from version control. Key design choices:

- **Three-mode option** (`UseEnvForConnections`): `Auto` (default, only when UID/PWD detected), `Always` (all connection strings), `Never` (disabled). Enum uses `uec` prefix per project convention.
  > **⚠ Partially superseded** (2026-06-19): `Never` no longer keeps *complete* strings in source. As of export format 5.1.0, raw passwords are stripped from source files in every mode (including `Never`); credentials live only in `.env`. See "Never write raw passwords to source files (any mode)" above.
- **Gated behind `EFV_5_0_0`**: No new export format version needed since v5 hasn't shipped.
- **Scope**: Linked tables (JSON), pass-through queries (.qdef via `clsSourceParser`), `db-connection.json`. Forms/reports deferred — investigation showed they don't directly embed connection strings.
- **Auto-population**: First export auto-creates `.env` with header comments explaining multi-developer workflow, and adds a descriptive comment above each entry (`# tblCustomers (linked table)`).
- **No auto-pruning**: `.env` is user-managed. Unused `conn_*` entry names are logged to the log file (not console) during full export so users can clean up manually.
- **Import resilience**: Missing `.env` keys log a warning; Access prompts for credentials at runtime.
- **Multi-line dbMemo handling**: Pass-through query connection strings can span multiple continuation lines in SaveAsText format. `clsSourceParser.SubstituteEnvConnect` collects all quoted fragments before substitution.
- **Cached .env reader**: Module-level `clsDotEnv` instance in `modConnect.bas` avoids re-reading the file for every table/query during a single operation.

**What this rules out**: Connection strings in source files are no longer guaranteed to be complete when `UseEnvForConnections` is not `Never`. Build/import workflows require a `.env` file with correct credentials. The `.env` file format follows standard `KEY=VALUE` conventions compatible with Docker, Node.js, and other ecosystems. If forms/reports are later found to embed connection strings directly (not via linked tables or named queries), `clsSourceParser` would need additional patterns. The `conn_` key prefix is reserved — `.env` entries with other prefixes (e.g., from external schema databases) are unaffected.

**Relevant files**:
- `Version Control.accda.src/modules/API/modAPI.bas` — `eUseEnvConnections` enum
- `Version Control.accda.src/modules/Infrastructure/clsOptions.cls` — `UseEnvForConnections` property, `GetUseEnvConnectionsName`, category hash classification
- `Version Control.accda.src/modules/Utility/modConnect.bas` — `GetConnectionEnvKey`, `ShouldUseEnvForConnection`, `SaveConnectionToEnv`, `ResolveEnvConnection`, `IsEnvReference`, `ResolveEnvReferencesInText`, `LogUnusedEnvEntries`, `CheckGitignoreForEnv`, `ClearEnvCache`, `GetEnvFilePath`, cached `clsDotEnv`
- `Version Control.accda.src/modules/Components/clsDbTableDef.cls` — env substitution on export, resolution on import
- `Version Control.accda.src/modules/Core/clsSourceParser.cls` — `SubstituteEnvConnect`, multi-line dbMemo handling
- `Version Control.accda.src/modules/Core/modLoadSaveText.bas` — `acQuery` case resolving env refs before `LoadFromText`
- `Version Control.accda.src/modules/Components/clsDbConnection.cls` — env refs in `GetSource`/`IDbComponent_Import`
- `Version Control.accda.src/modules/Core/modExport.bas` — `LogUnusedEnvEntries`, `CheckGitignoreForEnv`, `ClearEnvCache` calls
- `Version Control.accda.src/modules/Core/modBuild.bas` — `ClearEnvCache` calls
- `Version Control.accda.src/forms/frmVCSOptions.cls` — combo box population for `cboUseEnvForConnections`

---

## 2026-03-13 — @Folder annotation caching: Static per-instance vs modObjects-level cache

**Trigger**: After implementing `@Folder` annotation support (EFV 5.0.0), export logs from `C:\Repos\db-sec` showed "Clear Orphaned Files" consistently at 5-6 seconds on fast saves, even with zero modified objects. Root cause: `GetFolderAnnotation` reads the entire VBE code module via `cmpItem.CodeModule.Lines(1, 999999)` on every call, and `SourceFile` (which calls `GetFolderAnnotation`) was accessed multiple times per object per export — up to ~1,558 VBE COM calls for db-sec's 779 VBA-backed objects.

**Options explored**:

- **Approach A — modObjects-level Dictionary cache**: Add a `FolderAnnotations As Dictionary` to `udtObjects` in `modObjects.bas`, keyed by VBE component name. Provides cross-instance caching within a session. Initially planned, but analysis showed minimal benefit: Phase 1 (`GetAllFromDB`) has all unique keys (zero cache hits); Phase 2 (`ClearOrphanedSourceFiles`) is eliminated by the `varKey` fix; Phase 3 (export loop) reuses the same class instances (handled by instance-level caching). `ReleaseObjects` clears the cache between operations, preventing cross-operation persistence. Adds UDT member, accessor function, and cleanup code for ~12-90ms savings. Rejected.
- **Approach B — `Static` in `SourceFile` + `varKey` fix + Perf instrumentation**: Three small, self-contained changes: (1) `Static strCached` in each component's `SourceFile` Property Get caches the path for the lifetime of the instance; (2) `ClearOrphanedSourceFiles` uses `varKey` (the dictionary key, already the SourceFile path) instead of re-accessing `cItem.SourceFile`; (3) `Perf.OperationStart/End` around the VBE COM read in `GetFolderAnnotation` for measurement. Chosen.
- **Approach C — Batch-read all @Folder annotations in one pass**: Pre-scan all VBE components at the start of export, building a complete annotation map. Most efficient for VBE reads, but requires a new infrastructure function, changes the call pattern, and is premature without Perf data showing the ~779 reads are actually a bottleneck. Deferred pending Perf data.

**Decision**: Applied Approach B. The `Static` in `SourceFile` prevents repeated `GetFolderAnnotation` calls on the same instance (Export alone accesses `SourceFile` 4-6 times per object). The `varKey` fix eliminates ~779 redundant calls in `ClearOrphanedSourceFiles`. The Perf instrumentation will show the actual cost of the remaining ~779 VBE reads in `GetAllFromDB`, informing whether Approach A or C is worth revisiting.

**What this rules out**: A modObjects-level cache is not needed for the current workflow because dual-populate (`4f7f9c8`) shares class instances across export phases. If a future change introduces code paths that create separate instances for the same VBE component (breaking the shared-instance assumption), revisit the modObjects cache. If Perf data shows the ~779 VBE reads in `GetAllFromDB` are a significant bottleneck (>3 seconds), consider batch-reading annotations (Approach C).

**Relevant files**:
- `Version Control.accda.src/modules/Core/modVbeUtility.bas` — Perf instrumentation in `GetFolderAnnotation`
- `Version Control.accda.src/modules/Core/modOrphaned.bas` — `varKey` fix in `ClearOrphanedSourceFiles`
- `Version Control.accda.src/modules/Components/clsDbForm.cls` — `Static` cache in `SourceFile`
- `Version Control.accda.src/modules/Components/clsDbReport.cls` — `Static` cache in `SourceFile`
- `Version Control.accda.src/modules/Components/clsDbModule.cls` — `Static` cache in `SourceFile`
- `Version Control.accda.src/modules/Components/clsDbVbeForm.cls` — `Static` cache in `SourceFile`

---

## 2026-03-12 — Single-loop dual-populate for component cache slots

**Trigger**: During fast-save export, each `IDbComponent` class's `GetAllFromDB` was called twice per category: first with `blnModifiedOnly=True` (scan for changes), then with `blnModifiedOnly=False` (orphan detection via `ClearOrphanedSourceFiles`). Each call independently iterated the full Access collection and instantiated new `clsDb*` objects. Performance logs from `C:\Repos\db-sec` (~412 forms, ~3694 queries, ~392 tables) showed "Clear Orphaned Files" consistently taking 5.2-6.0 seconds — pure waste from re-enumerating objects already visited during the scan phase. Combined with "Scan DB Objects" (6.2-28.3s), these two passes consumed 34-54% of total fast-save runtime.

**Options explored**:

- **Approach A — Single-loop dual-populate**: When `GetAllFromDB(True)` iterates the collection, always populate `m_Items(False)` (all items) alongside `m_Items(True)` (modified items). The subsequent `GetAllFromDB(False)` call from `ClearOrphanedSourceFiles` hits the warm cache. A `blnNeedAll` flag prevents resetting `m_Items(False)` if it was already populated. Chosen.
- **Approach B — Lazy IsModified flag on instances**: Replace two-slot cache with a single dictionary of all items; cache `IsModified` results per instance and filter on demand. Conceptually clean, but filtering creates a new dictionary each time unless cached — reintroducing two-slot complexity. More invasive with no benefit over Approach A. Rejected.
- **Approach C — Lightweight orphan detection (no full instantiation)**: `ClearOrphanedSourceFiles` only needs base names, not full component instances. A new interface method could return just names. Initially dismissed as over-engineered, but db-sec logs proved orphan detection IS a bottleneck (5-6s consistently). However, Approach A eliminates the cost entirely without requiring interface changes, making Approach C unnecessary. Rejected.

**Decision**: Applied the single-loop dual-populate pattern to all 29 component classes implementing `IDbComponent`. Three implementation variants based on how each class determines modification:

1. **Per-item IsModified** (20 classes including all ADP classes): Single loop always adds to `m_Items(False)`, conditionally calls `IsModified` and adds to `m_Items(True)` only when `blnModifiedOnly=True`. Replaces `blnAdd` flag with `blnNeedAll` flag.
2. **Class-level IsModified** (7 classes: `clsDbConnection`, `clsDbDocument`, `clsDbNavPaneGroup`, `clsDbHiddenAttribute`, `clsDbProjProperty`, `clsDbVbeReference`): Uses `blnNeedAll` + `blnAddModified = IDbComponent_IsModified`. Iterates when either flag is set; adds to each slot based on its flag.
3. **Per-item with custom comparison** (2 classes: `clsDbProperty` with saved-vs-current dictionary comparison, `clsDbSharedImage` with duplicate detection against `m_Items(False)`): Retains specific filtering logic within the `blnModifiedOnly` branch.

Single-object classes (`clsDbProject`, `clsDbVbeProject`) also received the transform for consistency.

**What this rules out**: The `blnAdd` pattern (`blnAdd = True; If blnModifiedOnly Then blnAdd = ...; If blnAdd Then m_Items(blnModifiedOnly).Add ...`) is retired across all component classes. Future component classes should use the `blnNeedAll` single-loop pattern. The two-slot `m_Items(True To False)` declaration is unchanged — both slots still exist, but they are now populated in one pass instead of two. If a future calling pattern needs `GetAllFromDB(False)` first and then `GetAllFromDB(True)`, the `blnNeedAll` guard handles it correctly (iterates to build `m_Items(True)` from the existing objects without re-adding to `m_Items(False)`).

**Relevant files**:

- `Version Control.accda.src/modules/Components/clsDbForm.cls` — canonical example of per-item pattern
- `Version Control.accda.src/modules/Components/clsDbDocument.cls` — canonical example of class-level pattern
- `Version Control.accda.src/modules/Components/clsDbProperty.cls` — custom comparison pattern
- `Version Control.accda.src/modules/Components/clsDbSharedImage.cls` — duplicate detection pattern
- 25 additional component classes in `Components/` and `Components/ADP/` — same mechanical transform

---

## 2026-03-12 — SharedDb: shared CurrentDb reference across component classes

**Trigger**: Export of `sec.accdb` (~6,870 objects, ~567 with descriptions) took ~47s on fast save. Benchmarking revealed the bottleneck was **cold DAO property value reads** in `clsDbDocument.GetDictionary`: iterating Container/Document objects and reading `Description` values took ~18s due to physical disk I/O in the JET engine loading scattered property-value pages. Multiple component classes each called `Set dbs = CurrentDb` independently, and each new `CurrentDb` reference starts with a cold JET page cache (per-reference caching). This meant duplicate cold I/O penalties when multiple components accessed the same data.

**Options explored**:

- **MSysObjects SQL lookup**: Query the system table for descriptions instead of iterating DAO. Found only 16/567 descriptions — queries are stored under the "Tables" DAO container, not a "Queries" container. Even with correct mapping, this was not faster than DAO iteration for value reads.
- **Dictionary creation optimization**: Hypothesized that creating `Scripting.Dictionary` objects was expensive. Benchmarked at 0.008s for 1,200 dictionaries — negligible. Rejected.
- **Content hash via clsConcat**: Build a canonical string and hash it instead of building dictionaries. Fast for warm reads (0.33s) but doesn't avoid the cold I/O.
- **Shared CurrentDb reference (SharedDb)**: Cache a single `CurrentDb` reference in `modObjects` (lazy singleton pattern like FSO, Options, etc.). All component classes reuse the same reference, so the JET page cache stays warm after the first component pays the cold I/O cost. Chosen.
- **Separate warm-up pass (WarmDAOCache)**: Iterate all documents pre-scan to warm the cache, tracked as "Loading DB Objects". Implemented and then **reverted** — it added ~9s overhead by iterating all ~6,870 documents in a separate pass before the scan iterated them again. Total time increased from ~47s to ~63-71s.
- **Cold-start category annotation**: Tried annotating whichever category triggered the first SharedDb creation with a `*` footnote. The annotation landed on "DB Properties" (0.09s) because `clsDbProperty` runs before `clsDbDocument` in the scan order — but the actual cold I/O is paid later in "Doc Properties" (~18s). The annotation concept was correct but the trigger point was wrong. Removed the annotation call from `SharedDb()`; the `AddCategoryNote` mechanism remains available.

**Decision**: Added `SharedDb()` accessor to `modObjects.bas` following the existing singleton pattern (FSO, Options, VCSIndex). Replaced `Set dbs = CurrentDb` with `Set dbs = SharedDb` across 10 component classes. The key JET caching insights from 7 rounds of in-database benchmarks:

- **Per-reference caching**: Each `CurrentDb` call starts with a cold cache; references don't share warm state
- **Page-level caching**: Warming one property (Description) warms ALL properties on those documents (Owner reads: 0.051s for 4,942 docs after warming Description)
- **Cache pressure**: Aggressive full-property iteration causes exponential slowdown (500 docs: 0.07s → 2,000 docs: 261s) due to JET buffer pool saturation
- **LRU eviction**: Previously cached pages persist even after heavy I/O — targeted warm-up is safe

The separate `WarmDAOCache` warm-up pass was reverted because the first component to iterate (Doc Properties) naturally warms the cache for all subsequent components on the same `SharedDb` reference. **The real optimization opportunity discovered during this work**: commenting out Doc Properties entirely reduced export from ~47s to ~27s. This suggests the next step is making the Doc Properties scan conditional (skip when no objects are modified), not trying to make the cold I/O faster.

**What this rules out**: Components should use `SharedDb` instead of `CurrentDb` for DAO operations during export/scan. Do NOT add a separate warm-up pass — it's counterproductive. Do NOT try to annotate the cold-start category via `SharedDb()` creation — the reference creation and the cold I/O are separate events. The actual performance win for large databases will come from skipping the Doc Properties full scan when no objects have changed (future work).

**Relevant files**:

- `Version Control.accda.src/modules/Infrastructure/modObjects.bas` — `SharedDb()`, `Dbs` in `udtObjects`, cleared in `ReleaseObjects`
- `Version Control.accda.src/modules/Components/clsDbDocument.cls` — 5x `CurrentDb` → `SharedDb`
- `Version Control.accda.src/modules/Components/clsDbHiddenAttribute.cls` — 4x `CurrentDb` → `SharedDb`
- `Version Control.accda.src/modules/Components/clsDbProperty.cls` — 4x `CurrentDb` → `SharedDb`
- `Version Control.accda.src/modules/Components/clsDbTableDef.cls` — 6x `CurrentDb` → `SharedDb`
- `Version Control.accda.src/modules/Components/clsDbQuery.cls` — 3x `CurrentDb` → `SharedDb`
- `Version Control.accda.src/modules/Components/clsDbRelation.cls` — 3x `CurrentDb` → `SharedDb`
- `Version Control.accda.src/modules/Components/clsDbNavPaneGroup.cls` — 3x `CurrentDb` → `SharedDb`
- `Version Control.accda.src/modules/Components/clsDbImexSpec.cls` — 5x `CurrentDb` → `SharedDb`
- `Version Control.accda.src/modules/Components/clsDbTableData.cls` — 4x `CurrentDb` → `SharedDb`
- `Version Control.accda.src/modules/Components/clsDbTableDataMacro.cls` — 1x `CurrentDb` → `SharedDb`
- `Version Control.accda.src/modules/Core/modExport.bas` — WarmDAOCache added then removed

---

## 2026-03-12 — Generic category footnotes and TOTAL RUNTIME on clsPerformance

**Trigger**: During the SharedDb investigation, we wanted to annotate specific categories in the performance report with explanatory footnotes (e.g., marking which category paid the cold I/O cost). This required a mechanism on `clsPerformance` that was domain-agnostic, since the performance class is used for generic timing beyond just import/export.

**Options explored**:

- **Domain-specific property (ColdStartCategory)**: A single string property on `clsPerformance`. Simple but bakes import/export knowledge into a generic class. Rejected.
- **Generic CategoryNotes dictionary**: A single dictionary keyed by category name with note text as value. Supports one note per category. Considered but less flexible.
- **Two-dictionary footnote system with mark characters**: `FootnoteMarks` (mark → description) and `CategoryFootnotes` (category → accumulated marks string). Supports multiple distinct footnotes on the same category (e.g., `"*†"`), and different categories can share the same mark. Default mark is `"*"`. Chosen.

**Decision**: Added `AddCategoryNote(strCategory, strNote, Optional strMark = "*")` to `clsPerformance`. The method silently exits if `strCategory` is empty or perf is disabled. `GetReports` appends marks to category names in the table and renders footnote descriptions after the TOTALS row. Both dictionaries are cleared in `Reset()`. Also added a `TOTAL RUNTIME` line to the operations table footer, showing `this.Overall.Total` — makes it easy to see how operations add up to wall-clock time without referencing the "Done" line at the top of the log.

**What this rules out**: The footnote mechanism is fully generic — callers provide the mark character and description. There is no automatic detection built into `clsPerformance`; callers must explicitly call `AddCategoryNote`. Currently no callers use it (the `SharedDb` annotation was removed after proving the trigger point was wrong), but the mechanism is ready for future use.

**Relevant files**:

- `Version Control.accda.src/modules/Infrastructure/clsPerformance.cls` — `AddCategoryNote`, `FootnoteMarks`, `CategoryFootnotes` in `udtPerformance`, `GetReports` rendering, `TOTAL RUNTIME` line

---

## 2026-03-11 — Skip unavailable back-ends during export

**Trigger**: When exporting a database with many linked tables pointing to the same unavailable back-end (file missing, server down), the export tried and failed on every linked table individually. Each failure hit `TableExists()` → `tdf.Fields.Count`, which errors or times out, and logged a separate error per table. For ODBC connections, each failure could incur a full network timeout, multiplied by the number of linked tables.

**Options explored**:

- **Filter unavailable tables in `GetAllFromDB`**: Skip linked tables with unavailable back-ends during the scan phase so they never enter the export list. Would prevent the table from appearing in counts and progress, and would mix back-end availability concerns into the component discovery layer. Rejected as wrong abstraction level.
- **Pre-test all connection types proactively**: Extend `CacheBackEndConnections` to also test ODBC connections upfront. Would provide uniform proactive detection but risks triggering ODBC login prompts or long timeouts during the pre-scan for servers that the user hasn't configured for unattended access. Rejected for ODBC; kept for Access (already tested).
- **Proactive detection for Access + reactive detection with connection test for ODBC**: For Access back-ends, `CacheBackEndConnections` already opens each unique back-end file — just record failures instead of silently skipping them. For ODBC, on first `TableExists` failure, run a lightweight server-level connection test (`SELECT 1` via temp QueryDef) to distinguish "server down" from "single table missing." If the server is unreachable, mark the back-end as unavailable and skip remaining tables. If it responds, treat as a single-table error. Chosen.

**Decision**: Added `m_dUnavailableBackEnds` dictionary to `modConnect.bas`, keyed by normalized back-end identifier. Modified `CacheBackEndConnections` to record failed `DBEngine.OpenDatabase` attempts (with per-back-end table counts) and log a single `eelWarning` per unavailable Access back-end. Added four new functions: `IsBackEndUnavailable` (dictionary lookup), `MarkBackEndUnavailable` (reactive recording + warning log), `TestBackEndConnection` (lightweight `SELECT 1` for ODBC; checks `m_dBackEndConnections` for Access), and `GetBackEndKey` (normalizes connection strings to back-end identifiers — file path for Access, DSN or DRIVER+SERVER+DATABASE for ODBC). Modified `clsDbTableDef.Export` and `clsDbTableData.Export` to check `IsBackEndUnavailable` before `TableExists`, and to call `TestBackEndConnection` on failure to distinguish server-down from table-missing.

The back-end key normalization uses `UCase$` for case-insensitive matching. Access keys are file paths. ODBC keys use `ODBC:DSN=<name>` for DSN-based connections or `ODBC:<driver>;<server>;<database>` for DSN-less. `CloseBackEndConnections` clears both the connection cache and the unavailable dictionary.

**What this rules out**: The unavailable back-end tracking is session-scoped (cleared in `CloseBackEndConnections`). It does not persist across operations. ODBC detection is reactive — the first linked table to an unavailable ODBC server will still incur one timeout before the back-end is marked. Proactive ODBC testing could be reconsidered if users report that single-timeout cost is still too high, but it would need to handle credential prompts. `clsDbTableDataMacro` is not modified because its `GetAllFromDB` already filters out linked tables (`If Len(tdf.Connect) = 0`).

**Relevant files**:

- `Version Control.accda.src/modules/Utility/modConnect.bas` — `m_dUnavailableBackEnds`, `IsBackEndUnavailable`, `MarkBackEndUnavailable`, `TestBackEndConnection`, `GetBackEndKey`, `GetConnectPart`, modified `CacheBackEndConnections` and `CloseBackEndConnections`
- `Version Control.accda.src/modules/Components/clsDbTableDef.cls` — `IDbComponent_Export` modified with back-end availability check and reactive ODBC detection
- `Version Control.accda.src/modules/Components/clsDbTableData.cls` — `IDbComponent_Export` modified with same pattern

---

## 2026-03-11 — Persistent back-end database connection caching during export

> **⚠ Partially superseded** (2026-03-11): The claim "Inaccessible back-ends are silently skipped" is no longer true. `CacheBackEndConnections` now records unavailable back-ends in `m_dUnavailableBackEnds` and logs a warning per back-end with the count of affected tables. See "Skip unavailable back-ends during export" above.

**Trigger**: When exporting a database with linked tables pointing to Access back-end files (.accdb/.mdb), the Jet/ACE engine repeatedly opens and closes connections to the same back-end databases. Each access to a linked `TableDef`'s properties (`.Connect`, `.Fields`, `.Indexes`, `.SourceTableName`) or data (`OpenRecordset`, `ExportXML`) can trigger a separate connection cycle. With N linked tables pointing to the same back-end, this produces dozens of redundant open/close operations — especially costly when back-ends are on network shares.

**Options explored**:

- **Cache `TableDef` metadata in memory**: Instead of repeatedly accessing `tdf.Connect`, `tdf.Fields`, `tdf.Indexes`, cache these values in a dictionary on first access. Would reduce property-level overhead but wouldn't help with `OpenRecordset`/`ExportXML` operations, which are the heaviest. Rejected as partial solution.
- **Batch export operations by back-end database**: Group all linked tables by their back-end and process them together to maximize connection reuse within each batch. Would require significant restructuring of the export loop architecture. Rejected as too invasive.
- **Hold persistent `DAO.Database` references to back-end files**: Open each unique back-end database in shared read-only mode at the start of an operation, keeping the Jet/ACE internal connection pool warm. The engine reuses pooled connections for subsequent linked table operations. Mirrors the existing ODBC `CacheConnection` pattern in `modConnect.bas`. Chosen.

**Decision**: Added `CacheBackEndConnections()` and `CloseBackEndConnections()` to `modConnect.bas`, following the same architectural pattern as the existing ODBC `CacheConnection`/`CloseCachedConnections`. A module-level `m_dBackEndConnections` dictionary holds open `DAO.Database` references keyed by full file path. `CacheBackEndConnections` scans `CurrentDb.TableDefs` for links starting with `;DATABASE=`, extracts unique back-end paths, and opens each via `DBEngine.OpenDatabase(path, False, True)` (shared, read-only). Inaccessible back-ends are silently skipped. Performance timing is included via `Perf.OperationStart`/`OperationEnd`, and a log message reports how many connections were cached.

The cached read-only connection does not interfere with read-write operations on linked tables (e.g., `RunAfterExport` subs that write data) because linked table operations go through `CurrentDb`'s own connection path, which is independent.

Integration points: `CacheBackEndConnections` is called early in `ExportSource` (after `CloseDatabaseObjects`), `ExportSingleObject`, and `ExportMultipleObjects`. `CloseBackEndConnections` is called in the `CleanUp` section of all three export functions and in `modBuild.Build` (both startup and cleanup, alongside existing `CloseCachedConnections`).

**What this rules out**: This optimization targets only Access back-end links (`;DATABASE=` connection strings). ODBC links are already handled by the existing `CacheConnection` system. Excel, text file, and SharePoint links use different connection mechanisms and are not addressed. If back-end databases are moved or renamed during an operation, the cached connections become stale — but this is an unlikely scenario during export. The read-only open mode prevents write-locking conflicts but means the cache cannot be used to write to back-end tables (nor is it intended to).

**Relevant files**:

- `Version Control.accda.src/modules/Utility/modConnect.bas` — `CacheBackEndConnections()`, `CloseBackEndConnections()`, `m_dBackEndConnections`
- `Version Control.accda.src/modules/Core/modExport.bas` — cache/close calls in `ExportSource`, `ExportSingleObject`, `ExportMultipleObjects`
- `Version Control.accda.src/modules/Core/modBuild.bas` — close calls in startup and cleanup

---

## 2026-03-11 — Version-gate file extensions and @Folder paths for export format downgrade

**Trigger**: Switching `ExportFormatVersion` from 5.0.0 back to 4.1.2 left files in `@Folder` subfolders and with descriptive extensions (`.form`, `.report`, `.qdef`, `.macro`) instead of reverting to the original flat layout with `.bas` extensions. The `@Folder` subfolder path was already gated behind `EFV_5_0_0` in the `SourceFile` property (the Else branch omitted `GetFolderAnnotation`), but the file extension was always the new one in both branches. No reverse migration existed — only forward migration via `MigrateFileExtensions`. This caused all subfolder items to appear as orphaned files during export.

**Options explored**:

- **Gate only SourceFile extensions**: Would fix the export path but leave old files in subfolders with new extensions on disk, since orphan detection uses `FileExtensions` to decide which extensions to scan. Insufficient alone.
- **Gate SourceFile + FileExtensions + add reverse migration**: Ensures `SourceFile` returns `.bas` for format < 5.0.0, orphan detection scans for the correct extensions, and existing 5.0.0 files are actively moved/renamed back during export. Chosen.
- **Delete subfolder files instead of moving them**: Simpler but loses the user's source files, requiring a full re-export. Rejected.

**Decision**: Six files changed across four concerns:

1. **Extension gating in `SourceFile`**: The Else branch (format < 5.0.0) in `clsDbForm`, `clsDbReport`, `clsDbQuery`, `clsDbMacro` now uses `.bas` instead of the descriptive extension. Queries and macros gained a new `EFV_5_0_0` gate (they previously had no version gate at all).

2. **Extension gating in `FileExtensions`**: The primary extension returned by `FileExtensions` is version-gated in the same four classes (`"bas"` when < 5.0.0, descriptive extension when >= 5.0.0). This ensures orphan detection scans for the correct file types.

3. **Reverse migration** (`RevertFileExtensions` + `FlattenSubfolders` in `modSourceUpgrade.bas`): Counterpart to `MigrateFileExtensions`. `FlattenSubfolders` recursively moves all files from subfolders to the base folder for each @Folder-capable type (forms, reports, modules, VBE forms), then removes empty directories. `RevertFileExtensions` then renames `.form`/`.report`/`.qdef`/`.macro` back to `.bas` using the existing `RenameFilesInFolder` helper, and calls `VCSIndex.MigrateIndexExtension` (already bidirectional) to update index keys. Called from `modExport.ExportSource` when format < 5.0.0.

4. **Build backward compatibility**: `GetFileList` in all four classes now searches for both old (`.bas`) and new (`.form`/`.report`/`.qdef`/`.macro`) extensions using `MergeDictionary`, so builds work regardless of which format was used to export the source files.

**What this rules out**: Export format downgrade is now a supported operation — switching between 4.1.2 and 5.0.0 actively migrates files in both directions. The reverse migration runs on every export when format < 5.0.0 (same pattern as forward migration), but `RenameFilesInFolder` and `FlattenSubfolders` are no-ops when there's nothing to move. Future format versions that add new file organization features must also implement the reverse path. `GetFileList` searching for both extensions means the `forms/`, `reports/`, `queries/`, and `macros/` folders should not contain `.bas` files from other sources (e.g., stray VBA modules) — but this was already implicitly true since these folders are component-type-specific.

**Relevant files**:

- `Version Control.accda.src/modules/Components/clsDbForm.cls` — `SourceFile`, `FileExtensions`, `GetFileList`
- `Version Control.accda.src/modules/Components/clsDbReport.cls` — `SourceFile`, `FileExtensions`, `GetFileList`
- `Version Control.accda.src/modules/Components/clsDbQuery.cls` — `SourceFile`, `FileExtensions`, `GetFileList`
- `Version Control.accda.src/modules/Components/clsDbMacro.cls` — `SourceFile`, `FileExtensions`, `GetFileList`
- `Version Control.accda.src/modules/Core/modSourceUpgrade.bas` — `RevertFileExtensions`, `FlattenSubfolders`
- `Version Control.accda.src/modules/Core/modExport.bas` — conditional migration call

---

## 2026-03-10 — Organize 98 VBA source files into 10 architectural folders

**Trigger**: After the v5 module splits (Phases 1-6), the project has 98 modules and classes in a flat `modules/` directory. Finding related code requires prior knowledge or full-text search. With `@Folder` annotations now supported, the files can be organized into subfolders that reflect the architectural layers established during the reorganization.

**Options explored**:

- **Deep hierarchy (3+ levels)**: E.g., `Core/Export/`, `Core/Build/`, `Infrastructure/Logging/`, `Utility/FileIO/`, `Utility/String/`. More granular but adds folder overhead without improving discoverability for a project this size. Rejected.
- **Flat 10-folder structure**: One level of folders mapping to architectural roles: API, Components (with ADP and Schema sub-folders), Core, Infrastructure, Integration, Install, Utility, Lib, Tests. Balances organization with simplicity. Chosen.

**Decision**: Add `'@Folder("FolderName")` annotations to all 99 files (98 new + 1 existing `modUnitTesting`). Annotations are placed immediately after the `Option` statements (Option C). The 10 folders are:

- **API** (3): Public entry points — `modAPI`, `modAddInMenu`, `clsVersionControl`
- **Components** (25): `IDbComponent` interface and all standard implementations (`clsDbForm`, `clsDbQuery`, etc.)
- **Components.ADP** (5): ADP-specific components — `clsAdpFunction` through `clsAdpTrigger`
- **Components.Schema** (3): External database schema exporters — `IDbSchema`, `clsSchemaMsSql`, `clsSchemaMySql`
- **Core** (18): Export/build/merge orchestration and supporting logic — `modExport`, `modBuild`, `modContainers`, `clsSourceParser`, `clsPrinterSettings`, etc.
- **Infrastructure** (13): Global state, singletons, cross-cutting concerns — `modObjects`, `modErrorHandling`, `modConstants`, `clsOptions`, `clsVCSIndex`, `clsLog`, etc.
- **Integration** (4): External system interfaces — `clsGitIntegration`, `clsWorker`, `clsMCP`, `modExportOnSaveHook`
- **Install** (5): Add-in installation/deployment — `modInstall`, `modResource`, `modCOMAddIn`, `modRepair`, `modRibbonStrings`
- **Utility** (20): General-purpose helpers with no VCS-specific logic — `modFileAccess`, `modDatabase`, `modEncoding`, `modHash`, `modFunctions`, `clsConcat`, etc.
- **Lib** (2): Third-party code — `modJsonConverter`, `modUtcConverter`
- **Tests** (1): Already annotated — `modUnitTesting`

**Design rationale**: Components get their own tree (33 files total) because they are the largest and most uniform group. Core vs Infrastructure distinguishes "what the add-in does" from "how global state is managed." Utility stays flat at 20 files because these are leaf-level functions with no internal dependencies. Lib isolates third-party code so agents and developers know not to modify it. Integration groups external system interfaces (Git, MCP, export hooks) that depend on systems outside the VBA project.

**What this rules out**: The folder structure is enforced only via `@Folder` annotations and the `GetFolderAnnotation` parser — there is no build-time validation that a file's folder matches its actual dependencies. Moving a file to a different folder requires only changing its annotation and re-exporting.

**Relevant files**: All 99 `.bas` and `.cls` files in `Version Control.accda.src/modules/` were modified to add `'@Folder(...)` annotations.

---

## 2026-03-10 — @Folder annotation support for subfolder organization of exported source files

**Trigger**: With 30+ modules and classes in a project, the flat export structure (all modules in `modules/`, all forms in `forms/`) becomes hard to navigate. Rubberduck VBA already defines a `'@Folder("...")` annotation convention for logically grouping VBA components. Implementing this in the VCS add-in lets developers and AI agents organize source files into meaningful subfolders (e.g., `modules/Core/Utility/`, `forms/UI/`) while keeping each component type under its own root folder.

**Options explored**:

- **Combine component types into shared folders**: A single `src/Core/` folder could hold modules, classes, and forms together. Rejected — the existing architecture relies on component-type folders (`modules/`, `forms/`, `reports/`) for `BaseFolder`, `GetFileList`, file pattern matching, and orphan detection. Mixing types would require rewriting the entire component discovery system and break the `IDbComponent` contract.
- **Custom annotation format**: Invent a new syntax like `'!Folder:Core.Utility`. Rejected — Rubberduck's `'@Folder("...")` is already widely used by VBA developers, and compatibility means users don't need to learn a new convention or maintain two sets of annotations.
- **Line-by-line scan with 30-line limit**: Iterate `CodeModule.Lines(n, 1)` for the first 30 lines. Worked but made up to 30 COM calls per module and imposed an arbitrary cutoff. Rejected in favor of `InStr`.
- **Rubberduck-compatible `@Folder` with `InStr`-based search**: Read the full code module in a single `CodeModule.Lines(1, n)` call, prepend `vbCrLf`, and use `InStr` to find `vbCrLf & "'@FOLDER("`. No line-position limit, single COM call, and annotations must be on a comment line. Chosen.

**Decision**: Subfolder export is gated behind `Options.ExportFormatVersion >= EFV_5_0_0` (unreleased). Import always recurses into subfolders regardless of format version, ensuring backwards compatibility. Key design choices:

- **Annotation parser**: `GetFolderAnnotation()` in `modVbeUtility.bas` reads the entire code module in one COM call, prepends `vbCrLf` so line-1 annotations match, and searches for `vbCrLf & "'@FOLDER("` via `InStr`. Annotations must be on a comment line (preceded by `'`). Users can disable an annotation by removing the leading single quote. A second `InStr` past the first match detects duplicates.
- **Multiple annotations**: First `@Folder` annotation wins; duplicates log a warning via `Log.Add` with `ShowDebug` visibility.
- **Prefix parameter**: Forms use `"Form_"` prefix, reports use `"Report_"` prefix to match VBE component naming (e.g., `Form_frmMain`). Modules and VBE forms pass no prefix.
- **Index unaffected**: `clsVCSIndex` keys on `FSO.GetFileName()` (just the filename), so subfolder changes don't break index lookups.
- **Old file cleanup**: Each `Export` method deletes stale files at the base folder when `@Folder` moves them to a subfolder, handling annotation changes and format upgrades.
- **MoveSource**: All `MoveSource` implementations call `VerifyPath` on the destination, and `modExport.bas` passes `FSO.GetParentFolderName(cDbObject.SourceFile)` instead of `cDbObject.BaseFolder` to preserve subfolder structure during temp-file moves.
- **Orphan cleanup**: `modOrphaned.bas` recurses into subfolders and removes empty directories after cleanup.
- **File counting**: `GetQuickFileCount` in `modContainers.bas` counts files recursively for accurate progress bars.

**What this rules out**: Component types remain in separate root folders — `@Folder` only creates subfolders within each type's folder. The dot character in annotations is reserved as a path separator (consistent with Rubberduck). If Rubberduck changes its annotation syntax, this implementation would need updating. Annotations embedded in string literals or mid-line code will not match (the `vbCrLf & "'` prefix is required). There is no line-position limit for the annotation.

**Relevant files**:

- `Version Control.accda.src/modules/modVbeUtility.bas` — `GetFolderAnnotation()` parser
- `Version Control.accda.src/modules/modFileAccess.bas` — `GetFilePathsInFolderRecursive()`
- `Version Control.accda.src/modules/clsDbModule.cls` — SourceFile, GetFileList, Export, MoveSource updated
- `Version Control.accda.src/modules/clsDbForm.cls` — SourceFile, GetFileList, Export, MoveSource updated
- `Version Control.accda.src/modules/clsDbReport.cls` — SourceFile, GetFileList, Export, MoveSource updated
- `Version Control.accda.src/modules/clsDbVbeForm.cls` — SourceFile, GetFileList, Export, MoveSource updated
- `Version Control.accda.src/modules/modOrphaned.bas` — recursive `ScanFolderForOrphans`, empty folder cleanup
- `Version Control.accda.src/modules/modExport.bas` — subfolder-aware `MoveSource` destination paths
- `Version Control.accda.src/modules/modContainers.bas` — recursive `CountFilesRecursive` for `GetQuickFileCount`

---

## 2026-03-10 — Auto-batch split files when one source maps to multiple destinations

**Trigger**: `SplitFilesWithHistory` in `clsGitIntegration` uses `FSO.MoveFile` to rename each source file to its destination. When the same source file is listed multiple times (e.g., splitting `modVCSUtility.bas` into `modContainers.bas`, `modVbeUtility.bas`, and `modLoadSaveText.bas`), only the first move succeeds — subsequent entries are silently skipped because the source no longer exists. The `If FSO.FileExists(strOrig)` guard masks the failure.

**Options explored**:

- **Validate and block**: Detect duplicate source files during validation and show an error telling the user to manually split their list into batches. Simple, but pushes complexity onto the user.
- **Auto-batch with confirmation**: Automatically group entries into the minimum number of batches (one destination per source per batch) using round-robin distribution, then confirm the batch count with the user before executing. More complex, but transparent and user-friendly. Chosen.

**Decision**: Auto-batch in the form code (`frmVCSSplitFiles.cmdSplitFiles_Click`). A `Dictionary` counts occurrences of each source path; the max count determines batch count. Entries are distributed round-robin into `Collection` arrays by source. When batches > 1, a `MsgBox2` with OK/Cancel shows the batch count and number of additional commits. Each batch calls `SplitFilesWithHistory` independently. The single-batch path (no duplicate sources) remains unchanged. `SplitFilesWithHistory` itself is not modified — it already works correctly for one-destination-per-source batches.

**What this rules out**: The form no longer rejects duplicate source entries — it handles them. If `SplitFilesWithHistory` ever changes its branch naming (currently hardcoded `"split-files"`), the sequential batch execution would still work since each call deletes the temp branch before returning. If git operations fail mid-batch, only the completed batches are committed; partial recovery would require manual git intervention. Revisit if users report issues with large batch counts or if `SplitFilesWithHistory` gains its own multi-pass support.

**Relevant files**:

- `Version Control.accda.src/forms/frmVCSSplitFiles.cls` — batching logic added to `cmdSplitFiles_Click`
- `Wiki/Split-Files.md` — new "Splitting One File Into Multiple Files" section

---

## 2026-03-10 — Break modObjects/modErrorHandling circular dependency via ConfigureErrorHandling

**Trigger**: `modErrorHandling` called `Options.BreakOnError` and `OptionsLoaded` (from `modObjects`) to decide whether to break on errors. `modObjects.FSO` called `LogUnhandledErrors` and `CatchAny` (from `modErrorHandling`). This circular dependency meant: (1) FSO initialization could trigger Options loading through error handling, (2) error handling during Options loading could re-enter itself, requiring a fragile `blnInError` guard, and (3) an agent reading either module had to understand implicit initialization order.

**Options explored**:

- **Extract modErrorCore.bas**: Move core error functions to a leaf module. Partially breaks the cycle but splits a small, cohesive module for marginal gain. Rejected.
- **Callback pattern — cache BreakOnError locally**: `modErrorHandling` stores `blnBreakOnError` in its private UDT (defaults to False). `modObjects` pushes the value via `ConfigureErrorHandling` after options load. Eliminates `Options` and `OptionsLoaded` dependencies entirely. The remaining `Log.Error` coupling is documented as the single coupling point — it cannot be cleanly removed in VBA since there are no function pointers for object methods. Chosen.

**Decision**: Added `ConfigureErrorHandling(blnBreakOnError)` to `modErrorHandling`. Replaced `Options.BreakOnError` and `OptionsLoaded` references with `this.blnBreakOnError`. Added calls in `modObjects.Options` Property Get/Set to push the setting after options load. The `blnInError` re-entrancy guard is still present for `Log.Error` safety but the Options-triggered loop is fully eliminated.

**What this rules out**: `modErrorHandling` must not directly reference `Options` or `OptionsLoaded`. Any future BreakOnError changes at runtime must call `ConfigureErrorHandling` to take effect. The `Log.Error` coupling remains; removing it would require an event/callback mechanism that VBA does not natively support.

**Relevant files**:

- `Version Control.accda.src/modules/modErrorHandling.bas` — decoupled from Options
- `Version Control.accda.src/modules/modObjects.bas` — pushes BreakOnError after options load

---

## 2026-03-10 — Document IDbComponent contracts; change DbObject to Property Get/Set

**Trigger**: `DbObject` was declared as a public field on the `IDbComponent` interface, which is functionally equivalent to Property Get/Set in VBA but obscures the actual contract. The `IsModified` and `QuickCount` methods had no documented contract, making it unclear to new developers and agents which change-detection strategy each component uses or how `QuickCount` caching works.

**Options explored**:

- **Add Parent to IDbComponent interface**: Initially planned, but rejected after analysis. `Parent` is only useful from concrete-typed variables (e.g. `cForm.Parent.SourceFile`), which call the public property directly — the `IDbComponent_Parent` interface implementation would never be called since callers with an `IDbComponent`-typed variable already have the reference. Adding it would touch 29 classes for no practical benefit.
- **Change DbObject to Property Get/Set, document contracts**: Makes the interface declaration explicit and self-documenting. No implementing class changes needed since they already used property pairs. Chosen.

**Decision**: Changed `Public DbObject As Object` from a field declaration to explicit `Property Get`/`Property Set` on the interface. Added a contract documentation block to the `IDbComponent` header describing:

- **IsModified** strategies: date-only (9 classes), hash-only (17 classes), date+hash (2 classes: Form, Report), and special cases (TableData always True, SharedImage dual-hash).
- **QuickCount** caching semantics: approximate count cached via `Static` variable, suitable for progress bars only — not for exact tallies.

`Parent` was intentionally left off the interface — it remains as a `Public Property Get` on each concrete class where it serves its actual purpose.

**What this rules out**: `IsModified` implementations should follow one of the documented strategies and note any deviation. `Parent` will not be added to the interface unless a concrete use case through `IDbComponent`-typed variables emerges.

**Relevant files**:

- `Version Control.accda.src/modules/IDbComponent.cls` — interface updated

---

## 2026-03-10 — Fix naming inconsistencies; rename clsDevMode to clsPrinterSettings

**Trigger**: Four modules had stale `' Module :` header comments left over from earlier renames, creating confusion for both agents and developers scanning headers. Additionally, `clsDevMode` was named after the Windows API `DEVMODE` structure it wraps, but readers unfamiliar with the Win32 API assumed it meant "debug mode" or "developer mode." The class is actually a printer/page-layout settings parser.

**Options explored**:

- **Fix headers only, leave clsDevMode**: Fixes the copy-paste errors but leaves the most misleading name. Rejected — the v5 reorganization is the right time to rename.
- **Fix headers and rename clsDevMode to clsPrinterSettings**: Aligns the class name with its responsibility (parsing and applying printer settings for forms/reports). Internal variable names like `tDevMode` and `m_tDevMode` are kept because they directly reference the Windows `DEVMODE` structure and are appropriate at that level. Chosen.
- **Rename to clsPageLayout**: Considered but the class also handles printer name, paper bin, collation, and other non-layout settings. `clsPrinterSettings` is more accurate.

**Decision**: Fixed four header/filename mismatches (`modOrphaned` said `modVCSUtility`, `modFileWinAPI` said `modFileScan`, `modAddInMenu` said `modAddIn`, `modSqlFunctions` said `modAdpFunctions`). Renamed `clsDevMode.cls` to `clsPrinterSettings.cls` via `git mv` to preserve history, updated `Attribute VB_Name`, header comment, and all three callers (`clsVCSIndex`, `clsSourceParser`, `modLoadSaveText`).

**What this rules out**: The name `clsDevMode` is retired. Future printer/page-layout work goes in `clsPrinterSettings`. Header `' Module :` lines must always match `Attribute VB_Name`.

**Relevant files**:

- `Version Control.accda.src/modules/clsPrinterSettings.cls` — renamed from clsDevMode.cls
- `Version Control.accda.src/modules/clsVCSIndex.cls` — caller updated
- `Version Control.accda.src/modules/clsSourceParser.cls` — caller updated
- `Version Control.accda.src/modules/modLoadSaveText.bas` — caller updated
- `Version Control.accda.src/modules/modOrphaned.bas` — header fixed
- `Version Control.accda.src/modules/modFileWinAPI.bas` — header fixed
- `Version Control.accda.src/modules/modAddInMenu.bas` — header fixed
- `Version Control.accda.src/modules/modSqlFunctions.bas` — header fixed

---

## 2026-03-10 — Split modVCSUtility into modContainers, modVbeUtility, modLoadSaveText

**Trigger**: `modVCSUtility.bas` was a 1,527-line, 35-procedure catch-all module mixing component container registry, VBA editor operations, Access LoadFromText/SaveAsText wrappers, version helpers, schema filters, git file management, and command bar import. The name "modVCSUtility" gave no hint about which concern lived here.

**Options explored**:

- **Keep as one module**: Simple but the file mixed too many unrelated domains. A developer looking for "how does SaveAsText work?" had to wade through container setup and VBE compilation code. Rejected.
- **Split into two (containers vs everything else)**: Better but VBE operations and text I/O are distinct domains with different dependency profiles. Rejected as insufficient.
- **Split into four by responsibility**: Container registry (11 functions), VBE operations (7 functions), text I/O (4 functions), and remaining utility functions. Each module has a clear domain signaled by its name. Chosen.

**Decision**: Split into `modContainers.bas` (GetContainers, GetClassFromObject, GetComponentClass, ContainerHasObject, ContainerHasAnyObject, MergeIfChanged, GetQuickObjectCount, GetQuickFileCount, GetSourceModifiedDate, GetLastModifiedSourceFile, GetSourceFilesPropertyHash), `modVbeUtility.bas` (ExportCodeModule, OverlayCodeModule, RemoveNonBuiltInReferences, CompileAndSaveAllModules, PreloadVBE, GetAddInProject, LoadVCSAddIn), `modLoadSaveText.bas` (SaveComponentAsText, LoadComponentFromText, RequiresOverlay [Private], ReadSourceFile), and a slimmed `modVCSUtility.bas` (version helpers, path utilities, BuildJsonFile, CheckGitFiles, ShiftOpenDatabase, schema helpers, command bar import). `RequiresOverlay` was kept Private in `modLoadSaveText` with its only caller rather than moving to `modVbeUtility`.

Two existing module-qualified references (`modVCSUtility.GetVCSVersion` in clsVersionControl, `modVCSUtility.InteractionMode` in modAPI) both remain in the slimmed modVCSUtility — no caller updates needed. All other public functions are resolved by name within the project.

**What this rules out**: `modVCSUtility` no longer contains container management, VBE operations, or text I/O wrappers. Future container/component-related functions go in `modContainers`, VBE operations in `modVbeUtility`, and LoadFromText/SaveAsText wrappers in `modLoadSaveText`.

**Relevant files**:

- `Version Control.accda.src/modules/modContainers.bas` — new, split from modVCSUtility.bas
- `Version Control.accda.src/modules/modVbeUtility.bas` — new, split from modVCSUtility.bas
- `Version Control.accda.src/modules/modLoadSaveText.bas` — new, split from modVCSUtility.bas
- `Version Control.accda.src/modules/modVCSUtility.bas` — slimmed to remaining functions

---

## 2026-03-10 — Split modFunctions into modCollectionUtil, modStringUtil, modUIUtil

**Trigger**: `modFunctions.bas` was a 1,113-line, 41-function catch-all with no cohesion. An agent searching for "how to merge two dictionaries" had no reason to look in a file called `modFunctions`. The functions spanned collection/dictionary helpers, string manipulation, UI helpers, array utilities, null handling, date functions, and environment queries.

**Options explored**:

- **Keep as one module**: The generic name and mixed responsibilities made it the hardest module for new contributors to navigate. Rejected.
- **Split into two (data vs UI)**: Better but the data functions themselves span collections, strings, and arrays — very different concerns. Rejected as insufficient.
- **Split into four by domain**: Collection/dictionary helpers (9 functions), string manipulation (8 functions), UI/dialog helpers (4 functions), and remaining general utilities (20 functions). Each new module name immediately signals what it contains. Chosen.

**Decision**: Split into `modCollectionUtil.bas` (InCollection, MergeCollection, MergeDictionary, dNZ, KeyExists, SortCollectionByValue, SortDictionaryByKeys, DictionaryEqual, CloneDictionary), `modStringUtil.bas` (MultiReplace, Coalesce, DblQ, DeDupString, StartsWith, EndsWith, Repeat, LikeAny), `modUIUtil.bas` (ShowIDE, MsgBox2, MakeDialogResizable, ScaleColumns — includes window-style API declarations), and a slimmed `modFunctions.bas` (QuickSort, Pause, array helpers, null handling, file name encoding, SwapExtension, environment variables, etc.). The `Sleep` API declaration stays in `modFunctions` with `Pause`; the window-style API declarations move to `modUIUtil` with `MakeDialogResizable`.

No module-qualified references to `modFunctions` exist in the codebase — no caller updates needed.

**What this rules out**: `modFunctions` no longer contains collection/dictionary helpers, string manipulation, or UI code. Future collection/dictionary helpers go in `modCollectionUtil`, string utilities in `modStringUtil`, and UI/dialog helpers in `modUIUtil`.

**Relevant files**:

- `Version Control.accda.src/modules/modCollectionUtil.bas` — new, split from modFunctions.bas
- `Version Control.accda.src/modules/modStringUtil.bas` — new, split from modFunctions.bas
- `Version Control.accda.src/modules/modUIUtil.bas` — new, split from modFunctions.bas
- `Version Control.accda.src/modules/modFunctions.bas` — slimmed to remaining functions

---

## 2026-03-10 — Strengthen CRLF line ending preservation guidance for AI agents

**Trigger**: AI agents repeatedly converted CRLF line endings to LF when editing VBA source files. The existing documentation mentioned CRLF in a single table row in `Version Control.accda.src/AGENTS.md` with no explanation of consequences, no verification script, and no mention in the Cursor rule that activates during VBA file edits. By contrast, BOM encoding had extensive coverage (dedicated section, verification scripts, mandatory post-edit restoration). The `.gitattributes` file was also missing the newer file extensions (`.form`, `.report`, `.qdef`, `.macro`) introduced in export format 5.0.0, and no `.editorconfig` existed to enforce CRLF at the editor level.

**Options explored**:

- **Documentation-only fix (AGENTS.md + Cursor rule)**: Add warnings and verification scripts to the files agents actually read. Addresses the immediate problem but doesn't prevent editors from silently converting on save. Necessary but insufficient alone.
- **Config-file-only fix (.gitattributes + .editorconfig)**: Enforce CRLF via tooling. Git checkout would normalize, and editors with EditorConfig support would preserve CRLF. But AI agents don't always go through git checkout for their edits, and not all tools respect EditorConfig. Insufficient alone.
- **Both documentation and config files**: Belt-and-suspenders approach covering agent instructions, git normalization, and editor configuration. Chosen.

**Decision**: Four changes made in parallel: (1) Added "REQUIRED: Preserve CRLF Line Endings" section to `.cursor/rules/vba-source-files.mdc` with a PowerShell verification/restoration one-liner, matching the existing BOM restoration pattern. (2) Elevated CRLF from a table row to a full critical rule (Rule 2) in `Version Control.accda.src/AGENTS.md` with MUST/MUST NOT lists, verification script, and a new troubleshooting entry. Renumbered existing rules 2-3 to 3-4. (3) Added `eol=crlf` entries for `*.form`, `*.report`, `*.qdef`, `*.macro` to both `.gitattributes` and `.gitattributes.default`. (4) Created `.editorconfig` with `end_of_line = crlf` globally and `charset = utf-8-bom` for source file extensions, plus `trim_trailing_whitespace = false` and `insert_final_newline = false` to prevent editors from altering whitespace Access expects.

**What this rules out**: CRLF preservation is now a documented, enforced requirement at three levels (agent instructions, git config, editor config). Future source file extensions added to the project must be added to all three locations. If agents continue to introduce LF-only files despite these safeguards, the next step would be a pre-commit hook that rejects files with LF-only line endings.

**Relevant files**:

- `.cursor/rules/vba-source-files.mdc` — added CRLF section with verification script
- `Version Control.accda.src/AGENTS.md` — new Rule 2 (CRLF), troubleshooting entry, renumbered rules
- `.gitattributes` — added `.form`, `.report`, `.qdef`, `.macro` with `eol=crlf`
- `.gitattributes.default` — same additions (template distributed to users)
- `.editorconfig` — new file enforcing CRLF and UTF-8 BOM

---

## 2026-03-10 — Split modImportExport into modExport, modBuild, modSourceUpgrade

**Trigger**: `modImportExport.bas` was a 2,070-line, 20-procedure "god module" mixing export orchestration, build/merge orchestration, source file upgrade/migration, form initialization, legacy checks, and file format detection. As part of the v5 reorganization to improve navigability for new developers and AI agents, this was identified as the highest-impact split.

**Options explored**:

- **Keep as one module**: No change. Simple, but the module was doing too many things. A developer looking for "how does build work?" had to wade through export and migration code. Rejected.
- **Split into two (export vs build)**: Cleaner, but upgrade/migration logic is conceptually distinct from both export and build orchestration. Rejected as insufficient.
- **Split into three by responsibility**: Export (6 functions), Build/Merge (9 functions), Source Upgrade (4 functions). Each module has a clear single responsibility signaled by its name. Chosen.

**Decision**: Split into `modExport.bas` (ExportSource, ExportSingleObject, ExportMultipleObjects, ExportDependentObjects, ExportSchemas, RemoveThemeZipFiles), `modBuild.bas` (Build, LoadSingleObject, MergeAllSource, MergeDependentObjects, InitializeForms, OpenFormInCurrentDb, GetBackupFileName, GetFileFormat, PrepareRunBootstrap), and `modSourceUpgrade.bas` (CheckForLegacyModules, UpgradeSourceFiles, MigrateFileExtensions, RenameFilesInFolder). All three modules retain `Option Private Module` to stay hidden from external callers. Functions that were `Private` in the original module but are now called cross-module (UpgradeSourceFiles, MigrateFileExtensions, CheckForLegacyModules) were changed to `Public`, but `Option Private Module` keeps them internal to the add-in.

Git history was preserved using the project's built-in Split Files tool (`frmVCSSplitFiles` / `clsGitIntegration.SplitFilesWithHistory`) for `modBuild.bas` and `modSourceUpgrade.bas`, and `git mv` for `modExport.bas` (rename from `modImportExport.bas`). Three callers with explicit module-qualified references were updated: `clsVersionControl.cls`, `modExportOnSaveHook.bas`, `frmVCSMain.cls`.

**What this rules out**: `modImportExport.bas` no longer exists. All references to it should use the new module names. Future export-related functions go in `modExport`, build/merge functions in `modBuild`, and legacy/migration logic in `modSourceUpgrade`. If any of these modules grows beyond ~800 lines, consider further splitting by the same pattern.

**Relevant files**:

- `Version Control.accda.src/modules/modExport.bas` — renamed from modImportExport.bas
- `Version Control.accda.src/modules/modBuild.bas` — new, split from modImportExport.bas
- `Version Control.accda.src/modules/modSourceUpgrade.bas` — new, split from modImportExport.bas
- `Version Control.accda.src/modules/clsVersionControl.cls` — updated `modBuild.Build`, `modBuild.MergeAllSource`
- `Version Control.accda.src/modules/modExportOnSaveHook.bas` — updated `modExport.ExportMultipleObjects`
- `Version Control.accda.src/forms/frmVCSMain.cls` — updated `modExport.ExportSource`, `modExport.ExportSingleObject`

---

## 2026-03-10 — Rejected IDbComponent helper module for shared boilerplate

**Trigger**: During v5 reorganization review, the ~30 `clsDb*` classes implementing `IDbComponent` appeared to have significant boilerplate duplication. A `modComponentHelper.bas` was proposed to centralize shared logic.

**Options explored**:

- **Helper module with generic defaults** (DefaultGetAllFromDB, DefaultCount, DefaultMerge, etc.): Would centralize shared patterns. Initial analysis suggested 50-100 lines saved per class.
- **Composition/delegation pattern**: Each class holds a helper object that provides default implementations. More object-oriented, but VBA's lack of inheritance makes this awkward.
- **Keep boilerplate inline in each class**: Each class is self-contained and readable without jumping to another file.

**Decision**: After detailed comparison of every "boilerplate" method across 13+ classes, found that most methods have **meaningful per-class variations**: `GetAllFromDB` uses different collections, filters, and keys; `Merge` has 6+ distinct patterns (Forms protect add-in forms, TableDef stages relations, Property calls RemoveMissing, etc.); `MoveSource` moves different file sets; `DbObject` has custom loading for SharedImage/Theme. Only `Count` (always `GetAllFromDB(blnModifiedOnly).Count`) and `Parent` (always `Set Parent = Me`) are truly identical — but both are one-liners where extraction adds indirection without reducing code. Rejected the helper module entirely. The real improvement is **documentation** — adding comments to each class explaining its non-obvious Merge/GetAllFromDB/IsModified behavior.

**What this rules out**: No `modComponentHelper.bas` will be created. If a future refactoring introduces actual shared logic (e.g., a common conflict-detection step in Merge), a helper module can be reconsidered at that point. The per-class variations are real domain differences, not accidental duplication.

**Relevant files**: No files changed — this was a design decision to NOT create new abstraction.

---

## 2026-03-10 — Source file extension migration from .bas to descriptive extensions

> **⚠ Partially superseded** (2026-03-11): The claim that "Export, `GetFileList`, `FileExtensions`, and `SourceFile` use only the new extensions" is no longer true. These are now version-gated: format < 5.0.0 uses `.bas`, format >= 5.0.0 uses descriptive extensions. `GetFileList` searches for both. See "Version-gate file extensions and @Folder paths for export format downgrade" above.

> **⚠ Partially superseded** (2026-03-10): References to `modImportExport.bas` below should now read `modExport.bas` (export logic) and `modSourceUpgrade.bas` (migration logic). See "Split modImportExport into modExport, modBuild, modSourceUpgrade" above.

**Trigger**: The `.bas` extension was overloaded across five distinct content types: VBA standard modules (actual VBA code), forms, reports, queries, and macros (all Access `SaveAsText` proprietary format). This confused editors applying VBA syntax highlighting to non-VBA files, made it harder to distinguish file types at a glance, and conflicted with the legitimate `.bas` usage for VBA modules.

**Options explored**:

- **Per-type descriptive extensions** (`.form`, `.report`, `.macro`, `.qdef`): Full-word, unambiguous, zero collision with known formats. Parallels modern conventions (`.proto`, `.graphql`). Chosen.
- **Single unified extension** (`.axt` or `.sat` for all SaveAsText output): Simple to document, but loses per-file type distinction and relies entirely on subfolder names. Rejected as less intuitive.
- **Abbreviated extensions** (`.frm`, `.rpt`, `.mac`, `.qry`): Familiar feel, but `.frm` directly conflicts with VBE forms (`clsDbVbeForm` already uses `.frm`), `.rpt` is associated with Crystal Reports, `.mac` with macOS resource forks. Rejected due to collisions.

**Decision**: Forms use `.form`, reports use `.report`, queries use `.qdef` ("query definition" — distinguishes from the companion `.sql` file), macros use `.macro`. VBA modules keep `.bas`/`.cls` unchanged. Gated behind `EFV_5_0_0` (not a new version, since 5.0.0 hasn't shipped yet). Import methods accept both old `.bas` and new extensions for backward compatibility. Export, `GetFileList`, `FileExtensions`, and `SourceFile` use only the new extensions.

Fixed a latent bug in `clsDbQuery.Import`: two `Left$(strFile, Len(strFile) - 4)` calls hardcoded the `.bas` extension length (4 chars). With `.qdef` (5 chars) this would produce wrong paths. Replaced with `SwapExtension(strFile, "sql")`.

**What this rules out**: The `.bas` extension is no longer used for forms, reports, queries, or macros in export format >= 5.0.0. Adding new SaveAsText-based component types should follow this pattern of descriptive extensions. The abbreviations `.frm`, `.rpt`, `.mac`, `.qry` are ruled out — revisit only if a compelling external standard emerges. If a future component type's natural extension collides with an existing format, prefer full words.

**Relevant files**:

- `Version Control.accda.src/modules/clsDbForm.cls` — `.form` extension
- `Version Control.accda.src/modules/clsDbReport.cls` — `.report` extension
- `Version Control.accda.src/modules/clsDbQuery.cls` — `.qdef` extension, `SwapExtension` fix
- `Version Control.accda.src/modules/clsDbMacro.cls` — `.macro` extension
- `Version Control.accda.src/modules/modImportExport.bas` — migration logic, legacy cleanup
- `Version Control.accda.src/modules/clsOptions.cls` — `HasUnifiedLayoutFilesInGit` updated
- `Version Control.accda.src/modules/clsVCSIndex.cls` — `MigrateIndexExtension` method

---

## 2026-03-10 — Extension migration strategy: FSO.MoveFile + index key rename

> **⚠ Partially superseded** (2026-03-11): The reverse migration path (`.form`/`.report`/`.qdef`/`.macro` back to `.bas`) is now implemented via `RevertFileExtensions` in `modSourceUpgrade.bas`. See "Version-gate file extensions and @Folder paths for export format downgrade" above.

> **⚠ Partially superseded** (2026-03-10): `MigrateFileExtensions` and `RenameFilesInFolder` now live in `modSourceUpgrade.bas`, not `modImportExport.bas`. See "Split modImportExport into modExport, modBuild, modSourceUpgrade" above.

**Trigger**: When upgrading from old `.bas` extensions to new descriptive extensions, existing source files need to be renamed. For git repos, history preservation during the rename is desirable. The add-in already had `SplitFilesWithHistory` (branch-and-merge pattern) for splitting form layout from VBA code.

**Options explored**:

- **Reuse `SplitFilesWithHistory` branch-and-merge pattern**: Creates a temp branch, moves files, commits, restores originals, merges with `--no-ff`. Designed to produce two files that both have history. Overkill for a pure rename where the original should disappear. Rejected.
- **Add `git mv` command to `clsGitIntegration`**: Would stage renames atomically. But `RunGitCommand` is private, adding a new enum value requires modifying the class, and iterating hundreds of files one-at-a-time with shell calls is slow. Rejected as over-engineered.
- **`FSO.MoveFile` for all cases + index key rename**: Simple file rename, works with or without git. Git detects renames on commit via content similarity (100% match for identical content). Combined with renaming VCS index dictionary keys to prevent a full re-export. Chosen.

**Decision**: `MigrateFileExtensions` in `modImportExport.bas` runs on every export when `ExportFormatVersion >= EFV_5_0_0`. It scans each affected folder for old `.bas` files, renames them with `FSO.MoveFile`, then calls `VCSIndex.MigrateIndexExtension` to rename the corresponding dictionary keys. The `MigrateIndexExtension` method is generic and bidirectional — it takes a category name and target extension, iterates both `Components` and `AlternateExport` sections, and uses `Scripting.Dictionary.Key(old) = new` for in-place key rename. This supports reverting to `.bas` if a user drops back to a legacy export format version. `ClearFilesByExtension` calls in `UpgradeSourceFiles` serve as a safety net for any `.bas` stragglers missed by migration.

The index key rename was added specifically to avoid a costly full re-export on large projects. Without it, the stale `.bas` keys would cause the add-in to treat every form/report/query/macro as modified (no matching index entry), triggering `SaveAsText` for potentially hundreds of objects.

**What this rules out**: No git-specific commands are used for the migration — history preservation depends entirely on git's rename detection at commit time. This is reliable for identical content but could miss renames if the user also makes significant content changes in the same commit (similarity drops below git's 50% threshold). If this proves problematic, adding explicit `git mv` support to `clsGitIntegration` could be revisited. The `SplitFilesWithHistory` pattern remains available for future scenarios that genuinely need both files to retain history.

**Relevant files**:

- `Version Control.accda.src/modules/modImportExport.bas` — `MigrateFileExtensions`, `RenameFilesInFolder`, migration call in `ExportSource`
- `Version Control.accda.src/modules/clsVCSIndex.cls` — `MigrateIndexExtension`

---

## 2026-03-10 — Per-category option hashing for smart re-export

> **⚠ Partially superseded** (2026-03-10): `ExportSource()` with per-category stale detection now lives in `modExport.bas`, not `modImportExport.bas`. See "Split modImportExport into modExport, modBuild, modSourceUpgrade" above.

> **⚠ Supersedes** the `OptionsHash` mechanism described in "Export format versioning system" below. `OptionsHash` (single string) is replaced by `CategoryHashes` (per-category dictionary) in `clsVCSIndex`.

**Trigger**: Changing any export option (e.g., adding a table to `TablesToExportData`, toggling `ShowDebug`, adjusting print settings) triggered a full export of ALL database objects. On large databases this takes 30+ minutes, even when only a single component category is affected by the change.

**Options explored**:

- **Blacklist non-export options from hash**: Remove options like `ShowDebug`, `MaxLogFiles`, etc. from `GetHash()`. Simple, but still forces full export of everything when any remaining option changes — e.g., changing `ExtractThemeFiles` would still re-export all forms, reports, and queries. Rejected as insufficient.
- **Whitelist export options with flat hash**: Only hash the ~13 export-affecting options. Reduces false triggers but doesn't solve the cross-category problem. Rejected as a half-measure.
- **Per-category option hashing**: Compute a separate hash for each component category based on only the options that affect it. Store per-category hashes in the index. During export, only categories whose hash changed get full re-export; others use fast save. Chosen.

**Decision**: Replaced `OptionsHash` (single string) with `CategoryHashes` (Dictionary mapping category names to hashes) in `clsVCSIndex`. New `GetCategoryHashes()` function on `clsOptions` uses a `Select Case` that classifies every option into the categories it affects. Each category's hash includes its specific options plus global options (`ExportFormatVersion`, major Access version). A `Debug.Print` guard in `Case Else` names any unclassified option when a developer adds a new option to `m_colOptions` without classifying it.

Options are classified as:
- **Category-specific**: e.g., `SaveQuerySQL` affects Queries; `ExtractThemeFiles` affects Themes
- **Multi-category**: `SanitizeLevel` and `StripPublishOption` affect Forms, Reports, Queries, Macros, Tables, and Table Data Macros
- **Global**: `ExportFormatVersion` and major Access version — changing these triggers full export of all categories
- **Non-export**: 20 options (`ShowDebug`, `UseFastSave`, `TablesToExportData`, hooks, etc.) that don't affect exported file content and are excluded from all hashes

`TablesToExportData` is excluded because `clsDbTableData.IsModified` always returns `True` — table data is always exported regardless of fast save mode.

In `ExportSource()`, global hash changes set `blnFullExport = True` (same as user checking the Full Export box). Category-level changes build a `dStaleCategories` dictionary; the category loop checks `blnFullExport Or dStaleCategories.Exists(cCategory.Category)` per iteration.

**What this rules out**: The old `OptionsHash` string property on `clsVCSIndex` is removed. Old index files without `CategoryHashes` produce empty stored hashes, causing all categories to be treated as stale on first run (equivalent to the old full-export behavior). `GetHash()` still exists for backward compatibility, derived from `GetCategoryHashes()`. Future options must be added to the `Select Case` in `GetCategoryHashes()` — the `Debug.Print` guard catches omissions during development. When adding a new option that affects export output, add it to the appropriate category case(s); when adding a non-export option, add it to the skip case.

**Relevant files**:

- `Version Control.accda.src/modules/clsOptions.cls` — `GetCategoryHashes()`, `AddToCat()`, simplified `GetHash()`
- `Version Control.accda.src/modules/clsVCSIndex.cls` — `CategoryHashes` property, removed `OptionsHash`
- `Version Control.accda.src/modules/modImportExport.bas` — per-category stale detection in `ExportSource()`
- `Version Control.accda.src/forms/frmVCSMain.cls` — only force full export on global hash change

---

## 2026-03-06 — Export format versioning system

> **⚠ Partially superseded** (2026-03-10): References to `modImportExport.bas` below should now read `modExport.bas`. See "Split modImportExport into modExport, modBuild, modSourceUpgrade" above.

> **⚠ Partially superseded** (2026-03-10): The file extension migration was folded into `EFV_5_0_0` rather than adding a new `EFV_5_1_0`, since 5.0.0 has not shipped yet. The general pattern (add enum member, update `[_Last]`, gate with `>=`) remains correct for future post-release changes. See "Source file extension migration from .bas to descriptive extensions" above.

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

> **⚠ Partially superseded** (2026-03-10): References to `modImportExport.bas` below should now read `modExport.bas` (skip-count logging). See "Split modImportExport into modExport, modBuild, modSourceUpgrade" above.

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

## 2026-03-12 — Per-object companion .json for consolidated metadata

**Trigger**: `clsDbDocument` scans ~6,870 DAO documents to read the `Description` property on every export, costing ~18-20s of cold JET I/O. `clsDbHiddenAttribute` performs a similar full scan. Both produce monolithic singleton files (`documents.json`, `hidden-attributes.json`) because that mirrors how DAO exposes them via `Container.Documents`. However, document properties and hidden attributes are logically part of the objects they describe. During fast saves (the common case), only a handful of objects are modified, yet the full scan runs every time.

**Options explored**:

- **Skip the full scan during fast saves**: Only run the monolithic `clsDbDocument`/`clsDbHiddenAttribute` scan during full exports. Rejected because full exports are rare (days/weeks apart) while fast saves happen multiple times per day — descriptions would go stale for extended periods.
- **Targeted delta scan of modified objects against the monolithic file**: Scan only objects flagged as modified and merge into `documents.json`. Complex, and still suffers from the SingleFile limitation where every description change rewrites the entire file.
- **Per-object companion `.json` files** (chosen): Consolidate all per-object metadata (document properties, hidden attributes, print settings, linked table info) into companion `.json` files co-located with each object's primary source file. Each component's `Export` method performs O(1) lookups for its own metadata. The performance problem disappears by design.

**Decision**: Companion `.json` files use reserved keys under `"Items"`: `"Properties"` for document properties, `"Hidden"` for hidden attribute (only present when `True`). Existing keys (`"Printer"`, `"Margins"`, `"Connect"`, etc.) are unchanged. For forms/reports, metadata merges into the existing print settings `.json`. For linked tables, it merges into the existing linked table `.json`. For queries, macros, modules, and local tables, a new companion `.json` is created only when metadata exists.

`clsDbDocument` is reduced to only scan the "Databases" container (SummaryInfo, UserDefined) when `EFV >= 5.0.0`. `clsDbHiddenAttribute` returns an empty dictionary when `EFV >= 5.0.0`.

DAO container mapping: Forms→`"Forms"`, Reports→`"Reports"`, Queries→`"Tables"` (DAO quirk), Tables→`"Tables"`, Macros→`"Scripts"`, Modules→`"Modules"`.

**Change detection via MetaHash**: Access does not update an object's `DateModified` when its Description or Hidden attribute changes. Since companion `.json` files are only written during `Export`, and `Export` is only called for objects that `IsModified` returns `True` for, metadata-only changes would be silently missed. To address this, a lightweight `MetaHash` is stored in the VCS index during export. `GetMetadataHash()` reads just the Description property and Hidden attribute (two O(1) DAO calls) and returns a hash. Each component's `IsModified` compares the current `MetaHash` against the stored value as a final check after the existing DateModified/code-hash checks pass. This adds no file I/O — the comparison is entirely in-memory (VCS index) vs live DAO, and runs only for objects that appear unchanged by other checks.

When `SaveAllDocumentProperties = True`, all non-standard DAO properties are exported (not just Description). However, the `MetaHash` only covers Description + Hidden for fast-save detection. Custom property changes are captured on full export — an acceptable trade-off since custom properties are rare and typically accompany other object changes.

**Backward compatibility**: Import reads companion `.json` first; `clsDbDocument.Import` and `clsDbHiddenAttribute.Import` still process their singleton files for legacy source. A one-time migration in `modSourceUpgrade.UpgradeSourceFiles` distributes entries from `documents.json` and `hidden-attributes.json` into companion files.

**What this rules out**: The monolithic `documents.json` no longer contains per-object descriptions for `EFV >= 5.0.0` — only database-level properties (SummaryInfo, UserDefined). `hidden-attributes.json` is no longer written. Future per-object metadata should be added to the companion `.json` structure. Making the `.json` the primary source file for queries is deferred as a future direction.

**Relevant files**:

- `Version Control.accda.src/modules/Core/modLoadSaveText.bas` — `ExportObjectMetadata`, `ImportObjectMetadata`, `GetMetadataHash`, `HasNonMetadataKeys`
- `Version Control.accda.src/modules/Components/clsDbForm.cls` — Export/Import/IsModified with metadata helpers and MetaHash
- `Version Control.accda.src/modules/Components/clsDbReport.cls` — same pattern as forms
- `Version Control.accda.src/modules/Components/clsDbQuery.cls` — same pattern, add json to FileExtensions/MoveSource
- `Version Control.accda.src/modules/Components/clsDbTableDef.cls` — same pattern, update MoveSource
- `Version Control.accda.src/modules/Components/clsDbMacro.cls` — same pattern, add json to FileExtensions/MoveSource
- `Version Control.accda.src/modules/Components/clsDbModule.cls` — same pattern, add json to FileExtensions/MoveSource
- `Version Control.accda.src/modules/Components/clsDbDocument.cls` — reduced to Databases container only (EFV >= 5.0.0)
- `Version Control.accda.src/modules/Components/clsDbHiddenAttribute.cls` — returns empty dictionary (EFV >= 5.0.0)
- `Version Control.accda.src/modules/Core/modSourceUpgrade.bas` — `MigrateMetadataToCompanionFiles` migration logic
- `Version Control.accda.src/modules/Infrastructure/clsVCSIndex.cls` — `MetaHash` in `Update`, `LoadItem`
- `Version Control.accda.src/modules/Infrastructure/clsVCSIndexItem.cls` — `MetaHash` field

---

## 2026-04-03 — Remove BOM/CRLF workaround instructions from agent documentation

**Trigger**: Cursor fixed the underlying bug where `StrReplace` and `Write` tools stripped UTF-8 BOM bytes and converted CRLF line endings to LF. The extensive workaround instructions (mandatory post-edit PowerShell scripts, tool-distrust warnings, edit-size guidance to minimize corruption) added in earlier sessions were consuming significant token budget on every VBA source file edit with no remaining benefit.

**Decision**: Removed the workaround-specific content from agent documentation and Cursor rules while keeping the format requirements documented concisely as reference information. The `.editorconfig` and `.gitattributes` files (added as part of the 2026-03-10 belt-and-suspenders approach) remain in place as the primary enforcement mechanism.

Changes made: (1) Removed "Encoding", "REQUIRED: Restore BOM After Every Edit", "REQUIRED: Preserve CRLF Line Endings", and "Editing Safely" sections from `.cursor/rules/vba-source-files.mdc`. (2) Condensed Rules 1 and 2 in `Version Control.accda.src/AGENTS.md` from ~80 lines of MUST/MUST NOT lists, verification scripts, and warnings down to two brief sentences each, pointing to `.editorconfig` for enforcement. Removed repeated "Save with UTF-8 BOM encoding" steps from Common Tasks. (3) Removed the UTF-8 BOM reminder line from `.cursor/rules/project-guide.mdc`. (4) Added explanatory comments to `.editorconfig` since it is now the primary documentation point for these format constraints.

**What this rules out**: If Cursor regresses and reintroduces BOM stripping or CRLF conversion, the workaround instructions would need to be re-added. The `.editorconfig` and `.gitattributes` enforcement remains regardless.

**Relevant files**:

- `.cursor/rules/vba-source-files.mdc` — removed four workaround sections
- `Version Control.accda.src/AGENTS.md` — condensed Rules 1-2, removed Common Tasks encoding steps
- `.cursor/rules/project-guide.mdc` — removed BOM reminder line
- `.editorconfig` — added explanatory comments

---

## 2026-04-10 — Deterministic query export with performance optimization

**Trigger**: Query exports using `Application.SaveAsText` were non-deterministic (WHERE clause ordering, column metadata ordering varied between exports) causing VCS noise, and slow (~30 minutes for 2,800 queries due to per-query COM calls).

**Options explored**:

- **Keep `SaveAsText` and post-process for determinism**: Sanitize the output to normalize ordering. Rejected because it doesn't solve the performance problem (SaveAsText is the bottleneck) and the sanitization is fragile given the undocumented format.
- **Read `QueryDefs(name).SQL` directly**: Avoids SaveAsText but is still a slow per-query COM call. Doesn't capture design layout, column metadata, or properties without additional COM calls. Rejected.
- **Read MSysQueries + MSysObjects system tables directly** (chosen): Single SQL queries can bulk-read all query data. `MSysQueries` contains the decomposed query structure (one row per clause). `MSysObjects.LvProp` stores properties and column metadata in the same MR2 binary format already parsed for linked tables. `MSysObjects.LvExtra` stores Design View layout. Both blobs are sub-millisecond to read per query. SQL is reconstructed deterministically from the decomposed structure.

**Decision**: Replace `SaveAsText` + `QueryDefs.SQL` with direct reads from `MSysQueries` and `MSysObjects` system tables. Export produces `.sql` (source of truth for SQL text) + `.json` (metadata: properties, columns, design layout, description, hidden). The `.qdef` file is no longer exported.

**Architecture**:

- `clsQueryComposer`: Bidirectional SQL/structure translation class. `ReconstructSQL()` builds SQL from MSysQueries rows on export. `DecomposeSQL()` parses SQL back into structure on import. `GenerateQdef()` emits Design View or SQL View `.qdef` text for `LoadFromText`.
- `clsLvExtraParser`: Parses the LvExtra binary blob (magic `0x99 0x99 0xCE 0xAC`, window/pane RECTs, table positions as null-terminated UTF-16LE strings). Format reverse-engineered from live data.
- `clsLvPropParser`: Existing class, verified to work on query LvProp blobs (same MR2 format as linked tables).
- Import flow: `.sql` → `DecomposeSQL()` → check `IsDesignerCompatible()` → generate Design View `.qdef` (with layout from `.json`) or SQL View `.qdef` → `LoadFromText` → apply metadata from `.json`. Falls back to SQL View if Design View import fails.
- Backward compatibility: Legacy `.qdef`/`.bas` files are still accepted for import. `GetFileList` searches for `.sql` first, then `.qdef`/`.bas`. Legacy files are cleaned up on next export.

**LvExtra binary format** (reverse-engineered):

| Offset | Size | Content |
|--------|------|---------|
| 0-3 | 4 | Magic: `99 99 CE AC` |
| 4-15 | 12 | Padding: `0xAA` × 12 |
| 16-31 | 16 | Window RECT (Left, Top, Right, Bottom as Longs) |
| 32-35 | 4 | State (Long) |
| 36-51 | 16 | Designer pane RECT |
| 52-59 | 8 | Grid origin (Left, Top) |
| 60-63 | 4 | ColumnsShown (Long) |
| 64-67 | 4 | Table count (Long) |
| 68+ | var | Per table: 5 Longs (L,T,R,B,scrollTop) + 2 null-term UTF-16LE names |

**MSysQueries findings** (vs isladogs documentation):

- Attribute 6 (field references): Expression column, not Name1
- Attribute 11 (ORDER BY): Expression column, not Name2
- Undocumented columns: `Order` (Binary, 510 bytes), `LvExtra` (Long, always NULL)
- `MSysObjects.LvExtra IS NOT NULL` reliably indicates Design View save

**What this rules out**: `SaveAsText` is no longer used for query export (still used for forms, reports, macros). The `SaveQuerySQL` option and `ForceImportOriginalQuerySQL` option are superseded by the new format. The decomposed query structure is never stored in files — it exists only transiently during composition/decomposition. Future changes to Access SQL dialect (new keywords, syntax) may require updates to `clsQueryComposer`.

**Relevant files**:

- `Version Control.accda.src/modules/Utility/clsQueryComposer.cls` — new: bidirectional SQL/structure/qdef translation
- `Version Control.accda.src/modules/Utility/clsLvExtraParser.cls` — new: LvExtra binary parser
- `Version Control.accda.src/modules/Components/clsDbQuery.cls` — rewritten: Export reads system tables, Import generates .qdef on-the-fly
- `Version Control.accda.src/modules/Utility/clsLvPropParser.cls` — verified: works for query LvProp blobs as-is
- `Version Control.accda.src/AGENTS.md` — updated: Query Files section for .sql + .json format
- `docs/how-access-stores-queries.md` — corrections to MSysQueries attribute documentation

---

## 2026-04-15 — Session-scoped option overrides for MCP/API callers

**Trigger**: When the MCP agent sets an option (e.g., `BreakOnError = True`) via `SetOption`, the change was silently discarded because every operation entry point resets `Options` and reloads from `vcs-options.json`. The agent's overrides never survived past the first subsequent operation.

**Options explored**:
- **Edit `vcs-options.json` directly** — corrupts user config on failure, race conditions, violates thin-wrapper principle.
- **In-memory overrides dictionary** — lost on Access restart; invisible; complex `ReleaseObjects` coordination.
- **Pass options as operation parameters** — changes VBA API signatures, awkward across COM. Deferred as a possible future enhancement.
- **Skip reload when called via API** — agent operates with stale options for everything, not just its overrides.
- **Single shared override file** — no session isolation; stale overrides bleed into interactive use.
- **Session-scoped override files in `mcp/` subfolder (chosen)** — each MCP/API session gets its own override file. Files are `.gitignored`. The user's `vcs-options.json` is never touched.

**Decision**: `SetOption` now persists overrides to `mcp/options-{session_id}.json` alongside `vcs-options.json`. After every `LoadProjectOptions` call, if `Operation.Source` is `eosMCPTool` or `eosExternalAPI`, `LoadOptionOverrides` scans the `mcp/` subfolder and merges matching override files on top. Interactive ribbon operations never see them. Stale files are auto-cleaned after 30 days. The MCP server generates a random session ID at startup, registers it via `RegisterSession`, and calls `EndSession` on shutdown to delete the override file.

**What this rules out**: Overrides do not persist across MCP server restarts (the server generates a new session ID each time). If two agents concurrently interact with the same database, their override files may both be loaded — this is an accepted tradeoff. If the MCP spec adds persistent session IDs (SEP-1364), we can adopt them as the session component without changing the file-based mechanism.

**Relevant files**:
- `clsOptions.cls` — `LoadOptionOverrides`, `MergeOverrideFile`, `CleanupStaleOverrides`
- `clsVersionControl.cls` — `SetOption` (updated), `SaveOptionOverride`, `RegisterSession`, `EndSession`
- `modObjects.bas` — `SessionId` property (survives `ReleaseObjects`)
- `modExport.bas`, `modBuild.bas` — `LoadOptionOverrides` calls gated on `Operation.Source`
- `main.py` — session ID generation, `atexit` cleanup
- `tools.py` — `vcs_set_option` registers session, `vcs_end_session` tool added

---

## 2026-05-07 — Cross-table ON condition LeftTable/RightTable in Design View qdef

**Trigger**: A production database had four queries that passed SQL builder validation but failed with DAO error 3082 ("JOIN operation refers to a field that is not in one of the joined tables") after a full build from source. The queries used compound `ON` clauses where individual conditions referenced different table pairs, and one table was also used inside a saved subquery referenced in another condition.

**Root cause**: `clsQueryComposer.EmitDesignViewQdef` reused the parent join's `leftTable`/`rightTable` for all split conditions in a compound `ON` clause. Access stores each compound `ON` condition as a separate Attribute 7 row in `MSysQueries` with its own `Name1`/`Name2` (the specific table pair for that condition). The emitter's reuse of the parent join's tables produced a `.qdef` where the `RightTable` for a condition referencing table `C` was set to table `B` (the parent join's right table). `LoadFromText` accepted this silently, but the resulting internal storage confused Access's scope resolution at execution time.

**Options explored**:
- **Fall back to `QueryDefs(name).SQL` (legacy path)** — rejected: the new pipeline was designed to generate its own `.qdef` rather than receive a pre-baked one, and falling back to the legacy path would lose design layout. The bug was in the emitter, not in `LoadFromText`.
- **Store per-condition table pairs in the `.json` companion** — rejected: the table pair for each condition is derivable from the condition expression itself (e.g., `tblFunds.FundID = tblAssociates.fldFundID` clearly references `tblFunds` and `tblAssociates`). Adding explicit storage would be redundant.
- **Extract per-condition table pairs from the expression at emit time (chosen)** — the emitter already has `ExtractTableFromOnSide` available. Using it for each split condition, with a fallback to the parent join's tables if extraction fails, is correct, minimal, and preserves backward compatibility.

**Decision**: `EmitDesignViewQdef` now calls `ExtractTableFromOnSide(condition, True)` and `ExtractTableFromOnSide(condition, False)` for each individual condition in a split compound `ON` clause. Falls back to the parent join's `leftTable`/`rightTable` only if extraction returns empty.

**Why this was hard to diagnose**: The SQL builder validation compares `ReconstructSQL` output against `QueryDefs.SQL` — a text-level check. The bug was not in SQL reconstruction but in `.qdef` emission, and `LoadFromText` accepted the wrong structure silently. The error only surfaced at query execution time, where the misleading error message ("field not in one of the joined tables") pointed away from the actual root cause (wrong `LeftTable`/`RightTable` metadata).

**Relevant files**:
- `clsQueryComposer.cls` — `EmitDesignViewQdef`: per-condition `LeftTable`/`RightTable` extraction
- `docs/access-query-storage.md` § 6 — documents the finding
- `Testing/Fixtures/queries/regression/qryRegressionCrossTableOn.notes.md` — regression context

---

## 2026-05-05 — VBProject.Saved + DateModified fast path for VBA code hashing

**Trigger**: Fast-save exports were spending significant time hashing every VBA module's code (via `GetCodeModuleHash` → `CodeModule.Lines(1, 999999)` → SHA256) even when no VBA code had changed since the last export. For a project with 110+ modules, the "Get VBA Hash" operation dominated the scan phase.

**Key empirical findings** (tested against `Version Control.accda` with 110 modules, 17 forms):

1. `VBProject.Saved` (Boolean) reliably detects all unsaved VBE changes, including VBA's automatic case-sync propagation across modules. Goes `False` on any in-memory edit, `True` after any save.
2. `CurrentProject.AllModules(name).DateModified` is a VBE-level property (NOT from `MSysObjects`). Always identical across all modules. Updates in real-time from VBE memory, even without saving.
3. `MSysObjects.DateUpdate` is a separate DAO-level per-row write timestamp with millisecond precision. Only updates on actual disk writes. Does NOT reflect VBE code edits. DOES reflect DAO property changes (e.g., Description). These are two completely different dates from different subsystems.
4. Saving any single module triggers a full VBA project write that updates `DateModified` on all 110 modules simultaneously. Saving a form's code-behind also updates all 110 module dates, but only that form's `DateModified` changes.
5. `CurrentProject.AllModules` does NOT include form/report code-behind — those are `vbext_ct_Document` components in the VBE.

**Options explored for the fast-path guard**:
- **DateModified only** — rejected: VBA case-sync changes `CodeModule.Lines()` without updating `DateModified`, so the date alone could miss changes.
- **Force compile-and-save before export** — rejected: would fail on uncompilable code, which the add-in must support exporting.
- **VBProject.Saved + DateModified (chosen)** — `Saved = True` means no dirty VBE memory (covers case-sync); `DateModified` match confirms nothing was saved since last export. Both must pass to skip hashing.

**Options explored for index storage of module dates**:
- **Per-module ObjectDate (existing)** — rejected: all 110 values are always identical, and partial exports only update N entries, leaving the other 110-N stale until a full export "heals" them.
- **Per-module ObjectDate with post-export healing pass** — rejected: unnecessary iteration when a single value suffices.
- **Top-level VBAProjectDate (chosen)** — one value in the index, updated whenever any module is exported. Eliminates redundant storage, eliminates the healing problem, eliminates 110 per-module COM property reads during change detection.

**Decision**: Two-tier guard in `clsDbModule.IsModified`: (1) `CurrentVBProject.Saved = True`, (2) `AllModules(0).DateModified = VCSIndex.VBAProjectDate`. When both pass, skip `GetCodeModuleHash` entirely. `MetaHash` check always runs (metadata changes don't affect `Saved` or `DateModified`). For forms/reports, the same `VBProject.Saved` guard skips the code-behind hash when the layout `DateModified` also matches.

Additionally, unsaved VBA project changes are now persisted at the start of the export flow (alongside `CloseDatabaseObjects`), ensuring exported source always reflects the current VBE state and preventing the scenario where a user exports code then discards changes on close.

**Performance results** (no-change fast-save export):
- Before: 0.88s total, 127 `Get VBA Hash` calls (0.09s), 286 `Compute SHA256` calls (0.15s)
- After: 0.44s total, 0 `Get VBA Hash` calls, 159 `Compute SHA256` calls (0.05s)
- 50% faster overall; `Get VBA Hash` completely eliminated

**What this rules out**: Per-module `ObjectDate` is no longer written for module components (other types still use it). The binary index format version was bumped from 2 to 3, so existing index files are rebuilt on first use. `MSysObjects.DateUpdate` was investigated but provides no advantage over `AllModules.DateModified` for VBA change detection. `CompileAndSaveAllModules` is intentionally NOT added to the export flow — it would break on uncompilable code.

**Relevant files**:
- `clsVCSIndex.cls` — new `VBAProjectDate` top-level property, format version 3, `Update` sets `VBAProjectDate` instead of per-module `ObjectDate` for modules
- `clsDbModule.cls` — `IsModified` uses `VBProject.Saved` + `VBAProjectDate` fast path
- `clsDbForm.cls` — `IsModified` skips code-behind hash when `VBProject.Saved = True` and layout date matches
- `clsDbReport.cls` — same as `clsDbForm.cls`
- `modExport.bas` — saves VBA project before export scan, wraps `CloseDatabaseObjects` in `Perf.PauseTiming`/`ResumeTiming`, fixes `Exit Sub` → `GoTo CleanUp` with `eelCritical`

---
