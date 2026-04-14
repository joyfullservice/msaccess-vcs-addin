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
