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

Insert new entries immediately below the `<!-- END HEADER -->` marker and the
`---` that follows it, newest first — not below the introduction at the top of
this header, which splits the header in two. Do not modify or reorder existing
entries except to add supersession notes (see below).
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

## 2026-08-11 — Emit the parameter type keywords ACE accepts, not the ones that read best

**Trigger**: While pinning the type map for the `Begin Parameters` work below,
each `PARAMETERS` keyword was round-tripped through `CreateQueryDef` to learn
the DAO type it actually produces. Four of the keywords the exporter emitted
turned out not to parse at all: `Boolean`, `Memo`, `OLEObject` and `Decimal`
each raise error 3139, "Syntax error in PARAMETERS clause". Two others were
simply mis-mapped — `BigInt` reports as `dbNumeric` (19) rather than `dbBigInt`
(16), and `Value` is a spelling of `dbText` (10) rather than an untyped
parameter. The map had been written from the DAO type names, which look
authoritative but are not the keywords the SQL parser accepts.

This was not cosmetic. On the SQL View import path the whole statement is
handed to Access as one memo, so a query declaring a Yes/No parameter did not
merely lose its type — the object failed to import outright with "Could not
create or set the property SQL". Any query with a Yes/No, Memo or OLE Object
parameter was therefore unbuildable from source, and had been since the map was
written.

**Options explored**:
- **Keep the readable spellings and special-case the SQL View path** — rejected:
  it leaves exported `.sql` that Access itself cannot parse, which defeats the
  point of a text format users are expected to read, diff and hand-edit.
- **Emit only what ACE accepts, and accept the rest on import** (chosen) — the
  exporter emits the canonical keyword; the readable spellings stay in the
  reverse map so source written by earlier versions of the add-in still
  rebuilds. Import compatibility is a standing requirement, so nothing is lost
  by keeping them.

**Decision**: `ParameterTypeSql` and `ParameterFlagFromType` are now generated
from one list (`EnsureParameterTypeTables`) rather than being two hand-kept
`Select Case` blocks that had already drifted. Each entry names the canonical
keyword followed by its import-only aliases, so the two directions cannot
disagree; where two DAO types claim the same keyword the first registration
wins the reverse lookup. Only the keywords verified against the parser are
emitted: `Bit`, `LongText`, `LongBinary`, `BigInt` for the four that were
wrong, and the unchanged spellings elsewhere. `Single`/`Double`/`GUID` are kept
as-is even though Access normalizes them to `IEEESingle`/`IEEEDouble`/`Guid`,
because all three parse correctly and changing them would churn every existing
export for no functional gain.

**What this rules out**: Adding a type to the map from the DAO constant name
alone. A new entry needs the keyword confirmed against the parser first —
`CreateQueryDef` with `PARAMETERS [P] <keyword>;` either succeeds and reports a
type or raises 3139. `Decimal` in particular has no parseable spelling, so
there is no way to declare a decimal parameter in Access SQL at all.

**Relevant files**:
- `Version Control.accda.src/modules/Utility/clsQueryComposer.cls` —
  `EnsureParameterTypeTables`, `AddParameterType`, `ParameterTypeKey`
- `Version Control.accda.src/modules/Tests/SQL/clsTestQueryComposerParameters.cls`
  — `TestExportedTypeKeywordsAreParseable` closes the loop against a list of
  parser-accepted keywords
- `Testing/Fixtures/queries/regression/qryRegressionDesignViewParameterTypes.*`
  — fixture updated from `Boolean` to `Bit`

---

## 2026-08-11 — Design View queries carry parameters in a `Begin Parameters` block after `OutputColumns`

**Trigger**: A parameterized query rebuilt from source lost its `PARAMETERS`
declaration whenever it took the Design View import path (joyfullservice#744).
`clsQueryComposer.EmitDesignViewQdef` assembled the structured `.qdef` that
`Application.LoadFromText` consumes but never emitted the declared parameters,
so the rebuilt query came back with an empty parameter collection. The SQL View
path was unaffected — there the whole statement is handed to Access as one SQL
memo — so the loss only showed up for designer-built queries whose stored grid
layout forces `blnDesignView = True`. Every existing parameter fixture was
SQL-View shaped, so nothing exercised the gap.

**Options explored**:
- **Prepend a SQL `PARAMETERS ...;` line to the structured `.qdef`** — the
  obvious guess, mirroring the SQL memo grammar. Rejected empirically:
  `LoadFromText` fails with `Expected: 'Operation'. Found: PARAMETERS.` The
  structured format has no grammar for an inline parameters declaration.
- **Force every parameterized query onto the SQL View path** — sidesteps the
  emitter gap by never taking Design View for a query that declares parameters.
  Rejected: it discards the designer layout that Design View exists to
  preserve, regressing the queries this add-in tries hardest to round-trip.
- **Emit the native `Begin Parameters` block** (chosen) — parameters live in a
  single `Begin Parameters ... End` block of repeated `Name =` / `Flag =` pairs,
  where `Flag` is the DAO type. All parameters share one block; there is not
  one block per parameter.

**Decision**: `EmitDesignViewQdef` emits the block immediately after
`Begin OutputColumns ... End` and before `Begin Joins`. That position is not a
style choice — it is the only one Access accepts. Feeding `LoadFromText` the
same qdef with the block moved elsewhere fails every time: after `Joins` or
`OrderBy` with `Expected: End of file. Found: Parameters.`, after `Groups` with
`Expected: 'End'. Found: Parameters.`, and before `InputTables` with
`Expected: End of file. Found: InputTables.` The position was then confirmed
against native `Application.SaveAsText` output for six designer-built shapes —
single table, join with `ORDER BY`, `GROUP BY`, `TOP n`, parameterized `UPDATE`
and crosstab — which agree regardless of query type or how elaborate the output
columns block is.

Parameter names are emitted **verbatim**: Access records a parameter exactly as
declared, so `[Enter ID]` keeps its brackets while an unbracketed
`StatusFilter` stays bare, even though Access brackets the matching reference
inside the `WHERE` clause. Re-bracketing on the way out would not round-trip.
The clause is parsed once, in `ParseParametersClause`, into a structured
`m_colParameters`; `EmitParameters` is then a loop. Splitting reuses the
existing bracket- and paren-aware `SplitTopLevel` rather than a private
splitter, so a comma inside `[Last, First]` or `Text ( 255 )` does not break
the list apart.

**No export-format-version gate and no `GetExporterRevisions` bump.** Both
mechanisms govern *export* output; the `.qdef` is an import-only intermediate
generated to a temp file at import time (`clsDbQuery` exports `.sql` + `.json`
only). This change alters import behavior exclusively, and import stays
backward compatible: older sources that never carried parameters simply produce
no block.

**What this rules out**: The earlier working hypothesis — that native Access
always serializes a parameterized query as a SQL memo and the structured format
has no place for parameters — is disproven and should not be revisited.
Emitting parameters as a leading SQL line is a dead end. Any future change that
reorders the structured blocks must keep parameters immediately after
`OutputColumns`; "somewhere ahead of the properties" is not sufficient.

**Relevant files**:
- `Version Control.accda.src/modules/Utility/clsQueryComposer.cls` —
  `ParseParametersClause`, `EmitParameters`, `SplitParameterToken`
- `Version Control.accda.src/modules/Components/clsDbQuery.cls` — labels the
  composer before `DecomposeSQL` so parse-time warnings name their query
- `Version Control.accda.src/modules/Tests/modTestRoundtrip.bas` — `import_path`
  check, so a silent Design View → SQL View fallback fails by assertion
- `Testing/Fixtures/queries/regression/qryRegressionDesignViewParameters*.*` —
  round-trip fixtures for the single-table, join/`ORDER BY`, `GROUP BY`,
  `TOP n`, `UPDATE` and crosstab shapes
- `docs/access-query-storage.md` — `.qdef` block order and Attribute 2 reference

---

## 2026-08-11 — Legacy index entries without AllFilesHash treat multi-file components as modified

**Trigger**: Issue #748 — Merge reported "No changes found" after editing only a form's
companion `.cls` file. The reporter's diagnosis pointed at the primary-file-only
content-hash fallback in `GetModifiedSourceFiles`. A local repro against Database5
confirmed it, but only for an index whose entries were written before `AllFilesHash`
existed (5.0.1 and earlier). Fresh exports under current code already record the
combined hash and detect `.cls`-only edits correctly.

The failure mode is worse than a single missed merge. The legacy branch compared
only the primary `.form` content hash; when that matched, it *refreshed*
`FilePropertiesHash` to the edited tree's dates and sizes. The edited `.cls` was
then recorded as synced, every later merge took the clean fast path, and the next
export could overwrite the user's VBA with no conflict prompt. A fast-save export
only re-indexes changed database objects, so a form whose Access side never
changes never heals.

A secondary gotcha made this look like an export-format problem during diagnosis:
after the add-in self-updated on disk (`Updated VCS (5.0.1 -> 5.1.0)`), Access kept
running the stale in-memory 5.0.1 project until the instance was restarted. Full
exports under that session wrote no `AllFilesHash` and logged no
`Get File Content Hash` operations, even though the source tree already contained
the function. Always restart Access after an add-in self-update before trusting a
repro against new detection logic.

**Options explored**:
- **Document that upgrading users need one full export**: Rejected. Fast-save
  exports skip unchanged objects, so the advice would not heal a form whose
  database side never changes, and users already have a silent data-loss window
  between upgrade and that export.
- **Backfill only**: Silently populate `AllFilesHash` when the property hash still
  matches. Rejected as sole fix — an edit made before the first post-upgrade merge
  would still be missed and then recorded as synced by the legacy refresh.
- **Conservative only**: When a legacy entry lacks `AllFilesHash` and more than one
  indexed file exists, always report modified. Correct, but every multi-file
  component would re-import on every merge until something else populated the
  combined hash (a full export, or an actual content change).
- **Conservative plus backfill on the clean fast path (chosen)**: Report multi-file
  legacy entries as modified when the property hash differs (cannot prove clean),
  and when the property hash matches, record `AllFilesHash` now so later scans
  arbitrate companion edits precisely. Gate both on `GetSourceFileCount > 1`
  (files that *exist*, not `FileExtensions.Count`) so bare modules keep the
  precise primary-hash comparison.

**Decision**: The 2026-07-29/07-30 property-hash short-circuit is deliberately *not*
trusted as a content audit for legacy entries lacking `AllFilesHash`. Matching
dates and sizes prove the tree is in its last-synced state, which is enough to
*record* the combined hash, but a mismatch cannot be resolved by the primary file
alone. Correctness wins over the one-time scan cost: a no-change merge still saves
the index (`blnSuccess = True` before `VCSIndex.Save`), so both the backfill and
any one-time imports persist. Steady state after the first post-upgrade merge adds
only a `Len(AllFilesHash)` test per file on the fast path.

**What this rules out**: Refreshing `FilePropertiesHash` from the legacy
primary-hash branch when the component has companion files on disk. Treating
`FileExtensions.Count` as a multi-file signal — `clsDbModule` reports `bas`,
`cls`, and `json` but only one of `bas`/`cls` is ever present. Relying on "until
the next export re-syncs this entry" as a healing path for unchanged database
objects.

**Relevant files**: `Version Control.accda.src/modules/Infrastructure/clsVCSIndex.cls`,
`Version Control.accda.src/modules/Core/modContainers.bas`
(`GetSourceFileCount`, `GetSourceFilesContentHash`),
`Version Control.accda.src/modules/Tests/Core/modTestMergeDetection.bas`.
See also 2026-07-29 — Merge scan reads no file content when dates and sizes are
unchanged; issue #748.

---

## 2026-08-10 — SharedDb invalidation after single-object table create

**Trigger**: Table-definition round-trip fixtures raised Error 3265
(`Item not found in this collection`) on
`SharedDb.TableDefs(m_Table.Name).Connect` inside `clsDbTableDef.IDbComponent_Export`.
The harness creates sandbox tables through `modTableDefBuilder` (DAO on a fresh
`CurrentDb`) or `Application.ImportXML`, then exports through `SharedDb`, whose
`TableDefs` collection is a snapshot from when the cached handle was opened.
Fixture 1 usually survived (handle created after its table existed); later
fixtures failed. The DAO fast path's verification export hit the same 3265 under
`On Error Resume Next` and appeared to succeed with an empty `Connect`.

**Options explored**:
- **Per-collection `TableDefs.Refresh` on the cached handle**: Already rejected in
  2026-06-25 ("SharedDb invalidation during build/merge and database close") —
  other collections (`QueryDefs`, `Containers.Documents`) have the same staleness.
- **Harness-only `ReleaseDbReferences`**: Would mask the production gap for
  `LoadSingleObject` (no per-category release) and for
  `FixCorruptedBigIntFields` / metadata apply after `ImportXML`. Rejected as sole
  fix.
- **Invalidate at create boundaries plus harness catalog refresh** (chosen):
  `ReleaseDbReferences` after `CreateTableFromSchema` appends a table, and after a
  successful `Application.ImportXML` fallback in `IDbComponent_Import`. The
  round-trip harness also calls a `RefreshDbCatalog` helper (idle + release) after
  import, cleanup, and scaffold load — same pattern as
  `modTestTableDef.RefreshTableCollections`.

**Decision**: Single-object table create and the round-trip harness are their own
invalidation boundaries, in addition to the per-category release during
build/merge. The DAO fast path must invalidate between create and verify so
verification does not rely on swallowed 3265.

**What this rules out**: Relying on `On Error Resume Next` around export to paper
over a stale `SharedDb` after DAO create. Treating per-category release in
`modBuild` as sufficient for `LoadSingleObject` or fixture harnesses.

**Relevant files**: `Version Control.accda.src/modules/Components/modTableDefBuilder.bas`,
`Version Control.accda.src/modules/Components/clsDbTableDef.cls`,
`Version Control.accda.src/modules/Tests/modTestRoundtrip.bas`,
`Version Control.accda.src/modules/Infrastructure/modObjects.bas`.
See also 2026-06-25 — SharedDb invalidation during build/merge and database close.

---

## 2026-08-10 — Function-call operands in ON clauses must resolve against InputTables

**Trigger**: A production merge of `qryUserDynamoGainLossBySecurityYearEnd` logged
`Join reference 'DateAdd('yyyy', -1, cur' not found in InputTables block` and then
reported success. The stored query failed at runtime with DAO error 3080. The ON
clause had a multi-condition join whose third equality put `DateAdd(...)` on one
side.

**Root cause**: `ExtractTableFromOnSide` treated any text before the first
qualifying dot as a table name, so a function-call operand produced a garbage
token. The 2026-05-07 per-condition emit path fell back to the parent join's
tables only when extraction returned empty, so the garbage token was emitted as
`RightTable`. `LoadFromText` accepted it; the failure was deferred to execution.

**Options explored**:

- **Gate the shape out of Design View** (`IsDesignerCompatible = False` when an
  ON operand is an expression) — rejected: multi-condition ON *requires* Design
  View because SQL View `dbMemo "SQL"` is rejected by `LoadFromText` for that
  shape. Leaving Design View means we must emit valid join rows.
- **Always reuse the parent join's LeftTable/RightTable for every split
  condition** — rejected: that is exactly the 2026-05-07 cross-table ON bug;
  individual conditions can reference different table pairs.
- **Resolve each side independently, preferring a qualifier scan over the parent
  join** — implemented first, then rejected. It fixed the reported case but
  changed single-table predicates: for `tblCarsColour.ID > 0` the scan claims
  that table for the left side and the right side then falls to the parent's
  identical value, emitting `LeftTable = RightTable`. That drifted the
  `qryRegressionMultiCondJoin` baseline and reintroduces the pair collapse that
  breaks `BuildJoinChain` on the next export. Per-side resolution has no way to
  tell "this side is unknown" from "this condition only names one table".
- **Rank whole pairs by coverage of the condition's refs (chosen)** — a
  candidate pair is acceptable when every `InputTables` ref named in the
  condition equals its left or right value, which is the same invariant the
  round-trip harness enforces. `ResolveConditionJoinTables` tries per-condition
  extraction first (so a cross-table condition still wins over the parent pair,
  preserving the 2026-05-07 fix), then the parent pair (which covers
  single-table predicates and non-equi conditions), then extraction plus a ref
  named in the condition, normalized to parent orientation when it is merely a
  swap because outer-join `Flag` values are orientation-sensitive.

**Decision**: Per-condition join refs are resolved through
`ResolveConditionJoinTables`, which ranks whole candidate pairs by ref coverage,
rather than raw `ExtractTableFromOnSide` plus an empty-only fallback. This
supersedes the fallback rule in the 2026-05-07 cross-table ON entry.

`ValidateJoinRow` in the round-trip harness gained the three invariants that
would have caught both this bug and the per-side misstep above: refs must be
non-empty (an empty ref previously skipped validation entirely), refs must exist
in the `InputTables` block, and `LeftTable` must differ from `RightTable`. All 56
committed `.qdef` baselines already satisfy them.

**What this rules out**: Emitting any `LeftTable`/`RightTable` that is not in
`InputTables` when a better known ref can be recovered from the condition. Also
rules out treating "extraction returned something" as proof that the something is
a table name, and rules out per-side resolution as the shape of this fix.

**Relevant files**:
- `clsSqlSyntax.cls` — `ExtractTableFromOnSide` / `IsSimpleOnSideIdentifier`
- `clsQueryComposer.cls` — `ResolveConditionJoinTables`, `CollectConditionInputRefs`, `PairCoversRefs`, emit loop, parse-time join branches
- `modTestRoundtrip.bas` — `ValidateJoinRow`, `ParseQdefInputTableRefs`
- `Testing/Fixtures/queries/regression/qryRegressionFunctionInOnClause.*`
- `docs/access-query-storage.md` § 5 / § 6

---

## 2026-08-07 — Per-user toggle to disable the helper script (`Worker.vbs`)

**Trigger**: Issue #727. A user's endpoint protection (Sophos "Lockdown") blocks Access from launching `Worker.vbs`, the script `clsWorker` extracts into the install folder for the jobs that cannot run in-process. The visible failure was an uninstall that reported "Success!" and then left the add-in and its lock file behind, but the same block affects every worker consumer, and the user cannot whitelist anything to work around it. Without an escape hatch, v5 is unusable in that environment.

**Options explored**:

- **Storage: per-user registry under `PROJECT_NAME\Install` (chosen)** vs. per-project `vcs-options.json`. Uninstall and add-in rebuild are not scoped to a project, and the reason to disable is environmental rather than a property of any source tree, so a project option would be both wrong-scoped and unreachable at the moments it matters most. The install form already writes registry settings and does not itself use the worker, so the toggle is always reachable.
- **Gating location: `CallWorker` (chosen)** vs. a check at each of the five call sites. Gating the single launch choke point means a consumer added later that forgets to branch degrades to a no-op instead of launching `wscript` — fail-safe rather than fail-open. It is also the seam where a different out-of-process backend would attach without touching any call site.
- **Uninstall cleanup via `MoveFileEx` + `MOVEFILE_DELAY_UNTIL_REBOOT`** — this was the original plan's recommendation and was **rejected on review**: the flag writes `PendingFileRenameOperations` under HKLM and requires administrator rights. The add-in installs to `%AppData%` specifically so it never needs elevation, and a user locked down enough to have scripts blocked is the least likely to be an administrator. **Leave-and-notify** was chosen instead: name the files, open the folder, and quit Access so the user can delete them.
- **VBA project save: block the export** vs. **warn and continue (chosen)** vs. **stay silent (rejected outright)**. Per the 2026-07-29 entry, the worker is the only mechanism that reliably saves the project, and three of the four alternatives tried fail *silently*. Staying silent would mean an export omitting unsaved form and report class-module edits while reporting success — strictly worse than the antivirus alerts being fixed. Blocking was rejected as too aggressive when the user may not have unsaved work that matters to them. `SaveCurrentVBProject` already returns the project's real `Saved` state, so the failure is detectable; the export path was simply discarding it because `SaveUnsavedVbaProject` was a `Sub`.
- **Access-database worker as the fallback backend** — documented as a future alternative, not built. The worker logic is already VBA-compatible and the add-in self-routes `/cmd` actions via `AutoRun`, so copying the add-in to a temp file and launching `MSACCESS.EXE <copy> /cmd <ACTION>` would reuse all existing code and could restore the VBA project save. Deferred because it costs seconds of Access startup per call, adds a temp database to manage, and — decisively — it is **unverified** whether launching `MSACCESS.EXE` avoids the same heuristic, which cannot be tested without a reporter who can whitelist. If a dedicated worker database is ever wanted, ship it prebuilt like `Template/`: injecting code via `VBComponents.Import` requires "Trust access to the VBA project object model", commonly disabled by GPO in exactly these environments.

**Decision**: Add `blnUseWorkerScript` to `udtInstallSettings` (registry `Install\Use Worker Script`, default **on**), surfaced as **Use helper script (Worker.vbs)** in the installer's Advanced Options and read through `modInstall.UseWorkerScript`. `CallWorker` early-outs when it is off. Each consumer gets a script-free fallback: the accessibility probe reports "not accessible" through a new `modBuild.DatabaseAccessibleToOtherClients` so all three of its call sites take their existing reopen paths deliberately rather than by way of an implicit `CBool(Empty)`; uninstall lists the leftover files and quits; self-rebuild declines with a pointer to Build From Source; and the VBA project save warns. Turning the setting off also deletes `Worker.vbs`, because for these users the file's presence is part of the complaint, not just its execution.

**What this rules out**: Any future worker consumer must either tolerate a no-op or check the setting itself — `CallWorker` will not fail loudly. Locked-down users permanently forgo the in-place merge optimization and `RebuildAddIn`, and accept a manual VBE save before exporting a dirty project. The registry scope means the setting cannot vary per project, which is deliberate. Revisit the Access-database worker if the warn-and-continue save proves annoying in practice, or if anyone can confirm that `MSACCESS.EXE` escapes the heuristic.

**Relevant files**: `modInstall.bas` (setting, accessor, `NotifyManualAddInCleanup`), `clsWorker.cls` (`CallWorker` gate, `RemoveWorkerScript`), `modDatabase.bas` (`SaveUnsavedVbaProject` now a `Function`, `WarnUnsavedVbaProject`), `modBuild.bas` (`DatabaseAccessibleToOtherClients`), `clsVersionControl.cls` (`RebuildAddIn`), `frmVCSInstall` (`chkUseWorkerScript`), `docs/architecture.md`, `Wiki/Installation.md`, `Wiki/FAQs.md`.

---

## 2026-08-06 — In-memory error-break suppression for MCP/API calls

**Trigger**: MCP tools (`vcs_run_vba`, `vcs_export_database`, etc.) route through `modAPI.API` / `APIAsync`. When **Break on Error** is enabled, `LogUnhandledErrors` executes `Stop` on leftover `Err` before most `On Error` directives. That halts Access until a human continues — the MCP server sees a hung call with no JSON response.

**Options explored**:
- **Temporarily set `Options.BreakOnError = False`** (same pattern as `clsTestRunner`) — rejected: export calls `Options.SaveOptionsForProject`, which would persist `"BreakOnError": false` into `vcs-options.json`. The `Options` singleton can also be replaced mid-call via `LoadOptionOverrides`, dropping the suppression.
- **Switch VBE error trapping to Break on Unhandled Errors during MCP calls** — rejected: when a break does occur, Break in Class Modules stops at the line that raised the error; Break on Unhandled Errors surfaces it in the parent handler, which is less useful for diagnosis. The scope cannot prevent VBA-native unhandled breaks anyway.
- **In-memory nesting counter in `modErrorHandling`, honored by `LogUnhandledErrors` and `DebugMode` (chosen)** — `SuppressErrorBreaks` / `RestoreErrorBreaks` at MCP/API entry points. Cannot be serialized, survives options reload, nests correctly for `APIAsync` → `API` → `RunVBA`.

**Decision**: Add `SuppressErrorBreaks`, `RestoreErrorBreaks`, and `ErrorBreaksSuppressed` to `modErrorHandling`. Push at the first statement of `API`, `APIAsync`, `HandleAPIAsyncOperation`, and `RunVBA`; pop on every exit path including `ErrHandler`. When suppressed, log unhandled errors instead of `Stop`, and have `DebugMode` return `False`. Replace `clsTestRunner`'s `Options.BreakOnError` toggle with the same scope (avoids disk write-through during round-trip exports). Leave `Operation.Begin`'s Break in Class Modules trapping unchanged.

**What this rules out**: Mutating `Options.BreakOnError` for automation scopes. Changing VBE trapping mode during MCP calls to reduce breaks at the cost of break location. If a pop is skipped by a hard crash, breaks stay suppressed until Access restarts — acceptable degraded state.

**Relevant files**:
- `Version Control.accda.src/modules/Infrastructure/modErrorHandling.bas` — suppression primitive
- `Version Control.accda.src/modules/API/modAPI.bas` — `API`, `APIAsync` scopes
- `Version Control.accda.src/modules/Utility/modTimer.bas` — async timer scope
- `Version Control.accda.src/modules/API/clsVersionControl.cls` — `RunVBA` scope
- `Version Control.accda.src/modules/Tests/clsTestRunner.cls` — uses suppression instead of option mutation

---

## 2026-07-31 — One declarative list of export formats, guarded by a test that parses the enum

**Trigger**: Adding an export format version took three coordinated edits, and one of them failed silently. `frmVCSOptionsExport.Form_Load` populated its combo by looping `For lngFormat = EFV_4_1_2 To eExportFormatVersion.[_Last]` — 10,000 iterations across the sparse packed-integer space — and filtering through a hand-written `Case EFV_4_1_2, EFV_5_0_0, EFV_5_1_0`. Forget that `Case` and the new format still gates correctly everywhere in code; it just never appears in the UI, so no user can select it. `[_Last] = 50100` was a second hand-copied duplicate of the newest value.

**Options explored**:
- **Runtime reflection over the add-in's own source** — parse the enum from `GetCodeVBProject.VBComponents("modConstants").CodeModule` when the form loads. Truly one edit per version, and the precedent exists (`clsWorker.GetWorkerScriptContent` already reads its own code module at runtime). Rejected: the installer offers a compiled `.accde` install (`Install\Compile accde`), where `VBComponents` is unavailable, and "Trust access to the VBA project object model" is a Trust Center setting the add-in never manages. Both failure modes are silent and would need a fallback list anyway — the very thing being eliminated.
- **Deploy-time codegen** — regenerate the list between `BEGIN`/`END` markers, as `clsTestRunner.SyncFactoryEntries` does for `modTestAssert`. One edit per version and drift is impossible. Rejected as disproportionate: export formats appear roughly twice per major version, and the machinery would have to be understood by anyone touching `modConstants`.
- **Dictionary of version to description**, surfacing the enum's trailing comments in the combo. Rejected with the UI change: nothing consumes the descriptions, and a value slot nobody reads becomes a second thing to keep current.

**Decision**: `GetExportFormatVersions()` in `modConstants.bas` returns a `Collection` of packed values in ascending order, mirroring the `GetExporterRevisions()` shape already in that module. `LatestExportFormat()` returns its last entry, replacing the `LATEST_EXPORT_FORMAT` const, and `[_Last]` is deleted. The combo iterates the collection. Adding a version is now the enum member plus one `col.Add` line, adjacent in the same block.

VBA cannot enumerate enum members at runtime, so the list still repeats the enum. `modTestExportFormat` closes that gap where the cost is acceptable: it parses the `eExportFormatVersion` block out of the add-in's own source and asserts the enum and the list agree in both directions, that the list is strictly ascending (so `LatestExportFormat` is meaningful), and that each `EFV_a_b_c` name matches its packed value. Reflection in a test is safe in a way it is not at runtime — the test module only exists in the add-in's own project, so it only ever runs in a development session where the source is present and the VBA object model is reachable. When the code module cannot be read it fails with an actionable message rather than passing without checking anything.

Converting `LATEST_EXPORT_FORMAT` from a `Const` to a function was safe because all eight call sites are runtime expressions — no `Case` labels, no array bounds.

**What this rules out**: The enum-to-list duplication is now guarded rather than removed, so the guarantee is only as good as the test being run. If export formats ever become frequent enough that this friction bites, deploy-time codegen is the next step and the guard test becomes its verification. The parser assumes the current declaration shape (`EFV_x_y_z = nnnnn`, one per line, optional trailing comment); a reformatted enum silently parses fewer members, which the count assertion catches in one direction but not if the whole block stops matching — the "parsed at least one member" assertion covers that case.

**Relevant files**:
- `Version Control.accda.src/modules/Infrastructure/modConstants.bas` — `GetExportFormatVersions`, `LatestExportFormat`; `[_Last]` and `LATEST_EXPORT_FORMAT` removed
- `Version Control.accda.src/forms/frmVCSOptionsExport.cls` — combo populated from the list
- `Version Control.accda.src/modules/Tests/Infrastructure/modTestExportFormat.bas` — guard test

---

## 2026-07-31 — Compile gate in the rebuild worker, not in Build or in the installer

**Trigger**: `Rebuild Add-In` decides the rebuild succeeded by reading the build log for `Done. (` and the absence of `CRITICAL:`. That only proves source files imported. Nothing anywhere in the pipeline compiles the VBA, so an agent-introduced syntax error installs cleanly and only surfaces the next time someone opens the add-in.

**Options explored**:
- **Compile at the end of `modBuild.Build`** — reaches every build, but the 2026-07-29 entry already settled that a project which does not compile still has to be buildable and mergeable. Rejected: it would turn a broken project into an unrecoverable one.
- **Compile in `AfterBuild`** (`Options.RunAfterBuild`, add-in project only) — runs inside the newly built database, so it is naturally scoped to the add-in's own build. Rejected: `AfterBuild` executes *in* the project it would be compiling, and compiling a project with running code is exactly what raises the modal reset prompt.
- **Gate the `/cmd INSTALL` branch of `modInstall.AutoRun`** — would cover manual installs too. Rejected: an end user installing a downloaded `.accda` can fail to compile for environmental reasons (a missing reference on their machine), and blocking a legitimate install is worse than the problem being solved.
- **Open a clean Access instance to compile the built file** — a faithful simulation of what the user gets, with no leftover run-state. Rejected: `OpenCurrentDatabase` runs the target's AutoExec (the existing code already closes `frmVCSInstall` for that reason), and AutoExec on non-compiling code can raise a modal dialog with nothing present to dismiss it, hanging the disconnected worker.

**Decision**: Gate in `clsWorker.BuildAndInstall`, between the log-success check and the `/cmd INSTALL` launch, reusing the instance that performed the build. `CompileBuiltAddIn` activates the rebuilt project (the build closed and recreated the database, so the project `Main` activated no longer exists), clears run-state with the VBE Reset control, then issues `acCmdCompileAndSaveAllModules` (126) with a fall back to `acCmdCompileAllModules` (125) so a refused save is not read as a compile failure, and returns `Application.IsCompiled`. Compiling *and saving* means the installed copy ships compiled rather than compiling on the user's first load. On failure the instance is left open with the VBE showing, so the developer lands on the error instead of reopening the file to find it.

Both commands are issued from out of process, which is where VBE and project-save operations have proven reliable here (see `SaveVbaProject` in the same file and `modVbeUtility.SaveCurrentVBProject`). The reset is safe from this position for the same reason the 2026-07-29 crashes were not: no VBA is running in the target instance, so the reset cannot end its own caller.

**What this rules out**: The gate does not protect anything but the developer rebuild loop — a hand-built `.accda`, a downloaded release, or a `/cmd INSTALL` run by any other route still installs unchecked. Revisit if uncompilable add-ins start reaching users through those paths. Note also that `VerifyWorker` regenerates `Worker.vbs` from the *running* add-in's `clsWorker` code while `strInstalledLib` points at the *installed* add-in, so the gate only becomes live on the second rebuild after this change is installed; a first rebuild that installs a broken build is the bootstrap, not a failure of the gate.

**Relevant files**:
- `Version Control.accda.src/modules/Integration/clsWorker.cls` — `CompileBuiltAddIn` and `QuitAccessInstance` added; `BuildAndInstall` reordered so the build instance outlives the log checks
- `Wiki/Editing-and-Contributing.md` — development workflow no longer asks for a manual `Debug > Compile`

---

## 2026-07-31 — Property application order is not load-bearing; keep the sort alphabetical

**Trigger**: The canonical sort in the entry below orders property nodes alphabetically. That is the same ordering issue #691 blames for losing combo-box lookup properties on rebuild: alphabetically `BoundColumn`, `ColumnCount`, `ColumnHeads` and `ColumnWidths` all precede `DisplayControl`, `RowSourceType` and `RowSource`, and the theory was that Access re-derives lookup metadata when those three are set, resetting `ColumnCount` to 1 and discarding `ColumnWidths`. Since the sort sits behind an export format gate, it was worth proving before shipping.

**The corpus says the order looks load-bearing.** Across 444 tbldefs files from four databases — SecTbl (324 local tables), sec, Testing, and Northwind_dev, the last being Microsoft-authored and so the one sample with a creation history independent of this toolchain — Access's emission order is unstable in general (96 distinct field sequences, most property pairs appearing in both orders) but the lookup chain never inverts once: `DisplayControl` before `RowSourceType` (461 to 0), before `RowSource` (445 to 0), before `BoundColumn` (445 to 0), before `ColumnCount` (461 to 0), before `ColumnHeads` (461 to 0), before `ColumnWidths` (410 to 0). `SubdatasheetName` likewise always precedes `LinkChildFields` and `LinkMasterFields`, and `OrderByOn` always precedes `OrderBy`.

**Measurement says it is not.** Each scenario built the same combo lookup twice, once with properties in alphabetical order and once in Access's order, then read all ten lookup properties back:

| Path | Alphabetical result |
|---|---|
| `Application.ImportXML` | all ten correct |
| DAO local table, sequential `SetDAOProperty`-style apply | all ten correct |
| DAO linked table (Access backend) | all ten correct |
| Local and linked with a real `Table/Query` row source rather than a value list | all ten correct |

`ColumnCount` stayed 2 and `ColumnWidths` stayed `0;1440` everywhere. The invariance in the corpus reflects how Access serializes, not a constraint it enforces on read.

**Decision**: Keep the sort plain alphabetical. Do not add a dependency-ordered rank table, and do not reorder properties at apply time in `modTableDefBuilder` or in the linked-table restore. A rank table would mean a new helper, three call sites, and a procedure header asserting a dependency that does not exist.

**What this rules out**: Reading the corpus invariance as evidence of a dependency. It is real and reproducible, and it is still the wrong conclusion — the measurement above is the one that counts. `TestPropertyNodesSortedUnderNewFormat` now asserts the inversion explicitly (`BoundColumn` ahead of `DisplayControl`) so that a future reader who rediscovers the corpus pattern cannot quietly "fix" it.

**Issue #691 is therefore still unexplained.** The ordering hypothesis in that thread was never verified and now looks wrong, so nothing here fixes it. The reported loss of `ColumnCount` and `ColumnWidths` on rebuild has some other cause; candidates not ruled out are `SetDAOProperty`'s delete-and-recreate on type mismatch, behaviour specific to an ODBC/SQL Server link (only an Access backend was reachable for testing), or something in how the link is recreated during a rebuild.

**Also in this change.** The sort is now gated to `this.intObjectType = edbTableDef` and does a single `xsd:appinfo` scan handling both property kinds, rather than two full-document scans. `SanitizeXML` is shared with `clsDbTableData`, so the previous form charged every table-data export — which can run to many megabytes — two descendant scans for nodes that are almost never present. `TestPropertyNodesNotSortedForTableData` covers the gate. Two fixtures were added to `Testing/Fixtures/tabledefs/`, since the corpus had no lookup or subdatasheet coverage at all: `tblFixLookupCombo` (the full lookup chain) and `tblFixSubdatasheet`. The latter is also the only place `SubdatasheetHeight` and `SubdatasheetExpanded` appear anywhere in the corpus.

---

## 2026-07-30 — Canonicalize tbldefs property order so the DAO builder can match source

**Trigger**: The DAO table-def builder (entry below) worked, but a real merge in `sec.accdb` still fell back to `ImportXML` and paid 451.50 seconds. Verification was rejecting the table the builder produced.

**What the diff showed.** `LogDefinitionMismatch` was added to log the differing lines rather than just the verdict, and the cause was visible on the first run. Every difference is a rotation of the same property names, repeated once per field:

| Source (Access) | Rebuilt (DAO) |
|---|---|
| ColumnWidth, ColumnOrder, ColumnHidden, Required, AllowZeroLength | AllowZeroLength, Required, ColumnWidth, ColumnOrder, ColumnHidden |

Same names, same values, same count — the lines realign after each field, so nothing is added or dropped. Only sequence differs, and verification is a byte comparison.

**Why the order cannot be fixed in the builder.** Three probe tables in `Testing.accdb`, one field with the same five properties, exported with `acExportAllTableAndFieldProperties`:

| Creation path | Emitted order |
|---|---|
| `Application.ImportXML` | ColumnWidth, ColumnOrder, ColumnHidden, Required, AllowZeroLength |
| DAO, natives assigned before the save | AllowZeroLength, Required, ColumnWidth, ColumnOrder, ColumnHidden |
| DAO, natives assigned last, after the save | AllowZeroLength, Required, ColumnWidth, ColumnOrder, ColumnHidden |

`ImportXML` reproduces the document order of the file it read. For a DAO-created field, `Required` and `AllowZeroLength` are intrinsic members of the Properties collection and always lead. Assignment order is irrelevant — the second and third rows are identical. A fourth probe never assigned them at all and Access still emitted both, at default values, in the leading position. Appending them as ordinary properties to move them fails with error 3367, "an object with that name already exists in the collection", and intrinsic properties cannot be deleted. There is no lever on the DAO side.

**Decision**: Sort `od:tableProperty` and `od:fieldProperty` siblings by name in the export sanitizer (`clsSourceParser.SortXmlPropertyNodes`), gated behind `EFV_5_1_0`. Both creation paths then produce identical bytes. `od:index` nodes share the `appinfo` block and are left alone; the sorted properties are re-inserted at the position the first one occupied, so index order is untouched.

**The fast path now requires the new format.** `FastTableDefImportApplies` returns False below `EFV_5_1_0`. On an older format the DAO build would import correctly but re-export in a different order from the file it came from, rewriting `tbldefs/` on the user's next export. Trading source churn for speed is not a trade we want to make silently, so the path waits for the format the ordering fix lives in.

`EFV_5_1_0` was initially left dormant, which turned out to make the whole fast path unreachable: `[_Last]` was still 50000 and the options combo whitelisted only 4.1.2 and 5.0.0, so no project could be on 5.1.0 and `FastTableDefImportApplies` always returned False. A `tblListYN` import into `sec.accdb` on 2026-07-31 still took the full `ImportXML` cost for exactly that reason. It is now activated: `[_Last] = 50100` and the combo offers 5.1.0. The format also carries the sidecar `Info.Class` change, which goes live with it.

Note for anyone migrating a project: switching a project to 5.1.0 is not enough on its own. The DAO builder verifies by re-exporting the table it built and comparing bytes against the source file, so a project whose `tbldefs/` are still in unsorted 5.0.0 order will fail verification on every table and fall back. A full export has to run first to migrate the source, after which the fast path engages.

**Measured end to end.** Once `sec.accdb` was on 5.1.0 with its source migrated, importing `tblListYN` with a changed definition logged `Built tblListYN directly, without importing the XML` and completed in **1.14 seconds**, against **462.48 seconds** for the same import an hour earlier on the `ImportXML` fallback. The DAO build itself is 0.10s (`Create Table (DAO)`), with 0.01s each for parsing the XML and verifying the result. The sanitizer's sort costs nothing measurable: single-object exports of the same table run 1.43–1.45s on 5.1.0 against 1.29–1.52s historically, with `Sanitize XML` reporting 0.00s.

The one-time migration export is visible and worth expecting: the add-in's own export jumped from ~2.5s over 40 objects to 10.05s over 266, because changing the format invalidates the global option hash and marks every category stale. All six of its `tbldefs/` diffs verified as pure reordering — identical line multisets, resequenced.

**What this rules out**: Making verification order-insensitive instead. It is a smaller change and would work on every format, but it leaves the database and its source file genuinely disagreeing on order, so every subsequent export rewrites those files. The byte comparison is also the guard the builder's own header leans on ("even a construct we recognize but reproduce imperfectly is caught before it can be committed"), and weakening it to accommodate a difference we know is meaningless makes it weaker against differences that are not.

**Harness bug found along the way.** `RunTableDefRoundtrip` inferred "the DAO path was used" from `GetLastDeclineReason()` being empty, but that reason is set by the parser, not by the caller's verification step. A build that parsed cleanly and was then rejected and discarded left no reason behind, so the fixture passed on the `ImportXML` fallback — the exact failure the check was written to prevent, and why `tblListYN` appeared to round-trip cleanly while the same file failed in a real merge. `clsDbTableDef` now reports rejection through `modTableDefBuilder.RecordVerificationFailure`, so the reason describes the outcome of the import rather than only the parse. The fixtures also force `ExportFormatVersion = EFV_5_1_0` for the duration of the run, since the DAO path only matches source under the canonical ordering.

---

## 2026-07-30 — Build local table definitions through DAO instead of Application.ImportXML

**Trigger**: The follow-on from the entry below. Skipping the rebuild when the definition already matches source removed the cost for unchanged tables, but a table whose definition genuinely changed still paid the full 276–281 seconds of `Application.ImportXML ... acStructureOnly`. That entry closed by naming DDL generation as the next option; this is it.

**Scope correction from the build log.** The obvious worry — that a full build of 372 tables pays this 372 times — turns out to be wrong, and it changes what the fix should target. `Build_20260727_105856_191.log` (963 seconds total) attributes only **21.80 seconds to Tables**, about 0.06 per table. The reason is ordering: `GetContainers` in `modContainers.bas` creates every table (#14) before the first query (#15), so a full build creates tables into an empty query catalog, which is precisely the condition under which `ImportXML` is cheap. The real full-build cost is `modLoadFromText.LoadFromText` at 493 seconds over 4,480 calls, plus 163 in `Other Operations` — unrelated, and still open.

What *is* pathological is any import into an already-populated catalog: a merge build (which pays per table, so several tables is the worst case, not a lesser one), `ImportByType("tables")`, and single-object import.

**Decision**: A new `modTableDefBuilder.bas` parses the exported table definition XML into a schema model and creates the table with `CreateTableDef` / `CreateField` / `CreateIndex`, applying properties through the existing `SetDAOProperty`. `clsDbTableDef.IDbComponent_Import` tries it before `Application.ImportXML` and falls back on any doubt. Import-side only — no export format change, no new option, and every prior export still imports.

Three mechanisms carry the risk, because the failure mode we are guarding against is not an error but a table that is quietly missing something (exactly how `TransferDatabase` dropped `ColumnOrder`, see below):

1. **Strict parsing.** The parser walks the whole document and returns `Nothing` on any element or `od:` attribute it does not explicitly recognize, rather than building a partial table. Attachments, multi-value fields (`od:jetType="complex"`, `od:jetComplexType`) and calculated columns (`od:expression`) are refused by name.
2. **Verification.** The finished table is exported through `IDbComponent_Export` into the conflict-detection temp folder and hash-compared against the source file, reusing `StoredDefinitionMatchesSource`. `Application.ExportXML` costs 0.00 seconds even in the large database, so this is nearly free. On mismatch the table is dropped and `ImportXML` runs as before. The worst case is therefore the old timing plus a few milliseconds.
3. **Circuit breaker.** Source predating the current export format could fail verification on every table, and a merge touching many of them should not pay build-plus-export-plus-delete on top of each `ImportXML`. Three consecutive expensive failures stand the path down for the rest of the operation. A parse decline does not count — it costs about a millisecond and says nothing about the next table.

**Property names are deliberately not whitelisted**, unlike elements and attributes. Everything except five native DAO `Field` members (`Required`, `AllowZeroLength`, `DefaultValue`, `ValidationRule`, `ValidationText`) goes through `SetDAOProperty`, which is generic and works for names we have never seen. A native member we failed to route would either raise (caught) or fail verification (caught). Whitelisting them would trade that for a fallback every time Access adds a property.

**Big Integer comes out right the first time.** `ExportXML` has no schema representation for `dbBigInt` and writes a bare `xsd:decimal` restricted to `totalDigits="0"`, which is why `ImportXML` mis-creates those fields as `dbDecimal(38,0)` and `FixCorruptedBigIntFields` has to repair them afterwards. The parser recognizes the sentinel and creates `dbBigInt` directly. The repair stays for the fallback path.

**When it runs.** Two cheap gates, resolved once per operation and cached on the component instance (one instance serves one operation — `GetContainers` builds a fresh set per build, and `LoadSingleObject` is handed a fresh one per call):

- `Operation.OperationType = eotBuild` skips it. A full build is already fast, and the proven call is preferable where there is nothing to gain.
- Within `eotMerge`, saved query count must reach `MIN_QUERIES_FOR_FAST_TABLEDEF` (500). Query count is the cheapest available proxy for what actually drives the cost — total table references across the catalog. The measurements in the entry below set the scale: 4 queries costs 0.02 seconds, ~2,000 costs 2 to 19 depending on table references, ~4,000 costs 4.5 to 43, and 3,692 real ones cost 276–281. Below the line `ImportXML` is cheap; above it, every table avoided is seconds to minutes.

**Options reconsidered**: `DoCmd.TransferDatabase` out of a scratch database remains rejected for the fidelity reasons recorded below. This approach reaches the same speed without leaving the current database, and unlike `TransferDatabase` its output is checked rather than trusted.

**Testing.** Two layers, because neither covers the other:

- `Testing/Fixtures/tabledefs/` — round-trip fixtures seeded from real Access exports, proving that create-then-re-export is byte-identical. Each asserts an `import_path` check, so a fixture cannot quietly pass on the fallback and prove nothing. Fixtures under `tabledefs/fallback/` invert it and assert the builder *refuses* the construct. The eligibility gate is forced open for the run (`FastPathTestOverride`) because no test database is large enough to open it naturally.
- `modTestTableDefBuilder.bas` — parser tests from hand-written schema fragments, covering the data types no sample database happens to contain and every rejection path. Hand-written XML is unsuitable for round-trip fixtures (the drift check compares against what Access actually emits) but is exactly right for testing the type map.

**What this rules out**: Trusting a DAO-built table without re-exporting and comparing it. Attributing full-build time to table imports. Whitelisting property names, which would make the path brittle against future Access versions for no safety gain.

**First end-to-end proof.** The table that started this — `tblListYN` from `sec.accdb`, three fields, seventeen table properties, no indexes — was run through the round-trip harness against `Testing.accdb` from a fixture root outside the repo. All four checks passed in 0.298 seconds: `import_path` (the DAO builder handled it, no fallback), `xml_vs_fixture` (the table it created re-exported byte-identical to sec's own source file), and `xml_pass2_idempotent`. So for this shape the builder reproduces what `ImportXML` produces, and the verification step confirms it rather than taking it on trust.

Note the gate had to be forced open: `Testing.accdb` has 4 saved queries against a threshold of 500, which is exactly why `FastPathTestOverride` exists.

**Known open failure.** In `sec.accdb`, a single-object import of `tblListYN` with a genuinely changed definition (one `BackTint` value, 100 → 200) took the fallback: `Merge_20260730_163854_113.log` shows `Parse Table Def XML` 0.01 s and `Create Table (DAO)` 0.11 s, then `Verify Table Def (DAO)` rejecting the result and `App.ImportXML() Structure` spending 451.50 s. So the builder is installed and runs; verification refuses its output. The same source file, byte for byte, round-trips cleanly through the harness in `Testing.accdb`. The difference between the two runs is that sec's table already existed and was being replaced through `IDbComponent_Merge` (stage relations, delete, rebuild), where the harness only ever creates a fresh sandbox table — that path is currently untested.

`LogDefinitionMismatch` was added for this: a rejected rebuild now lists the differing lines in the log. Previously the only record was "the rebuilt table did not match the source file", and the temp export is swept at the end of the operation, so the evidence was gone before anyone could read it — leaving reproduction as the only option on exactly the databases where reproducing costs minutes per attempt.

**Still to verify** against `sec.accdb`, next to the numbers in the entry below: the `tblListYN` single-object import (baseline 455 s end to end, 276–281 of it in `ImportXML`); a merge build touching several tables, where the per-table cost compounds and the gate has to earn its keep; and the Tables category of a full build (baseline 21.80 s across 372 tables), confirming the `eotBuild` gate leaves it untouched.

**Drive the harness from MCP with `vcs_call_vba`, not `vcs_run_vba`.** `vcs_call_vba` is a single `Application.Run` against the add-in's API and works:

```
vcs_call_vba(<target.accdb>,
             "<AppData>\MSAccessVCS\Version Control.API",
             ["RunRoundtripTests", "C:\path\to\fixtures\"])
```

Qualify with the **full path**, which also loads the add-in on demand. The bare file name `"Version Control.API"` — the tool's own documented example — never resolves, because `Application.Run` matches on the VBA project name (`MSAccessVCS`, per `PROJECT_NAME`) rather than the file name (`Version Control`, per `ADDIN_BASENAME`). `"MSAccessVCS.API"` does work, but only after something else has loaded the add-in, which is why `RunInAddIn` calls `LoadVCSAddIn` before using that form. The full path is the only one that is correct from a cold start.

`vcs_run_vba` cannot be used for this. MCP-executed VBA is itself delivered through `modAPI.API`, so the submitted code runs *inside* an `API` call and any nested `Application.Run(... ".API", ...)` hits the `Static IsRunning` re-entrancy guard. The guard used to return `Empty` without executing, which is indistinguishable from a method that legitimately returned nothing — it read as a broken add-in rather than a refused call, and cost hours. Routing around it through `HandleRibbonCommand` is worse: it reliably kills the temp module with error 2517 and leaves the host VBA project needing a close and reopen.

`API` and `APIAsync` now return an explanatory message instead, prefixed with `API_REFUSED_PREFIX` so callers can tell a refusal from data. Note that this deliberately is *not* an `Err.Raise`: an error raised inside a library database does not propagate across `Application.Run` into the calling project's handler, so even a caller with `On Error GoTo` active gets a modal dialog that blocks Access until someone dismisses it. That was tried first and reverted.

**Relevant files**:

- `Version Control.accda.src/modules/Components/modTableDefBuilder.bas` — new
- `Version Control.accda.src/modules/Components/clsDbTableDef.cls` — `TryFastTableDefImport`, `UseFastTableDefImport`, `FastTableDefImportApplies`, `RecordFastPathFailure`, `StoredDefinitionMatchesSource` timer label
- `Version Control.accda.src/modules/Utility/modEncoding.bas` — `UnescapeXmlName`
- `Version Control.accda.src/modules/Tests/modTestRoundtrip.bas` — `RunTableDefFixtures`
- `Version Control.accda.src/modules/Tests/Components/modTestTableDefBuilder.bas` — new

---

## 2026-07-30 — Skip the table rebuild when the stored definition already matches source

**Trigger**: Reloading `tblListYN` — three fields, two rows, no relationships — from source took 455 seconds in a database holding roughly 5,000 objects (3,692 queries, 514 table definitions, 416 forms, 318 reports). The `Merge_*.log` performance report accounted for 0.76 seconds and left 454.85 in `Other Operations`, because nothing on the merge path carried a `Perf` timer. Exporting the same object was instant.

Timed each candidate call directly against the live database:

| Call | Seconds |
|---|---|
| `Application.ImportXML ... acStructureOnly` | 276.07, 281.16 on repeat |
| `Application.ExportXML` | 0.00 |
| `DoCmd.DeleteObject acTable` | 0.00 |
| `SELECT ... INTO` (same table via DDL) | 0.00 |
| `Application.ImportXML ... acAppendData` (2 rows) | 0.01 |

So it is not object creation that is slow — DDL creates the same table instantly. It is `ImportXML` with `acStructureOnly` specifically, and the cost belongs to the target database rather than to the file.

**What drives the cost.** Importing the *same file* into `Testing.accdb` (44 objects) took 0.02 seconds — roughly four orders of magnitude faster. Building up `Testing.accdb` with synthetic objects isolated which part of the catalog is responsible:

| Target database contents | `acStructureOnly` seconds |
|---|---|
| 44 objects (baseline) | 0.02 |
| + 2,000 queries reading `SELECT 1 AS X` (no table reference) | 2.00 |
| + 4,000 such queries | 4.54 |
| + 400 linked tables, 4 queries | 0.03 |
| + 400 linked tables, 2,000 queries each selecting from one of them | 19.19 |
| + 400 linked tables, 4,000 such queries | 43.13 |
| `sec.accdb` (3,692 real queries, 527 tables) | 276–281 |

Object count alone is not the driver, and linked tables are not either — 400 of them cost nothing on their own, and scanning all 440 `Connect` strings in `sec.accdb` takes 0.02 seconds. What costs is **saved queries that reference tables**: at a fixed query count, giving each query a table reference multiplied the time by roughly ten. The remaining gap to the live database is consistent with its queries being real ones carrying joins and multiple references rather than a single `SELECT *`. The working model is that adding a table invalidates the query-to-table name resolution, and `ImportXML` pays to rebuild it, at a cost proportional to the total number of table references across all saved queries. `DoCmd.DeleteObject` and `SELECT ... INTO` change the catalog too and stay free, so whatever the mechanism is, it is reached from `ImportXML` specifically.

**Options explored**:

- **Generate DDL from the table-definition XML instead of calling `ImportXML`**: rejected. The XML carries field properties, indexes, lookup metadata and `od:` annotations that the existing importer reproduces faithfully. Reimplementing that is a large change with a wide blast radius, to work around an engine cost we do not fully understand.
- **Suppress catalog churn around the call** (hide the Navigation Pane, `Application.Echo False`): not pursued. `DoCmd.DeleteObject` and DDL creation both complete in under 0.01 seconds in the same database, so the cost is inside `ImportXML`, not in Access reacting to the object list changing.
- **Import into a scratch database and copy the table across with `DoCmd.TransferDatabase`**: rejected on fidelity, despite being dramatically faster. Since the cost lives in the target catalog, importing into an empty database sidesteps it entirely: spawning a second Access instance, creating a database and importing the XML there took 2.89 seconds, and transferring the finished table into `sec.accdb` took 0.05 — about 2.9 seconds against 281, and it would work even when the definition genuinely changed. But the transferred table is not the same table. A/B-importing `Testing/Testing.accdb.src/tbldefs/tblInternal.xml` both ways and comparing every field, index and datasheet property found `TransferDatabase` silently dropping `ColumnOrder` (source says `1` for `ID`, transfer produced `0`) and, worse, setting `Required = True` on `ObjectType`, which the schema declares `minOccurs="0"`. A field that source says is optional arriving as mandatory would reject valid inserts. Worth revisiting only if the divergences turn out to be a short, enumerable list that can be replayed onto the table afterward.
- **Compare the source file against the current table and skip the rebuild when they match** (chosen): `Application.ExportXML` costs nothing even here, so the check is close to free relative to what it avoids.

**Decision**: `clsDbTableDef.IDbComponent_Merge` calls `StoredDefinitionMatchesSource` before staging relations or deleting anything. That exports the current table to the conflict-detection temp folder through the normal `IDbComponent_Export` path — so the comparison is against a file produced the same way the source file was — and compares content hashes. On a match it applies metadata and updates the index (the same tail `Import` runs, factored into `ApplyMetadataAndUpdateIndex`) and returns without touching the table.

Limited to local tables (`.xml` source). Linked tables import from `.json` without `ImportXML`, and exporting one can reach across to the back end, so the check could cost more than the import it would skip. False negatives are harmless — the caller rebuilds exactly as before — so a missing table, a source file outside the export folder, or an export that produces no file all fall through to the old path.

Dependent table data still merges afterward, because `LoadSingleObject` calls `MergeDependentObjects` separately from `Merge`. Reloading two rows into an existing table costs 0.01 seconds.

**Measured**: the `tblListYN` single-object import that started this, re-run against `sec.accdb` with the check in place, went from **455.61 seconds to 0.98** (`Merge_20260730_163347_309.log`). The log reports `Compare Table Definition` at 0.04 seconds and `Merge Table Data` at 0.04; the two largest remaining entries are `Save Index` at 0.71 and `Load Index` at 0.60, so the index round trip is now the floor for a single-object import into a database this size, not the table work. That run exercised this skip, not the DAO builder below — the definition was unchanged, so `IDbComponent_Merge` returned before reaching the import path at all.

**Consequence for table data — `MergeDependentObjects` had to switch from `Import` to `Merge`.** It called `cItem.Parent.Import`, which was correct only because the definition merge above it *always* deleted and recreated the table first: data was being loaded into a guaranteed-empty table. Skipping the rebuild breaks that assumption, and the table now arrives at the data step with its rows intact. `Import` in that state is wrong both ways it can run — the XML path uses `acAppendData`, so rows are appended alongside the ones already there (duplicate keys, or silently duplicated rows on an unkeyed table), and the tab-delimited path issues `delete from [table]` first, which fails outright against any table on the child side of a relationship. `clsDbTableData.IDbComponent_Merge` already handles exactly this — its own header says "Import cannot be reused here" — by loading a staging table and reconciling against the key, and it degrades to inserting everything when the table is empty because the definition genuinely was rebuilt. The call now goes there.

**Also fixed**, since both were what made this hard to diagnose:

- The merge path now carries `Perf` timers (`App.ImportXML() Structure`, `App.ImportXML() Data`, `Import Table Data (TDF)`, `Stage Relations`, `Delete Table Object`, `Restore Relations`, `Compare Table Definition`), so a slow merge names its own bottleneck instead of reporting `Other Operations`.
- `LoadSingleObject` saved the index *after* `Perf.EndTiming`, so the existing `Save Index` timer never reached the report. The save now happens before timing stops. `ClearTempExportFolder` also runs unconditionally, because the new comparison writes there even when the index is disabled on the MCP/API path.

**What this rules out**: Treating a table merge as unconditionally destructive. Attributing the cost to object count, to linked tables, or to anything about the table being imported. Swapping `ImportXML` for `TransferDatabase` out of a scratch database without first proving property-level fidelity. Any future change that makes `IDbComponent_Export` expensive for local tables would undermine this check and should revisit it.

Does not address the underlying `ImportXML` cost — a table whose definition genuinely changed still pays it, and a full build of many tables into a large database still pays it per table. That is the case to watch: a full build is exactly where the definition always differs, so the skip never fires. If it becomes the complaint, generating DDL from the XML is the next option, and the measurements above are the baseline to beat.

**Relevant files**:

- `Version Control.accda.src/modules/Components/clsDbTableDef.cls` — `IDbComponent_Merge`, new `StoredDefinitionMatchesSource`, new `ApplyMetadataAndUpdateIndex`
- `Version Control.accda.src/modules/Components/clsDbTableData.cls` — `IDbComponent_Import` timers
- `Version Control.accda.src/modules/Core/modBuild.bas` — `LoadSingleObject` cleanup ordering, `MergeDependentObjects` now merges table data instead of importing it

---

## 2026-07-30 — Split the table data reconcile update when the engine refuses it

**Trigger**: A merge build reported `Error 3360: Query is too complex` for `tblWorkforce` (67 fields) and rolled the whole table back, leaving it unmerged. The reconcile builds one `UPDATE` carrying an assignment per non-key field plus a two-term comparison per non-key field, which for that table is 66 assignments and 132 `OR`-ed comparisons. Reproduced read-only against the live database: the comparison chain alone fails on its own in a `SELECT`, and the same shape truncated to 20 fields runs. The table holds no rows, so this is the engine declining to compile the statement, not anything to do with data volume. The merge log's `Perf` table showed the same thing — `Reconcile: Insert` ran 17 times against `Reconcile: Update` 16.

**Options explored**:

- **Gate on a field count**: rejected. There is no published ceiling, and there cannot be a reliable one: the limit is on the size of the parsed expression tree, so joins, calculated columns, and function arguments all spend from the same budget. Reported thresholds in the wild sit anywhere from the high 20s to past 100 depending on what else the query carries. Any constant we picked would be wrong for some table in both directions.
- **Drop the comparison and update every matched row**: rejected. It halves the expression count but not reliably enough to matter, rewrites rows that never changed, and turns the "N changed" count into "N matched".
- **Row-by-row DAO compare and edit**: rejected for the same reason the 2026-07-28 entry rejected it for the whole reconcile — the per-row cost is what the set-based design exists to avoid, and wide tables are not necessarily small ones.
- **Try the one statement, split into field groups only when the engine refuses** (chosen): keeps today's single statement, and its timing, for every table that compiles.

**Decision**: `ReconcileTableData` delegates the update to `UpdateChangedRows`, which attempts the all-fields statement and inspects the error. Only 3360 is recoverable; anything else is re-raised so the caller's handler rolls the transaction back. Retrying is safe because the refusal happens while compiling, before any row is touched, so the surrounding transaction is intact.

The fallback assigns and compares `UPDATE_FIELD_GROUP_SIZE` (12) fields per statement, keeping the two-table shape that already works rather than introducing a join the engine might refuse to update through. Because no group size is *known* to be safe, a further 3360 halves the size and repeats the pass; repeating is harmless, since groups already applied no longer differ from the staging table and match nothing on the way through again.

**Changed rows are counted by collecting keys, not by summing statements.** Per-statement `RecordsAffected` would count a row differing in two groups twice. Each group reads the keys it is about to change into a `Dictionary` immediately before running its update — before, because the rows stop differing once it runs — and the count is the number of distinct keys. A temporary keys table was the obvious alternative and was rejected: it would have to be created inside the open transaction, and Jet does not treat DDL as transactable, so the cleanup story is worse than holding the keys in memory. Memory is bounded by the number of *changed* rows, not table size.

`Reconcile: Update (Grouped)` appears in the `Perf` table only when the split path ran, so a log tells you which path a table took. The caller stops holding a timer across the update, since the helper now owns and closes its own.

**What this rules out**: Treating a field count as a proxy for what the engine will compile, anywhere in this path. Reporting the changed-row count from `RecordsAffected` once more than one statement is involved. Creating temporary tables inside the reconcile transaction. If a table ever proves too wide even at one field per statement, the failure surfaces as the original error rather than silently doing nothing — that would be the point to revisit, most likely with a DAO row walk reserved for that case.

**Relevant files**:

- `Version Control.accda.src/modules/Components/clsDbTableData.cls` — `ReconcileTableData`, new `UpdateChangedRows`, `UpdateChangedRowsInGroups`, `UpdateFieldGroups`, `BuildFieldAssignments`, `BuildFieldComparisons`, `BuildKeyColumns`, `GetRowKey`
- `Version Control.accda.src/modules/Tests/Components/modTestTableData.bas` — `TestTableDataMerge_WideTableUpdatesInGroups` and its 80-field table builders

---

## 2026-07-30 — Repair dbBigInt fields after Application.ImportXML

**Trigger**: Issue #734 / PR #735. `Application.ExportXML`/`ImportXML` has no schema representation for `dbBigInt` (Large Number) fields. On Full Build or Merge, `ImportXML` re-creates them as `dbDecimal(38,0)`; once corrupted, writes succeed but always store `NULL`. A separate cosmetic bug mislabeled `dbBigInt` as `COUNTER` in optional `.sql` companion files because `Case dbAutoIncrField` (value 16) collided with `dbBigInt` (also 16) in `GetTypeString()`.

**Options explored**:
- **Detect via `od:jetType` / `od:sqlSType`**: rejected after live probe. Access exports `dbBigInt` with no `od:jetType` on the element and an `xsd:decimal` restriction carrying `totalDigits value="0"`. There is no `bigint` jetType to key on. `longinteger`, `text`, and other common types all carry explicit jetType values.
- **Post-import DAO `Field.Type` reassignment**: rejected. DAO raises "Illegal operation" when changing `Type` on an appended field.
- **Delete and recreate the field**: rejected. Loses field ordinal position and complicates indexes/relationships.
- **Scan source XML for `totalDigits value="0"` on `xsd:decimal`, then `ALTER TABLE ... ALTER COLUMN ... BIGINT` after `ImportXML`**: chosen. Verified live: repair succeeds, field order is preserved, primary keys on bigint columns survive, and `od:fieldProperty` values such as Caption survive the ALTER.
- **Full DOM parse on every table `.xml` during build**: rejected for the common case. An `InStr` pre-check for `totalDigits value="0"` skips DOM construction when the signature is absent; DOM/XPath runs only on the rare matching files.

**Decision**: Before `Application.ImportXML`, scan the table-definition XML (from disk, before any temp copy is deleted) with `GetBigIntRepairFieldNamesFromTableDefXml` in `modDatabase.bas` using namespace-agnostic XPath (same pattern as `clsSourceParser.cls`). After a successful import, run `FixCorruptedBigIntFields` in `clsDbTableDef.cls`, which issues `ALTER COLUMN ... BIGINT` per affected field and verifies the result through a fresh `CurrentDb` handle (a `SharedDb` handle held since before the table rebuild can report the pre-ALTER type even after `TableDefs.Refresh`). Also fix `GetTypeString()` / `SaveTableSqlDef` to emit `BIGINT` and `DECIMAL(p,s)` correctly.

**Residual risk accepted**: Legacy `dbNumeric` fields export with the same `totalDigits value="0"` signature as `dbBigInt` (no `od:jetType` on either). A table using `dbNumeric` could be mis-identified and altered to `BIGINT`. `dbNumeric` is uncommon relative to Large Number fields, and genuine `dbDecimal` fields export with their real precision (e.g. `totalDigits="38"`), so they do not match.

**What this rules out**: Fixing this at export time only (ImportXML would still corrupt on the next build). Relying on Access to add `od:jetType="bigint"` in a future release without revisiting the XPath. Using `SharedDb` for the post-ALTER type verification.

**Relevant files**: `Version Control.accda.src/modules/Components/clsDbTableDef.cls`, `Version Control.accda.src/modules/Utility/modDatabase.bas`, `Version Control.accda.src/modules/Tests/Components/modTestTableDef.bas`, `Testing/Testing.accdb.src/tbldefs/tblBigInt.xml`.

---

## 2026-07-30 — UTF-8 conversion streams are cached, not rebuilt per hash

**Trigger**: After the 2026-07-29 scan work took a zero-change merge from 50.9s to 7.35s, the remaining per-file cost was uneven across categories: the benchmark showed Queries at ~1.09 ms per file against 0.38–0.39 ms for Forms and Modules. The initial hypothesis — that un-memoized `clsDbQuery.SourceFile` was driving it through repeated `m_Query.Name` COM reads — turned out to be wrong, and measuring the alternatives against a live 3,681-query database redirected the work.

**Options explored**:
- **Read query names from `MSysObjects` instead of `CurrentData.AllQueries`**: rejected on measurement. Enumerating all 3,681 queries and reading `.Name` costs 10 ms via the COM collection and 10 ms via a `MSysObjects` snapshot — no difference. DAO `CurrentDb.QueryDefs` is worse on both counts: 44 ms, and it returns 4,918 objects because it includes the 1,237 `~sq_`-prefixed hidden system queries that `AllQueries` correctly omits. The COM collection was never the bottleneck; a repeated `AccessObject.Name` read is 0.12 µs.
- **Replace `Application.GetHiddenAttribute` with a pre-built `MSysObjects.Flags` map** (the remaining per-object call in `GetQueryMetadataHash`, since Descriptions were already batch-loaded from `LvProp` on 2026-07-13): rejected as a measured regression. `GetHiddenAttribute` costs 1.06 µs per call, 4 ms for the whole set — *faster* than a `Scripting.Dictionary` lookup at 2.12 µs. The two sources were verified to agree (zero mismatches across 3,681 queries), but only the non-hidden case was exercised, so the equivalence is unproven for hidden objects and there is now no reason to rely on it.
- **Replace `ADODB.Stream` UTF-8 conversion with `WideCharToMultiByte`**: deferred. Almost certainly faster still (sub-microsecond), but it is a different encoder, and every hash stored in `vcs-index.idx` depends on byte-identical output. The surrogate and lone-surrogate edge cases would need to match exactly; the stream-reuse option gets most of the win with the same encoder and therefore zero encoding risk.
- **Cache the two `ADODB.Stream` objects and rewind them between calls**: chosen. See below.

**Decision**: `GetUTF8Bytes` constructed **two fresh `ADODB.Stream` COM objects on every call** — opening each, setting `Charset` and `Type`, then `WriteText`/`CopyTo`/`Read`. Since `GetStringHash` routes through it, every property hash, content hash, and metadata hash in the product paid that construction cost. Measured at 39.1 µs per call; reusing a module-level pair, rewound with `Position = 0` followed by `SetEOS`, costs 3.9 µs. The pair is opened on first use by `EnsureUtf8Streams` and torn down by `ReleaseUtf8Streams`, called from `modObjects.ReleaseObjects` and from the `GetUTF8Bytes` error handler so a transient stream failure cannot poison every later call.

The `SetEOS` truncation is the load-bearing detail: without it a short input inherits the tail of a longer preceding one. Byte-for-byte equality with the fresh-stream form was verified across both BOM modes for ASCII, accented, CJK, surrogate pairs, lone high and low surrogates, empty strings, a 5,000-character string, and — the dangerous direction — short inputs immediately following long ones.

Separately and much smaller, `clsDbQuery.SourceFile` is now memoized to match `clsDbForm`, cleared whenever `m_Query` is rebound (the `DbObject` setter plus the two import paths that assign `m_Query` directly). An uncached read costs 5.9 µs because it rebuilds the path through an un-memoized `Options.GetExportFolder` — including a `CurrentProject.FullName` COM read — plus the ten `Replace` calls in `GetSafeFileName`; a memoized read is 0.04 µs. Worth ~22 ms per pass over 3,681 queries: real, but not the seconds the category benchmark shows.

**Residual risk accepted**: the cached streams are process-wide module state, so any future reentrant use of `GetUTF8Bytes` (a hash computed from inside a hash) would corrupt both results. Nothing in the current call graph does this, and VBA is single-threaded. `Options.GetExportFolder` remains un-memoized and is still called once per component per category via `BaseFolder`; memoizing it needs care because it depends on `CurrentProject`, which changes when the database being operated on changes.

**Measured**: zero-change merge of the same project, against the 2026-07-29 result. Two consecutive runs are shown because the first had an outlier folder walk; the folder scan is the noisiest line in this report and must be checked before reading a total.

| Operation | 2026-07-29 | Run 1 | Run 2 |
|---|---|---|---|
| **Total runtime** | **7.35s** | **7.23s** | **6.68s** |
| `Get File Property Hash` (exclusive) | 0.75s / 5,094 | 0.44s / 5,094 | 0.44s / 5,094 |
| `Compute SHA256` | 0.95s / 5,111 | 0.87s / 5,111 | 0.85s / 5,111 |
| `Scan Source Files` (loop, exclusive) | 2.08s | 2.02s | 1.94s |
| `Scan Folder Metadata (API)` | 0.30s / 1 | 0.54s / 1 | 0.27s / 1 |
| `Other Operations` | — | 1.49s | 1.38s |

Reading run 2, where the folder walk is back in line with the baseline, the net is **~0.67s (9%)**. The largest single component is exactly the intended one and is reproducible across both runs: `Get File Property Hash` — the operation wrapping `GetStringHash`, since the `Compute SHA256` timer covers only `HashBytes` and never included UTF-8 conversion — fell 41%, from 0.75s to 0.44s, about 61 µs per call. `Scan Source Files` fell 0.14s, consistent with the `clsDbQuery.SourceFile` memoization being worth ~22 ms per pass plus incidental gains.

The 0.31s hash-path saving is precisely the ceiling that a call count predicted in advance: ~5,100 string hashes at ~60 µs each. That ceiling is low because the 2026-07-29 work had already cut hashing from 20,095 calls to 5,111, so this follow-up could never have been a large win no matter how fast the primitive became. The lesson worth carrying forward is to multiply per-call cost by call count *before* choosing the work, and to take two runs before quoting a total — the first run here understated the result by a factor of five purely through folder-walk variance.

Note also that `GetFileHash` goes through `GetFileBytes`, not `GetUTF8Bytes`, so file content hashing does not benefit at all — only string hashes do (property hashes, combined content hashes, metadata hashes, dictionary hashes), and those run roughly once or twice per component regardless of operation. This is not a change that scales up on full exports.

**Where the remaining 6.68s sits**: `Scan Source Files` loop overhead 1.94s and unattributed `Other Operations` 1.38s now dominate, together half the runtime — the loop cost is `GetFileList` enumeration plus `GetAllFromDB`, which must enumerate every database object for orphan detection. Beyond that there is ~2.1s of fixed overhead that is not scanning at all (`CheckDatabaseAccessible` 0.77s, `Wait for Job Queue` 0.71s, `Load Index` 0.64s). Hashing is no longer a meaningful target: the two hash lines together are 1.29s, and the property-hash half of that is now near its floor. The deferred per-category aggregate fingerprint remains the only option identified that attacks the loop itself, and the 1.38s of untimed work has never been instrumented.

**What this rules out**: reaching for `MSysObjects` as a general substitute for Access COM object collections on performance grounds. For queries it is measurably no better, and it costs the correct hidden-object filtering that `AllQueries` gives for free. It also rules out treating the query category's per-file cost as query-specific — the dominant term is shared by every category, so component-class tuning is not where the remaining time is. And it largely rules out further hash-path work: with hashing down to 1.29s across both timers on a 6.68s run, `WideCharToMultiByte` could recover at most a couple of tenths and would require a byte-equality corpus at least as broad as the one used here to be safe. It is only worth revisiting if a future change pushes string-hash call counts back up by an order of magnitude.

**Relevant files**: `Version Control.accda.src/modules/Utility/modHash.bas`, `Version Control.accda.src/modules/Infrastructure/modObjects.bas`, `Version Control.accda.src/modules/Components/clsDbQuery.cls`, `Version Control.accda.src/modules/Tests/FileIO/modTestHash.bas`, `Version Control.accda.src/modules/Tests/Core/modTestContainers.bas`

---

## 2026-07-29 — Merge scan reads no file content when dates and sizes are unchanged

**Trigger**: A merge build on a ~5,000-component project with zero changes took 50.9 seconds, 45.3 of it in scanning. The perf report showed 20,095 SHA-256 computations, 9,893 whole-file reads, and 24 recursive folder walks — every indexed source file on disk was read and hashed even though nothing had changed. Root cause: `GetModifiedSourceFiles` computed the date+size property hash but used it only to *refresh* stale index metadata, never to skip the content read. The 2026-07-20 `AllFilesHash` work (which fixed companion-`.json` edits being dismissed) removed the property-hash-match short-circuit along with the broken primary-file-only fallback; only the broken fallback needed replacing.

**Options explored**:
- **Keep content hashing unconditional, optimize only the mechanics** (raw Win32 reads instead of ADODB.Stream, cached hash provider): rejected as the primary fix. It reduces the constant factor but still reads every byte of every source file on every merge, so the cost stays proportional to project size rather than to change count.
- **Per-category aggregate fingerprint in the index** (one hash over all files in a category, skipping the per-file loop entirely): deferred. Strictly faster still, but requires an index format addition and complicates orphan reconciliation, which needs the per-file list regardless.
- **Property-hash match short-circuits, content hash arbitrates on mismatch**: chosen. Restores the pre-2026-07-20 fast path while keeping the combined `AllFilesHash` (not the primary file alone) as the arbiter, so the bug that motivated `AllFilesHash` stays fixed. A file is clean when the date and size of every indexed file match the index; only when one moved is content read, and then the content hash decides — catching companion-only edits and dismissing timestamp-only drift such as a checkout that rewrote mtimes.

**Decision**: Three-tier precedence in `GetModifiedSourceFiles` (property hash, then combined content hash, then the legacy primary-file fallback for entries predating `AllFilesHash`), plus supporting work that stands on its own:

- **One shared folder scan.** Nine component types report the export root as their `BaseFolder`, and `ScanFolderMetadata` is recursive, so each re-walked the entire source tree. `modBuild.GetSharedScanMetadata` builds one map for the whole scan phase; when the container list contains no root-folder category (a narrowly scoped sync), only the distinct folders needed are walked, so a small operation does not pay for a full-tree walk. `GetModifiedSourceFiles` accepts it as an optional argument and falls back to its own per-category scan for other callers.
- **`GetSourceBasePath`** replaces `FSO.BuildPath(FSO.GetParentFolderName(p), FSO.GetBaseName(p))` in the four places that built a per-extension base path, and `GetSourceFilesContentHash` takes the scan map so per-extension existence is a dictionary lookup rather than `FSO.FileExists`.
- **Cached CNG provider.** `NGHash` opened an algorithm provider, queried two size properties, and closed the provider on *every* hash. At ~0.84 ms per hash for inputs that are mostly short strings, setup and teardown dominated. The handle, its size properties, and the hash object buffer are now cached per algorithm and released via `modObjects.ReleaseObjects`. The digest is formatted through a 256-entry hex lookup table instead of `Hex`/`Right`/`LCase` per byte.
- **Index item resolved once.** The loop called `Me.Item(cCategory, strFile)` up to five times per file; it now resolves one `clsVCSIndexItem` through `GetExistingIndexItem`. This also fixes a latent bug: `Exists` honors the legacy table-def `.xml`/`.json` key alias but `Me.Item` does not, so aliased entries silently got a blank item and were reported modified on every merge.

**Residual risk accepted**: a content edit that preserves both the exact byte size *and* the recorded modification timestamp is missed until the next full export. File dates are second-precision, so the window is an edit landing in the same second as the recorded date at identical length. This is the behavior the add-in had before 2026-07-20 and the same assumption git and rsync make. `VCS.FullExport` remains the escape hatch.

**Measured**: Zero-change merge of the same ~5,000-component project (10,047 indexed files), before and after. Perf timers are exclusive, so nested time is attributed to the innermost operation.

| Operation | Before | After |
|---|---|---|
| **Total runtime** | **50.92s** | **7.35s** |
| `Scan Source Files` (loop, exclusive) | 6.29s | 2.08s |
| `Compute SHA256` | 16.92s / 20,095 calls | 0.95s / 5,111 calls |
| `Get File Content Hash` | 7.60s / 5,094 calls | *absent — 0 calls* |
| `Read File Bytes` | 8.01s / 9,893 calls | 0.00s / 3 calls |
| `Scan Folder Metadata (API)` | 3.99s / 24 calls | 0.30s / 1 call |
| `Get File Property Hash` | 2.47s / 5,094 calls | 0.75s / 5,094 calls |

The hash count drops to exactly one property hash per component (5,094) plus 17 incidental — no source file content is read or hashed at all on a clean merge. Per-hash cost fell from ~0.84 ms to 0.186 ms, which an isolated `modTestPerf` measurement of `GetStringHash` reproduces at 0.185 ms, confirming the provider open/close was the dominant cost rather than the digest itself. The same property hash computed from the scan map instead of `FSO.GetFile` is 0.25 ms vs 0.70 ms. Threading the map into `GetSourceFilesContentHash` is worth little on its own (0.87 → 0.73 ms) because the cost there is reading bytes, not the existence check — it matters only on the mismatch path, which is now rare.

Remaining scan cost is ~4.1s, of which 2.08s is loop overhead (`GetFileList` enumeration plus `GetAllFromDB`, which must enumerate every database object for orphan detection) and 0.95s is the 5,094 property hashes. Both are what the deferred per-category aggregate fingerprint would target.

**What this rules out**: `AllFilesHash` is no longer consulted on every merge, so it cannot be relied on as a content audit of the source tree — it is a tiebreaker for files whose metadata moved. Any future change to how `FilePropertiesHash` is computed must keep the FSO branch and the Win32 scan branch byte-identical; they are written on the export path through FSO and compared on the merge path through the scan map, and a divergence would silently disable the fast path rather than fail. `modTestContainers.TestPropertyHashIdenticalWithAndWithoutMetaScan` exists to make that a visible failure. Revisit with the per-category aggregate fingerprint if scan time again becomes the dominant cost on very large projects.

**Relevant files**: `clsVCSIndex.cls`, `modBuild.bas`, `modContainers.bas`, `modHash.bas`, `modObjects.bas`, `modTestPerf.bas` (new benchmark harness), `modTestMergeDetection.bas`, `modTestContainers.bas`, `modTestHash.bas`.

---

## 2026-07-29 — Opt-in in-place merge preparation instead of the pre-merge reopen

**Trigger**: A merge build unconditionally closes and shift-reopens the database before scanning source files, costing roughly 23 seconds on a ~7,300-file project. Merges are run frequently while working with AI tools, so that fixed cost dominates the loop. Two earlier attempts to remove it were reverted: an in-place VBE reset (2026-07-06, crashed) and a deferred reopen (2026-06-09, reverted in `0e4b93b0` for stale component references).

**Probe results** (control 228 executed against a scratch database through out-of-process COM):
- The VBE `Reset` control is **id 228**, present on the Run menu, Standard, Debug, Watch, Immediate, and Locals bars. **Id 645 does not exist** in the VBE command bars, so it is not an alternative.
- The reset does clear the target project's run-state: a module-level `Long` and a global object reference went from `counter=42 obj=alive:1` to `counter=0 obj=nothing`.
- `VBE.ActiveVBProject` points at the **add-in** project whenever the add-in is loaded, not the database being merged. Anything issuing a reset must set the active project first (`ResetCurrentVBProjectState` already does). Modifying the add-in's own components by mistake hung Access on a modal.
- Importing a component while the project holds run-state raises the modal "this will reset your project" prompt, confirming that a merge into an unprepared project gets an *implicit, mid-merge* reset — the mechanism that invalidated cached references in the 2026-07-06 crash.
- The prompt still appeared after an explicit reset when the import was driven from **out-of-process COM**. Out-of-process automation is therefore not a faithful harness for the prompt behavior of the add-in's in-process import path, and the prompt question cannot be settled this way.

**Options explored**:
- **Issue the reset from the `Worker.vbs` VBScript** (it already attaches to the running instance with `GetObject` and sets `ActiveVBProject`): rejected. The probe showed out-of-process component modification behaves differently from the in-process path, so a worker-issued reset cannot be validated by the same evidence it would rely on, and it adds a round-trip plus a second failure mode for no demonstrated benefit. (This applies to the *reset* only. The VBA project *save* does go through the worker, for the opposite reason: in process it cannot be made to work at all — see the save findings below.)
- **In-place reset with no other changes** (the 2026-07-06 shape): still rejected — that is the reverted crash.
- **In-place reset, plus releasing everything the reset invalidates, plus resuming the merge on a fresh call stack**: chosen. The 2026-07-06 entry named exactly these two shapes as the prerequisites for revisiting.
- **Make it the default**: rejected. The reopen is the conservative path, and the differences (startup code no longer bypassed, run-state cleared in place) are behavioral, not just faster.

**Decision**: Add `Options.SkipReopenBeforeMerge` (default **False**). When on, the merge runs as three timer stages (see below) so that the target project's run-state is cleared deliberately, in isolation, rather than implicitly part way through the merge. Any failed step falls back to `ReopenBeforeMerge`, so a merge never proceeds in an unknown state. The resumed stage skips `Log.Clear` and `Perf.StartTiming` so all three stages share one log and one performance report.

**The reset needs isolation, not just a different stack** — learned by crashing Access on the first real merge (silent disappearance; Windows logged `MSACCESS.EXE` faulting in `VBE7.DLL`, `0xc0000005`). A VBE reset is equivalent to the `End` statement for the project it acts on, and there is no trappable error, so this has to be prevented rather than handled.

Because the crash leaves no log — Access dies without unwinding, so buffered log output is never written — diagnosing it needed `LogCrashTrace` (in `modErrorHandling`), which persists a breadcrumb to the log file at each step. That trace disproved two successive hypotheses and is worth keeping for anything that manipulates a VBA project:

1. **Self-merge was not the cause.** The first suspicion was that the reset target was the executing project. The crash happened with the add-in loaded as a library and the merge started from the ribbon, so the target project was idle — the same arrangement that has always been safe in `RunVBA`.
2. **The reset call itself was not the cause.** The trace showed it completing and returning `True`. The fault came *after* it.
3. **The reset's teardown is asynchronous.** `CommandBarControl.Execute` returns immediately; the teardown lands later, when the thread next reaches a message pump. Both crashes were in the first substantial work done after the Execute *on the same call stack* — continuing the merge in the first design, then `RestoreMainForm` in the second. Opening a form pumps messages, so it collided with the teardown. Cheap statements (recording a result, appending to the log file) proved survivable across both runs; anything that pumps did not.

So the rule is not just "reset on a stack the target project does not own" but "**do nothing on that stack after the Execute**". The reset stage therefore arms the next stage's timer *before* executing the reset, then returns to the message loop, letting the teardown land with no VBA of ours in progress.

The result is a three-stage pipeline, each stage on its own call stack: `PrepareMergeInPlace` closes objects, saves VBA, releases cached references, and stages the main form (releasing the form instance and `Log`'s console binding, exactly as a database reopen does), then arms a timer and unwinds completely; the `MergeReset` stage arms the next timer and resets the project as its final statement; `MergeResume` runs the merge with nothing surviving from either earlier stage. It does not restore the main form — the merge stage reopens it, and `frmVCSMain.ResetForOperation` rebinds the log console and clears the console text regardless, so there is nothing worth restoring first. Two further constraints:

- **The reset must run on a stack the target project does not own**, which the middle stage also satisfies.
- **There is no in-place path at all when the add-in is open as the current database**, because the project being reset is then the project running the merge, on any stack. `ResetWouldEndOurOwnCode` (in `modVbeUtility`) detects this and forces the fallback. This is the add-in's *own* development workflow, so the option cannot be exercised by self-merging this repo — it needs a separate target database.

`ReleaseScanState` is the helper `0e4b93b0` referred to: it drops component classes from a category dictionary (they cache database objects and cannot cross the boundary), releases `SharedDb`, closes cached and back-end connections, clears the `.env`/connection caches, and resets `modLoadFromText`. It takes the dictionary optionally so it also serves the deferred-reopen shape if that is revisited.

**What this rules out**: Do not collapse the three stages back into one, do not move the reset back inside `Build`, do not add work after the reset Execute on its own stack (including anything that opens a form or otherwise pumps messages), and do not remove the `ResetWouldEndOurOwnCode` guard — each crashes Access outright rather than failing safely. When debugging anything in this path, reach for `LogCrashTrace` first: a fault here produces no log at all, so reasoning without it is guesswork. Do not make this the default without validating the behavioral differences. Do not route the reset through the worker on the strength of the probe above. The option only removes the *pre*-merge reopen; the post-merge shared-mode reopen (16s measured here, worker-probed) is untouched, and skipping the pre-merge reopen makes it *more* likely to fire rather than less — see the lock-state finding under validation below.

**Validation status**: On a ~7,300-file project, once the three constraints above were in place, three consecutive merges completed without incident **in a single Access instance, with no restart between them** — which matters as much as any one run passing, since a reset that left the project subtly damaged would be expected to accumulate across repeated use in one session:

1. No changed source files (ribbon) — confirmed the stage choreography and that the log and performance report stay continuous across the two timer hops. `Prepare Merge In Place` = 0.02 s in place of the reopen.
2. Imported a standard module, a query, and two forms, then ran `InitializeForms` (ribbon) — the VBA-bearing import that crashed in 2026-07-06 and drove the 2026-06-09 revert survived a reset standing in for a reopen. 28 s total.
3. Open forms and live run-state present, invoked through the External API rather than the ribbon — `Close Open Objects` = 0.37 s and `Prepare Merge In Place` = 1.04 s (against 0.02 s when there was nothing to close), so the preparation demonstrably did real work rather than short-circuiting.

**Debug → Compile succeeded in every project afterwards.** This was the failure mode most worth ruling out: a reset leaves the project loaded rather than rebuilt, so a merge that imported modules could plausibly have left a project that no longer compiles even though the import reported success. It compiles.

Worth noting for anyone optimizing further: with the reopen gone, these runs are dominated by change detection, not by merging. Of run 2's 28 s, roughly 18 s was hashing and scanning 7,300 source files (20,136 SHA-256 computations at 6.6 s, plus file reads, folder metadata, and content hashes) against 3.9 s of actual merging.

**The post-merge shared-mode reopen is triggered by lock state, not by importing.** Run 3 imported nothing yet still reopened in shared mode (16.04 s of its 61 s), while run 2 imported four objects and did not. This corrects the model recorded in the 2026-06-09 entry, which read the reopen as a consequence of schema-modifying imports. The check is `Worker.IsDatabaseAccessible`, an out-of-process probe of the engine lock state, and it makes no reference to what the merge did.

**Consequence: the preparation falls back to a reopen when the database is not accessible.** The option's benefit and the lock are mutually exclusive — a database other clients cannot open is going to be reopened whether the preparation does it now or the post-merge check does it later, so there is no saving left to protect. Reopening up front is strictly better than reopening after:

- The pre-merge `ShiftOpenDatabase` leaves the database accessible, so the post-merge check then finds nothing to do. Run 3's shape would be roughly 33 s (reopen, fast scan, no second reopen) rather than 61 s (fast scan, then reopen).
- The merge takes its backup **mid-flight** (`FSO.CopyFile` in `Build`, `eelCritical` on failure), and an exclusive lock is documented in `IsFileOpenExclusive` as preventing exactly that copy. Run 3 had no changes and never reached that line. A locked session merging *with* changes could therefore abort the merge at the backup — a far worse outcome than a slow reopen, and the reason this fall back is a correctness measure and not only an optimization.

The check is placed at the *end* of the preparation, after objects are closed and the VBA project saved, since those steps affect the lock state themselves — as the finding below shows, one of them was creating it.

The probe costs a worker round trip (~0.95 s measured), and the in-place path would otherwise pay it twice. When the preparation confirmed accessibility and the merge then found no changed files, nothing has happened in between that could have escalated the lock, so the post-merge check is skipped (`m_blnVerifiedAccessible` + `blnNoChanges`). The no-change merge is the case the fast path exists for, so the saving lands where it matters.

An option to skip the post-merge check outright was considered and deferred rather than rejected. With the fall back in place a locked session takes the old path anyway, so the check rarely leads to a reopen; what remains is the ~1 s probe, and the baseline timing should say whether that is worth an option. The argument against is attribution: skipping it leaves a locked database behind, and the resulting failure surfaces much later in an MCP call, a worker job, or the *next* merge's backup, where nothing points back to the setting.

**The lock is escalated by the preparation itself, not inherited from the session.** The initial reading was that a used session arrives already locked and the in-place path merely inherits it. Probing on entry to the preparation as well as at the end disproved that. Two consecutive runs, minutes apart on the same database:

| Session | On entry | After preparation | Outcome |
|---|---|---|---|
| Startup code had run, forms opened | accessible | **not accessible** | fell back, `Reopen DB before Merge` 36.91 s, 101.64 s total |
| Immediately after that reopen | accessible | accessible | in place, no reopen anywhere, 56.41 s total |

Both sessions were accessible when the merge began. The first became inaccessible during preparation. Probing between the individual steps then identified the culprit exactly:

```
[trace] lock state after closing objects: accessible
[trace] prep: saving VBA project (dirty: True)
[trace] lock state after saving VBA project: NOT accessible to other clients
```

**A *partial* save of the VBA project locks the database against other clients.** `SaveUnsavedVbaProject` issued `DoCmd.Save acModule, <first standard module>` on the long-held assumption that saving one module saves the whole project. It does not, when form and report class modules are dirty — which is the usual state after startup code has run, where dozens of form classes report unsaved. Closing open objects, by contrast, is harmless.

Saving is not the problem; saving *incompletely* is. Manual verification: after `AutoExec` ran, `CurrentVBProject.Saved` was False; pressing **Save** on the VBE toolbar returned it to True, and a merge then ran fully in place with no reopen. The same sequence through `DoCmd.Save acModule` left the project dirty and the database locked.

**The same bug was already on record in two other places, mislabelled in one and invisible in the other.** `modLetterCasing.StandardizeLetterCasing` had been logging "VBA project still has unsaved changes after letter casing corrections" for a long time, directly under a comment asserting that "saving one module saves the whole project" — the warning was reporting the bug and the comment was denying it. Correcting casing in `clsStandardLetterCasing` propagates project-wide and dirties form and report classes, which the single-module save cannot reach. That also resolves an open puzzle: the same correction (`fldConvertId` → `fldConvertID`) recurred across six runs spanning three days without converging, which looked like something restoring the non-canonical casing between runs and was simply the correction never being persisted.

`modExport.ExportSource` made the same call with no warning at all, to ensure "exported source reflects the current state of the code" — which did not hold for form and report class code, so an export could silently omit unsaved class-module edits it believed it had captured. That is a correctness bug independent of merging. Both now call `SaveCurrentVBProject`; fixing them alongside the merge path was preferable to leaving two known-wrong copies of a mechanism this hard to get right for the next reader to copy.

**Decision: `SaveCurrentVBProject` owns saving the project, executes the VBE Save command (ID 3) *from the worker script*, and returns the project's actual `Saved` state rather than assuming the save worked.** `SaveUnsavedVbaProject` delegates to it, so the merge preparation, export, letter casing, and category-scoped sync share one implementation. Evidence that the worker is not incidental complexity:

```
[trace] save: VBE Save control executed (&Save sec, window visible: True), saved: False
[trace] save: retrying out of process
Worker job SaveVbaProject (1fce8f0) completed in 1.46 seconds.
[trace] save: worker returned, saved: True
```

Identical command, same instance, same project — refused in process, succeeds out of process. A hand-pressed button has no VBA frame beneath it, and that turns out to be the thing that matters. The merge that produced those lines ran fully in place after `AutoExec` had dirtied the project: 58.94 s against the 101.64 s reopen baseline, accessibility recheck passing. That is the case that motivated the whole investigation.

`clsWorker` needed very little, because two things it already does are exactly what this requires: it attaches to the *specific* instance via `GetObject(<database path>)` rather than an ambiguous `GetObject(, "Access.Application")`, and `Main` already sets `ActiveVBProject` to the current database's project for every job. (The VBA wrapper and the script function live in the same class module, hence the `Run_` prefix convention `Run_SaveVbaProject` shares with `Run_BuildAndInstall`.)

**Four mechanisms were tried and dropped. Do not reintroduce any of them without new evidence** — each looked correct, and three of the four fail *silently*, which is why this took a full day to pin down:

| Mechanism | Why it was dropped |
|---|---|
| `DoCmd.Save acModule, <one module>` | Cannot save form or report class modules at all. Reports success while leaving dirty precisely the components that matter, and locks the database. The original bug. |
| `DoCmd.RunCommand acCmdSaveAllModules` (280) | Raises 2046 ("isn't available now") unless a module window is active, and is widely reported to do nothing even when it runs. |
| VBE Save command in process | Reports success and saves nothing. No error, correct project active, caption naming the right document, before *and* after a project reset alike. |
| `acCmdCompileAndSaveAllModules` | Compiles, and a project that does not compile still has to be mergeable. |

Two theories were disproved along the way and are recorded because both are plausible enough to be re-derived:

- **Run-state is not what defeated the save.** The reading was that saving a project holding run-state requires resetting it, and that the VBE was silently declining rather than raising its "this action will reset your project" prompt to a programmatic caller. Moving the save into `FlushVbaProjectAfterReset`, after the reset has cleared run-state, produced the same silent failure — so run-state was not the cause. Run-state does still matter, just for a different reason: it makes that prompt possible where no reset has happened, which is why the worker resets first outside the merge path (below). The merge's save stays after the reset regardless, since the reset clears run-state rather than editor buffers and nothing is lost by waiting.
- **Wrong-document targeting does not explain it either.** `Execute` *is* bound to the VBE's active document rather than `ActiveVBProject`, so an add-in code pane in focus would have meant saving the already-clean add-in — a no-op indistinguishable from a refusal. Tracing `ctl.Caption` settled it: `&Save sec`, right project, right document, still nothing saved. The in-process attempt was then removed entirely rather than kept as a free fast path, because targeting it correctly requires showing a code pane, which pops the VBE window open mid-merge for a step that cannot succeed.

**Trap worth remembering: never log before clearing `Err`.** Tracing the 2046 before clearing it meant `LogCrashTrace` → `LogUnhandledErrors` reported it, putting a modal dialog in front of what is meant to be an unattended merge. It also corrupted a measurement: `Prepare Merge In Place` read 13.68 s, which looked like a slow save and was the time taken to read and dismiss that dialog. Capture the error text, clear `Err`, *then* log.

A working save also removes the reason for the pessimistic guard that briefly lived here (fall back immediately whenever `CurrentVBProject.Saved` was False). The accessibility probe at the end of the preparation remains the authority: save properly, then ask. If a save ever does leave the project dirty, a breadcrumb records it and the probe decides.

**Rejected: resetting the project before saving it outside the merge path.** Export, letter casing, and category-scoped sync call the save with no preceding reset, so the project can still hold run-state, and saving such a project should raise the VBE's modal "this action will reset your project" prompt — which nobody is present to dismiss during an unattended export. (That prompt *was* observed on importing into a project with run-state, earlier in this entry.) Resetting first, in the same worker job so that no caller VBA sits between the two steps, looked like the safe way to prevent it.

It broke export immediately, and the failure is instructive: **a VBE reset ends whatever code is running, and setting `ActiveVBProject` does not scope that away from the caller.** During an export the running code is the add-in itself, waiting in `Worker.WaitForQueue`'s `DoEvents` loop for the very job that issues the reset. So the reset terminated its own caller and took the add-in's module-level state with it:

```
ERROR: Returned worker not found in job queue: 1fce8f0   (clsWorker.ReturnWorker)
Failed to run ribbon command for btnExport
40040: The expression you entered refers to an object that is closed or doesn't exist.
```

The job queue was gone by the time the worker called back. The same run of the merge was unaffected, which isolates the cause precisely — the merge passed through the identical code with the reset suppressed, having already reset in its own stage.

This is also why the merge's three-stage choreography is not over-engineering. The merge survives a reset because its next stage arrives on a Windows timer, so nothing of ours has to live through it. A save called from the middle of `ExportSource` has no such re-entry, and giving it one would mean restructuring export around the reset — a large change to prevent a prompt that has never actually been observed on this path. Left as: save without resetting, and treat the prompt as a hypothesis awaiting a run where a dirty project is exported.

**Knock-on: this explains the most expensive reopen observed anywhere in this investigation.** `StandardizeLetterCasing` runs at the end of every export and every build/merge, and when it applies corrections it deliberately saves the project (`DoCmd.Save acModule`) so the user is not prompted at shutdown. By the measurement above, that save locks the database. The accessibility check runs a few lines later in the same procedure, so the reopen follows directly. Observed on the same project, in one merge:

```
1 letter casing correction(s) applied:
  fldConvertId -> fldConvertID
WARNING: VBA project still has unsaved changes after letter casing corrections
Reopening database in shared mode...
Reopen DB (shared mode)       1         82.96
```

The merge that produced those lines **imported nothing** — `No changes found` — so 141 seconds were spent because one identifier's capitalization was corrected in a database that was not otherwise modified.

**Decision: the casing pass runs on full builds only, not on merges.** Source consistency never depended on the merge doing it. Export standardizes casing *before* it writes source files, so the source tree is authoritative and self-healing regardless of what the database holds; a merge's incoming code comes from those already-standardized files. The worst case from skipping it is that the database carries non-canonical casing until the next export corrects it — against a reopen that has been measured at 16, 48, and 83 seconds on the same project. Export and full build are unchanged: export must run it to keep source consistent, and a full build is producing a database from scratch where the cost is proportionate and no user session is disrupted.

The cost also compounded: because the save left the project dirty, the *next* merge found a dirty project, saved it during preparation, and locked the database again. Both that and the non-convergence trace back to the partial save, so `SaveCurrentVBProject` addresses the cause; skipping the casing pass on merges makes the two paths independent regardless.

**What dirties the project in practice: `AutoExec`.** Running any VBA compiles it on demand, and the compiled state is part of the saved project, so a database whose startup macro runs code has a dirty project from the moment it finishes opening — before the user does anything. On such a database the first merge after a normal open always declines.

Even before the save worked this self-healed, because the fall back uses `ShiftOpenDatabase`, which bypasses the startup code: the reopened session has a clean project, and subsequent merges take the in-place path. Observed twice — a merge that fell back at 13:21 was followed by an in-place merge at 13:23 with no reopen at all. So even the worst case was one reopen per *run of the application*, not one per merge.

With the worker save in place that reopen is gone as well: a merge immediately after `AutoExec` saved the project and completed in place in 58.94 s. So the option pays off both in a session opened for merging and in one where startup code has run.

**Measured saving**: the two runs above are close to a matched pair — same database, no changed files, two minutes apart — and differ by 45 s (101.64 s against 56.41 s), of which 36.91 s is the reopen and the remainder is the first run's slower I/O (`Scan Folder Metadata` 8.02 s against 4.80 s). So the option removes roughly a 37 s reopen from a 100 s no-change merge on a 7,300-file project. The same pair also confirms the fall back end to end and the skipped post-merge probe: the second run made two worker calls, both during preparation, and none after the merge.

Still outstanding: the equivalence check (merge the same source into two copies of a database, one path each, export both, diff the trees), the rest of the matrix (open objects of the remaining types, VBE open with unsaved edits, themes, table data, startup form), and one run of the *export* path with run-state present, to confirm the worker's reset-then-save neither prompts nor faults. Keep the option default-off and experimental until those are done.

**Relevant files**: `Version Control.accda.src/modules/Core/modBuild.bas` (`Build`, new `ReopenBeforeMerge`, `PrepareMergeInPlace`, `ResetProjectForInPlaceMerge`, `FlushVbaProjectAfterReset`, `ReleaseScanState`, `TraceInPlaceMerge`), `modules/Core/modVbeUtility.bas` (new `ResetWouldEndOurOwnCode` and `SaveCurrentVBProject`, tracing in `ResetCurrentVBProjectState`), `modules/Integration/clsWorker.cls` (new `Run_SaveVbaProject` and the `SaveVbaProject` script action), `modules/Infrastructure/modErrorHandling.bas` (new `LogCrashTrace`), `modules/Utility/modTimer.bas` (`MergeReset`, `MergeResume`), `modules/Utility/modDatabase.bas` (new `SaveUnsavedVbaProject`), `modules/Core/modExport.bas` and `modules/Core/modLetterCasing.bas` (partial saves replaced), `modules/Infrastructure/clsOptions.cls`, `modules/Tests/Infrastructure/clsTestOptions.cls`, `forms/frmVCSOptionsBuild.form` + `.cls`, `Wiki/Options.md`.

---

## 2026-07-28 — Merge table data through a staging table and set-based reconcile

**Trigger**: Table data was skipped on merge builds, so a developer who added a record to a versioned internal table (release/version info was the motivating case) could only get it into another database with a full build. Merges are used precisely on the large databases where a full build is expensive, so "just do a full build" was not a real answer.

**Options explored**:

- **Reuse `IDbComponent_Import`**: It already loads a source file into the table. Rejected on both formats. The tab-delimited path runs `DELETE FROM [table]` first, which fails outright once any other table references this one — the normal case for a versioned lookup table — and even when it succeeds it discards AutoNumber values that child rows point at. The XML path uses `acAppendData`, which duplicates every existing row on a populated table.
- **Delete and reload inside a transaction**: Fixes the "half done" problem but not the referential-integrity failure or the AutoNumber loss, and it rewrites every row of a large table to apply a one-record change.
- **Row-by-row DAO reconcile**: Read the source file, seek each key in the table, and compare fields in VBA. Correct, but the per-row cost is exactly what makes this unusable on the large databases that motivated the feature.
- **Staging table plus set-based reconcile** (chosen): Load the source file into a temporary local table, then apply one `INSERT`, one `UPDATE`, and one `DELETE` keyed on the table's primary key. Only differing rows are written, key values and AutoNumber values survive, and referential integrity is never exercised for rows that did not change.

**Decision**: Merge reconciles table data against a temporary staging table (`vcs_tmp_merge_data*`), created with `SELECT ... INTO ... WHERE (1 = 0)` so every column keeps its exact type and an AutoNumber key is demoted to Long. A unique index on the merge key is added because the engine refuses an `UPDATE` across a join unless the joined side is provably unique. The three statements run in one `BeginTrans`/`CommitTrans`, so the expected failure — a delete blocked by a child record — rolls the table back to its prior state, logs an error, and lets the rest of the merge continue.

Behavior established by this decision:

- **Source is authoritative, including deletions.** A row absent from the source file is removed. The user opted this table into version control, so the file is the record of what the table should contain. This also matches `clsVCSIndex.IsMergeConflict`, which has always returned `ercNone` for table data rather than raising a conflict.
- **Default on** (`Options.MergeTableData`, Build options). Getting the data a developer committed is the expected outcome of a merge; the option exists for projects that treat records as environment-specific. It is in the non-export skip list in `GetCategoryHashes` — folding a build-side option into export category hashes would trigger spurious re-exports.
- **Tables without a merge key are reloaded wholesale.** `GetTableMergeKey` returns primary key or unique+required index fields and nothing else — `GetTableSortFields` falls back to "all non-binary fields" when there is no key, which is fine for ordering but would let one source row match several table rows. Keyless tables are common in practice (a production database contributed 16 of 32 exported tables with no key), and skipping them left merge unable to do the one thing it was built for. Since there is no key, there is also no identity or AutoNumber value anything could hold a reference to, so `DELETE` followed by `INSERT ... SELECT` inside the same transaction is equivalent to what a full build already does — and it preserves duplicate rows, which a key-based reconcile could not. It is refused when a relationship points at the table (`GetFirstDependentTable`), because the delete would fail and roll the whole table back; that reads `Relation.Table` as the referenced side and `Relation.ForeignTable` as the referencing one. The count line says `N row(s) reloaded (no key to compare on)` rather than added/changed/removed, because without a key those numbers cannot be established.
- **A source file with no rows skips `ImportXML` entirely.** Not an optimization. `Application.ImportXML` spends about 95 seconds on a document containing no row elements, regardless of how small the file is, while a 2,164-row 715 KB file loads in 0.32 seconds — measured across six real tables in a production database. Two empty exported tables were enough to turn a merge into a three-and-a-half minute operation and made the feature look unusable per-table; per-phase `Perf` timers isolated it to the single call. The staging table is already empty, which is exactly what "the source has no records" means, so the reconcile can proceed straight to deleting the table's rows.
- **Binary, complex, and calculated columns are skipped with a warning.** They cannot take part in a SQL comparison, and calculated columns cannot be assigned. Those tables still export, and still import on a full build.
- **A missing source file never deletes rows.** It means the table was dropped from `TablesToExportData` or the file was deleted, so only the index entry is removed.
- **XML rows are relabeled through the DOM, not by replacing tag text.** `ImportXML` takes the target table from the row element names, so the rows have to be renamed to reach the staging table — but a table is allowed to have a field with the same name as the table, and a textual replacement would rename that field element too. Verified empirically before building on it: `SELECT INTO` preserves column types, `ImportXML ... acAppendData` loads renamed rows into a pre-created staging table with nulls and long memo values intact, and Jet's `<>` detects a memo difference past 255 characters.

**What this rules out**: Table data no longer needs a full build to move records between databases. `ComponentTypeSupportsScopedImport` still rejects `edbTableData`, so `VCS.ImportByType("table_data")` continues to error — scoped sync takes no database backup, which deserves its own decision rather than being inherited from this one. Revisit if a project needs per-table control (the option is deliberately global) or an "insert and update but never delete" mode; both were considered unnecessary until someone has the use case.

**Relevant files**:

- `Version Control.accda.src/modules/Components/clsDbTableData.cls` — `IDbComponent_Merge`, `GetMergeStrategy`, `LoadStagingTable`, `WriteStagingXml`, `ReconcileTableData`
- `Version Control.accda.src/modules/Utility/modDatabase.bas` — `GetTableMergeKey`, `GetTableMergeStrategy`, `GetFirstDependentTable`, staging table lifecycle
- `Version Control.accda.src/modules/Core/modBuild.bas` — merge skip now gated on the option
- `Version Control.accda.src/modules/Infrastructure/clsOptions.cls`, `forms/frmVCSOptionsBuild.*` — the new option
- `Version Control.accda.src/modules/Tests/Components/modTestTableData.bas` — reconcile, composite key, null/memo, non-mergeable, missing file, and rollback tests

---

## 2026-07-28 — Deterministic table data export row order

**Trigger**: Exported table data (especially XML format) could appear in different row orders between exports even when no records changed, producing noisy git diffs. Tab-delimited export already used `ORDER BY` but sorted on every non-binary column (expensive and fragile on linked tables with unsortable memo/text columns).

*(Consolidated 2026-07-28: this entry absorbs the original sanitization-sort decision and the three same-day revisions that followed it — engine-side sorting, the sort-key rewrite, and the `EscapeXmlName` fix. The parser-side DOM sort described here is now a fallback, not the primary path.)*

**Options explored**:
- **Sort in post-export XML sanitization** — reorders the parsed DOM after `Application.ExportXML`, which has no ordering parameter. Retained only as the fallback for schema-bearing tables; see the known limitation below.
- **Export via temporary `ORDER BY` query with `acExportQuery` (chosen for XML)** — initially rejected on the belief that losing embedded-schema annotations would break `ImportXML`. Re-probed and adopted; see below.
- **Sort on all non-binary columns (status quo for TDF)** — rejected for performance: replaced with primary-key sort when available (index-backed scan).
- **Use the table's saved datasheet `OrderBy`** — rejected on technical grounds: `ExportXML` ignores the property entirely (see probe table).
- **Gate behind `eExportFormatVersion`** — rejected: treated as a bug fix; `clsDbTableData.IsModified` already forces re-export of all table data on every run, so users get a one-time reorder diff regardless.

**Decision**: Shared `GetTableSortFields` in `modDatabase.bas` picks sort fields (primary key → unique+required index → all non-binary fields). Tab-delimited export uses it for `ORDER BY`, with a warning and unsorted fallback when the sort query fails (e.g. linked SQL Server memo columns). XML export routes through a temporary `ORDER BY` query so the database engine does the ordering; schema-bearing tables fall back to `acExportTable` plus the parser-side DOM sort. `IndexAvailable` moved to `TableIndexesAvailable` in `modDatabase` for reuse.

**Probe findings — why `acExportQuery`**: tested against a live Access instance on a table whose rows were inserted deliberately out of key order.

| Export mode | Row order |
|---|---|
| `acExportTable` | insertion order |
| `acExportTable` with `OrderBy` + `OrderByOn` set | insertion order (**property ignored**) |
| `acExportQuery` with `ORDER BY` | correctly sorted |

The schema objection turned out to be mostly moot: a query export does lose `<od:index primary="yes">` and downgrades the key field from `minOccurs="1"` / `od:nonNullable="yes"` to `minOccurs="0"`, but `SanitizeXML` already discards the entire schema unless it contains `od:expression`, `od:jetType="complex"`, or `od:jetType="oleobject"`. For every other table the schema never reaches disk either way. A round-trip probe confirmed the renamed query output re-imports through `ImportXML` to the correct table name with identical row content, and that its `<dataroot>` element and row markup match the table export byte for byte.

Row elements take the query name and are renamed back to the table name with a string replace, which is safe because XML escapes `<` in character data, so `<queryname>` can only occur as markup. `TableRequiresXmlSchema` routes calculated, complex, and OLE object tables to `acExportTable`, and a post-sanitization check for a surviving `xsd:schema` acts as a correctness net in case that detection is ever wrong. Any failure — including a read-only database that cannot host the temp query — falls back to the table export.

**Temp query lifecycle — one per operation, not one per table**: the first implementation created and dropped a temp `QueryDef` for every table, calling `QueryDefs.Refresh` on each side. On a database with 3681 saved queries that was catastrophic: 64 full-collection refreshes pushed the Table Data category from 5.95s to 19.97s, with 15.46s of it (77%) invisible because none of the new code was instrumented.

The query is now created once per export operation and repointed by assigning `.SQL` per table — empirically confirmed that `ExportXML acExportQuery` observes a reassigned `.SQL` without any collection refresh, so `QueryDefs.Refresh` runs once at creation instead of twice per table. `PrepareTableDataSortExport` sweeps leftovers from an interrupted run; `ReleaseTableDataSortExport` drops the query and is called from every export cleanup block.

`AssignTableDataSortQuery` must return an empty string on any failure, because the query would otherwise still hold the *previous* table's SQL and the caller would export that table's rows into this table's source file. Note that the engine defers table-name resolution, so SQL naming a missing table is accepted at assignment and only fails later inside `ExportXML`, where the existing fallback handles it. A lookup failure (query deleted mid-operation) is distinguished from an assignment failure and recreates the query in the same call.

**Sort key construction**: the initial XML sort-key builder prefixed each normalized field value with its *length* (`Format$(Len(strPart), "0000")`), making string length the primary sort dimension — tables with variable-length text primary keys exported rows grouped by `Len(value)` rather than by value. Fixed with `ComposeXmlSortKey` / `XmlSortKeyOrdinal` in `modEncoding.bas`: each normalized part is terminated with `vbNullChar` (unambiguous without length-prefixing) and a fixed-width ordinal suffix supports stable sort. Comparison uses `QuickSortStringsBinary` with `vbBinaryCompare` so `Chr$(1)`–`Chr$(5)` sentinels are not ignored under `Option Compare Database` collation. XML text ordering is therefore ordinal while tab-delimited uses Access `ORDER BY` collation; each table uses one format only, so per-table determinism holds.

**`EscapeXmlName` underscore rule**: Access escapes `_` to `_x005F_` when, and only when, it is immediately followed by a **lowercase** `x` (probed: `a_x1` → `a_x005F_x1`, `b_X2` → `b_X2`, and `a_xZZZZ_b` → `a_x005F_xZZZZ_b` even though the hex digits are invalid). Treating `_` as a plain name character meant field names containing `_x` never matched when building sort keys, so every row collapsed to the null sentinel and sorted arbitrarily — a silent wrong answer rather than a visible failure.

**Performance** (same database, 3681 queries, table-data-only fast-save export):

| Implementation | Table Data | Per table | Uninstrumented |
|---|---|---|---|
| No sorting (baseline) | 5.95s / 30 | 0.198s | 0.38s (6%) |
| Parser-side DOM sort | 7.45s / 30 | 0.248s | 0.34s (5%) |
| Temp query per table | 19.97s / 32 | 0.624s | 15.46s (77%) |
| Temp query per operation | **4.90s / 32** | **0.153s** | ~0.3s (6%) |

The final result beats the unsorted baseline because `acExportQuery` is itself cheaper than `acExportTable` (0.113s vs 0.163–0.190s per call); engine-side `ORDER BY` costs less than the DOM sort it replaced. `Application.ExportXML` is now 72% of the remaining category time, so further gains would require emitting the XML directly from a sorted recordset — not attempted, given the `ImportXML` compatibility risk.

**Instrumentation**: `Assign Temp Sort Query`, `Drop Temp Sort Query`, `Check Calculated Fields`, and `Build Table Sort Fields` were added so this path can never again hide its cost in `Other Operations`. A `Drop Temp Sort Query` count above 1 for a whole export means the per-operation lifecycle has regressed.

**Known limitation**: the parser-side DOM sort retained for schema-bearing tables indexes rows with `objRows.Item(i)` on a live `ChildNodes` list, which is a sibling-chain walk. Measured per-access cost scales linearly with row count (16.7 µs at 5K rows rising to 104.1 µs at 40K), making that loop O(N²); caching nodes during a single `For Each` pass measured flat at ~2.3 µs/node. This path is now reached only by calculated, complex, and OLE object tables, which is why it was left alone.

**What this rules out**: Relying on engine iteration order for XML table data. Using saved datasheet `OrderBy` for export ordering (Access ignores it). Creating or dropping the temp query per table, or calling `QueryDefs.Refresh` anywhere in the per-table path. Sorting on every column when a primary key exists. Length-prefixing normalized sort-key parts. Treating `_` as always safe in exported XML element names. Memoizing `TableRequiresXmlSchema` or `IsLocalTable` per operation — both are called once per table, so the caches measured 0% hit rate and were removed; `Check Calculated Fields` is only 0.06s across 31 tables.

**Relevant files**:
- `modDatabase.bas` — `GetTableSortFields`, `TableIndexesAvailable`, `IsBinaryTableFieldType`, `TableRequiresXmlSchema`, temp query lifecycle (`PrepareTableDataSortExport`, `AssignTableDataSortQuery`, `ReleaseTableDataSortExport`)
- `clsDbTableData.cls` — shared sort SQL, `ExportTableDataAsXml` query/table routing, TDF fallback
- `modExport.bas` / `clsVersionControl.cls` — `Prepare`/`Release` calls in each export entry point and cleanup block
- `clsSourceParser.cls` — `RowSortFields`, `SortXmlDataRows`
- `modEncoding.bas` — `EscapeXmlName`, `NormalizeXmlSortValue`, `NormalizeNumericXmlSortValue`, `ComposeXmlSortKey`, `XmlSortKeyOrdinal`
- `modFunctions.bas` — `QuickSortStringsBinary`
- `modTestTableData.bas` — unit/integration tests, including temp query SQL reuse and recovery

---

## 2026-07-28 — Conditional formatting: crash-proof decode, and which Boolean encoding to emit

**Trigger**: Issue #730 — data bar conditional formatting was silently lost on export.
Investigating it against a purpose-built corpus of 42 controls captured from Access
(`SaveAsText` output for every rule shape the add-in handles) turned up three separate
crashes in the decode path plus a systematically wrong legacy emitter, and then a genuine
ambiguity in how Access encodes Booleans.

Any unhandled error in the decode path is the *same bug as #730*, because a failed decode
means the control's formatting does not reach the companion JSON. Three were found:

1. The data bar record layout assumed a per-rule prefix on rule 0, which has none.
2. `ReadLong` computed the high byte as `byte3 * &H1000000`, which overflows a signed
   `Long` once `byte3 >= &H80`. A rule with `FontUnderline` set from VBA stores `&HFF`
   there, so every such form raised error 6.
3. `HexToBytes` drove its loop from the hex string length rather than the byte count, so an
   odd-length block (a truncated or hand-edited `.form`) raised "subscript out of range".

**The ambiguity**: Access writes `1` for a flag byte it computes itself, and `&HFF` for a
Boolean it copied from VBA (VBA `True` is `-1`, truncated to a byte). Both appear in real
databases, Access preserves whichever it finds, and it never rewrites the blocks on a plain
save — confirmed by round-tripping a form through design view. The same split applies to the
legacy expression slots: dialog-authored rules allocate one null unit less per expression
rule (1 rule / 13 chars → 126 bytes from the dialog vs 128 from VBA; 3 rules / 29 chars →
282 vs 288). Field-value rules match, since both populate two slots. So a rebuilt block
cannot be byte-identical to both encodings.

**Options explored**:
- **Preserve the original encoding in the companion JSON** as optional fidelity fields
  (the existing `TrailerColor` precedent) — byte-exact for both families, at the cost of two
  new JSON fields carrying an Access implementation detail. Rejected as not worth the schema
  surface: the choice has no functional or source-churn consequence (see below).
- **Always emit the design-view encoding** (`1`, slots `len + 2`) — byte-exact for typical
  user forms and for the pre-existing fixtures and docs. Rejected.
- **Always emit the VBA encoding** (`&HFF`, slots `max(len + 1, 2)` per slot) — chosen. It is
  what the 42-control corpus verifies byte-for-byte, and the corpus is the artifact future
  changes will be tested against.
- **Treat the legacy block as optional and stop emitting it** — not pursued now, though
  Access demonstrably tolerates its absence. Worth revisiting separately.

Why the choice is low-stakes either way: the exported source stores decoded rules as JSON
booleans, not hex, so neither encoding causes git churn on re-export; Access reads any
non-zero byte as true; and CF14 is the authoritative block, with the legacy block existing
only for Access 2007 compatibility.

**Decision**:
- Read the format flags **byte-wise**, never as a dword. They are four independent Boolean
  bytes, not a bitfield.
- `ReadLong` folds the sign bit in separately, and `WriteLong` masks each byte before
  dividing, so the two are exact inverses over all 2³² patterns (integer division truncates
  toward zero and produced wrong high bytes for negative values).
- Screen hex input with `IsHexBytes` before converting; malformed input sets `DecodeFailed`
  so the caller keeps the inline binary block rather than raising.
- `FlagsToLong` replaced by `WriteFlags`, emitting bytes directly: `Enabled` as `1`, each set
  font flag as `&HFF`. Same convention for the data bar `showBarOnly` byte.
- `BuildLegacyHex` uses a unified slot model — every rule owns two slots sized
  `max(len + 1, 2)` — with descriptor dwords naming the next rule's slot window, and emits
  no block at all for data-bar-only controls.
- Byte-exact assertions live in the new `modTestConditionalFormatCorpus` (42 captures,
  generated from capture files rather than transcribed). `modTestConditionalFormat` keeps the
  design-view fixtures to assert the *other* encoding decodes and survives a rebuild with its
  model unchanged, and pins the slot-size arithmetic so a change to the convention is visible.

**What this rules out**: Do not read the format flags, or any CF dword, as a single `Long`
without accounting for the sign bit. Do not "fix" the design-view fixtures to be byte-exact —
they document a second valid encoding on purpose. Do not add an export format version gate
for these changes: the emitted artifact is the companion JSON, which is unaffected. Do not
transcribe hex fixtures by hand; generate them from captures.

**Relevant files**: `clsConditionalFormat.cls`, `modTestConditionalFormat.bas`,
`modTestConditionalFormatCorpus.bas`, `docs/access-conditional-format.md`

---

## 2026-07-27 — Pre-operation hook timing and letter-casing save reliability

**Trigger**: Review of when `RunBeforeExport` and `RunBeforeMerge` execute relative to
`CloseDatabaseObjects` / the merge shift-reopen. `RunBeforeExport` ran before objects
were closed, so a hook that modified an open object could have its change overwritten
when `DoCmd.Close` saved the stale in-memory design, and hook side effects persisted even
when the subsequent close failure aborted the export. `RunBeforeMerge` never fired: the
build read `dNZ(Options.GitSettings, "RunBeforeMerge")`, but `GitSettings` was a vestigial
dictionary (default-populated only, never serialized to `vcs-options.json`, never written
by any form). The options UI reads/writes `Options.RunBeforeMerge`, which nothing
consumed. Separately, intermittent save prompts during builds pointed at
`StandardizeLetterCasing`: its post-correction `DoCmd.Save` did not set
`VBE.ActiveVBProject` and swallowed failures with `On Error Resume Next`.

**Options explored**:
- **Leave `RunBeforeExport` before close** — rejected: clobber risk on open objects and
  wasted hook side effects when close aborts the export.
- **Re-close only when a hook is configured** — chosen for export: avoids extra close
  work on the common path (no hook), but sweeps up anything the hook opened.
- **Run `RunBeforeMerge` before shift-reopen** — rejected: session-scoped hook state is
  destroyed by `CloseCurrentDatabase2` / `ShiftOpenDatabase` before merge work starts.
- **Run `RunBeforeMerge` after shift-reopen** — chosen: matches `RunBeforeBuild`, which
  runs after the target database exists in the state the operation will use.
- **Keep `GitSettings` with a comment** — rejected: nothing serializes or reads it;
  delete the dead surface.
- **Letter-casing save: silent `Err.Clear`** — rejected: failures left the VBA project
  dirty with no log entry; set active project before save, log via `CatchAny`, warn if
  `CurrentVBProject.Saved` is still False.

**Decision**:
- `modExport.ExportSource`: close objects first; run `RunBeforeExport`; re-close only
  when a hook ran; then save unsaved VBA. Extracted `CloseObjectsOrAbort` helper.
- `modBuild.Build` (merge): remove dead `GitSettings` read; run `Options.RunBeforeMerge`
  after the close/shift-reopen, with `Log.Flush` and trailing `CatchAny`.
- `modLetterCasing.StandardizeLetterCasing`: `Set VBE.ActiveVBProject = CurrentVBProject`
  before `DoCmd.Save`; log save failures; warn when project remains dirty.
- Remove `Options.GitSettings` from `clsOptions`.

> **⚠ Partially superseded** (2026-07-29): The letter-casing and export saves no longer use
> `DoCmd.Save acModule`. Saving one module does not save the project when form or report
> class modules are dirty, so those saves silently failed — which is what the "project
> remains dirty" warning added here had been reporting all along. Both now call
> `SaveCurrentVBProject`, which saves via the worker and returns the resulting `Saved`
> state. The casing pass also no longer runs on merges. See "Opt-in in-place merge
> preparation instead of the pre-merge reopen" above.

**What this rules out**: Do not read merge hook names from `GitSettings` again. Do not
move `RunBeforeMerge` before the shift-reopen without re-evaluating session lifetime.
Do not restore silent error swallowing on the letter-casing save.

**Relevant files**: `modExport.bas`, `modBuild.bas`, `modLetterCasing.bas`,
`clsOptions.cls`

---

## 2026-07-27 — Schema export cache-bust via per-schema fingerprint in the index

**Trigger**: Installing or removing `sp_GetDDL` on a SQL Server switches every exported
object between rich SP-generated DDL and the built-in `object_definition()` fallback, but
nothing re-exported. Schema exports bypass `VCSIndex` and `CategoryHashes` entirely;
their only change signal is timestamp equality (`ExportObject` stamps each `.sql` with the
server's `last_modified`, and the next scan compares the two). A capability change leaves
both sides of that comparison identical, so `IDbSchema_Export` short-circuits before
opening a connection. Investigation also found `blnFullExport` was declared in
`IDbSchema.Export` and both implementations but **never referenced** — so no supported
way to force a schema re-export existed at all.

**Options explored**:
- **Reuse `GetExporterRevisions` / `CategoryHashes`** — rejected on two counts:
  `modExport` replaces `VCSIndex.CategoryHashes` wholesale after each export
  (`Set .CategoryHashes = dCurrentHashes`), so extra `"Schema:<name>"` keys are clobbered;
  and probing for the SP needs a live connection, while `GetCategoryHashes` runs offline
  for `frmVCSMain` UI state.
- **Sidecar state file** in `databases/<name>/` — viable (verified safe from the orphan
  and empty-folder cleanup passes, which only consider `*.sql` inside known base
  subfolders), but adds another gitignore rule to ship to user projects.
- **Marker comment in each exported `.sql`** naming the DDL source — rejected: changes
  exported content, which is `eExportFormatVersion` territory, and adds per-file I/O.
- **New `SchemaState` section in `vcs-index.idx` (chosen)** — the index is already the
  gitignored home for local derived sync state, and a per-key accessor avoids the
  wholesale-replacement problem that rules out `CategoryHashes`.

**Decision**:
- Index format bumped 3 → 4, appending `SchemaState` (`{schema name → fingerprint}`).
  The reader accepts versions 3 and 4 via `cintIdxMinReadVersion`, so upgrading does not
  discard an existing index — the previous `<> cintIdxFormatVersion` check would have
  deleted it and forced a full project re-export.
- Fingerprint covers both an exporter revision (`SCHEMA_EXPORTER_REVISION_MSSQL` /
  `_MYSQL` in `modConstants.bas`) and any runtime capability affecting output. Only MSSQL
  has the latter today (`sp_GetDDL`); MySQL always uses `show create ...`.
- `VCSIndex.SchemaState` is a **parameterized property**, not a whole-dictionary
  Get/Set like `CategoryHashes`, precisely because this section is not rebuilt from
  options on each export. `TextCompare` matches `Options.SchemaExports` name handling.
- Both exporters now honor `blnFullExport`, and an unrecorded fingerprint forces one
  baseline re-export. Content is normally byte-identical and files are re-stamped with
  the same server dates, so git shows nothing.
- The fingerprint is recorded only when `Operation.ErrorLevel < eelError`, so a canceled
  or failed export does not claim credit for files it never wrote. An undeterminable SP
  status (failed connection) returns an empty fingerprint and forces nothing, rather than
  guessing "builtin" and re-exporting everything.
- When `VCSIndex.Disabled` there is nowhere to record the result, so the check is skipped
  — otherwise every run would re-export the whole schema.
- `CanUseGetDDL`'s cached status moved from a procedure `Static` to instance-level
  `m_intSpStatus` so `GetFingerprint` can distinguish *unavailable* from *unknown*.
  Behaviorally identical (VBA `Static` in a class procedure is already per-instance, and
  a fresh exporter instance is created per schema per export).

**What this rules out**: Treating timestamp equality as a complete change signal for
external schema exports. Any future change to how DDL is generated — in our code or in
the server environment we probe — must be reflected in `GetFingerprint`, or it will not
reach existing projects. Also rules out storing per-schema derived state in
`CategoryHashes`, which is rebuilt from options and cannot hold it.

**Relevant files**: `clsVCSIndex.cls`, `clsSchemaMsSql.cls`, `clsSchemaMySql.cls`,
`modConstants.bas`, `modTestIndex.bas`, `AGENTS.md`.

---

## 2026-07-27 — Per-category exporter revisions for cache-bust bug fixes

**Trigger**: Bug fixes to command-bar `_Images` sidecar export changed output without
changing primary `GetSource` bytes or `DateModified`. The change index reported no
modification, so fast save skipped re-export and the fix never reached existing
projects. Bumping `eExportFormatVersion` for each such fix would force a full
project export and proliferate permanent `If >= EFV_...` branches.

**Options explored**:
- **Patch-level `eExportFormatVersion`** — rejected: `_Global` hash forces full export;
  every gate is a permanent opt-in branch users must carry forever.
- **Per-object predicate in exporter code** (`If IsUnionQuery Then ...`) — rejected:
  permanent per-fix branches; no better git outcome than category-level re-export
  because unaffected objects serialize to byte-identical output.
- **Explicit per-category revision in index + opt-in UI** — rejected for v1: honors
  deferral but adds index format field and UI affordance; category re-export cost is
  bounded (~minutes worst case) and git diffs stay surgical.
- **Fold revision into `CategoryHashes` (chosen)** — `GetExporterRevisions()` in
  `modConstants.bas` returns `{categoryName → revision}`; `GetCategoryHashes` seeds
  `ExporterRevision` into each listed category (creating the dict on demand so
  categories with no classified export options, e.g. `CommandBars`, still get a hash).
  Existing `dStaleCategories` path in `modExport` re-exports that category once;
  new hash is persisted self-clearingly.

**Decision**:
- `GetExporterRevisions()` is the single registry; bump on blind-spot fixes only
  (sidecars, date-fast-path). Do not bump for content-hashed primary output
  (`IsModified` self-heals) or opt-in format changes (`eExportFormatVersion`).
- Initial entry: `CommandBars = 1` for `_Images` sidecar export fix.
- Pattern mirrors `LAYOUT_SVG_GENERATOR_VERSION` intent but uses category hash
  invalidation instead of a dedicated comparison in each component.

---

## 2026-07-27 — Hydrate prior test metadata in the web runner after tree publish

**Trigger**: `test-state.json` already stores per-test `durationMs`, assertion
counts, and `lastRunAt`, but the web runner showed test names only until Run on
the `VCS.RunTests` path. `LoadInto` was gated behind `m_blnStandalone`
(`VCS.OpenTestRunner` only).

**Options explored**:
- **Embed metadata in the tree JSON** — rejected: duplicates the batch replay
  path and enlarges every `onReady` payload with full assertion arrays.
- **Replace `LoadStateTests` before scan** — rejected: `LoadStateTests` wipes
  `this.Tests`; scan must run first so discovered tests match the live VBA project.
- **Publish tree, then overlay + one `onResultsBatch` (chosen)** — scan and
  `onReady` paint names immediately; `modTestState.MergeInto` overlays durable
  state onto discovered keys; one batch streams prior results. Cost is one
  `ParseJson` (~240 KB in this repo) — negligible vs VBIDE scan.

**Decision**:
- `clsTestRunner.MergeStateResults` overlays disk state by `Module.Proc` key;
  `ApplyStateResultsToTest` is shared with `LoadStateTests`.
- In-memory `fromPriorRun` (distinct from persisted `stale`) marks hydrated rows
  for the UI; cleared when a test actually starts running.
- JS keeps a `priorResults` map that survives `onRunStart` reset so not-yet-run
  rows still show last session's duration during a long run.
- `m_blnPriorStateLoaded` prevents re-reading disk on toolbar Refresh (merge-scan
  preserves in-memory results) **and across `DocumentComplete` re-fires**. It is
  deliberately not reset there: the page is wiped but the singleton is not, so
  `RefreshWebTestTreeDeferred` replays from memory instead of parsing again. Resetting
  it meant any spurious reload paid for a second parse and blinked the indicator a
  second time. Only a real form unload (`ResetWebRunnerHostState`) clears it.
- The parse is visible on open, and the tree is already painted when it starts, so
  it gets an in-page indicator rather than relying on the Access hourglass (easy to
  miss over a WebView2 pop-up): `onHydrateStart` / `onHydrateEnd` drive the header
  status badge (pulsing **Loading previous results…**, sharing the run-status slot)
  plus a stats-bar chip. A chip alone tested as too easy to miss.
- JS enforces `HYDRATE_MIN_VISIBLE_MS` (600 ms) before hiding. The parse often
  finishes fast enough that a truthful indicator is unreadable, so the hold buys
  legibility; `onRunStart` passes `force` to drop it immediately, since the run's own
  status must own the badge.
- **The overlay itself is deferred to `Form_Timer`**, not run inline. An inline parse
  with a `DoEvents` before it was tried first and the chip never appeared: WebView2
  composites no frames while VBA holds the thread, so the chip was created and removed
  without ever being painted. `ScheduleHydratePriorResults` only sets a flag and pushes
  `onHydrateStart`; `PumpDeferredHydrate` runs after `DrainOutbox` on the next tick, so
  user commands keep priority. A run supersedes a pending hydrate (`AcceptBridgeRun`
  clears the flag).
- **Reaching the timer is still not sufficient** — `PumpDeferredHydrate` waits for the
  page to *confirm the paint*. The diagnostic trace (7/27) showed only 49 ms between the
  `onHydrateStart` push and the first tick, followed by a **1.76 s** parse with the
  thread held, so the indicator was pushed, never composited, then removed. JS sets
  `window.__hydratePainted` from a **double `requestAnimationFrame`** (the first
  callback still precedes the frame carrying the change; the second follows it), and VBA
  polls that flag, capped by `HYDRATE_PAINT_TIMEOUT_MS` so a bridge failure cannot stall
  the data. Wall-clock delays were rejected: they guess at compositor timing, and this
  is the one signal that actually means "on screen."
- The warm-reuse path now clears the hourglass *after* refresh, not before.

**What this rules out**: Showing prior pass/fail as authoritative without the
dimmed/stale styling — hydrated rows are explicitly "previous run" until re-run.

**Relevant files**: `clsTestRunner.cls`, `modTestState.bas`, `modTestRunnerUI.bas`,
`TestRunner/runner.html`.

---

## 2026-07-27 — Grid origin is stored Top-first in LvExtra, Left-first in the qdef

**Trigger**: Query `.json` files flipped `GridLeft` and `GridTop` on every round trip
(243 of ~1800 queries in one production database; the rest have `0, 0`, which is why
no fixture caught it).

**Decision**: The `LvExtra` blob stores the grid origin as `(Top, Left)` — reversed
relative to every other RECT in that format. `clsLvExtraParser` now assigns the first
Long to `gridTop`. The qdef layout block keeps the opposite order: `EmitDesignLayout`
writes `Left` before `Top`, because that is what `LoadFromText` demands.

**Alternatives considered**: Reversing the qdef emitter instead. Rejected empirically —
Access rejects the whole Design View import with `Expected: 'Left'. Found: Top.`, and
the importer silently falls back to SQL View, dropping the layout for every Design View
query. The round-trip harness caught this across 21 fixtures.

**What this rules out**: Assuming the blob and the qdef agree on coordinate order. They
do not, for this one field. Do not "tidy" either side into matching the other.

**Fixture**: `qryRegressionGridOrigin` pins an asymmetric nonzero origin; every other
fixture in the corpus has `0, 0` and cannot detect a swap.

---

## 2026-07-27 — Layout pipeline gets DB-free unit tests, not a nested round-trip subset

**Trigger**: The grid origin swap survived every normal test run and only surfaced under
`VCS.RunRoundtripTests`. Layer 1 covered SQL text and `MSysQueries` rows; nothing below
the round-trip harness touched the binary `LvExtra` blob or the qdef layout block, so the
one field where the reader and the writer disagree had no unit-level coverage at all.

**Decision**: Add `clsTestQueryLayout` — a database-free class that synthesizes an
`LvExtra` blob, parses it with `clsLvExtraParser`, and feeds the result through the public
`clsQueryComposer.GenerateQdef`. Three tests pin the asymmetry: the blob reads Top-then-Left,
the qdef emits Left-then-Top, and the composition of the two preserves the original values.
The probe values are asymmetric and distinct from every other coordinate, so a swap cannot
hide behind a coincidentally matching number.

**Alternatives considered**: Running a curated subset of the round-trip corpus inside the
normal suite. Rejected for now — `RunObjectRoundtripTests` calls `Operation.Begin` and
`Log.Clear`, so nesting it inside a test run (itself an `eotTestRun` operation) would fail
to start and would wipe the console output. Making the harness re-entrant is a real option,
but it is a change to singleton ownership and not one to make immediately before a release.

**What this rules out**: Treating "it round-trips through Access" as the only available
check for binary-blob fields. Where a reader and a writer must disagree about ordering,
pin each end separately against ground truth — a composed test alone passes when both
ends flip together.

---

## 2026-07-27 — Bracket-aware table-ref parsing in clsSqlSyntax

**Trigger**: Production database build failures when join operands used bracketed
multi-word table names (`[Car Models]`). `TryExtractSimpleTable` un-bracketed
then split on the first space, truncating names and corrupting `MSysQueries`.

**Decision**: Extract a pure `clsSqlSyntax` functional core with shared `SplitTableRef`
(bracket-aware `AS` detection, no space split). Both `AddInputTable` and
`TryExtractSimpleTable` derive the same reference key from identical operand text.
`clsQueryComposer` delegates parsing helpers to a per-instance `m_syntax` member.

**Alternatives considered**:
- Surgical fix (only guard the space split for bracketed names) — rejected; left
  embedded-`AS`, embedded-`JOIN`, and implicit-alias inconsistencies in place.
- Standard module (`modSqlSyntax`) — rejected; would add ~15 public names to an
  already crowded global namespace; class scoping keeps IntelliSense clean.

**What this rules out**: Naive `InStr`/`Split` on un-bracketed table operands in
join parsing. Any new table-ref extraction must go through `SplitTableRef` or share
its bracket-aware keyword scanning.

**Fixtures**: `qryRegressionSpacedTableNameJoin` (Layer 2), `clsTestSqlSyntax` and
`clsTestQueryComposerJoins` (Layer 1 matrix + closed loop).

---

## 2026-07-27 — Web runner: copy test-state path and save TestRun log

**Trigger**: Users want to paste a test-results file path into an agent chat after a
web-runner run. Investigation also found web runs never called `Log.SaveFile`, so
`TestRun_*.log` was not written despite `Log.Active = True` during bridge runs.

**Options explored**:
- **`test-results.html`** — rejected for agent handoff: same data as state JSON but
  inlined in a large HTML/CSS/JS shell; poor token efficiency for agents.
- **`logs/TestResults_<timestamp>.json`** — rejected as the copy target: ephemeral
  per-run filename changes every run; `test-state.json` has a stable path and merges
  partial runs.
- **`logs/TestRun_*.log`** — useful for human debugging but not the primary agent
  artifact; log save is fixed separately so both tiers exist after web runs.
- **`test-results/test-state.json` (chosen)** — stable path, always written by
  `MergeAndSave`, rich per-test fields (`moduleName`, `procName`, `line`, assertions,
  `loggedErrors`, tags).

**Decision**: Add **Copy path** toolbar button and `CopyResultsPath` bridge callback
that copies the bare `GetStateFilePath()` string via `SetClipboardText`. Fix web-runner
teardown to call `SaveWebRunnerRunLog` (`Perf.EndTiming`, perf report, `Log.SaveFile`)
in `EndInteractiveBridgeRun` and on execute-phase errors before `Operation.Finish`.
Add a self-describing run heading when `blnInvokeSetup` is true.

**What this rules out**: Using `navigator.clipboard` in WebView2 for this action (VBA
clipboard helper is already tested). Copying HTML report or timestamped JSON paths as
the default agent handoff. Leaving web runs without a persisted `TestRun_*.log`.

**Relevant files**: `TestRunner/runner.html`, `modTestRunnerUI.bas`, `AGENTS.md`,
`.cursor/rules/testing.mdc`.

---

## 2026-07-24 — Defer dedicated decision-document (ADR) infrastructure

**Trigger**: Some decisions involve far more reasoning and trade-off analysis than a `DECISIONS.md` entry seems able to hold (a long deliberative session on per-component "exporter revision" invalidation was the prompting example). Question raised: should heavyweight decisions get comprehensive standalone ADRs under a new folder, with the lightweight log linking out to them? Concern was that an agent lacking full context might reopen a settled decision.

**Options explored**:
- **Curated ADR folder** (`docs/decisions/`, dated files, template, log links out via a `Full rationale` field): explored in detail and initially planned. Rejected after auditing the log: only ~3–5 of ~102 entries would clear a sensible promotion bar, and the strongest candidates (e.g. the round-trip harness entry with 9 options) already carry their reasoning fully in the log. Standing up a folder + template + README + cross-references in four files to serve ~3% of cases is disproportionate, and adds permanent cost of carry (a second place to keep in sync, drift risk, per-doc sanitization, and exactly the "more data to sift through" problem we were trying to avoid).
- **Reconstruct ADRs retroactively from agent transcripts**: rejected. Backfilling settled history yields low-value archive; second-hand distillations look authoritative but can quietly encode a wrong rationale (fidelity risk); each transcript needs the same sanitization pass fixtures get (production names). Transcripts already preserve the full raw record and can be mined on demand.
- **Link the originating transcript from a log entry**: rejected as a committed convention. Transcripts are local, per-user, and uncommitted, so a link in the committed log is a dangling reference on any clean clone and may point at unsanitized content.
- **Do nothing / rely on the existing log**: chosen. The premise (the log format can't hold big decisions) did not survive scrutiny — the "aim for 10–50 lines" guideline is a soft default the log already breaks when a decision earns it. The real worry ("an agent reverses a decision without context") is what the `What this rules out` section already exists to prevent; a thin entry is fixed by writing a better paragraph, not by new infrastructure.

**Decision**: Do not build ADR infrastructure now. This is a YAGNI situation — no concrete case has yet shown a log entry failing in a way an ADR would have prevented. Keep `DECISIONS.md` as the single home for architectural rationale, and invest in richer `What this rules out` / `Options explored` sections when a decision warrants depth. If a log entry ever genuinely proves insufficient for a specific contested decision, create one standalone doc at that moment and let the real need define its format.

**What this rules out**: Creating `docs/decisions/`, an ADR template, or a bulk transcript-reconstruction effort without a demonstrated, concrete failure of the log first. Revisit only if a real case arises where the log measurably fell short (an agent reopened a well-documented decision, or a clean-clone contributor lacked reasoning that only a transcript held). `docs/` remains for sustained internal *reference* material about how systems work — not one-shot decision rationale, which stays in this log (see 2026-04-27 entry).

**Relevant files**: None (documentation/process decision). Supersedes the shelved plan "ADR convention for heavyweight decisions."

---

## 2026-07-20 — Index companion `.json` for merge detection and `AllFilesHash`

> **⚠ Partially superseded** (2026-07-29): `AllFilesHash` is no longer consulted on
> every merge. It is now reached only when the date+size property hash does not match
> the index, which is what made a no-change merge read every source file. The fix this
> entry describes is intact — when content *is* hashed, all indexed files are hashed,
> not the primary file alone. See "Merge scan reads no file content when dates and
> sizes are unchanged" above.

**Trigger**: Metadata-only edits to a form/report companion `.json` (Description, Hidden) were not picked up by `MergeBuild`. Root causes: (1) form/report `.json` was excluded from indexed `FileExtensions`; (2) even where `.json` was indexed (modules, queries, table defs, macros), `GetModifiedSourceFiles` confirmed timestamp drift via primary-file content hash only, masking companion-only changes. Macros also gated `.json` emission to the real export path while indexing it — a latent `GetDifferingFiles` false-conflict risk.

**Options explored**:
- **Rely on `MetaHash` only**: rejected. `MetaHash` reads live DB state and does not detect hand-edited source `.json`; the `.json` can also hold `ConditionalFormatting` and other sections beyond Description/Hidden.
- **Add `.json` to indexed set without fixing fallback**: rejected. Property-hash drift would still be dismissed when the primary file was unchanged.
- **Index `.json` + combined `AllFilesHash` content fallback**: chosen. Forms/reports add `json` to `efesIndexed` (`efesAll` adds `svg` only). All three components that gated metadata writes (form, report, macro) now emit `.json` on alternate/temp exports (matching modules/table defs/queries). `GetSourceFilesContentHash` hashes all indexed files; stored as `AllFilesHash` on index entries (flag bit 32, no index format-version bump). `GetModifiedSourceFiles` uses `AllFilesHash` when present; legacy entries fall back to primary-only until re-synced.

**Decision**: Comprehensive fix — the fallback benefits every multi-file component. Component-specific edits limited to form/report (index + alt emit) and macro (alt emit only).

**What this rules out**: Treating form/report `.json` as derived-only (`efesAll` sidecar). Derived `.svg` previews remain unindexed. Index format version was not bumped — `AllFilesHash` populates on next export/import per object.

**Relevant files**: `clsVCSIndex.cls`, `clsVCSIndexItem.cls`, `modContainers.bas`, `clsDbForm.cls`, `clsDbReport.cls`, `clsDbMacro.cls`, `modTestMergeDetection.bas`, `modTestOrphaned.bas`.

---

## 2026-07-20 — Scoped `FileExtensions` for artifact cleanup and moves

> **⚠ Partially superseded** (2026-07-20): Form/report companion `.json` is now in `efesIndexed` (authoritative metadata), not `efesAll` only. `.svg` remains `efesAll`-only. See "Index companion `.json` for merge detection and `AllFilesHash`" above.

**Trigger**: Orphan cleanup and `MoveSource` duplicated hardcoded extension lists (form/report `.json`/`.svg`, query legacy files, etc.) separate from `FileExtensions`, which intentionally excludes derived sidecars from the index because `GetDifferingFiles` uses a strict file-count match (see 2026-05-05 entry). A single declaration site was needed for “all files this component writes” without polluting change detection.

**Options explored**:
- **Add sidecars to `FileExtensions`**: rejected at the time for derived `.svg` and conflict noise; form/report `.json` was later indexed separately once alternate-path emission was aligned (see superseding entry).
- **Separate `ArtifactExtensions` property**: rejected. Third parallel list to maintain; same drift risk as hardcoded cleanup arrays.
- **Optional `Scope` on `FileExtensions`** (`efesIndexed` default, `efesAll` adds sidecars): chosen. Indexed consumers unchanged; `ClearOrphanedComponentArtifacts`, `MoveComponentSource`, and tests read `efesAll`. Folder artifacts (`_Images`, theme folders) are not object-named flat files and stay on the folder cleanup path.

**Decision**: `eFileExtensionScope` in `modConstants`; `clsDbForm` / `clsDbReport` branch on scope (`efesAll` adds `svg` only; `json` in `efesIndexed` since 2026-07-20). Shared helpers: `ClearOrphanedComponentArtifacts` and `MoveComponentSource`.

**Orphan-cleanup dispatch — no interface method**: An initial pass added an `IDbComponent.ClearOrphanedArtifacts` hook with 29 implementations. Once file cleanup became fully data-driven from `FileExtensions(efesAll)`, 27 of those implementations were identical no-ops and the hook had a single call site — pure boilerplate that VBA's lack of default interface methods forces onto every class. It was removed. `modOrphaned.ClearOrphanedSourceFiles` now calls `ClearOrphanedComponentArtifacts cType, dBaseNames` directly (data-driven files, covers form/report and any future sidecar automatically), plus `ClearOrphanedComponentFolders cType, dBaseNames` — a small `TypeOf` switch handling the only two folder-producing types (`clsDbCommandBar` → `_Images`, `clsDbTheme` → extracted folder). Rejected keeping the interface hook for uniformity: 25+ no-op overrides is worse maintenance than one localized 2-branch switch. This introduces a minor Core→Components reference in `modOrphaned`, accepted as the pragmatic home for the one bit of per-type folder knowledge (the suffix).

**What this rules out**: Using `FileExtensions` without a scope for cleanup/move — callers must pass `efesAll` explicitly when they need the full file set. Adding a new derived sidecar requires updating the component's `FileExtensions(efesAll)` branch only; it must not be added to `efesIndexed` unless it becomes authoritative tracked state. A new folder-producing component adds one branch to `ClearOrphanedComponentFolders` (and its own MoveSource folder handling) rather than an interface override.

**Relevant files**: `modConstants.bas`, `IDbComponent.cls`, `modOrphaned.bas` (`ClearOrphanedComponentArtifacts`, `ClearOrphanedComponentFolders`), `modContainers.bas` (`MoveComponentSource`), `clsDbForm.cls`, `clsDbReport.cls`, `modTestOrphaned.bas`, `modTestComponentInvariants.bas`.

---

## 2026-07-14 — Conditional formatting field-value operator decode (issue #725)

**Trigger**: Issue #725 — exporting a form with conditional formatting decoded every
field-value rule to `"Operator": "Between"` in the companion JSON, regardless of the real
operator (Equal, GreaterThan, etc.). `clsConditionalFormat.ParseStandardRule` hardcoded
`dRule.Add "Operator", "Between"`, and `LegacyOperator` always returned `0`. The original
#725 analysis correctly identified the hardcode but did not know where the operator lived
in the binary blocks, because the only field-value fixture (Text25) used operator `Between`
(value 0) and a white BackColor (zero trailer echo), so nothing exercised the difference.

**Empirical method**: Against a sample DB (`frmExample.txtFormatted`), drove
`FormatConditions.Add` through all eight `AcFormatConditionOperator` values plus a mixed
three-rule case, ran `Application.SaveAsText`, and diffed the exported `ConditionalFormat14`
/ `ConditionalFormat` hex with an authoritative byte dump. Findings:
- CF14 operator for rule 0 is a 2-byte LE value at header **offset 10** (previously labeled
  "reserved"); for later rules it is the **second dword of the 8-byte per-rule prefix**
  (previously labeled "reserved").
- Legacy operator is the dword at **offset 16** (location was already documented; the value
  was just never read/written as anything but 0).
- Bonus bug: the field-value CF14 trailer BackColor echo sits at **+5**, not +9. The echo is
  always `trailingLen - 12` into the trailer (+9 for the 21-byte expression/focus trailer,
  +5 for the 17-byte field-value trailer). Never caught because Text25 is white.

**Options explored**:
- **Gate behind a new `eExportFormatVersion`** (per the export-format-change rule): rejected.
  The JSON *structure* is unchanged — the `Operator` key already exists; only its value is
  corrected. Gating would leave existing 5.0.0 users exporting and building wrong operators
  until they opt into a new format. This is a correctness fix inside the already-5.0.0-gated
  `DecodeConditionalFormatting` feature, not a new formatting behavior.
- **Fix decode only** (populate JSON correctly, leave rebuild hardcoded): rejected — import
  round-trip would still rebuild Between, so a build from source would lose the operator.

**Decision**: Read the operator on decode (header offset 10 for rule 0, prefix+4 for later
rules) and map it to a name via new `OperatorToName`/`NameToOperator`; write it back on
rebuild in the CF14 header, the CF14 per-rule prefix, and the legacy header (offset 16).
Also corrected the field-value trailer echo offset in both decode and emit. Not gated behind
a new export format version (correctness fix to existing 5.0.0 output). Re-exporting a
project with non-Between conditional formatting will produce a one-time JSON diff (correct
operator, plus `TrailerColor` for colored field-value rules).

**What this rules out**: The CF14 header/prefix "reserved" bytes at offsets 10 and prefix+4
are no longer free — future format probing must not reuse them. Multi-rule *legacy*
field-value blocks embed per-rule operators in their descriptors and remain best-effort;
CF14 stays the authoritative decode source. Since the add-in cannot be rebuilt via MCP, the
byte-exact fixtures added to `modTestConditionalFormat` (all eight operators + a mixed
three-rule block, captured from real SaveAsText output) must be verified by running the test
suite after a manual add-in rebuild.

**Relevant files**: `modules/Core/clsConditionalFormat.cls` (decode/rebuild + operator
maps), `modules/Tests/Core/modTestConditionalFormat.bas` (operator + trailer-echo fixtures
and tests), `docs/access-conditional-format.md` (§4.1, §4.2, §4.3, §5.1, §10).

---

## 2026-07-13 — Fast-save query metadata scan via MSysObjects.LvProp

**Trigger**: On a production database with ~3,675 queries, a fast-save export sat at 1% for nearly a minute during "Scanning for changes...". Profiling (`Perf` ops added to `clsDbQuery.IDbComponent_IsModified` and `GetMetadataHash`) showed `Meta: Read Description` consuming ~38s — ~82% of the ~46s export. Each query's metadata hash read its `Description` via `dbs.Containers("Tables").Documents(name).Properties("Description")`, and that DAO/COM access forced Access to lazily materialize each query definition. Two symptoms: no progress feedback during the scan, and the scan itself being dominated by per-object Description reads.

**Options explored**:
- **Leave as-is, add progress only** — surfaces the stall but leaves the 38s cost. Rejected: treats the symptom, not the cause.
- **Cache Descriptions per `clsDbQuery` instance** — useless: each query is a fresh instance, so no sharing across the scan.
- **Batched `MSysObjects.LvProp` read (chosen)** — one snapshot recordset of `Name, LvProp` for all `Type=5` rows, parse the table-level `Description` from each `LvProp` blob (`clsLvPropParser`), cache by name for the duration of one scan. `GetHiddenAttribute` stays a per-object call (negligibly fast). Verified byte-for-byte equal to the DAO Description across all query rows (0 mismatches) with a throwaway `modScanDiagnostics.VerifyQueryMetadataSource` diagnostic. Batched read cost ~0.8s vs ~38s; total export ~46s → ~10s.

**Decision**: Added `BuildQueryDescriptionCache`/`ClearQueryDescriptionCache`/`GetQueryMetadataHash` to `modLoadSaveText.bas`; `clsDbQuery` builds the cache before the modified-scan loop and clears it after. The scan-time metadata hash uses the shared `HashMetadataValues` formula so the cached fast path and the generic `GetMetadataHash` produce identical hashes (unmodified queries are never falsely re-exported); when no cache is active, `GetQueryMetadataHash` falls back to the per-object DAO read. The cache always rebuilds fresh and drops itself on any build error, so it can never serve stale/partial data. Auto-generated `~sq_*` queries (present in `MSysObjects` but not `CurrentData.AllQueries`) are skipped. Separately, progress feedback was added: `Log.IncrementObjectScanProgress` increments once per object in every component's `GetAllFromDB`, the scan bar is sized to `GetQuickObjectCount` only, and `ClearOrphanedSourceFiles` no longer advances the bar (its near-instant per-file increments made the bar leap in bulk bursts). The one-time `modScanDiagnostics.VerifyQueryMetadataSource` verifier is dropped rather than carried forward — it was never committed, so its finding (LvProp == DAO, 0 mismatches) is recorded here; the verifier is simple enough to reconstruct from this entry if ever needed.

**What this rules out**: The fast path trusts that `LvProp`'s table-level `Description` equals the DAO document `Description`. If a future Access version diverges, only the metadata hash (Description/Hidden) is affected — a query's SQL change is still caught by `DateModified`, so the worst case is a missed description/hidden-only change or a harmless extra export, not data loss. That low severity is why no permanent CI guard was added; re-run the git-history diagnostic (or add a targeted `modTest*` check) if the invariant is ever suspected. The per-object scan increment only fires on the fast-save path, so a full export's scan phase shows a static bar briefly before the export loop reports progress.

**Relevant files**:
- `modLoadSaveText.bas` — `BuildQueryDescriptionCache`, `ClearQueryDescriptionCache`, `GetQueryMetadataHash`, `HashMetadataValues`, `ParseLvPropDescription`
- `clsDbQuery.cls` — batch cache around the modified-scan loop; `IsModified` uses `GetQueryMetadataHash` + `Perf` split
- `clsLog.cls` — `IncrementObjectScanProgress`
- `modExport.bas` — scan bar sized to `GetQuickObjectCount` only
- `modOrphaned.bas` — orphan-file scan no longer increments the bar
- `clsDbForm/Report/Module/Macro/TableDef/…` and `clsAdp*` — per-object scan progress increments

---

## 2026-07-13 — Oracle ODBC connectivity probe SQL

**Trigger**: Issue #723 — `CacheConnection` used `SELECT 1;` for all ODBC probes. Oracle rejects that syntax (requires `SELECT 1 FROM DUAL;`), causing ODBC error 3146 and a false Retry/Ignore/Abort dialog during build. A second bug: `clsDbConnection.Import` read `Err.Description` after `CacheConnection` had cleared `Err`, so the failure dialog showed no detail.

**Options explored**:
- **Reactive 3146 fallback** — retry with `FROM DUAL` when `OpenRecordset` fails. Rejected: 3146 is a generic ODBC wrapper that also covers auth and server failures, so retry conflates SQL syntax errors with real connection problems.
- **Registry DSN→driver lookup** for DSN-only strings without `DRIVER=`. Rejected for now: bitness-sensitive, heavier, and DSN-only Oracle links are rare with Access.
- **Proactive DRIVER-based detection (chosen)** — if `GetConnectPart(strConnect, "DRIVER")` contains `"Oracle"` (case-insensitive), use `SELECT 1 FROM DUAL;`; otherwise `SELECT 1;`. Covers DSN-less strings Access stores after linking.

**Decision**: Added `IsOracleOdbcConnect` and `GetConnectivityProbeSql` to `modConnect.bas`. `CacheConnection` and `TestBackEndConnection` call `GetConnectivityProbeSql` before `OpenRecordset`. `CacheConnection` now returns `strErrDesc` via ByRef (required parameter) so `HandleConnectionFailure` shows the real ODBC message.

**What this rules out**: DSN-only Oracle connections (`ODBC;DSN=...` without `DRIVER=`) still use `SELECT 1;` — same as before, not a regression. Registry lookup or reactive fallback can be added later if reported.

**Relevant files**:
- `modConnect.bas` — `IsOracleOdbcConnect`, `GetConnectivityProbeSql`, `CacheConnection`, `TestBackEndConnection`
- `clsDbConnection.cls` — pass `strErrDesc` from `CacheConnection` to `HandleConnectionFailure`
- `modTestConnect.bas` — unit tests for detection and probe SQL

---

## 2026-07-09 — Ship test-results/ gitignore to user projects

**Trigger**: Review of test-results persistence found the add-in repo's `.gitignore`
includes `test-results/`, but the shipped `.gitignore.default` template and
`CheckGitFiles` upgrade path did not. Existing user projects could commit durable
test artifacts (`test-state.json`, `test-results.xml`, `test-results.html`) after
running tests with default export options.

**Options explored**:
- **Document-only** — rejected: docs already claimed test-results were gitignored;
  users would still commit artifacts until they edited `.gitignore` manually.
- **Template + CheckGitFiles upgrade (chosen)** — mirror the existing `logs/` pattern:
  add `test-results/` to `.gitignore.default` and ensure it on export via
  `EnsureGitignoreLineRespectComment` (respects deliberately commented-out lines).

**Decision**: New projects get `test-results/` from the default template. Existing
projects with Git Integration pick it up on the next export when `CheckGitFiles` runs.
Already-committed files are not auto-untracked.

**What this rules out**: Auto-removing tracked `test-results/` from Git history; users
who committed those files still need `git rm -r --cached test-results/` once.

**Relevant files**: `.gitignore.default`, `modVCSUtility.bas` (`CheckGitFiles`).

---

## 2026-07-09 — TestResults JSON retention uses MaxLogFiles

**Trigger**: `clsTestRunner.SaveResults` pruned ephemeral `TestResults_*.json` dumps
with a hard-coded keep-10 constant and a private `PruneResultDumps` helper, while
`TestRun_*.log` files in the same `logs/` folder already honor `Options.MaxLogFiles`
via `clsLog.CleanupOldLogs`.

**Options explored**:
- **Separate `MaxTestResultFiles` option** — rejected: duplicates Advanced options UI
  and splits retention policy for artifacts that live in the same folder.
- **Parallel prune helper with MaxLogFiles** — rejected: duplicates `CleanupOldLogs`.
- **Reuse `Log.CleanupOldLogs` + `MaxLogFiles` (chosen)** — generalize the existing
  helper to accept a `Like` pattern; one code path and one knob for both
  `TestRun_*.log` and `TestResults_*.json`; `0` keeps all.

**Decision**: `SaveResults` calls `Log.CleanupOldLogs` with `"TestResults_*.json"`.
`CleanupOldLogs` is public, pattern-based, and sorts by name so oldest timestamped
files are deleted first. Enumeration uses `ScanFolderContents` (Win32) with VBA
`Like` filtering — same pattern as `GetMatchingFilePaths`. Durable `test-results/`
artifacts are unchanged.

**What this rules out**: Per-artifact retention counts or a second prune
implementation for test run history in `logs/`. Revisit only if JSON dumps and
console logs need different lifetimes.

**Relevant files**: `clsLog.cls` (`CleanupOldLogs`), `clsTestRunner.cls` (`SaveResults`)

---

## 2026-07-09 — Bridge run commands resolve at acceptance, not completion

**Trigger**: Any web-runner test run longer than 30 seconds toasted "Run failed: VBA call timed out: RunSelected" even while the run kept streaming results. The JS promise for `RunSelected`/`RunAll`/`RunFailed` was only resolved after the entire synchronous VBA run finished, but every `VBA.call` arms a 30 s timeout. A compile-error abort was worse: it skipped all stream events, leaving the page silent until the timeout fired.

**Options explored**:
- **Raise the JS timeout for run commands** — rejected: any fixed value just moves the false-failure threshold; a full suite with slow-tagged tests has no meaningful upper bound.
- **Suppress the timeout entirely for run commands** — rejected: a genuinely lost dispatch (stalled `RetrieveJavascriptValue`) would leave the UI waiting forever with no feedback.
- **Resolve at acceptance (chosen)** — VBA validates the request (`AcceptBridgeRun`: not already running, keys present, `Operation.Begin` succeeds), resolves the promise with `{"ok":true,"accepted":true}`, then executes the blocking run (`ExecutePendingBridgeRun`). Completion arrives via the streamed `onRunComplete` / `onRunCancelled` / new `onRunError` events the page already consumes. A JS watchdog (`RUN_START_WATCHDOG_MS`, 60 s) covers the pathological accepted-but-never-started case.

**Decision**: Run commands are ack-then-stream; request/response calls (`Cancel`, `RefreshTests`, `OpenTestSource`, `OpenResultsReport`, `CopyResultsPath`, `ReportJsError`) keep resolve-with-result under the 30 s `VBA_CALL_TIMEOUT_MS`. `TestRunner.InvokeGlobalTestSetup` (unbounded user hook) moved from the accept phase to the execute phase so the ack itself cannot stall. The compile-error abort in `RunSelected` now streams `onRunError` before bailing. Post-ack failures never reject the promise (it is already consumed) — they stream `onRunError` and guarantee `Operation.Finish` so the next run is not blocked.

**What this rules out**: The `.then()` of a run-command `VBA.call` no longer means "run finished" — UI state transitions must key off streamed events only. Any new long-running bridge command should follow the same accept/execute split rather than raising the shared timeout. Live per-assertion streaming (`onAssertionResult`) was also evaluated and removed: one `ExecuteJavascript` round-trip per assertion (suites run thousands) for no visible gain, since `onTestComplete` already carries the full assertion list; the per-assertion contract is `seq`/`passed`/`context` — VBA cannot capture call-site source text or line numbers through `Application.Run`.

**Relevant files**:
- `modTestRunnerUI.bas` — `IsRunCommand`, `AcceptBridgeRun`, `ExecutePendingBridgeRun`, `StreamRunError`; `BridgeRun*`/`ExecuteBridgeRun`/`BeginInteractiveBridgeRun` removed
- `frmVCSTestRunner.cls` — `DispatchRequest` branches run commands to ack-then-execute
- `clsTestRunner.cls` — compile-error branch streams `onRunError`; `HasFailedTests`
- `TestRunner/runner.html` — `dispatchRun` wrapper, `onRunError` handler, pending-button state, named timeout constants

---

## 2026-07-08 — Test runner HTML: repo-root `TestRunner/` packaging folder

**Trigger**: `runner.html` lived under `Version Control.accda.src/TestRunner/`, which
reads like an exported Access object folder but is actually an embedded packaging
asset (like `Ribbon.xml`). A dedicated repo-root folder also leaves room for future
test HTML (e.g. `test-results.html`).

**Options explored**:
- **Keep under `.accda.src/`** — works but misleads contributors; couples HTML to the
  export tree.
- **Install-folder extraction like `Ribbon.xml`** — rejected: no external consumer
  reads the file from `App.Path`; the Edge control needs a space-free temp path for
  `https://msaccess/` navigation anyway.
- **Repo-root `TestRunner/` + build-time embed + temp-cache runtime (chosen)** —
  same build-time `modResource.VerifyResources` path as `Ribbon.xml`; runtime
  extraction stays in `GetTempFolder` via `ResolveRunnerHtmlPath`.

**Decision**: Author `TestRunner/runner.html` at the repository root (parallel to
`Ribbon/` and `Hook/`). Embed with `VerifyResource "Test Runner HTML",
"\TestRunner\runner.html"`. Dev live-edit copies from `CodeProject.Path\TestRunner\`.
Reserve the folder for additional embedded test HTML assets.

**What this rules out**: Treating `TestRunner/` as part of the Access export tree;
install-folder deployment of runner HTML.

**Relevant files**: `TestRunner/runner.html`, `modResource.bas`, `modTestRunnerUI.bas`,
`AGENTS.md`, `.cursor/rules/testing.mdc`.

---

## 2026-07-08 — Standalone HTML test-results report (embedded snapshot)

**Trigger**: Users need a shareable, double-clickable test results view outside
Access (after closing the database, on builds without the Edge web runner, or for
emailing/archiving). Dynamic `fetch()` of sibling `test-state.json` fails under
`file://` due to browser CORS restrictions.

**Options explored**:
- **JS sidecar loaded via `<script src>`** — works on `file://` but still two
  files; awkward when emailing a single artifact.
- **Dynamic fetch of `test-state.json` / `test-results.xml`** — rejected for
  double-click use: blocked on `file://` in modern browsers.
- **Self-contained HTML with inlined JSON snapshot (chosen)** — add-in reads
  `test-state.json`, escapes `<` as `\u003c`, injects into embedded
  `TestRunner/results.html` template, writes `test-results/test-results.html`.

**Decision**:
- Report scope mirrors durable merged state (fresh + stale + pending), not just
  the last run — consistent with `test-state.json` and `test-results.xml`.
- Template at repo-root `TestRunner/results.html`, embedded via
  `VerifyResource "Test Results HTML", "\TestRunner\results.html"`.
- `modTestReport.ExportResultsHtml` called from `modTestState.PersistAfterRun`
  when `Options.ExportTestResultsHtml` is on (default); on-demand via
  `VCS.ExportTestResultsHtml`.
- Surfaced via plain console log path after generation, **Open Test Results...**
  on `frmVCSMain` after console test runs, and **Open report** toolbar button in
  the web runner (`OpenResultsReport` bridge → `FollowHyperlink`).

**Relevant files**: `TestRunner/results.html`, `modTestReport.bas`, `modTestState.bas`,
`modResource.bas`, `modTestRunnerUI.bas`, `clsOptions.cls`, `clsVersionControl.cls`,
`frmVCSMain`, `frmVCSOptionsAdvanced`, `AGENTS.md`.

---

## 2026-07-08 — Test results persistence: three-tier artifacts, durable state, JUnit export

**Trigger**: Need to reload the last test run when reopening the web runner after an
Access restart or VBA state reset; need a single file reflecting current status
across full and partial runs; need JUnit XML for CI without committing noisy
per-run artifacts.

**Options explored**:
- **JUnit XML as the persistence format** — rejected: JUnit lacks per-assertion
  detail, tags, logged errors, and source location fields the web runner needs.
- **Keep everything in `logs/`** — rejected: `logs/` is wholesale gitignored and
  semantically ephemeral/timestamped; durable state and CI export are different
  lifecycle tiers.
- **Dedicated `test-results/` folder with custom JSON + JUnit projection (chosen)** —
  three tiers: ephemeral per-run history in `logs/`, durable merged state in
  `test-results/test-state.json`, JUnit XML as an optional projection of state.

**Decision**:
- `modTestState.MergeAndSave` merges each run into `test-state.json`: executed
  tests get a fresh `lastRunAt`; non-executed tests keep prior status and are
  flagged `stale`.
- `modTestJUnit.ExportFromState` projects state to `test-results.xml` (on by default
  via `Options.ExportTestResultsJUnit`; regeneratable via `VCS.ExportTestResultsJUnit`).
- `RehydrateWebRunner` loads from disk when the singleton is empty.
- Both `test-results/` artifacts are gitignored (local working state / CI input).

**Relevant files**: `modTestState.bas`, `modTestJUnit.bas`, `clsTestRunner.cls`,
`modTestRunnerUI.bas`, `clsVersionControl.cls`, `clsOptions.cls`,
`frmVCSOptionsAdvanced`, `.gitignore`, `AGENTS.md`.

---

## 2026-07-09 — Web test runner: folder run scope and VCS.RunTests-style filters

**Trigger**: UX feedback — folder headers only expand/collapse with no way to run all
tests under a folder; exclusions like `-slow` work in `VCS.RunTests` but not in the
web UI; the primary Run button should respect composed navigation + filter scope.

**Decision**:

- **Folder select + run**: chevron toggles collapse; clicking the folder name selects
  that `@Folder` path (toggle off to clear). Descendant suites are included (e.g.
  `Tests` includes `Tests.SQL`). A ▶ on the folder header runs that folder immediately
  without changing selection. Primary Run label becomes **Run folder (N)**.
- **Tag include/exclude**: sidebar Tags cycle off → include → exclude (`-tag`) → off.
  Multiple tags compose like `ResolveFilters` (includes OR, exclusions AND-subtract).
- **Filter text box**: single sidebar input (no separate suite search) accepts the same
  token syntax as `VCS.RunTests` (module, folder/suite, procedure, tag; `-` prefix
  excludes). Composes with folder/suite/failed navigation base; drives the main list,
  Run scope, and tag chips. Sidebar tree is navigation-only (click folder/suite/tag).
- **Run scope**: `getVisibleTestKeys()` resolves navigation base then filter tokens;
  primary Run always calls `RunSelected` with that key set.
- **DefaultTestFilter seed**: when opening via `VCS.RunTests(...)` or ribbon
  `RunFilteredTests`, non-empty filters are passed through `setContext.defaultFilter`
  and prefill the filter box (editable, not auto-run).
- **Recent snapshots**: Recent entries store `{folder, suite, filterText}` (not a single
  kind/key) so combinations like `Tests.SQL` + `-slow` restore together. Legacy
  `{kind,key}` entries are migrated on load. Filter-box typing persists on blur/Enter
  (not every keystroke).
- **No Operation on web open**: `ExecuteTests` no longer calls `Operation.Begin` before
  opening the web runner (bridge runs own the lifecycle). `ClearOrphanedTestOperation`
  finishes a leftover `eotTestRun` when reopening/hiding if no test is actually running,
  fixing "Another Operation Already Running" after hide-and-reopen.
- **Access build probe**: `GetAccessFileBuild` uses `FSO.GetFileVersion` on
  `MSACCESS.EXE` instead of WMI `CIM_DataFile` (WMI often raised a one-shot Automation
  error dialog on first ribbon open despite fail-open support detection).

**Relevant files**: `TestRunner/runner.html`, `modTestRunnerUI.bas`,
`clsVersionControl.cls`, `frmVCSTestRunner.cls`, `AGENTS.md`.

---

## 2026-07-08 — Web test runner: status filters, clear filter, nested folders, cancel poll, focus restore

**Trigger**: UX feedback — stats-bar status counts should filter the list; a single
control should clear any active filter; `@Folder` paths should nest in the sidebar
(not flat dotted headers); **Stop** had no effect mid-run; pre-run compile stole
focus to the VBE when it was already open.

**Decision**:

- **Status filters**: passed / failed / errored / skipped counts in the stats bar
  are clickable (toggle off to All). Reuses `#test-list[data-filter]` CSS; failed
  includes assertion failures and runtime errors; errored is errors only.
- **Clear filter**: a **Clear filter** chip (and Esc when a filter is active)
  resets suite/tag/failed focus and status filter via `clearAllFilters()`.
- **Nested folders**: sidebar splits `suite.folder` on `.` into nested
  `.folder-group` nodes; full path strings in headers/tooltips use `/`.
- **Cancel poll**: `Form_Timer` cannot re-enter while a run started from
  `DrainOutbox` is still on the stack, so Cancel sat in `__vbaOutbox` until the
  suite finished. `RunSelected` now calls `PollBridgeCancel` after each `DoEvents`;
  the form exposes `DrainCancelOutbox` to splice Cancel commands without disturbing
  other queued bridge calls. `BridgeCancel` sets `CancelRequested` only (no eager
  `StreamRunCancelled`); the run loop streams cancelled when it actually exits.
  Cooperative cancel: the in-flight test still finishes.
- **Stop button window (2026-07-09)**: `#btn-cancel` is shown and enabled for the
  full in-flight window (`pendingRun` through `running`), not only after
  `onRunStart`. `onRunComplete` resets `disabled` so a prior cancel cannot leave
  Stop greyed out on the next run. `BridgeCancel` honors cancel whenever a test
  `Operation` is active (covers `GlobalTestSetup` / compile before
  `etrsRunning`); the form timer sets `CancelRequested`, and `RunSelected`
  aborts before `StreamRunStart` when that flag is already set.
- **Focus restore**: after `acCmdCompileAllModules`, and again at the end of
  `EndInteractiveBridgeRun` (after teardown / `ActiveVBProject` switches that can
  re-steal focus), `RefocusWebRunner` calls `BringAccessToForeground` and
  `ShowRunner` when the web runner is active so the run finishes on the form.
- **Default window size**: form design size raised from ~1320×768 to ~1600×900
  (24000×13500 twips) so the main column is usable on 1080p with sidebar + detail
  panes open; users can still resize.

**Relevant files**: `TestRunner/runner.html`, `frmVCSTestRunner.cls`,
`frmVCSTestRunner.form`, `modTestRunnerUI.bas`, `clsTestRunner.cls`,
`modVCSUtility.bas`, `AGENTS.md`.

---

## 2026-07-08 — Web test runner: run scope matches visible selection

**Trigger**: UX feedback — the primary **Run all** button did not match what the
list showed when a suite or tag was selected (suite focus ran the whole project;
only tag focus was scoped). Progress bar, header, and completion toast used
project-wide totals after partial runs (e.g. re-running one test toasts "All 10
tests passed" after a prior suite run).

**Decision**: **What you see is what the primary Run button runs.**

- Primary Run resolves `getVisibleTestKeys()` (tag/failed focus → those keys;
  suite selected → that suite; otherwise all tests) and always calls
  `RunSelected` — never `RunAll` with a cleared selection.
- Button label is dynamic: **Run all (N)**, **Run suite (N)**, **Run tag (N)**,
  **Run failed (N)**.
- Header always shows project total; appends `showing N (scope)` when filtered;
  last-run summary uses outcomes from the most recent run only (`lastRunOutcome`).
- Progress bar during a run sizes to `runKeys.length`; idle bar remains
  project-wide health. Duration shows `completed / total · elapsed` while running.
- Completion toast counts only tests in the current run.

**Relevant files**: `TestRunner/runner.html`, `AGENTS.md`.

---

## 2026-07-08 — Web test runner: hide-on-close, no auto-run, Escape key

> **⚠ Partially superseded** (2026-07-09): The `BridgeRunAll`/`BridgeRunSelected`/
> `BridgeRunFailed` callbacks were replaced by the `AcceptBridgeRun` /
> `ExecutePendingBridgeRun` split so the JS promise resolves at acceptance rather
> than after the blocking run. The operation lifecycle is still owned by the
> bridge run path (not by `ExecuteTests`), as decided here. See "Bridge run
> commands resolve at acceptance, not completion" above.

**Trigger**: UX feedback — closing the test runner form with the X button destroys
the WebView2 control (expensive cold-start on reopen), and the web runner was
auto-running tests immediately on open rather than letting the user click Run.

**Options explored**:
- **Hide vs unload** — considered keeping the old unload-on-X behavior. The
  WebView2 control's cold start takes 5–15+ seconds on first init; hiding the
  form keeps it warm and instant on reopen. The trade-off is a hidden form using
  some memory, but the WebView2 process is small and the warm control is worth it.
- **QueryClose vs Unload** — Access forms do not expose `Form_QueryClose` with
  `CloseMode` the way Excel/UserForms do. Used `Form_Unload` with an
  `m_blnAllowClose` flag instead: user X-click and Escape always cancel the unload
  and call `HideRunner`; programmatic `CloseWebTestRunner` sets `AllowClose = True`
  before `DoCmd.Close`, allowing the real unload to proceed for Access shutdown.
- **Auto-run** — the previous `ExecuteTests` scanned, opened the web runner, and
  immediately ran all tests in the same call. This made the UI feel like it was
  doing something behind the scenes before the user could see what was happening.
  Now `ExecuteTests` (web runner path) finishes the operation early after publishing
  the test tree; bridge callbacks (`BridgeRunAll`, `BridgeRunSelected`,
  `BridgeRunFailed`) handle the full operation lifecycle when the user clicks Run.

**What this rules out**: The web runner form cannot be programmatically unloaded
without setting `AllowClose` first — any new code paths that need to close it
must go through `CloseWebTestRunner` or set the flag manually. If Access shutdown
hangs because the flag isn't set, that's a bug (would need an emergency fallback
in `Form_Unload` based on `Application.Quit` detection).

**Relevant files**: `frmVCSTestRunner.cls`, `frmVCSTestRunner.form`,
`modTestRunnerUI.bas`, `clsVersionControl.cls`, `AGENTS.md`.

---

## 2026-07-08 — Web test runner: option toggle, pop-up host, quiet mode, state rehydration

**Trigger**: Follow-up polish on the web test runner. Requests: don't echo results
to the Immediate window when the HTML runner is showing them; make the runner
opt-out-able (legacy console fallback); make the Edge control fill the window and
open as a stand-alone window rather than a docked tab; preserve results so the
runner can be reopened to see/re-run the failed set; add an "All tests" / "Failed
tests (N)" focus affordance and surface assertion totals.

**Options explored**:
- **Suppress console noise** — considered skipping the per-test `Log.Add` calls in
  `clsTestRunner` (would also drop them from the TestRun log file). Chose instead a
  `Log.SuppressDebugOutput` flag that only gates the `Debug.Print` echo in
  `clsLog.Add` (fires when no GUI console is bound, i.e. the web-runner case). The
  log file and console buffer are unaffected. Set/cleared around the run in
  `ExecuteTests`.
- **Fill the form** — the Edge control has no design-time anchoring, so a
  `Form_Resize` handler sizes it to `InsideWidth/InsideHeight` (called once from
  `Form_Load` too). `OnResize` is a standard form event token (unlike the Edge
  control's `OnBeforeNavigation`/`OnDocumentComplete`, which had to be wired via
  the Property Sheet), so hand-authoring it in the `.form` imported cleanly.
- **Stand-alone window** — set the form `PopUp=1` so it opens as an overlapping
  window instead of a tabbed document. Cheapest reliable way to "pop out".
- **State persistence** — rejected keeping the form alive across close (reverted
  earlier: it blocks Access shutdown). Chose to rehydrate on open from the
  `clsTestRunner` singleton's `this.Tests` (tree + per-test `StreamTestComplete`),
  driven by a new standalone open path (`VCS.OpenTestRunner` → `m_blnStandalone`
  → rehydrate in `NotifyDocumentReady`). In-memory only; a VBA state reset clears
  it. A file-based rehydrate (from the saved `TestResults_*.json`) is the escalation
  if cross-restart persistence is needed.
- **Assertion numbering** — followed the PHPUnit "Tests: N, Assertions: N" convention
  (added `tests` and `assertions` totals to the stats bar) rather than inventing a
  coverage metric.

**Decision**: Gate the whole web-runner routing behind `Options.UseWebTestRunner`
(default True, Advanced options → Automated Testing). Keep the runner a pop-up whose
Edge control fills the window; suppress Immediate-window echo during web runs; and
rehydrate last results on a view-only reopen.

**What this rules out**: Results do not survive a VBA state reset or Access restart
(would need the file-based rehydrate). The runner is a single shared pop-up, not
multiple concurrent windows.

**Relevant files**: `clsOptions.cls` (option), `clsVersionControl.cls`
(`ExecuteTests` routing, `OpenTestRunner`), `clsLog.cls` (`SuppressDebugOutput`),
`modTestRunnerUI.bas` (`OpenTestRunnerForResults`, `RehydrateWebRunner`),
`frmVCSTestRunner.form`/`.cls` (`PopUp`, `Form_Resize`), `frmVCSOptionsAdvanced.*`
(checkbox), `TestRunner/runner.html` (Failed-tests focus entry, assertion totals,
focus-preserving filter).

---

## 2026-07-08 — Web test runner: warm reuse without page reload + merge-scan refresh

**Trigger**: Reopening the hidden test runner briefly showed prior run data, then
wiped it a few seconds later. Root cause: every reuse path called `ReloadRunnerHtml`
(full WebView2 navigate), and `RunTests` called destructive `Scan` before open.

**Options explored**:
- **Always reload on reuse** — rejected. Defeats hide-to-keep-warm; the multi-second
  navigate is what users perceived as a mysterious refresh.
- **Never rescan on reuse** — rejected. New/renamed/retagged tests would be stale in
  the sidebar until Access restart.
- **Destructive `Scan` on every open** — rejected for warm reuse. Wipes in-memory
  pass/fail even when the page is still loaded.
- **Show warm page + deferred `ScanMergingPriorResults`** — chosen. `ShowRunner`
  paints cached UI immediately; `DoEvents` then merge-scan republishes the tree
  only (`onReady` leaves `state.results` intact). `ReloadRunnerHtml` is fallback
  when `window.TestUI` is missing. Forced reloads replay completed results via
  `StreamCompletedTestResults` (not standalone-only).
- **Manual Refresh** — toolbar button + `RefreshTests` bridge callback runs the same
  merge-scan path for agents/users who added tests while the runner stayed open.

**Decision**: Warm reuse skips navigate when the document is healthy. Tree refresh
uses merge-scan; result rehydrate streams after any page wipe. `VCS.RunTests` no
longer calls blocking `Scan` before open — discovery is owned by the deferred refresh.

**Relevant files**: `modTestRunnerUI.bas` (`ReuseOrReloadRunner`,
`RefreshWebTestTreeDeferred`, `BridgeRefreshTests`), `clsTestRunner.cls`
(`ScanMergingPriorResults`), `clsVersionControl.cls` (`ExecuteTests` web path),
`frmVCSTestRunner.cls` (`RetrieveRunnerJsValue`), `TestRunner/runner.html`
(Refresh button).

---

## 2026-07-07 — Web test runner: EdgeBrowserControl + BeforeNavigate bridge

**Trigger**: Add a modern HTML/JS test-runner UI on Microsoft 365 while keeping
the existing `frmVCSMain` log console as the fallback on older Access builds.

**Options explored**:
- **Classic WebBrowser (IE) control**: rejected. Cannot run Promise-based bridge
  code; CSS/JS would require emulation downgrades.
- **Rebuild handoff VCSTest framework**: rejected. `clsTestRunner`, `TestAssert`,
  and `VCS.RunTests` already exist and are mature; only the UI shell was missing.
- **JS→VBA via `BeforeNavigate` + `vba://` iframe (primary)**: chosen. JS
  `VBA.call()` navigates a hidden iframe; VBA cancels navigation, pulls JSON
  payload via `RetrieveJavascriptValue`, replies with `ExecuteJavascript`.
- **Timer polling `__vbaOutbox` via `RetrieveJavascriptValue` (fallback)**: kept
  as opt-in via `EnableBridgePollingFallback` if `BeforeNavigate` interception
  fails on a given build.

**Decision**: Ship `frmVCSTestRunner` with a late-bound `webTestRunner` Edge
control (`As Object` everywhere), host `TestRunner/runner.html` via
`https://msaccess/` (space-free path, copy-to-cache when needed), stream run
events from `clsTestRunner` through `modTestRunnerUI`, and route
`ExecuteTests` to the web UI when `EdgeTestRunnerSupported()` (Access file
build ≥ 16327 / M365 2304+). Allowlisted inbound callbacks only:
`RunAll`, `RunSelected`, `RunFailed`, `Cancel`, `OpenTestSource`.

**What this rules out**: Early-bound `As Access.Edge` references (breaks compile
on older Access); `Application.Run` with page-supplied procedure names;
`RetrieveJavascriptValue` on the streaming hot path; static HTML report files
as a third UI mode.

**Relevant files**: `frmVCSTestRunner.*`, `modTestRunnerUI.bas`,
`clsTestRunner.cls`, `clsVersionControl.cls`, `TestRunner/runner.html`.

> **⚠ Implementation note** (2026-07-07): The Edge control's event property
> tokens are **`OnBeforeNavigation`** and **`OnDocumentComplete`** — note the
> first is `OnBeforeNavigation`, *not* `OnBeforeNavigate`. Using the wrong name
> is what makes `LoadFromText` fail ("This property does not apply to this
> control"); with the correct tokens the events round-trip through the `.form`
> normally. Wire them in the Access designer (Property Sheet → Event; Code
> Builder on this control crashes Access) and export. The event-handler subs are
> `webTestRunner_BeforeNavigate(Cancel As Integer, URL As String)` and
> `webTestRunner_DocumentComplete(URL As Variant)` — Access normalizes `Cancel`
> to `Integer` (not `Boolean`). `DocumentComplete` drives readiness and
> `BeforeNavigate` is the primary inbound bridge. `RetrieveJavascriptValue` is
> latency-prone and must NOT be polled continuously (an early
> `document.readyState` poll timed out and raised a modal warning), so there is
> no automatic readyState/outbox polling — the timer is off unless
> `EnableBridgePollingFallback` is called (last-resort `__vbaOutbox` drain). A
> request-id dedup guard ensures a call delivered via both channels runs only
> once. The exported `.form` must retain the `CodeBehindForm` marker so
> `MergeVBA` can splice the `.cls` on import.

> **⚠ Superseded** (2026-07-08): Authoring path moved to repo-root `TestRunner/`
> (see "Test runner HTML: repo-root `TestRunner/` packaging folder" above).
> Build-time embed and temp-cache runtime are unchanged.
>
> **⚠ HTML delivery** (2026-07-07): `runner.html` lives in the source tree
> (`Version Control.accda.src\TestRunner\`) but is NOT co-located with the
> compiled/installed add-in, so `CodeProject.Path\TestRunner\...` does not exist
> at runtime (grey/blank control). It is delivered like `Ribbon.xml`: embedded
> in `tblResources` via `modResource.VerifyResources` at build time and extracted
> at runtime with `ExtractResource` into a per-session temp folder
> (`GetTempFolder`). A dev fallback copies from the source tree, and re-copies
> when the source is newer (live HTML edits without a rebuild). The extraction
> folder must be space-free — the Edge control silently fails to load
> `https://msaccess/` URLs containing spaces — so `ResolveRunnerNavigateUrl`
> converts any path with a space to its 8.3 short form via `GetShortPath`
> (`GetShortPathNameW`, added to `modFileWinAPI`).

> **⚠ Final inbound decision: outbox polling** (2026-07-08): After exhaustively
> testing every navigation-based inbound signal (captured in the diagnostic
> trace), the JS->VBA bridge uses **lightweight polling**, not events:
> - JS `VBA.call()` enqueues `{id, fn, params}` in `window.__vbaOutbox` and does
>   NOT navigate.
> - The form timer (`POLL_INTERVAL_MS = 500`) drains the outbox via one
>   `RetrieveJavascriptValue` and dispatches each command; a `Cancel` is
>   dispatched even mid-run (the timer re-enters during the run's `DoEvents`),
>   other commands only when idle; request-id dedup prevents double-runs.
> - VBA->JS (streaming + `__vbaResolve`/`__vbaReject`) stays on `ExecuteJavascript`.
>
> Why not event-driven (the ruling-out, all confirmed by `beforenavigate.raw`
> logging): (a) `vba://` custom scheme — WebView2 swallows it, no `BeforeNavigate`;
> (b) hidden **iframe** to any URL — sub-frame navigations never reach the Access
> control's `BeforeNavigate`; (c) **top-level** navigation to `https://msaccess/` —
> `BeforeNavigate` fires, but `Cancel=True` does not cleanly abort a scripted
> main-frame navigation, so the control tries to load the URL, times out (~5-7s),
> and reloads the whole page. Polling is the only inbound path that never reloads
> the page. Also note: `RetrieveJavascriptValue` must be called only from the form
> timer / VBA flow, never from inside a control navigation event (it times out
> there). `DocumentComplete` is still used for readiness + cold-start re-navigate.
> The historical event-driven analysis below is retained for context.

> **⚠ Superseded by diagnostics** (2026-07-07, same day): The diagnostic trace
> (`beforenavigate.raw` logging + JS breadcrumbs) resolved the bridge questions
> empirically, replacing the deferred-dispatch/polling machinery below:
> 1. **`BeforeNavigate` DOES fire** for main-frame navigations (logged for both
>    `about:blank` and the runner URL), so the inbound bridge is event-driven —
>    no timer, no polling. Dispatch happens inline in `BeforeNavigate`; the whole
>    timer/`m_strPendingFn`/`DrainOutbox`/`EnableBridgePollingFallback` layer was
>    removed.
> 2. The JS signal must be a **top-level navigation to an `https://msaccess/`
>    URL** (`window.location.href = 'https://msaccess/__vba__/<fn>/<id>'`), which
>    VBA cancels in `BeforeNavigate`. Two things that do NOT work, both confirmed
>    by the trace (no `beforenavigate.raw` line appears for them): (a) a hidden
>    **iframe** — WebView2 raises the navigation event for the main frame only,
>    not sub-frames; (b) a **custom `vba://` scheme** — WebView2 swallows/bounces
>    unknown-scheme navigations without raising `BeforeNavigate` (and the page
>    reloads). The `https://msaccess/` host is the same one the runner HTML is
>    served from, so `BeforeNavigate` fires for it reliably; the `/__vba__/` path
>    marker distinguishes a bridge signal from the page load.
> 3. **Cold-start grey screen**: the control's `ControlSource ="about:blank"`
>    loads `about:blank` when the (cold) WebView2 finishes initializing, AFTER
>    Form_Load's `Navigate` was lost — so about:blank won. Fix: `DocumentComplete`
>    only treats the `https://msaccess/...` runner URL as ready; any other landing
>    (about:blank or stray) triggers a re-`Navigate` to the runner. This doubles
>    as the reliable "control is now initialized" trigger. Event-driven, no timer.
>
> The original deferred-dispatch reasoning (kept below for history):

> **⚠ Deferred bridge dispatch** (2026-07-07): Page-initiated bridge commands
> (Run all / Rerun failed) must NOT run the test suite synchronously inside the
> `BeforeNavigate` handler. `ExecuteJavascript` calls issued while a navigation
> callback is executing (both the per-test streaming pushes and the final
> promise-resolve) do not run until the handler returns, so a full run started
> inline leaves the JS promise unresolved until it times out (~30s). Fix:
> `BeforeNavigate` handles only `Cancel` inline (fast, must interrupt a running
> suite) and defers every other command to the form timer via `m_strPendingFn`/
> `m_strPendingId`; the timer executes it OUTSIDE the callback (guarded by
> `m_blnDispatchBusy` against re-entrancy from the run's own `DoEvents`), where
> streaming and resolve work. The ribbon path is unaffected because it drives the
> run directly from VBA, not from a navigation event. GENERAL RULE (confirmed by
> the diagnostic trace: a `RetrieveJavascriptValue` call inside `DocumentComplete`
> logged `js.drain.timeout` and delayed readiness ~1s): never call the Edge
> control's JS methods (`ExecuteJavascript` OR `RetrieveJavascriptValue`) from
> inside its own `BeforeNavigate`/`DocumentComplete` event handlers — WebView2
> needs the message pump that the blocked callback is holding, so the call hits
> its internal timeout. Do such calls from the timer, run flow, or other
> VBA-driven context. Also: the readiness wait
> (`WaitForWebRunnerReady`) default was raised to 30s because a cold WebView2
> first-init can exceed 15s, which otherwise starts the run against a blank page.

> **⚠ Warm-singleton lifecycle** (2026-07-07): The WebView2 cold-start (first
> control init per Access session) is the dominant startup cost. The form is a warm
> singleton **only while open** — `OpenWebTestRunner` reuses it if already open (no
> cold-start), so back-to-back runs are instant as long as the window stays open.
>
> > **⚠ Superseded** (2026-07-07, same day): An earlier version HID the form on
> > close (`Form_Unload` set `Me.Visible = False` + `Cancel = True`) to keep the
> > control warm across closes. This **blocked Access from shutting down** — the
> > host's quit unloads forms, and the cancelled unload aborts the quit — and left
> > `msedgewebview2.exe` processes alive. There is no reliable way to distinguish a
> > user form-close from an app quit in `Form_Unload`, so hide-on-close was removed.
> > `Form_Unload` now always tears down cleanly (navigates the control to
> > `about:blank`, releases refs) so WebView2 exits with the form. Cost: closing the
> > window means the next run re-initializes (cold). Do NOT reintroduce
> > `Cancel = True` in this form's unload.

> **⚠ Diagnostic trace log** (2026-07-07): Debugging the bridge by round-tripping
> rebuilds is slow because the actual VBA↔JS flow is invisible. `modTestRunnerDiag`
> closes the loop: it writes a single agent-readable trace
> (`<ExportFolder>\logs\TestRunnerDiag.log`, truncated per session by `DiagStart`)
> interleaving VBA lifecycle/bridge events with JS breadcrumbs drained from
> `window.__diag` via a single `RetrieveJavascriptValue` at checkpoints
> (DocumentComplete, after each deferred dispatch). The `navigate.call` →
> `documentcomplete` gap reveals WebView2 load/cold-start time; `beforenavigate`
> presence proves the JS→VBA event bridge fired; `wait.ready`/`wait.timeout`,
> `resolve`/`reject`, and `push.dropped` pinpoint where a run stalls. Diagnostics
> must never perturb the observed flow: `Diag` no-ops when disabled and never
> raises.

---

## 2026-07-06 — Surgical VBE reset in RunVBA; rejected for merge (crashes)

> **⚠ Partially superseded** (2026-07-29): The reset is now used in the merge
> path behind the opt-in `Options.SkipReopenBeforeMerge`, in one of the two
> shapes this entry named as prerequisites — references released before the
> reset, and the merge resumed on a fresh stack via the timer. The rejection of
> a *bare* in-place reset in merge still stands. See "Opt-in in-place merge
> preparation instead of the pre-merge reopen" above.

**Trigger**: Agents driving `vcs_run_vba` (and repeated add/remove of modules
via MCP) intermittently hit the modal "This action will reset your project,
proceed anyway?" prompt, which blocks the automation thread until a human
clicks it. Root cause: modifying a VBA project's `VBComponents` (add/remove a
module) while that project holds in-memory run-state raises the prompt. It is
raised by the VBA/VBE engine, not the Access action layer, so
`DoCmd.SetWarnings`, `Application.Echo`, and `Application.SetOption` do not
suppress it, and there is no documented flag to turn it off — you can only
avoid triggering it or dismiss the modal after the fact.

**Options explored**:
- **`DoCmd.SetWarnings False`**: rejected. Only gates Access action-query/UI
  confirmations, not the VBE engine's project-reset prompt.
- **Auto-dismiss the modal (SendKeys / UI Automation watcher)**: not pursued.
  The dialog is modal and blocks the thread that raised it, so it needs a
  pre-queued keystroke or an external watcher; agent VBA is user-monitored, so
  forcing the dialog closed was deemed unnecessary.
- **`End` before running / before build**: rejected as primary. `End` is
  global — it resets every loaded project (including the add-in's singletons)
  and cannot "End then continue" in one call. Usable only as an external,
  between-calls action.
- **`DoCmd.RunCommand acCmdReset` (AcCommand 124)**: viable but its scope
  (global/abortive vs. active-project) was uncertain. Kept only as a
  documented fallback.
- **VBE Standard toolbar Reset control, resolved by language-independent ID
  228 (`Application.VBE.CommandBars.FindControl(, 228).Execute`)**: chosen for
  `RunVBA`. Confirmed (IDE and programmatically via MCP) to reset only the
  *active* project — never the library add-in. `RunVBA` sets
  `VBE.ActiveVBProject = CurrentVBProject` first, resets before running agent
  code and again after removing the temp module. Validated with an
  `OptionsLoaded` sentinel (add-in `Options` singleton stayed loaded) across
  basic, DB-access, and runtime-error probes, plus repeated add/remove cycles.
- **Surfaced `McpResetProjectForRunVBA` option to toggle the reset**: added,
  then removed (YAGNI) — the reset is now unconditional in `RunVBA`.
- **Reuse the same reset for merge, replacing `CloseCurrentDatabase2` +
  `ShiftOpenDatabase` with `CloseDatabaseObjects()` + reset (shift-reopen kept
  as fallback)**: **rejected — crashes Access.** Plain version and a variant
  with `DoEvents` between close and reset both crashed on merge in the Testing
  database (full build was unaffected). Cause: unlike `RunVBA` (reset → run one
  function → return), merge continues heavy, sustained work against the same
  project after the reset (deleting/re-importing many `VBComponents`,
  `DoCmd.Close/Save`, DAO refreshes) while holding cached references
  (`CurrentVBProject`/`VBComponents`, `SharedDb`, DAO handles) that the reset
  invalidates. Merge reverted to the stable shift-reopen.

**Decision**: Use the VBE Reset control (id 228) via
`modVbeUtility.ResetCurrentVBProjectState()` in `clsVersionControl.RunVBA`
only. Merge and full build remain on the close/shift-reopen path unchanged.

**What this rules out**: Do not drop an in-place VBE reset into the merge (or
build) sequence — it corrupts the long-lived cached references those flows
depend on and hard-crashes Access. Revisiting the merge performance win
(avoiding the physical close/reopen) requires a different shape: either
re-acquire all target-DB/VBE handles after the reset, or stage → reset/`End` →
resume on a fresh stack via the existing `SetTimer`/`APIAsyncOperation`
pattern. `acCmdReset` and the surfaced toggle remain deliberately unused.

**Relevant files**: `Version Control.accda.src/modules/Core/modVbeUtility.bas`
(new `ResetCurrentVBProjectState`), `Version
Control.accda.src/modules/API/clsVersionControl.cls` (`RunVBA` calls it before
run and after temp-module removal). Reverted: `Version
Control.accda.src/modules/Core/modBuild.bas` (merge stays on shift-reopen).
`Version Control.accda.src/modules/Infrastructure/clsOptions.cls` unchanged
(surfaced option added then removed).

---

## 2026-06-26 — Table-less Design View SELECT: parse field list and emit empty InputTables

**Trigger**: `qryFormControl` in `Testing.accdb` lost its output columns on
build-from-source + re-export. The query is a table-less scalar SELECT
(form-control references, no `FROM`) last saved in Design View. After import
via the native `.sql`/`.json` pipeline, MSysQueries had Attribute 0/1/3/4
rows but no Attribute 6 output-column rows; export faithfully reconstructed
`SELECT FROM ;`.

**Options explored**:
- **Force SQL View for table-less SELECTs** — guaranteed lossless (raw SQL
  stored verbatim) but drops Design View fidelity and the trivial layout
  block; rejected because the user chose to preserve Design View.
- **Parser fix only** — necessary but insufficient; `EmitDesignViewQdef`
  skipped the empty `InputTables` block when `m_colInputTables.Count = 0`,
  and Access `LoadFromText` silently dropped `OutputColumns` without it.
- **Parser fix + always emit empty InputTables block** — matches Access
  `SaveAsText` output for Design View queries; chosen.

**Decision**: `ParseSelectQuery` gains an `Else` branch (mirroring
`ParseInsertQuery`) that calls `ParseFieldList` when no `FROM` keyword is
present. `EmitDesignViewQdef` always writes `Begin InputTables` / `End`,
even when the table collection is empty.

**What this rules out**: Assuming SQL View import masks parser gaps for
table-less queries — `qryRegressionScalarNoTable` passes on SQL View only;
Design View requires the separate fixture
`qryRegressionScalarNoTableDesignView`. If `LoadFromText` still drops
columns after both fixes, fall back to forcing SQL View for this shape.

**Relevant files**: `clsQueryComposer.cls` (`ParseSelectQuery`,
`EmitDesignViewQdef`); fixture
`Testing/Fixtures/queries/regression/qryRegressionScalarNoTableDesignView.*`;
`docs/access-query-storage.md`.

---

## 2026-06-26 — CF companion JSON colors as RGB(R,G,B) strings

**Trigger**: Conditional formatting rules in the companion `.json` stored colors as raw
Access Long integers (e.g. `16777215` for white). These values are opaque to humans and
AI agents editing source files; diffs are meaningless without knowing the
`R + G*256 + B*65536` formula.

**Options explored**:
- *Keep numeric Longs* — lossless and trivial, but poor readability. Rejected for the JSON
  layer (binary blocks still use Longs unchanged).
- *Hex strings (`#RRGGBB`)* — familiar on the web, but not VBA-native syntax.
- *RGB(R,G,B) strings (chosen)* — matches VBA's built-in `RGB()` function; agents and
  developers can read and edit colors directly.

**Decision**: At the JSON dictionary boundary in `clsConditionalFormat`, export converts
Long color values to `"RGB(R,G,B)"` strings. Access flattens automatic and theme colors to
literal RGB at save time, so CF binary blocks and JSON values are always in `0..16777215`
with no sentinel. Import parses `RGB(...)` strings and legacy numeric Longs for backward
compatibility. No new export format version gate — CF decode-to-JSON has not shipped in a
release yet.

**Relevant files**: `modules/Core/clsConditionalFormat.cls` (`LongToRGB` / `RGBToLong`),
`modules/Tests/Core/modTestConditionalFormat.bas`, `docs/access-conditional-format.md`.

---

## 2026-06-25 — SharedDb invalidation during build/merge and database close

**Trigger**: Full build failed when importing table data: `clsDbTableData.ImportTableDataTDF`
accesses `SharedDb.TableDefs(strTable).Fields` after `clsDbTableDef` created tables via
`Application.ImportXML`. `IsLocalTable` passed (queries `MSysObjects` directly), but the
cached `CurrentDb` handle's `TableDefs` collection was stale. Same pattern affects
`QueryDefs`, `Containers.Documents`, and `Relations` when imports use `Application.*` APIs
instead of DAO on the cached reference.

**Options explored**:
- **`TableDefs.Refresh` on the cached handle**: Fixes one collection only; `QueryDefs` and
  `Containers.Documents` have the same problem. Rejected.
- **Allowlist of schema-changing component types** (`edbTableDef`, `edbQuery`, etc.):
  Fragile — most object-creating categories use `Application.LoadFromText` or `ImportXML`;
  future component types would need manual updates. Rejected.
- **Unconditional `ReleaseDbReferences` after every import category** (chosen): One
  `Set this.dbs = Nothing` per category in `modBuild.Build`; next `SharedDb()` obtains
  fresh `CurrentDb`. Negligible cost; JET page cache lives at engine level, not on the
  DAO handle.
- **Also call `ReleaseDbReferences` inside `CloseCurrentDatabase2`**: Covers all close/reopen
  paths (full build, merge pre-reopen, theme reopen, shared-mode reopen). Explicit call
  before shared-mode reopen in `modBuild` kept as belt-and-suspenders.

**Decision**: `ReleaseDbReferences` runs unconditionally after each category in the build
loop and at the start of `CloseCurrentDatabase2` (before `WizHook.CloseCurrentDatabase`).

**What this rules out**: Per-component `TableDefs.Refresh` / collection refresh calls as
the primary fix. Do not rely on `SharedDb` seeing objects created via `Application.*`
import APIs within the same category without invalidating between categories. Export perf
optimization via `SharedDb` is unchanged — invalidation applies only during build/merge and
close/reopen, not export.

**Relevant files**: `Version Control.accda.src/modules/Infrastructure/modObjects.bas`
(`ReleaseDbReferences`, `SharedDb`), `Version Control.accda.src/modules/Utility/modWizHook.bas`
(`CloseCurrentDatabase2`), `Version Control.accda.src/modules/Core/modBuild.bas`.

---

## 2026-06-25 — Command bar built-in classification: runtime probe + template copy + replicas

**Trigger**: Menu round-trip work (issues #584, #583) exposed that Access "custom built-in"
controls — object-openers (Open Table/Query/Form/Report 1835-1838, ADP 3885/3886/3888),
Edit Hyperlink 3227, etc. — cannot be recreated with `CommandBarControls.Add(Type, Id)`.
PR #588 handled them by copying from a shipped binary template (`Template/CommandBars.bin`),
which was fragile when copying into popup parents and unavailable on consecutive add-in builds.
#584 also reported built-in controls losing their pictures and action parameters on round-trip
when exported as minimal built-ins.

**Options explored**:
- **Template copy only (PR #588)**: copy a fixed set of non-addable built-ins from a binary
  template bar. Rejected as sole strategy: hand-maintained Id list, type-mismatch when copying
  into popup parents, and no path for controls absent from the template.
- **Static whitelist ∩ ¬blacklist (tried, replaced)**: classify by `ControlIdToName`
  membership minus a hand-maintained `IsNonAddableControl` list (`IsBuiltInControlRecreatable`).
  Rejected: endless whack-a-mole as users hit Ids absent from our small sample, and Id-only
  keying is provably wrong — `Undo` Id 128 is non-addable as `Type=6` (split-button) but
  addable as `Type=1` (plain button).
- **Runtime addability probe + replicas only (tried, partially superseded)**: classify each
  built-in at export by attempting `Controls.Add(Type, Id)` on a temp bar, cached per
  `(Type, Id)`. Comprehensive for detection and #584 customized-control replicas, but
  replicas with empty `OnAction` are inert for intrinsic object-opener wiring.
- **Hybrid probe + template copy + replica fallback (chosen)**: probe for comprehensive
  addability detection; template copy on import for non-addable controls present on the
  shipped template bar (full intrinsic behavior); visual replica fallback when the control
  is non-addable and absent from the template.

**Decision**: `IsBuiltInControlAddable(Type, Id)` in `modCommandBarNames.bas` wraps
`ProbeBuiltInControlAddable` with a session-lifetime `(Type, Id)` cache. Export classification
in `clsDbCommandBar.BuildElementDictionary`: `IsNonAddableControl` manual override -> replica;
else probe addable + `IsBuiltInControlCustomized` (custom `OnAction`/`Parameter`, or pasted
face: FaceId=0 with a picture) -> replica (#584); else probe addable + clean -> minimal
`BuiltIn=true` block; else non-addable + `IsTemplateControl(Id)` -> full-property built-in
(`BuiltIn=true`, `Id`, all props); else -> replica. Import: `BuiltIn=true` -> `.Add(Type,Id)`;
on failure -> template copy by `Id` (popup-safe via `objParent.CommandBar`) + `BuildControls`;
on failure -> `eelError` (visible on build screen). `BuiltIn=false` -> custom build (replicas).
The template ships in `Template/CommandBars.bin`, baked into the add-in by
`ImportCommandBarsTemplate` in `AfterBuild`; `m_TemplateCommandBar` is looked up at class init
and skipped during export enumeration. When the template is not loaded (consecutive builds),
`IsTemplateControl` returns False and non-addable controls degrade to replicas until the
add-in is rebuilt. Gated at `EFV_5_0_0` (v5 unreleased).

**What this rules out / caveats**:
- ADP database diagrams (3887): unsupported even by the template (cannot be built in Access
  2003 per #588). Stays a documented limitation.
- Classification reflects the running Access version; a borderline control can serialize as a
  minimal built-in on one version and a replica on another (accepted).
- Custom controls added directly into an *addable* built-in popup's submenu are not preserved
  (native submenu regenerates on import). Extension path: `AddAfter` anchor for custom children.
- Performance: command bars have no timestamp, so `IDbComponent_IsModified` rebuilds the
  dictionary each scan; the `(Type, Id)` session cache keeps the probe to one lookup per
  built-in after warmup.

**Relevant files**:
- `modCommandBarNames.bas` — `IsBuiltInControlAddable`, `ResetAddableCache`,
  `IsNonAddableControl` (empty override), `ProbeBuiltInControlAddable`/`DumpNonAddableControls`
- `clsDbCommandBar.cls` — `BuildElementDictionary` three-way classification,
  `IsTemplateControl`, `IsBuiltInControlCustomized`, `AddChildControl` template copy,
  `m_TemplateCommandBar`
- `modConstants.bas` — `strTemplateCommandBarName`, `EFV_5_0_0`
- `modVCSUtility.bas` — `ImportCommandBarsTemplate` in `AfterBuild`
- `Template/CommandBars.bin` — shipped template instances
- `modTestCommandBarNames.bas` — probe and template-aware tests

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

---

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

---

## 2026-06-09 — Batch file metadata (date+size) for source property hashing

> **⚠ Partially superseded** (2026-07-29): The map is no longer built "once per
> category." A merge build now builds one map for the whole scan phase and shares it
> across every category, because several component types report the export root as
> their `BaseFolder` and each recursive walk therefore covered the entire tree. The
> "capture during enumeration" design was not needed to fix that. The DST caveat and
> its reliance on the content-hash fallback still hold — that fallback is now the
> second tier of a three-tier precedence. See "Merge scan reads no file content when
> dates and sizes are unchanged" above.

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

> **⚠ Partially superseded** (2026-07-29): This decision was reverted in
> `0e4b93b0` (the reopen is unconditional again). The `ReleaseScanState` helper
> it depended on now exists, and the pre-merge reopen can instead be skipped
> entirely via the opt-in `Options.SkipReopenBeforeMerge`. Deferring the reopen
> based on change count remains unimplemented. See "Opt-in in-place merge
> preparation instead of the pre-merge reopen" above.

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

*(Correction, 2026-07-29: the link to importing was coincidental. The post-merge reopen is
triggered purely by the engine lock state that `Worker.IsDatabaseAccessible` probes, which
makes no reference to what the merge did. A later run imported four objects without
triggering it, and another triggered it having imported nothing. See the 2026-07-29 in-place
merge entry.)*

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

## 2026-05-07 — Cross-table ON condition LeftTable/RightTable in Design View qdef

> **⚠ Partially superseded** (2026-08-10): The "fall back to the parent join's
> tables only if extraction returns empty" rule is incomplete — a non-empty
> garbage token from a function-call operand bypassed it. See
> "Function-call operands in ON clauses must resolve against InputTables" above.
> Per-condition extraction itself remains correct for simple column equalities.

**Trigger**: A production database had four queries that passed SQL builder validation but failed with DAO error 3082 ("JOIN operation refers to a field that is not in one of the joined tables") after a full build from source. The queries used compound `ON` clauses where individual conditions referenced different table pairs, and one table was also used inside a saved subquery referenced in another condition.

**Root cause**: `clsQueryComposer.EmitDesignViewQdef` reused the parent join's `leftTable`/`rightTable` for all split conditions in a compound `ON` clause. Access stores each compound `ON` condition as a separate Attribute 7 row in `MSysQueries` with its own `Name1`/`Name2` (the specific table pair for that condition). The emitter's reuse of the parent join's tables produced a `.qdef` where the `RightTable` for a condition referencing table `C` was set to table `B` (the parent join's right table). `LoadFromText` accepted this silently, but the resulting internal storage confused Access's scope resolution at execution time.

**Options explored**:
- **Fall back to `QueryDefs(name).SQL` (legacy path)** — rejected: the new pipeline was designed to generate its own `.qdef` rather than receive a pre-baked one, and falling back to the legacy path would lose design layout. The bug was in the emitter, not in `LoadFromText`.
- **Store per-condition table pairs in the `.json` companion** — rejected: the table pair for each condition is derivable from the condition expression itself (e.g., `tblCars.ID = tblCarsModel.CarID` clearly references `tblCars` and `tblCarsModel`). Adding explicit storage would be redundant.
- **Extract per-condition table pairs from the expression at emit time (chosen)** — the emitter already has `ExtractTableFromOnSide` available. Using it for each split condition, with a fallback to the parent join's tables if extraction fails, is correct, minimal, and preserves backward compatibility.

**Decision**: `EmitDesignViewQdef` now calls `ExtractTableFromOnSide(condition, True)` and `ExtractTableFromOnSide(condition, False)` for each individual condition in a split compound `ON` clause. Falls back to the parent join's `leftTable`/`rightTable` only if extraction returns empty.

**Why this was hard to diagnose**: The SQL builder validation compares `ReconstructSQL` output against `QueryDefs.SQL` — a text-level check. The bug was not in SQL reconstruction but in `.qdef` emission, and `LoadFromText` accepted the wrong structure silently. The error only surfaced at query execution time, where the misleading error message ("field not in one of the joined tables") pointed away from the actual root cause (wrong `LeftTable`/`RightTable` metadata).

**Relevant files**:
- `clsQueryComposer.cls` — `EmitDesignViewQdef`: per-condition `LeftTable`/`RightTable` extraction
- `docs/access-query-storage.md` § 6 — documents the finding
- `Testing/Fixtures/queries/regression/qryRegressionCrossTableOn.notes.md` — regression context

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

## 2026-05-01 — Pass-through queries bypass SQL formatter; SQL sourced from MSysQueries

**Trigger**: Exporting a database containing `dbQSQLPassThrough` queries crashed `clsSqlFormatter` with "Unable to parse SQL after position N" — the formatter's tokenizer is designed for Access SQL syntax and cannot handle T-SQL, PL/SQL, or other server-side dialects that pass-through queries may contain.

**Options explored**:
- **Teach the formatter about T-SQL/PL-SQL**: rejected. Scope explosion — every server dialect has its own syntax, reserved words, quoting rules, and comment styles. The formatter would become a multi-dialect parser with no clear boundary.
- **Format only the SELECT-like subset** (heuristic detection of "looks like Access SQL"): rejected. Fragile — any heuristic would produce false positives on server SQL that happens to resemble Access SQL, silently corrupting the stored query text.
- **Detect and bypass formatter; reconstruct SQL from MSysQueries** (chosen): Check `MSysObjects.Flags` for `dbQSQLPassThrough` (112) and `dbQSPTBulk` (144), or MSysQueries Attribute 1 Flag 8/10. `clsQueryComposer.ReconstructSQL` returns the verbatim Attribute 1 `Expression` for Flag 7 (DDL), 8 (pass-through, returns records), and 10 (pass-through, no records). Connect comes from Attribute 1 `Name1` (or Attribute 4 `Expression` when present). `clsSqlFormatter` is skipped for pass-through types.

**Decision**: Pass-through export reads SQL and connect from the system tables only (no `QueryDef.SQL` / `QueryDef.Connect` round-trip on the hot path). `ReturnsRecords` is not in `LvProp`; when Attribute 1 Flag = 10, export writes `QueryProperties.ReturnsRecords = false` to the `.json` companion (issue #724). Import generates a `.qdef` via `clsQueryComposer` and `LoadFromText`, which honors `dbBoolean "ReturnsRecords" ="0"`.

**What this rules out**: Future attempts to "fix" the formatter or composer for non-Access SQL dialects — pass-through SQL must always be stored verbatim from MSysQueries. If a future need arises to pretty-print server SQL (e.g. for diff readability), it must be a separate, opt-in formatter that does not share code paths with the Access SQL formatter.

**Relevant files**: `Version Control.accda.src/modules/Components/clsDbQuery.cls` (export/import), `Version Control.accda.src/modules/Utility/clsQueryComposer.cls` (`ReconstructSQL`, `ConnectString`, `EmitAllProperties`), `Testing/Fixtures/queries/passthrough/` (round-trip fixtures including `qryPassThroughNoRecords` for Flag 10).

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

**Trigger**: Export profiling on a large production database (~3,500 components) showed orphan scanning and file-extension migration checks dominated the "no changes" export time. Two separate problems: (1) `Dir()` does not support Unicode filenames — it silently skips or fails on paths containing non-ASCII characters, which Access databases frequently contain (accented characters, CJK object names). (2) `Scripting.FileSystemObject` (FSO) folder enumeration is correct but slow — each `oFolder.Files` / `oFolder.SubFolders` iteration creates COM proxy objects with per-item round-trip overhead.

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

**Trigger**: On a large production database (~3,500 component entries), the `vcs-index.json` file grew to 1.5MB / 40K lines. Parsing it via `modJsonConverter.ParseJson` consumed 1.5-2.2s per export — nearly half the total runtime for a no-changes export. The bottleneck was threefold: three `Replace()` calls stripping whitespace from a 1.5MB string, character-by-character recursive descent creating ~10,000 `Scripting.Dictionary` COM objects, and ~3,500 ISO 8601 date string parses.

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

**Trigger**: Post-`clsQueryComposer` work on the SQL/JSON query format surfaced ~723 affected queries in a single production database from a self-join alias bug (`qryCurrencyCrossRates` archetype). Manual repro-and-fix is unsustainable as more edge cases land. Traditional VBA unit testing (Rubberduck-style or hand-rolled) would require hundreds of fixture queries hard-coded into the add-in — thousands of lines of VBA permanently loaded into memory in every running instance, for code paths that are only exercised during development. A different shape was needed.

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

**Decision**: Implement `modTestRoundtrip.bas` inside the add-in with `Option Private Module` and `RunObjectRoundtripTests(Optional strFixtureFolder, Optional blnRebaseline)` as its single in-project entry point. Expose this externally through one public delegate, `clsVersionControl.RunRoundtripTests`, alongside the other dev/agent tools (`RunVBA`, `ExecuteSQL`, `CompileVBA`). External invocation: Immediate Window uses `?VCS.RunRoundtripTests`; MCP/CI uses `vcs_run_vba` with `MCP_TempFunction = VCS.RunRoundtripTests()` (gated by `McpAllowRunVBA`). Fixtures live in `Testing/Fixtures/<component>/<category>/` as plain text (`.sql` + `.json` for queries today; the slot is reserved for `forms/`, `reports/`, etc.) with a `_scaffold/` sibling folder for shared supporting objects loaded once per session. Each fixture runs through a two-pass round trip (import to `vcs_test_<name>_<hash>` sandbox, export, re-import, re-export) with three independent SHA-256 comparisons: Pass 1 vs. fixture, Pass 1 vs. Pass 2 (idempotency), JSON-with-`Info`-stripped both directions. Bloat is addressed structurally: random-suffix sandbox names allow parallel runs and unambiguous leftover detection, every fixture cleans up via `DoCmd.DeleteObject` + `DBEngine.Idle dbRefreshCache`, the run starts with a `CleanupStaleObjects` sweep over any `vcs_test_*` survivors from a crashed prior run, and `VCSIndex.Disabled = True` for the entire run prevents test operations from polluting `vcs-index.json`. Output flows through the existing `Log` singleton (live console in `frmVCSMain` + per-session `ObjectRoundtrip_<opId>.log` with full inline diffs) and a structured JSON return for programmatic parsing. Bug-as-fixture is the canonical contribution path: real-world failures from production validation or user reports are distilled into a fixture under `regression/` with a `.notes.md` companion documenting the failure mode and resolution status — `qryCurrencyCrossRates` is the seed entry, currently failing as expected.

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

> **⚠ Partially superseded** (2026-06-25): `ReleaseDbReferences` is now also called
> unconditionally after each import category in `modBuild.Build` and at the start of
> `CloseCurrentDatabase2` (not only before the shared-mode reopen). See
> "SharedDb invalidation during build/merge and database close" above.

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

**Trigger**: After implementing `@Folder` annotation support (EFV 5.0.0), export logs from a large production database showed "Clear Orphaned Files" consistently at 5-6 seconds on fast saves, even with zero modified objects. Root cause: `GetFolderAnnotation` reads the entire VBE code module via `cmpItem.CodeModule.Lines(1, 999999)` on every call, and `SourceFile` (which calls `GetFolderAnnotation`) was accessed multiple times per object per export — up to ~1,558 VBE COM calls for that database's 779 VBA-backed objects.

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

**Trigger**: During fast-save export, each `IDbComponent` class's `GetAllFromDB` was called twice per category: first with `blnModifiedOnly=True` (scan for changes), then with `blnModifiedOnly=False` (orphan detection via `ClearOrphanedSourceFiles`). Each call independently iterated the full Access collection and instantiated new `clsDb*` objects. Performance logs from a large production database (~412 forms, ~3694 queries, ~392 tables) showed "Clear Orphaned Files" consistently taking 5.2-6.0 seconds — pure waste from re-enumerating objects already visited during the scan phase. Combined with "Scan DB Objects" (6.2-28.3s), these two passes consumed 34-54% of total fast-save runtime.

**Options explored**:

- **Approach A — Single-loop dual-populate**: When `GetAllFromDB(True)` iterates the collection, always populate `m_Items(False)` (all items) alongside `m_Items(True)` (modified items). The subsequent `GetAllFromDB(False)` call from `ClearOrphanedSourceFiles` hits the warm cache. A `blnNeedAll` flag prevents resetting `m_Items(False)` if it was already populated. Chosen.
- **Approach B — Lazy IsModified flag on instances**: Replace two-slot cache with a single dictionary of all items; cache `IsModified` results per instance and filter on demand. Conceptually clean, but filtering creates a new dictionary each time unless cached — reintroducing two-slot complexity. More invasive with no benefit over Approach A. Rejected.
- **Approach C — Lightweight orphan detection (no full instantiation)**: `ClearOrphanedSourceFiles` only needs base names, not full component instances. A new interface method could return just names. Initially dismissed as over-engineered, but production logs proved orphan detection IS a bottleneck (5-6s consistently). However, Approach A eliminates the cost entirely without requiring interface changes, making Approach C unnecessary. Rejected.

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

**Trigger**: Export of a large production database (~6,870 objects, ~567 with descriptions) took ~47s on fast save. Benchmarking revealed the bottleneck was **cold DAO property value reads** in `clsDbDocument.GetDictionary`: iterating Container/Document objects and reading `Description` values took ~18s due to physical disk I/O in the JET engine loading scattered property-value pages. Multiple component classes each called `Set dbs = CurrentDb` independently, and each new `CurrentDb` reference starts with a cold JET page cache (per-reference caching). This meant duplicate cold I/O penalties when multiple components accessed the same data.

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

> **⚠ Partially superseded** (2026-07-31): `[_Last]` and the `LATEST_EXPORT_FORMAT` constant no longer exist. The maintenance pattern is now: add the enum member, add a matching `col.Add` line to `GetExportFormatVersions()`, and gate with `>=`. `LatestExportFormat()` and the options combo derive from that list. See "One declarative list of export formats, guarded by a test that parses the enum" above.

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

**Trigger**: After building a database from source, a subsequent "fast save" export re-exported every single object (e.g., all 3,673 queries in a large production database, taking ~1,600s). The existing `IsModified` logic compared `DateModified > ExportDate`, but every object received a fresh `DateModified` from Access during import, making all objects appear modified.

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
