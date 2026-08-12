# Testing strategy

The add-in has three distinct testing layers, each catching a different bug
class. Keeping them separate (rather than one giant test database) is
deliberate — each layer trades scope for fidelity.

| Layer | What it tests | Lives in |
|---|---|---|
| **1. VBA logic tests** | "Given inputs, does this function return the right output?" | `modTestSuite.bas` |
| **2. Object round-trip tests** | "Does this database object survive a serialize/deserialize cycle unchanged?" | `modTestRoundtrip.bas` + `Testing/Fixtures/` |
| **3. Whole-database integration** | "Does building an entire database from source produce a working database?" | `Testing/Testing.accdb.src` |

For how to *write* a test — module and class shapes, discovery rules,
`TestAssert`, tags, global hooks — see
[.cursor/rules/testing.mdc](../.cursor/rules/testing.mdc). For the runner's web
UI and its VBA-to-JavaScript bridge, see
[web-test-runner.md](web-test-runner.md). For the filter syntax used by
`VCS.RunTests`, see [AGENTS.md](../AGENTS.md).

---

## `modTest*` naming convention

All test-infrastructure modules use the `modTest*` family prefix. This matches
the family grouping already used in the codebase (`clsDb*` for component
classes, `clsLv*` for ListView property parsers) and gives "Test" maximum
prominence for developer and agent discoverability.

- `modTestSuite` — heterogeneous unit tests (encoding, JSON, sanitization,
  formatter, hashing, `IDbComponent` invariants).
- `modTestRoundtrip` — generalized object round-trip regression harness.
- Future siblings (e.g., `modTestPerf`, `modTestFixtures`, `modTestEncoding`)
  adopt the same prefix automatically.

New test modules go under `Version Control.accda.src/modules/Tests/`. If a module
wraps an existing concept (encoding, hashing, sanitization), prefer extracting it
from `modTestSuite` into a focused `modTest<Topic>` module rather than letting
`modTestSuite` grow indefinitely.

---

## Object round-trip harness (`modTestRoundtrip`)

The Layer 2 harness is generic over `IDbComponent`. It currently covers queries
(`Testing/Fixtures/queries/`) and local table definitions
(`Testing/Fixtures/tabledefs/`); forms, reports, modules, and table data follow
the same pattern by adding a per-type helper.

Table definition fixtures additionally assert **which import path ran**: fixtures
under `tabledefs/` must be built by `modTableDefBuilder` through DAO, while those
under `tabledefs/fallback/` must be refused by it and fall back to
`Application.ImportXML`. Expected path is decided only by that folder placement —
there is no per-fixture declaration. Put attachment, multi-value, and other
deliberately unsupported constructs under `fallback/`. One case is easy to miss: a
text or memo field whose XML carries no `AllowZeroLength` property also belongs
under `fallback/`, because `CreateField` always materializes that property and it
cannot be deleted, so the DAO build can never re-export identically. Because the DAO builder's
output only matches source under the canonical property ordering, the harness
forces `ExportFormatVersion = EFV_5_1_0` for the duration of the run and restores
it afterwards.

For each fixture under `Testing/Fixtures/`, the harness:

1. Imports the fixture into the running database under a sandboxed name
   (`vcs_test_<basename>_<hash>`).
2. Validates the emitter's `.qdef` output:
   - `import_path` — for a fixture carrying a `DesignLayout`, asserts Access
     really stored a designer grid (`MSysObjects.LvExtra` populated), so a lost
     layout fails here rather than as a later layout diff.
   - `qdef_joins` — structural check: each join row's `LeftTable`/`RightTable`
     matches its `Expression` (Design View only).
   - `qdef_vs_fixture` — drift check: compares generated `.qdef` against the
     stored `.qdef` baseline (if present).
3. Exports it twice (Pass 1 and Pass 2), into a per-run scratch folder.
4. Asserts Pass 2 == Pass 1 (idempotency, hard requirement).
5. Asserts Pass 1 == fixture (drift check, soft requirement).
6. Drops the sandboxed object and moves on.

Output goes to three coordinated channels:

- **`frmVCSMain` console** — live progress (one line per fixture), then a summary
  that echoes the log file and scratch folder paths as plain text so they can be
  copied straight into an agent session.
- **Per-session log file** — `Testing/Fixtures/logs/ObjectRoundtrip_<opId>.log`,
  with full unified diffs for any failures.
- **JSON return value** — machine-parseable summary suitable for `vcs_run_vba`
  callers and CI. Includes `logPath` and `scratchFolder`.

All external invocations go through the public API method
`VCS.RunRoundtripTests`. The implementation in `modTestRoundtrip.bas` uses
`Option Private Module` so test internals stay hidden from cross-project
`Application.Run` lookups, matching the rest of the add-in.

```vba
?VCS.RunRoundtripTests                             ' Immediate Window
?VCS.RunRoundtripTests("C:\path\to\my-fixtures\")  ' A user's own corpus
```

Via MCP (requires `McpAllowRunVBA`, same as any other agent-driven execution):

```
vcs_run_vba(<addin-path>, "MCP_TempFunction = VCS.RunRoundtripTests()")
```

Pass `True` as the second argument to rebaseline mismatched fixtures (review the
resulting git diff before committing). When working inside the add-in's own VBE —
debugging the harness itself — the in-project entry point
`?modTestRoundtrip.RunObjectRoundtripTests()` also works, since
`Option Private Module` only blocks cross-project callers.

### Bug-as-fixture: the contribution workflow

The harness supports a contribution workflow uniquely enabled by the add-in's
text-source format. When a user hits an edge case where an object fails to
round-trip:

1. They reproduce the bug in their own database.
2. They sanitize the failing object's `.sql` + `.json` pair (strip
   business-sensitive names, replace `Connect` strings with `env:` references,
   remove embedded data).
3. They drop the pair into `Testing/Fixtures/queries/regression/` (or the
   appropriate category) on a branch.
4. They optionally add a sibling `<name>.notes.md` describing the bug and linking
   to the issue.
5. They open a PR against the add-in.

The fixture becomes a permanent regression test against every future change. The
bug report and the regression test are literally the same artifact.

When adding regression fixtures from a user or production database, sanitize both
the fixture and its `.notes.md`. Do not include source database names, source
query names, table/field names, business concepts, file paths, customer names, or
server names. Use generic parser-shape language such as "production validation
exposed a cross-subtree join predicate placement bug."

See [Testing/Fixtures/README.md](../Testing/Fixtures/README.md) for the full
workflow, the `_scaffold/` convention (fixtures with shared dependencies), and a
sanitization checklist.

---

## Headless runs (CI / automation)

`VCS.RunTestsHeadless` takes the same filter arguments as `VCS.RunTests` but is
designed for unattended sessions: no forms are opened, no prompts are raised (a
missing `modTestAssert` is installed silently), the web runner is bypassed, and
JUnit XML is always exported regardless of the `ExportTestResultsJUnit` option.
The returned JSON includes `allPassed`, `cancelled`, `junitPath`, and `statePath`
in the root, so a pipeline can assert the outcome without parsing per-test detail.

```powershell
# Drive from PowerShell via COM automation (Access stays invisible).
# Application.Run on the add-in's public API function loads the add-in library
# and routes the call to clsVersionControl (see modAPI.bas `API`); up to three
# filter arguments are supported through this route.
$addin = "$env:AppData\MSAccessVCS\Version Control.API"
$access = New-Object -ComObject Access.Application
$access.OpenCurrentDatabase("C:\path\to\Database.accdb")
$json = $access.Run($addin, "RunTestsHeadless", "-slow")
$access.Quit()
$result = $json | ConvertFrom-Json
if (-not $result.allPassed) { exit 1 }
```

From an MCP session, `MCP_TempFunction = VCS.RunTestsHeadless()` via
`vcs_run_vba` works the same way.

---

## Where results and logs land

Results from any run are:

- Streamed live to the `frmVCSMain` console (color-coded: green PASS, red
  FAIL/ERROR, gray EMPTY, orange SKIP) or to the web runner UI.
- Saved as JSON to `<export-folder>/logs/TestResults_<timestamp>.json`
  (ephemeral per-run history).
- Merged into `<export-folder>/test-results/test-state.json` (single durable
  current state; survives Access restarts, partial runs mark untouched tests
  `stale`).
- Exported as JUnit XML to `<export-folder>/test-results/test-results.xml` when
  `Options.ExportTestResultsJUnit` is enabled (default on).
- Exported as a self-contained HTML report to
  `<export-folder>/test-results/test-results.html` when
  `Options.ExportTestResultsHtml` is enabled (default on).
- Returned as a JSON string from `VCS.RunTests`, with a `"tags"` array per entry.

`VCS.ExportTestResultsJUnit` and `VCS.ExportTestResultsHtml` regenerate those
artifacts from current state without re-running tests.

**Log files are gitignored.** All `logs/` directories and `*.log` files are
excluded by `.gitignore`, along with `test-results/`. Agent tools that respect
`.gitignore` (Glob, Grep, semantic search) will silently skip them — use shell
commands instead:

```powershell
# Find log files (run from repository root)
Get-ChildItem -Recurse -Include "*.log","*.json" | Where-Object { $_.DirectoryName -like "*logs*" }
```

Key locations:

- `Version Control.accda.src/logs/` — build, export, merge, and ephemeral test
  run logs (`TestResults_*.json`, `TestRun_*.log`).
- `Version Control.accda.src/test-results/` — durable test state, JUnit XML, and
  the HTML report.
- `Testing/Fixtures/logs/` — object round-trip logs (`ObjectRoundtrip_*.log`).

Operation logs are named for the operation type, not the entry point:
`Export_*.log` for exports, `Build_*.log` for full builds, and `Merge_*.log` for
merge builds **and single-object imports** (`LoadSelected` / `ImportObject` begin
an `eotMerge` operation, and `clsLog.LogFilePath` maps that to the `Merge`
prefix). A single-object import that appears to have produced no log has almost
certainly written a `Merge_*.log`.

---

## Fixture corpus vs. sample database

The canonical object round-trip fixture corpus is `Testing/Fixtures/`, as plain
text files. Query fixtures live under `Testing/Fixtures/queries/` as `.sql` +
`.json` pairs, with optional `.notes.md` files for regression context.

`Testing/Testing.accdb.src` is the sample/integration database used for full
build/export flows; do not look there for the primary `VCS.RunRoundtripTests`
fixture corpus. `MSysQueriesExamples` and `db-analysis-tools` are useful sources
or validation projects for query shapes, but they are not the add-in's canonical
regression fixture store.
