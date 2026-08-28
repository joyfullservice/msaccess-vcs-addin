# AGENTS.md — MSAccess VCS Add-in

A Microsoft Access add-in that exports database objects (forms, reports, queries,
modules, table data) to text-based source files suitable for Git, and rebuilds or
merges databases back from those files. It tracks changes through a binary index
for fast incremental exports, resolves database-versus-source conflicts, and
supports ADP projects and external SQL schema export. The add-in's own source
lives in `Version Control.accda.src/`, exported by the add-in itself, so the
repository is its own largest test case.

This file loads on every turn and is budgeted at **150 lines**. Depth belongs in
[docs/](docs/README.md); read [docs/agent-docs-maintenance.md](docs/agent-docs-maintenance.md) first.

## Where to read next

| Working on | Read |
|---|---|
| Orienting; adding a component type; finding the class for an object type | [docs/architecture.md](docs/architecture.md) |
| Anything that changes exported file content or layout | [docs/export-format-versioning.md](docs/export-format-versioning.md) |
| Diagnosing an error in a log file | [docs/error-handling.md](docs/error-handling.md) |
| Why an operation was slow; reading performance data as JSON | [docs/perf-diagnostics.md](docs/perf-diagnostics.md) |
| Inspecting a database's schema over MCP; a failing `vcs_run_vba` call | [docs/mcp-runvba.md](docs/mcp-runvba.md) |
| Rebuilding the add-in unattended; a refused or stalled rebuild | [docs/agentic-rebuild.md](docs/agentic-rebuild.md) |
| Running this repo's own tests over MCP; an all-`EMPTY` result | [docs/agent-test-runs.md](docs/agent-test-runs.md) |
| Test layers, fixtures, round-trip harness, CI | [docs/testing-strategy.md](docs/testing-strategy.md) |
| The test runner UI or its VBA/JS bridge | [docs/web-test-runner.md](docs/web-test-runner.md) |
| Writing an individual test | [.cursor/rules/testing.mdc](.cursor/rules/testing.mdc) |
| The query parser (`clsQueryComposer`, `clsDbQuery`) | [docs/access-query-storage.md](docs/access-query-storage.md) |
| Conditional formatting binary blobs | [docs/access-conditional-format.md](docs/access-conditional-format.md) |
| Editing exported source files by hand | [Version Control.accda.src/AGENTS.md](Version%20Control.accda.src/AGENTS.md) |
| Why something was built this way | [DECISIONS.md](DECISIONS.md) (append-only; search it) |

## Development workflow

**Agents: edit the source files, then rebuild — never touch the installed
add-in.** Prefer `vcs_rebuild_addin("<source>")`; it streams detailed build
callbacks, then watches compile/install status. `vcs_call_vba` stays
launch-only. The rebuild refuses when another Access process holds a file it
must replace; then test. See
[docs/agentic-rebuild.md](docs/agentic-rebuild.md) and [docs/agent-test-runs.md](docs/agent-test-runs.md).

By hand: install a recent release, clone the repo, use **Build From Source** on
`Version Control.accda.src`, then modify the running `Version Control.accda`,
test, and run `Deploy` in the VBA Immediate Window — it increments the version,
exports to source files, and installs the development build. Commit the source
files and open a pull request against `dev`; releases are cut from `master`.

## Invariants

None are compiler-enforced; missing one causes silent data loss, churn in user
projects, or untranslatable UI.

- **Never use `Dir()`.** It does not support Unicode filenames and silently skips
  or fails on non-ASCII paths, which Access object names frequently produce. Use
  `Scripting.FileSystemObject`, or `modFileWinAPI` for performance-critical scans.
- **Wrap every user-facing string in `T()`**, using `{0}`-style placeholders for
  substitution: `Log.Add T("Error in file: {0}", var0:=strFileName)`.
- **Use library constants, not magic numbers.** All modules share the same
  references (DAO, VBE, Scripting), so write `dbQSQLPassThrough`, `acQuery`, and
  `vbTextCompare` rather than their numeric values.
- **Gate any change to exported output behind an export format version.** Users
  must be able to upgrade the add-in without their source files churning. See
  [docs/export-format-versioning.md](docs/export-format-versioning.md) for the
  three mechanisms. Import logic is never gated — it stays backwards compatible.
- **Never create a second file with the same basename** in another folder of an
 export tree. Placement is driven by the `'@Folder` annotation inside the file;
 duplicates cause last-one-wins behavior on build and corrupt the change index.
- **Never patch the loaded add-in in place.** `VBComponents.Remove`/`.Import`,
 `LoadFromText`, and `vcs_import_object` aimed at `Version Control.accda` reset a
 VBA project that is currently executing, killing the MCP session and possibly
 leaving the add-in broken. Edit the source and rebuild ([docs/agentic-rebuild.md](docs/agentic-rebuild.md)).

## Error handling

Procedures opt into a structured inline pattern rather than relying on a
top-level handler:

```vba
Public Sub SomeOperation()
    If DebugMode(True) Then On Error GoTo 0 Else On Error Resume Next

    ' ... operation code ...

    CatchAny eelError, "Error description", ModuleName & ".SomeOperation", True, True
    If Operation.ErrorLevel = eelCritical Then GoTo CleanUp

CleanUp:
End Sub
```

`DebugMode(True)` reports whether debug mode is on and internally calls
`LogUnhandledErrors`, which must run before any `On Error` directive because
`On Error` silently clears `Err`. `CatchAny` logs and optionally clears; `Catch`
tests for specific error numbers. A log entry reading ``Unhandled error, likely
before `On Error` directive`` does **not** mean the error originated there — it
came from whatever ran immediately before, which
[docs/error-handling.md](docs/error-handling.md) explains how to trace.

## Naming conventions

| Element | Convention | Example |
| --- | --- | --- |
| Modules, classes, interfaces, forms | `mod` / `cls` / `I` / `frm` | `modImportExport`, `clsDbForm`, `IDbComponent`, `frmVCSMain` |
| Test modules and classes | `modTest` / `clsTest` | `modTestRoundtrip` |
| Private module vars; UDT instance | `m_`; `this` | `m_Items`; `Private this As udtThis` |
| Constants; enums | `UPPER_CASE` or `PascalCase`; `e` | `CHUNK_SIZE`; `eErrorLevel` |
| Boolean / String / numeric params | `bln` / `str` / `lng` / `int` | `blnModifiedOnly`, `strFile` |
| Dictionary / Collection / class object | `d` / `col` / `c` | `dFiles`, `colCategories`, `cDbObject` |

Every module opens with a header block and `Option Compare Database` /
`Option Explicit`; [docs/architecture.md](docs/architecture.md) has the template.

## Running tests

Over MCP, `database_path` decides which project the runner scans, so **run this
repo's tests on the development copy** — `vcs_run_tests("<repo>\Version Control.accda")`.
The installed copy under `%AppData%` is only ever loaded as a library, where it
supplies the runner and `TestAssert`; hosting a run on it is refused. Nothing in
`Testing/` hosts a run either, and an all-`EMPTY` result is a broken harness, not
a pass. [docs/agent-test-runs.md](docs/agent-test-runs.md) covers the traps.

`VCS.RunTests` takes filters resolved in priority order: module name, suite or
`@Folder` value (exact or final segment), procedure or `Module.Procedure` key,
then `'@Tag`. Prefix with `-` to exclude; inclusions OR, exclusions AND.

```vba
?VCS.RunTests("modTestEncoding")   ' One module (omit all filters to run everything)
?VCS.RunTests("SQL", "-slow")      ' Run SQL suite, skip slow tests
?VCS.RunTestsHeadless("-slow")     ' Unattended: no forms, always writes JUnit
?VCS.RunRoundtripTests             ' Object round-trip fixture corpus
```

## Key files in an export folder

| File | Purpose |
|------|---------|
| `vcs-options.json` | Per-project configuration (export folder, options) |
| `vcs-index.idx` | Change tracking index, binary. Do not hand-edit; dump it with `VCSIndex.DumpToJson` |
| `project.json` | Database file format version |
| `vbe-references.json`, `dbs-properties.json` | VBA library references; database properties |

## Documentation that ships to users

`Version Control.accda.src/AGENTS.md` and `vcs-agent-docs/` beside it ship into
**every user's export folder on every export** — a different audience, tighter
budgets, own gate. Read [docs/agent-docs-maintenance.md](docs/agent-docs-maintenance.md) first.

## Resources
[Repository](https://github.com/joyfullservice/msaccess-vcs-addin) · [Wiki](https://github.com/joyfullservice/msaccess-vcs-addin/wiki) · [Issues](https://github.com/joyfullservice/msaccess-vcs-addin/issues) ·
[Releases](https://github.com/joyfullservice/msaccess-vcs-addin/releases)
