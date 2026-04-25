# Object Round-Trip Regression Fixtures

This folder contains the regression test corpus consumed by the
`modTestRoundtrip` harness in the add-in. Every fixture is the canonical
exported form of a single database object — typically a `.sql` plus `.json`
pair for queries — and is treated as the source of truth.

The harness:

1. Imports each fixture into the running database under a sandboxed name
   (`vcs_test_<basename>_<hash>`).
2. Exports it twice (Pass 1 and Pass 2).
3. Asserts Pass 2 equals Pass 1 (idempotency, hard requirement).
4. Asserts Pass 1 equals the fixture (drift check, soft requirement).
5. Drops the sandboxed object and moves on.

Fixtures themselves are **never modified** by a normal run. The only way to
overwrite them is to deliberately invoke the harness with `blnRebaseline:=True`
and review the resulting `git diff` before committing.

## How to run

All external invocations go through the public API method
`VCS.RunRoundtripTests`. The harness implementation in `modTestRoundtrip.bas`
uses `Option Private Module`, so test internals stay hidden from cross-project
`Application.Run` lookups (matching the rest of the add-in).

From the VBA Immediate Window (open `frmVCSMain` first to see live progress):

```vba
?VCS.RunRoundtripTests
```

From an MCP `vcs_run_vba` call (requires `McpAllowRunVBA` enabled in the
add-in's Options form, same as any other agent-driven VBA execution):

```vba
MCP_TempFunction = VCS.RunRoundtripTests()
```

End-user / external project (regression-test your own database against your
own corpus, no add-in modifications required):

```vba
?VCS.RunRoundtripTests("C:\path\to\my-fixtures\")
```

To rebaseline (overwrite mismatched fixture files with the actual export —
review `git diff` before committing):

```vba
?VCS.RunRoundtripTests(, True)
```

A dedicated session log file is written to
`Testing\Fixtures\logs\ObjectRoundtrip_<opId>.log`, and the function returns a
JSON document summarizing every fixture (parseable by automation).

## Folder layout

```
Testing\Fixtures\
  _scaffold\           Shared supporting objects loaded once per session
                       (queries/UDFs/tables that fixtures may reference).
                       Loaded with their original names; dropped at end of run.
  queries\
    select\            Plain SELECT queries (single- and multi-table).
    crosstab\          TRANSFORM / PIVOT queries.
    append\            INSERT INTO ... SELECT.
    update\            UPDATE with or without joins.
    delete\            DELETE with or without joins.
    passthrough\       SQL pass-through to an external server.
    union\             UNION / UNION ALL.
    ddl\               CREATE / ALTER / DROP DDL queries.
    regression\        Specific bugs that must never come back. Each fixture
                       in here ideally has a sibling .notes.md linking to
                       the issue it pins down.
  forms\               (future) form fixtures
  reports\             (future) report fixtures
  modules\             (future) standard / class module fixtures
  tabledefs\           (future) table definitions
  scratch\             Per-run intermediate files (gitignored).
  logs\                Per-session log files (gitignored).
```

Inside `queries\` the choice of subfolder is *organizational*; the harness
walks the entire tree. Pick the most specific category that fits.

## Fixture format

For queries (v1):

- **`<name>.sql`** — the formatted SQL statement, exactly as the add-in's
  exporter would write it.
- **`<name>.json`** — companion metadata: query properties, columns, and
  design layout, plus an `Info` block (descriptive metadata: class and query
  name). The `Info` block is purely informational — the import path reads the
  query name from the filename, not from the JSON — so the harness strips it
  entirely before comparison to keep diffs name-agnostic.
- **`<name>.notes.md`** *(optional)* — short prose describing why this
  fixture exists. Especially valuable for regression fixtures: link to the
  GitHub issue, summarize the failure mode, and note any constraints (e.g.
  "this fixture depends on `_scaffold/qryHelper`").

The harness uses the file basename (`<name>`) only to derive the sandbox
object name. Embedded names inside the .sql are unchanged.

## The `_scaffold/` convention

If a fixture references another query, table, or UDF that isn't shipped with
the add-in, place that supporting object in `_scaffold/` (using the same
`.sql`/`.json` pair format for queries). The harness imports every file in
`_scaffold/` once at the start of a session — under each scaffold object's
**original name** so fixtures can reference them directly — and drops them at
the end.

If a same-named object already exists in the host database when the session
starts, the harness emits a warning and skips that scaffold file rather than
clobbering the user's data.

For v1 (queries only), most fixtures are self-contained because Access does
not validate references on import. The convention is established now to avoid
retrofitting it when forms / reports / modules join the corpus.

## Bug-as-fixture: contributing a regression case

The single most useful thing you can do as a user of the VCS add-in is to
contribute a fixture for any object that round-trips incorrectly. The
workflow is intentionally low-friction:

1. **Reproduce the bug** in your own database. If `Export` produces a file
   that, when re-imported and re-exported, doesn't match — you have a
   round-trip bug.
2. **Sanitize the artifact.** Strip anything sensitive: customer-identifying
   table or column names, embedded data, and especially `Connect` strings on
   pass-through queries (replace the connection with the `env:` reference
   convention or a placeholder). Keep the structural shape that triggers the
   bug; rename anything else.
3. **Drop the `.sql` + `.json` pair into the appropriate folder.** Most
   reproductions go in `queries\regression\`. Add a sibling `<name>.notes.md`
   describing what the bug was and linking to the issue.
4. **Run `?VCS.RunRoundtripTests`** locally to confirm the new fixture fails
   (or passes, if you are pinning a recently-fixed case).
5. **Open a PR.** The two text files are the entire change. The fixture
   becomes a permanent regression test against every future add-in change.

This workflow is what makes the harness genuinely valuable: the user's bug
report and the permanent regression test are literally the same artifact.

### Sanitization checklist

- [ ] Connection strings replaced with `env:` references or placeholders.
- [ ] Table / column names rewritten if they leak business context.
- [ ] Embedded sample data removed (it isn't needed — the harness exercises
      query *definitions*, not result sets).
- [ ] Any user names / paths / server names removed from comments.
- [ ] If the fixture depends on a scaffold object, that object is added to
      `_scaffold/` and is also sanitized.

## Naming convention

Test-infrastructure modules use the `modTest*` family prefix
(`modTestRoundtrip`, `modTestSuite`, ...). Fixture files follow whatever the
add-in's exporter produces — typically the original Access object name. Don't
rename fixture files just for organization; let the folder structure carry
that load.

See also: `AGENTS.md` (in the add-in repo root) for the full naming
convention and the broader testing strategy.
