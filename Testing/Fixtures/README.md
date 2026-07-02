# Object Round-Trip Regression Fixtures

This folder contains the regression test corpus consumed by the
`modTestRoundtrip` harness in the add-in. Every fixture is the canonical
exported form of a single database object — typically a `.sql` plus `.json`
pair for queries — and is treated as the source of truth.

The harness:

1. Imports each fixture into the running database under a sandboxed name
   (`vcs_test_<basename>_<hash>`).
2. Validates the emitter's `.qdef` output (`qdef_joins` check — see below).
3. Exports it twice (Pass 1 and Pass 2).
4. Asserts Pass 2 equals Pass 1 (idempotency, hard requirement).
5. Asserts Pass 1 equals the fixture (drift check, soft requirement).
6. Drops the sandboxed object and moves on.

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

## Qdef emitter validation

The harness runs two checks on the `.qdef` that the emitter would produce for
each fixture. The `.qdef` is a structured text file that `Application.LoadFromText`
consumes to create the query — bugs in the emitter can cause `LoadFromText` to
silently create broken queries that fail at execution time.

### `qdef_joins` — structural invariant check

For Design View `.qdef` fixtures, the harness validates that every join row's
`LeftTable` and `RightTable` reference tables that appear in that row's
`Expression`. This catches emitter bugs where split compound `ON` conditions
get the wrong table pair — bugs that `Application.LoadFromText` accepts silently
but cause DAO error 3082 at query execution time.

Skipped for SQL-View-only queries (UNION, DDL, pass-through, etc.) since those
have no `Joins` block. See
[regression/qryRegressionCrossTableOnEmitter.notes.md](queries/regression/qryRegressionCrossTableOnEmitter.notes.md)
for the fixture that pins this invariant.

### `qdef_vs_fixture` — drift comparison

For fixtures that have a `.qdef` baseline file (alongside the `.sql` and
`.json`), the harness compares the generated `.qdef` against the baseline and
reports any differences. This catches *any* change to the emitter's output —
properties, input tables, output columns, joins, ordering, etc. — not just the
specific invariant that `qdef_joins` checks.

The `.qdef` baseline is generated without the Layout section (layout is a
pass-through from the JSON's `DesignLayout` blob, already validated by the
`json_vs_fixture` check).

To create baselines for all fixtures, run with `blnRebaseline:=True`:

```vba
?VCS.RunRoundtripTests(, True)
```

Review the resulting `git diff` before committing — each new `.qdef` file
should be inspected for correctness. Subsequent runs without rebaselining will
detect drift against these baselines.

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
- **`<name>.qdef`** *(optional)* — the `.qdef` text the emitter generates
  for this fixture. Stored without the Layout section (layout drift is
  covered by the `.json` comparison). When present, the harness compares
  the freshly-generated `.qdef` against this baseline and reports any
  drift. Created automatically when running with `blnRebaseline:=True`.
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

Fixture notes (`*.notes.md`) must be sanitized too. Do not name the source
database, source query, customer/project, table, field, path, or business
domain that exposed the bug. Describe the generic parser shape instead, e.g.
"a production validation run exposed a cross-subtree join predicate placement
bug" rather than naming the production query or its business objects.

## Documenting parser invariants and edge cases

A round-trip fixture pins a specific shape, but most non-trivial parser
work also produces durable knowledge that doesn't fit inside one fixture
file: cross-cutting design decisions, Access-API quirks, format-level
references, and the running list of "things we know we don't yet handle."
Future agents lose this knowledge silently if it isn't captured somewhere
discoverable.

The repo uses **four layers** of documentation, each with a distinct
trigger. Every non-obvious parser fact should land in at least one layer;
cross-references between layers are encouraged.

| Layer | Lives at | Trigger |
|-------|----------|---------|
| **Reference doc** | [docs/access-query-storage.md](../../docs/access-query-storage.md) | "Is this a stable fact about how Access stores queries, or about which shapes our parser does/doesn't handle?" |
| **DECISIONS.md** | [DECISIONS.md](../../DECISIONS.md) at repo root | "Would the next agent reasonably try the rejected approach again? Did we evaluate alternatives that aren't obvious from the code?" Append-only journal. |
| **`.notes.md` beside a fixture** | `Testing/Fixtures/queries/<category>/<name>.notes.md` | "Is there a single SQL pattern that demonstrates this? What would re-break it?" |
| **Procedure-header comment** | Inside the `.cls` / `.bas` source | "Would someone reading just this function be confused without context the body itself doesn't carry?" Add an `INVARIANT:` or `EDGE CASE:` callout. |

The four layers are complementary, not exclusive. A new finding often
warrants entries in two or three of them — for example, the
multi-condition `ON` `LoadFromText` rejection finding lives as a
fact in [docs/access-query-storage.md § 6](../../docs/access-query-storage.md),
as a decision rationale (alternatives ruled out) in `DECISIONS.md`, and as
a regression pin in
[regression/qryRegressionMultiCondJoin.notes.md](queries/regression/qryRegressionMultiCondJoin.notes.md).

### Before you change `clsQueryComposer.cls` or `clsDbQuery.cls`

The parser carries hard-won decisions in places that are not always
obvious from a casual read. Before modifying either class, do this short
read pass:

1. **Read the parser reference.** Skim
   [docs/access-query-storage.md](../../docs/access-query-storage.md) §§ 4
   and 5 for the shape you're about to alter. § 4 tells you which
   fixtures exercise it; § 5 tells you whether we have a known gap that
   touches it.
2. **Search `DECISIONS.md`.** Run `rg "clsQueryComposer" DECISIONS.md -A 30`
   (or the corresponding search for `clsDbQuery`). Recent entries cover
   SQL reconstruction fidelity, the Design / SQL view fallback, column
   metadata serialization, and the composer's error-handling shape.
3. **Skim the regression fixtures.**
   `Testing/Fixtures/queries/regression/*.notes.md` — each one is a
   1-paragraph description of what would re-break if a careful decision
   were reverted.
4. **Read the procedure-header comments.** Functions like
   `RequiresDesignView`, `IsDesignerCompatible`, `HasTopLevelBoolean`,
   `ParseJoinExpression`, `SafeBreak`, and `EmitDbMemoSql` carry
   constraints in their headers that the body alone does not convey.

### Worked example: multi-condition `ON`

Concrete example of the four layers working together for a single
finding:

- **Reference doc.**
  [docs/access-query-storage.md § 6](../../docs/access-query-storage.md)
  records the empirical fact: Access's
  `Application.SaveAsText` produces a `dbMemo "SQL"` qdef for queries
  with multi-condition `ON`, but `Application.LoadFromText` rejects the
  same file. This is not in the upstream Riddington documentation;
  it's specific to our use of `LoadFromText` for import.
- **DECISIONS.md.** The rationale entry (when added in the deferred
  follow-up pass) records the alternatives that were ruled out — for
  example, "fall back to setting `QueryDefs(name).SQL` directly like the
  legacy path does" was rejected because the new pipeline can choose its
  qdef shape rather than receive one.
- **`.notes.md`.** [regression/qryRegressionMultiCondJoin.notes.md](queries/regression/qryRegressionMultiCondJoin.notes.md)
  pins the specific SQL shape and explains what would re-break it.
- **Procedure-header comment.** `clsQueryComposer.RequiresDesignView`
  carries an `' INVARIANT:` callout that names the shape and points back
  at both the fixture and the reference doc.

### `INVARIANT:` and `EDGE CASE:` comment convention

Two lightweight tags for in-code callouts that are intentionally
greppable:

- `' INVARIANT:` — "this must remain true; reverting this is a known
  regression." Used at the call site or the procedure header for any
  property the implementation is deliberately upholding.
- `' EDGE CASE:` — "Access (or some other dependency) does this weird
  thing here; the surrounding code is the workaround." Used at the
  workaround site.

Both end with a pointer to the other layer(s) carrying the full story:

```vba
' INVARIANT: multi-condition ON cannot survive a SQL View qdef round-trip;
' Access's LoadFromText rejects what its own SaveAsText produces. We force
' Design View import for this shape via RequiresDesignView.
' See: docs/access-query-storage.md § 6, regression/qryRegressionMultiCondJoin
```

The convention is already in informal use on the procedure headers for
`RequiresDesignView`, `SafeBreak`, and `EmitDbMemoSql`; the tags above
are the canonical form for new callouts.

## Naming convention

Test-infrastructure modules use the `modTest*` family prefix
(`modTestRoundtrip`, `modTestSuite`, ...). Fixture files follow whatever the
add-in's exporter produces — typically the original Access object name. Don't
rename fixture files just for organization; let the folder structure carry
that load.

See also: `AGENTS.md` (in the add-in repo root) for the full naming
convention and the broader testing strategy.
