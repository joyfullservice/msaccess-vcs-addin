# Regression Testing

The add-in ships with a generalized regression-test harness for verifying that database objects survive a serialize / deserialize cycle unchanged. This page covers what it does, how to run it, and how to contribute new fixtures when you hit an edge case.

## Why it exists

Every change to the export / import pipeline (sanitization rules, formatter quirks, JSON layout, edge-case handling) carries the risk of breaking some object somewhere. With thousands of users and an open-ended object grammar, the only sustainable way to defend against regressions is a corpus of real-world fixtures that can be re-checked on every change.

The harness is generic over `IDbComponent`, so the same machinery covers queries today and forms, reports, modules, table data, etc. as fixtures are added.

## What it does

For each fixture under `Testing/Fixtures/`, the harness:

1. Imports the fixture into the running database under a sandboxed name (`vcs_test_<basename>_<hash>`).
2. For **queries**, validates generated `.qdef` structure:
   - **`qdef_joins`** — each join row's tables match its `Expression` (Design View structural check).
   - **`qdef_vs_fixture`** — compares generated `.qdef` to a stored `.qdef` baseline when present.
3. Exports twice (Pass 1 and Pass 2) into a per-run scratch folder.
4. Asserts **Pass 2 == Pass 1** (idempotency — hard requirement).
5. Asserts **Pass 1 == fixture** (drift check — soft requirement; warnings when rebaselining).
6. Drops the sandboxed object and refreshes the cache (`DBEngine.Idle dbRefreshCache`) before moving on.

Stale sandbox objects from a prior crashed run are detected and cleaned up at the start of every session, so the database does not accumulate cruft over time.

## Running it

### From the VBA Immediate Window

Open `frmVCSMain` first if you want live progress in the console.

```vba
?VCS.RunRoundtripTests
```

To run against your own corpus instead of the shipped fixtures, pass a folder path:

```vba
?VCS.RunRoundtripTests("C:\path\to\my-fixtures\")
```

To rebaseline (overwrite fixtures with the actual export when comparisons mismatch — review the resulting git diff carefully before committing):

```vba
?VCS.RunRoundtripTests(, True)
```

### Via MCP / `vcs_run_vba`

```
vcs_run_vba(<addin-path>, "MCP_TempFunction = VCS.RunRoundtripTests()")
```

This requires **Allow Arbitrary VBA Execution** (`McpAllowRunVBA`) under **Options** → **MCP**. See [MCP and Automation](MCP-and-Automation).

For unit tests (not round-trip), use [Testing](Testing) and `VCS.RunTests`.

Inside the add-in's own development VBE, the harness is also callable via `?modTestRoundtrip.RunObjectRoundtripTests()` for in-project debugging.

## What you get back

The harness produces output on three coordinated channels:

| Channel | Audience | Contents |
|---|---|---|
| `frmVCSMain` console | Developer running interactively | One line per fixture (✔/✖), running totals, summary block |
| `Testing/Fixtures/logs/ObjectRoundtrip_<opId>.log` | Post-mortem inspection | Full log with unified diffs for every failure |
| JSON return value | CI / `vcs_run_vba` callers | Machine-parseable summary with per-fixture results, hashes, and diff payloads |

The JSON shape is:

```json
{
  "success": true,
  "fixtureFolder": "...",
  "scratchFolder": "...",
  "logPath": "...",
  "stats": { "total": 15, "passed": 15, "failed": 0, "skipped": 0, "errors": 0, "elapsedSeconds": 4.2 },
  "results": [
    { "fixture": "qryCars", "category": "select", "status": "pass", "checks": [...] },
    ...
  ]
}
```

## Fixture corpus layout

```
Testing/Fixtures/
├── README.md                ← contributor-facing usage
├── .gitignore               ← excludes scratch/ and logs/
├── _scaffold/               ← shared supporting objects (loaded once per session)
│   └── .gitkeep
├── queries/
│   ├── select/              ← qryCars.sql + qryCars.json, ...
│   ├── crosstab/
│   ├── append/
│   ├── update/
│   ├── delete/
│   ├── union/
│   ├── passthrough/
│   ├── ddl/
│   └── regression/          ← bug-fix pin-downs, with sibling .notes.md files
├── scratch/                 ← per-run output (gitignored)
└── logs/                    ← per-session log files (gitignored)
```

The `_scaffold/` folder is a convention for fixtures that depend on shared tables or supporting queries. Anything in `_scaffold/` is imported once at the start of a run and dropped at the end. v1 doesn't ship any scaffold objects (the shipped corpus exercises only standalone queries), but the convention is in place for future fixtures that need it.

## The bug-as-fixture workflow

This is the contribution pattern the harness was built around. When you hit an object that fails to round-trip:

### 1. Reproduce in your own database

Confirm the bug reproduces with the current add-in. If you can isolate it to a single object, great — that's your fixture candidate.

### 2. Sanitize the fixture

Strip anything you can't share publicly:

- Replace business-sensitive table / field / query names with generic ones (`tblCustomer` → `tblA`, `lngCustomerId` → `lngId`).
- Replace `Connect` strings on linked tables with `env:` references or remove the linked tables entirely if they aren't required to reproduce the bug.
- Remove embedded sample data unless the bug is data-dependent.
- Strip comments that reference internal projects or people.

The goal is the smallest possible reproducer that still fails the round-trip in the same way.

### 3. Drop the pair into the right category

The sanitized `<name>.sql` + `<name>.json` pair goes into the most appropriate subfolder under `Testing/Fixtures/queries/`. If you're pinning down a specific bug, use `regression/`.

### 4. Add a `.notes.md` (recommended for `regression/`)

Sibling markdown documenting the bug. Format:

```markdown
# qryYourFixture

**Bug:** Brief description of the failure mode.

**Symptom:** What the user sees (compile error, runtime error, silent corruption, etc.).

**Root cause:** Brief explanation if known.

**Fixed in:** Issue / PR link.
```

### 5. Open a PR

Title it `regression: <one-line description>` against `dev`. The fixture itself is the test — once it passes locally, it'll guard against the bug forever.

## What makes this approach work

The unique enabler here is the add-in's text-source format. Because every Access object can be expressed as a `.sql` + `.json` (or `.bas`, `.cls`, `.form`, etc.) pair, a "bug report" and a "regression test" are literally the same artifact. Other Access projects can't easily do this because they lack a canonical text representation of database objects.

This means:

- Bug reproduction is a file copy, not a screen recording.
- Regression coverage grows organically with the user base, not just with the maintainers' bandwidth.
- The corpus doubles as a portable benchmark for evaluating proposed changes to the export pipeline (you can rerun it across versions).

## Adding new component types

v1 supports queries. To add support for forms (or any other `IDbComponent`):

1. Create a `Testing/Fixtures/forms/` subtree mirroring the `queries/` layout.
2. Add a `RunFormFixtures` helper to `modTestRoundtrip.bas` modeled on `RunQueryFixtures`. It needs to know how to enumerate the fixture pairs, import them under a sandboxed name, export twice, compare, and clean up. Most of the existing helpers (`ProvisionScratchFolder`, `MakeUnifiedDiff`, `AddCheck`, `RollUpStatus`, `LogFixtureResult`) are component-agnostic and reusable.
3. Wire the new helper into `RunObjectRoundtripTests`.

The public API method (`VCS.RunRoundtripTests`) doesn't change — it's deliberately component-agnostic.

See `modTestRoundtrip.bas` for the existing query implementation as a template.
