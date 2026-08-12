# Automated Testing

The add-in includes a test runner that discovers and executes tests in whatever
database is open. Tests assert with `TestAssert`, a drop-in replacement for
`Debug.Assert`. Nothing in the project needs a compile-time reference to the add-in.

## Setup

Run `VCS.InstallTestAssertModule` from the Immediate Window to inject the
`modTestAssert` module. If test code already uses `Debug.Assert`,
`VCS.MigrateDebugAssert` converts it in bulk.

## Writing tests

A test is a parameterless `Public Sub` in a test module:

```vba
Option Compare Database
Option Explicit
Option Private Module
'@Folder("Tests")

Public Sub TestDoubleInput()
    TestAssert MyFunction(42) = 84, "MyFunction should double input"
    TestAssert MyFunction(0) = 0, "Zero input returns zero"
End Sub
```

The second `TestAssert` argument is optional context that identifies which
assertion failed, which matters inside loops and shared helpers.

A module counts as a test module if it carries `'@Folder("...Tests...")` (in
projects that use `@Folder` annotations at all) or its name contains `Test`. Within
it, only parameterless `Public Sub` procedures are registered, and only if the
module contains at least one `TestAssert` call. To keep a helper out of the suite,
make it `Private` or give it a parameter.

**Class modules** work the same way and add per-test setup and teardown: every test
method gets a fresh instance, so `Class_Initialize` runs before it and
`Class_Terminate` after. Use parameterless `Public Sub` or `Public Function`.

Name test modules `modTest*` for standard modules or `clsTest*` for classes, and
mark standard test modules `Option Private Module`.

## Running tests

Run `?VCS.RunTests` from the Immediate Window, or use **Tools > Run Tests** on the
ribbon. After a completed run the runner can re-run only the failures.

`RunTests` takes an optional `ParamArray` of filters. Each argument resolves in
priority order: exact module name, then suite or `@Folder` value (matching the full
path or its final segment, so `"SQL"` matches `"Tests.SQL"`), then procedure name or
full `Module.Procedure` key, then tag.

Prefix an argument with `-` to exclude it. Inclusions combine with OR, exclusions
with AND, and a filter list containing only exclusions starts from all tests.

```vba
?VCS.RunTests("-slow")                   ' Everything except slow-tagged tests
?VCS.RunTests("Reporting", "-slow")      ' The Reporting suite, skipping slow tests
?VCS.RunTests("TestInvoiceTotal")        ' One specific procedure
```

`RunTests` returns a JSON summary with per-test status, assertion detail, and a
`tags` array for each test.

## Tagging

`'@Tag("name")` annotations categorize tests and are case-insensitive. A
module-level tag sits in the first ~30 lines, before any procedure, and applies to
every test in the module. A procedure-level tag sits at the very top of the
procedure body, before any executable line including `Dim`. The two sets merge.

```vba
'@Tag("slow")           ' Module-level: inherited by every test here

Public Sub TestExpensiveQuery()
    '@Tag("database")   ' Procedure-level: this test only
    TestAssert RunCheck(), "check passes"
End Sub
```

## Global suite hooks

`modTestAssert` may define two optional once-per-run hooks:

```vba
Public Sub GlobalTestSetup()    ' Before the first test, when at least one is selected
Public Sub GlobalTestTeardown() ' After all tests; the results JSON already exists
```

`VCS.InstallTestAssertModule` writes empty stubs with inline guidance. If they are
absent the runner skips them silently. They do not run when no tests are discovered
or a filter matches nothing. An error inside a hook is non-fatal — it goes to the
console, the run continues, and teardown still executes. Per-test
`Class_Initialize` and `Class_Terminate` nest inside these hooks unchanged.

## Where results land

| File | Contents |
|------|----------|
| `logs/TestResults_<timestamp>.json` | Per-run results with per-assertion detail |
| `logs/TestRun_<timestamp>.log` | Full console output including timing |
| `test-results/test-state.json` | Merged current state; a partial run updates only the tests it executed and flags the rest `stale` |
| `test-results/test-results.xml` | JUnit XML projection of the state file |
| `test-results/test-results.html` | Self-contained HTML dashboard |

Both folders are gitignored, so search tools will not find them. See
`troubleshooting.md` for how to list and read them.
