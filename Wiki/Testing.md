# Testing

The add-in uses **three layers** of tests. Contributors should run the layers relevant to their change before opening a pull request.

---

## Layer 1 — Unit / logic tests (`VCS.RunTests`)

Hundreds of assertions across `modTest*` modules (encoding, JSON, hashing, conflicts, query builder, etc.).

### Run from Access

Open the add-in or development build, then in the Immediate Window:

```vba
?VCS.RunTests
```

Filter examples:

```vba
?VCS.RunTests("modTestEncoding")
?VCS.RunTests("SQL", "-slow")
?VCS.RunTests("TestParseJoinExpression")
```

Tags use `'@Tag("name")` in module or procedure headers. Prefix `-` to exclude.

### Ribbon

Set **Default Test Filter** under **Options** → **Advanced**, then click **Run Tests** on the ribbon. Leave blank to run all tests.

![Run Tests Button](img/ribbon-run-tests.png)

### Output

- Progress in `frmVCSMain` console
- JSON summary and `TestRun_*.log` under the add-in `logs/` folder

When the run finishes, the console shows a summary of passed, failed, and skipped tests along with timing:

![Completed test run summary](img/tests-complete.png)

---

## Layer 2 — Object round-trip (`VCS.RunRoundtripTests`)

Imports each fixture, exports twice, checks idempotency and drift. **Queries** are fully covered today; other object types follow the same harness pattern.

See [Regression Testing](Regression-Testing) for fixtures, rebaseline mode, and contribution workflow.

```vba
?VCS.RunRoundtripTests
?VCS.RunRoundtripTests("C:\path\to\fixtures\", True)  ' rebaseline — review diff!
```

---

## Layer 3 — Integration database

[`Testing.accdb.src`](https://github.com/joyfullservice/msaccess-vcs-addin/tree/dev/Testing/Testing.accdb.src) in the repository — full build/export scenarios for the add-in itself and sample projects.

Use after large import/export or build pipeline changes.

---

## MCP / agents

When **Allow Arbitrary VBA Execution** is enabled:

```
vcs_run_vba(<addin-path>, "MCP_TempFunction = VCS.RunTests(""SQL"", ""-slow"")")
```

See [MCP and Automation](MCP-and-Automation).

---

## PR expectations

| Change type | Minimum testing |
|-------------|-----------------|
| Options / UI copy | Manual smoke export |
| Export/import logic | `RunTests` + affected `RunRoundtripTests` |
| Query parser | `RunRoundtripTests` on `Testing/Fixtures/queries/` |
| Build/merge | Integration build + targeted unit tests |

---

## Related

- [Editing and Contributing](Editing-and-Contributing)
- [Regression Testing](Regression-Testing)
- Repository [`AGENTS.md`](https://github.com/joyfullservice/msaccess-vcs-addin/blob/dev/AGENTS.md)
