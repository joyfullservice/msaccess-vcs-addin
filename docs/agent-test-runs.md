# Running the add-in's own tests

How an agent runs the test suite that lives in this repository. Two things make
it different from testing a user database, and both have sent agents somewhere
else entirely: the run must be hosted on **this repo's development copy** of the
add-in, and it must go through the MCP server rather than the add-in's own window.

For writing a test, see [.cursor/rules/testing.mdc](../.cursor/rules/testing.mdc).
For the layers, the round-trip harness, and where results land, see
[testing-strategy.md](testing-strategy.md).

## Host the run on the development copy

```
vcs_run_tests("C:\Repos\msaccess-vcs-addin\Version Control.accda", "clsTestInstall")
```

MCP progress is best-effort in Cursor. For live per-test output, keep this
CLI command in the foreground:

```text
msaccess-vcs run-tests "C:\Repos\msaccess-vcs-addin\Version Control.accda" --filter clsTestInstall
```

The stream is pytest-style: dots for fast passes, a named line after a test
that took ≥ 1s, and full FAIL/ERROR/EMPTY lines. Assertion detail stays
in the TestRun log and in the `vcs_run_tests` MCP result. The CLI prints a
compact JSON summary (no `tests` map) and a last line such as
`Tests passed. 12 subs, 40 assertions in 1.48s`.

Headless means no add-in UI (no web runner, no console form, silent dialogs),
not a hidden Access window. The host instance stays visible so a dialog or a
VBA break is on screen.

`database_path` is the development copy in the repository — the `.accda` beside
`Version Control.accda.src`. The runner scans `CurrentVBProject`, so whichever
database hosts the run is the one whose tests are found: point a run at a user
database and you get that database's tests, reported as a clean pass because
nothing you were looking for was there to fail.

**Never host a run on the installed add-in.** The copy under `%AppData%\MSAccessVCS`
exists to be loaded as an add-in and nothing else; it is not to be opened as a
database. It also has no source tree beside it — its export folder holds only
`logs`, `mcp`, `tables`, and `test-results`, no `modules` and no `forms` — and a
good part of the suite reads real exported files or asks about the enclosing Git
repository. A run hosted there produces failures that say nothing about the code
under test, which is its own reason not to believe one.

`vcs_run_tests` opens the development copy when no instance has it open, `.accda`
included. The server binds a file moniker first, which Access only honours for a
database extension, and falls back to `OpenCurrentDatabase` when that bind fails.
`AutoRun` stands down when `Application.UserControl` is False, so a COM client can
open it without the install message box closing the instance or the installer form
stranding it.

## Nothing in `Testing\` hosts this

`Testing\Testing.accdb` and `Testing\Testing.accdb.src` are a sample database
used as a build and export integration fixture. They are not where the add-in's
tests live and not a host for running them — a run pointed there searches that
sample database for tests. `Testing\Fixtures\` is the object round-trip corpus
driven by `VCS.RunRoundtripTests`, which is a different entry point again. See
[Testing/AGENTS.md](../Testing/AGENTS.md).

Agents have historically arrived at `Testing.accdb` while looking for a database
to host a call from, after `VCS_API_REFUSED` made hosting look like the problem.
If a refusal says a call was dispatched to the installed add-in and arrived back
where it started, that is a defect in `modAPI` and no choice of host will move it
— see [Where that refusal came from](#where-that-refusal-came-from).

## Run through the MCP server, not the add-in's window

The runner singleton that records assertions lives in whichever project received
the `RunTests` call, while `modTestAssert.TestAssert` always routes through
`Application.Run` to the *installed* add-in path. Invoking `VCS.RunTests` from
inside a development copy puts those two in different projects: every assertion
is silently discarded, every test reports `EMPTY`, and the run looks clean while
proving nothing.

**Treat an all-`EMPTY` result as a broken harness, not a pass.**

## Guards that skip instead of failing

`modTestRepoDocs` and `modTestAgentDocs` read the repository working tree, which
they locate from `CodeProject.Path`. Where there is no `AGENTS.md` beside that path,
`RepoIsAvailable` returns False and every check in both modules reports a passing
note instead. That guard is there for the end user whose install has no checkout,
not as a configuration to run in: hosted on the development copy the checks are
live, and if they ever report skips you are running somewhere you should not be.
Line budgets from [agent-docs-maintenance.md](agent-docs-maintenance.md) are worth
confirming directly either way, since a count is cheaper than a run:

```powershell
Get-ChildItem 'AGENTS.md','.cursor\rules\*.mdc' |
    ForEach-Object { "{0,-24} {1}" -f $_.Name, (Get-Content $_.FullName).Count }
```

## Reaching an already-open instance

`vcs_run_vba` is the exception to the tool opening files for you. It attaches
through the Running Object Table without opening anything, so it reports
`Cannot find Access instance ... may have been closed` unless the file is already
open — for a `.accdb` as much as a `.accda`. Open it first, handing ownership to
the desktop so the instance outlives the launching script:

```powershell
$app = New-Object -ComObject Access.Application
$app.OpenCurrentDatabase("C:\Repos\msaccess-vcs-addin\Version Control.accda")
$app.Visible = $true
$app.UserControl = $true   # after opening, so AutoRun still sees automation
```

A rebuild ends with no Access process running, so reopen it before the next
iteration. See [agentic-rebuild.md](agentic-rebuild.md).

## Where that refusal came from

`modAPI.API` redirects to the installed add-in whenever the current database and
the running code are the same file, which is what carries a call made from the
development copy across to the library. Until August 2026 it did that without
checking where the call would land, so a call whose redirect target *was* the
running file re-entered the entry point while the outer call still held its
`Static IsRunning`, and came back `VCS_API_REFUSED`. The message named
`vcs_run_vba` nesting as the cause, which was wrong, and pointed at the choice of
host — which is how agents ended up hunting for some other database to call from.
`RedirectTargetIsSelf` now redirects only when the target is a different file, and
a self-dispatched refusal says so in its own words.

The lesson worth keeping: a refusal that says a call was dispatched to the
installed add-in and arrived back where it started is a defect in `modAPI`. No
choice of host database fixes one, and looking for a different host is what leads
into the two mistakes above.
