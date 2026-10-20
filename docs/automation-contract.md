# Automation contract

What an unattended caller can rely on from the add-in. A CI pipeline is out of
scope for this project ([Project Scope](../Wiki/Project-Scope.md)); this file
defines the surface such a pipeline is built against, so it can be relied on
without reading the implementation.

The tests that guard this surface are in
`Version Control.accda.src/modules/Tests/Core/modTestHeadlessBuild.bas`.

## Two tiers of entry point

The `VCS` object has grown two kinds of build entry point, and the difference
matters more than the names suggest.

| | `Build` / `MergeBuild` | `BuildHeadless` / `MergeHeadless` |
|---|---|---|
| Shape | `Sub` | `Function` returning JSON |
| Progress form | opens `frmVCSMain` | none |
| Returns | when the build is *scheduled* | when the build is *finished* |
| Failure signal | none | `success: false` plus a reason |

`Build` opens the form, then hands the work to `SetTimer "Build"` and returns
immediately. Nothing about the outcome is available to the caller — which is
why external builders written against it resort to polling log files or
counting objects. Use the headless pair for anything unattended.

`RunTestsHeadless` is the same idea for the test suite, and predates these; see
[testing-strategy.md](testing-strategy.md).

## VBE Error Trapping

`BuildHeadless`, `MergeHeadless`, and `RunTestsHeadless` (including test runs
that API or MCP force onto the headless path) raise VBE **Error Trapping** to
**Break on Unhandled Errors** for the whole call, including preflight that runs
before an operation begins. A caller already at that mode is left alone. The
original setting is restored when the call returns, including on a refused
preflight. Attended ribbon and non-headless API/MCP operations still use
**Break in Class Module** as their floor and do not downgrade a more
permissive caller setting.

## Result shape

```json
{
  "success": true,
  "logPath": "C:\\proj\\source\\logs\\Build_20260812_101500.log",
  "errorCount": 0,
  "warningCount": 2,
  "durationMs": 48210,
  "databasePath": "C:\\proj\\MyApp.accdb"
}
```

A failure adds `error` (always) and `errorNumber` (only when a VBA runtime error
escaped, rather than the build deciding it had failed):

```json
{
  "success": false,
  "error": "Build did not complete successfully. Check the log for details.",
  "logPath": "C:\\proj\\source\\logs\\Build_20260812_101500.log",
  "errorCount": 3,
  "warningCount": 0,
  "durationMs": 12044,
  "databasePath": "C:\\proj\\MyApp.accdb"
}
```

A request refused before the build started carries only `success` and `error` —
there is no log to point at. Refusals: building the add-in from inside itself,
merging with no database open, a source folder holding no `vcs-options.json`,
and another operation already running.

`success: false` with `errorCount: 0` is possible and is not a contradiction. It
means the build was cancelled or abandoned rather than erroring — for example a
critical condition that stopped the run before anything was logged as an error.
Treat `success` as the verdict and the counts as detail.

## Why the log path is an out-param

`modBuild.Build` finishes by calling `Operation.Finish`, which calls
`modObjects.ReleaseObjects` and drops the `Log`, `Perf`, and `Options`
singletons. By the time control returns to the caller, `Log.SavedLogFilePath` is
an empty string on a freshly constructed object.

So the values in the result are captured inside `Build`'s `CleanUp` block,
through the optional `dOutcome As Dictionary` parameter, before `Operation.Finish`
runs. Any new field in the result shape has to be collected there too, not
computed afterwards. `LoadSingleObject` solves the same problem the same way with
its `strSavedLogPath` out-param.

## Headless merges take the slow path

`Options.SkipReopenBeforeMerge` lets a merge prepare the database in place rather
than closing and reopening it, which is much faster on a large project. It works
by handing off to `SetTimer "MergeReset"` and returning — the merge then resumes
on a fresh call stack.

That is incompatible with returning a result, because the call that would return
it has already unwound. `modBuild.Build` therefore ignores the option when
`dOutcome` was supplied, and takes the reopen path instead. A headless merge is
slower than the same merge from the ribbon, by design.

## Hook ordering

Hooks are named procedures in the *target* database, run through
`RunProcInCurrentProject`. They must be `Public` in a standard module and take no
parameters.

Full build:

1. `RunBeforeBuild` — before any object is imported, in the new blank database
2. (import)
3. `RunAfterBuild`
4. `ValidateAfterBuild`

Merge build:

1. `RunBeforeMerge` — after the database is prepared, before source is merged
2. (merge)
3. `RunAfterMerge`
4. `ValidateAfterBuild`

`ValidateAfterBuild` differs from the others in one way that matters: **it gates
success**. The `RunAfter*` hooks log an error if they raise one, but the build
still reports success. Validation is for the case where a build produced every
object and still is not shippable, and a pipeline needs to hear about it before
the file reaches users.

It must be a `Function` returning `True`. Anything else fails the build: `False`,
a raised error, a missing procedure, no return value (a `Sub`), or a return that
will not coerce to a Boolean. Silence is not approval when the answer decides
whether a release ships.

```vba
Public Function ValidateBuild() As Boolean
    ' Cheap smoke checks that a compile cannot catch.
    If DCount("*", "MSysObjects", "Type = -32768") = 0 Then Exit Function
    If Not IsAppConfigLoadable() Then Exit Function
    ValidateBuild = True
End Function
```

Set the option in `vcs-options.json` (`"ValidateAfterBuild": "ValidateBuild"`),
through `VCS.Options`, or on the options form as **Validate Build With** under
**Build Hooks**.

## Silent install

`/cmd "INSTALL SILENT"` installs without prompting. A runner installing a
downloaded release should name a status file, because nothing in a fresh
environment wrote the registry value that an add-in rebuilding itself relies on:

```
MSACCESS.EXE "Version Control.accda" /cmd "INSTALL SILENT C:\runner\install-status.json"
```

The path is everything after the keyword — quotes are optional and stripped,
and interior spaces are preserved, so no inner quoting is needed. The file is
written twice: `installing` as soon as the add-in's code starts, then `complete`
or `install-failed` with an `error`.

That first write is the useful one. **No file at all** means VBA never ran, which
in practice means the `.accda` was not in an Access trusted location and Access
is sitting behind a security prompt with nobody to dismiss it. Nothing inside the
add-in can report that condition, because its own code is what did not run.
Trusting the folder is a precondition of unattended install, and an absent status
file is how a pipeline detects that it was not met.

## Prompts

An unattended caller is never prompted: an API or MCP call, or a headless entry point
(`ForceUnattended`). That holds in normal interaction mode as well as silent, so a
caller no longer has to call `SetInteractionMode` to keep a `MsgBox2` from blocking
it. The prompt goes to the log, marked `[**MessageBox Not Displayed**]`, and
`MsgBox2` returns its `intDefaultResult`. Not covered:

- dialogs that do not go through `MsgBox2`: VBA runtime errors, Access's own prompts,
  and the add-in's folder and file pickers (`BuildAs`, and `Build` when the open
  database has no source folder);
- a root-less API call made while another caller's attended root is live, such as the
  options form, which reads that root's `Attended`;
- a person at a form that was open, or opened by an API call, when the last caller
  was automated: `Source` is reset only by the ribbon and the Add-Ins menu.

Every question passes that answer explicitly, and
`modTestPromptDefaults.TestEveryQuestionPassesAnUnattendedAnswer` fails on one that
does not. The answers are the ones silent mode already gave, so a caller that was
suppressed before sees no change. The ones an automation caller is most likely to
meet:

| Question | Answer | Effect |
|---|---|---|
| *Could not import* (`LoadComponentFromText`) | Abort | the build or import fails |
| *Connection Failed* (`HandleConnectionFailure`) | Abort | the build fails |
| *Newer VCS Version Detected*, build or merge | Yes | the build goes on |
| *Newer VCS Version Detected*, export | No | the export stops and fails |
| Database file name differs from the source (full build) | OK | builds under the source's name |
| *Save Connection Settings?* | No | nothing saved to `.env` |

A suppressed question that ends an operation reports it:

- `BuildHeadless` and `MergeHeadless` return `success: false`, with the log path when
  the log was started.
- `Export`, `FullExport` and `ExportVBA` return `True` only for an export that ran to
  the end. An export stopped before it began, by a newer add-in version in the
  project's options or by a locked project, saves its log and leaves the options and
  the index as they were.
- An export or a build that ends at a critical error completes its root as
  `eorFailed`, so an MCP caller gets an `error` callback with the log path rather
  than `complete`.
- `MigrateDebugAssert` and `StandardizeLetterCasing` return `False` when their
  confirmation was answered for nobody.
- `Build` and `MergeBuild` still return when the build is scheduled; their outcome
  reaches an MCP caller through the callback only. Use the headless pair for a result.
  An `APIAsync` call whose method never begins a root, such as `MergeBuild` before any
  full build, posts no callback at all.

## Operation ownership and interaction

Root work is owned by a **root lease** (`Operation.TryBeginRoot` →
`clsRootOperationLease.Complete`) or by the synchronous pair **`Operation.Begin` /
`Operation.Finish`**. Both create the same root; the difference is what the caller
holds and what guarantees apply.

| Use a **lease** when | Use **Begin / Finish** when |
|---|---|
| Ownership crosses a timer continuation (`DetachForContinuation`, `ResumeRoot`) | Work begins and ends in one procedure on the same stack |
| A module holds the root across an async boundary (web test bridge) | The caller is the only code that will complete the root |
| The site may not own the root at all (conditional harness entry) | No token needs to survive a form unload or timer tick |

`Begin` wraps `TryBeginRoot` and discards the lease handle. **`Finish` supplies the
live root token to itself**, so it performs no ownership check — only call it from the
procedure that called `Begin`. A lease's `Complete` validates the token, and
`Class_Terminate` calls `AbandonRootLease` if the lease is dropped without completing,
which is load-bearing for paths such as `ExecuteTests` that have no error handler
around the run body.

Core routines such as `modBuild.Build` and `LoadSingleObject` perform work only; the
orchestration boundary that acquired the lease completes it. Timer continuations carry
an opaque **root token** through `SetTimer` and resume with `Operation.ResumeRoot` /
`RunBuildFromContinuation`.

There is no nested operation object. A routine that can run either standalone or
inside a larger operation tests for the enclosing one and, when it finds it, owns
nothing: it does not take a lease, does not clear the log or console the root is
writing to, and completes nothing on the way out. `modTestRoundtrip` and
`modTestQuerySqlBuilder` both do this, keyed on `Operation.OperationType =
eotTestRun`.

Work that hands control to foreign code — user VBA through `Application.Run`, or
closing a form whose unload handler would read a running operation as a cancel —
takes a **pause** (`Operation.TryPause` → `clsOperationPause.ResumePause`). A pause
suspends the root on the same stack that started it and issues no continuation
token; pauses nest, and only the outermost resume restores the root.

**Effective interaction mode** belongs to the root and only ratchets tighter
(`eimSilent` wins over `eimNormal`). While a root is active a looser mode is
ignored, so nothing deeper in an operation can revoke silence a caller asked for;
between roots the value is only the default for the next root, so relaxing it is
allowed. Completing a root resets it, which is why headless callers
no longer restore it by hand. **`Attended`** is a root capability (`Not
(AutomationSource Or ForceUnattended)`) captured once at root creation and immutable
for the life of the operation, so automation and headless entry points must set
`Operation.ForceUnattended = True` *before* `TryBeginRoot`.
`Operation.TryBeginUnattendedRoot` does both, and puts the flag back when the root
is refused: no root of that call will finish and clear it.
Outside a live root, `Attended` describes the incoming caller with the same rule,
since the value the last root captured belongs to the previous operation. The
interactive entry points, the ribbon through `HandleRibbonCommand` and the Add-Ins
menu items, call `Operation.ResetForInteractiveCommand`, which resets `Source` and,
when no root is live, clears `ForceUnattended`.

`PromptWouldDisplay` decides whether a prompt blocks or is logged; `MsgBox2` asks it,
and so does `modConnect.CacheConnection` before a probe that can raise the ODBC
driver dialog. The policy itself is `Operation.AllowsPrompt`. An unattended caller
is never prompted, whatever the interaction mode: `MsgBox2` logs the prompt and
returns its `intDefaultResult` (see [Prompts](#prompts)). An attended caller is
prompted unless the operation is silent. `blnUserGesturePrompt` lets a prompt through
silent mode when the root or the caller acting now is attended, so a user-initiated
cancel confirmation can appear during an attended silent test run, and a refused
ribbon command can say so while an unattended root runs in normal mode, such as an
MCP export. A refused root logs its refusal instead of prompting whenever the
caller is automated or explicitly unattended.

Interaction mode reaches only as far as the operation instance does, which is one
VBA project. A test run drives a *second* project, so the mode the driver set is
invisible there; `modTestAssert.TestRunActive`, pushed by `clsTestRunner`, carries
the guarantee across and suppresses every prompt while it is set. See
[testing-strategy.md](testing-strategy.md) for how a run spans two projects.

Everything that reaches outside the operation instance — registry state for crash
recovery, the MCP completion callback, VBE error trapping, and `ReleaseObjects` — is
gated on being the instance published by `modObjects`, and only that instance adopts
restored state from the registry. A `New clsOperation` is therefore inert with
respect to the session, which is what makes the lifecycle testable: see
`modTestOperationLifecycle`, which drives private instances rather than the root
that is running the suite.

### Known gaps (not fixed by this refactor)

Two cross-form pairs still use `Begin` in one module and `Finish` in a form unload
handler, with no lease token carried across:

- `clsVersionControl.ShowOptions` → `frmVCSOptions.Form_Close`
- `clsVersionControl.SplitFiles` → `frmVCSSplitFiles.Form_Unload`

`Finish` ends whatever root is active; these are safe only because `Begin` refuses
when another operation is running and each form is the expected owner.

The in-place merge timer chain detaches via `Operation.CurrentRootToken` at
`modBuild.bas` because `Build` deliberately holds no lease object. The token is still
validated on resume through `ResumeRoot`.

## Related

- [agentic-rebuild.md](agentic-rebuild.md) — rebuilding the add-in itself, and the
  status-file protocol this borrows
- [testing-strategy.md](testing-strategy.md) — `RunTestsHeadless`, JUnit output
- [Wiki/Continuous-Integration.md](../Wiki/Continuous-Integration.md) — the
  user-facing version of this material
