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

## Interaction mode

`Operation.InteractionMode = eimSilent` routes `MsgBox2` prompts to the log and
returns their default answer. It is **sticky** — it is not reset by
`Operation.Begin` or `Operation.Finish`, so whoever sets it restores it. The
headless entry points set it before `Operation.Begin` (which raises its own modal
dialog when it refuses) and restore it on every path out.

Silent mode also suppresses the progress form, so a caller that wants a truly
headless build does not need to arrange anything beyond calling these methods.

## Related

- [agentic-rebuild.md](agentic-rebuild.md) — rebuilding the add-in itself, and the
  status-file protocol this borrows
- [testing-strategy.md](testing-strategy.md) — `RunTestsHeadless`, JUnit output
- [Wiki/Continuous-Integration.md](../Wiki/Continuous-Integration.md) — the
  user-facing version of this material
