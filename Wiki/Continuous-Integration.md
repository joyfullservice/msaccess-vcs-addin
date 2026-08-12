# Continuous Integration

Build a database from source, test it, and publish it — without anyone clicking
anything. This page covers the add-in side of that: the methods you call and what
they tell you.

The pipeline itself is yours to write. A full CI/CD product inside the add-in is
[out of scope](Project-Scope); what the add-in owes you is an API that reports
what happened clearly enough to automate against.

## Why the ordinary methods are not enough

`VCS.Build` and `VCS.MergeBuild` are the same calls the ribbon buttons make. They
open the progress window and start the build on a timer, then return right away —
before the build has done anything. There is no return value, so a script has no
way to tell a finished build from a failed one except by reading log files and
guessing.

`VCS.BuildHeadless` and `VCS.MergeHeadless` exist for scripts. They open no
window, they do not return until the build is over, and they return a JSON string
saying what happened.

## Calling them

From VBA in another database:

```vba
Dim strResult As String
strResult = VCS.BuildHeadless("C:\proj\MyApp.accdb.src\")
```

From PowerShell, through the add-in's public API function:

```powershell
$addin  = "$env:AppData\MSAccessVCS\Version Control.API"
$access = New-Object -ComObject Access.Application
$json   = $access.Run($addin, "BuildHeadless", "C:\proj\MyApp.accdb.src\")
$access.Quit()

$result = $json | ConvertFrom-Json
if (-not $result.success) {
    Write-Error "Build failed: $($result.error)"
    Write-Host  "Log: $($result.logPath)"
    exit 1
}
```

Omit the folder argument to use the export folder configured for the database
that is already open.

## What comes back

```json
{
  "success": true,
  "logPath": "C:\\proj\\MyApp.accdb.src\\logs\\Build_20260812_101500.log",
  "errorCount": 0,
  "warningCount": 2,
  "durationMs": 48210,
  "databasePath": "C:\\proj\\MyApp.accdb"
}
```

Check `success` and nothing else to decide pass or fail. `errorCount` and
`warningCount` are useful for a build summary; `logPath` is what to attach as an
artifact when something goes wrong. On failure an `error` field explains why.

A merge that cannot start at all — no database open, no source files in the
folder given, another operation still running — comes back with just `success`
and `error`, since there is no log yet to point at.

Merges run slower here than from the ribbon. The in-place merge optimization
(**Skip reopen before merge**) finishes on a background timer, which cannot report
a result, so headless merges always use the reopen path.

## Failing a build on your own criteria

A build can produce every object and still be broken in a way no compile catches:
a missing lookup table, a configuration record that did not come across, a linked
table pointing at the wrong server. The **Validate After Build** option lets you
decide.

Add a function to the database being built:

```vba
Public Function ValidateBuild() As Boolean
    If DCount("*", "tblConfig") = 0 Then Exit Function
    If Not CanReachBackEnd() Then Exit Function
    ValidateBuild = True
End Function
```

Then set the option in `vcs-options.json`:

```json
"ValidateAfterBuild": "ValidateBuild"
```

It runs last, after `RunAfterBuild` or `RunAfterMerge`. Returning `True` lets the
build succeed. Anything else fails it — `False`, an error, a missing procedure, or
a `Sub` with no return value — and the reason lands in the log and in the `error`
field. This is deliberately strict: a validation step that stays quiet when
something goes wrong is worse than no validation step.

The options form carries it as **Validate Build With**, under **Build Hooks**, so
it can also be set interactively rather than only in `vcs-options.json`.

## Running tests

`VCS.RunTestsHeadless` runs the test suite with no forms and no prompts, always
writing JUnit XML that CI can read:

```powershell
$json   = $access.Run($addin, "RunTestsHeadless", "-slow")
$result = $json | ConvertFrom-Json
if (-not $result.allPassed) { exit 1 }
```

The XML lands at `test-results\test-results.xml` under the export folder — publish
it as a test report artifact. See [Testing](Testing) for filters and tags.

## Installing the add-in on a runner

A fresh runner needs the add-in installed before any of this works:

```powershell
& $msaccess "$downloaded\Version Control.accda" /cmd "INSTALL SILENT C:\runner\install-status.json"
```

The status file is written twice — `installing` when the install starts, then
`complete` or `install-failed`. Poll it after Access exits.

**If the file never appears, the add-in's code never ran.** That almost always
means the downloaded `.accda` is not in an Access
[trusted location](Security-Considerations), so Access opened it with macros
disabled and is waiting behind a prompt nobody will answer. Add the download
folder to the Trust Center on the runner, once, as part of provisioning.

## Things that will bite you

- **Access must run in an interactive session.** As a Windows service (Session 0)
  it has no desktop, and any dialog that does slip through hangs forever with no
  way to see it. Configure the runner to log on as a user.
- **Only one build at a time.** The add-in refuses a second operation while one
  is running. Do not run parallel jobs against one runner.
- **Access processes outlive `Quit`.** `MSACCESS.EXE` often stays resident after
  automation releases it. Record the process ID your script started and clean up
  that one; never kill every `MSACCESS.EXE`, which would take out anything else
  running on the machine.
- **`/decompile` and Compact & Repair both discard compiled VBA.** If your
  pipeline does either, compile again afterwards before you ship the file.

## Further reading

- [MCP and Automation](MCP-and-Automation) — agent-driven automation and permissions
- [Testing](Testing) — writing and filtering tests
- [Merge Build](Merge-Build) — what a merge does and when to use one
- [Options](Options) — every setting in `vcs-options.json`
- [Security Considerations](Security-Considerations) — trusted locations and Trust Center
