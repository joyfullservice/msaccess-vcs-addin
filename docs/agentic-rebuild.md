# Agentic add-in rebuild

How an MCP agent rebuilds `Version Control.accda` from `Version Control.accda.src`
without waiting on the user. No new MCP tool: the existing `vcs_call_vba` entry
point launches `VCS.RebuildAddIn`, then the agent polls a status file after
Access exits. The one Access process it ever closes is its own.

This is **not** `vcs_rebuild_database`, which rebuilds a *user* project from
source. This path rebuilds the add-in itself.

## Preconditions

- The helper script is enabled (`Use Worker Script` on the installer form).
- This Access process is the only `MSACCESS.EXE` in the current Windows session.
  The guard **refuses** otherwise, and also refuses if the process query itself
  fails. It never quits or terminates another process — see
  [Why it refuses instead of closing them](#why-it-refuses-instead-of-closing-them).
- The folder that contains the repo `Version Control.accda` (the parent of the
  source folder) is an Access trusted location. Otherwise `/cmd INSTALL` hangs
  on autoexec's native untrusted-file `MsgBox`.
- Pass the source folder when it cannot be derived from the current database or
  the saved `Install\Source Path` registry value. Unattended callers never see
  a folder picker.

## Call

```
vcs_call_vba(database_path, "VCS.API", ["RebuildAddIn", "<source folder>"])
```

`database_path` is whatever database the MCP session already has open (often
`Testing.accdb`). The JSON result looks like:

```json
{
  "success": true,
  "status": "launched",
  "statusFile": "C:\\Repos\\msaccess-vcs-addin\\Version Control.accda.src\\logs\\rebuild-status.json",
  "logFolder": "C:\\Repos\\msaccess-vcs-addin\\Version Control.accda.src\\logs"
}
```

A refusal returns `"success": false`, `"status": "refused"`, and an `error`
string. Other running instances are listed in `otherInstances`, one line each:

```
PID 27184: C:\Repos\Testing.accdb (hidden, responded to automation)
PID 31002: no database open (hidden, responded to automation)
PID 8820: open database unknown (visible, did not respond to automation)
```

`open database unknown` means the instance never answered, so nothing is known
about it — not that it is empty. Close the listed processes and call again.

The worker then quits this Access instance. A COM error on that same
`vcs_call_vba` call is expected. `vcs_call_vba` has no timeout; the worker
sleeps a few seconds before quitting so the JSON can return first.

## Status file

`<source folder>\logs\rebuild-status.json` is gitignored with the rest of
`logs/`. The agent already knows the source folder, so it does not need the
Access instance to stay alive. Poll it with the Read tool.

| `status` | Meaning |
|---|---|
| `launched` | Worker started; Access is about to quit |
| `building` | New Access instance is building from source |
| `compiling` | Compile gate on the rebuilt project |
| `installing` | `/cmd INSTALL SILENT` is running |
| `complete` | Installed add-in file is newer than when install started |
| `refused` | Guard or preflight failed; nothing was rebuilt |
| `build-failed` | Build timed out or the build log reported failure |
| `compile-failed` | Rebuilt project does not compile; Access is left open on that file |
| `install-failed` | Silent install aborted, or the installed file was not updated |

Other fields: `error`, `buildLog` (path of the build log when one exists),
`phaseStarted`, `updated`. Terminal states are `complete` and the `*-failed` /
`refused` values.

After `complete`, later MCP calls against a user database spawn a fresh Access
and load the newly installed add-in. Re-run the verification that needed the
rebuild.

## Running the add-in's own tests

The add-in's tests only run when the add-in itself is the current database,
because `TestRunner.Scan` walks `CurrentVBProject`. Pointing a test run at a user
database finds that database's tests instead.

```
vcs_run_tests("C:\Repos\msaccess-vcs-addin\Version Control.accda", "clsTestInstall")
```

`AutoRun` stands down when `Application.UserControl` is False, so a COM client can
open either copy without the install message box closing the instance or the
installer form stranding it.

The add-in must already be open in an Access instance before the call. The MCP
server attaches to a running instance through the Running Object Table and does not
open a database itself, so it reports `Cannot find Access instance ... may have been
closed` when nothing has that file open. Open it, then hand ownership to the desktop
so the instance outlives the launching script:

```powershell
$app = New-Object -ComObject Access.Application
$app.OpenCurrentDatabase("C:\Repos\msaccess-vcs-addin\Version Control.accda")
$app.Visible = $true
$app.UserControl = $true   # after opening, so AutoRun still sees automation
```

A rebuild ends with no Access process running, so reopen it before the next
iteration.

Run the tests **through the MCP server**, not from the add-in's own window. The
runner singleton that records assertions lives in whichever project received the
`RunTests` call, while `modTestAssert.TestAssert` always routes through
`Application.Run` to the *installed* add-in path. Invoking `VCS.RunTests` from
inside a development copy puts those two in different projects: every assertion
is silently discarded, every test reports `EMPTY`, and the run looks clean while
proving nothing. Treat an all-`EMPTY` result as a broken harness, not a pass.

## Why it refuses instead of closing them

A hidden Access instance with nothing open is almost certainly a leftover
automation process, and closing it would let more rebuilds through. The guard
still refuses, because it cannot reliably tell that case apart from the one that
must not be touched.

Whether a database is open cannot be read from the command line: `vcs_*` tools
open databases through COM `OpenCurrentDatabase`, which leaves no argument
behind. The only way to ask is to reach the other process's object model. A
*busy* instance rejects those calls, so an instance in the middle of a long
export answers exactly like an empty one.

Measured against real instances on Access 16.0, two of the three pieces of
evidence behave as hoped and one does not:

- `AccessibleObjectFromWindow` with `OBJID_NATIVEOM` against Access's `OMain`
  window does reach a foreign instance's object model. Every instance probed
  returned `Application.Version`, so the mechanism itself is no longer in doubt.
- Visibility and PID enumeration were accurate.
- The open-database answer is **not** trustworthy. An instance launched with a
  database path it never managed to open was reported as having that file open,
  because the command-line fallback names an argument rather than an open
  database. The error runs toward "occupied", which is the safe direction for
  refusing and the wrong direction for closing.

Still unmeasured is the case the whole decision rests on: an instance genuinely
busy inside a long operation. Until something distinguishes that from an idle
one, closing on this evidence would risk terminating live work, so the guard
classifies read-only and reports.

## What this does not do

- It does not increment the add-in version or export source first. The agent
  edits source files directly; an export would overwrite that work.
- It is not `Deploy`, which copies the live development database rather than
  building from source.
- It does not close, quit, or terminate any Access process. If another one
  exists, it refuses and names it.
