# Agentic add-in rebuild

How an MCP agent rebuilds `Version Control.accda` from `Version Control.accda.src`
without waiting on the user. No new MCP tool: the existing `vcs_call_vba` entry
point launches `VCS.RebuildAddIn`, then the agent polls a status file after
Access exits. The one Access process it ever closes is its own.

This is **not** `vcs_rebuild_database`, which rebuilds a *user* project from
source. This path rebuilds the add-in itself.

## Preconditions

- The helper script is enabled (`Use Worker Script` on the installer form).
- No other `MSACCESS.EXE` in this Windows session holds a file the rebuild must
  replace — the installed add-in or the build target. Another instance with an
  unrelated database open does not block anything, because only a loaded VBA
  project keeps a file open. The guard **refuses** when one does hold such a
  file, when an instance cannot be asked, or when the process query itself fails.
  It never quits or terminates another process — see
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
string. Every other running instance is listed in `otherInstances`, one line
each, whether or not it was the one that blocked:

```
PID 37340: C:\Repos\Testing.accdb (visible, responded to automation, holds C:\USERS\ME\APPDATA\ROAMING\MSACCESSVCS\VERSION CONTROL.ACCDA)
PID 31002: no database open (hidden, responded to automation, holds no rebuild file)
PID 8820: open database unknown (visible, did not respond to automation, loaded projects unknown)
```

The last clause is the one that decides the refusal. `holds <path>` blocks it.
`loaded projects unknown` also blocks it, because the instance never answered and
silence must not be read as safe. `holds no rebuild file` does not block, so an
instance reported that way needs no attention. Close whatever blocked and call
again.

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

Closing another process is never necessary, because the guard does not need the
instance gone — it needs the *files* free. Only a loaded VBA project holds one
open, so `ClassifyAccessInstance` reads `VBE.VBProjects` in the other instance and
compares each `FileName` against the installed add-in and the build target. An
instance that never loaded the add-in locks nothing no matter how long it has been
running, which is why an unrelated database open next door does not block a
rebuild.

Do not substitute the command line for this. `vcs_*` tools open databases through
COM `OpenCurrentDatabase`, which leaves no argument behind, and the reverse error
also occurs: an instance launched pointing at a database it never managed to open
was reported as having that file open, because the fallback names an argument
rather than a loaded file. `VBProjects` reflects what is actually loaded.

Measured on Access 16.0:

- `AccessibleObjectFromWindow` with `OBJID_NATIVEOM` against Access's `OMain`
  window does reach a foreign instance's object model. Every instance probed
  returned `Application.Version`.
- `VBProjects` discriminates cleanly. The same instance reported one project
  before the add-in was invoked and two after, the second being the installed
  add-in path — reported in upper case, which is why the comparison is
  case-insensitive.
- A rebuild ran to completion with an unrelated Access instance open the whole
  time, confirming that such an instance holds no lock on the add-in.

Two answers still cannot be trusted, and both keep refusing. An instance that
never responds is indistinguishable from an idle one, so silence blocks. And
reading `VBE` can be refused by the target instance's "Trust access to the VBA
project object model" setting, which reports as `loaded projects unknown` rather
than as nothing loaded. Since a wrong guess would terminate live work, the guard
classifies read-only and reports either way.

The check is a precondition, not a guarantee: an instance could load the add-in
between the check and the file replace. The install then fails and the status file
says so.

## What this does not do

- It does not increment the add-in version or export source first. The agent
  edits source files directly; an export would overwrite that work.
- It is not `Deploy`, which copies the live development database rather than
  building from source.
- It does not close, quit, or terminate any Access process. If another one holds
  a file it needs, it refuses and names both the process and the file.
