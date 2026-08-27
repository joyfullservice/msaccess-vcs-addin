# Agentic add-in rebuild

How an MCP agent rebuilds `Version Control.accda` from `Version Control.accda.src`
without waiting on the user. Prefer `vcs_rebuild_addin(source_dir)`: it launches
`VCS.RebuildAddIn`, streams the builder's existing HTTP log/progress callbacks,
while the MCP server internally watches `rebuild-status.json` through compile
and install. The calling agent waits for the tool or CLI process, not the status
file. `vcs_call_vba` remains a launch-only escape hatch. The status file is also
durable recovery state after a timeout. The one Access process the rebuild ever
closes is its own.

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
vcs_rebuild_addin("<source folder>")
```

The MCP server derives the development copy beside that folder, launches
`RebuildAddIn`, and registers a callback operation. The callback URL and
operation ID pass through `Worker.vbs` to the builder Access process, which
invokes the same `APIAsync(..., "Build", source)` entry point as an ordinary
database build. `Log.Add` and `Log.Progress` therefore provide detailed output
without a second logging implementation. Internally, the tool uses the status
file for the later compile/install phases and terminal result. For guaranteed
live output in a terminal (Cursor often shows only "Running..."):

```
msaccess-vcs rebuild-addin "<source folder>"
```

Keep the CLI in the foreground so its stream stays in the primary chat. It
exits when the operation reaches terminal status; that process exit is the
completion signal. Do not background it just to wait on a notification, and
do not add a second timer wait, fixed-duration sleep, or independently poll
`rebuild-status.json` after it has already finished. The MCP tool already
watches the status file internally.

The launch-only escape hatch is still:

```
vcs_call_vba(database_path, "VCS.API", ["RebuildAddIn", "<source folder>"])
```

`database_path` only decides which Access instance hosts the call; what gets
rebuilt is the source folder argument. **Host it on the development copy of the
add-in** — the `Version Control.accda` in the repository root, beside the source
folder. That host holds the build target, so it closes itself once the handoff is
confirmed; nothing refuses the call for that reason.

Do not host it on the installed copy under `%AppData%`, which is only ever loaded
as an add-in and never opened as a database, and do not open some unrelated
database — a user project or `Testing\Testing.accdb` — to satisfy the parameter.
Rebuilding the add-in is a repository operation and belongs to the repository's
own copy. The JSON result looks like:

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

`"status": "launch-failed"` means the worker process was created but never got
far enough to report for duty. Nothing was rebuilt and the host Access instance
is still open, so this is safe to retry once the cause is fixed — most often the
helper script missing from the add-in folder. `RebuildAddIn` waits up to 15
seconds for the worker to write its first status before deciding this, so an
answer either way arrives within that window and a `launched` result is real.

Access then exits a few seconds later, once the launch result has had time to
reach the caller. The callback server remains alive; the builder Access process
posts detailed messages to it until the build phase completes.
`vcs_rebuild_addin` then continues waiting on the status file while the worker
compiles and installs. `vcs_call_vba` does not wait; a COM error on that
launch-only call is possible if the timing is tight.

## Status file

`<source folder>\logs\rebuild-status.json` is gitignored with the rest of
`logs/`. `vcs_rebuild_addin` watches it with directory-change notifications.
Read it yourself only to recover after a client timeout, a stalled wait, or a
launch-only `vcs_call_vba` call. Identify *your* attempt by `phaseStarted`.

| `status` | Meaning |
|---|---|
| `starting` | Attempt claimed the file; preflight has not finished |
| `launched` | Worker launched; not yet confirmed running |
| `building` | Worker reported for duty; a new Access instance is building from source |
| `compiling` | Compile gate on the rebuilt project |
| `installing` | `/cmd INSTALL SILENT` is running |
| `complete` | Installed add-in file is newer than when install started |
| `refused` | Guard or preflight failed; nothing was rebuilt |
| `launch-failed` | Worker never started; nothing was rebuilt and Access stayed open |
| `build-failed` | Build timed out, the build log reported failure, or the host instance never exited |
| `compile-failed` | Rebuilt project does not compile; Access is left open on that file |
| `install-failed` | Silent install aborted, or the installed file was not updated |

Other fields: `error`, `buildLog` (path of the build log when one exists),
`phaseStarted`, `updated`. Terminal states are `complete` and the `*-failed` /
`refused` values.

Every attempt that gets as far as an existing source folder stamps `starting`
with a fresh `phaseStarted` before running any check that could refuse, and
records its own `refused` verdict if one comes. `phaseStarted` then holds still
for the rest of the run, so it identifies the attempt: the call returns the same
value it wrote, and a record carrying a different one belongs to somebody else's
run. A refusal reached before that — no source folder, or none that exists —
leaves no record, and there is no folder to watch in that case either.

A `refused` or `launch-failed` call returns that verdict in its own JSON, so
there is nothing to wait for. Only a `launched` result has a watch phase.

**A file that stops changing is not the same as a build in progress.** Compare
`updated` against the clock before waiting any longer, and check
`Get-Process MSACCESS,wscript`: a rebuild that is genuinely running always has at
least one of them. Neither present, with a non-terminal status, means the run
died without being able to say so — capture the state and treat it as a bug in
this path rather than retrying blindly.

After `complete`, later MCP calls against a user database spawn a fresh Access
and load the newly installed add-in. Re-run the verification that needed the
rebuild.

## Verifying the rebuild

A rebuild ends with no Access process running, so the next run reopens the file.
Running the add-in's own test suite afterwards is its own topic, including which
database has to host the run and why an all-`EMPTY` result is a broken harness
rather than a pass: see [agent-test-runs.md](agent-test-runs.md).

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
