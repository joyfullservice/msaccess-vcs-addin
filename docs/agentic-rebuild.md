# Agentic add-in rebuild

How an MCP agent rebuilds `Version Control.accda` from `Version Control.accda.src`
without waiting on the user, and without closing any Access instance it does not
own. No new MCP tool: the existing `vcs_call_vba` entry point launches
`VCS.RebuildAddIn`, then the agent polls a status file after Access exits.

This is **not** `vcs_rebuild_database`, which rebuilds a *user* project from
source. This path rebuilds the add-in itself.

## Preconditions

- The helper script is enabled (`Use Worker Script` on the installer form).
- This Access process is the only `MSACCESS.EXE` in the current Windows session.
  The guard enumerates processes via WMI and **refuses** if the count of others
  is non-zero or if the query itself fails. It never quits or kills anything it
  finds.
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
string. Other running instances are listed in `otherInstances`.

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

## What this does not do

- It does not increment the add-in version or export source first. The agent
  edits source files directly; an export would overwrite that work.
- It is not `Deploy`, which copies the live development database rather than
  building from source.
- It does not close other people's Access windows. If another instance is
  running, it refuses and names the open databases so they can be closed.
