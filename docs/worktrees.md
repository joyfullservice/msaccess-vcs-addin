# Working in a git worktree

What an agent gets, and what it shares, when it builds and tests the add-in from a
second working copy of this repository: a `git worktree add` beside the repo, or a
worktree Cursor creates under `%USERPROFILE%\.cursor\worktrees\<repo>\<name>`. The
same questions apply to any Access project, and the MCP server's
`vcs_tips("worktrees")` topic answers them for user databases.

## What the worktree builds

A build writes the database beside the source folder it was given
(`GetOriginalDbFullPathFromSource`), so `vcs_rebuild_addin("<worktree>\Version
Control.accda.src")` builds `<worktree>\Version Control.accda`, not the main
checkout's copy. The change index, `logs/`, and `test-results/` are per worktree too.

A full build of the add-in takes 15 to 25 seconds, and each copy is about 20 MB. Each
full build also renames the previous file to `Version Control_VCSBackupN.accda`; they
pile up until someone deletes them, so clear them out of a long-lived worktree.

## What it does not get

Everything gitignored stays behind:

- The root `.env`. It holds `ACCESS_VCS_DATABASE` and the database-inspector paths,
  and in the main checkout they name the main checkout's files. Copied verbatim, they
  still do.
- `Version Control.accda.src\.env*` and any `.env` in a test database's export
  folder: the connection strings that `env:conn_*` references resolve against.
- Built databases: `Version Control.accda` and `Testing\Testing.accdb`, their
  `vcs-index.idx` change indexes, and the back-end or library files they link to.
  Build them from source, or copy them with `-CopyDatabase` (below).

`.cursor/worktrees.json` runs [tools/setup-worktree.ps1](../tools/setup-worktree.ps1)
when Cursor creates a worktree. For a worktree you made yourself, run it from the
worktree root with the main checkout's path:

```powershell
$env:ROOT_WORKTREE_PATH = 'C:\path\to\msaccess-vcs-addin'
.\tools\setup-worktree.ps1
```

It copies the `.env` files and rewrites any value that points into the main checkout
so it points into the worktree. It does not build: a build needs Access, and
`vcs_rebuild_addin` would install.

`-CopyDatabase` also copies each project's built database and its `vcs-index.idx`,
so the first export in the worktree is incremental, and the files the database links
to by a `rel:` path: Access, Excel, and text back ends, and VBA library references.
A file that a `.env` value names inside the main checkout is copied to the rewritten
path. Projects are found by their export folders (a `vcs-options.json` beside a
`vbe-project.json`). It skips, with a warning:

- a project whose `ExportFolder` is absolute, or climbs above the database
  (`\..\src\`), because the copy could not export into the worktree;
- a `rel:` path that leads outside the main checkout;
- a database with a `.laccdb` or `.ldb` lock file beside it, unless `-Force` is
  given, and even then the copy may be inconsistent.

## Export folders

A blank `ExportFolder` (`<database>.src` beside the database) or a relative one
(`\src\`) resolves from wherever the database is, so every copy exports into its
own working tree. An absolute one names one folder for every copy. The add-in warns
on every export when a database inside a git working tree exports outside it, and
the MCP server refuses build, export, and import there with
`export_folder_outside_working_tree`. Change the option in the main checkout; never
use `\..\src\`, because a build cannot find the database from it.

A build from a copy of a project whose `ExportFolder` is absolute (a clone, a
worktree, or a folder copied elsewhere) writes the database beside that copy and
skips importing the `VCS Source Path` and `VCS Build Path` properties, which name the
original's folders (`IsSourceFolderCopy`).

To export somewhere other than the configured folder, pass a folder to `Export`,
`FullExport`, `ExportVBA`, `ExportByType`, or `ExportObject`, or `output_dir` over
MCP (which requires it to be inside the working tree, when there is one). That
export is always full, writes no change index, and leaves the configured folder, its
index, the options, and the path properties untouched. A folder that
already holds a `vcs-index.idx` belongs to another project and is refused
(`export_folder_has_index`).

Always pass the worktree's own paths to MCP tools. Paths in docs, transcripts, and
work packets usually name the main checkout.

## What it shares with every other copy

**The installed add-in.** `vcs_rebuild_addin` builds the worktree's copy and then
installs it to `%APPDATA%\MSAccessVCS`. Every Access session on the machine loads
that file, so after a rebuild from a worktree:

- every export, import, and merge on any database, including the user's own work,
  runs the branch build;
- every test run's runner and `TestAssert` come from the branch build, whichever
  worktree hosts the run;
- the rebuild has closed every Access window the MCP server created that had the
  add-in loaded, which can be another agent's session.

So before rebuilding from a worktree, tell the user, and do not rebuild while another
agent is testing. When the branch is merged or dropped, rebuild from the main
checkout to put `dev` back. `vcs_rebuild_addin` reports `installed_from_worktree`
and names the main checkout when this happens.

**Machine-wide settings.** The add-in's options under
`HKCU\Software\VB and VBA Program Settings\MSAccessVCS`,
`%APPDATA%\MSAccessVCS\vcs-options.json`, and Access trusted locations. The folder
holding the worktree's `Version Control.accda` must be a trusted location, or the
silent install hangs on Access's untrusted-file prompt
([agentic-rebuild.md](agentic-rebuild.md#preconditions)). Cursor's worktree folder
usually is not one.

**Linked data.** The add-in's own databases link to nothing, but a test database or a
user project may. Each build logs a "Linked data sources" summary that says, for every
back end and library, whether it lies inside the database's folder or is shared with
every other copy. See [Linked data sources](#linked-data-sources).

**The MCP server.** One server process serves every Cursor window and every agent.
Calls on the same database run one at a time; calls on different databases run in
parallel, up to `ACCESS_VCS_MAX_PARALLEL` instances. A call that waits too long
returns `server_busy`; retry it.

## Several copies open at once

Measured on Access 365 (2026-11-27): two databases with the same file name and the
same VBA project name, in different folders, opened side by side in separate
instances. The file moniker bind, the Running Object Table scan, the Restart
Manager, and `Application.Run` inside each instance all resolved to the right copy
by full path. The MCP concurrency suite builds its three databases the same way
(`App.accdb` in three folders) and drives them in parallel.

Window titles are the one thing that does not tell copies apart. Access's default
title includes the full path, but a database that sets its own application title
(as this add-in does: "Version Control System") shows the same title in every
copy, so name a copy by its path or PID, never its title. Each agent needs its own
copy: two agents on one file share an Access window and overwrite each other's
imports.

## Running tests from a worktree

Host the run on the worktree's development copy:

```
vcs_run_tests("<worktree>\Version Control.accda", "clsTestOptions")
```

The installed runner scans whichever database hosts the run, so the tests exercise the
worktree's code even when the installed add-in came from somewhere else. Anything that
goes through `VCS.API` is redirected to the installed add-in, so it runs the installed
build, not the worktree's ([agent-test-runs.md](agent-test-runs.md)).

Changes to the headless export, merge, build, or test paths can collide when several
Access processes run them at once. After such a change, run the MCP repository's
multi-instance suite:
`pytest tests/test_concurrency_integration.py --run-concurrency -v`.

## Linked data sources

At the end of a full or merge build the log lists each distinct linked data source
once, with what kind of location it is:

| Kind | Meaning | Shared with other copies? |
|---|---|---|
| `project` | A file under the database's own folder | No; each copy needs its own |
| `outside` | A local file outside the database's folder | Yes |
| `network` | A UNC path or mapped network drive | Yes |
| `server` | An ODBC connection | Yes |
| `missing` | A file that does not exist | It was not copied, or the path is wrong |

`BuildHeadless` returns the same list as `dataSources`, and `vcs_rebuild_database`
returns it as `data_sources`. Before running code that writes data in a worktree,
check that nothing it writes to is shared, and ask the user when a shared source
might hold live data.
