# Using Version Control in Microsoft Access

In the process of developing a more complex Microsoft Access database, you may find yourself wondering what changed when, or how you are going to collaborate with other developers. The VCS add-in exports your Access database into individual files representing tables, queries, forms, and other objects so you can track changes in Git (or another VCS) and rebuild or update the database from those files.

## Export and Save

For some developers, the goal is simply to **track progress** over time. **Export** the project to source files, then commit those changes to GitHub, GitLab, or another system. [GitHub Desktop](https://desktop.github.com/) is a friendly UI if you are new to Git.

**Fast Save** (on by default) exports only objects that changed since the last export, using a local `vcs-index.idx` file paired with your database. Keep the index with the database file and exclude both from version control (default `.gitignore` template).

## Collaborative Development

Teams often work on separate copies of the database, **exporting** and **committing** changes to a shared repository. Changes are reviewed at the **source file** level, then combined by:

1. **Full build** — Replaces the database from source (recommended for release validation).
2. **[Merge build](Merge-Build)** — Imports only changed source files into the existing database (faster for day-to-day sync).

A merge build no longer requires a prior full build. As long as your database exports are reasonably current (so the `vcs-index.idx` reflects the database), you can run a merge build directly to pull in source changes. Keep the index paired with the database so the add-in can compare database vs source reliably.

When the same object changed in both the database and source, the add-in shows a **conflict** dialog so you can skip, overwrite from source, or delete as appropriate. Multi-file objects (for example queries with separate `.sql` and `.json` files) can show per-file diffs.

## Connection strings and secrets

Linked tables and pass-through queries often contain server names and credentials. With **Use .env For Connection Strings** enabled, exports replace secrets with `env:conn_*` placeholders; real values live in a gitignored `.env` file. See [Connections](Connections).

## Queries in version 5

New and upgraded projects typically export queries as a **`.sql` + `.json` pair** (deterministic export). See [Query Source Files](Query-Source-Files) and [Version 5 Overview](Version-5-Overview).

## Options

Settings are stored per project in `vcs-options.json` under your `.src` folder, with optional machine-wide defaults.

[Options reference](Options)

## Related pages

| Topic | Page |
|-------|------|
| First-time setup | [Quick Start](Quick-Start) |
| Partial updates | [Merge Build](Merge-Build) |
| Upgrading export format | [Version 5 Overview](Version-5-Overview) |
| Agent automation | [MCP and Automation](MCP-and-Automation) |
