# Query Source Files

From export format **5.0** with **Use Deterministic Query Export** enabled (default), each query is exported as two files:

| File | Role |
|------|------|
| `MyQuery.sql` | SQL text — **edit this in Git** for logic changes |
| `MyQuery.json` | Query type, properties, column metadata, Design View layout |

Import combines them back into an Access query object.

---

## What to edit in Git

- **Change SQL** → edit the `.sql` file.
- **Rare layout-only changes** → edit `.json` only if you understand Design View metadata (table positions, join lines). Most teams edit SQL and re-export once to refresh layout.

Pass-through queries store SQL verbatim in `.sql`; formatter and designer rules do not apply.

---

## Design View vs SQL View import

If the SQL is compatible with the query designer and layout data exists in `.json`, the query imports in **Design View** (diagram positions preserved). Otherwise the add-in uses **SQL View**. Some complex SQL (certain unions, subqueries, or non-designer expressions) always import as SQL View — the database still runs correctly.

---

## Legacy formats

These still **import** as a compatibility bridge; they are not the actively developed query format:

- `*.bas` — primary query source on export format 4.1.2
- `*.qdef` — SaveAsText-style export when format 5.0+ has **Use Deterministic Query Export** turned off

A paired `.sql` written by **Save Query SQL** is applied on import only when **Force import of original SQL for queries** is enabled (Options → Build). That option does not change which file the build lists under "Importing queries".

On the next export with format 5.0 and deterministic query export (the default), leftover `.bas`/`.qdef` files are replaced by `.sql` + `.json`. To make `.sql` authoritative, switch the project to export format 5.0+ and run a full export.

---

## Troubleshooting

| Symptom | Things to try |
|---------|----------------|
| Large unexpected diff after export | One-time migration to `.sql`/`.json`; ensure queries were not re-saved in Designer vs SQL View unnecessarily |
| Query fails after build | Enable **Force import of original SQL for queries** (Options → Build) for edge cases; see [issue #76](https://github.com/joyfullservice/msaccess-vcs-addin/issues/76) |
| Drift every export | Compare with [Regression Testing](Regression-Testing); report a fixture if it is a parser bug |

---

## For contributors (parser internals)

Maintainer reference: [`docs/access-query-storage.md`](https://github.com/joyfullservice/msaccess-vcs-addin/blob/dev/docs/access-query-storage.md) in the repository (not synced to this wiki).

Regression fixtures: `Testing/Fixtures/queries/` in the add-in repo.
