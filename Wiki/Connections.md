# Connection Strings and `.env` Files

The add-in stores linked-table and pass-through query connection strings in a
per-developer `.env` file instead of committing raw ODBC strings to source control.
Exported source files reference these entries as `env:conn_*` keys.

## Where `env:` references appear

| Location | Property |
|----------|----------|
| `tbldefs/*.json` | `Connect` on SQL-linked tables |
| `queries/*.json` | `Connect` on pass-through queries |
| `db-connection.json` | Keys used to prime Access's ODBC cache during import |

During **export**, the add-in replaces connection strings with `env:conn_*`
references when `UseEnvForConnections` is enabled in `vcs-options.json`. During
**import/build**, each reference is resolved back to the full ODBC string from
the merged `.env` config.

## `db-connection.json` (priming only)

This file does **not** store connection strings. It lists which `env:conn_*`
connections to touch during the "Importing db connections" step so Access caches
ODBC credentials once before importing dozens of linked tables.

Example:

```json
{
  "Items": {
    "env:conn_mydb": {
      "env:conn_mydb": "dbo_tblCustomers"
    }
  }
}
```

The inner value is a linked table name used to open a temporary query against that
connection. If a referenced key is missing from `.env`, Access may show an ODBC
login dialog during import.

## Base `.env` file

Path: `{ExportFolder}/.env` (gitignored).

The add-in auto-creates this file on export when connection sanitization is active.
Each developer maintains their own copy with machine-specific server names and
credentials.

When the add-in manages your `.gitignore`, it excludes `*.env` plus all layered
variants (`.env.*`) while keeping `*.example` templates committed:

```
*.env
.env.*
!.env*.example
```

Commit templates such as `.env.example` and `.env.offline.example` so other
developers know which keys to fill in; never commit the real `.env*` files.

Named connections (`EnvConnectionNames` in `vcs-options.json`) let projects use
stable keys like `conn_production` instead of auto-generated names.

## Layered environments (`APP_ENV`)

For projects that target different backends (dev/staging/production) from the same
source tree, use a selector file plus environment-specific overlays.

**Load order** (later overrides earlier):

1. `.env` — shared base config; set `APP_ENV` here
2. `.env.local` — developer-local overrides (gitignored)
3. `.env.{APP_ENV}` — environment-specific values (e.g. `.env.offline`)
4. `.env.{APP_ENV}.local` — environment-specific local overrides (gitignored)

`APP_ENV` is read from the OS environment first (`Environ("APP_ENV")`), then from
the merged base level (`.env` + `.env.local`). If `APP_ENV` is empty, only levels
1–2 load — fully compatible with single-file `.env` projects.

### Example layout

```
MyProject.accdb.src/
  .env                 # APP_ENV=offline + shared keys
  .env.local           # gitignored — your machine overrides
  .env.offline         # Purple_Offline connection strings
  .env.offline.local   # gitignored — offline machine overrides
  .env.live            # production connection strings
```

Base `.env`:

```
APP_ENV=offline
# shared non-secret config...
```

`.env.offline`:

```
conn_purple=ODBC;DRIVER={ODBC Driver 18 for SQL Server};SERVER=localhost;DATABASE=Purple_Offline;Trusted_Connection=Yes;
```

### Switching environments

1. Edit `APP_ENV` in `.env` (or set the OS environment variable).
2. Ensure the matching `.env.{APP_ENV}` file exists with the correct `conn_*` keys.
3. Rebuild or merge from source so linked tables relink with the new connection strings.

**Note:** Layered `.env` resolution applies at **build/import** time. Runtime VBA
code that opens ADODB connections must read `.env` itself (or use a project-specific
loader). Linked tables bake in the resolved ODBC string when imported.

### Reads vs writes

- **Import/build resolution** uses the merged config (all four levels).
- **Export writes** go to the base `.env` only — auto-generated `conn_*` entries
  are not written into `.env.{APP_ENV}` files.

## Troubleshooting auth dialogs during build

An ODBC login prompt usually means:

1. The `env:conn_*` key referenced in source is missing from the resolved `.env`
   config for the active `APP_ENV`.
2. The SQL Server (or other ODBC target) is unreachable — ODBC falls back to a
   login dialog on timeout.

Check the build log for:

```
Connection key not found in .env file: conn_...
```

Compare `db-connection.json` priming keys and `tbldefs`/`queries` `Connect` values
against the keys present in your merged `.env` files.

## Related options

| Option | Location | Purpose |
|--------|----------|---------|
| `UseEnvForConnections` | `vcs-options.json` | `Auto` / `Always` / `Never` for export sanitization |
| `EnvConnectionNames` | `vcs-options.json` | Stable named keys (Tier 2 matching) |

See also [Options](Options) and [Export-Import File Types](Export-Import-File-Types).
