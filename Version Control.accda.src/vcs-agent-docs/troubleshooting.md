# Troubleshooting and Logs

## Reading the logs

Every operation writes a timestamped log to the `logs/` subfolder. Check the most
recent one first; it carries the detailed output, timing, and error messages.

| Pattern | Operation |
|---------|-----------|
| `Export_<timestamp>.log` | Export |
| `Build_<timestamp>.log` | Full build |
| `Merge_<timestamp>.log` | Merge build **and** single-object import |
| `TestRun_<timestamp>.log` | Test run console output |

A single-object import that seems to have produced no log wrote a `Merge_*.log`.

**`logs/` and `test-results/` are gitignored**, so tools that respect `.gitignore`
— glob, grep, and semantic search — will silently skip them. Use the shell:

```powershell
Get-ChildItem -Path "logs" | Sort-Object LastWriteTime -Descending | Select-Object -First 3
Get-ChildItem -Path "test-results"
```

## Import fails with no clear error

Almost always the encoding changed. Confirm the file still begins with the bytes
`EF BB BF` and re-save it as UTF-8 with BOM.

## Import fails or objects are corrupted after an edit

Check the line endings. The file needs CRLF (`0D 0A`), not bare LF. Restore it from
git or re-save with CRLF.

## Unresolved Git conflict markers

After a `git merge`, `git stash pop`, or branch switch, a source file may still
contain lines starting with `<<<<<<<` or `>>>>>>>`. The add-in skips that file and
logs:

`Unresolved Git conflict markers in '<file>' (line N). Resolve the conflict in this
file, then merge again.`

Open the named file, resolve every conflict section, remove the marker lines, save,
and run merge again.

## "Object not found" after import

The `Attribute VB_Name` inside the file no longer matches the filename, usually
because the file was renamed without updating the name stored in it. Make them
agree, without the extension.

## VBA compile error after import

Typical causes are a syntax error introduced by an edit, a missing declaration, or
a reference to something that does not exist.

The compile tools available over MCP report only success or failure — never the
failing module or line. When `vcs_compile_vba` returns `success: false`, or
`vcs_check_vba_compiled` returns `compiled: false`:

1. Stop. Do not edit more source files or guess at a fix.
2. Ask the user to open the database in Access, open the VBA editor, and choose
   **Debug > Compile**. Access jumps to the first error.
3. Ask them to paste the few lines around it. The snippet is usually enough; the
   error text is optional.
4. Propose a targeted fix, then re-import and re-compile.

Working in Access directly, the highlighted line after **Debug > Compile** tells
you the same thing.

## XML parse error

Invalid syntax such as an unclosed tag, or an unescaped special character. Validate
the structure and check how special characters are encoded.

## Build prompts for SQL authentication

Linked tables and pass-through queries store `env:conn_*` references rather than raw
ODBC strings, in `tbldefs/*.json`, the `Connect` value in `queries/*.json`, and
`db-connection.json`. At build time the add-in resolves each reference against the
`.env` file at this folder's root, which is per-developer and gitignored.

A prompt means the key is missing from the resolved configuration or the server is
unreachable. Search the build log for `Connection key not found in .env file:`.

`db-connection.json` only **primes** Access's ODBC cache during import, running one
temporary query per connection so later linked-table imports reuse the cached
credentials. It stores `env:` keys, not connection strings, so do not treat it as a
registry of connections.

Per-environment layering through an `APP_ENV` selector is supported and documented
in the
[Connections wiki page](https://github.com/joyfullservice/msaccess-vcs-addin/wiki/Connections).
