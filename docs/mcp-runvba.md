# Running VBA over MCP

`clsVersionControl.RunVBA` (exposed to agents via the `vcs_run_vba` MCP tool)
wraps caller-supplied VBA code in a temporary module, compiles it, runs it, and
returns a JSON result. Read this before using it to inspect a database, and when
an agent-authored snippet fails and you need to know what the returned JSON is
telling you.

For the general `DebugMode` / `CatchAny` / `LogUnhandledErrors` system used
throughout the add-in, see [error-handling.md](error-handling.md).

---

## Reading schema facts you cannot query

`vcs_execute_sql` reaches data, not schema metadata. Index membership,
uniqueness, primary keys, AutoNumber attributes, and field properties are not
exposed to SQL in ACE: `MSysIndexes` was a Jet-era table and does not exist in an
`.accdb`, so querying it fails with error 3078. DAO through `vcs_run_vba` is the
only way to read them.

**Read these facts rather than deducing them.** A DAO error number tells you that
a constraint was violated, not which one. And row order is not evidence of an
index at all: without an explicit `ORDER BY` the order is undefined, and in
practice ACE commonly returns physical insertion order, so a result that comes
back looking sorted may reflect nothing but the sequence the rows were written
in. Both invite a confident wrong answer, and one call settles it:

```vba
Dim dbs As DAO.Database
Dim tdf As DAO.TableDef
Dim idx As DAO.Index
Dim strOut As String
Dim i As Long, j As Long
Set dbs = CurrentDb
Set tdf = dbs.TableDefs("MSysIMEXSpecs")
For i = 0 To tdf.Fields.Count - 1
    strOut = strOut & "F " & tdf.Fields(i).Name & " auto=" & _
        CBool((tdf.Fields(i).Attributes And dbAutoIncrField) <> 0) & vbCrLf
Next i
For i = 0 To tdf.Indexes.Count - 1
    Set idx = tdf.Indexes(i)
    strOut = strOut & "I " & idx.Name & " unique=" & idx.Unique & _
        " primary=" & idx.Primary & " flds="
    For j = 0 To idx.Fields.Count - 1
        strOut = strOut & idx.Fields(j).Name & ","
    Next j
    strOut = strOut & vbCrLf
Next i
MCP_TempFunction = strOut
```

Two traps this surfaces that nothing else will. A column missing from
`TableDef.Indexes` has **no** uniqueness enforcement even when it is an
AutoNumber, so an explicit insert can duplicate it silently. And ODBC-sourced
column metadata (the `db-inspector` MCP server) reports AutoNumber as plain
`Long` and nullability unreliably for system tables, so it cannot answer either
question.

## Hold a database reference

`Set tdf = CurrentDb.TableDefs("x")` fails on the *next* statement with error
3420, "Object invalid or no longer set". `CurrentDb` returns a fresh `Database`
object that is released when the statement ends, taking the `TableDef` with it.
Assign it to a variable first, as above. The add-in's own code caches one
reference through `SharedDb` in `modObjects` for the same reason.

---

## MCP/API calls never break into the debugger

All MCP and external API entry points (`modAPI.API`, `modAPI.APIAsync`, and the
async timer handler in `modTimer`) open an in-memory **error-break suppression**
scope via `SuppressErrorBreaks` in `modErrorHandling`. While that scope is active:

- `LogUnhandledErrors` **never** executes its `Stop` statement, even when
  **Break on Error** is enabled in Options.
- `DebugMode` returns `False`, so add-in code uses `On Error Resume Next` rather
  than `On Error GoTo 0`.

Leftover errors are logged (``Unhandled error, likely before `On Error` directive``)
or returned as JSON (`RunVBA`) instead of halting Access until a human dismisses a
debugger break. Scopes nest (e.g. `APIAsync` to `API` to `RunVBA`) via a counter,
not a Boolean.

This scope does **not** mutate `Options.BreakOnError` — export persists options to
disk, and toggling the option would write through to `vcs-options.json`. VBE error
trapping during operations remains **Break in Class Modules** so any break that
does survive lands where the error was raised.

After an MCP call completes, interactive debugging from the ribbon behaves
normally.

## Reset and temporary-module recovery

Before calling `RunVBA`, the MCP worker queues a reset of the current database's
VBA project as a separate API call. It then performs a harmless built-in COM
property read as a message-pump barrier, reacquires the Access and add-in COM
references, and only then creates or executes the temporary module. The separation
is required because the VBE Reset command returns before its teardown runs; doing
payload work on the same VBA stack can end that work inside `Application.Run` and
surface as error 2517 ("cannot find the procedure `MCP_TempFunction`").

The reset is fail-closed. A response with `error_pattern` `reset_refused` or
`reset_failed` means no submitted VBA ran. Do not retry through another execution
path that skips the reset; resolve the stated Access/VBE condition first.

`RunVBA` also removes stale standard modules whose names begin `MCP_Temp_` before
creating a new wrapper. A previous failed call can leave one behind, and two
modules declaring `MCP_TempFunction` make the unqualified name unresolvable:

- `sweptModules` lists stale modules removed before this call.
- `temp_module_sweep_failed` and `orphanModules` mean cleanup was incomplete, so
  the payload was not started.
- `temp_module_unresolvable` means the post-compile canary could not resolve the
  generated module, so the payload was not started.
- `temp_module_cleanup_failed`, `cleanupFailed`, and `orphanModule` mean the
  payload may have completed but its wrapper survived. If a return value was
  already available it is carried as `payloadResult`.

Stop and tell the user when cleanup fails rather than repeatedly issuing calls.
The next properly reset call will attempt the stale-module sweep, but a module
that resists removal needs inspection in the VBE.

---

## Auto-injected line numbers and `errorLine`

Before the wrapper is built, `RunVBA` runs the submitted code through the private
helper `AddVbaLineNumbers` (in the same class). That helper prepends sequential
1-based VBA line numbers to every executable statement. The number equals the
1-based ordinal of the line within the original `code` string — the counter
advances on every physical input line, blanks, comments, and continuations
included — so when a runtime error fires, the JSON result carries an `errorLine`
that maps directly back to the caller's source.

```json
{
  "success": false,
  "error": "Type mismatch",
  "errorNumber": 13,
  "errorLine": 7
}
```

`errorLine: 7` literally means "line 7 of what I submitted" — no offset math
required. The field is omitted when no `Erl` value is available (the wrapper
itself failed to compile, or the error fired before any numbered line ran).

Lines that cannot legally carry a VBA line number — blank lines, pure comments,
continuations of a prior `_`-terminated line, and lines the caller already
pre-numbered — pass through unchanged. The counter still advances over them so the
numbers stay in sync with the original text.

## Concise multi-error test procedures

The default wrapper uses `On Error Resume Next` and reports the **last** runtime
error. When you want a test to keep going past the first failure and report every
problem in a single round-trip, write your own handler that exploits the
auto-injected line numbers:

```vba
Dim col As New Collection
On Error GoTo H
CurrentDb.Execute "DELETE * FROM tblA"
CurrentDb.Execute "INSERT INTO tblB SELECT * FROM nope"
CurrentDb.Execute "UPDATE tblC SET x = 1"
MCP_TempFunction = "errors=" & col.Count & " | " & Join(CollectionToArray(col), "; ")
Exit Function
H: col.Add Erl & ": " & Err.Number & " " & Err.Description
Resume Next
```

Each `Erl` value collected inside the `H:` label is meaningful because the wrapper
auto-numbered every line for you. You do not need to write `10`, `20`, `30`
yourself.

The choice between default single-error capture and an explicit multi-error
handler is per-test: pick whichever shape matches what the test is verifying.

---

## VBA error-handler state: `Err.Clear` is not enough

When execution is inside an active `On Error GoTo Handler` block, `Err.Clear`
clears the error object but does **not** reset the active exception/handler state.
Expected cleanup errors inside that handler can still break or poison the wrapper
if you only write `On Error Resume Next`.

Use this pattern before any expected-error cleanup inside a handler:

```vba
Handler:
    strMsg = Err.Description
    Err.Clear
    On Error GoTo -1      ' clear active handler state
    On Error Resume Next  ' now expected cleanup errors are safe
    CurrentDb.QueryDefs.Delete "__temp_query__"
    Err.Clear
    On Error GoTo 0
    GoTo ContinueAfterHandler
```

Do not use `Resume` after `On Error GoTo -1`; jump to a continuation label
instead. Prefer explicit existence checks over expected-error cleanup when the
code is simple enough.
