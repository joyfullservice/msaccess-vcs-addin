# Debugging RunVBA failures

`clsVersionControl.RunVBA` (exposed to agents via the `vcs_run_vba` MCP tool)
wraps caller-supplied VBA code in a temporary module, compiles it, runs it, and
returns a JSON result. Read this when an agent-authored snippet fails and you
need to know what the returned JSON is telling you.

For the general `DebugMode` / `CatchAny` / `LogUnhandledErrors` system used
throughout the add-in, see [error-handling.md](error-handling.md).

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
