# Machine-readable performance diagnostics

The `PERFORMANCE REPORTS` section at the end of every operation log is written for a
person skimming a build. It has to fit in a fixed-width column layout, so it shows one
view: operations aggregated by name, timed exclusive of anything nested inside them.

That view answers one question well and another one not at all. Setting
`ExportPerfJson` writes the same measurements to a separate JSON file with several
views, including the call paths that the text table flattens away.

## Turning it on

The option is off by default, because collecting call paths costs a string build on
every `OperationStart` and `OperationEnd`.

Persistently, for a project — add to its `vcs-options.json`:

```json
"ExportPerfJson": true
```

Over MCP, for one session:

```python
vcs_set_option(db, "ExportPerfJson", True)
```

**The session override only holds inside the Access process that received it.** The
session ID lives in a module variable in that project (`modObjects.SessionId`, set by
`RegisterSession`), so a tool call that starts a *new* Access process falls back to
`options-default.json` and your override is not applied. If the file does not appear,
that is the likely reason; use `vcs-options.json` instead. This applies to every
option, not just this one.

## Where it lands

Beside the operation log it belongs to, with the same base name:

```
logs/Build_20260827_173418_105.log
logs/Build_20260827_173418_105.perf.json
```

Every operation that writes a log writes the JSON — export, build, merge, test run.
Tools return the log path, so the JSON path is `<log path minus .log>.perf.json`. Both
are under `logs/`, which the add-in gitignores, so **Glob and Grep silently skip them**;
read the path directly.

## What is in it

| Key | View |
|---|---|
| `operations` | Exclusive seconds and call count per operation name, aggregated over every call site |
| `callPaths` | One row per distinct call path, with `exclusiveSeconds` and `inclusiveSeconds` |
| `callers` | Per operation, which parent drove it and how many of the calls came from each |
| `categories` | Per component type, as in the text report |
| `totals` | `runtimeSeconds`, `accountedSeconds`, and the unattributed `otherSeconds` |
| `stream` | Effective progress interval plus candidate, sent, and suppressed counts |
| `mcp` | Callback attempts by type, failures, pruning, pending requests, and peak in-flight requests |
| `unfinishedOperations` | Operations still open when the report was taken; if non-empty, distrust the rest |
| `trackedCallPaths` | False means path tracking was off, so `callPaths` and `callers` are empty |

`path` is an array of segments rather than a joined string, so an operation name
containing a separator cannot be mis-split.

## Reading it without drawing the wrong conclusion

**Exclusive seconds sum; inclusive seconds overlap.** Every `operations` row and every
`callPaths` row is exclusive of nested work, which is why they add up to
`accountedSeconds`. `inclusiveSeconds` deliberately double-counts — a parent's
inclusive figure contains its children's. Summing a column of inclusive figures
produces a number larger than the runtime and means nothing.

**A feature's cost is usually split across two names.** `PostCallback` calls
`ConvertToJson`, which starts its own `Convert to JSON` operation, so the row named
after the feature covers transport only and the payload cost sits under the
serializer. In the text report there is no way to put those back together. Here,
`callers` does it: `Convert to JSON` showing `73` of its `105` calls arriving from
`MCP Callback` is what tells you the streaming callbacks cost the sum of the two rows.

**Callback volume has its own controls.** `McpProgressIntervalMs` defaults to 500,
so a long tight component loop emits progress at most twice per second. `Log.Add`
messages are separate, unthrottled callbacks. Compare `stream.progressSent` with
`mcp.logPostAttempts` before assuming the progress throttle controls most of the HTTP
traffic.

**Use `callers` for shared leaf operations and `callPaths` for features.** `Compute
SHA256` is called from four different hashing paths; the interesting question is which
of them is driving the count, which is a `callers` question. "What does the whole
merge step cost" is a `callPaths` question.

**`otherSeconds` is not overhead you can hunt directly.** It is whatever ran outside
any `OperationStart`. Turning path tracking on inflates it slightly, since building a
path string happens between a child's stop and its parent's restart, when no timer is
running. Compare `otherSeconds` only between runs with the same `trackedCallPaths`.

## Comparing two runs

The counts are the control. When two runs of the same work show identical call counts
and different seconds, nothing changed in what the code did and the difference is
environmental — machine state, core scheduling, or contention. When the counts differ,
the code took a different path and that is where to look first.

MCP-launched Access now requests full-power QoS (EcoQoS off, Above Normal) so a
hybrid CPU is less likely to park the single Access thread on an LP-E core. A
remaining gap versus a foreground ribbon build can still be cold-process start
or MCP JSON work, not core class.
