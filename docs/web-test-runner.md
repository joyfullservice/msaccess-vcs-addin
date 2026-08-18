# Web test runner

On Microsoft 365 builds with the Edge browser control (file build >= 16327) and
with the `UseWebTestRunner` option enabled (default **on**), `VCS.RunTests` and
the ribbon **Run Tests** button open `frmVCSTestRunner`, merge-scan for tests, and
publish the test tree to `TestRunner/runner.html` via `modTestRunnerUI`
(`ExecuteJavascript` to `window.TestUI.*`). Tests are **not** auto-run; the user
clicks the primary **Run** button (its label reflects scope: all, folder, suite,
filter, or failed) or a per-suite / per-test play button. When the option is off,
or the Access build is older, it falls back to the `frmVCSMain` console unchanged.

`Options.UseWebTestRunner` lives under Advanced options > Automated Testing and
gates the whole routing in `clsVersionControl.ExecuteTests`. The Edge control is
referenced as `As Object` (late binding) only so the add-in still compiles on
older Access.

See DECISIONS.md (2026-07-07 through 2026-07-09, Web test runner) for the
rationale behind these choices.

---

## Packaging and entry points

- **HTML**: [`TestRunner/runner.html`](../TestRunner/runner.html) — a repo-root
  packaging asset, embedded at build like `Ribbon.xml`, extracted to a stable
  `TestRunnerCache\` folder under the add-in install path so WebView2
  `localStorage` (Recent filters, column widths, theme) persists across sessions.
- **`VCS.RunTests`** — show the form and defer the scan; the user clicks Run.
- **`VCS.OpenTestRunner`** — open to view last results, rehydrating from the
  in-memory singleton or from `test-state.json` when the singleton is empty.

Opening the web runner does not begin an `Operation`; the bridge Run callbacks do.

## Inbound bridge

The bridge is **outbox polling**: JS enqueues commands in `window.__vbaOutbox`
and the form timer drains them via `RetrieveJavascriptValue`. No navigation is
involved — see the `frmVCSTestRunner` header comment and the DECISIONS.md entry
of 2026-07-08.

Allowlisted callbacks: `RunAll`, `RunSelected`, `RunFailed`, `Cancel`,
`OpenTestSource`, `RefreshTests`, `OpenResultsReport`, `CopyResultsPath`.

### Run command protocol

Run commands (`RunAll` / `RunSelected` / `RunFailed`) resolve the JS promise with
an **acceptance ack** (`AcceptBridgeRun`) *before* the blocking run executes
(`ExecutePendingBridgeRun`); completion arrives later via streamed
`onRunComplete` / `onRunCancelled` / `onRunError`. This is why long runs never
trip the JS `VBA_CALL_TIMEOUT_MS` (30 s, which applies to request/response calls
only). A promise rejection means the run was refused before starting — already
running, no matching keys, or `Operation.Begin` failed.

## Form lifecycle

The form opens as a **pop-up** window (`PopUp=1`). Closing via the X button or
Escape **hides** it: the timer is disabled and WebView2 stays warm. Re-opening
**reuses** the hidden instance without reloading the page when it is healthy (see
`open.reuse.warm` in the diagnostic log); a forced reload replays completed
results. Programmatic `CloseWebTestRunner` sets `AllowClose` and issues
`DoCmd.Close` for a real unload.

## Tree refresh and hydration

After show, or on **Refresh** in the toolbar, VBA runs `ScanMergingPriorResults`
— rediscovering tests while preserving pass/fail for unchanged `Module.Proc`
keys — publishes the tree, then overlays durable state from `test-state.json` in
one `onResultsBatch`. Names paint first; prior durations and assertion counts
follow without blocking the scan.

On a cold open the parse is normally already paid for. `WaitForWebRunnerReady`
spins on `DoEvents` for the whole WebView2 first-init, so `PrefetchDurableState`
spends that otherwise idle window reading and parsing `test-state.json` into the
`modTestState` session cache (keyed on path, modified time, and size). WebView2
initializes in its own processes and keeps making progress while VBA holds the
thread, so most of that parse is genuinely hidden — but its COM callbacks queue
until we pump, so `DocumentComplete` slips by part of whatever runs here. The
window is now full; adding more work to it mostly lands back on the critical
path. When `modTestState.StateCached` is then True,
`RefreshWebTestTreeDeferred` merges **inline** and pushes only `onHydrateEnd`:
raising the indicator would cost a JS round trip and a minimum-visible hold to
announce work that is already finished.

Otherwise the overlay is **deferred to `Form_Timer`** (`ScheduleHydratePriorResults` sets a
flag; `PumpDeferredHydrate` does the work on the next tick) so Access reaches its
message loop and the page paints before VBA blocks on the parse — WebView2
composites no frames while VBA holds the thread. `PumpDeferredHydrate`
additionally waits for the page to confirm the paint (`window.__hydratePainted`,
set from a double `requestAnimationFrame`, capped by
`HYDRATE_PAINT_TIMEOUT_MS`), because the next timer tick alone arrives roughly
50 ms after the push — far too soon.

`onHydrateStart` / `onHydrateEnd` cover the gap by switching the header status
badge to a pulsing **Loading previous results...** plus a stats-bar chip; JS holds
both up for `HYDRATE_MIN_VISIBLE_MS` so a fast parse does not flash them past
unread. The overlay runs **once per open** — a re-fired `DocumentComplete`
replays the already-merged state from memory instead of parsing the file (and
showing the indicator) a second time.

## Quiet mode

While the web runner hosts a run, `Log.SuppressDebugOutput` is set so per-test
results are not echoed to the Immediate window. The UI shows them; the log file
is unaffected.

## UI affordances

The sidebar has **All tests** and **Failed tests (N)** focus entries, a nested
**@Folder** tree with folder select (click the name) and a play button, a **Tags**
section with an include/exclude cycle, and **Recent**, which stores full
`{folder, suite, filterText}` snapshots so combinations restore on click.

A single filter box uses `VCS.RunTests` token syntax (`SQL -slow`) and scopes both
the test list and Run; the sidebar tree is navigation-only, not a second filter.
The stats bar shows PHPUnit-style **tests** and **assertions** totals. The primary
Run button executes the **visible scope** (composed navigation plus filter), and
per-test, per-suite, and per-folder play buttons run narrower scopes.

**Copy path** copies the bare `test-results/test-state.json` path to the clipboard
for pasting into an agent chat. Clicking a location opens the VBE at that
procedure. `VCS.RunTests(...)` and the ribbon `DefaultTestFilter` prefill the
filter box when the runner opens, without auto-running.

`Options.ExportTestResultsHtml` writes a self-contained HTML dashboard reachable
from **Open report** in the toolbar; see
[testing-strategy.md](testing-strategy.md) for where all result artifacts land.

---

## Diagnostic trace log

`modTestRunnerDiag` writes a single agent-readable trace of the real
bridge/lifecycle flow to `<ExportFolder>\logs\TestRunnerDiag_<timestamp>.log`,
falling back to a temp folder when Options are not loaded. Tracing is **off by
default**: `VCS.TestRunnerDiag True` (or `modTestRunnerDiag.DiagEnabled = True`
in the Immediate Window) and reopen the runner to capture a session. Each line
is `[+elapsed ms +delta] TAG | detail`. Nested `phase.begin` / `phase.end`
spans carry `ms=`; a per-phase summary table is written at hide or when tracing
is turned off. Writes are buffered (not one file open per line) and flushed
every 32 lines or 250 ms, so a VBA state reset cannot swallow the tail. Phases
called too often to print — `log.flush`, `js.exec`, `js.retrieve` — use a quiet
span: they always accrue into the summary but only print a single line when they
exceed their threshold, so a slow one still shows up in place.

A span is not free: each one allocates a `Scripting.Dictionary`, which costs
~0.5 ms in an interactive Access session. Instrument phases and per-test steps,
never a loop over thousands of items — at that scale the trace measures itself.

| Tag group | Tags | Meaning |
|---|---|---|
| Form lifecycle | `form.load`, `form.unload`, `form.hide`, `form.show`, `hide.*` | Open/close/hide transitions; hide spans split cancel / clear-op / timer / Visible |
| Navigation | `navigate.url`, `navigate.call`, `documentcomplete` | What URL the control was given; the gap from `navigate.call` to `documentcomplete` is WebView2 load / cold-start time |
| Readiness | `wait.ready`, `wait.timeout` | Outcome of the readiness wait |
| Open / scan | `open`, `refresh.tree`, `scan.*`, `tree.json`, `publish.tree`, `state.prefetch`, `hydrate`, `hydrate.inline`, `merge.counts` | Cold/warm open, merge-scan, tree JSON, state parse overlapped with the Edge start, prior-result overlay |
| Teardown | `teardown`, `save.results`, `persist`, `state.*`, `junit.export`, `html.export`, `log.save` | Post-run artifact writes; `form.hide.entry` reports `sinceIdleMs` to tell queued input from a slow handler |
| Bridge | `beforenavigate`, `defer.exec`, `dispatch.begin`, `dispatch.end`, `resolve`, `reject` | Command received, deferred, dispatched, and settled |
| Streaming | `push`, `push.dropped`, `push.<handler>`, `js.exec`, `js.retrieve` | VBA-to-JS result streaming and each ExecuteJavascript / RetrieveJavascriptValue |
| JS breadcrumbs | `js.call`, `js.onReady`, `js.renderTestList`, `js.tick`, `js.onResultsBatch` | Drained from `window.__diag`; `ts` is mapped onto the session clock |

Read this file first when the page does not load, a call times out, or the
runner feels slow — it shows exactly where the flow diverged or the time went.
