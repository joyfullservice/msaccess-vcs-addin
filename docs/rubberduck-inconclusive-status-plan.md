# Rubberduck "Inconclusive" outcome support in the VCS test runner

**Owner:** pflugs30 fork · branch `feat/rubberduck-test-support`
**Parent effort:** issue #308 — making the existing Rubberduck (RD) test suite runnable
under the VCS add-in test runner (see the reconciliation plan in the `original-system2`
repo, §4). This document is the detailed, living plan for the Inconclusive sub-effort and
is updated as each phase lands.

---

## Progress

| Phase | Scope                                                   | Status                                                            |
| ----- | ------------------------------------------------------- | ----------------------------------------------------------------- |
| I-1   | Model + tri-state reporting channel (add-in)            | **Verified** (headless run green)                                 |
| I-2   | Shim emulation of RD `AssertClass` (`StubRdAssert.cls`) | **Verified**                                                      |
| I-3   | Serialization surfaces (JSON, state, JUnit)             | **Verified** (JSON confirmed; JUnit pending inspection)           |
| I-3b  | UI surfaces (HTML report, web runner badge/filter)      | **Implemented** (pending add-in rebuild to verify)                |
| I-4   | Parity verification on the practice DB                  | **VCS side verified**; RD Test Explorer parity pending user check |
| 0-B*  | Multi-RD-module dispatcher (shared lifecycle names)     | **Verified** (`5.1.0-pflugs30-c`, combined run, no err 91)        |

Update this table (and the per-phase checklists) on every meaningful change.

---

## Execution model (confirmed): emulation, not delegation

Under the VCS runner we **do not** invoke Rubberduck's engine. Two separate paths exist,
reconciled by the shim (`StubRdAssert` via the `CreateTestAssert()` factory):

- **RD Test Explorer path:** RD runs each test via `IVBETypeLibsAPI.ExecuteCode` and captures
  assertion outcomes through `AssertHandler.OnAssertCompleted` events that only RD subscribes
  to. When RD is present, the shim **delegates** to the real `Rubberduck.AssertClass`.
- **VCS runner path:** nobody subscribes to those RD events, so the shim **computes the
  outcome itself** (emulating `AssertClass` semantics) and reports it to the add-in. Pure
  emulation — no Rubberduck engine involved.

The only reason RD stays runnable is the shim's dual mode; the VCS runner never calls RD.

## Why this work is needed

The classification channel is currently **binary**. `clsTestRunner.RecordAssertion` takes
`blnCondition As Boolean`; the outcome classifier knows only Errored / Empty / Failed /
Passed; `eTestStatus` has no Inconclusive member. RD, by contrast, treats several cases as
**Inconclusive, not Failed** (from the RD `AssertClass` source):

- `AreEqual` / `AreNotEqual` where operand **types differ** (`AreEqual 42, "42"`, `AreEqual 42, Null`)
- `AreSame` / `AreNotSame` on **value types** (not object references)
- an explicit `Assert.Inconclusive([message])` call

Without support, each shows up as a false **FAIL** under the VCS runner.

## Accepted decisions

1. **Precedence** when one test accumulates both Failed and Inconclusive: **Failed wins**
   (`Errored > Failed > Inconclusive > Passed`). Inconclusive must never mask a real failure.
2. **Record-and-continue** (VCS model) is kept over RD's stop-at-first. A rare edge case
   diverges (inconclusive-then-would-fail reads Inconclusive in RD, Failed in VCS); accepted
   because the production suite is engineered to stay conclusive. Stop-at-first on the shim
   path is a deferred optional fidelity upgrade.
3. **CI `allPassed`**: Inconclusive is **non-failing** (does not set failure). A run of only
   inconclusive tests is not `allPassed` (nothing passed). `TreatInconclusiveAsFailure`
   option deferred.
4. **JUnit mapping**: `INCONCLUSIVE` → `<skipped/>` (NUnit convention; JUnit has no native
   inconclusive).

No `eExportFormatVersion` gating is required: test results live in gitignored
`test-results/` artifacts, not exported source, so they are outside that contract.

---

## Phase I-1 — Model + reporting channel (add-in only)

Files: `clsTestRunner.cls`, `clsVersionControl.cls`, `modAPI.bas`.

- [x] Add `etsInconclusive = 5` to `eTestStatus`.
- [x] Add `eAssertOutcome` enum (`eaoPassed=1, eaoFailed=2, eaoInconclusive=3`).
- [x] `clsVersionControl.HandleTestOutcome(lngOutcome As Long, varContext) As Boolean` —
      returns `False` when `TestRunner.State <> etrsRunning`, else
      `TestRunner.RecordAssertionOutcome`. Mirrors `HandleTestAssertion`.
- [x] `modAPI.HandleTestOutcome` public wrapper (so the shim reaches it via `Application.Run`).
- [x] `clsTestRunner.RecordAssertionOutcome(lngOutcome, ctx)`; existing Boolean
      `RecordAssertion` delegates (True→Passed, False→Failed). Store `"outcome"` per
      assertion; keep `"passed"` (= outcome is Passed) for backward compat.
- [x] New counters in `udtTestRunner`: `InconclusiveAssertions`, `InconclusiveCount`;
      reset alongside the others.
- [x] Classifier: insert Inconclusive between Failed and Passed. Add
      `AnyAssertionInconclusive`; make `AnyAssertionFailed` match `eaoFailed` specifically
      (falling back to `"passed"` when `"outcome"` absent).
- [x] `HandleTestAssertion(Boolean, …)` stays **unchanged** (add-in's own `TestAssert` tests).

## Phase I-2 — Shim emulation (`StubRdAssert.cls`, practice DB)

- [x] Implement RD `AssertClass` surface computing `eAssertOutcome`:
  - `IsTrue/IsFalse/IsNothing/IsNotNothing/Fail/Succeed` → Passed/Failed.
  - `AreEqual/AreNotEqual` → null-coercion (`Null`→`""`), `Null==Null` passes, `Null` vs `""`
    passes, **types differ → Inconclusive**.
  - `AreSame/AreNotSame` → `ReferenceEquals`; both-null passes for `AreSame`;
    **value types → Inconclusive**.
  - new `Inconclusive([message])` method → always Inconclusive.
- [x] Per call: compute outcome → `If HandleTestOutcome(outcome, ctx)` (VCS run active) done;
      `ElseIf` real `AssertClass` present → delegate (RD path); `Else` `Debug.Assert`.

Note: the shim already computed all three outcome codes internally; this phase replaced the
Boolean `HandleTestAssertion` bridge (which collapsed Inconclusive to a failure) with the new
tri-state `HandleTestOutcome` bridge via a `VcsOutcome` mapping helper. The `Debug.Assert`
last-resort fallback now only breaks on a genuine failure, not on Inconclusive.

## Phase I-3 — Serialization surfaces

- [x] `clsTestRunner.GetResultsAsJson`: add `"inconclusive"` to summary + `"outcome"` per
      assertion; `StatusToString`/`StringToStatus` gain `INCONCLUSIVE`.
- [x] `modTestState.BuildSummaryFromState`: `INCONCLUSIVE` case + count; assertion `"outcome"`
      round-trip (`OutcomeCodeToString`/`OutcomeCodeFromString`); `CloneAssertionsFromState`
      carries `"outcome"`.
- [x] `modTestJUnit`: `INCONCLUSIVE` → `<skipped/>`.

## Phase I-3b — UI surfaces

- [x] Web runner streaming (`modTestRunnerUI`): `WebStatusFromRunnerStatus` maps
      `etsInconclusive` -> `"inconc"` (was falling through to `"pending"`, which made the
      inconclusive tests leak into "New tests" with no icon); per-assertion `"outcome"`
      token (`WebOutcomeToken`) passed through so inconclusive assertions render correctly.
- [x] Web runner (`TestRunner/runner.html`): `inconc` status class + ⚠ icon, amber
      CSS vars (light/dark), status-icon/detail-status/badge/stale/assertion-block styles,
      `data-filter="inconclusive"` list-filter rules, stats-bar chip (`stat-inconclusive`),
      "Inconclusive" filter tab, progress-bar segment (`prog-inconc`), and `inconclusive`
      folded into every `counts` object / increment branch / suite summary.
- [x] HTML report (`TestRunner/results.html`): `inconc` webStatus + ⚠ icon, amber CSS vars,
      donut segment, per-suite bar segment, summary stat pill, "Inconclusive" filter tab,
      `matchesFilter` case, and assertion-level outcome rendering.

## Phase I-4 — Parity verification

- [x] Add a test module / practice-DB fixture exercising `AreEqual 42, "42"`,
      `AreEqual 42, Null`, `Assert.Inconclusive("reason")`, and mixed tests
      (`InconclusiveTestModule.bas`, 5 tests incl. a Fail-beats-Inconclusive precedence case).
- [x] VCS headless run classifies them Inconclusive with the accepted `allPassed` behavior.
      Result: `passed 1, failed 1, inconclusive 4`, `allPassed false`; the four inconclusive
      cases → INCONCLUSIVE, the Inconclusive-then-Fail case → FAILED (precedence holds).
- [ ] RD Test Explorer marks the same tests Inconclusive (parity) — pending user check in Access.

> **Blocker found (separate from Inconclusive):** the runner invokes every RD proc via
> `Application.Run` by name, and Access cannot resolve a proc name that is duplicated across
> standard modules (bare, project-qualified, and module-qualified all return "cannot find the
> procedure"). Every RD module shares lifecycle names (`ModuleInitialize`/`TestInitialize`/…),
> so a multi-RD-module run fails: the ambiguous lifecycle call is swallowed, the module's
> `Assert` is never set, and the first `Assert.*` raises runtime error 91. The real 44-module
> suite is blocked by this.
>
> **Fix implemented (pending add-in rebuild): generated dispatcher.** `clsTestRunner.BuildDispatcher`
> injects a temporary standard module (`modVCSTestDispatch`) into the target DB containing
> uniquely-named wrappers (`zz_VCS_Dispatch_N`) that call each RD lifecycle/test proc
> module-qualified (`[Module].[Proc]` — legal and unambiguous in compiled VBA). The run loop,
> `RunTestPhase`, and `InvokeModuleLifecycle` route RD invocations through `DispatchName` →
> `BuildRunCmd(wrapper)`; the module is removed in `CleanUp`. Best-effort (falls back to direct
> `Application.Run` on injection failure); only standard modules are wrapped; the dispatcher also
> reaches `Option Private Module` procs. The fixture's lifecycle procs were restored to the
> standard RD names so `InconclusiveTestModule` + `ExampleTestModule` (both using
> `ModuleInitialize`, etc.) form the collision regression test. **Verified** on add-in
> `5.1.0-pflugs30-c`: a combined headless run of both modules classified `ExampleTestModule`
> PASSED and the Inconclusive/Failed cases correctly with **no error 91**, and the temp
> `modVCSTestDispatch` module was removed after the run (object list clean).

---

## Change log

- 2026-08-02 — Plan created; decisions 1–4 accepted. Phases not yet started.
- 2026-08-02 — Phases I-1, I-2, I-3 implemented (code-complete, pending add-in rebuild).
  - I-1 (`clsTestRunner.cls`, `clsVersionControl.cls`, `modAPI.bas`): `etsInconclusive`,
    `eAssertOutcome`, `RecordAssertionOutcome`, `HandleTestOutcome` (class + modAPI wrapper),
    inconclusive counters, `AnyAssertionInconclusive`, outcome-aware `AnyAssertionFailed`,
    `OutcomeToString`/`StringToOutcome`, `INCONCLUSIVE` in `StatusToString`/`StringToStatus`,
    summary + per-assertion `outcome` in `GetResultsAsJson`, `AccumulateCountsFromTest`,
    `CloneAssertionsFromState`, `ResetCounts`, `LogRunSummary`.
  - I-2 (`StubRdAssert.cls`, practice DB): retargeted the VCS bridge from Boolean
    `HandleTestAssertion` to tri-state `HandleTestOutcome` via `VcsOutcome` mapping; softened
    the `Debug.Assert` fallback to break only on failure.
  - I-3 (`modTestJUnit.bas`, `modTestState.bas`): `INCONCLUSIVE` → `<skipped/>`; durable-state
    assertion `outcome` round-trip + inconclusive counts in `BuildSummaryFromState`.
  - Next: rebuild + reinstall the add-in (Access closed), then Phase I-4 parity verification.
- 2026-08-02 — Add-in rebuilt/reinstalled (`5.1.0-pflugs30-b`). Phases I-1/I-2/I-3 verified via
  headless run of the existing passing module (backward-compatible schema: `inconclusive: 0`,
  per-assertion `outcome: passed`, `allPassed: true`). Phase I-4 fixture (`InconclusiveTestModule`)
  added and VCS side verified (see Phase I-4). Discovered a **multi-RD-module `Application.Run`
  name-collision blocker** (lifecycle procs share names → err 91); documented under Phase I-4 as a
  separate Phase 0-B follow-up requiring a generated dispatcher. Interim: fixture lifecycle procs
  renamed to unique names so the Inconclusive work could be verified independently.
- 2026-08-02 — **Multi-RD-module dispatcher implemented** (`clsTestRunner.cls`, pending rebuild).
  Root cause: Access `Application.Run` cannot resolve a proc name duplicated across standard
  modules, so shared RD lifecycle names collide. Added `BuildDispatcher`/`DispatchName`/
  `RemoveDispatcher`/`AddDispatchPair`/`RemoveDispatchModuleIfPresent`/`IsStdModule` and
  `DispatchMap`/`DispatchModuleName` state; `RunSelected` builds the dispatcher before the
  pre-run compile and removes it in `CleanUp`; `RunTestPhase`/`InvokeModuleLifecycle`/the
  standard-module invocation route RD procs through it. Reverted the fixture's lifecycle procs
  to standard RD names so the two-module collision is now the regression scenario. Next:
  rebuild + reinstall, then verify a combined `ExampleTestModule` + `InconclusiveTestModule`
  headless run (no error 91; Inconclusive classification intact).
- 2026-08-02 — **Dispatcher verified** on add-in `5.1.0-pflugs30-c`. Combined headless run of
  `ExampleTestModule` + `InconclusiveTestModule` (both with standard RD lifecycle names):
  `passed 3, failed 1, inconclusive 4`, `allPassed false`, **no error 91**. `ExampleTestModule`
  `ModuleInitialize` ran correctly, the four Inconclusive cases classified INCONCLUSIVE, and the
  Inconclusive-then-Fail case FAILED (precedence). `vcs_list_objects` confirms the temp
  `modVCSTestDispatch` was cleaned up. Multi-RD-module support (issue #308) is now unblocked.
- 2026-08-02 — **Phase I-3b (UI surfaces) implemented** (pending add-in rebuild). Root fix:
  `modTestRunnerUI.WebStatusFromRunnerStatus` mapped `etsInconclusive` to `"pending"`, so the
  web runner rendered inconclusive tests with no icon and leaked them into "New tests". Added
  the `"inconc"` token + `WebOutcomeToken` assertion passthrough. `TestRunner/runner.html` and
  `TestRunner/results.html` (HTML report) both gained the `inconc` status class, amber ⚠ icon,
  light/dark CSS vars, badges/detail-status/assertion styling, an "Inconclusive" filter tab,
  a stats chip / summary pill, a progress/donut segment, and `inconclusive` folded into every
  counts object, increment branch, and suite summary. Both HTML files are build-embedded assets,
  so a rebuild + reinstall is required before the UI reflects the change. Next: rebuild, then
  open the web runner and confirm the four inconclusive tests show the ⚠ icon, an orange count
  chip, and a working "Inconclusive" filter.
