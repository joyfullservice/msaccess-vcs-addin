# Running Rubberduck test modules under the VCS test runner

Internal reference for the parallel test-discovery path that lets the VCS runner
(`clsTestRunner`) discover and execute test modules written to Rubberduck's
`@TestModule` / `@TestMethod` convention, and for the tri-state `Inconclusive`
outcome the runner models to match Rubberduck's `AssertClass` semantics.

This is maintainer/agent-facing internals. For the end-user how-to (enabling it in a
project, running it from the web runner / headless / MCP), see
[`Wiki/Rubberduck-Testing-Support.md`](../Wiki/Rubberduck-Testing-Support.md).

Rubberduck references:

- [Rubberduck Unit Testing (wiki)](https://github.com/rubberduck-vba/Rubberduck/wiki/Unit-Testing)
- [Rubberduck source](https://github.com/rubberduck-vba/Rubberduck) — `AssertClass`,
  `AssertHandler`, `IVBETypeLibsAPI.ExecuteCode`

---

## Why a parallel path (not the native convention)

The VCS runner's **native** convention is any parameterless `Public Sub` in a test
module containing at least one `TestAssert` line (see the 2026-05-08 DECISIONS.md
entry, which deliberately rejected `@Test`-style annotations for the native path).
That decision stands. The three hard rules of native discovery each block a standard
Rubberduck module:

1. **Assertion-line gate.** `ScanModuleForTests` skips a module with zero `TestAssert `
   lines. RD tests call `Assert.AreEqual` / `Assert.Succeed` / `Assert.Fail`, so they
   register none.
2. **Public parameterless subs only.** RD test methods are declared `Private Sub`.
3. **`Application.Run` invocation.** `Option Private Module` (in the RD template) makes
   the procs unreachable by name.

Rather than relax the native rules (which would change discovery for every project),
RD support is a **second discovery path that activates only when a module declares
`'@TestModule`**. Modules without that annotation stay entirely on the native fast
path.

---

## Execution model: emulation, not delegation

The VCS runner **never invokes Rubberduck's engine.** Rubberduck captures assertion
outcomes through `AssertHandler.OnAssertCompleted`, an event only its own engine
subscribes to; under a VCS run nobody is listening, so results made against a real
`Rubberduck.AssertClass` would be silently lost.

The bridge is a per-project shim, `StubRdAssert.cls`, obtained through a
`basTestHelpers.CreateTestAssert()` factory. It replicates the `AssertClass` surface the
suite uses and, per call, branches on whether a VCS run is active:

- **VCS run active** → the shim **computes the outcome itself** (emulating `AssertClass`
  value semantics) and reports it to the add-in via `HandleTestOutcome`.
- **Not a VCS run** (RD Test Explorer, or the add-in isn't loaded) → the shim
  **delegates** to a real `Rubberduck.AssertClass` created late-bound, so RD's
  `AssertHandler` captures the result exactly as it does today.

The discriminator is the existing `HandleTestOutcome` / `HandleTestAssertion` public,
which returns `False` when `TestRunner.State <> etrsRunning`. Because the shim owns both
modes, a module runs **identically** under both runners with its test bodies untouched,
and **Rubberduck is never a required dependency** — the shim falls back to `Debug.Assert`
when neither a VCS run nor Rubberduck is present.

The shim and factory live in the *consuming* project (in build-stripped `Stub*` /
`basTestHelpers` modules), not in the add-in. The add-in ships only the discovery,
lifecycle, dispatch, and reporting machinery.

---

## Lifecycle

For an `@TestModule`, the run loop honors the RD lifecycle order: `@ModuleInitialize`
once on entry → for each `@TestMethod`: `@TestInitialize` → test → `@TestCleanup` →
`@ModuleCleanup` once on exit. A lifecycle error folds into the affected test's outcome.
`@IgnoreTest` excludes a method. RD lifecycle is driven for standard-module
`@TestModule`s; class-module `@TestModule`s are out of scope.

## Multi-module dispatch (name-collision workaround)

Every RD module shares lifecycle proc names (`ModuleInitialize`, `TestInitialize`, …).
Access `Application.Run` cannot resolve a proc name duplicated across standard modules
(bare, project-qualified, and module-qualified by-name all fail), so a naive
multi-module run raises runtime error 91 when the ambiguous lifecycle call is swallowed
and the module's `Assert` is never set.

`clsTestRunner.BuildDispatcher` works around this by injecting a temporary standard
module (`modVCSTestDispatch`) containing uniquely-named wrappers (`zz_VCS_Dispatch_N`),
each calling one RD proc **module-qualified** (`[Module].[Proc]` — legal and unambiguous
in compiled VBA). The run loop, `RunTestPhase`, and `InvokeModuleLifecycle` route RD
invocations through `DispatchName` → `BuildRunCmd(wrapper)`; `CleanUp` removes the
module. It is best-effort (falls back to direct `Application.Run` on injection failure),
wraps only standard modules, and reaches `Option Private Module` procs.

---

## Inconclusive outcome (tri-state)

Rubberduck classifies several cases as **Inconclusive, not Failed**:

- `AreEqual` / `AreNotEqual` where the operands' **types differ** (`AreEqual 42, "42"`,
  `AreEqual 42, Null`)
- `AreSame` / `AreNotSame` on **value types** (not object references)
- an explicit `Assert.Inconclusive([message])` call

The runner models this end to end:

- **Model** — `etsInconclusive` in `eTestStatus`; an `eAssertOutcome`
  (`eaoPassed` / `eaoFailed` / `eaoInconclusive`) per assertion. `RecordAssertion`
  (Boolean) still maps True→Passed / False→Failed; `RecordAssertionOutcome` carries the
  tri-state. Each assertion stores `"outcome"` and keeps `"passed"` for backward compat.
- **Classifier** — a test's status is the highest-precedence outcome among its
  assertions: **`Errored > Failed > Inconclusive > Passed`**.
- **Serialization** — `GetResultsAsJson` (summary `"inconclusive"` + per-assertion
  `"outcome"`), `test-state.json` round-trip (`OutcomeCodeToString` /
  `OutcomeCodeFromString`), and JUnit (`INCONCLUSIVE` → `<skipped/>`, the closest
  standard-schema equivalent).
- **UI** — the web runner (`TestRunner/runner.html`) and HTML report
  (`TestRunner/results.html`) render an amber ⚠ status with its own count chip, filter
  tab, and progress/donut segment. `modTestRunnerUI.WebStatusFromRunnerStatus` maps
  `etsInconclusive` → `"inconc"`; `WebOutcomeToken` carries the per-assertion outcome.

### Locked decisions

1. **Precedence:** `Errored > Failed > Inconclusive > Passed`. Inconclusive must never
   mask a real failure.
2. **Record-and-continue:** the shim records each assertion and continues rather than
   stopping at the first non-passing assert. This is the one intentional divergence from
   the RD Test Explorer — an inconclusive-then-fail test reads **Failed** in VCS (later
   Fail wins under precedence) but **Inconclusive** in RD (RD stops at the first
   inconclusive assert). Stop-at-first on the shim path is a deferred fidelity option.
3. **Inconclusive is non-failing:** it does not flip `allPassed` or fail a headless CI
   run, matching RD's default. A run of only inconclusive tests is not `allPassed`
   (nothing passed). A `TreatInconclusiveAsFailure` toggle is deferred.
4. **JUnit mapping:** `INCONCLUSIVE` → `<skipped/>` (NUnit convention; JUnit has no
   native inconclusive).

No `eExportFormatVersion` gating is required: test results live in gitignored
`test-results/` artifacts, not exported source.

---

## Key source references

- `clsTestRunner.cls` — `eTestStatus` (`etsInconclusive`), `eAssertOutcome`,
  `ScanModuleForTests` (annotation-first RD path), the run loop / lifecycle
  (`IsRdModule`, `LifecycleProc`, `RunTestPhase`, `InvokeModuleLifecycle`),
  `RecordAssertionOutcome`, `BuildDispatcher` / `DispatchName` / `RemoveDispatcher`,
  `StatusToString` / `StringToStatus`, `GetResultsAsJson`.
- `clsVersionControl.cls` / `modAPI.bas` — `HandleTestOutcome` (tri-state bridge, returns
  `False` when no run is active).
- `modTestState.bas` — `BuildSummaryFromState`, `OutcomeCodeToString` /
  `OutcomeCodeFromString` round-trip.
- `modTestJUnit.bas` — `INCONCLUSIVE` → `<skipped/>`.
- `modTestRunnerUI.bas` — `WebStatusFromRunnerStatus`, `WebOutcomeToken`.
- `TestRunner/runner.html`, `TestRunner/results.html` — amber ⚠ status rendering.

The shim + codemod (`StubRdAssert.cls`, `basTestHelpers.CreateTestAssert`) live in the
consuming project, not this repo.
