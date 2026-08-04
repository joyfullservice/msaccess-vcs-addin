# Rubberduck Testing Support

The VCS test runner can discover and run [Rubberduck](https://github.com/rubberduck-vba/Rubberduck) unit-test suites written with the `@TestModule` / `@TestMethod` convention. A suite can then run headlessly, through the web runner, or via MCP and report alongside native `TestAssert` tests, without rewriting its test bodies.

This is a second, opt-in discovery path: it activates only for a module that declares `'@TestModule`. Modules without that annotation use the native convention documented in [Testing](Testing). The VCS runner recognizes Rubberduck annotations but is not a replacement for the Rubberduck Test Explorer, and Rubberduck is never a required dependency.

See the Rubberduck [Unit Testing wiki](https://github.com/rubberduck-vba/Rubberduck/wiki/Unit-Testing) for how to author these tests.

## How It Works

Rubberduck reports each assertion through an internal event that only its own engine listens to, so the VCS runner cannot reuse the real `Rubberduck.AssertClass` directly. Instead, each project supplies a small **shim** (`StubRdAssert.cls`, obtained through a `CreateTestAssert()` factory) that replicates the supported `AssertClass` surface. Per call the shim branches:

- **Under a VCS run**: the shim computes the outcome itself and reports it to the add-in.
- **Under the Rubberduck Test Explorer** (or with the add-in not loaded): the shim delegates to a real `Rubberduck.AssertClass`, so Rubberduck captures results exactly as before.

Because the shim owns both modes, the same module normally produces matching results under both runners. The VCS runner intentionally differs when an Inconclusive assertion is followed by a later failure; see [The `Inconclusive` Status](#the-inconclusive-status). The shim lives in your project, in a build-stripped `Stub*` module, not in the add-in.

When it runs an `@TestModule`, the VCS runner honors the Rubberduck lifecycle: `@ModuleInitialize` once; then, for each test, `@TestInitialize`, test, and `@TestCleanup`; and finally `@ModuleCleanup`. `@IgnoreTest` methods are skipped.

## Migrate One Module

Start with one fast, deterministic **standard-module** test suite. Keep your existing Rubberduck module and test bodies; make the mechanical changes below, then verify both runners before converting more modules.

1. Copy the maintained [Rubberduck-to-VCS template](https://github.com/joyfullservice/msaccess-vcs-addin/tree/dev/Testing/Templates/Rubberduck-Vcs) into your source project: `StubRdAssert.cls` and `basTestHelpers.bas`. The template's `ExampleTestModule.bas` is a complete converted example.
2. Exclude `StubRdAssert` and `basTestHelpers` from your production build using your project's normal test-only build rules.
3. In the test module, replace the compile-time Rubberduck assertion field with a late-bound one and initialize it through the factory:

   ```vba
   ' Before
   Private Assert As Rubberduck.AssertClass

   '@ModuleInitialize
   Private Sub ModuleInitialize()
     Set Assert = New Rubberduck.AssertClass
   End Sub

   ' After
   Private Assert As Object

   '@ModuleInitialize
   Public Sub ModuleInitialize()
     Set Assert = CreateTestAssert()
   End Sub
   ```

4. Change each annotated `@TestMethod`, `@ModuleInitialize`, `@ModuleCleanup`, `@TestInitialize`, and `@TestCleanup` procedure to `Public` so the VCS runner's temporary dispatcher can call them. Each must be callable with no arguments; parameterized Rubberduck test methods are not supported by the VCS runner. Leave unannotated helpers `Private`. Keep `Option Private Module` if your test project already uses it: it still prevents the module's public members from being exposed outside the project. Preserve every Rubberduck annotation and test body.
5. Compile the project. Run the converted module in the Rubberduck Test Explorer first, then run that module by name through `VCS.RunTests` or `VCS.RunTestsHeadless`.
6. Compare results before migrating the next module. If the module uses an assertion method, `FakesProvider`, or mocks not covered by the template, extend the project-side shim and verify that behavior in both runners first.

The template supports the listed `AssertClass` subset only; it does not make every Rubberduck API available. It also supports standard-module `@TestModule`s only; class-module Rubberduck lifecycles are not driven by the VCS runner.

## Verify The First Conversion

Use one converted module as the acceptance check before migrating a suite:

1. Compile the database in Access. The VCS runner stops before executing tests when the project has compile errors.
2. Run that module in the Rubberduck Test Explorer and record its pass, fail, inconclusive, and skipped counts.
3. In the Access Immediate Window, run the same module through the VCS runner:

   ```vba
   ?VCS.RunTests("ExampleTestModule")
   ```

   With the web runner enabled, this opens the test tree; select the module and choose **Run**. Otherwise, follow the result in the `frmVCSMain` console.
4. For an unattended check, run the same filter through `VCS.RunTestsHeadless("ExampleTestModule")`. Its returned JSON includes `allPassed`, `cancelled`, `junitPath`, and `statePath`; it always writes JUnit XML.
5. Compare the two runners' results, allowing for the documented record-and-continue Inconclusive case below. Inspect `<export-folder>\test-results\test-results.xml`, `test-state.json`, and, when enabled, `test-results.html` for the persisted result.

For MCP-driven runs, enable the relevant permissions under **Options** -> **MCP** first; see [MCP and Automation](MCP-and-Automation). Run against a development copy, not a production database.

## Filtering And Skipped Tests

VCS filters use module names, `@Folder`, procedure names, and explicit `'@Tag("...")` annotations. A Rubberduck `@TestMethod("Category")` argument is **not** automatically a VCS tag. Add a `@Tag` annotation when a Rubberduck category must also be selectable through `VCS.RunTests` or `VCS.RunTestsHeadless`:

```vba
'@TestMethod("Integration")
Public Sub TestRemoteService()
  '@Tag("integration")
  ' ...
End Sub
```

`@IgnoreTest` is discovered as **Skipped** rather than omitted. The VCS runner does not execute it, preserves an optional quoted reason, shows it in the test tree, and emits it as JUnit `<skipped/>`. A module containing only ignored tests does not run its lifecycle hooks.

## The `Inconclusive` Status

Rubberduck treats some assertions as **Inconclusive** rather than passed or failed:

- `AreEqual` / `AreNotEqual` where the two operands have **different types**, such as `AreEqual 42, "42"` or `AreEqual 42, Null`
- `AreSame` / `AreNotSame` on **value types**, not object references
- An explicit `Assert.Inconclusive([message])` call

The VCS runner models this as a first-class status, shown amber in the web runner and HTML report, with its own count, filter, and progress segment. Behavior to know:

- **Precedence** within a test: `Errored > Failed > Inconclusive > Passed`; an inconclusive assertion never masks a real failure.
- **Inconclusive is non-failing.** It does not flip `allPassed` or fail a headless CI run, matching Rubberduck's default. A run of only inconclusive tests is still not `allPassed`, because nothing passed.
- **JUnit:** an inconclusive test serializes to `<skipped/>`, the closest standard-schema equivalent.

> **One intentional divergence from the Rubberduck Test Explorer:** the VCS runner records every assertion and continues, whereas Rubberduck stops at the first non-passing assertion. A test that is inconclusive and then would fail reports **Failed** under the VCS runner, because the later failure wins under precedence, but **Inconclusive** under Rubberduck. Author tests to reach a single conclusive outcome and this never arises.

## Running Tests

Once a module is shim-migrated, it runs through the same entry points as any other VCS test: [`VCS.RunTests`](Testing#run-from-access), [`VCS.RunTestsHeadless`](Testing#headless-ci--automation), the ribbon **Run Tests** button, and [MCP](Testing#mcp--agents). Filter by module, folder, procedure, or tag exactly as with native tests.

## Internals

Maintainers and agents: the discovery path, lifecycle driver, multi-module dispatcher, which sidesteps an Access `Application.Run` name collision on shared lifecycle procedure names, and the tri-state model are documented in [`docs/rubberduck-test-runner.md`](https://github.com/joyfullservice/msaccess-vcs-addin/blob/dev/docs/rubberduck-test-runner.md).
