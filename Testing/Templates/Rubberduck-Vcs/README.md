# Rubberduck-to-VCS test template

Copy these three exported VBA files into an Access project to make a standard-module
Rubberduck test module runnable through both the Rubberduck Test Explorer and the VCS
test runner. Follow the migration procedure in [Wiki/Testing.md](../../../Wiki/Testing.md).

## Contents

- `StubRdAssert.cls` -- project-side assertion shim.
- `basTestHelpers.bas` -- supplies `CreateTestAssert()`.
- `ExampleTestModule.bas` -- a converted Rubberduck test module.

The shim is not part of the add-in. Keep it and its helper in test-only source and
exclude them from a production build using the consuming project's normal build rules.

## Supported assertion methods

The template implements these `Rubberduck.AssertClass` methods:

- `AreEqual`, `AreNotEqual`
- `AreSame`, `AreNotSame`
- `IsTrue`, `IsFalse`
- `IsNothing`, `IsNotNothing`
- `Succeed`, `Fail`, `Inconclusive`

A suite that uses another assertion method, `Rubberduck.FakesProvider`, or mocks must
extend the project-side shim and verify behavior in both runners before migrating the
suite. The shim implements the VCS-run path itself; when the Rubberduck Test Explorer
runs, it delegates to the installed `Rubberduck.AssertClass`.

This template supports standard-module `@TestModule`s. The VCS runner does not drive
Rubberduck lifecycle methods for class-module `@TestModule`s.
