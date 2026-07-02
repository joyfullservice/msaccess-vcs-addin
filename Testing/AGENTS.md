# AGENTS.md - Testing Folder Guide

This folder contains two different testing layers. Keep their roles separate:

- `Fixtures/` is the canonical object round-trip regression corpus used by
  `VCS.RunRoundtripTests`. Query fixtures live under `Fixtures/queries/` as
  `.sql` + `.json` pairs, with optional `.notes.md` files for regression
  context. Add bug-as-fixture cases here.
- `Testing.accdb.src/` is the sample Access database source used for
  whole-database build/export integration testing. Do not treat it as the
  primary fixture store for `VCS.RunRoundtripTests`.

If you are working on query export/import, `clsDbQuery`, `clsQueryComposer`, or
the round-trip harness, start with `Fixtures/README.md` and the query fixtures
under `Fixtures/queries/`.
