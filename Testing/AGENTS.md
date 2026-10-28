# AGENTS.md - Testing Folder Guide

**Nothing in this folder runs the add-in's own tests, or rebuilds the add-in.**
Both are done against the development copy of `Version Control.accda` in the
repository root, driven over MCP; see
[../docs/agent-test-runs.md](../docs/agent-test-runs.md) and
[../docs/agentic-rebuild.md](../docs/agentic-rebuild.md). A test run pointed at
`Testing.accdb` searches that sample database for tests and reports a vacuous
pass, and opening it to host a rebuild is never the process — if an API call was
refused, those documents explain what the refusal actually means.

What is here is two different testing layers. Keep their roles separate:

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
