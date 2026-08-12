# `docs/` — Internal reference documentation

Long-form reference docs for maintainers and AI agents working on the
add-in's internals. Distinct from the other documentation venues in
this repo:

- **`Wiki/`** — public-facing user docs (syncs to the GitHub Wiki).
  Audience: end users learning to use the add-in.
- **`AGENTS.md`** (root) — the always-loaded entry point for contributors
  and agents: workflow, invariants, and a routing table into this folder.
  It is budgeted at 150 lines, so depth belongs here, not there.
- **`Version Control.accda.src/AGENTS.md`** and its `vcs-agent-docs/`
  siblings — shipped into every user's export folder. A different
  audience and a separate set of rules; see
  [`agent-docs-maintenance.md`](agent-docs-maintenance.md).
- **`.cursor/rules/*.mdc`** — short, glob-scoped triggers that fire when
  you touch a matching file. They point here rather than restating.
- **`DECISIONS.md`** — append-only journal of architectural decisions
  and the alternatives evaluated.
- **Per-fixture `.notes.md`** under `Testing/Fixtures/` — bug-specific
  context tied to a single test artifact.

## What lives here

Two kinds of reference material: the external systems and formats the
add-in parses, and the add-in's own internals. Docs in this folder are
sustained — updated as understanding evolves rather than written once
and abandoned.

### The add-in's internals

| Doc | Topic |
|---|---|
| [`architecture.md`](architecture.md) | Repository layout, component diagram, `IDbComponent`, the `clsVersionControl` public API and category-scoped sync, core modules, key enums, query-parser entry points, ribbon and hook add-ins. |
| [`export-format-versioning.md`](export-format-versioning.md) | How to change what gets exported: `eExportFormatVersion` gates, `GetExporterRevisions`, schema fingerprints, plus the checklists for adding a component type or an option. |
| [`error-handling.md`](error-handling.md) | The `DebugMode` + `LogUnhandledErrors` + `CatchAny` pattern, and how to read "Unhandled error, likely before `On Error` directive" in a log. |
| [`mcp-runvba.md`](mcp-runvba.md) | Debugging `vcs_run_vba` failures: error-break suppression, the auto-injected `errorLine`, multi-error handlers, and `On Error GoTo -1`. |
| [`testing-strategy.md`](testing-strategy.md) | The three testing layers, the `modTest*` convention, the object round-trip harness, the bug-as-fixture workflow, headless CI runs, and where results land. |
| [`web-test-runner.md`](web-test-runner.md) | `frmVCSTestRunner`, the outbox-polling bridge, the run-command protocol, hydration, and the `modTestRunnerDiag` trace log. |
| [`agent-docs-maintenance.md`](agent-docs-maintenance.md) | Rules for the agent documentation that ships to users, and for this repo's own always-loaded docs. |

### External formats and systems

| Doc | Topic |
|---|---|
| [`access-query-storage.md`](access-query-storage.md) | How Access stores queries (MSysQueries fields, Design View vs SQL View, `LoadFromText` / `SaveAsText` asymmetries, parser invariants and known gaps). |
| [`access-conditional-format.md`](access-conditional-format.md) | The undocumented `ConditionalFormat` / `ConditionalFormat14` binary properties, and how the add-in decodes, stores, and rebuilds them. |

## Plausible future siblings

None of these exist yet — add them when the need arises:

- `access-form-storage.md` — how Access stores forms internally.
- `access-report-storage.md` — same, for reports.
- `access-binary-formats.md` — the `LvProp`, `LvExtra`, and MR2 binary
  blobs the add-in parses.
- `com-ribbon-addin.md` — twinBASIC ribbon DLL architecture, if the
  summary in `architecture.md` outgrows its section.
- `hook-dll-architecture.md` — the export-on-save hook DLLs in `Hook/`.

## When to add a doc here vs. elsewhere

Use this folder when the content is:

- A **long-form reference** (not a short how-to or a one-shot note).
- About **internals, dependencies, or formats** — not user-facing usage.
- **Sustained** — expected to be updated as the system or the team's
  understanding of it evolves.

If the content is one-time architectural rationale, log it in
`DECISIONS.md`. If it's a public how-to for end users, add it to
`Wiki/`. If it's bug-specific context tied to a single test fixture,
use a `.notes.md` companion next to the fixture. If it ships to users
in their export folder, read
[`agent-docs-maintenance.md`](agent-docs-maintenance.md) first.

Every doc in this folder must be reachable — linked from the root
`AGENTS.md` routing table, from a `.cursor` rule, or from the index
above. `modTestRepoDocs` fails when one is orphaned.

The decision to split internal reference material into `docs/` and
keep `Wiki/` for user-facing content is recorded in `DECISIONS.md`
under the 2026-04-27 entry.
