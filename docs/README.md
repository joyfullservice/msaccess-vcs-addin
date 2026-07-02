# `docs/` — Internal reference documentation

Long-form reference docs for maintainers and AI agents working on the
add-in's internals. Distinct from the other documentation venues in
this repo:

- **`Wiki/`** — public-facing user docs (syncs to the GitHub Wiki).
  Audience: end users learning to use the add-in.
- **`AGENTS.md`** (root and `Version Control.accda.src/`) — workflow,
  coding standards, and agent-specific guidance.
- **`DECISIONS.md`** — append-only journal of architectural decisions
  and the alternatives evaluated.
- **Per-fixture `.notes.md`** under `Testing/Fixtures/` — bug-specific
  context tied to a single test artifact.

## What lives here

Reference material about the systems the add-in depends on, the formats
it parses, and the underlying behaviors that constrain its design. Docs
in this folder are sustained — updated as understanding evolves rather
than written once and abandoned.

| Doc | Topic |
|---|---|
| [`access-query-storage.md`](access-query-storage.md) | How Access stores queries (MSysQueries fields, Design View vs SQL View, `LoadFromText` / `SaveAsText` asymmetries, parser invariants and known gaps). |

## Plausible future siblings

None of these exist yet — add them when the need arises:

- `access-form-storage.md` — how Access stores forms internally.
- `access-report-storage.md` — same, for reports.
- `access-binary-formats.md` — the `LvProp`, `LvExtra`, and MR2 binary
  blobs the add-in parses.
- `com-ribbon-addin.md` — twinBASIC ribbon DLL architecture and the
  `Ribbon/` folder.
- `hook-dll-architecture.md` — the export-on-save hook DLLs in `Hook/`.
- `error-handling-pattern.md` — the `DebugMode` + `LogUnhandledErrors`
  + `CatchAny` system, currently distributed across module/class
  comments and `AGENTS.md`.

## When to add a doc here vs. elsewhere

Use this folder when the content is:

- A **long-form reference** (not a short how-to or a one-shot note).
- About **internals or dependencies** (not user-facing usage).
- **Sustained** — expected to be updated as the system or the team's
  understanding of it evolves.

If the content is one-time architectural rationale, log it in
`DECISIONS.md`. If it's a public how-to for end users, add it to
`Wiki/`. If it's bug-specific context tied to a single test fixture,
use a `.notes.md` companion next to the fixture.

The decision to split internal reference material into `docs/` and
keep `Wiki/` for user-facing content is recorded in `DECISIONS.md`
under the 2026-04-27 entry.
