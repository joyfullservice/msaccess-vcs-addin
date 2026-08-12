# Maintaining agent documentation

This repository maintains two separate bodies of agent-facing documentation with
two different audiences and two different sets of rules:

| | Audience | Governed by |
|---|---|---|
| `Version Control.accda.src/AGENTS.md` + `vcs-agent-docs/` | an agent in **somebody else's** Access database | [Part 1](#part-1--documentation-that-ships-to-users) |
| Root `AGENTS.md`, `docs/`, `.cursor/rules/` | contributors and agents in **this** repository | [Part 2](#part-2--this-repositorys-own-documentation) |

Both are budgeted and both are enforced by a test. The shared principle is
progressive disclosure: the always-loaded entry file carries only what is always
relevant, and each reference is pulled in only by an agent with a reason to read
it.

---

# Part 1 — Documentation that ships to users

`Version Control.accda.src/AGENTS.md` and everything under
`Version Control.accda.src/vcs-agent-docs/` are **not** documentation for this
repository. They are embedded as resources and extracted into every user's export
folder on every export (`modResource.ExtractAgentDocs`, called from
`modExport.ExportSource`, `modExport.ExportSingleObject`, and
`clsVersionControl.ExportObject`).

Their audience is an AI agent editing source files in somebody else's Access
database, and they are loaded into that agent's context on essentially every turn.
That makes them a shared, permanently-paid cost, which is why they get a budget and
this file exists.

Read this before adding anything to them.

## The gate

Answer both questions before touching a shipped document. **The default answer is
no change**, and most add-in work does not reach the second question.

1. **Does this change what a user's source files look like, or how they must be
   edited?** Internal refactors, new options, UI work, performance tuning, and
   changes to how the add-in reaches a result do not.
2. **Would an agent editing source files make a mistake without knowing this?** If
   it is discoverable by listing a folder or opening a single file, leave it out.
   Content the model can derive costs tokens and buys nothing.

## Where it goes

Only if both answers are yes:

| The change is | It belongs in |
|---|---|
| An invariant whose violation silently corrupts data, or a change to which file is the source of truth for an object type | the entry `AGENTS.md` |
| A format detail, a procedure, or a symptom-to-fix | the matching `vcs-agent-docs/` reference |
| How the add-in works internally, or why a decision was made | root `AGENTS.md` or `DECISIONS.md` — never the shipped docs |
| Conceptual or how-to material written for people | `Wiki/` — never the shipped docs |

## The budget

**Entry `AGENTS.md`: 150 lines. Each `vcs-agent-docs/*.md`: 110 lines.**

An addition that breaks a budget must remove something in the same edit. This is
the rule that actually holds the line. Without a ceiling and an obligation to pay
for new content, every individually reasonable addition wins and the file ratchets
upward — which is exactly how the entry file reached 605 lines before this was
written.

`modTestAgentDocs` enforces the budgets, so this is a failing test rather than a
convention.

## Edit in place

Find the section that already covers the topic and revise it. Adding a new section
for a variation of something already documented is the main way these files grow.
The pre-split file explained "edit the `.sql`, the `.json` preserves layout, the
add-in generates the `.qdef`" in four separate places because each author added a
section without reading the others.

If you are adding a document, add its filename to
`modResource.GetAgentDocFiles` and link it from the entry file's reference table.
Unlinked documents are read in under 10% of sessions; linked ones in over 90%.

## Never ship these

Each of these was a real defect in the pre-split file.

- **Add-in internals.** `VCSIndex.DumpToJson` was recommended to users, but
  `VCSIndex` lives in the add-in's own VBA project and is unreachable from a user
  database. Do not name modules, classes, or functions that only exist here.
- **Examples taken from this repo.** `modTestEncoding` and
  `TestParseJoinExpression` are our test names, meaningless in a user's project.
  Invent neutral examples instead.
- **Repo-relative paths.** `../Wiki/Connections.md` and
  `../docs/access-conditional-format.md` resolve to nothing once the file is sitting
  in a user's export folder. Use absolute GitHub or wiki URLs.
- **Changelog voice.** "Build *now* auto-removes misplaced duplicates" tells the
  reader about a release they were not present for. Shipped docs describe current
  behavior in the present tense; version history belongs in `DECISIONS.md`.
- **Directory listings and architecture overviews.** Measured as net-negative: they
  raise cost and pull agents into exploring files they did not need to open.
- **Warnings with no alternative.** A wall of prohibitions makes an agent verify its
  work against every one of them. Pair each "don't" with the "do" that replaces it.

## Sizing rationale

The 150 and 110 line budgets are not arbitrary. Published evaluations converge on a
100-150 line entry file paired with a handful of focused reference documents as the
configuration that measurably improves agent performance, with the gains reversing
past that point. Anthropic's Agent Skills specification independently lands in the
same place, capping a `SKILL.md` body at roughly 500 lines and pushing detail into
`references/`.

---

# Part 2 — This repository's own documentation

The root `AGENTS.md` is loaded on **every turn of every session in this repo**, so
it is subject to the same ratchet problem as the shipped entry file — and reached
934 lines before this was written. `docs/` and `.cursor/rules/` carry the depth it
sheds.

## The budgets

| File | Budget | Why |
|---|---|---|
| Root `AGENTS.md` | **150 lines** | Loaded every turn. It is a router, not a manual. |
| Each `.cursor/rules/*.mdc` | **120 lines** | Auto-attached by glob; cheap only while short. |
| Each `docs/*.md` | no limit | Read on demand by an agent that chose to. |

`modTestRepoDocs` enforces the first two, plus two structural rules: every
`docs/` link in the root `AGENTS.md` must resolve, and no `docs/*.md` may be
orphaned from both the root `AGENTS.md` and the `docs/README.md` index.

An addition that breaks a budget must remove something in the same edit, or move
the content to `docs/`.

## Where it goes

| The content is | It belongs in |
|---|---|
| An invariant the compiler does not enforce, whose violation loses data or churns user projects | root `AGENTS.md`, Invariants |
| A pointer telling an agent which reference to open for a given task | root `AGENTS.md`, routing table |
| A procedure, a mechanism, a format, a protocol, or a table of internals | a `docs/*.md` reference |
| Guidance needed only while editing a particular kind of file | a glob-scoped `.cursor/rules/*.mdc` that links to the `docs/` reference |
| Why a decision was made, and what was rejected | `DECISIONS.md` |
| A how-to for end users | `Wiki/` |

Substance goes in `docs/`. Rules stay thin and exist to *trigger* the read at the
right moment, because they are Cursor-only — an agent running under Claude Code or
another harness never sees them, but it will follow a link from `AGENTS.md`.

## Two rules that do most of the work

- **Anything recoverable by listing a folder or opening one source file does not
  belong in the always-loaded set.** Class inventories, folder trees, and enum
  member lists drift out of date and cost tokens on every turn. The `clsDb*` table
  in the old root `AGENTS.md` was already incomplete when it was removed.
- **New depth goes in `docs/`.** Add the file, add a routing-table row in the root
  `AGENTS.md` pointing at it, and add a row to the `docs/README.md` index.
  Unlinked documents are read in under 10% of sessions; linked ones in over 90%.

## Do not duplicate across venues

The single largest source of drift is the same content living in two places and
one copy being updated. Before writing, check whether the topic already has a
home: `rg -l "<distinctive phrase>" AGENTS.md docs .cursor/rules`. If it does,
edit that copy and link to it.

When consolidating, prefer the version that cites a `DECISIONS.md` date or a
specific identifier — it is almost always the more recently maintained one.
