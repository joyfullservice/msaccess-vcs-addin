# Version Control System

Welcome! This wiki documents installation and usage of the **MSAccess Version Control System (VCS) Add-in** for Microsoft Access — export database objects to text source, collaborate with Git, and build or merge databases from those files.

**Current release:** v5.x (export format 5.0.0). See [What's New in Version 5](Version-5-Overview) for migration notes from older projects.

---

## Getting started

| Page | Description |
|------|-------------|
| [Quick Start](Quick-Start) | Five-minute install → export → build walkthrough |
| [Installation](Installation) | Download, install options, ribbon COM add-in, uninstall |
| [Documentation](Documentation) | Export-only vs collaborative development workflows |

---

## Daily workflow

| Page | Description |
|------|-------------|
| [Options](Options) | Per-project settings (`vcs-options.json`) — export, build, merge, connections |
| [Merge Build](Merge-Build) | Update an existing database from changed source files |
| [Connections](Connections) | `.env` files, `env:conn_*` references, and `APP_ENV` layering |
| [Split Files](Split-Files) | Separate form/report layout from VBA code-behind |
| [Export / Import File Types](Export-Import-File-Types) | Source folder layout and file-type reference |
| [Query Source Files](Query-Source-Files) | `.sql` + `.json` query pairs — what to edit in Git |

---

## Reference

| Page | Description |
|------|-------------|
| [Supported Objects](Supported-Objects) | What can be exported, imported, and built |
| [FAQs](FAQs) | Common questions (ribbon, drift, table data, queries) |
| [Security Considerations](Security-Considerations) | Trust Center, export/build risks, MCP and hooks |

---

## What's new in v5

| Page | Description |
|------|-------------|
| [Version 5 Overview](Version-5-Overview) | Export format 5.0, deterministic queries, binary index, migration steps |

---

## Power features

| Page | Description |
|------|-------------|
| [MCP and Automation](MCP-and-Automation) | AI/agent integration — permissions, security, when to enable |
| [Export on Save Hook](Export-on-Save-Hook) | Experimental, community-contributed DLL to export objects when saved in Access |

---

## Contributing

| Page | Description |
|------|-------------|
| [Editing and Contributing](Editing-and-Contributing) | Fork, build add-in from source, `Deploy`, pull requests |
| [Testing](Testing) | Unit tests, round-trip fixtures, integration database |
| [Regression Testing](Regression-Testing) | Query round-trip harness and fixture contribution workflow |
| [Project Scope](Project-Scope) | What belongs in this add-in vs out of scope |
| [Terminology and Style Guide](Terminology-and-Style-Guide) | Wiki and UI writing conventions |
| [Translation](Translation) | Localization via `T()` and `.po` files |

---

## Documentation map

| Audience | Where to look |
|----------|----------------|
| **End users** | This wiki (synced from the repo [`Wiki/`](https://github.com/joyfullservice/msaccess-vcs-addin/tree/dev/Wiki) folder on `main`) |
| **Contributors** | [CONTRIBUTING.md](https://github.com/joyfullservice/msaccess-vcs-addin/blob/dev/CONTRIBUTING.md) in the repository |
| **Maintainers / AI agents** | [AGENTS.md](https://github.com/joyfullservice/msaccess-vcs-addin/blob/dev/AGENTS.md) — architecture, coding standards, MCP API |
| **Parser / internals** | [`docs/`](https://github.com/joyfullservice/msaccess-vcs-addin/tree/dev/docs) in the repository (not synced to this wiki) |

**Wiki updates:** Edit markdown under `Wiki/` in the GitHub repository and merge to the `main` branch; [GitHub Actions](https://github.com/joyfullservice/msaccess-vcs-addin/blob/main/.github/workflows/update-wiki.yml) publishes changes to this wiki automatically.
