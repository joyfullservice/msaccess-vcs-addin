# MCP and Automation

The add-in exposes a **public API** (`VCS` / `clsVersionControl`) for scripting, CI, and **MCP** (Model Context Protocol) agents. Automation can export/import objects, run tests, and execute controlled VBA — but **every sensitive capability is disabled by default**.

This page is for **database owners and administrators** deciding what to enable. API details for developers and agents are in the repository [`AGENTS.md`](https://github.com/joyfullservice/msaccess-vcs-addin/blob/dev/AGENTS.md).

---

## What MCP can do (when permitted)

| Capability | Option gate | Risk if misused |
|------------|-------------|-----------------|
| Import objects from source files | **Allow Object Import from Source** (`McpAllowImport`) | Overwrites database objects from untrusted source |
| Call existing public VBA procedures | Part of import/VBA policy | Runs code already in the database |
| Read-only `SELECT` via DAO | **Allow Read-Only SQL Queries** (`McpAllowExecuteSQL`) | Reads data exposed to the add-in connection |
| Arbitrary VBA in a temp module | **Allow Arbitrary VBA Execution** (`McpAllowRunVBA`) | **Highest risk** — full VBA execution as the interactive user |
| Run test suites / round-trip harness | Requires **Allow Arbitrary VBA Execution** | Can modify sandbox objects and fixture folders |

Configure gates under **Options** → **MCP**.

---

## Default posture: all Off

New projects and upgrades leave MCP permissions **Off**. Agents connecting through an MCP server cannot import, run SQL, or execute arbitrary VBA until an administrator explicitly enables the needed flags in Options and saves `vcs-options.json`.

Treat enabling **Allow Arbitrary VBA Execution** like granting **macro/VBA trust** to an automated operator.

---

## Threat model (practical)

1. **Trust the add-in file** — Only install `Version Control.accda` from official [releases](https://github.com/joyfullservice/msaccess-vcs-addin/releases) or builds you compiled from source.
2. **Trust source files** — Import/build runs VBA hooks and object definitions from your `.src` tree. Protect the repository branch permissions.
3. **Session scope** — MCP can use session-scoped option overrides; they should not persist beyond the session without your intent.
4. **UI skipped** — Agent operations may auto-resolve merge conflicts and skip dialogs. Human review happens in Git, not in Access prompts.
5. **Index** — Single-object agent export/import may skip updating `vcs-index.idx` by design; run a normal export afterward for consistent Fast Save.

---

## When to enable each permission

| Permission | Enable when… |
|------------|----------------|
| **Object Import** | You use an agent to apply reviewed PRs to a dev database; source repo is trusted. |
| **Read-Only SQL** | Agent needs to inspect data for debugging; no write access required. |
| **Arbitrary VBA** | You run `VCS.RunTests`, `VCS.RunRoundtripTests`, or custom automation via `RunVBA`; dev machine only. |
| **VBA function calls** | Agent should call existing project APIs without injecting new code. |

For production databases, prefer **no MCP permissions** or a dedicated sandbox copy.

---

## Related security topics

- [Security Considerations](Security-Considerations) — Trust Center, hooks, export/build
- [Connections](Connections) — keep `.env` out of Git
- [Export on Save Hook](Export-on-Save-Hook) — optional DLL

---

## Further reading

- Repository [`AGENTS.md`](https://github.com/joyfullservice/msaccess-vcs-addin/blob/dev/AGENTS.md) — `RunVBA`, `errorLine`, `ExportObject`, `RunTests` filters
- [`clsVersionControl`](https://github.com/joyfullservice/msaccess-vcs-addin/blob/dev/Version%20Control.accda.src/modules/API/clsVersionControl.cls) — public API implementation
