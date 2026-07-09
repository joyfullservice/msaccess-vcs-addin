# Editing and Contributing

Thank you for helping improve the MSAccess VCS Add-in. This page summarizes how to set up a development environment, submit changes, and update documentation.

For policy details see [CONTRIBUTING.md](https://github.com/joyfullservice/msaccess-vcs-addin/blob/dev/CONTRIBUTING.md) in the repository.

---

## Quick rules

1. **Pull requests target `dev`** — active development branch.
2. **One focus per PR** — separate bug fixes from unrelated features when possible.
3. **Use a feature branch** on your fork (`feature/my-fix`), not direct commits to your fork's `dev`.
4. **Wiki changes** — edit files under `Wiki/` in the repo; they publish to GitHub Wiki when merged to **`main`** (see below).

---

## Fork and clone

1. Fork [joyfullservice/msaccess-vcs-addin](https://github.com/joyfullservice/msaccess-vcs-addin).
2. Clone your fork locally.
3. Add upstream: `git remote add upstream https://github.com/joyfullservice/msaccess-vcs-addin.git`
4. Branch from `dev`: `git checkout dev` → `git pull upstream dev` → `git checkout -b feature/my-change`

---

## Build the add-in from source

You need a recent **released** add-in installed first (to run **Build From Source**).

1. Clone the repository and open the `dev` branch.
2. In Access, **Build From Source** pointing at `Version Control.accda.src`.
3. Run the built `Version Control.accda` to install the development build.
4. To edit the add-in database itself, open the dev `.accda` and click **Cancel** on the install prompt, then work with the modules and forms directly.

### Development workflow

After making changes (in source files or in the add-in's own VBA project):

1. **Check for compile errors** in the VBA project (**Debug → Compile**).
2. Click **Rebuild Add-In** on the ribbon (or run `VCS.RebuildAddIn`). This rebuilds the add-in from source and installs it locally. Access closes and reopens as part of the rebuild.
3. Click **Run Tests** on the ribbon to run the unit-test suite. The full suite runs quickly, so there is no need to filter it down.
4. If everything passes, click **Export** to write your changes back to `Version Control.accda.src`.
5. Commit **only intentional source changes** (avoid unrelated generated noise) with a clear commit message.
6. Push and open a pull request to **`dev`**.

---

## Testing before you PR

| Layer | Command |
|-------|---------|
| Unit tests | `?VCS.RunTests` or filtered — see [Testing](Testing) |
| Query round-trip | `?VCS.RunRoundtripTests` — see [Regression Testing](Regression-Testing) |
| Export/import change | Relevant `RunTests` modules + round-trip fixtures |

For the rules that decide what counts as a test, see [Testing — How tests are discovered](Testing#how-tests-are-discovered).

---

## Updating the wiki

1. Edit markdown in the repository `Wiki/` folder (same content as [GitHub Wiki](https://github.com/joyfullservice/msaccess-vcs-addin/wiki)).
2. Open a PR to `dev` (or `main` if wiki-only).
3. After merge to **`main`**, the [Update Wiki](https://github.com/joyfullservice/msaccess-vcs-addin/blob/main/.github/workflows/update-wiki.yml) workflow syncs to the live wiki.

See `Wiki/README.md` in the repo for what belongs in wiki vs `docs/` vs `AGENTS.md`.

---

## Scope and design decisions

- [Project Scope](Project-Scope) — what we accept
- [DECISIONS.md](https://github.com/joyfullservice/msaccess-vcs-addin/blob/dev/DECISIONS.md) — architectural journal (repository only)

---

## Related

- [Issues](https://github.com/joyfullservice/msaccess-vcs-addin/issues)
- [Pull requests](https://github.com/joyfullservice/msaccess-vcs-addin/pulls)
- [AGENTS.md](https://github.com/joyfullservice/msaccess-vcs-addin/blob/dev/AGENTS.md) — agent/MCP and coding standards
