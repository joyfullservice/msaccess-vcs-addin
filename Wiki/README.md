# Wiki folder (repository only)

User-facing markdown in this folder syncs to the [GitHub Wiki](https://github.com/joyfullservice/msaccess-vcs-addin/wiki) when changes merge to the `main` branch (see `.github/workflows/update-wiki.yml`).

**Meta docs (not published):** `README.md`, `PUBLISH.md`, and `img/CAPTURE_CHECKLIST.md` are maintainer-only files excluded from the wiki sync.

## When to update the wiki

- New user-visible option or workflow change
- Export format version bump with user-facing migration steps
- Security-sensitive feature (connections, MCP, hooks)

## When not to update the wiki

- Parser internals → `docs/` in the repository
- Architectural rationale → `DECISIONS.md`
- Agent API details → `AGENTS.md`

## Screenshots

UI screenshots live in `Wiki/img/`. After options UI changes, recapture images referenced from wiki pages and commit them here.

## Release checklist

Before tagging a release, verify [Options](Options.md), [Version 5 Overview](Version-5-Overview.md), and [Home](Home.md) reflect the release version and any new options.
