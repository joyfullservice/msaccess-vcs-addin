# Publishing wiki updates

User-facing wiki pages in this folder sync to https://github.com/joyfullservice/msaccess-vcs-addin/wiki when changes land on the **`main`** branch. This file, `README.md`, and `img/CAPTURE_CHECKLIST.md` are **excluded** from sync (maintainer docs only).

## Steps

1. Merge wiki changes to `main` (wiki-only PR is fine).
2. Confirm the [Update Wiki](https://github.com/joyfullservice/msaccess-vcs-addin/actions/workflows/update-wiki.yml) workflow ran successfully.
3. Spot-check [Home](https://github.com/joyfullservice/msaccess-vcs-addin/wiki) and new pages (Version 5 Overview, MCP, Testing, etc.).
4. If sync did not run, use **workflow_dispatch** on Update Wiki manually.

## Screenshots

Screenshots live under `Wiki/img/` (PNG preferred). When the UI changes, recapture the affected images and commit them. See [img/CAPTURE_CHECKLIST.md](img/CAPTURE_CHECKLIST.md) for the current image inventory and capture conventions.

## Link check (local)

From repo root, verify relative `Page-Name` links match a `Wiki/Page-Name.md` file.
