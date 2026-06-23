# Wiki screenshot capture checklist

Maintainer-only — this file is **not** published to the GitHub Wiki (see `.github/workflows/update-wiki.yml`).

Images in this folder are referenced from wiki pages (`![](img/...)`). Capture from a machine with the latest add-in installed and commit here, then merge `Wiki/` to `main` to update the live GitHub Wiki.

**Use PNG** for all UI captures (crisper text than JPEG). JPEG is only acceptable for photos.

## Current images (post–options UI redesign, v5)

| Filename | Source screen | Used on page |
|----------|----------------|--------------|
| `install.png` | Install form | Installation, Quick Start |
| `install-advanced.png` | Install advanced options | Installation |
| `ribbon-export.png` | Ribbon — Export | Quick Start |
| `build-from-source-ribbon.png` | Ribbon — Build | Quick Start |
| `full-export-finished.png` | Console after full export | Quick Start |
| `quick-export-finished.png` | Console after fast export | Quick Start |
| `build-finished.png` | Console after build | Quick Start |
| `github-desktop-changes.png` | Git client example | Quick Start |
| `options-general.png` | Options → General (representative layout) | Options |
| `options-translation.png` | Options → Translation | Translation |
| `export-conflicts.png` | Merge conflict dialog | Merge Build |
| `ribbon-run-tests.png` | Ribbon — Run Tests | Testing |
| `tests-complete.png` | Test run summary screen | Testing |

## Main README (repository root)

The repository's top-level `README.md` uses one animated capture, stored in the **root** `img/` folder (not `Wiki/img/`):

| Filename | Source | Used on |
|----------|--------|---------|
| `img/gui-demo.gif` | Animated export demo | Repository `README.md` |

Recapture this when the ribbon or export console changes significantly.

## Notes

- The redesigned Options dialog uses a left-navigation layout where only part of each section is visible at once. We intentionally show **one representative section** (General) on the Options page rather than a screenshot per section.
- Use consistent Access window scale and theme across shots.
- Update this checklist when images are added or removed.
