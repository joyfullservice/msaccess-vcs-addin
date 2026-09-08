# Access DPI Layout Probe

This probe measures how Microsoft Access rewrites form layout geometry at
different Windows display scaling settings. It imports the same code-free form
sources into a new scratch database on every run and captures two paths:

1. `plain/` — `LoadFromText` followed immediately by `SaveAsText`.
2. `design-save/` — the same import followed by the open/dirty/close sequence
   used by `InitializeForms`, then `SaveAsText`.

The probe never opens or modifies a production database. Each run gets its own
new `probe.accdb` under the ignored `results/` directory.

## Fixtures

Harvested cases (CRLF, UTF-16-LE; `RightAnchored` is an Access design-save):

- `TabularGrid` — two-row, five-column tabular layout (cross-section group).
- `RightAnchored` — the tabular fixture with Access-saved right anchors.
- `MultiRowGrid` — six-row layout from the add-in's conflict form.
- `StretchAnchored` — three-row layout with horizontally stretched controls.
- `SpacerGrid` — a real multi-row layout containing `EmptyCell` spacers.
- `StackedGrid` — a form containing a genuine stacked layout.

Compact oracle cases: `CustomGap`, `ZeroGap`, `ColumnSpan`, `CrossSection`,
`TabPage`.

All fixtures have code-behind, event handlers, and data bindings removed.
`Caption`, `ControlTipText`, and `StatusBarText` are rewritten to
`DPI Layout Probe`; `Tag` becomes `probe`; hyperlink addresses become a
no-op macro. Distinctive source-database colors are replaced with neutral
dark gray and white, and business control names are remapped.
`prepare_fixtures.py` applies these rules on every harvest. Private source
paths and source-specific substitutions live in the gitignored
`fixture-sources.json`; its shape is:

```json
{
  "sources": {"SpacerGrid.form": "C:\\private\\spacer.form"},
  "replacements": {
    "Name =\"sourceSpecificName\"": "Name =\"genericName\"",
    "BackColor =<source value>": "BackColor =4210752"
  }
}
```

`write_compact_fixtures.py` writes the compact ones. Normal probe runs
consume `fixtures/` only.

## Canonicalizer oracle

`form_geometry.py` is the executable spec for `clsFormGeometryCanonicalizer`.
Why the algorithm looks this way, and the measurements that justified it,
are in [`docs/access-form-geometry.md`](../../docs/access-form-geometry.md).

```text
python .\form_geometry.py .\fixtures --json
python .\validate_canonical.py --corpus <optional-forms-folder>
python .\export_canonical_fixtures.py
```

After a probe run whose `-FixtureDirectory` is
`canonical-fixtures/access-proof`:

```text
python .\Prove-CanonicalRoundtrip.py .\results\<label>
```

Required: `N(P_d(C)) = C` on both paths, where equality means the
geometry signature rather than byte-identical files. The signature covers
named blocks' geometry, form `Width`, and section `Height`; Page blocks
are excluded. The script compares `C` to `N(P_d(C))` and reports `C`
versus `P_d(C)` as a projection diagnostic.
`invariantIdentical` is the proof verdict; `projectionIdentical` is
informational. The proof also requires canonical inputs, unchanged
control and `EmptyCell` counts, idempotence, and no skipped groups.
Proven at 96, 120, 128, 144, 168, and 192 DPI for the six Access-proof
fixtures.
`P_120(C)`, `P_128(C)`, `P_144(C)`, and `P_192(C)` did not equal `C`
(max 108, 60, 90, and 68 twips on `SpacerGrid`); `N` recovered `C`.
At 96 and 168, `P_d(C)` already matched `C`. 128 DPI is custom
133.33%, not a 25-point step; that run required a sign-out. Confirm
`metadata.json` `accessEffectiveDpi` before trusting a run — the
folder label is not the sensor.

On the laptop panel, standard scale changes reached Access immediately
during the original series (often alongside a resolution change). A
133% custom scale required a sign-out. On the lid-closed 34" through a
KVM, one slider-only change did not (`dpi-sanity-125` stayed at 96),
while the later 120-DPI proof reached Access without a sign-out.
For a reproducible series, use the sign-out protocol below. For a
one-off run, skipping it is acceptable only when `metadata.json`
reports the intended DPI. Close every Access window either way.

## First test series

Keep the selected display's resolution fixed (prefer its native resolution)
and vary only Windows scaling.
Recommended labels:

1. `panel-100-a`
2. `panel-125`
3. `panel-150`
4. The display's recommended scaling, such as `panel-200`
5. `panel-100-b`

The repeated 100% run checks that any differences are caused by scaling rather
than an unstable fixture or Access session.

For every run:

1. Save work and close **all** Microsoft Access windows.
2. Change scaling for the selected display.
3. Sign out of Windows and sign back in. Do not leave Access running while
   changing scaling.
4. Keep that display as the primary display, or select it explicitly as shown
   below.
5. Run the probe and enter the run label.

Double-click:

```text
Run-DpiLayoutProbe.cmd
```

The script opens a fresh Access instance, places it on the selected monitor,
records `GetDpiForWindow`, processes every fixture, closes Access, and writes
`metadata.json`. The measured Access DPI—not the label—is the authoritative
scaling value.

## Selecting a non-primary display

From PowerShell:

```powershell
.\Run-DpiLayoutProbe.ps1 -ListMonitors
.\Run-DpiLayoutProbe.ps1 -Label panel-150 -Monitor 1
```

The script refuses to start while another Access process exists. This is
intentional: reusing an existing process could retain the wrong DPI context.

## Compare completed runs

After at least two runs, double-click:

```text
Compare-DpiLayoutProbes.cmd
```

Or specify baseline-first ordering explicitly:

```powershell
python .\Compare-DpiLayoutProbes.py `
  .\results\panel-100-a `
  .\results\panel-125 `
  .\results\panel-150 `
  .\results\panel-200 `
  .\results\panel-100-b
```

The comparison writes both human-readable and JSON reports. It reports:

- Exact and geometry-only equality.
- Added or lost controls.
- Geometry changes and their twip deltas by property.
- Separate `LayoutCached*` changes.
- Inferred inter-column gap distributions.
- Plain import versus the `InitializeForms` design-save path.

Do not rename a completed run directory or edit its contents. If its
label is wrong, leave it in place and use `metadata.json` as the
authoritative DPI record.

## Optional A→B→A test

After the baseline series, a previous run's `plain/` folder can become the next
run's fixture directory:

```powershell
.\Run-DpiLayoutProbe.ps1 `
  -Label chain-150 `
  -Monitor 1 `
  -FixtureDirectory .\results\panel-100-a\plain
```

After changing back to 100%, use `results\chain-150\plain` as the fixture
directory for `chain-100`. This directly tests whether 100%→150%→100% returns
to the original geometry.
