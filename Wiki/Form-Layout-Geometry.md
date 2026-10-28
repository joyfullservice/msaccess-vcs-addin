# Form Layout Geometry

From export format **5.1**, the add-in rewrites the position and size values inside `.form` files onto a fixed 60-twip grid so that the same form exports identically on machines with different display scaling.

You only need this page if you have seen a warning like this in an export log:

```text
WARNING: Form layout geometry left unchanged in frmInvoiceTotals: layout
table 4, group 4, columns 0 to 2. These tracks do not record the sizes
needed to compute coordinates that are independent of display scaling, so
this group is exported exactly as Access saved it. Its values may differ
between machines that run at different display scaling (DPI).
```

**Nothing is broken and no data is lost.** It is a notice that one part of one form could not be made scaling-independent.

---

## Why form dimensions drift at all

Access does not store layout geometry the way you authored it. When a form using a **layout** (Tabular or Stacked, the grid you get from *Arrange → Tabular*) is opened in Design View, Access re-solves every cell's `Left`, `Top`, `Width` and `Height` from the display scaling in effect at that moment, then saves those numbers.

Export opens forms in Design View, so the values written to the `.form` file are the ones your monitor produced. A colleague at a different scaling exports the same untouched form and gets different numbers. That produced diffs nobody made — historically the most common source of noise in a forms-heavy project.

Export format 5.1 fixes this by rounding the values back to the grid Access would have used at 100% scaling. Free-positioned controls (those not in a layout) are authored, not solved, so they never drifted and are left alone.

---

## What the warning means

The add-in does not round each number independently. Two values from different scalings can land either side of a rounding midpoint, which would make the churn worse rather than better. Instead it reconstructs the layout's **grid** — the position of each column and row boundary and the size of each track — and then recomputes each control from that grid.

Reconstructing the grid needs evidence. A column's width is known because some control occupies that column alone. When a control spans several columns and no other control occupies any of them individually, and another spanning control overlaps the same columns with different boundaries, there is no way to tell where one column ends and the next begins. The total is known; the split is not.

In that situation the add-in **stops and changes nothing** in that group, rather than guessing at a boundary and moving your controls. The affected controls keep the exact values Access wrote, which means they may still differ between machines at different scaling — the original problem, confined to that one group.

Reading the message:

| Part | Meaning |
|------|---------|
| `layout table 4, group 4` | Which layout on the form. These are Access's internal `GroupTable` and `LayoutGroup` identifiers, not a count or a fraction |
| `columns 0 to 2` / `rows 0 to 2` | The range of grid tracks that could not be resolved |
| The axis implied by "columns" or "rows" | Horizontal or vertical; each is solved separately |

One warning appears per layout group per axis, so a form with two affected layouts logs two warnings on every export.

---

## What to do about it

Most people should do nothing. The warning is informational, and in a typical project it affects a small number of layouts.

It is worth acting if the form's geometry keeps showing up in diffs and you want that to stop. The cause is almost always ambiguity in the layout's grid, and giving the grid something to measure resolves it.

**Open the form in Design View and look at the layout with *Arrange → Select Layout*.** Then try one of these:

1. **Reduce spans that overlap each other.** Look for two controls that each cover several of the same columns but start and end in different places. Making them share the same boundaries, or giving one of them a single column, is usually enough.
2. **Merge the columns you are not using.** If a control spans three columns only because the columns to its left are empty leftovers from an earlier edit, select those cells and use *Arrange → Merge*. One cell instead of three removes the ambiguity entirely.
3. **Rebuild the layout.** Select the controls, *Arrange → Remove Layout*, then re-apply *Tabular* or *Stacked*. Access rebuilds the grid cleanly. Check the form's appearance afterwards, as this can change spacing.

After changing the form, export again and confirm the warning is gone.

If you would rather not touch the form, you can also turn the whole feature off by setting the export format back to **5.0** in *Options → Export*. That stops all layout geometry normalization for the project, not just for the affected group, so the drift returns everywhere. Only do this if 5.1 causes a problem you cannot otherwise resolve.

---

## Why the first 5.1 export produces a large diff

The first export after adopting format 5.1 rewrites layout geometry throughout the project, because it is moving values onto the grid for the first time. In a measured 400-form project, 43 forms changed, and the largest single control movement was 180 twips (about 3 mm).

Review that diff once, commit it, and subsequent exports are stable. This is a one-time cost, not ongoing churn.

---

## Troubleshooting

| Symptom | Things to try |
|---------|----------------|
| The warning appears on every export | Expected until the form's layout is changed. See *What to do about it* above |
| Form dimensions still differ between developers | Check whether the affected controls are in a group the log warns about. If there is no warning, make sure every developer is on export format 5.1 (*Options → Export*) |
| A control looks slightly different after upgrading to 5.1 | Movement of up to a few millimetres is expected on the first export. If a control now clips or overlaps, report it as a bug with the form's `.form` file |
| Report dimensions still drift | Reports are not normalized. The same layout properties exist on reports, but the behavior has not been verified for them |

---

## See also

- [Options](Options) — export format version
- [Export-Import File Types](Export-Import-File-Types) — what `.form` files contain
- [FAQs](FAQs) — other sources of unexpected diffs
- [Merge / Build](Merge-Build) — rebuilding a database from source
