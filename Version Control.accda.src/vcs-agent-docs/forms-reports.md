# Form and Report Layout Files

Covers `.form` and `.report` files (legacy projects use `.bas`) and their companion
`.json`. For the `.cls` code-behind, see `vba-modules.md`.

## The SaveAsText format

These files use Access's `SaveAsText` format. It is not VBA, and it is largely
undocumented:

```
Version =21                              <- Access version marker
VersionRequired =20
Begin Form                               <- Form or Report
    RecordSelectors = NotDefault
    Caption ="My Form Title"
    Begin
        Begin Label                      <- A control
            Left =1234
            Top =567
            Width =2000
            Height =300
            Name ="lblTitle"
            Caption ="Welcome"
        End
        Begin CommandButton
            Name ="cmdOK"
            Caption ="OK"
            OnClick ="[Event Procedure]"
        End
    End
End
CodeBehindForm                           <- VBA section, when layout is not split
Attribute VB_Name = "Form_frmMyForm"
```

## What is safe to change

Property edits are safe wherever you can infer the expected format from
neighboring values: caption text inside quotes, geometry such as `Left`, `Top`,
`Width`, and `Height`, colors, and font sizes. Guessing at a property whose
format you cannot infer breaks the layout or fails the import, so leave those
alone, along with the `Version` lines at the top. Keep `Begin`/`End` balanced
throughout.

## Adding, removing, and renaming controls

Structural edits work if you copy rather than compose. The property set Access
expects for a control type is not something to invent, so use an existing
control of the same type in the same file as your template.

To add a control:

1. Copy an entire `Begin <Type>` ... `End` block, including any nested
   `Begin`/`End` subtree. An attached label lives *inside* its text box, combo
   box, or check box, and an option group's buttons live inside the group.
2. Paste it into the intended section's inner `Begin` block. Appending at the
   end of the section is the safest placement.
3. Give the new control, and every nested control copied with it, a `Name` that
   is unique within the file.
4. Set the geometry, and update the matching `LayoutCached*` values to agree
   with it — including on any control you shift to make room. The form's `Width`
   and the enclosing section's `Height` bound the visible area and do not grow
   to fit, so raise them if the new control extends past them. On a form laid
   out in Layout view, keep the grid bookkeeping (`LayoutGroup`, `GroupTable`,
   `RowStart`/`ColumnEnd`) consistent with the row you copied into.
5. Set `TabIndex` to a value not already used in that section. `TabIndex`, not
   file order, determines tab order.
6. If you kept an event property such as `OnClick ="[Event Procedure]"`, add the
   handler to the `.cls` code-behind. Access binds handlers by control name, so
   `Private Sub cmdSave_Click()` pairs with `Name ="cmdSave"`.

To remove a control, delete its whole block including nested children, then
remove its event handlers and any references to it from the `.cls`.

Renaming means updating every reference to the old name: handler procedure names
and code in the `.cls`, and any key for that control in the companion `.json`.

## Conditional formatting

With **Decode conditional formatting to JSON** enabled (export format 5.0.0 and
later), the opaque `ConditionalFormat` and `ConditionalFormat14` hex blocks are
stripped out of each control and stored as readable rules in the companion `.json`,
under `Items.ConditionalFormatting` keyed by control name. Edit the rules there
rather than the hex.

With the option disabled, the binary blocks stay inline in the layout file. Either
way the add-in rebuilds them on import.

On import the JSON wins. If a control has a JSON entry, any stale inline block for
that control is stripped before the rebuilt block is injected, which keeps repeated
imports idempotent. A control with an inline block but no JSON entry is left
untouched.

The binary layout and the exact boundary of what survives a decode/rebuild cycle
are documented in
[docs/access-conditional-format.md](https://github.com/joyfullservice/msaccess-vcs-addin/blob/dev/docs/access-conditional-format.md).

## Recovering from a bad edit

If a form or report looks wrong after import, the usual causes are an unbalanced
`Begin`/`End` pair or a mangled property value. Restore the `.form` or `.report`
file from git and redo the change in smaller steps.

To check a structural edit, ask for a re-export after the merge build. A
`git diff` on the file then shows what Access normalized or dropped.
