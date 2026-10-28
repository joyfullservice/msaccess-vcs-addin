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

Small, local edits are feasible: caption text inside quotes, geometry such as
`Left`, `Top`, `Width`, and `Height`, and visual properties like colors and font
sizes. Anything whose expected format you can infer from neighboring values is
usually fine.

Leave these alone:

- The `Version` lines at the top.
- The `Begin`/`End` nesting, which must stay balanced.
- Control `Name` values, which the VBA code binds to by name.
- The order of controls, which drives control indexing and tab order.
- Properties whose format you cannot infer. Guessing here breaks the layout or
  fails the import outright.

For anything structural — adding or removing controls, reordering them, changing
a subform binding — ask the user to make the change in Access instead.

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
