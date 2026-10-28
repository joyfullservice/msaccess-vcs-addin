# custom_combo_list

Regression fixture for a **user-defined** combo box list (as opposed to the
Access-managed Zoom combo in `office_links_zoom`).

- **Custom Combo** (Type 4, `BuiltIn:false`) carries a `List` array
  (`Alpha`, `Beta`, `Gamma`). The list *items* are the data that must survive a
  round-trip; `ListCount` is derived and intentionally **not** serialized (it is
  read-only, redundant with `List`, and reading it on Access-managed combos
  raises `-2147467259`).

What this pins:

1. **Import repopulates the list.** The combo-list import branch in
   `clsDbCommandBar.BuildControls` must assign the live control to the typed
   `cbcItem` reference (`Set cbcItem = objItem`) before calling `AddItem`. A
   prior reversed assignment (`Set objItem = cbcItem`) left `cbcItem` as
   `Nothing`, so `AddItem` failed silently under `On Error Resume Next` and the
   list was dropped. The test asserts `ListCount = 3` and the item values after
   import.
2. **The `List` round-trips idempotently.** Two consecutive exports must be
   byte-identical, and the exported control must still carry a 3-item `List`.
