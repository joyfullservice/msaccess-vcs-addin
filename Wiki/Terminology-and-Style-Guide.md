This page contains style and terminology for how to describe elements within the Wiki, Code, and Interface for a consistent approach.

This is a constant work in progress, and is not complete. Items added here were added for clarity where a universal consensus may not exist.

Use active voice for documentation.

Where not found here, use of [Google's Style Guide](https://developers.google.com/style/) is generally accepted; though some deviations from their guide have been made in favor of other methods where appropriate.
<p>
<p>

# Action / Descriptor Terms
The following are either excerpted or differ from the above style guide. Use these methods first. If you don't find what you need here, go to the above guide(s).

|Item|Term|When used
|-|:-:|-
|Checkbox **State**|**On** (☑) <p>Within `Code`: `True` |Checkbox has a check in it. <p>**For user-facing content:** Use **On** (TagTips, Form Text, Wiki Documentation, etc.). <p>**For `code` comments:** You may use `True`, as you will be working with the `boolean` value directly.<p>**NOTE:** This differs from the *Action* performed on the **Checkbox**.
||**Off** (☐) <p>Within `Code`: `False`|Checkbox is cleared. <p>**For user-facing content:** Use **Off** (TagTips, Form Text, Wiki Documentation, etc.). <p>**For `code` comments:** You may use `False`, as you will be working with the `boolean` value directly.<p>**NOTE:** This differs from the *Action* performed on the **Checkbox**.
||**`Null`**|In some cases, [checkboxes](https://docs.microsoft.com/en-us/office/vba/api/access.checkbox.triplestate) may allow a `TripleState` of `Null`; this may also be referred to as `Indeterminate`. See also [Tri-State](https://en.wikipedia.org/wiki/Checkbox#Tri-state%20checkbox) `Null`. <p>*Example:* In permission lists, if a parent permission includes children, and some of the children are set to allow, and some are off/cleared, the parent may appear in an indeterminate state.<p>All cases use: `Null`.<p> Appears as a square or gray interior inside the checkbox instead of a checkmark or cleared state.
|UI **Action**|Select|Chose the **(Option name)**, set **(Option name)** checkbox state to `True`.<p>*Example:* Select **Use Fast Save** to enable a significant performance boost.
||Clear|Clear **(Option name)** checkbox (set state to `False`), remove text.<p>*Example:* Clear **Use Fast Save** to export all elements in your file every time.
|Keyboard Keystrokes|\<kbd>|Enclose keystrokes with the \<kbd></kbd> tag. To group simultaneous keystrokes, use the + charachter.<p>*Example:*<p> `Press <kbd>[Shift]+[Q]</kbd> to do a fancy thing.` <p>will render like this:<p> Press <kbd>[Shift]+[Q]</kbd> to do a fancy thing.
||<kbd>[]</kbd>|Use of Square Brackets ([ ]) is desirable, especially within code comments as the \<kbd> tag is not rendered in the VBA IDE, and to disambiguate between combined keystrokes and the <kbd>[+]</kbd> key.<p> *Example in Code:* ` ' To break the execution of this function, Press [Ctrl]+[Break] or [Esc]` <p> *Example Wiki:* To close Access immediately, press <kbd>[Alt]+[F4]</kbd>
|Text entry in fields||Enclose text entered by users with \` charachter (code formatting). <p> *Example:* Type `Potato` into the **Desired Food** box.
|Referencing Field|**Bold**|To refer to a field, use **Bold** (Markdown= \**[stuff]**)
|Action||Generally, the action performed does not need to be highlighted. Instead, highlight the **Field** or the **Menu Item**. <P>*Example:* Open **Options**, select the **Export** section in the left navigation, and set **Sanitize Level**.

# Version 5 terminology

Use these terms consistently in wiki pages, UI strings (`T()`), and contributor docs:

| Term | Meaning |
|------|---------|
| **Export format version** | Numeric setting (`ExportFormatVersion` in `vcs-options.json`) that gates export behavior changes. Import remains backward compatible with older formats. |
| **Deterministic query export** | Default query export as a `.sql` + `.json` pair reconstructed from MSysQueries (not legacy `.qdef` as the primary artifact). |
| **Merge build** | Partial import that updates only changed source files into an existing database (not a “new in v4” feature). |
| **`vcs-index.idx`** | Binary change-tracking index paired with the database file; typically gitignored. |
| **`env:conn_*`** | Placeholder in exported JSON for connection strings resolved from `.env` at build time. |
| **MCP session** | Scoped agent automation with optional permission flags (`McpAllowImport`, etc.); all off by default. |
| **Options sections** | Left-nav areas: General, Export, Build, Table Data, Databases, MCP, Translation, Advanced, Defaults (not legacy tab names only). |
| **Fast Save** | Export only objects changed since the last export (uses the index). |
