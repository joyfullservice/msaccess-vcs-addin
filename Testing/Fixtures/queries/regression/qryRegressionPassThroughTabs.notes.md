# qryRegressionPassThroughTabs

**Pins:** [Issue #786](https://github.com/joyfullservice/msaccess-vcs-addin/issues/786)

Pass-through SQL imported through a SQL View `.qdef` must preserve leading tabs.
Literal tabs inside `dbMemo "SQL"` are consumed as qdef whitespace by
`Application.LoadFromText`, so the emitter must serialize them as octal `\011`.

The fixture mixes four-space, one-tab, and two-tab indentation. Its `.qdef`
baseline also places an escape near a wrapped segment boundary.
