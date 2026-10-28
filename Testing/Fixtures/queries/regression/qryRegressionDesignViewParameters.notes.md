# qryRegressionDesignViewParameters

Pins the fix for a parameterized query losing its `PARAMETERS` clause when
rebuilt from source through the **Design View** import path.

## The bug

`clsQueryComposer.EmitDesignViewQdef` assembled the structured `.qdef` that
`Application.LoadFromText` consumes, but never emitted a `Begin Parameters`
block. A query that declares parameters *and* takes the Design View import
path (because it carries a `DesignLayout`, or because its shape otherwise
requires Design View) was therefore rebuilt with an empty parameter
collection — the `PARAMETERS` declaration in the `.sql` source was silently
dropped.

The SQL View import path was unaffected, because there the whole statement
(including `PARAMETERS ...;`) is handed to Access as a single SQL memo. So a
parameterized query only lost its parameters when something forced it onto the
Design View path — which is exactly what a designer-built query with a stored
grid layout does. That made this a blind spot: the existing parameter fixtures
were SQL-View shaped and never exercised it.

## What this fixture pins

- A `DesignLayout` in the companion `.json` forces `blnDesignView = True`, so
  the query is rebuilt through `EmitDesignViewQdef` (not the SQL memo path).
- The `.sql` declares two typed parameters.
- Round-trip (import → export) must preserve both parameters, and the
  generated `.qdef` baseline must contain a `Begin Parameters` block whose
  `Flag` values match the declared parameter types.

If `EmitParameters` is removed or the `Begin Parameters` block is emitted in
the wrong position, the `.qdef` drift check fails and the re-exported `.sql`
loses its `PARAMETERS` line.
