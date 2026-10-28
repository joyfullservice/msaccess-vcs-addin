# qryRegressionDesignViewParameterTypes

Companion to `qryRegressionDesignViewParameters`. That fixture pins the core
bug — parameters being dropped on the **Design View** import path. This one
widens the *type* and *name* coverage so a future change to parameter
serialization can't silently regress the less common flags.

## What this fixture adds

`qryRegressionDesignViewParameters` covers only `Long` (flag 4) and
`Text` (flag 10). This fixture exercises four more of the
`ParameterFlagFromType` / `ParameterTypeSql` mappings through the same Design
View round trip:

| Parameter      | Declared type | DAO flag |
| -------------- | ------------- | -------- |
| `pActive`      | `Bit`         | 1        |
| `[Start Date]` | `DateTime`    | 8        |
| `[Min Price]`  | `Currency`    | 5        |
| `[Max Weight]` | `Double`      | 7        |

### Why `Bit` and not `Boolean`

`Bit` is the spelling Access itself writes for a Yes/No parameter. `Boolean`
looks like the natural keyword and reads better, but the ACE parser rejects it
outright with error 3139 ("Syntax error in PARAMETERS clause"), as it does
`Memo`, `OLEObject` and `Decimal`.

This matters beyond cosmetics. On the **SQL View** import path the entire
statement — `PARAMETERS` clause included — is handed to Access as one SQL memo,
so a query declaring a `Boolean` parameter does not merely lose its type: the
whole object fails to import with "Could not create or set the property SQL".
The composer therefore emits `Bit` and accepts `Boolean` only on the way in, so
that source exported by earlier versions of the add-in still rebuilds.

It also pins **name handling on both branches of `SplitParameterToken`**:
`pActive` is declared *unbracketed* (space-delimited name/type split), while
the other three are bracketed. All four names must survive the round trip
verbatim — the unbracketed name must not be over-bracketed on re-export, and
the bracketed names must keep their brackets.

## What this fixture pins

- A `DesignLayout` in the companion `.json` forces `blnDesignView = True`, so
  the query rebuilds through `EmitDesignViewQdef`.
- Round-trip (import → export) must preserve all four parameters with their
  declared types.
- The generated `.qdef` baseline must contain a `Begin Parameters` block whose
  `Flag` values are `1`, `8`, `5`, `7` in declaration order, with `pActive`
  stored unbracketed and the remaining names bracketed.

If `ParameterFlagFromType` or `ParameterTypeSql` mis-maps any of these types,
or `SplitParameterToken` mishandles the unbracketed name, the `.qdef` drift
check fails and the re-exported `.sql` loses or corrupts the affected
parameter.
