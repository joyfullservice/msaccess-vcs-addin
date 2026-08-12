# Query Source Files

Each query in `queries/` is exported as two files: a `.sql` holding the statement,
and a `.json` holding everything that SQL cannot express.

## Editing query logic

Edit the `.sql`. It is the sole source of truth for the query's SQL text, and
changing it is the normal way to change what a query does.

```sql
SELECT
    CustomerID,
    CustomerName,
    Email
FROM tblCustomers
WHERE Active = True
ORDER BY CustomerName;
```

On import the add-in reads both files, generates a temporary `.qdef` in memory, and
hands it to `LoadFromText`. When the SQL is designer-compatible and the `.json`
carries layout data, the query imports in Design View format with table positions
preserved; otherwise it imports as SQL View. Making the SQL incompatible with the
designer — adding a `UNION` or a subquery, for example — causes the layout data to
be ignored and logs a warning, which is expected rather than an error.

## The companion `.json`

| Key | Meaning |
|-----|---------|
| `QueryType` | 0 Select, 16 Crosstab, 32 Delete, 48 Update, 64 Append |
| `Connect` | Connection string for pass-through queries, usually an `env:` reference |
| `QueryProperties` | Non-default query properties from the `LvProp` binary blob |
| `Columns` | Column metadata, sorted alphabetically for deterministic output |
| `DesignLayout` | Table positions and window dimensions; present only for queries last saved in Design View |
| `Description` | Object description from DAO document properties |
| `Hidden` | True when the query is hidden in the Navigation Pane |

Edit these values directly when you need to change a property the SQL cannot carry.
Keep the file valid JSON.

Pass-through `Connect` values use `env:` references rather than raw ODBC strings;
see `troubleshooting.md` for how those resolve.

## Legacy files

Older projects may still contain `.qdef` or `.bas` query files. Import still
accepts them, and the next export replaces each one with the `.sql` plus `.json`
pair.
