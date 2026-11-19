# Query Source Files

Each query in `queries/` is exported as two files: a `.sql` holding the statement,
and a `.json` holding everything that SQL cannot express.

## Editing query logic

Edit the `.sql`. It is the sole source of truth for the query's SQL text, and
changing it is the normal way to change what a query does. The file is Access SQL:
it must paste into the Access SQL View and run unchanged.

```sql
SELECT
  tblCustomers.CustomerName,
  IIf([tblOrders].[Shipped], "Yes", "No") AS ShipStatus,
  [tblCustomers].[FirstName] & " " & [tblCustomers].[LastName] AS FullName
FROM
  (tblCustomers
  INNER JOIN tblOrders
    ON tblCustomers.CustomerID = tblOrders.CustomerID)
  INNER JOIN tblOrderLines
    ON (
      (tblOrders.OrderID = tblOrderLines.OrderID)
      AND (tblOrders.Region = tblOrderLines.Region)
    )
WHERE
  (tblOrders.OrderDate >= #1/1/2024#)
  AND (tblCustomers.CustomerName Like "A*");
```

Keep the Access syntax that standard SQL would spell differently; Access rejects or
misreads the standard forms:

- Parentheses around each nested join (`(A INNER JOIN B ON ...) INNER JOIN C`)
- Parentheses around each condition of a multi-condition `ON`, and one more pair
  around the conditions on the same two tables: `ON (a.x = b.x) AND (b.y > 0)` is
  rejected, `ON ((a.x = b.x) AND (b.y > 0))` is not
- `#date#` literals, `*` and `?` wildcards in `Like`, `IIf()` rather than `CASE`,
  and `&` for string concatenation
- Brackets around names with spaces or reserved words (`AS [Option]`)

Layout is yours to choose; the next export rewrites it in the style above without
changing meaning. With export format 5.2.0 or later, export also drops redundant
parentheses from `WHERE` and `HAVING`, including the extra pairs the query designer
adds when it saves a query, so a designer save leaves at most one pair per criterion.
A designer save can still reorder criteria or distribute `AND` over `OR`
(`(A OR B) AND C` becomes `(A AND C) OR (C AND B)`); the logic is unchanged.

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
accepts them. On export format 4.1.2 the `.bas` file is the import source; a
paired `.sql` is applied only when **Force import of original SQL for queries**
is enabled. To make `.sql` the source of truth, switch to export format 5.0+
and run a full export — that writes `.sql` plus `.json` and removes the legacy
files.
