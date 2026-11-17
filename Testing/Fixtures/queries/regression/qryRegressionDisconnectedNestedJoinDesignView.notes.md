# qryRegressionDisconnectedNestedJoinDesignView

**Pins:** a designer query whose join tree is nested on the right
(`A INNER JOIN (B INNER JOIN C ON ...) ON ...`) and which also has a table
outside every join, plus the designer's parenthesized WHERE and HAVING.

From export format 5.2.0:

- The unjoined table comes first and the join tree follows without outer
  parentheses, `FROM tblCurrencies, tblCars INNER JOIN (...) ON ...`, which
  is how Access writes it. The import parser has to split that FROM list at
  the top-level comma before reading the join tree.
- `clsSqlCanonicalizer` removes the designer's wrappers:
  `WHERE (((tblCars.Year)=2011) AND ...)` exports as
  `WHERE (tblCars.Year = 2011) AND ...`. HAVING loses its clause wrapper but
  keeps `(Count(tblCarsColour.Colour)) > 1`: only single values are unwrapped,
  and a function call is not one. The nested join keeps its parentheses.

The fixture input is written the way the designer saves it; the baseline
`.sql` is the 5.2.0 export.
