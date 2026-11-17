# qryRegressionUpdateCartesianDesignView

**Pins:** an `UPDATE` saved from the designer over a join plus an unjoined
lookup table that supplies a value to the `SET` clause.

From export format 5.2.0 the unjoined table comes first,
`UPDATE tblCurrencies, tblCars INNER JOIN tblCarsPrice ON ...`, matching
Access. Earlier formats wrote `(tblCars INNER JOIN ...), tblCurrencies`.
Both forms are valid and update the same rows; only the order differs.

The `UPDATE` target list goes through the same FROM parser as a `SELECT`,
so this fixture also guards the import side of the top-level comma split.
The designer-style WHERE exports without its redundant parentheses.
