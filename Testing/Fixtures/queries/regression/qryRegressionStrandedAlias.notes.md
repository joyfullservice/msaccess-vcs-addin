# qryRegressionStrandedAlias

**Pins:** the "stranded cartesian-style" data-loss pattern observed
across ~26 production queries during a full-database build / re-export.

A multi-table inner-join chain whose outermost join is a self-aliased
LEFT JOIN to one of the already-joined tables (`tblCars AS carCopy`).
When the Pass 1 export drops the `Alias = carCopy` line from the
.qdef input table list, the importer rebuilds an MSysQueries
Attribute 7 row where `leftTable == rightTable == tblCars`. On the
next export `BuildJoinChain` cannot attach this self-referential join
to the accumulator (the algorithm has no way to distinguish "tblCars"
the base input from "tblCars" the aliased input), and
`BuildFromClause` appends the trailing tables Cartesian-style with no
`ON` clause -- silently dropping the join condition.

The original production failure was a 14-table LEFT JOIN chain with a
trailing `<table> AS <table>_1` self-aliased join. This fixture
distills the same shape down to four tables.

If this fixture starts producing a Pass 2 != Pass 1 diff that loses
the `AS carCopy` clause and appends `tblCars` to a comma list at the
bottom of the FROM, the bug is back.
