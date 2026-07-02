# qryRegressionSelectStar

**Pins:** Designed-view `SELECT * FROM <table>;` queries that store zero
output-column rows in MSysQueries.

When the user designs a query in Design View and selects only the table-level
`*` row from a single input table, Access stores **no Attribute=6 (output
column) rows** in MSysQueries — only Attribute 0 (designed-marker, NULL
Expression), Attribute 3 (options), Attribute 5 (input table), and Attribute
255 (terminator). The `*` projection is implicit.

Without the empty-field-list guard in `clsQueryComposer.BuildFieldList`, the
composer treats `colFields.Count = 0` as "emit nothing" and produces invalid
SQL of the shape:

```
SELECT
FROM
  tblCars;
```

This corrupts the export and breaks subsequent imports. The fix: when the
field list is empty AND the FROM clause has at least one input table or
join, emit `*`. (Truly tableless `SELECT FROM ;` queries — see
`select/qryNoTable` — keep their empty field list and are unaffected.)

This shape was first observed in the wild on a customer database whose
linked-view-backed queries are all `SELECT * FROM ifportal_vw_<name>;` —
five queries, every one rendered as `SELECT \nFROM ...`. The
`Validate Query SQL Builder` Advanced Tools harness flagged them as the
first non-passing batch when run against that database.
