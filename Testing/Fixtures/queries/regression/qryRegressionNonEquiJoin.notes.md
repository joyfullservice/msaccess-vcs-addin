# qryRegressionNonEquiJoin

**Pins:** Self-join with a compound ON clause containing an inequality (`>`).

Existing self-join fixtures (`qryCurrencyCrossRates`,
`qryRegressionSelfJoinUnaliased`) use only equality conditions in the ON
clause. This fixture adds an inequality operator (`A.Reading > B.ReadingDate`)
as the second condition in a compound ON, exercising a different branch in
`ParseJoinExpression`.

Additional patterns exercised:
- Computed expressions referencing both alias prefixes
  (`[a].[readingdate] - [b].[readingdate] AS days`)
- Date literal in a boolean sub-expression (`#11/12/2020#`)
- Inequality in WHERE (`<= 1`)

Derived from `qryMeterReadingsNonEquiJoin` in the MSysQueriesExamples v2
database. Column metadata (ColumnWidth/ColumnOrder) stripped during
sanitization since those values are layout-specific.
