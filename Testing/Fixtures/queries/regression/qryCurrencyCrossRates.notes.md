# qryCurrencyCrossRates

**Pins:** self-join with table aliases must round-trip with `Alias =`
emission preserved.

This query joins `tblCurrencyExchange` to itself (`AS a` / `AS b`) to compute
cross-rate pairs. It is the smallest reproducer for the bug that affected
~723 queries in `db-sec` (including `qryBucketCombineRandUCos` and
`qdelTempInvByQtrSubMergerAcquisition3`).

The failure mode: when an aliased input table's `Alias =` line is dropped
during export, MSysQueries Attribute 7 collapses `leftTable == rightTable`
on the next import, and `BuildJoinChain` crashes when re-exporting. The
fixture exercises the alias-emission path on both join inputs.

The harness's pass-1 vs pass-2 idempotency check would have caught this bug
the first time it shipped.
