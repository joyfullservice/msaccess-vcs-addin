# qryRegressionTopPercent

**Pins:** `TOP N PERCENT` must round-trip with the `PERCENT` keyword preserved.

Earlier MSysQueries handling collapsed `TOP 5 PERCENT` to `TOP 5` because the
PERCENT bit (encoded in MSysQueries Attribute 6 / OptionFlag) was not being
serialized into the `.json` metadata, so the importer rebuilt the query
without it.
