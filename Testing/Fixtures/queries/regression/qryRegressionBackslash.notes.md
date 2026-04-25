# qryRegressionBackslash

**Pins:** literal backslash inside a string-concatenation expression must
survive the formatter unchanged.

The expression `tblCars.Manufacturer & "\" & tblCars.Year AS PathLike`
includes a backslash inside a double-quoted string literal. Earlier
formatter logic could mishandle the backslash either as an escape character
or as a path separator, producing a non-idempotent export.
