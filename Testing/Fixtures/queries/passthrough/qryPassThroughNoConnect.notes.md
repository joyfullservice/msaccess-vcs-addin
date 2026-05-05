# qryPassThroughNoConnect

**Pins:** Pass-through queries must be detected from `QueryType`, even when the
only connection string is Access's minimal `ODBC;` placeholder.

The SQL intentionally contains server-specific variable syntax that Access SQL
formatting cannot parse. Export should treat the query as raw pass-through SQL
and skip `clsSqlFormatter`.
