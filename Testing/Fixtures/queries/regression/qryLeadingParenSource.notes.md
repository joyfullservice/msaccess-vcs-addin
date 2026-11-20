# qryLeadingParenSource

Issue [#815](https://github.com/joyfullservice/msaccess-vcs-addin/issues/815):
a source name that starts with `(` was treated as already bracketed, so the
export wrote `FROM (Old) Clients`, which Access cannot parse.

Depends on `_scaffold/(Old) Clients`, so the source resolves when the SQL is
pasted back into Access.
