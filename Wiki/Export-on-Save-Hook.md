# Export on Save Hook

> **Experimental / community-contributed.** This feature was contributed by a single user and is not widely used. It may have known bugs or rough edges. Most users should rely on **Export** (and **Fast Save**) from the ribbon instead. This page is kept for reference for those who want to experiment with it.

Optional **Hook DLLs** export database objects to source automatically when you save them in Access — useful when you want Git to stay current without remembering to click **Export**.

---

## Components

| Item | Location |
|------|----------|
| Hook DLLs | `Hook/` folder in the repository (`MSAccessVCSHook_win32.dll`, `MSAccessVCSHook_win64.dll`) |
| Add-in integration | `modExportOnSaveHook` in the add-in |
| Upstream project | [AccessAppHook](https://github.com/bclothier/AccessAppHook) (LGPL-2.1) |

---

## Activation

1. Install the VCS add-in and ribbon as usual.
2. Use **Activate Hook** from the ribbon (or `VCS.ActivateHook` in code) when prompted.
3. Use **Verify Hook** if export-on-save does not trigger.

Exact menu labels may vary by add-in version; see the ribbon **Tools** area.

---

## Security and trust

Hook DLLs load into the Access process. Your environment may require:

- Trusted install location for the add-in and hook DLLs
- COM add-ins enabled ([Installation](Installation))
- Approval of unsigned DLLs in strict environments

Read [Security Considerations](Security-Considerations) before deploying hooks on locked-down PCs.

---

## When to use

| Good fit | Poor fit |
|----------|----------|
| Solo dev who saves often and commits frequently | Shared database with many concurrent editors |
| Small projects | Very large exports where each save would be slow |
| Ensuring source never lags far behind design | CI that expects explicit export timestamps |

Fast Save still applies — only changed objects export, but frequent saves on large forms can add overhead.

---

## Related

- [Quick Start](Quick-Start) — manual export workflow
- [Options](Options) — export settings that affect hook output
