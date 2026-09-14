# VBA performance techniques

This document covers **how to make hot paths faster** in the add-in. To find where
time went in a completed export or build, read
[perf-diagnostics.md](perf-diagnostics.md) instead — that file explains the JSON
views in `logs/*.perf.json`; this one explains the idioms that showed up when those
views pointed at VBA loops.

Every performance claim below was validated by measurement on a real object through
the real call path, not by reasoning alone. Some neutral changes are retained because
they simplify state or make invalid behavior harder to express; those are labeled as
maintainability choices rather than speedups. Several initial hypotheses were wrong
and are listed under [Dead ends](#dead-ends) so they are not re-run.

## Measurement discipline

**Do not optimize without a before and after.** Single-shot timings on a busy
desktop can vary by ±40%; interleave variants across rounds and report the minimum.
Change one thing at a time so an improvement is attributable.

**Export log rows are not an A/B instrument.** `Sanitize File` over the same 800
objects in a private benchmark corpus read 33.9 s, 5.9 s, 7.7 s, and 10.8 s across
four consecutive daily exports whose total runtime ranged from 257 s to 936 s.
Identical call counts do not make two runs comparable when machine load differs by
3×. Use the log to find which phase to look at, and a harness to attribute a change
to it.

**Confirm the instrument is measuring work, not blocking.** A phase timer that
includes a modal `MsgBox` from `Log.Error` will report seconds of “computation” while
Access waits for a click. Harness code sets `Operation.InteractionMode = eimSilent`
so warnings stay in the log.

**Prove byte identity, not just speed.** Exported output must not change unless an
export format version gate intentionally allows it. The perf harnesses in
`modTestPerf` record an output hash beside each timing.

**Check a corpus, not just a hash and the unit suite.** Committed source is a fixed
point of the sanitizer that produced it, so re-sanitizing a real project's files and
diffing is the strongest available equivalence check — and it is cheap. Both a green
unit suite and an unchanged single-file hash missed a report regression that
`BenchmarkSanitizeCorpus` catches on the first file. Cover each object type
separately; forms and reports do not share every rule. If the corpus predates a
format gate, compare against `git show HEAD:<path>` rather than the working tree.

## Techniques

### Hoist nested Dictionary access into typed locals

**Prefer:** `Set dProps = dBlock("Props")` once, then `dProps(strName)`.

**Why:** `Dictionary.Item` returns a `Variant`. A chained expression like
`dBlock("Props")(strName)` re-dispatches through the Variant on every access.

**Measured:** `ParseDocument` on a 5,800-line form dropped from **857 ms to 197 ms**
(4.3×) when early binding alone was bisected out of the canonicalizer probe.
Harness: `modTestPerf.BenchmarkFormGeometry`.

### Avoid TypeName and per-character escaping in JSON serialization

**Prefer:** dispatch known object nodes with `TypeOf value Is Dictionary` /
`TypeOf value Is Collection`. Before escaping a string, scan its UTF-16 bytes for
characters that actually require escaping and return the original string when none
are present.

**Why:** `TypeName()` on a live `Scripting.Dictionary` measured about 410 µs, so
nested metadata trees can spend most of their serialization time identifying nodes.
The original encoder also allocated a one-character string and appended it to a
buffer for every character, even though most keys and values contain no escapes.
The byte scan falls back to that established encoder on the first quote, backslash,
configured solidus, control character, or Unicode character that options require it
to escape.

**Measured (interleaved, minimum-of-rounds):** a 75 KB compact synthetic tree dropped
from **386 ms to 17 ms** and its 144 KB pretty form from **389 ms to 19 ms**. A
4,960-file, 6.23-million-character exported JSON corpus dropped from **22.97 s to
1.54 s** (**93.3%**) with **0 byte differences**. Harnesses:
`modTestPerf.BenchmarkJsonSerialization` and `BenchmarkJsonCorpus`.

In two user-run full exports, the serializer call count matched at 6,165 and dropped
from **116.09 s to 52.40 s** (**54.9% raw**). Unchanged operations indicate that the
second run was generally 11-17% faster. Normalizing by those controls puts the
serializer improvement at **46-49% (about 48%)**, an attributable saving of roughly
44-51 s, or 7.8-9.0% of the baseline export's total runtime. Use the interleaved
harness above for causal comparison; use these normalized figures for realistic
end-to-end expectations.

### Reuse query state already parsed during deterministic export

**Prefer:** derive companion `OptionFlag` bits from the `clsQueryComposer` that
reconstructed generated SQL, falling back to a SQL parse only for raw-SQL queries.
Likewise, reuse the Description property already parsed from `MSysObjects.LvProp`
instead of rereading it through DAO. Retain the DAO path when
`SaveAllDocumentProperties` requests a deep property scan.

**Why:** deterministic query export already has both sources in memory. Reparsing
every formatted SQL statement and reopening every DAO document repeats substantial
work inside the `Write JSON` phase.

**Measured:** across a 3,752-query production corpus, option-flag work dropped from
**15.71 s to 0.11 s** and metadata collection from **10.68 s to 0.29 s**. The
non-serialization portion of `Write JSON` fell from **27.22 s to 1.16 s** (**95.7%
raw, about 95.9% normalized**); matched controls were 3-5% slower in the final run.
Reference and optimized exports produced the same fingerprint across all 3,752 JSON
files. Applied to the prior user-run full export, this predicts roughly 56 s saved,
or 12.5% of that run; confirm that extrapolation with another full export.

### Use typed String() line buffers and cache bounds in hot loops

**Prefer:** assign `Split` to a typed `String()` array. Keep it local when one
procedure owns the operation and pass it `ByRef` to helpers. Use module-level storage
only when many tightly coupled helpers operate on the same buffer and measurement
justifies the implicit state. Cache `UBound` in genuinely hot loops.

**Why:** A `String()` array stored as a UDT member and read through `this` measured
materially slower than the same array at module scope. Re-evaluating `UBound` on every
loop pass adds up on 25,000-line forms. That measurement does not imply module state
is preferable to a local typed array.

**Measured:** Moving the canonicalizer line buffer out of the UDT took production
parse from **627 ms to 128 ms** on the same 5,800-line form.

### Gate locale-aware comparisons by length first

**Prefer:** nested `If Len(strTrimmed) = 3 Then If strTrimmed = "End"` rather than
`If strTrimmed = "End"` alone on every line.

**Why:** `Option Compare Database` makes every `=` string comparison locale-aware.
Comparing tens of thousands of long hex payload lines against `"End"` dominated parse
time on blob-heavy forms.

**Measured:** a 25,471-line form with 24,531 lines inside one `PictureData` block
went from **>60 s (never completed)** to **62 ms**. Harness: direct `Canonicalize`
timing via `modTestPerf` / `vcs_run_vba`.

### VBA does not short-circuit And

**Prefer:** nested `If` blocks or a precomputed flag instead of
`If cheap And expensive Then`.

**Why:** `If Len(strTLine) > 60 And StartsWith(strTLine, "0x")` still calls
`StartsWith` on every line. The same applies to `blnFlag And StartsWith(...)`.

**Applied in:** `clsFormGeometryCanonicalizer.ParseBlock` and
`clsSourceParser.IsBinaryPayloadLine`.

**Trap — do not nest inside an `ElseIf` chain.** Rewriting
`ElseIf blnFlag And StartsWith(...) Then` as `ElseIf blnFlag Then If StartsWith(...)`
changes behavior, not just speed: the outer `ElseIf` now claims the line whenever the
flag is set, so every later branch in the chain is skipped. In `SanitizeObject` that
made one report-only flag swallow `Checksum`, `WebImagePadding`, and the theme color
rules for every line between `Begin Report` and the report's `Bottom` — 327 report
files in one project quietly kept lines they should have dropped. Either compute the
condition into a `Boolean` **above** the chain, or handle an independent side effect
in a separate `If` and then continue into the chain:

```vba
If blnRemoveBottomRight Then
    If StartsWith(strLine, "    Right =") Then SkipLine lngLine
End If
' Continue into the ordinary ElseIf chain; no line has been claimed.
```

Covered by `clsTestSourceParser.TestSanitizeReport_*`, which places those rules inside
the flag's window.

### Prefer prefix comparison over InStr for prefix tests

**Prefer:** `Len(strText) >= Len(prefix)` then
`StrComp(Left$(strText, Len(prefix)), prefix, vbBinaryCompare) = 0`.

**Why:** `StartsWith` historically used `InStr(1, strText, prefix) = 1`, which scans
the entire line. The `Case Else` chain in `SanitizeObject` calls `StartsWith` many
times per property line.

**Scope:** the fast path applies only when `Compare = vbBinaryCompare` (the default).
Callers passing `vbTextCompare` keep the existing `InStr` path.

**Measured (simplified bundle, min-of-rounds):** a representative large form
**22.9 ms/call**, `modStringUtil.bas` **0.32 ms/call**, and `tblConflicts.xml`
**2.06 ms/call** at `EFV_5_0_0` (canonicalizer off). Harness:
`modTestPerf.BenchmarkSanitize`.

### Break Dictionary and Collection reference cycles

**Prefer:** remove back-reference keys (`Parent`, `Children`, …) before dropping the
root, or scope graphs to a single procedure without cycles.

**Why:** `Dictionary` and `Collection` are reference-counted. A parent holding
children that hold `Parent` is never reclaimed.

**Applied in:** `clsFormGeometryCanonicalizer` via `ReleaseBlocks` at the end of each
`Canonicalize` call.

### Represent line membership directly

**Prefer:** when line numbers are dense and bounded, use a `Boolean()` mask where the
array index is the line number and the value says whether to omit it.

**Why:** The previous skip representation appended line numbers, sorted them, walked
two indexes in parallel, and used an out-of-range sentinel to restore failed
conditional-format decodes. A Boolean mask expresses the actual question directly:
skip, keep, or restore a line with one assignment. It also removes a dedicated
`QuickSortLongs` implementation and duplicate-entry concerns.

**Measured:** this is a maintainability change, not a claimed speedup. The large-form
benchmark remained flat (**22.8 → 22.9 ms/call**) and all reference hashes remained
identical.

### Parse simple property lines once

**Prefer:** locate the single `" ="` delimiter and slice with `Left$` / `Mid$`
instead of allocating a Variant array with `Split`. Keep the parsing inside the
helper that owns the property format.

**Why:** This avoids an allocation while centralizing the format knowledge.
Module-level option and generated-key caches were removed after they showed no
corpus-level benefit; cheap setup stays local unless measurement justifies hidden
state.

**Applied in:** `clsSourceParser.CheckColorProperties`.

**Measured (corpus idempotence):** `BenchmarkFormCorpus` on a private corpus of 416
forms — **0 changed, 0 errors**, **1.33 s** total (**3.2 ms/form**). Full-sanitize
equivalence against the same corpus's freshly exported source, via
`BenchmarkSanitizeCorpus`: 319 reports and 416 forms both reported **0 differs,
0 errors**.

### Pass String() buffers ByRef between sanitize and canonicalize

**Prefer:** `clsSourceParser` holds a local `astrLines()` and passes it `ByRef` into
`clsFormGeometryCanonicalizer.Canonicalize(ByRef astrLines() As String)`, copying out
only when `ChangedLines > 0`.

**Why:** The old `Variant` parameter forced `CStr` on every line during copy-in.
Cross-class Variant array aliasing does not write back to the caller's buffer, so the
canonicalizer keeps a local `String()` working copy.

**Applied in:** `clsSourceParser.SanitizeObject` and `Canonicalize` signature change.

## Dead ends

These hypotheses were tested and **cleared by measurement** during the form geometry
work. Do not re-investigate without new evidence.

The original broad sanitize pass belongs here too, in part. Its per-line techniques
were plausible and the single-file harness numbers were real, but at corpus scale
they bought nothing measurable: two load-matched private-corpus exports (`Convert to
JSON` 43.15 s vs 43.14 s, same 800 objects) put `Sanitize File` at **5.86 s before
the pass and 6.34 s after**. The export's 800 objects average ~8 ms of sanitize each,
so fixed per-object overhead dominates and per-line savings only show up on the rare
large form. The implementation was simplified afterward, retaining only local,
low-complexity improvements. The real export-level win in that period —
`Sanitize File` **33.9 s → 5.9 s** — came from the form geometry canonicalizer work,
not from the broad sanitize pass. Optimize per-line cost when a single large object is
slow; do not expect it to move a full export.

| Hypothesis | Result |
|---|---|
| Variant-array-in-UDT read cost dominates | All 5,800 line reads ≈ 1.5 ms |
| Hidden quadratic re-scan in parse | Exactly one visit per line (5,794 visits / 254 blocks) |
| Block-graph bookkeeping alone | Mode 5 still 754 ms with full graph |
| O(n²) layout-group sorting on private corpus | Largest group (211 cells) sanitizes in ~250 ms |
| Modal dialog time is compute time | 13.5 s “CanonicalizeGroups” was a blocked warning dialog |

## How to re-measure

Run from the Immediate Window on the development copy of the add-in, or via
`vcs_run_vba` with `Operation.InteractionMode = eimSilent`:

| Harness | What it times |
|---|---|
| `modTestPerf.BenchmarkSanitize` | Full `Sanitize` at `EFV_5_0_0` (canonicalizer off) for a form, module, and table-def XML; reports hashes |
| `modTestPerf.BenchmarkFormGeometry` | Sanitize with canonicalizer on vs off, plus isolated canonicalizer phases |
| `modTestPerf.BenchmarkFormCorpus` | Canonicalizer only, every `.form` in a folder; **0 changed** means byte-identical to committed output |
| `modTestPerf.BenchmarkSanitizeCorpus` | Full `Sanitize` over any extension/object type in a folder; **0 differs** means byte-identical. Use for reports and macros, which the form harness does not cover |

For end-to-end export time, enable `ExportPerfJson` and read
[perf-diagnostics.md](perf-diagnostics.md). Compare runs only when call counts match.
