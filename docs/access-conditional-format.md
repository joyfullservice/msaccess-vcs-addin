# Access Conditional Format Binary Specification

This document describes the undocumented `ConditionalFormat` and `ConditionalFormat14`
properties as they appear in Microsoft Access `SaveAsText` / `LoadFromText` exports, and
how the VCS add-in decodes, stores, and rebuilds them.

**Status:** Reverse-engineered from fixture data. Suitable as a working reference for
parsing and rebuilding; some legacy-block fields remain unverified (see §14).

---

## 0. Implementation in this repo

The decoder/encoder lives in
[`clsConditionalFormat.cls`](../Version%20Control.accda.src/modules/Core/clsConditionalFormat.cls).
It is wired into the export/import pipeline by
[`clsSourceParser.cls`](../Version%20Control.accda.src/modules/Core/clsSourceParser.cls)
(capture/strip on export, `MergeConditionalFormat` on import) and
[`modLoadSaveText.bas`](../Version%20Control.accda.src/modules/Core/modLoadSaveText.bas)
(`WriteConditionalFormatting` into the companion JSON).

Behavior:

- On **export** (forms and reports), each control's `ConditionalFormat` /
  `ConditionalFormat14` blocks are stripped from the `.form` / `.report` source and the
  decoded rules are written to the companion `.json` under `Items.ConditionalFormatting`,
  keyed by control name.
- On **import**, both binary blocks are rebuilt from the JSON model and reinserted into
  each matching control, anchored immediately after the control's `Name` property line.
- **The JSON is authoritative.** If a control has a JSON entry, any inline
  `ConditionalFormat` / `ConditionalFormat14` block still present in the source for that
  control is stripped before the rebuilt block is injected. This prevents duplicate blocks
  when source is in a mixed state (e.g. a branch merge across the option boundary, or
  hand-edited files) and makes the merge **idempotent** — re-importing cannot accumulate
  blocks. A control that has an inline block but **no** JSON entry is left untouched (so
  option-off / un-migrated source round-trips unchanged, and will be decoded to JSON on the
  next export if the option is enabled).
- The feature is gated behind export format version **`EFV_5_0_0`** and the
  **`DecodeConditionalFormatting`** option (default **on**). The rules are always
  preserved; the option only chooses whether they are decoded to JSON (on) or left as
  the inline binary blocks (off). Import is always backward compatible: older source
  with inline blocks simply has nothing to merge.

**Round-trip fidelity boundary (important):**

- **CF14 is the authoritative block** and rebuilds **byte-for-byte** for every rule shape
  (expression, field-value/between, field-has-focus, data bar), including the per-rule
  trailer BackColor echo (§4.2). This is the complete copy and the source of truth for
  decoding.
- The **legacy `ConditionalFormat` block** rebuilds **byte-for-byte** for both single-rule
  and multi-rule expression/focus controls, and for single field-value rules of any
  operator (both bounds for Between/NotBetween, single bound for the rest). The multi-rule
  layout (28-byte non-last records with descriptor dwords, 12-byte last record,
  expression-window chaining, and the Access 2007 3-rule cap) is fully decoded from the
  AlertText fixture (§5.2). The legacy header operator (offset 16) is now decoded and
  rebuilt from the first rule; multi-rule field-value legacy blocks embed per-rule
  operators in their descriptors and remain best-effort.
- **Field-value operator** (`AcFormatConditionOperator`) is fully round-tripped from CF14
  for every one of the eight operators, single- and multi-rule (§4.3).
- Validated by [`modTestConditionalFormat.bas`](../Version%20Control.accda.src/modules/Tests/Core/modTestConditionalFormat.bas):
  byte-exact CF14 and legacy for all exercised shapes (single expression, single between,
  2-rule expression, 3-rule expression with 3-rule cap, trailer echo colors, all eight
  field-value operators, and a mixed-operator 3-rule field-value block).

### Companion JSON schema

```json
{
  "Info": { "Class": "clsSourceParser", "Description": "<object> Conditional Formatting" },
  "Items": {
    "ConditionalFormatting": {
      "Text9": {
        "Rules": [
          {
            "Type": "Expression",
            "Enabled": true,
            "FontBold": false,
            "FontItalic": false,
            "FontUnderline": false,
            "ForeColor": "RGB(0,0,0)",
            "BackColor": "RGB(255,255,255)",
            "Expression1": "[fraOption]=1"
          }
        ]
      }
    }
  }
}
```

Rule field sets by `Type`:

- `Expression` — font flags, colors, `Expression1`. May include `TrailerColor`.
- `FieldValue` — adds `Operator` and `Expression2`. `Operator` is one of `Between`,
  `NotBetween`, `Equal`, `NotEqual`, `GreaterThan`, `LessThan`, `GreaterThanOrEqual`,
  `LessThanOrEqual`. `Expression2` is only populated for `Between` / `NotBetween`; the
  single-value operators leave it empty.
- `FieldHasFocus` — font flags and colors only (no expression).
- `DataBar` — `FillColor`, `BarColor`, `ShowBarOnly`, `ShortestValue`, `ShortestLimit`,
  `LongestValue`, `LongestLimit` (limit values: `automatic`, `number`, `percent`).

`TrailerColor` is an optional `"RGB(R,G,B)"` string present when the CF14 trailing region
contains a non-zero BackColor echo (§4.2). Absent for white BackColors (echo is zero) and
for rules that exist only in CF14 (not also in the legacy block). Import produces zeros when
absent (safe; Access tolerates it).

Color fields (`ForeColor`, `BackColor`, `FillColor`, `BarColor`) use VBA-style
`RGB(R,G,B)` strings (e.g. `"RGB(255,0,0)"` for red). Access flattens automatic and
theme colors to literal RGB at save time, so exported values are always in `0..16777215`.
Import accepts both `RGB(...)` strings and legacy numeric Long values.

## 1. Overview

Conditional formatting on text boxes and combo boxes is stored as opaque binary blobs when
a form is exported to text. Access also exposes the same rules through the VBA
`FormatConditions` collection on each control.

| Export property | Typical Access versions | Notes |
|-----------------|-------------------------|-------|
| `ConditionalFormat` | 2007+ | Legacy layout; includes self-referential total byte length |
| `ConditionalFormat14` | 2010+ | Compact layout; preferred for analysis and editing |

Both properties are present on the same control in Access 365 exports. **Round-trip import
keeps both blocks consistent.**

### Relationship to VBA

| Binary concept | VBA `FormatCondition` property |
|----------------|-------------------------------|
| Condition type | `.Type` (`AcFormatConditionType`) |
| Comparison operator | `.Operator` (`AcFormatConditionOperator`) |
| First value / expression | `.Expression1` |
| Second value (between) | `.Expression2` |
| Bold / italic / underline | `.FontBold`, `.FontItalic`, `.FontUnderline` |
| Control enabled state | `.Enabled` |
| Text color | `.ForeColor` (Long; `-1` = automatic) |
| Background color | `.BackColor` (Long; `-1` = automatic) |

In the companion JSON, colors are stored as `"RGB(R,G,B)"` strings rather than raw Long
integers.

---

## 2. Text-file envelope

In `.form` / `.report` files, binary properties use this syntax:

```
ConditionalFormat14 = Begin
    0x01000100000001000000000000000101000000000000ffffff000d0000005b00 ,
    0x6600720061004f007000740069006f006e005d003d0031000000000000000000 ,
    0x00000000000000000000000000
End
```

### Parsing algorithm

1. Collect all lines between `Begin` and `End`.
2. Trim whitespace and trailing commas.
3. Remove the `0x` prefix from each line.
4. Split into pairs of hex digits and convert each pair to a byte.
5. Concatenate into a single `byte[]`.

Access emits **32 bytes (64 hex characters) per line**, all lines but the last terminated
with ` ,`. The rebuilder reproduces this format.

---

## 3. Enumerations

### AcFormatConditionType

| Name | Value | Use |
|------|-------|-----|
| `acFieldValue` | 0 | Compare control value using `.Operator` |
| `acExpression` | 1 | True/false expression in `.Expression1` |
| `acFieldHasFocus` | 2 | Applies when control has focus |
| `acDataBar` | 3 | Data bar; **CF14-only** (see §4.4) |

### AcFormatConditionOperator

| Name | Value |
|------|-------|
| `acBetween` | 0 |
| `acNotBetween` | 1 |
| `acEqual` | 2 |
| `acNotEqual` | 3 |
| `acGreaterThan` | 4 |
| `acLessThan` | 5 |
| `acGreaterThanOrEqual` | 6 |
| `acLessThanOrEqual` | 7 |

When `.Type` is `acExpression`, `.Operator` is ignored.

---

## 4. ConditionalFormat14 layout

### 4.1 File header (14 bytes)

The first rule record always begins at **offset 14**.

| Offset | Size | Field | Description |
|--------|------|-------|-------------|
| 0 | 2 | `version` | Always `0x0001` |
| 2 | 2 | `ruleCount` | Number of format rules |
| 4 | 2 | `reserved` | Always `0x0000` |
| 6 | 2 | `conditionType` | `AcFormatConditionType` of the **first** rule only |
| 8 | 2 | `reserved` | Always zero |
| 10 | 2 | `operator` | `AcFormatConditionOperator` of the **first** rule (field-value only; `0` otherwise) |
| 12 | 2 | `reserved` | Always zero |
| 14 | … | first rule record | See §4.2 |

`conditionType` and `operator` report the **first** rule's type and comparison operator
only. For rules beyond the first, the type and operator both come from that rule's 8-byte
prefix (§4.2), not from the header. `operator` is meaningful only when the rule type is
`acFieldValue`; Access ignores it (and writes `0`) for expression and focus rules.

> **Empirically verified** by exporting one text box eight times, once per
> `AcFormatConditionOperator`: only the two bytes at offset 10 changed, tracking the
> operator value `0`–`7`.

### 4.2 Rule record (expression / focus types)

| Offset (rel.) | Size | Field | Description |
|---------------|------|-------|-------------|
| 0 | 4 | `formatFlags` | Font and enabled flags (§6) |
| 4 | 4 | `foreColor` | BGR color (`0x00BBGGRR`) |
| 8 | 4 | `backColor` | BGR color (`0x00BBGGRR`) |
| 12 | 4 | `expr1Length` | Length of expression 1 in UTF-16 code units |
| 16 | `expr1Length × 2` | `expression1` | UTF-16LE text |
| … | variable | padding | Zero bytes (incl. null terminator) to the next rule / block end |

**Verified body-length rule:** a non-data-bar rule body is
`37 + 2 × (total UTF-16 code units across its expression fields)` bytes (excluding the
optional 8-byte prefix). This holds for expression (one length field), focus (zero-length),
and field-value (two length fields) rules.

**Trailing region layout:** after the expression data, each rule body has a fixed trailing
region (21 bytes for expression/focus, 17 bytes for field-value). Within this region, a
**3-byte BackColor echo** — the low 3 bytes of the rule's BackColor in BGR byte order —
sits at offset **`trailingLen - 12`**: **+9** for the 21-byte expression/focus trailer and
**+5** for the 17-byte field-value trailer, always followed by 9 trailing bytes. This echo
is non-zero only for rules that also appear in the legacy block (i.e., the first 3 rules of
the block's primary type). Rules that exist only in CF14 (e.g., the 4th expression rule, or
data bars) have zeros at this position. Access tolerates zeros here on import, so the echo
is a fidelity detail, not a correctness requirement. The companion JSON stores this as
`"TrailerColor": "RGB(R,G,B)"` when non-zero.

> **Empirically verified** for field-value rules by exporting a control with red, green,
> and blue field-value rules: each echo appeared 3 bytes at trailer offset +5 (the earlier
> +9-only assumption held for expression/focus rules but was never exercised for
> field-value because the only prior fixture used a white — zero-echo — BackColor).

**Multi-rule blocks:** rule 0 starts at offset 14 and its type/operator come from the
header (offsets 6 and 10). Each subsequent rule is preceded by an **8-byte prefix** (a
per-rule `conditionType` dword followed by an `operator` dword — the latter previously
assumed reserved), then the normal rule record. This is how one CF14 block carries rules of
different types and operators.

### 4.3 Field-value / between rules (`acFieldValue`)

`conditionType = 0` stores both bounds inline, each length-prefixed:

```
formatFlags(4) + foreColor(4) + backColor(4)
  + expr1Length(4) + expression1
  + expr2Length(4) + expression2
  + padding
```

The comparison operator is **not** stored in the rule record — it lives in the header
(offset 10) for the first rule, or in the 8-byte prefix's second dword for later rules
(§4.1, §4.2). All eight `AcFormatConditionOperator` values are supported. `Between` and
`NotBetween` populate both `expr1` and `expr2`; the single-value operators (`Equal`,
`NotEqual`, `GreaterThan`, `LessThan`, `GreaterThanOrEqual`, `LessThanOrEqual`) write
`expr2Length = 0` and no `expression2` bytes.

> **Historical note (issue #725):** earlier versions hardcoded the decoded operator to
> `Between`, so every field-value rule exported to JSON as `"Operator": "Between"`
> regardless of the real operator. Empirical capture of all eight operators located the
> operator bytes (header offset 10 / prefix second dword) and fixed both decode and
> rebuild.

### 4.4 Data bar rules (`acDataBar`, type 3)

Data bars exist **only in CF14** (the legacy block is single-type and omits them). Each data
bar rule begins with a `03 00 00 00` type dword; offsets below are relative to that marker.
The trailer starts at **`P = 32 + 2·(shortestLen + longestLen)`** and the record length is
**`P + 13`**.

| Offset (rel.) | Size | Field | Notes |
|---------------|------|-------|-------|
| 0 | 4 | `type` | `0x00000003` |
| 4 | 4 | `reserved` | `0` |
| 8 | 4 | `unk1` | always `1` |
| 12 | 4 | `reserved` | `0` |
| 16 | 4 | `fillColor` | BGR |
| 20 | 4 | `shortestLen` | UTF-16 code-unit count (`0` if unset) |
| 24 | `shortestLen × 2` | `shortestValue` | UTF-16LE typed value, e.g. `"10"`, `"-100"` |
| +A | 4 | `longestLen` | UTF-16 code-unit count |
| +A+4 | `longestLen × 2` | `longestValue` | UTF-16LE typed value |
| +B | 4 | `unk2` | always `1` |
| P | 1 | `showBarOnly` | `0` / `1` → VBA `ShowBarOnly` |
| P+1 | 3 | `barColor` | BGR (high byte `0` at P+4) → VBA `BackColor` |
| P+5 | 1 | `shortestLimit` | limit-type enum → VBA `ShortestBarLimit` |
| P+9 | 1 | `longestLimit` | limit-type enum → VBA `LongestBarLimit` |

> **Parsing note:** do **not** split records by searching for the `03 00 00 00` marker — a
> value length of 3 (e.g. `"-20"`) produces a false marker. Walk records using `P + 13`.

**Limit-type enum:** `0` = automatic, `1` = number, `2` = percent.

---

## 5. ConditionalFormat (legacy) layout

### 5.1 Block header (20 bytes)

| Offset | Size | Field | Description |
|--------|------|-------|-------------|
| 0 | 4 | `version` | Always `1` |
| 4 | 4 | **`blockSize`** | Total byte length of the entire block |
| 8 | 4 | `ruleCount` | Number of rules **of the block's single type** |
| 12 | 4 | `conditionType` | One type for the whole block |
| 16 | 4 | `operator` | `AcFormatConditionOperator` of the first rule (field-value only; `0` for expression/focus). Verified for all eight operators. |

**The legacy block is single-type.** When a control mixes rule types, the legacy block
stores only the rules matching its header type and omits the rest; CF14 holds the complete
set. (Text11: 2 of 3 rules; Text38: 2 of 10 rules.)

### 5.2 Rule body (fully decoded)

After the 20-byte header, offset 20 onward:

| Offset | Size | Field | Description |
|--------|------|-------|-------------|
| 20 | 4 | `reserved` | Always `0` |
| 24 | 4 | `offset24` | End unit index of the first expression window (= `exprLen[0] + 1` for expression/focus; `capacity` for between) |
| 28 | variable | rule records | Per-rule format records (see below) |
| … | variable | padding | Zero bytes so expression buffer begins at offset 96 |
| 96 | variable | expression buffer | Concatenated expression windows |

**3-rule cap.** The legacy block stores at most 3 rules (the Access 2007 limit). A control
with 4+ same-type rules stores only the first 3 in the legacy block; the rest survive in
CF14 only. Verified from AlertText (4 expression rules, legacy `ruleCount = 3`).

**Per-rule records.** The record size depends on position:

- **Non-last rules** use 28-byte records (7 dwords):

  | Offset (rel.) | Size | Field | Description |
  |---------------|------|-------|-------------|
  | 0 | 4 | `formatFlags` | Font and enabled flags (§6) |
  | 4 | 4 | `foreColor` | BGR color |
  | 8 | 4 | `backColor` | BGR color |
  | 12 | 4 | `constant` | Always `1` |
  | 16 | 4 | `constant` | Always `0` |
  | 20 | 4 | `nextExprStart` | Start unit index of the *next* rule's expression window |
  | 24 | 4 | `nextExprEnd` | End unit index (inclusive) of the *next* rule's expression window |

- The **last rule** uses a 12-byte record (3 dwords):

  | Offset (rel.) | Size | Field |
  |---------------|------|-------|
  | 0 | 4 | `formatFlags` |
  | 4 | 4 | `foreColor` |
  | 8 | 4 | `backColor` |

**Expression-window chaining.** Each expression/focus rule's expression occupies a tight
window of `exprLen + 2` UTF-16 code units (expression text + 2 null units). Windows are
chained sequentially from unit 0. The descriptor dwords on non-last rules point to the
*next* rule's window:

```
Window 0: units 0 .. (exprLen[0]+1)     ← offset24 = exprLen[0]+1
Window 1: units (W0 size) .. (W0+W1-1)  ← nextExprStart/End from rule 0
Window 2: units (W0+W1) .. (W0+W1+W2-1) ← nextExprStart/End from rule 1
```

**Verified from AlertText** (3 rules, each exprLen = 29):

| Rule | Record size | nextExprStart | nextExprEnd |
|------|-------------|---------------|-------------|
| 0 (non-last) | 28 | 31 | 61 |
| 1 (non-last) | 28 | 62 | 92 |
| 2 (last) | 12 | — | — |

Header + body = 28 + 28 + 28 + 12 = 96 (= `LEGACY_EXPR_OFFSET`, no padding needed).
Expression buffer = 3 × 31 × 2 = 186 bytes. `blockSize` = 96 + 186 = 282. ✓

**Between rules.** Field-value (between) rules use capacity-based padded slots instead of
tight windows: two slots of `capacity` code units each per rule (where `capacity` =
`max(exprLen) + 1`). `offset24` = `capacity`. Verified for single-rule (Text25 = 124 bytes).
Multi-rule between layout is unverified.

The rebuilder reproduces single-rule blocks byte-for-byte (Text9 = 126, Text23 = 100,
Text25 = 124 bytes) and multi-rule expression blocks byte-for-byte (Text11 = 156,
AlertText = 282 bytes).

---

## 6. Format flags dword (4 bytes)

| Byte index | Meaning | `1` = |
|------------|---------|-------|
| 0 | `Enabled` | enabled (control active) |
| 1 | `FontBold` | bold |
| 2 | `FontItalic` | italic |
| 3 | `FontUnderline` | underline |

Examples: `01 00 00 00` enabled/not bold; `01 01 00 00` enabled+bold; `00 01 00 00`
disabled+bold; `01 00 00 01` enabled+underline; `01 01 01 01` all flags.

---

## 7. Color encoding

Colors are stored as a **4-byte little-endian** value equal to the VBA RGB `Long`
(`0x00BBGGRR` byte order). The companion JSON stores them as `"RGB(R,G,B)"` strings.

| Appearance | Bytes | VBA Long |
|------------|-------|----------|
| Blue text | `00 00 FF 00` | 16711680 |
| White | `FF FF FF 00` | 16777215 |
| Dark Blue 5 | `17 36 5D 00` | 6108695 |
| Black | `00 00 00 00` | 0 |

Theme colors are flattened to literal RGB at save time; "Automatic" resolves to black text
on white background. There is no `-1` sentinel in the fixtures.

---

## 8. Expression encoding

- **Encoding:** UTF-16 little-endian.
- **Length field:** count of UTF-16 code units (not bytes).
- **Termination:** `00 00` after string content, then zero padding.

---

## 9. Verified fixtures

| Control | Source | Rules | Key properties |
|---------|--------|-------|----------------|
| Text9 | frmMain | 1 expression | Baseline single-rule expression (bold removed) |
| Text11 | frmMain | 2 expression + 1 focus | 2-rule multi-rule legacy; focus dropped |
| Text23 | frmMain | 1 focus | Single-rule focus |
| Text25 | frmMain | 1 between | Single-rule field-value |
| Text38 | frmMain | 2 expression + 8 data bar | Mixed-type; data bars CF14-only |
| **AlertText** | rAlertList | **4 expression** | Non-white BackColors (`RGB(219,219,183)`), 3-rule legacy cap, 28-byte non-last records, trailer BackColor echo |

## 10. Remaining unknowns

- CF14 trailing region: the BackColor echo is located (at `trailingLen - 12`: +9 for
  expression/focus, +5 for field-value); the purpose of the remaining trailing bytes is
  unknown. Reproduced verbatim (echo + zeros).
- Data bar `unk1`/`unk2` dwords (both always `1`) and the `fillColor` field — purpose
  unconfirmed; reproduced verbatim.
- Multi-rule legacy layout for field-value rules: the header operator (offset 16) is the
  first rule's; later field-value rules embed their operators in the per-rule descriptors.
  Rebuild remains best-effort for multi-rule field-value legacy blocks (CF14 is the
  authoritative decode source and round-trips fully).
- Whether the legacy block alone imports cleanly when no data bars are present.

---

## 11. References

- [FormatCondition object (Access)](https://learn.microsoft.com/en-us/office/vba/api/access.formatcondition)
- [AcFormatConditionType](https://learn.microsoft.com/en-us/office/vba/api/access.acformatconditiontype)
- [AcFormatConditionOperator](https://learn.microsoft.com/en-us/office/vba/api/access.acformatconditionoperator)
- [Application.SaveAsText](https://learn.microsoft.com/en-us/office/client-developer/access/desktop-database-reference/application-save-as-text)
- [Stack Overflow: SaveAsText hex / ConditionalFormat](https://stackoverflow.com/questions/63839201/access-application-saveastext-read-hex-values)
