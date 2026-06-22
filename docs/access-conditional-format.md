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
  (expression, field-value/between, field-has-focus, data bar). This is the complete copy
  and the source of truth for decoding.
- The **legacy `ConditionalFormat` block** rebuilds **byte-for-byte for single-rule
  controls** (the common case). Its **multi-rule** per-rule layout is not fully documented
  (§14), so multi-rule legacy blocks are rebuilt best-effort (correct header, flags,
  colors, and expressions) but may not be byte-identical to Access's original.
- Validated by [`modTestConditionalFormat.bas`](../Version%20Control.accda.src/modules/Tests/Core/modTestConditionalFormat.bas):
  byte-exact CF14 for all shapes, byte-exact legacy for single-rule shapes, and semantic
  round-trip stability for the rest.

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
            "ForeColor": 0,
            "BackColor": 16777215,
            "Expression1": "[fraOption]=1"
          }
        ]
      }
    }
  }
}
```

Rule field sets by `Type`:

- `Expression` — font flags, colors, `Expression1`.
- `FieldValue` — adds `Operator` ("Between") and `Expression2`.
- `FieldHasFocus` — font flags and colors only (no expression).
- `DataBar` — `FillColor`, `BarColor`, `ShowBarOnly`, `ShortestValue`, `ShortestLimit`,
  `LongestValue`, `LongestLimit` (limit values: `automatic`, `number`, `percent`).

---

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
| 8 | 6 | `reserved` | Always zero |
| 14 | … | first rule record | See §4.2 |

`conditionType` reports the **first** rule's type only. Rule types beyond the first are
read from each rule record's 8-byte prefix, not from the header.

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
and field-value (two length fields) rules. The rebuilder uses this to emit exact trailing
padding.

**Multi-rule blocks:** rule 0 starts at offset 14 and its type comes from the header. Each
subsequent rule is preceded by an **8-byte prefix** (a per-rule `conditionType` dword + a
reserved dword), then the normal rule record. This is how one CF14 block carries rules of
different types.

### 4.3 Field-value / between rules (`acFieldValue`)

`conditionType = 0` stores both bounds inline, each length-prefixed:

```
formatFlags(4) + foreColor(4) + backColor(4)
  + expr1Length(4) + expression1
  + expr2Length(4) + expression2
  + padding
```

Operator `acBetween` is implied (the only field-value shape exercised); other operators are
single-value and currently treated as a single expression slot.

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
| 16 | 4 | `operator` | `AcFormatConditionOperator` (0 for expression/focus) |

**The legacy block is single-type.** When a control mixes rule types, the legacy block
stores only the rules matching its header type and omits the rest; CF14 holds the complete
set. (Text11: 2 of 3 rules; Text38: 2 of 10 rules.)

### 5.2 Rule body (as rebuilt by this add-in)

After the 20-byte header:

- `reserved` dword (`0`), then `exprCapacity` dword — the expression slot capacity in
  UTF-16 code units (longest expression + 1 for the null).
- Per-rule `formatFlags` + `foreColor` + `backColor` dwords.
- A reserved region of zeros so the concatenated expression buffer begins at **offset 96**.
- Expression buffer. **Verified slot sizing:** expression / focus rules reserve one slot of
  `(capacity + 1)` code units; field-value (between) rules reserve two slots of `capacity`
  code units each.

The rebuilder reproduces single-rule blocks byte-for-byte (Text9 = 126, Text23 = 100,
Text25 = 124 bytes). Multi-rule legacy blocks insert additional per-rule descriptor bytes
that are not fully decoded (§14); the rebuilder approximates them.

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
(`0x00BBGGRR` byte order). The add-in stores them as the VBA `Long` integer in JSON.

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

## 9. Remaining unknowns

- Exact trailing padding inside each CF14 rule record is derived from the verified
  `37 + 2·units` formula; the underlying meaning of the constant is not documented.
- Full legacy per-rule header bytes between `operator` and the color dwords (multi-rule
  legacy layout) — the reason multi-rule legacy rebuild is best-effort, not byte-exact.
- Data bar `unk1`/`unk2` dwords (both always `1`) and the `fillColor` field — purpose
  unconfirmed; reproduced verbatim.
- Whether the legacy block alone imports cleanly when no data bars are present.

---

## 10. References

- [FormatCondition object (Access)](https://learn.microsoft.com/en-us/office/vba/api/access.formatcondition)
- [AcFormatConditionType](https://learn.microsoft.com/en-us/office/vba/api/access.acformatconditiontype)
- [AcFormatConditionOperator](https://learn.microsoft.com/en-us/office/vba/api/access.acformatconditionoperator)
- [Application.SaveAsText](https://learn.microsoft.com/en-us/office/client-developer/access/desktop-database-reference/application-save-as-text)
- [Stack Overflow: SaveAsText hex / ConditionalFormat](https://stackoverflow.com/questions/63839201/access-application-saveastext-read-hex-values)
