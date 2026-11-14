# name-to-code-matcher

A small, dependency-free **VBA macro for Excel-compatible spreadsheets** that
maps a free-text item name (plus an optional attribute such as material) to a
standardized code, by looking the name up against one or more configurable
reference tables.

> **Unofficial.** Not affiliated with, endorsed by, or sponsored by Microsoft
> (Excel) or any other spreadsheet software vendor. The word "workbook" is used
> only as a generic reference to a spreadsheet file.

## Why this exists

Many everyday tasks boil down to the same problem: you have a list of messy,
free-text names and you want each one tagged with a canonical code from some
reference list. Examples that share the exact same shape:

- tagging items with a category code,
- mapping descriptions to account or class codes,
- normalizing user-entered labels to a controlled vocabulary.

Doing this by hand is slow and inconsistent. This macro automates it with a
robust matching engine and writes the result straight back into the sheet.

## What the engine does

The macro is fully generic. It implements four reusable ideas:

1. **Word-boundary-aware substring matching.** A reference name only matches a
   source name when it is surrounded by non-word characters (spaces, punctuation)
   or string edges. So `Widget` matches `Blue Widget Stand` but not `Widgetron`.
   This works for Latin text and CJK ideographs alike.
2. **Attribute-category filtering.** If both the source row and a reference row
   carry an attribute (e.g. material), a candidate is only accepted when the two
   attributes fall in the same coarse category (both "Metal", both "Plastic", …).
3. **Multi-tier fallback.** Reference tables are tried in order; the first table
   that yields a code wins. This lets you layer a precise table over broader ones.
4. **Weighted best-candidate resolver.** When several codes match a name, the
   macro picks the best one using a score that combines match frequency (0.3)
   and row recency — newer rows rank higher (0.7).

All of the above is driven by configuration constants at the top of the module;
nothing about a specific domain is hard-coded.

## How it works

1. The macro reads the source sheet (the active sheet, or a named one).
2. For each data row it takes the item name and optional attribute.
3. It walks the configured tiers in order. Within a tier it scans the reference
   table from the bottom up, collecting exact and boundary-aware substring hits,
   filters them by attribute category, and resolves ties with the weighted score.
4. The first tier that returns a code wins; the code is written to the output
   column, an optional source label is recorded, and the cell is tinted by tier.
5. Rows with no match get a `N/A` marker; rows that error are flagged in red.

### Illustrative example (fictional)

With the bundled sample the source row `Blue Widget Stand` (attribute `Metal`)
matches `Widget Stand` in `RefPrimary` exactly and receives code `W-100`. A name
like `Industrial Widget` would not falsely match a reference `Widgetron`, because
the boundary check requires `Widget` to sit between non-word characters.

## Input data format

The source sheet layout is configurable. Defaults used by the sample:

| Column (default) | Meaning | Required |
|------------------|---------|----------|
| C (NAME_COL)     | Item name (free text) | Yes |
| Y (ATTR_COL)     | Attribute, e.g. material (optional) | No |
| E (CODE_COL)     | Output: the matched code | Yes |
| F (SOURCE_COL)   | Output: which tier matched (optional) | No |
| C3 (header)      | Expected header text `Item name` | No (configurable) |

- `HEADER_ROW` = `3`, `DATA_START_ROW` = `4`.
- Reference tables live in the **same workbook that holds the macro**
  (`ThisWorkbook`), each with a header row and data starting at row 2.
- A reference row's name may use `-` to drop a suffix and `/` to list aliases;
  each alias is matched independently.

## Configuration

All knobs are constants at the top of `src/NameToCode.bas`:

| Constant | Default | Meaning |
|----------|---------|---------|
| `SRC_SHEET` | `""` | `""` = active sheet, else the sheet name |
| `NAME_COL` | `"C"` | Item name column |
| `ATTR_COL` | `"Y"` | Attribute column (`""` to disable) |
| `CODE_COL` | `"E"` | Output code column |
| `SOURCE_COL` | `"F"` | Match-source column (`""` to disable) |
| `HEADER_ROW` | `3` | Row holding the name header |
| `DATA_START_ROW` | `4` | First data row |
| `NAME_HEADER_TEXT` | `"Item name"` | Header to expect (`""` to skip) |
| `TIERS` | `RefPrimary\|A\|B\|C, RefSecondary\|A\|C\|H, RefTertiary\|A\|B\|D` | Ordered reference tables |
| `CATEGORY_SHEET` | `"Categories"` | Keyword→category sheet (optional) |
| `ATTR_FILTER_ON` | `True` | Filter matches by attribute category |
| `W_COUNT` | `0.3` | Weight of match frequency |
| `W_RECENCY` | `0.7` | Weight of row recency |
| `COLOR_ON` | `True` | Tint cells by tier |

### Tier format

Each tier is `SheetName|NameCol|CodeCol|AttrCol`, separated by commas:

```
RefPrimary|A|B|C, RefSecondary|A|C|H, RefTertiary|A|B|D
```

- `SheetName` — the reference table sheet in the macro workbook.
- `NameCol` — column holding the reference names.
- `CodeCol` — column holding the codes to output.
- `AttrCol` — column holding the reference attribute (`""` if none).

### Attribute taxonomy

If a sheet named by `CATEGORY_SHEET` exists, it is read as `keyword, category`
rows and used for attribute filtering. If it is missing, a built-in sample
taxonomy of physical materials (plastic, metal, wood, paper, textile, glass,
mineral, composite) is used. Replace either with your own categories.

## Installation

1. Open your workbook in Excel or any compatible spreadsheet program.
2. Press `Alt + F11` to open the VBA editor, then `File ▸ Import File…` and
   select `src/NameToCode.bas`.
3. Place your reference tables in the **same workbook** that now holds the macro.
4. Edit the configuration constants at the top of the module to match your layout.
5. Save the workbook as a macro-enabled file (`.xlsm`).

## Usage

1. Prepare the source sheet in the [input format](#input-data-format) above.
2. Select the sheet you want to process (or leave `SRC_SHEET = ""` for the active
   sheet).
3. Run the `MatchCodes` macro.
4. Review the codes in the output column and the optional source column; adjust
   manually if needed.

A ready-to-run sample lives in
[`examples/sample-data.xlsx`](examples/sample-data.xlsx).

## File structure

```
name-to-code-matcher/
├── src/
│   └── NameToCode.bas         # the macro (clean-room, generic reimplementation)
├── docs/
│   └── Configuration.md       # configuration walkthrough with a second domain
├── examples/
│   └── sample-data.xlsx       # neutral, fictional sample workbook
├── LICENSE                    # MIT
├── README.md                  # this file
├── PUSH_GUIDE.md              # how this repo was published to GitHub
├── .gitignore
└── .gitattributes
```

## Performance notes

Matching scans each reference table once per source row, bottom-up, so cost grows
with (source rows) × (reference rows across tiers). For typical lookup tables
(thousands of rows) this is fast. For very large tables, consider trimming the
reference data or reducing the number of tiers.

## Limitations

- Reference tables must reside in the same workbook as the macro (`ThisWorkbook`).
- The engine matches on whole-name equality or boundary-aware substring; it does
  not do fuzzy/phonetic matching.
- The scoring objective (frequency + recency) is fixed; a different objective
  would require editing `BestCode`.
- This macro was not executed in the author's environment; test it on a copy of
  your data before relying on it.

## Non-association & Legal

This repository is an **independent, clean-room reimplementation** of a generic
matching pattern. It contains no proprietary data, no internal nomenclature, and
no business-specific logic from any third party or former employer. The author
is not affiliated with, and does not speak for, any company whose data may have
originally inspired a similar workflow.

The MIT license covers only the author's own contributions; it does not purport
to license any rights of others. Users are responsible for ensuring their use of
this tool complies with applicable laws and their own organization's policies.

## License

Released under the [MIT License](LICENSE). Copyright (c) 2026 ivy92888-glitch.
