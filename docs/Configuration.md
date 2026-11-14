# Configuration walkthrough

Everything in `name-to-code-matcher` is controlled by constants at the top of
`src/NameToCode.bas`. This page shows the defaults and then walks through a
second, completely different domain to prove the tool is domain-agnostic.

## Defaults (as shipped)

```vb
SRC_SHEET        = ""          ' active sheet
NAME_COL         = "C"         ' item name
ATTR_COL         = "Y"         ' attribute (material)
CODE_COL         = "E"         ' output code
SOURCE_COL       = "F"         ' match source (optional)
HEADER_ROW       = 3
DATA_START_ROW   = 4
NAME_HEADER_TEXT = "Item name"

TIERS = "RefPrimary|A|B|C,RefSecondary|A|C|H,RefTertiary|A|B|D"

CATEGORY_SHEET = "Categories"
ATTR_FILTER_ON = True

W_COUNT   = 0.3
W_RECENCY = 0.7

COLOR_ON = True
TIER_COLORS = "198,224,180|255,242,204|255,183,183"
```

## Worked example: expense description → account code

Suppose you receive bank-export rows and want each `Description` tagged with a
general-ledger account code. You have three reference lists of increasing
specificity.

### 1. Source sheet (`Expenses`)

| A (row) | C = Description | Y = Department | E = Account (output) |
|---------|-----------------|----------------|----------------------|
| 4 | Office paper purchase | Admin | (empty) |
| 5 | Cloud hosting fee | IT | (empty) |
| 6 | Client dinner | Sales | (empty) |

### 2. Reference tiers (in the macro workbook)

- `AcctExact` — exact description → account (Name col A, Code col B, no attribute).
- `AcctKeyword` — keyword → account (Name col A, Code col B, no attribute).
- `AcctDept` — department → account (Name col A, Code col B, attribute col C).

### 3. Edit the constants

```vb
NAME_COL   = "C"
ATTR_COL   = "Y"
CODE_COL   = "E"
SOURCE_COL = "F"
HEADER_ROW = 3
DATA_START_ROW = 4
NAME_HEADER_TEXT = "Description"

TIERS = "AcctExact|A|B|,AcctKeyword|A|B|,AcctDept|A|B|C"

CATEGORY_SHEET = "DeptGroups"   ' optional: keyword -> department group
ATTR_FILTER_ON = True
```

Notes:

- An empty `AttrCol` is written as a trailing empty field: `"AcctExact|A|B|"`. The
  parser treats a missing/empty attribute column as "no attribute filtering for
  that tier".
- Because `AcctExact` is listed first, an exact description match always wins over
  a looser keyword match in `AcctKeyword`, which in turn wins over the
  department-level fallback in `AcctDept`. That is the multi-tier fallback in
  action.
- `W_COUNT` and `W_RECENCY` still decide ties when several rows in one tier match
  the same description.

### 4. Result

| C = Description | Y = Department | E = Account |
|-----------------|----------------|-------------|
| Office paper purchase | Admin | (from AcctKeyword, e.g. `ACC-5400`) |
| Cloud hosting fee | IT | (from AcctExact, e.g. `ACC-6100`) |
| Client dinner | Sales | (from AcctDept, e.g. `ACC-7200`) |

Same engine, different vocabulary — no code changes beyond the constants.

## Tuning the resolver

- Raise `W_RECENCY` (toward `1.0`) to prefer codes defined on newer reference
  rows; lower it to `0.0` to rank purely by how often a code appears.
- Set `COLOR_ON = False` to skip cell tinting when you only want the codes.
- Set `ATTR_FILTER_ON = False` to ignore attributes entirely and match on names
  alone.
