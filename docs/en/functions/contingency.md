# Contingency Tables

## `CONTINGENCY.T`

Analyzes a contingency table entered directly as a frequency matrix.

### Syntax

```excel
=CONTINGENCY.T(table; [has_header]; [alpha])
```

### Arguments

- `table`: a 2D range of observed frequencies
- `has_header`: optional header mode code; default is `0`
- `alpha`: significance level

### `has_header` Codes

| Code | Meaning |
| --- | --- |
| `0` | autodetect labels |
| `1` | top row and left column are labels |
| `2` | the table is purely numeric with no labels |

### Notes

- the function requires at least a `2 × 2` table
- frequencies must be non-negative integers
- in autodetect mode, labels are used only if the top row and left column actually look like textual headers

### Output

The spill output contains:

- the observed contingency table including marginal totals
- the expected frequency table
- a test summary with `n`, `df`, `α`, `χ²`, critical value, and `p`
- association measures `Pearson C`, `Cramér V`, and `phi` for `2x2` tables

### Example

```excel
=CONTINGENCY.T(A1:C3;1;0,05)
```

## `CONTINGENCY.G`

Analyzes a contingency table built from grouped data.

### Syntax

```excel
=CONTINGENCY.G(columns; rows; [count]; [alpha]; [has_header])
```

### Arguments

- `columns`: categories of the future contingency table columns
- `rows`: categories of the future contingency table rows
- `count`: optional pair frequencies; if omitted, each pair has weight `1`
- `alpha`: significance level
- `has_header`: optional header mode code; default is `0`

### `has_header` Codes

| Code | Meaning |
| --- | --- |
| `0` | autodetect header |
| `1` | first row is a header |
| `2` | input has no header |

### Notes

- the function requires at least two distinct row categories and two distinct column categories
- blank rows are skipped
- if `count` is provided, it must contain non-negative integer frequencies
- the output structure is the same as `CONTINGENCY.T`

### Output

The spill output contains:

- the observed contingency table
- expected frequencies
- a `χ²` test summary
- association measures `Pearson C`, `Cramér V`, and possibly `phi`

### Examples

```excel
=CONTINGENCY.G(A2:A100;B2:B100)
=CONTINGENCY.G(A2:A100;B2:B100;C2:C100;0,05;1)
```
