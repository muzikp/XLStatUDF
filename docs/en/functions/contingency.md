# Contingency Tables

## `CONTINGENCY.T`

Contingency-table analysis from a directly supplied matrix of counts.

### Syntax

```excel
=CONTINGENCY.T(table; [ma_zahlavi]; [alpha])
```

### Arguments

- `table`: 2D range with observed counts
- `ma_zahlavi`: optional header-mode code; default is `0`
- `alpha`: significance level

### `ma_zahlavi` Codes

| Code | Meaning |
| --- | --- |
| `0` | auto-detect labels |
| `1` | top row and left column are labels |
| `2` | table is purely numeric without labels |

### Output

The spill output contains:

- the observed contingency table including margins
- the expected-frequency table
- a test summary with `n`, `df`, `α`, `χ²`, the critical value, and `p`
- association measures `Pearson C`, `Cramer V`, and `phi` for `2x2` tables

### Example

```excel
=CONTINGENCY.T(A1:C3;1;0,05)
```

## `CONTINGENCY.G`

Contingency-table analysis from grouped columns.

### Syntax

```excel
=CONTINGENCY.G(columns; rows; [count]; [alpha]; [ma_zahlavi])
```

### Arguments

- `columns`: categories of future contingency-table columns
- `rows`: categories of future contingency-table rows
- `count`: optional frequencies for each pair; when omitted, each pair has weight `1`
- `alpha`: significance level
- `ma_zahlavi`: optional header-mode code; default is `0`

### `ma_zahlavi` Codes

| Code | Meaning |
| --- | --- |
| `0` | auto-detect header |
| `1` | first row is a header |
| `2` | input has no header |

### Output

The spill output contains the same sections as `CONTINGENCY.T`:

- the observed contingency table
- expected frequencies
- the `χ²` test summary
- association measures `Pearson C`, `Cramer V`, and optionally `phi`

### Examples

```excel
=CONTINGENCY.G(A2:A100;B2:B100)
=CONTINGENCY.G(A2:A100;B2:B100;C2:C100;0,05;1)
```
