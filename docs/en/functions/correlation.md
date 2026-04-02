# Correlation

## `CORREL.SPEARMAN`

Spearman rank correlation with significance test.

### Syntax

```excel
=CORREL.SPEARMAN(x_values; y_values; [direction]; [alpha]; [ma_zahlavi])
```

### Arguments

- `x_values`: first numeric range
- `y_values`: second numeric range
- `direction`: optional direction code; default is `0`
- `alpha`: significance level
- `ma_zahlavi`: optional header-mode code; default is `0`

### `direction` Codes

| Code | Meaning |
| --- | --- |
| `0` | two-sided test |
| `1` | left-sided test |
| `2` | right-sided test |

### `ma_zahlavi` Codes

| Code | Meaning |
| --- | --- |
| `0` | auto-detect header |
| `1` | first row is a header |
| `2` | input has no header |

Blank cells are excluded pairwise.

### Output

The spill output contains:

- `ρ`
- `n`
- `α`
- `t`
- `df`
- critical `t` value
- `p`
