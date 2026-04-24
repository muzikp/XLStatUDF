# Correlation

## `CORREL.SPEARMAN`

Computes Spearman's rank correlation coefficient and performs its significance test.

### Syntax

```excel
=CORREL.SPEARMAN(x_values; y_values; [direction]; [alpha]; [has_header])
```

### Arguments

- `x_values`: first numeric range
- `y_values`: second numeric range
- `direction`: optional direction code; default is `0`
- `alpha`: significance level
- `has_header`: optional header mode code; default is `0`

### `direction` Codes

| Code | Meaning |
| --- | --- |
| `0` | two-sided test |
| `1` | left-sided test |
| `2` | right-sided test |

### `has_header` Codes

| Code | Meaning |
| --- | --- |
| `0` | autodetect header |
| `1` | first row is a header |
| `2` | input has no header |

### Notes

- blank cells are excluded pairwise
- at least three valid pairs are required
- if either variable has zero rank variance, the function returns a numeric error

### Output

The spill output contains:

- `ρ`
- `n`
- `α`
- `t`
- `df`
- critical `t`
- `p`

## `CORREL.MATRIX`

Builds a correlation matrix for data with two or more columns.

### Syntax

```excel
=CORREL.MATRIX(data; [method]; [output]; [p_minimum]; [has_header])
```

### Arguments

- `data`: input matrix; columns are variables
- `method`: optional computation method code; default is `0`
- `output`: optional output type code; default is `0`
- `p_minimum`: optional filter; only links with `p < p_minimum` are returned
- `has_header`: optional header mode code; default is `0`

### `method` Codes

| Code | Meaning |
| --- | --- |
| `0` | Pearson |
| `1` | Spearman |

### `output` Codes

| Code | Meaning |
| --- | --- |
| `0` | coefficients only |
| `1` | two-tailed p-values only |
| `2` | coefficient and p-value on the row below |
| `3` | coefficient, p-value below it, and significance stars on the third row |
| `4` | coefficients only, with significance stars appended in the same cell |

### `has_header` Codes

| Code | Meaning |
| --- | --- |
| `0` | autodetect header |
| `1` | first row is a header |
| `2` | input has no header |

### Notes

- the function requires at least two columns and at least three complete rows
- incomplete rows are skipped as complete-case rows
- if any column is constant, the function returns a numeric error
- if `p_minimum` is supplied, non-matching links are left blank in the output, including diagonal cells
- the add-in still accepts the legacy argument order `([p_minimum]; data; [method]; [output]; [has_header])` for backward compatibility

### Output

For outputs `0`, `1`, and `4`, the spill output returns a square matrix with variable names.

For outputs `2` and `3`, the spill output returns a stacked layout with two leading label columns and one row block per variable:

- coefficient
- `p`
- optionally `sig.`

For output `4`, each cell contains text in a form such as `0.30156***`.

### Examples

```excel
=CORREL.MATRIX(B1:E30)
=CORREL.MATRIX(B1:E30;1;0;;1)
=CORREL.MATRIX(B1:E30;0;4;0,05;1)
```
