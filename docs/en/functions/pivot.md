# PIVOT

## `PIVOT.*`

The `PIVOT.*` family builds a statistical pivot from row and column categories, and each function computes exactly one metric.

### Syntax

```excel
=PIVOT.SUM(rows; columns; values)
=PIVOT.AVERAGE(rows; columns; values)
=PIVOT.PERCENTILE(rows; columns; values; quantile)
=PIVOT.CONF.T(rows; columns; values; alpha; [direction])
```

### Arguments

- `rows`: one or more category columns for rows, always including headers
- `columns`: optionally one or more category columns for columns, always including headers; may be blank
- `values`: one value column including a header
- `quantile`: requested percentile in `(0,1)`
- `alpha`: alpha level in `(0,1)`
- `direction`: optional direction for `CONF.*`; `0 = two-sided`, `-1 = left-sided`, `1 = right-sided`

### Available functions

| Function | Description | Detail |
|---|---|---|
| `PIVOT.COUNT(rows; columns; values)` | Count of non-empty values | no detail |
| `PIVOT.SUM(rows; columns; values)` | Sum | no detail |
| `PIVOT.AVERAGE(rows; columns; values)` | Arithmetic mean | no detail |
| `PIVOT.MIN(rows; columns; values)` | Minimum | no detail |
| `PIVOT.MAX(rows; columns; values)` | Maximum | no detail |
| `PIVOT.MEDIAN(rows; columns; values)` | Median | no detail |
| `PIVOT.PERCENTILE(rows; columns; values; quantile)` | Percentile | `0 < quantile < 1` |
| `PIVOT.STDEV.S(rows; columns; values)` | Sample standard deviation | no detail |
| `PIVOT.STDEV.P(rows; columns; values)` | Population standard deviation | no detail |
| `PIVOT.VAR.S(rows; columns; values)` | Sample variance | no detail |
| `PIVOT.VAR.P(rows; columns; values)` | Population variance | no detail |
| `PIVOT.VARCOEF.S(rows; columns; values)` | Sample coefficient of variation | no detail |
| `PIVOT.VARCOEF.P(rows; columns; values)` | Population coefficient of variation | no detail |
| `PIVOT.CONF.T(rows; columns; values; alpha; [direction])` | Half-width of the confidence interval based on the t distribution | `0 < alpha < 1`; `direction`: `0`, `-1`, `1` |
| `PIVOT.CONF.NORM(rows; columns; values; alpha; [direction])` | Half-width of the confidence interval based on the normal approximation | `0 < alpha < 1`; `direction`: `0`, `-1`, `1` |
| `PIVOT.MAD(rows; columns; values)` | Median absolute deviation from the median | no detail |
| `PIVOT.IQR(rows; columns; values)` | Interquartile range | no detail |

### Validation

- `rows`, `columns`, and `values` are expected to include a header in the first row
- `count` can work with non-numeric values and counts non-empty cells
- all other functions require quantitative data
- `varcoef.s` and `varcoef.p` require a non-zero mean
- `stdev.s`, `var.s`, `conf.t`, and `conf.norm` require at least two valid numeric observations
- incompatible data and function combinations return a value error
- `conf.*` only accept `direction` from `{-1, 0, 1}`

### Notes

- `columns` may be blank; in that case a one-dimensional pivot is created with a single `Total` column block
- row and column categories are sorted alphabetically in the output
- the output always includes a `TOTAL` summary row and a `TOTAL` summary column
- the bottom-right cell contains the overall aggregate for the selected metric across all data
- the output no longer includes a dedicated metric-label row
- rows for which every resulting value is blank are omitted from the output
- columns for which every resulting value is blank are omitted from the output, except for the `TOTAL` summary column
- `conf.t` and `conf.norm` return the half-width of the interval, not both bounds separately
- with `direction = 0`, the critical value `1 - alpha/2` is used
- with `direction = -1` or `1`, the critical value `1 - alpha` is used

### Output

Multi-row spill output in a pivot-like layout:

- row category columns on the left
- column category levels at the top
- the last header row above numeric columns now only contains row-variable names
- the last column is always the `TOTAL` summary column
- the last row is always the `TOTAL` summary row

### Examples

```excel
=PIVOT.SUM(E:E;F:F;G:G)
=PIVOT.AVERAGE(E:E;F:F;G:G)
=PIVOT.PERCENTILE(E:E;F:F;G:G;0.9)
=PIVOT.CONF.T(E:E;F:F;G:G;0.05)
=PIVOT.CONF.T(E:E;F:F;G:G;0.05;1)
```
