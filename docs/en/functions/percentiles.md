# Percentiles

## `PERCENTILE.INC.IFS`

Computes an inclusive percentile with `SUMIFS`-style filtering.

### Syntax

```excel
=PERCENTILE.INC.IFS(values; quantile; [criteria_range_1; criteria_1; ...])
```

### Arguments

- `values`: numeric data
- `quantile`: a value from `0` to `1`
- `criteria_range_n`: range used for filtering
- `criteria_n`: exact match, relational expression, or wildcard expression

### Notes

- filter arguments must come in even `range + criteria` pairs
- if filtering leaves no values, the function returns `#N/A`
- the percentile is computed using the inclusive method, matching `PERCENTILE.INC`

### Output

A scalar percentile.

### Example

```excel
=PERCENTILE.INC.IFS(A2:A100;0,75;B2:B100;"A";A2:A100;">10")
```

## `PERCENTILE.EXC.IFS`

Computes an exclusive percentile with `SUMIFS`-style filtering.

### Syntax

```excel
=PERCENTILE.EXC.IFS(values; quantile; [criteria_range_1; criteria_1; ...])
```

### Arguments

Same as `PERCENTILE.INC.IFS`, but `quantile` must be strictly between `0` and `1`.

### Notes

- filter arguments must come in even `range + criteria` pairs
- if filtering leaves no values, the function returns `#N/A`
- if the exclusive percentile formula falls outside its valid domain for the given quantile, the function returns a numeric error

### Output

A scalar percentile.
