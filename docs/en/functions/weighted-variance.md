# Weighted Variance And Standard Deviation

Common behavior for all four functions:

- `values` and `weights` must have the same length
- weights must be non-negative
- blank weight cells are treated as `0`
- if the sum of weights is zero, the function returns a numeric error

## `VAR.P.W`

Computes the weighted population variance.

### Syntax

```excel
=VAR.P.W(values; weights)
```

### Output

A scalar weighted population variance.

### Notes

- uses the denominator `Σw`

## `VAR.S.W`

Computes the weighted sample variance.

### Syntax

```excel
=VAR.S.W(values; weights)
```

### Output

A scalar weighted sample variance.

### Notes

- uses the denominator `Σw - 1`
- if `Σw <= 1`, the function returns `#COUNT!`

## `STDEV.P.W`

Computes the weighted population standard deviation.

### Syntax

```excel
=STDEV.P.W(values; weights)
```

### Output

A scalar weighted population standard deviation `σ`.

### Notes

- defined as the square root of `VAR.P.W`

## `STDEV.S.W`

Computes the weighted sample standard deviation.

### Syntax

```excel
=STDEV.S.W(values; weights)
```

### Output

A scalar weighted sample standard deviation `sₓ`.

### Notes

- defined as the square root of `VAR.S.W`
