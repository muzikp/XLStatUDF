# Coefficients Of Variation

## `VARCOEF`

Computes the population coefficient of variation.

### Syntax

```excel
=VARCOEF(values)
```

### Notes

- blank cells are ignored
- at least one valid value is required
- if the mean equals zero, the function returns a divide-by-zero error

### Output

A scalar `σ / μ`.

## `VARCOEF.S`

Computes the sample coefficient of variation.

### Syntax

```excel
=VARCOEF.S(values)
```

### Notes

- blank cells are ignored
- at least two valid values are required
- if the mean equals zero, the function returns a divide-by-zero error

### Output

A scalar `sₓ / x̄`.

## `VARCOEF.W`

Computes the weighted population coefficient of variation.

### Syntax

```excel
=VARCOEF.W(values; weights)
```

### Notes

- `values` and `weights` must have the same length
- weights must be non-negative
- blank weight cells are treated as `0`
- if the weighted mean equals zero, the function returns a divide-by-zero error

### Output

A scalar `σ_w / x̄_w`.

## `VARCOEF.S.W`

Computes the weighted sample coefficient of variation.

### Syntax

```excel
=VARCOEF.S.W(values; weights)
```

### Notes

- `values` and `weights` must have the same length
- weights must be non-negative
- blank weight cells are treated as `0`
- if `Σw <= 1`, the function returns `#COUNT!`
- if the weighted mean equals zero, the function returns a divide-by-zero error

### Output

A scalar `sₓ,w / x̄_w`.
