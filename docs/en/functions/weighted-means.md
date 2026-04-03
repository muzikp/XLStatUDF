# Weighted Means

## `AVERAGE.W`

Computes the weighted arithmetic mean.

### Syntax

```excel
=AVERAGE.W(values; weights)
```

### Arguments

- `values`: numeric observations
- `weights`: non-negative weights; blank weight cells are treated as `0`

### Notes

- `values` and `weights` must have the same length
- blank cells in `values` are skipped together with their matching weight
- if all weights are zero or any weight is negative, the function returns a numeric error

### Output

A scalar weighted arithmetic mean.

### Example

```excel
=AVERAGE.W(A2:A10;B2:B10)
```

## `HARMEAN.W`

Computes the weighted harmonic mean.

### Syntax

```excel
=HARMEAN.W(values; weights)
```

### Arguments

- `values`: positive numeric observations
- `weights`: non-negative weights; blank weight cells are treated as `0`

### Notes

- `values` and `weights` must have the same length
- values with zero weight do not affect the result
- if any value with positive weight is `<= 0`, the function returns a numeric error

### Output

A scalar weighted harmonic mean.

## `GEOMEAN.W`

Computes the weighted geometric mean.

### Syntax

```excel
=GEOMEAN.W(values; weights)
```

### Arguments

- `values`: positive numeric observations
- `weights`: non-negative weights; blank weight cells are treated as `0`

### Notes

- `values` and `weights` must have the same length
- values with zero weight do not affect the result
- if any value with positive weight is `<= 0`, the function returns a numeric error

### Output

A scalar weighted geometric mean.
