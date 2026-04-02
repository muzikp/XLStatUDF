# Distributions

## `NORM.DIST.RANGE`

Returns the probability that a normally distributed random variable falls within a given interval.

### Syntax

```excel
=NORM.DIST.RANGE(mean; standard_deviation; lower_bound; upper_bound)
```

## `GENERATE.NORM`

Generates a spill column of random numbers from a normal distribution.

### Syntax

```excel
=GENERATE.NORM(mean; stdev; count)
```

## `FILL`

Creates a single-column spill by repeating one value or by repeatedly evaluating a formula.

### Syntax

```excel
=FILL(what; count)
```

### Arguments

- `what`: a scalar value to repeat, or a text formula starting with `=`
- `count`: number of returned rows; integer `>= 1`

### Output

A single-column spill range with `count` values.

### Notes

- if you pass a regular value, the function simply repeats it in every row
- if you pass a text formula starting with `=`, the formula is evaluated separately for each row
- this is useful for random generators; a direct argument like `RANDBETWEEN(...)` would otherwise reach the UDF as an already computed single value

### Examples

```excel
=FILL("A";5)
=FILL(123;4)
=FILL("=RANDBETWEEN(1;10)";20)
```
