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

## `GENERATE.INT`

Generates a spill column of random integers from a given interval.

### Syntax

```excel
=GENERATE.INT([count]; [minimum]; [maximum])
```

### Arguments

- `count`: optional number of generated values; integer `>= 1`; default is `1`
- `minimum`: optional lower bound; default is `-2147483648`
- `maximum`: optional upper bound; default is `2147483647`

### Output

A single-column spill range with random integers from the closed interval `<minimum; maximum>`.

### Notes

- the function is volatile, so recalculation generates a new sequence
- if `minimum > maximum`, the function returns a numeric error
- if bounds are omitted, the full practical 32-bit range is used

### Examples

```excel
=GENERATE.INT()
=GENERATE.INT(10)
=GENERATE.INT(20;1;6)
```

## `FILL`

Creates a single-column spill by repeating one value or by repeatedly evaluating a formula.

### Syntax

```excel
=FILL(what; count; [what2]; [count2]; ...)
```

### Arguments

- `what`: a scalar value to repeat, or a text formula starting with `=`
- `count`: number of returned rows; integer `>= 1`
- any additional arguments must be supplied as `what + count` pairs

### Output

A single-column spill range with `count` values.

### Notes

- if you pass a regular value, the function simply repeats it in every row
- if you pass a text formula starting with `=`, the formula is evaluated separately for each row
- you can pass multiple `what + count` pairs; the function appends the generated blocks under each other in the given order
- this is useful for random generators; a direct argument like `RANDBETWEEN(...)` would otherwise reach the UDF as an already computed single value

### Examples

```excel
=FILL("A";5)
=FILL(123;4)
=FILL("male";100;"female";100;"child";90)
=FILL("=RANDBETWEEN(1;10)";20)
```

## `FILL.RANDOM`

Creates a single-column spill like `FILL`, but randomly shuffles the generated sequence before returning it.

### Syntax

```excel
=FILL.RANDOM(what; count; [what2]; [count2]; ...)
```

### Arguments

- `what`: a scalar value to repeat, or a text formula starting with `=`
- `count`: number of returned rows; integer `>= 1`
- any additional arguments must be supplied as `what + count` pairs

### Output

A single-column spill range containing all generated values in random order.

### Notes

- the function first builds the full sequence just like `FILL`
- it then shuffles the sequence randomly
- this is useful for randomized category generation or shuffled stimulus orders

### Examples

```excel
=FILL.RANDOM("male";100;"female";100;"child";90)
=FILL.RANDOM("A";5;"B";5)
```
