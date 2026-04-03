# Distributions

## `NORM.DIST.RANGE`

Computes the probability that a normally distributed random variable falls within a given interval.

### Syntax

```excel
=NORM.DIST.RANGE(mean; standard_deviation; lower_bound; upper_bound)
```

### Arguments

- `mean`: distribution mean
- `standard_deviation`: positive standard deviation
- `lower_bound`: lower interval bound; a blank cell means minus infinity
- `upper_bound`: upper interval bound; a blank cell means plus infinity

### Notes

- if `lower_bound > upper_bound`, the function returns a numeric error
- if `standard_deviation <= 0`, the function returns a numeric error
- blank bounds can be used for one-sided intervals

### Output

A scalar value in the range `[0;1]`.

### Example

```excel
=NORM.DIST.RANGE(0;1;-1;1)
```

## `GENERATE.NORM`

Generates a single random value from a normal distribution with optional perturbation.

### Syntax

```excel
=GENERATE.NORM(mean; stdev; [outlier_rate])
```

### Arguments

- `mean`: desired mean
- `stdev`: positive standard deviation
- `outlier_rate`: optional probability of additional random perturbation in the interval `<0;1>`; default is `0`

### Notes

- the function is volatile, so recalculation generates a new draw
- when `outlier_rate = 0`, the function returns a regular draw from `N(μ, σ)`
- when perturbation occurs, additional noise from `N(0, 3σ)` is added to the generated value
- invalid `outlier_rate` or non-numeric inputs return a value error

### Output

A single scalar value.

### Example

```excel
=GENERATE.NORM(100;15)
=GENERATE.NORM(0;1;0,2)
```

## `GENERATE.INT`

Generates a single random integer from a given interval with optional perturbation.

### Syntax

```excel
=GENERATE.INT([minimum]; [maximum]; [outlier_rate])
```

### Arguments

- `minimum`: optional lower bound; default is `-2147483648`
- `maximum`: optional upper bound; default is `2147483647`
- `outlier_rate`: optional probability of additional random perturbation in the interval `<0;1>`; default is `0`

### Notes

- the function is volatile, so recalculation generates a new value
- if `minimum > maximum`, the function returns a numeric error
- if bounds are omitted, the full practical 32-bit range is used
- when `outlier_rate = 0`, the function returns a regular draw from the closed interval `<minimum; maximum>`
- when perturbation occurs, an additional random integer offset from `⟨-(maximum-minimum); +(maximum-minimum)⟩` is added

### Output

A single integer value.

### Examples

```excel
=GENERATE.INT()
=GENERATE.INT(1;6)
=GENERATE.INT(1;6;0,2)
```

## `FILL`

Repeats one or more values, or repeatedly evaluates a text formula, into a single-column spill output.

### Syntax

```excel
=FILL(what; count; [what2]; [count2]; ...)
```

### Arguments

- `what`: a scalar value to repeat, or a text formula starting with `=`
- `count`: number of returned rows; integer `>= 1`
- additional arguments must be provided as `what + count` pairs

### Notes

- if you pass a regular value, the function simply copies it into every row
- if you pass a text formula starting with `=`, the formula is evaluated separately for each row
- the number of additional arguments must be even; otherwise the function returns a value error
- a direct argument such as `GENERATE.NORM(...)` or `RANDBETWEEN(...)` is evaluated by Excel before `FILL` is called, so repeated recalculation requires a text formula

### Output

A single-column spill range with the resulting sequence.

### Examples

```excel
=FILL("A";5)
=FILL(123;4)
=FILL("male";100;"female";100;"child";90)
=FILL("=RANDBETWEEN(1;10)";20)
=FILL("=GENERATE.NORM(0;1;0,2)";100)
```

## `FILL.RANDOM`

Builds a sequence like `FILL`, then shuffles it randomly before returning it.

### Syntax

```excel
=FILL.RANDOM(what; count; [what2]; [count2]; ...)
```

### Arguments

Same as `FILL`.

### Notes

- the full sequence is first created in the same way as `FILL`
- the completed sequence is then shuffled randomly
- the same validation rules as `FILL` apply here as well

### Output

A single-column spill range containing all generated values in random order.

### Examples

```excel
=FILL.RANDOM("male";100;"female";100;"child";90)
=FILL.RANDOM("A";5;"B";5)
```
