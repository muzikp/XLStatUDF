# One-Sample Tests

## `T.TEST.1S`

Performs a one-sample t-test.

### Syntax

```excel
=T.TEST.1S(values; mu_0; [direction]; [alpha]; [has_header])
```

### Arguments

- `values`: numeric sample
- `mu_0`: hypothetical mean
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
| `1` | first cell is a header |
| `2` | input has no header |

### Notes

- blank cells are ignored
- at least two valid values are required
- the critical `t` value is computed according to the selected test direction

### Output

The spill output contains:

- `x̄`
- `μ₀`
- `sₓ`
- `n`
- `α`
- `t`
- `df`
- critical `t`
- `p`

## `PROP.TEST.1S`

Performs a one-sample z-test for a proportion.

### Syntax

```excel
=PROP.TEST.1S(values; pi_0; [direction]; [alpha]; [has_header])
```

### Arguments

- `values`: binary sample in `0/1` or `FALSE/TRUE` form
- `pi_0`: hypothetical population proportion
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
| `1` | first cell is a header |
| `2` | input has no header |

### Notes

- `pi_0` must lie strictly between `0` and `1`
- the function accepts only binary data `0/1` or `FALSE/TRUE`
- at least one valid observation is required

### Output

The spill output contains:

- `p̂`
- `π₀`
- `x`
- `n`
- `α`
- `z`
- critical `z`
- `p`

## `WILCOXON.PAIRED`

Performs the Wilcoxon signed-rank test for dependent pairs.

### Syntax

```excel
=WILCOXON.PAIRED(x; y; [has_header]; [alpha]; [direction])
```

### Arguments

- `x`: first measurement
- `y`: second measurement
- `has_header`: optional header mode code; default is `0`
- `alpha`: significance level
- `direction`: optional direction code; default is `0`

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
| `1` | first cell is a header |
| `2` | input has no header |

### Notes

- the function works with paired differences `x - y`
- incomplete pairs are skipped pairwise
- zero differences are excluded from the test
- if too few values remain after removing zero differences, the function returns a count error

### Output

The spill output contains:

- `n`
- `med(d)`
- `α`
- `W+`
- `W-`
- `W`
- `z`
- critical `z`
- `p`
- effect size `r`

### Example

```excel
=WILCOXON.PAIRED(B2:B20;C2:C20)
=WILCOXON.PAIRED(B2:B20;C2:C20;1;0,05;0)
```
