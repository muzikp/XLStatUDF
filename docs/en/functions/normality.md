# Normality And Fit Tests

## `SHAPIRO.WILK`

Performs the Shapiro-Wilk normality test.

### Syntax

```excel
=SHAPIRO.WILK(values; [has_header])
```

### Arguments

- `values`: numeric sample; blank cells are ignored
- `has_header`: `0=autodetect`, `1=first cell is a header`, `2=no header`

### Notes

- the function requires a sample size from `3` to `5000`
- the data are sorted in ascending order before the statistic is computed
- if all observations are identical, the statistic becomes degenerate and interpretation is not meaningful

### Output

A two-row spill output:

- `W`: test statistic
- `p`: p-value

### Example

```excel
=SHAPIRO.WILK(B2:B18)
```

## `KOLMOGOROV.SMIRNOV`

Performs the one-sample Kolmogorov-Smirnov goodness-of-fit test.

### Syntax

```excel
=KOLMOGOROV.SMIRNOV(values; [distribution]; [has_header])
```

### Arguments

- `values`: numeric sample; blank cells are ignored
- `distribution`: optional code of the tested distribution; default is `0`
- `has_header`: `0=autodetect`, `1=first cell is a header`, `2=no header`

### `distribution` Codes

| Code | Distribution | What is tested |
| --- | --- | --- |
| `0` | `normal` | whether the data follow a normal distribution with mean and standard deviation estimated from the sample |
| `1` | `lognormal` | whether the data follow a lognormal distribution; all values must be positive |
| `2` | `exponential` | whether the data follow an exponential distribution; all values must be non-negative |
| `3` | `uniform` | whether the data follow a continuous uniform distribution on the interval determined by the sample minimum and maximum |
| `4` | `weibull` | whether the data follow a Weibull distribution with parameters estimated from the sample |

### Notes

- the function requires at least `5` valid observations
- the data are sorted in ascending order before the statistic is computed
- some distributions impose additional input constraints:
  - `lognormal` and `weibull`: positive values only
  - `exponential`: non-negative values only
  - `uniform`: the sample must not be constant
- the p-value for `normal` uses a different correction than the other distributions

### Output

A two-row spill output:

- `D`: test statistic
- `p`: p-value

### Examples

```excel
=KOLMOGOROV.SMIRNOV(B2:B18)
=KOLMOGOROV.SMIRNOV(B2:B18;1)
=KOLMOGOROV.SMIRNOV(B2:B18;4;1)
```
