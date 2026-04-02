# Normality Tests

## `SHAPIRO.WILK`

Shapiro-Wilk normality test.

### Syntax

```excel
=SHAPIRO.WILK(values; [ma_zahlavi])
```

### Arguments

- `values`: numeric sample; blank cells are ignored
- `ma_zahlavi`: optional header-mode code; default is `0`

### `ma_zahlavi` Codes

| Code | Meaning |
| --- | --- |
| `0` | auto-detect header |
| `1` | first cell is a header |
| `2` | input has no header |

### Output

2-row spill output:

- `W`: test statistic
- `p`: p-value

## `KOLMOGOROV.SMIRNOV`

One-sample Kolmogorov-Smirnov goodness-of-fit test.

### Syntax

```excel
=KOLMOGOROV.SMIRNOV(values; [distribution]; [ma_zahlavi])
```

### Arguments

- `values`: numeric sample; blank cells are ignored
- `distribution`: optional distribution code; default is `0`
- `ma_zahlavi`: optional header-mode code; default is `0`

### `distribution` Codes

| Code | Distribution | What is tested |
| --- | --- | --- |
| `0` | `normal` | sample fit to a normal distribution with mean and standard deviation estimated from the sample |
| `1` | `lognormal` | sample fit to a lognormal distribution; all values must be positive |
| `2` | `exponential` | sample fit to an exponential distribution; all values must be non-negative |
| `3` | `uniform` | sample fit to a continuous uniform distribution on the sample min-max interval |
| `4` | `weibull` | sample fit to a Weibull distribution with parameters estimated from the sample |

### `ma_zahlavi` Codes

| Code | Meaning |
| --- | --- |
| `0` | auto-detect header |
| `1` | first cell is a header |
| `2` | input has no header |
