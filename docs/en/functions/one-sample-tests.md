# One-Sample Tests

## `T.TEST.1S`

One-sample t-test.

### Syntax

```excel
=T.TEST.1S(values; mu_0; [direction]; [alpha]; [ma_zahlavi])
```

### Arguments

- `values`: numeric sample
- `mu_0`: hypothesized mean
- `direction`: optional direction code; default is `0`
- `alpha`: significance level
- `ma_zahlavi`: optional header-mode code; default is `0`

### `direction` Codes

| Code | Meaning |
| --- | --- |
| `0` | two-sided test |
| `1` | left-sided test |
| `2` | right-sided test |

### `ma_zahlavi` Codes

| Code | Meaning |
| --- | --- |
| `0` | auto-detect header |
| `1` | first cell is a header |
| `2` | input has no header |

### Output

The spill output contains:

- `x̄`
- `μ₀`
- `sₓ`
- `n`
- `α`
- `t`
- `df`
- critical `t` value
- `p`

## `PROP.TEST.1S`

One-sample z-test for a population proportion.

### Syntax

```excel
=PROP.TEST.1S(values; pi_0; [direction]; [alpha]; [ma_zahlavi])
```

### Arguments

- `values`: binary sample with `0/1` or `FALSE/TRUE`
- `pi_0`: hypothesized proportion
- `direction`: optional direction code; default is `0`
- `alpha`: significance level
- `ma_zahlavi`: optional header-mode code; default is `0`

### `direction` Codes

| Code | Meaning |
| --- | --- |
| `0` | two-sided test |
| `1` | left-sided test |
| `2` | right-sided test |

### `ma_zahlavi` Codes

| Code | Meaning |
| --- | --- |
| `0` | auto-detect header |
| `1` | first cell is a header |
| `2` | input has no header |

### Output

The spill output contains:

- `p̂`
- `π₀`
- `x`
- `n`
- `α`
- `z`
- critical `z` value
- `p`
