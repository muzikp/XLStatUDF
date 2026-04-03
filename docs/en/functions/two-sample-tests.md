# Two-Sample Tests

## `WELCH.TEST.2S.G`

Performs Welch's two-sample t-test for two independent groups.

### Syntax

```excel
=WELCH.TEST.2S.G(categories; values; [has_header]; [alpha]; [direction])
```

### Arguments

- `categories`: labels defining exactly two groups
- `values`: numeric observations
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
| `1` | first row is a header |
| `2` | input has no header |

### Notes

- the function requires exactly two groups
- each group must contain at least two values
- groups are internally sorted by label, which affects the sign of the difference and the `t` statistic
- descriptive statistics are returned in a compact table with groups as rows

### Output

The spill output contains:

- a descriptive statistics table by group
- `α`
- `t`
- Welch-Satterthwaite `df`
- critical `t`
- `p`
- Cohen's `d`
- effect size `r`

### Example

```excel
=WELCH.TEST.2S.G(A2:A40;B2:B40;1;0,05;0)
```

## `MANN.WHITNEY.G`

Performs the nonparametric Mann-Whitney test for two independent groups.

### Syntax

```excel
=MANN.WHITNEY.G(categories; values; [has_header]; [alpha]; [direction])
```

### Arguments

- `categories`: labels of exactly two groups
- `values`: numeric observations
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
| `1` | first row is a header |
| `2` | input has no header |

### Notes

- the function requires exactly two groups
- each group must contain at least one value
- mid-ranks are used for ties
- if the variance term becomes zero after tie correction, the function returns a numeric error

### Output

The spill output contains:

- a descriptive statistics table by group
- `U`
- `U₁`
- `U₂`
- `z`
- critical `z`
- `p`
- effect size `r`

### Example

```excel
=MANN.WHITNEY.G(A2:A20;B2:B20)
=MANN.WHITNEY.G(A2:A20;B2:B20;1;0,05;0)
```
