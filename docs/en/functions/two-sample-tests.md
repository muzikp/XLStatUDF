# Two-Sample Tests

## `WELCH.TEST.2S.G`

Welch two-sample t-test for two independent groups with unequal variances.

### Syntax

```excel
=WELCH.TEST.2S.G(categories; values; [ma_zahlavi]; [alpha]; [direction])
```

### Arguments

- `categories`: labels defining exactly two groups, typically the first column
- `values`: numeric observations, typically the second adjacent column with measured values
- `ma_zahlavi`: `0=autodetect`, `1=first row is a header`, `2=no header`
- `alpha`: significance level
- `direction`: `0=two-sided test`, `1=left-sided test`, `2=right-sided test`

### Output

The spill output contains:

- descriptive statistics for group 1
- descriptive statistics for group 2
- `α`
- `t`
- Welch-Satterthwaite `df`
- critical `t` region
- `p`
- Cohen's `d`
- effect size `r` (normalized Cohen's `d`)

### Example

```excel
=WELCH.TEST.2S.G(A2:A40;B2:B40;1;0,05;0)
```
