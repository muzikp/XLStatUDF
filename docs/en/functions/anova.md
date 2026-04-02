# ANOVA

## `ANOVA.G`

One-way ANOVA for grouped data.

### Syntax

```excel
=ANOVA.G(categories; values; [ma_zahlavi]; [alpha]; [post_hoc])
```

### Arguments

- `categories`: group labels
- `values`: numeric observations
- `ma_zahlavi`: `0=autodetect`, `1=header present`, `2=no header`
- `alpha`: significance level
- `post_hoc`: optional post-hoc procedure code; default is `0`

### `post_hoc` Codes

| Code | Name | Description |
| --- | --- | --- |
| `0` | `none` | no post-hoc comparisons; only the main ANOVA report is returned |
| `1` | `tukey` | Tukey HSD; currently implemented as a conservative Bonferroni-style fallback |
| `2` | `bonferroni` | pairwise comparisons with Bonferroni correction |
| `3` | `scheffe` | Scheffe multiple-comparison procedure |
| `4` | `games-howell` | pairwise comparisons without equal-variance assumption |

### Output

The spill report contains:

- descriptive statistics by group
- the main ANOVA table
- Levene homogeneity-of-variance test
- effect sizes `η²`, `ω²`, and `f`
- an optional post-hoc section

Sample standard deviation is labeled as `sₓ`.

### Examples

```excel
=ANOVA.G(A2:A40;B2:B40)
=ANOVA.G(A2:A40;B2:B40;1;0,05;2)
```
