# ANOVA

## `ANOVA.G`

Performs one-way analysis of variance on grouped data.

### Syntax

```excel
=ANOVA.G(categories; values; [has_header]; [alpha]; [post_hoc])
```

### Arguments

- `categories`: group labels
- `values`: numeric observations
- `has_header`: `0=autodetect`, `1=first row is a header`, `2=no header`
- `alpha`: significance level
- `post_hoc`: optional post-hoc procedure code; default is `0`

### `post_hoc` Codes

| Code | Name | Description |
| --- | --- | --- |
| `0` | `none` | no post-hoc comparison; only the main ANOVA report is returned |
| `1` | `tukey` | Tukey HSD; currently implemented as a conservative Bonferroni fallback |
| `2` | `bonferroni` | pairwise comparisons with Bonferroni correction |
| `3` | `scheffe` | Scheffé method for multiple comparisons |
| `4` | `games-howell` | pairwise comparisons without assuming equal variances |

### Notes

- the function requires at least three groups
- each group must contain at least two values
- the output includes Levene's test of variance homogeneity
- some post-hoc options are currently implemented via Bonferroni fallback and are explicitly labeled as such in the output

### Output

The spill output contains:

- descriptive statistics by group
- the main ANOVA table
- Levene's homogeneity test
- effect sizes `η²`, `ω²`, and `f`
- an optional post-hoc section

### Examples

```excel
=ANOVA.G(A2:A40;B2:B40)
=ANOVA.G(A2:A40;B2:B40;1;0,05;2)
=ANOVA.G(A2:A40;B2:B40;0;0,05;4)
```

## `ANOVA.RM`

Performs one-factor repeated-measures ANOVA where columns represent conditions and rows represent subjects.

### Syntax

```excel
=ANOVA.RM(values; [has_header]; [alpha]; [post_hoc])
```

### Arguments

- `values`: value matrix; rows are subjects and columns are conditions
- `has_header`: `0=autodetect`, `1=first row is a header`, `2=no header`
- `alpha`: significance level
- `post_hoc`: optional post-hoc procedure code; default is `0`

### `post_hoc` Codes

| Code | Name | Description |
| --- | --- | --- |
| `0` | `none` | no post-hoc comparison; only the main RM ANOVA report is returned |
| `1` | `tukey` | currently implemented as a conservative Bonferroni fallback |
| `2` | `bonferroni` | pairwise condition comparisons via paired t-tests with Bonferroni correction |
| `3` | `scheffe` | currently implemented as a conservative Bonferroni fallback |
| `4` | `games-howell` | currently implemented as a conservative Bonferroni fallback |

### Notes

- the function requires at least two columns and at least two complete rows
- incomplete rows are skipped as complete-case rows
- sphericity is not currently tested; this is also stated in the output

### Output

The spill output contains:

- descriptive statistics by condition
- an RM ANOVA table with rows `Conditions`, `Subjects`, `Residual`, `Total`
- effect sizes `η²`, `η²p`, `ω²`, `ω²p` for the repeated-measures factor
- a note about untested sphericity
- an optional post-hoc section with pairwise condition comparisons

### Example

```excel
=ANOVA.RM(B2:D25)
=ANOVA.RM(B1:D25;1;0,05;2)
```
