# ANCOVA

## `ANCOVA.G`

Performs analysis of covariance on grouped data with one factor and one or more covariates.

### Syntax

```excel
=ANCOVA.G(factor; dependent_variable; covariates; [post_hoc]; [alpha]; [has_header])
```

### Arguments

- `factor`: factor categories
- `dependent_variable`: dependent variable
- `covariates`: one or more covariates arranged in columns
- `post_hoc`: optional post-hoc procedure code; default is `0`
- `alpha`: significance level
- `has_header`: optional header mode code; default is `0`

### `post_hoc` Codes

| Code | Name | Description |
| --- | --- | --- |
| `0` | `none` | no post-hoc comparison |
| `1` | `tukey` | conservative Bonferroni fallback |
| `2` | `bonferroni` | pairwise comparisons of adjusted means with Bonferroni correction |
| `3` | `scheffe` | Scheffé-style approximation over adjusted means |
| `4` | `games-howell` | currently implemented as a Bonferroni fallback |

### `has_header` Codes

| Code | Meaning |
| --- | --- |
| `0` | autodetect header |
| `1` | first row is a header |
| `2` | input has no header |

### Notes

- the function requires at least two groups
- covariates are supplied as one or more columns
- incomplete rows are skipped as complete-case rows
- adjusted means are computed at the global means of the covariates
- the main table includes interactions `group × covariate`; if an interaction is significant, a warning about violated slope homogeneity is shown
- effect sizes `η²`, `η²p`, `ω²`, and `ω²p` are returned for the factor, covariates, and interactions

### Output

The spill output contains:

- descriptive statistics by group
- one common ANCOVA table for the factor, individual covariates, and interactions
- an optional warning about violated homogeneity of regression slopes
- adjusted means by group
- an optional post-hoc section

### Example

```excel
=ANCOVA.G(A2:A100;B2:B100;C2:D100;2;0,05;1)
```
