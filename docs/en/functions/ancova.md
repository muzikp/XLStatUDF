# ANCOVA

## `ANCOVA.G`

Analysis of covariance for grouped data with one factor and one or more covariates.

### Syntax

```excel
=ANCOVA.G(faktor; zavisla_promenna; kovariaty; [post_hoc]; [alpha]; [ma_zahlavi])
```

### Arguments

- `faktor`: factor categories
- `zavisla_promenna`: dependent variable
- `kovariaty`: one or more covariates in columns
- `post_hoc`: optional post-hoc procedure code; default is `0`
- `alpha`: significance level
- `ma_zahlavi`: optional header-mode code; default is `0`

### `post_hoc` Codes

| Code | Name | Description |
| --- | --- | --- |
| `0` | `none` | no post-hoc comparisons |
| `1` | `tukey` | conservative fallback via Bonferroni |
| `2` | `bonferroni` | pairwise comparisons of adjusted means with Bonferroni correction |
| `3` | `scheffe` | Scheffe-style approximation on adjusted means |
| `4` | `games-howell` | currently implemented as a Bonferroni fallback |

### `ma_zahlavi` Codes

| Code | Meaning |
| --- | --- |
| `0` | auto-detect header |
| `1` | first row is a header |
| `2` | input has no header |

### Output

The spill output contains:

- descriptive statistics by group
- one joint ANCOVA table for the factor, individual covariates, and `group × covariate` interactions
- effect sizes `η²`, `η²p`, `ω²`, and `ω²p` for all model terms
- an optional warning when homogeneity of regression slopes is violated
- adjusted means by group
- an optional post-hoc section

### Example

```excel
=ANCOVA.G(A2:A100;B2:B100;C2:D100;2;0,05;1)
```
