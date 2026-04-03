# Goodness-Of-Fit Test

## `CHISQ.GOF`

Performs the chi-square goodness-of-fit test.

### Syntax

```excel
=CHISQ.GOF(observed; expected; [categories]; [alpha]; [has_header])
```

### Arguments

- `observed`: observed category frequencies
- `expected`: expected frequencies or probabilities
- `categories`: optional category labels
- `alpha`: significance level
- `has_header`: optional header mode code; default is `0`

### `has_header` Codes

| Code | Meaning |
| --- | --- |
| `0` | autodetect header |
| `1` | first row is a header |
| `2` | input has no header |

### Notes

- `observed` must contain non-negative integer frequencies
- `expected` may be provided either as frequencies or as probabilities summing to `1`
- if `expected` is given as probabilities, the function automatically rescales them to expected counts using the sample size
- if `categories` is omitted, categories are labeled `1, 2, 3, ...`

### Output

The spill output contains:

- a test summary with `χ²`, `df`, `α`, critical value, and `p`
- a category table with `O`, `E`, and contribution `(O−E)² / E`
