# Goodness-Of-Fit

## `CHISQ.GOF`

Chi-square goodness-of-fit test.

### Syntax

```excel
=CHISQ.GOF(observed; expected; [categories]; [alpha]; [ma_zahlavi])
```

### Arguments

- `observed`: observed category counts
- `expected`: expected counts or probabilities
- `categories`: optional category labels
- `alpha`: significance level
- `ma_zahlavi`: optional header-mode code; default is `0`

### `ma_zahlavi` Codes

| Code | Meaning |
| --- | --- |
| `0` | auto-detect header |
| `1` | first row is a header |
| `2` | input has no header |

If `expected` sums to `1`, values are treated as probabilities and converted to expected counts automatically.

### Output

The spill output contains:

- a summary block with `χ²`, `df`, `α`, the critical value, and `p`
- a category contribution block with `O`, `E`, and `(O−E)² / E`
