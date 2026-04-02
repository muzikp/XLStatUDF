# ANOVA

## `ANOVA.G`

Jednofaktorová analýza rozptylu nad groupovanými daty.

### Syntaxe

```excel
=ANOVA.G(categories; values; [post_hoc]; [alpha]; [ma_zahlavi])
```

### Argumenty

- `categories`: štítky skupin
- `values`: číselná pozorování
- `post_hoc`: volitelný kód post-hoc procedury; výchozí hodnota je `0`
- `alpha`: hladina významnosti
- `ma_zahlavi`: `0=autodetect`, `1=první řádek je záhlaví`, `2=bez záhlaví`

### Kódy `post_hoc`

| Kód | Název | Popis |
| --- | --- | --- |
| `0` | `none` | bez post-hoc porovnání; vrátí se jen hlavní ANOVA report |
| `1` | `tukey` | Tukey HSD; v aktuální implementaci konzervativní fallback přes Bonferroni |
| `2` | `bonferroni` | párová porovnání s Bonferroniho korekcí |
| `3` | `scheffe` | Scheffého metoda pro vícenásobná porovnání |
| `4` | `games-howell` | párová porovnání bez předpokladu shodných rozptylů |

### Výstup

Spill výstup obsahuje:

- popisné statistiky po skupinách
- hlavní ANOVA tabulku
- Leveneho test homogenity rozptylů
- velikosti účinku `η²`, `ω²` a `f`
- volitelnou post-hoc část

### Příklady

```excel
=ANOVA.G(A2:A40;B2:B40)
=ANOVA.G(A2:A40;B2:B40;2;0,05;1)
=ANOVA.G(A2:A40;B2:B40;4;0,05;0)
```
