# ANOVA

## `ANOVA.G`

Jednofaktorová analýza rozptylu nad groupovanými daty.

### Syntaxe

```excel
=ANOVA.G(categories; values; [ma_zahlavi]; [alpha]; [post_hoc])
```

### Argumenty

- `categories`: štítky skupin
- `values`: číselná pozorování
- `ma_zahlavi`: `0=autodetect`, `1=první řádek je záhlaví`, `2=bez záhlaví`
- `alpha`: hladina významnosti
- `post_hoc`: volitelný kód post-hoc procedury; výchozí hodnota je `0`

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
=ANOVA.G(A2:A40;B2:B40;1;0,05;2)
=ANOVA.G(A2:A40;B2:B40;0;0,05;4)
```

## `ANOVA.RM`

Jednofaktorová ANOVA s opakovaným měřením, kde sloupce představují podmínky a řádky subjekty.

### Syntaxe

```excel
=ANOVA.RM(hodnoty; [ma_zahlavi]; [alpha]; [post_hoc])
```

### Argumenty

- `hodnoty`: matice hodnot; řádky jsou subjekty, sloupce podmínky
- `ma_zahlavi`: `0=autodetect`, `1=první řádek je záhlaví`, `2=bez záhlaví`
- `alpha`: hladina významnosti
- `post_hoc`: volitelný kód post-hoc procedury; výchozí hodnota je `0`

### Kódy `post_hoc`

| Kód | Název | Popis |
| --- | --- | --- |
| `0` | `none` | bez post-hoc porovnání; vrátí se jen hlavní RM ANOVA report |
| `1` | `tukey` | v aktuální implementaci konzervativní fallback přes Bonferroni |
| `2` | `bonferroni` | párová porovnání podmínek přes párové t-testy s Bonferroniho korekcí |
| `3` | `scheffe` | v aktuální implementaci konzervativní fallback přes Bonferroni |
| `4` | `games-howell` | v aktuální implementaci konzervativní fallback přes Bonferroni |

### Výstup

Spill výstup obsahuje:

- popisné statistiky po podmínkách
- tabulku RM ANOVA s řádky `Podminky`, `Subjekty`, `Reziduum`, `Celkem`
- velikosti účinku `η²`, `η²p`, `ω²`, `ω²p` pro faktor v rámci opakovaného měření
- stručnou poznámku, že sphericita zatím není testována
- volitelnou post-hoc část s párovými porovnáními podmínek

### Příklad

```excel
=ANOVA.RM(B2:D25)
=ANOVA.RM(B1:D25;1;0,05;2)
```
