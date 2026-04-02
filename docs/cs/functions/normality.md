# Testy Normality A Shody

## `SHAPIRO.WILK`

Shapiro-Wilkův test normality.

### Syntaxe

```excel
=SHAPIRO.WILK(values; [ma_zahlavi])
```

### Argumenty

- `values`: číselný výběr; prázdné buňky jsou ignorovány
- `ma_zahlavi`: `0=autodetect`, `1=první buňka je záhlaví`, `2=bez záhlaví`

### Výstup

Dvouřádkový spill výstup:

- `W`: testová statistika
- `p`: p-hodnota

### Příklad

```excel
=SHAPIRO.WILK(B2:B18)
```

## `KOLMOGOROV.SMIRNOV`

Jednovýběrový Kolmogorov-Smirnovův test dobré shody.

### Syntaxe

```excel
=KOLMOGOROV.SMIRNOV(values; [distribution]; [ma_zahlavi])
```

### Argumenty

- `values`: číselný výběr; prázdné buňky jsou ignorovány
- `distribution`: volitelný kód testovaného rozdělení; výchozí hodnota je `0`
- `ma_zahlavi`: `0=autodetect`, `1=první buňka je záhlaví`, `2=bez záhlaví`

### Kódy Rozdělení

| Kód | Rozdělení | Co se testuje |
| --- | --- | --- |
| `0` | `normal` | zda data odpovídají normálnímu rozdělení s průměrem a směrodatnou odchylkou odhadnutou ze vzorku |
| `1` | `lognormal` | zda data odpovídají lognormálnímu rozdělení; všechny hodnoty musí být kladné |
| `2` | `exponential` | zda data odpovídají exponenciálnímu rozdělení; všechny hodnoty musí být nezáporné |
| `3` | `uniform` | zda data odpovídají spojitému rovnoměrnému rozdělení na intervalu daném minimem a maximem vzorku |
| `4` | `weibull` | zda data odpovídají Weibullovu rozdělení s parametry odhadnutými ze vzorku |

### Výstup

Dvouřádkový spill výstup:

- `D`: testová statistika
- `p`: p-hodnota

### Příklady

```excel
=KOLMOGOROV.SMIRNOV(B2:B18)
=KOLMOGOROV.SMIRNOV(B2:B18;1)
=KOLMOGOROV.SMIRNOV(B2:B18;4;1)
```
