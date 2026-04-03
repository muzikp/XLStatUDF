# Testy Normality A Shody

## `SHAPIRO.WILK`

Provádí Shapiro-Wilkův test normality.

### Syntaxe

```excel
=SHAPIRO.WILK(values; [ma_záhlaví])
```

### Argumenty

- `values`: číselný výběr; prázdné buňky jsou ignorovány
- `ma_záhlaví`: `0=autodetect`, `1=první buňka je záhlaví`, `2=bez záhlaví`

### Poznámky

- funkce vyžaduje velikost vzorku od `3` do `5000`
- data se před výpočtem seřadí vzestupně
- pokud mají všechna pozorování stejnou hodnotu, statistika vyjde hraničně a interpretace není smysluplná

### Výstup

Dvouřádkový spill výstup:

- `W`: testová statistika
- `p`: p-hodnota

### Příklad

```excel
=SHAPIRO.WILK(B2:B18)
```

## `KOLMOGOROV.SMIRNOV`

Provádí jednovýběrový Kolmogorov-Smirnovův test dobré shody.

### Syntaxe

```excel
=KOLMOGOROV.SMIRNOV(values; [distribution]; [ma_záhlaví])
```

### Argumenty

- `values`: číselný výběr; prázdné buňky jsou ignorovány
- `distribution`: volitelný kód testovaného rozdělení; výchozí hodnota je `0`
- `ma_záhlaví`: `0=autodetect`, `1=první buňka je záhlaví`, `2=bez záhlaví`

### Kódy `distribution`

| Kód | Rozdělení | Co se testuje |
| --- | --- | --- |
| `0` | `normal` | zda data odpovídají normálnímu rozdělení s průměrem a směrodatnou odchylkou odhadnutou ze vzorku |
| `1` | `lognormal` | zda data odpovídají lognormálnímu rozdělení; všechny hodnoty musí být kladné |
| `2` | `exponential` | zda data odpovídají exponenciálnímu rozdělení; všechny hodnoty musí být nezáporné |
| `3` | `uniform` | zda data odpovídají spojitému rovnoměrnému rozdělení na intervalu daném minimem a maximem vzorku |
| `4` | `weibull` | zda data odpovídají Weibullovu rozdělení s parametry odhadnutými ze vzorku |

### Poznámky

- funkce vyžaduje alespoň `5` platných pozorování
- data se před výpočtem seřadí vzestupně
- některá rozdělení mají dodatečné podmínky na vstup:
  - `lognormal` a `weibull`: pouze kladné hodnoty
  - `exponential`: pouze nezáporné hodnoty
  - `uniform`: data nesmí být konstantní
- p-hodnota pro `normal` používá korekci odlišnou od ostatních rozdělení

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
