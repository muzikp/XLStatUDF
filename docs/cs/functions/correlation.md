# Korelace

## `CORREL.SPEARMAN`

Spearmanův pořadový korelační koeficient s testem významnosti.

### Syntaxe

```excel
=CORREL.SPEARMAN(x_values; y_values; [direction]; [alpha]; [ma_zahlavi])
```

### Argumenty

- `x_values`: první číselný rozsah
- `y_values`: druhý číselný rozsah
- `direction`: volitelný kód směru testu; výchozí hodnota je `0`
- `alpha`: hladina významnosti
- `ma_zahlavi`: volitelný kód režimu záhlaví; výchozí hodnota je `0`

### Kódy `direction`

| Kód | Význam |
| --- | --- |
| `0` | oboustranný test |
| `1` | levostranný test |
| `2` | pravostranný test |

### Kódy `ma_zahlavi`

| Kód | Význam |
| --- | --- |
| `0` | autodetekce záhlaví |
| `1` | první řádek je záhlaví |
| `2` | vstup je bez záhlaví |

Prázdné buňky se vylučují po dvojicích.

### Výstup

Spill výstup obsahuje:

- `ρ`
- `n`
- `α`
- `t`
- `df`
- kritickou hodnotu `t`
- `p`
