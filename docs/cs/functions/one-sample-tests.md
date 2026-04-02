# Jednovýběrové Testy

## `T.TEST.1S`

Jednovýběrový t-test.

### Syntaxe

```excel
=T.TEST.1S(values; mu_0; [direction]; [alpha]; [ma_zahlavi])
```

### Argumenty

- `values`: číselný výběr
- `mu_0`: hypotetická střední hodnota
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
| `1` | první buňka je záhlaví |
| `2` | vstup je bez záhlaví |

### Výstup

Spill výstup obsahuje:

- `x̄`
- `μ₀`
- `sₓ`
- `n`
- `α`
- `t`
- `df`
- kritickou hodnotu `t`
- `p`

## `PROP.TEST.1S`

Jednovýběrový z-test podílu.

### Syntaxe

```excel
=PROP.TEST.1S(values; pi_0; [direction]; [alpha]; [ma_zahlavi])
```

### Argumenty

- `values`: binární výběr ve formátu `0/1` nebo `FALSE/TRUE`
- `pi_0`: hypotetický populační podíl
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
| `1` | první buňka je záhlaví |
| `2` | vstup je bez záhlaví |

### Výstup

Spill výstup obsahuje:

- `p̂`
- `π₀`
- `x`
- `n`
- `α`
- `z`
- kritickou hodnotu `z`
- `p`
