# Jednovýběrové Testy

## `T.TEST.1S`

Provádí jednovýběrový t-test.

### Syntaxe

```excel
=T.TEST.1S(values; mu_0; [direction]; [alpha]; [ma_záhlaví])
```

### Argumenty

- `values`: číselný výběr
- `mu_0`: hypotetická střední hodnota
- `direction`: volitelný kód směru testu; výchozí hodnota je `0`
- `alpha`: hladina významnosti
- `ma_záhlaví`: volitelný kód režimu záhlaví; výchozí hodnota je `0`

### Kódy `direction`

| Kód | Význam |
| --- | --- |
| `0` | oboustranný test |
| `1` | levostranný test |
| `2` | pravostranný test |

### Kódy `ma_záhlaví`

| Kód | Význam |
| --- | --- |
| `0` | autodetekce záhlaví |
| `1` | první buňka je záhlaví |
| `2` | vstup je bez záhlaví |

### Poznámky

- prázdné buňky se ignorují
- jsou potřeba alespoň dvě platné hodnoty
- kritická hodnota `t` se počítá podle zvoleného směru testu

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

Provádí jednovýběrový z-test podílu.

### Syntaxe

```excel
=PROP.TEST.1S(values; pi_0; [direction]; [alpha]; [ma_záhlaví])
```

### Argumenty

- `values`: binární výběr ve formátu `0/1` nebo `FALSE/TRUE`
- `pi_0`: hypotetický populační podíl
- `direction`: volitelný kód směru testu; výchozí hodnota je `0`
- `alpha`: hladina významnosti
- `ma_záhlaví`: volitelný kód režimu záhlaví; výchozí hodnota je `0`

### Kódy `direction`

| Kód | Význam |
| --- | --- |
| `0` | oboustranný test |
| `1` | levostranný test |
| `2` | pravostranný test |

### Kódy `ma_záhlaví`

| Kód | Význam |
| --- | --- |
| `0` | autodetekce záhlaví |
| `1` | první buňka je záhlaví |
| `2` | vstup je bez záhlaví |

### Poznámky

- `pi_0` musí ležet přísně mezi `0` a `1`
- funkce přijímá pouze binární data `0/1` nebo `FALSE/TRUE`
- jsou potřeba alespoň jedna platná pozorování

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

## `WILCOXON.PAIRED`

Provádí Wilcoxonův párový znaménkový test pro závislé dvojice.

### Syntaxe

```excel
=WILCOXON.PAIRED(x; y; [ma_záhlaví]; [alpha]; [smer])
```

### Argumenty

- `x`: první měření
- `y`: druhé měření
- `ma_záhlaví`: volitelný kód režimu záhlaví; výchozí hodnota je `0`
- `alpha`: hladina významnosti
- `smer`: volitelný kód směru testu; výchozí hodnota je `0`

### Kódy `smer`

| Kód | Význam |
| --- | --- |
| `0` | oboustranný test |
| `1` | levostranný test |
| `2` | pravostranný test |

### Kódy `ma_záhlaví`

| Kód | Význam |
| --- | --- |
| `0` | autodetekce záhlaví |
| `1` | první buňka je záhlaví |
| `2` | vstup je bez záhlaví |

### Poznámky

- funkce pracuje s párovými rozdíly `x - y`
- nekompletní dvojice se vynechávají po dvojicích
- nulové rozdíly se z testu vyřazují
- pokud po vyřazení nulových rozdílů nezbude dost dat, funkce vrátí chybu počtu

### Výstup

Spill výstup obsahuje:

- `n`
- `med(d)`
- `α`
- `W+`
- `W-`
- `W`
- `z`
- kritickou hodnotu `z`
- `p`
- efekt velikosti `r`

### Příklad

```excel
=WILCOXON.PAIRED(B2:B20;C2:C20)
=WILCOXON.PAIRED(B2:B20;C2:C20;1;0,05;0)
```
