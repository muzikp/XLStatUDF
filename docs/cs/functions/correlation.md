# Korelace

## `CORREL.SPEARMAN`

Počítá Spearmanův pořadový korelační koeficient a provádí jeho test významnosti.

### Syntaxe

```excel
=CORREL.SPEARMAN(x_values; y_values; [direction]; [alpha]; [ma_záhlaví])
```

### Argumenty

- `x_values`: první číselný rozsah
- `y_values`: druhý číselný rozsah
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
| `1` | první řádek je záhlaví |
| `2` | vstup je bez záhlaví |

### Poznámky

- prázdné buňky se vylučují po dvojicích
- funkce vyžaduje alespoň tři platné dvojice
- při nulové variabilitě pořadí v některé proměnné vrací funkce numerickou chybu

### Výstup

Spill výstup obsahuje:

- `ρ`
- `n`
- `α`
- `t`
- `df`
- kritickou hodnotu `t`
- `p`

## `CORREL.MATRIX`

Sestavuje korelační matici pro data se dvěma a více sloupci.

### Syntaxe

```excel
=CORREL.MATRIX(data; [metoda]; [vystup]; [p_minimum]; [ma_záhlaví])
```

### Argumenty

- `data`: vstupní matice; sloupce jsou proměnné
- `metoda`: volitelný kód výpočetní metody; výchozí hodnota je `0`
- `vystup`: volitelný kód typu výstupu; výchozí hodnota je `0`
- `p_minimum`: volitelný filtr; vrátí jen vazby s `p < p_minimum`
- `ma_záhlaví`: volitelný kód režimu záhlaví; výchozí hodnota je `0`

### Kódy `metoda`

| Kód | Význam |
| --- | --- |
| `0` | Pearson |
| `1` | Spearman |

### Kódy `vystup`

| Kód | Význam |
| --- | --- |
| `0` | jen korelační koeficienty |
| `1` | jen oboustranné p-hodnoty |
| `2` | korelační koeficient a v řádku pod ním p-hodnota |
| `3` | korelační koeficient, pod ním p-hodnota a na třetím řádku signifikance hvězdičkami |
| `4` | jen korelační koeficienty a za nimi v téže buňce signifikance hvězdičkami |

### Kódy `ma_záhlaví`

| Kód | Význam |
| --- | --- |
| `0` | autodetekce záhlaví |
| `1` | první řádek je záhlaví |
| `2` | vstup je bez záhlaví |

### Poznámky

- funkce vyžaduje alespoň dva sloupce a alespoň tři kompletní řádky
- nekompletní řádky se vynechávají jako complete-case
- pokud je některý sloupec konstantní, funkce vrátí numerickou chybu
- pokud je zadáno `p_minimum`, nevyhovující vazby se ve výstupu nechají prázdné včetně diagonály
- doplněk kvůli zpětné kompatibilitě stále přijímá i původní pořadí argumentů `([p_minimum]; data; [metoda]; [vystup]; [ma_záhlaví])`

### Výstup

U výstupů `0`, `1` a `4` vrací spill výstup čtvercovou matici s názvy proměnných.

U výstupů `2` a `3` vrací spill výstup skládané rozvržení se dvěma úvodními popisnými sloupci a jedním blokem řádků pro každou proměnnou:

- koeficient
- `p`
- volitelně `sig.`

U výstupu `4` je v každé buňce text ve formátu například `0,30156***`.

### Příklady

```excel
=CORREL.MATRIX(B1:E30)
=CORREL.MATRIX(B1:E30;1;0;;1)
=CORREL.MATRIX(B1:E30;0;4;0,05;1)
```
