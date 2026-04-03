# Dvouvýběrové Testy

## `WELCH.TEST.2S.G`

Provádí Welchův dvouvýběrový t-test pro dvě nezávislé skupiny.

### Syntaxe

```excel
=WELCH.TEST.2S.G(categories; values; [ma_záhlaví]; [alpha]; [direction])
```

### Argumenty

- `categories`: štítky definující právě dvě skupiny
- `values`: číselná pozorování
- `ma_záhlaví`: volitelný kód režimu záhlaví; výchozí hodnota je `0`
- `alpha`: hladina významnosti
- `direction`: volitelný kód směru testu; výchozí hodnota je `0`

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

- funkce vyžaduje právě dvě skupiny
- v každé skupině musí být alespoň dvě hodnoty
- skupiny jsou interně seřazeny podle názvu, což ovlivní znaménko rozdílu i statistiky `t`
- ve výstupu jsou popisné statistiky zhuštěny do tabulky se skupinami po řádcích

### Výstup

Spill výstup obsahuje:

- tabulku popisných statistik po skupinách
- `α`
- `t`
- Welch-Satterthwaite `df`
- kritickou hodnotu `t`
- `p`
- Cohenovo `d`
- velikost účinku `r`

### Příklad

```excel
=WELCH.TEST.2S.G(A2:A40;B2:B40;1;0,05;0)
```

## `MANN.WHITNEY.G`

Provádí Mann-Whitneyho neparametrický test pro dvě nezávislé skupiny.

### Syntaxe

```excel
=MANN.WHITNEY.G(kategorie; hodnoty; [ma_záhlaví]; [alpha]; [smer])
```

### Argumenty

- `kategorie`: štítky právě dvou skupin
- `hodnoty`: číselná pozorování
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
| `1` | první řádek je záhlaví |
| `2` | vstup je bez záhlaví |

### Poznámky

- funkce vyžaduje právě dvě skupiny
- v každé skupině musí být alespoň jedna hodnota
- používají se střední pořadí při shodách
- při nulové rozptylové složce po tie correction vrátí funkce numerickou chybu

### Výstup

Spill výstup obsahuje:

- tabulku popisných statistik po skupinách
- `U`
- `U₁`
- `U₂`
- `z`
- kritickou hodnotu `z`
- `p`
- efekt velikosti `r`

### Příklad

```excel
=MANN.WHITNEY.G(A2:A20;B2:B20)
=MANN.WHITNEY.G(A2:A20;B2:B20;1;0,05;0)
```
