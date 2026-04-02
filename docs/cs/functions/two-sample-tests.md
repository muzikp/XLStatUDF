# Dvouvýběrové Testy

## `WELCH.TEST.2S.G`

Welchův dvouvýběrový t-test pro dvě nezávislé skupiny s nerovností rozptylů.

### Syntaxe

```excel
=WELCH.TEST.2S.G(categories; values; [ma_zahlavi]; [alpha]; [direction])
```

### Argumenty

- `categories`: štítky definující právě dvě skupiny, typicky první sloupec
- `values`: číselná pozorování, typicky druhý sousední sloupec s měřenými hodnotami
- `ma_zahlavi`: volitelný kód režimu záhlaví; výchozí hodnota je `0`
- `alpha`: hladina významnosti
- `direction`: volitelný kód směru testu; výchozí hodnota je `0`

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

### Výstup

Spill výstup obsahuje:

- tabulku popisných statistik po skupinách
- `α`
- `t`
- Welch-Satterthwaite `df`
- kritický obor `t`
- `p`
- Cohenovo `d`
- velikost účinku `r` (normalizované Cohenovo `d`)

### Příklad

```excel
=WELCH.TEST.2S.G(A2:A40;B2:B40;1;0,05;0)
```
