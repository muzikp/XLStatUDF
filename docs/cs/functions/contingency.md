# Kontingenční Tabulky

## `CONTINGENCY.T`

Analýza kontingenční tabulky zadané přímo jako matice četností.

### Syntaxe

```excel
=CONTINGENCY.T(tabulka; [ma_zahlavi]; [alpha])
```

### Argumenty

- `tabulka`: 2D rozsah s pozorovanými četnostmi
- `ma_zahlavi`: volitelný kód režimu záhlaví; výchozí hodnota je `0`
- `alpha`: hladina významnosti

### Kódy `ma_zahlavi`

| Kód | Význam |
| --- | --- |
| `0` | autodetekce popisků |
| `1` | horní řádek i levý sloupec jsou popisky |
| `2` | tabulka je čistě numerická bez popisků |

### Výstup

Spill výstup obsahuje:

- pozorovanou kontingenční tabulku včetně marginálních součtů
- tabulku očekávaných četností
- souhrn testu `n`, `df`, `α`, `χ²`, kritickou hodnotu a `p`
- míry asociace `Pearson C`, `Cramer V` a `phi` pro tabulku `2x2`

### Příklad

```excel
=CONTINGENCY.T(A1:C3;1;0,05)
```

## `CONTINGENCY.G`

Analýza kontingenční tabulky z groupovaných dat.

### Syntaxe

```excel
=CONTINGENCY.G(sloupce; radky; [pocet]; [alpha]; [ma_zahlavi])
```

### Argumenty

- `sloupce`: kategorie budoucích sloupců kontingenční tabulky
- `radky`: kategorie budoucích řádků kontingenční tabulky
- `pocet`: volitelné četnosti dvojic; pokud chybí, každá dvojice má váhu `1`
- `alpha`: hladina významnosti
- `ma_zahlavi`: volitelný kód režimu záhlaví; výchozí hodnota je `0`

### Kódy `ma_zahlavi`

| Kód | Význam |
| --- | --- |
| `0` | autodetekce záhlaví |
| `1` | první řádek je záhlaví |
| `2` | vstup je bez záhlaví |

### Výstup

Spill výstup obsahuje stejné části jako `CONTINGENCY.T`:

- pozorovanou kontingenční tabulku
- očekávané četnosti
- souhrn testu `χ²`
- míry asociace `Pearson C`, `Cramer V` a případně `phi`

### Příklady

```excel
=CONTINGENCY.G(A2:A100;B2:B100)
=CONTINGENCY.G(A2:A100;B2:B100;C2:C100;0,05;1)
```
