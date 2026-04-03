# Kontingenční Tabulky

## `CONTINGENCY.T`

Analyzuje kontingenční tabulku zadanou přímo jako matici četností.

### Syntaxe

```excel
=CONTINGENCY.T(tabulka; [ma_záhlaví]; [alpha])
```

### Argumenty

- `tabulka`: 2D rozsah s pozorovanými četnostmi
- `ma_záhlaví`: volitelný kód režimu záhlaví; výchozí hodnota je `0`
- `alpha`: hladina významnosti

### Kódy `ma_záhlaví`

| Kód | Význam |
| --- | --- |
| `0` | autodetekce popisků |
| `1` | horní řádek i levý sloupec jsou popisky |
| `2` | tabulka je čistě numerická bez popisků |

### Poznámky

- funkce vyžaduje alespoň tabulku `2 × 2`
- četnosti musí být nezáporná celá čísla
- v autodetekci se popisky použijí jen tehdy, když horní řádek i levý sloupec skutečně vypadají jako textové hlavičky

### Výstup

Spill výstup obsahuje:

- pozorovanou kontingenční tabulku včetně marginálních součtů
- tabulku očekávaných četností
- souhrn testu `n`, `df`, `α`, `χ²`, kritickou hodnotu a `p`
- míry asociace `Pearson C`, `Cramér V` a `phi` pro tabulku `2x2`

### Příklad

```excel
=CONTINGENCY.T(A1:C3;1;0,05)
```

## `CONTINGENCY.G`

Analyzuje kontingenční tabulku sestavenou z groupovaných dat.

### Syntaxe

```excel
=CONTINGENCY.G(sloupce; řádky; [počet]; [alpha]; [ma_záhlaví])
```

### Argumenty

- `sloupce`: kategorie budoucích sloupců kontingenční tabulky
- `řádky`: kategorie budoucích řádků kontingenční tabulky
- `počet`: volitelné četnosti dvojic; pokud chybí, každá dvojice má váhu `1`
- `alpha`: hladina významnosti
- `ma_záhlaví`: volitelný kód režimu záhlaví; výchozí hodnota je `0`

### Kódy `ma_záhlaví`

| Kód | Význam |
| --- | --- |
| `0` | autodetekce záhlaví |
| `1` | první řádek je záhlaví |
| `2` | vstup je bez záhlaví |

### Poznámky

- funkce vyžaduje alespoň dvě různé řádkové i sloupcové kategorie
- prázdné řádky se přeskočí
- pokud je `počet` zadán, musí obsahovat nezáporné celé četnosti
- výstup má stejnou strukturu jako `CONTINGENCY.T`

### Výstup

Spill výstup obsahuje:

- pozorovanou kontingenční tabulku
- očekávané četnosti
- souhrn testu `χ²`
- míry asociace `Pearson C`, `Cramér V` a případně `phi`

### Příklady

```excel
=CONTINGENCY.G(A2:A100;B2:B100)
=CONTINGENCY.G(A2:A100;B2:B100;C2:C100;0,05;1)
```
