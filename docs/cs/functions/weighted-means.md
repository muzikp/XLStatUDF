# Vážené Průměry

## `AVERAGE.W`

Počítá vážený aritmetický průměr.

### Syntaxe

```excel
=AVERAGE.W(values; weights)
```

### Argumenty

- `values`: číselná pozorování
- `weights`: nezáporné váhy; prázdné buňky vah se berou jako `0`

### Poznámky

- rozsahy `values` a `weights` musí mít stejnou délku
- prázdné buňky ve `values` se přeskočí; odpovídající váha se tím také vynechá
- pokud jsou všechny váhy nulové nebo některá váha záporná, funkce vrátí numerickou chybu

### Výstup

Skalární vážený aritmetický průměr.

### Příklad

```excel
=AVERAGE.W(A2:A10;B2:B10)
```

## `HARMEAN.W`

Počítá vážený harmonický průměr.

### Syntaxe

```excel
=HARMEAN.W(values; weights)
```

### Argumenty

- `values`: kladná číselná pozorování
- `weights`: nezáporné váhy; prázdné buňky vah se berou jako `0`

### Poznámky

- rozsahy `values` a `weights` musí mít stejnou délku
- hodnoty s nulovou vahou do výsledku nevstupují
- pokud má některá hodnota s kladnou vahou hodnotu `<= 0`, funkce vrátí numerickou chybu

### Výstup

Skalární vážený harmonický průměr.

## `GEOMEAN.W`

Počítá vážený geometrický průměr.

### Syntaxe

```excel
=GEOMEAN.W(values; weights)
```

### Argumenty

- `values`: kladná číselná pozorování
- `weights`: nezáporné váhy; prázdné buňky vah se berou jako `0`

### Poznámky

- rozsahy `values` a `weights` musí mít stejnou délku
- hodnoty s nulovou vahou do výsledku nevstupují
- pokud má některá hodnota s kladnou vahou hodnotu `<= 0`, funkce vrátí numerickou chybu

### Výstup

Skalární vážený geometrický průměr.
