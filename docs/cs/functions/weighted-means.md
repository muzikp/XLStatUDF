# Vážené Průměry

## `AVERAGE.W`

Vážený aritmetický průměr.

### Syntaxe

```excel
=AVERAGE.W(values; weights)
```

### Argumenty

- `values`: číselná pozorování
- `weights`: nezáporné váhy; prázdné buňky vah se berou jako `0`

### Výstup

Skalární vážený průměr.

### Příklad

```excel
=AVERAGE.W(A2:A10;B2:B10)
```

## `HARMEAN.W`

Vážený harmonický průměr.

### Syntaxe

```excel
=HARMEAN.W(values; weights)
```

### Argumenty

- `values`: kladná číselná pozorování
- `weights`: nezáporné váhy

### Výstup

Skalární vážený harmonický průměr.

## `GEOMEAN.W`

Vážený geometrický průměr.

### Syntaxe

```excel
=GEOMEAN.W(values; weights)
```

### Argumenty

- `values`: kladná číselná pozorování
- `weights`: nezáporné váhy

### Výstup

Skalární vážený geometrický průměr.
