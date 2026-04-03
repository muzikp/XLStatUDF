# Percentily

## `PERCENTILE.INC.IFS`

Počítá inkluzivní percentil s filtrováním ve stylu `SUMIFS`.

### Syntaxe

```excel
=PERCENTILE.INC.IFS(values; quantile; [criteria_range_1; criteria_1; ...])
```

### Argumenty

- `values`: číselná data
- `quantile`: hodnota od `0` do `1`
- `criteria_range_n`: rozsah, podle kterého se filtruje
- `criteria_n`: přesná shoda, relační výraz nebo wildcard výraz

### Poznámky

- počet argumentů filtru musí být sudý po dvojicích `rozsah + kritérium`
- pokud po filtrování nezůstane žádná hodnota, funkce vrátí `#N/A`
- percentil se počítá inkluzivní metodou stejně jako `PERCENTILE.INC`

### Výstup

Skalární percentil.

### Příklad

```excel
=PERCENTILE.INC.IFS(A2:A100;0,75;B2:B100;"A";A2:A100;">10")
```

## `PERCENTILE.EXC.IFS`

Počítá exkluzivní percentil s filtrováním ve stylu `SUMIFS`.

### Syntaxe

```excel
=PERCENTILE.EXC.IFS(values; quantile; [criteria_range_1; criteria_1; ...])
```

### Argumenty

Stejné jako u `PERCENTILE.INC.IFS`, ale `quantile` musí být přísně mezi `0` a `1`.

### Poznámky

- počet argumentů filtru musí být sudý po dvojicích `rozsah + kritérium`
- pokud po filtrování nezůstane žádná hodnota, funkce vrátí `#N/A`
- pokud je pro zadaný kvantil exkluzivní percentil mimo definiční obor vzorce, funkce vrátí numerickou chybu

### Výstup

Skalární percentil.
