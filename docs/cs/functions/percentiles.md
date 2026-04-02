# Percentily

## `PERCENTILE.INC.IFS`

Inkluzivní percentil s volitelným filtrováním ve stylu `SUMIFS`.

### Syntaxe

```excel
=PERCENTILE.INC.IFS(values; quantile; [criteria_range_1; criteria_1; ...])
```

### Argumenty

- `values`: číselná data
- `quantile`: hodnota od `0` do `1`
- `criteria_range_n`: rozsah, podle kterého se filtruje
- `criteria_n`: přesná shoda, relační výraz nebo wildcard výraz

### Výstup

Skalární percentil.

### Příklad

```excel
=PERCENTILE.INC.IFS(A2:A100;0,75;B2:B100;"A";A2:A100;">10")
```

## `PERCENTILE.EXC.IFS`

Exkluzivní percentil s volitelným filtrováním ve stylu `SUMIFS`.

### Syntaxe

```excel
=PERCENTILE.EXC.IFS(values; quantile; [criteria_range_1; criteria_1; ...])
```

### Argumenty

Stejné jako u `PERCENTILE.INC.IFS`, ale `quantile` musí být přísně mezi `0` a `1`.

### Výstup

Skalární percentil.
