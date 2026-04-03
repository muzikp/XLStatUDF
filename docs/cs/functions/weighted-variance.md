# Vážený Rozptyl A Směrodatná Odchylka

Společné chování pro všechny čtyři funkce:

- `values` a `weights` musí mít stejnou délku
- váhy musí být nezáporné
- prázdné buňky ve vahách se berou jako `0`
- pokud je součet vah nulový, funkce vrátí numerickou chybu

## `VAR.P.W`

Počítá vážený populační rozptyl.

### Syntaxe

```excel
=VAR.P.W(values; weights)
```

### Výstup

Skalární vážený populační rozptyl.

### Poznámky

- používá jmenovatel `Σw`

## `VAR.S.W`

Počítá vážený výběrový rozptyl.

### Syntaxe

```excel
=VAR.S.W(values; weights)
```

### Výstup

Skalární vážený výběrový rozptyl.

### Poznámky

- používá jmenovatel `Σw - 1`
- pokud `Σw <= 1`, funkce vrátí chybu `#POČET!`

## `STDEV.P.W`

Počítá váženou populační směrodatnou odchylku.

### Syntaxe

```excel
=STDEV.P.W(values; weights)
```

### Výstup

Skalární vážená populační směrodatná odchylka `σ`.

### Poznámky

- je definována jako druhá odmocnina z `VAR.P.W`

## `STDEV.S.W`

Počítá váženou výběrovou směrodatnou odchylku.

### Syntaxe

```excel
=STDEV.S.W(values; weights)
```

### Výstup

Skalární vážená výběrová směrodatná odchylka `sₓ`.

### Poznámky

- je definována jako druhá odmocnina z `VAR.S.W`
