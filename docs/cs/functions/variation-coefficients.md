# Variační Koeficienty

## `VARCOEF`

Počítá populační variační koeficient.

### Syntaxe

```excel
=VARCOEF(values)
```

### Poznámky

- prázdné buňky se ignorují
- je potřeba alespoň jedna platná hodnota
- pokud je průměr roven nule, funkce vrátí chybu dělení nulou

### Výstup

Skalární `σ / μ`.

## `VARCOEF.S`

Počítá výběrový variační koeficient.

### Syntaxe

```excel
=VARCOEF.S(values)
```

### Poznámky

- prázdné buňky se ignorují
- jsou potřeba alespoň dvě platné hodnoty
- pokud je průměr roven nule, funkce vrátí chybu dělení nulou

### Výstup

Skalární `sₓ / x̄`.

## `VARCOEF.W`

Počítá vážený populační variační koeficient.

### Syntaxe

```excel
=VARCOEF.W(values; weights)
```

### Poznámky

- `values` a `weights` musí mít stejnou délku
- váhy musí být nezáporné
- prázdné váhy se berou jako `0`
- pokud je vážený průměr roven nule, funkce vrátí chybu dělení nulou

### Výstup

Skalární `σ_w / x̄_w`.

## `VARCOEF.S.W`

Počítá vážený výběrový variační koeficient.

### Syntaxe

```excel
=VARCOEF.S.W(values; weights)
```

### Poznámky

- `values` a `weights` musí mít stejnou délku
- váhy musí být nezáporné
- prázdné váhy se berou jako `0`
- pokud `Σw <= 1`, funkce vrátí chybu `#POČET!`
- pokud je vážený průměr roven nule, funkce vrátí chybu dělení nulou

### Výstup

Skalární `sₓ,w / x̄_w`.
