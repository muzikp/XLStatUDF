# Test Dobrého Souladu

## `CHISQ.GOF`

Provádí chí-kvadrát test dobré shody.

### Syntaxe

```excel
=CHISQ.GOF(observed; expected; [categories]; [alpha]; [ma_záhlaví])
```

### Argumenty

- `observed`: pozorované četnosti kategorií
- `expected`: očekávané četnosti nebo pravděpodobnosti
- `categories`: volitelné názvy kategorií
- `alpha`: hladina významnosti
- `ma_záhlaví`: volitelný kód režimu záhlaví; výchozí hodnota je `0`

### Kódy `ma_záhlaví`

| Kód | Význam |
| --- | --- |
| `0` | autodetekce záhlaví |
| `1` | první řádek je záhlaví |
| `2` | vstup je bez záhlaví |

### Poznámky

- `observed` musí obsahovat nezáporné celé četnosti
- `expected` může být zadáno jako četnosti nebo jako pravděpodobnosti se součtem `1`
- pokud `expected` tvoří pravděpodobnosti, funkce je automaticky přepočte na očekávané četnosti podle velikosti vzorku
- pokud `categories` chybí, kategorie se očíslují `1, 2, 3, ...`

### Výstup

Spill výstup obsahuje:

- souhrn testu `χ²`, `df`, `α`, kritickou hodnotu a `p`
- tabulku kategorií s `O`, `E` a příspěvkem `(O−E)² / E`
