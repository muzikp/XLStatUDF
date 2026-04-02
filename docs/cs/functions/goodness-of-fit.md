# Test Dobré Shody

## `CHISQ.GOF`

Chí-kvadrát test dobré shody.

### Syntaxe

```excel
=CHISQ.GOF(observed; expected; [categories]; [alpha]; [ma_zahlavi])
```

### Argumenty

- `observed`: pozorované četnosti kategorií
- `expected`: očekávané četnosti nebo pravděpodobnosti
- `categories`: volitelné názvy kategorií
- `alpha`: hladina významnosti
- `ma_zahlavi`: volitelný kód režimu záhlaví; výchozí hodnota je `0`

### Kódy `ma_zahlavi`

| Kód | Význam |
| --- | --- |
| `0` | autodetekce záhlaví |
| `1` | první řádek je záhlaví |
| `2` | vstup je bez záhlaví |

### Výstup

Spill výstup obsahuje:

- souhrn testu `χ²`, `df`, `α`, kritickou hodnotu a `p`
- tabulku kategorií s `O`, `E` a příspěvkem `(O−E)² / E`
