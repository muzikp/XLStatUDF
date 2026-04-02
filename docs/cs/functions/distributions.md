# Rozdělení

## `NORM.DIST.RANGE`

Vrací pravděpodobnost, že náhodná veličina s normálním rozdělením padne do zadaného intervalu.

### Syntaxe

```excel
=NORM.DIST.RANGE(mean; standard_deviation; lower_bound; upper_bound)
```

### Argumenty

- `mean`: střední hodnota rozdělení
- `standard_deviation`: kladná směrodatná odchylka
- `lower_bound`: dolní mez intervalu; prázdná buňka znamená mínus nekonečno
- `upper_bound`: horní mez intervalu; prázdná buňka znamená plus nekonečno

### Výstup

Skalární hodnota z intervalu `[0;1]`.

### Příklad

```excel
=NORM.DIST.RANGE(0;1;-1;1)
```

## `GENERATE.NORM`

Vygeneruje spill sloupec náhodných hodnot z normálního rozdělení.

### Syntaxe

```excel
=GENERATE.NORM(mean; stdev; count)
```

### Argumenty

- `mean`: požadovaný průměr
- `stdev`: kladná směrodatná odchylka
- `count`: počet generovaných hodnot; celé číslo `>= 1`

### Výstup

Jednosloupcový spill rozsah s `count` hodnotami.

### Poznámky

- funkce je volatilní, takže po přepočtu vygeneruje nový vzorek

### Příklad

```excel
=GENERATE.NORM(100;15;20)
```

## `GENERATE.INT`

Vygeneruje spill sloupec náhodných celých čísel ze zadaného intervalu.

### Syntaxe

```excel
=GENERATE.INT([pocet]; [minimum]; [maximum])
```

### Argumenty

- `pocet`: volitelný počet generovaných hodnot; celé číslo `>= 1`; výchozí hodnota je `1`
- `minimum`: volitelná dolní mez intervalu; výchozí hodnota je `-2147483648`
- `maximum`: volitelná horní mez intervalu; výchozí hodnota je `2147483647`

### Výstup

Jednosloupcový spill rozsah s náhodnými celými čísly z uzavřeného intervalu `<minimum; maximum>`.

### Poznámky

- funkce je volatilní, takže po přepočtu vygeneruje novou řadu
- pokud `minimum > maximum`, funkce vrátí číselnou chybu
- pokud hranice nezadáš, použije se praktický strop přes celé 32bitové rozmezí

### Příklady

```excel
=GENERATE.INT()
=GENERATE.INT(10)
=GENERATE.INT(20;1;6)
```

## `FILL`

Vytvoří jednosloupcový spill opakováním jedné hodnoty nebo opakovaným vyhodnocením vzorce.

### Syntaxe

```excel
=FILL(co; pocet; [co2]; [pocet2]; ...)
```

### Argumenty

- `co`: skalární hodnota, která se má opakovat, nebo textový vzorec začínající `=`
- `pocet`: počet vrácených řádků; celé číslo `>= 1`
- další argumenty se zadávají po dvojicích `co + pocet`

### Výstup

Jednosloupcový spill rozsah s `pocet` hodnotami.

### Poznámky

- pokud zadáš běžnou hodnotu, funkce ji pouze zkopíruje do všech řádků
- pokud zadáš textový vzorec začínající `=`, funkce ho vyhodnotí zvlášť pro každý řádek
- můžeš zadat více dvojic `co + pocet`; funkce jednotlivé bloky spojí pod sebe v pořadí zadání
- to je praktické například pro náhodné generátory; přímý argument `RANDBETWEEN(...)` by se jinak do UDF předal už jen jako jedna spočtená hodnota

### Příklady

```excel
=FILL("A";5)
=FILL(123;4)
=FILL("muž";100;"žena";100;"dítě";90)
=FILL("=RANDBETWEEN(1;10)";20)
```

## `FILL.RANDOM`

Vytvoří jednosloupcový spill stejně jako `FILL`, ale výslednou řadu ještě před vrácením náhodně promíchá.

### Syntaxe

```excel
=FILL.RANDOM(co; pocet; [co2]; [pocet2]; ...)
```

### Argumenty

- `co`: skalární hodnota, která se má opakovat, nebo textový vzorec začínající `=`
- `pocet`: počet vrácených řádků; celé číslo `>= 1`
- další argumenty se zadávají po dvojicích `co + pocet`

### Výstup

Jednosloupcový spill rozsah se všemi vygenerovanými hodnotami v náhodném pořadí.

### Poznámky

- funkce nejprve sestaví celou řadu stejně jako `FILL`
- potom ji náhodně promíchá
- hodí se například pro generování promíchaných kategorií nebo náhodného pořadí stimulů

### Příklady

```excel
=FILL.RANDOM("muž";100;"žena";100;"dítě";90)
=FILL.RANDOM("A";5;"B";5)
```
