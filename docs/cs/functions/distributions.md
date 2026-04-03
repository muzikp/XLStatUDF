# Rozdělení

## `NORM.DIST.RANGE`

Počítá pravděpodobnost, že náhodná veličina s normálním rozdělením padne do zadaného intervalu.

### Syntaxe

```excel
=NORM.DIST.RANGE(mean; standard_deviation; lower_bound; upper_bound)
```

### Argumenty

- `mean`: střední hodnota rozdělení
- `standard_deviation`: kladná směrodatná odchylka
- `lower_bound`: dolní mez intervalu; prázdná buňka znamená mínus nekonečno
- `upper_bound`: horní mez intervalu; prázdná buňka znamená plus nekonečno

### Poznámky

- pokud `lower_bound > upper_bound`, funkce vrátí numerickou chybu
- pokud `standard_deviation <= 0`, funkce vrátí numerickou chybu
- prázdné meze lze použít pro jednostranné intervaly

### Výstup

Skalární hodnota z intervalu `[0;1]`.

### Příklad

```excel
=NORM.DIST.RANGE(0;1;-1;1)
```

## `GENERATE.NORM`

Generuje jednu náhodnou hodnotu z normálního rozdělení s volitelnou perturbací.

### Syntaxe

```excel
=GENERATE.NORM(stredni_hodnota; směrodatná_odchylka; [outlier_rate])
```

### Argumenty

- `stredni_hodnota`: požadovaný průměr
- `směrodatná_odchylka`: kladná směrodatná odchylka
- `outlier_rate`: volitelná pravděpodobnost dodatečné náhodné perturbace v intervalu `<0;1>`; výchozí hodnota je `0`

### Poznámky

- funkce je volatilní, takže po přepočtu vygeneruje nový vzorek
- při `outlier_rate = 0` se generuje běžná hodnota z `N(μ, σ)`
- pokud náhodně nastane perturbace, k vygenerované hodnotě se přičte dodatečný šum z `N(0, 3σ)`
- neplatný `outlier_rate` nebo nečíselné argumenty vrací chybu hodnoty

### Výstup

Jedna skalární hodnota.

### Příklad

```excel
=GENERATE.NORM(100;15)
=GENERATE.NORM(0;1;0,2)
```

## `GENERATE.INT`

Generuje jedno náhodné celé číslo ze zadaného intervalu s volitelnou perturbací.

### Syntaxe

```excel
=GENERATE.INT([minimum]; [maximum]; [outlier_rate])
```

### Argumenty

- `minimum`: volitelná dolní mez intervalu; výchozí hodnota je `-2147483648`
- `maximum`: volitelná horní mez intervalu; výchozí hodnota je `2147483647`
- `outlier_rate`: volitelná pravděpodobnost dodatečné náhodné perturbace v intervalu `<0;1>`; výchozí hodnota je `0`

### Poznámky

- funkce je volatilní, takže po přepočtu vygeneruje novou hodnotu
- pokud `minimum > maximum`, funkce vrátí numerickou chybu
- pokud hranice nezadáš, použije se celé praktické 32bitové rozmezí
- při `outlier_rate = 0` se generuje běžná hodnota z uzavřeného intervalu `<minimum; maximum>`
- pokud náhodně nastane perturbace, k vygenerované hodnotě se přičte dodatečný náhodný celočíselný posun z intervalu `⟨-(maximum-minimum); +(maximum-minimum)⟩`

### Výstup

Jedno celé číslo.

### Příklady

```excel
=GENERATE.INT()
=GENERATE.INT(1;6)
=GENERATE.INT(1;6;0,2)
```

## `FILL`

Opakuje jednu nebo více hodnot, případně opakovaně vyhodnocuje textový vzorec, do jednosloupcového spill výstupu.

### Syntaxe

```excel
=FILL(co; počet; [co2]; [počet2]; ...)
```

### Argumenty

- `co`: skalární hodnota, která se má opakovat, nebo textový vzorec začínající `=`
- `počet`: počet vrácených řádků; celé číslo `>= 1`
- další argumenty se zadávají po dvojicích `co + počet`

### Poznámky

- pokud zadáš běžnou hodnotu, funkce ji pouze zkopíruje do všech řádků
- pokud zadáš textový vzorec začínající `=`, funkce ho vyhodnotí zvlášť pro každý řádek
- počet dodatečných argumentů musí být sudý; jinak funkce vrátí chybu hodnoty
- přímý argument typu `GENERATE.NORM(...)` nebo `RANDBETWEEN(...)` Excel vyhodnotí ještě před vstupem do `FILL`, proto je pro opakované přepočítání nutný textový vzorec

### Výstup

Jednosloupcový spill rozsah s výslednou řadou.

### Příklady

```excel
=FILL("A";5)
=FILL(123;4)
=FILL("muž";100;"žena";100;"dítě";90)
=FILL("=RANDBETWEEN(1;10)";20)
=FILL("=GENERATE.NORM(0;1;0,2)";100)
```

## `FILL.RANDOM`

Sestaví řadu stejně jako `FILL`, ale před vrácením ji náhodně promíchá.

### Syntaxe

```excel
=FILL.RANDOM(co; počet; [co2]; [počet2]; ...)
```

### Argumenty

Stejné jako u `FILL`.

### Poznámky

- nejprve se vytvoří celá řada stejně jako u `FILL`
- teprve poté se hotová řada náhodně promíchá
- stejné validační podmínky jako u `FILL` platí i zde

### Výstup

Jednosloupcový spill rozsah se všemi vygenerovanými hodnotami v náhodném pořadí.

### Příklady

```excel
=FILL.RANDOM("muž";100;"žena";100;"dítě";90)
=FILL.RANDOM("A";5;"B";5)
```
