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

## `FILL`

Vytvoří jednosloupcový spill opakováním jedné hodnoty nebo opakovaným vyhodnocením vzorce.

### Syntaxe

```excel
=FILL(co; pocet)
```

### Argumenty

- `co`: skalární hodnota, která se má opakovat, nebo textový vzorec začínající `=`
- `pocet`: počet vrácených řádků; celé číslo `>= 1`

### Výstup

Jednosloupcový spill rozsah s `pocet` hodnotami.

### Poznámky

- pokud zadáš běžnou hodnotu, funkce ji pouze zkopíruje do všech řádků
- pokud zadáš textový vzorec začínající `=`, funkce ho vyhodnotí zvlášť pro každý řádek
- to je praktické například pro náhodné generátory; přímý argument `RANDBETWEEN(...)` by se jinak do UDF předal už jen jako jedna spočtená hodnota

### Příklady

```excel
=FILL("A";5)
=FILL(123;4)
=FILL("=RANDBETWEEN(1;10)";20)
```
