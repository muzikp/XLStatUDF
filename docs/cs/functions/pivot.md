# PIVOT

## `PIVOT.*`

Rodina funkcí `PIVOT.*` sestavuje statistický pivot z řádkových a sloupcových kategorií a pro každou funkci počítá právě jeden ukazatel.

### Syntaxe

```excel
=PIVOT.SUM(řádky; sloupce; hodnoty)
=PIVOT.AVERAGE(řádky; sloupce; hodnoty)
=PIVOT.PERCENTILE(řádky; sloupce; hodnoty; kvantil)
=PIVOT.CONF.T(řádky; sloupce; hodnoty; alfa; [smer])
```

### Argumenty

- `řádky`: jeden nebo více sloupců kategorií pro řádky, vždy včetně záhlaví
- `sloupce`: volitelně jeden nebo více sloupců kategorií pro sloupce, vždy včetně záhlaví; může být prázdné
- `hodnoty`: jeden sloupec hodnot včetně záhlaví
- `kvantil`: hledaný percentil v intervalu `(0;1)`
- `alfa`: hladina alfa v intervalu `(0;1)`
- `smer`: volitelný směr pro `CONF.*`; `0 = oboustranný`, `-1 = levostranný`, `1 = pravostranný`

### Dostupné funkce

| Funkce | Popis | Detail |
| --- | --- | --- |
| `PIVOT.COUNT(řádky; sloupce; hodnoty)` | Počet neprázdných hodnot | bez detailu |
| `PIVOT.SUM(řádky; sloupce; hodnoty)` | Součet | bez detailu |
| `PIVOT.AVERAGE(řádky; sloupce; hodnoty)` | Aritmetický průměr | bez detailu |
| `PIVOT.MIN(řádky; sloupce; hodnoty)` | Minimum | bez detailu |
| `PIVOT.MAX(řádky; sloupce; hodnoty)` | Maximum | bez detailu |
| `PIVOT.MEDIAN(řádky; sloupce; hodnoty)` | Medián | bez detailu |
| `PIVOT.PERCENTILE(řádky; sloupce; hodnoty; kvantil)` | Percentil | `0 < kvantil < 1` |
| `PIVOT.STDEV.S(řádky; sloupce; hodnoty)` | Výběrová směrodatná odchylka | bez detailu |
| `PIVOT.STDEV.P(řádky; sloupce; hodnoty)` | Populační směrodatná odchylka | bez detailu |
| `PIVOT.VAR.S(řádky; sloupce; hodnoty)` | Výběrový rozptyl | bez detailu |
| `PIVOT.VAR.P(řádky; sloupce; hodnoty)` | Populační rozptyl | bez detailu |
| `PIVOT.VARCOEF.S(řádky; sloupce; hodnoty)` | Výběrový variační koeficient | bez detailu |
| `PIVOT.VARCOEF.P(řádky; sloupce; hodnoty)` | Populační variační koeficient | bez detailu |
| `PIVOT.CONF.T(řádky; sloupce; hodnoty; alfa; [smer])` | Poloviční šířka intervalu spolehlivosti pro t-rozdělení | `0 < alfa < 1`; `smer`: `0`, `-1`, `1` |
| `PIVOT.CONF.NORM(řádky; sloupce; hodnoty; alfa; [smer])` | Poloviční šířka intervalu spolehlivosti pro normální aproximaci | `0 < alfa < 1`; `smer`: `0`, `-1`, `1` |
| `PIVOT.MAD(řádky; sloupce; hodnoty)` | Medián absolutních odchylek od mediánu | bez detailu |
| `PIVOT.IQR(řádky; sloupce; hodnoty)` | Mezikvartilové rozpětí | bez detailu |

### Validace dat

- `řádky`, `sloupce` i `hodnoty` se očekávají se záhlavím v prvním řádku
- `count` může pracovat i s nenumerickými hodnotami; počítá neprázdné buňky
- všechny ostatní funkce vyžadují kvantitativní data
- `varcoef.s` a `varcoef.p` vyžadují nenulový průměr
- `stdev.s`, `var.s`, `conf.t` a `conf.norm` vyžadují alespoň dva platné numerické záznamy
- nekompatibilní kombinace typu dat a výpočtu vracejí chybu hodnoty
- `conf.*` akceptují jen `smer` z množiny `{-1, 0, 1}`

### Poznámky

- `sloupce` může být prázdné; v tom případě vznikne jednorozměrný pivot s jediným sloupcovým blokem `Celkem`
- řádkové i sloupcové kategorie se ve výstupu řadí abecedně
- řádky, pro které jsou všechny výsledné hodnoty prázdné, se z výstupu vynechají
- `conf.t` a `conf.norm` vracejí poloviční šířku intervalu, ne obě meze zvlášť
- při `smer = 0` se používá kritická hodnota `1 - alfa/2`
- při `smer = -1` nebo `1` se používá kritická hodnota `1 - alfa`

### Výstup

Víceřádkový spill výstup ve stylu pivotu:

- vlevo jsou sloupce řádkových kategorií
- nahoře jsou úrovně sloupcových kategorií
- poslední řádek záhlaví nad numerickými sloupci obsahuje názvy řádkových proměnných

### Příklady

```excel
=PIVOT.SUM(E:E;F:F;G:G)
=PIVOT.AVERAGE(E:E;F:F;G:G)
=PIVOT.PERCENTILE(E:E;F:F;G:G;0,9)
=PIVOT.CONF.T(E:E;F:F;G:G;0,05)
=PIVOT.CONF.T(E:E;F:F;G:G;0,05;1)
```
