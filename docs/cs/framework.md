# Přehled Frameworku

## Účel

XLStatUDF je Excel add-in se statistickými uživatelskými funkcemi napsanými v C# nad `Excel-DNA`.

## Technologie

- cílový runtime: `.NET 8`
- integrace s Excelem: `Excel-DNA`
- numerická knihovna: `MathNet.Numerics`
- výstup pro Excel: zabalený `.xll` doplněk pro 64bit Excel

## Hlavní Build

Pro běžné použití v Excelu je určený jediný soubor:

- [`artifacts/main/publish/XLStatUDF-AddIn64-packed.xll`](/c:/Users/pavel/Documents/github/XLStatUDF/artifacts/main/publish/XLStatUDF-AddIn64-packed.xll)

## Poznámky K Použití

- funkce jsou v kategorii `XLStatUDF`
- nápověda argumentů při psaní je zajištěna přes `ExcelDna.IntelliSense`
- většina funkcí ignoruje prázdné buňky v numerických vstupech
- relevantní funkce podporují poslední volitelný argument `ma_zahlavi`

## Režim `ma_zahlavi`

| Kód | Význam |
| --- | --- |
| `0` | autodetekce záhlaví |
| `1` | vstup obsahuje záhlaví |
| `2` | vstup záhlaví neobsahuje |

## Směr Testu

Funkce se směrem alternativní hypotézy používají tyto kódy:

| Kód | Význam |
| --- | --- |
| `0` | oboustranný test |
| `1` | levostranný test |
| `2` | pravostranný test |

## Chybové Stavy

| Chyba | Význam |
| --- | --- |
| `#HODNOTA!` | neplatný typ argumentu nebo nepodporovaná volba |
| `#ČÍSLO!` | neplatný číselný rozsah nebo omezení rozdělení |
| `#POČET!` | nedostatek pozorování |
| `#DÉLKA!` | rozsahy mají různou délku |
| `#DĚLENÍ0!` | dělení nulou v odvozené statistice |
| `#NENÍ_K_DISP!` | filtrování odstranilo všechny hodnoty |

## Hladina Významnosti

Funkce, které přijímají `alpha`, ji současně vracejí i ve spill výstupu jako `α`, aby byl výsledek samopopisný.
