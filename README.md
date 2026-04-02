# XLStatUDF

Excel-DNA add-in se sadou statistických UDF funkcí pro Excel.

## Aktuální stav

Implementované a otestované funkce:

- `NORM.DIST.RANGE`
- `GENERATE.NORM`
- `GENERATE.INT`
- `FILL`
- `FILL.RANDOM`
- `AVERAGE.W`
- `HARMEAN.W`
- `GEOMEAN.W`
- `VAR.P.W`
- `VAR.S.W`
- `STDEV.P.W`
- `STDEV.S.W`
- `VARCOEF`
- `VARCOEF.S`
- `VARCOEF.S.W`
- `PERCENTILE.INC.IFS`
- `PERCENTILE.EXC.IFS`
- `SHAPIRO.WILK`
- `KOLMOGOROV.SMIRNOV`
- `T.TEST.1S`
- `PROP.TEST.1S`
- `CORREL.SPEARMAN`
- `WELCH.TEST.2S.G`
- `CHISQ.GOF`
- `ANOVA.G`
- `ANOVA.RM`
- `ANCOVA.G`
- `CONTINGENCY.T`
- `CONTINGENCY.G`

## Build

```powershell
.\build.ps1
```

Jediný hlavní build pro Excel je vždy:

`artifacts\main\publish\XLStatUDF-AddIn64-packed.xll`

Stejný soubor se po buildu kopíruje i do:

`installer\XLStatUDF-packed.xll`

Instalátory se po buildu generují sem:

- `artifacts\installer\XLStatUDF_CS_Setup.exe`
- `artifacts\installer\XLStatUDF_EN_Setup.exe`
