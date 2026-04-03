# XLStatUDF

Excel-DNA doplněk se sadou statistických uživatelských funkcí pro Microsoft Excel.

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
- `VARCOEF.W`
- `VARCOEF.S.W`
- `PERCENTILE.INC.IFS`
- `PERCENTILE.EXC.IFS`
- `SHAPIRO.WILK`
- `KOLMOGOROV.SMIRNOV`
- `T.TEST.1S`
- `PROP.TEST.1S`
- `CORREL.SPEARMAN`
- `CORREL.MATRIX`
- `WELCH.TEST.2S.G`
- `MANN.WHITNEY.G`
- `WILCOXON.PAIRED`
- `CHISQ.GOF`
- `CONTINGENCY.T`
- `CONTINGENCY.G`
- `ANOVA.G`
- `ANOVA.RM`
- `ANCOVA.G`
- `PIVOT.*`

## Build

```powershell
.\build.ps1
```

Hlavní build pro testování v Excelu:

`artifacts\main\publish\XLStatUDF-AddIn64-packed.xll`

Po buildu se generují také oba instalační soubory:

- `artifacts\installer\XLStatUDF_CS_Setup.exe`
- `artifacts\installer\XLStatUDF_EN_Setup.exe`

## Dokumentace

Dokumentace je rozdělená podle jazyka:

- česká: [`docs/cs/README.md`](/c:/Users/pavel/Documents/github/XLStatUDF/docs/cs/README.md)
- anglická: [`docs/en/README.md`](/c:/Users/pavel/Documents/github/XLStatUDF/docs/en/README.md)
