# Index Funkcí

## Obecné

- [GENERATE.NORM](./distributions.md#generatenorm): Generuje jednu náhodnou hodnotu z normálního rozdělení s volitelnou perturbací.
- [GENERATE.INT](./distributions.md#generateint): Generuje jedno náhodné celé číslo ze zadaného intervalu s volitelnou perturbací.
- [FILL](./distributions.md#fill): Opakuje hodnoty nebo textové vzorce do jednosloupcového spill výstupu.
- [FILL.RANDOM](./distributions.md#fillrandom): Sestavuje řadu jako `FILL` a následně ji náhodně promíchá.

## Popisné

- [AVERAGE.W](./weighted-means.md#averagew): Počítá vážený aritmetický průměr.
- [HARMEAN.W](./weighted-means.md#harmeanw): Počítá vážený harmonický průměr.
- [GEOMEAN.W](./weighted-means.md#geomeanw): Počítá vážený geometrický průměr.
- [VAR.P.W](./weighted-variance.md#varpw): Počítá vážený populační rozptyl.
- [VAR.S.W](./weighted-variance.md#varsw): Počítá vážený výběrový rozptyl.
- [STDEV.P.W](./weighted-variance.md#stdevpw): Počítá váženou populační směrodatnou odchylku.
- [STDEV.S.W](./weighted-variance.md#stdevsw): Počítá váženou výběrovou směrodatnou odchylku.
- [VARCOEF](./variation-coefficients.md#varcoef): Počítá populační variační koeficient.
- [VARCOEF.S](./variation-coefficients.md#varcoefs): Počítá výběrový variační koeficient.
- [VARCOEF.W](./variation-coefficients.md#varcoefw): Počítá vážený populační variační koeficient.
- [VARCOEF.S.W](./variation-coefficients.md#varcoefsw): Počítá vážený výběrový variační koeficient.
- [PERCENTILE.INC.IFS](./percentiles.md#percentileincifs): Počítá inkluzivní percentil po aplikaci filtrů.
- [PERCENTILE.EXC.IFS](./percentiles.md#percentileexcifs): Počítá exkluzivní percentil po aplikaci filtrů.
- [PIVOT.*](./pivot.md#pivot): Sestavuje statistický pivot a v každé funkci počítá právě jeden zvolený ukazatel.

## Testy

- [NORM.DIST.RANGE](./distributions.md#normdistrange): Počítá pravděpodobnost intervalu normálního rozdělení.
- [SHAPIRO.WILK](./normality.md#shapirowilk): Provádí Shapiro-Wilkův test normality.
- [KOLMOGOROV.SMIRNOV](./normality.md#kolmogorovsmirnov): Provádí Kolmogorov-Smirnovův test dobré shody pro zvolené rozdělení.
- [T.TEST.1S](./one-sample-tests.md#ttest1s): Provádí jednovýběrový t-test vůči zadané hypotetické střední hodnotě.
- [PROP.TEST.1S](./one-sample-tests.md#proptest1s): Provádí jednovýběrový test podílu.
- [WILCOXON.PAIRED](./one-sample-tests.md#wilcoxonpaired): Provádí Wilcoxonův párový znaménkový test pro závislá měření.
- [WELCH.TEST.2S.G](./two-sample-tests.md#welchtest2sg): Provádí Welchův dvouvýběrový t-test pro dvě nezávislé skupiny.
- [MANN.WHITNEY.G](./two-sample-tests.md#mannwhitneyg): Provádí Mann-Whitneyho test pro dvě nezávislé skupiny.
- [CHISQ.GOF](./goodness-of-fit.md#chisqgof): Provádí chí-kvadrát test dobré shody.
- [ANOVA.G](./anova.md#anovag): Provádí jednofaktorovou ANOVA nad groupovanými daty.
- [ANOVA.RM](./anova.md#anovarm): Provádí ANOVA s opakovaným měřením nad sloupci.
- [ANCOVA.G](./ancova.md#ancovag): Provádí ANCOVA s jedním faktorem a jednou nebo více kovariátami.
- [CONTINGENCY.T](./contingency.md#contingencyt): Analyzuje kontingenční tabulku zadanou přímo jako matici četností.
- [CONTINGENCY.G](./contingency.md#contingencyg): Analyzuje kontingenční tabulku sestavenou z groupovaných sloupců.
- [CORREL.SPEARMAN](./correlation.md#correlspearman): Počítá Spearmanův korelační koeficient a test jeho významnosti.
- [CORREL.MATRIX](./correlation.md#correlmatrix): Sestavuje korelační matici včetně p-hodnot a značek signifikance.
