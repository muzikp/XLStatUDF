# Functions Index

## General

- [GENERATE.NORM](./distributions.md#generatenorm): Generates a single random value from a normal distribution with optional perturbation.
- [GENERATE.INT](./distributions.md#generateint): Generates a single random integer from a specified interval with optional perturbation.
- [FILL](./distributions.md#fill): Repeats values or text formulas into a single-column spill output.
- [FILL.RANDOM](./distributions.md#fillrandom): Builds a sequence like `FILL` and then shuffles it randomly.

## Descriptive

- [AVERAGE.W](./weighted-means.md#averagew): Computes the weighted arithmetic mean.
- [HARMEAN.W](./weighted-means.md#harmeanw): Computes the weighted harmonic mean.
- [GEOMEAN.W](./weighted-means.md#geomeanw): Computes the weighted geometric mean.
- [VAR.P.W](./weighted-variance.md#varpw): Computes the weighted population variance.
- [VAR.S.W](./weighted-variance.md#varsw): Computes the weighted sample variance.
- [STDEV.P.W](./weighted-variance.md#stdevpw): Computes the weighted population standard deviation.
- [STDEV.S.W](./weighted-variance.md#stdevsw): Computes the weighted sample standard deviation.
- [VARCOEF](./variation-coefficients.md#varcoef): Computes the population coefficient of variation.
- [VARCOEF.S](./variation-coefficients.md#varcoefs): Computes the sample coefficient of variation.
- [VARCOEF.W](./variation-coefficients.md#varcoefw): Computes the weighted population coefficient of variation.
- [VARCOEF.S.W](./variation-coefficients.md#varcoefsw): Computes the weighted sample coefficient of variation.
- [PERCENTILE.INC.IFS](./percentiles.md#percentileincifs): Computes the inclusive percentile after applying filters.
- [PERCENTILE.EXC.IFS](./percentiles.md#percentileexcifs): Computes the exclusive percentile after applying filters.
- [PIVOT.*](./pivot.md#pivot): Builds a statistical pivot and computes exactly one selected metric per function.

## Tests

- [NORM.DIST.RANGE](./distributions.md#normdistrange): Computes the probability of an interval under the normal distribution.
- [SHAPIRO.WILK](./normality.md#shapirowilk): Performs the Shapiro-Wilk normality test.
- [KOLMOGOROV.SMIRNOV](./normality.md#kolmogorovsmirnov): Performs the Kolmogorov-Smirnov goodness-of-fit test for the selected distribution.
- [T.TEST.1S](./one-sample-tests.md#ttest1s): Performs a one-sample t-test against a specified hypothetical mean.
- [PROP.TEST.1S](./one-sample-tests.md#proptest1s): Performs a one-sample proportion test.
- [WILCOXON.PAIRED](./one-sample-tests.md#wilcoxonpaired): Performs the Wilcoxon signed-rank test for dependent measurements.
- [WELCH.TEST.2S.G](./two-sample-tests.md#welchtest2sg): Performs Welch's two-sample t-test for two independent groups.
- [MANN.WHITNEY.G](./two-sample-tests.md#mannwhitneyg): Performs the Mann-Whitney test for two independent groups.
- [CHISQ.GOF](./goodness-of-fit.md#chisqgof): Performs the chi-square goodness-of-fit test.
- [ANOVA.G](./anova.md#anovag): Performs one-way ANOVA on grouped data.
- [ANOVA.RM](./anova.md#anovarm): Performs repeated-measures ANOVA across columns.
- [ANCOVA.G](./ancova.md#ancovag): Performs ANCOVA with one factor and one or more covariates.
- [CONTINGENCY.T](./contingency.md#contingencyt): Analyzes a contingency table entered directly as a frequency matrix.
- [CONTINGENCY.G](./contingency.md#contingencyg): Analyzes a contingency table built from grouped columns.
- [CORREL.SPEARMAN](./correlation.md#correlspearman): Computes Spearman correlation and its significance test.
- [CORREL.MATRIX](./correlation.md#correlmatrix): Builds a correlation matrix including p-values and significance markers.
