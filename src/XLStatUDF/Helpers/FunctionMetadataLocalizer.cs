namespace XLStatUDF.Helpers
{
    using System.Collections.Generic;
    using ExcelDna.Integration;
    using ExcelDna.Registration;

    internal static partial class FunctionMetadataLocalizer
    {
        private const string DocsRoot = "https://github.com/muzikp/XLStatUDF/blob/main/docs/";

        private static readonly IReadOnlyDictionary<string, FunctionMetadata> CzechMetadata = BuildCzechMetadata();
        private static readonly IReadOnlyDictionary<string, FunctionMetadata> EnglishMetadata = BuildEnglishMetadata();

        public static IEnumerable<ExcelFunctionRegistration> Apply(IEnumerable<ExcelFunctionRegistration> registrations)
        {
            var metadata = AddInLanguage.IsCzech ? CzechMetadata : EnglishMetadata;
            foreach (var registration in registrations)
            {
                var functionName = registration.FunctionAttribute?.Name;
                if (functionName is null || !metadata.TryGetValue(functionName, out var localized))
                {
                    yield return registration;
                    continue;
                }

                registration.FunctionAttribute!.Category = localized.Category;
                registration.FunctionAttribute.Description = localized.Description;
                registration.FunctionAttribute.HelpTopic = localized.HelpTopic;

                for (var i = 0; i < registration.ParameterRegistrations.Count && i < localized.Arguments.Length; i++)
                {
                    var argument = registration.ParameterRegistrations[i].ArgumentAttribute;
                    if (argument is null)
                    {
                        continue;
                    }

                    argument.Name = localized.Arguments[i].Name;
                    argument.Description = localized.Arguments[i].Description;
                }

                if (string.Equals(functionName, "CORREL.MATRIX", StringComparison.OrdinalIgnoreCase))
                {
                    ApplyCorrelMatrixArgumentOverrideClean(registration);
                }

                yield return registration;
            }
        }

        private static void ApplyCorrelMatrixArgumentOverride(ExcelFunctionRegistration registration)
        {
            var arguments = AddInLanguage.IsCzech
                ? new[]
                {
                    Arg("data", "VstupnĂ­ data; alespoĹ dva sloupce."),
                    Arg("metoda", "VolitelnÄ›: 0=Pearson, 1=Spearman."),
                    Arg("vĂ˝stup", "VolitelnÄ›: 0=koeficienty, 1=p-hodnoty, 2=koeficient+p, 3=koeficient+p+sig, 4=koeficient+sig v jednĂ© buĹce."),
                    Arg("p_minimum", "VolitelnÄ›: zobrazĂ­ jen vazby s p < zadanĂˇ hodnota."),
                    HeaderArg()
                }
                : new[]
                {
                    Arg("data", "Input data; at least two columns."),
                    Arg("method", "Optional: 0=Pearson, 1=Spearman."),
                    Arg("output", "Optional: 0=coefficients, 1=p-values, 2=coefficient+p, 3=coefficient+p+sig, 4=coefficient+sig in one cell."),
                    Arg("p_minimum", "Optional: show only relationships with p < specified value."),
                    HeaderArgEn()
                };

            for (var i = 0; i < registration.ParameterRegistrations.Count && i < arguments.Length; i++)
            {
                var argument = registration.ParameterRegistrations[i].ArgumentAttribute;
                if (argument is null)
                {
                    continue;
                }

                argument.Name = arguments[i].Name;
                argument.Description = arguments[i].Description;
            }
        }

        private static void ApplyCorrelMatrixArgumentOverrideClean(ExcelFunctionRegistration registration)
        {
            var arguments = AddInLanguage.IsCzech
                ? new[]
                {
                    Arg("data", "Vstupní data; alespoň dva sloupce."),
                    Arg("metoda", "Volitelně: 0=Pearson, 1=Spearman."),
                    Arg("výstup", "Volitelně: 0=koeficienty, 1=p-hodnoty, 2=koeficient+p, 3=koeficient+p+sig, 4=koeficient+sig v jedné buňce."),
                    Arg("p_minimum", "Volitelně: zobrazí jen vazby s p < zadaná hodnota."),
                    HeaderArg()
                }
                : new[]
                {
                    Arg("data", "Input data; at least two columns."),
                    Arg("method", "Optional: 0=Pearson, 1=Spearman."),
                    Arg("output", "Optional: 0=coefficients, 1=p-values, 2=coefficient+p, 3=coefficient+p+sig, 4=coefficient+sig in one cell."),
                    Arg("p_minimum", "Optional: show only relationships with p < specified value."),
                    HeaderArgEn()
                };

            for (var i = 0; i < registration.ParameterRegistrations.Count && i < arguments.Length; i++)
            {
                var argument = registration.ParameterRegistrations[i].ArgumentAttribute;
                if (argument is null)
                {
                    continue;
                }

                argument.Name = arguments[i].Name;
                argument.Description = arguments[i].Description;
            }
        }

        private static void Add(IDictionary<string, FunctionMetadata> dictionary, string functionName, string category, string docsRoot, string page, string description, params ArgumentMetadata[] arguments)
            => dictionary[functionName] = new FunctionMetadata(description, category, docsRoot + page, arguments);

        private static ArgumentMetadata Arg(string name, string description) => new(name, description);
        private static ArgumentMetadata WeightedValuesArg() => Arg("hodnoty", "Pozorování.");
        private static ArgumentMetadata WeightedWeightsArg() => Arg("váhy", "Váhy; prázdné buňky jsou brány jako 0.");
        private static ArgumentMetadata HeaderArg() => Arg("ma_záhlaví", "Volitelně: 0=autodetect, 1=má záhlaví, 2=nemá záhlaví.");
        private static ArgumentMetadata AlphaArg() => Arg("alpha", "Hladina významnosti.");
        private static ArgumentMetadata DirectionArg() => Arg("směr", "Volitelně: 0=two, 1=left, 2=right.");
        private static ArgumentMetadata PostHocArg() => Arg("post_hoc", "Volitelně: 0=none, 1=tukey, 2=bonferroni, 3=scheffe, 4=games-howell.");
        private static ArgumentMetadata HeaderArgEn() => Arg("has_header", "Optional: 0=autodetect, 1=has header, 2=no header.");
        private static ArgumentMetadata AlphaArgEn() => Arg("alpha", "Significance level.");
        private static ArgumentMetadata DirectionArgEn() => Arg("direction", "Optional: 0=two, 1=left, 2=right.");
        private static ArgumentMetadata PostHocArgEn() => Arg("post_hoc", "Optional: 0=none, 1=tukey, 2=bonferroni, 3=scheffe, 4=games-howell.");

        private sealed record FunctionMetadata(string Description, string Category, string HelpTopic, ArgumentMetadata[] Arguments);
        private sealed record ArgumentMetadata(string Name, string Description);

        private static IReadOnlyDictionary<string, FunctionMetadata> BuildCzechMetadata()
        {
            var general = "Obecné";
            var descriptive = "Popisné";
            var tests = "Testy";
            var docs = DocsRoot + "cs/functions/";
            var dictionary = new Dictionary<string, FunctionMetadata>();

            Add(dictionary, "GENERATE.NORM", general, docs, "distributions.md", "Generuje jednu náhodnou hodnotu z normálního rozdělení s volitelnou perturbací.",
                Arg("střední_hodnota", "Střední hodnota normálního rozdělení."),
                Arg("směrodatná_odchylka", "Kladná směrodatná odchylka."),
                Arg("outlier_rate", "Volitelně: pravděpodobnost dodatečné náhodné perturbace v intervalu <0;1>."));
            Add(dictionary, "GENERATE.INT", general, docs, "distributions.md", "Generuje jedno náhodné celé číslo z intervalu s volitelnou perturbací.",
                Arg("minimum", "Volitelně: dolní mez intervalu."),
                Arg("maximum", "Volitelně: horní mez intervalu."),
                Arg("outlier_rate", "Volitelně: pravděpodobnost dodatečné náhodné perturbace v intervalu <0;1>."));
            Add(dictionary, "FILL", general, docs, "distributions.md", "Opakuje jednu nebo více hodnot nebo textových vzorců do jednosloupcového spill výstupu.",
                Arg("co", "Hodnota nebo textový vzorec začínající znakem =."),
                Arg("počet", "Počet opakování; celé číslo >= 1."),
                Arg("další_pary", "Volitelně: další dvojice co + počet."));
            Add(dictionary, "FILL.RANDOM", general, docs, "distributions.md", "Opakuje hodnoty nebo textové vzorce do spill výstupu a výslednou řadu náhodně promíchá.",
                Arg("co", "Hodnota nebo textový vzorec začínající znakem =."),
                Arg("počet", "Počet opakování; celé číslo >= 1."),
                Arg("další_pary", "Volitelně: další dvojice co + počet."));

            Add(dictionary, "AVERAGE.W", descriptive, docs, "weighted-means.md", "Počítá vážený aritmetický průměr.",
                WeightedValuesArg(), WeightedWeightsArg());
            Add(dictionary, "HARMEAN.W", descriptive, docs, "weighted-means.md", "Počítá vážený harmonický průměr.",
                Arg("hodnoty", "Kladná pozorování."), WeightedWeightsArg());
            Add(dictionary, "GEOMEAN.W", descriptive, docs, "weighted-means.md", "Počítá vážený geometrický průměr.",
                Arg("hodnoty", "Kladná pozorování."), WeightedWeightsArg());
            Add(dictionary, "VAR.P.W", descriptive, docs, "weighted-variance.md", "Počítá vážený populační rozptyl.",
                WeightedValuesArg(), Arg("váhy", "Nezáporné váhy."));
            Add(dictionary, "VAR.S.W", descriptive, docs, "weighted-variance.md", "Počítá vážený výběrový rozptyl.",
                WeightedValuesArg(), Arg("váhy", "Nezáporné váhy."));
            Add(dictionary, "STDEV.P.W", descriptive, docs, "weighted-variance.md", "Počítá váženou populační směrodatnou odchylku.",
                WeightedValuesArg(), Arg("váhy", "Nezáporné váhy."));
            Add(dictionary, "STDEV.S.W", descriptive, docs, "weighted-variance.md", "Počítá váženou výběrovou směrodatnou odchylku.",
                WeightedValuesArg(), Arg("váhy", "Nezáporné váhy."));
            Add(dictionary, "VARCOEF", descriptive, docs, "variation-coefficients.md", "Počítá populační variační koeficient.",
                Arg("hodnoty", "Pozorování; prázdné buňky jsou ignorovány."));
            Add(dictionary, "VARCOEF.S", descriptive, docs, "variation-coefficients.md", "Počítá výběrový variační koeficient.",
                Arg("hodnoty", "Pozorování; prázdné buňky jsou ignorovány."));
            Add(dictionary, "VARCOEF.W", descriptive, docs, "variation-coefficients.md", "Počítá vážený populační variační koeficient.",
                WeightedValuesArg(), Arg("váhy", "Nezáporné váhy."));
            Add(dictionary, "VARCOEF.S.W", descriptive, docs, "variation-coefficients.md", "Počítá vážený výběrový variační koeficient.",
                WeightedValuesArg(), Arg("váhy", "Nezáporné váhy."));
            Add(dictionary, "PERCENTILE.INC.IFS", descriptive, docs, "percentiles.md", "Počítá inkluzivní percentil po aplikaci filtrů.",
                Arg("hodnoty", "Zdrojová data."),
                Arg("kvantil", "Požadovaný kvantil z intervalu <0;1>."));
            Add(dictionary, "PERCENTILE.EXC.IFS", descriptive, docs, "percentiles.md", "Počítá exkluzivní percentil po aplikaci filtrů.",
                Arg("hodnoty", "Zdrojová data."),
                Arg("kvantil", "Požadovaný kvantil z intervalu (0;1)."));

            AddPivotMetadata(dictionary, descriptive, docs);

            Add(dictionary, "NORM.DIST.RANGE", tests, docs, "distributions.md", "Počítá pravděpodobnost intervalu normálního rozdělení.",
                Arg("střední_hodnota", "Střední hodnota normálního rozdělení."),
                Arg("směrodatná_odchylka", "Kladná směrodatná odchylka."),
                Arg("dolní_hranice", "Dolní mez intervalu; prázdná buňka znamená minus nekonečno."),
                Arg("horní_hranice", "Horní mez intervalu; prázdná buňka znamená plus nekonečno."));
            Add(dictionary, "SHAPIRO.WILK", tests, docs, "normality.md", "Provádí Shapiro-Wilkův test normality.",
                Arg("rozsah_hodnot", "Číselná data; prázdné buňky jsou ignorovány."),
                HeaderArg());
            Add(dictionary, "KOLMOGOROV.SMIRNOV", tests, docs, "normality.md", "Provádí Kolmogorov-Smirnovův test dobré shody pro zvolené rozdělení.",
                Arg("rozsah_hodnot", "Číselná data; prázdné buňky jsou ignorovány."),
                Arg("typ_rozdělení", "Volitelně: 0=normal, 1=lognormal, 2=exponential, 3=uniform, 4=weibull."),
                HeaderArg());
            Add(dictionary, "T.TEST.1S", tests, docs, "one-sample-tests.md", "Provádí jednovýběrový t-test vůči zadané hypotetické střední hodnotě.",
                Arg("hodnoty", "Výběrová data; prázdné buňky jsou ignorovány."),
                Arg("mu_0", "Hypotetická střední hodnota."),
                DirectionArg(),
                AlphaArg(),
                HeaderArg());
            Add(dictionary, "PROP.TEST.1S", tests, docs, "one-sample-tests.md", "Provádí jednovýběrový test podílu.",
                Arg("hodnoty", "Binární data 0/1 nebo TRUE/FALSE."),
                Arg("pi_0", "Hypotetický populační podíl."),
                DirectionArg(),
                AlphaArg(),
                HeaderArg());
            Add(dictionary, "WILCOXON.PAIRED", tests, docs, "one-sample-tests.md", "Provádí Wilcoxonův párový znaménkový test pro závislá měření.",
                Arg("x", "První měření."),
                Arg("y", "Druhé měření."),
                HeaderArg(),
                AlphaArg(),
                DirectionArg());
            Add(dictionary, "WELCH.TEST.2S.G", tests, docs, "two-sample-tests.md", "Provádí Welchův dvouvýběrový t-test pro dvě nezávislé skupiny.",
                Arg("kategorie", "Štítky dvou skupin."),
                Arg("hodnoty", "Číselná pozorování."),
                HeaderArg(),
                AlphaArg(),
                DirectionArg());
            Add(dictionary, "MANN.WHITNEY.G", tests, docs, "two-sample-tests.md", "Provádí Mann-Whitneyho test pro dvě nezávislé skupiny.",
                Arg("kategorie", "Štítky dvou skupin."),
                Arg("hodnoty", "Číselná pozorování."),
                HeaderArg(),
                AlphaArg(),
                DirectionArg());
            Add(dictionary, "CHISQ.GOF", tests, docs, "goodness-of-fit.md", "Provádí chí-kvadrát test dobré shody.",
                Arg("pozorované", "Pozorované četnosti kategorií."),
                Arg("očekávané", "Očekávané četnosti nebo pravděpodobnosti."),
                Arg("kategorie", "Volitelné názvy kategorií."),
                AlphaArg(),
                HeaderArg());
            Add(dictionary, "ANOVA.G", tests, docs, "anova.md", "Provádí jednofaktorovou ANOVA nad groupovanými daty.",
                Arg("kategorie", "Štítky skupin."),
                Arg("hodnoty", "Číselná pozorování."),
                HeaderArg(),
                AlphaArg(),
                PostHocArg());
            Add(dictionary, "ANOVA.RM", tests, docs, "anova.md", "Provádí jednofaktorovou ANOVA s opakovaným měřením nad sloupci.",
                Arg("hodnoty", "Matice hodnot: řádky = subjekty, sloupce = podmínky."),
                HeaderArg(),
                AlphaArg(),
                PostHocArg());
            Add(dictionary, "ANCOVA.G", tests, docs, "ancova.md", "Provádí ANCOVA s jedním faktorem a jednou nebo více kovariátami.",
                Arg("faktor", "Kategorie faktoru."),
                Arg("závislá_proměnná", "Závislá proměnná."),
                Arg("kovariáty", "Jedna nebo více kovariát ve sloupcích."),
                PostHocArg(),
                AlphaArg(),
                HeaderArg());
            Add(dictionary, "CONTINGENCY.T", tests, docs, "contingency.md", "Analyzuje kontingenční tabulku zadanou přímo jako matici četností.",
                Arg("tabulka", "Kontingenční tabulka; volitelně i s horním řádkem a levým sloupcem popisků."),
                Arg("ma_záhlaví", "Volitelně: 0=autodetect, 1=horní řádek i levý sloupec jsou popisky, 2=bez popisků."),
                AlphaArg());
            Add(dictionary, "CONTINGENCY.G", tests, docs, "contingency.md", "Analyzuje kontingenční tabulku sestavenou z groupovaných sloupců.",
                Arg("sloupce", "Kategorie sloupců."),
                Arg("řádky", "Kategorie řádků."),
                Arg("počet", "Volitelně: četnosti dvojic; když chybí, každý záznam má váhu 1."),
                AlphaArg(),
                HeaderArg());
            Add(dictionary, "CORREL.SPEARMAN", tests, docs, "correlation.md", "Počítá Spearmanův korelační koeficient a test jeho významnosti.",
                Arg("rozsah_x", "První proměnná."),
                Arg("rozsah_y", "Druhá proměnná."),
                DirectionArg(),
                AlphaArg(),
                HeaderArg());
            Add(dictionary, "CORREL.MATRIX", tests, docs, "correlation.md", "Sestavuje korelační matici včetně p-hodnot a značek signifikance.",
                Arg("p_minimum", "Volitelně: zobrazí jen vazby s p < zadaná hodnota."),
                Arg("data", "Vstupní data; alespoň dva sloupce."),
                Arg("metoda", "Volitelně: 0=Pearson, 1=Spearman."),
                Arg("výstup", "Volitelně: 0=koeficienty, 1=p-hodnoty, 2=koeficient+p, 3=koeficient+p+sig, 4=koeficient+sig v jedné buňce."),
                HeaderArg());

            return dictionary;
        }

        private static IReadOnlyDictionary<string, FunctionMetadata> BuildEnglishMetadata()
        {
            var general = "General";
            var descriptive = "Descriptive";
            var tests = "Tests";
            var docs = DocsRoot + "en/functions/";
            var dictionary = new Dictionary<string, FunctionMetadata>();

            Add(dictionary, "GENERATE.NORM", general, docs, "distributions.md", "Generates a single random value from a normal distribution with optional perturbation.",
                Arg("mean", "Mean of the normal distribution."),
                Arg("stdev", "Positive standard deviation."),
                Arg("outlier_rate", "Optional: probability of additional random perturbation in the <0,1> interval."));
            Add(dictionary, "GENERATE.INT", general, docs, "distributions.md", "Generates a single random integer from an interval with optional perturbation.",
                Arg("minimum", "Optional: lower interval bound."),
                Arg("maximum", "Optional: upper interval bound."),
                Arg("outlier_rate", "Optional: probability of additional random perturbation in the <0,1> interval."));
            Add(dictionary, "FILL", general, docs, "distributions.md", "Repeats one or more values or text formulas into a single-column spill output.",
                Arg("what", "Value or text formula starting with =."),
                Arg("count", "Number of repetitions; integer >= 1."),
                Arg("more_pairs", "Optional: additional what + count pairs."));
            Add(dictionary, "FILL.RANDOM", general, docs, "distributions.md", "Repeats values or text formulas into a spill output and shuffles the final sequence.",
                Arg("what", "Value or text formula starting with =."),
                Arg("count", "Number of repetitions; integer >= 1."),
                Arg("more_pairs", "Optional: additional what + count pairs."));

            Add(dictionary, "AVERAGE.W", descriptive, docs, "weighted-means.md", "Computes the weighted arithmetic mean.",
                Arg("values", "Observed values."), Arg("weights", "Weights; blank cells are treated as 0."));
            Add(dictionary, "HARMEAN.W", descriptive, docs, "weighted-means.md", "Computes the weighted harmonic mean.",
                Arg("values", "Positive observed values."), Arg("weights", "Weights; blank cells are treated as 0."));
            Add(dictionary, "GEOMEAN.W", descriptive, docs, "weighted-means.md", "Computes the weighted geometric mean.",
                Arg("values", "Positive observed values."), Arg("weights", "Weights; blank cells are treated as 0."));
            Add(dictionary, "VAR.P.W", descriptive, docs, "weighted-variance.md", "Computes the weighted population variance.",
                Arg("values", "Observed values."), Arg("weights", "Non-negative weights."));
            Add(dictionary, "VAR.S.W", descriptive, docs, "weighted-variance.md", "Computes the weighted sample variance.",
                Arg("values", "Observed values."), Arg("weights", "Non-negative weights."));
            Add(dictionary, "STDEV.P.W", descriptive, docs, "weighted-variance.md", "Computes the weighted population standard deviation.",
                Arg("values", "Observed values."), Arg("weights", "Non-negative weights."));
            Add(dictionary, "STDEV.S.W", descriptive, docs, "weighted-variance.md", "Computes the weighted sample standard deviation.",
                Arg("values", "Observed values."), Arg("weights", "Non-negative weights."));
            Add(dictionary, "VARCOEF", descriptive, docs, "variation-coefficients.md", "Computes the population coefficient of variation.",
                Arg("values", "Observed values; blank cells are ignored."));
            Add(dictionary, "VARCOEF.S", descriptive, docs, "variation-coefficients.md", "Computes the sample coefficient of variation.",
                Arg("values", "Observed values; blank cells are ignored."));
            Add(dictionary, "VARCOEF.W", descriptive, docs, "variation-coefficients.md", "Computes the weighted population coefficient of variation.",
                Arg("values", "Observed values."), Arg("weights", "Non-negative weights."));
            Add(dictionary, "VARCOEF.S.W", descriptive, docs, "variation-coefficients.md", "Computes the weighted sample coefficient of variation.",
                Arg("values", "Observed values."), Arg("weights", "Non-negative weights."));
            Add(dictionary, "PERCENTILE.INC.IFS", descriptive, docs, "percentiles.md", "Computes the inclusive percentile after applying filters.",
                Arg("values", "Source data."),
                Arg("quantile", "Requested quantile in the <0,1> interval."));
            Add(dictionary, "PERCENTILE.EXC.IFS", descriptive, docs, "percentiles.md", "Computes the exclusive percentile after applying filters.",
                Arg("values", "Source data."),
                Arg("quantile", "Requested quantile in the (0,1) interval."));

            AddPivotMetadataEnglish(dictionary, descriptive, docs);

            Add(dictionary, "NORM.DIST.RANGE", tests, docs, "distributions.md", "Computes the probability of an interval under the normal distribution.",
                Arg("mean", "Mean of the normal distribution."),
                Arg("stdev", "Positive standard deviation."),
                Arg("lower_bound", "Lower interval bound; blank means minus infinity."),
                Arg("upper_bound", "Upper interval bound; blank means plus infinity."));
            Add(dictionary, "SHAPIRO.WILK", tests, docs, "normality.md", "Performs the Shapiro-Wilk normality test.",
                Arg("values", "Numeric data; blank cells are ignored."),
                HeaderArgEn());
            Add(dictionary, "KOLMOGOROV.SMIRNOV", tests, docs, "normality.md", "Performs the Kolmogorov-Smirnov goodness-of-fit test for the selected distribution.",
                Arg("values", "Numeric data; blank cells are ignored."),
                Arg("distribution", "Optional: 0=normal, 1=lognormal, 2=exponential, 3=uniform, 4=weibull."),
                HeaderArgEn());
            Add(dictionary, "T.TEST.1S", tests, docs, "one-sample-tests.md", "Performs a one-sample t-test against a specified hypothetical mean.",
                Arg("values", "Sample data; blank cells are ignored."),
                Arg("mu_0", "Hypothetical mean."),
                DirectionArgEn(),
                AlphaArgEn(),
                HeaderArgEn());
            Add(dictionary, "PROP.TEST.1S", tests, docs, "one-sample-tests.md", "Performs a one-sample proportion test.",
                Arg("values", "Binary data 0/1 or TRUE/FALSE."),
                Arg("pi_0", "Hypothetical population proportion."),
                DirectionArgEn(),
                AlphaArgEn(),
                HeaderArgEn());
            Add(dictionary, "WILCOXON.PAIRED", tests, docs, "one-sample-tests.md", "Performs the Wilcoxon signed-rank test for paired measurements.",
                Arg("x", "First measurement."),
                Arg("y", "Second measurement."),
                HeaderArgEn(),
                AlphaArgEn(),
                DirectionArgEn());
            Add(dictionary, "WELCH.TEST.2S.G", tests, docs, "two-sample-tests.md", "Performs Welch's two-sample t-test for two independent groups.",
                Arg("categories", "Labels of the two groups."),
                Arg("values", "Numeric observations."),
                HeaderArgEn(),
                AlphaArgEn(),
                DirectionArgEn());
            Add(dictionary, "MANN.WHITNEY.G", tests, docs, "two-sample-tests.md", "Performs the Mann-Whitney test for two independent groups.",
                Arg("categories", "Labels of the two groups."),
                Arg("values", "Numeric observations."),
                HeaderArgEn(),
                AlphaArgEn(),
                DirectionArgEn());
            Add(dictionary, "CHISQ.GOF", tests, docs, "goodness-of-fit.md", "Performs the chi-square goodness-of-fit test.",
                Arg("observed", "Observed category frequencies."),
                Arg("expected", "Expected frequencies or probabilities."),
                Arg("categories", "Optional category names."),
                AlphaArgEn(),
                HeaderArgEn());
            Add(dictionary, "ANOVA.G", tests, docs, "anova.md", "Performs one-way ANOVA on grouped data.",
                Arg("categories", "Group labels."),
                Arg("values", "Numeric observations."),
                HeaderArgEn(),
                AlphaArgEn(),
                PostHocArgEn());
            Add(dictionary, "ANOVA.RM", tests, docs, "anova.md", "Performs one-factor repeated-measures ANOVA across columns.",
                Arg("values", "Value matrix: rows = subjects, columns = conditions."),
                HeaderArgEn(),
                AlphaArgEn(),
                PostHocArgEn());
            Add(dictionary, "ANCOVA.G", tests, docs, "ancova.md", "Performs ANCOVA with one factor and one or more covariates.",
                Arg("factor", "Factor categories."),
                Arg("dependent_variable", "Dependent variable."),
                Arg("covariates", "One or more covariates in columns."),
                PostHocArgEn(),
                AlphaArgEn(),
                HeaderArgEn());
            Add(dictionary, "CONTINGENCY.T", tests, docs, "contingency.md", "Analyzes a contingency table entered directly as a frequency matrix.",
                Arg("table", "Contingency table; optionally with top-row and left-column labels."),
                Arg("has_header", "Optional: 0=autodetect, 1=top row and left column are labels, 2=no labels."),
                AlphaArgEn());
            Add(dictionary, "CONTINGENCY.G", tests, docs, "contingency.md", "Analyzes a contingency table built from grouped columns.",
                Arg("columns", "Column categories."),
                Arg("rows", "Row categories."),
                Arg("count", "Optional: pair frequencies; if omitted, each record has weight 1."),
                AlphaArgEn(),
                HeaderArgEn());
            Add(dictionary, "CORREL.SPEARMAN", tests, docs, "correlation.md", "Computes Spearman correlation and its significance test.",
                Arg("range_x", "First variable."),
                Arg("range_y", "Second variable."),
                DirectionArgEn(),
                AlphaArgEn(),
                HeaderArgEn());
            Add(dictionary, "CORREL.MATRIX", tests, docs, "correlation.md", "Builds a correlation matrix including p-values and significance markers.",
                Arg("p_minimum", "Optional: show only relationships with p < specified value."),
                Arg("data", "Input data; at least two columns."),
                Arg("method", "Optional: 0=Pearson, 1=Spearman."),
                Arg("output", "Optional: 0=coefficients, 1=p-values, 2=coefficient+p, 3=coefficient+p+sig, 4=coefficient+sig in one cell."),
                HeaderArgEn());

            return dictionary;
        }

        private static void AddPivotMetadata(IDictionary<string, FunctionMetadata> dictionary, string category, string docsRoot)
        {
            var pivotArgs = new[]
            {
                Arg("řádky", "Jeden nebo více sloupců s kategoriemi řádků včetně záhlaví."),
                Arg("sloupce", "Volitelně: jeden nebo více sloupců s kategoriemi sloupců včetně záhlaví."),
                Arg("hodnoty", "Jeden sloupec hodnot včetně záhlaví.")
            };

            Add(dictionary, "PIVOT.COUNT", category, docsRoot, "pivot.md", "Sestaví pivot a spočítá počet neprázdných hodnot.", pivotArgs);
            Add(dictionary, "PIVOT.SUM", category, docsRoot, "pivot.md", "Sestaví pivot a spočítá součet.", pivotArgs);
            Add(dictionary, "PIVOT.AVERAGE", category, docsRoot, "pivot.md", "Sestaví pivot a spočítá aritmetický průměr.", pivotArgs);
            Add(dictionary, "PIVOT.MIN", category, docsRoot, "pivot.md", "Sestaví pivot a spočítá minimum.", pivotArgs);
            Add(dictionary, "PIVOT.MAX", category, docsRoot, "pivot.md", "Sestaví pivot a spočítá maximum.", pivotArgs);
            Add(dictionary, "PIVOT.MEDIAN", category, docsRoot, "pivot.md", "Sestaví pivot a spočítá medián.", pivotArgs);
            Add(dictionary, "PIVOT.PERCENTILE", category, docsRoot, "pivot.md", "Sestaví pivot a spočítá percentil.",
                pivotArgs[0], pivotArgs[1], pivotArgs[2], Arg("kvantil", "Hledaný kvantil v intervalu (0;1)."));
            Add(dictionary, "PIVOT.STDEV.S", category, docsRoot, "pivot.md", "Sestaví pivot a spočítá výběrovou směrodatnou odchylku.", pivotArgs);
            Add(dictionary, "PIVOT.STDEV.P", category, docsRoot, "pivot.md", "Sestaví pivot a spočítá populační směrodatnou odchylku.", pivotArgs);
            Add(dictionary, "PIVOT.VAR.S", category, docsRoot, "pivot.md", "Sestaví pivot a spočítá výběrový rozptyl.", pivotArgs);
            Add(dictionary, "PIVOT.VAR.P", category, docsRoot, "pivot.md", "Sestaví pivot a spočítá populační rozptyl.", pivotArgs);
            Add(dictionary, "PIVOT.VARCOEF.S", category, docsRoot, "pivot.md", "Sestaví pivot a spočítá výběrový variační koeficient.", pivotArgs);
            Add(dictionary, "PIVOT.VARCOEF.P", category, docsRoot, "pivot.md", "Sestaví pivot a spočítá populační variační koeficient.", pivotArgs);
            Add(dictionary, "PIVOT.CONF.T", category, docsRoot, "pivot.md", "Sestaví pivot a spočítá poloviční šířku intervalu spolehlivosti pro t-rozdělení.",
                pivotArgs[0], pivotArgs[1], pivotArgs[2], Arg("alfa", "Hladina alfa v intervalu (0;1)."), Arg("směr", "Volitelně: 0=oboustranný, -1=levostranný, 1=pravostranný."));
            Add(dictionary, "PIVOT.CONF.NORM", category, docsRoot, "pivot.md", "Sestaví pivot a spočítá poloviční šířku intervalu spolehlivosti pro normální aproximaci.",
                pivotArgs[0], pivotArgs[1], pivotArgs[2], Arg("alfa", "Hladina alfa v intervalu (0;1)."), Arg("směr", "Volitelně: 0=oboustranný, -1=levostranný, 1=pravostranný."));
            Add(dictionary, "PIVOT.MAD", category, docsRoot, "pivot.md", "Sestaví pivot a spočítá medián absolutních odchylek od mediánu.", pivotArgs);
            Add(dictionary, "PIVOT.IQR", category, docsRoot, "pivot.md", "Sestaví pivot a spočítá mezikvartilové rozpětí.", pivotArgs);
        }

        private static void AddPivotMetadataEnglish(IDictionary<string, FunctionMetadata> dictionary, string category, string docsRoot)
        {
            var pivotArgs = new[]
            {
                Arg("rows", "One or more row-category columns including their headers."),
                Arg("columns", "Optional: one or more column-category columns including their headers."),
                Arg("values", "One value column including its header.")
            };

            Add(dictionary, "PIVOT.COUNT", category, docsRoot, "pivot.md", "Builds a pivot and counts non-empty values.", pivotArgs);
            Add(dictionary, "PIVOT.SUM", category, docsRoot, "pivot.md", "Builds a pivot and computes the sum.", pivotArgs);
            Add(dictionary, "PIVOT.AVERAGE", category, docsRoot, "pivot.md", "Builds a pivot and computes the arithmetic mean.", pivotArgs);
            Add(dictionary, "PIVOT.MIN", category, docsRoot, "pivot.md", "Builds a pivot and computes the minimum.", pivotArgs);
            Add(dictionary, "PIVOT.MAX", category, docsRoot, "pivot.md", "Builds a pivot and computes the maximum.", pivotArgs);
            Add(dictionary, "PIVOT.MEDIAN", category, docsRoot, "pivot.md", "Builds a pivot and computes the median.", pivotArgs);
            Add(dictionary, "PIVOT.PERCENTILE", category, docsRoot, "pivot.md", "Builds a pivot and computes a percentile.",
                pivotArgs[0], pivotArgs[1], pivotArgs[2], Arg("quantile", "Requested quantile in the (0,1) interval."));
            Add(dictionary, "PIVOT.STDEV.S", category, docsRoot, "pivot.md", "Builds a pivot and computes the sample standard deviation.", pivotArgs);
            Add(dictionary, "PIVOT.STDEV.P", category, docsRoot, "pivot.md", "Builds a pivot and computes the population standard deviation.", pivotArgs);
            Add(dictionary, "PIVOT.VAR.S", category, docsRoot, "pivot.md", "Builds a pivot and computes the sample variance.", pivotArgs);
            Add(dictionary, "PIVOT.VAR.P", category, docsRoot, "pivot.md", "Builds a pivot and computes the population variance.", pivotArgs);
            Add(dictionary, "PIVOT.VARCOEF.S", category, docsRoot, "pivot.md", "Builds a pivot and computes the sample coefficient of variation.", pivotArgs);
            Add(dictionary, "PIVOT.VARCOEF.P", category, docsRoot, "pivot.md", "Builds a pivot and computes the population coefficient of variation.", pivotArgs);
            Add(dictionary, "PIVOT.CONF.T", category, docsRoot, "pivot.md", "Builds a pivot and computes the half-width of a t-based confidence interval.",
                pivotArgs[0], pivotArgs[1], pivotArgs[2], Arg("alpha", "Alpha level in the (0,1) interval."), Arg("direction", "Optional: 0=two-sided, -1=left-sided, 1=right-sided."));
            Add(dictionary, "PIVOT.CONF.NORM", category, docsRoot, "pivot.md", "Builds a pivot and computes the half-width of a normal-approximation confidence interval.",
                pivotArgs[0], pivotArgs[1], pivotArgs[2], Arg("alpha", "Alpha level in the (0,1) interval."), Arg("direction", "Optional: 0=two-sided, -1=left-sided, 1=right-sided."));
            Add(dictionary, "PIVOT.MAD", category, docsRoot, "pivot.md", "Builds a pivot and computes the median absolute deviation from the median.", pivotArgs);
            Add(dictionary, "PIVOT.IQR", category, docsRoot, "pivot.md", "Builds a pivot and computes the interquartile range.", pivotArgs);
        }
    }
}
