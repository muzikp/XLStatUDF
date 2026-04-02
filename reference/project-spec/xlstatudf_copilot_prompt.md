# XLStatUDF – Zadání pro implementaci Excel UDF doplňku

> Tento soubor je zadání pro AI asistenta (Claude Sonnet 4.6 přes GitHub Copilot ve VSCode).  
> Spolu s tímto souborem máš k dispozici `xlstatudf_udf_dokumentace.md` — přečti ho celý před začátkem implementace, obsahuje přesné specifikace všech funkcí včetně argumentů, výstupů a chybových stavů.

---

## 1. Cíl

Vytvořit Excel doplněk **XLStatUDF** jako sadu statistických UDF (User-Defined Functions). Doplněk bude distribuován jako instalační balíček (`.exe`), který na cílovém počítači nainstaluje `.xll` soubor a zaregistruje ho v Excelu.

---

## 2. Technologický stack

| Vrstva | Technologie | Poznámka |
|--------|-------------|----------|
| Jazyk | C# (.NET 8) | |
| Excel integrace | **Excel-DNA** (NuGet: `ExcelDna.Integration`, `ExcelDna.AddIn`) | Produces `.xll` |
| Balení add-inu | **ExcelDnaPack** (součást Excel-DNA toolingu) | Single-file `.xll` |
| Statistické distribuce | **MathNet.Numerics** (NuGet: `MathNet.Numerics`) | Normal, StudentT, FisherSnedecor, ChiSquared, aj. |
| Instalátor | **InnoSetup 6** (`XLStatUDF.iss`) | Kompiluje do `.exe` instalátoru |

---

## 3. Struktura projektu

```
XLStatUDF/
├── src/
│   └── XLStatUDF/
│       ├── XLStatUDF.csproj
│       ├── XLStatUDF.dna                  ← Excel-DNA manifest
│       ├── Functions/
│       │   ├── Descriptive/
│       │   │   ├── WeightedMeans.cs     ← AVERAGE.W, HARMEAN.W, GEOMEAN.W
│       │   │   ├── WeightedVariance.cs  ← VAR.P.W, VAR.S.W, STDEV.P.W, STDEV.S.W
│       │   │   └── VarCoef.cs           ← VARCOEF, VARCOEF.S, VARCOEF.S.W
│       │   ├── Distributions/
│       │   │   └── NormDistRange.cs     ← NORM.DIST.RANGE
│       │   ├── Normality/
│       │   │   ├── ShapiroWilk.cs       ← SHAPIRO.WILK
│       │   │   └── KolmogorovSmirnov.cs ← KOLMOGOROV.SMIRNOV
│       │   ├── Tests/
│       │   │   ├── TTest1S.cs           ← T.TEST.1S
│       │   │   ├── WelchTest2S.cs       ← WELCH.TEST.2S
│       │   │   ├── PropTest1S.cs        ← PROP.TEST.1S
│       │   │   ├── Anova.cs             ← ANOVA
│       │   │   └── ChisqGof.cs          ← CHISQ.GOF
│       │   ├── Correlation/
│       │   │   └── SpearmanCorrel.cs    ← CORREL.SPEARMAN
│       │   └── Percentiles/
│       │       └── PercentileIfs.cs     ← PERCENTILE.INC.IFS, PERCENTILE.EXC.IFS
│       └── Helpers/
│           ├── SpillBuilder.cs          ← sestavení object[,] výstupů
│           ├── WeightHelper.cs          ← normalizace a validace vah
│           ├── RankHelper.cs            ← midrank pro ties (Spearman, Mann-Whitney)
│           ├── FilterHelper.cs          ← SUMIFS-style filtrování (Percentile.IFS)
│           ├── CriticalValues.cs        ← tα/zα/Fα/χ²α + dynamické názvy
│           └── ExcelErrors.cs           ← konstanty pro chybové návratové hodnoty
├── installer/
│   ├── XLStatUDF.iss                      ← InnoSetup skript
│   └── assets/
│       └── XLStatUDF_logo.bmp             ← volitelné, pro instalátor
├── build.ps1                            ← PowerShell build skript
└── README.md
```

---

## 4. Excel-DNA — klíčové konvence

### 4.1 Podpis funkce

```csharp
using ExcelDna.Integration;

public static class WeightedMeans
{
    [ExcelFunction(
        Name = "AVERAGE.W",
        Description = "Vážený aritmetický průměr",
        Category = "XLStatUDF")]
    public static object AverageW(
        [ExcelArgument(Name = "hodnoty", Description = "Pozorování")] double[] values,
        [ExcelArgument(Name = "váhy",    Description = "Váhy (≥ 0)")] double[] weights)
    {
        // implementace
    }
}
```

- Všechny UDF metody musí být `public static` v `public` třídě.
- Návratový typ skalárních funkcí: `object` (vrátí `double`, `string` nebo `ExcelError`).
- Návratový typ spill funkcí: `object[,]` — Excel-DNA automaticky rozleje přes dynamic array.

### 4.2 Spill range — konvence výstupu

Všechny víceřádkové výstupy jsou `object[,]` (řádky × sloupce). Sestavuj je přes `SpillBuilder`:

```csharp
// Příklad: 2-sloupcový spill
public static class SpillBuilder
{
    // Přidá řádek [název, hodnota] do 2-sloupcové tabulky
    public static void AddRow(List<object[]> rows, string label, object value)
        => rows.Add(new object[] { label, value });

    // Přidá prázdný oddělovací řádek
    public static void AddSeparator(List<object[]> rows, int cols)
        => rows.Add(Enumerable.Repeat<object>("", cols).ToArray());

    // Přidá záhlaví bloku (1. buňka, zbytek prázdné)
    public static void AddHeader(List<object[]> rows, string title, int cols)
    {
        var row = Enumerable.Repeat<object>("", cols).ToArray();
        row[0] = title;
        rows.Add(row);
    }

    // Převede List<object[]> na object[,]
    public static object[,] Build(List<object[]> rows)
    {
        int rowCount = rows.Count;
        int colCount = rows.Max(r => r.Length);
        var result = new object[rowCount, colCount];
        for (int i = 0; i < rowCount; i++)
            for (int j = 0; j < colCount; j++)
                result[i, j] = j < rows[i].Length ? rows[i][j] : "";
        return result;
    }
}
```

### 4.3 Chybové hodnoty

```csharp
using ExcelDna.Integration;

public static class ExcelErrors
{
    public static readonly object Value   = ExcelError.ExcelErrorValue;   // #HODNOTA!
    public static readonly object Num     = ExcelError.ExcelErrorNum;     // #ČÍSLO!
    public static readonly object NA      = ExcelError.ExcelErrorNA;      // #NENÍ_K_DISP!
    public static readonly object DivZero = ExcelError.ExcelErrorDiv0;    // #DĚLENÍ0!
    // #POČET! a #DÉLKA! nemají přímý ekvivalent — použij ExcelErrorValue nebo vlastní string
}
```

### 4.4 Dynamický název kritické hodnoty

```csharp
public static class CriticalValues
{
    // Vrátí název buňky pro kritickou hodnotu podle směru testu
    // distribution: "t", "z", "F", "chi2"
    public static string LabelForDirection(string distribution, string smer)
    {
        string sub = distribution switch
        {
            "t"    => "t",
            "z"    => "z",
            "F"    => "F",
            "chi2" => "χ²",
            _      => distribution
        };
        return smer == "two"
            ? $"{sub}₁₋α/₂"
            : $"{sub}₁₋α";
    }
}
```

### 4.5 Validace vah (WeightHelper)

```csharp
public static class WeightHelper
{
    // Ověří váhy, vrátí null pokud OK nebo chybový object
    public static object? Validate(double[] values, double[] weights)
    {
        if (values.Length != weights.Length) return "DÉLKA";
        if (weights.Any(w => w < 0))        return ExcelErrors.Num;
        if (weights.Sum() == 0)             return ExcelErrors.Num;
        return null;
    }

    // Vrátí normalizované váhy (wᵢ / Σwᵢ), prázdné buňky (NaN) jako 0
    public static double[] Normalize(double[] weights)
    {
        var w = weights.Select(x => double.IsNaN(x) ? 0 : x).ToArray();
        double sum = w.Sum();
        return w.Select(x => x / sum).ToArray();
    }
}
```

---

## 5. Statistické implementace — klíčové algoritmy

### 5.1 MathNet.Numerics — použití pro distribuce

```csharp
using MathNet.Numerics.Distributions;

// p-hodnota pro t-test (oboustranný)
double pValue = 2 * (1 - StudentT.CDF(0, 1, df, Math.Abs(tStat)));

// Kritická hodnota F
double fCrit = FisherSnedecor.InvCDF(df1, df2, 1 - alpha);

// Kritická hodnota χ²
double chiCrit = ChiSquared.InvCDF(df, 1 - alpha);

// Normální CDF pro NORM.DIST.RANGE
double prob = Normal.CDF(mean, stddev, upper) - Normal.CDF(mean, stddev, lower);
```

### 5.2 Shapiro-Wilk (Royston 1992)

Implementuj Roystonovu aproximaci pro n = 3..5000:
- Koeficienty `a[]` pro výpočet W se generují z normálních pořadových statistik (approximace přes `phi^{-1}((i - 3/8) / (n + 1/4))`).
- Testová statistika: W = (Σ aᵢ x_(i))² / Σ(xᵢ - x̄)²
- p-hodnota přes aproximaci normálního rozdělení aplikovanou na transformaci z = ((1-W)^λ - μ) / σ, kde λ, μ, σ závisí na n (tabelovány v Royston 1992, J. Applied Statistics).
- Referenční implementace: [Algorithm AS R94, Royston 1995](https://www.jstor.org/stable/2347973) — koeficienty jsou veřejně dostupné v R zdrojovém kódu (`shapiro.test`).

### 5.3 Kolmogorov-Smirnov (jednovýběrový)

- Testová statistika D = max|F_emp(x) - F_teor(x)|
- Parametry odhadni z dat (MLE) — pro `"normal"`: μ = x̄, σ = s; pro `"exponential"`: λ = 1/x̄ atd.
- p-hodnota pro n ≥ 35: asymptotická aproximace přes Kolmogorovovo rozdělení.
- Pro `"normal"` s odhadnutými parametry aplikuj **Lillieforsovu korekci**: kritické hodnoty z Lillieforsových tabulek (ne standardní KS tabulky).

### 5.4 Spearmanův korelační koeficient

```csharp
// 1. Párově vyřaď NaN
// 2. Přiřaď pořadí s midrank pro ties (průměrné pořadí)
// 3. Spočítej Pearsonovu korelaci nad pořadími
// 4. t = rho * sqrt(n-2) / sqrt(1 - rho^2), df = n-2
// 5. p-hodnota přes StudentT

public static double[] MidRank(double[] values)
{
    // Seřaď indexy, skupinám se stejnou hodnotou přiřaď průměrné pořadí
    var indexed = values.Select((v, i) => (v, i)).OrderBy(x => x.v).ToArray();
    var ranks = new double[values.Length];
    int i = 0;
    while (i < indexed.Length)
    {
        int j = i;
        while (j < indexed.Length && indexed[j].v == indexed[i].v) j++;
        double avgRank = (i + j - 1) / 2.0 + 1; // 1-based
        for (int k = i; k < j; k++) ranks[indexed[k].i] = avgRank;
        i = j;
    }
    return ranks;
}
```

### 5.5 Tukey HSD (post-hoc pro ANOVA)

MathNet.Numerics nemá Studentized Range Distribution. Použij jednu z možností:
- **Doporučeno**: implementuj numerickou aproximaci Studentized Range CDF (algoritmus AS 190, Lund & Lund 1983) — zdrojový kód v Fortran/C je volně dostupný.
- **Alternativa**: pokud je implementace příliš komplexní, použij Bonferroniho korekci jako fallback pro `"tukey"` a do popisu přidej poznámku — Bonferroni je konzervativnější, ale správný.

### 5.6 SUMIFS-style filtrování (Percentile.IFS)

```csharp
// criteriaArgs: střídají se rozsah a kritérium jako object[]
public static double[] ApplyFilters(double[] values, object[] criteriaArgs)
{
    // Parsuj páry (rozsah[], kritérium)
    // Kritérium může být: číslo, text s operátorem (">5", "<=10", "<>0"), wildcard ("A*")
    // Vrať values[], kde všechny podmínky platí (AND logika)
}
```

---

## 6. XLStatUDF.dna — manifest

```xml
<DnaLibrary Name="XLStatUDF" RuntimeVersion="v4.0" xmlns="http://schemas.excel-dna.net/addin/2020/07/dnalibrary">
  <ExternalLibrary Path="XLStatUDF.dll" ExplicitExports="false" LoadFromBytes="true" Pack="true" />
</DnaLibrary>
```

- `Pack="true"` zajistí, že ExcelDnaPack zabalí DLL + NuGet závislosti do jednoho `.xll`.

---

## 7. Build skript (build.ps1)

```powershell
# build.ps1 — spustit z kořene repozitáře
param([string]$Configuration = "Release")

Set-Location src/XLStatUDF
dotnet restore
dotnet build -c $Configuration

# ExcelDnaPack — zabalí XLStatUDF.xll + závislosti do XLStatUDF-packed.xll
$packTool = "$env:USERPROFILE\.nuget\packages\exceldna.addin\*\tools\ExcelDnaPack.exe"
& (Resolve-Path $packTool | Select-Object -Last 1) XLStatUDF.xll /O

# Zkopíruj výstup do installer/
Copy-Item "bin/$Configuration/net8.0-windows/XLStatUDF-packed.xll" `
          "../../installer/XLStatUDF-packed.xll" -Force

Write-Host "Build hotový. Spusť InnoSetup pro vytvoření instalátoru."
```

---

## 8. InnoSetup skript (XLStatUDF.iss)

```iss
[Setup]
AppName=XLStatUDF
AppVersion=1.0
DefaultDirName={autopf}\XLStatUDF
DefaultGroupName=XLStatUDF
OutputBaseFilename=XLStatUDF_Setup
Compression=lzma2
SolidCompression=yes

[Files]
Source: "XLStatUDF-packed.xll"; DestDir: "{userappdata}\Microsoft\Excel\XLSTART"; Flags: ignoreversion

[Icons]
Name: "{group}\Odinstalovat XLStatUDF"; Filename: "{uninstallexe}"

[Code]
// Volitelně: zápis do registru pro automatické načtení při startu Excelu
procedure RegisterAddin();
var
  RegKey: string;
begin
  RegKey := 'Software\Microsoft\Office\16.0\Excel\Options';
  RegWriteStringValue(HKCU, RegKey, 'OPEN', '/R "XLStatUDF-packed.xll"');
end;

procedure CurStepChanged(CurStep: TSetupStep);
begin
  if CurStep = ssPostInstall then
    RegisterAddin();
end;
```

> Poznámka: Číslo verze Office v cestě registru (`16.0` = Office 2016/2019/365) může být potřeba detekovat dynamicky pro širší kompatibilitu.

---

## 9. Pořadí implementace

Implementuj funkce v tomto pořadí — od jednodušších ke komplexnějším, přičemž každá skupina reusuje infrastrukturu té předchozí:

1. **Infrastruktura** — `SpillBuilder`, `WeightHelper`, `RankHelper`, `FilterHelper`, `CriticalValues`, `ExcelErrors`
2. **Skalární funkce** — `NORM.DIST.RANGE`, `AVERAGE.W`, `HARMEAN.W`, `GEOMEAN.W`
3. **Varianční funkce** — `VAR.P.W`, `VAR.S.W`, `STDEV.P.W`, `STDEV.S.W`, `VARCOEF`, `VARCOEF.S`, `VARCOEF.S.W`
4. **Percentily** — `PERCENTILE.INC.IFS`, `PERCENTILE.EXC.IFS` (vyžaduje `FilterHelper`)
5. **Normality testy** — `SHAPIRO.WILK`, `KOLMOGOROV.SMIRNOV`
6. **Jednoduché testy** — `T.TEST.1S`, `PROP.TEST.1S`, `CORREL.SPEARMAN`
7. **Složené testy** — `WELCH.TEST.2S`, `CHISQ.GOF`
8. **ANOVA** — včetně Levene + post-hoc (nejkomplexnější)
9. **Build & installer** — `build.ps1`, `XLStatUDF.iss`

---

## 10. Požadavky na kód

- Každý soubor začíná XML doc-commentem `/// <summary>` popisujícím funkci.
- Validace argumentů vždy jako první krok — při chybě okamžitě `return ExcelErrors.XYZ`.
- Žádné globální stavy — všechny funkce jsou čisté (pure), bez side-effectů.
- Kultura: čísla vždy formátuj s `CultureInfo.InvariantCulture` (Excel si locale řeší sám).
- Cílový framework: `net8.0-windows` (Excel-DNA vyžaduje Windows).
- Testy: ke každé funkci vlož alespoň jeden xUnit test s ověřením výstupu oproti referenční hodnotě (R nebo Python/scipy).

---

## 11. Dokumentace funkcí uvnitř Excelu

### 11.1 Vrstvy nápovědy

Excel-DNA podporuje tři vrstvy dokumentace viditelné přímo v Excelu:

| Vrstva | Kde se zobrazí | Jak nastavit |
|--------|---------------|--------------|
| Popis funkce | Průvodce vložením funkce (Shift+F3), tooltip při psaní | `[ExcelFunction(Description = "...")]` |
| Popis argumentů | Průvodce funkcí, spodní panel při psaní | `[ExcelArgument(Description = "...")]` |
| Nápověda (F1) | Otevře URL v prohlížeči | `[ExcelFunction(HelpTopic = "...")]` |

```csharp
[ExcelFunction(
    Name        = "VARCOEF.S",
    Description = "Výběrový variační koeficient: s / x̄",
    Category    = "XLStatUDF",
    HelpTopic   = "https://your-docs-url/varcoef-s")]
public static object VarCoefS(
    [ExcelArgument(Name = "hodnoty", Description = "Číselný rozsah; prázdné buňky jsou ignorovány")]
    double[] values) { ... }
```

Všechny `Description` a `Name` argumentů jsou **výhradně česky**. `HelpTopic` odkazuje na URL — ideálně GitHub Pages nebo obdobný hostovaný soubor odvozený z `xlstatudf_udf_dokumentace.md`.

### 11.2 Kategorie funkcí

Všechny UDF patří do kategorie `"XLStatUDF"`. Podkategorii simuluj prefixem v `Description`:

```csharp
[ExcelFunction(Category = "XLStatUDF", Description = "[Vážené] Vážený aritmetický průměr")]
[ExcelFunction(Category = "XLStatUDF", Description = "[Testy] Jednovýběrový t-test")]
[ExcelFunction(Category = "XLStatUDF", Description = "[Normalita] Shapiro-Wilkův test")]
[ExcelFunction(Category = "XLStatUDF", Description = "[Korelace] Spearmanův korelační koeficient")]
[ExcelFunction(Category = "XLStatUDF", Description = "[Percentily] Percentil s filtrováním (INC)")]
```

---

## 12. Reference

- Excel-DNA dokumentace: https://excel-dna.net/docs/
- MathNet.Numerics: https://numerics.mathdotnet.com/
- Royston (1992) Shapiro-Wilk: Algorithm AS R94, Applied Statistics
- R zdrojový kód `shapiro.test` a `ks.test` pro referenční koeficienty a algoritmy
- InnoSetup: https://jrsoftware.org/ishelp/
