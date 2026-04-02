# XLStatUDF – UDF pro statistickou analýzu v Excelu

> Dokumentace vlastních funkcí (User-Defined Functions) pro doplněk Excelu.  
> Účel: specifikace pro implementaci v prostředí Excel-DNA / VBA / Office JS Add-in.

---

## Obecné konvence

### Oddělovač argumentů
Funkce používají středník (`;`) jako oddělovač argumentů (CZ/SK locale Excel).  
V anglickém locale nahradit čárkou (`,`).

### Typy výstupu

| Typ | Popis | Příklad |
|-----|-------|---------|
| **Skalár** | Vrátí jedinou hodnotu do buňky | `NORM.DIST.RANGE` |
| **Spill 2-sloupcový** | Rozepíše se dolů od aktivní buňky: levý sloupec = název ukazatele, pravý sloupec = hodnota. Vhodné pro skalární výsledky testů. | `SHAPIRO.WILK`, `KOLMOGOROV.SMIRNOV`, `WELCH.TEST.2S` |
| **Spill víc-sloupcový** | Rozepíše se do obdélníkového rozsahu. Počet sloupců se přizpůsobuje struktuře výstupu (tabulka, párové srovnání). Jednotlivé bloky výstupu jsou odděleny prázdným řádkem a záhlavím. | `ANOVA` |

> **Poznámka k implementaci Spill range:** Funkce musí být zadána do buňky s dostatkem volného místa (vpravo i dolů). Pokud je rozsah blokován, Excel vrátí `#SPILL!`. U funkcí s dynamickým výstupem (`WELCH.TEST.2S`, `ANOVA`) závisí počet řádků na počtu skupin k a zvoleném post-hoc testu — viz dokumentace příslušné funkce. Šířka spill rozsahu je vždy fixní (uvedena v dokumentaci funkce) a nesmí být blokována.

### Formátování víc-sloupcových bloků

Víc-sloupcový spill strukturuje výstup do logických bloků oddělených záhlavím:

```
[ZÁHLAVÍ BLOKU]   (1. sloupec, ostatní prázdné)
[hlavička sl. 1]  [hlavička sl. 2]  ...
[data řádek 1]    [data řádek 1]    ...
[data řádek 2]    [data řádek 2]    ...
                                        ← prázdný řádek = vizuální oddělovač
[ZÁHLAVÍ BLOKU 2] ...
```

Záhlaví bloku jsou psána velkými písmeny, záhlaví sloupců jsou popisná. Implementace vloží záhlaví jako textové hodnoty bez formátování (formátování tabulky je zodpovědností volajícího / šablony sešitu).

### Konvence pro nepovinné argumenty
Nepovinné argumenty jsou označeny hranatými závorkami `[argument=výchozí_hodnota]`.

### Konvence pro název kritické hodnoty
Název buňky s kritickou hodnotou je dynamický a odráží skutečný kvantil použitého rozdělení:

| Situace | Název buňky |
|---------|-------------|
| Oboustranný test (`smer="two"`) | `"t₁₋α/₂"`, `"z₁₋α/₂"` |
| Jednostranný test (`smer="left"` nebo `"right"`) | `"t₁₋α"`, `"z₁₋α"` |
| Vždy pravostranné (F, χ²) | `"F₁₋α"`, `"χ²₁₋α"` |

> Poznámka k Unicode: ₁, ₋, ₂ jsou standardní znaky (U+2081, U+208B, U+2082). α zůstává v základní velikosti, protože subscript α v Unicode neexistuje — výsledný tvar `"t₁₋α/₂"` je přesto jednoznačný.

### Konvence pro váhy (suffix `.W`)
Funkce se suffixem `.W` přijímají argument `váhy` — rozsah nezáporných čísel stejné délky jako `hodnoty`. Platí:
- Váhy nemusí být normalizovány (funkce normalizuje interně: wᵢ' = wᵢ / Σwᵢ).
- Váha = 0 znamená vyloučení pozorování z výpočtu.
- Záporná váha → `#ČÍSLO!`
- Σwᵢ = 0 → `#ČÍSLO!`
- Prázdné buňky ve `váhy` jsou tratovány jako 0 (pozorování vyloučeno).
- Prázdné buňky ve `hodnoty` jsou ignorovány párově s odpovídající vahou.

### Řecká písmena a Unicode ve výstupu
UDF vrací hodnoty buněk jako Unicode řetězce. Řecká písmena (η, ω, μ, σ, α, …) a matematické symboly (², √, ≤, ∈, …) jsou plně podporovány — Excel je zobrazí správně ve všech standardních fontech (Calibri, Arial, Times New Roman). Názvy řádků výstupu proto používají přímo řecké znaky místo latinských přepisů (např. `"η²"` místo `"eta kvadrát"`).

---

## Funkce

---

### `NORM.DIST.RANGE`

**Syntaxe**
```
NORM.DIST.RANGE(střední_hodnota; smerodatna_odchylka; dolní_hranice; horní_hranice)
```

**Popis**  
Vrátí pravděpodobnost P(dolní_hranice ≤ X ≤ horní_hranice) pro náhodnou veličinu X s normálním rozdělením N(střední_hodnota, smerodatna_odchylka). Analogie funkce `BINOM.DIST.RANGE` pro normální (spojité) rozdělení. Kumulativní integrace je vždy předpokládána (bodová pravděpodobnost u spojitého rozdělení je nulová).

**Argumenty**

| Argument | Typ | Povinný | Popis |
|----------|-----|---------|-------|
| `střední_hodnota` | číslo | ✓ | Střední hodnota (μ) normálního rozdělení |
| `smerodatna_odchylka` | číslo > 0 | ✓ | Směrodatná odchylka (σ); musí být kladná |
| `dolní_hranice` | číslo | ✓ | Dolní mez intervalu. Pro −∞ zadat `""`  nebo `-1E+308` |
| `horní_hranice` | číslo | ✓ | Horní mez intervalu. Pro +∞ zadat `""` nebo `1E+308` |

**Výstup – Skalár**

Vrátí číslo v intervalu [0, 1] — pravděpodobnost, že X padne do zadaného intervalu.

**Chybové stavy**

| Chyba | Příčina |
|-------|---------|
| `#HODNOTA!` | Argumenty nejsou číselné |
| `#ČÍSLO!` | `smerodatna_odchylka` ≤ 0 |
| `#ČÍSLO!` | `dolní_hranice` > `horní_hranice` |

**Příklad**
```
=NORM.DIST.RANGE(0; 1; -1; 1)   → 0,6827  (pravděpodobnost intervalu ±1σ)
=NORM.DIST.RANGE(50; 10; ""; 60) → 0,8413  (P(X ≤ 60) pro N(50,10))
```

---

### `SHAPIRO.WILK`

**Syntaxe**
```
SHAPIRO.WILK(rozsah_hodnot)
```

**Popis**  
Provede Shapiro-Wilkův test normality na zadaném výběru. Testuje nulovou hypotézu H₀, že data pochází z normálního rozdělení. Implementace využívá Roystonovu aproximaci (1992), platnou pro rozsahy n = 3 až 5 000.

**Argumenty**

| Argument | Typ | Povinný | Popis |
|----------|-----|---------|-------|
| `rozsah_hodnot` | číselný rozsah (1 sloupec nebo 1 řádek) | ✓ | Hodnoty testovaného výběru; prázdné buňky jsou ignorovány |

**Výstup – Spill range (2 řádky × 2 sloupce)**

| Řádek | Levý sloupec (název) | Pravý sloupec (hodnota) |
|-------|----------------------|------------------------|
| 1 | `"W"` | Testová statistika W |
| 2 | `"p"` | p-hodnota testu |

**Chybové stavy**

| Chyba | Příčina |
|-------|---------|
| `#POČET!` | n < 3 nebo n > 5 000 |
| `#HODNOTA!` | Rozsah obsahuje nečíselné hodnoty (kromě prázdných buněk) |

**Příklad**
```
=SHAPIRO.WILK(A2:A31)
```
Výstup (rozlití od aktivní buňky):
```
W   │ 0,9741
p   │ 0,6082
```

---

### `KOLMOGOROV.SMIRNOV`

**Syntaxe**
```
KOLMOGOROV.SMIRNOV(rozsah_hodnot; [typ_rozdeleni="normal"])
```

**Popis**  
Provede Kolmogorov-Smirnovův test dobré shody (goodness-of-fit). Testuje, zda data pochází ze zadaného teoretického rozdělení. Parametry rozdělení jsou odhadnuty z dat metodou maximální věrohodnosti (MLE), pokud nejsou zadány explicitně. Jedná se o jednovýběrový KS test (srovnání s teoretickým rozdělením, nikoli dvou výběrů).

**Argumenty**

| Argument | Typ | Povinný | Popis |
|----------|-----|---------|-------|
| `rozsah_hodnot` | číselný rozsah (1 sloupec nebo 1 řádek) | ✓ | Hodnoty testovaného výběru; prázdné buňky jsou ignorovány |
| `typ_rozdeleni` | text | ne | Typ testovaného teoretického rozdělení (viz tabulka níže). Výchozí: `"normal"` |

**Podporované typy rozdělení pro `typ_rozdeleni`**

| Hodnota argumentu | Rozdělení | Odhadované parametry z dat |
|-------------------|-----------|---------------------------|
| `"normal"` (výchozí) | Normální N(μ, σ) | μ = výběrový průměr, σ = výběrová sm. odch. |
| `"lognormal"` | Lognormální | μ a σ na log-škále |
| `"exponential"` | Exponenciální Exp(λ) | λ = 1 / výběrový průměr |
| `"uniform"` | Rovnoměrné U(a, b) | a = min, b = max výběru |
| `"weibull"` | Weibullovo W(k, λ) | MLE odhad k a λ |

**Výstup – Spill range (2 řádky × 2 sloupce)**

| Řádek | Levý sloupec (název) | Pravý sloupec (hodnota) |
|-------|----------------------|------------------------|
| 1 | `"D"` | Testová statistika D |
| 2 | `"p"` | p-hodnota testu |

**Poznámka k implementaci**  
KS test má v základní podobě konzervativní p-hodnoty, pokud jsou parametry odhadnuty z dat (místo zadání a priori). Doporučujeme implementovat Lillieforsovu korekci pro případ `"normal"`, kde je tato korekce standardizovaná.

**Chybové stavy**

| Chyba | Příčina |
|-------|---------|
| `#HODNOTA!` | `typ_rozdeleni` není platná hodnota z enumerace |
| `#ČÍSLO!` | Data nesplňují podmínky rozdělení (např. záporné hodnoty pro `"lognormal"` nebo `"exponential"`) |
| `#POČET!` | n < 5 |

**Příklad**
```
=KOLMOGOROV.SMIRNOV(A2:A101)
=KOLMOGOROV.SMIRNOV(A2:A101; "lognormal")
```
Výstup (rozlití od aktivní buňky):
```
D   │ 0,0712
p   │ 0,3184
```

---

### `WELCH.TEST.2S`

**Syntaxe**
```
WELCH.TEST.2S(kategorie; hodnoty; [smer="two"]; [alpha=0,05])
```

**Popis**  
Provede Welchův dvouvýběrový t-test pro dva nezávislé výběry s potenciálně různými rozptyly (nepředpokládá homoskedasticitu, na rozdíl od Studentova t-testu). Vstupem jsou dva sousední sloupce — kategorický identifikátor skupiny a číselné hodnoty. Stupně volnosti jsou vypočteny Welch-Satterthwaitovou aproximací.

**Argumenty**

| Argument | Typ | Povinný | Popis |
|----------|-----|---------|-------|
| `kategorie` | textový nebo číselný rozsah | ✓ | Rozsah štítků skupin (do které ze dvou skupin jednotlivé záznamy patří). Musí mít právě 2 unikátní hodnoty. |
| `hodnoty` | číselný rozsah | ✓ | Rozsah číselných hodnot; musí mít stejnou délku jako `kategorie` |
| `smer` | text | ne | Alternativní hypotéza: `"two"` = oboustranný (výchozí), `"left"` = levostranný (H₁: μ₁ < μ₂), `"right"` = pravostranný (H₁: μ₁ > μ₂) |
| `alpha` | číslo ∈ (0, 1) | ne | Hladina významnosti. Výchozí: `0,05` |

**Výstup – Spill range (21 řádků × 2 sloupce)**

Skupiny jsou označeny hodnotami, které se vyskytují v `kategorie` (seřazeny abecedně / numericky; skupina 1 = první unikátní hodnota).

| Řádek | Levý sloupec (název) | Pravý sloupec (hodnota) |
|-------|----------------------|------------------------|
| 1 | `"=== Skupina: <název_1> ==="` | `""` |
| 2 | `"n₁"` | počet pozorování skupiny 1 |
| 3 | `"x̄₁"` | výběrový průměr skupiny 1 |
| 4 | `"med₁"` | výběrový medián skupiny 1 |
| 5 | `"s₁"` | výběrová směrodatná odchylka skupiny 1 |
| 6 | `"min₁"` | minimum skupiny 1 |
| 7 | `"max₁"` | maximum skupiny 1 |
| 8 | `"=== Skupina: <název_2> ==="` | `""` |
| 9 | `"n₂"` | počet pozorování skupiny 2 |
| 10 | `"x̄₂"` | výběrový průměr skupiny 2 |
| 11 | `"med₂"` | výběrový medián skupiny 2 |
| 12 | `"s₂"` | výběrová směrodatná odchylka skupiny 2 |
| 13 | `"min₂"` | minimum skupiny 2 |
| 14 | `"max₂"` | maximum skupiny 2 |
| 15 | `"=== Výsledky testu ==="` | `""` |
| 16 | `"t"` | testová statistika t |
| 17 | `"df"` | Welch-Satterthwaite df (nemusí být celé číslo) |
| 18 | `"t₁₋α"` nebo `"t₁₋α/₂"` | kritická hodnota t; název závisí na `smer` (viz konvence pro název kritické hodnoty) |
| 19 | `"p"` | p-hodnota testu |
| 20 | `"d"` | Cohenovo d |
| 21 | `"r"` | r = d / √(d² + 4); pro silně nevyvážené skupiny použít r = √(t² / (t² + df)) |

**Chybové stavy**

| Chyba | Příčina |
|-------|---------|
| `#HODNOTA!` | `smer` není `"two"`, `"left"`, nebo `"right"` |
| `#POČET!` | `kategorie` neobsahuje právě 2 unikátní hodnoty |
| `#POČET!` | Některá skupina má méně než 2 pozorování |
| `#DÉLKA!` | `kategorie` a `hodnoty` mají různou délku |
| `#ČÍSLO!` | `alpha` není v intervalu (0, 1) |

**Příklad**
```
=WELCH.TEST.2S(A2:A61; B2:B61)
=WELCH.TEST.2S(A2:A61; B2:B61; "right"; 0,01)
```

Příklad výstupu (rozlití od aktivní buňky):
```
=== Skupina: Kontrolní ===   │
n₁                           │ 30
x̄₁                           │ 48,3
med₁                         │ 47,5
s₁                           │ 6,12
min₁                         │ 35
max₁                         │ 62
=== Skupina: Experimentální ===  │
n₂                           │ 31
x̄₂                           │ 54,7
med₂                         │ 55,0
s₂                           │ 7,84
min₂                         │ 39
max₂                         │ 71
=== Výsledky testu ===       │
t                            │ -3,621
df                           │ 56,4
t₁₋α/₂                       │ ±2,003
p                            │ 0,0006
d                            │ -0,91
r                            │ -0,41
```

---

### `ANOVA`

**Syntaxe**
```
ANOVA(kategorie; hodnoty; [post_hoc="tukey"]; [alpha=0,05])
```

**Popis**  
Provede jednofaktorovou analýzu rozptylu (one-way ANOVA). Testuje nulovou hypotézu H₀, že střední hodnoty všech k skupin jsou stejné (μ₁ = μ₂ = ... = μₖ). Předpokládá normalitu reziduálů a homogenitu rozptylů — obojí je automaticky ověřeno a zahrnuto ve výstupu. Při zamítnutí H₀ provede zvolený post-hoc test pro identifikaci konkrétních rozdílných dvojic.

Výstup se skládá ze čtyř bloků (viz níže); šířka spill rozsahu je **7 sloupců**.

**Argumenty**

| Argument | Typ | Povinný | Popis |
|----------|-----|---------|-------|
| `kategorie` | textový nebo číselný rozsah | ✓ | Štítky skupin; musí mít ≥ 3 unikátní hodnoty |
| `hodnoty` | číselný rozsah | ✓ | Číselné hodnoty; stejná délka jako `kategorie` |
| `post_hoc` | text | ne | Zvolený post-hoc test (viz tabulka níže). Výchozí: `"tukey"`. Hodnota `"none"` post-hoc přeskočí. |
| `alpha` | číslo ∈ (0, 1) | ne | Hladina významnosti pro ANOVA i post-hoc. Výchozí: `0,05` |

**Podporované post-hoc testy**

| Hodnota `post_hoc` | Test | Vhodné použití |
|--------------------|------|----------------|
| `"tukey"` (výchozí) | Tukey HSD (Tukey-Kramer pro nevyvážené skupiny) | Všechny párové srovnání; standardní volba |
| `"bonferroni"` | Bonferroniho korekce | Malý počet plánovaných srovnání; konzervativnější |
| `"scheffe"` | Scheffého test | Neomezené kontrasty; nejkonzervativnější |
| `"games-howell"` | Games-Howellův test | Nevyvážené skupiny nebo heterogenní rozptyly (Levene p < α) |
| `"none"` | — | Bez post-hoc testu |

> **Doporučení k implementaci:** Pokud Levenův test vrátí p < `alpha`, automaticky upozornit ve výstupu a doporučit `"games-howell"`, protože ostatní testy předpokládají homogenitu rozptylů.

---

**Výstup – Spill range (dynamický počet řádků × 7 sloupců)**

Výstup tvoří čtyři bloky oddělené prázdnými řádky. Počet řádků závisí na počtu skupin k:
- Blok 1 (popisné statistiky): k + 2 řádky
- Blok 2 (ANOVA tabulka + předpoklady): 8 řádků (fixní)
- Blok 3 (velikost účinku): 4 řádky (fixní)
- Blok 4 (post-hoc): k*(k−1)/2 + 2 řádky (0 pokud `post_hoc="none"`)

---

**Blok 1 – Popisné statistiky (šířka: 7 sloupců)**

| Sloupec 1 | Sloupec 2 | Sloupec 3 | Sloupec 4 | Sloupec 5 | Sloupec 6 | Sloupec 7 |
|-----------|-----------|-----------|-----------|-----------|-----------|-----------|
| `"POPISNÉ STATISTIKY"` | | | | | | |
| `"Skupina"` | `"n"` | `"x̄"` | `"med"` | `"s"` | `"min"` | `"max"` |
| `<název skupiny 1>` | n₁ | x̄₁ | med₁ | s₁ | min₁ | max₁ |
| `<název skupiny 2>` | n₂ | x̄₂ | med₂ | s₂ | min₂ | max₂ |
| … | … | … | … | … | … | … |
| `<název skupiny k>` | nₖ | x̄ₖ | medₖ | sₖ | minₖ | maxₖ |
| *(prázdný řádek)* | | | | | | |

---

**Blok 2 – ANOVA tabulka a ověření předpokladů (šířka: 6 sloupců, 7. prázdný)**

| Sl. 1 | Sl. 2 | Sl. 3 | Sl. 4 | Sl. 5 | Sl. 6 |
|-------|-------|-------|-------|-------|-------|
| `"ANOVA"` | | | | | |
| `"Zdroj variability"` | `"SS"` | `"df"` | `"MS"` | `"F"` | `"p-hodnota"` |
| `"Mezi skupinami"` | SS_b | k−1 | MS_b | F | p |
| `"Uvnitř skupin (rezidua)"` | SS_w | N−k | MS_w | | |
| `"Celkem"` | SS_t | N−1 | | | |
| `"F₁₋α"` | F_crit | | | | |
| *(prázdný řádek)* | | | | | |
| `"OVĚŘENÍ PŘEDPOKLADŮ"` | | | | | |
| `"Levenův test (homogenita rozptylů)"` | `"F"` | `"df₁"` | `"df₂"` | `"p-hodnota"` | `"⚠ heterogenní rozptyly"` |
| | L_stat | k−1 | N−k | L_p | `TRUE` nebo `FALSE` |
| *(prázdný řádek)* | | | | | |

---

**Blok 3 – Velikost účinku (šířka: 2 sloupce)**

| Sloupec 1 | Sloupec 2 |
|-----------|-----------|
| `"VELIKOST ÚČINKU"` | |
| `"η²"` | η² = SS_b / SS_t |
| `"ω²"` | ω² = (SS_b − df_b·MS_w) / (SS_t + MS_w) |
| `"f"` | f = √(η² / (1−η²)) |
| *(prázdný řádek)* | |

---

**Blok 4 – Post-hoc test (šířka: 5 sloupců; vynechán pokud `post_hoc="none"`)**

| Sl. 1 | Sl. 2 | Sl. 3 | Sl. 4 | Sl. 5 |
|-------|-------|-------|-------|-------|
| `"POST-HOC: <NÁZEV TESTU>"` | | | | |
| `"Skupina A"` | `"Skupina B"` | `"Δ průměrů (A−B)"` | `"p-hodnota"` | `"Sig."` |
| `<skupina i>` | `<skupina j>` | x̄ᵢ − x̄ⱼ | p_ij | `"***"` / ``"**"` / `"*"` / `"ns"` |
| … (všechny páry k*(k−1)/2) | | | | |

**Konvence pro sloupec `Sig.`:**

| Značka | Podmínka |
|--------|----------|
| `***` | p < 0,001 |
| `**` | p < 0,01 |
| `*` | p < α (zadaná hodnota) |
| `ns` | p ≥ α (not significant) |

---

**Chybové stavy**

| Chyba | Příčina |
|-------|---------|
| `#POČET!` | `kategorie` má méně než 3 unikátní hodnoty (pro 2 skupiny použijte `WELCH.TEST.2S`) |
| `#DÉLKA!` | `kategorie` a `hodnoty` mají různou délku |
| `#POČET!` | Některá skupina má méně než 2 pozorování |
| `#HODNOTA!` | `post_hoc` není platná hodnota z enumerace |
| `#ČÍSLO!` | `alpha` není v intervalu (0, 1) |

**Příklad**
```
=ANOVA(A2:A91; B2:B91)
=ANOVA(A2:A91; B2:B91; "games-howell"; 0,01)
=ANOVA(A2:A91; B2:B91; "none")
```

Příklad výstupu (k = 3 skupiny, post-hoc Tukey, zkráceno):
```
POPISNÉ STATISTIKY
Skupina        n    x̄       med     s      min   max
Kontrolní      30   48,3    47,5    6,12   35    62
Dávka nízká    30   53,1    52,0    7,04   38    68
Dávka vysoká   30   61,8    62,5    6,88   47    77

ANOVA
Zdroj variability       SS        df   MS        F       p
Mezi skupinami          2841,4    2    1420,7    29,84   < 0,0001
Uvnitř skupin (rez.)    4142,1    87   47,61
Celkem                  6983,5    89
F₁₋α                    3,101

OVĚŘENÍ PŘEDPOKLADŮ
Levenův test            F        df₁  df₂  p       ⚠ heterogenní rozptyly
                        0,312    2    87   0,7325  FALSE

VELIKOST ÚČINKU
η²   0,407
ω²   0,391
f    0,828

POST-HOC: TUKEY HSD
Skupina A       Skupina B      Δ průměrů   p-hodnota   Sig.
Kontrolní       Dávka nízká    -4,8        0,0241      *
Kontrolní       Dávka vysoká   -13,5       < 0,0001    ***
Dávka nízká     Dávka vysoká   -8,7        0,0003      ***
```

---

### `CORREL.SPEARMAN`

**Syntaxe**
```
CORREL.SPEARMAN(rozsah_x; rozsah_y; [smer="two"]; [alpha=0,05])
```

**Popis**  
Vypočítá Spearmanův korelační koeficient ρ a provede test jeho statistické významnosti. Výpočet probíhá jako Pearsonova korelace nad pořadími obou proměnných — tento přístup správně zpracovává shody v hodnotách (ties) přiřazením průměrného pořadí (midrank metoda), zatímco zjednodušený vzorec ρ = 1 − 6Σd²/n(n²−1) při ties zkresluje výsledek.

Významnost se testuje přes t-aproximaci: t = ρ√(n−2) / √(1−ρ²), df = n−2. Aproximace je spolehlivá pro n ≥ 10.

**Argumenty**

| Argument | Typ | Povinný | Popis |
|----------|-----|---------|-------|
| `rozsah_x` | číselný rozsah (1 sloupec nebo 1 řádek) | ✓ | První proměnná; prázdné buňky jsou ignorovány párově s `rozsah_y` |
| `rozsah_y` | číselný rozsah (1 sloupec nebo 1 řádek) | ✓ | Druhá proměnná; musí mít stejnou délku jako `rozsah_x` |
| `smer` | text | ne | Alternativní hypotéza: `"two"` = oboustranný (výchozí), `"left"` = levostranný (H₁: ρ < 0), `"right"` = pravostranný (H₁: ρ > 0) |
| `alpha` | číslo ∈ (0, 1) | ne | Hladina významnosti. Výchozí: `0,05` |

**Výstup – Spill range (6 řádků × 2 sloupce)**

| Řádek | Levý sloupec (název) | Pravý sloupec (hodnota) |
|-------|----------------------|------------------------|
| 1 | `"ρ"` | Spearmanův korelační koeficient ∈ (−1, 1) |
| 2 | `"n"` | počet párů (po vyloučení prázdných hodnot) |
| 3 | `"t"` | testová statistika |
| 4 | `"df"` | n−2 |
| 5 | `"t₁₋α"` nebo `"t₁₋α/₂"` | kritická hodnota t; název závisí na `smer` (viz konvence pro název kritické hodnoty) |
| 6 | `"p"` | p-hodnota testu |

**Poznámka k implementaci**  
Párové vyloučení prázdných hodnot: pokud je buňka prázdná v `rozsah_x` nebo `rozsah_y`, vyloučí se odpovídající pozice v obou rozsazích před výpočtem pořadí. Pořadí se přiřazují zvlášť pro x a zvlášť pro y po vyloučení. Pro n < 10 upozornit na nespolehlivost t-aproximace (vrátit výsledek, ale přidat `#WARN` nebo obdobný příznak — dle konvence doplňku).

**Chybové stavy**

| Chyba | Příčina |
|-------|---------|
| `#DÉLKA!` | `rozsah_x` a `rozsah_y` mají různou délku |
| `#POČET!` | n < 3 po vyloučení prázdných hodnot |
| `#ČÍSLO!` | Rozptyl pořadí jedné z proměnných je nulový (všechny hodnoty identické) |
| `#HODNOTA!` | `smer` není `"two"`, `"left"`, nebo `"right"` |
| `#ČÍSLO!` | `alpha` není v intervalu (0, 1) |

**Příklad**
```
=CORREL.SPEARMAN(A2:A51; B2:B51)
=CORREL.SPEARMAN(A2:A51; B2:B51; "right")
```
Výstup (rozlití od aktivní buňky):
```
ρ    │ 0,7134
n    │ 50
t    │ 7,081
df   │ 48
t₁₋α/₂  │ ±2,011
p    │ < 0,0001
```

---

### `AVERAGE.W`

**Syntaxe**
```
AVERAGE.W(hodnoty; váhy)
```

**Popis**  
Vážený aritmetický průměr: x̄_w = Σ(wᵢ·xᵢ) / Σwᵢ.

**Argumenty**

| Argument | Typ | Povinný | Popis |
|----------|-----|---------|-------|
| `hodnoty` | číselný rozsah | ✓ | Pozorování |
| `váhy` | číselný rozsah ≥ 0 | ✓ | Váhy; viz konvence pro váhy |

**Výstup – Skalár.** Vážený průměr.

**Chybové stavy**

| Chyba | Příčina |
|-------|---------|
| `#DÉLKA!` | `hodnoty` a `váhy` mají různou délku |
| `#ČÍSLO!` | Záporná váha nebo Σwᵢ = 0 |

**Příklad**
```
=AVERAGE.W(A2:A11; B2:B11)
```

---

### `HARMEAN.W`

**Syntaxe**
```
HARMEAN.W(hodnoty; váhy)
```

**Popis**  
Vážený harmonický průměr: H_w = Σwᵢ / Σ(wᵢ / xᵢ).

**Argumenty**

| Argument | Typ | Povinný | Popis |
|----------|-----|---------|-------|
| `hodnoty` | číselný rozsah > 0 | ✓ | Pozorování; všechny hodnoty musí být kladné |
| `váhy` | číselný rozsah ≥ 0 | ✓ | Váhy; viz konvence pro váhy |

**Výstup – Skalár.** Vážený harmonický průměr.

**Chybové stavy**

| Chyba | Příčina |
|-------|---------|
| `#DÉLKA!` | `hodnoty` a `váhy` mají různou délku |
| `#ČÍSLO!` | Záporná váha nebo Σwᵢ = 0 |
| `#ČÍSLO!` | Některá hodnota ≤ 0 (s nenulovou vahou) |

**Příklad**
```
=HARMEAN.W(A2:A11; B2:B11)
```

---

### `GEOMEAN.W`

**Syntaxe**
```
GEOMEAN.W(hodnoty; váhy)
```

**Popis**  
Vážený geometrický průměr: G_w = exp(Σ(wᵢ · ln xᵢ) / Σwᵢ).

**Argumenty**

| Argument | Typ | Povinný | Popis |
|----------|-----|---------|-------|
| `hodnoty` | číselný rozsah > 0 | ✓ | Pozorování; všechny hodnoty musí být kladné |
| `váhy` | číselný rozsah ≥ 0 | ✓ | Váhy; viz konvence pro váhy |

**Výstup – Skalár.** Vážený geometrický průměr.

**Chybové stavy**

| Chyba | Příčina |
|-------|---------|
| `#DÉLKA!` | `hodnoty` a `váhy` mají různou délku |
| `#ČÍSLO!` | Záporná váha nebo Σwᵢ = 0 |
| `#ČÍSLO!` | Některá hodnota ≤ 0 (s nenulovou vahou) |

**Příklad**
```
=GEOMEAN.W(A2:A11; B2:B11)
```

---

### `VAR.P.W`

**Syntaxe**
```
VAR.P.W(hodnoty; váhy)
```

**Popis**  
Vážený populační rozptyl: σ²_w = Σwᵢ(xᵢ − x̄_w)² / Σwᵢ.

**Argumenty**

| Argument | Typ | Povinný | Popis |
|----------|-----|---------|-------|
| `hodnoty` | číselný rozsah | ✓ | Pozorování |
| `váhy` | číselný rozsah ≥ 0 | ✓ | Váhy; viz konvence pro váhy |

**Výstup – Skalár.** Vážený populační rozptyl.

**Chybové stavy**

| Chyba | Příčina |
|-------|---------|
| `#DÉLKA!` | `hodnoty` a `váhy` mají různou délku |
| `#ČÍSLO!` | Záporná váha nebo Σwᵢ = 0 |

**Příklad**
```
=VAR.P.W(A2:A11; B2:B11)
```

---

### `VAR.S.W`

**Syntaxe**
```
VAR.S.W(hodnoty; váhy)
```

**Popis**  
Vážený výběrový rozptyl s Besselovou korekcí: s²_w = Σwᵢ(xᵢ − x̄_w)² / (Σwᵢ − 1). Předpokládá frekvenční váhy (reliability weights), kde Σwᵢ odpovídá efektivní velikosti výběru.

**Argumenty**

| Argument | Typ | Povinný | Popis |
|----------|-----|---------|-------|
| `hodnoty` | číselný rozsah | ✓ | Pozorování |
| `váhy` | číselný rozsah ≥ 0 | ✓ | Váhy; viz konvence pro váhy |

**Výstup – Skalár.** Vážený výběrový rozptyl.

**Chybové stavy**

| Chyba | Příčina |
|-------|---------|
| `#DÉLKA!` | `hodnoty` a `váhy` mají různou délku |
| `#ČÍSLO!` | Záporná váha nebo Σwᵢ = 0 |
| `#POČET!` | Σwᵢ ≤ 1 (jmenovatel ≤ 0) |

**Příklad**
```
=VAR.S.W(A2:A11; B2:B11)
```

---

### `STDEV.P.W`

**Syntaxe**
```
STDEV.P.W(hodnoty; váhy)
```

**Popis**  
Vážená populační směrodatná odchylka: σ_w = √VAR.P.W(hodnoty; váhy).

**Argumenty**

| Argument | Typ | Povinný | Popis |
|----------|-----|---------|-------|
| `hodnoty` | číselný rozsah | ✓ | Pozorování |
| `váhy` | číselný rozsah ≥ 0 | ✓ | Váhy; viz konvence pro váhy |

**Výstup – Skalár.** Vážená populační směrodatná odchylka.

**Chybové stavy** — shodné s `VAR.P.W`.

**Příklad**
```
=STDEV.P.W(A2:A11; B2:B11)
```

---

### `STDEV.S.W`

**Syntaxe**
```
STDEV.S.W(hodnoty; váhy)
```

**Popis**  
Vážená výběrová směrodatná odchylka: s_w = √VAR.S.W(hodnoty; váhy).

**Argumenty**

| Argument | Typ | Povinný | Popis |
|----------|-----|---------|-------|
| `hodnoty` | číselný rozsah | ✓ | Pozorování |
| `váhy` | číselný rozsah ≥ 0 | ✓ | Váhy; viz konvence pro váhy |

**Výstup – Skalár.** Vážená výběrová směrodatná odchylka.

**Chybové stavy** — shodné s `VAR.S.W`.

**Příklad**
```
=STDEV.S.W(A2:A11; B2:B11)
```

---

### `VARCOEF`

**Syntaxe**
```
VARCOEF(hodnoty)
```

**Popis**  
Populační variační koeficient: CV = σ / μ. Bezrozměrná míra relativní variability. Výsledek je vrácen jako desetinné číslo (nikoli procento); případné násobení ×100 je na volajícím.

**Argumenty**

| Argument | Typ | Povinný | Popis |
|----------|-----|---------|-------|
| `hodnoty` | číselný rozsah | ✓ | Pozorování; prázdné buňky jsou ignorovány |

**Výstup – Skalár.** σ / μ.

**Chybové stavy**

| Chyba | Příčina |
|-------|---------|
| `#DĚLENÍ0!` | μ = 0 |
| `#POČET!` | n < 1 |

**Příklad**
```
=VARCOEF(A2:A51)         → 0,1823  (tj. ~18,2 %)
```

---

### `VARCOEF.S`

**Syntaxe**
```
VARCOEF.S(hodnoty)
```

**Popis**  
Výběrový variační koeficient: CV = s / x̄ (s = výběrová směrodatná odchylka s dělením n−1).

**Argumenty**

| Argument | Typ | Povinný | Popis |
|----------|-----|---------|-------|
| `hodnoty` | číselný rozsah | ✓ | Pozorování; prázdné buňky jsou ignorovány |

**Výstup – Skalár.** s / x̄.

**Chybové stavy**

| Chyba | Příčina |
|-------|---------|
| `#DĚLENÍ0!` | x̄ = 0 |
| `#POČET!` | n < 2 |

**Příklad**
```
=VARCOEF.S(A2:A51)
```

---

### `VARCOEF.S.W`

**Syntaxe**
```
VARCOEF.S.W(hodnoty; váhy)
```

**Popis**  
Vážený výběrový variační koeficient: CV_w = s_w / x̄_w, kde s_w = STDEV.S.W a x̄_w = AVERAGE.W.

**Argumenty**

| Argument | Typ | Povinný | Popis |
|----------|-----|---------|-------|
| `hodnoty` | číselný rozsah | ✓ | Pozorování |
| `váhy` | číselný rozsah ≥ 0 | ✓ | Váhy; viz konvence pro váhy |

**Výstup – Skalár.** s_w / x̄_w.

**Chybové stavy**

| Chyba | Příčina |
|-------|---------|
| `#DÉLKA!` | `hodnoty` a `váhy` mají různou délku |
| `#ČÍSLO!` | Záporná váha nebo Σwᵢ = 0 |
| `#POČET!` | Σwᵢ ≤ 1 |
| `#DĚLENÍ0!` | x̄_w = 0 |

**Příklad**
```
=VARCOEF.S.W(A2:A11; B2:B11)
```

---

### `PERCENTILE.INC.IFS`

**Syntaxe**
```
PERCENTILE.INC.IFS(hodnoty; kvantil; [rozsah_kriteria_1; kriterium_1; rozsah_kriteria_2; kriterium_2; …])
```

**Popis**  
Vrátí k-tý percentil hodnot filtrovaných podle zadaných kritérií. Interpolační metoda je shodná s `PERCENTILE.INC` (Excel): kvantil k ∈ [0, 1], pořadová pozice = k·(n−1), lineární interpolace. Kriteriální logika odpovídá `SUMIFS` — páry (rozsah_kriteria; kriterium) jsou nepovinné a lze jich zadat libovolný počet; výsledkem je průnik všech podmínek (AND). Bez kritérií se chová identicky jako `PERCENTILE.INC`.

**Argumenty**

| Argument | Typ | Povinný | Popis |
|----------|-----|---------|-------|
| `hodnoty` | číselný rozsah | ✓ | Zdrojová data pro výpočet percentilu |
| `kvantil` | číslo ∈ [0, 1] | ✓ | Požadovaný kvantil |
| `rozsah_kriteria_N` | rozsah | ne | Rozsah, na který se aplikuje N-té kriterium; musí mít stejnou délku jako `hodnoty` |
| `kriterium_N` | hodnota nebo text | ne | Podmínka ve formátu SUMIFS: přesná shoda, operátory (`">5"`, `"<=10"`, `"<>0"`), zástupné znaky (`"A*"`, `"?B"`) |

> Volitelné argumenty musí být zadány vždy v párech `(rozsah_kriteria_N; kriterium_N)`. Lichý počet volitelných argumentů → `#HODNOTA!`.

**Výstup – Skalár.** k-tý percentil filtrovaných hodnot.

**Chybové stavy**

| Chyba | Příčina |
|-------|---------|
| `#ČÍSLO!` | `kvantil` < 0 nebo > 1 |
| `#HODNOTA!` | Lichý počet volitelných argumentů |
| `#DÉLKA!` | Některý `rozsah_kriteria_N` má jinou délku než `hodnoty` |
| `#NENÍ_K_DISP!` | Po aplikaci filtrů nezůstane žádná hodnota |

**Příklad**
```
' Medián bez filtrů — ekvivalent PERCENTILE.INC(A2:A101; 0,5)
=PERCENTILE.INC.IFS(A2:A101; 0,5)

' 75. percentil pro skupinu "A" s hodnotou > 10
=PERCENTILE.INC.IFS(A2:A101; 0,75; B2:B101; "A"; A2:A101; ">10")

' 90. percentil pro rok 2023
=PERCENTILE.INC.IFS(A2:A101; 0,9; C2:C101; 2023)
```

---

### `PERCENTILE.EXC.IFS`

**Syntaxe**
```
PERCENTILE.EXC.IFS(hodnoty; kvantil; [rozsah_kriteria_1; kriterium_1; rozsah_kriteria_2; kriterium_2; …])
```

**Popis**  
Identické chování jako `PERCENTILE.INC.IFS`, ale s interpolační metodou `PERCENTILE.EXC`: kvantil k ∈ (0, 1), pořadová pozice = k·(n+1)−1, exkluduje krajní hodnoty při extrapolaci. Vyžaduje n ≥ 1/k a n ≥ 1/(1−k) (jinak by byl výsledek mimo rozsah dat).

**Argumenty**

Shodné s `PERCENTILE.INC.IFS`; jediný rozdíl je v povoleném rozsahu `kvantil`: musí být striktně v (0, 1).

| Argument | Typ | Povinný | Popis |
|----------|-----|---------|-------|
| `hodnoty` | číselný rozsah | ✓ | Zdrojová data pro výpočet percentilu |
| `kvantil` | číslo ∈ (0, 1) | ✓ | Požadovaný kvantil; striktně exkluzivní krajní hodnoty |
| `rozsah_kriteria_N` | rozsah | ne | viz `PERCENTILE.INC.IFS` |
| `kriterium_N` | hodnota nebo text | ne | viz `PERCENTILE.INC.IFS` |

**Výstup – Skalár.** k-tý percentil filtrovaných hodnot (EXC metoda).

**Chybové stavy**

| Chyba | Příčina |
|-------|---------|
| `#ČÍSLO!` | `kvantil` ≤ 0 nebo ≥ 1 |
| `#ČÍSLO!` | n po filtraci příliš malé pro zadaný kvantil (extrapolace mimo data) |
| `#HODNOTA!` | Lichý počet volitelných argumentů |
| `#DÉLKA!` | Některý `rozsah_kriteria_N` má jinou délku než `hodnoty` |
| `#NENÍ_K_DISP!` | Po aplikaci filtrů nezůstane žádná hodnota |

**Příklad**
```
=PERCENTILE.EXC.IFS(A2:A101; 0,5)
=PERCENTILE.EXC.IFS(A2:A101; 0,25; B2:B101; "Praha"; C2:C101; ">="&DATE(2023;1;1))
```

---

### `T.TEST.1S`

**Syntaxe**
```
T.TEST.1S(hodnoty; mu_0; [smer="two"]; [alpha=0,05])
```

**Popis**  
Jednovýběrový t-test. Testuje H₀: μ = μ₀ oproti alternativě dané argumentem `smer`. Testová statistika: t = (x̄ − μ₀) / (s / √n), df = n − 1.

**Argumenty**

| Argument | Typ | Povinný | Popis |
|----------|-----|---------|-------|
| `hodnoty` | číselný rozsah | ✓ | Výběrová data; prázdné buňky jsou ignorovány |
| `mu_0` | číslo | ✓ | Hypotetická populační střední hodnota μ₀ |
| `smer` | text | ne | `"two"` = oboustranný (výchozí), `"left"` = levostranný (H₁: μ < μ₀), `"right"` = pravostranný (H₁: μ > μ₀) |
| `alpha` | číslo ∈ (0, 1) | ne | Hladina významnosti. Výchozí: `0,05` |

**Výstup – Spill range (8 řádků × 2 sloupce)**

| Řádek | Název | Hodnota |
|-------|-------|---------|
| 1 | `"x̄"` | výběrový průměr |
| 2 | `"μ₀"` | hypotetická střední hodnota |
| 3 | `"s"` | výběrová směrodatná odchylka |
| 4 | `"n"` | počet pozorování |
| 5 | `"t"` | testová statistika |
| 6 | `"df"` | n − 1 |
| 7 | `"t₁₋α"` nebo `"t₁₋α/₂"` | kritická hodnota t; název závisí na `smer` (viz konvence pro název kritické hodnoty) |
| 8 | `"p"` | p-hodnota |

**Chybové stavy**

| Chyba | Příčina |
|-------|---------|
| `#POČET!` | n < 2 |
| `#HODNOTA!` | `smer` není `"two"`, `"left"`, nebo `"right"` |
| `#ČÍSLO!` | `alpha` není v intervalu (0, 1) |

**Příklad**
```
=T.TEST.1S(A2:A31; 50)
=T.TEST.1S(A2:A31; 50; "right"; 0,01)
```
Výstup:
```
x̄   │ 53,2
μ₀  │ 50
s   │ 8,14
n   │ 30
t   │ 2,153
df  │ 29
t₁₋α/₂  │ ±2,045
p   │ 0,0398
```

---

### `PROP.TEST.1S`

**Syntaxe**
```
PROP.TEST.1S(hodnoty; pi_0; [smer="two"]; [alpha=0,05])
```

**Popis**  
Jednovýběrový z-test pro populační podíl. Testuje H₀: π = π₀. Vstupem je rozsah binárních hodnot (0/1 nebo FALSE/TRUE), kde 1/TRUE = úspěch. Testová statistika: z = (p̂ − π₀) / √(π₀(1−π₀)/n). Používá normální aproximaci — viz poznámka k implementaci.

**Argumenty**

| Argument | Typ | Povinný | Popis |
|----------|-----|---------|-------|
| `hodnoty` | rozsah binárních hodnot (0/1, TRUE/FALSE) | ✓ | Výběrová data; prázdné buňky jsou ignorovány |
| `pi_0` | číslo ∈ (0, 1) | ✓ | Hypotetický populační podíl π₀ |
| `smer` | text | ne | `"two"` (výchozí), `"left"` (H₁: π < π₀), `"right"` (H₁: π > π₀) |
| `alpha` | číslo ∈ (0, 1) | ne | Hladina významnosti. Výchozí: `0,05` |

**Výstup – Spill range (7 řádků × 2 sloupce)**

| Řádek | Název | Hodnota |
|-------|-------|---------|
| 1 | `"p̂"` | výběrový podíl x/n |
| 2 | `"π₀"` | hypotetický podíl |
| 3 | `"x"` | počet úspěchů |
| 4 | `"n"` | celkový počet pozorování |
| 5 | `"z"` | testová statistika |
| 6 | `"z₁₋α"` nebo `"z₁₋α/₂"` | kritická hodnota z; název závisí na `smer` (viz konvence pro název kritické hodnoty) |
| 7 | `"p"` | p-hodnota |

**Poznámka k implementaci**  
Normální aproximace je spolehlivá při n·π₀ ≥ 5 a n·(1−π₀) ≥ 5. Pokud tato podmínka není splněna, vrátit výsledek a zároveň přidat varovný příznak (dle konvence doplňku). Pro agregovaná vstupní data (znám pouze x a n, ne individuální záznamy) uživatel sestaví sloupec x jedniček a n−x nul.

**Chybové stavy**

| Chyba | Příčina |
|-------|---------|
| `#POČET!` | n < 1 |
| `#ČÍSLO!` | `pi_0` ∉ (0, 1) |
| `#HODNOTA!` | `hodnoty` obsahují jiné hodnoty než 0, 1, TRUE, FALSE |
| `#HODNOTA!` | `smer` není `"two"`, `"left"`, nebo `"right"` |
| `#ČÍSLO!` | `alpha` není v intervalu (0, 1) |

**Příklad**
```
=PROP.TEST.1S(A2:A201; 0,3)
=PROP.TEST.1S(A2:A201; 0,3; "right"; 0,01)
```
Výstup:
```
p̂   │ 0,365
π₀  │ 0,3
x   │ 73
n   │ 200
z   │ 2,121
z₁₋α/₂  │ ±1,960
p   │ 0,0340
```

---

### `CHISQ.GOF`

**Syntaxe**
```
CHISQ.GOF(pozorované; očekávané; [kategorie]; [alpha=0,05])
```

**Popis**  
Chí-kvadrát test dobré shody (goodness-of-fit). Testuje H₀, že pozorované četnosti odpovídají zadanému teoretickému rozdělení. Testová statistika: χ² = Σ(Oᵢ − Eᵢ)² / Eᵢ, df = k − 1.

Argument `očekávané` přijímá buď přímo očekávané četnosti (Σ Eᵢ = Σ Oᵢ), nebo pravděpodobnosti (Σ pᵢ = 1) — funkce detekuje typ automaticky podle součtu a převede pravděpodobnosti na četnosti: Eᵢ = pᵢ · Σ Oᵢ.

**Argumenty**

| Argument | Typ | Povinný | Popis |
|----------|-----|---------|-------|
| `pozorované` | číselný rozsah celých čísel ≥ 0 | ✓ | Pozorované četnosti kategorií (k hodnot) |
| `očekávané` | číselný rozsah > 0 | ✓ | Očekávané četnosti nebo pravděpodobnosti (k hodnot; stejná délka jako `pozorované`) |
| `kategorie` | textový nebo číselný rozsah | ne | Popisky kategorií (k hodnot); není-li zadán, použijí se indexy 1, 2, … k |
| `alpha` | číslo ∈ (0, 1) | ne | Hladina významnosti. Výchozí: `0,05` |

**Výstup – Spill range (dynamický × 4 sloupce)**

Výstup tvoří dva bloky; šířka je fixně **4 sloupce**.

**Blok 1 – Výsledky testu (fixní, 6 řádků)**

| Sl. 1 | Sl. 2 | Sl. 3 | Sl. 4 |
|-------|-------|-------|-------|
| `"χ² GOF"` | | | |
| `"χ²"` | hodnota | | |
| `"df"` | k − 1 | | |
| `"χ²α"` | kritická hodnota | | |
| `"p"` | p-hodnota | | |
| *(prázdný řádek)* | | | |

**Blok 2 – Příspěvky kategorií (k + 2 řádků)**

| Sl. 1 | Sl. 2 | Sl. 3 | Sl. 4 |
|-------|-------|-------|-------|
| `"KATEGORIE"` | | | |
| `"Kategorie"` | `"O"` | `"E"` | `"(O−E)²/E"` |
| `<kat 1>` | O₁ | E₁ | (O₁−E₁)²/E₁ |
| … | … | … | … |
| `<kat k>` | Oₖ | Eₖ | (Oₖ−Eₖ)²/Eₖ |

**Poznámka k implementaci**  
Standardní pravidlo spolehlivosti aproximace: všechna Eᵢ ≥ 5. Pokud některé Eᵢ < 5, vrátit výsledek a přidat varovný příznak. df = k − 1 platí při plně specifikovaném očekávaném rozdělení (bez odhadovaných parametrů); pokud jsou parametry rozdělení odhadnuty z dat, df = k − 1 − m, kde m = počet odhadovaných parametrů — tuto korekci implementátor aplikuje manuálně nebo přidá volitelný argument `[pocet_parametru=0]`.

**Chybové stavy**

| Chyba | Příčina |
|-------|---------|
| `#DÉLKA!` | `pozorované`, `očekávané` (nebo `kategorie`) mají různou délku |
| `#POČET!` | k < 2 |
| `#ČÍSLO!` | Některé Oᵢ < 0 nebo není celé číslo |
| `#ČÍSLO!` | Některé Eᵢ ≤ 0 |
| `#ČÍSLO!` | `alpha` není v intervalu (0, 1) |

**Příklad**
```
' Pravděpodobnosti (automaticky převedeny na četnosti)
=CHISQ.GOF(A2:A6; B2:B6; C2:C6)

' Bez popisků kategorií
=CHISQ.GOF(A2:A6; B2:B6)
```
Výstup (k = 4, s popisky):
```
χ² GOF
χ²      │ 6,342   │         │
df      │ 3       │         │
χ²α     │ 7,815   │         │
p       │ 0,0962  │         │

KATEGORIE
Kategorie  │ O    │ E     │ (O−E)²/E
Červená    │ 42   │ 50    │ 1,280
Modrá      │ 63   │ 50    │ 3,380
Zelená     │ 48   │ 50    │ 0,080
Žlutá      │ 47   │ 50    │ 0,180
```

---

## Navrhovaná rozšíření (backlog)

Níže jsou funkce, které by logicky doplnily sadu:

| Funkce | Popis |
|--------|-------|
| `LEVENE.TEST(kategorie; hodnoty)` | Test homogenity rozptylů jako samostatná funkce (je součástí výstupu `ANOVA`, ale může být potřeba i před `WELCH.TEST.2S`) |
| `MANN.WHITNEY(kategorie; hodnoty; [smer]; [alpha])` | Neparametrická alternativa k Welchovu testu pro data bez normality; výstup: U statistika, p-hodnota, rank-biserial correlation |
| `KRUSKAL.WALLIS(kategorie; hodnoty; [post_hoc]; [alpha])` | Neparametrická alternativa k `ANOVA`; vhodné jako záložní test při zamítnutí normality ve skupinách |
| `SHAPIRO.WILK.GROUP(kategorie; hodnoty)` | Shapiro-Wilkův test zvlášť pro každou skupinu — předzpracovací krok před `WELCH.TEST.2S` nebo `ANOVA` |
| `NORM.DIST.RANGE.2T(stredni_hodnota; smerodatna_odchylka; hranice)` | Symetrická verze — vrátí P(μ−hranice ≤ X ≤ μ+hranice) |

---

## Změnová historie

| Verze | Datum | Změna |
|-------|-------|-------|
| 1.0 | — | Počáteční návrh (DOCX) |
| 1.1 | — | Rozšíření argumentů, definice spill range výstupů, chybové stavy, navrhovaná rozšíření |
| 1.4 | — | Výstupní popisky přepsány na symboly (x̄, med, s, tα, Fα, d, r, p, …); odstraněny interpretační poznámky pro studenty; blok velikosti účinku ANOVA redukován na 2 sloupce |
| 1.8 | — | Přidány funkce `T.TEST.1S`, `PROP.TEST.1S`, `CHISQ.GOF` |
| 1.9 | — | Přejmenování `VAR.X` → `VARCOEF`, `VAR.X.S` → `VARCOEF.S`, `VAR.X.S.W` → `VARCOEF.S.W` |
