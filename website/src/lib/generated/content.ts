export type GeneratedDoc = {
  lang: 'cs' | 'en';
  slug: string;
  section: string;
  title: string;
  summary: string;
  sourcePath: string;
  html: string;
};

export type GeneratedInstaller = {
  lang: 'cs' | 'en';
  label: string;
  fileName: string;
  href: string;
  exists: boolean;
};

export type FunctionIndexEntry = {
  lang: 'cs' | 'en';
  name: string;
  summary: string;
  href: string;
};

export const generatedAt = "2026-04-24T16:58:10.250Z";
export const docsByLanguage: Record<'cs' | 'en', GeneratedDoc[]> = {
  "cs": [
    {
      "lang": "cs",
      "slug": "framework",
      "section": "framework",
      "title": "Přehled Frameworku",
      "summary": "Přehled Frameworku Účel XLStatUDF je Excel add in se statistickými uživatelskými funkcemi napsanými v C nad Excel DNA. Technologie cílový runtime: .NET 8 integrace s Excelem: Excel",
      "sourcePath": "docs/cs/framework.md",
      "html": "<h1 id=\"prehled-frameworku\">Přehled Frameworku</h1>\n<h2 id=\"ucel\">Účel</h2>\n<p>XLStatUDF je Excel add-in se statistickými uživatelskými funkcemi napsanými v C# nad <code>Excel-DNA</code>.</p>\n<h2 id=\"technologie\">Technologie</h2>\n<ul><li>cílový runtime: <code>.NET 8</code></li><li>integrace s Excelem: <code>Excel-DNA</code></li><li>numerická knihovna: <code>MathNet.Numerics</code></li><li>výstup pro Excel: zabalený <code>.xll</code> doplněk pro 64bitový Excel</li></ul>\n<h2 id=\"hlavni-build\">Hlavní Build</h2>\n<p>Pro běžné použití v Excelu je určený jediný soubor:</p>\n<ul><li><a href=\"#\"><code>artifacts/main/publish/XLStatUDF-AddIn64-packed.xll</code></a></li></ul>\n<h2 id=\"poznamky-k-pouziti\">Poznámky K Použití</h2>\n<ul><li>funkce jsou rozdělené do kategorií <code>Obecné</code>, <code>Popisné</code> a <code>Testy</code></li><li>nápověda argumentů při psaní je zajištěna přes <code>ExcelDna.IntelliSense</code></li><li>většina funkcí ignoruje prázdné buňky v numerických vstupech</li><li>relevantní funkce podporují poslední volitelný argument <code>ma_záhlaví</code></li></ul>\n<h2 id=\"rezim-ma-zahlavi\">Režim <code>ma_záhlaví</code></h2>\n<table><thead><tr><th>Kód</th><th>Význam</th></tr></thead><tbody><tr><td><code>0</code></td><td>autodetekce záhlaví</td></tr><tr><td><code>1</code></td><td>vstup obsahuje záhlaví</td></tr><tr><td><code>2</code></td><td>vstup záhlaví neobsahuje</td></tr></tbody></table>\n<h2 id=\"smer-testu\">Směr Testu</h2>\n<p>Funkce se směrem alternativní hypotézy používají tyto kódy:</p>\n<table><thead><tr><th>Kód</th><th>Význam</th></tr></thead><tbody><tr><td><code>0</code></td><td>oboustranný test</td></tr><tr><td><code>1</code></td><td>levostranný test</td></tr><tr><td><code>2</code></td><td>pravostranný test</td></tr></tbody></table>\n<h2 id=\"chybove-stavy\">Chybové Stavy</h2>\n<table><thead><tr><th>Chyba</th><th>Význam</th></tr></thead><tbody><tr><td><code>#HODNOTA!</code></td><td>neplatný typ argumentu nebo nepodporovaná volba</td></tr><tr><td><code>#ČÍSLO!</code></td><td>neplatný číselný rozsah nebo omezení rozdělení</td></tr><tr><td><code>#POČET!</code></td><td>nedostatek pozorování</td></tr><tr><td><code>#DÉLKA!</code></td><td>rozsahy mají různou délku</td></tr><tr><td><code>#DĚLENÍ0!</code></td><td>dělení nulou v odvozené statistice</td></tr><tr><td><code>#NENÍ_K_DISP!</code></td><td>filtrování odstranilo všechny hodnoty</td></tr></tbody></table>\n<h2 id=\"hladina-vyznamnosti\">Hladina Významnosti</h2>\n<p>Funkce, které přijímají <code>alpha</code>, ji současně vracejí i ve spill výstupu jako <code>α</code>, aby byl výsledek samopopisný.</p>"
    },
    {
      "lang": "cs",
      "slug": "functions/ancova",
      "section": "Functions",
      "title": "ANCOVA",
      "summary": "ANCOVA ANCOVA.G Provádí analýzu kovariance nad groupovanými daty s jedním faktorem a jednou nebo více kovariátami. Syntaxe Argumenty faktor: kategorie faktoru zavisla promenna: záv",
      "sourcePath": "docs/cs/functions/ancova.md",
      "html": "<h1 id=\"ancova\">ANCOVA</h1>\n<h2 id=\"ancovag\"><code>ANCOVA.G</code></h2>\n<p>Provádí analýzu kovariance nad groupovanými daty s jedním faktorem a jednou nebo více kovariátami.</p>\n<h3 id=\"syntaxe\">Syntaxe</h3>\n<pre><code class=\"language-excel\">=ANCOVA.G(faktor; zavisla_promenna; kovariaty; [post_hoc]; [alpha]; [ma_záhlaví])</code></pre>\n<h3 id=\"argumenty\">Argumenty</h3>\n<ul><li><code>faktor</code>: kategorie faktoru</li><li><code>zavisla_promenna</code>: závislá proměnná</li><li><code>kovariaty</code>: jedna nebo více kovariát ve sloupcích</li><li><code>post_hoc</code>: volitelný kód post-hoc procedury; výchozí hodnota je <code>0</code></li><li><code>alpha</code>: hladina významnosti</li><li><code>ma_záhlaví</code>: volitelný kód režimu záhlaví; výchozí hodnota je <code>0</code></li></ul>\n<h3 id=\"kody-post-hoc\">Kódy <code>post_hoc</code></h3>\n<table><thead><tr><th>Kód</th><th>Název</th><th>Popis</th></tr></thead><tbody><tr><td><code>0</code></td><td><code>none</code></td><td>bez post-hoc porovnání</td></tr><tr><td><code>1</code></td><td><code>tukey</code></td><td>konzervativní fallback přes Bonferroni</td></tr><tr><td><code>2</code></td><td><code>bonferroni</code></td><td>párová porovnání adjusted means s Bonferroniho korekcí</td></tr><tr><td><code>3</code></td><td><code>scheffe</code></td><td>Scheffého aproximace nad adjusted means</td></tr><tr><td><code>4</code></td><td><code>games-howell</code></td><td>v aktuální implementaci fallback přes Bonferroni</td></tr></tbody></table>\n<h3 id=\"kody-ma-zahlavi\">Kódy <code>ma_záhlaví</code></h3>\n<table><thead><tr><th>Kód</th><th>Význam</th></tr></thead><tbody><tr><td><code>0</code></td><td>autodetekce záhlaví</td></tr><tr><td><code>1</code></td><td>první řádek je záhlaví</td></tr><tr><td><code>2</code></td><td>vstup je bez záhlaví</td></tr></tbody></table>\n<h3 id=\"poznamky\">Poznámky</h3>\n<ul><li>funkce vyžaduje alespoň dvě skupiny</li><li>kovariáty se zadávají jako jeden nebo více sloupců</li><li>nekompletní řádky se vynechávají jako complete-case</li><li>adjusted means se počítají při globálních průměrech kovariát</li><li>hlavní tabulka zahrnuje i interakce <code>skupina × kovariáta</code>; při významné interakci se zobrazí upozornění na porušení homogenity regresních sklonů</li><li>efektové velikosti <code>η²</code>, <code>η²p</code>, <code>ω²</code>, <code>ω²p</code> jsou ve výstupu pro faktor, kovariáty i interakce</li></ul>\n<h3 id=\"vystup\">Výstup</h3>\n<p>Spill výstup obsahuje:</p>\n<ul><li>popisné statistiky po skupinách</li><li>jednu společnou ANCOVA tabulku pro faktor, jednotlivé kovariáty i interakce</li><li>případné upozornění na porušení homogenity regresních sklonů</li><li>adjusted means po skupinách</li><li>volitelnou post-hoc část</li></ul>\n<h3 id=\"priklad\">Příklad</h3>\n<pre><code class=\"language-excel\">=ANCOVA.G(A2:A100;B2:B100;C2:D100;2;0,05;1)</code></pre>"
    },
    {
      "lang": "cs",
      "slug": "functions/anova",
      "section": "Functions",
      "title": "ANOVA",
      "summary": "ANOVA ANOVA.G Provádí jednofaktorovou analýzu rozptylu nad groupovanými daty. Syntaxe Argumenty categories: štítky skupin values: číselná pozorování ma záhlaví: 0=autodetect, 1=prv",
      "sourcePath": "docs/cs/functions/anova.md",
      "html": "<h1 id=\"anova\">ANOVA</h1>\n<h2 id=\"anovag\"><code>ANOVA.G</code></h2>\n<p>Provádí jednofaktorovou analýzu rozptylu nad groupovanými daty.</p>\n<h3 id=\"syntaxe\">Syntaxe</h3>\n<pre><code class=\"language-excel\">=ANOVA.G(categories; values; [ma_záhlaví]; [alpha]; [post_hoc])</code></pre>\n<h3 id=\"argumenty\">Argumenty</h3>\n<ul><li><code>categories</code>: štítky skupin</li><li><code>values</code>: číselná pozorování</li><li><code>ma_záhlaví</code>: <code>0=autodetect</code>, <code>1=první řádek je záhlaví</code>, <code>2=bez záhlaví</code></li><li><code>alpha</code>: hladina významnosti</li><li><code>post_hoc</code>: volitelný kód post-hoc procedury; výchozí hodnota je <code>0</code></li></ul>\n<h3 id=\"kody-post-hoc\">Kódy <code>post_hoc</code></h3>\n<table><thead><tr><th>Kód</th><th>Název</th><th>Popis</th></tr></thead><tbody><tr><td><code>0</code></td><td><code>none</code></td><td>bez post-hoc porovnání; vrátí se jen hlavní ANOVA report</td></tr><tr><td><code>1</code></td><td><code>tukey</code></td><td>Tukey HSD; v aktuální implementaci konzervativní fallback přes Bonferroni</td></tr><tr><td><code>2</code></td><td><code>bonferroni</code></td><td>párová porovnání s Bonferroniho korekcí</td></tr><tr><td><code>3</code></td><td><code>scheffe</code></td><td>Scheffého metoda pro vícenásobná porovnání</td></tr><tr><td><code>4</code></td><td><code>games-howell</code></td><td>párová porovnání bez předpokladu shodných rozptylů</td></tr></tbody></table>\n<h3 id=\"poznamky\">Poznámky</h3>\n<ul><li>funkce vyžaduje alespoň tři skupiny</li><li>v každé skupině musí být alespoň dvě hodnoty</li><li>součástí výstupu je i Leveneho test homogenity rozptylů</li><li>některé post-hoc volby jsou v aktuální implementaci řešeny přes Bonferroni fallback; výstup to výslovně uvádí</li></ul>\n<h3 id=\"vystup\">Výstup</h3>\n<p>Spill výstup obsahuje:</p>\n<ul><li>popisné statistiky po skupinách</li><li>hlavní ANOVA tabulku</li><li>Leveneho test homogenity rozptylů</li><li>velikosti účinku <code>η²</code>, <code>ω²</code> a <code>f</code></li><li>volitelnou post-hoc část</li></ul>\n<h3 id=\"priklady\">Příklady</h3>\n<pre><code class=\"language-excel\">=ANOVA.G(A2:A40;B2:B40)\n=ANOVA.G(A2:A40;B2:B40;1;0,05;2)\n=ANOVA.G(A2:A40;B2:B40;0;0,05;4)</code></pre>\n<h2 id=\"anovarm\"><code>ANOVA.RM</code></h2>\n<p>Provádí jednofaktorovou ANOVA s opakovaným měřením, kde sloupce představují podmínky a řádky subjekty.</p>\n<h3 id=\"syntaxe\">Syntaxe</h3>\n<pre><code class=\"language-excel\">=ANOVA.RM(hodnoty; [ma_záhlaví]; [alpha]; [post_hoc])</code></pre>\n<h3 id=\"argumenty\">Argumenty</h3>\n<ul><li><code>hodnoty</code>: matice hodnot; řádky jsou subjekty, sloupce podmínky</li><li><code>ma_záhlaví</code>: <code>0=autodetect</code>, <code>1=první řádek je záhlaví</code>, <code>2=bez záhlaví</code></li><li><code>alpha</code>: hladina významnosti</li><li><code>post_hoc</code>: volitelný kód post-hoc procedury; výchozí hodnota je <code>0</code></li></ul>\n<h3 id=\"kody-post-hoc\">Kódy <code>post_hoc</code></h3>\n<table><thead><tr><th>Kód</th><th>Název</th><th>Popis</th></tr></thead><tbody><tr><td><code>0</code></td><td><code>none</code></td><td>bez post-hoc porovnání; vrátí se jen hlavní RM ANOVA report</td></tr><tr><td><code>1</code></td><td><code>tukey</code></td><td>v aktuální implementaci konzervativní fallback přes Bonferroni</td></tr><tr><td><code>2</code></td><td><code>bonferroni</code></td><td>párová porovnání podmínek přes párové t-testy s Bonferroniho korekcí</td></tr><tr><td><code>3</code></td><td><code>scheffe</code></td><td>v aktuální implementaci konzervativní fallback přes Bonferroni</td></tr><tr><td><code>4</code></td><td><code>games-howell</code></td><td>v aktuální implementaci konzervativní fallback přes Bonferroni</td></tr></tbody></table>\n<h3 id=\"poznamky\">Poznámky</h3>\n<ul><li>funkce vyžaduje alespoň dva sloupce a alespoň dva kompletní řádky</li><li>nekompletní řádky se vynechávají jako complete-case</li><li>sphericita zatím není testována; tato informace je uvedena i ve výstupu</li></ul>\n<h3 id=\"vystup\">Výstup</h3>\n<p>Spill výstup obsahuje:</p>\n<ul><li>popisné statistiky po podmínkách</li><li>tabulku RM ANOVA s řádky <code>Podmínky</code>, <code>Subjekty</code>, <code>Reziduum</code>, <code>Celkem</code></li><li>velikosti účinku <code>η²</code>, <code>η²p</code>, <code>ω²</code>, <code>ω²p</code> pro faktor</li><li>poznámku o netestované sphericitě</li><li>volitelnou post-hoc část s párovými porovnáními podmínek</li></ul>\n<h3 id=\"priklad\">Příklad</h3>\n<pre><code class=\"language-excel\">=ANOVA.RM(B2:D25)\n=ANOVA.RM(B1:D25;1;0,05;2)</code></pre>"
    },
    {
      "lang": "cs",
      "slug": "functions/contingency",
      "section": "Functions",
      "title": "Kontingenční Tabulky",
      "summary": "Kontingenční Tabulky CONTINGENCY.T Analyzuje kontingenční tabulku zadanou přímo jako matici četností. Syntaxe Argumenty tabulka: 2D rozsah s pozorovanými četnostmi ma záhlaví: voli",
      "sourcePath": "docs/cs/functions/contingency.md",
      "html": "<h1 id=\"kontingencni-tabulky\">Kontingenční Tabulky</h1>\n<h2 id=\"contingencyt\"><code>CONTINGENCY.T</code></h2>\n<p>Analyzuje kontingenční tabulku zadanou přímo jako matici četností.</p>\n<h3 id=\"syntaxe\">Syntaxe</h3>\n<pre><code class=\"language-excel\">=CONTINGENCY.T(tabulka; [ma_záhlaví]; [alpha])</code></pre>\n<h3 id=\"argumenty\">Argumenty</h3>\n<ul><li><code>tabulka</code>: 2D rozsah s pozorovanými četnostmi</li><li><code>ma_záhlaví</code>: volitelný kód režimu záhlaví; výchozí hodnota je <code>0</code></li><li><code>alpha</code>: hladina významnosti</li></ul>\n<h3 id=\"kody-ma-zahlavi\">Kódy <code>ma_záhlaví</code></h3>\n<table><thead><tr><th>Kód</th><th>Význam</th></tr></thead><tbody><tr><td><code>0</code></td><td>autodetekce popisků</td></tr><tr><td><code>1</code></td><td>horní řádek i levý sloupec jsou popisky</td></tr><tr><td><code>2</code></td><td>tabulka je čistě numerická bez popisků</td></tr></tbody></table>\n<h3 id=\"poznamky\">Poznámky</h3>\n<ul><li>funkce vyžaduje alespoň tabulku <code>2 × 2</code></li><li>četnosti musí být nezáporná celá čísla</li><li>v autodetekci se popisky použijí jen tehdy, když horní řádek i levý sloupec skutečně vypadají jako textové hlavičky</li></ul>\n<h3 id=\"vystup\">Výstup</h3>\n<p>Spill výstup obsahuje:</p>\n<ul><li>pozorovanou kontingenční tabulku včetně marginálních součtů</li><li>tabulku očekávaných četností</li><li>souhrn testu <code>n</code>, <code>df</code>, <code>α</code>, <code>χ²</code>, kritickou hodnotu a <code>p</code></li><li>míry asociace <code>Pearson C</code>, <code>Cramér V</code> a <code>phi</code> pro tabulku <code>2x2</code></li></ul>\n<h3 id=\"priklad\">Příklad</h3>\n<pre><code class=\"language-excel\">=CONTINGENCY.T(A1:C3;1;0,05)</code></pre>\n<h2 id=\"contingencyg\"><code>CONTINGENCY.G</code></h2>\n<p>Analyzuje kontingenční tabulku sestavenou z groupovaných dat.</p>\n<h3 id=\"syntaxe\">Syntaxe</h3>\n<pre><code class=\"language-excel\">=CONTINGENCY.G(sloupce; řádky; [počet]; [alpha]; [ma_záhlaví])</code></pre>\n<h3 id=\"argumenty\">Argumenty</h3>\n<ul><li><code>sloupce</code>: kategorie budoucích sloupců kontingenční tabulky</li><li><code>řádky</code>: kategorie budoucích řádků kontingenční tabulky</li><li><code>počet</code>: volitelné četnosti dvojic; pokud chybí, každá dvojice má váhu <code>1</code></li><li><code>alpha</code>: hladina významnosti</li><li><code>ma_záhlaví</code>: volitelný kód režimu záhlaví; výchozí hodnota je <code>0</code></li></ul>\n<h3 id=\"kody-ma-zahlavi\">Kódy <code>ma_záhlaví</code></h3>\n<table><thead><tr><th>Kód</th><th>Význam</th></tr></thead><tbody><tr><td><code>0</code></td><td>autodetekce záhlaví</td></tr><tr><td><code>1</code></td><td>první řádek je záhlaví</td></tr><tr><td><code>2</code></td><td>vstup je bez záhlaví</td></tr></tbody></table>\n<h3 id=\"poznamky\">Poznámky</h3>\n<ul><li>funkce vyžaduje alespoň dvě různé řádkové i sloupcové kategorie</li><li>prázdné řádky se přeskočí</li><li>pokud je <code>počet</code> zadán, musí obsahovat nezáporné celé četnosti</li><li>výstup má stejnou strukturu jako <code>CONTINGENCY.T</code></li></ul>\n<h3 id=\"vystup\">Výstup</h3>\n<p>Spill výstup obsahuje:</p>\n<ul><li>pozorovanou kontingenční tabulku</li><li>očekávané četnosti</li><li>souhrn testu <code>χ²</code></li><li>míry asociace <code>Pearson C</code>, <code>Cramér V</code> a případně <code>phi</code></li></ul>\n<h3 id=\"priklady\">Příklady</h3>\n<pre><code class=\"language-excel\">=CONTINGENCY.G(A2:A100;B2:B100)\n=CONTINGENCY.G(A2:A100;B2:B100;C2:C100;0,05;1)</code></pre>"
    },
    {
      "lang": "cs",
      "slug": "functions/correlation",
      "section": "Functions",
      "title": "Korelace",
      "summary": "Korelace CORREL.SPEARMAN Počítá Spearmanův pořadový korelační koeficient a provádí jeho test významnosti. Syntaxe Argumenty x values: první číselný rozsah y values: druhý číselný r",
      "sourcePath": "docs/cs/functions/correlation.md",
      "html": "<h1 id=\"korelace\">Korelace</h1>\n<h2 id=\"correlspearman\"><code>CORREL.SPEARMAN</code></h2>\n<p>Počítá Spearmanův pořadový korelační koeficient a provádí jeho test významnosti.</p>\n<h3 id=\"syntaxe\">Syntaxe</h3>\n<pre><code class=\"language-excel\">=CORREL.SPEARMAN(x_values; y_values; [direction]; [alpha]; [ma_záhlaví])</code></pre>\n<h3 id=\"argumenty\">Argumenty</h3>\n<ul><li><code>x_values</code>: první číselný rozsah</li><li><code>y_values</code>: druhý číselný rozsah</li><li><code>direction</code>: volitelný kód směru testu; výchozí hodnota je <code>0</code></li><li><code>alpha</code>: hladina významnosti</li><li><code>ma_záhlaví</code>: volitelný kód režimu záhlaví; výchozí hodnota je <code>0</code></li></ul>\n<h3 id=\"kody-direction\">Kódy <code>direction</code></h3>\n<table><thead><tr><th>Kód</th><th>Význam</th></tr></thead><tbody><tr><td><code>0</code></td><td>oboustranný test</td></tr><tr><td><code>1</code></td><td>levostranný test</td></tr><tr><td><code>2</code></td><td>pravostranný test</td></tr></tbody></table>\n<h3 id=\"kody-ma-zahlavi\">Kódy <code>ma_záhlaví</code></h3>\n<table><thead><tr><th>Kód</th><th>Význam</th></tr></thead><tbody><tr><td><code>0</code></td><td>autodetekce záhlaví</td></tr><tr><td><code>1</code></td><td>první řádek je záhlaví</td></tr><tr><td><code>2</code></td><td>vstup je bez záhlaví</td></tr></tbody></table>\n<h3 id=\"poznamky\">Poznámky</h3>\n<ul><li>prázdné buňky se vylučují po dvojicích</li><li>funkce vyžaduje alespoň tři platné dvojice</li><li>při nulové variabilitě pořadí v některé proměnné vrací funkce numerickou chybu</li></ul>\n<h3 id=\"vystup\">Výstup</h3>\n<p>Spill výstup obsahuje:</p>\n<ul><li><code>ρ</code></li><li><code>n</code></li><li><code>α</code></li><li><code>t</code></li><li><code>df</code></li><li>kritickou hodnotu <code>t</code></li><li><code>p</code></li></ul>\n<h2 id=\"correlmatrix\"><code>CORREL.MATRIX</code></h2>\n<p>Sestavuje korelační matici pro data se dvěma a více sloupci.</p>\n<h3 id=\"syntaxe\">Syntaxe</h3>\n<pre><code class=\"language-excel\">=CORREL.MATRIX(data; [metoda]; [vystup]; [p_minimum]; [ma_záhlaví])</code></pre>\n<h3 id=\"argumenty\">Argumenty</h3>\n<ul><li><code>data</code>: vstupní matice; sloupce jsou proměnné</li><li><code>metoda</code>: volitelný kód výpočetní metody; výchozí hodnota je <code>0</code></li><li><code>vystup</code>: volitelný kód typu výstupu; výchozí hodnota je <code>0</code></li><li><code>p_minimum</code>: volitelný filtr; vrátí jen vazby s <code>p &lt; p_minimum</code></li><li><code>ma_záhlaví</code>: volitelný kód režimu záhlaví; výchozí hodnota je <code>0</code></li></ul>\n<h3 id=\"kody-metoda\">Kódy <code>metoda</code></h3>\n<table><thead><tr><th>Kód</th><th>Význam</th></tr></thead><tbody><tr><td><code>0</code></td><td>Pearson</td></tr><tr><td><code>1</code></td><td>Spearman</td></tr></tbody></table>\n<h3 id=\"kody-vystup\">Kódy <code>vystup</code></h3>\n<table><thead><tr><th>Kód</th><th>Význam</th></tr></thead><tbody><tr><td><code>0</code></td><td>jen korelační koeficienty</td></tr><tr><td><code>1</code></td><td>jen oboustranné p-hodnoty</td></tr><tr><td><code>2</code></td><td>korelační koeficient a v řádku pod ním p-hodnota</td></tr><tr><td><code>3</code></td><td>korelační koeficient, pod ním p-hodnota a na třetím řádku signifikance hvězdičkami</td></tr><tr><td><code>4</code></td><td>jen korelační koeficienty a za nimi v téže buňce signifikance hvězdičkami</td></tr></tbody></table>\n<h3 id=\"kody-ma-zahlavi\">Kódy <code>ma_záhlaví</code></h3>\n<table><thead><tr><th>Kód</th><th>Význam</th></tr></thead><tbody><tr><td><code>0</code></td><td>autodetekce záhlaví</td></tr><tr><td><code>1</code></td><td>první řádek je záhlaví</td></tr><tr><td><code>2</code></td><td>vstup je bez záhlaví</td></tr></tbody></table>\n<h3 id=\"poznamky\">Poznámky</h3>\n<ul><li>funkce vyžaduje alespoň dva sloupce a alespoň tři kompletní řádky</li><li>nekompletní řádky se vynechávají jako complete-case</li><li>pokud je některý sloupec konstantní, funkce vrátí numerickou chybu</li><li>pokud je zadáno <code>p_minimum</code>, nevyhovující vazby se ve výstupu nechají prázdné včetně diagonály</li><li>doplněk kvůli zpětné kompatibilitě stále přijímá i původní pořadí argumentů <code>([p_minimum]; data; [metoda]; [vystup]; [ma_záhlaví])</code></li></ul>\n<h3 id=\"vystup\">Výstup</h3>\n<p>U výstupů <code>0</code>, <code>1</code> a <code>4</code> vrací spill výstup čtvercovou matici s názvy proměnných.</p>\n<p>U výstupů <code>2</code> a <code>3</code> vrací spill výstup skládané rozvržení se dvěma úvodními popisnými sloupci a jedním blokem řádků pro každou proměnnou:</p>\n<ul><li>koeficient</li><li><code>p</code></li><li>volitelně <code>sig.</code></li></ul>\n<p>U výstupu <code>4</code> je v každé buňce text ve formátu například <code>0,30156***</code>.</p>\n<h3 id=\"priklady\">Příklady</h3>\n<pre><code class=\"language-excel\">=CORREL.MATRIX(B1:E30)\n=CORREL.MATRIX(B1:E30;1;0;;1)\n=CORREL.MATRIX(B1:E30;0;4;0,05;1)</code></pre>"
    },
    {
      "lang": "cs",
      "slug": "functions/distributions",
      "section": "Functions",
      "title": "Rozdělení",
      "summary": "Rozdělení NORM.DIST.RANGE Počítá pravděpodobnost, že náhodná veličina s normálním rozdělením padne do zadaného intervalu. Syntaxe Argumenty mean: střední hodnota rozdělení standard",
      "sourcePath": "docs/cs/functions/distributions.md",
      "html": "<h1 id=\"rozdeleni\">Rozdělení</h1>\n<h2 id=\"normdistrange\"><code>NORM.DIST.RANGE</code></h2>\n<p>Počítá pravděpodobnost, že náhodná veličina s normálním rozdělením padne do zadaného intervalu.</p>\n<h3 id=\"syntaxe\">Syntaxe</h3>\n<pre><code class=\"language-excel\">=NORM.DIST.RANGE(mean; standard_deviation; lower_bound; upper_bound)</code></pre>\n<h3 id=\"argumenty\">Argumenty</h3>\n<ul><li><code>mean</code>: střední hodnota rozdělení</li><li><code>standard_deviation</code>: kladná směrodatná odchylka</li><li><code>lower_bound</code>: dolní mez intervalu; prázdná buňka znamená mínus nekonečno</li><li><code>upper_bound</code>: horní mez intervalu; prázdná buňka znamená plus nekonečno</li></ul>\n<h3 id=\"poznamky\">Poznámky</h3>\n<ul><li>pokud <code>lower_bound &gt; upper_bound</code>, funkce vrátí numerickou chybu</li><li>pokud <code>standard_deviation &lt;= 0</code>, funkce vrátí numerickou chybu</li><li>prázdné meze lze použít pro jednostranné intervaly</li></ul>\n<h3 id=\"vystup\">Výstup</h3>\n<p>Skalární hodnota z intervalu <code>[0;1]</code>.</p>\n<h3 id=\"priklad\">Příklad</h3>\n<pre><code class=\"language-excel\">=NORM.DIST.RANGE(0;1;-1;1)</code></pre>\n<h2 id=\"generatenorm\"><code>GENERATE.NORM</code></h2>\n<p>Generuje jednu náhodnou hodnotu z normálního rozdělení s volitelnou perturbací.</p>\n<h3 id=\"syntaxe\">Syntaxe</h3>\n<pre><code class=\"language-excel\">=GENERATE.NORM(stredni_hodnota; směrodatná_odchylka; [outlier_rate])</code></pre>\n<h3 id=\"argumenty\">Argumenty</h3>\n<ul><li><code>stredni_hodnota</code>: požadovaný průměr</li><li><code>směrodatná_odchylka</code>: kladná směrodatná odchylka</li><li><code>outlier_rate</code>: volitelná pravděpodobnost dodatečné náhodné perturbace v intervalu <code>&lt;0;1&gt;</code>; výchozí hodnota je <code>0</code></li></ul>\n<h3 id=\"poznamky\">Poznámky</h3>\n<ul><li>funkce je volatilní, takže po přepočtu vygeneruje nový vzorek</li><li>při <code>outlier_rate = 0</code> se generuje běžná hodnota z <code>N(μ, σ)</code></li><li>pokud náhodně nastane perturbace, k vygenerované hodnotě se přičte dodatečný šum z <code>N(0, 3σ)</code></li><li>neplatný <code>outlier_rate</code> nebo nečíselné argumenty vrací chybu hodnoty</li></ul>\n<h3 id=\"vystup\">Výstup</h3>\n<p>Jedna skalární hodnota.</p>\n<h3 id=\"priklad\">Příklad</h3>\n<pre><code class=\"language-excel\">=GENERATE.NORM(100;15)\n=GENERATE.NORM(0;1;0,2)</code></pre>\n<h2 id=\"generateint\"><code>GENERATE.INT</code></h2>\n<p>Generuje jedno náhodné celé číslo ze zadaného intervalu s volitelnou perturbací.</p>\n<h3 id=\"syntaxe\">Syntaxe</h3>\n<pre><code class=\"language-excel\">=GENERATE.INT([minimum]; [maximum]; [outlier_rate])</code></pre>\n<h3 id=\"argumenty\">Argumenty</h3>\n<ul><li><code>minimum</code>: volitelná dolní mez intervalu; výchozí hodnota je <code>-2147483648</code></li><li><code>maximum</code>: volitelná horní mez intervalu; výchozí hodnota je <code>2147483647</code></li><li><code>outlier_rate</code>: volitelná pravděpodobnost dodatečné náhodné perturbace v intervalu <code>&lt;0;1&gt;</code>; výchozí hodnota je <code>0</code></li></ul>\n<h3 id=\"poznamky\">Poznámky</h3>\n<ul><li>funkce je volatilní, takže po přepočtu vygeneruje novou hodnotu</li><li>pokud <code>minimum &gt; maximum</code>, funkce vrátí numerickou chybu</li><li>pokud hranice nezadáš, použije se celé praktické 32bitové rozmezí</li><li>při <code>outlier_rate = 0</code> se generuje běžná hodnota z uzavřeného intervalu <code>&lt;minimum; maximum&gt;</code></li><li>pokud náhodně nastane perturbace, k vygenerované hodnotě se přičte dodatečný náhodný celočíselný posun z intervalu <code>⟨-(maximum-minimum); +(maximum-minimum)⟩</code></li></ul>\n<h3 id=\"vystup\">Výstup</h3>\n<p>Jedno celé číslo.</p>\n<h3 id=\"priklady\">Příklady</h3>\n<pre><code class=\"language-excel\">=GENERATE.INT()\n=GENERATE.INT(1;6)\n=GENERATE.INT(1;6;0,2)</code></pre>\n<h2 id=\"fill\"><code>FILL</code></h2>\n<p>Opakuje jednu nebo více hodnot, případně opakovaně vyhodnocuje textový vzorec, do jednosloupcového spill výstupu.</p>\n<h3 id=\"syntaxe\">Syntaxe</h3>\n<pre><code class=\"language-excel\">=FILL(co; počet; [co2]; [počet2]; ...)</code></pre>\n<h3 id=\"argumenty\">Argumenty</h3>\n<ul><li><code>co</code>: skalární hodnota, která se má opakovat, nebo textový vzorec začínající <code>=</code></li><li><code>počet</code>: počet vrácených řádků; celé číslo <code>&gt;= 1</code></li><li>další argumenty se zadávají po dvojicích <code>co + počet</code></li></ul>\n<h3 id=\"poznamky\">Poznámky</h3>\n<ul><li>pokud zadáš běžnou hodnotu, funkce ji pouze zkopíruje do všech řádků</li><li>pokud zadáš textový vzorec začínající <code>=</code>, funkce ho vyhodnotí zvlášť pro každý řádek</li><li>počet dodatečných argumentů musí být sudý; jinak funkce vrátí chybu hodnoty</li><li>přímý argument typu <code>GENERATE.NORM(...)</code> nebo <code>RANDBETWEEN(...)</code> Excel vyhodnotí ještě před vstupem do <code>FILL</code>, proto je pro opakované přepočítání nutný textový vzorec</li></ul>\n<h3 id=\"vystup\">Výstup</h3>\n<p>Jednosloupcový spill rozsah s výslednou řadou.</p>\n<h3 id=\"priklady\">Příklady</h3>\n<pre><code class=\"language-excel\">=FILL(&quot;A&quot;;5)\n=FILL(123;4)\n=FILL(&quot;muž&quot;;100;&quot;žena&quot;;100;&quot;dítě&quot;;90)\n=FILL(&quot;=RANDBETWEEN(1;10)&quot;;20)\n=FILL(&quot;=GENERATE.NORM(0;1;0,2)&quot;;100)</code></pre>\n<h2 id=\"fillrandom\"><code>FILL.RANDOM</code></h2>\n<p>Sestaví řadu stejně jako <code>FILL</code>, ale před vrácením ji náhodně promíchá.</p>\n<h3 id=\"syntaxe\">Syntaxe</h3>\n<pre><code class=\"language-excel\">=FILL.RANDOM(co; počet; [co2]; [počet2]; ...)</code></pre>\n<h3 id=\"argumenty\">Argumenty</h3>\n<p>Stejné jako u <code>FILL</code>.</p>\n<h3 id=\"poznamky\">Poznámky</h3>\n<ul><li>nejprve se vytvoří celá řada stejně jako u <code>FILL</code></li><li>teprve poté se hotová řada náhodně promíchá</li><li>stejné validační podmínky jako u <code>FILL</code> platí i zde</li></ul>\n<h3 id=\"vystup\">Výstup</h3>\n<p>Jednosloupcový spill rozsah se všemi vygenerovanými hodnotami v náhodném pořadí.</p>\n<h3 id=\"priklady\">Příklady</h3>\n<pre><code class=\"language-excel\">=FILL.RANDOM(&quot;muž&quot;;100;&quot;žena&quot;;100;&quot;dítě&quot;;90)\n=FILL.RANDOM(&quot;A&quot;;5;&quot;B&quot;;5)</code></pre>"
    },
    {
      "lang": "cs",
      "slug": "functions/goodness-of-fit",
      "section": "Functions",
      "title": "Test Dobrého Souladu",
      "summary": "Test Dobrého Souladu CHISQ.GOF Provádí chí kvadrát test dobré shody. Syntaxe Argumenty observed: pozorované četnosti kategorií expected: očekávané četnosti nebo pravděpodobnosti ca",
      "sourcePath": "docs/cs/functions/goodness-of-fit.md",
      "html": "<h1 id=\"test-dobreho-souladu\">Test Dobrého Souladu</h1>\n<h2 id=\"chisqgof\"><code>CHISQ.GOF</code></h2>\n<p>Provádí chí-kvadrát test dobré shody.</p>\n<h3 id=\"syntaxe\">Syntaxe</h3>\n<pre><code class=\"language-excel\">=CHISQ.GOF(observed; expected; [categories]; [alpha]; [ma_záhlaví])</code></pre>\n<h3 id=\"argumenty\">Argumenty</h3>\n<ul><li><code>observed</code>: pozorované četnosti kategorií</li><li><code>expected</code>: očekávané četnosti nebo pravděpodobnosti</li><li><code>categories</code>: volitelné názvy kategorií</li><li><code>alpha</code>: hladina významnosti</li><li><code>ma_záhlaví</code>: volitelný kód režimu záhlaví; výchozí hodnota je <code>0</code></li></ul>\n<h3 id=\"kody-ma-zahlavi\">Kódy <code>ma_záhlaví</code></h3>\n<table><thead><tr><th>Kód</th><th>Význam</th></tr></thead><tbody><tr><td><code>0</code></td><td>autodetekce záhlaví</td></tr><tr><td><code>1</code></td><td>první řádek je záhlaví</td></tr><tr><td><code>2</code></td><td>vstup je bez záhlaví</td></tr></tbody></table>\n<h3 id=\"poznamky\">Poznámky</h3>\n<ul><li><code>observed</code> musí obsahovat nezáporné celé četnosti</li><li><code>expected</code> může být zadáno jako četnosti nebo jako pravděpodobnosti se součtem <code>1</code></li><li>pokud <code>expected</code> tvoří pravděpodobnosti, funkce je automaticky přepočte na očekávané četnosti podle velikosti vzorku</li><li>pokud <code>categories</code> chybí, kategorie se očíslují <code>1, 2, 3, ...</code></li></ul>\n<h3 id=\"vystup\">Výstup</h3>\n<p>Spill výstup obsahuje:</p>\n<ul><li>souhrn testu <code>χ²</code>, <code>df</code>, <code>α</code>, kritickou hodnotu a <code>p</code></li><li>tabulku kategorií s <code>O</code>, <code>E</code> a příspěvkem <code>(O−E)² / E</code></li></ul>"
    },
    {
      "lang": "cs",
      "slug": "functions/index",
      "section": "Functions",
      "title": "Index Funkcí",
      "summary": "Index Funkcí Obecné GENERATE.NORM: Generuje jednu náhodnou hodnotu z normálního rozdělení s volitelnou perturbací. GENERATE.INT: Generuje jedno náhodné celé číslo ze zadaného inter",
      "sourcePath": "docs/cs/functions/index.md",
      "html": "<h1 id=\"index-funkci\">Index Funkcí</h1>\n<h2 id=\"obecne\">Obecné</h2>\n<ul><li><a href=\"./distributions.md#generatenorm\">GENERATE.NORM</a>: Generuje jednu náhodnou hodnotu z normálního rozdělení s volitelnou perturbací.</li><li><a href=\"./distributions.md#generateint\">GENERATE.INT</a>: Generuje jedno náhodné celé číslo ze zadaného intervalu s volitelnou perturbací.</li><li><a href=\"./distributions.md#fill\">FILL</a>: Opakuje hodnoty nebo textové vzorce do jednosloupcového spill výstupu.</li><li><a href=\"./distributions.md#fillrandom\">FILL.RANDOM</a>: Sestavuje řadu jako <code>FILL</code> a následně ji náhodně promíchá.</li></ul>\n<h2 id=\"popisne\">Popisné</h2>\n<ul><li><a href=\"./weighted-means.md#averagew\">AVERAGE.W</a>: Počítá vážený aritmetický průměr.</li><li><a href=\"./weighted-means.md#harmeanw\">HARMEAN.W</a>: Počítá vážený harmonický průměr.</li><li><a href=\"./weighted-means.md#geomeanw\">GEOMEAN.W</a>: Počítá vážený geometrický průměr.</li><li><a href=\"./weighted-variance.md#varpw\">VAR.P.W</a>: Počítá vážený populační rozptyl.</li><li><a href=\"./weighted-variance.md#varsw\">VAR.S.W</a>: Počítá vážený výběrový rozptyl.</li><li><a href=\"./weighted-variance.md#stdevpw\">STDEV.P.W</a>: Počítá váženou populační směrodatnou odchylku.</li><li><a href=\"./weighted-variance.md#stdevsw\">STDEV.S.W</a>: Počítá váženou výběrovou směrodatnou odchylku.</li><li><a href=\"./variation-coefficients.md#varcoef\">VARCOEF</a>: Počítá populační variační koeficient.</li><li><a href=\"./variation-coefficients.md#varcoefs\">VARCOEF.S</a>: Počítá výběrový variační koeficient.</li><li><a href=\"./variation-coefficients.md#varcoefw\">VARCOEF.W</a>: Počítá vážený populační variační koeficient.</li><li><a href=\"./variation-coefficients.md#varcoefsw\">VARCOEF.S.W</a>: Počítá vážený výběrový variační koeficient.</li><li><a href=\"./percentiles.md#percentileincifs\">PERCENTILE.INC.IFS</a>: Počítá inkluzivní percentil po aplikaci filtrů.</li><li><a href=\"./percentiles.md#percentileexcifs\">PERCENTILE.EXC.IFS</a>: Počítá exkluzivní percentil po aplikaci filtrů.</li><li><a href=\"./pivot.md#pivot\">PIVOT.*</a>: Sestavuje statistický pivot a v každé funkci počítá právě jeden zvolený ukazatel.</li></ul>\n<h2 id=\"testy\">Testy</h2>\n<ul><li><a href=\"./distributions.md#normdistrange\">NORM.DIST.RANGE</a>: Počítá pravděpodobnost intervalu normálního rozdělení.</li><li><a href=\"./normality.md#shapirowilk\">SHAPIRO.WILK</a>: Provádí Shapiro-Wilkův test normality.</li><li><a href=\"./normality.md#kolmogorovsmirnov\">KOLMOGOROV.SMIRNOV</a>: Provádí Kolmogorov-Smirnovův test dobré shody pro zvolené rozdělení.</li><li><a href=\"./one-sample-tests.md#ttest1s\">T.TEST.1S</a>: Provádí jednovýběrový t-test vůči zadané hypotetické střední hodnotě.</li><li><a href=\"./one-sample-tests.md#proptest1s\">PROP.TEST.1S</a>: Provádí jednovýběrový test podílu.</li><li><a href=\"./one-sample-tests.md#wilcoxonpaired\">WILCOXON.PAIRED</a>: Provádí Wilcoxonův párový znaménkový test pro závislá měření.</li><li><a href=\"./two-sample-tests.md#welchtest2sg\">WELCH.TEST.2S.G</a>: Provádí Welchův dvouvýběrový t-test pro dvě nezávislé skupiny.</li><li><a href=\"./two-sample-tests.md#mannwhitneyg\">MANN.WHITNEY.G</a>: Provádí Mann-Whitneyho test pro dvě nezávislé skupiny.</li><li><a href=\"./goodness-of-fit.md#chisqgof\">CHISQ.GOF</a>: Provádí chí-kvadrát test dobré shody.</li><li><a href=\"./anova.md#anovag\">ANOVA.G</a>: Provádí jednofaktorovou ANOVA nad groupovanými daty.</li><li><a href=\"./anova.md#anovarm\">ANOVA.RM</a>: Provádí ANOVA s opakovaným měřením nad sloupci.</li><li><a href=\"./ancova.md#ancovag\">ANCOVA.G</a>: Provádí ANCOVA s jedním faktorem a jednou nebo více kovariátami.</li><li><a href=\"./contingency.md#contingencyt\">CONTINGENCY.T</a>: Analyzuje kontingenční tabulku zadanou přímo jako matici četností.</li><li><a href=\"./contingency.md#contingencyg\">CONTINGENCY.G</a>: Analyzuje kontingenční tabulku sestavenou z groupovaných sloupců.</li><li><a href=\"./correlation.md#correlspearman\">CORREL.SPEARMAN</a>: Počítá Spearmanův korelační koeficient a test jeho významnosti.</li><li><a href=\"./correlation.md#correlmatrix\">CORREL.MATRIX</a>: Sestavuje korelační matici včetně p-hodnot a značek signifikance.</li></ul>"
    },
    {
      "lang": "cs",
      "slug": "functions/normality",
      "section": "Functions",
      "title": "Testy Normality A Shody",
      "summary": "Testy Normality A Shody SHAPIRO.WILK Provádí Shapiro Wilkův test normality. Syntaxe Argumenty values: číselný výběr; prázdné buňky jsou ignorovány ma záhlaví: 0=autodetect, 1=první",
      "sourcePath": "docs/cs/functions/normality.md",
      "html": "<h1 id=\"testy-normality-a-shody\">Testy Normality A Shody</h1>\n<h2 id=\"shapirowilk\"><code>SHAPIRO.WILK</code></h2>\n<p>Provádí Shapiro-Wilkův test normality.</p>\n<h3 id=\"syntaxe\">Syntaxe</h3>\n<pre><code class=\"language-excel\">=SHAPIRO.WILK(values; [ma_záhlaví])</code></pre>\n<h3 id=\"argumenty\">Argumenty</h3>\n<ul><li><code>values</code>: číselný výběr; prázdné buňky jsou ignorovány</li><li><code>ma_záhlaví</code>: <code>0=autodetect</code>, <code>1=první buňka je záhlaví</code>, <code>2=bez záhlaví</code></li></ul>\n<h3 id=\"poznamky\">Poznámky</h3>\n<ul><li>funkce vyžaduje velikost vzorku od <code>3</code> do <code>5000</code></li><li>data se před výpočtem seřadí vzestupně</li><li>pokud mají všechna pozorování stejnou hodnotu, statistika vyjde hraničně a interpretace není smysluplná</li></ul>\n<h3 id=\"vystup\">Výstup</h3>\n<p>Dvouřádkový spill výstup:</p>\n<ul><li><code>W</code>: testová statistika</li><li><code>p</code>: p-hodnota</li></ul>\n<h3 id=\"priklad\">Příklad</h3>\n<pre><code class=\"language-excel\">=SHAPIRO.WILK(B2:B18)</code></pre>\n<h2 id=\"kolmogorovsmirnov\"><code>KOLMOGOROV.SMIRNOV</code></h2>\n<p>Provádí jednovýběrový Kolmogorov-Smirnovův test dobré shody.</p>\n<h3 id=\"syntaxe\">Syntaxe</h3>\n<pre><code class=\"language-excel\">=KOLMOGOROV.SMIRNOV(values; [distribution]; [ma_záhlaví])</code></pre>\n<h3 id=\"argumenty\">Argumenty</h3>\n<ul><li><code>values</code>: číselný výběr; prázdné buňky jsou ignorovány</li><li><code>distribution</code>: volitelný kód testovaného rozdělení; výchozí hodnota je <code>0</code></li><li><code>ma_záhlaví</code>: <code>0=autodetect</code>, <code>1=první buňka je záhlaví</code>, <code>2=bez záhlaví</code></li></ul>\n<h3 id=\"kody-distribution\">Kódy <code>distribution</code></h3>\n<table><thead><tr><th>Kód</th><th>Rozdělení</th><th>Co se testuje</th></tr></thead><tbody><tr><td><code>0</code></td><td><code>normal</code></td><td>zda data odpovídají normálnímu rozdělení s průměrem a směrodatnou odchylkou odhadnutou ze vzorku</td></tr><tr><td><code>1</code></td><td><code>lognormal</code></td><td>zda data odpovídají lognormálnímu rozdělení; všechny hodnoty musí být kladné</td></tr><tr><td><code>2</code></td><td><code>exponential</code></td><td>zda data odpovídají exponenciálnímu rozdělení; všechny hodnoty musí být nezáporné</td></tr><tr><td><code>3</code></td><td><code>uniform</code></td><td>zda data odpovídají spojitému rovnoměrnému rozdělení na intervalu daném minimem a maximem vzorku</td></tr><tr><td><code>4</code></td><td><code>weibull</code></td><td>zda data odpovídají Weibullovu rozdělení s parametry odhadnutými ze vzorku</td></tr></tbody></table>\n<h3 id=\"poznamky\">Poznámky</h3>\n<ul><li>funkce vyžaduje alespoň <code>5</code> platných pozorování</li><li>data se před výpočtem seřadí vzestupně</li><li>některá rozdělení mají dodatečné podmínky na vstup:</li><li><code>lognormal</code> a <code>weibull</code>: pouze kladné hodnoty</li><li><code>exponential</code>: pouze nezáporné hodnoty</li><li><code>uniform</code>: data nesmí být konstantní</li><li>p-hodnota pro <code>normal</code> používá korekci odlišnou od ostatních rozdělení</li></ul>\n<h3 id=\"vystup\">Výstup</h3>\n<p>Dvouřádkový spill výstup:</p>\n<ul><li><code>D</code>: testová statistika</li><li><code>p</code>: p-hodnota</li></ul>\n<h3 id=\"priklady\">Příklady</h3>\n<pre><code class=\"language-excel\">=KOLMOGOROV.SMIRNOV(B2:B18)\n=KOLMOGOROV.SMIRNOV(B2:B18;1)\n=KOLMOGOROV.SMIRNOV(B2:B18;4;1)</code></pre>"
    },
    {
      "lang": "cs",
      "slug": "functions/one-sample-tests",
      "section": "Functions",
      "title": "Jednovýběrové Testy",
      "summary": "Jednovýběrové Testy T.TEST.1S Provádí jednovýběrový t test. Syntaxe Argumenty values: číselný výběr mu 0: hypotetická střední hodnota direction: volitelný kód směru testu; výchozí",
      "sourcePath": "docs/cs/functions/one-sample-tests.md",
      "html": "<h1 id=\"jednovyberove-testy\">Jednovýběrové Testy</h1>\n<h2 id=\"ttest1s\"><code>T.TEST.1S</code></h2>\n<p>Provádí jednovýběrový t-test.</p>\n<h3 id=\"syntaxe\">Syntaxe</h3>\n<pre><code class=\"language-excel\">=T.TEST.1S(values; mu_0; [direction]; [alpha]; [ma_záhlaví])</code></pre>\n<h3 id=\"argumenty\">Argumenty</h3>\n<ul><li><code>values</code>: číselný výběr</li><li><code>mu_0</code>: hypotetická střední hodnota</li><li><code>direction</code>: volitelný kód směru testu; výchozí hodnota je <code>0</code></li><li><code>alpha</code>: hladina významnosti</li><li><code>ma_záhlaví</code>: volitelný kód režimu záhlaví; výchozí hodnota je <code>0</code></li></ul>\n<h3 id=\"kody-direction\">Kódy <code>direction</code></h3>\n<table><thead><tr><th>Kód</th><th>Význam</th></tr></thead><tbody><tr><td><code>0</code></td><td>oboustranný test</td></tr><tr><td><code>1</code></td><td>levostranný test</td></tr><tr><td><code>2</code></td><td>pravostranný test</td></tr></tbody></table>\n<h3 id=\"kody-ma-zahlavi\">Kódy <code>ma_záhlaví</code></h3>\n<table><thead><tr><th>Kód</th><th>Význam</th></tr></thead><tbody><tr><td><code>0</code></td><td>autodetekce záhlaví</td></tr><tr><td><code>1</code></td><td>první buňka je záhlaví</td></tr><tr><td><code>2</code></td><td>vstup je bez záhlaví</td></tr></tbody></table>\n<h3 id=\"poznamky\">Poznámky</h3>\n<ul><li>prázdné buňky se ignorují</li><li>jsou potřeba alespoň dvě platné hodnoty</li><li>kritická hodnota <code>t</code> se počítá podle zvoleného směru testu</li></ul>\n<h3 id=\"vystup\">Výstup</h3>\n<p>Spill výstup obsahuje:</p>\n<ul><li><code>x̄</code></li><li><code>μ₀</code></li><li><code>sₓ</code></li><li><code>n</code></li><li><code>α</code></li><li><code>t</code></li><li><code>df</code></li><li>kritickou hodnotu <code>t</code></li><li><code>p</code></li></ul>\n<h2 id=\"proptest1s\"><code>PROP.TEST.1S</code></h2>\n<p>Provádí jednovýběrový z-test podílu.</p>\n<h3 id=\"syntaxe\">Syntaxe</h3>\n<pre><code class=\"language-excel\">=PROP.TEST.1S(values; pi_0; [direction]; [alpha]; [ma_záhlaví])</code></pre>\n<h3 id=\"argumenty\">Argumenty</h3>\n<ul><li><code>values</code>: binární výběr ve formátu <code>0/1</code> nebo <code>FALSE/TRUE</code></li><li><code>pi_0</code>: hypotetický populační podíl</li><li><code>direction</code>: volitelný kód směru testu; výchozí hodnota je <code>0</code></li><li><code>alpha</code>: hladina významnosti</li><li><code>ma_záhlaví</code>: volitelný kód režimu záhlaví; výchozí hodnota je <code>0</code></li></ul>\n<h3 id=\"kody-direction\">Kódy <code>direction</code></h3>\n<table><thead><tr><th>Kód</th><th>Význam</th></tr></thead><tbody><tr><td><code>0</code></td><td>oboustranný test</td></tr><tr><td><code>1</code></td><td>levostranný test</td></tr><tr><td><code>2</code></td><td>pravostranný test</td></tr></tbody></table>\n<h3 id=\"kody-ma-zahlavi\">Kódy <code>ma_záhlaví</code></h3>\n<table><thead><tr><th>Kód</th><th>Význam</th></tr></thead><tbody><tr><td><code>0</code></td><td>autodetekce záhlaví</td></tr><tr><td><code>1</code></td><td>první buňka je záhlaví</td></tr><tr><td><code>2</code></td><td>vstup je bez záhlaví</td></tr></tbody></table>\n<h3 id=\"poznamky\">Poznámky</h3>\n<ul><li><code>pi_0</code> musí ležet přísně mezi <code>0</code> a <code>1</code></li><li>funkce přijímá pouze binární data <code>0/1</code> nebo <code>FALSE/TRUE</code></li><li>jsou potřeba alespoň jedna platná pozorování</li></ul>\n<h3 id=\"vystup\">Výstup</h3>\n<p>Spill výstup obsahuje:</p>\n<ul><li><code>p̂</code></li><li><code>π₀</code></li><li><code>x</code></li><li><code>n</code></li><li><code>α</code></li><li><code>z</code></li><li>kritickou hodnotu <code>z</code></li><li><code>p</code></li></ul>\n<h2 id=\"wilcoxonpaired\"><code>WILCOXON.PAIRED</code></h2>\n<p>Provádí Wilcoxonův párový znaménkový test pro závislé dvojice.</p>\n<h3 id=\"syntaxe\">Syntaxe</h3>\n<pre><code class=\"language-excel\">=WILCOXON.PAIRED(x; y; [ma_záhlaví]; [alpha]; [smer])</code></pre>\n<h3 id=\"argumenty\">Argumenty</h3>\n<ul><li><code>x</code>: první měření</li><li><code>y</code>: druhé měření</li><li><code>ma_záhlaví</code>: volitelný kód režimu záhlaví; výchozí hodnota je <code>0</code></li><li><code>alpha</code>: hladina významnosti</li><li><code>smer</code>: volitelný kód směru testu; výchozí hodnota je <code>0</code></li></ul>\n<h3 id=\"kody-smer\">Kódy <code>smer</code></h3>\n<table><thead><tr><th>Kód</th><th>Význam</th></tr></thead><tbody><tr><td><code>0</code></td><td>oboustranný test</td></tr><tr><td><code>1</code></td><td>levostranný test</td></tr><tr><td><code>2</code></td><td>pravostranný test</td></tr></tbody></table>\n<h3 id=\"kody-ma-zahlavi\">Kódy <code>ma_záhlaví</code></h3>\n<table><thead><tr><th>Kód</th><th>Význam</th></tr></thead><tbody><tr><td><code>0</code></td><td>autodetekce záhlaví</td></tr><tr><td><code>1</code></td><td>první buňka je záhlaví</td></tr><tr><td><code>2</code></td><td>vstup je bez záhlaví</td></tr></tbody></table>\n<h3 id=\"poznamky\">Poznámky</h3>\n<ul><li>funkce pracuje s párovými rozdíly <code>x - y</code></li><li>nekompletní dvojice se vynechávají po dvojicích</li><li>nulové rozdíly se z testu vyřazují</li><li>pokud po vyřazení nulových rozdílů nezbude dost dat, funkce vrátí chybu počtu</li></ul>\n<h3 id=\"vystup\">Výstup</h3>\n<p>Spill výstup obsahuje:</p>\n<ul><li><code>n</code></li><li><code>med(d)</code></li><li><code>α</code></li><li><code>W+</code></li><li><code>W-</code></li><li><code>W</code></li><li><code>z</code></li><li>kritickou hodnotu <code>z</code></li><li><code>p</code></li><li>efekt velikosti <code>r</code></li></ul>\n<h3 id=\"priklad\">Příklad</h3>\n<pre><code class=\"language-excel\">=WILCOXON.PAIRED(B2:B20;C2:C20)\n=WILCOXON.PAIRED(B2:B20;C2:C20;1;0,05;0)</code></pre>"
    },
    {
      "lang": "cs",
      "slug": "functions/percentiles",
      "section": "Functions",
      "title": "Percentily",
      "summary": "Percentily PERCENTILE.INC.IFS Počítá inkluzivní percentil s filtrováním ve stylu SUMIFS. Syntaxe Argumenty values: číselná data quantile: hodnota od 0 do 1 criteria range n: rozsah",
      "sourcePath": "docs/cs/functions/percentiles.md",
      "html": "<h1 id=\"percentily\">Percentily</h1>\n<h2 id=\"percentileincifs\"><code>PERCENTILE.INC.IFS</code></h2>\n<p>Počítá inkluzivní percentil s filtrováním ve stylu <code>SUMIFS</code>.</p>\n<h3 id=\"syntaxe\">Syntaxe</h3>\n<pre><code class=\"language-excel\">=PERCENTILE.INC.IFS(values; quantile; [criteria_range_1; criteria_1; ...])</code></pre>\n<h3 id=\"argumenty\">Argumenty</h3>\n<ul><li><code>values</code>: číselná data</li><li><code>quantile</code>: hodnota od <code>0</code> do <code>1</code></li><li><code>criteria_range_n</code>: rozsah, podle kterého se filtruje</li><li><code>criteria_n</code>: přesná shoda, relační výraz nebo wildcard výraz</li></ul>\n<h3 id=\"poznamky\">Poznámky</h3>\n<ul><li>počet argumentů filtru musí být sudý po dvojicích <code>rozsah + kritérium</code></li><li>pokud po filtrování nezůstane žádná hodnota, funkce vrátí <code>#N/A</code></li><li>percentil se počítá inkluzivní metodou stejně jako <code>PERCENTILE.INC</code></li></ul>\n<h3 id=\"vystup\">Výstup</h3>\n<p>Skalární percentil.</p>\n<h3 id=\"priklad\">Příklad</h3>\n<pre><code class=\"language-excel\">=PERCENTILE.INC.IFS(A2:A100;0,75;B2:B100;&quot;A&quot;;A2:A100;&quot;&gt;10&quot;)</code></pre>\n<h2 id=\"percentileexcifs\"><code>PERCENTILE.EXC.IFS</code></h2>\n<p>Počítá exkluzivní percentil s filtrováním ve stylu <code>SUMIFS</code>.</p>\n<h3 id=\"syntaxe\">Syntaxe</h3>\n<pre><code class=\"language-excel\">=PERCENTILE.EXC.IFS(values; quantile; [criteria_range_1; criteria_1; ...])</code></pre>\n<h3 id=\"argumenty\">Argumenty</h3>\n<p>Stejné jako u <code>PERCENTILE.INC.IFS</code>, ale <code>quantile</code> musí být přísně mezi <code>0</code> a <code>1</code>.</p>\n<h3 id=\"poznamky\">Poznámky</h3>\n<ul><li>počet argumentů filtru musí být sudý po dvojicích <code>rozsah + kritérium</code></li><li>pokud po filtrování nezůstane žádná hodnota, funkce vrátí <code>#N/A</code></li><li>pokud je pro zadaný kvantil exkluzivní percentil mimo definiční obor vzorce, funkce vrátí numerickou chybu</li></ul>\n<h3 id=\"vystup\">Výstup</h3>\n<p>Skalární percentil.</p>"
    },
    {
      "lang": "cs",
      "slug": "functions/pivot",
      "section": "Functions",
      "title": "PIVOT",
      "summary": "PIVOT PIVOT. Rodina funkcí PIVOT. sestavuje statistický pivot z řádkových a sloupcových kategorií a pro každou funkci počítá právě jeden ukazatel. Syntaxe Argumenty řádky: jeden ne",
      "sourcePath": "docs/cs/functions/pivot.md",
      "html": "<h1 id=\"pivot\">PIVOT</h1>\n<h2 id=\"pivot\"><code>PIVOT.*</code></h2>\n<p>Rodina funkcí <code>PIVOT.*</code> sestavuje statistický pivot z řádkových a sloupcových kategorií a pro každou funkci počítá právě jeden ukazatel.</p>\n<h3 id=\"syntaxe\">Syntaxe</h3>\n<pre><code class=\"language-excel\">=PIVOT.SUM(řádky; sloupce; hodnoty)\n=PIVOT.AVERAGE(řádky; sloupce; hodnoty)\n=PIVOT.PERCENTILE(řádky; sloupce; hodnoty; kvantil)\n=PIVOT.CONF.T(řádky; sloupce; hodnoty; alfa; [smer])</code></pre>\n<h3 id=\"argumenty\">Argumenty</h3>\n<ul><li><code>řádky</code>: jeden nebo více sloupců kategorií pro řádky, vždy včetně záhlaví</li><li><code>sloupce</code>: volitelně jeden nebo více sloupců kategorií pro sloupce, vždy včetně záhlaví; může být prázdné</li><li><code>hodnoty</code>: jeden sloupec hodnot včetně záhlaví</li><li><code>kvantil</code>: hledaný percentil v intervalu <code>(0;1)</code></li><li><code>alfa</code>: hladina alfa v intervalu <code>(0;1)</code></li><li><code>smer</code>: volitelný směr pro <code>CONF.*</code>; <code>0 = oboustranný</code>, <code>-1 = levostranný</code>, <code>1 = pravostranný</code></li></ul>\n<h3 id=\"dostupne-funkce\">Dostupné funkce</h3>\n<table><thead><tr><th>Funkce</th><th>Popis</th><th>Detail</th></tr></thead><tbody><tr><td><code>PIVOT.COUNT(řádky; sloupce; hodnoty)</code></td><td>Počet neprázdných hodnot</td><td>bez detailu</td></tr><tr><td><code>PIVOT.SUM(řádky; sloupce; hodnoty)</code></td><td>Součet</td><td>bez detailu</td></tr><tr><td><code>PIVOT.AVERAGE(řádky; sloupce; hodnoty)</code></td><td>Aritmetický průměr</td><td>bez detailu</td></tr><tr><td><code>PIVOT.MIN(řádky; sloupce; hodnoty)</code></td><td>Minimum</td><td>bez detailu</td></tr><tr><td><code>PIVOT.MAX(řádky; sloupce; hodnoty)</code></td><td>Maximum</td><td>bez detailu</td></tr><tr><td><code>PIVOT.MEDIAN(řádky; sloupce; hodnoty)</code></td><td>Medián</td><td>bez detailu</td></tr><tr><td><code>PIVOT.PERCENTILE(řádky; sloupce; hodnoty; kvantil)</code></td><td>Percentil</td><td><code>0 &lt; kvantil &lt; 1</code></td></tr><tr><td><code>PIVOT.STDEV.S(řádky; sloupce; hodnoty)</code></td><td>Výběrová směrodatná odchylka</td><td>bez detailu</td></tr><tr><td><code>PIVOT.STDEV.P(řádky; sloupce; hodnoty)</code></td><td>Populační směrodatná odchylka</td><td>bez detailu</td></tr><tr><td><code>PIVOT.VAR.S(řádky; sloupce; hodnoty)</code></td><td>Výběrový rozptyl</td><td>bez detailu</td></tr><tr><td><code>PIVOT.VAR.P(řádky; sloupce; hodnoty)</code></td><td>Populační rozptyl</td><td>bez detailu</td></tr><tr><td><code>PIVOT.VARCOEF.S(řádky; sloupce; hodnoty)</code></td><td>Výběrový variační koeficient</td><td>bez detailu</td></tr><tr><td><code>PIVOT.VARCOEF.P(řádky; sloupce; hodnoty)</code></td><td>Populační variační koeficient</td><td>bez detailu</td></tr><tr><td><code>PIVOT.CONF.T(řádky; sloupce; hodnoty; alfa; [smer])</code></td><td>Poloviční šířka intervalu spolehlivosti pro t-rozdělení</td><td><code>0 &lt; alfa &lt; 1</code>; <code>smer</code>: <code>0</code>, <code>-1</code>, <code>1</code></td></tr><tr><td><code>PIVOT.CONF.NORM(řádky; sloupce; hodnoty; alfa; [smer])</code></td><td>Poloviční šířka intervalu spolehlivosti pro normální aproximaci</td><td><code>0 &lt; alfa &lt; 1</code>; <code>smer</code>: <code>0</code>, <code>-1</code>, <code>1</code></td></tr><tr><td><code>PIVOT.MAD(řádky; sloupce; hodnoty)</code></td><td>Medián absolutních odchylek od mediánu</td><td>bez detailu</td></tr><tr><td><code>PIVOT.IQR(řádky; sloupce; hodnoty)</code></td><td>Mezikvartilové rozpětí</td><td>bez detailu</td></tr></tbody></table>\n<h3 id=\"validace-dat\">Validace dat</h3>\n<ul><li><code>řádky</code>, <code>sloupce</code> i <code>hodnoty</code> se očekávají se záhlavím v prvním řádku</li><li><code>count</code> může pracovat i s nenumerickými hodnotami; počítá neprázdné buňky</li><li>všechny ostatní funkce vyžadují kvantitativní data</li><li><code>varcoef.s</code> a <code>varcoef.p</code> vyžadují nenulový průměr</li><li><code>stdev.s</code>, <code>var.s</code>, <code>conf.t</code> a <code>conf.norm</code> vyžadují alespoň dva platné numerické záznamy</li><li>nekompatibilní kombinace typu dat a výpočtu vracejí chybu hodnoty</li><li><code>conf.*</code> akceptují jen <code>smer</code> z množiny <code>{-1, 0, 1}</code></li></ul>\n<h3 id=\"poznamky\">Poznámky</h3>\n<ul><li><code>sloupce</code> může být prázdné; v tom případě vznikne jednorozměrný pivot s jediným sloupcovým blokem <code>Celkem</code></li><li>řádkové i sloupcové kategorie se ve výstupu řadí abecedně</li><li>výstup vždy obsahuje souhrnný řádek <code>CELKEM</code> a souhrnný sloupec <code>CELKEM</code></li><li>pravá dolní buňka obsahuje celkový agregát zvoleného ukazatele za všechna data</li><li>řádky, pro které jsou všechny výsledné hodnoty prázdné, se z výstupu vynechají</li><li>sloupce, pro které jsou všechny výsledné hodnoty prázdné, se z výstupu vynechají s výjimkou souhrnného sloupce <code>CELKEM</code></li><li><code>conf.t</code> a <code>conf.norm</code> vracejí poloviční šířku intervalu, ne obě meze zvlášť</li><li>při <code>smer = 0</code> se používá kritická hodnota <code>1 - alfa/2</code></li><li>při <code>smer = -1</code> nebo <code>1</code> se používá kritická hodnota <code>1 - alfa</code></li></ul>\n<h3 id=\"vystup\">Výstup</h3>\n<p>Víceřádkový spill výstup ve stylu pivotu:</p>\n<ul><li>vlevo jsou sloupce řádkových kategorií</li><li>nahoře jsou úrovně sloupcových kategorií</li><li>poslední řádek záhlaví nad numerickými sloupci obsahuje názvy řádkových proměnných</li><li>poslední sloupec je vždy souhrnný sloupec <code>CELKEM</code></li><li>poslední řádek je vždy souhrnný řádek <code>CELKEM</code></li></ul>\n<h3 id=\"priklady\">Příklady</h3>\n<pre><code class=\"language-excel\">=PIVOT.SUM(E:E;F:F;G:G)\n=PIVOT.AVERAGE(E:E;F:F;G:G)\n=PIVOT.PERCENTILE(E:E;F:F;G:G;0,9)\n=PIVOT.CONF.T(E:E;F:F;G:G;0,05)\n=PIVOT.CONF.T(E:E;F:F;G:G;0,05;1)</code></pre>"
    },
    {
      "lang": "cs",
      "slug": "functions/two-sample-tests",
      "section": "Functions",
      "title": "Dvouvýběrové Testy",
      "summary": "Dvouvýběrové Testy WELCH.TEST.2S.G Provádí Welchův dvouvýběrový t test pro dvě nezávislé skupiny. Syntaxe Argumenty categories: štítky definující právě dvě skupiny values: číselná",
      "sourcePath": "docs/cs/functions/two-sample-tests.md",
      "html": "<h1 id=\"dvouvyberove-testy\">Dvouvýběrové Testy</h1>\n<h2 id=\"welchtest2sg\"><code>WELCH.TEST.2S.G</code></h2>\n<p>Provádí Welchův dvouvýběrový t-test pro dvě nezávislé skupiny.</p>\n<h3 id=\"syntaxe\">Syntaxe</h3>\n<pre><code class=\"language-excel\">=WELCH.TEST.2S.G(categories; values; [ma_záhlaví]; [alpha]; [direction])</code></pre>\n<h3 id=\"argumenty\">Argumenty</h3>\n<ul><li><code>categories</code>: štítky definující právě dvě skupiny</li><li><code>values</code>: číselná pozorování</li><li><code>ma_záhlaví</code>: volitelný kód režimu záhlaví; výchozí hodnota je <code>0</code></li><li><code>alpha</code>: hladina významnosti</li><li><code>direction</code>: volitelný kód směru testu; výchozí hodnota je <code>0</code></li></ul>\n<h3 id=\"kody-direction\">Kódy <code>direction</code></h3>\n<table><thead><tr><th>Kód</th><th>Význam</th></tr></thead><tbody><tr><td><code>0</code></td><td>oboustranný test</td></tr><tr><td><code>1</code></td><td>levostranný test</td></tr><tr><td><code>2</code></td><td>pravostranný test</td></tr></tbody></table>\n<h3 id=\"kody-ma-zahlavi\">Kódy <code>ma_záhlaví</code></h3>\n<table><thead><tr><th>Kód</th><th>Význam</th></tr></thead><tbody><tr><td><code>0</code></td><td>autodetekce záhlaví</td></tr><tr><td><code>1</code></td><td>první řádek je záhlaví</td></tr><tr><td><code>2</code></td><td>vstup je bez záhlaví</td></tr></tbody></table>\n<h3 id=\"poznamky\">Poznámky</h3>\n<ul><li>funkce vyžaduje právě dvě skupiny</li><li>v každé skupině musí být alespoň dvě hodnoty</li><li>skupiny jsou interně seřazeny podle názvu, což ovlivní znaménko rozdílu i statistiky <code>t</code></li><li>ve výstupu jsou popisné statistiky zhuštěny do tabulky se skupinami po řádcích</li></ul>\n<h3 id=\"vystup\">Výstup</h3>\n<p>Spill výstup obsahuje:</p>\n<ul><li>tabulku popisných statistik po skupinách</li><li><code>α</code></li><li><code>t</code></li><li>Welch-Satterthwaite <code>df</code></li><li>kritickou hodnotu <code>t</code></li><li><code>p</code></li><li>Cohenovo <code>d</code></li><li>velikost účinku <code>r</code></li></ul>\n<h3 id=\"priklad\">Příklad</h3>\n<pre><code class=\"language-excel\">=WELCH.TEST.2S.G(A2:A40;B2:B40;1;0,05;0)</code></pre>\n<h2 id=\"mannwhitneyg\"><code>MANN.WHITNEY.G</code></h2>\n<p>Provádí Mann-Whitneyho neparametrický test pro dvě nezávislé skupiny.</p>\n<h3 id=\"syntaxe\">Syntaxe</h3>\n<pre><code class=\"language-excel\">=MANN.WHITNEY.G(kategorie; hodnoty; [ma_záhlaví]; [alpha]; [smer])</code></pre>\n<h3 id=\"argumenty\">Argumenty</h3>\n<ul><li><code>kategorie</code>: štítky právě dvou skupin</li><li><code>hodnoty</code>: číselná pozorování</li><li><code>ma_záhlaví</code>: volitelný kód režimu záhlaví; výchozí hodnota je <code>0</code></li><li><code>alpha</code>: hladina významnosti</li><li><code>smer</code>: volitelný kód směru testu; výchozí hodnota je <code>0</code></li></ul>\n<h3 id=\"kody-smer\">Kódy <code>smer</code></h3>\n<table><thead><tr><th>Kód</th><th>Význam</th></tr></thead><tbody><tr><td><code>0</code></td><td>oboustranný test</td></tr><tr><td><code>1</code></td><td>levostranný test</td></tr><tr><td><code>2</code></td><td>pravostranný test</td></tr></tbody></table>\n<h3 id=\"kody-ma-zahlavi\">Kódy <code>ma_záhlaví</code></h3>\n<table><thead><tr><th>Kód</th><th>Význam</th></tr></thead><tbody><tr><td><code>0</code></td><td>autodetekce záhlaví</td></tr><tr><td><code>1</code></td><td>první řádek je záhlaví</td></tr><tr><td><code>2</code></td><td>vstup je bez záhlaví</td></tr></tbody></table>\n<h3 id=\"poznamky\">Poznámky</h3>\n<ul><li>funkce vyžaduje právě dvě skupiny</li><li>v každé skupině musí být alespoň jedna hodnota</li><li>používají se střední pořadí při shodách</li><li>při nulové rozptylové složce po tie correction vrátí funkce numerickou chybu</li></ul>\n<h3 id=\"vystup\">Výstup</h3>\n<p>Spill výstup obsahuje:</p>\n<ul><li>tabulku popisných statistik po skupinách</li><li><code>U</code></li><li><code>U₁</code></li><li><code>U₂</code></li><li><code>z</code></li><li>kritickou hodnotu <code>z</code></li><li><code>p</code></li><li>efekt velikosti <code>r</code></li></ul>\n<h3 id=\"priklad\">Příklad</h3>\n<pre><code class=\"language-excel\">=MANN.WHITNEY.G(A2:A20;B2:B20)\n=MANN.WHITNEY.G(A2:A20;B2:B20;1;0,05;0)</code></pre>"
    },
    {
      "lang": "cs",
      "slug": "functions/variation-coefficients",
      "section": "Functions",
      "title": "Variační Koeficienty",
      "summary": "Variační Koeficienty VARCOEF Počítá populační variační koeficient. Syntaxe Poznámky prázdné buňky se ignorují je potřeba alespoň jedna platná hodnota pokud je průměr roven nule, fu",
      "sourcePath": "docs/cs/functions/variation-coefficients.md",
      "html": "<h1 id=\"variacni-koeficienty\">Variační Koeficienty</h1>\n<h2 id=\"varcoef\"><code>VARCOEF</code></h2>\n<p>Počítá populační variační koeficient.</p>\n<h3 id=\"syntaxe\">Syntaxe</h3>\n<pre><code class=\"language-excel\">=VARCOEF(values)</code></pre>\n<h3 id=\"poznamky\">Poznámky</h3>\n<ul><li>prázdné buňky se ignorují</li><li>je potřeba alespoň jedna platná hodnota</li><li>pokud je průměr roven nule, funkce vrátí chybu dělení nulou</li></ul>\n<h3 id=\"vystup\">Výstup</h3>\n<p>Skalární <code>σ / μ</code>.</p>\n<h2 id=\"varcoefs\"><code>VARCOEF.S</code></h2>\n<p>Počítá výběrový variační koeficient.</p>\n<h3 id=\"syntaxe\">Syntaxe</h3>\n<pre><code class=\"language-excel\">=VARCOEF.S(values)</code></pre>\n<h3 id=\"poznamky\">Poznámky</h3>\n<ul><li>prázdné buňky se ignorují</li><li>jsou potřeba alespoň dvě platné hodnoty</li><li>pokud je průměr roven nule, funkce vrátí chybu dělení nulou</li></ul>\n<h3 id=\"vystup\">Výstup</h3>\n<p>Skalární <code>sₓ / x̄</code>.</p>\n<h2 id=\"varcoefw\"><code>VARCOEF.W</code></h2>\n<p>Počítá vážený populační variační koeficient.</p>\n<h3 id=\"syntaxe\">Syntaxe</h3>\n<pre><code class=\"language-excel\">=VARCOEF.W(values; weights)</code></pre>\n<h3 id=\"poznamky\">Poznámky</h3>\n<ul><li><code>values</code> a <code>weights</code> musí mít stejnou délku</li><li>váhy musí být nezáporné</li><li>prázdné váhy se berou jako <code>0</code></li><li>pokud je vážený průměr roven nule, funkce vrátí chybu dělení nulou</li></ul>\n<h3 id=\"vystup\">Výstup</h3>\n<p>Skalární <code>σ_w / x̄_w</code>.</p>\n<h2 id=\"varcoefsw\"><code>VARCOEF.S.W</code></h2>\n<p>Počítá vážený výběrový variační koeficient.</p>\n<h3 id=\"syntaxe\">Syntaxe</h3>\n<pre><code class=\"language-excel\">=VARCOEF.S.W(values; weights)</code></pre>\n<h3 id=\"poznamky\">Poznámky</h3>\n<ul><li><code>values</code> a <code>weights</code> musí mít stejnou délku</li><li>váhy musí být nezáporné</li><li>prázdné váhy se berou jako <code>0</code></li><li>pokud <code>Σw &lt;= 1</code>, funkce vrátí chybu <code>#POČET!</code></li><li>pokud je vážený průměr roven nule, funkce vrátí chybu dělení nulou</li></ul>\n<h3 id=\"vystup\">Výstup</h3>\n<p>Skalární <code>sₓ,w / x̄_w</code>.</p>"
    },
    {
      "lang": "cs",
      "slug": "functions/weighted-means",
      "section": "Functions",
      "title": "Vážené Průměry",
      "summary": "Vážené Průměry AVERAGE.W Počítá vážený aritmetický průměr. Syntaxe Argumenty values: číselná pozorování weights: nezáporné váhy; prázdné buňky vah se berou jako 0 Poznámky rozsahy",
      "sourcePath": "docs/cs/functions/weighted-means.md",
      "html": "<h1 id=\"vazene-prumery\">Vážené Průměry</h1>\n<h2 id=\"averagew\"><code>AVERAGE.W</code></h2>\n<p>Počítá vážený aritmetický průměr.</p>\n<h3 id=\"syntaxe\">Syntaxe</h3>\n<pre><code class=\"language-excel\">=AVERAGE.W(values; weights)</code></pre>\n<h3 id=\"argumenty\">Argumenty</h3>\n<ul><li><code>values</code>: číselná pozorování</li><li><code>weights</code>: nezáporné váhy; prázdné buňky vah se berou jako <code>0</code></li></ul>\n<h3 id=\"poznamky\">Poznámky</h3>\n<ul><li>rozsahy <code>values</code> a <code>weights</code> musí mít stejnou délku</li><li>prázdné buňky ve <code>values</code> se přeskočí; odpovídající váha se tím také vynechá</li><li>pokud jsou všechny váhy nulové nebo některá váha záporná, funkce vrátí numerickou chybu</li></ul>\n<h3 id=\"vystup\">Výstup</h3>\n<p>Skalární vážený aritmetický průměr.</p>\n<h3 id=\"priklad\">Příklad</h3>\n<pre><code class=\"language-excel\">=AVERAGE.W(A2:A10;B2:B10)</code></pre>\n<h2 id=\"harmeanw\"><code>HARMEAN.W</code></h2>\n<p>Počítá vážený harmonický průměr.</p>\n<h3 id=\"syntaxe\">Syntaxe</h3>\n<pre><code class=\"language-excel\">=HARMEAN.W(values; weights)</code></pre>\n<h3 id=\"argumenty\">Argumenty</h3>\n<ul><li><code>values</code>: kladná číselná pozorování</li><li><code>weights</code>: nezáporné váhy; prázdné buňky vah se berou jako <code>0</code></li></ul>\n<h3 id=\"poznamky\">Poznámky</h3>\n<ul><li>rozsahy <code>values</code> a <code>weights</code> musí mít stejnou délku</li><li>hodnoty s nulovou vahou do výsledku nevstupují</li><li>pokud má některá hodnota s kladnou vahou hodnotu <code>&lt;= 0</code>, funkce vrátí numerickou chybu</li></ul>\n<h3 id=\"vystup\">Výstup</h3>\n<p>Skalární vážený harmonický průměr.</p>\n<h2 id=\"geomeanw\"><code>GEOMEAN.W</code></h2>\n<p>Počítá vážený geometrický průměr.</p>\n<h3 id=\"syntaxe\">Syntaxe</h3>\n<pre><code class=\"language-excel\">=GEOMEAN.W(values; weights)</code></pre>\n<h3 id=\"argumenty\">Argumenty</h3>\n<ul><li><code>values</code>: kladná číselná pozorování</li><li><code>weights</code>: nezáporné váhy; prázdné buňky vah se berou jako <code>0</code></li></ul>\n<h3 id=\"poznamky\">Poznámky</h3>\n<ul><li>rozsahy <code>values</code> a <code>weights</code> musí mít stejnou délku</li><li>hodnoty s nulovou vahou do výsledku nevstupují</li><li>pokud má některá hodnota s kladnou vahou hodnotu <code>&lt;= 0</code>, funkce vrátí numerickou chybu</li></ul>\n<h3 id=\"vystup\">Výstup</h3>\n<p>Skalární vážený geometrický průměr.</p>"
    },
    {
      "lang": "cs",
      "slug": "functions/weighted-variance",
      "section": "Functions",
      "title": "Vážený Rozptyl A Směrodatná Odchylka",
      "summary": "Vážený Rozptyl A Směrodatná Odchylka Společné chování pro všechny čtyři funkce: values a weights musí mít stejnou délku váhy musí být nezáporné prázdné buňky ve vahách se berou jak",
      "sourcePath": "docs/cs/functions/weighted-variance.md",
      "html": "<h1 id=\"vazeny-rozptyl-a-smerodatna-odchylka\">Vážený Rozptyl A Směrodatná Odchylka</h1>\n<p>Společné chování pro všechny čtyři funkce:</p>\n<ul><li><code>values</code> a <code>weights</code> musí mít stejnou délku</li><li>váhy musí být nezáporné</li><li>prázdné buňky ve vahách se berou jako <code>0</code></li><li>pokud je součet vah nulový, funkce vrátí numerickou chybu</li></ul>\n<h2 id=\"varpw\"><code>VAR.P.W</code></h2>\n<p>Počítá vážený populační rozptyl.</p>\n<h3 id=\"syntaxe\">Syntaxe</h3>\n<pre><code class=\"language-excel\">=VAR.P.W(values; weights)</code></pre>\n<h3 id=\"vystup\">Výstup</h3>\n<p>Skalární vážený populační rozptyl.</p>\n<h3 id=\"poznamky\">Poznámky</h3>\n<ul><li>používá jmenovatel <code>Σw</code></li></ul>\n<h2 id=\"varsw\"><code>VAR.S.W</code></h2>\n<p>Počítá vážený výběrový rozptyl.</p>\n<h3 id=\"syntaxe\">Syntaxe</h3>\n<pre><code class=\"language-excel\">=VAR.S.W(values; weights)</code></pre>\n<h3 id=\"vystup\">Výstup</h3>\n<p>Skalární vážený výběrový rozptyl.</p>\n<h3 id=\"poznamky\">Poznámky</h3>\n<ul><li>používá jmenovatel <code>Σw - 1</code></li><li>pokud <code>Σw &lt;= 1</code>, funkce vrátí chybu <code>#POČET!</code></li></ul>\n<h2 id=\"stdevpw\"><code>STDEV.P.W</code></h2>\n<p>Počítá váženou populační směrodatnou odchylku.</p>\n<h3 id=\"syntaxe\">Syntaxe</h3>\n<pre><code class=\"language-excel\">=STDEV.P.W(values; weights)</code></pre>\n<h3 id=\"vystup\">Výstup</h3>\n<p>Skalární vážená populační směrodatná odchylka <code>σ</code>.</p>\n<h3 id=\"poznamky\">Poznámky</h3>\n<ul><li>je definována jako druhá odmocnina z <code>VAR.P.W</code></li></ul>\n<h2 id=\"stdevsw\"><code>STDEV.S.W</code></h2>\n<p>Počítá váženou výběrovou směrodatnou odchylku.</p>\n<h3 id=\"syntaxe\">Syntaxe</h3>\n<pre><code class=\"language-excel\">=STDEV.S.W(values; weights)</code></pre>\n<h3 id=\"vystup\">Výstup</h3>\n<p>Skalární vážená výběrová směrodatná odchylka <code>sₓ</code>.</p>\n<h3 id=\"poznamky\">Poznámky</h3>\n<ul><li>je definována jako druhá odmocnina z <code>VAR.S.W</code></li></ul>"
    },
    {
      "lang": "cs",
      "slug": "overview",
      "section": "Overview",
      "title": "Dokumentace XLStatUDF",
      "summary": "Dokumentace XLStatUDF Tato větev obsahuje uživatelskou dokumentaci doplňku XLStatUDF v češtině. Hlavní build Pro běžné testování v Excelu používej tento soubor: artifacts/main/publ",
      "sourcePath": "docs/cs/README.md",
      "html": "<h1 id=\"dokumentace-xlstatudf\">Dokumentace XLStatUDF</h1>\n<p>Tato větev obsahuje uživatelskou dokumentaci doplňku XLStatUDF v češtině.</p>\n<h2 id=\"hlavni-build\">Hlavní build</h2>\n<p>Pro běžné testování v Excelu používej tento soubor:</p>\n<ul><li><a href=\"#\"><code>artifacts/main/publish/XLStatUDF-AddIn64-packed.xll</code></a></li></ul>\n<h2 id=\"obsah\">Obsah</h2>\n<ul><li><a href=\"/cs/docs/framework\">Přehled frameworku</a></li><li><a href=\"/cs/docs/functions/index\">Index funkcí</a></li></ul>\n<h2 id=\"referencni-podklady\">Referenční podklady</h2>\n<p>Původní zadávací dokumenty jsou uložené zde:</p>\n<ul><li><a href=\"#\"><code>reference/project-spec/xlstatudf_copilot_prompt.md</code></a></li><li><a href=\"#\"><code>reference/project-spec/xlstatudf_udf_dokumentace.md</code></a></li></ul>"
    }
  ],
  "en": [
    {
      "lang": "en",
      "slug": "framework",
      "section": "framework",
      "title": "Framework Overview",
      "summary": "Framework Overview Purpose XLStatUDF is an Excel add in that provides statistical user defined functions implemented in C with Excel DNA. Runtime target runtime: .NET 8 Excel integ",
      "sourcePath": "docs/en/framework.md",
      "html": "<h1 id=\"framework-overview\">Framework Overview</h1>\n<h2 id=\"purpose\">Purpose</h2>\n<p>XLStatUDF is an Excel add-in that provides statistical user-defined functions implemented in C# with Excel-DNA.</p>\n<h2 id=\"runtime\">Runtime</h2>\n<ul><li>target runtime: <code>.NET 8</code></li><li>Excel integration: <code>Excel-DNA</code></li><li>numeric/statistical library: <code>MathNet.Numerics</code></li><li>output for Excel: packed <code>.xll</code> add-in for 64-bit Excel</li></ul>\n<h2 id=\"main-build-path\">Main Build Path</h2>\n<p>The only build file intended for regular Excel testing is:</p>\n<ul><li><a href=\"#\"><code>artifacts/main/publish/XLStatUDF-AddIn64-packed.xll</code></a></li></ul>\n<h2 id=\"header-mode\">Header Mode</h2>\n<p>Relevant functions support an optional final argument <code>ma_záhlaví</code>:</p>\n<table><thead><tr><th>Code</th><th>Meaning</th></tr></thead><tbody><tr><td><code>0</code></td><td>auto-detect header</td></tr><tr><td><code>1</code></td><td>header is present</td></tr><tr><td><code>2</code></td><td>header is not present</td></tr></tbody></table>"
    },
    {
      "lang": "en",
      "slug": "functions/ancova",
      "section": "Functions",
      "title": "ANCOVA",
      "summary": "ANCOVA ANCOVA.G Performs analysis of covariance on grouped data with one factor and one or more covariates. Syntax Arguments factor: factor categories dependent variable: dependent",
      "sourcePath": "docs/en/functions/ancova.md",
      "html": "<h1 id=\"ancova\">ANCOVA</h1>\n<h2 id=\"ancovag\"><code>ANCOVA.G</code></h2>\n<p>Performs analysis of covariance on grouped data with one factor and one or more covariates.</p>\n<h3 id=\"syntax\">Syntax</h3>\n<pre><code class=\"language-excel\">=ANCOVA.G(factor; dependent_variable; covariates; [post_hoc]; [alpha]; [has_header])</code></pre>\n<h3 id=\"arguments\">Arguments</h3>\n<ul><li><code>factor</code>: factor categories</li><li><code>dependent_variable</code>: dependent variable</li><li><code>covariates</code>: one or more covariates arranged in columns</li><li><code>post_hoc</code>: optional post-hoc procedure code; default is <code>0</code></li><li><code>alpha</code>: significance level</li><li><code>has_header</code>: optional header mode code; default is <code>0</code></li></ul>\n<h3 id=\"post-hoc-codes\"><code>post_hoc</code> Codes</h3>\n<table><thead><tr><th>Code</th><th>Name</th><th>Description</th></tr></thead><tbody><tr><td><code>0</code></td><td><code>none</code></td><td>no post-hoc comparison</td></tr><tr><td><code>1</code></td><td><code>tukey</code></td><td>conservative Bonferroni fallback</td></tr><tr><td><code>2</code></td><td><code>bonferroni</code></td><td>pairwise comparisons of adjusted means with Bonferroni correction</td></tr><tr><td><code>3</code></td><td><code>scheffe</code></td><td>Scheffé-style approximation over adjusted means</td></tr><tr><td><code>4</code></td><td><code>games-howell</code></td><td>currently implemented as a Bonferroni fallback</td></tr></tbody></table>\n<h3 id=\"has-header-codes\"><code>has_header</code> Codes</h3>\n<table><thead><tr><th>Code</th><th>Meaning</th></tr></thead><tbody><tr><td><code>0</code></td><td>autodetect header</td></tr><tr><td><code>1</code></td><td>first row is a header</td></tr><tr><td><code>2</code></td><td>input has no header</td></tr></tbody></table>\n<h3 id=\"notes\">Notes</h3>\n<ul><li>the function requires at least two groups</li><li>covariates are supplied as one or more columns</li><li>incomplete rows are skipped as complete-case rows</li><li>adjusted means are computed at the global means of the covariates</li><li>the main table includes interactions <code>group × covariate</code>; if an interaction is significant, a warning about violated slope homogeneity is shown</li><li>effect sizes <code>η²</code>, <code>η²p</code>, <code>ω²</code>, and <code>ω²p</code> are returned for the factor, covariates, and interactions</li></ul>\n<h3 id=\"output\">Output</h3>\n<p>The spill output contains:</p>\n<ul><li>descriptive statistics by group</li><li>one common ANCOVA table for the factor, individual covariates, and interactions</li><li>an optional warning about violated homogeneity of regression slopes</li><li>adjusted means by group</li><li>an optional post-hoc section</li></ul>\n<h3 id=\"example\">Example</h3>\n<pre><code class=\"language-excel\">=ANCOVA.G(A2:A100;B2:B100;C2:D100;2;0,05;1)</code></pre>"
    },
    {
      "lang": "en",
      "slug": "functions/anova",
      "section": "Functions",
      "title": "ANOVA",
      "summary": "ANOVA ANOVA.G Performs one way analysis of variance on grouped data. Syntax Arguments categories: group labels values: numeric observations has header: 0=autodetect, 1=first row is",
      "sourcePath": "docs/en/functions/anova.md",
      "html": "<h1 id=\"anova\">ANOVA</h1>\n<h2 id=\"anovag\"><code>ANOVA.G</code></h2>\n<p>Performs one-way analysis of variance on grouped data.</p>\n<h3 id=\"syntax\">Syntax</h3>\n<pre><code class=\"language-excel\">=ANOVA.G(categories; values; [has_header]; [alpha]; [post_hoc])</code></pre>\n<h3 id=\"arguments\">Arguments</h3>\n<ul><li><code>categories</code>: group labels</li><li><code>values</code>: numeric observations</li><li><code>has_header</code>: <code>0=autodetect</code>, <code>1=first row is a header</code>, <code>2=no header</code></li><li><code>alpha</code>: significance level</li><li><code>post_hoc</code>: optional post-hoc procedure code; default is <code>0</code></li></ul>\n<h3 id=\"post-hoc-codes\"><code>post_hoc</code> Codes</h3>\n<table><thead><tr><th>Code</th><th>Name</th><th>Description</th></tr></thead><tbody><tr><td><code>0</code></td><td><code>none</code></td><td>no post-hoc comparison; only the main ANOVA report is returned</td></tr><tr><td><code>1</code></td><td><code>tukey</code></td><td>Tukey HSD; currently implemented as a conservative Bonferroni fallback</td></tr><tr><td><code>2</code></td><td><code>bonferroni</code></td><td>pairwise comparisons with Bonferroni correction</td></tr><tr><td><code>3</code></td><td><code>scheffe</code></td><td>Scheffé method for multiple comparisons</td></tr><tr><td><code>4</code></td><td><code>games-howell</code></td><td>pairwise comparisons without assuming equal variances</td></tr></tbody></table>\n<h3 id=\"notes\">Notes</h3>\n<ul><li>the function requires at least three groups</li><li>each group must contain at least two values</li><li>the output includes Levene&#39;s test of variance homogeneity</li><li>some post-hoc options are currently implemented via Bonferroni fallback and are explicitly labeled as such in the output</li></ul>\n<h3 id=\"output\">Output</h3>\n<p>The spill output contains:</p>\n<ul><li>descriptive statistics by group</li><li>the main ANOVA table</li><li>Levene&#39;s homogeneity test</li><li>effect sizes <code>η²</code>, <code>ω²</code>, and <code>f</code></li><li>an optional post-hoc section</li></ul>\n<h3 id=\"examples\">Examples</h3>\n<pre><code class=\"language-excel\">=ANOVA.G(A2:A40;B2:B40)\n=ANOVA.G(A2:A40;B2:B40;1;0,05;2)\n=ANOVA.G(A2:A40;B2:B40;0;0,05;4)</code></pre>\n<h2 id=\"anovarm\"><code>ANOVA.RM</code></h2>\n<p>Performs one-factor repeated-measures ANOVA where columns represent conditions and rows represent subjects.</p>\n<h3 id=\"syntax\">Syntax</h3>\n<pre><code class=\"language-excel\">=ANOVA.RM(values; [has_header]; [alpha]; [post_hoc])</code></pre>\n<h3 id=\"arguments\">Arguments</h3>\n<ul><li><code>values</code>: value matrix; rows are subjects and columns are conditions</li><li><code>has_header</code>: <code>0=autodetect</code>, <code>1=first row is a header</code>, <code>2=no header</code></li><li><code>alpha</code>: significance level</li><li><code>post_hoc</code>: optional post-hoc procedure code; default is <code>0</code></li></ul>\n<h3 id=\"post-hoc-codes\"><code>post_hoc</code> Codes</h3>\n<table><thead><tr><th>Code</th><th>Name</th><th>Description</th></tr></thead><tbody><tr><td><code>0</code></td><td><code>none</code></td><td>no post-hoc comparison; only the main RM ANOVA report is returned</td></tr><tr><td><code>1</code></td><td><code>tukey</code></td><td>currently implemented as a conservative Bonferroni fallback</td></tr><tr><td><code>2</code></td><td><code>bonferroni</code></td><td>pairwise condition comparisons via paired t-tests with Bonferroni correction</td></tr><tr><td><code>3</code></td><td><code>scheffe</code></td><td>currently implemented as a conservative Bonferroni fallback</td></tr><tr><td><code>4</code></td><td><code>games-howell</code></td><td>currently implemented as a conservative Bonferroni fallback</td></tr></tbody></table>\n<h3 id=\"notes\">Notes</h3>\n<ul><li>the function requires at least two columns and at least two complete rows</li><li>incomplete rows are skipped as complete-case rows</li><li>sphericity is not currently tested; this is also stated in the output</li></ul>\n<h3 id=\"output\">Output</h3>\n<p>The spill output contains:</p>\n<ul><li>descriptive statistics by condition</li><li>an RM ANOVA table with rows <code>Conditions</code>, <code>Subjects</code>, <code>Residual</code>, <code>Total</code></li><li>effect sizes <code>η²</code>, <code>η²p</code>, <code>ω²</code>, <code>ω²p</code> for the repeated-measures factor</li><li>a note about untested sphericity</li><li>an optional post-hoc section with pairwise condition comparisons</li></ul>\n<h3 id=\"example\">Example</h3>\n<pre><code class=\"language-excel\">=ANOVA.RM(B2:D25)\n=ANOVA.RM(B1:D25;1;0,05;2)</code></pre>"
    },
    {
      "lang": "en",
      "slug": "functions/contingency",
      "section": "Functions",
      "title": "Contingency Tables",
      "summary": "Contingency Tables CONTINGENCY.T Analyzes a contingency table entered directly as a frequency matrix. Syntax Arguments table: a 2D range of observed frequencies has header: optiona",
      "sourcePath": "docs/en/functions/contingency.md",
      "html": "<h1 id=\"contingency-tables\">Contingency Tables</h1>\n<h2 id=\"contingencyt\"><code>CONTINGENCY.T</code></h2>\n<p>Analyzes a contingency table entered directly as a frequency matrix.</p>\n<h3 id=\"syntax\">Syntax</h3>\n<pre><code class=\"language-excel\">=CONTINGENCY.T(table; [has_header]; [alpha])</code></pre>\n<h3 id=\"arguments\">Arguments</h3>\n<ul><li><code>table</code>: a 2D range of observed frequencies</li><li><code>has_header</code>: optional header mode code; default is <code>0</code></li><li><code>alpha</code>: significance level</li></ul>\n<h3 id=\"has-header-codes\"><code>has_header</code> Codes</h3>\n<table><thead><tr><th>Code</th><th>Meaning</th></tr></thead><tbody><tr><td><code>0</code></td><td>autodetect labels</td></tr><tr><td><code>1</code></td><td>top row and left column are labels</td></tr><tr><td><code>2</code></td><td>the table is purely numeric with no labels</td></tr></tbody></table>\n<h3 id=\"notes\">Notes</h3>\n<ul><li>the function requires at least a <code>2 × 2</code> table</li><li>frequencies must be non-negative integers</li><li>in autodetect mode, labels are used only if the top row and left column actually look like textual headers</li></ul>\n<h3 id=\"output\">Output</h3>\n<p>The spill output contains:</p>\n<ul><li>the observed contingency table including marginal totals</li><li>the expected frequency table</li><li>a test summary with <code>n</code>, <code>df</code>, <code>α</code>, <code>χ²</code>, critical value, and <code>p</code></li><li>association measures <code>Pearson C</code>, <code>Cramér V</code>, and <code>phi</code> for <code>2x2</code> tables</li></ul>\n<h3 id=\"example\">Example</h3>\n<pre><code class=\"language-excel\">=CONTINGENCY.T(A1:C3;1;0,05)</code></pre>\n<h2 id=\"contingencyg\"><code>CONTINGENCY.G</code></h2>\n<p>Analyzes a contingency table built from grouped data.</p>\n<h3 id=\"syntax\">Syntax</h3>\n<pre><code class=\"language-excel\">=CONTINGENCY.G(columns; rows; [count]; [alpha]; [has_header])</code></pre>\n<h3 id=\"arguments\">Arguments</h3>\n<ul><li><code>columns</code>: categories of the future contingency table columns</li><li><code>rows</code>: categories of the future contingency table rows</li><li><code>count</code>: optional pair frequencies; if omitted, each pair has weight <code>1</code></li><li><code>alpha</code>: significance level</li><li><code>has_header</code>: optional header mode code; default is <code>0</code></li></ul>\n<h3 id=\"has-header-codes\"><code>has_header</code> Codes</h3>\n<table><thead><tr><th>Code</th><th>Meaning</th></tr></thead><tbody><tr><td><code>0</code></td><td>autodetect header</td></tr><tr><td><code>1</code></td><td>first row is a header</td></tr><tr><td><code>2</code></td><td>input has no header</td></tr></tbody></table>\n<h3 id=\"notes\">Notes</h3>\n<ul><li>the function requires at least two distinct row categories and two distinct column categories</li><li>blank rows are skipped</li><li>if <code>count</code> is provided, it must contain non-negative integer frequencies</li><li>the output structure is the same as <code>CONTINGENCY.T</code></li></ul>\n<h3 id=\"output\">Output</h3>\n<p>The spill output contains:</p>\n<ul><li>the observed contingency table</li><li>expected frequencies</li><li>a <code>χ²</code> test summary</li><li>association measures <code>Pearson C</code>, <code>Cramér V</code>, and possibly <code>phi</code></li></ul>\n<h3 id=\"examples\">Examples</h3>\n<pre><code class=\"language-excel\">=CONTINGENCY.G(A2:A100;B2:B100)\n=CONTINGENCY.G(A2:A100;B2:B100;C2:C100;0,05;1)</code></pre>"
    },
    {
      "lang": "en",
      "slug": "functions/correlation",
      "section": "Functions",
      "title": "Correlation",
      "summary": "Correlation CORREL.SPEARMAN Computes Spearman's rank correlation coefficient and performs its significance test. Syntax Arguments x values: first numeric range y values: second num",
      "sourcePath": "docs/en/functions/correlation.md",
      "html": "<h1 id=\"correlation\">Correlation</h1>\n<h2 id=\"correlspearman\"><code>CORREL.SPEARMAN</code></h2>\n<p>Computes Spearman&#39;s rank correlation coefficient and performs its significance test.</p>\n<h3 id=\"syntax\">Syntax</h3>\n<pre><code class=\"language-excel\">=CORREL.SPEARMAN(x_values; y_values; [direction]; [alpha]; [has_header])</code></pre>\n<h3 id=\"arguments\">Arguments</h3>\n<ul><li><code>x_values</code>: first numeric range</li><li><code>y_values</code>: second numeric range</li><li><code>direction</code>: optional direction code; default is <code>0</code></li><li><code>alpha</code>: significance level</li><li><code>has_header</code>: optional header mode code; default is <code>0</code></li></ul>\n<h3 id=\"direction-codes\"><code>direction</code> Codes</h3>\n<table><thead><tr><th>Code</th><th>Meaning</th></tr></thead><tbody><tr><td><code>0</code></td><td>two-sided test</td></tr><tr><td><code>1</code></td><td>left-sided test</td></tr><tr><td><code>2</code></td><td>right-sided test</td></tr></tbody></table>\n<h3 id=\"has-header-codes\"><code>has_header</code> Codes</h3>\n<table><thead><tr><th>Code</th><th>Meaning</th></tr></thead><tbody><tr><td><code>0</code></td><td>autodetect header</td></tr><tr><td><code>1</code></td><td>first row is a header</td></tr><tr><td><code>2</code></td><td>input has no header</td></tr></tbody></table>\n<h3 id=\"notes\">Notes</h3>\n<ul><li>blank cells are excluded pairwise</li><li>at least three valid pairs are required</li><li>if either variable has zero rank variance, the function returns a numeric error</li></ul>\n<h3 id=\"output\">Output</h3>\n<p>The spill output contains:</p>\n<ul><li><code>ρ</code></li><li><code>n</code></li><li><code>α</code></li><li><code>t</code></li><li><code>df</code></li><li>critical <code>t</code></li><li><code>p</code></li></ul>\n<h2 id=\"correlmatrix\"><code>CORREL.MATRIX</code></h2>\n<p>Builds a correlation matrix for data with two or more columns.</p>\n<h3 id=\"syntax\">Syntax</h3>\n<pre><code class=\"language-excel\">=CORREL.MATRIX(data; [method]; [output]; [p_minimum]; [has_header])</code></pre>\n<h3 id=\"arguments\">Arguments</h3>\n<ul><li><code>data</code>: input matrix; columns are variables</li><li><code>method</code>: optional computation method code; default is <code>0</code></li><li><code>output</code>: optional output type code; default is <code>0</code></li><li><code>p_minimum</code>: optional filter; only links with <code>p &lt; p_minimum</code> are returned</li><li><code>has_header</code>: optional header mode code; default is <code>0</code></li></ul>\n<h3 id=\"method-codes\"><code>method</code> Codes</h3>\n<table><thead><tr><th>Code</th><th>Meaning</th></tr></thead><tbody><tr><td><code>0</code></td><td>Pearson</td></tr><tr><td><code>1</code></td><td>Spearman</td></tr></tbody></table>\n<h3 id=\"output-codes\"><code>output</code> Codes</h3>\n<table><thead><tr><th>Code</th><th>Meaning</th></tr></thead><tbody><tr><td><code>0</code></td><td>coefficients only</td></tr><tr><td><code>1</code></td><td>two-tailed p-values only</td></tr><tr><td><code>2</code></td><td>coefficient and p-value on the row below</td></tr><tr><td><code>3</code></td><td>coefficient, p-value below it, and significance stars on the third row</td></tr><tr><td><code>4</code></td><td>coefficients only, with significance stars appended in the same cell</td></tr></tbody></table>\n<h3 id=\"has-header-codes\"><code>has_header</code> Codes</h3>\n<table><thead><tr><th>Code</th><th>Meaning</th></tr></thead><tbody><tr><td><code>0</code></td><td>autodetect header</td></tr><tr><td><code>1</code></td><td>first row is a header</td></tr><tr><td><code>2</code></td><td>input has no header</td></tr></tbody></table>\n<h3 id=\"notes\">Notes</h3>\n<ul><li>the function requires at least two columns and at least three complete rows</li><li>incomplete rows are skipped as complete-case rows</li><li>if any column is constant, the function returns a numeric error</li><li>if <code>p_minimum</code> is supplied, non-matching links are left blank in the output, including diagonal cells</li><li>the add-in still accepts the legacy argument order <code>([p_minimum]; data; [method]; [output]; [has_header])</code> for backward compatibility</li></ul>\n<h3 id=\"output\">Output</h3>\n<p>For outputs <code>0</code>, <code>1</code>, and <code>4</code>, the spill output returns a square matrix with variable names.</p>\n<p>For outputs <code>2</code> and <code>3</code>, the spill output returns a stacked layout with two leading label columns and one row block per variable:</p>\n<ul><li>coefficient</li><li><code>p</code></li><li>optionally <code>sig.</code></li></ul>\n<p>For output <code>4</code>, each cell contains text in a form such as <code>0.30156***</code>.</p>\n<h3 id=\"examples\">Examples</h3>\n<pre><code class=\"language-excel\">=CORREL.MATRIX(B1:E30)\n=CORREL.MATRIX(B1:E30;1;0;;1)\n=CORREL.MATRIX(B1:E30;0;4;0,05;1)</code></pre>"
    },
    {
      "lang": "en",
      "slug": "functions/distributions",
      "section": "Functions",
      "title": "Distributions",
      "summary": "Distributions NORM.DIST.RANGE Computes the probability that a normally distributed random variable falls within a given interval. Syntax Arguments mean: distribution mean standard",
      "sourcePath": "docs/en/functions/distributions.md",
      "html": "<h1 id=\"distributions\">Distributions</h1>\n<h2 id=\"normdistrange\"><code>NORM.DIST.RANGE</code></h2>\n<p>Computes the probability that a normally distributed random variable falls within a given interval.</p>\n<h3 id=\"syntax\">Syntax</h3>\n<pre><code class=\"language-excel\">=NORM.DIST.RANGE(mean; standard_deviation; lower_bound; upper_bound)</code></pre>\n<h3 id=\"arguments\">Arguments</h3>\n<ul><li><code>mean</code>: distribution mean</li><li><code>standard_deviation</code>: positive standard deviation</li><li><code>lower_bound</code>: lower interval bound; a blank cell means minus infinity</li><li><code>upper_bound</code>: upper interval bound; a blank cell means plus infinity</li></ul>\n<h3 id=\"notes\">Notes</h3>\n<ul><li>if <code>lower_bound &gt; upper_bound</code>, the function returns a numeric error</li><li>if <code>standard_deviation &lt;= 0</code>, the function returns a numeric error</li><li>blank bounds can be used for one-sided intervals</li></ul>\n<h3 id=\"output\">Output</h3>\n<p>A scalar value in the range <code>[0;1]</code>.</p>\n<h3 id=\"example\">Example</h3>\n<pre><code class=\"language-excel\">=NORM.DIST.RANGE(0;1;-1;1)</code></pre>\n<h2 id=\"generatenorm\"><code>GENERATE.NORM</code></h2>\n<p>Generates a single random value from a normal distribution with optional perturbation.</p>\n<h3 id=\"syntax\">Syntax</h3>\n<pre><code class=\"language-excel\">=GENERATE.NORM(mean; stdev; [outlier_rate])</code></pre>\n<h3 id=\"arguments\">Arguments</h3>\n<ul><li><code>mean</code>: desired mean</li><li><code>stdev</code>: positive standard deviation</li><li><code>outlier_rate</code>: optional probability of additional random perturbation in the interval <code>&lt;0;1&gt;</code>; default is <code>0</code></li></ul>\n<h3 id=\"notes\">Notes</h3>\n<ul><li>the function is volatile, so recalculation generates a new draw</li><li>when <code>outlier_rate = 0</code>, the function returns a regular draw from <code>N(μ, σ)</code></li><li>when perturbation occurs, additional noise from <code>N(0, 3σ)</code> is added to the generated value</li><li>invalid <code>outlier_rate</code> or non-numeric inputs return a value error</li></ul>\n<h3 id=\"output\">Output</h3>\n<p>A single scalar value.</p>\n<h3 id=\"example\">Example</h3>\n<pre><code class=\"language-excel\">=GENERATE.NORM(100;15)\n=GENERATE.NORM(0;1;0,2)</code></pre>\n<h2 id=\"generateint\"><code>GENERATE.INT</code></h2>\n<p>Generates a single random integer from a given interval with optional perturbation.</p>\n<h3 id=\"syntax\">Syntax</h3>\n<pre><code class=\"language-excel\">=GENERATE.INT([minimum]; [maximum]; [outlier_rate])</code></pre>\n<h3 id=\"arguments\">Arguments</h3>\n<ul><li><code>minimum</code>: optional lower bound; default is <code>-2147483648</code></li><li><code>maximum</code>: optional upper bound; default is <code>2147483647</code></li><li><code>outlier_rate</code>: optional probability of additional random perturbation in the interval <code>&lt;0;1&gt;</code>; default is <code>0</code></li></ul>\n<h3 id=\"notes\">Notes</h3>\n<ul><li>the function is volatile, so recalculation generates a new value</li><li>if <code>minimum &gt; maximum</code>, the function returns a numeric error</li><li>if bounds are omitted, the full practical 32-bit range is used</li><li>when <code>outlier_rate = 0</code>, the function returns a regular draw from the closed interval <code>&lt;minimum; maximum&gt;</code></li><li>when perturbation occurs, an additional random integer offset from <code>⟨-(maximum-minimum); +(maximum-minimum)⟩</code> is added</li></ul>\n<h3 id=\"output\">Output</h3>\n<p>A single integer value.</p>\n<h3 id=\"examples\">Examples</h3>\n<pre><code class=\"language-excel\">=GENERATE.INT()\n=GENERATE.INT(1;6)\n=GENERATE.INT(1;6;0,2)</code></pre>\n<h2 id=\"fill\"><code>FILL</code></h2>\n<p>Repeats one or more values, or repeatedly evaluates a text formula, into a single-column spill output.</p>\n<h3 id=\"syntax\">Syntax</h3>\n<pre><code class=\"language-excel\">=FILL(what; count; [what2]; [count2]; ...)</code></pre>\n<h3 id=\"arguments\">Arguments</h3>\n<ul><li><code>what</code>: a scalar value to repeat, or a text formula starting with <code>=</code></li><li><code>count</code>: number of returned rows; integer <code>&gt;= 1</code></li><li>additional arguments must be provided as <code>what + count</code> pairs</li></ul>\n<h3 id=\"notes\">Notes</h3>\n<ul><li>if you pass a regular value, the function simply copies it into every row</li><li>if you pass a text formula starting with <code>=</code>, the formula is evaluated separately for each row</li><li>the number of additional arguments must be even; otherwise the function returns a value error</li><li>a direct argument such as <code>GENERATE.NORM(...)</code> or <code>RANDBETWEEN(...)</code> is evaluated by Excel before <code>FILL</code> is called, so repeated recalculation requires a text formula</li></ul>\n<h3 id=\"output\">Output</h3>\n<p>A single-column spill range with the resulting sequence.</p>\n<h3 id=\"examples\">Examples</h3>\n<pre><code class=\"language-excel\">=FILL(&quot;A&quot;;5)\n=FILL(123;4)\n=FILL(&quot;male&quot;;100;&quot;female&quot;;100;&quot;child&quot;;90)\n=FILL(&quot;=RANDBETWEEN(1;10)&quot;;20)\n=FILL(&quot;=GENERATE.NORM(0;1;0,2)&quot;;100)</code></pre>\n<h2 id=\"fillrandom\"><code>FILL.RANDOM</code></h2>\n<p>Builds a sequence like <code>FILL</code>, then shuffles it randomly before returning it.</p>\n<h3 id=\"syntax\">Syntax</h3>\n<pre><code class=\"language-excel\">=FILL.RANDOM(what; count; [what2]; [count2]; ...)</code></pre>\n<h3 id=\"arguments\">Arguments</h3>\n<p>Same as <code>FILL</code>.</p>\n<h3 id=\"notes\">Notes</h3>\n<ul><li>the full sequence is first created in the same way as <code>FILL</code></li><li>the completed sequence is then shuffled randomly</li><li>the same validation rules as <code>FILL</code> apply here as well</li></ul>\n<h3 id=\"output\">Output</h3>\n<p>A single-column spill range containing all generated values in random order.</p>\n<h3 id=\"examples\">Examples</h3>\n<pre><code class=\"language-excel\">=FILL.RANDOM(&quot;male&quot;;100;&quot;female&quot;;100;&quot;child&quot;;90)\n=FILL.RANDOM(&quot;A&quot;;5;&quot;B&quot;;5)</code></pre>"
    },
    {
      "lang": "en",
      "slug": "functions/goodness-of-fit",
      "section": "Functions",
      "title": "Goodness-Of-Fit Test",
      "summary": "Goodness Of Fit Test CHISQ.GOF Performs the chi square goodness of fit test. Syntax Arguments observed: observed category frequencies expected: expected frequencies or probabilitie",
      "sourcePath": "docs/en/functions/goodness-of-fit.md",
      "html": "<h1 id=\"goodness-of-fit-test\">Goodness-Of-Fit Test</h1>\n<h2 id=\"chisqgof\"><code>CHISQ.GOF</code></h2>\n<p>Performs the chi-square goodness-of-fit test.</p>\n<h3 id=\"syntax\">Syntax</h3>\n<pre><code class=\"language-excel\">=CHISQ.GOF(observed; expected; [categories]; [alpha]; [has_header])</code></pre>\n<h3 id=\"arguments\">Arguments</h3>\n<ul><li><code>observed</code>: observed category frequencies</li><li><code>expected</code>: expected frequencies or probabilities</li><li><code>categories</code>: optional category labels</li><li><code>alpha</code>: significance level</li><li><code>has_header</code>: optional header mode code; default is <code>0</code></li></ul>\n<h3 id=\"has-header-codes\"><code>has_header</code> Codes</h3>\n<table><thead><tr><th>Code</th><th>Meaning</th></tr></thead><tbody><tr><td><code>0</code></td><td>autodetect header</td></tr><tr><td><code>1</code></td><td>first row is a header</td></tr><tr><td><code>2</code></td><td>input has no header</td></tr></tbody></table>\n<h3 id=\"notes\">Notes</h3>\n<ul><li><code>observed</code> must contain non-negative integer frequencies</li><li><code>expected</code> may be provided either as frequencies or as probabilities summing to <code>1</code></li><li>if <code>expected</code> is given as probabilities, the function automatically rescales them to expected counts using the sample size</li><li>if <code>categories</code> is omitted, categories are labeled <code>1, 2, 3, ...</code></li></ul>\n<h3 id=\"output\">Output</h3>\n<p>The spill output contains:</p>\n<ul><li>a test summary with <code>χ²</code>, <code>df</code>, <code>α</code>, critical value, and <code>p</code></li><li>a category table with <code>O</code>, <code>E</code>, and contribution <code>(O−E)² / E</code></li></ul>"
    },
    {
      "lang": "en",
      "slug": "functions/index",
      "section": "Functions",
      "title": "Functions Index",
      "summary": "Functions Index General GENERATE.NORM: Generates a single random value from a normal distribution with optional perturbation. GENERATE.INT: Generates a single random integer from a",
      "sourcePath": "docs/en/functions/index.md",
      "html": "<h1 id=\"functions-index\">Functions Index</h1>\n<h2 id=\"general\">General</h2>\n<ul><li><a href=\"./distributions.md#generatenorm\">GENERATE.NORM</a>: Generates a single random value from a normal distribution with optional perturbation.</li><li><a href=\"./distributions.md#generateint\">GENERATE.INT</a>: Generates a single random integer from a specified interval with optional perturbation.</li><li><a href=\"./distributions.md#fill\">FILL</a>: Repeats values or text formulas into a single-column spill output.</li><li><a href=\"./distributions.md#fillrandom\">FILL.RANDOM</a>: Builds a sequence like <code>FILL</code> and then shuffles it randomly.</li></ul>\n<h2 id=\"descriptive\">Descriptive</h2>\n<ul><li><a href=\"./weighted-means.md#averagew\">AVERAGE.W</a>: Computes the weighted arithmetic mean.</li><li><a href=\"./weighted-means.md#harmeanw\">HARMEAN.W</a>: Computes the weighted harmonic mean.</li><li><a href=\"./weighted-means.md#geomeanw\">GEOMEAN.W</a>: Computes the weighted geometric mean.</li><li><a href=\"./weighted-variance.md#varpw\">VAR.P.W</a>: Computes the weighted population variance.</li><li><a href=\"./weighted-variance.md#varsw\">VAR.S.W</a>: Computes the weighted sample variance.</li><li><a href=\"./weighted-variance.md#stdevpw\">STDEV.P.W</a>: Computes the weighted population standard deviation.</li><li><a href=\"./weighted-variance.md#stdevsw\">STDEV.S.W</a>: Computes the weighted sample standard deviation.</li><li><a href=\"./variation-coefficients.md#varcoef\">VARCOEF</a>: Computes the population coefficient of variation.</li><li><a href=\"./variation-coefficients.md#varcoefs\">VARCOEF.S</a>: Computes the sample coefficient of variation.</li><li><a href=\"./variation-coefficients.md#varcoefw\">VARCOEF.W</a>: Computes the weighted population coefficient of variation.</li><li><a href=\"./variation-coefficients.md#varcoefsw\">VARCOEF.S.W</a>: Computes the weighted sample coefficient of variation.</li><li><a href=\"./percentiles.md#percentileincifs\">PERCENTILE.INC.IFS</a>: Computes the inclusive percentile after applying filters.</li><li><a href=\"./percentiles.md#percentileexcifs\">PERCENTILE.EXC.IFS</a>: Computes the exclusive percentile after applying filters.</li><li><a href=\"./pivot.md#pivot\">PIVOT.*</a>: Builds a statistical pivot and computes exactly one selected metric per function.</li></ul>\n<h2 id=\"tests\">Tests</h2>\n<ul><li><a href=\"./distributions.md#normdistrange\">NORM.DIST.RANGE</a>: Computes the probability of an interval under the normal distribution.</li><li><a href=\"./normality.md#shapirowilk\">SHAPIRO.WILK</a>: Performs the Shapiro-Wilk normality test.</li><li><a href=\"./normality.md#kolmogorovsmirnov\">KOLMOGOROV.SMIRNOV</a>: Performs the Kolmogorov-Smirnov goodness-of-fit test for the selected distribution.</li><li><a href=\"./one-sample-tests.md#ttest1s\">T.TEST.1S</a>: Performs a one-sample t-test against a specified hypothetical mean.</li><li><a href=\"./one-sample-tests.md#proptest1s\">PROP.TEST.1S</a>: Performs a one-sample proportion test.</li><li><a href=\"./one-sample-tests.md#wilcoxonpaired\">WILCOXON.PAIRED</a>: Performs the Wilcoxon signed-rank test for dependent measurements.</li><li><a href=\"./two-sample-tests.md#welchtest2sg\">WELCH.TEST.2S.G</a>: Performs Welch&#39;s two-sample t-test for two independent groups.</li><li><a href=\"./two-sample-tests.md#mannwhitneyg\">MANN.WHITNEY.G</a>: Performs the Mann-Whitney test for two independent groups.</li><li><a href=\"./goodness-of-fit.md#chisqgof\">CHISQ.GOF</a>: Performs the chi-square goodness-of-fit test.</li><li><a href=\"./anova.md#anovag\">ANOVA.G</a>: Performs one-way ANOVA on grouped data.</li><li><a href=\"./anova.md#anovarm\">ANOVA.RM</a>: Performs repeated-measures ANOVA across columns.</li><li><a href=\"./ancova.md#ancovag\">ANCOVA.G</a>: Performs ANCOVA with one factor and one or more covariates.</li><li><a href=\"./contingency.md#contingencyt\">CONTINGENCY.T</a>: Analyzes a contingency table entered directly as a frequency matrix.</li><li><a href=\"./contingency.md#contingencyg\">CONTINGENCY.G</a>: Analyzes a contingency table built from grouped columns.</li><li><a href=\"./correlation.md#correlspearman\">CORREL.SPEARMAN</a>: Computes Spearman correlation and its significance test.</li><li><a href=\"./correlation.md#correlmatrix\">CORREL.MATRIX</a>: Builds a correlation matrix including p-values and significance markers.</li></ul>"
    },
    {
      "lang": "en",
      "slug": "functions/normality",
      "section": "Functions",
      "title": "Normality And Fit Tests",
      "summary": "Normality And Fit Tests SHAPIRO.WILK Performs the Shapiro Wilk normality test. Syntax Arguments values: numeric sample; blank cells are ignored has header: 0=autodetect, 1=first ce",
      "sourcePath": "docs/en/functions/normality.md",
      "html": "<h1 id=\"normality-and-fit-tests\">Normality And Fit Tests</h1>\n<h2 id=\"shapirowilk\"><code>SHAPIRO.WILK</code></h2>\n<p>Performs the Shapiro-Wilk normality test.</p>\n<h3 id=\"syntax\">Syntax</h3>\n<pre><code class=\"language-excel\">=SHAPIRO.WILK(values; [has_header])</code></pre>\n<h3 id=\"arguments\">Arguments</h3>\n<ul><li><code>values</code>: numeric sample; blank cells are ignored</li><li><code>has_header</code>: <code>0=autodetect</code>, <code>1=first cell is a header</code>, <code>2=no header</code></li></ul>\n<h3 id=\"notes\">Notes</h3>\n<ul><li>the function requires a sample size from <code>3</code> to <code>5000</code></li><li>the data are sorted in ascending order before the statistic is computed</li><li>if all observations are identical, the statistic becomes degenerate and interpretation is not meaningful</li></ul>\n<h3 id=\"output\">Output</h3>\n<p>A two-row spill output:</p>\n<ul><li><code>W</code>: test statistic</li><li><code>p</code>: p-value</li></ul>\n<h3 id=\"example\">Example</h3>\n<pre><code class=\"language-excel\">=SHAPIRO.WILK(B2:B18)</code></pre>\n<h2 id=\"kolmogorovsmirnov\"><code>KOLMOGOROV.SMIRNOV</code></h2>\n<p>Performs the one-sample Kolmogorov-Smirnov goodness-of-fit test.</p>\n<h3 id=\"syntax\">Syntax</h3>\n<pre><code class=\"language-excel\">=KOLMOGOROV.SMIRNOV(values; [distribution]; [has_header])</code></pre>\n<h3 id=\"arguments\">Arguments</h3>\n<ul><li><code>values</code>: numeric sample; blank cells are ignored</li><li><code>distribution</code>: optional code of the tested distribution; default is <code>0</code></li><li><code>has_header</code>: <code>0=autodetect</code>, <code>1=first cell is a header</code>, <code>2=no header</code></li></ul>\n<h3 id=\"distribution-codes\"><code>distribution</code> Codes</h3>\n<table><thead><tr><th>Code</th><th>Distribution</th><th>What is tested</th></tr></thead><tbody><tr><td><code>0</code></td><td><code>normal</code></td><td>whether the data follow a normal distribution with mean and standard deviation estimated from the sample</td></tr><tr><td><code>1</code></td><td><code>lognormal</code></td><td>whether the data follow a lognormal distribution; all values must be positive</td></tr><tr><td><code>2</code></td><td><code>exponential</code></td><td>whether the data follow an exponential distribution; all values must be non-negative</td></tr><tr><td><code>3</code></td><td><code>uniform</code></td><td>whether the data follow a continuous uniform distribution on the interval determined by the sample minimum and maximum</td></tr><tr><td><code>4</code></td><td><code>weibull</code></td><td>whether the data follow a Weibull distribution with parameters estimated from the sample</td></tr></tbody></table>\n<h3 id=\"notes\">Notes</h3>\n<ul><li>the function requires at least <code>5</code> valid observations</li><li>the data are sorted in ascending order before the statistic is computed</li><li>some distributions impose additional input constraints:</li><li><code>lognormal</code> and <code>weibull</code>: positive values only</li><li><code>exponential</code>: non-negative values only</li><li><code>uniform</code>: the sample must not be constant</li><li>the p-value for <code>normal</code> uses a different correction than the other distributions</li></ul>\n<h3 id=\"output\">Output</h3>\n<p>A two-row spill output:</p>\n<ul><li><code>D</code>: test statistic</li><li><code>p</code>: p-value</li></ul>\n<h3 id=\"examples\">Examples</h3>\n<pre><code class=\"language-excel\">=KOLMOGOROV.SMIRNOV(B2:B18)\n=KOLMOGOROV.SMIRNOV(B2:B18;1)\n=KOLMOGOROV.SMIRNOV(B2:B18;4;1)</code></pre>"
    },
    {
      "lang": "en",
      "slug": "functions/one-sample-tests",
      "section": "Functions",
      "title": "One-Sample Tests",
      "summary": "One Sample Tests T.TEST.1S Performs a one sample t test. Syntax Arguments values: numeric sample mu 0: hypothetical mean direction: optional direction code; default is 0 alpha: sig",
      "sourcePath": "docs/en/functions/one-sample-tests.md",
      "html": "<h1 id=\"one-sample-tests\">One-Sample Tests</h1>\n<h2 id=\"ttest1s\"><code>T.TEST.1S</code></h2>\n<p>Performs a one-sample t-test.</p>\n<h3 id=\"syntax\">Syntax</h3>\n<pre><code class=\"language-excel\">=T.TEST.1S(values; mu_0; [direction]; [alpha]; [has_header])</code></pre>\n<h3 id=\"arguments\">Arguments</h3>\n<ul><li><code>values</code>: numeric sample</li><li><code>mu_0</code>: hypothetical mean</li><li><code>direction</code>: optional direction code; default is <code>0</code></li><li><code>alpha</code>: significance level</li><li><code>has_header</code>: optional header mode code; default is <code>0</code></li></ul>\n<h3 id=\"direction-codes\"><code>direction</code> Codes</h3>\n<table><thead><tr><th>Code</th><th>Meaning</th></tr></thead><tbody><tr><td><code>0</code></td><td>two-sided test</td></tr><tr><td><code>1</code></td><td>left-sided test</td></tr><tr><td><code>2</code></td><td>right-sided test</td></tr></tbody></table>\n<h3 id=\"has-header-codes\"><code>has_header</code> Codes</h3>\n<table><thead><tr><th>Code</th><th>Meaning</th></tr></thead><tbody><tr><td><code>0</code></td><td>autodetect header</td></tr><tr><td><code>1</code></td><td>first cell is a header</td></tr><tr><td><code>2</code></td><td>input has no header</td></tr></tbody></table>\n<h3 id=\"notes\">Notes</h3>\n<ul><li>blank cells are ignored</li><li>at least two valid values are required</li><li>the critical <code>t</code> value is computed according to the selected test direction</li></ul>\n<h3 id=\"output\">Output</h3>\n<p>The spill output contains:</p>\n<ul><li><code>x̄</code></li><li><code>μ₀</code></li><li><code>sₓ</code></li><li><code>n</code></li><li><code>α</code></li><li><code>t</code></li><li><code>df</code></li><li>critical <code>t</code></li><li><code>p</code></li></ul>\n<h2 id=\"proptest1s\"><code>PROP.TEST.1S</code></h2>\n<p>Performs a one-sample z-test for a proportion.</p>\n<h3 id=\"syntax\">Syntax</h3>\n<pre><code class=\"language-excel\">=PROP.TEST.1S(values; pi_0; [direction]; [alpha]; [has_header])</code></pre>\n<h3 id=\"arguments\">Arguments</h3>\n<ul><li><code>values</code>: binary sample in <code>0/1</code> or <code>FALSE/TRUE</code> form</li><li><code>pi_0</code>: hypothetical population proportion</li><li><code>direction</code>: optional direction code; default is <code>0</code></li><li><code>alpha</code>: significance level</li><li><code>has_header</code>: optional header mode code; default is <code>0</code></li></ul>\n<h3 id=\"direction-codes\"><code>direction</code> Codes</h3>\n<table><thead><tr><th>Code</th><th>Meaning</th></tr></thead><tbody><tr><td><code>0</code></td><td>two-sided test</td></tr><tr><td><code>1</code></td><td>left-sided test</td></tr><tr><td><code>2</code></td><td>right-sided test</td></tr></tbody></table>\n<h3 id=\"has-header-codes\"><code>has_header</code> Codes</h3>\n<table><thead><tr><th>Code</th><th>Meaning</th></tr></thead><tbody><tr><td><code>0</code></td><td>autodetect header</td></tr><tr><td><code>1</code></td><td>first cell is a header</td></tr><tr><td><code>2</code></td><td>input has no header</td></tr></tbody></table>\n<h3 id=\"notes\">Notes</h3>\n<ul><li><code>pi_0</code> must lie strictly between <code>0</code> and <code>1</code></li><li>the function accepts only binary data <code>0/1</code> or <code>FALSE/TRUE</code></li><li>at least one valid observation is required</li></ul>\n<h3 id=\"output\">Output</h3>\n<p>The spill output contains:</p>\n<ul><li><code>p̂</code></li><li><code>π₀</code></li><li><code>x</code></li><li><code>n</code></li><li><code>α</code></li><li><code>z</code></li><li>critical <code>z</code></li><li><code>p</code></li></ul>\n<h2 id=\"wilcoxonpaired\"><code>WILCOXON.PAIRED</code></h2>\n<p>Performs the Wilcoxon signed-rank test for dependent pairs.</p>\n<h3 id=\"syntax\">Syntax</h3>\n<pre><code class=\"language-excel\">=WILCOXON.PAIRED(x; y; [has_header]; [alpha]; [direction])</code></pre>\n<h3 id=\"arguments\">Arguments</h3>\n<ul><li><code>x</code>: first measurement</li><li><code>y</code>: second measurement</li><li><code>has_header</code>: optional header mode code; default is <code>0</code></li><li><code>alpha</code>: significance level</li><li><code>direction</code>: optional direction code; default is <code>0</code></li></ul>\n<h3 id=\"direction-codes\"><code>direction</code> Codes</h3>\n<table><thead><tr><th>Code</th><th>Meaning</th></tr></thead><tbody><tr><td><code>0</code></td><td>two-sided test</td></tr><tr><td><code>1</code></td><td>left-sided test</td></tr><tr><td><code>2</code></td><td>right-sided test</td></tr></tbody></table>\n<h3 id=\"has-header-codes\"><code>has_header</code> Codes</h3>\n<table><thead><tr><th>Code</th><th>Meaning</th></tr></thead><tbody><tr><td><code>0</code></td><td>autodetect header</td></tr><tr><td><code>1</code></td><td>first cell is a header</td></tr><tr><td><code>2</code></td><td>input has no header</td></tr></tbody></table>\n<h3 id=\"notes\">Notes</h3>\n<ul><li>the function works with paired differences <code>x - y</code></li><li>incomplete pairs are skipped pairwise</li><li>zero differences are excluded from the test</li><li>if too few values remain after removing zero differences, the function returns a count error</li></ul>\n<h3 id=\"output\">Output</h3>\n<p>The spill output contains:</p>\n<ul><li><code>n</code></li><li><code>med(d)</code></li><li><code>α</code></li><li><code>W+</code></li><li><code>W-</code></li><li><code>W</code></li><li><code>z</code></li><li>critical <code>z</code></li><li><code>p</code></li><li>effect size <code>r</code></li></ul>\n<h3 id=\"example\">Example</h3>\n<pre><code class=\"language-excel\">=WILCOXON.PAIRED(B2:B20;C2:C20)\n=WILCOXON.PAIRED(B2:B20;C2:C20;1;0,05;0)</code></pre>"
    },
    {
      "lang": "en",
      "slug": "functions/percentiles",
      "section": "Functions",
      "title": "Percentiles",
      "summary": "Percentiles PERCENTILE.INC.IFS Computes an inclusive percentile with SUMIFS style filtering. Syntax Arguments values: numeric data quantile: a value from 0 to 1 criteria range n: r",
      "sourcePath": "docs/en/functions/percentiles.md",
      "html": "<h1 id=\"percentiles\">Percentiles</h1>\n<h2 id=\"percentileincifs\"><code>PERCENTILE.INC.IFS</code></h2>\n<p>Computes an inclusive percentile with <code>SUMIFS</code>-style filtering.</p>\n<h3 id=\"syntax\">Syntax</h3>\n<pre><code class=\"language-excel\">=PERCENTILE.INC.IFS(values; quantile; [criteria_range_1; criteria_1; ...])</code></pre>\n<h3 id=\"arguments\">Arguments</h3>\n<ul><li><code>values</code>: numeric data</li><li><code>quantile</code>: a value from <code>0</code> to <code>1</code></li><li><code>criteria_range_n</code>: range used for filtering</li><li><code>criteria_n</code>: exact match, relational expression, or wildcard expression</li></ul>\n<h3 id=\"notes\">Notes</h3>\n<ul><li>filter arguments must come in even <code>range + criteria</code> pairs</li><li>if filtering leaves no values, the function returns <code>#N/A</code></li><li>the percentile is computed using the inclusive method, matching <code>PERCENTILE.INC</code></li></ul>\n<h3 id=\"output\">Output</h3>\n<p>A scalar percentile.</p>\n<h3 id=\"example\">Example</h3>\n<pre><code class=\"language-excel\">=PERCENTILE.INC.IFS(A2:A100;0,75;B2:B100;&quot;A&quot;;A2:A100;&quot;&gt;10&quot;)</code></pre>\n<h2 id=\"percentileexcifs\"><code>PERCENTILE.EXC.IFS</code></h2>\n<p>Computes an exclusive percentile with <code>SUMIFS</code>-style filtering.</p>\n<h3 id=\"syntax\">Syntax</h3>\n<pre><code class=\"language-excel\">=PERCENTILE.EXC.IFS(values; quantile; [criteria_range_1; criteria_1; ...])</code></pre>\n<h3 id=\"arguments\">Arguments</h3>\n<p>Same as <code>PERCENTILE.INC.IFS</code>, but <code>quantile</code> must be strictly between <code>0</code> and <code>1</code>.</p>\n<h3 id=\"notes\">Notes</h3>\n<ul><li>filter arguments must come in even <code>range + criteria</code> pairs</li><li>if filtering leaves no values, the function returns <code>#N/A</code></li><li>if the exclusive percentile formula falls outside its valid domain for the given quantile, the function returns a numeric error</li></ul>\n<h3 id=\"output\">Output</h3>\n<p>A scalar percentile.</p>"
    },
    {
      "lang": "en",
      "slug": "functions/pivot",
      "section": "Functions",
      "title": "PIVOT",
      "summary": "PIVOT PIVOT. The PIVOT. family builds a statistical pivot from row and column categories, and each function computes exactly one metric. Syntax Arguments rows: one or more category",
      "sourcePath": "docs/en/functions/pivot.md",
      "html": "<h1 id=\"pivot\">PIVOT</h1>\n<h2 id=\"pivot\"><code>PIVOT.*</code></h2>\n<p>The <code>PIVOT.*</code> family builds a statistical pivot from row and column categories, and each function computes exactly one metric.</p>\n<h3 id=\"syntax\">Syntax</h3>\n<pre><code class=\"language-excel\">=PIVOT.SUM(rows; columns; values)\n=PIVOT.AVERAGE(rows; columns; values)\n=PIVOT.PERCENTILE(rows; columns; values; quantile)\n=PIVOT.CONF.T(rows; columns; values; alpha; [direction])</code></pre>\n<h3 id=\"arguments\">Arguments</h3>\n<ul><li><code>rows</code>: one or more category columns for rows, always including headers</li><li><code>columns</code>: optionally one or more category columns for columns, always including headers; may be blank</li><li><code>values</code>: one value column including a header</li><li><code>quantile</code>: requested percentile in <code>(0,1)</code></li><li><code>alpha</code>: alpha level in <code>(0,1)</code></li><li><code>direction</code>: optional direction for <code>CONF.*</code>; <code>0 = two-sided</code>, <code>-1 = left-sided</code>, <code>1 = right-sided</code></li></ul>\n<h3 id=\"available-functions\">Available functions</h3>\n<table><thead><tr><th>Function</th><th>Description</th><th>Detail</th></tr></thead><tbody><tr><td><code>PIVOT.COUNT(rows; columns; values)</code></td><td>Count of non-empty values</td><td>no detail</td></tr><tr><td><code>PIVOT.SUM(rows; columns; values)</code></td><td>Sum</td><td>no detail</td></tr><tr><td><code>PIVOT.AVERAGE(rows; columns; values)</code></td><td>Arithmetic mean</td><td>no detail</td></tr><tr><td><code>PIVOT.MIN(rows; columns; values)</code></td><td>Minimum</td><td>no detail</td></tr><tr><td><code>PIVOT.MAX(rows; columns; values)</code></td><td>Maximum</td><td>no detail</td></tr><tr><td><code>PIVOT.MEDIAN(rows; columns; values)</code></td><td>Median</td><td>no detail</td></tr><tr><td><code>PIVOT.PERCENTILE(rows; columns; values; quantile)</code></td><td>Percentile</td><td><code>0 &lt; quantile &lt; 1</code></td></tr><tr><td><code>PIVOT.STDEV.S(rows; columns; values)</code></td><td>Sample standard deviation</td><td>no detail</td></tr><tr><td><code>PIVOT.STDEV.P(rows; columns; values)</code></td><td>Population standard deviation</td><td>no detail</td></tr><tr><td><code>PIVOT.VAR.S(rows; columns; values)</code></td><td>Sample variance</td><td>no detail</td></tr><tr><td><code>PIVOT.VAR.P(rows; columns; values)</code></td><td>Population variance</td><td>no detail</td></tr><tr><td><code>PIVOT.VARCOEF.S(rows; columns; values)</code></td><td>Sample coefficient of variation</td><td>no detail</td></tr><tr><td><code>PIVOT.VARCOEF.P(rows; columns; values)</code></td><td>Population coefficient of variation</td><td>no detail</td></tr><tr><td><code>PIVOT.CONF.T(rows; columns; values; alpha; [direction])</code></td><td>Half-width of the confidence interval based on the t distribution</td><td><code>0 &lt; alpha &lt; 1</code>; <code>direction</code>: <code>0</code>, <code>-1</code>, <code>1</code></td></tr><tr><td><code>PIVOT.CONF.NORM(rows; columns; values; alpha; [direction])</code></td><td>Half-width of the confidence interval based on the normal approximation</td><td><code>0 &lt; alpha &lt; 1</code>; <code>direction</code>: <code>0</code>, <code>-1</code>, <code>1</code></td></tr><tr><td><code>PIVOT.MAD(rows; columns; values)</code></td><td>Median absolute deviation from the median</td><td>no detail</td></tr><tr><td><code>PIVOT.IQR(rows; columns; values)</code></td><td>Interquartile range</td><td>no detail</td></tr></tbody></table>\n<h3 id=\"validation\">Validation</h3>\n<ul><li><code>rows</code>, <code>columns</code>, and <code>values</code> are expected to include a header in the first row</li><li><code>count</code> can work with non-numeric values and counts non-empty cells</li><li>all other functions require quantitative data</li><li><code>varcoef.s</code> and <code>varcoef.p</code> require a non-zero mean</li><li><code>stdev.s</code>, <code>var.s</code>, <code>conf.t</code>, and <code>conf.norm</code> require at least two valid numeric observations</li><li>incompatible data and function combinations return a value error</li><li><code>conf.*</code> only accept <code>direction</code> from <code>{-1, 0, 1}</code></li></ul>\n<h3 id=\"notes\">Notes</h3>\n<ul><li><code>columns</code> may be blank; in that case a one-dimensional pivot is created with a single <code>Total</code> column block</li><li>row and column categories are sorted alphabetically in the output</li><li>the output always includes a <code>TOTAL</code> summary row and a <code>TOTAL</code> summary column</li><li>the bottom-right cell contains the overall aggregate for the selected metric across all data</li><li>the output no longer includes a dedicated metric-label row</li><li>rows for which every resulting value is blank are omitted from the output</li><li>columns for which every resulting value is blank are omitted from the output, except for the <code>TOTAL</code> summary column</li><li><code>conf.t</code> and <code>conf.norm</code> return the half-width of the interval, not both bounds separately</li><li>with <code>direction = 0</code>, the critical value <code>1 - alpha/2</code> is used</li><li>with <code>direction = -1</code> or <code>1</code>, the critical value <code>1 - alpha</code> is used</li></ul>\n<h3 id=\"output\">Output</h3>\n<p>Multi-row spill output in a pivot-like layout:</p>\n<ul><li>row category columns on the left</li><li>column category levels at the top</li><li>the last header row above numeric columns now only contains row-variable names</li><li>the last column is always the <code>TOTAL</code> summary column</li><li>the last row is always the <code>TOTAL</code> summary row</li></ul>\n<h3 id=\"examples\">Examples</h3>\n<pre><code class=\"language-excel\">=PIVOT.SUM(E:E;F:F;G:G)\n=PIVOT.AVERAGE(E:E;F:F;G:G)\n=PIVOT.PERCENTILE(E:E;F:F;G:G;0.9)\n=PIVOT.CONF.T(E:E;F:F;G:G;0.05)\n=PIVOT.CONF.T(E:E;F:F;G:G;0.05;1)</code></pre>"
    },
    {
      "lang": "en",
      "slug": "functions/two-sample-tests",
      "section": "Functions",
      "title": "Two-Sample Tests",
      "summary": "Two Sample Tests WELCH.TEST.2S.G Performs Welch's two sample t test for two independent groups. Syntax Arguments categories: labels defining exactly two groups values: numeric obse",
      "sourcePath": "docs/en/functions/two-sample-tests.md",
      "html": "<h1 id=\"two-sample-tests\">Two-Sample Tests</h1>\n<h2 id=\"welchtest2sg\"><code>WELCH.TEST.2S.G</code></h2>\n<p>Performs Welch&#39;s two-sample t-test for two independent groups.</p>\n<h3 id=\"syntax\">Syntax</h3>\n<pre><code class=\"language-excel\">=WELCH.TEST.2S.G(categories; values; [has_header]; [alpha]; [direction])</code></pre>\n<h3 id=\"arguments\">Arguments</h3>\n<ul><li><code>categories</code>: labels defining exactly two groups</li><li><code>values</code>: numeric observations</li><li><code>has_header</code>: optional header mode code; default is <code>0</code></li><li><code>alpha</code>: significance level</li><li><code>direction</code>: optional direction code; default is <code>0</code></li></ul>\n<h3 id=\"direction-codes\"><code>direction</code> Codes</h3>\n<table><thead><tr><th>Code</th><th>Meaning</th></tr></thead><tbody><tr><td><code>0</code></td><td>two-sided test</td></tr><tr><td><code>1</code></td><td>left-sided test</td></tr><tr><td><code>2</code></td><td>right-sided test</td></tr></tbody></table>\n<h3 id=\"has-header-codes\"><code>has_header</code> Codes</h3>\n<table><thead><tr><th>Code</th><th>Meaning</th></tr></thead><tbody><tr><td><code>0</code></td><td>autodetect header</td></tr><tr><td><code>1</code></td><td>first row is a header</td></tr><tr><td><code>2</code></td><td>input has no header</td></tr></tbody></table>\n<h3 id=\"notes\">Notes</h3>\n<ul><li>the function requires exactly two groups</li><li>each group must contain at least two values</li><li>groups are internally sorted by label, which affects the sign of the difference and the <code>t</code> statistic</li><li>descriptive statistics are returned in a compact table with groups as rows</li></ul>\n<h3 id=\"output\">Output</h3>\n<p>The spill output contains:</p>\n<ul><li>a descriptive statistics table by group</li><li><code>α</code></li><li><code>t</code></li><li>Welch-Satterthwaite <code>df</code></li><li>critical <code>t</code></li><li><code>p</code></li><li>Cohen&#39;s <code>d</code></li><li>effect size <code>r</code></li></ul>\n<h3 id=\"example\">Example</h3>\n<pre><code class=\"language-excel\">=WELCH.TEST.2S.G(A2:A40;B2:B40;1;0,05;0)</code></pre>\n<h2 id=\"mannwhitneyg\"><code>MANN.WHITNEY.G</code></h2>\n<p>Performs the nonparametric Mann-Whitney test for two independent groups.</p>\n<h3 id=\"syntax\">Syntax</h3>\n<pre><code class=\"language-excel\">=MANN.WHITNEY.G(categories; values; [has_header]; [alpha]; [direction])</code></pre>\n<h3 id=\"arguments\">Arguments</h3>\n<ul><li><code>categories</code>: labels of exactly two groups</li><li><code>values</code>: numeric observations</li><li><code>has_header</code>: optional header mode code; default is <code>0</code></li><li><code>alpha</code>: significance level</li><li><code>direction</code>: optional direction code; default is <code>0</code></li></ul>\n<h3 id=\"direction-codes\"><code>direction</code> Codes</h3>\n<table><thead><tr><th>Code</th><th>Meaning</th></tr></thead><tbody><tr><td><code>0</code></td><td>two-sided test</td></tr><tr><td><code>1</code></td><td>left-sided test</td></tr><tr><td><code>2</code></td><td>right-sided test</td></tr></tbody></table>\n<h3 id=\"has-header-codes\"><code>has_header</code> Codes</h3>\n<table><thead><tr><th>Code</th><th>Meaning</th></tr></thead><tbody><tr><td><code>0</code></td><td>autodetect header</td></tr><tr><td><code>1</code></td><td>first row is a header</td></tr><tr><td><code>2</code></td><td>input has no header</td></tr></tbody></table>\n<h3 id=\"notes\">Notes</h3>\n<ul><li>the function requires exactly two groups</li><li>each group must contain at least one value</li><li>mid-ranks are used for ties</li><li>if the variance term becomes zero after tie correction, the function returns a numeric error</li></ul>\n<h3 id=\"output\">Output</h3>\n<p>The spill output contains:</p>\n<ul><li>a descriptive statistics table by group</li><li><code>U</code></li><li><code>U₁</code></li><li><code>U₂</code></li><li><code>z</code></li><li>critical <code>z</code></li><li><code>p</code></li><li>effect size <code>r</code></li></ul>\n<h3 id=\"example\">Example</h3>\n<pre><code class=\"language-excel\">=MANN.WHITNEY.G(A2:A20;B2:B20)\n=MANN.WHITNEY.G(A2:A20;B2:B20;1;0,05;0)</code></pre>"
    },
    {
      "lang": "en",
      "slug": "functions/variation-coefficients",
      "section": "Functions",
      "title": "Coefficients Of Variation",
      "summary": "Coefficients Of Variation VARCOEF Computes the population coefficient of variation. Syntax Notes blank cells are ignored at least one valid value is required if the mean equals zer",
      "sourcePath": "docs/en/functions/variation-coefficients.md",
      "html": "<h1 id=\"coefficients-of-variation\">Coefficients Of Variation</h1>\n<h2 id=\"varcoef\"><code>VARCOEF</code></h2>\n<p>Computes the population coefficient of variation.</p>\n<h3 id=\"syntax\">Syntax</h3>\n<pre><code class=\"language-excel\">=VARCOEF(values)</code></pre>\n<h3 id=\"notes\">Notes</h3>\n<ul><li>blank cells are ignored</li><li>at least one valid value is required</li><li>if the mean equals zero, the function returns a divide-by-zero error</li></ul>\n<h3 id=\"output\">Output</h3>\n<p>A scalar <code>σ / μ</code>.</p>\n<h2 id=\"varcoefs\"><code>VARCOEF.S</code></h2>\n<p>Computes the sample coefficient of variation.</p>\n<h3 id=\"syntax\">Syntax</h3>\n<pre><code class=\"language-excel\">=VARCOEF.S(values)</code></pre>\n<h3 id=\"notes\">Notes</h3>\n<ul><li>blank cells are ignored</li><li>at least two valid values are required</li><li>if the mean equals zero, the function returns a divide-by-zero error</li></ul>\n<h3 id=\"output\">Output</h3>\n<p>A scalar <code>sₓ / x̄</code>.</p>\n<h2 id=\"varcoefw\"><code>VARCOEF.W</code></h2>\n<p>Computes the weighted population coefficient of variation.</p>\n<h3 id=\"syntax\">Syntax</h3>\n<pre><code class=\"language-excel\">=VARCOEF.W(values; weights)</code></pre>\n<h3 id=\"notes\">Notes</h3>\n<ul><li><code>values</code> and <code>weights</code> must have the same length</li><li>weights must be non-negative</li><li>blank weight cells are treated as <code>0</code></li><li>if the weighted mean equals zero, the function returns a divide-by-zero error</li></ul>\n<h3 id=\"output\">Output</h3>\n<p>A scalar <code>σ_w / x̄_w</code>.</p>\n<h2 id=\"varcoefsw\"><code>VARCOEF.S.W</code></h2>\n<p>Computes the weighted sample coefficient of variation.</p>\n<h3 id=\"syntax\">Syntax</h3>\n<pre><code class=\"language-excel\">=VARCOEF.S.W(values; weights)</code></pre>\n<h3 id=\"notes\">Notes</h3>\n<ul><li><code>values</code> and <code>weights</code> must have the same length</li><li>weights must be non-negative</li><li>blank weight cells are treated as <code>0</code></li><li>if <code>Σw &lt;= 1</code>, the function returns <code>#COUNT!</code></li><li>if the weighted mean equals zero, the function returns a divide-by-zero error</li></ul>\n<h3 id=\"output\">Output</h3>\n<p>A scalar <code>sₓ,w / x̄_w</code>.</p>"
    },
    {
      "lang": "en",
      "slug": "functions/weighted-means",
      "section": "Functions",
      "title": "Weighted Means",
      "summary": "Weighted Means AVERAGE.W Computes the weighted arithmetic mean. Syntax Arguments values: numeric observations weights: non negative weights; blank weight cells are treated as 0 Not",
      "sourcePath": "docs/en/functions/weighted-means.md",
      "html": "<h1 id=\"weighted-means\">Weighted Means</h1>\n<h2 id=\"averagew\"><code>AVERAGE.W</code></h2>\n<p>Computes the weighted arithmetic mean.</p>\n<h3 id=\"syntax\">Syntax</h3>\n<pre><code class=\"language-excel\">=AVERAGE.W(values; weights)</code></pre>\n<h3 id=\"arguments\">Arguments</h3>\n<ul><li><code>values</code>: numeric observations</li><li><code>weights</code>: non-negative weights; blank weight cells are treated as <code>0</code></li></ul>\n<h3 id=\"notes\">Notes</h3>\n<ul><li><code>values</code> and <code>weights</code> must have the same length</li><li>blank cells in <code>values</code> are skipped together with their matching weight</li><li>if all weights are zero or any weight is negative, the function returns a numeric error</li></ul>\n<h3 id=\"output\">Output</h3>\n<p>A scalar weighted arithmetic mean.</p>\n<h3 id=\"example\">Example</h3>\n<pre><code class=\"language-excel\">=AVERAGE.W(A2:A10;B2:B10)</code></pre>\n<h2 id=\"harmeanw\"><code>HARMEAN.W</code></h2>\n<p>Computes the weighted harmonic mean.</p>\n<h3 id=\"syntax\">Syntax</h3>\n<pre><code class=\"language-excel\">=HARMEAN.W(values; weights)</code></pre>\n<h3 id=\"arguments\">Arguments</h3>\n<ul><li><code>values</code>: positive numeric observations</li><li><code>weights</code>: non-negative weights; blank weight cells are treated as <code>0</code></li></ul>\n<h3 id=\"notes\">Notes</h3>\n<ul><li><code>values</code> and <code>weights</code> must have the same length</li><li>values with zero weight do not affect the result</li><li>if any value with positive weight is <code>&lt;= 0</code>, the function returns a numeric error</li></ul>\n<h3 id=\"output\">Output</h3>\n<p>A scalar weighted harmonic mean.</p>\n<h2 id=\"geomeanw\"><code>GEOMEAN.W</code></h2>\n<p>Computes the weighted geometric mean.</p>\n<h3 id=\"syntax\">Syntax</h3>\n<pre><code class=\"language-excel\">=GEOMEAN.W(values; weights)</code></pre>\n<h3 id=\"arguments\">Arguments</h3>\n<ul><li><code>values</code>: positive numeric observations</li><li><code>weights</code>: non-negative weights; blank weight cells are treated as <code>0</code></li></ul>\n<h3 id=\"notes\">Notes</h3>\n<ul><li><code>values</code> and <code>weights</code> must have the same length</li><li>values with zero weight do not affect the result</li><li>if any value with positive weight is <code>&lt;= 0</code>, the function returns a numeric error</li></ul>\n<h3 id=\"output\">Output</h3>\n<p>A scalar weighted geometric mean.</p>"
    },
    {
      "lang": "en",
      "slug": "functions/weighted-variance",
      "section": "Functions",
      "title": "Weighted Variance And Standard Deviation",
      "summary": "Weighted Variance And Standard Deviation Common behavior for all four functions: values and weights must have the same length weights must be non negative blank weight cells are tr",
      "sourcePath": "docs/en/functions/weighted-variance.md",
      "html": "<h1 id=\"weighted-variance-and-standard-deviation\">Weighted Variance And Standard Deviation</h1>\n<p>Common behavior for all four functions:</p>\n<ul><li><code>values</code> and <code>weights</code> must have the same length</li><li>weights must be non-negative</li><li>blank weight cells are treated as <code>0</code></li><li>if the sum of weights is zero, the function returns a numeric error</li></ul>\n<h2 id=\"varpw\"><code>VAR.P.W</code></h2>\n<p>Computes the weighted population variance.</p>\n<h3 id=\"syntax\">Syntax</h3>\n<pre><code class=\"language-excel\">=VAR.P.W(values; weights)</code></pre>\n<h3 id=\"output\">Output</h3>\n<p>A scalar weighted population variance.</p>\n<h3 id=\"notes\">Notes</h3>\n<ul><li>uses the denominator <code>Σw</code></li></ul>\n<h2 id=\"varsw\"><code>VAR.S.W</code></h2>\n<p>Computes the weighted sample variance.</p>\n<h3 id=\"syntax\">Syntax</h3>\n<pre><code class=\"language-excel\">=VAR.S.W(values; weights)</code></pre>\n<h3 id=\"output\">Output</h3>\n<p>A scalar weighted sample variance.</p>\n<h3 id=\"notes\">Notes</h3>\n<ul><li>uses the denominator <code>Σw - 1</code></li><li>if <code>Σw &lt;= 1</code>, the function returns <code>#COUNT!</code></li></ul>\n<h2 id=\"stdevpw\"><code>STDEV.P.W</code></h2>\n<p>Computes the weighted population standard deviation.</p>\n<h3 id=\"syntax\">Syntax</h3>\n<pre><code class=\"language-excel\">=STDEV.P.W(values; weights)</code></pre>\n<h3 id=\"output\">Output</h3>\n<p>A scalar weighted population standard deviation <code>σ</code>.</p>\n<h3 id=\"notes\">Notes</h3>\n<ul><li>defined as the square root of <code>VAR.P.W</code></li></ul>\n<h2 id=\"stdevsw\"><code>STDEV.S.W</code></h2>\n<p>Computes the weighted sample standard deviation.</p>\n<h3 id=\"syntax\">Syntax</h3>\n<pre><code class=\"language-excel\">=STDEV.S.W(values; weights)</code></pre>\n<h3 id=\"output\">Output</h3>\n<p>A scalar weighted sample standard deviation <code>sₓ</code>.</p>\n<h3 id=\"notes\">Notes</h3>\n<ul><li>defined as the square root of <code>VAR.S.W</code></li></ul>"
    },
    {
      "lang": "en",
      "slug": "overview",
      "section": "Overview",
      "title": "XLStatUDF Documentation",
      "summary": "XLStatUDF Documentation This branch contains the English user facing documentation for XLStatUDF. Main Build Use this file for regular Excel testing: artifacts/main/publish/XLStatU",
      "sourcePath": "docs/en/README.md",
      "html": "<h1 id=\"xlstatudf-documentation\">XLStatUDF Documentation</h1>\n<p>This branch contains the English user-facing documentation for XLStatUDF.</p>\n<h2 id=\"main-build\">Main Build</h2>\n<p>Use this file for regular Excel testing:</p>\n<ul><li><a href=\"#\"><code>artifacts/main/publish/XLStatUDF-AddIn64-packed.xll</code></a></li></ul>\n<h2 id=\"contents\">Contents</h2>\n<ul><li><a href=\"/en/docs/framework\">Framework Overview</a></li><li><a href=\"/en/docs/functions/index\">Functions Index</a></li></ul>"
    }
  ]
};

export const functionIndexByLanguage: Record<'cs' | 'en', FunctionIndexEntry[]> = {
  "cs": [
    {
      "lang": "cs",
      "name": "GENERATE.NORM",
      "summary": "Generuje jednu náhodnou hodnotu z normálního rozdělení s volitelnou perturbací.",
      "href": "/cs/docs/distributions#generatenorm"
    },
    {
      "lang": "cs",
      "name": "GENERATE.INT",
      "summary": "Generuje jedno náhodné celé číslo ze zadaného intervalu s volitelnou perturbací.",
      "href": "/cs/docs/distributions#generateint"
    },
    {
      "lang": "cs",
      "name": "FILL",
      "summary": "Opakuje hodnoty nebo textové vzorce do jednosloupcového spill výstupu.",
      "href": "/cs/docs/distributions#fill"
    },
    {
      "lang": "cs",
      "name": "FILL.RANDOM",
      "summary": "Sestavuje řadu jako `FILL` a následně ji náhodně promíchá.",
      "href": "/cs/docs/distributions#fillrandom"
    },
    {
      "lang": "cs",
      "name": "AVERAGE.W",
      "summary": "Počítá vážený aritmetický průměr.",
      "href": "/cs/docs/weighted-means#averagew"
    },
    {
      "lang": "cs",
      "name": "HARMEAN.W",
      "summary": "Počítá vážený harmonický průměr.",
      "href": "/cs/docs/weighted-means#harmeanw"
    },
    {
      "lang": "cs",
      "name": "GEOMEAN.W",
      "summary": "Počítá vážený geometrický průměr.",
      "href": "/cs/docs/weighted-means#geomeanw"
    },
    {
      "lang": "cs",
      "name": "VAR.P.W",
      "summary": "Počítá vážený populační rozptyl.",
      "href": "/cs/docs/weighted-variance#varpw"
    },
    {
      "lang": "cs",
      "name": "VAR.S.W",
      "summary": "Počítá vážený výběrový rozptyl.",
      "href": "/cs/docs/weighted-variance#varsw"
    },
    {
      "lang": "cs",
      "name": "STDEV.P.W",
      "summary": "Počítá váženou populační směrodatnou odchylku.",
      "href": "/cs/docs/weighted-variance#stdevpw"
    },
    {
      "lang": "cs",
      "name": "STDEV.S.W",
      "summary": "Počítá váženou výběrovou směrodatnou odchylku.",
      "href": "/cs/docs/weighted-variance#stdevsw"
    },
    {
      "lang": "cs",
      "name": "VARCOEF",
      "summary": "Počítá populační variační koeficient.",
      "href": "/cs/docs/variation-coefficients#varcoef"
    },
    {
      "lang": "cs",
      "name": "VARCOEF.S",
      "summary": "Počítá výběrový variační koeficient.",
      "href": "/cs/docs/variation-coefficients#varcoefs"
    },
    {
      "lang": "cs",
      "name": "VARCOEF.W",
      "summary": "Počítá vážený populační variační koeficient.",
      "href": "/cs/docs/variation-coefficients#varcoefw"
    },
    {
      "lang": "cs",
      "name": "VARCOEF.S.W",
      "summary": "Počítá vážený výběrový variační koeficient.",
      "href": "/cs/docs/variation-coefficients#varcoefsw"
    },
    {
      "lang": "cs",
      "name": "PERCENTILE.INC.IFS",
      "summary": "Počítá inkluzivní percentil po aplikaci filtrů.",
      "href": "/cs/docs/percentiles#percentileincifs"
    },
    {
      "lang": "cs",
      "name": "PERCENTILE.EXC.IFS",
      "summary": "Počítá exkluzivní percentil po aplikaci filtrů.",
      "href": "/cs/docs/percentiles#percentileexcifs"
    },
    {
      "lang": "cs",
      "name": "PIVOT.*",
      "summary": "Sestavuje statistický pivot a v každé funkci počítá právě jeden zvolený ukazatel.",
      "href": "/cs/docs/pivot#pivot"
    },
    {
      "lang": "cs",
      "name": "NORM.DIST.RANGE",
      "summary": "Počítá pravděpodobnost intervalu normálního rozdělení.",
      "href": "/cs/docs/distributions#normdistrange"
    },
    {
      "lang": "cs",
      "name": "SHAPIRO.WILK",
      "summary": "Provádí Shapiro-Wilkův test normality.",
      "href": "/cs/docs/normality#shapirowilk"
    },
    {
      "lang": "cs",
      "name": "KOLMOGOROV.SMIRNOV",
      "summary": "Provádí Kolmogorov-Smirnovův test dobré shody pro zvolené rozdělení.",
      "href": "/cs/docs/normality#kolmogorovsmirnov"
    },
    {
      "lang": "cs",
      "name": "T.TEST.1S",
      "summary": "Provádí jednovýběrový t-test vůči zadané hypotetické střední hodnotě.",
      "href": "/cs/docs/one-sample-tests#ttest1s"
    },
    {
      "lang": "cs",
      "name": "PROP.TEST.1S",
      "summary": "Provádí jednovýběrový test podílu.",
      "href": "/cs/docs/one-sample-tests#proptest1s"
    },
    {
      "lang": "cs",
      "name": "WILCOXON.PAIRED",
      "summary": "Provádí Wilcoxonův párový znaménkový test pro závislá měření.",
      "href": "/cs/docs/one-sample-tests#wilcoxonpaired"
    },
    {
      "lang": "cs",
      "name": "WELCH.TEST.2S.G",
      "summary": "Provádí Welchův dvouvýběrový t-test pro dvě nezávislé skupiny.",
      "href": "/cs/docs/two-sample-tests#welchtest2sg"
    },
    {
      "lang": "cs",
      "name": "MANN.WHITNEY.G",
      "summary": "Provádí Mann-Whitneyho test pro dvě nezávislé skupiny.",
      "href": "/cs/docs/two-sample-tests#mannwhitneyg"
    },
    {
      "lang": "cs",
      "name": "CHISQ.GOF",
      "summary": "Provádí chí-kvadrát test dobré shody.",
      "href": "/cs/docs/goodness-of-fit#chisqgof"
    },
    {
      "lang": "cs",
      "name": "ANOVA.G",
      "summary": "Provádí jednofaktorovou ANOVA nad groupovanými daty.",
      "href": "/cs/docs/anova#anovag"
    },
    {
      "lang": "cs",
      "name": "ANOVA.RM",
      "summary": "Provádí ANOVA s opakovaným měřením nad sloupci.",
      "href": "/cs/docs/anova#anovarm"
    },
    {
      "lang": "cs",
      "name": "ANCOVA.G",
      "summary": "Provádí ANCOVA s jedním faktorem a jednou nebo více kovariátami.",
      "href": "/cs/docs/ancova#ancovag"
    },
    {
      "lang": "cs",
      "name": "CONTINGENCY.T",
      "summary": "Analyzuje kontingenční tabulku zadanou přímo jako matici četností.",
      "href": "/cs/docs/contingency#contingencyt"
    },
    {
      "lang": "cs",
      "name": "CONTINGENCY.G",
      "summary": "Analyzuje kontingenční tabulku sestavenou z groupovaných sloupců.",
      "href": "/cs/docs/contingency#contingencyg"
    },
    {
      "lang": "cs",
      "name": "CORREL.SPEARMAN",
      "summary": "Počítá Spearmanův korelační koeficient a test jeho významnosti.",
      "href": "/cs/docs/correlation#correlspearman"
    },
    {
      "lang": "cs",
      "name": "CORREL.MATRIX",
      "summary": "Sestavuje korelační matici včetně p-hodnot a značek signifikance.",
      "href": "/cs/docs/correlation#correlmatrix"
    }
  ],
  "en": [
    {
      "lang": "en",
      "name": "GENERATE.NORM",
      "summary": "Generates a single random value from a normal distribution with optional perturbation.",
      "href": "/en/docs/distributions#generatenorm"
    },
    {
      "lang": "en",
      "name": "GENERATE.INT",
      "summary": "Generates a single random integer from a specified interval with optional perturbation.",
      "href": "/en/docs/distributions#generateint"
    },
    {
      "lang": "en",
      "name": "FILL",
      "summary": "Repeats values or text formulas into a single-column spill output.",
      "href": "/en/docs/distributions#fill"
    },
    {
      "lang": "en",
      "name": "FILL.RANDOM",
      "summary": "Builds a sequence like `FILL` and then shuffles it randomly.",
      "href": "/en/docs/distributions#fillrandom"
    },
    {
      "lang": "en",
      "name": "AVERAGE.W",
      "summary": "Computes the weighted arithmetic mean.",
      "href": "/en/docs/weighted-means#averagew"
    },
    {
      "lang": "en",
      "name": "HARMEAN.W",
      "summary": "Computes the weighted harmonic mean.",
      "href": "/en/docs/weighted-means#harmeanw"
    },
    {
      "lang": "en",
      "name": "GEOMEAN.W",
      "summary": "Computes the weighted geometric mean.",
      "href": "/en/docs/weighted-means#geomeanw"
    },
    {
      "lang": "en",
      "name": "VAR.P.W",
      "summary": "Computes the weighted population variance.",
      "href": "/en/docs/weighted-variance#varpw"
    },
    {
      "lang": "en",
      "name": "VAR.S.W",
      "summary": "Computes the weighted sample variance.",
      "href": "/en/docs/weighted-variance#varsw"
    },
    {
      "lang": "en",
      "name": "STDEV.P.W",
      "summary": "Computes the weighted population standard deviation.",
      "href": "/en/docs/weighted-variance#stdevpw"
    },
    {
      "lang": "en",
      "name": "STDEV.S.W",
      "summary": "Computes the weighted sample standard deviation.",
      "href": "/en/docs/weighted-variance#stdevsw"
    },
    {
      "lang": "en",
      "name": "VARCOEF",
      "summary": "Computes the population coefficient of variation.",
      "href": "/en/docs/variation-coefficients#varcoef"
    },
    {
      "lang": "en",
      "name": "VARCOEF.S",
      "summary": "Computes the sample coefficient of variation.",
      "href": "/en/docs/variation-coefficients#varcoefs"
    },
    {
      "lang": "en",
      "name": "VARCOEF.W",
      "summary": "Computes the weighted population coefficient of variation.",
      "href": "/en/docs/variation-coefficients#varcoefw"
    },
    {
      "lang": "en",
      "name": "VARCOEF.S.W",
      "summary": "Computes the weighted sample coefficient of variation.",
      "href": "/en/docs/variation-coefficients#varcoefsw"
    },
    {
      "lang": "en",
      "name": "PERCENTILE.INC.IFS",
      "summary": "Computes the inclusive percentile after applying filters.",
      "href": "/en/docs/percentiles#percentileincifs"
    },
    {
      "lang": "en",
      "name": "PERCENTILE.EXC.IFS",
      "summary": "Computes the exclusive percentile after applying filters.",
      "href": "/en/docs/percentiles#percentileexcifs"
    },
    {
      "lang": "en",
      "name": "PIVOT.*",
      "summary": "Builds a statistical pivot and computes exactly one selected metric per function.",
      "href": "/en/docs/pivot#pivot"
    },
    {
      "lang": "en",
      "name": "NORM.DIST.RANGE",
      "summary": "Computes the probability of an interval under the normal distribution.",
      "href": "/en/docs/distributions#normdistrange"
    },
    {
      "lang": "en",
      "name": "SHAPIRO.WILK",
      "summary": "Performs the Shapiro-Wilk normality test.",
      "href": "/en/docs/normality#shapirowilk"
    },
    {
      "lang": "en",
      "name": "KOLMOGOROV.SMIRNOV",
      "summary": "Performs the Kolmogorov-Smirnov goodness-of-fit test for the selected distribution.",
      "href": "/en/docs/normality#kolmogorovsmirnov"
    },
    {
      "lang": "en",
      "name": "T.TEST.1S",
      "summary": "Performs a one-sample t-test against a specified hypothetical mean.",
      "href": "/en/docs/one-sample-tests#ttest1s"
    },
    {
      "lang": "en",
      "name": "PROP.TEST.1S",
      "summary": "Performs a one-sample proportion test.",
      "href": "/en/docs/one-sample-tests#proptest1s"
    },
    {
      "lang": "en",
      "name": "WILCOXON.PAIRED",
      "summary": "Performs the Wilcoxon signed-rank test for dependent measurements.",
      "href": "/en/docs/one-sample-tests#wilcoxonpaired"
    },
    {
      "lang": "en",
      "name": "WELCH.TEST.2S.G",
      "summary": "Performs Welch's two-sample t-test for two independent groups.",
      "href": "/en/docs/two-sample-tests#welchtest2sg"
    },
    {
      "lang": "en",
      "name": "MANN.WHITNEY.G",
      "summary": "Performs the Mann-Whitney test for two independent groups.",
      "href": "/en/docs/two-sample-tests#mannwhitneyg"
    },
    {
      "lang": "en",
      "name": "CHISQ.GOF",
      "summary": "Performs the chi-square goodness-of-fit test.",
      "href": "/en/docs/goodness-of-fit#chisqgof"
    },
    {
      "lang": "en",
      "name": "ANOVA.G",
      "summary": "Performs one-way ANOVA on grouped data.",
      "href": "/en/docs/anova#anovag"
    },
    {
      "lang": "en",
      "name": "ANOVA.RM",
      "summary": "Performs repeated-measures ANOVA across columns.",
      "href": "/en/docs/anova#anovarm"
    },
    {
      "lang": "en",
      "name": "ANCOVA.G",
      "summary": "Performs ANCOVA with one factor and one or more covariates.",
      "href": "/en/docs/ancova#ancovag"
    },
    {
      "lang": "en",
      "name": "CONTINGENCY.T",
      "summary": "Analyzes a contingency table entered directly as a frequency matrix.",
      "href": "/en/docs/contingency#contingencyt"
    },
    {
      "lang": "en",
      "name": "CONTINGENCY.G",
      "summary": "Analyzes a contingency table built from grouped columns.",
      "href": "/en/docs/contingency#contingencyg"
    },
    {
      "lang": "en",
      "name": "CORREL.SPEARMAN",
      "summary": "Computes Spearman correlation and its significance test.",
      "href": "/en/docs/correlation#correlspearman"
    },
    {
      "lang": "en",
      "name": "CORREL.MATRIX",
      "summary": "Builds a correlation matrix including p-values and significance markers.",
      "href": "/en/docs/correlation#correlmatrix"
    }
  ]
};

export const installersByLanguage: Record<'cs' | 'en', GeneratedInstaller[]> = {
  "cs": [
    {
      "lang": "cs",
      "label": "Czech installer",
      "fileName": "XLStatUDF_CS_Setup.exe",
      "href": "/downloads/XLStatUDF_CS_Setup.exe",
      "exists": true
    }
  ],
  "en": [
    {
      "lang": "en",
      "label": "English installer",
      "fileName": "XLStatUDF_EN_Setup.exe",
      "href": "/downloads/XLStatUDF_EN_Setup.exe",
      "exists": true
    }
  ]
};
