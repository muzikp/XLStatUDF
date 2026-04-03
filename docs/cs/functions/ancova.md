# ANCOVA

## `ANCOVA.G`

Provádí analýzu kovariance nad groupovanými daty s jedním faktorem a jednou nebo více kovariátami.

### Syntaxe

```excel
=ANCOVA.G(faktor; zavisla_promenna; kovariaty; [post_hoc]; [alpha]; [ma_záhlaví])
```

### Argumenty

- `faktor`: kategorie faktoru
- `zavisla_promenna`: závislá proměnná
- `kovariaty`: jedna nebo více kovariát ve sloupcích
- `post_hoc`: volitelný kód post-hoc procedury; výchozí hodnota je `0`
- `alpha`: hladina významnosti
- `ma_záhlaví`: volitelný kód režimu záhlaví; výchozí hodnota je `0`

### Kódy `post_hoc`

| Kód | Název | Popis |
| --- | --- | --- |
| `0` | `none` | bez post-hoc porovnání |
| `1` | `tukey` | konzervativní fallback přes Bonferroni |
| `2` | `bonferroni` | párová porovnání adjusted means s Bonferroniho korekcí |
| `3` | `scheffe` | Scheffého aproximace nad adjusted means |
| `4` | `games-howell` | v aktuální implementaci fallback přes Bonferroni |

### Kódy `ma_záhlaví`

| Kód | Význam |
| --- | --- |
| `0` | autodetekce záhlaví |
| `1` | první řádek je záhlaví |
| `2` | vstup je bez záhlaví |

### Poznámky

- funkce vyžaduje alespoň dvě skupiny
- kovariáty se zadávají jako jeden nebo více sloupců
- nekompletní řádky se vynechávají jako complete-case
- adjusted means se počítají při globálních průměrech kovariát
- hlavní tabulka zahrnuje i interakce `skupina × kovariáta`; při významné interakci se zobrazí upozornění na porušení homogenity regresních sklonů
- efektové velikosti `η²`, `η²p`, `ω²`, `ω²p` jsou ve výstupu pro faktor, kovariáty i interakce

### Výstup

Spill výstup obsahuje:

- popisné statistiky po skupinách
- jednu společnou ANCOVA tabulku pro faktor, jednotlivé kovariáty i interakce
- případné upozornění na porušení homogenity regresních sklonů
- adjusted means po skupinách
- volitelnou post-hoc část

### Příklad

```excel
=ANCOVA.G(A2:A100;B2:B100;C2:D100;2;0,05;1)
```
