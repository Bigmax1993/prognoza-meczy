# AUDYT DANYCH — aleks_ligi_stats.xlsx

**Plik:** `aleks_ligi_stats.xlsx`  
**Wiersze:** 6422  
**Kolumny:** 27  
**Lata:** [2024, 2025, 2026]  
**Ligi:** 9

## 0. Walidacja okna predykcji

**Reguła:** mecze w `predykcje_2026.xlsx` od **13/08/2026** (włącznie).  
Historia sprzed tej daty w `aleks_ligi_stats.xlsx` jest dozwolona (forma / backtest).

- Źródło: **38** meczów od 13/08/2026 (z 6422 łącznie)
- `Mecze_2026`: **OK** — 38 meczów, wszystkie od 13/08/2026
- `Predykcje`: **OK** — 38 meczów, wszystkie od 13/08/2026
- Wynik walidacji: **OK**

## 1. Kompletność per liga × kolumna

| ліга | mecze | результат | оз | фоли_господар | кутові_господар | жовті_картки_господар | удари_господар | удари_в_площину_господар | фоли | кутові | жовті_картки | удари | удари_в_площину |
| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |
| Allsvenskan | 617 | 100.0 | 100.0 | 0.0 | 0.0 | 0.0 | 0.0 | 0.0 | 0.0 | 0.0 | 0.0 | 0.0 | 0.0 |
| Bundesliga | 684 | 100.0 | 100.0 | 99.9 | 99.9 | 99.9 | 99.9 | 99.9 | 99.9 | 99.9 | 99.9 | 99.9 | 99.9 |
| Bundesliga 2 | 702 | 100.0 | 100.0 | 100.0 | 100.0 | 100.0 | 100.0 | 100.0 | 100.0 | 100.0 | 100.0 | 100.0 | 100.0 |
| Eliteserien | 618 | 100.0 | 100.0 | 0.0 | 0.0 | 0.0 | 0.0 | 0.0 | 0.0 | 0.0 | 0.0 | 0.0 | 0.0 |
| Eredivisie | 702 | 100.0 | 100.0 | 100.0 | 100.0 | 100.0 | 100.0 | 100.0 | 100.0 | 100.0 | 100.0 | 100.0 | 100.0 |
| La Liga | 855 | 100.0 | 100.0 | 100.0 | 100.0 | 100.0 | 100.0 | 100.0 | 100.0 | 100.0 | 100.0 | 100.0 | 100.0 |
| Premier League | 857 | 100.0 | 100.0 | 100.0 | 100.0 | 100.0 | 100.0 | 100.0 | 100.0 | 100.0 | 100.0 | 100.0 | 100.0 |
| Serie A | 851 | 100.0 | 100.0 | 100.0 | 100.0 | 100.0 | 100.0 | 100.0 | 100.0 | 100.0 | 100.0 | 100.0 | 100.0 |
| Super League | 536 | 100.0 | 100.0 | 0.0 | 0.0 | 0.0 | 0.0 | 0.0 | 0.0 | 0.0 | 0.0 | 0.0 | 0.0 |

### Ligi z niepełnymi statystykami per drużyna

- **Allsvenskan**: brak stat per team (~0.0% фоли_господар) — tylko wynik + BTTS + gole
- **Bundesliga**: pełne statystyki (99.9% кутові)
- **Bundesliga 2**: pełne statystyki (100.0% кутові)
- **Eliteserien**: brak stat per team (~0.0% фоли_господар) — tylko wynik + BTTS + gole
- **Eredivisie**: pełne statystyki (100.0% кутові)
- **La Liga**: pełne statystyki (100.0% кутові)
- **Premier League**: pełne statystyki (100.0% кутові)
- **Serie A**: pełne statystyki (100.0% кутові)
- **Super League**: brak stat per team (~0.0% фоли_господар) — tylko wynik + BTTS + gole

## 2. Rozkłady (średnia / mediana / p75) per liga

### Allsvenskan

| metryka | mean | median | p75 |
|---------|------|--------|-----|
| _total_goals | 2.89 | 3.0 | 4.0 |

### Bundesliga

| metryka | mean | median | p75 |
|---------|------|--------|-----|
| _total_goals | 3.19 | 3.0 | 4.0 |
| кутові | 9.72 | 10.0 | 12.0 |
| жовті_картки | 3.82 | 4.0 | 5.0 |
| фоли | 21.11 | 21.0 | 25.0 |
| удари | 26.37 | 26.0 | 30.0 |
| удари_в_площину | 9.47 | 9.0 | 11.0 |

### Bundesliga 2

| metryka | mean | median | p75 |
|---------|------|--------|-----|
| _total_goals | 2.98 | 3.0 | 4.0 |
| кутові | 10.22 | 10.0 | 12.0 |
| жовті_картки | 4.57 | 5.0 | 6.0 |
| фоли | 23.93 | 24.0 | 28.0 |
| удари | 26.89 | 27.0 | 31.0 |
| удари_в_площину | 9.31 | 9.0 | 11.0 |

### Eliteserien

| metryka | mean | median | p75 |
|---------|------|--------|-----|
| _total_goals | 3.01 | 3.0 | 4.0 |

### Eredivisie

| metryka | mean | median | p75 |
|---------|------|--------|-----|
| _total_goals | 3.14 | 3.0 | 4.0 |
| кутові | 10.27 | 10.0 | 12.8 |
| жовті_картки | 3.03 | 3.0 | 4.0 |
| фоли | 20.92 | 21.0 | 24.0 |
| удари | 27.49 | 27.5 | 31.0 |
| удари_в_площину | 9.89 | 10.0 | 12.0 |

### La Liga

| metryka | mean | median | p75 |
|---------|------|--------|-----|
| _total_goals | 2.66 | 2.0 | 3.5 |
| кутові | 9.49 | 9.0 | 12.0 |
| жовті_картки | 4.46 | 4.0 | 6.0 |
| фоли | 24.93 | 25.0 | 29.0 |
| удари | 24.31 | 24.0 | 28.0 |
| удари_в_площину | 8.43 | 8.0 | 10.0 |

### Premier League

| metryka | mean | median | p75 |
|---------|------|--------|-----|
| _total_goals | 2.90 | 3.0 | 4.0 |
| кутові | 10.24 | 10.0 | 12.0 |
| жовті_картки | 3.87 | 4.0 | 5.0 |
| фоли | 21.85 | 22.0 | 25.0 |
| удари | 25.82 | 26.0 | 29.0 |
| удари_в_площину | 8.92 | 9.0 | 11.0 |

### Serie A

| metryka | mean | median | p75 |
|---------|------|--------|-----|
| _total_goals | 2.51 | 2.0 | 3.0 |
| кутові | 9.11 | 9.0 | 11.0 |
| жовті_картки | 3.75 | 4.0 | 5.0 |
| фоли | 25.08 | 25.0 | 29.0 |
| удари | 24.51 | 24.0 | 28.0 |
| удари_в_площину | 8.10 | 8.0 | 10.0 |

### Super League

| metryka | mean | median | p75 |
|---------|------|--------|-----|
| _total_goals | 3.08 | 3.0 | 4.0 |

## 3. Korelacje ze wynikiem

- **Różnica SoT vs różnica goli (global):** 0.609
- **Suma rożnych vs over 9.5 (proxy):** 0.802
- **Suma kartek vs over 3.5 (proxy):** 0.789

### Home win rate per liga

| liga | home_win_pct | avg_goli | avg_rozne | avg_kartki |
|------|--------------|----------|-----------|------------|
| Allsvenskan | 42.3% | 2.89 | nan | nan |
| Bundesliga | 41.4% | 3.19 | 9.72 | 3.82 |
| Bundesliga 2 | 43.3% | 2.98 | 10.22 | 4.57 |
| Eliteserien | 46.4% | 3.01 | nan | nan |
| Eredivisie | 44.6% | 3.14 | 10.27 | 3.03 |
| La Liga | 46.2% | 2.66 | 9.49 | 4.46 |
| Premier League | 42.2% | 2.90 | 10.24 | 3.87 |
| Serie A | 38.9% | 2.51 | 9.11 | 3.75 |
| Super League | 43.7% | 3.08 | nan | nan |

## 4. Leakage / jakość danych

- **Brak daty:** 0 wierszy
- **Duplikaty meczów (liga+data+home+away):** 0
- **Max жовті_картки:** 15.0 | outliery >30: 0
- **Max кутові:** 26.0 | outliery >30: 0
- **Max фоли:** 47.0 | outliery >30: 465

## 5. Co da się realnie przewidywać

| Rynek | Pokrycie danych | Sensowność | Uwagi |
|-------|-----------------|------------|-------|
| 1X2 + wynik | результат 100.0% (6422/6422) | TAK | Pełna historia 2024–2026 |
| BTTS так/ні | оз 100.0% (6422/6422) | TAK | 100% wypełnienie, baseline ~52% |
| O/U goli 1.5/2.5/3.5 | результат 100.0% | TAK | Poisson-lite z exp GF/GA |
| O/U rożne 8.5/9.5/10.5 | кутові 72.4% | WARUNKOWO | Nordic bez stat — pomiń |
| O/U kartki 2.5/3.5/4.5 | жовті 72.4% | WARUNKOWO | Nordic bez stat — pomiń |
| O/U fauli | фоли 72.4% | WARUNKOWO | Próg = mediana ligi (4650 meczy) |
| Team totals rożne H/A | split 72.4% (4650/6422) | WARUNKOWO | Tylko top-6 lig |
| Handicap 0/-0.5 | результат + forma 100.0% | TAK | Z exp goals Poisson-lite |
| Dokładny wynik Poisson | результат 100.0% | WARUNKOWO | Niska hit-rate, info value |
| SoT / strzały heurystyka | SoT split 72.4% | WARUNKOWO | Fallback gdy brak wyniku historycznego |
| H2H + ranking punktowy | 100% klucze meczów | TAK | Feature store as_of |