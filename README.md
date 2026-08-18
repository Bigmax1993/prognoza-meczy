# Projekt stawki na mecze

Scraper i eksport statystyk piłkarskich pod analizę kursów / stawek.

**Lokalizacja:** `C:\Users\svinc\Documents\Projekt stawki na mecze`

---

## Spis treści

1. [Co robi projekt](#co-robi-projekt)
2. [Wymagania i instalacja](#wymagania-i-instalacja)
3. [Szybki start](#szybki-start)
4. [Moduły](#moduły)
5. [Pliki wyjściowe](#pliki-wyjściowe)
6. [Cache i logi](#cache-i-logi)
7. [Testy](#testy)
8. [Struktura katalogów](#struktura-katalogów)
9. [Ograniczenia](#ograniczenia)

---

## Co robi projekt

| Moduł | Opis |
|--------|------|
| `scrape_footystats.py` | Pobiera dzisiejsze mecze z [footystats.org](https://footystats.org/ru/), zapisuje **Excel** + cache JSON |
| `export_aleks_stats.py` | Eksportuje **tylko wybrane ligi i 6 statystyk** do CSV (lista Aleksa) |
| `preview_pandas.py` | Moduł podglądu danych w **pandas** (load / preview / summarize) |
| `predykcje.py` | Predykcje meczów 2026 + Excel z pivotami i wykresami |
| `predykcje_max.py` | Predykcje MAX — rozszerzone rynki + backtest + benchmark |
| `enrich_scores.py` | Wynik **результат** + statystyki per drużyna z football-data.co.uk |
| `monthly_summary.py` | Miesięczne średnie statystyk per drużyna |
| `preview_data.ipynb` | Podgląd w **Jupyter Lab** (cache / Excel / CSV Aleksa) |
| `preview_data.py` | Ten sam podgląd jako skrypt / Interactive Window |
| `tests/` | Testy jednostkowe, integracyjne i regresyjne |

### Ligi Aleksa (eksport CSV)

- Premier League (APL)
- La Liga
- Serie A
- Bundesliga
- Bundesliga 2
- Eredivisie
- Super League (Szwajcaria)
- Allsvenskan (Szwecja)
- Eliteserien (Norwegia)

### Statystyki w CSV Aleksa

Nagłówki kolumn są po **ukraińsku**:

1. **оз** (BTTS) — обидві заб'ють (`так` / `ні`)
2. **фоли** — suma fauli; dodatkowo **фоли_господар** / **фоли_гість** (per drużyna)
3. **кутові** — suma rożnych; **кутові_господар** / **кутові_гість**
4. **жовті_картки** — suma kartek; **жовті_картки_господар** / **жовті_картки_гість**
5. **удари** — suma strzałów; **удари_господар** / **удари_гість**
6. **удари_в_площину** — strzały w bramkę; **удари_в_площину_господар** / **удари_в_площину_гість**

Dodatkowo (identyfikacja meczu): `ліга`, `дата`, `господар`, `гість`, **`результат`** (np. `2:1`).  
Statystyki per drużyna i wynik dociągasz z [football-data.co.uk](https://www.football-data.co.uk/):

```powershell
python enrich_scores.py
python monthly_summary.py --year 2026
```

---

## Wymagania i instalacja

- Python **3.10+** (testowane na 3.13)
- Windows (ścieżki projektu)

```powershell
cd "C:\Users\svinc\Documents\Projekt stawki na mecze"
python -m pip install -r requirements.txt
python -m playwright install chromium
```

### Zależności (`requirements.txt`)

- `beautifulsoup4`, `lxml` — parsowanie HTML
- `requests`, `curl_cffi` — pobieranie HTTP
- `playwright` — Chromium **headless** (bez okna przeglądarki)
- `openpyxl`, `pandas` — Excel / podgląd
- `jupyterlab`, `ipykernel` — Jupyter Lab / notebooki
- `pytest` — testy

### Biblioteki wbudowane Pythona

- **`json`** — zapis i odczyt cache (`cache/*.json`, w tym `aleks_stats_cache.json`)
- **`logging`** — logi do `logs/scraper_*.log` oraz `logs/aleks_export_*.log`

---

## Szybki start

### 1. Scraper FootyStats (mecze dnia → Excel)

```powershell
cd "C:\Users\svinc\Documents\Projekt stawki na mecze"
python scrape_footystats.py
```

- Bez `--refresh` korzysta z cache (jeśli świeży).
- Wymuszenie ponownego pobrania:

```powershell
python scrape_footystats.py --refresh
python scrape_footystats.py --refresh --log-level DEBUG
```

**Playwright działa w tle** (`headless=True`, bez otwierania okna Chrome).

Kolejność pobierania HTML:

1. Playwright (headless)
2. `curl_cffi` (impersonacja Chrome)
3. `requests`

### 2. Eksport lig Aleksa → CSV

```powershell
python export_aleks_stats.py
```

Wynik: `aleks_ligi_stats.xlsx` w katalogu projektu.

Źródło: [football-data.co.uk](https://www.football-data.co.uk/) (sezony 2024/25 i 2025/26).

### 3. Podgląd danych (pandas / Jupyter Lab)

Moduł `preview_pandas.py`:

```python
from preview_pandas import load_matches, load_aleks, preview_dataframe, summarize

df = load_matches()
preview_dataframe(df, n=10)
summarize(df)
```

Jupyter Lab:

```powershell
cd "C:\Users\svinc\Documents\Projekt stawki na mecze"
python -m jupyter lab
```

W przeglądarce otwórz `preview_data.ipynb` i uruchamiaj komórki (`Shift+Enter`).

Alternatywa bez Jupyter:

```powershell
python preview_data.py
# albo bezpośrednio moduł:
python preview_pandas.py
```

### 4. Predykcje 2026 → Excel

Moduł `predykcje.py` — predykcje zwycięzcy, wyniku i BTTS **tylko dla roku 2026**:

```powershell
python predykcje.py
```

Wynik: `predykcje_2026.xlsx` z arkuszami:
- `Mecze_2026`, `Predykcje`, `Srednie_druzyny`, `Srednie_ligi`
- `Pivot_{Liga}` — tabela przestawna per liga
- `Podsumowanie_miesiaca` — średnie miesięczne per drużyna (2026)
- `Wykresy` — wykresy kołowe per liga

Metoda: forma ostatnich 5 meczów (70%) + średnie sezonowe (30%); Over/Under rożne 9.5 i kartki 3.5; backtest na 2025 w arkuszach `Backtest_2025`.

### 5. Predykcje MAX → Excel

Rozszerzony moduł `predykcje_max.py` — wszystkie rynki z audytu danych:

```powershell
python predykcje_max.py
```

Wynik: `predykcje_max_2026.xlsx` z arkuszami:
- **Osobny arkusz per liga** (np. `Premier League`, `La Liga`) — prognozy w tabeli Excel
- `Топ_типи` — najlepsze typy dnia
- `Ознаки_{Liga}` — macierz cech (forma, H2H, ranking)
- `Середні_команди`, `Середні_ліги`, `Підсумок_місяця`
- `Бектест_2025`, `Бектест_{Liga}`, `Порівняння`, `Аудит`, `Графіки`

Wszystkie arkusze: **tabele Excel** + **nagłówki i wartości po ukraińsku** (`більше`/`менше`, `так`/`ні`, `висока`/`середня`/`низька`).

Metoda: Warstwa 1 (heurystyka + Poisson-lite) + Warstwa 2 (kalibracja progów per liga na 2024).
Nordic (Allsvenskan/Eliteserien/Super League): tylko rynki wynik/BTTS/O/U goli.

Audyt danych przed predykcją:

```powershell
python scripts/audyt_danych.py
```

---

## Moduły

### `scrape_footystats.py`

| Funkcja / obszar | Opis |
|------------------|------|
| `_is_cloudflare_block` | Wykrywa stronę Cloudflare |
| `fetch_with_playwright` | Pobieranie headless (bez UI) |
| `fetch_html` | Łańcuch metod pobierania |
| `parse_matches` | HTML → lista meczów |
| `save_excel` / cache JSON | Zapis wyników |
| `main` | CLI: `--refresh`, `--log-level` |

Kolumny Excela / cache meczów:

`Data`, `Kraj`, `Liga`, `Godzina`, `Gospodarz`, `Forma gospodarza`, `Gość`, `Forma gościa`, `Kurs 1`, `Kurs X`, `Kurs 2`, `Status`, `Link`

### `export_aleks_stats.py`

- Pobiera CSV z football-data.co.uk
- Filtruje ligi z listy Aleksa
- Liczy sumy meczowe (faule, rożne, kartki, strzały) oraz BTTS
- Zapisuje `aleks_ligi_stats.xlsx`

**Uwaga:** dla Super League / Allsvenskan / Eliteserien publiczny format „new” zawiera głównie wyniki → zwykle wypełnione jest tylko **оз** (BTTS); pozostałe pola mogą być puste.

---

## Pliki wyjściowe

| Plik | Opis |
|------|------|
| `predykcje_2026.xlsx` | Predykcje baseline (forma + O/U) |
| `predykcje_max_2026.xlsx` | Predykcje MAX (wszystkie rynki + benchmark) |
| `AUDYT_DANYCH.md` | Raport audytu danych (kompletność, korelacje) |
| `aleks_ligi_stats.xlsx` | Statystyki lig Aleksa |
| `preview_head.csv` / `.txt` | Podgląd 10 wierszy |
| `cache/matches_cache.json` | Cache meczów |
| `cache/html_cache.json` | Ostatni HTML |
| `cache/cookies_cache.json` | Cookies (Cloudflare / sesja) |
| `logs/scraper_YYYY-MM-DD.log` | Logi dziennika |

---

## Cache i logi

- Cache meczów: domyślnie ważny **30 minut** (TTL).
- Cookies: TTL ok. **60 minut**.
- `--refresh` pomija cache meczów i ponownie pobiera stronę.
- Logi: katalog `logs/`, poziom przez `--log-level INFO|DEBUG|WARNING`.

---

## Testy

```powershell
cd "C:\Users\svinc\Documents\Projekt stawki na mecze"
python -m pytest tests -v
```

### Rodzaje testów

| Plik | Typ | Zakres |
|------|-----|--------|
| `tests/test_unit.py` | jednostkowe | Cloudflare, helpery, cookies, cache JSON |
| `tests/test_integration.py` | integracyjne | parse → cache → Excel, fetch (mock), `main()` |
| `tests/test_regression.py` | regresyjne | schemat kolumn, golden CSV, minimalna liczba meczów |
| pozostałe `test_*.py` | szczegółowe | cache, fetch, parse, pipeline, excel, logging |

Odświeżenie golden fixture:

```powershell
python tests/generate_golden.py
```

---

## Struktura katalogów

```
Projekt stawki na mecze/
├── scrape_footystats.py      # scraper FootyStats
├── export_aleks_stats.py      # CSV lig Aleksa
├── preview_pandas.py           # moduł podglądu pandas
├── predykcje.py                # predykcje 2026 → Excel
├── predykcje_max.py            # predykcje MAX (rozszerzone rynki)
├── scripts/audyt_danych.py     # audyt aleks_ligi_stats.xlsx
├── enrich_scores.py            # wynik + statystyki per drużyna
├── monthly_summary.py          # podsumowanie miesięczne drużyn
├── preview_data.ipynb          # podgląd w Jupyter Lab
├── preview_data.py             # podgląd pandas (skrypt)
├── requirements.txt
├── pytest.ini
├── aleks_ligi_stats.xlsx
├── footystats_mecze.csv
├── cache/
├── logs/
└── tests/
    ├── conftest.py
    ├── test_unit.py
    ├── test_integration.py
    ├── test_regression.py
    ├── fixtures/
    │   ├── sample_html.py
    │   └── golden/sample_mecze.csv
    └── ...
```

---

## Ograniczenia

1. **Cloudflare** na FootyStats może zablokować pobieranie — Playwright headless + cookies łagodzą problem, ale nie gwarantują 100% sukcesu.
2. **Darmowe API FootyStats** (`key=example`) ogranicza dostęp do części lig (głównie Premier League).
3. **Super League / Allsvenskan / Eliteserien** — w publicznym CSV football-data często brak fauli/rożnych/kartek/strzałów (tylko wynik → BTTS).
4. Dane historyczne zależą od dostępności plików na football-data.co.uk (sezony `2425`, `2526`).

---

## Typowy workflow

```powershell
cd "C:\Users\svinc\Documents\Projekt stawki na mecze"

# 1) Mecze dnia z FootyStats
python scrape_footystats.py --refresh

# 2) Statystyki lig pod stawki (Aleks)
python export_aleks_stats.py

# 3) Podgląd w Jupyter Lab
python -m jupyter lab

# 4) Testy po zmianach w kodzie
python -m pytest tests -q
```

W notebooku `preview_data.ipynb` filtruj po `ліга`, `оз`, progach `кутові` / `жовті_картки` itd. Alternatywnie otwórz `aleks_ligi_stats.xlsx` w Excelu.
