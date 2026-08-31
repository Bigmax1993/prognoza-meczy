# Prognoza meczy

Pipeline predykcji piłkarskich: historia lig Aleksa → braki statystyk (JSON → Serper → strona → Claude) → Excel po ukraińsku → mail w poniedziałek.

Repozytorium: [github.com/Bigmax1993/prognoza-meczy](https://github.com/Bigmax1993/prognoza-meczy)

---

## Spis treści

1. [Co robi](#co-robi)
2. [Wynik: Excel](#wynik-excel)
3. [Szybki start (lokalnie)](#szybki-start-lokalnie)
4. [Klucze API i sekrety](#klucze-api-i-sekrety)
5. [Gmail](#gmail)
6. [GitHub Actions (cron)](#github-actions-cron)
7. [Jak liczona jest predykcja](#jak-liczona-jest-predykcja)
8. [Moduły](#moduły)
9. [Testy](#testy)
10. [Ograniczenia](#ograniczenia)

---

## Co robi

1. Bierze mecze lig Aleksa **od 13.08.2026** (`--od`), dociąga wyniki FT z BBC za lukę po ostatnim meczu w źródle i dokleja nadchodzące (domyślnie 7 dni).
2. **Zawsze** (bez flagi do wyłączenia) weryfikuje braki, ponownie odpytuje `cache/missing_data.json` i uzupełnia luki (faule, rożne, kartki, strzały) z JSON, a resztę z Serper + strony + Claude — **bez zmyślania liczb**. Pusty Excel po weryfikacji jest błędem.
3. Liczy 1X2 (Poisson z oczekiwanych goli), BTTS, O/U rożnych **9.5** i żółtych **3.5**.
4. Zapisuje `predykcje_2026.xlsx` (nagłówki/ligi po ukraińsku, **nazwy klubów bez zmian**). W arkuszu **Прогнози** nie ma kolumny `результат` — typowany wynik to `прогноз_рахунок`.
5. W poniedziałek wysyła ten plik na Gmail (na Actions: artifact z niedzielnego pipeline).

Kolejność uzupełniania luk (**zawsze**, na każdym `python predykcje.py`):

Excel → skan pustych komórek → ponowne odpytanie `cache/missing_data.json` → **API tylko gdy w JSON też pusto** → Serper + HTML + Claude → walidacja → JSON → Excel → **ponowna weryfikacja zapisanego pliku**.

Pipeline **wywala się**, jeśli po weryfikacji zostaną puste statystyki. Nie ma flagi, która to pomija.

Na GitHub Actions ten JSON jest w repo + w cache workflow, żeby niedzielny run nie walił w Claude od zera.

Przyszłe mecze **nie muszą** mieć faktu FT. Kalendarz idzie do osobnego arkusza.

---

## Wynik: Excel

Plik: **`predykcje_2026.xlsx`**

| Arkusz | Zawartość |
|--------|-----------|
| **Матчі_2026** | Rozegrane mecze **od 13.08.2026** (wynik, **Чи обидві забили?**, faule / rożne / kartki / strzały) |
| **Майбутні_матчі** | Kalendarz (ліга, дата, господар, гість) — bez wyniku |
| **Прогнози** | Typy: `ліга`, `дата`, `господар`, `гість`, `статус`, `причина`, `прогноз_переможець`, prawdopodobieństwa 1X2, `прогноз_рахунок`, BTTS, O/U. **Bez** kolumny `результат` — wynik FT jest tylko w Матчі_2026 |

Linie O/U są stałe celowo: `лінія_кутові` = 9.5, `лінія_жовті` = 3.5. Różne per mecz są `очікувані_*` i `прогноз_*` (більше / менше).

Jeśli Excel jest otwarty (Permission denied), zapis idzie do `predykcje_2026_wypelnione.xlsx`.

---

## Szybki start (lokalnie)

Python **3.10+** (testowane na 3.13).

```powershell
git clone https://github.com/Bigmax1993/prognoza-meczy.git
cd prognoza-meczy
python -m pip install -r requirements.txt
```

Klucze **User env** (nie `.env` w gicie):

```powershell
$env:ANTHROPIC_API_KEY = [Environment]::GetEnvironmentVariable('ANTHROPIC_API_KEY','User')
$env:SERPER_API_KEY    = [Environment]::GetEnvironmentVariable('SERPER_API_KEY','User')
python predykcje.py --fill-missing
```

Opcje:

```powershell
python predykcje.py --fill-missing --send-mail   # pipeline + mail (weryfikacja i tak zawsze)
python send_mail.py                              # sam Excel na MAIL_TO
python -m pytest tests -q
```

Źródło historii: `aleks_ligi_stats.xlsx`. Dociągnięcie wyników/statystyk:

```powershell
python enrich_scores.py
python monthly_summary.py --year 2026
```

---

## Klucze API i sekrety

**Nie commituj** `.env` ani haseł. Lokalnie: zmienne User w Windows. Na GitHubie: **Settings → Secrets and variables → Actions**.

| Secret | Do czego |
|--------|----------|
| `ANTHROPIC_API_KEY` | Claude: wyciąga liczby ze strony (null, jeśli ich nie ma) |
| `SERPER_API_KEY` | Szukanie stron ze statystykami meczu |
| `GMAIL_USER` | Nadawca SMTP, np. `svinchak1993@gmail.com` |
| `GMAIL_APP_PASSWORD` | Hasło do aplikacji Gmail (16 znaków), nie zwykłe hasło |
| `MAIL_TO` | Odbiorca Excela, np. `Swinczakaleksy@gmail.com` |

Przy 401 Anthropic pipeline **się zatrzymuje** (bez spamu requestów).

---

## Gmail

Wysyłka: SMTP `smtp.gmail.com` + załącznik `predykcje_2026.xlsx`.  
Kopia ląduje w **Wysłanych** nadawcy (IMAP, folder Wysłane / Sent Mail).  
Na GitHub Actions załącznik to artifact z ostatniego udanego **Pipeline niedziela**, nie Excel z checkoutu gita.

Włącz IMAP: Gmail → Ustawienia → Przekazywanie i POP/IMAP → **Włącz IMAP**.

Hasło do aplikacji: [myaccount.google.com/apppasswords](https://myaccount.google.com/apppasswords) (wymagana weryfikacja dwuetapowa).

---

## GitHub Actions (cron)

Czas poniżej: **Polska, lato (CEST = UTC+2)**. Cron GitHuba jest w UTC. Zimą (CET) będzie +1 h.

| Workflow | Kiedy | Co |
|----------|--------|-----|
| [Pipeline niedziela](.github/workflows/pipeline.yml) | niedziela **20:00** | `python predykcje.py --fill-missing` + artifact `predykcje-xlsx` (Excel, 7 dni). Limit czasu joba: **4 h** (`timeout-minutes: 240`). |
| [Wysyłka Gmail poniedziałek](.github/workflows/send-mail.yml) | poniedziałek **03:00** (zimą **02:00**) | ściąga artifact z ostatniego udanego pipeline i wysyła mail |
| [Testy](.github/workflows/test.yml) | push na `main` (też ręcznie) | `pytest tests` |

Ręcznie: **Actions** → wybrany workflow → **Run workflow**.  
Najpierw odpal pipeline, potem wysyłkę — mail bez niedzielnego artifactu się wywali.  
Duży `--fill-missing` (wiele luk w JSON) może trwać **ponad godzinę** — stąd 4-godzinny limit w pipeline.

W checkoutcie jest już `predykcje_2026.xlsx`. `gh run download` **nie nadpisuje** plików (`file exists`), więc wysyłka ściąga artifact do pustego `artifacts/`, a potem kopiuje go na `predykcje_2026.xlsx`. Mail idzie z Excela z niedzielnego pipeline, nie z gita.

---

## Jak liczona jest predykcja

- Forma ostatnich **5** meczów (70%) + sezon (30%).
- 1X2: Poisson z oczekiwanych goli (remis **nie** jest stały).
- BTTS: średni % «tak» z formy obu drużyn, próg 50%.
- Rożne / kartki: suma średnich obu drużyn vs linia 9.5 / 3.5.
- Za mało historii (`< 3` mecze): status `пропуск`, przyczyna `замало_даних`.

Rozszerzony wariant: `python predykcje_max.py` → `predykcje_max_2026.xlsx`.

---

## Moduły

| Plik | Rola |
|------|------|
| `predykcje.py` | Pipeline 2026, Excel UA, `--fill-missing`, `--send-mail` |
| `fill_missing.py` | JSON → Serper → BS4 → Claude → walidacja |
| `send_mail.py` | Gmail SMTP + kopia w Wysłanych |
| `upcoming.py` | Nadchodzące mecze (BBC) |
| `team_names.py` | Aliasy klubów (bez tłumaczenia nazw) |
| `enrich_scores.py` | Wyniki/statystyki z football-data.co.uk |
| `export_aleks_stats.py` | Eksport lig Aleksa |
| `scrape_footystats.py` | Scraper FootyStats (mecze dnia) |
| `predykcje_max.py` | Predykcje MAX |
| `scripts/audyt_danych.py` | Audyt źródła |

Ligi: Premier League, La Liga, Serie A, Bundesliga, Bundesliga 2, Eredivisie, Super League, Allsvenskan, Eliteserien.

---

## Testy

```powershell
python -m pytest tests -q
```

Kluczowe: `tests/test_predykcje.py`, `tests/test_fill_missing.py`, `tests/test_send_mail.py`.  
Na GitHubie to samo robi workflow [Testy](.github/workflows/test.yml) przy pushu na `main`.

---

## Ograniczenia

- Claude **nie zgaduje** fauli/rożnych/kartek — brak na stronie = puste pole / residual.
- Nordic (Allsvenskan, Eliteserien, Super League): w publicznym CSV często tylko wynik; reszta z JSON/API albo puste.
- Cloudflare na FootyStats może blokować scraper.
- Nie tłumacz nazw klubów (Arsenal, Sarpsborg 08, Elfsborg…).

Typowy przebieg tygodnia: niedziela 20:00 pipeline (do 4 h) → poniedziałek 03:00 mail z `predykcje_2026.xlsx`.
