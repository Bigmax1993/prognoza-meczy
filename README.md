# Prognoza meczy

Pipeline predykcji piłkarskich: historia lig Aleksa → braki statystyk (JSON → Serper → strona → Claude) → Excel po ukraińsku → mail w poniedziałek.

Repozytorium: [github.com/Bigmax1993/prognoza-meczy](https://github.com/Bigmax1993/prognoza-meczy)  
Actions: [github.com/Bigmax1993/prognoza-meczy/actions](https://github.com/Bigmax1993/prognoza-meczy/actions)

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
2. **Zawsze** weryfikuje braki w **ostatnim tygodniu** rozegranych meczów (domyślnie `--fill-days 7`): ponownie odpytuje `cache/missing_data.json` i uzupełnia luki (faule, rożne, kartki, strzały) z JSON, a resztę z Serper + strony + Claude — **bez zmyślania liczb**. Starsze mecze (od 13.08) zostają w Excelu, ale bez wołania API.
3. Liczy 1X2 (Poisson z oczekiwanych goli), BTTS, O/U rożnych **9.5** i żółtych **3.5** — **tylko dla nadchodzących meczów** (bez wyniku, data ≥ dziś).
4. Zapisuje `predykcje_2026.xlsx` (nagłówki/ligi po ukraińsku, **nazwy klubów bez zmian**). Rozegrane od 13.08 trafiają do **Матчі_2026**; nadchodzące do **Майбутні_матчі** i **Прогнози** (bez kolumny `результат` — typowany wynik to `прогноз_рахунок`).
5. W poniedziałek **05:00** wysyła finalny Excel na Gmail (na Actions: artifact `predykcje-xlsx` z **Pipeline poniedziałek fill**).

Kolejność uzupełniania luk (**zawsze**, na każdym `python predykcje.py`):

Excel → skan pustych komórek **(ostatnie 7 dni)** → ponowne odpytanie `cache/missing_data.json` → **API tylko gdy w JSON też pusto** → Serper + HTML + Claude → walidacja → JSON → Excel → **ponowna weryfikacja zapisanego pliku (te same 7 dni)**.

**Kompletność:** domyślnie pipeline **wywala się**, jeśli po weryfikacji zostaną puste statystyki w ostatnim tygodniu. Na Actions część 1 (discovery) używa `--allow-incomplete` + `--fill-budget-minutes`, żeby zapisać postęp przed limitem 4 h; część 2 (fill) wymaga kompletnego Excela.

Na GitHub Actions JSON braków jest w artifactach / cache workflow, żeby poniedziałkowy fill nie zaczynał Claude od zera.

Przyszłe mecze **nie muszą** mieć faktu FT. Kalendarz idzie do osobnego arkusza.

---

## Wynik: Excel

Plik: **`predykcje_2026.xlsx`**

| Arkusz | Zawartość |
|--------|-----------|
| **Матчі_2026** | Rozegrane mecze **od 13.08.2026** (wynik, **Чи обидві забили?**, faule / rożne / kartki / strzały) |
| **Майбутні_матчі** | Kalendarz (ліга, дата, господар, гість) — bez wyniku |
| **Прогнози** | **Tylko nadchodzące** mecze (jak **Майбутні_матчі**): typy 1X2, BTTS, O/U. **Bez** kolumny `результат` — wynik FT jest tylko w Матчі_2026 |

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
python predykcje.py --fill-missing
python predykcje.py --fill-missing --fill-budget-minutes 210 --allow-incomplete   # jak discovery na Actions
python predykcje.py --fill-missing --restore-excel artifacts/previous_predykcje_2026.xlsx
python predykcje.py --fill-missing --send-mail
python send_mail.py                              # sam Excel na MAIL_TO
python -m pytest tests -q
```

| Flaga | Znaczenie |
|-------|-----------|
| `--fill-missing` | JSON + Serper/Claude + weryfikacja Excela (w CI zawsze włączone) |
| `--fill-days N` | Ile dni wstecz uzupełniać braki (domyślnie 7) |
| `--fill-budget-minutes N` | Miękki limit czasu fill — zapisuje postęp i kończy przed hard limitem joba |
| `--allow-incomplete` | Nie wymagaj pustego zestawu braków (część 1 / discovery) |
| `--restore-excel …` | Uzupełnij puste pola z poprzedniego Excela przed API |
| `--send-mail` | Wyślij gotowy xlsx na `MAIL_TO` |

Źródło historii: `aleks_ligi_stats.xlsx`. Dociągnięcie wyników/statystyk:

```powershell
python enrich_scores.py
python monthly_summary.py --year 2026
```

---

## Klucze API i sekrety

**Nie commituj** `.env` ani haseł. Lokalnie: zmienne User w Windows.  
Na GitHubie sekrety są już w repo: **Settings → Secrets and variables → Actions**.

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
Na GitHub Actions załącznik to artifact z ostatniego udanego **Pipeline poniedziałek fill**, nie Excel z checkoutu gita.

Włącz IMAP: Gmail → Ustawienia → Przekazywanie i POP/IMAP → **Włącz IMAP**.

Hasło do aplikacji: [myaccount.google.com/apppasswords](https://myaccount.google.com/apppasswords) (wymagana weryfikacja dwuetapowa).

---

## GitHub Actions (cron)

Czas poniżej: **Polska, lato (CEST = UTC+2)**. Cron GitHuba jest w UTC. Zimą (CET) crony przesuwają się o **+1 h** względem poniższych godzin PL.

### Harmonogram tygodnia

```text
niedziela 20:00  →  Pipeline niedziela discovery   (max 4 h, soft-stop ~210 min)
poniedziałek 01:00  →  Pipeline poniedziałek fill     (max 4 h, dokończenie luk)
poniedziałek 05:00  →  Wysyłka Gmail                  (~kilka sekund)
```

| Workflow | Cron (UTC) | Kiedy (PL, lato) | Co |
|----------|------------|------------------|-----|
| [Pipeline niedziela discovery](.github/workflows/pipeline.yml) | `0 18 * * 0` | nd **20:00** | Build + Serper/Claude z `--fill-budget-minutes 210 --allow-incomplete` → artifact **`predykcje-partial`** |
| [Pipeline poniedziałek fill](.github/workflows/pipeline-czesc2.yml) | `0 23 * * 0` | pn **01:00** | Bierze partial z discovery → dokańcza fill → artifact **`predykcje-xlsx`** |
| [Wysyłka Gmail poniedziałek](.github/workflows/send-mail.yml) | `0 3 * * 1` | pn **05:00** | Ściąga `predykcje-xlsx` i wysyła mail |
| [Testy](.github/workflows/test.yml) | — (push) | przy pushu na `main` | `pytest tests` |

### Dlaczego dwa runy po 4 h

GitHub **ubija** pojedynczy job po **4 h**. Duży `--fill-missing` często tego nie mieści.  
Dlatego:

1. **Discovery** kończy się miękko przed limitem (`--fill-budget-minutes`) i zostawia Excel + `cache/missing_data.json` w `predykcje-partial`.
2. **Fill** startuje kilka godzin później, wznawia z tego artifactu i publikuje finalne `predykcje-xlsx`.

Ręcznie: **Actions** → najpierw discovery → po sukcesie fill → potem wysyłka.  
Mail bez udanego fillu się wywali.

### Artifacty

| Artifact | Kto publikuje | Retention | Użycie |
|----------|---------------|-----------|--------|
| `predykcje-partial` | discovery | 3 dni | wejście do fill |
| `predykcje-xlsx` | fill | 7 dni | wejście do maila + restore w kolejnym tygodniu |

W checkoutcie jest już `predykcje_2026.xlsx`. `gh run download` **nie nadpisuje** istniejących plików (`file exists`), więc workflowy najpierw `rm` / kopiują do `artifacts/`.

Discovery przed API pobiera poprzednie `predykcje_2026.xlsx` z udanych fillów i **uzupełnia puste statystyki** — mecze starsze niż 7 dni nie tracą danych z poprzedniego tygodnia.

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
| `predykcje.py` | Pipeline 2026, Excel UA, `--fill-missing`, `--fill-budget-minutes`, `--allow-incomplete`, `--send-mail` |
| `fill_missing.py` | JSON → Serper → BS4 → Claude → walidacja + soft deadline |
| `send_mail.py` | Gmail SMTP + kopia w Wysłanych |
| `upcoming.py` | Nadchodzące mecze (BBC) |
| `team_names.py` | Aliasy klubów (bez tłumaczenia nazw) |
| `enrich_scores.py` | Wyniki/statystyki z football-data.co.uk |
| `export_aleks_stats.py` | Eksport lig Aleksa |
| `scrape_footystats.py` | Scraper FootyStats (mecze dnia) |
| `predykcje_max.py` | Predykcje MAX |
| `scripts/audyt_danych.py` | Audyt źródła |
| `.github/workflows/pipeline.yml` | Discovery (nd 20:00) |
| `.github/workflows/pipeline-czesc2.yml` | Fill (pn 01:00) |
| `.github/workflows/send-mail.yml` | Mail (pn 05:00) |

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
- Limit joba GitHub Actions: **4 h** — stąd dwa etapy i `--fill-budget-minutes`.

Typowy przebieg tygodnia: niedziela **20:00** discovery → poniedziałek **01:00** fill → poniedziałek **05:00** mail z `predykcje_2026.xlsx`.
