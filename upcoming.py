# -*- coding: utf-8 -*-
"""Nadchodzące mecze lig Aleksa z FootyStats (bez wyniku).

Nie nadpisuje aleks_ligi_stats.xlsx — tylko cache + DataFrame do predykcji.

Uruchomienie:
  python upcoming.py
  python upcoming.py --days 5 --refresh
"""
from __future__ import annotations

import argparse
import json
import logging
import re
import sys
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any

import pandas as pd
import requests
from bs4 import BeautifulSoup

import export_aleks_stats as aleks
import scrape_footystats as sf
from team_names import TEAM_ALIASES, canonical_name, known_teams_by_league, map_team_to_known

ROOT = Path(__file__).resolve().parent
CACHE_PATH = ROOT / "cache" / "upcoming_fixtures.json"
CACHE_TTL_HOURS = 6
URL_EN = "https://footystats.org/"
BBC_FIXTURES = "https://www.bbc.com/sport/football/scores-fixtures/{iso}"
BBC_UA = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
)
VERSUS_RE = re.compile(r"^(.+?)\s+versus\s+(.+?)\s+kick off\s+", re.I)
SCORE_RE = re.compile(
    r"^(.+?)\s+(\d+)\s*,\s*(.+?)\s+(\d+)\s+at\s+Full time\b",
    re.I,
)
MAX_LOOKBACK_DAYS = 21

BBC_LEAGUE_MAP = {
    "Premier League": "Premier League",
    "Spanish La Liga": "La Liga",
    "Italian Serie A": "Serie A",
    "German Bundesliga": "Bundesliga",
    "German 2. Bundesliga": "Bundesliga 2",
    "Dutch Eredivisie": "Eredivisie",
    "Swedish Allsvenskan": "Allsvenskan",
    "Norwegian Eliteserien": "Eliteserien",
    "Swiss Super League": "Super League",
}

logger = logging.getLogger("upcoming")

COL_LIGA = "ліга"
COL_DATA = "дата"
COL_HOME = "господар"
COL_AWAY = "гість"
COL_RESULT = "результат"

STAT_EMPTY = {
    "фоли_господар": "",
    "фоли_гість": "",
    "фоли": "",
    "кутові_господар": "",
    "кутові_гість": "",
    "кутові": "",
    "жовті_картки_господар": "",
    "жовті_картки_гість": "",
    "жовті_картки": "",
    "удари_господар": "",
    "удари_гість": "",
    "удари": "",
    "удари_в_площину_господар": "",
    "удари_в_площину_гість": "",
    "удари_в_площину": "",
    "оз": "",
}


def _safe_print(text: str) -> None:
    try:
        if hasattr(sys.stdout, "reconfigure"):
            try:
                sys.stdout.reconfigure(errors="replace")
            except Exception:
                pass
        print(text)
    except UnicodeEncodeError:
        enc = sys.stdout.encoding or "utf-8"
        print(text.encode(enc, errors="replace").decode(enc, errors="replace"))


def urls_for_date(day: date) -> list[str]:
    """Kandydaci URL FootyStats na dany dzień (EN, potem RU)."""
    iso = day.strftime("%Y-%m-%d")
    dmy = day.strftime("%d-%m-%Y")
    month = day.strftime("%B").lower()
    return [
        f"{URL_EN}?date={iso}",
        f"{URL_EN}{day.day}-{month}-{day.year}",
        f"{sf.URL}?date={iso}",
        f"{sf.URL}?date={dmy}",
        URL_EN if day == date.today() else "",
        sf.URL if day == date.today() else "",
    ]


def complete_match_date(dd_mm: str, *, year: int | None = None, fallback: date | None = None) -> str:
    """Zamienia 18/08 na 18/08/2026."""
    fb = fallback or date.today()
    y = year or fb.year
    text = (dd_mm or "").strip()
    m = re.match(r"^(\d{1,2})[./-](\d{1,2})(?:[./-](\d{2,4}))?$", text)
    if not m:
        return fb.strftime("%d/%m/%Y")
    d, mo, yr = int(m.group(1)), int(m.group(2)), m.group(3)
    if yr:
        y = int(yr)
        if y < 100:
            y += 2000
    try:
        return date(y, mo, d).strftime("%d/%m/%Y")
    except ValueError:
        return fb.strftime("%d/%m/%Y")


def _bbc_match_count(el) -> int:
    n = 0
    for span in el.select("span.visually-hidden"):
        text = span.get_text(" ", strip=True)
        if VERSUS_RE.match(text) or SCORE_RE.match(text):
            n += 1
    return n


def _league_block(h2):
    p = h2
    for _ in range(8):
        p = getattr(p, "parent", None)
        if p is None:
            return None
        n_h2 = len(p.select("h2[class*=GroupHeader]"))
        if _bbc_match_count(p) > 0 and n_h2 == 1:
            return p
    return None


def _abs_bbc_url(href: str) -> str:
    h = (href or "").strip()
    if not h:
        return ""
    if h.startswith("http://") or h.startswith("https://"):
        return h
    if h.startswith("/"):
        return "https://www.bbc.com" + h
    return "https://www.bbc.com/" + h


def _team_pair_key(home: str, away: str) -> tuple[str, str]:
    return (canonical_name(str(home or "")), canonical_name(str(away or "")))


def _live_url_map_from_box(box) -> dict[tuple[str, str], str]:
    """Mapa (canonical home, away) → URL live ze scorera BBC."""
    out: dict[tuple[str, str], str] = {}
    if box is None:
        return out
    for a in box.select('a[href*="/sport/football/live/"]'):
        href = _abs_bbc_url(a.get("href") or "")
        if not href:
            continue
        candidates: list[str] = []
        for span in a.select("span.visually-hidden"):
            candidates.append(span.get_text(" ", strip=True))
        candidates.append(a.get_text(" ", strip=True))
        for text in candidates:
            m = SCORE_RE.match(text) or SCORE_RE.search(text)
            if m:
                out[_team_pair_key(m.group(1), m.group(3))] = href
                break
            m = VERSUS_RE.match(text) or VERSUS_RE.search(text)
            if m:
                out[_team_pair_key(m.group(1), m.group(2))] = href
                break
    return out


def parse_bbc_fixtures(html: str, match_day: date) -> list[dict]:
    """Parsuje mecze lig Aleksa ze strony BBC Sport (FT + nadchodzące + live URL)."""
    soup = BeautifulSoup(html or "", "lxml")
    rows: list[dict] = []
    date_str = match_day.strftime("%d/%m/%Y")
    for h2 in soup.select("h2[class*=GroupHeader]"):
        title = h2.get_text(" ", strip=True)
        liga = BBC_LEAGUE_MAP.get(title)
        if not liga:
            continue
        box = _league_block(h2)
        if box is None:
            continue
        live_map = _live_url_map_from_box(box)
        for span in box.select("span.visually-hidden"):
            text = span.get_text(" ", strip=True)
            m = VERSUS_RE.match(text)
            if m:
                home, away = m.group(1).strip(), m.group(2).strip()
                rows.append(
                    {
                        "Kraj": "",
                        "Liga": liga,
                        "_aleks_liga": liga,
                        "Gospodarz": home,
                        "Gość": away,
                        "Data": date_str,
                        "Wynik": "",
                        "bbc_live_url": live_map.get(_team_pair_key(home, away), ""),
                    }
                )
                continue
            m = SCORE_RE.match(text)
            if not m:
                continue
            home, away = m.group(1).strip(), m.group(3).strip()
            hg, ag = int(m.group(2)), int(m.group(4))
            rows.append(
                {
                    "Kraj": "",
                    "Liga": liga,
                    "_aleks_liga": liga,
                    "Gospodarz": home,
                    "Gość": away,
                    "Data": date_str,
                    "Wynik": f"{hg}:{ag}",
                    "bbc_live_url": live_map.get(_team_pair_key(home, away), ""),
                }
            )
    # span.visually-hidden bywa i w linku live, i obok — usuń duplikaty
    deduped: list[dict] = []
    seen: set[tuple[str, str, str, str]] = set()
    for row in rows:
        key = (
            str(row.get("Liga") or ""),
            str(row.get("Gospodarz") or ""),
            str(row.get("Gość") or ""),
            str(row.get("Wynik") or ""),
        )
        if key in seen:
            # uzupełnij URL, jeśli wcześniejszy wiersz go nie miał
            if row.get("bbc_live_url"):
                for prev in deduped:
                    pkey = (
                        str(prev.get("Liga") or ""),
                        str(prev.get("Gospodarz") or ""),
                        str(prev.get("Gość") or ""),
                        str(prev.get("Wynik") or ""),
                    )
                    if pkey == key and not prev.get("bbc_live_url"):
                        prev["bbc_live_url"] = row["bbc_live_url"]
            continue
        seen.add(key)
        deduped.append(row)
    return deduped


def parse_bbc_initial_data(html: str) -> dict[str, Any] | None:
    """Wyciąga JSON z window.__INITIAL_DATA__ na stronie BBC live."""
    text = html or ""
    marker = 'window.__INITIAL_DATA__="'
    start = text.find(marker)
    if start < 0:
        marker = "window.__INITIAL_DATA__='"
        start = text.find(marker)
        if start < 0:
            return None
        quote = marker[-1]
    else:
        quote = '"'
    start += len(marker)
    end = text.find(f"{quote};</script>", start)
    if end < 0:
        end = text.find(f"{quote};", start)
    if end < 0:
        return None
    raw = text[start:end]
    try:
        decoded = raw.encode("utf-8").decode("unicode_escape")
        return json.loads(decoded)
    except (UnicodeDecodeError, json.JSONDecodeError, ValueError):
        try:
            # fallback: ręczna zamiana \" → "
            decoded = (
                raw.replace(r"\\", "\\")
                .replace(r"\"", '"')
                .replace(r"\/", "/")
                .replace(r"\n", "\n")
                .replace(r"\r", "\r")
                .replace(r"\t", "\t")
            )
            return json.loads(decoded)
        except (json.JSONDecodeError, ValueError):
            return None


def _bbc_stat_total(stats: dict[str, Any] | None, *path: str) -> int | None:
    node: Any = stats or {}
    for key in path:
        if not isinstance(node, dict):
            return None
        node = node.get(key)
    if node is None:
        return None
    if isinstance(node, dict) and "total" in node:
        node = node.get("total")
    try:
        return int(round(float(node)))
    except (TypeError, ValueError):
        return None


def _find_bbc_team_stats_block(data: Any) -> dict[str, Any] | None:
    """Szuka węzła z homeTeam.stats / awayTeam.stats (match-stats)."""
    if isinstance(data, dict):
        home = data.get("homeTeam")
        away = data.get("awayTeam")
        if isinstance(home, dict) and isinstance(away, dict):
            if isinstance(home.get("stats"), dict) or isinstance(away.get("stats"), dict):
                return data
        for value in data.values():
            found = _find_bbc_team_stats_block(value)
            if found is not None:
                return found
    elif isinstance(data, list):
        for value in data:
            found = _find_bbc_team_stats_block(value)
            if found is not None:
                return found
    return None


def extract_bbc_match_stats(html: str) -> dict[str, Any] | None:
    """Mapuje boxscore BBC live → kolumny UA (faule/rożne/kartki/strzały)."""
    payload = parse_bbc_initial_data(html)
    if not payload:
        return None
    block = _find_bbc_team_stats_block(payload)
    if not block:
        return None
    home_stats = ((block.get("homeTeam") or {}).get("stats")) or {}
    away_stats = ((block.get("awayTeam") or {}).get("stats")) or {}
    if not home_stats and not away_stats:
        return None

    def pair(h: int | None, a: int | None) -> tuple[int | None, int | None, int | None]:
        if h is None and a is None:
            return None, None, None
        if h is not None and a is not None:
            return h, a, h + a
        return h, a, None

    fh = _bbc_stat_total(home_stats, "foulsCommitted")
    fa = _bbc_stat_total(away_stats, "foulsCommitted")
    ch = _bbc_stat_total(home_stats, "cornersWon")
    ca = _bbc_stat_total(away_stats, "cornersWon")
    yh = _bbc_stat_total(home_stats, "defence", "totalYellowCard")
    if yh is None:
        yh = _bbc_stat_total(home_stats, "totalYellowCard")
    ya = _bbc_stat_total(away_stats, "defence", "totalYellowCard")
    if ya is None:
        ya = _bbc_stat_total(away_stats, "totalYellowCard")
    sh = _bbc_stat_total(home_stats, "shotsTotal")
    sa = _bbc_stat_total(away_stats, "shotsTotal")
    soh = _bbc_stat_total(home_stats, "shotsOnTarget")
    soa = _bbc_stat_total(away_stats, "shotsOnTarget")

    out: dict[str, Any] = {}
    mapping = [
        ("фоли", *pair(fh, fa)),
        ("кутові", *pair(ch, ca)),
        ("жовті_картки", *pair(yh, ya)),
        ("удари", *pair(sh, sa)),
        ("удари_в_площину", *pair(soh, soa)),
    ]
    for base, hv, av, tot in mapping:
        if hv is not None:
            out[f"{base}_господар"] = hv
        if av is not None:
            out[f"{base}_гість"] = av
        if tot is not None:
            out[base] = tot
    return out or None


def fetch_bbc_match_stats(url: str, *, timeout: int = 45) -> dict[str, Any] | None:
    """GET strony BBC live → dict kolumn UA albo None."""
    href = _abs_bbc_url(url)
    if not href:
        return None
    try:
        r = requests.get(href, timeout=timeout, headers={"User-Agent": BBC_UA})
        if r.status_code != 200 or len(r.content) < 500:
            logger.warning("BBC live %s: HTTP %s", href, r.status_code)
            return None
        html = r.content.decode("utf-8", errors="replace")
        stats = extract_bbc_match_stats(html)
        if not stats:
            logger.info("BBC live %s: brak match-stats w HTML", href)
        return stats
    except Exception as exc:
        logger.warning("BBC live fail %s: %s", href, exc)
        return None


def resolve_bbc_live_url(
    match_day: date,
    home: str,
    away: str,
    *,
    html: str | None = None,
    rows: list[dict] | None = None,
) -> str:
    """Szuka URL live dla meczu (HTML dnia, lista rows albo cache fixtures)."""
    key = _team_pair_key(home, away)
    if rows:
        for row in rows:
            if _team_pair_key(row.get("Gospodarz") or "", row.get("Gość") or "") == key:
                url = str(row.get("bbc_live_url") or "").strip()
                if url:
                    return url
    if html:
        for row in parse_bbc_fixtures(html, match_day):
            if _team_pair_key(row.get("Gospodarz") or "", row.get("Gość") or "") == key:
                url = str(row.get("bbc_live_url") or "").strip()
                if url:
                    return url
    day_key = match_day.strftime("%Y-%m-%d")
    cached = load_upcoming_cache(ttl_hours=24 * 14)
    if cached and isinstance(cached.get("days"), dict):
        for row in cached["days"].get(day_key) or []:
            if _team_pair_key(row.get("Gospodarz") or "", row.get("Gość") or "") == key:
                url = str(row.get("bbc_live_url") or "").strip()
                if url:
                    return url
    fetched = fetch_bbc_day(match_day)
    for row in fetched:
        if _team_pair_key(row.get("Gospodarz") or "", row.get("Gość") or "") == key:
            return str(row.get("bbc_live_url") or "").strip()
    return ""


def fetch_bbc_day(day: date) -> list[dict]:
    url = BBC_FIXTURES.format(iso=day.strftime("%Y-%m-%d"))
    try:
        r = requests.get(url, timeout=45, headers={"User-Agent": BBC_UA})
        if r.status_code != 200 or len(r.content) < 500:
            logger.warning("BBC %s: HTTP %s", day, r.status_code)
            return []
        html = r.content.decode("utf-8", errors="replace")
        rows = parse_bbc_fixtures(html, day)
        logger.info("BBC %s: %s meczow lig Aleksa", day, len(rows))
        return rows
    except Exception as exc:
        logger.warning("BBC %s fail: %s", day, exc)
        return []


def rows_to_aleks_fixtures(
    rows: list[dict],
    *,
    match_day: date,
    known: dict[str, list[str]] | None = None,
) -> pd.DataFrame:
    """Filtr lig Aleksa → wiersze (wynik FT, jeśli BBC go podał)."""
    known = known or {}
    ready = [r for r in rows if r.get("_aleks_liga")]
    rest = [r for r in rows if not r.get("_aleks_liga")]
    aleks_rows = list(ready) + (aleks.filter_aleks_matches(rest) if rest else [])
    out: list[dict[str, Any]] = []
    date_str = match_day.strftime("%d/%m/%Y")
    for m in aleks_rows:
        liga = str(m.get("_aleks_liga") or "")
        home = map_team_to_known(str(m.get("Gospodarz") or ""), liga, known)
        away = map_team_to_known(str(m.get("Gość") or ""), liga, known)
        if not home or not away:
            continue
        score = str(m.get("Wynik") or m.get(COL_RESULT) or "").strip()
        parsed = re.match(r"^(\d+)\s*[:\-]\s*(\d+)$", score)
        btts = ""
        if parsed:
            hg, ag = int(parsed.group(1)), int(parsed.group(2))
            score = f"{hg}:{ag}"
            btts = "так" if hg > 0 and ag > 0 else "ні"
        else:
            score = ""
        stats = dict(STAT_EMPTY)
        stats["оз"] = btts
        live_url = str(m.get("bbc_live_url") or "").strip()
        row_out: dict[str, Any] = {
            COL_LIGA: liga,
            COL_DATA: date_str,
            COL_HOME: home,
            COL_AWAY: away,
            COL_RESULT: score,
            **stats,
        }
        if live_url:
            row_out["_bbc_live_url"] = live_url
        out.append(row_out)
    if not out:
        return pd.DataFrame(columns=[COL_LIGA, COL_DATA, COL_HOME, COL_AWAY, COL_RESULT])
    return pd.DataFrame(out).drop_duplicates(subset=[COL_LIGA, COL_DATA, COL_HOME, COL_AWAY])


def fetch_day_html(day: date) -> tuple[str, str] | None:
    """Pobiera HTML dnia; pierwsza URL z meczami wygrywa."""
    for url in urls_for_date(day):
        if not url:
            continue
        try:
            html, source = sf.fetch_html(url)
        except SystemExit:
            continue
        except Exception as exc:
            logger.warning("Fetch fail %s: %s", url, exc)
            continue
        if html and not sf._is_cloudflare_block(html, sf._extract_title(html)):
            _, rows = sf.parse_matches(html)
            if rows:
                logger.info("Dzien %s: %s meczow z %s (%s)", day, len(rows), url, source)
                return html, source
            logger.info("Dzien %s: HTML OK ale 0 meczow (%s)", day, url)
    return None


def fetch_day_rows(day: date) -> list[dict]:
    """BBC Sport — dni bez lig Aleksa zwracają [] (to normalne, nie fallback do CF)."""
    return fetch_bbc_day(day)


def load_upcoming_cache(*, ttl_hours: int = CACHE_TTL_HOURS) -> dict | None:
    try:
        if not CACHE_PATH.exists():
            return None
        payload = json.loads(CACHE_PATH.read_text(encoding="utf-8"))
        cached_at = payload.get("cached_at")
        if not cached_at:
            return None
        ts = datetime.fromisoformat(cached_at)
        if datetime.now() - ts > timedelta(hours=ttl_hours):
            return None
        return payload
    except Exception:
        return None


def save_upcoming_cache(by_day: dict[str, list[dict]]) -> Path:
    CACHE_PATH.parent.mkdir(parents=True, exist_ok=True)
    payload = {
        "cached_at": datetime.now().isoformat(timespec="seconds"),
        "days": by_day,
    }
    CACHE_PATH.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    return CACHE_PATH


def lookback_days_from_history(
    history: pd.DataFrame | None,
    *,
    start: date,
    max_lookback: int = MAX_LOOKBACK_DAYS,
) -> int:
    """Ile dni wstecz dociągnąć wyniki BBC (luka od ostatniego meczu w źródle)."""
    if history is None or history.empty or COL_DATA not in history.columns:
        return 0
    last = pd.to_datetime(history[COL_DATA], dayfirst=True, errors="coerce").max()
    if pd.isna(last):
        return 0
    gap = (start - last.date()).days
    return int(max(0, min(max_lookback, gap)))


def fetch_upcoming_fixtures(
    *,
    days: int = 7,
    start: date | None = None,
    refresh: bool = False,
    history: pd.DataFrame | None = None,
    lookback_days: int | None = None,
    max_lookback: int = MAX_LOOKBACK_DAYS,
) -> pd.DataFrame:
    """Pobiera mecze BBC: luka wstecz (wyniki FT) + `days` dni do przodu."""
    start = start or date.today()
    if lookback_days is None:
        lookback_days = lookback_days_from_history(
            history, start=start, max_lookback=max_lookback
        )
    lookback_days = max(0, int(lookback_days))
    first = start - timedelta(days=lookback_days)
    n_days = lookback_days + max(1, days)
    wanted = [first + timedelta(days=i) for i in range(n_days)]
    wanted_keys = [d.strftime("%Y-%m-%d") for d in wanted]

    by_day: dict[str, list[dict]] = {}
    if not refresh:
        cached = load_upcoming_cache()
        if cached and isinstance(cached.get("days"), dict):
            by_day = {k: v for k, v in cached["days"].items() if k in wanted_keys}

    missing = [d for d, k in zip(wanted, wanted_keys) if k not in by_day]
    if missing:
        for day in missing:
            try:
                rows = fetch_day_rows(day)
            except Exception:
                logger.exception("Nie udalo sie pobrac %s", day)
                rows = []
            by_day[day.strftime("%Y-%m-%d")] = rows
        save_upcoming_cache(by_day)

    known = known_teams_by_league(history) if history is not None else {}
    frames: list[pd.DataFrame] = []
    for day in wanted:
        rows = by_day.get(day.strftime("%Y-%m-%d")) or []
        df = rows_to_aleks_fixtures(rows, match_day=day, known=known)
        if not df.empty:
            frames.append(df)
    if not frames:
        return pd.DataFrame(columns=[COL_LIGA, COL_DATA, COL_HOME, COL_AWAY, COL_RESULT])
    return pd.concat(frames, ignore_index=True).drop_duplicates(
        subset=[COL_LIGA, COL_DATA, COL_HOME, COL_AWAY]
    )


def _result_blank(value: object) -> bool:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return True
    return str(value).strip() in {"", "nan", "None", "<NA>", "—"}


def _match_keys(df: pd.DataFrame) -> list[tuple[str, str, str, str]]:
    d = pd.to_datetime(df[COL_DATA], dayfirst=True, errors="coerce")
    return list(
        zip(
            df[COL_LIGA].astype(str),
            d.dt.strftime("%Y-%m-%d"),
            df[COL_HOME].map(lambda x: canonical_name(str(x))),
            df[COL_AWAY].map(lambda x: canonical_name(str(x))),
        )
    )


def merge_upcoming(existing: pd.DataFrame, upcoming: pd.DataFrame) -> pd.DataFrame:
    """Dokleja brakujące mecze i uzupełnia puste wyniki FT z BBC."""
    if upcoming is None or upcoming.empty:
        return existing
    if existing.empty:
        return upcoming.copy()

    out = existing.copy()
    have = {k: pos for pos, k in enumerate(_match_keys(out))}
    add_pos: list[int] = []
    up_keys = _match_keys(upcoming)
    for i, key in enumerate(up_keys):
        incoming = upcoming.iloc[i]
        score = incoming.get(COL_RESULT)
        if key in have:
            if _result_blank(score):
                continue
            idx = out.index[have[key]]
            if COL_RESULT not in out.columns or _result_blank(out.at[idx, COL_RESULT]):
                out.at[idx, COL_RESULT] = str(score).strip()
                parsed = re.match(r"^(\d+)\s*[:\-]\s*(\d+)$", str(score).strip())
                if parsed and "оз" in out.columns and _result_blank(out.at[idx, "оз"]):
                    hg, ag = int(parsed.group(1)), int(parsed.group(2))
                    out.at[idx, "оз"] = "так" if hg > 0 and ag > 0 else "ні"
            continue
        add_pos.append(i)
    if not add_pos:
        return out
    add = upcoming.iloc[add_pos].copy()
    for col in out.columns:
        if col not in add.columns:
            add[col] = pd.NA
    return pd.concat([out, add[list(out.columns)]], ignore_index=True)


def upsert_matches(existing: pd.DataFrame, updated: pd.DataFrame) -> pd.DataFrame:
    """Zastępuje wiersze o tym samym kluczu (liga+data+drużyny), resztę zostawia."""
    if updated is None or updated.empty:
        return existing
    if existing is None or existing.empty:
        return updated.copy()
    drop = set(_match_keys(updated))
    keep_mask = [k not in drop for k in _match_keys(existing)]
    keep = existing.loc[keep_mask].copy()
    add = updated.copy()
    for col in keep.columns:
        if col not in add.columns:
            add[col] = pd.NA
    cols = list(keep.columns) if not keep.empty else list(add.columns)
    if keep.empty:
        return add[cols].reset_index(drop=True)
    return pd.concat([keep, add[cols]], ignore_index=True)


def main(argv: list[str] | None = None) -> None:
    parser = argparse.ArgumentParser(description="Nadchodzace mecze FootyStats")
    parser.add_argument("--days", type=int, default=7)
    parser.add_argument("--refresh", action="store_true")
    args = parser.parse_args(argv)
    logging.basicConfig(level=logging.INFO, format="%(levelname)s %(message)s")
    from preview_pandas import read_aleks_table, resolve_aleks_path

    hist = pd.DataFrame()
    src = resolve_aleks_path()
    if src is not None:
        hist = read_aleks_table(src)
    df = fetch_upcoming_fixtures(days=args.days, refresh=args.refresh, history=hist)
    _safe_print(f"Nadchodzace (ligi Aleksa): {len(df)}")
    if not df.empty:
        _safe_print(df.groupby(COL_LIGA).size().to_string())
        _safe_print(df[[COL_DATA, COL_LIGA, COL_HOME, COL_AWAY]].head(20).to_string(index=False))


if __name__ == "__main__":
    main()
