# -*- coding: utf-8 -*-
"""Uzupełnianie luk: JSON (braki) → Serper → BS4 → Claude → walidacja → Excel.

1. Skanuje mecze i zapisuje braki do cache/missing_data.json
2. Serper szuka stron ze statystykami (faule, rożne, kartki, strzały)
3. requests+BeautifulSoup pobiera HTML
4. Claude akceptuje TYLKO liczby z tekstu strony (bez zmyślania)
5. Walidacja → dopisanie do JSON w lukę → zapis Excel
6. Po zapisie Excela: odczyt pliku → puste komórki z JSON, reszta Serper/Claude → JSON → Excel



Klucze w środowisku / .env:
  SERPER_API_KEY
  ANTHROPIC_API_KEY
"""
from __future__ import annotations

import json
import logging
import os
import re
from datetime import datetime
from pathlib import Path
from typing import Any, Callable

import pandas as pd
import requests
from bs4 import BeautifulSoup

from enrich_scores import clean_name
from team_names import TEAM_ALIASES, canonical_name, canonicalize_teams

ROOT = Path(__file__).resolve().parent
CACHE_JSON = ROOT / "cache" / "missing_matches.json"
STATS_JSON = ROOT / "cache" / "missing_data.json"
SERPER_URL = "https://google.serper.dev/search"
MIN_MATCHES = 3
FORM_N = 5
MAX_URLS_PER_TEAM = 4
MAX_URLS_PER_GAP = 4
PAGE_TEXT_LIMIT = 12000

COL_LIGA = "ліга"
COL_DATA = "дата"
COL_HOME = "господар"
COL_AWAY = "гість"
COL_RESULT = "результат"
COL_BTTS = "оз"
COL_BTTS_UA = "Чи обидві забили?"
COL_BTTS_UA_OLD = "Чи обидві заб'ють?"

STAT_BASES = ["фоли", "кутові", "жовті_картки", "удари", "удари_в_площину"]
STAT_COLUMNS = [c for b in STAT_BASES for c in (f"{b}_господар", f"{b}_гість", b)]
SKIP_HOSTS = (
    "facebook.com",
    "instagram.com",
    "twitter.com",
    "x.com",
    "tiktok.com",
    "youtube.com",
    "sportsgambler.com",
    "betdiary.io",
    "bet365.com",
    "oddsportal.com",
)

CLAUDE_TO_COL = {
    "fouls_home": "фоли_господар",
    "fouls_away": "фоли_гість",
    "corners_home": "кутові_господар",
    "corners_away": "кутові_гість",
    "cards_home": "жовті_картки_господар",
    "cards_away": "жовті_картки_гість",
    "yellow_home": "жовті_картки_господар",
    "yellow_away": "жовті_картки_гість",
    "shots_home": "удари_господар",
    "shots_away": "удари_гість",
    "shots_on_home": "удари_в_площину_господар",
    "shots_on_away": "удари_в_площину_гість",
}

UA = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
)

logger = logging.getLogger("fill_missing")

CLAUDE_PROMPT = """
You audit football result pages. You do not invent data.

INPUT: club we need history for, as_of_date, and PAGE TEXT scraped from one URL.

DECIDE:
- match=true only if this page is about THAT senior club's finished competitive results
  (not a namesake, not youth/reserves, not predictions, not odds).
- Extract up to 8 finished matches BEFORE as_of_date.
- Full-time 90-minute scores only.
- If a field is not on the page, use null. Never guess corners/cards/shots.
- If you cannot verify at least the date, both teams and FT score, omit that match.
- Our club in every extracted row must be written as canonical_name.

OUTPUT JSON only:
{
  "match": true,
  "reason": "short",
  "source_url": "...",
  "matches": [
    {
      "date": "dd/mm/yyyy",
      "home": "...",
      "away": "...",
      "home_goals": 0,
      "away_goals": 0,
      "corners_home": null,
      "corners_away": null,
      "cards_home": null,
      "cards_away": null,
      "competition": "Championship"
    }
  ]
}
If the page is wrong, return {"match": false, "reason": "...", "matches": []}.
""".strip()

CLAUDE_STATS_PROMPT = """
You extract football MATCH STATISTICS from PAGE TEXT. You do not invent numbers.

INPUT: one fixture (date, home, away, league, known FT score if any),
the list of MISSING fields, and PAGE TEXT from one URL.

DECIDE:
- match=true only if the page is about THIS exact senior first-team match
  (same date, these two clubs, not youth/reserves, not odds, not a preview).
- Extract ONLY numbers that appear on the page.
- If a field is not on the page, omit it or use null. Never guess.
- Home/away split must match THIS fixture (home is the home club).
- If a known FT score is provided and the page shows a different score, return match=false.

OUTPUT JSON only:
{
  "match": true,
  "reason": "short",
  "home_goals": null,
  "away_goals": null,
  "fouls_home": null,
  "fouls_away": null,
  "corners_home": null,
  "corners_away": null,
  "cards_home": null,
  "cards_away": null,
  "shots_home": null,
  "shots_away": null,
  "shots_on_home": null,
  "shots_on_away": null
}
If the page is wrong or has no usable stats: {"match": false, "reason": "..."}
""".strip()


class ClaudeAuthError(RuntimeError):
    """Nieprawidłowy ANTHROPIC_API_KEY — nie ponawiaj requestów."""


def _load_dotenv() -> None:
    path = ROOT / ".env"
    if not path.exists():
        return
    for raw in path.read_text(encoding="utf-8").splitlines():
        line = raw.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, val = line.split("=", 1)
        key, val = key.strip(), val.strip().strip('"').strip("'")
        if key and key not in os.environ:
            os.environ[key] = val


def _hydrate_user_env() -> None:
    """Windows: dopisz klucze API z User env, jeśli nie ma ich w procesie."""
    names = (
        "SERPER_API_KEY",
        "ANTHROPIC_API_KEY",
        "CLAUDE_API_KEY",
        "ANTHROPIC_AUTH_TOKEN",
        "FILL_MISSING",
    )
    try:
        import winreg
    except ImportError:
        return
    for name in names:
        if (os.environ.get(name) or "").strip():
            continue
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Environment") as key:
                val, _ = winreg.QueryValueEx(key, name)
        except OSError:
            continue
        if val:
            os.environ[name] = str(val).strip()


def _env(name: str) -> str:
    _hydrate_user_env()
    _load_dotenv()
    return (os.environ.get(name) or "").strip()


def _anthropic_key() -> str:
    """Klucz z PowerShella / User env, potem .env. Alias: CLAUDE_API_KEY."""
    _hydrate_user_env()
    for name in ("ANTHROPIC_API_KEY", "CLAUDE_API_KEY", "ANTHROPIC_AUTH_TOKEN"):
        val = (os.environ.get(name) or "").strip().strip('"').strip("'")
        if val:
            return val
    _load_dotenv()
    for name in ("ANTHROPIC_API_KEY", "CLAUDE_API_KEY", "ANTHROPIC_AUTH_TOKEN"):
        val = (os.environ.get(name) or "").strip().strip('"').strip("'")
        if val:
            return val
    return ""


def has_keys() -> bool:
    return bool(_env("SERPER_API_KEY") and _anthropic_key())


def fill_enabled() -> bool:
    if _env("FILL_MISSING").lower() in {"0", "false", "no"}:
        return False
    return has_keys()


def _norm(name: str) -> str:
    return clean_name(canonical_name(name))


def _team_keys(name: str) -> set[str]:
    can = canonical_name(name)
    keys = {_norm(name), _norm(can), clean_name(name)}
    for src, dst in TEAM_ALIASES.items():
        if _norm(dst) == _norm(can) or _norm(src) == _norm(can):
            keys.add(clean_name(src))
            keys.add(_norm(dst))
    return {k for k in keys if k}


def search_query(team: str, league: str, as_of: str) -> str:
    return (
        f"{team} {league} football results 2025 2026 "
        f"before {as_of} Championship Segunda Eerste Divisie 3. Liga"
    )


def thin_teams(
    target: pd.DataFrame,
    history: pd.DataFrame,
    *,
    min_matches: int = MIN_MATCHES,
) -> list[dict[str, Any]]:
    if target is None or target.empty:
        return []
    dates = pd.to_datetime(target[COL_DATA], dayfirst=True, errors="coerce")
    as_of = dates.min()
    hist = history.copy()
    hist["_d"] = pd.to_datetime(hist[COL_DATA], dayfirst=True, errors="coerce") if COL_DATA in hist.columns else pd.NaT
    if pd.notna(as_of) and "_d" in hist.columns:
        hist = hist[hist["_d"] < as_of]
    seen: set[str] = set()
    out: list[dict[str, Any]] = []
    for _, row in target.iterrows():
        liga = str(row.get(COL_LIGA) or "")
        for side in (COL_HOME, COL_AWAY):
            raw = str(row.get(side) or "").strip()
            if not raw:
                continue
            can = canonical_name(raw)
            key = _norm(can)
            if key in seen:
                continue
            seen.add(key)
            keys = _team_keys(can)
            if hist.empty:
                n = 0
            else:
                h = hist[COL_HOME].map(lambda x, ks=keys: _norm(str(x)) in ks)
                a = hist[COL_AWAY].map(lambda x, ks=keys: _norm(str(x)) in ks)
                n = int((h | a).sum())
            if n < min_matches:
                out.append(
                    {
                        "team": can,
                        "query_name": raw,
                        "league": liga,
                        "n": n,
                        "need": max(0, FORM_N - n),
                        "as_of": pd.Timestamp(as_of).strftime("%d/%m/%Y") if pd.notna(as_of) else "",
                    }
                )
    return out


def serper_search(query: str, *, num: int = 8) -> list[dict[str, str]]:
    key = _env("SERPER_API_KEY")
    if not key:
        raise RuntimeError("Brak SERPER_API_KEY")
    r = requests.post(
        SERPER_URL,
        headers={"X-API-KEY": key, "Content-Type": "application/json"},
        json={"q": query, "num": num},
        timeout=30,
    )
    r.raise_for_status()
    organic = r.json().get("organic") or []
    hits: list[dict[str, str]] = []
    for item in organic:
        link = str(item.get("link") or "").strip()
        if not link.startswith("http"):
            continue
        if any(h in link.lower() for h in SKIP_HOSTS):
            continue
        hits.append(
            {
                "title": str(item.get("title") or ""),
                "link": link,
                "snippet": str(item.get("snippet") or ""),
            }
        )
    return hits


def html_to_text(html: str) -> str:
    try:
        soup = BeautifulSoup(html or "", "lxml")
    except Exception:
        soup = BeautifulSoup(html or "", "html.parser")
    for tag in soup(["script", "style", "noscript", "svg", "iframe"]):
        tag.decompose()
    text = soup.get_text(" ", strip=True)
    text = re.sub(r"\s+", " ", text)
    return text[:PAGE_TEXT_LIMIT]


def fetch_page(url: str, *, timeout: int = 25) -> str:
    r = requests.get(url, timeout=timeout, headers={"User-Agent": UA})
    r.raise_for_status()
    r.encoding = r.apparent_encoding or "utf-8"
    return html_to_text(r.text)


def extract_json(text: str) -> dict[str, Any] | None:
    if not (text or "").strip():
        return None
    raw = text.strip()
    fence = re.search(r"```(?:json)?\s*([\s\S]*?)```", raw)
    if fence:
        raw = fence.group(1).strip()
    try:
        start, end = raw.index("{"), raw.rindex("}") + 1
        return json.loads(raw[start:end])
    except (ValueError, json.JSONDecodeError):
        return None


def _is_auth_error(exc: BaseException) -> bool:
    msg = str(exc).lower()
    return "401" in msg or "authentication_error" in msg or "api key is invalid" in msg


def _claude_message(*, system: str, user: str, max_tokens: int = 2500) -> dict[str, Any]:
    key = _anthropic_key()
    if not key:
        raise RuntimeError("Brak ANTHROPIC_API_KEY w PowerShell (albo CLAUDE_API_KEY)")
    try:
        import anthropic
    except ImportError as exc:
        raise RuntimeError("Zainstaluj: pip install anthropic") from exc
    model = _env("ANTHROPIC_MODEL") or "claude-sonnet-4-6"
    client = anthropic.Anthropic(api_key=key)
    try:
        msg = client.messages.create(
            model=model,
            max_tokens=max_tokens,
            system=system,
            messages=[{"role": "user", "content": user}],
        )
    except Exception as exc:
        if _is_auth_error(exc):
            raise ClaudeAuthError(
                "PowerShell ANTHROPIC_API_KEY jest odrzucony przez Anthropic (401 API key is invalid). "
                "Klucz jest odczytywany, ale serwer go nie akceptuje — wygeneruj nowy w console.anthropic.com "
                "i ustaw: $env:ANTHROPIC_API_KEY = 'sk-ant-...'"
            ) from exc
        raise
    parts = [getattr(b, "text", "") for b in (msg.content or [])]
    return extract_json("\n".join(parts)) or {"match": False, "reason": "claude_json_fail"}


def claude_decide(
    *,
    team: str,
    league: str,
    as_of: str,
    url: str,
    page_text: str,
) -> dict[str, Any]:
    user = (
        f"canonical_name: {canonical_name(team)}\n"
        f"query_name: {team}\n"
        f"league_context: {league}\n"
        f"as_of_date: {as_of}\n"
        f"source_url: {url}\n\n"
        f"PAGE TEXT:\n{page_text}"
    )
    parsed = _claude_message(system=CLAUDE_PROMPT, user=user)
    parsed["source_url"] = url
    if "matches" not in parsed:
        parsed["matches"] = []
    return parsed


def _num(v: Any) -> int | None:
    if v is None or v == "":
        return None
    try:
        if pd.isna(v):
            return None
    except (TypeError, ValueError):
        pass
    try:
        n = int(float(v))
    except (TypeError, ValueError):
        return None
    return n if n >= 0 else None


def validate_match(
    raw: dict[str, Any],
    *,
    team: str,
    as_of: pd.Timestamp | None,
) -> dict[str, Any] | None:
    """Odrzuca zmyślone / niepełne wiersze. Zwraca znormalizowany mecz albo None."""
    d = pd.to_datetime(str(raw.get("date") or ""), dayfirst=True, errors="coerce")
    if pd.isna(d):
        return None
    if as_of is not None and d >= as_of:
        return None
    hg, ag = _num(raw.get("home_goals")), _num(raw.get("away_goals"))
    home, away = str(raw.get("home") or "").strip(), str(raw.get("away") or "").strip()
    if hg is None or ag is None or not home or not away:
        return None
    keys = _team_keys(team)
    can = canonical_name(team)
    if _norm(home) not in keys and _norm(away) not in keys:
        return None
    if _norm(home) in keys:
        home = can
    if _norm(away) in keys:
        away = can
    if home == away:
        return None
    return {
        "date": d.strftime("%d/%m/%Y"),
        "home": home,
        "away": away,
        "home_goals": hg,
        "away_goals": ag,
        "corners_home": _num(raw.get("corners_home")),
        "corners_away": _num(raw.get("corners_away")),
        "cards_home": _num(raw.get("cards_home")),
        "cards_away": _num(raw.get("cards_away")),
        "competition": str(raw.get("competition") or "").strip() or "uzupelnienie",
        "source_url": str(raw.get("source_url") or ""),
    }


def matches_to_aleks(matches: list[dict[str, Any]]) -> pd.DataFrame:
    rows: list[dict[str, Any]] = []
    for m in matches:
        hg, ag = int(m["home_goals"]), int(m["away_goals"])
        rec: dict[str, Any] = {
            COL_LIGA: m.get("competition") or "uzupelnienie",
            COL_DATA: m["date"],
            COL_HOME: m["home"],
            COL_AWAY: m["away"],
            COL_RESULT: f"{hg}:{ag}",
            "оз": "так" if hg > 0 and ag > 0 else "ні",
        }
        ch, ca = m.get("corners_home"), m.get("corners_away")
        yh, ya = m.get("cards_home"), m.get("cards_away")
        rec["кутові_господар"] = ch if ch is not None else pd.NA
        rec["кутові_гість"] = ca if ca is not None else pd.NA
        rec["кутові"] = (ch + ca) if ch is not None and ca is not None else pd.NA
        rec["жовті_картки_господар"] = yh if yh is not None else pd.NA
        rec["жовті_картки_гість"] = ya if ya is not None else pd.NA
        rec["жовті_картки"] = (yh + ya) if yh is not None and ya is not None else pd.NA
        for col in ("фоли", "удари", "удари_в_площину"):
            rec[f"{col}_господар"] = pd.NA
            rec[f"{col}_гість"] = pd.NA
            rec[col] = pd.NA
        rows.append(rec)
    if not rows:
        return pd.DataFrame()
    return pd.DataFrame(rows).drop_duplicates(subset=[COL_DATA, COL_HOME, COL_AWAY])


def save_missing_json(payload: dict[str, Any], path: Path | None = None) -> Path:
    out = path or CACHE_JSON
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    return out


def load_missing_json(path: Path | None = None) -> dict[str, Any]:
    p = path or CACHE_JSON
    if not p.exists():
        return {}
    try:
        return json.loads(p.read_text(encoding="utf-8"))
    except json.JSONDecodeError:
        return {}


def merge_history(history: pd.DataFrame, extra: pd.DataFrame) -> pd.DataFrame:
    if extra is None or extra.empty:
        return history
    hist = history.copy()
    extra = extra.copy()

    def keys(df: pd.DataFrame) -> pd.Series:
        d = pd.to_datetime(df[COL_DATA], dayfirst=True, errors="coerce")
        return (
            d.dt.strftime("%Y-%m-%d")
            + "|"
            + df[COL_HOME].map(lambda x: _norm(str(x)))
            + "|"
            + df[COL_AWAY].map(lambda x: _norm(str(x)))
        )

    hist["_k"] = keys(hist)
    extra["_k"] = keys(extra)
    add = extra[~extra["_k"].isin(set(hist["_k"].dropna()))].drop(columns=["_k"])
    hist = hist.drop(columns=["_k"])
    if add.empty:
        return hist
    for c in hist.columns:
        if c not in add.columns:
            add[c] = pd.NA
    return pd.concat([hist, add[list(hist.columns)]], ignore_index=True)


def fill_gaps(
    target: pd.DataFrame,
    history: pd.DataFrame,
    *,
    search_fn: Callable[[str], list[dict[str, str]]] | None = None,
    fetch_fn: Callable[[str], str] | None = None,
    decide_fn: Callable[..., dict[str, Any]] | None = None,
    cache_path: Path | None = None,
) -> tuple[pd.DataFrame, dict[str, Any]]:
    """Serper → strona → Claude → walidacja. Zwraca (wiersze Aleksa, raport JSON)."""
    teams = thin_teams(target, history)
    report: dict[str, Any] = {
        "updated_at": datetime.now().isoformat(timespec="seconds"),
        "teams": [],
        "accepted_matches": 0,
        "rejected_pages": 0,
    }
    if not teams:
        save_missing_json(report, cache_path)
        return pd.DataFrame(), report

    search_fn = search_fn or serper_search
    fetch_fn = fetch_fn or fetch_page
    decide_fn = decide_fn or (
        lambda **kw: claude_decide(
            team=kw["team"],
            league=kw["league"],
            as_of=kw["as_of"],
            url=kw["url"],
            page_text=kw["page_text"],
        )
    )

    accepted: list[dict[str, Any]] = []
    for t in teams:
        as_of = pd.to_datetime(t["as_of"], dayfirst=True, errors="coerce")
        query = search_query(t["query_name"] or t["team"], t["league"], t["as_of"])
        entry: dict[str, Any] = {
            "team": t["team"],
            "league": t["league"],
            "n_history": t["n"],
            "query": query,
            "pages": [],
            "matches": [],
        }
        try:
            hits = search_fn(query)[:MAX_URLS_PER_TEAM]
        except Exception as exc:
            logger.warning("Serper fail %s: %s", t["team"], exc)
            entry["error"] = str(exc)
            report["teams"].append(entry)
            continue
        for hit in hits:
            url = hit["link"]
            page_info = {"url": url, "title": hit.get("title"), "accepted": False}
            try:
                text = fetch_fn(url)
            except Exception as exc:
                page_info["error"] = f"fetch: {exc}"
                report["rejected_pages"] += 1
                entry["pages"].append(page_info)
                continue
            try:
                decided = decide_fn(
                    team=t["team"],
                    league=t["league"],
                    as_of=t["as_of"],
                    url=url,
                    page_text=text,
                )
            except ClaudeAuthError:
                raise
            except Exception as exc:
                page_info["error"] = f"claude: {exc}"
                report["rejected_pages"] += 1
                entry["pages"].append(page_info)
                continue
            if not decided.get("match"):
                page_info["reason"] = decided.get("reason") or "no_match"
                report["rejected_pages"] += 1
                entry["pages"].append(page_info)
                continue
            page_info["accepted"] = True
            page_info["reason"] = decided.get("reason") or "ok"
            got = 0
            for raw in decided.get("matches") or []:
                raw = dict(raw)
                raw["source_url"] = url
                ok = validate_match(raw, team=t["team"], as_of=as_of)
                if ok:
                    accepted.append(ok)
                    entry["matches"].append(ok)
                    got += 1
            if got == 0:
                page_info["accepted"] = False
                page_info["reason"] = "validation_dropped_all"
                report["rejected_pages"] += 1
            entry["pages"].append(page_info)
            if len(entry["matches"]) >= FORM_N:
                break
        report["teams"].append(entry)

    report["accepted_matches"] = len(accepted)
    save_missing_json(report, cache_path)
    extra = matches_to_aleks(accepted)
    if not extra.empty:
        extra = canonicalize_teams(extra)
    return extra, report


def _is_blank(v: Any) -> bool:
    if v is None:
        return True
    try:
        if pd.isna(v):
            return True
    except (TypeError, ValueError):
        pass
    s = str(v).strip().lower()
    return s in {"", "nan", "none", "<na>"}


def _fmt_date(v: Any) -> str:
    d = pd.to_datetime(v, dayfirst=True, errors="coerce")
    if pd.isna(d):
        return str(v or "").strip()
    return pd.Timestamp(d).strftime("%d/%m/%Y")


def row_key(date: Any, home: Any, away: Any) -> str:
    d = pd.to_datetime(date, dayfirst=True, errors="coerce")
    ds = pd.Timestamp(d).strftime("%Y-%m-%d") if pd.notna(d) else str(date or "")
    return f"{ds}|{_norm(str(home or ''))}|{_norm(str(away or ''))}"


def missing_fields(row: pd.Series) -> list[str]:
    out: list[str] = []
    if COL_RESULT in row.index and _is_blank(row.get(COL_RESULT)):
        out.append(COL_RESULT)
    for col in STAT_COLUMNS:
        if col in row.index and _is_blank(row.get(col)):
            out.append(col)
    return out


def scan_missing(
    df: pd.DataFrame,
    *,
    as_of: pd.Timestamp | None = None,
    from_date: pd.Timestamp | None = None,
) -> list[dict[str, Any]]:
    """Braki tylko w meczach już rozegranych (wynik albo data < as_of)."""
    if df is None or df.empty:
        return []
    cut = (as_of or pd.Timestamp.now()).normalize()
    start = pd.Timestamp(from_date).normalize() if from_date is not None else None
    gaps: list[dict[str, Any]] = []
    for _, row in df.iterrows():
        date = pd.to_datetime(row.get(COL_DATA), dayfirst=True, errors="coerce")
        if start is not None and pd.notna(date) and date.normalize() < start:
            continue
        has_score = not _is_blank(row.get(COL_RESULT))
        played = has_score or (pd.notna(date) and date.normalize() < cut)
        if not played:
            continue
        miss = missing_fields(row)
        if not miss:
            continue
        home, away = str(row.get(COL_HOME) or "").strip(), str(row.get(COL_AWAY) or "").strip()
        liga = str(row.get(COL_LIGA) or "").strip()
        date_s = _fmt_date(date)
        query = (
            f"{home} vs {away} {date_s} {liga} football match statistics "
            f"corners fouls yellow cards shots"
        )
        gaps.append(
            {
                "key": row_key(date, home, away),
                "league": liga,
                "date": date_s,
                "home": home,
                "away": away,
                "result": "" if _is_blank(row.get(COL_RESULT)) else str(row.get(COL_RESULT)).strip(),
                "missing": miss,
                "status": "pending",
                "filled": {},
                "query": query,
                "pages": [],
                "source_url": None,
                "reason": None,
            }
        )
    return gaps


def _merge_inventories(
    scanned: list[dict[str, Any]],
    previous: dict[str, Any] | None,
) -> list[dict[str, Any]]:
    """Scala nowy skan z JSON: nowe braki + stare wypełnienia (nawet gdy Excel jest kompletny)."""
    old_by_key = {g.get("key"): g for g in (previous or {}).get("gaps") or [] if g.get("key")}
    merged: list[dict[str, Any]] = []
    seen: set[str] = set()
    for gap in scanned:
        key = gap["key"]
        seen.add(key)
        prev = old_by_key.get(key)
        if not prev:
            merged.append(gap)
            continue
        keep = dict(gap)
        keep["pages"] = list(prev.get("pages") or [])
        keep["source_url"] = prev.get("source_url")
        keep["reason"] = prev.get("reason")
        filled = dict(prev.get("filled") or {})
        still = [c for c in gap["missing"] if c not in filled]
        keep["filled"] = filled
        keep["missing"] = still
        if still:
            keep["status"] = "pending"
        else:
            keep["status"] = "filled"
        merged.append(keep)
    for key, prev in old_by_key.items():
        if key not in seen:
            merged.append(prev)
    return merged


def inventory_summary(gaps: list[dict[str, Any]], *, rows: int, future_skipped: int) -> dict[str, Any]:
    pending = sum(1 for g in gaps if g.get("status") == "pending")
    filled = sum(1 for g in gaps if g.get("status") == "filled")
    return {
        "excel_rows": rows,
        "gaps": len(gaps),
        "pending": pending,
        "filled": filled,
        "skipped_future": future_skipped,
    }


def build_inventory(
    df: pd.DataFrame,
    *,
    as_of: pd.Timestamp | None = None,
    from_date: pd.Timestamp | None = None,
    path: Path | None = None,
) -> dict[str, Any]:
    """Najpierw JSON: skanuje braki, scala z poprzednim plikiem, zapisuje."""
    cut = (as_of or pd.Timestamp.now()).normalize()
    dates = pd.to_datetime(df[COL_DATA], dayfirst=True, errors="coerce") if not df.empty else pd.Series(dtype="datetime64[ns]")
    future = int(((dates.dt.normalize() >= cut) & df[COL_RESULT].map(_is_blank)).sum()) if not df.empty else 0
    scanned = scan_missing(df, as_of=cut, from_date=from_date)
    prev = load_missing_json(path or STATS_JSON)
    gaps = _merge_inventories(scanned, prev)
    report = {
        "updated_at": datetime.now().isoformat(timespec="seconds"),
        "as_of": cut.strftime("%d/%m/%Y"),
        "summary": inventory_summary(gaps, rows=len(df), future_skipped=future),
        "gaps": gaps,
    }
    save_missing_json(report, path or STATS_JSON)
    return report


def stats_search_query(gap: dict[str, Any]) -> str:
    return stats_search_queries(gap)[0]


def stats_search_queries(gap: dict[str, Any], *, round_n: int = 0) -> list[str]:
    home = str(gap.get("home") or "")
    away = str(gap.get("away") or "")
    date = str(gap.get("date") or "")
    liga = str(gap.get("league") or "")
    miss = " ".join(gap.get("missing") or [])
    queries = [
        f"{home} vs {away} {date} {liga} football match statistics corners fouls yellow cards shots",
    ]
    if "кутові" in miss or "фоли" in miss:
        queries.extend(
            [
                f"{home} vs {away} {date} sofascore match statistics fouls corners",
                f"{home} vs {away} {date} foxsports boxscore fouls corners",
                f"{home} vs {away} {date} {liga} flashscore statistics corners fouls",
            ]
        )
    if "удари_в_площину" in miss:
        queries.append(f"{home} vs {away} {date} shots on target sofascore foxsports")
    if round_n >= 1:
        queries.extend(
            [
                f"{home} vs {away} {date} valuestats match statistics",
                f"{home} vs {away} {date} soccerway match stats",
                f"{home} vs {away} {date} {liga} team stats boxscore",
            ]
        )
    seen: set[str] = set()
    out: list[str] = []
    for q in queries:
        if q not in seen:
            seen.add(q)
            out.append(q)
    return out


def claude_decide_stats(
    *,
    gap: dict[str, Any],
    url: str,
    page_text: str,
) -> dict[str, Any]:
    user = (
        f"date: {gap.get('date')}\n"
        f"home: {gap.get('home')}\n"
        f"away: {gap.get('away')}\n"
        f"league: {gap.get('league')}\n"
        f"known_ft_score: {gap.get('result') or 'unknown'}\n"
        f"missing_fields: {gap.get('missing')}\n"
        f"source_url: {url}\n\n"
        f"PAGE TEXT:\n{page_text}"
    )
    parsed = _claude_message(system=CLAUDE_STATS_PROMPT, user=user, max_tokens=1200)
    parsed["source_url"] = url
    return parsed


def _score_tuple(text: str) -> tuple[int, int] | None:
    m = re.match(r"^(\d+)\s*[:\-]\s*(\d+)$", str(text or "").strip())
    if not m:
        return None
    return int(m.group(1)), int(m.group(2))


def validate_stats(
    raw: dict[str, Any],
    gap: dict[str, Any],
) -> dict[str, Any] | None:
    """Odrzuca zmyślone liczby. Zwraca mapę kolumn Aleksa albo None."""
    if not raw.get("match"):
        return None
    known = _score_tuple(str(gap.get("result") or ""))
    hg, ag = _num(raw.get("home_goals")), _num(raw.get("away_goals"))
    if known is not None and hg is not None and ag is not None and (hg, ag) != known:
        return None
    filled: dict[str, Any] = {}
    missing = set(gap.get("missing") or [])
    if COL_RESULT in missing and hg is not None and ag is not None:
        filled[COL_RESULT] = f"{hg}:{ag}"
        filled[COL_BTTS] = "так" if hg > 0 and ag > 0 else "ні"
    for src, col in CLAUDE_TO_COL.items():
        if col not in missing:
            continue
        n = _num(raw.get(src))
        if n is not None:
            filled[col] = n
    existing = dict(gap.get("filled") or {})
    return _complete_totals(filled, existing, missing) or None


def _complete_totals(
    filled: dict[str, Any],
    existing: dict[str, Any],
    missing: set[str] | list[str],
) -> dict[str, Any]:
    miss = set(missing)
    merged = {**existing, **filled}
    for base in STAT_BASES:
        hcol, acol, tcol = f"{base}_господар", f"{base}_гість", base
        hv, av = merged.get(hcol), merged.get(acol)
        if hv is not None and av is not None and tcol in miss and tcol not in filled:
            try:
                filled[tcol] = int(hv) + int(av)
            except (TypeError, ValueError):
                pass
    return filled


def apply_inventory(df: pd.DataFrame, inventory: dict[str, Any]) -> pd.DataFrame:
    out = df.copy()
    index: dict[str, list[int]] = {}
    for i, row in out.iterrows():
        index.setdefault(row_key(row.get(COL_DATA), row.get(COL_HOME), row.get(COL_AWAY)), []).append(i)
    for gap in inventory.get("gaps") or []:
        filled = gap.get("filled") or {}
        if not filled:
            continue
        for idx in index.get(gap.get("key") or "", []):
            for col, val in filled.items():
                if col not in out.columns:
                    continue
                if _is_blank(out.at[idx, col]):
                    out.at[idx, col] = val
    return out


def fill_stats(
    df: pd.DataFrame,
    *,
    as_of: pd.Timestamp | None = None,
    from_date: pd.Timestamp | None = None,
    search_fn: Callable[[str], list[dict[str, str]]] | None = None,
    fetch_fn: Callable[[str], str] | None = None,
    decide_fn: Callable[..., dict[str, Any]] | None = None,
    cache_path: Path | None = None,
    live: bool = True,
    search_round: int = 0,
) -> tuple[pd.DataFrame, dict[str, Any]]:
    """JSON braków → Serper/Claude (opcjonalnie) → walidacja → DataFrame."""
    path = cache_path or STATS_JSON
    inventory = build_inventory(df, as_of=as_of, from_date=from_date, path=path)
    if not live:
        return apply_inventory(df, inventory), inventory

    search_fn = search_fn or serper_search
    fetch_fn = fetch_fn or fetch_page
    decide_fn = decide_fn or (
        lambda **kw: claude_decide_stats(gap=kw["gap"], url=kw["url"], page_text=kw["page_text"])
    )

    for gap in inventory["gaps"]:
        if not gap.get("missing"):
            continue
        extra = _complete_totals({}, gap.get("filled") or {}, set(gap["missing"]))
        if extra:
            gap["filled"].update(extra)
            gap["missing"] = [c for c in gap["missing"] if c not in gap["filled"]]
            if not gap["missing"]:
                gap["status"] = "filled"
                continue
        seen_urls = {str(p.get("url") or "") for p in gap.get("pages") or []}
        queries = stats_search_queries(gap, round_n=search_round)
        gap["query"] = queries[0]
        hits: list[dict[str, str]] = []
        try:
            for q in queries:
                for hit in search_fn(q):
                    url = hit.get("link") or ""
                    if not url or url in seen_urls:
                        continue
                    seen_urls.add(url)
                    hits.append(hit)
                    if "foxsports.com" in url and "boxscore" in url and "tab=" not in url:
                        extra = url + ("&" if "?" in url else "?") + "tab=boxscore"
                        if extra not in seen_urls:
                            seen_urls.add(extra)
                            hits.append({"title": hit.get("title") or "", "link": extra, "snippet": hit.get("snippet") or ""})
                    if len(hits) >= MAX_URLS_PER_GAP + 4:
                        break
                if len(hits) >= MAX_URLS_PER_GAP + 4:
                    break
        except ClaudeAuthError:
            save_missing_json(inventory, path)
            raise
        except Exception as exc:
            logger.warning("Serper fail %s: %s", gap["key"], exc)
            gap["reason"] = f"serper: {exc}"
            save_missing_json(inventory, path)
            continue
        for hit in hits:
            url = hit.get("link") or ""
            page_info = {"url": url, "title": hit.get("title"), "accepted": False}
            try:
                text = fetch_fn(url)
            except Exception as exc:
                page_info["error"] = f"fetch: {exc}"
                gap["pages"].append(page_info)
                continue
            try:
                decided = decide_fn(gap=gap, url=url, page_text=text)
            except ClaudeAuthError:
                gap["pages"].append(page_info | {"error": "claude: 401"})
                save_missing_json(inventory, path)
                raise
            except Exception as exc:
                page_info["error"] = f"claude: {exc}"
                gap["pages"].append(page_info)
                continue
            if not decided.get("match"):
                page_info["reason"] = decided.get("reason") or "no_match"
                gap["pages"].append(page_info)
                continue
            ok = validate_stats(decided, gap)
            if not ok:
                page_info["reason"] = "validation_dropped"
                gap["pages"].append(page_info)
                continue
            page_info["accepted"] = True
            page_info["reason"] = decided.get("reason") or "ok"
            gap["pages"].append(page_info)
            gap["filled"].update(ok)
            gap["missing"] = [c for c in gap["missing"] if c not in gap["filled"]]
            gap["source_url"] = url
            gap["reason"] = page_info["reason"]
            if not gap["missing"]:
                gap["status"] = "filled"
                break
        if gap["filled"] and gap["missing"]:
            gap["status"] = "partial"
        save_missing_json(inventory, path)

    inventory["updated_at"] = datetime.now().isoformat(timespec="seconds")
    inventory["summary"] = inventory_summary(
        inventory["gaps"],
        rows=len(df),
        future_skipped=int((inventory.get("summary") or {}).get("skipped_future") or 0),
    )
    save_missing_json(inventory, path)
    return apply_inventory(df, inventory), inventory


def complete_row_totals(df: pd.DataFrame) -> pd.DataFrame:
    """Uzupełnia sumy (faule/rożne/kartki/strzały) z home+away, bez API."""
    if df is None or df.empty:
        return df
    out = df.copy()
    for base in STAT_BASES:
        hcol, acol, tcol = f"{base}_господар", f"{base}_гість", base
        if not all(c in out.columns for c in (hcol, acol, tcol)):
            continue
        h = pd.to_numeric(out[hcol], errors="coerce")
        a = pd.to_numeric(out[acol], errors="coerce")
        need = out[tcol].map(_is_blank) & h.notna() & a.notna()
        if need.any():
            out.loc[need, tcol] = (h + a).loc[need]
    return out


def audit_missing(
    df: pd.DataFrame,
    *,
    as_of: pd.Timestamp | None = None,
    from_date: pd.Timestamp | None = None,
) -> dict[str, Any]:
    """Skanuje rozegrane mecze (opcjonalnie od from_date) pod kątem pustych pól."""
    checked = complete_row_totals(df)
    gaps = scan_missing(checked, as_of=as_of, from_date=from_date)
    return {
        "matches": len(gaps),
        "fields": int(sum(len(g.get("missing") or []) for g in gaps)),
        "gaps": [
            {
                "date": g.get("date"),
                "home": g.get("home"),
                "away": g.get("away"),
                "league": g.get("league"),
                "missing": g.get("missing"),
            }
            for g in gaps
        ],
    }


def verify_and_fill(
    df: pd.DataFrame,
    *,
    as_of: pd.Timestamp | None = None,
    from_date: pd.Timestamp | None = None,
    search_fn: Callable[[str], list[dict[str, str]]] | None = None,
    fetch_fn: Callable[[str], str] | None = None,
    decide_fn: Callable[..., dict[str, Any]] | None = None,
    cache_path: Path | None = None,
    live: bool = True,
    max_rounds: int = 3,
) -> tuple[pd.DataFrame, dict[str, Any]]:
    """Weryfikuje braki i dociąga je w rundach, aż nie ma postępu."""
    path = cache_path or STATS_JSON
    out = complete_row_totals(df)
    inventory: dict[str, Any] = load_missing_json(path)
    out = apply_inventory(out, inventory)
    out = complete_row_totals(out)
    last_fields: int | None = None
    used_rounds = 0
    for rnd in range(max(1, max_rounds)):
        audit = audit_missing(out, as_of=as_of, from_date=from_date)
        n = int(audit["fields"])
        if n == 0:
            used_rounds = rnd
            break
        if last_fields is not None and n >= last_fields:
            used_rounds = rnd
            break
        last_fields = n
        used_rounds = rnd + 1
        out, inventory = fill_stats(
            out,
            as_of=as_of,
            from_date=from_date,
            search_fn=search_fn,
            fetch_fn=fetch_fn,
            decide_fn=decide_fn,
            cache_path=path,
            live=live,
            search_round=rnd,
        )
        out = complete_row_totals(out)
        if not live:
            break
    out = complete_row_totals(out)
    audit = audit_missing(out, as_of=as_of, from_date=from_date)
    if not inventory:
        inventory = build_inventory(out, as_of=as_of, from_date=from_date, path=path)
    inventory["verification"] = {
        "updated_at": datetime.now().isoformat(timespec="seconds"),
        "rounds": used_rounds,
        "remaining_matches": audit["matches"],
        "remaining_fields": audit["fields"],
        "remaining": audit["gaps"],
    }
    save_missing_json(inventory, path)
    return out, inventory


def read_excel_mecze(path: Path) -> tuple[pd.DataFrame, pd.DataFrame]:
    xlsx = Path(path)
    xl = pd.ExcelFile(xlsx)
    mecze_name = next((n for n in ("Матчі_2026", "Mecze_2026") if n in xl.sheet_names), None)
    pred_name = next((n for n in ("Прогнози", "Predykcje") if n in xl.sheet_names), None)
    if not mecze_name:
        raise ValueError(f"Brak arkusza meczow w {xlsx.name}: {xl.sheet_names}")
    mecze = pd.read_excel(xlsx, sheet_name=mecze_name)
    preds = pd.read_excel(xlsx, sheet_name=pred_name) if pred_name else pd.DataFrame()
    for alias in (COL_BTTS_UA, COL_BTTS_UA_OLD):
        if alias in mecze.columns and COL_BTTS not in mecze.columns:
            mecze = mecze.rename(columns={alias: COL_BTTS})
        if not preds.empty and alias in preds.columns and COL_BTTS not in preds.columns:
            preds = preds.rename(columns={alias: COL_BTTS})
    return mecze, preds


def verify_exported_workbook(
    xlsx_path: Path,
    *,
    json_path: Path | None = None,
    as_of: pd.Timestamp | None = None,
    from_date: pd.Timestamp | None = None,
    search_fn: Callable[[str], list[dict[str, str]]] | None = None,
    fetch_fn: Callable[[str], str] | None = None,
    decide_fn: Callable[..., dict[str, Any]] | None = None,
    live: bool = True,
    max_rounds: int = 2,
) -> tuple[pd.DataFrame, pd.DataFrame, dict[str, Any]]:
    """Po pipeline: Excel → puste komórki → JSON, a gdy JSON też pusty → Serper/Claude.

    1. Czyta Mecze_2026 z zapisanego xlsx
    2. Puste pola najpierw z cache/missing_data.json
    3. Reszta: Serper + HTML + Claude + walidacja → JSON → DataFrame
    Braki skanowane od `from_date` (okno predykcji), nie cały sezon.
    """
    cache = json_path or STATS_JSON
    mecze, preds = read_excel_mecze(xlsx_path)
    mecze = complete_row_totals(mecze)
    before = audit_missing(mecze, as_of=as_of, from_date=from_date)
    report: dict[str, Any] = {
        "excel": str(xlsx_path),
        "empty_before": {"matches": before["matches"], "fields": before["fields"]},
        "filled_from_json": 0,
        "filled_from_api": 0,
        "empty_after": {"matches": before["matches"], "fields": before["fields"]},
        "remaining": before["gaps"],
    }
    if before["fields"] == 0:
        return mecze, preds, report

    inv = load_missing_json(cache)
    mecze = apply_inventory(mecze, inv)
    mecze = complete_row_totals(mecze)
    after_json = audit_missing(mecze, as_of=as_of, from_date=from_date)
    report["filled_from_json"] = int(before["fields"] - after_json["fields"])

    after = after_json
    if after_json["fields"] > 0 and live:
        mecze, inv = verify_and_fill(
            mecze,
            as_of=as_of,
            from_date=from_date,
            search_fn=search_fn,
            fetch_fn=fetch_fn,
            decide_fn=decide_fn,
            cache_path=cache,
            live=True,
            max_rounds=max_rounds,
        )
        mecze = complete_row_totals(mecze)
        after = audit_missing(mecze, as_of=as_of, from_date=from_date)
        report["filled_from_api"] = int(after_json["fields"] - after["fields"])

    report["empty_after"] = {"matches": after["matches"], "fields": after["fields"]}
    report["remaining"] = after["gaps"]
    return mecze, preds, report
