# -*- coding: utf-8 -*-
"""Wzbogaca aleks_ligi_stats.xlsx o wynik meczu i statystyki per drużyna.

Źródło: https://www.football-data.co.uk/
  - wynik: FTHG/FTAG (classic) lub HG/AG (new)
  - statystyki per drużyna (classic): HF/AF, HC/AC, HY/AY, HS/AS, HST/AST

Kolumny per mecz (ukraińskie):
  фоли_господар, фоли_гість, фоли (suma)
  кутові_господар, кутові_гість, кутові
  ... itd.

Uruchomienie:
  python enrich_scores.py
"""
from __future__ import annotations

import logging
import re
import sys
import unicodedata
from difflib import SequenceMatcher
from io import StringIO
from pathlib import Path
from typing import Any

import pandas as pd
import requests

from preview_pandas import ALEKS_XLSX, read_aleks_table, resolve_aleks_path, save_aleks_table

ROOT = Path(__file__).resolve().parent
CACHE_DIR = ROOT / "cache" / "football_data"

logger = logging.getLogger("enrich_scores")

COL_LIGA = "ліга"
COL_DATA = "дата"
COL_HOME = "господар"
COL_AWAY = "гість"
COL_RESULT = "результат"

HOME_SUFFIX = "_господар"
AWAY_SUFFIX = "_гість"

# mapowanie football-data → baza kolumny UA
STAT_MAP: list[tuple[str, str, str]] = [
    ("фоли", "hf", "af"),
    ("кутові", "hc", "ac"),
    ("жовті_картки", "hy", "ay"),
    ("удари", "hs", "as_col"),  # AS to away shots — rename in fetch
    ("удари_в_площину", "hst", "ast"),
]

SEASONS = ("2324", "2425", "2526", "2627")

CLASSIC: dict[str, str] = {
    "Premier League": "E0",
    "La Liga": "SP1",
    "Serie A": "I1",
    "Bundesliga": "D1",
    "Bundesliga 2": "D2",
    "Eredivisie": "N1",
}

NEW_FORMAT: dict[str, tuple[str, str]] = {
    "Allsvenskan": ("SWE", "Allsvenskan"),
    "Eliteserien": ("NOR", "Eliteserien"),
    "Super League": ("SWZ", "Super League"),
}

UA = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
)


def all_csv_columns() -> list[str]:
    cols = [COL_LIGA, COL_DATA, COL_HOME, COL_AWAY, COL_RESULT, "оз"]
    for base, _, _ in STAT_MAP:
        cols.extend([f"{base}{HOME_SUFFIX}", f"{base}{AWAY_SUFFIX}", base])
    return cols


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


def clean_name(value: object) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    text = str(value).strip()
    text = unicodedata.normalize("NFKD", text)
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    out = []
    for ch in text:
        if ch.isascii() and (ch.isalnum() or ch in " ._/-:"):
            out.append(ch)
    return re.sub(r"\s+", " ", "".join(out)).strip().casefold()


def _similarity(a: str, b: str) -> float:
    if not a or not b:
        return 0.0
    if a == b:
        return 1.0
    if a in b or b in a:
        return 0.92
    return SequenceMatcher(None, a, b).ratio()


def _download(url: str, dest: Path) -> Path | None:
    dest.parent.mkdir(parents=True, exist_ok=True)
    try:
        r = requests.get(url, timeout=45, headers={"User-Agent": UA})
        if r.status_code != 200 or len(r.content) < 200:
            logger.warning("Pobieranie nieudane (%s): %s", r.status_code, url)
            return None
        dest.write_bytes(r.content)
        return dest
    except Exception as exc:
        logger.warning("Błąd pobierania %s: %s", url, exc)
        return None


def _read_fd_csv(path: Path) -> pd.DataFrame | None:
    try:
        raw = path.read_text(encoding="utf-8", errors="replace")
        return pd.read_csv(StringIO(raw))
    except Exception:
        try:
            return pd.read_csv(path, encoding="latin-1")
        except Exception as exc:
            logger.warning("Nie mogę wczytać %s: %s", path.name, exc)
            return None


def _num(series: pd.Series) -> pd.Series:
    return pd.to_numeric(series, errors="coerce")


def _classic_part(df: pd.DataFrame, league: str) -> pd.DataFrame:
    part: dict[str, Any] = {
        COL_LIGA: league,
        "date_raw": df["Date"],
        "home_raw": df["HomeTeam"],
        "away_raw": df["AwayTeam"],
        "hg": _num(df["FTHG"]),
        "ag": _num(df["FTAG"]),
    }
    fd_cols = {
        "hf": "HF", "af": "AF",
        "hc": "HC", "ac": "AC",
        "hy": "HY", "ay": "AY",
        "hs": "HS", "as_col": "AS",
        "hst": "HST", "ast": "AST",
    }
    for key, src in fd_cols.items():
        if src in df.columns:
            part[key] = _num(df[src])
    return pd.DataFrame(part)


def fetch_classic_results() -> pd.DataFrame:
    frames: list[pd.DataFrame] = []
    for league, code in CLASSIC.items():
        for season in SEASONS:
            url = f"https://www.football-data.co.uk/mmz4281/{season}/{code}.csv"
            dest = CACHE_DIR / f"{code}_{season}.csv"
            path = _download(url, dest) or (dest if dest.exists() else None)
            if path is None:
                continue
            df = _read_fd_csv(path)
            if df is None or "FTHG" not in df.columns:
                continue
            part = _classic_part(df, league)
            frames.append(part)
            logger.info("Classic %s %s: %s wierszy", league, season, len(part))
    return pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()


def fetch_new_results() -> pd.DataFrame:
    frames: list[pd.DataFrame] = []
    for league, (code, league_filter) in NEW_FORMAT.items():
        url = f"https://www.football-data.co.uk/new/{code}.csv"
        dest = CACHE_DIR / f"{code}_new.csv"
        path = _download(url, dest) or (dest if dest.exists() else None)
        if path is None:
            continue
        df = _read_fd_csv(path)
        if df is None or "HG" not in df.columns:
            continue
        if "League" in df.columns:
            df = df[df["League"].astype(str).str.contains(league_filter, case=False, na=False)]
        part = pd.DataFrame(
            {
                COL_LIGA: league,
                "date_raw": df["Date"],
                "home_raw": df["Home"],
                "away_raw": df["Away"],
                "hg": _num(df["HG"]),
                "ag": _num(df["AG"]),
            }
        )
        frames.append(part)
        logger.info("New %s: %s wierszy", league, len(part))
    return pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()


def _parse_dates(series: pd.Series) -> pd.Series:
    s = series.astype(str).str.strip()
    d = pd.to_datetime(s, format="%d/%m/%Y", errors="coerce")
    miss = d.isna()
    if miss.any():
        d2 = pd.to_datetime(s[miss], format="%d/%m/%y", errors="coerce")
        d.loc[miss] = d2
    miss = d.isna()
    if miss.any():
        d3 = pd.to_datetime(s[miss], dayfirst=True, errors="coerce")
        d.loc[miss] = d3
    return d


def prepare_results_index(results: pd.DataFrame) -> pd.DataFrame:
    out = results.copy()
    out["date"] = _parse_dates(out["date_raw"])
    out["home_n"] = out["home_raw"].map(clean_name)
    out["away_n"] = out["away_raw"].map(clean_name)
    out = out.dropna(subset=["date", "hg", "ag"])
    out["hg"] = out["hg"].astype(int)
    out["ag"] = out["ag"].astype(int)
    out[COL_RESULT] = out["hg"].astype(str) + ":" + out["ag"].astype(str)
    return out


def _best_match(
    home_n: str,
    away_n: str,
    candidates: pd.DataFrame,
    min_score: float = 0.72,
) -> pd.Series | None:
    if candidates.empty:
        return None
    best: tuple[float, pd.Series] | None = None
    for _, row in candidates.iterrows():
        score = (
            _similarity(home_n, row["home_n"]) + _similarity(away_n, row["away_n"])
        ) / 2
        if best is None or score > best[0]:
            best = (score, row)
    if best is None or best[0] < min_score:
        return None
    return best[1]


def _ensure_columns(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    for col in all_csv_columns():
        if col not in out.columns:
            out[col] = pd.NA
    return out


def _apply_hit_to_row(df: pd.DataFrame, idx: int, hit: pd.Series) -> None:
    df.at[idx, COL_RESULT] = hit[COL_RESULT]
    for base, hkey, akey in STAT_MAP:
        hcol = f"{base}{HOME_SUFFIX}"
        acol = f"{base}{AWAY_SUFFIX}"
        hv = hit.get(hkey)
        av = hit.get(akey)
        if pd.notna(hv) and pd.notna(av):
            df.at[idx, hcol] = int(hv)
            df.at[idx, acol] = int(av)
            df.at[idx, base] = int(hv) + int(av)


def enrich_dataframe(
    aleks: pd.DataFrame,
    results: pd.DataFrame,
) -> tuple[pd.DataFrame, dict[str, int]]:
    """Dopasowuje wynik + statystyki per drużyna."""
    df = _ensure_columns(aleks)

    res = prepare_results_index(results)
    by_key: dict[tuple[str, pd.Timestamp], pd.DataFrame] = {}
    for (liga, date), grp in res.groupby([COL_LIGA, "date"], sort=False):
        if pd.isna(date):
            continue
        by_key[(str(liga), pd.Timestamp(date).normalize())] = grp

    matched_result = 0
    matched_stats = 0
    unmatched = 0

    dates = _parse_dates(df[COL_DATA])
    for i, row in df.iterrows():
        date = dates.loc[i]
        if pd.isna(date):
            unmatched += 1
            continue

        liga = str(row[COL_LIGA])
        key = (liga, pd.Timestamp(date).normalize())
        cand = by_key.get(key)
        if cand is None or cand.empty:
            unmatched += 1
            continue

        hit = _best_match(clean_name(row[COL_HOME]), clean_name(row[COL_AWAY]), cand)
        if hit is None:
            unmatched += 1
            continue

        existing = row.get(COL_RESULT)
        had_result = bool(
            pd.notna(existing) and re.match(r"^\d+:\d+$", str(existing).strip())
        )
        _apply_hit_to_row(df, i, hit)
        matched_result += 0 if had_result else 1

        if any(pd.notna(hit.get(hk)) and pd.notna(hit.get(ak)) for _, hk, ak in STAT_MAP):
            matched_stats += 1

    stats = {
        "matched_result": matched_result,
        "matched_stats": matched_stats,
        "unmatched": unmatched,
        "total": len(df),
    }
    return df, stats


def load_all_results() -> pd.DataFrame:
    parts = [p for p in (fetch_classic_results(), fetch_new_results()) if not p.empty]
    return pd.concat(parts, ignore_index=True) if parts else pd.DataFrame()


def results_as_aleks(results: pd.DataFrame) -> pd.DataFrame:
    """Zamienia wiersze football-data na schemat aleks_ligi_stats."""
    idx = prepare_results_index(results)
    if idx.empty:
        return pd.DataFrame(columns=all_csv_columns())
    out = pd.DataFrame(
        {
            COL_LIGA: idx[COL_LIGA].astype(str),
            COL_DATA: idx["date"].dt.strftime("%d/%m/%Y"),
            COL_HOME: idx["home_raw"].astype(str),
            COL_AWAY: idx["away_raw"].astype(str),
            COL_RESULT: idx[COL_RESULT],
            "оз": [
                "так" if int(h) > 0 and int(a) > 0 else "ні"
                for h, a in zip(idx["hg"], idx["ag"])
            ],
        }
    )
    for base, hk, ak in STAT_MAP:
        hcol, acol = f"{base}{HOME_SUFFIX}", f"{base}{AWAY_SUFFIX}"
        if hk in idx.columns and ak in idx.columns:
            hv = pd.to_numeric(idx[hk], errors="coerce")
            av = pd.to_numeric(idx[ak], errors="coerce")
            out[hcol] = hv
            out[acol] = av
            out[base] = hv.add(av, fill_value=0)
        else:
            out[hcol] = pd.NA
            out[acol] = pd.NA
            out[base] = pd.NA
    return out


def append_missing_matches(
    aleks: pd.DataFrame,
    results: pd.DataFrame,
    *,
    from_date: pd.Timestamp | None = None,
) -> pd.DataFrame:
    """Dopisuje mecze z football-data, których nie ma jeszcze w tabeli Aleksa."""
    fd = results_as_aleks(results)
    if fd.empty:
        return aleks
    min_date = pd.Timestamp(from_date) if from_date is not None else pd.Timestamp("2024-03-01")
    fd_dates = pd.to_datetime(fd[COL_DATA], dayfirst=True, errors="coerce")
    fd = fd.loc[fd_dates >= min_date].copy()
    if fd.empty:
        return aleks

    def keys(df: pd.DataFrame) -> list[tuple[str, str, str, str]]:
        d = pd.to_datetime(df[COL_DATA], dayfirst=True, errors="coerce")
        return list(
            zip(
                df[COL_LIGA].astype(str),
                d.dt.strftime("%Y-%m-%d"),
                df[COL_HOME].map(clean_name),
                df[COL_AWAY].map(clean_name),
            )
        )

    existing = set(keys(aleks)) if not aleks.empty else set()
    mask = [k not in existing for k in keys(fd)]
    new_rows = fd.loc[mask]
    if new_rows.empty:
        return aleks
    logger.info("Nowe mecze z football-data: %s", len(new_rows))
    return pd.concat([aleks, new_rows], ignore_index=True)


def _integer_stat_columns() -> list[str]:
    cols: list[str] = []
    for base, _, _ in STAT_MAP:
        cols.extend([f"{base}{HOME_SUFFIX}", f"{base}{AWAY_SUFFIX}", base])
    return cols


def normalize_btts(value: object) -> str:
    """Mapuje yes/no/tak/ні → так/ні (ukraiński)."""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    v = str(value).strip().casefold()
    if v in {"yes", "tak", "так", "y", "1", "true"}:
        return "так"
    if v in {"no", "nie", "ні", "n", "0", "false"}:
        return "ні"
    return str(value).strip()


def format_stats_as_ints(df: pd.DataFrame) -> pd.DataFrame:
    """Zapisuje faule/rożne/kartki/strzały jako liczby całkowite (13, nie 13.0)."""
    out = df.copy()
    for col in _integer_stat_columns():
        if col not in out.columns:
            continue
        nums = pd.to_numeric(out[col], errors="coerce")
        out[col] = [
            "" if pd.isna(v) else str(int(round(float(v))))
            for v in nums
        ]
    if "оз" in out.columns:
        out["оз"] = out["оз"].map(normalize_btts)
    return out


def enrich_csv(csv_path: Path | None = None) -> Path:
    path = csv_path or resolve_aleks_path() or ALEKS_XLSX
    if not Path(path).exists() and csv_path is None:
        # świeży start: pozwól odczytać legacy CSV przez resolve
        legacy = ROOT / "aleks_ligi_stats.csv"
        if legacy.exists():
            path = legacy
        else:
            raise FileNotFoundError(ALEKS_XLSX)

    path = Path(path)
    aleks = read_aleks_table(path)
    logger.info("Wczytano %s (%s wierszy)", path.name, len(aleks))

    results = load_all_results()
    if results.empty:
        raise SystemExit("Brak wyników z football-data.co.uk")

    before = len(aleks)
    aleks = append_missing_matches(aleks, results)
    if len(aleks) > before:
        logger.info("Dopisano meczy: %s (teraz %s)", len(aleks) - before, len(aleks))

    enriched, stats = enrich_dataframe(aleks, results)
    preferred = all_csv_columns()
    cols = [c for c in preferred if c in enriched.columns] + [
        c for c in enriched.columns if c not in preferred
    ]
    enriched = format_stats_as_ints(enriched[cols])

    out_path = save_aleks_table(enriched, ALEKS_XLSX)

    _safe_print(
        f"Zapisano {out_path.name}: wyniki+={stats['matched_result']} "
        f"statystyki={stats['matched_stats']} "
        f"unmatched={stats['unmatched']} total={stats['total']}"
    )
    return out_path


def main() -> None:
    logging.basicConfig(level=logging.INFO, format="%(levelname)s %(message)s")
    CACHE_DIR.mkdir(parents=True, exist_ok=True)
    enrich_csv()


if __name__ == "__main__":
    main()
