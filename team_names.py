# -*- coding: utf-8 -*-
"""Jedna kanoniczna nazwa drużyny: BBC / football-data / literówki w historii."""
from __future__ import annotations

import re

import pandas as pd

from enrich_scores import _similarity, clean_name

COL_LIGA = "ліга"
COL_HOME = "господар"
COL_AWAY = "гість"

# Prefiksy/sufiksy, które nie rozróżniają klubów (IFK, FC, PEC…).
_NOISE = frozenset(
    {
        "fc",
        "cf",
        "bk",
        "ifk",
        "if",
        "isk",
        "ado",
        "pec",
        "the",
        "de",
        "club",
        "afc",
        "ac",
        "as",
        "ssc",
        "us",
        "sv",
        "tsg",
        "vfl",
        "vfb",
        "bsc",
        "sc",
        "rcd",
        "real",  # tylko jako osobny token; Real Madrid zostaje po "madrid"
    }
)

# Klucz = clean_name (bez ogonków). Wartość = nazwa z football-data / historii Aleksa.
TEAM_ALIASES: dict[str, str] = {
    # Premier League
    "manchester united": "Man United",
    "man utd": "Man United",
    "man united": "Man United",
    "manchester city": "Man City",
    "man city": "Man City",
    "tottenham hotspur": "Tottenham",
    "spurs": "Tottenham",
    "nottingham forest": "Nott'm Forest",
    "nottm forest": "Nott'm Forest",
    "nott m forest": "Nott'm Forest",
    "wolverhampton wanderers": "Wolves",
    "brighton and hove albion": "Brighton",
    "brighton hove albion": "Brighton",
    "west ham united": "West Ham",
    "newcastle united": "Newcastle",
    "leicester city": "Leicester",
    "leeds united": "Leeds",
    "ipswich town": "Ipswich",
    "hull city": "Hull",
    "coventry city": "Coventry",
    # La Liga
    "athletic club": "Ath Bilbao",
    "athletic bilbao": "Ath Bilbao",
    "ath bilbao": "Ath Bilbao",
    "atletico madrid": "Ath Madrid",
    "ath madrid": "Ath Madrid",
    "atletico de madrid": "Ath Madrid",
    "real sociedad": "Sociedad",
    "real betis": "Betis",
    "celta vigo": "Celta",
    "celta de vigo": "Celta",
    "rayo vallecano": "Vallecano",
    "espanyol": "Espanol",
    "rcd espanyol": "Espanol",
    "deportivo de a coruna": "Dep. A Coruna",
    "deportivo la coruna": "Dep. A Coruna",
    "de a coruna": "Dep. A Coruna",
    "dep. a coruna": "Dep. A Coruna",
    "racing de santander": "Santander",
    "racing santander": "Santander",
    "malaga cf": "Malaga",
    "malaga": "Malaga",
    "deportivo alaves": "Alaves",
    "alaves": "Alaves",
    # Serie A
    "inter milan": "Inter",
    "internazionale": "Inter",
    "fc internazionale": "Inter",
    "ac milan": "Milan",
    "as roma": "Roma",
    "hellas verona": "Verona",
    # Bundesliga
    "borussia dortmund": "Dortmund",
    "borussia monchengladbach": "M'gladbach",
    "borussia mgladbach": "M'gladbach",
    "monchengladbach": "M'gladbach",
    "mgladbach": "M'gladbach",
    "eintracht frankfurt": "Ein Frankfurt",
    "bayer leverkusen": "Leverkusen",
    "tsg hoffenheim": "Hoffenheim",
    "mainz 05": "Mainz",
    "1. fsv mainz 05": "Mainz",
    "fc koln": "FC Koln",
    "1. fc koln": "FC Koln",
    "cologne": "FC Koln",
    "bayern munchen": "Bayern Munich",
    "fc bayern munich": "Bayern Munich",
    "fc bayern munchen": "Bayern Munich",
    "rasenballsport leipzig": "RB Leipzig",
    "st. pauli": "St Pauli",
    "fc st. pauli": "St Pauli",
    # Eredivisie
    "psv eindhoven": "PSV Eindhoven",
    "psv": "PSV Eindhoven",
    "az": "AZ Alkmaar",
    "az alkmaar": "AZ Alkmaar",
    "fortuna sittard": "For Sittard",
    "for sittard": "For Sittard",
    "pec zwolle": "Zwolle",
    "zwolle": "Zwolle",
    "ado den haag": "Den Haag",
    "den haag": "Den Haag",
    "nec nijmegen": "Nijmegen",
    "nec": "Nijmegen",
    # Allsvenskan
    "ifk goteborg": "Goteborg",
    "ifk gbg": "Goteborg",
    "goteborg": "Goteborg",
    "bk hacken": "Hacken",
    "hacken": "Hacken",
    "orgryte is": "Orgryte",
    "orgryte": "Orgryte",
    "vasteras sk": "Vasteras SK",
    "vasteras": "Vasteras SK",
    "malmo ff": "Malmo FF",
    "malmo": "Malmo FF",
    "djurgardens if": "Djurgarden",
    "djurgarden": "Djurgarden",
    "ifk norrkoping": "Norrkoping",
    "norrkoping": "Norrkoping",
    # Super League
    "zurich": "Zurich",
    "fc zurich": "Zurich",
    "lausanne-sport": "Lausanne",
    "lausanne sport": "Lausanne",
    "fc lausanne-sport": "Lausanne",
    "lausanne": "Lausanne",
    "grasshopper": "Grasshoppers",
    "grasshopper club": "Grasshoppers",
    "grasshoppers": "Grasshoppers",
    "fc luzern": "Luzern",
    "fc basel": "Basel",
    # Eliteserien
    "bodo/glimt": "Bodo/Glimt",
    "bodo glimt": "Bodo/Glimt",
    "stromsgodset": "Stromsgodset",
    "tromso": "Tromso",
    "valerenga": "Valerenga",
    "kristiansund bk": "Kristiansund",
    "kfum-kameratene": "KFUM Oslo",
    "kfum oslo": "KFUM Oslo",
}


def fix_mojibake(text: str) -> str:
    """Naprawia UTF-8 odczytane jako Latin-1 (ZÃ¼rich → Zürich)."""
    if not text:
        return text
    if "Ã" not in text and "Â" not in text:
        return text
    try:
        fixed = text.encode("latin-1").decode("utf-8")
    except (UnicodeDecodeError, UnicodeEncodeError):
        return text
    return fixed if fixed else text


def _norm(name: str) -> str:
    return clean_name(fix_mojibake(name)).replace("-", " ")


def _core_tokens(normed: str) -> tuple[str, ...]:
    parts = [p for p in re.split(r"[\s./]+", normed) if p and p not in _NOISE]
    return tuple(parts)


def canonical_name(name: str) -> str:
    """Alias → nazwa z football-data; inaczej oryginał po naprawie kodowania."""
    raw = fix_mojibake((name or "").strip())
    if not raw:
        return raw
    alias = TEAM_ALIASES.get(_norm(raw))
    return alias if alias else raw


def known_teams_by_league(history: pd.DataFrame) -> dict[str, list[str]]:
    if history is None or history.empty:
        return {}
    out: dict[str, list[str]] = {}
    for liga, sub in history.groupby(COL_LIGA):
        teams = pd.unique(pd.concat([sub[COL_HOME], sub[COL_AWAY]], ignore_index=True).dropna())
        out[str(liga)] = [str(t) for t in teams]
    return out


def map_team_to_known(
    name: str,
    league: str,
    known: dict[str, list[str]],
    *,
    min_score: float = 0.72,
) -> str:
    """Sprowadza nazwę (BBC / alias / mojibake) do nazwy z historii ligi."""
    raw = fix_mojibake((name or "").strip())
    if not raw:
        return raw
    teams = known.get(league) or []
    n = _norm(raw)

    alias = TEAM_ALIASES.get(n)
    if alias:
        for t in teams:
            if t == alias:
                return str(t)
        return alias

    for t in teams:
        if _norm(t) == n:
            return str(t)

    raw_core = _core_tokens(n)
    if raw_core:
        for t in teams:
            if _core_tokens(_norm(t)) == raw_core:
                return str(t)

    best = raw
    best_score = 0.0
    for t in teams:
        tn = _norm(t)
        score = _similarity(n, tn)
        t_core = " ".join(_core_tokens(tn))
        r_core = " ".join(raw_core)
        if t_core and r_core:
            score = max(score, _similarity(r_core, t_core))
        if score > best_score:
            best, best_score = str(t), score
    return best if best_score >= min_score else raw


def canonicalize_teams(
    df: pd.DataFrame,
    known: dict[str, list[str]] | None = None,
) -> pd.DataFrame:
    """Przepisuje gospodarz/gość na kanoniczne nazwy (ta sama drużyna = jedna etykieta)."""
    if df is None or df.empty or COL_HOME not in df.columns:
        return df
    known = known or known_teams_by_league(df)
    out = df.copy()
    ligas = out[COL_LIGA].astype(str) if COL_LIGA in out.columns else [""] * len(out)
    out[COL_HOME] = [
        map_team_to_known(h, lg, known) for h, lg in zip(out[COL_HOME].astype(str), ligas)
    ]
    out[COL_AWAY] = [
        map_team_to_known(a, lg, known) for a, lg in zip(out[COL_AWAY].astype(str), ligas)
    ]
    return out
