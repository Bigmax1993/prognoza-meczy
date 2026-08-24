"""Testy upcoming.py (bez sieci)."""

from __future__ import annotations

from datetime import date

import pandas as pd
import pytest

import upcoming as up


def test_complete_match_date():
    assert up.complete_match_date("18/08", year=2026) == "18/08/2026"
    assert up.complete_match_date("18/08/2026") == "18/08/2026"


def test_map_team_to_known():
    known = {
        "Premier League": ["Arsenal", "Man City", "Chelsea", "Nott'm Forest"],
        "Super League": ["Zurich", "Basel"],
        "Allsvenskan": ["Goteborg", "Elfsborg"],
        "La Liga": ["Ath Bilbao", "Sevilla"],
    }
    assert up.map_team_to_known("Arsenal", "Premier League", known) == "Arsenal"
    assert up.map_team_to_known("Arsenal FC", "Premier League", known) == "Arsenal"
    assert up.map_team_to_known("Manchester City", "Premier League", known) == "Man City"
    assert up.map_team_to_known("Nottingham Forest", "Premier League", known) == "Nott'm Forest"
    assert up.map_team_to_known("ZÃ¼rich", "Super League", known) == "Zurich"
    assert up.map_team_to_known("IFK Gbg", "Allsvenskan", known) == "Goteborg"
    assert up.map_team_to_known("Athletic Club", "La Liga", known) == "Ath Bilbao"


def test_rows_to_aleks_fixtures_filters_leagues():
    rows = [
        {
            "Kraj": "England",
            "Liga": "Premier League",
            "Gospodarz": "Arsenal",
            "Gość": "Chelsea",
            "Data": "18/08",
        },
        {
            "Kraj": "England",
            "Liga": "National League",
            "Gospodarz": "Barnet",
            "Gość": "York",
            "Data": "18/08",
        },
    ]
    known = {"Premier League": ["Arsenal", "Chelsea"]}
    df = up.rows_to_aleks_fixtures(rows, match_day=date(2026, 8, 18), known=known)
    assert len(df) == 1
    assert df.iloc[0]["ліга"] == "Premier League"
    assert df.iloc[0]["дата"] == "18/08/2026"
    assert df.iloc[0]["результат"] == ""


def test_rows_to_aleks_keeps_premapped_la_liga():
    rows = [
        {
            "Kraj": "",
            "Liga": "La Liga",
            "_aleks_liga": "La Liga",
            "Gospodarz": "Real Madrid",
            "Gość": "Osasuna",
            "Data": "19/08/2026",
        }
    ]
    df = up.rows_to_aleks_fixtures(rows, match_day=date(2026, 8, 19), known={})
    assert len(df) == 1
    assert df.iloc[0]["ліга"] == "La Liga"


def test_parse_bbc_fixtures():
    html = """
    <html><body>
      <div class="ssrcss-1ox7t1a-Container">
        <h2 class="ssrcss-m22vfq-GroupHeader">Premier League</h2>
        <span class="visually-hidden">Arsenal versus Chelsea kick off 15:00</span>
        <span class="visually-hidden">Everton versus Crystal Palace kick off 17:30</span>
      </div>
      <div class="ssrcss-1ox7t1a-Container">
        <h2 class="ssrcss-m22vfq-GroupHeader">Brazilian Serie A</h2>
        <span class="visually-hidden">Flamengo versus Palmeiras kick off 20:00</span>
      </div>
    </body></html>
    """
    rows = up.parse_bbc_fixtures(html, date(2026, 8, 22))
    assert len(rows) == 2
    assert {r["Liga"] for r in rows} == {"Premier League"}
    assert rows[0]["Gospodarz"] == "Arsenal"
    assert rows[0]["Gość"] == "Chelsea"


def test_parse_bbc_full_time_scores():
    html = """
    <html><body>
      <div>
        <h2 class="ssrcss-m22vfq-GroupHeader">Premier League</h2>
        <span class="visually-hidden">Brighton & Hove Albion 4 , Aston Villa 0 at Full time</span>
        <span class="visually-hidden">Manchester City 2 , Bournemouth 1 at Full time</span>
      </div>
      <div>
        <h2 class="ssrcss-m22vfq-GroupHeader">Spanish La Liga</h2>
        <span class="visually-hidden">Elche 0 , Barcelona 5 at Full time</span>
      </div>
    </body></html>
    """
    rows = up.parse_bbc_fixtures(html, date(2026, 8, 23))
    assert len(rows) == 3
    by_home = {r["Gospodarz"]: r for r in rows}
    assert by_home["Brighton & Hove Albion"]["Wynik"] == "4:0"
    assert by_home["Manchester City"]["Wynik"] == "2:1"
    assert by_home["Elche"]["Liga"] == "La Liga"
    assert by_home["Elche"]["Wynik"] == "0:5"


def test_rows_to_aleks_keeps_full_time_score():
    rows = [
        {
            "Kraj": "",
            "Liga": "Premier League",
            "_aleks_liga": "Premier League",
            "Gospodarz": "Man City",
            "Gość": "Bournemouth",
            "Data": "23/08/2026",
            "Wynik": "2:1",
        }
    ]
    known = {"Premier League": ["Man City", "Bournemouth"]}
    df = up.rows_to_aleks_fixtures(rows, match_day=date(2026, 8, 23), known=known)
    assert len(df) == 1
    assert df.iloc[0]["результат"] == "2:1"
    assert df.iloc[0]["оз"] == "так"


def test_merge_upcoming_skips_duplicates():
    existing = pd.DataFrame(
        [
            {
                "ліга": "Premier League",
                "дата": pd.Timestamp("2026-08-18"),
                "господар": "Arsenal",
                "гість": "Chelsea",
                "результат": "1:0",
            }
        ]
    )
    upcoming = pd.DataFrame(
        [
            {
                "ліга": "Premier League",
                "дата": "18/08/2026",
                "господар": "Arsenal",
                "гість": "Chelsea",
                "результат": "",
            },
            {
                "ліга": "La Liga",
                "дата": "18/08/2026",
                "господар": "Real Madrid",
                "гість": "Osasuna",
                "результат": "",
            },
        ]
    )
    out = up.merge_upcoming(existing, upcoming)
    assert len(out) == 2
    assert "Osasuna" in set(out["гість"].astype(str))


def test_merge_upcoming_fills_blank_score():
    existing = pd.DataFrame(
        [
            {
                "ліга": "Premier League",
                "дата": pd.Timestamp("2026-08-23"),
                "господар": "Man City",
                "гість": "Bournemouth",
                "результат": "",
                "оз": "",
            }
        ]
    )
    upcoming = pd.DataFrame(
        [
            {
                "ліга": "Premier League",
                "дата": "23/08/2026",
                "господар": "Man City",
                "гість": "Bournemouth",
                "результат": "2:1",
                "оз": "так",
            }
        ]
    )
    out = up.merge_upcoming(existing, upcoming)
    assert len(out) == 1
    assert out.iloc[0]["результат"] == "2:1"
    assert out.iloc[0]["оз"] == "так"


def test_lookback_days_from_history():
    hist = pd.DataFrame({"дата": [pd.Timestamp("2026-08-17")]})
    assert up.lookback_days_from_history(hist, start=date(2026, 8, 24)) == 7
    assert up.lookback_days_from_history(hist, start=date(2026, 8, 17)) == 0
    assert up.lookback_days_from_history(None, start=date(2026, 8, 24)) == 0
    wide = pd.DataFrame({"дата": [pd.Timestamp("2026-01-01")]})
    assert up.lookback_days_from_history(wide, start=date(2026, 8, 24)) == up.MAX_LOOKBACK_DAYS


def test_upsert_matches_replaces_row():
    existing = pd.DataFrame(
        [
            {
                "ліга": "Premier League",
                "дата": pd.Timestamp("2026-01-10"),
                "господар": "Arsenal",
                "гість": "Chelsea",
                "результат": "1:0",
                "кутові": 8,
            },
            {
                "ліга": "Premier League",
                "дата": pd.Timestamp("2026-08-16"),
                "господар": "Everton",
                "гість": "Leeds",
                "результат": "0:0",
                "кутові": "",
            },
        ]
    )
    updated = pd.DataFrame(
        [
            {
                "ліга": "Premier League",
                "дата": pd.Timestamp("2026-08-16"),
                "господар": "Everton",
                "гість": "Leeds",
                "результат": "0:0",
                "кутові": 11,
            }
        ]
    )
    out = up.upsert_matches(existing, updated)
    assert len(out) == 2
    eve = out[out["господар"].astype(str) == "Everton"].iloc[0]
    assert eve["кутові"] == 11
    assert "Arsenal" in set(out["господар"].astype(str))
