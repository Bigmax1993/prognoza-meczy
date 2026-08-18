"""Testy statystyk per drużyna i podsumowania miesięcznego."""

from __future__ import annotations

import pandas as pd

import enrich_scores as es
import monthly_summary as ms


def test_apply_hit_splits_team_stats():
    aleks = pd.DataFrame(
        [
            {
                "ліга": "Premier League",
                "дата": "15/08/2025",
                "господар": "Liverpool",
                "гість": "Bournemouth",
                "оз": "yes",
            }
        ]
    )
    results = pd.DataFrame(
        [
            {
                "ліга": "Premier League",
                "date_raw": "15/08/2025",
                "home_raw": "Liverpool",
                "away_raw": "Bournemouth",
                "hg": 4,
                "ag": 2,
                "hf": 7,
                "af": 10,
                "hc": 6,
                "ac": 7,
                "hy": 1,
                "ay": 2,
                "hs": 19,
                "as_col": 10,
                "hst": 10,
                "ast": 3,
            }
        ]
    )
    out, stats = es.enrich_dataframe(aleks, results)
    assert out.iloc[0]["результат"] == "4:2"
    assert out.iloc[0]["фоли_господар"] == 7
    assert out.iloc[0]["фоли_гість"] == 10
    assert out.iloc[0]["фоли"] == 17
    assert out.iloc[0]["кутові_господар"] == 6
    assert out.iloc[0]["удари_гість"] == 10
    assert stats["matched_stats"] == 1


def test_format_stats_as_ints_no_decimals():
    df = pd.DataFrame(
        {
            "фоли": [13.0, None],
            "фоли_господар": [7.0, 10.0],
            "фоли_гість": [6.0, pd.NA],
            "кутові": [12.0, 5.5],
            "кутові_господар": [5.0, 2.0],
            "кутові_гість": [7.0, 3.0],
            "жовті_картки": [2.0, 1.0],
            "жовті_картки_господар": [1.0, 0.0],
            "жовті_картки_гість": [1.0, 1.0],
            "удари": [21.0, 10.0],
            "удари_господар": [11.0, 4.0],
            "удари_гість": [10.0, 6.0],
            "удари_в_площину": [7.0, 3.0],
            "удари_в_площину_господар": [4.0, 1.0],
            "удари_в_площину_гість": [3.0, 2.0],
        }
    )
    out = es.format_stats_as_ints(df)
    assert out.loc[0, "фоли"] == "13"
    assert out.loc[1, "фоли"] == ""
    assert out.loc[0, "кутові"] == "12"
    assert "." not in str(out.loc[0, "удари"])


def test_monthly_team_summary():
    df = pd.DataFrame(
        [
            {
                "ліга": "Premier League",
                "дата": "01/01/2026",
                "господар": "Arsenal",
                "гість": "Chelsea",
                "фоли_господар": 10,
                "фоли_гість": 12,
                "кутові_господар": 5,
                "кутові_гість": 4,
            },
            {
                "ліга": "Premier League",
                "дата": "15/01/2026",
                "господар": "Chelsea",
                "гість": "Arsenal",
                "фоли_господар": 14,
                "фоли_гість": 8,
                "кутові_господар": 6,
                "кутові_гість": 3,
            },
        ]
    )
    df["дата"] = pd.to_datetime(df["дата"], format="%d/%m/%Y")
    summary = ms.compute_monthly_team_summary(df, year=2026)
    assert not summary.empty
    assert "srednia_фоли" in summary.columns
    arsenal = summary[summary["druzyna"] == "Arsenal"].iloc[0]
    assert arsenal["mecze"] == 2
    assert arsenal["srednia_фоли"] == 9.0  # (10+8)/2
    assert summary["miesiac"].iloc[0] == "2026-01"
