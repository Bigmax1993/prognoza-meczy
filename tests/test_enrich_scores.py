"""Testy enrich_scores + wynik meczu w predykcjach."""

from __future__ import annotations

from pathlib import Path

import pandas as pd

import enrich_scores as es
import predykcje as pred


def test_parse_score():
    assert pred.parse_score("2:1") == (2, 1)
    assert pred.parse_score("0-0") == (0, 0)
    assert pred.parse_score("") is None
    assert pred.parse_score(None) is None


def test_enrich_dataframe_matches_by_date_and_team(tmp_path):
    aleks = pd.DataFrame(
        [
            {
                "ліга": "Premier League",
                "дата": "15/08/2025",
                "господар": "Liverpool",
                "гість": "Bournemouth",
                "оз": "yes",
                "фоли": 17,
                "кутові": 13,
                "жовті_картки": 3,
                "удари": 29,
                "удари_в_площину": 13,
            },
            {
                "ліга": "Premier League",
                "дата": "16/08/2025",
                "господар": "Aston Villa",
                "гість": "Newcastle",
                "оз": "no",
                "фоли": 24,
                "кутові": 9,
                "жовті_картки": 2,
                "удари": 19,
                "удари_в_площину": 6,
            },
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
            },
            {
                "ліга": "Premier League",
                "date_raw": "16/08/2025",
                "home_raw": "Aston Villa",
                "away_raw": "Newcastle",
                "hg": 0,
                "ag": 0,
            },
        ]
    )
    out, stats = es.enrich_dataframe(aleks, results)
    assert "результат" in out.columns
    assert out.iloc[0]["результат"] == "4:2"
    assert out.iloc[1]["результат"] == "0:0"
    assert stats["matched_result"] == 2


def test_predict_uses_goals_when_available():
    rows = []
    for i, (h, a, score) in enumerate(
        [
            ("Arsenal", "Chelsea", "3:1"),
            ("Chelsea", "Arsenal", "1:2"),
            ("Arsenal", "Liverpool", "2:0"),
            ("Liverpool", "Chelsea", "1:1"),
        ]
    ):
        rows.append(
            {
                "ліга": "Premier League",
                "дата": f"0{i+1}/01/2026",
                "господар": h,
                "гість": a,
                "результат": score,
                "оз": "yes" if "0" not in score.replace(":", "") else "no",
                "фоли": 20,
                "кутові": 10,
                "жовті_картки": 3,
                "удари": 25,
                "удари_в_площину": 10,
            }
        )
    path = Path("cache") / "_tmp_goals_test.csv"
    path.parent.mkdir(exist_ok=True)
    pd.DataFrame(rows).to_csv(path, index=False, encoding="utf-8-sig")
    try:
        df_2026, hist = pred.load_2026_data(path)
        avg = pred.compute_team_averages(hist)
        assert "srednia_goli_strzelonych" in avg.columns
        assert avg["mecze_z_wynikiem"].sum() > 0
        result = pred.predict_match("Arsenal", "Chelsea", "Premier League", hist, avg)
        assert result["metoda"] in {"model_wynikow", "forma_5"}
        assert ":" in result["przewidywany_wynik"]
    finally:
        path.unlink(missing_ok=True)


def test_append_missing_matches():
    aleks = pd.DataFrame(
        [
            {
                "ліга": "Premier League",
                "дата": "15/08/2025",
                "господар": "Liverpool",
                "гість": "Bournemouth",
                "результат": "2:1",
                "оз": "так",
            }
        ]
    )
    results = pd.DataFrame(
        {
            "ліга": ["Premier League", "Premier League"],
            "date_raw": ["15/08/2025", "13/08/2026"],
            "home_raw": ["Liverpool", "Arsenal"],
            "away_raw": ["Bournemouth", "Chelsea"],
            "hg": [2, 1],
            "ag": [1, 0],
        }
    )
    out = es.append_missing_matches(aleks, results)
    assert len(out) == 2
    dates = pd.to_datetime(out["дата"], dayfirst=True)
    assert pd.Timestamp("2026-08-13") in set(dates)


def test_append_missing_matches_skips_old_history():
    aleks = pd.DataFrame(
        [
            {
                "ліга": "Allsvenskan",
                "дата": "15/08/2025",
                "господар": "AIK",
                "гість": "Malmo",
                "результат": "1:0",
                "оз": "ні",
            }
        ]
    )
    results = pd.DataFrame(
        {
            "ліга": ["Allsvenskan"],
            "date_raw": ["01/05/2018"],
            "home_raw": ["AIK"],
            "away_raw": ["Malmo"],
            "hg": [2],
            "ag": [1],
        }
    )
    out = es.append_missing_matches(aleks, results)
    assert len(out) == 1
