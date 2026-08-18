"""Aliasy i różne nazwy tej samej drużyny."""

from __future__ import annotations

import pandas as pd

import predykcje as pred
from team_names import (
    canonical_name,
    canonicalize_teams,
    fix_mojibake,
    map_team_to_known,
)


def test_fix_mojibake_utf8_as_latin1():
    assert fix_mojibake("ZÃ¼rich") == "Zürich"
    assert fix_mojibake("IFK GÃ¶teborg") == "IFK Göteborg"
    assert fix_mojibake("Zurich") == "Zurich"


def test_canonical_aliases():
    assert canonical_name("Zürich") == "Zurich"
    assert canonical_name("ZÃ¼rich") == "Zurich"
    assert canonical_name("IFK Göteborg") == "Goteborg"
    assert canonical_name("IFK Gbg") == "Goteborg"
    assert canonical_name("IFK GÃ¶teborg") == "Goteborg"
    assert canonical_name("Athletic Club") == "Ath Bilbao"
    assert canonical_name("PSV") == "PSV Eindhoven"
    assert canonical_name("Fortuna Sittard") == "For Sittard"
    assert canonical_name("Nottm Forest") == "Nott'm Forest"
    assert canonical_name("Mgladbach") == "M'gladbach"
    assert canonical_name("Lausanne-Sport") == "Lausanne"
    assert canonical_name("ADO Den Haag") == "Den Haag"
    assert canonical_name("Västerås") == "Vasteras SK"


def test_map_team_to_known_uses_history_spelling():
    known = {
        "Super League": ["Zurich", "Basel", "Lausanne"],
        "Allsvenskan": ["Goteborg", "Elfsborg", "Hacken"],
        "Eredivisie": ["PSV Eindhoven", "For Sittard", "AZ Alkmaar"],
    }
    assert map_team_to_known("ZÃ¼rich", "Super League", known) == "Zurich"
    assert map_team_to_known("IFK Gbg", "Allsvenskan", known) == "Goteborg"
    assert map_team_to_known("Lausanne-Sport", "Super League", known) == "Lausanne"
    assert map_team_to_known("PSV", "Eredivisie", known) == "PSV Eindhoven"
    assert map_team_to_known("Fortuna Sittard", "Eredivisie", known) == "For Sittard"


def test_canonicalize_merges_history_spellings():
    df = pd.DataFrame(
        {
            "ліга": ["Bundesliga", "Bundesliga"],
            "господар": ["Mgladbach", "M'gladbach"],
            "гість": ["Dortmund", "Bayern Munich"],
        }
    )
    out = canonicalize_teams(df)
    assert set(out["господар"]) == {"M'gladbach"}


def test_predict_match_finds_form_via_alias():
    rows = []
    for i, (h, a, score) in enumerate(
        [
            ("Zurich", "Basel", "1:0"),
            ("Basel", "Zurich", "0:2"),
            ("Zurich", "Luzern", "2:1"),
            ("Young Boys", "Zurich", "1:1"),
            ("Basel", "Luzern", "3:0"),
            ("Luzern", "Basel", "0:1"),
            ("Young Boys", "Basel", "1:2"),
            ("Luzern", "Young Boys", "0:0"),
        ]
    ):
        rows.append(
            {
                "ліга": "Super League",
                "дата": pd.Timestamp(f"2026-03-{i + 1:02d}"),
                "господар": h,
                "гість": a,
                "результат": score,
                "оз": "так" if "0" not in score else "ні",
                "фоли": 20,
                "кутові": 10,
                "жовті_картки": 3,
                "удари": 22,
                "удари_в_площину": 8,
            }
        )
    hist = pred._attach_goals(pd.DataFrame(rows))
    result = pred.predict_match("Zürich", "Basel", "Super League", hist)
    assert result["status"] == "ok"
    assert result["forma_gospodarz_mecze"] >= 3
    assert result["powod"] != "za_malo_danych"


def test_index_merges_gladbach_spellings():
    rows = []
    pairs = [
        ("Mgladbach", "Dortmund", "1:0"),
        ("Dortmund", "Mgladbach", "0:1"),
        ("M'gladbach", "Bayern Munich", "2:2"),
        ("Bayern Munich", "M'gladbach", "1:0"),
        ("Dortmund", "Bayern Munich", "2:1"),
        ("Bayern Munich", "Dortmund", "3:1"),
    ]
    for i, (h, a, score) in enumerate(pairs):
        rows.append(
            {
                "ліга": "Bundesliga",
                "дата": pd.Timestamp(f"2026-02-{i + 1:02d}"),
                "господар": h,
                "гість": a,
                "результат": score,
                "оз": "так",
                "фоли": 18,
                "кутові": 9,
                "жовті_картки": 3,
                "удари": 20,
                "удари_в_площину": 7,
            }
        )
    hist = pred._attach_goals(pd.DataFrame(rows))
    form = pred.compute_form(hist, "Mgladbach", index=pred.get_team_index(hist))
    assert form["mecze"] == 4
    result = pred.predict_match("Mgladbach", "Dortmund", "Bundesliga", hist)
    assert result["status"] == "ok"
