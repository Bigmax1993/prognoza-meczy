"""Serper + BS4 + Claude: uzupełnianie luk (bez sieci w testach)."""

from __future__ import annotations

import pandas as pd

import fill_missing as fm


def test_search_query_contains_team():
    q = fm.search_query("Hull", "Premier League", "13/08/2026")
    assert "Hull" in q
    assert "2026" in q


def test_html_to_text_strips_scripts():
    html = "<html><script>void 0</script><p>Hull 2-1 Leeds 10/08/2025</p></html>"
    text = fm.html_to_text(html)
    assert "void" not in text
    assert "Hull" in text


def test_extract_json_from_fence():
    raw = '```json\n{"match": true, "matches": []}\n```'
    assert fm.extract_json(raw) == {"match": True, "matches": []}


def test_validate_match_rejects_future_and_wrong_club():
    as_of = pd.Timestamp("2026-08-13")
    bad_date = fm.validate_match(
        {"date": "20/08/2026", "home": "Hull", "away": "Leeds", "home_goals": 1, "away_goals": 0},
        team="Hull",
        as_of=as_of,
    )
    assert bad_date is None
    wrong = fm.validate_match(
        {"date": "01/05/2026", "home": "Arsenal", "away": "Chelsea", "home_goals": 1, "away_goals": 0},
        team="Hull",
        as_of=as_of,
    )
    assert wrong is None


def test_validate_match_accepts_alias_and_canonicalizes():
    ok = fm.validate_match(
        {
            "date": "10/05/2026",
            "home": "Hull City",
            "away": "Leeds",
            "home_goals": 2,
            "away_goals": 1,
            "competition": "Championship",
        },
        team="Hull",
        as_of=pd.Timestamp("2026-08-13"),
    )
    assert ok is not None
    assert ok["home"] == "Hull"
    assert ok["away"] == "Leeds"
    assert ok["home_goals"] == 2


def test_matches_to_aleks_and_merge():
    matches = [
        {
            "date": "10/05/2026",
            "home": "Hull",
            "away": "Leeds",
            "home_goals": 2,
            "away_goals": 1,
            "corners_home": None,
            "corners_away": None,
            "cards_home": None,
            "cards_away": None,
            "competition": "Championship",
        }
    ]
    extra = fm.matches_to_aleks(matches)
    assert len(extra) == 1
    assert extra.iloc[0]["результат"] == "2:1"
    hist = pd.DataFrame(
        [
            {
                "ліга": "Premier League",
                "дата": "01/01/2026",
                "господар": "Arsenal",
                "гість": "Chelsea",
                "результат": "1:0",
            }
        ]
    )
    out = fm.merge_history(hist, extra)
    assert len(out) == 2


def test_thin_teams_finds_club_with_no_history():
    target = pd.DataFrame(
        [
            {
                "ліга": "Premier League",
                "дата": "22/08/2026",
                "господар": "Hull",
                "гість": "Man United",
            }
        ]
    )
    hist = pd.DataFrame(
        [
            {
                "ліга": "Premier League",
                "дата": "01/01/2026",
                "господар": "Arsenal",
                "гість": "Chelsea",
            }
        ]
    )
    thin = fm.thin_teams(target, hist)
    names = {t["team"] for t in thin}
    assert "Hull" in names


def test_fill_gaps_uses_serper_bs4_claude_mocks(tmp_path):
    target = pd.DataFrame(
        [
            {
                "ліга": "Premier League",
                "дата": "22/08/2026",
                "господар": "Hull",
                "гість": "Man United",
            }
        ]
    )
    hist = pd.DataFrame(
        [
            {
                "ліга": "Premier League",
                "дата": "01/01/2026",
                "господар": "Arsenal",
                "гість": "Chelsea",
            }
        ]
    )

    def search_fn(query: str):
        assert "Hull" in query
        return [{"title": "Hull results", "link": "https://example.com/hull", "snippet": "2-1"}]

    def fetch_fn(url: str):
        assert url.endswith("/hull")
        return "Hull 2-1 Leeds 10 May 2026 Championship"

    def decide_fn(**kwargs):
        assert kwargs["team"] == "Hull"
        assert "Hull" in kwargs["page_text"]
        return {
            "match": True,
            "reason": "results list",
            "matches": [
                {
                    "date": "10/05/2026",
                    "home": "Hull",
                    "away": "Leeds",
                    "home_goals": 2,
                    "away_goals": 1,
                    "competition": "Championship",
                }
            ],
        }

    cache = tmp_path / "missing_matches.json"
    extra, report = fm.fill_gaps(
        target,
        hist,
        search_fn=search_fn,
        fetch_fn=fetch_fn,
        decide_fn=decide_fn,
        cache_path=cache,
    )
    assert report["accepted_matches"] == 1
    assert len(extra) == 1
    assert extra.iloc[0]["господар"] == "Hull"
    assert extra.iloc[0]["результат"] == "2:1"
    assert cache.exists()
    dumped = fm.load_missing_json(cache)
    assert dumped["accepted_matches"] == 1


def test_scan_missing_skips_future_and_finds_played_stats_gap():
    df = pd.DataFrame(
        [
            {
                "ліга": "Allsvenskan",
                "дата": "14/08/2026",
                "господар": "Elfsborg",
                "гість": "Vasteras SK",
                "результат": "3:0",
                "кутові_господар": pd.NA,
                "кутові_гість": pd.NA,
                "кутові": pd.NA,
            },
            {
                "ліга": "Premier League",
                "дата": "22/08/2026",
                "господар": "Arsenal",
                "гість": "Chelsea",
                "результат": pd.NA,
                "кутові_господар": pd.NA,
                "кутові_гість": pd.NA,
                "кутові": pd.NA,
            },
        ]
    )
    gaps = fm.scan_missing(df, as_of=pd.Timestamp("2026-08-18"))
    assert len(gaps) == 1
    assert gaps[0]["home"] == "Elfsborg"
    assert "кутові_господар" in gaps[0]["missing"]


def test_scan_missing_from_date_skips_older_played():
    df = pd.DataFrame(
        [
            {
                "ліга": "Allsvenskan",
                "дата": "01/05/2026",
                "господар": "Elfsborg",
                "гість": "Hacken",
                "результат": "1:0",
                "кутові_господар": pd.NA,
                "кутові_гість": pd.NA,
                "кутові": pd.NA,
            },
            {
                "ліга": "Allsvenskan",
                "дата": "16/08/2026",
                "господар": "Malmo FF",
                "гість": "AIK",
                "результат": "2:1",
                "кутові_господар": pd.NA,
                "кутові_гість": pd.NA,
                "кутові": pd.NA,
            },
        ]
    )
    gaps = fm.scan_missing(
        df,
        as_of=pd.Timestamp("2026-08-24"),
        from_date=pd.Timestamp("2026-08-13"),
    )
    assert len(gaps) == 1
    assert gaps[0]["home"] == "Malmo FF"


def test_validate_stats_rejects_wrong_score_and_fills_totals():
    gap = {
        "result": "3:0",
        "missing": ["кутові_господар", "кутові_гість", "кутові"],
    }
    assert (
        fm.validate_stats(
            {"match": True, "home_goals": 1, "away_goals": 1, "corners_home": 4, "corners_away": 2},
            gap,
        )
        is None
    )
    ok = fm.validate_stats(
        {"match": True, "corners_home": 5, "corners_away": 3},
        gap,
    )
    assert ok is not None
    assert ok["кутові_господар"] == 5
    assert ok["кутові"] == 8


def test_validate_stats_completes_total_from_already_filled():
    gap = {
        "result": "3:0",
        "missing": ["жовті_картки"],
        "filled": {"жовті_картки_господар": 2, "жовті_картки_гість": 1},
    }
    ok = fm.validate_stats({"match": True}, gap)
    assert ok is not None
    assert ok["жовті_картки"] == 3


def test_stats_search_queries_adds_sofascore_for_corners():
    qs = fm.stats_search_queries(
        {
            "home": "Djurgarden",
            "away": "AIK",
            "date": "16/08/2026",
            "league": "Allsvenskan",
            "missing": ["кутові_господар", "фоли_господар"],
        }
    )
    blob = " ".join(qs).lower()
    assert "sofascore" in blob
    assert "foxsports" in blob


def test_complete_row_totals_from_sides():
    df = pd.DataFrame(
        [
            {
                "ліга": "Allsvenskan",
                "дата": "16/08/2026",
                "господар": "Degerfors",
                "гість": "Goteborg",
                "результат": "3:0",
                "жовті_картки_господар": 2,
                "жовті_картки_гість": 1,
                "жовті_картки": pd.NA,
            }
        ]
    )
    out = fm.complete_row_totals(df)
    assert int(out.iloc[0]["жовті_картки"]) == 3


def test_verify_and_fill_second_round_fills_remaining(tmp_path):
    df = pd.DataFrame(
        [
            {
                "ліга": "Allsvenskan",
                "дата": "16/08/2026",
                "господар": "Djurgarden",
                "гість": "AIK",
                "результат": "1:3",
                "кутові_господар": pd.NA,
                "кутові_гість": pd.NA,
                "кутові": pd.NA,
                "фоли_господар": pd.NA,
                "фоли_гість": pd.NA,
                "фоли": pd.NA,
            }
        ]
    )
    urls = iter(
        [
            "https://example.com/a",
            "https://example.com/b",
            "https://example.com/c",
            "https://example.com/d",
        ]
    )

    def search_fn(query: str):
        return [{"title": query, "link": next(urls), "snippet": ""}]

    def fetch_fn(url: str):
        return f"Djurgarden 1-3 AIK stats {url}"

    def decide_fn(**kwargs):
        if kwargs["url"].endswith("/a"):
            return {"match": True, "corners_home": 5, "corners_away": 1}
        return {"match": True, "fouls_home": 15, "fouls_away": 17}

    filled, inv = fm.verify_and_fill(
        df,
        as_of=pd.Timestamp("2026-08-18"),
        search_fn=search_fn,
        fetch_fn=fetch_fn,
        decide_fn=decide_fn,
        cache_path=tmp_path / "missing_data.json",
        live=True,
        max_rounds=3,
    )
    assert int(filled.iloc[0]["кутові"]) == 6
    assert int(filled.iloc[0]["фоли"]) == 32
    assert inv["verification"]["remaining_fields"] == 0


def test_apply_inventory_does_not_overwrite_existing():
    df = pd.DataFrame(
        [
            {
                "ліга": "Allsvenskan",
                "дата": "14/08/2026",
                "господар": "Elfsborg",
                "гість": "Vasteras SK",
                "результат": "3:0",
                "кутові_господар": 7,
                "кутові_гість": pd.NA,
                "кутові": pd.NA,
            }
        ]
    )
    inv = {
        "gaps": [
            {
                "key": fm.row_key("14/08/2026", "Elfsborg", "Vasteras SK"),
                "filled": {"кутові_господар": 1, "кутові_гість": 4, "кутові": 5},
            }
        ]
    }
    out = fm.apply_inventory(df, inv)
    assert int(out.iloc[0]["кутові_господар"]) == 7
    assert int(out.iloc[0]["кутові_гість"]) == 4


def test_fill_stats_json_first_then_claude_mock(tmp_path):
    df = pd.DataFrame(
        [
            {
                "ліга": "Allsvenskan",
                "дата": "14/08/2026",
                "господар": "Elfsborg",
                "гість": "Vasteras SK",
                "результат": "3:0",
                "кутові_господар": pd.NA,
                "кутові_гість": pd.NA,
                "кутові": pd.NA,
                "фоли_господар": pd.NA,
                "фоли_гість": pd.NA,
                "фоли": pd.NA,
            }
        ]
    )
    cache = tmp_path / "missing_data.json"
    inventory = fm.build_inventory(df, as_of=pd.Timestamp("2026-08-18"), path=cache)
    assert cache.exists()
    assert inventory["summary"]["pending"] == 1

    def decide_fn(**kwargs):
        assert "Elfsborg" in kwargs["page_text"]
        return {
            "match": True,
            "reason": "sofascore stats",
            "corners_home": 6,
            "corners_away": 2,
            "fouls_home": 11,
            "fouls_away": 9,
        }

    filled, report = fm.fill_stats(
        df,
        as_of=pd.Timestamp("2026-08-18"),
        search_fn=lambda q: [{"title": "stats", "link": "https://example.com/elf", "snippet": ""}],
        fetch_fn=lambda url: "Elfsborg 3-0 Vasteras SK corners 6-2 fouls 11-9",
        decide_fn=decide_fn,
        cache_path=cache,
        live=True,
    )
    assert int(filled.iloc[0]["кутові"]) == 8
    assert int(filled.iloc[0]["фоли"]) == 20
    dumped = fm.load_missing_json(cache)
    assert dumped["gaps"][0]["status"] == "filled"
    assert dumped["summary"]["filled"] == 1


def test_fill_gaps_rejects_when_claude_says_no_match(tmp_path):
    target = pd.DataFrame(
        [
            {
                "ліга": "Premier League",
                "дата": "22/08/2026",
                "господар": "Hull",
                "гість": "Man United",
            }
        ]
    )
    hist = pd.DataFrame(
        [
            {
                "ліга": "Premier League",
                "дата": "01/01/2026",
                "господар": "Arsenal",
                "гість": "Chelsea",
            }
        ]
    )
    extra, report = fm.fill_gaps(
        target,
        hist,
        search_fn=lambda q: [{"title": "odds", "link": "https://example.com/odds", "snippet": ""}],
        fetch_fn=lambda url: "Betting odds Hull vs United 1.90",
        decide_fn=lambda **kw: {"match": False, "reason": "odds page", "matches": []},
        cache_path=tmp_path / "missing.json",
    )
    assert extra.empty
    assert report["accepted_matches"] == 0
    assert report["rejected_pages"] >= 1


def _write_test_xlsx(path, mecze, preds=None):
    preds = preds if preds is not None else pd.DataFrame([{"liga": "x"}])
    with pd.ExcelWriter(path) as w:
        mecze.to_excel(w, sheet_name="Mecze_2026", index=False)
        preds.to_excel(w, sheet_name="Predykcje", index=False)


def test_verify_excel_fills_empty_cells_from_json_without_api(tmp_path):
    mecze = pd.DataFrame(
        [
            {
                "ліга": "Eliteserien",
                "дата": "16/08/2026",
                "господар": "Sarpsborg 08",
                "гість": "Sandefjord",
                "результат": "1:2",
                "фоли_господар": pd.NA,
                "фоли_гість": pd.NA,
                "фоли": pd.NA,
                "кутові_господар": 4,
                "кутові_гість": 5,
                "кутові": 9,
            }
        ]
    )
    xlsx = tmp_path / "predykcje_2026.xlsx"
    _write_test_xlsx(xlsx, mecze)
    cache = tmp_path / "missing_data.json"
    fm.save_missing_json(
        {
            "gaps": [
                {
                    "key": fm.row_key("16/08/2026", "Sarpsborg 08", "Sandefjord"),
                    "filled": {"фоли_господар": 12, "фоли_гість": 12, "фоли": 24},
                }
            ]
        },
        cache,
    )
    calls = {"n": 0}

    def search_fn(query: str):
        calls["n"] += 1
        return []

    out, _, report = fm.verify_exported_workbook(
        xlsx,
        json_path=cache,
        as_of=pd.Timestamp("2026-08-18"),
        search_fn=search_fn,
        live=True,
    )
    assert int(out.iloc[0]["фоли"]) == 24
    assert report["filled_from_json"] == 3
    assert report["filled_from_api"] == 0
    assert calls["n"] == 0


def test_verify_excel_uses_api_when_json_lacks_value(tmp_path):
    mecze = pd.DataFrame(
        [
            {
                "ліга": "Eliteserien",
                "дата": "16/08/2026",
                "господар": "Sarpsborg 08",
                "гість": "Sandefjord",
                "результат": "1:2",
                "фоли_господар": pd.NA,
                "фоли_гість": pd.NA,
                "фоли": pd.NA,
            }
        ]
    )
    xlsx = tmp_path / "predykcje_2026.xlsx"
    _write_test_xlsx(xlsx, mecze)
    cache = tmp_path / "missing_data.json"
    fm.save_missing_json({"gaps": []}, cache)
    out, _, report = fm.verify_exported_workbook(
        xlsx,
        json_path=cache,
        as_of=pd.Timestamp("2026-08-18"),
        search_fn=lambda q: [{"title": "box", "link": "https://example.com/sarp", "snippet": ""}],
        fetch_fn=lambda url: "Sarpsborg 08 1-2 Sandefjord fouls 12-12",
        decide_fn=lambda **kw: {"match": True, "fouls_home": 12, "fouls_away": 12},
        live=True,
        max_rounds=1,
    )
    assert int(out.iloc[0]["фоли"]) == 24
    assert report["filled_from_json"] == 0
    assert report["filled_from_api"] >= 1
