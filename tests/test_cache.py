import json
from datetime import datetime, timedelta

import scrape_footystats as sf


def test_save_and_load_matches_cache(tmp_paths, sample_rows):
    path = sf.save_matches_cache("11/08", sample_rows, "test")
    assert path.exists()
    assert path == tmp_paths["matches_cache"]

    loaded = sf.load_matches_cache(ttl_minutes=30)
    assert loaded is not None
    match_date, rows = loaded
    assert match_date == "11/08"
    assert len(rows) == len(sample_rows)
    assert rows[0]["Gospodarz"] == sample_rows[0]["Gospodarz"]


def test_matches_cache_expires(tmp_paths, sample_rows):
    sf.save_matches_cache("11/08", sample_rows, "test")
    payload = json.loads(tmp_paths["matches_cache"].read_text(encoding="utf-8"))
    payload["cached_at"] = (datetime.now() - timedelta(hours=2)).isoformat(timespec="seconds")
    tmp_paths["matches_cache"].write_text(json.dumps(payload), encoding="utf-8")

    assert sf.load_matches_cache(ttl_minutes=30) is None


def test_save_and_load_cookies_cache(tmp_paths):
    cookies = [{"name": "cf_clearance", "value": "abc", "domain": "footystats.org", "path": "/"}]
    sf.save_cookies_cache(cookies)
    loaded = sf.load_cookies_cache(ttl_minutes=60)
    assert loaded == cookies


def test_cookies_cache_expires(tmp_paths):
    sf.save_cookies_cache([{"name": "a", "value": "1", "domain": "x", "path": "/"}])
    payload = json.loads(tmp_paths["cookies_cache"].read_text(encoding="utf-8"))
    payload["cached_at"] = (datetime.now() - timedelta(hours=5)).isoformat(timespec="seconds")
    tmp_paths["cookies_cache"].write_text(json.dumps(payload), encoding="utf-8")
    assert sf.load_cookies_cache(ttl_minutes=60) is None


def test_html_cache(tmp_paths, sample_html):
    sf.save_html_cache(sample_html, "unit-test")
    data = sf.load_json_cache(tmp_paths["html_cache"])
    assert data is not None
    assert data["source"] == "unit-test"
    assert data["html"] == sample_html
    assert data["length"] == len(sample_html)


def test_corrupted_json_cache(tmp_paths):
    tmp_paths["matches_cache"].write_text("{not-json", encoding="utf-8")
    assert sf.load_json_cache(tmp_paths["matches_cache"]) is None


def test_missing_cache_returns_none(tmp_paths):
    assert sf.load_matches_cache() is None
    assert sf.load_cookies_cache() is None
