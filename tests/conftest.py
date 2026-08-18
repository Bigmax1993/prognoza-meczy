"""Wspólne fixture'y — izolowany katalog cache/logs dla każdego testu."""

from __future__ import annotations

import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

import scrape_footystats as sf  # noqa: E402
from tests.fixtures.sample_html import SAMPLE_MATCHES_HTML  # noqa: E402


@pytest.fixture
def tmp_paths(tmp_path, monkeypatch):
    """Przekierowuje cache/logs/output do tymczasowego katalogu."""
    cache_dir = tmp_path / "cache"
    logs_dir = tmp_path / "logs"
    cache_dir.mkdir()
    logs_dir.mkdir()

    monkeypatch.setattr(sf, "OUTPUT_DIR", tmp_path)
    monkeypatch.setattr(sf, "CACHE_DIR", cache_dir)
    monkeypatch.setattr(sf, "LOGS_DIR", logs_dir)
    monkeypatch.setattr(sf, "MATCHES_CACHE", cache_dir / "matches_cache.json")
    monkeypatch.setattr(sf, "COOKIES_CACHE", cache_dir / "cookies_cache.json")
    monkeypatch.setattr(sf, "HTML_CACHE", cache_dir / "html_cache.json")

    return {
        "root": tmp_path,
        "cache": cache_dir,
        "logs": logs_dir,
        "matches_cache": cache_dir / "matches_cache.json",
        "cookies_cache": cache_dir / "cookies_cache.json",
        "html_cache": cache_dir / "html_cache.json",
    }


@pytest.fixture
def sample_html() -> str:
    return SAMPLE_MATCHES_HTML


@pytest.fixture
def sample_rows(sample_html):
    _, rows = sf.parse_matches(sample_html)
    return rows


@pytest.fixture
def configured_logger(tmp_paths):
    path = sf.setup_logging("DEBUG")
    yield path
    sf.logger.handlers.clear()
