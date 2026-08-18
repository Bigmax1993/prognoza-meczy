# -*- coding: utf-8 -*-
"""Testy jednostkowe — czyste funkcje bez I/O sieciowego i bez pełnego pipeline."""
from __future__ import annotations

from types import SimpleNamespace

import pytest

import scrape_footystats as sf
from tests.fixtures.sample_html import (
    CLOUDFLARE_HTML,
    CLOUDFLARE_RU_HTML,
    SAMPLE_MATCHES_HTML,
)


class TestCloudflareDetection:
    def test_english_challenge(self):
        assert sf._is_cloudflare_block(CLOUDFLARE_HTML, "Just a moment...") is True

    def test_russian_challenge(self):
        assert sf._is_cloudflare_block(CLOUDFLARE_RU_HTML, "Один момент…") is True

    def test_attention_required_title(self):
        assert (
            sf._is_cloudflare_block("<html></html>", "Attention Required! | Cloudflare")
            is True
        )

    def test_real_page_not_blocked(self):
        assert sf._is_cloudflare_block(SAMPLE_MATCHES_HTML, "FootyStats Test") is False


class TestTextHelpers:
    def test_text_strips_and_joins(self):
        soup_span = SimpleNamespace(get_text=lambda strip=True: "  A  B  ")
        # BeautifulSoup-compatible: use real tag via parse
        from bs4 import BeautifulSoup

        el = BeautifulSoup("<span>  Жальгирис  </span>", "lxml").span
        assert sf._text(el) == "Жальгирис"

    def test_text_none_safe(self):
        assert sf._text(None) == ""

    def test_odd_value_reads_leading_number(self):
        from bs4 import BeautifulSoup

        span = BeautifulSoup(
            '<span class="ac">6.35<span class="hover-modal-content">Home</span></span>',
            "lxml",
        ).span
        assert sf._odd_value(span) == "6.35"

    def test_odd_value_empty(self):
        assert sf._odd_value(None) == ""


class TestLeagueTitle:
    def test_strips_country_prefix(self):
        assert (
            sf._strip_country_prefix("Англия - Кубок Англии", "Англия")
            == "Кубок Англии"
        )

    def test_keeps_title_without_prefix(self):
        assert sf._strip_country_prefix("Лига чемпионов УЕФА", "Европа") == (
            "Лига чемпионов УЕФА"
        )


class TestCookiesForPlaywright:
    def test_from_dict_like_jar(self):
        session = SimpleNamespace(cookies={"cf_clearance": "token123", "other": "x"})
        cookies = sf._cookies_for_playwright(session, "https://footystats.org/ru/")
        names = {c["name"] for c in cookies}
        assert "cf_clearance" in names
        assert all(c["domain"] == "footystats.org" for c in cookies)

    def test_from_cookie_objects(self):
        cookie = SimpleNamespace(
            name="cf_clearance",
            value="abc",
            domain=".footystats.org",
            path="/",
        )
        session = SimpleNamespace(cookies=SimpleNamespace(jar=[cookie]))
        cookies = sf._cookies_for_playwright(session, "https://footystats.org/ru/")
        assert cookies == [
            {
                "name": "cf_clearance",
                "value": "abc",
                "domain": ".footystats.org",
                "path": "/",
            }
        ]


class TestJsonCacheUnit:
    def test_missing_returns_none(self, tmp_paths):
        assert sf.load_json_cache(tmp_paths["matches_cache"]) is None

    def test_corrupted_returns_none(self, tmp_paths):
        tmp_paths["matches_cache"].write_text("{not-json", encoding="utf-8")
        assert sf.load_json_cache(tmp_paths["matches_cache"]) is None

    @pytest.mark.parametrize("ttl", [1, 30, 60])
    def test_cache_age_ok_fresh(self, ttl):
        from datetime import datetime

        now = datetime.now().isoformat(timespec="seconds")
        assert sf._cache_age_ok(now, ttl) is True
