import scrape_footystats as sf
from tests.fixtures.sample_html import (
    CLOUDFLARE_HTML,
    CLOUDFLARE_RU_HTML,
    SAMPLE_MATCHES_HTML,
)


def test_detects_english_cloudflare_challenge():
    assert sf._is_cloudflare_block(CLOUDFLARE_HTML, "Just a moment...") is True


def test_detects_russian_cloudflare_challenge():
    assert sf._is_cloudflare_block(CLOUDFLARE_RU_HTML, "Один момент…") is True


def test_real_page_is_not_cloudflare():
    assert sf._is_cloudflare_block(SAMPLE_MATCHES_HTML, "FootyStats Test") is False


def test_attention_required_title():
    assert sf._is_cloudflare_block("<html></html>", "Attention Required! | Cloudflare") is True
