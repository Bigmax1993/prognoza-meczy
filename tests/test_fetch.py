from types import SimpleNamespace
from unittest.mock import MagicMock, patch

import scrape_footystats as sf


def test_fetch_with_requests_success(monkeypatch):
    mock_resp = SimpleNamespace(status_code=200, text="<html>ok matches</html>")
    monkeypatch.setattr(
        "requests.get",
        lambda *a, **k: mock_resp,
    )
    # force import path used inside function
    with patch("requests.get", return_value=mock_resp):
        html = sf.fetch_with_requests("https://example.com")
    assert html == "<html>ok matches</html>"


def test_fetch_with_requests_cloudflare(monkeypatch):
    mock_resp = SimpleNamespace(status_code=403, text="Just a moment...")
    with patch("requests.get", return_value=mock_resp):
        assert sf.fetch_with_requests("https://example.com") is None


def test_fetch_with_requests_exception():
    with patch("requests.get", side_effect=RuntimeError("timeout")):
        assert sf.fetch_with_requests("https://example.com") is None


def test_fetch_html_uses_first_success(tmp_paths, sample_html):
    with (
        patch.object(sf, "fetch_with_playwright", return_value=sample_html),
        patch.object(sf, "fetch_with_curl_cffi", return_value=None) as curl_mock,
        patch.object(sf, "fetch_with_requests", return_value=None) as req_mock,
    ):
        html, source = sf.fetch_html("https://example.com")
    assert html == sample_html
    assert source == "playwright"
    curl_mock.assert_not_called()
    req_mock.assert_not_called()
    assert tmp_paths["html_cache"].exists()


def test_fetch_html_falls_back_to_curl(tmp_paths, sample_html):
    with (
        patch.object(sf, "fetch_with_playwright", return_value=None),
        patch.object(sf, "fetch_with_curl_cffi", return_value=sample_html),
        patch.object(sf, "fetch_with_requests", return_value=None),
    ):
        html, source = sf.fetch_html("https://example.com")
    assert source == "curl_cffi"
    assert "league" in html


def test_fetch_html_all_fail(tmp_paths):
    with (
        patch.object(sf, "fetch_with_playwright", return_value=None),
        patch.object(sf, "fetch_with_curl_cffi", return_value=None),
        patch.object(sf, "fetch_with_requests", return_value=None),
    ):
        try:
            sf.fetch_html("https://example.com")
            assert False, "expected SystemExit"
        except SystemExit as exc:
            assert exc.code == 1


def test_playwright_launch_args_are_headless():
    args = sf._playwright_launch_args()
    assert "--disable-gpu" in args
    assert "--no-sandbox" in args
    assert any(a.startswith("--window-position") for a in args)
    # headless jest parametrem launch/persistent_context, nie musi byc w args
    assert "--disable-dev-shm-usage" in args


def test_cookie_consent_selectors_and_texts_defined():
    assert "#onetrust-accept-btn-handler" in sf._COOKIE_BUTTON_SELECTORS
    assert "Accept all" in sf._COOKIE_BUTTON_TEXTS
    assert "Принять" in sf._COOKIE_BUTTON_TEXTS
    assert "Akceptuj" in sf._COOKIE_BUTTON_TEXTS
    assert callable(sf._accept_cookies_and_consents)

def test_cookies_for_playwright_from_dict_like_jar():
    session = SimpleNamespace(
        cookies={"cf_clearance": "token123", "other": "x"},
    )
    cookies = sf._cookies_for_playwright(session, "https://footystats.org/ru/")
    names = {c["name"] for c in cookies}
    assert "cf_clearance" in names
    assert all(c["domain"] == "footystats.org" for c in cookies)


def test_cookies_for_playwright_from_cookie_objects():
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


def test_fetch_with_curl_cffi_saves_cookies(tmp_paths, sample_html):
    cookie = SimpleNamespace(name="cf_clearance", value="xyz", domain="footystats.org", path="/")
    session = SimpleNamespace(cookies=SimpleNamespace(jar=[cookie]))

    with patch.object(sf, "_curl_cffi_session", return_value=(session, sample_html)):
        html = sf.fetch_with_curl_cffi("https://footystats.org/ru/")

    assert html == sample_html
    loaded = sf.load_cookies_cache()
    assert loaded is not None
    assert loaded[0]["name"] == "cf_clearance"
