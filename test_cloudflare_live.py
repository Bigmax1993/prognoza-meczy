# -*- coding: utf-8 -*-
"""Jednorazowy test Cloudflare vs Playwright (FootyStats)."""
from __future__ import annotations

from pathlib import Path

import scrape_footystats as sf

OUT_DIR = Path(__file__).resolve().parent / "cache"


def main() -> None:
    sf.setup_logging("INFO")
    url = sf.URL
    sf.logger.info("TEST Cloudflare — URL: %s", url)

    html = sf.fetch_with_playwright(url)
    result = {
        "ok": bool(html),
        "len": len(html or ""),
        "cf_block": None,
        "title": "",
        "has_league": False,
        "has_match": False,
        "match_date": "",
        "matches": 0,
    }

    if html:
        title = sf._extract_title(html)
        result["title"] = title
        result["cf_block"] = sf._is_cloudflare_block(html, title)
        low = html.lower()
        result["has_league"] = "league" in low and ("div" in low)
        result["has_match"] = 'class="match' in html or "a.match" in low
        (OUT_DIR / "cf_test_page.html").write_text(html, encoding="utf-8")

        if not result["cf_block"]:
            match_date, rows = sf.parse_matches(html)
            result["match_date"] = match_date
            result["matches"] = len(rows)
            sf.save_html_cache(html, "cf_test_playwright")
            if rows:
                sf.save_matches_cache(match_date, rows, "cf_test_playwright")
            sf.logger.info(
                "PASS: Cloudflare OK | mecze=%s | data=%s",
                len(rows),
                match_date,
            )
        else:
            sf.logger.error("FAIL: nadal strona Cloudflare (title=%s)", title)
    else:
        sf.logger.error("FAIL: Playwright zwrocil None (Cloudflare / blad)")

    lines = [f"{k}={v}" for k, v in result.items()]
    (OUT_DIR / "cf_test_result.txt").write_text("\n".join(lines) + "\n", encoding="utf-8")
    for line in lines:
        print(line)

    raise SystemExit(0 if html and not result["cf_block"] else 1)


if __name__ == "__main__":
    main()
