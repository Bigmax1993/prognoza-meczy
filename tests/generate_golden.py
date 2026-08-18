# -*- coding: utf-8 -*-
"""Generuje golden CSV dla testów regresyjnych."""
from __future__ import annotations

import csv
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

import scrape_footystats as sf  # noqa: E402
from tests.fixtures.sample_html import SAMPLE_MATCHES_HTML  # noqa: E402

OUT = Path(__file__).parent / "fixtures" / "golden" / "sample_mecze.csv"

COLUMNS = [
    "Data",
    "Kraj",
    "Liga",
    "Godzina",
    "Gospodarz",
    "Forma gospodarza",
    "Gość",
    "Forma gościa",
    "Kurs 1",
    "Kurs X",
    "Kurs 2",
    "Status",
    "Link",
]


def main() -> None:
    _, rows = sf.parse_matches(SAMPLE_MATCHES_HTML)
    OUT.parent.mkdir(parents=True, exist_ok=True)
    with OUT.open("w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=COLUMNS, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(rows)
    print(f"wrote {len(rows)} rows -> {OUT}")


if __name__ == "__main__":
    main()
