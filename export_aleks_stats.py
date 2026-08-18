# -*- coding: utf-8 -*-
"""
Eksport CSV wyłącznie z https://footystats.org/ru/

Używa:
  - json     → zapis/odczyt cache (cache/aleks_stats_cache.json)
  - logging  → logi do logs/aleks_export_YYYY-MM-DD.log

Ligi Aleksa + kolumny (nagłówki po ukraińsku):
  ліга, дата, господар, гість, результат,
  фоли_господар, фоли_гість, фоли,
  кутові_господар, кутові_гість, кутові, ...
  оз
"""
from __future__ import annotations

import json
import logging
import re
import sys
import time
import unicodedata
from datetime import datetime
from pathlib import Path

import pandas as pd
from bs4 import BeautifulSoup

import scrape_footystats as sf

ROOT = Path(__file__).resolve().parent
OUT = ROOT / "aleks_ligi_stats.xlsx"
CACHE_DIR = ROOT / "cache"
LOGS_DIR = ROOT / "logs"
ALEKS_CACHE = CACHE_DIR / "aleks_stats_cache.json"
SOURCE_URL = "https://footystats.org/ru/"

logger = logging.getLogger("aleks_export")

LEAGUE_RULES: list[tuple[str, list[str], list[str]]] = [
    ("Premier League", ["премьер-лига", "premier league"], ["ган", "ghana", "u18", "u19", "u21", "u23"]),
    ("La Liga", ["ла лига", "la liga"], ["2", "сегунда", "segunda"]),
    ("Serie A", ["серия а", "serie a"], ["бразил", "brazil", "женщин", "women", "b ", "б "]),
    ("Bundesliga 2", ["2. бундеслига", "бундеслига 2", "2. bundesliga", "2 bundesliga"], []),
    ("Bundesliga", ["бундеслига", "bundesliga"], ["2.", "u19", "u17", "3.", "регион", "oberliga"]),
    ("Eredivisie", ["эредивиз", "eredivisie"], ["women", "женщин"]),
    ("Super League", ["суперлига", "super league"], ["китай", "china", "грец", "greece", "турц", "turkey"]),
    ("Allsvenskan", ["аллсвенскан", "allsvenskan"], ["dam"]),
    ("Eliteserien", ["элитсери", "eliteserien"], []),
]

SUPER_LEAGUE_COUNTRIES = ("швейцар", "switzerland", "swiss")

# Nagłówki CSV po ukraińsku
FIELDS = [
    "ліга",
    "дата",
    "господар",
    "гість",
    "результат",
    "фоли_господар",
    "фоли_гість",
    "фоли",
    "кутові_господар",
    "кутові_гість",
    "кутові",
    "жовті_картки_господар",
    "жовті_картки_гість",
    "жовті_картки",
    "удари_господар",
    "удари_гість",
    "удари",
    "удари_в_площину_господар",
    "удари_в_площину_гість",
    "удари_в_площину",
    "оз",
]


# ---------------------------------------------------------------------------
# logging
# ---------------------------------------------------------------------------

def setup_logging(level: str = "INFO") -> Path:
    """Konfiguruje logging → plik w logs/ + konsola."""
    LOGS_DIR.mkdir(parents=True, exist_ok=True)
    log_path = LOGS_DIR / f"aleks_export_{datetime.now():%Y-%m-%d}.log"

    level_no = getattr(logging, level.upper(), logging.INFO)
    logger.handlers.clear()
    logger.setLevel(level_no)
    logger.propagate = False

    fmt = logging.Formatter(
        fmt="%(asctime)s | %(levelname)-8s | %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )
    fh = logging.FileHandler(log_path, encoding="utf-8")
    fh.setLevel(level_no)
    fh.setFormatter(fmt)
    logger.addHandler(fh)

    sh = logging.StreamHandler(sys.stdout)
    sh.setLevel(level_no)
    sh.setFormatter(fmt)
    logger.addHandler(sh)

    logger.info("Start eksportu Aleks | logi: %s", log_path)
    return log_path


# ---------------------------------------------------------------------------
# json cache
# ---------------------------------------------------------------------------

def save_json_cache(path: Path, payload: dict) -> Path:
    """Zapis cache przez bibliotekę json."""
    CACHE_DIR.mkdir(parents=True, exist_ok=True)
    path.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    logger.info("Zapisano cache JSON: %s (%s bajtow)", path.name, path.stat().st_size)
    return path


def load_json_cache(path: Path) -> dict | None:
    """Odczyt cache przez bibliotekę json."""
    try:
        if not path.exists():
            logger.info("Brak cache: %s", path.name)
            return None
        data = json.loads(path.read_text(encoding="utf-8"))
        logger.info("Wczytano cache JSON: %s", path.name)
        return data
    except Exception as exc:
        logger.warning("Nie udalo sie wczytac cache %s: %s", path, exc)
        return None


def save_aleks_cache(
    match_date: str,
    source: str,
    raw_count: int,
    rows: list[dict],
) -> Path:
    payload = {
        "cached_at": datetime.now().isoformat(timespec="seconds"),
        "url": SOURCE_URL,
        "match_date": match_date,
        "fetch_source": source,
        "raw_matches": raw_count,
        "count": len(rows),
        "rows": rows,
    }
    return save_json_cache(ALEKS_CACHE, payload)


# ---------------------------------------------------------------------------
# helpers / parse
# ---------------------------------------------------------------------------

def clean(value: object) -> str:
    if value is None:
        return ""
    text = str(value).strip()
    text = unicodedata.normalize("NFKD", text)
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    out = []
    for ch in text:
        if ch.isascii() and (ch.isalnum() or ch in " ._/-:"):
            out.append(ch)
    return re.sub(r"\s+", " ", "".join(out)).strip()


def clean_btts(value: object) -> str:
    """BTTS jako tak/ні (ukraiński) — bez stripowania cyrylicy."""
    if value is None:
        return ""
    v = str(value).strip().casefold()
    if v in {"yes", "tak", "так", "y", "1", "true"}:
        return "так"
    if v in {"no", "nie", "ні", "n", "0", "false"}:
        return "ні"
    # czasem % z H2H — nie mapujemy na tak/nie
    if re.search(r"\d", v):
        return ""
    return ""


def _norm(text: str) -> str:
    return (text or "").casefold().strip()


def match_aleks_league(kraj: str, liga: str) -> str | None:
    blob = f"{kraj} {liga}"
    n = _norm(blob)
    liga_n = _norm(liga)

    ordered = [r for r in LEAGUE_RULES if r[0] == "Bundesliga 2"] + [
        r for r in LEAGUE_RULES if r[0] != "Bundesliga 2"
    ]

    for label, needles, banned in ordered:
        if not any(needle in n or needle in liga_n for needle in needles):
            continue
        if any(b and b in n for b in banned):
            continue
        if label == "Premier League":
            if "английс" not in liga_n and "premier league" not in liga_n:
                if "англи" not in n and "england" not in n:
                    if "премьер-лига" in liga_n and "англи" not in n:
                        continue
        if label == "Serie A":
            if "бразил" in n:
                continue
        if label == "Super League":
            if not any(c in n for c in SUPER_LEAGUE_COUNTRIES):
                continue
        if label == "La Liga":
            if "испан" not in n and "spain" not in n and "ла лига" not in liga_n:
                continue
        if label == "Bundesliga" and ("2." in liga_n or liga_n.startswith("2 ")):
            continue
        return label
    return None


def _first_number(text: str) -> str:
    m = re.search(r"(\d+(?:[.,]\d+)?)", text.replace(",", "."))
    return m.group(1) if m else ""


def parse_h2h_stats(html: str) -> dict[str, str]:
    soup = BeautifulSoup(html or "", "lxml")
    text = soup.get_text("\n", strip=True)
    out = {
        "btts": "",
        "faule": "",
        "rozne": "",
        "zolte_kartki": "",
        "strzaly": "",
        "strzaly_w_stwor": "",
    }

    patterns = {
        "btts": [
            r"(?:BTTS|ОЗ|обе\s*команды[^%\n]{0,40}?заб)[^%\n]{0,40}?(\d+(?:[.,]\d+)?)\s*%",
            r"Both Teams[^%\n]{0,30}?(\d+(?:[.,]\d+)?)\s*%",
        ],
        "rozne": [
            r"(?:Corners|угловые|Угловые)[^%\n]{0,40}?(\d+(?:[.,]\d+)?)",
            r"AVG\s*Corners[^0-9]{0,20}(\d+(?:[.,]\d+)?)",
        ],
        "zolte_kartki": [
            r"(?:Yellow Cards|желт[^%\n]{0,20}карт)[^0-9]{0,30}(\d+(?:[.,]\d+)?)",
            r"Cards\s*/\s*match[^0-9]{0,20}(\d+(?:[.,]\d+)?)",
        ],
        "faule": [
            r"(?:Fouls|Фол[^%\n]{0,10})[^0-9]{0,30}(\d+(?:[.,]\d+)?)",
        ],
        "strzaly": [
            r"(?:Shots|Удар[^%\n]{0,15})(?!\s*on\s*Target)(?!\s*в\s*створ)[^0-9]{0,30}(\d+(?:[.,]\d+)?)",
        ],
        "strzaly_w_stwor": [
            r"(?:Shots on Target|удары?\s*в\s*створ|В створ)[^0-9]{0,30}(\d+(?:[.,]\d+)?)",
        ],
    }

    for key, pats in patterns.items():
        for pat in pats:
            m = re.search(pat, text, re.I | re.M)
            if m:
                out[key] = m.group(1).replace(",", ".")
                break

    for el in soup.select("[data-stat], .stat-box, .average, .w100"):
        label = _norm(el.get_text(" ", strip=True))
        val = _first_number(el.get_text(" ", strip=True))
        if not val:
            continue
        if "btts" in label or "оз" in label:
            out["btts"] = out["btts"] or val
        elif "foul" in label or "фол" in label:
            out["faule"] = out["faule"] or val
        elif "corner" in label or "угл" in label:
            out["rozne"] = out["rozne"] or val
        elif "yellow" in label or "желт" in label:
            out["zolte_kartki"] = out["zolte_kartki"] or val
        elif "on target" in label or "створ" in label:
            out["strzaly_w_stwor"] = out["strzaly_w_stwor"] or val
        elif "shot" in label or "удар" in label:
            out["strzaly"] = out["strzaly"] or val

    return out


def filter_aleks_matches(rows: list[dict]) -> list[dict]:
    out = []
    for r in rows:
        label = match_aleks_league(r.get("Kraj", ""), r.get("Liga", ""))
        if not label:
            continue
        out.append({**r, "_aleks_liga": label})
    logger.info("Filtr lig Aleksa: %s / %s meczow", len(out), len(rows))
    return out


def enrich_from_h2h(matches: list[dict], max_h2h: int = 80) -> list[dict]:
    enriched = []
    total = len(matches)
    for i, m in enumerate(matches):
        stats = {
            "btts": "",
            "faule": "",
            "rozne": "",
            "zolte_kartki": "",
            "strzaly": "",
            "strzaly_w_stwor": "",
        }
        link = m.get("Link") or ""
        if link and i < max_h2h:
            try:
                html = sf.fetch_with_playwright(link)
                if html:
                    stats = parse_h2h_stats(html)
                    logger.debug("H2H OK: %s", link)
                else:
                    logger.warning("H2H pusty/HTML zablokowany: %s", link)
                time.sleep(0.4)
            except Exception as exc:
                logger.exception("H2H fail %s: %s", link, exc)

        row = {
            "ліга": clean(m["_aleks_liga"]),
            "дата": clean(m.get("Data")),
            "господар": clean(m.get("Gospodarz")),
            "гість": clean(m.get("Gość")),
            "результат": clean(m.get("Wynik") or m.get("Status") or ""),
            "фоли_господар": "",
            "фоли_гість": "",
            "фоли": clean(stats.get("faule")),
            "кутові_господар": "",
            "кутові_гість": "",
            "кутові": clean(stats.get("rozne")),
            "жовті_картки_господар": "",
            "жовті_картки_гість": "",
            "жовті_картки": clean(stats.get("zolte_kartki")),
            "удари_господар": "",
            "удари_гість": "",
            "удари": clean(stats.get("strzaly")),
            "удари_в_площину_господар": "",
            "удари_в_площину_гість": "",
            "удари_в_площину": clean(stats.get("strzaly_w_stwor")),
            "оз": clean_btts(stats.get("btts")),
        }
        enriched.append(row)
        logger.info(
            "[%s/%s] %s: %s vs %s",
            i + 1,
            total,
            row["ліга"],
            row["господар"],
            row["гість"],
        )
    return enriched


def write_excel(rows: list[dict], path: Path) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    df = pd.DataFrame(rows, columns=FIELDS)
    try:
        df.to_excel(path, index=False, sheet_name="aleks_ligi_stats")
        out = path
    except PermissionError:
        out = path.with_name(path.stem + "_nowy.xlsx")
        df.to_excel(out, index=False, sheet_name="aleks_ligi_stats")
        logger.warning("Plik zablokowany — zapisano: %s", out.name)
    logger.info("Zapisano Excel: %s (%s wierszy)", out, len(rows))
    return out


# kompatybilność wsteczna
def write_csv(rows: list[dict], path: Path) -> Path:
    if path.suffix.lower() == ".csv":
        path = path.with_suffix(".xlsx")
    return write_excel(rows, path)


def main() -> None:
    log_path = setup_logging("INFO")
    # wspolne logi scrapera tez do pliku projektu
    sf.setup_logging("INFO")

    logger.info("Pobieram wylacznie: %s", SOURCE_URL)
    try:
        html, source = sf.fetch_html(SOURCE_URL)
    except SystemExit:
        # awaryjnie: sprobuj HTML z cache JSON scrapera (tez z FootyStats)
        cached = load_json_cache(sf.HTML_CACHE)
        if not cached or not cached.get("html"):
            logger.error("Brak zywego pobrania i brak html_cache.json")
            raise
        html = cached["html"]
        source = f"html_cache:{cached.get('source', '?')}"
        logger.warning("Uzyto awaryjnego cache HTML: %s", source)

    logger.info("Zrodlo pobrania: %s (%s bajtow)", source, len(html))
    # cache surowego HTML tez przez json (scraper)
    sf.save_html_cache(html, source)

    match_date, rows = sf.parse_matches(html)
    logger.info("Wszystkie mecze ze strony: %s (data %s)", len(rows), match_date)
    sf.save_matches_cache(match_date, rows, source)

    aleks = filter_aleks_matches(rows)
    if not aleks:
        seen = sorted({(r.get("Kraj"), r.get("Liga")) for r in rows})
        logger.error("Brak dopasowan lig Aleksa. Przyklad lig na stronie:")
        for k, l in seen[:40]:
            logger.error("  - %s | %s", k, l)
        write_csv([], OUT)
        save_aleks_cache(match_date, source, len(rows), [])
        raise SystemExit(2)

    logger.info("Dociagam H2H (FootyStats) pod 6 statystyk...")
    out_rows = enrich_from_h2h(aleks)
    path = write_csv(out_rows, OUT)
    save_aleks_cache(match_date, source, len(rows), out_rows)

    by_liga: dict[str, int] = {}
    for r in out_rows:
        by_liga[r["ліга"]] = by_liga.get(r["ліга"], 0) + 1

    logger.info("Gotowe: rows=%s out=%s log=%s", len(out_rows), path, log_path)
    for liga, n in sorted(by_liga.items()):
        logger.info("  %s: %s", liga, n)


if __name__ == "__main__":
    main()
