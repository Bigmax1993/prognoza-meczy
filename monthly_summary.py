# -*- coding: utf-8 -*-
"""Miesięczne średnie statystyk per drużyna.

Wymaga kolumn per drużyna (фоли_господар, фоли_гість, ...) z enrich_scores.py.

Uruchomienie:
  python monthly_summary.py
  python monthly_summary.py --year 2026
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

import pandas as pd

from preview_pandas import ALEKS_XLSX, read_aleks_table, resolve_aleks_path

ROOT = Path(__file__).resolve().parent
OUT_XLSX = ROOT / "podsumowanie_miesiaca.xlsx"

COL_LIGA = "ліга"
COL_DATA = "дата"
COL_HOME = "господар"
COL_AWAY = "гість"

STAT_BASES = ["фоли", "кутові", "жовті_картки", "удари", "удари_в_площину"]
HOME_SUFFIX = "_господар"
AWAY_SUFFIX = "_гість"


def _team_stat_columns(base: str) -> tuple[str, str]:
    return f"{base}{HOME_SUFFIX}", f"{base}{AWAY_SUFFIX}"


def load_data(csv_path: Path | None = None) -> pd.DataFrame:
    path = Path(csv_path) if csv_path is not None else resolve_aleks_path()
    if path is None:
        raise FileNotFoundError(ALEKS_XLSX)
    df = read_aleks_table(path)
    df[COL_DATA] = pd.to_datetime(df[COL_DATA], format="%d/%m/%Y", errors="coerce")
    return df


def compute_monthly_team_summary(
    df: pd.DataFrame,
    *,
    year: int | None = None,
) -> pd.DataFrame:
    """Średnie statystyk per drużyna, liga i miesiąc (YYYY-MM)."""
    data = df.copy()
    if year is not None:
        data = data[data[COL_DATA].dt.year == year]
    data = data.dropna(subset=[COL_DATA])
    if data.empty:
        return pd.DataFrame()

    data["miesiac"] = data[COL_DATA].dt.to_period("M").astype(str)

    appearances: list[dict] = []
    for _, m in data.iterrows():
        for team_col, suffix in ((COL_HOME, HOME_SUFFIX), (COL_AWAY, AWAY_SUFFIX)):
            row: dict = {
                "druzyna": m[team_col],
                COL_LIGA: m[COL_LIGA],
                "miesiac": m["miesiac"],
            }
            has_stat = False
            for base in STAT_BASES:
                col = f"{base}{suffix}"
                val = m[col] if col in m.index else pd.NA
                row[base] = pd.to_numeric(val, errors="coerce")
                if pd.notna(row[base]):
                    has_stat = True
            if has_stat:
                appearances.append(row)

    if not appearances:
        return pd.DataFrame()

    long_df = pd.DataFrame(appearances)
    agg: dict = {"mecze": ("druzyna", "count")}
    for base in STAT_BASES:
        if base in long_df.columns and long_df[base].notna().any():
            agg[f"srednia_{base}"] = (base, "mean")

    summary = (
        long_df.groupby(["druzyna", COL_LIGA, "miesiac"], as_index=False)
        .agg(**agg)
        .sort_values(["miesiac", COL_LIGA, "druzyna"])
        .reset_index(drop=True)
    )

    for col in summary.columns:
        if col.startswith("srednia_"):
            summary[col] = summary[col].round(2)
    return summary


def export_excel(summary: pd.DataFrame, path: Path | None = None) -> Path:
    out = path or OUT_XLSX
    summary.to_excel(out, index=False, sheet_name="Podsumowanie_miesiaca")
    return out


def main(argv: list[str] | None = None) -> None:
    parser = argparse.ArgumentParser(description="Miesięczne średnie statystyk per drużyna")
    parser.add_argument("--year", type=int, default=None, help="Filtr roku (np. 2026)")
    parser.add_argument("--csv", type=Path, default=None, help="Ścieżka do .xlsx/.csv")
    parser.add_argument("--out", type=Path, default=OUT_XLSX)
    args = parser.parse_args(argv)

    df = load_data(args.csv)
    summary = compute_monthly_team_summary(df, year=args.year)
    if summary.empty:
        raise SystemExit("Brak danych do podsumowania (czy uruchomiono enrich_scores.py?)")

    path = export_excel(summary, args.out)
    try:
        if hasattr(sys.stdout, "reconfigure"):
            sys.stdout.reconfigure(errors="replace")
        print(f"Wierszy podsumowania: {len(summary)}")
        print(f"Zapisano: {path}")
    except UnicodeEncodeError:
        print(f"Rows: {len(summary)} -> {path}")


if __name__ == "__main__":
    main()
