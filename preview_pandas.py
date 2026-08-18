# -*- coding: utf-8 -*-
"""Moduł podglądu danych projektu w pandas.

Źródła:
  - cache/matches_cache.json
  - footystats_mecze_*.xlsx
  - aleks_ligi_stats.xlsx

Przykład (Jupyter / skrypt):
  from preview_pandas import load_matches, load_aleks, preview_dataframe, summarize

  df = load_matches()
  preview_dataframe(df, n=10)
  summarize(df)

  aleks = load_aleks()
  if aleks is not None:
      preview_dataframe(aleks, n=10, prefix="preview_aleks_head")
"""
from __future__ import annotations

import json
import sys
from pathlib import Path
from typing import Any

import pandas as pd

ROOT = Path(__file__).resolve().parent
CACHE = ROOT / "cache" / "matches_cache.json"
ALEKS_XLSX = ROOT / "aleks_ligi_stats.xlsx"
ALEKS_CSV = ROOT / "aleks_ligi_stats.csv"  # legacy fallback


def _safe_print(text: str) -> None:
    """Print odporny na kodowanie konsoli Windows."""
    try:
        if hasattr(sys.stdout, "reconfigure"):
            try:
                sys.stdout.reconfigure(errors="replace")
            except Exception:
                pass
        print(text)
    except UnicodeEncodeError:
        enc = sys.stdout.encoding or "utf-8"
        print(text.encode(enc, errors="replace").decode(enc, errors="replace"))


def resolve_aleks_path(root: Path | None = None, path: Path | None = None) -> Path | None:
    """Zwraca ścieżkę do aleks_ligi_stats.xlsx (albo legacy .csv)."""
    if path is not None:
        return path if path.exists() else None
    base = root or ROOT
    xlsx = base / "aleks_ligi_stats.xlsx"
    if xlsx.exists():
        return xlsx
    csv_path = base / "aleks_ligi_stats.csv"
    if csv_path.exists():
        return csv_path
    return None


def read_aleks_table(path: Path) -> pd.DataFrame:
    """Czyta Excel lub CSV Aleksa."""
    if path.suffix.lower() in {".xlsx", ".xls"}:
        return pd.read_excel(path)
    return pd.read_csv(path)


def save_aleks_table(df: pd.DataFrame, path: Path | None = None) -> Path:
    """Zapisuje tabelę Aleksa jako Excel (.xlsx)."""
    out = path or ALEKS_XLSX
    if out.suffix.lower() not in {".xlsx", ".xls"}:
        out = out.with_suffix(".xlsx")
    out.parent.mkdir(parents=True, exist_ok=True)
    try:
        df.to_excel(out, index=False, sheet_name="aleks_ligi_stats")
        return out
    except PermissionError:
        alt = out.with_name(out.stem + "_nowy.xlsx")
        df.to_excel(alt, index=False, sheet_name="aleks_ligi_stats")
        _safe_print(f"UWAGA: {out.name} zablokowany (zamknij Excel). Zapisano: {alt.name}")
        return alt


def load_matches(
    root: Path | None = None,
    cache_path: Path | None = None,
) -> pd.DataFrame:
    """Wczytuje mecze z cache JSON, a gdy brak — z najnowszego Excela."""
    base = root or ROOT
    cache = cache_path or (base / "cache" / "matches_cache.json")

    if cache.exists():
        data = json.loads(cache.read_text(encoding="utf-8"))
        matches = data.get("matches", data if isinstance(data, list) else [])
        df = pd.DataFrame(matches)
        _safe_print(f"Źródło: {cache.name} | wierszy: {len(df)} | shape: {df.shape}")
        return df

    excel_files = sorted(base.glob("footystats_mecze_*.xlsx"))
    if not excel_files:
        raise FileNotFoundError(
            "Brak cache i Excela — najpierw uruchom: python scrape_footystats.py"
        )
    path = excel_files[-1]
    df = pd.read_excel(path)
    _safe_print(f"Źródło: {path.name} | wierszy: {len(df)} | shape: {df.shape}")
    return df


def load_aleks(root: Path | None = None, csv_path: Path | None = None) -> pd.DataFrame | None:
    """Wczytuje aleks_ligi_stats.xlsx (fallback: .csv) albo zwraca None."""
    path = resolve_aleks_path(root=root, path=csv_path)
    if path is None:
        _safe_print("Brak aleks_ligi_stats.xlsx — uruchom: python export_aleks_stats.py")
        return None
    df = read_aleks_table(path)
    _safe_print(f"Źródło: {path.name} | wierszy: {len(df)} | shape: {df.shape}")
    return df


def rows_to_dataframe(rows: list[dict[str, Any]] | pd.DataFrame) -> pd.DataFrame:
    """Zamienia listę słowników lub DataFrame na DataFrame."""
    if isinstance(rows, pd.DataFrame):
        return rows.copy()
    return pd.DataFrame(rows)


def preview_dataframe(
    data: list[dict[str, Any]] | pd.DataFrame,
    n: int = 10,
    *,
    output_dir: Path | None = None,
    prefix: str = "preview_head",
    save: bool = True,
    show: bool = True,
) -> pd.DataFrame:
    """Buduje DataFrame, wypisuje head(n) i opcjonalnie zapisuje pliki podglądu.

    Zapisuje:
      - {prefix}.txt
      - {prefix}.csv
    """
    df = rows_to_dataframe(data)
    head = df.head(n)
    banner = f"df.head({n}) | shape={df.shape}"
    body = head.to_string(index=False)
    text = f"{banner}\n\n{body}\n"

    if save:
        out = Path(output_dir) if output_dir is not None else ROOT
        out.mkdir(parents=True, exist_ok=True)
        txt_path = out / f"{prefix}.txt"
        csv_path = out / f"{prefix}.csv"
        txt_path.write_text(text, encoding="utf-8")
        head.to_csv(csv_path, index=False, encoding="utf-8-sig")
        if show:
            _safe_print(f"Zapisano: {txt_path.name}, {csv_path.name}")

    if show:
        _safe_print(banner)
        _safe_print(body)

    return df


def summarize(df: pd.DataFrame, *, show: bool = True) -> dict[str, Any]:
    """Zwraca krótkie podsumowanie kolumn / dtypes / nulli."""
    info = {
        "shape": df.shape,
        "columns": df.columns.tolist(),
        "dtypes": {c: str(t) for c, t in df.dtypes.items()},
        "nulls": {c: int(v) for c, v in df.isna().sum().items()},
    }
    if show:
        _safe_print(f"shape: {info['shape']}")
        _safe_print(f"Kolumny: {info['columns']}")
        _safe_print("dtypes:")
        _safe_print(df.dtypes.to_string())
        nulls = df.isna().sum()
        if int(nulls.sum()) > 0:
            _safe_print("Nulls:")
            _safe_print(nulls[nulls > 0].to_string())
    return info


def filter_aleks(
    df: pd.DataFrame,
    *,
    liga: str | None = None,
    btts: str | None = None,
) -> pd.DataFrame:
    """Filtruje statystyki Aleksa po lidze i/lub ОЗ (BTTS: так/ні).

    Obsługuje zarówno stare nagłówki PL (`liga`, `btts`),
    jak i nowe UA (`ліга`, `оз`). Akceptuje yes/no oraz так/ні.
    """
    out = df.copy()
    liga_col = "ліга" if "ліга" in out.columns else "liga"
    btts_col = "оз" if "оз" in out.columns else "btts"
    if liga is not None and liga_col in out.columns:
        out = out[out[liga_col] == liga]
    if btts is not None and btts_col in out.columns:
        needle = str(btts).strip().casefold()
        yes_vals = {"yes", "tak", "так"}
        no_vals = {"no", "nie", "ні"}
        series = out[btts_col].astype(str).str.strip().str.casefold()
        if needle in yes_vals:
            out = out[series.isin(yes_vals)]
        elif needle in no_vals:
            out = out[series.isin(no_vals)]
        else:
            out = out[out[btts_col] == btts]
    return out


def preview_project(
    root: Path | None = None,
    n: int = 10,
    *,
    save: bool = True,
) -> tuple[pd.DataFrame, pd.DataFrame | None]:
    """Ładuje mecze + Aleks i robi standardowy podgląd projektu."""
    base = root or ROOT
    df = load_matches(root=base)
    preview_dataframe(df, n=n, output_dir=base, prefix="preview_head", save=save)
    summarize(df)

    aleks = load_aleks(root=base)
    if aleks is not None:
        preview_dataframe(
            aleks,
            n=n,
            output_dir=base,
            prefix="preview_aleks_head",
            save=save,
        )
        summarize(aleks)

    return df, aleks


def main() -> None:
    preview_project()


if __name__ == "__main__":
    main()
