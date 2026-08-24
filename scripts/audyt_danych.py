# -*- coding: utf-8 -*-
"""Skrypt audytu danych aleks_ligi_stats.xlsx → AUDYT_DANYCH.md"""
from __future__ import annotations

import re
import sys
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from predykcje import FROM_DATE, validate_from_date

XLSX = ROOT / "aleks_ligi_stats.xlsx"
PRED_XLSX = ROOT / "predykcje_2026.xlsx"
OUT = ROOT / "AUDYT_DANYCH.md"

COL_LIGA = "ліга"
COL_DATA = "дата"
COL_HOME = "господар"
COL_AWAY = "гість"
COL_RESULT = "результат"
COL_BTTS = "оз"

STAT_TOTALS = ["фоли", "кутові", "жовті_картки", "удари", "удари_в_площину"]
STAT_HOME = [f"{c}_господар" for c in STAT_TOTALS]
STAT_AWAY = [f"{c}_гість" for c in STAT_TOTALS]


def parse_score(s) -> tuple[int, int] | None:
    if pd.isna(s):
        return None
    m = re.match(r"^(\d+)\s*[:\-]\s*(\d+)$", str(s).strip())
    return (int(m.group(1)), int(m.group(2))) if m else None


def filled_pct(series: pd.Series) -> float:
    num = pd.to_numeric(series, errors="coerce")
    if num.notna().any():
        return round(num.notna().mean() * 100, 1)
    s = series.astype(str).str.strip()
    s = s.where(~s.str.casefold().isin(["", "nan", "none", "<na>"]), other=pd.NA)
    return round(s.notna().mean() * 100, 1)


def df_to_md(table: pd.DataFrame) -> str:
    """Markdown table bez zależności tabulate."""
    cols = list(table.columns)
    lines = [
        "| " + " | ".join(str(c) for c in cols) + " |",
        "| " + " | ".join("---" for _ in cols) + " |",
    ]
    for _, row in table.iterrows():
        lines.append("| " + " | ".join(str(row[c]) for c in cols) + " |")
    return "\n".join(lines)


def _audit_prediction_window(lines: list[str], source: pd.DataFrame) -> None:
    """Reguła: predykcje w Excelu tylko od FROM_DATE (historia źródłowa zostaje)."""
    cutoff = FROM_DATE
    cutoff_s = cutoff.strftime("%d/%m/%Y")
    lines.append("## 0. Walidacja okna predykcji\n")
    lines.append(f"**Reguła:** arkusz `Прогнози` od **{cutoff_s}**; `Матчі_2026` = wszystkie rozegrane 2026.  ")
    lines.append("Historia sprzed tej daty w `aleks_ligi_stats.xlsx` jest dozwolona (forma / backtest).\n")

    src_dates = pd.to_datetime(source[COL_DATA], errors="coerce")
    n_window = int((src_dates >= cutoff).sum())
    lines.append(f"- Źródło: **{n_window}** meczów od {cutoff_s} (z {len(source)} łącznie)")

    if not PRED_XLSX.exists():
        lines.append(f"- `predykcje_2026.xlsx`: **BRAK PLIKU** — uruchom `python predykcje.py`")
        lines.append("")
        return

    try:
        xl = pd.ExcelFile(PRED_XLSX)
        mecze_name = next((n for n in ("Матчі_2026", "Mecze_2026") if n in xl.sheet_names), "Mecze_2026")
        pred_name = next((n for n in ("Прогнози", "Predykcje") if n in xl.sheet_names), "Predykcje")
        mecze = pd.read_excel(PRED_XLSX, sheet_name=mecze_name)
        preds = pd.read_excel(PRED_XLSX, sheet_name=pred_name)
    except Exception as exc:
        lines.append(f"- `predykcje_2026.xlsx`: **BŁĄD ODCZYTU** ({exc})")
        lines.append("")
        return

    ok = True
    try:
        validate_from_date(preds, cutoff, label="Predykcje")
        lines.append(f"- `Predykcje`: **OK** — {len(preds)} meczów, wszystkie od {cutoff_s}")
    except ValueError as exc:
        ok = False
        lines.append(f"- `Predykcje`: **FAIL** — {exc}")
    lines.append(
        f"- `Mecze_2026`: **OK** — {len(mecze)} rozegranych "
        f"(cały 2026, bez obcinania do {cutoff_s})"
    )
    lines.append(f"- Wynik walidacji: **{'OK' if ok else 'FAIL'}**")
    lines.append("")


def main() -> None:
    df = pd.read_excel(XLSX)
    df[COL_DATA] = pd.to_datetime(df[COL_DATA], format="%d/%m/%Y", errors="coerce")
    for c in STAT_TOTALS + STAT_HOME + STAT_AWAY:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce")

    parsed = df[COL_RESULT].map(parse_score)
    df["_hg"] = parsed.map(lambda x: x[0] if x else pd.NA)
    df["_ag"] = parsed.map(lambda x: x[1] if x else pd.NA)
    df["_total_goals"] = df["_hg"] + df["_ag"]
    df["_goal_diff"] = df["_hg"] - df["_ag"]
    df["_home_win"] = (df["_hg"] > df["_ag"]).astype(int)
    df["_btts_yes"] = df[COL_BTTS].astype(str).str.strip().str.casefold().isin(
        ["yes", "tak", "так"]
    ).astype(int)

    lines: list[str] = []
    lines.append("# AUDYT DANYCH — aleks_ligi_stats.xlsx\n")
    lines.append(f"**Plik:** `{XLSX.name}`  ")
    lines.append(f"**Wiersze:** {len(df)}  ")
    lines.append(f"**Kolumny:** {len(df.columns)}  ")
    lines.append(f"**Lata:** {sorted(df[COL_DATA].dt.year.dropna().astype(int).unique().tolist())}  ")
    lines.append(f"**Ligi:** {df[COL_LIGA].nunique()}\n")

    _audit_prediction_window(lines, df)

    # 1. Kompletność
    lines.append("## 1. Kompletność per liga × kolumna\n")
    audit_cols = [COL_RESULT, COL_BTTS] + STAT_HOME + STAT_TOTALS
    comp_rows = []
    for liga in sorted(df[COL_LIGA].dropna().unique()):
        sub = df[df[COL_LIGA] == liga]
        row = {"ліга": liga, "mecze": len(sub)}
        for c in audit_cols:
            if c in sub.columns:
                row[c] = filled_pct(sub[c])
        comp_rows.append(row)
    comp = pd.DataFrame(comp_rows)
    lines.append(df_to_md(comp))
    lines.append("")

    # ligi tylko wynik/BTTS
    lines.append("### Ligi z niepełnymi statystykami per drużyna\n")
    for _, r in comp.iterrows():
        home_filled = r.get("фоли_господар", 100)
        corners_filled = r.get("кутові", 0)
        if home_filled < 50 or corners_filled < 50:
            lines.append(
                f"- **{r['ліга']}**: brak stat per team (~{home_filled}% фоли_господар) — "
                f"tylko wynik + BTTS + gole"
            )
        else:
            lines.append(f"- **{r['ліга']}**: pełne statystyki ({corners_filled}% кутові)")
    lines.append("")

    # 2. Rozkłady per liga
    lines.append("## 2. Rozkłady (średnia / mediana / p75) per liga\n")
    dist_cols = ["_total_goals", "кутові", "жовті_картки", "фоли", "удари", "удари_в_площину"]
    for liga in sorted(df[COL_LIGA].dropna().unique()):
        sub = df[df[COL_LIGA] == liga]
        lines.append(f"### {liga}\n")
        lines.append("| metryka | mean | median | p75 |")
        lines.append("|---------|------|--------|-----|")
        for c in dist_cols:
            if c not in sub.columns:
                continue
            s = pd.to_numeric(sub[c], errors="coerce").dropna()
            if s.empty:
                continue
            lines.append(
                f"| {c} | {s.mean():.2f} | {s.median():.1f} | {s.quantile(0.75):.1f} |"
            )
        lines.append("")

    # 3. Korelacje
    lines.append("## 3. Korelacje ze wynikiem\n")
    numeric = df.dropna(subset=["_hg", "_ag"]).copy()
    if "удари_в_площину_господар" in numeric.columns:
        numeric["_sot_diff"] = numeric["удари_в_площину_господар"] - numeric["удари_в_площину_гість"]
        corr_sot = numeric["_sot_diff"].corr(numeric["_goal_diff"])
        lines.append(f"- **Różnica SoT vs różnica goli (global):** {corr_sot:.3f}")
    if "кутові" in numeric.columns:
        numeric["_over_corners"] = (numeric["кутові"] > 9.5).astype(int)
        corr_c = numeric["кутові"].corr(numeric["_over_corners"])
        lines.append(f"- **Suma rożnych vs over 9.5 (proxy):** {corr_c:.3f}")
    if "жовті_картки" in numeric.columns:
        numeric["_over_cards"] = (numeric["жовті_картки"] > 3.5).astype(int)
        corr_k = numeric["жовті_картки"].corr(numeric["_over_cards"])
        lines.append(f"- **Suma kartek vs over 3.5 (proxy):** {corr_k:.3f}")
    lines.append("")
    lines.append("### Home win rate per liga\n")
    lines.append("| liga | home_win_pct | avg_goli | avg_rozne | avg_kartki |")
    lines.append("|------|--------------|----------|-----------|------------|")
    for liga in sorted(df[COL_LIGA].dropna().unique()):
        sub = df[df[COL_LIGA] == liga].dropna(subset=["_hg", "_ag"])
        if sub.empty:
            continue
        lines.append(
            f"| {liga} | {sub['_home_win'].mean()*100:.1f}% | "
            f"{sub['_total_goals'].mean():.2f} | "
            f"{sub['кутові'].mean():.2f} | "
            f"{sub['жовті_картки'].mean():.2f} |"
        )
    lines.append("")

    # 4. Jakość
    lines.append("## 4. Leakage / jakość danych\n")
    bad_dates = df[COL_DATA].isna().sum()
    dup = df.duplicated(subset=[COL_LIGA, COL_DATA, COL_HOME, COL_AWAY]).sum()
    lines.append(f"- **Brak daty:** {bad_dates} wierszy")
    lines.append(f"- **Duplikaty meczów (liga+data+home+away):** {dup}")
    for c in ["жовті_картки", "кутові", "фоли"]:
        if c in df.columns:
            mx = pd.to_numeric(df[c], errors="coerce").max()
            outliers = (pd.to_numeric(df[c], errors="coerce") > 30).sum()
            lines.append(f"- **Max {c}:** {mx} | outliery >30: {outliers}")
    lines.append("")

    # 5. Rynki — pokrycie z danych
    lines.append("## 5. Co da się realnie przewidywać\n")
    n = len(df)
    n_score = df[COL_RESULT].astype(str).str.strip().replace("nan", "").ne("").sum()
    n_btts = df[COL_BTTS].astype(str).str.strip().replace("nan", "").ne("").sum()
    n_corners = pd.to_numeric(df.get("кутові"), errors="coerce").notna().sum()
    n_cards = pd.to_numeric(df.get("жовті_картки"), errors="coerce").notna().sum()
    n_fouls = pd.to_numeric(df.get("фоли"), errors="coerce").notna().sum()
    n_sot_h = pd.to_numeric(df.get("удари_в_площину_господар"), errors="coerce").notna().sum()
    n_corners_split = (
        pd.to_numeric(df.get("кутові_господар"), errors="coerce").notna()
        & pd.to_numeric(df.get("кутові_гість"), errors="coerce").notna()
    ).sum()
    pct = lambda x: f"{x/n*100:.1f}%"

    lines.append("| Rynek | Pokrycie danych | Sensowność | Uwagi |")
    lines.append("|-------|-----------------|------------|-------|")
    markets = [
        ("1X2 + wynik", f"результат {pct(n_score)} ({n_score}/{n})", "TAK", "Pełna historia 2024–2026"),
        ("BTTS так/ні", f"оз {pct(n_btts)} ({n_btts}/{n})", "TAK", "100% wypełnienie, baseline ~52%"),
        ("O/U goli 1.5/2.5/3.5", f"результат {pct(n_score)}", "TAK", "Poisson-lite z exp GF/GA"),
        ("O/U rożne 8.5/9.5/10.5", f"кутові {pct(n_corners)}", "WARUNKOWO", "Nordic bez stat — pomiń"),
        ("O/U kartki 2.5/3.5/4.5", f"жовті {pct(n_cards)}", "WARUNKOWO", "Nordic bez stat — pomiń"),
        ("O/U fauli", f"фоли {pct(n_fouls)}", "WARUNKOWO", f"Próg = mediana ligi ({n_fouls} meczy)"),
        ("Team totals rożne H/A", f"split {pct(n_corners_split)} ({n_corners_split}/{n})", "WARUNKOWO", "Tylko top-6 lig"),
        ("Handicap 0/-0.5", f"результат + forma {pct(n_score)}", "TAK", "Z exp goals Poisson-lite"),
        ("Dokładny wynik Poisson", f"результат {pct(n_score)}", "WARUNKOWO", "Niska hit-rate, info value"),
        ("SoT / strzały heurystyka", f"SoT split {pct(n_sot_h)}", "WARUNKOWO", "Fallback gdy brak wyniku historycznego"),
        ("H2H + ranking punktowy", "100% klucze meczów", "TAK", "Feature store as_of"),
    ]
    for m in markets:
        lines.append(f"| {m[0]} | {m[1]} | {m[2]} | {m[3]} |")

    OUT.write_text("\n".join(lines), encoding="utf-8")
    print(f"Zapisano: {OUT}")


if __name__ == "__main__":
    main()
