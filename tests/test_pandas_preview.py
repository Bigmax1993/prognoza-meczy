import pandas as pd
import pytest

import preview_pandas as pp
import scrape_footystats as sf


@pytest.fixture
def preview_root(tmp_path, monkeypatch):
    """Izolowany katalog wyjścia dla preview_pandas."""
    monkeypatch.setattr(pp, "ROOT", tmp_path)
    return tmp_path


def test_load_matches_from_cache(preview_root, sample_rows):
    cache_dir = preview_root / "cache"
    cache_dir.mkdir()
    cache_path = cache_dir / "matches_cache.json"
    cache_path.write_text(
        '{"matches": [{"Gospodarz": "A", "Gość": "B"}]}',
        encoding="utf-8",
    )

    df = pp.load_matches(root=preview_root, cache_path=cache_path)
    assert isinstance(df, pd.DataFrame)
    assert len(df) == 1
    assert df.iloc[0]["Gospodarz"] == "A"


def test_load_matches_from_excel(preview_root):
    df_src = pd.DataFrame([{"Gospodarz": "X", "Gość": "Y"}])
    df_src.to_excel(preview_root / "footystats_mecze_01-01.xlsx", index=False)

    df = pp.load_matches(root=preview_root)
    assert len(df) == 1
    assert df.iloc[0]["Gospodarz"] == "X"


def test_preview_dataframe_saves_files(preview_root, sample_rows, capsys):
    df = pp.preview_dataframe(sample_rows, n=2, output_dir=preview_root, save=True)

    assert isinstance(df, pd.DataFrame)
    assert len(df) == len(sample_rows)
    assert (preview_root / "preview_head.txt").exists()
    assert (preview_root / "preview_head.csv").exists()

    txt = (preview_root / "preview_head.txt").read_text(encoding="utf-8")
    assert "df.head(2)" in txt
    assert sample_rows[0]["Gospodarz"] in txt

    out = capsys.readouterr().out
    assert "df.head(2)" in out


def test_summarize_returns_info(sample_rows):
    df = pd.DataFrame(sample_rows)
    info = pp.summarize(df, show=False)

    assert info["shape"] == df.shape
    assert info["columns"] == df.columns.tolist()
    assert set(info["dtypes"]) == set(df.columns)


def test_filter_aleks():
    df = pd.DataFrame(
        [
            {"liga": "Premier League", "btts": "yes"},
            {"liga": "La Liga", "btts": "no"},
            {"liga": "Premier League", "btts": "no"},
        ]
    )
    filtered = pp.filter_aleks(df, liga="Premier League", btts="yes")
    assert len(filtered) == 1
    assert filtered.iloc[0]["btts"] == "yes"


def test_preview_dataframe_scraper_wrapper(tmp_paths, sample_rows, capsys):
    """Scraper nadal eksponuje preview_dataframe() przez wrapper."""
    df = sf.preview_dataframe(sample_rows, n=2)
    assert isinstance(df, pd.DataFrame)
    assert (tmp_paths["root"] / "preview_head.txt").exists()
    assert (tmp_paths["root"] / "preview_head.csv").exists()

    out = capsys.readouterr().out
    assert "df.head(2)" in out
