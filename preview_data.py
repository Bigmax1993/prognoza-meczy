"""
Podgląd danych FootyStats — wariant skryptowy / Interactive Window.

Preferowane w Jupyter Lab: otwórz preview_data.ipynb
  python -m jupyter lab

Jak użyć tego pliku (.py):
  1. Otwórz preview_data.ipynb w Jupyter Lab (zalecane)
  2. Albo uruchom komórki (# %%) w Cursor / VS Code
  3. Albo: python preview_data.py
"""

# %% Importy
from preview_pandas import load_aleks, load_matches, preview_dataframe, summarize

# %% Wczytaj mecze (cache JSON albo Excel)
df = load_matches()

# %% Podgląd — df.head(10)
preview_dataframe(df, n=10)

# %% Info o kolumnach
summarize(df)

# %% Statystyki Aleksa (jeśli CSV istnieje)
aleks = load_aleks()
if aleks is not None:
    preview_dataframe(aleks, n=10, prefix="preview_aleks_head")
    summarize(aleks)
