import scrape_footystats as sf


def test_parse_match_date(sample_html):
    match_date, rows = sf.parse_matches(sample_html)
    assert match_date == "11/08"
    assert len(rows) == 3


def test_parse_first_match_fields(sample_html):
    _, rows = sf.parse_matches(sample_html)
    first = rows[0]
    assert first["Kraj"] == "Европа"
    assert first["Liga"] == "Лига чемпионов УЕФА"
    assert first["Godzina"] == "18:00"
    assert first["Gospodarz"] == "Жальгирис"
    assert first["Forma gospodarza"] == "1.60"
    assert first["Gość"] == "Динамо Загреб"
    assert first["Forma gościa"] == "2.33"
    assert first["Kurs 1"] == "6.35"
    assert first["Kurs X"] == "4.62"
    assert first["Kurs 2"] == "1.39"
    assert first["Status"] == "incomplete"
    assert first["Link"].endswith("/ru/europe/team-a-vs-team-b-h2h-stats")


def test_parse_strips_country_from_league_title(sample_html):
    _, rows = sf.parse_matches(sample_html)
    england = [r for r in rows if r["Kraj"] == "Англия"]
    assert len(england) == 1
    assert england[0]["Liga"] == "Кубок Англии"
    assert england[0]["Gospodarz"] == "Холлен"


def test_parse_empty_html():
    match_date, rows = sf.parse_matches("<html><body></body></html>")
    assert rows == []
    assert match_date  # fallback na dzisiejszą datę
