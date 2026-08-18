import scrape_footystats as sf


def test_setup_logging_creates_file(tmp_paths):
    sf.logger.handlers.clear()
    path = sf.setup_logging("INFO")
    assert path.exists()
    assert path.parent == tmp_paths["logs"]
    assert path.name.startswith("scraper_")
    assert path.suffix == ".log"

    sf.logger.info("test-message-xyz")
    for handler in sf.logger.handlers:
        handler.flush()

    content = path.read_text(encoding="utf-8")
    assert "test-message-xyz" in content
    sf.logger.handlers.clear()


def test_setup_logging_sets_level(tmp_paths):
    sf.logger.handlers.clear()
    sf.setup_logging("WARNING")
    assert sf.logger.level == 30  # WARNING
    sf.logger.handlers.clear()
