"""Testy wysyłki Gmail (bez SMTP)."""

from __future__ import annotations

from email.message import EmailMessage
from pathlib import Path

import pandas as pd
import pytest

import send_mail as sm


def test_build_message_attaches_xlsx(tmp_path: Path):
    xlsx = tmp_path / "predykcje_2026.xlsx"
    pd.DataFrame({"a": [1]}).to_excel(xlsx, index=False)
    msg = sm.build_message(sender="a@gmail.com", to="b@gmail.com", path=xlsx)
    assert msg["From"] == "a@gmail.com"
    assert msg["To"] == "b@gmail.com"
    payloads = [p.get_filename() for p in msg.iter_attachments()]
    assert payloads == ["predykcje_2026.xlsx"]


def test_send_excel_uses_env_and_smtp_hook(tmp_path: Path, monkeypatch):
    xlsx = tmp_path / "predykcje_2026.xlsx"
    pd.DataFrame({"a": [1]}).to_excel(xlsx, index=False)
    monkeypatch.setenv("GMAIL_USER", "svinchak1993@gmail.com")
    monkeypatch.setenv("GMAIL_APP_PASSWORD", "abcd efgh ijkl mnop")
    monkeypatch.setenv("MAIL_TO", "Swinczakaleksy@gmail.com")
    seen: dict[str, object] = {}

    def fake_smtp(msg: EmailMessage, *, sender: str, to: str) -> None:
        seen["msg"] = msg
        seen["sender"] = sender
        seen["to"] = to

    info = sm.send_excel(xlsx, smtp_send=fake_smtp)
    assert info["to"] == "Swinczakaleksy@gmail.com"
    assert seen["sender"] == "svinchak1993@gmail.com"
    assert isinstance(seen["msg"], EmailMessage)


def test_mail_config_requires_keys(monkeypatch):
    monkeypatch.delenv("GMAIL_USER", raising=False)
    monkeypatch.delenv("GMAIL_APP_PASSWORD", raising=False)
    monkeypatch.delenv("MAIL_TO", raising=False)
    with pytest.raises(RuntimeError, match="Brak zmiennych"):
        sm.mail_config()
