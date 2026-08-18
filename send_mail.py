# -*- coding: utf-8 -*-
"""Wysyłka predykcje_2026.xlsx przez Gmail SMTP (hasło do aplikacji)."""
from __future__ import annotations

import argparse
import os
import smtplib
from email.message import EmailMessage
from pathlib import Path

ROOT = Path(__file__).resolve().parent
OUT_XLSX = ROOT / "predykcje_2026.xlsx"
SMTP_HOST = "smtp.gmail.com"
SMTP_PORT = 587


def _env(name: str) -> str:
    return (os.environ.get(name) or "").strip().strip('"').strip("'")


def _hydrate_user_env() -> None:
    """Windows: dopisz GMAIL_* z User env, jeśli nie ma w procesie."""
    names = ("GMAIL_USER", "GMAIL_APP_PASSWORD", "MAIL_TO")
    try:
        import winreg
    except ImportError:
        return
    for name in names:
        if _env(name):
            continue
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Environment") as key:
                val, _ = winreg.QueryValueEx(key, name)
        except OSError:
            continue
        if val:
            os.environ[name] = str(val).strip()


def mail_config() -> tuple[str, str, str]:
    user = _env("GMAIL_USER")
    password = _env("GMAIL_APP_PASSWORD").replace(" ", "")
    to = _env("MAIL_TO")
    if not user or not password or not to:
        missing = [
            n
            for n, v in (
                ("GMAIL_USER", user),
                ("GMAIL_APP_PASSWORD", password),
                ("MAIL_TO", to),
            )
            if not v
        ]
        raise RuntimeError("Brak zmiennych: " + ", ".join(missing))
    return user, password, to


def build_message(
    *,
    sender: str,
    to: str,
    path: Path,
    subject: str | None = None,
) -> EmailMessage:
    xlsx = Path(path)
    if not xlsx.is_file():
        raise FileNotFoundError(f"Brak pliku: {xlsx}")
    msg = EmailMessage()
    msg["From"] = sender
    msg["To"] = to
    msg["Subject"] = subject or f"Prognozy meczów — {xlsx.name}"
    msg.set_content(
        "W załączniku aktualny plik prognoz (predykcje_2026.xlsx).\n"
        "Arkusze: Матчі_2026, Майбутні_матчі, Прогнози.\n"
    )
    msg.add_attachment(
        xlsx.read_bytes(),
        maintype="application",
        subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=xlsx.name,
    )
    return msg


def send_excel(
    path: Path | None = None,
    *,
    smtp_send=None,
) -> dict[str, str]:
    _hydrate_user_env()
    sender, password, to = mail_config()
    xlsx = Path(path) if path is not None else OUT_XLSX
    msg = build_message(sender=sender, to=to, path=xlsx)
    if smtp_send is None:
        with smtplib.SMTP(SMTP_HOST, SMTP_PORT, timeout=45) as smtp:
            smtp.starttls()
            smtp.login(sender, password)
            smtp.send_message(msg)
    else:
        smtp_send(msg, sender=sender, to=to)
    return {"from": sender, "to": to, "file": str(xlsx.resolve())}


def main(argv: list[str] | None = None) -> None:
    parser = argparse.ArgumentParser(description="Wyślij predykcje_2026.xlsx na Gmail")
    parser.add_argument("--plik", type=Path, default=OUT_XLSX, help="Ścieżka do xlsx")
    args = parser.parse_args(argv)
    info = send_excel(args.plik)
    print(f"Wysłano {Path(info['file']).name} → {info['to']}")


if __name__ == "__main__":
    main()
