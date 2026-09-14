#!/usr/bin/env python3
"""Forward new Librus messages by email.

The program is intended to be run periodically by a systemd timer.  It keeps
Librus API keys on disk, records discovered messages in SQLite, and marks a
message as sent only after SMTP accepts it.
"""

from __future__ import annotations

import argparse
import base64
import json
import logging
from logging.handlers import RotatingFileHandler
import os
from dataclasses import dataclass
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import formataddr, formatdate
from pathlib import Path
import smtplib
import sqlite3
import ssl
import sys
import time
from typing import Iterable

import msal
import msal_extensions
from librus_apix.client import Token, new_client
from librus_apix.exceptions import AuthorizationError, TokenError, TokenKeyError
from librus_apix.messages import (
    Message,
    get_received,
    message_content,
)


APP_DIR = Path(__file__).resolve().parent
SMTP_SCOPES = ["https://outlook.office.com/SMTP.Send"]
DEFAULT_CLIENT_ID = ""
DEFAULT_AUTHORITY = "https://login.microsoftonline.com/common"
DEFAULT_SENDER = ""


@dataclass(frozen=True)
class Account:
    username: str
    password: str
    name: str
    recipients: tuple[str, ...]
    token_file: Path


@dataclass(frozen=True)
class PendingMessage:
    key: str
    href: str
    title: str
    author: str
    date: str


class MessageStore:
    def __init__(self, path: Path):
        path.parent.mkdir(parents=True, exist_ok=True, mode=0o700)
        os.chmod(path.parent, 0o700)
        self.connection = sqlite3.connect(path)
        self.connection.execute(
            """
            CREATE TABLE IF NOT EXISTS messages (
                account TEXT NOT NULL,
                message_key TEXT NOT NULL,
                href TEXT NOT NULL,
                title TEXT NOT NULL,
                author TEXT NOT NULL,
                message_date TEXT NOT NULL,
                status TEXT NOT NULL DEFAULT 'pending',
                discovered_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
                sent_at TEXT,
                last_error TEXT,
                PRIMARY KEY (account, message_key)
            )
            """
        )
        self.connection.commit()
        os.chmod(path, 0o600)

    def close(self) -> None:
        self.connection.close()

    def discover(self, account: str, messages: Iterable[Message]) -> int:
        before = self.connection.total_changes
        with self.connection:
            for message in messages:
                if not message.unread or not message.href:
                    continue
                self.connection.execute(
                    """
                    INSERT OR IGNORE INTO messages
                        (account, message_key, href, title, author, message_date)
                    VALUES (?, ?, ?, ?, ?, ?)
                    """,
                    (
                        account,
                        message.href,
                        message.href,
                        message.title,
                        message.author,
                        message.date,
                    ),
                )
        return self.connection.total_changes - before

    def pending(self, account: str) -> list[PendingMessage]:
        rows = self.connection.execute(
            """
            SELECT message_key, href, title, author, message_date
            FROM messages
            WHERE account = ? AND status = 'pending'
            ORDER BY discovered_at, message_key
            """,
            (account,),
        ).fetchall()
        return [PendingMessage(*row) for row in rows]

    def mark_sent(self, account: str, key: str) -> None:
        with self.connection:
            self.connection.execute(
                """
                UPDATE messages
                SET status = 'sent', sent_at = CURRENT_TIMESTAMP, last_error = NULL
                WHERE account = ? AND message_key = ?
                """,
                (account, key),
            )

    def mark_error(self, account: str, key: str, error: Exception) -> None:
        with self.connection:
            self.connection.execute(
                """
                UPDATE messages SET last_error = ?
                WHERE account = ? AND message_key = ?
                """,
                (f"{type(error).__name__}: {error}"[:500], account, key),
            )


class SmtpTokenProvider:
    def __init__(self, client_id: str, authority: str, cache_path: Path):
        cache_path.parent.mkdir(parents=True, exist_ok=True, mode=0o700)
        self.cache_path = cache_path
        self.cache = msal_extensions.PersistedTokenCache(
            msal_extensions.FilePersistence(str(cache_path))
        )
        self.app = msal.PublicClientApplication(
            client_id=client_id,
            authority=authority,
            token_cache=self.cache,
        )

    def acquire(self, *, interactive: bool = False, force_refresh: bool = False) -> str:
        accounts = self.app.get_accounts()
        result = None
        if accounts:
            result = self.app.acquire_token_silent(
                scopes=SMTP_SCOPES,
                account=accounts[0],
                force_refresh=force_refresh,
            )
        if (not result or "access_token" not in result) and interactive:
            result = self.app.acquire_token_interactive(scopes=SMTP_SCOPES)
        if not result or "access_token" not in result:
            detail = result.get("error_description", "no cached Outlook token") if result else "no cached Outlook token"
            raise RuntimeError(
                f"SMTP OAuth unavailable: {detail}. Run once with --bootstrap-smtp."
            )
        if self.cache_path.exists():
            os.chmod(self.cache_path, 0o600)
        return result["access_token"]


def configure_logging(log_path: Path) -> None:
    log_path.parent.mkdir(parents=True, exist_ok=True)
    formatter = logging.Formatter("%(asctime)s - %(levelname)s - %(message)s")
    file_handler = RotatingFileHandler(
        log_path, maxBytes=5_000_000, backupCount=3, encoding="utf-8"
    )
    file_handler.setFormatter(formatter)
    stream_handler = logging.StreamHandler()
    stream_handler.setFormatter(formatter)
    logging.basicConfig(level=logging.INFO, handlers=[file_handler, stream_handler], force=True)
    os.chmod(log_path, 0o600)


def load_config(config_path: Path) -> list[Account]:
    config_path = config_path.expanduser().resolve()
    with config_path.open(encoding="utf-8") as config_file:
        raw = json.load(config_file)
    entries = raw.get("accounts") if isinstance(raw, dict) else raw
    if not isinstance(entries, list) or not entries:
        raise ValueError("Configuration must contain a non-empty accounts list")

    accounts = []
    for entry in entries:
        if not isinstance(entry, dict):
            raise ValueError("Every account must be an object")
        required = ("username", "password", "konto", "recipients", "token_file")
        missing = [field for field in required if not entry.get(field)]
        if missing:
            raise ValueError(f"Missing fields for an account: {', '.join(missing)}")
        recipients = entry["recipients"]
        if not isinstance(recipients, list) or not all(
            isinstance(address, str) and "@" in address for address in recipients
        ):
            raise ValueError(f"Invalid recipients for account {entry['konto']}")
        password = os.path.expandvars(str(entry["password"]))
        if password.startswith("$"):
            raise ValueError(f"Unresolved password variable for account {entry['konto']}")
        token_file = Path(entry["token_file"])
        if not token_file.is_absolute():
            token_file = config_path.parent / token_file
        accounts.append(
            Account(
                username=str(entry["username"]),
                password=password,
                name=str(entry["konto"]),
                recipients=tuple(recipients),
                token_file=token_file,
            )
        )
    return accounts


def configure_timeout(client) -> None:
    original_request = client._session.request

    def request_with_timeout(method, url, **kwargs):
        kwargs.setdefault("timeout", (10, 30))
        return original_request(method, url, **kwargs)

    client._session.request = request_with_timeout


def read_token(path: Path) -> str | None:
    try:
        token = path.read_text(encoding="utf-8").strip()
        return token or None
    except FileNotFoundError:
        return None


def save_token(path: Path, api_key: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_suffix(path.suffix + ".tmp")
    temporary.write_text(api_key, encoding="utf-8")
    os.chmod(temporary, 0o600)
    os.replace(temporary, path)


def fetch_headers(client) -> list[Message]:
    return get_received(client, page=0)


def connect_librus(account: Account):
    api_key = read_token(account.token_file)
    if api_key:
        token_age_seconds = max(0, int(time.time() - account.token_file.stat().st_mtime))
        client = new_client(token=Token(API_Key=api_key))
        configure_timeout(client)
        try:
            headers = fetch_headers(client)
            logging.info(
                "Reused cached Librus session for %s (token age: %s)",
                account.name,
                format_duration(token_age_seconds),
            )
            return client, headers
        except (AuthorizationError, TokenError, TokenKeyError):
            logging.info(
                "Cached Librus session expired for %s after %s",
                account.name,
                format_duration(token_age_seconds),
            )

    client = new_client()
    configure_timeout(client)
    client.get_token(account.username, account.password)
    save_token(account.token_file, client.token.API_Key)
    logging.info("Authenticated with Librus for %s", account.name)
    return client, fetch_headers(client)


def format_duration(seconds: int) -> str:
    days, remainder = divmod(seconds, 86_400)
    hours, remainder = divmod(remainder, 3_600)
    minutes, seconds = divmod(remainder, 60)
    parts = []
    if days:
        parts.append(f"{days}d")
    if hours or days:
        parts.append(f"{hours}h")
    if minutes or hours or days:
        parts.append(f"{minutes}m")
    parts.append(f"{seconds}s")
    return " ".join(parts)


def create_email_message(
    message: PendingMessage,
    content: str,
    account_name: str,
    sender: str,
    recipients: tuple[str, ...],
) -> MIMEMultipart:
    body = (
        f"Title: {message.title}\n"
        f"From: {message.author}\n"
        f"Date: {message.date}\n"
        f"Content:\n{content}\n"
        f"{'-' * 50}\n"
    )
    email = MIMEMultipart()
    email["Subject"] = f"Librus {account_name}: {message.title}"
    email["From"] = formataddr(("Librus AutoNotifier", sender))
    email["To"] = ", ".join(recipients)
    email["Date"] = formatdate(localtime=True)
    email.attach(MIMEText(body, "plain", "utf-8"))
    return email


def send_email(
    email: MIMEMultipart,
    sender: str,
    recipients: tuple[str, ...],
    token_provider: SmtpTokenProvider,
) -> None:
    access_token = token_provider.acquire()
    server = smtplib.SMTP("smtp-mail.outlook.com", 587, timeout=30)
    try:
        server.ehlo()
        server.starttls(context=ssl.create_default_context())
        server.ehlo()

        def authenticate(token: str):
            auth = f"user={sender}\x01auth=Bearer {token}\x01\x01"
            encoded = base64.b64encode(auth.encode("utf-8")).decode("ascii")
            return server.docmd("AUTH", "XOAUTH2 " + encoded)

        code, response = authenticate(access_token)
        if code != 235:
            code, response = authenticate(token_provider.acquire(force_refresh=True))
        if code != 235:
            text = response.decode(errors="replace") if isinstance(response, bytes) else str(response)
            raise smtplib.SMTPAuthenticationError(code, text)

        refused = server.sendmail(sender, list(recipients), email.as_string())
        if refused:
            raise smtplib.SMTPRecipientsRefused(refused)
    finally:
        try:
            server.quit()
        except Exception:
            server.close()


def process_account(
    account: Account,
    store: MessageStore,
    token_provider: SmtpTokenProvider,
    sender: str,
) -> bool:
    try:
        client, headers = connect_librus(account)
        discovered = store.discover(account.username, headers)
        pending = store.pending(account.username)
        logging.info(
            "%s: discovered %d new, %d pending", account.name, discovered, len(pending)
        )
    except Exception as error:
        logging.error("Failed to check account %s: %s", account.name, error)
        return False

    success = True
    for message in pending:
        try:
            details = message_content(client, message.href)
            email = create_email_message(
                message, details.content, account.name, sender, account.recipients
            )
            send_email(email, sender, account.recipients, token_provider)
            store.mark_sent(account.username, message.key)
            logging.info("%s: forwarded message %s", account.name, message.key)
        except Exception as error:
            store.mark_error(account.username, message.key, error)
            logging.error(
                "%s: failed to forward message %s: %s",
                account.name,
                message.key,
                error,
            )
            success = False
    return success


def process_accounts(
    accounts: list[Account],
    store: MessageStore,
    token_provider: SmtpTokenProvider,
    sender: str,
    account_delay: float,
    retry_delay: float,
) -> int:
    failed_accounts = []
    for index, account in enumerate(accounts):
        if not process_account(account, store, token_provider, sender):
            failed_accounts.append(account)
        if index + 1 < len(accounts):
            time.sleep(account_delay)

    if not failed_accounts:
        return 0

    logging.warning(
        "Retrying %d failed account(s) after all accounts were processed",
        len(failed_accounts),
    )
    time.sleep(retry_delay)
    still_failed = []
    for index, account in enumerate(failed_accounts):
        if not process_account(account, store, token_provider, sender):
            still_failed.append(account)
        if index + 1 < len(failed_accounts):
            time.sleep(account_delay)
    return len(still_failed)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--check",
        action="store_true",
        help="validate Outlook and Librus access without reading content or sending email",
    )
    parser.add_argument(
        "--bootstrap-smtp",
        action="store_true",
        help="perform interactive Outlook authorization and exit",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    config_path = Path(os.getenv("LIBRUS_CONFIG_FILE", APP_DIR / "librus_konta.json"))
    state_dir = Path(os.getenv("LIBRUS_STATE_DIR", APP_DIR / "state"))
    log_path = Path(os.getenv("LIBRUS_LOG_FILE", APP_DIR / "librus_email.log"))
    cache_path = Path(
        os.getenv("SMTP_TOKEN_CACHE_PATH", APP_DIR / "smtp_token_cache.bin")
    )
    client_id = os.getenv("SMTP_CLIENT_ID", DEFAULT_CLIENT_ID)
    authority = os.getenv("SMTP_AUTHORITY", DEFAULT_AUTHORITY)
    sender = os.getenv("SMTP_EMAIL", DEFAULT_SENDER)
    account_delay = float(os.getenv("LIBRUS_ACCOUNT_DELAY", "10"))
    retry_delay = float(os.getenv("LIBRUS_RETRY_DELAY", "10"))

    configure_logging(log_path)
    if not client_id or not sender:
        logging.critical("SMTP_CLIENT_ID and SMTP_EMAIL must be configured")
        return 1
    token_provider = SmtpTokenProvider(client_id, authority, cache_path)
    if args.bootstrap_smtp:
        token_provider.acquire(interactive=True)
        logging.info("Outlook authorization cache initialized")
        return 0

    try:
        accounts = load_config(config_path)
        token_provider.acquire()
    except Exception as error:
        logging.critical("Startup failed: %s", error)
        return 1

    if args.check:
        failures = 0
        for index, account in enumerate(accounts):
            try:
                _, headers = connect_librus(account)
                unread = sum(1 for message in headers if message.unread)
                logging.info(
                    "%s: check OK, %d message(s), %d unread",
                    account.name,
                    len(headers),
                    unread,
                )
            except Exception as error:
                failures += 1
                logging.error("%s: check failed: %s", account.name, error)
            if index + 1 < len(accounts):
                time.sleep(account_delay)
        if failures:
            logging.error("Check completed with %d failed account(s)", failures)
            return 1
        logging.info("Check completed successfully for %d account(s)", len(accounts))
        return 0

    try:
        store = MessageStore(state_dir / "messages.sqlite3")
    except Exception as error:
        logging.critical("Failed to open message state: %s", error)
        return 1

    try:
        failures = process_accounts(
            accounts,
            store,
            token_provider,
            sender,
            account_delay,
            retry_delay,
        )
    finally:
        store.close()

    if failures:
        logging.error("Run completed with %d failed account(s)", failures)
        return 1
    logging.info("Run completed successfully for %d account(s)", len(accounts))
    return 0


if __name__ == "__main__":
    sys.exit(main())
