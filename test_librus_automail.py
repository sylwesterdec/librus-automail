import json
import os
from pathlib import Path
import tempfile
import unittest
from unittest.mock import Mock, patch

from librus_apix.messages import Message

import librus_automail as app


class LibrusAutoMailTests(unittest.TestCase):
    def test_format_duration(self):
        self.assertEqual(app.format_duration(59), "59s")
        self.assertEqual(app.format_duration(3_661), "1h 1m 1s")
        self.assertEqual(app.format_duration(90_061), "1d 1h 1m 1s")

    def test_load_config_resolves_environment_and_relative_token_path(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            config_path = root / "accounts.json"
            config_path.write_text(
                json.dumps(
                    {
                        "accounts": [
                            {
                                "username": "student",
                                "password": "$LIBRUS_TEST_PASSWORD",
                                "konto": "Student",
                                "recipients": ["parent@example.com"],
                                "token_file": "student.token",
                            }
                        ]
                    }
                ),
                encoding="utf-8",
            )
            with patch.dict(os.environ, {"LIBRUS_TEST_PASSWORD": "secret"}):
                accounts = app.load_config(config_path)

            self.assertEqual(accounts[0].password, "secret")
            self.assertEqual(accounts[0].token_file, (root / "student.token").resolve())

    def test_store_discovers_unread_message_only_once(self):
        with tempfile.TemporaryDirectory() as directory:
            store = app.MessageStore(Path(directory) / "messages.sqlite3")
            unread = Message("Teacher", "Title", "Date", "123", True, False)
            read = Message("Teacher", "Old", "Date", "456", False, False)

            self.assertEqual(store.discover("student", [unread, read]), 1)
            self.assertEqual(store.discover("student", [unread]), 0)
            self.assertEqual([item.key for item in store.pending("student")], ["123"])

            store.mark_sent("student", "123")
            self.assertEqual(store.pending("student"), [])
            store.close()

    def test_message_is_not_sent_twice(self):
        with tempfile.TemporaryDirectory() as directory:
            store = app.MessageStore(Path(directory) / "messages.sqlite3")
            account = app.Account(
                "student",
                "secret",
                "Student",
                ("parent@example.com",),
                Path(directory) / "student.token",
            )
            header = Message("Teacher", "Title", "Date", "123", True, False)
            details = Mock(content="Message body")
            provider = Mock()

            with (
                patch.object(app, "connect_librus", return_value=(Mock(), [header])),
                patch.object(app, "message_content", return_value=details),
                patch.object(app, "send_email") as send,
            ):
                self.assertTrue(app.process_account(account, store, provider, "sender@example.com"))
                self.assertTrue(app.process_account(account, store, provider, "sender@example.com"))

            send.assert_called_once()
            store.close()

    def test_failed_account_is_retried_after_all_other_accounts(self):
        first = app.Account("first", "secret", "First", ("a@example.com",), Path("a"))
        second = app.Account("second", "secret", "Second", ("b@example.com",), Path("b"))

        with (
            patch.object(app, "process_account", side_effect=[False, True, True]) as process,
            patch.object(app.time, "sleep") as sleep,
        ):
            failures = app.process_accounts(
                [first, second], Mock(), Mock(), "sender@example.com", 2, 10
            )

        self.assertEqual(failures, 0)
        self.assertEqual([call.args[0].username for call in process.call_args_list], ["first", "second", "first"])
        self.assertEqual([call.args[0] for call in sleep.call_args_list], [2, 10])

    def test_fetch_headers_reads_only_first_page(self):
        first = Message("Teacher", "First", "Date", "123", True, False)
        with patch.object(app, "get_received", return_value=[first]) as get_received:
            messages = app.fetch_headers(Mock())

        self.assertEqual(messages, [first])
        get_received.assert_called_once()
        self.assertEqual(get_received.call_args.kwargs["page"], 0)

    def test_failed_send_remains_pending_for_next_run(self):
        with tempfile.TemporaryDirectory() as directory:
            store = app.MessageStore(Path(directory) / "messages.sqlite3")
            account = app.Account(
                "student",
                "secret",
                "Student",
                ("parent@example.com",),
                Path(directory) / "student.token",
            )
            header = Message("Teacher", "Title", "Date", "123", True, False)
            details = Mock(content="Message body")

            with (
                patch.object(app, "connect_librus", return_value=(Mock(), [header])),
                patch.object(app, "message_content", return_value=details),
                patch.object(app, "send_email", side_effect=RuntimeError("SMTP down")),
            ):
                self.assertFalse(app.process_account(account, store, Mock(), "sender@example.com"))

            self.assertEqual([item.key for item in store.pending("student")], ["123"])
            store.close()

    def test_email_format_matches_current_application(self):
        message = app.PendingMessage("123", "123", "Title", "Teacher", "Date")
        email = app.create_email_message(
            message,
            "Message body",
            "Student",
            "sender@example.com",
            ("parent@example.com",),
        )
        payload = email.get_payload()[0].get_payload(decode=True).decode("utf-8")

        self.assertEqual(email["Subject"], "Librus Student: Title")
        self.assertIn("Title: Title\n", payload)
        self.assertIn("From: Teacher\n", payload)
        self.assertIn("Date: Date\n", payload)
        self.assertIn("Content:\nMessage body\n", payload)


if __name__ == "__main__":
    unittest.main()
