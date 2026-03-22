"""Tests for EML parsing, folder paths, filtering, sorting, and preview extraction."""
import os
import tempfile
import unittest
from pathlib import Path

from email import policy
from email.parser import BytesParser

import PythonEmail_browse as eb

PROJECT_ROOT = Path(__file__).resolve().parent.parent
EXAMPLE = PROJECT_ROOT / "ExampleEmails"


class TestFolderDisplay(unittest.TestCase):
    def test_file_in_root_shows_slash(self):
        with tempfile.TemporaryDirectory() as tmp:
            p = Path(tmp) / "solo.eml"
            p.write_bytes(b"Subject: x\n\n")
            self.assertEqual(eb.folder_display_for_path(str(p), tmp), "/")

    def test_nested_folders_use_forward_slashes(self):
        with tempfile.TemporaryDirectory() as tmp:
            sub = Path(tmp) / "Inbox" / "archive"
            sub.mkdir(parents=True)
            p = sub / "m.eml"
            p.write_bytes(b"Subject: y\n\n")
            self.assertEqual(eb.folder_display_for_path(str(p), tmp), "Inbox/archive")


class TestLoadEmailRow(unittest.TestCase):
    def test_welcome_plain_no_attachments(self):
        path = EXAMPLE / "Inbox" / "welcome.eml"
        row = eb.load_email_row(str(path), str(PROJECT_ROOT))
        self.assertIsNotNone(row)
        self.assertEqual(row["folder"], "ExampleEmails/Inbox")
        self.assertEqual(row["from"], "alice@example.com")
        self.assertIn("Welcome", row["subject"])
        self.assertEqual(row["attach_label"], "No")
        self.assertFalse(row["has_attachments"])
        self.assertIn("plain-text message", row["search_blob"])
        self.assertIn("alice", row["search_blob"])

    def test_report_has_attachment(self):
        path = EXAMPLE / "Inbox" / "report.eml"
        row = eb.load_email_row(str(path), str(PROJECT_ROOT))
        self.assertIsNotNone(row)
        self.assertEqual(row["attach_label"], "Yes")
        self.assertTrue(row["has_attachments"])
        self.assertIn("summary.txt", row["search_blob"])

    def test_html_multipart_search_blob_includes_plain_and_cleaned_html(self):
        path = EXAMPLE / "Inbox" / "html_note.eml"
        row = eb.load_email_row(str(path), str(PROJECT_ROOT))
        self.assertIsNotNone(row)
        self.assertIn("rendered", row["search_blob"])
        self.assertIn("plain fallback", row["search_blob"].lower())


class TestFilterAndSort(unittest.TestCase):
    def setUp(self):
        self.rows = [
            {
                "folder": "A",
                "date": "2024-01-01 00:00",
                "from": "a@test",
                "subject": "alpha",
                "attach_label": "No",
                "has_attachments": False,
                "sort_key": 100.0,
                "path": "/x",
                "search_blob": "a@test alpha uniquebodyone",
            },
            {
                "folder": "B",
                "date": "2024-06-01 00:00",
                "from": "b@test",
                "subject": "beta",
                "attach_label": "Yes",
                "has_attachments": True,
                "sort_key": 200.0,
                "path": "/y",
                "search_blob": "b@test beta uniquebodytwo",
            },
        ]

    def test_filter_by_subject(self):
        f = {"Folder": "", "Date": "", "From": "", "Subject": "alp", "Attachments": ""}
        out = eb.filter_emails(self.rows, f)
        self.assertEqual(len(out), 1)
        self.assertEqual(out[0]["subject"], "alpha")

    def test_filter_attachments_yes(self):
        f = {c: "" for c in eb.COLUMNS}
        f["Attachments"] = "yes"
        out = eb.filter_emails(self.rows, f)
        self.assertEqual(len(out), 1)
        self.assertEqual(out[0]["subject"], "beta")

    def test_sorted_by_date_descending(self):
        f = {c: "" for c in eb.COLUMNS}
        out = eb.filtered_sorted_emails(self.rows, "Date", True, f)
        self.assertEqual([r["sort_key"] for r in out], [200.0, 100.0])

    def test_global_search_matches_body_substring(self):
        f = {c: "" for c in eb.COLUMNS}
        out = eb.filter_emails(self.rows, f, global_search="uniquebodytwo")
        self.assertEqual(len(out), 1)
        self.assertEqual(out[0]["subject"], "beta")

    def test_global_search_combines_with_column_filter(self):
        f = {c: "" for c in eb.COLUMNS}
        f["From"] = "a@"
        out = eb.filter_emails(self.rows, f, global_search="uniquebodyone")
        self.assertEqual(len(out), 1)
        self.assertEqual(out[0]["subject"], "alpha")


class TestMessageHelpers(unittest.TestCase):
    def test_clean_html_strips_tags(self):
        self.assertEqual(eb.clean_html("<p>Hi</p>").strip(), "Hi")

    def test_prepare_html_strips_assets_and_empty_wrappers(self):
        html = (
            '<div><img src="https://x/y.png" alt=""></div>'
            '<div class="spacer"><br/></div>'
            '<p>Visible text</p>'
        )
        out = eb.prepare_html_for_mail_preview(html)
        self.assertNotIn("img", out.lower())
        self.assertIn("Visible text", out)

    def test_clean_html_drops_scripts_and_keeps_paragraphs(self):
        raw = (
            "<html><body><script>bad()</script>"
            "<p>First</p><p>Second</p><br/>Line</body></html>"
        )
        out = eb.clean_html(raw)
        self.assertNotIn("bad", out)
        self.assertNotIn("<p>", out)
        self.assertIn("First", out)
        self.assertIn("Second", out)

    def test_message_has_attachments_false_for_plain(self):
        path = EXAMPLE / "Inbox" / "welcome.eml"
        with open(path, "rb") as f:
            msg = BytesParser(policy=policy.default).parse(f)
        self.assertFalse(eb.message_has_attachments(msg))

    def test_extract_body_plain(self):
        path = EXAMPLE / "Inbox" / "welcome.eml"
        with open(path, "rb") as f:
            msg = BytesParser(policy=policy.default).parse(f)
        body, parts, html = eb.extract_body_and_attachments(msg)
        self.assertIsNone(html)
        self.assertIn("plain-text message", body)
        self.assertEqual(parts, [])

    def test_extract_multipart_alternative_includes_html(self):
        path = EXAMPLE / "Inbox" / "html_note.eml"
        with open(path, "rb") as f:
            msg = BytesParser(policy=policy.default).parse(f)
        plain, parts, html = eb.extract_body_and_attachments(msg)
        self.assertIsNotNone(html)
        self.assertIn("rendered", plain)
        self.assertIn("<html>", html.lower())


if __name__ == "__main__":
    unittest.main()
