# Unit tests for the pure helpers in common.py — the logic that decides
# which columns get matched, which addresses get suppressed, and which
# replies count as unsubscribes. Run with:  python -m unittest discover tests
import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from common import (  # noqa: E402
    build_header_map, col_index_to_letter, format_phone_us,
    is_unsubscribe_text, normalize_email, normalize_header, normalize_phone,
    split_name, strip_quoted_text,
)


class TestNormalizers(unittest.TestCase):
    def test_normalize_header(self):
        self.assertEqual(normalize_header("  First-Name_1 "), "first name 1")
        self.assertEqual(normalize_header("E-Mail!"), "e mail")
        self.assertEqual(normalize_header(None), "")

    def test_normalize_email(self):
        self.assertEqual(normalize_email(" John.Doe+x@Example.COM "), "john.doe+x@example.com")
        self.assertEqual(normalize_email("mailto:jay@ex.io (work)"), "jay@ex.io")
        self.assertEqual(normalize_email("not an email"), "")
        self.assertEqual(normalize_email(""), "")

    def test_normalize_phone_valid(self):
        self.assertEqual(normalize_phone("(555) 123-4567"), "5551234567")
        self.assertEqual(normalize_phone("1-555-123-4567"), "5551234567")

    def test_normalize_phone_invalid_kept_by_default(self):
        self.assertEqual(normalize_phone("55-1234"), "55-1234")

    def test_normalize_phone_invalid_dropped_for_match_sets(self):
        self.assertEqual(normalize_phone("55-1234", keep_original=False), "")
        self.assertEqual(normalize_phone("", keep_original=False), "")

    def test_format_phone_us(self):
        self.assertEqual(format_phone_us("5551234567"), "(555) 123-4567")
        self.assertEqual(format_phone_us(""), "")

    def test_split_name(self):
        self.assertEqual(split_name("john ronald smith"), ("John", "Ronald Smith"))
        self.assertEqual(split_name("cher"), ("Cher", ""))
        self.assertEqual(split_name(""), ("", ""))

    def test_col_index_to_letter(self):
        self.assertEqual(col_index_to_letter(0), "A")
        self.assertEqual(col_index_to_letter(25), "Z")
        self.assertEqual(col_index_to_letter(26), "AA")
        self.assertEqual(col_index_to_letter(27), "AB")


class TestHeaderMap(unittest.TestCase):
    def test_basic_mapping(self):
        m = build_header_map(["First Name", "Last Name", "Phone", "Email", "email_sent"])
        self.assertEqual(m["first_name"], 0)
        self.assertEqual(m["last_name"], 1)
        self.assertEqual(m["phone"], 2)
        self.assertEqual(m["email"], 3)
        self.assertEqual(m["email_sent"], 4)

    def test_email_sent_does_not_steal_email(self):
        # An "Email Sent" column must never be mistaken for the email column.
        m = build_header_map(["Email Sent", "Email Address"])
        self.assertEqual(m["email_sent"], 0)
        self.assertEqual(m["email"], 1)

    def test_mailing_address_is_not_email(self):
        m = build_header_map(["Mailing Address", "Email"])
        self.assertEqual(m["email"], 1)
        self.assertEqual(m.get("address"), 0)

    def test_exact_state_beats_earlier_substring(self):
        # "Statement Date" used to hijack the state column.
        m = build_header_map(["Statement Date", "State"])
        self.assertEqual(m["state"], 1)

    def test_contact_does_not_beat_phone(self):
        # A "Contact" (name) column before "Phone" used to hijack phone.
        m = build_header_map(["Contact", "Phone"])
        self.assertEqual(m["phone"], 1)

    def test_first_name_not_matched_as_state(self):
        m = build_header_map(["First Name", "Email"])
        self.assertNotIn("state", m)

    def test_full_name_only_sheet(self):
        m = build_header_map(["Name", "Email"])
        self.assertEqual(m["full_name"], 0)

    def test_emailed_and_emailed_date_both_map(self):
        m = build_header_map(["emailed", "emailed_date"])
        self.assertEqual(m["email_sent"], 0)
        self.assertEqual(m["emailed_date"], 1)

    def test_followup_and_replied(self):
        m = build_header_map(["email", "email_sent", "followup_sent", "replied"])
        self.assertEqual(m["followup_sent"], 2)
        self.assertEqual(m["replied"], 3)


class TestUnsubscribeDetection(unittest.TestCase):
    FOOTER_QUOTE = (
        "Sounds good, call me tomorrow!\n"
        "\n"
        "On Mon, Jul 27, 2026 at 9:00 AM Justin <agent@gmail.com> wrote:\n"
        "> If you'd prefer not to receive emails from me, just reply with\n"
        '> "unsubscribe" and I will take you off my list.\n'
    )

    def test_quoted_footer_is_not_an_optout(self):
        own = strip_quoted_text(self.FOOTER_QUOTE)
        self.assertFalse(is_unsubscribe_text(own))

    def test_real_optout_detected(self):
        self.assertTrue(is_unsubscribe_text("Please UNSUBSCRIBE me."))
        self.assertTrue(is_unsubscribe_text("take me off your list"))
        self.assertTrue(is_unsubscribe_text("stop emailing me"))
        self.assertTrue(is_unsubscribe_text("I want to opt out"))

    def test_normal_reply_not_flagged(self):
        self.assertFalse(is_unsubscribe_text("Yes, I'd like a quote please"))
        self.assertFalse(is_unsubscribe_text(""))

    def test_strip_quoted_gt_lines(self):
        body = "remove me please\n> old quoted text with unsubscribe\n"
        own = strip_quoted_text(body)
        self.assertIn("remove me", own)
        self.assertNotIn("old quoted", own)


if __name__ == "__main__":
    unittest.main()
