# leademailblast.py
# Sends the first-touch email to every unsent lead on one tab of the leads
# spreadsheet, marking each row's email_sent cell as it goes.
#
# Before every run it:
#   - validates the .env config (no more "CA License: None" emails)
#   - loads the suppression list (Invalid_Email + Unsubscribes tabs) and
#     skips anyone on it — CAN-SPAM requires honoring opt-outs
#   - dedupes addresses, so the same person is never emailed twice
#
# Every message includes the CAN-SPAM essentials: physical postal address,
# ad identification, working unsubscribe instructions, a List-Unsubscribe
# header, and the license number adjacent to the agent's name (CA Ins. Code
# 1725.5). Messages are multipart plain-text + HTML.
#
# Useful .env knobs (see .env.example): DRY_RUN=true prints instead of
# sending; TEST_SEND_TO=you@gmail.com sends every email to yourself.

import os
import sys
import base64
import html
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

from googleapiclient.errors import HttpError

from common import (
    ALIASES, build_header_map, col_index_to_letter, ensure_column, env_flag,
    env_int, get_cell, get_values, gmail_service, jitter_sleep,
    load_suppression_set, normalize_email, now_timestamp, require_env,
    sheets_service, split_name, titlecase_name, update_values, with_backoff,
)

# --- Required configuration (from .env) ---
AGENT_NAME, WORK_PHONE, WORK_EMAIL, POSTAL_ADDRESS = require_env(
    "AGENT_NAME", "WORK_PHONE", "WORK_EMAIL", "POSTAL_ADDRESS"
)
# AGENT_NUMBER kept as a fallback: older .env files used that name.
AGENT_LICENSE = (os.getenv("AGENT_LICENSE") or os.getenv("AGENT_NUMBER") or "").strip()
if not AGENT_LICENSE:
    raise SystemExit("❌ Missing AGENT_LICENSE in .env (your insurance license number).")

# --- Optional configuration ---
AGENCY_NAME = os.getenv("AGENCY_NAME", "Family First Life")
LICENSE_STATE = os.getenv("LICENSE_STATE", "CA")
BOOKING_URL = (os.getenv("BOOKING_URL") or "").strip()
TARGET_SHEET_NAME = os.getenv("TARGET_SHEET_NAME", "Bad_Numbers_Email")
TARGET_RANGE = "A1:ZZ"
DAILY_LIMIT = env_int("DAILY_LIMIT", 25)
DELAY_MIN_SECONDS = env_int("DELAY_MIN_SECONDS", 60)
DELAY_MAX_SECONDS = env_int("DELAY_MAX_SECONDS", 180)
DRY_RUN = env_flag("DRY_RUN")
TEST_SEND_TO = (os.getenv("TEST_SEND_TO") or "").strip()

BUSINESS_CARD_PATH = os.getenv("BUSINESS_CARD_PATH", os.path.join("images", "Justin_Cheung.png"))
CARRIERS_CARD_PATH = os.getenv("CARRIERS_CARD_PATH", os.path.join("images", "Carriers.png"))

GMAIL_SCOPES = ["https://www.googleapis.com/auth/gmail.send"]
GMAIL_TOKEN_FILE = "token.json"


# ---------------------------------------------------------------------------
# Message content
# ---------------------------------------------------------------------------

def build_subject(first_name):
    if first_name:
        return f"{first_name}, your life insurance request"
    return "your life insurance request"


def signature_text():
    return (
        f"{AGENT_NAME} — {LICENSE_STATE} License #{AGENT_LICENSE}\n"
        f"Life Insurance & Annuities Broker, {AGENCY_NAME}\n"
        f"Phone: {WORK_PHONE}\n"
        f"Email: {WORK_EMAIL}"
    )


def footer_text():
    return (
        "This is an advertisement for life insurance products. "
        "I'm contacting you because you responded to an ad or online form "
        "requesting life insurance information.\n"
        "If you'd prefer not to receive emails from me, just reply with "
        '"unsubscribe" and I will take you off my list.\n'
        f"{AGENT_NAME}, {POSTAL_ADDRESS}"
    )


def build_body_text(greeting_name):
    booking = (
        f"\n\nYou can also book a time directly on my calendar: {BOOKING_URL}"
        if BOOKING_URL else ""
    )
    return f"""Hi {greeting_name},

My name is {AGENT_NAME}, a licensed life insurance broker with {AGENCY_NAME}. I work with the top-rated carriers across the United States, and I'm reaching out because you previously requested information about life insurance coverage options.

Whether you're a young family looking for income replacement to secure your family's future, or a senior on a fixed income who needs affordable final expense coverage, my job is to find the right fit for your situation — at no cost to you.

Would a quick 10-minute call to look at your options be worth it? Just reply with a good time to talk.{booking}

Sincerely,
{signature_text()}

--
{footer_text()}
"""


def build_body_html(greeting_name, include_images):
    e = html.escape
    booking_html = (
        f"""
    <p>
      You can also book a time directly on my calendar:<br>
      <a href="{e(BOOKING_URL, quote=True)}" style="color: #0066cc;">{e(BOOKING_URL)}</a>
    </p>"""
        if BOOKING_URL else ""
    )
    images_html = (
        """
    <p style="margin-top:20px;">
      <img src="cid:businesscard" alt="Business card" style="max-width:420px;width:100%;border-radius:6px;display:block;margin-bottom:10px;">
      <img src="cid:carrierscard" alt="Our carriers" style="max-width:420px;width:100%;border-radius:6px;display:block;">
    </p>"""
        if include_images else ""
    )
    return f"""
<html>
  <body style="font-family: Arial, Helvetica, sans-serif; font-size: 14px; color: #000; line-height: 1.6;">
    <p>Hi {e(greeting_name)},</p>

    <p>
      My name is <strong>{e(AGENT_NAME)}</strong>, a licensed life insurance broker with
      <strong>{e(AGENCY_NAME)}</strong>. I work with the top-rated carriers across the
      United States, and I'm reaching out because you previously requested information
      about life insurance coverage options.
    </p>

    <p>
      Whether you're a young family looking for <strong>income replacement</strong> to
      secure your family's future, or a senior on a fixed income who needs
      <strong>affordable final expense coverage</strong>, my job is to find the right
      fit for your situation &mdash; at no cost to you.
    </p>

    <p>
      Would a quick 10-minute call to look at your options be worth it?
      Just reply with a good time to talk.
    </p>{booking_html}

    <p>Sincerely,</p>
    <p>
      <strong>{e(AGENT_NAME)}</strong> &mdash; {e(LICENSE_STATE)} License #{e(AGENT_LICENSE)}<br>
      Life Insurance &amp; Annuities Broker, {e(AGENCY_NAME)}<br>
      &#128222; {e(WORK_PHONE)}<br>
      &#128231; {e(WORK_EMAIL)}
    </p>{images_html}

    <hr style="border:none;border-top:1px solid #ddd;margin:20px 0;">

    <p style="font-size: 12px; color: #555;">
      This is an advertisement for life insurance products. I'm contacting you because
      you responded to an ad or online form requesting life insurance information.<br>
      If you'd prefer not to receive emails from me, just reply with
      &ldquo;unsubscribe&rdquo; and I will take you off my list.<br>
      {e(AGENT_NAME)}, {e(POSTAL_ADDRESS)}
    </p>
  </body>
</html>
"""


# ---------------------------------------------------------------------------
# Gmail
# ---------------------------------------------------------------------------

def available_inline_images():
    """Only attach images that actually exist, and only reference the CIDs we
    attach — no more emails with broken image placeholders."""
    cards = {"businesscard": BUSINESS_CARD_PATH, "carrierscard": CARRIERS_CARD_PATH}
    found = {cid: path for cid, path in cards.items() if os.path.exists(path)}
    missing = [path for path in cards.values() if not os.path.exists(path)]
    return found, missing


def create_message(to, subject, body_text, body_html, images):
    message = MIMEMultipart("related")
    message["to"] = to
    message["subject"] = subject
    message["reply-to"] = WORK_EMAIL
    # Lets Gmail show an Unsubscribe button and turn would-be spam reports
    # into harmless unsubscribes.
    message["List-Unsubscribe"] = f"<mailto:{WORK_EMAIL}?subject=unsubscribe>"

    alternative = MIMEMultipart("alternative")
    message.attach(alternative)
    alternative.attach(MIMEText(body_text, "plain", "utf-8"))
    alternative.attach(MIMEText(body_html, "html", "utf-8"))

    for cid, path in images.items():
        with open(path, "rb") as f:
            img = MIMEImage(f.read())
        img.add_header("Content-ID", f"<{cid}>")
        img.add_header("Content-Disposition", "inline", filename=os.path.basename(path))
        message.attach(img)

    raw = base64.urlsafe_b64encode(message.as_bytes()).decode()
    return {"raw": raw}


def is_quota_exceeded(err):
    """Gmail's daily-quota block: retrying is pointless (the account is
    send-blocked for up to 24h), so the run must stop."""
    text = str(err).lower()
    return "quota exceeded" in text or "5.4.5" in text


def send_email(gmail, to_name, to_email, images):
    subject = build_subject(to_name)
    body_text = build_body_text(to_name or "there")
    body_html = build_body_html(to_name or "there", include_images=bool(images))
    msg = create_message(to_email, subject, body_text, body_html, images)
    with_backoff(
        lambda: gmail.users().messages().send(userId="me", body=msg).execute(),
        what=f"send to {to_email}",
    )


# ---------------------------------------------------------------------------
# Run
# ---------------------------------------------------------------------------

def main():
    sheets_svc = sheets_service()

    images, missing_images = available_inline_images()
    for path in missing_images:
        print(f"⚠️ Image not found, sending without it: {path}")

    rows = get_values(sheets_svc, TARGET_SHEET_NAME, TARGET_RANGE)
    if not rows or len(rows) < 2:
        raise SystemExit(f"Sheet '{TARGET_SHEET_NAME}' is empty or missing data rows.")

    header_map = build_header_map(rows[0])
    if "email" not in header_map:
        raise SystemExit(
            f"Couldn't find an EMAIL column in '{TARGET_SHEET_NAME}'.\n"
            f"Header row was: {rows[0]}\n"
            f"Rename a header to one of: {ALIASES['email']}"
        )

    header_map, _ = ensure_column(sheets_svc, TARGET_SHEET_NAME, rows[0], header_map, "email_sent")
    print("✅ Detected header mapping:", header_map)

    first_idx = header_map.get("first_name")
    last_idx = header_map.get("last_name")
    full_idx = header_map.get("full_name")
    email_idx = header_map.get("email")
    email_sent_idx = header_map.get("email_sent")
    sent_col_letter = col_index_to_letter(email_sent_idx)

    print("🔎 Loading suppression list (bounces + unsubscribes)...")
    suppressed = load_suppression_set(sheets_svc)
    print(f"✅ {len(suppressed)} suppressed address(es) will be skipped.")

    # Addresses already emailed anywhere on this tab (so a duplicate row
    # can't trigger a second copy).
    already_sent = set()
    for row in rows[1:]:
        if str(get_cell(row, email_sent_idx)).strip():
            addr = normalize_email(get_cell(row, email_idx))
            if addr:
                already_sent.add(addr)

    if DRY_RUN:
        print("🧪 DRY_RUN is on — nothing will be sent and nothing written to the sheet.")
        gmail = None
    else:
        gmail = gmail_service(GMAIL_SCOPES, GMAIL_TOKEN_FILE)
    if TEST_SEND_TO:
        print(f"🧪 TEST_SEND_TO is set — every email goes to {TEST_SEND_TO}; "
              "the sheet will NOT be marked.")

    sent = skipped_suppressed = skipped_duplicate = errors = 0
    emailed_this_run = set()

    def mark_row(row_number, value):
        update_values(sheets_svc, TARGET_SHEET_NAME,
                      f"{sent_col_letter}{row_number}", [[value]])

    for row_number, row in enumerate(rows[1:], start=2):
        if sent >= DAILY_LIMIT:
            print(f"🛑 Reached the per-run limit of {DAILY_LIMIT} emails. "
                  f"Stopping — {sent} sent this run.")
            break

        email = normalize_email(get_cell(row, email_idx))
        if not email:
            continue
        if str(get_cell(row, email_sent_idx)).strip():
            continue

        if email in suppressed:
            skipped_suppressed += 1
            print(f"⛔ Row {row_number}: {email} is on the suppression list "
                  "(bounced or unsubscribed) — skipping.")
            if not DRY_RUN and not TEST_SEND_TO:
                mark_row(row_number, "suppressed (bounce/unsubscribe)")
            continue

        if email in already_sent or email in emailed_this_run:
            skipped_duplicate += 1
            print(f"↪️ Row {row_number}: {email} already emailed — skipping duplicate.")
            if not DRY_RUN and not TEST_SEND_TO:
                mark_row(row_number, "duplicate — already emailed")
            continue

        first = titlecase_name(get_cell(row, first_idx)) if first_idx is not None else ""
        last = titlecase_name(get_cell(row, last_idx)) if last_idx is not None else ""
        if (not first and not last) and full_idx is not None:
            first, _ = split_name(get_cell(row, full_idx))
        greeting_name = first or ""

        to_address = TEST_SEND_TO or email

        if DRY_RUN:
            print(f"🧪 Would send to {to_address} (row {row_number}, "
                  f"subject: {build_subject(greeting_name)!r})")
            emailed_this_run.add(email)
            sent += 1
            continue

        try:
            send_email(gmail, greeting_name, to_address, images)
        except HttpError as e:
            if is_quota_exceeded(e):
                print(f"🛑 Gmail daily sending quota exceeded — stopping the run. "
                      f"({sent} sent before the block.)")
                break
            errors += 1
            print(f"❌ Row {row_number}: failed to send to {email}: {e}")
            continue

        emailed_this_run.add(email)
        sent += 1

        if not TEST_SEND_TO:
            ts = now_timestamp()
            try:
                mark_row(row_number, ts)
            except HttpError as e:
                # The send DID happen. Surface it loudly so the row can be
                # marked by hand — otherwise the next run re-emails this lead.
                print(f"🚨 SENT to {email} but FAILED to mark row {row_number} "
                      f"({sent_col_letter}{row_number}). Mark it manually before "
                      f"the next run! Error: {e}")
                errors += 1

        print(f"✅ Sent #{sent} to {greeting_name or '(no name)'} at {to_address} "
              f"(row {row_number})")

        if sent < DAILY_LIMIT:
            jitter_sleep(DELAY_MIN_SECONDS, DELAY_MAX_SECONDS)

    print(
        f"\n📊 Done. Sent: {sent} | suppressed: {skipped_suppressed} | "
        f"duplicates: {skipped_duplicate} | errors: {errors}"
    )
    if errors:
        sys.exit(1)


if __name__ == "__main__":
    main()
