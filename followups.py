# followups.py
# Sends ONE follow-up to leads who got the first touch (leademailblast.py)
# at least FOLLOWUP_DAYS ago and never replied. Research on outreach
# consistently shows 40%+ of all replies come from follow-ups, so this is
# the single highest-leverage script in the toolkit.
#
# Safety rails, in order, before any send:
#   1. skips anyone on the suppression list (bounced / unsubscribed)
#   2. checks Gmail for ANY message from the lead — if they ever wrote back,
#      they're marked replied and never followed up automatically
#   3. sends at most one follow-up per lead, marked in followup_sent
#
# Uses the same DRY_RUN / TEST_SEND_TO / DAILY_LIMIT knobs as the blast.

import os
import sys
import base64
import html
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

from googleapiclient.errors import HttpError

from common import (
    ALIASES, build_header_map, col_index_to_letter, ensure_column, env_flag,
    env_int, get_cell, get_values, gmail_service, jitter_sleep,
    load_suppression_set, normalize_email, now_timestamp, require_env,
    sheets_service, split_name, titlecase_name, update_values, with_backoff,
)

AGENT_NAME, WORK_PHONE, WORK_EMAIL, POSTAL_ADDRESS = require_env(
    "AGENT_NAME", "WORK_PHONE", "WORK_EMAIL", "POSTAL_ADDRESS"
)
AGENT_LICENSE = (os.getenv("AGENT_LICENSE") or os.getenv("AGENT_NUMBER") or "").strip()
if not AGENT_LICENSE:
    raise SystemExit("❌ Missing AGENT_LICENSE in .env (your insurance license number).")

AGENCY_NAME = os.getenv("AGENCY_NAME", "Family First Life")
LICENSE_STATE = os.getenv("LICENSE_STATE", "CA")
BOOKING_URL = (os.getenv("BOOKING_URL") or "").strip()
TARGET_SHEET_NAME = os.getenv("TARGET_SHEET_NAME", "Bad_Numbers_Email")
FOLLOWUP_DAYS = env_int("FOLLOWUP_DAYS", 4)
DAILY_LIMIT = env_int("DAILY_LIMIT", 25)
DELAY_MIN_SECONDS = env_int("DELAY_MIN_SECONDS", 60)
DELAY_MAX_SECONDS = env_int("DELAY_MAX_SECONDS", 180)
DRY_RUN = env_flag("DRY_RUN")
TEST_SEND_TO = (os.getenv("TEST_SEND_TO") or "").strip()

# send + readonly: readonly is what lets us detect replies and stop.
GMAIL_SCOPES = [
    "https://www.googleapis.com/auth/gmail.send",
    "https://www.googleapis.com/auth/gmail.readonly",
]
GMAIL_TOKEN_FILE = "token_followups.json"


def parse_sent_timestamp(value):
    value = str(value or "").strip()
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d",
                "%m/%d/%Y %H:%M:%S", "%m/%d/%Y", "%m/%d/%y"):
        try:
            return datetime.strptime(value, fmt)
        except ValueError:
            continue
    return None


def build_subject(first_name):
    return f"following up, {first_name}" if first_name else "following up"


def build_body_text(greeting_name):
    booking = (
        f"\n\nIf it's easier, grab a time on my calendar: {BOOKING_URL}"
        if BOOKING_URL else ""
    )
    return f"""Hi {greeting_name},

I reached out a few days ago about the life insurance information you requested and didn't want my note to get buried. I help families compare options from the top-rated carriers, and it usually takes about 10 minutes to see what you'd qualify for.

Would a quick call this week be worth it? Just reply with a time that works.{booking}

Sincerely,
{AGENT_NAME} — {LICENSE_STATE} License #{AGENT_LICENSE}
Life Insurance & Annuities Broker, {AGENCY_NAME}
Phone: {WORK_PHONE}
Email: {WORK_EMAIL}

--
This is an advertisement for life insurance products. I'm contacting you because you responded to an ad or online form requesting life insurance information.
If you'd prefer not to receive emails from me, just reply with "unsubscribe" and I will take you off my list.
{AGENT_NAME}, {POSTAL_ADDRESS}
"""


def build_body_html(greeting_name):
    e = html.escape
    booking_html = (
        f"""
    <p>If it's easier, grab a time on my calendar:<br>
      <a href="{e(BOOKING_URL, quote=True)}" style="color: #0066cc;">{e(BOOKING_URL)}</a></p>"""
        if BOOKING_URL else ""
    )
    return f"""
<html>
  <body style="font-family: Arial, Helvetica, sans-serif; font-size: 14px; color: #000; line-height: 1.6;">
    <p>Hi {e(greeting_name)},</p>
    <p>
      I reached out a few days ago about the life insurance information you requested
      and didn't want my note to get buried. I help families compare options from the
      top-rated carriers, and it usually takes about 10 minutes to see what you'd
      qualify for.
    </p>
    <p>Would a quick call this week be worth it? Just reply with a time that works.</p>{booking_html}
    <p>Sincerely,</p>
    <p>
      <strong>{e(AGENT_NAME)}</strong> &mdash; {e(LICENSE_STATE)} License #{e(AGENT_LICENSE)}<br>
      Life Insurance &amp; Annuities Broker, {e(AGENCY_NAME)}<br>
      &#128222; {e(WORK_PHONE)}<br>
      &#128231; {e(WORK_EMAIL)}
    </p>
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


def create_message(to, subject, body_text, body_html):
    message = MIMEMultipart("alternative")
    message["to"] = to
    message["subject"] = subject
    message["reply-to"] = WORK_EMAIL
    message["List-Unsubscribe"] = f"<mailto:{WORK_EMAIL}?subject=unsubscribe>"
    message.attach(MIMEText(body_text, "plain", "utf-8"))
    message.attach(MIMEText(body_html, "html", "utf-8"))
    raw = base64.urlsafe_b64encode(message.as_bytes()).decode()
    return {"raw": raw}


def has_replied(gmail, lead_email):
    """True if ANY message from this address exists in the mailbox."""
    results = with_backoff(
        lambda: gmail.users().messages().list(
            userId="me", q=f"from:{lead_email}", maxResults=1
        ).execute(),
        what=f"check replies from {lead_email}",
    )
    return bool(results.get("messages"))


def is_quota_exceeded(err):
    text = str(err).lower()
    return "quota exceeded" in text or "5.4.5" in text


def main():
    sheets_svc = sheets_service()

    rows = get_values(sheets_svc, TARGET_SHEET_NAME)
    if not rows or len(rows) < 2:
        raise SystemExit(f"Sheet '{TARGET_SHEET_NAME}' is empty or missing data rows.")

    header_map = build_header_map(rows[0])
    if "email" not in header_map:
        raise SystemExit(
            f"Couldn't find an EMAIL column in '{TARGET_SHEET_NAME}'.\n"
            f"Rename a header to one of: {ALIASES['email']}"
        )
    if "email_sent" not in header_map:
        raise SystemExit(
            f"No email_sent column on '{TARGET_SHEET_NAME}' — run "
            "leademailblast.py first; follow-ups only go to leads who "
            "received the first touch."
        )

    header = rows[0]
    if DRY_RUN:
        # DRY_RUN must not write anything — simulate missing columns instead.
        width = max(len(header), max(len(r) for r in rows))
        header_map = dict(header_map)
        for field in ("followup_sent", "replied"):
            if field not in header_map:
                print(f"🧪 Would create a {field} column (DRY_RUN: not writing).")
                header_map[field] = width
                width += 1
    else:
        header_map, header = ensure_column(sheets_svc, TARGET_SHEET_NAME, header,
                                           header_map, "followup_sent", rows=rows)
        header_map, header = ensure_column(sheets_svc, TARGET_SHEET_NAME, header,
                                           header_map, "replied", rows=rows)

    first_idx = header_map.get("first_name")
    full_idx = header_map.get("full_name")
    email_idx = header_map.get("email")
    email_sent_idx = header_map.get("email_sent")
    followup_idx = header_map.get("followup_sent")
    replied_idx = header_map.get("replied")

    followup_col = col_index_to_letter(followup_idx)
    replied_col = col_index_to_letter(replied_idx)

    print("🔎 Loading suppression list (bounces + unsubscribes)...")
    suppressed = load_suppression_set(sheets_svc)
    print(f"✅ {len(suppressed)} suppressed address(es) will be skipped.")

    # Addresses already followed up (or known to have replied) ANYWHERE on
    # the tab — so a duplicated row can never trigger a second follow-up.
    already_followed = set()
    for row in rows[1:]:
        if (str(get_cell(row, header_map.get("followup_sent"))).strip()
                or str(get_cell(row, header_map.get("replied"))).strip()):
            addr = normalize_email(get_cell(row, header_map.get("email")))
            if addr:
                already_followed.add(addr)

    gmail = gmail_service(GMAIL_SCOPES, GMAIL_TOKEN_FILE)
    if DRY_RUN:
        print("🧪 DRY_RUN is on — nothing will be sent and nothing written to the sheet.")
    if TEST_SEND_TO:
        print(f"🧪 TEST_SEND_TO is set — every email goes to {TEST_SEND_TO}; "
              "the sheet will NOT be marked.")

    now = datetime.now()
    sent = skipped_replied = skipped_suppressed = errors = unparseable = 0
    emailed_this_run = set()

    def mark_followup(row_number, value):
        update_values(sheets_svc, TARGET_SHEET_NAME,
                      f"{followup_col}{row_number}", [[value]])

    for row_number, row in enumerate(rows[1:], start=2):
        if sent >= DAILY_LIMIT:
            print(f"🛑 Reached the per-run limit of {DAILY_LIMIT} follow-ups. Stopping.")
            break

        email = normalize_email(get_cell(row, email_idx))
        if not email:
            continue
        if str(get_cell(row, followup_idx)).strip():
            continue  # already followed up
        if str(get_cell(row, replied_idx)).strip():
            continue  # already known to have replied

        if email in already_followed or email in emailed_this_run:
            print(f"↪️ Row {row_number}: {email} already followed up on another "
                  "row — marking this duplicate.")
            if not DRY_RUN and not TEST_SEND_TO:
                mark_followup(row_number, "duplicate — already followed up")
            continue

        sent_cell = str(get_cell(row, email_sent_idx)).strip()
        sent_ts = parse_sent_timestamp(sent_cell)
        if sent_ts is None:
            # Empty = never sent; a skip marker is expected. Anything else
            # is a date we couldn't read — count it so it isn't invisible.
            if sent_cell and not sent_cell.startswith(("suppressed", "duplicate")):
                unparseable += 1
                print(f"⚠️ Row {row_number}: can't read email_sent value "
                      f"{sent_cell!r} — no follow-up will be sent for it.")
            continue
        if (now - sent_ts).days < FOLLOWUP_DAYS:
            continue  # too soon

        if email in suppressed:
            skipped_suppressed += 1
            if not DRY_RUN and not TEST_SEND_TO:
                mark_followup(row_number, "suppressed (bounce/unsubscribe)")
            continue

        if has_replied(gmail, email):
            skipped_replied += 1
            print(f"💬 Row {row_number}: {email} has replied — no follow-up, marking.")
            if not DRY_RUN and not TEST_SEND_TO:
                update_values(sheets_svc, TARGET_SHEET_NAME,
                              f"{replied_col}{row_number}", [[now_timestamp()]])
            continue

        first = titlecase_name(get_cell(row, first_idx)) if first_idx is not None else ""
        if not first and full_idx is not None:
            first, _ = split_name(get_cell(row, full_idx))

        to_address = TEST_SEND_TO or email
        subject = build_subject(first)

        if DRY_RUN:
            print(f"🧪 Would follow up with {to_address} (row {row_number}, "
                  f"subject: {subject!r})")
            emailed_this_run.add(email)
            sent += 1
            continue

        msg = create_message(to_address, subject,
                             build_body_text(first or "there"),
                             build_body_html(first or "there"))
        try:
            with_backoff(
                lambda: gmail.users().messages().send(userId="me", body=msg).execute(),
                what=f"send follow-up to {email}",
            )
        except HttpError as e:
            if is_quota_exceeded(e):
                print("🛑 Gmail daily sending quota exceeded — stopping the run.")
                break
            errors += 1
            print(f"❌ Row {row_number}: failed to send to {email}: {e}")
            continue
        except Exception as e:
            errors += 1
            print(f"❌ Row {row_number}: send to {email} failed with "
                  f"{type(e).__name__}: {e}\n"
                  f"   ⚠️ The message MAY still have been delivered — check "
                  f"Sent mail and mark the row by hand before re-running.")
            continue

        emailed_this_run.add(email)
        sent += 1

        if not TEST_SEND_TO:
            try:
                mark_followup(row_number, now_timestamp())
            except Exception as e:
                print(f"🚨 SENT follow-up to {email} but FAILED to mark row "
                      f"{row_number} ({followup_col}{row_number}). Mark it manually! "
                      f"Error: {e}")
                errors += 1

        print(f"✅ Follow-up #{sent} to {first or '(no name)'} at {to_address} "
              f"(row {row_number})")

        if sent < DAILY_LIMIT:
            jitter_sleep(DELAY_MIN_SECONDS, DELAY_MAX_SECONDS)

    print(
        f"\n📊 Done. Follow-ups sent: {sent} | replied (skipped+marked): "
        f"{skipped_replied} | suppressed: {skipped_suppressed} | "
        f"unreadable email_sent: {unparseable} | errors: {errors}"
    )
    if errors:
        sys.exit(1)


if __name__ == "__main__":
    main()
