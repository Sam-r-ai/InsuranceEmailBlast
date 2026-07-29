# bounces.py
# Harvests mailer-daemon bounce notifications from Gmail and records the
# failed addresses in the spreadsheet:
#   - hard bounces (address doesn't exist)  -> Invalid_Email tab (suppression)
#   - soft bounces (inbox full / temporary) -> Soft_Bounces tab (retryable)
#   - unparseable notifications             -> Bounce_Review tab (check by hand)
#
# The sheet is written FIRST and messages are only trashed after their data
# is safely recorded — a crash can no longer lose bounce data (Gmail search
# doesn't see Trash, so trashed-but-unrecorded bounces used to vanish).
#
# leademailblast.py and followups.py read the Invalid_Email tab as part of
# the suppression list, so running this regularly keeps bad addresses from
# being emailed again (protecting your Gmail sender reputation).

import base64
import re

from googleapiclient.errors import HttpError

from common import (
    BOUNCE_REVIEW_TAB, INVALID_EMAIL_TAB, SOFT_BOUNCES_TAB, append_values,
    ensure_tab, get_values, gmail_service, normalize_email, now_timestamp,
    sheets_service, with_backoff,
)

# gmail.modify = read + trash. The old full-mailbox scope was far broader
# than this script needs.
GMAIL_SCOPES = ["https://www.googleapis.com/auth/gmail.modify"]
GMAIL_TOKEN_FILE = "token_cleanup.json"

BOUNCE_QUERY = "from:mailer-daemon"

FAILED_EMAIL_RE = re.compile(
    r"(?:to|address)\s+<?([a-z0-9._%+-]+@[a-z0-9.-]+\.[a-z]{2,})>?", re.I
)

HARD_BOUNCE_PATTERNS = (
    "address not found", "does not exist", "user unknown", "no such user",
    "address rejected", "recipient not found", "550 5.1.1", "5.1.1",
)
SOFT_BOUNCE_PATTERNS = (
    "inbox full", "inbox is full", "mailbox full", "quota exceeded",
    "delivery incomplete", "temporary", "try again later", "421",
)


def get_email_body(payload):
    body = ""
    if "parts" in payload:
        for part in payload["parts"]:
            body += get_email_body(part)
    elif payload.get("mimeType") == "text/plain":
        data = payload["body"].get("data")
        if data:
            body = base64.urlsafe_b64decode(data).decode("utf-8", errors="ignore")
    return body


def classify(body_lower, subject_lower):
    text = body_lower + " " + subject_lower
    if any(p in text for p in HARD_BOUNCE_PATTERNS):
        return "hard", "Address not found"
    if any(p in text for p in SOFT_BOUNCE_PATTERNS):
        return "soft", "Inbox full / temporary problem"
    return "unknown", "Unclassified bounce"


def list_bounce_messages(gmail):
    messages, page_token = [], None
    while True:
        results = with_backoff(
            lambda: gmail.users().messages().list(
                userId="me", q=BOUNCE_QUERY, maxResults=500, pageToken=page_token
            ).execute(),
            what="list bounces",
        )
        messages.extend(results.get("messages", []))
        page_token = results.get("nextPageToken")
        if not page_token:
            break
    return messages


def main():
    gmail = gmail_service(GMAIL_SCOPES, GMAIL_TOKEN_FILE)
    sheets_svc = sheets_service()

    print("Searching for bounce-backs...")
    messages = list_bounce_messages(gmail)
    if not messages:
        print("No bounce-back messages found.")
        return

    total = len(messages)
    print(f"Found {total} bounce-back(s). Processing...")

    hard_rows, soft_rows, review_rows = [], [], []
    processed_ids, failed = [], 0

    for index, msg_meta in enumerate(messages, start=1):
        print(f"Processing {index} of {total}...", end="\r")
        try:
            msg = with_backoff(
                lambda: gmail.users().messages().get(
                    userId="me", id=msg_meta["id"], format="full"
                ).execute(),
                what="fetch bounce",
            )
        except HttpError as e:
            failed += 1
            print(f"\n⚠️ Could not fetch message {msg_meta['id']}: {e}")
            continue

        payload = msg.get("payload", {})
        headers = payload.get("headers", [])
        subject = next(
            (h["value"] for h in headers if h["name"].lower() == "subject"),
            "Unknown Subject",
        )
        body = get_email_body(payload) or msg.get("snippet", "")

        match = FAILED_EMAIL_RE.search(body.lower())
        failed_email = match.group(1) if match else ""
        kind, reason = classify(body.lower(), subject.lower())
        ts = now_timestamp()

        if not failed_email:
            # Keep the message OUT of the trash so it can be read by hand.
            review_rows.append(
                ["(could not extract address)", reason, subject, ts, msg_meta["id"]]
            )
            continue

        row = [failed_email, reason, subject, ts, msg_meta["id"]]
        if kind == "hard":
            hard_rows.append(row)
        elif kind == "soft":
            soft_rows.append(row)
        else:
            review_rows.append(row)
        processed_ids.append(msg_meta["id"])

    print(f"\nParsed {len(hard_rows)} hard, {len(soft_rows)} soft, "
          f"{len(review_rows)} for review ({failed} fetch failures).")

    # 1) Record everything in the spreadsheet FIRST.
    HEADER = ["email", "reason", "subject", "recorded_at", "message_id"]

    def upload(tab, rows, dedupe_by="email"):
        if not rows:
            return 0
        actual_tab = ensure_tab(sheets_svc, tab, header=HEADER)
        # Dedupe on address (suppression tabs) or on Gmail message id (the
        # review tab, whose unparseable messages stay in the inbox and would
        # otherwise be re-appended every run).
        seen = set()
        for existing in get_values(sheets_svc, actual_tab):
            if dedupe_by == "email":
                key = normalize_email(existing[0]) if existing else ""
            else:
                key = existing[4] if len(existing) > 4 else ""
            if key:
                seen.add(key)
        fresh = []
        for r in rows:
            key = normalize_email(r[0]) if dedupe_by == "email" else r[4]
            if key and key in seen:
                continue
            if key:
                seen.add(key)
            fresh.append(r)
        if fresh:
            append_values(sheets_svc, actual_tab, fresh)
        return len(fresh)

    n_hard = upload(INVALID_EMAIL_TAB, hard_rows)
    n_soft = upload(SOFT_BOUNCES_TAB, soft_rows)
    n_review = upload(BOUNCE_REVIEW_TAB, review_rows, dedupe_by="message_id")
    print(f"Uploaded: {n_hard} to {INVALID_EMAIL_TAB}, {n_soft} to "
          f"{SOFT_BOUNCES_TAB}, {n_review} to {BOUNCE_REVIEW_TAB}.")

    # 2) Only now is it safe to trash the processed notifications.
    trashed = 0
    for msg_id in processed_ids:
        try:
            with_backoff(
                lambda: gmail.users().messages().trash(userId="me", id=msg_id).execute(),
                what="trash bounce",
            )
            trashed += 1
        except HttpError as e:
            print(f"⚠️ Could not trash message {msg_id}: {e}")

    print(f"Done. Recorded {n_hard + n_soft + n_review} bounce(s), "
          f"trashed {trashed} notification(s).")
    if review_rows:
        print(f"👀 {len(review_rows)} bounce(s) need manual review — see the "
              f"'{BOUNCE_REVIEW_TAB}' tab (unparseable ones were left in the inbox).")


if __name__ == "__main__":
    main()
