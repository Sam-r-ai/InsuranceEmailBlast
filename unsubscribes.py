# unsubscribes.py
# Finds opt-out replies in the Gmail inbox and records them on the
# Unsubscribes tab, which leademailblast.py / followups.py check before
# every send.
#
# The email footer promises "reply with unsubscribe and I'll take you off my
# list" — CAN-SPAM requires that promise to be honored within 10 business
# days, so run this before every blast (and at least weekly).
#
# False-positive guards — an address is only recorded when ALL of these hold:
#   - the message is NOT from the agent's own account
#   - the opt-out phrase appears in text the sender actually TYPED (quoted
#     history is stripped first — every reply quotes our own footer, which
#     contains the word "unsubscribe")
#   - the message sits in a thread that contains mail WE sent (i.e. it is a
#     genuine reply to our outreach, not a newsletter with its own
#     unsubscribe footer)
#
# Processed messages get a Gmail label so they are never scanned twice.

import base64
import os

from googleapiclient.errors import HttpError

from common import (
    UNSUBSCRIBES_TAB, append_values, ensure_tab, get_values, gmail_service,
    is_unsubscribe_text, normalize_email, now_timestamp, sheets_service,
    strip_quoted_text, with_backoff,
)

GMAIL_SCOPES = ["https://www.googleapis.com/auth/gmail.modify"]
GMAIL_TOKEN_FILE = "token_unsubscribes.json"

PROCESSED_LABEL = "UnsubProcessed"

# Candidate messages: inbox mail addressed to us mentioning an opt-out phrase
# (mail to the List-Unsubscribe mailto: has subject "unsubscribe").
SEARCH_QUERY = (
    f'in:inbox to:me -label:{PROCESSED_LABEL} '
    '("unsubscribe" OR "opt out" OR "remove me" OR "stop emailing" OR "take me off")'
)


def get_plain_body(payload):
    body = ""
    if "parts" in payload:
        for part in payload["parts"]:
            body += get_plain_body(part)
    elif payload.get("mimeType") == "text/plain":
        data = payload["body"].get("data")
        if data:
            body = base64.urlsafe_b64decode(data).decode("utf-8", errors="ignore")
    return body


def get_or_create_label(gmail, name):
    labels = with_backoff(
        lambda: gmail.users().labels().list(userId="me").execute(),
        what="list labels",
    ).get("labels", [])
    for label in labels:
        if label["name"] == name:
            return label["id"]
    created = with_backoff(
        lambda: gmail.users().labels().create(
            userId="me", body={"name": name}
        ).execute(),
        what="create label",
    )
    return created["id"]


def header_value(headers, name):
    return next((h["value"] for h in headers if h["name"].lower() == name.lower()), "")


def is_reply_to_our_outreach(gmail, msg):
    """True when the message's thread contains at least one message WE sent
    (other than this message itself)."""
    thread_id = msg.get("threadId")
    if not thread_id:
        return False
    thread = with_backoff(
        lambda: gmail.users().threads().get(
            userId="me", id=thread_id, format="minimal"
        ).execute(),
        what="fetch thread",
    )
    for m in thread.get("messages", []):
        if m.get("id") != msg.get("id") and "SENT" in m.get("labelIds", []):
            return True
    return False


def main():
    gmail = gmail_service(GMAIL_SCOPES, GMAIL_TOKEN_FILE)
    sheets_svc = sheets_service()
    label_id = get_or_create_label(gmail, PROCESSED_LABEL)

    profile = with_backoff(
        lambda: gmail.users().getProfile(userId="me").execute(),
        what="read profile",
    )
    own_addresses = {profile.get("emailAddress", "").lower()}
    work_email = (os.getenv("WORK_EMAIL") or "").strip().lower()
    if work_email:
        own_addresses.add(work_email)

    print("Searching inbox for opt-out replies...")
    messages, page_token = [], None
    while True:
        results = with_backoff(
            lambda: gmail.users().messages().list(
                userId="me", q=SEARCH_QUERY, maxResults=500, pageToken=page_token
            ).execute(),
            what="search inbox",
        )
        messages.extend(results.get("messages", []))
        page_token = results.get("nextPageToken")
        if not page_token:
            break

    if not messages:
        print("No new opt-out candidates found.")
        return

    print(f"Found {len(messages)} candidate message(s). Checking each...")

    tab = ensure_tab(sheets_svc, UNSUBSCRIBES_TAB,
                     header=["email", "unsubscribed_at", "subject"])
    already = set()
    for row in get_values(sheets_svc, tab):
        if row:
            addr = normalize_email(row[0])
            if addr:
                already.add(addr)

    new_rows, labeled = [], []

    for msg_meta in messages:
        try:
            msg = with_backoff(
                lambda: gmail.users().messages().get(
                    userId="me", id=msg_meta["id"], format="full"
                ).execute(),
                what="fetch message",
            )
        except HttpError as e:
            print(f"⚠️ Could not fetch message {msg_meta['id']}: {e}")
            continue

        payload = msg.get("payload", {})
        headers = payload.get("headers", [])
        sender = normalize_email(header_value(headers, "From"))
        subject = header_value(headers, "Subject")

        if sender in own_addresses:
            # Our own mail (e.g. a TEST_SEND_TO preview, whose footer
            # contains "unsubscribe"). Never an opt-out.
            labeled.append(msg_meta["id"])
            continue

        body = get_plain_body(payload) or msg.get("snippet", "")
        own_text = strip_quoted_text(body)

        if not (is_unsubscribe_text(own_text) or is_unsubscribe_text(subject)):
            # Keyword only appeared in quoted history (our own footer) —
            # a normal reply, not an opt-out. Label it so it isn't rescanned.
            labeled.append(msg_meta["id"])
            continue

        if not is_reply_to_our_outreach(gmail, msg):
            # Not part of a conversation we started — a newsletter or other
            # third-party mail with its own unsubscribe wording.
            labeled.append(msg_meta["id"])
            continue

        if not sender:
            print(f"⚠️ Opt-out-looking message {msg_meta['id']} has no parseable "
                  "From address — review it by hand.")
            continue

        labeled.append(msg_meta["id"])
        if sender in already:
            continue
        already.add(sender)
        new_rows.append([sender, now_timestamp(), subject])
        print(f"⛔ Opt-out: {sender}  (subject: {subject!r})")

    if new_rows:
        append_values(sheets_svc, tab, new_rows)
    print(f"Recorded {len(new_rows)} new unsubscribe(s) on '{tab}'.")

    for msg_id in labeled:
        try:
            with_backoff(
                lambda: gmail.users().messages().modify(
                    userId="me", id=msg_id, body={"addLabelIds": [label_id]}
                ).execute(),
                what="label message",
            )
        except HttpError as e:
            print(f"⚠️ Could not label message {msg_id}: {e}")

    print("Done. These addresses will be skipped by every future send.")


if __name__ == "__main__":
    main()
