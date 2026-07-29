# Insurance Email Blast Automation

A Python toolkit for licensed insurance agents to manage lead lists in Google
Sheets and send personalized outreach email through their own Gmail account
(Gmail API + OAuth2).

It handles the whole loop: organizing messy lead sheets → sending the first
touch → sending one follow-up to non-responders → harvesting bounces →
recording unsubscribes → suppressing anyone who bounced or opted out from all
future sends.

---

## 🧰 The Scripts

Run them in roughly this order:

| Script | What it does | Credentials used |
| --- | --- | --- |
| `sheet_organizer.py` | Rewrites one messy lead tab into the standard column layout. Safe to re-run — send-tracking columns are preserved. | service account |
| `sheets_combiner.py` | Combines several lead tabs into one output tab, deduping by email/phone and carrying forward `email_sent` history. | service account |
| `leads_state_organizer.py` | Sorts a tab by State (atomic in-place sort). | service account |
| `sheets_scrubber.py` | Deletes existing clients (matched by phone **or** email) out of lead tabs so they never get cold outreach. | service account |
| `unsubscribes.py` | Scans the inbox for "unsubscribe"-type replies and records them on the `Unsubscribes` tab. **Run before every blast.** | `token_unsubscribes.json` |
| `bounces.py` | Harvests mailer-daemon bounces: hard bounces → `Invalid_Email` (suppression), soft bounces → `Soft_Bounces`, unparseable → `Bounce_Review`. | `token_cleanup.json` |
| `leademailblast.py` | Sends the first-touch email to unsent leads on the target tab, skipping suppressed/duplicate addresses, and timestamps `email_sent`. | `token.json` |
| `followups.py` | Sends ONE follow-up to leads emailed ≥ `FOLLOWUP_DAYS` ago who never replied (reply detection via Gmail search). | `token_followups.json` |
| `lead_finder.py` | Finds a lead by first + last name across every tab. | service account |

All scripts share `common.py` (auth, retries with exponential backoff, header
detection, suppression list, normalizers).

**Suppression list:** every send checks the `Invalid_Email` and
`Unsubscribes` tabs first. Once an address hard-bounces or opts out it is
never emailed again — that's both a CAN-SPAM requirement and how you protect
your Gmail sender reputation.

---

## ✅ Requirements

* Python 3.10 or higher
* A Gmail account (for sending)
* A Google Cloud project (for credentials)
* A Google Sheet with leads

---

## 🛠 Installation

```bash
git clone https://github.com/Sam-r-ai/InsuranceEmailBlast.git
cd InsuranceEmailBlast
pip install -r requirements.txt
```

---

## 🔐 Setup: Google Cloud + Gmail

### A. Gmail API OAuth2 credentials (for sending/reading mail)

1. Go to [Google Cloud Console](https://console.cloud.google.com/)
2. Create a new project (or use an existing one)
3. **APIs & Services → Library** → enable **Gmail API** (and **Google Sheets API**)
4. **APIs & Services → Credentials** → **Create Credentials → OAuth client ID**
   * Application type: **Desktop App**
5. Download the file and save it in the project folder as:

```
credentials.json
```

The first time each Gmail-using script runs, a browser window asks you to
authorize; the script then saves a token file (see the table above) so later
runs are non-interactive. If a token expires it refreshes silently; if it's
revoked you'll just be asked to sign in again.

### B. Google Sheets service-account credentials

1. Same project → **IAM & Admin → Service Accounts** → **Create Service Account**
2. Open it → **Keys** tab → **Add Key → Create new key → JSON**
3. Save the downloaded file in the project folder as:

```
sheet_service_account.json
```

4. Share your Google Sheet (Editor) with the service account's email address
   (ends in `@yourproject.iam.gserviceaccount.com`).

---

## 📄 Configuration: the .env file

Copy `.env.example` to `.env` and fill it in. Required values:

```env
AGENT_NAME=Your Name
AGENT_LICENSE=1234567
WORK_PHONE=5555555555
WORK_EMAIL=youremail@gmail.com
POSTAL_ADDRESS=123 Main St Suite 4, Sacramento, CA 95814
SPREADSHEET_ID=your_google_sheet_id_here
```

`POSTAL_ADDRESS` is not optional — CAN-SPAM requires a valid physical postal
address in every commercial email (a USPS-registered PO Box or registered
private mailbox works if you don't want to publish a street address).

Useful optional settings (defaults in `.env.example`): `TARGET_SHEET_NAME`,
`DAILY_LIMIT`, `DELAY_MIN_SECONDS`/`DELAY_MAX_SECONDS`, `FOLLOWUP_DAYS`,
`BOOKING_URL`, `AGENCY_NAME`, `LICENSE_STATE`, `DRY_RUN`, `TEST_SEND_TO`.

**Spreadsheet ID:** for
`https://docs.google.com/spreadsheets/d/1ABCDefGHIJKL.../edit#gid=0`
the ID is `1ABCDefGHIJKL...`.

### Sheet format

Column **order doesn't matter** — columns are found by header name, and most
common spellings are recognized (`First Name`, `fname`, `E-mail`,
`Phone Number`, `Cell`, ...). A recognizable **Email** header is required on
the blast tab. Tracking columns (`email_sent`, `followup_sent`, `replied`)
are created automatically.

---

## 🚀 Run

Always preview first:

```bash
DRY_RUN=true python leademailblast.py     # prints what would happen, sends nothing
TEST_SEND_TO=you@gmail.com python leademailblast.py   # real sends, but only to yourself
```

Then the real thing:

```bash
python unsubscribes.py     # record any new opt-outs first (CAN-SPAM)
python bounces.py          # record any new bounces
python leademailblast.py   # first touch, up to DAILY_LIMIT
python followups.py        # one follow-up to non-responders
```

Each send is spaced by a random 60–180 s delay (evenly-spaced robotic sending
is a spam signal), so a 25-email run takes ~30–75 minutes. The console shows
a per-run summary: sent / suppressed / duplicates / errors.

Run the unit tests with:

```bash
python -m unittest discover tests
```

---

## 📈 What actually makes this work

Sending the email is the easy part. The evidence on what drives replies and
booked appointments:

* **Follow-ups are ~40% of your replies.** Roughly 55–58% of replies come
  from the first email and the rest from follow-ups; a single follow-up
  lifts total replies by ~65%. That's why `followups.py` exists — run it.
  For insurance specifically, most sales take 5+ touches (email + phone).
* **Speed to lead dominates.** Contacting a fresh lead within 5 minutes
  makes qualification ~21x more likely than waiting 30 minutes, and most
  buyers go with the first responder. Email a fresh lead the day it arrives
  — and call it the same day. Aged leads (90+ days) contact at only 8–15%,
  so judge those campaigns accordingly.
* **Timing:** Tuesday–Thursday, ~9–11 am recipient-local time performs
  30–45% better than Monday/Friday sends.
* **Subject lines:** short (2–4 words) and specific beats clever; casual
  lowercase outperforms Title Case in cold-outreach tests; never imply
  something false — deceptive subjects are a CAN-SPAM violation and (in CA)
  a $1,000-per-email private lawsuit risk.
* **Body:** 75–125 words, one clear ask, personal plain-text look. Heavy
  image/HTML templates read as marketing blasts and get filtered; that's
  why the emails here always include a full plain-text part, and why you
  should consider skipping the inline images entirely (they're optional —
  the scripts send cleanly without them).
* **Measure replies and bookings, not opens.** Apple Mail Privacy makes
  open rates fiction. Realistic targets for warm/aged insurance leads:
  4–8% reply rate, 1–3% booked appointments per batch. If a batch bounces
  more than ~5%, stop and re-validate the list before continuing.
* **Volume:** a personal Gmail account allows ~500 recipients per rolling
  24 h across ALL mail, and unsolicited bulk mail from consumer accounts is
  against Gmail policy — keep `DAILY_LIMIT` modest (default 25), ramp up
  gradually on a fresh account, and keep day-to-day volume steady. If this
  becomes a real channel, move to Google Workspace on your own domain (2,000
  /day, real SPF/DKIM/DMARC control, Postmaster Tools spam-rate monitoring).

---

## ⚖️ Compliance checklist (read this)

The templates and scripts implement the mechanics, but compliance is
ultimately about how *you* use them:

* **CAN-SPAM** (every commercial email): valid physical postal address
  (`POSTAL_ADDRESS`), truthful From/subject, clear opt-out notice, and
  opt-outs honored within **10 business days**. The footer promises
  "reply unsubscribe" — `unsubscribes.py` is what keeps that promise, so run
  it at least weekly and before every blast. Penalties run to ~$53k per
  email.
* **Truthful lead-source claims:** the template says the recipient
  "previously requested information about life insurance coverage options."
  Only send to leads for whom that is actually true, and keep the lead
  vendor's opt-in/consent records. If a list can't be verified, don't imply
  an inquiry.
* **California Insurance Code 1725.5:** license number must appear adjacent
  to or directly below your name — the templates already format the
  signature this way. Set `LICENSE_STATE` + `AGENT_LICENSE`.
* **Senior marketing (CA Ins. Code 787):** final-expense outreach to 65+
  requires disclosing that the contact results from a lead/advertisement —
  the footer includes this disclosure; don't remove it.
* **State licensing:** only email leads in states where you're licensed
  (that's one reason `leads_state_organizer.py` exists).
* **Carrier approval:** naming your IMO/agency or showing carrier logos
  (the `Carriers.png` inline image) generally requires advertising
  pre-approval from the carriers/IMO. Get it in writing or drop the image.

---

## 🔒 Files that must never be committed

`.gitignore` already covers them: `.env`, `credentials.json`, every
`token*.json`, the service-account key, and the `images/` folder. If a
credential file ever leaks, revoke it in Google Cloud Console immediately.
