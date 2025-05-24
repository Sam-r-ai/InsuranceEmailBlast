# Insurance Email Blast Automation

This Python script pulls lead data from a Google Sheet and sends out personalized emails using the Gmail API via OAuth2.

Perfect for licensed agents looking to automate outreach to warm leads with FOMO-style messaging.

---

## 🎞 What This Script Does

* Connects to Google Sheets using a service account
* Reads first name and email of each lead
* Sends personalized emails using your Gmail account with Gmail API and OAuth2
* Keeps your sensitive info safe using a `.env` file

---

## ✅ Requirements

* Python 3.10 or higher
* A Gmail account (for sending)
* A Google Cloud account (for credentials)
* A Google Sheet with leads

---

## 🛠 Installation

1. **Clone the repo**

```bash
git clone https://github.com/yourusername/InsuranceEmailBlast.git
cd InsuranceEmailBlast
```

2. **Install dependencies**

```bash
pip install -r requirements.txt
```

If you don’t have `requirements.txt`, use:

```bash
pip install google-api-python-client google-auth google-auth-oauthlib python-dotenv
```

---

## 🔐 Setup: Google Cloud + Gmail

### 📌 A. Create Gmail API OAuth2 Credentials (for sending emails)

1. Go to [Google Cloud Console](https://console.cloud.google.com/)
2. Create a new project (or use an existing one)
3. Go to **APIs & Services > Library**
4. Enable **Gmail API**
5. Go to **APIs & Services > Credentials**
6. Click **Create Credentials → OAuth client ID**

   * Application type: **Desktop App**
   * Name: Gmail API Sender
7. Download the file and rename it to:

```
credentials.json
```

8. Place it in the project folder

---

### 📌 B. Create Google Sheets Service Account Credentials

1. In the same project, go to **IAM & Admin > Service Accounts**
2. Click **Create Service Account**

   * Name it something like Sheets Reader
3. Click into it after creating
4. Go to the **"Keys"** tab → **Add Key > Create new key > JSON**
5. Download the file and rename it:

```
sheets_service_account.json
```

6. Share your Google Sheet with the **service account email** (it ends in `@yourproject.iam.gserviceaccount.com`)

---

## 📄 Setting Up Your .env File

In the root of the project, create a file named `.env`:

```env
AGENT_NAME=Your Name
AGENT_LICENSE=1234567
WORK_PHONE=5555555555
WORK_EMAIL=youremail@gmail.com
SPREADSHEET_ID=your_google_sheet_id_here
```

### 🔍 How to Get Your Spreadsheet ID

Example Sheet URL:

```
https://docs.google.com/spreadsheets/d/1ABCDefGHIJKL1234567890XYZ/edit#gid=0
```

Your ID is: `1ABCDefGHIJKL1234567890XYZ`

---

## 📋 Sample Google Sheet Format

| First Name | Last Name | Phone  | Email                                           |
| ---------- | --------- | ------ | ----------------------------------------------- |
| Justin     | Cheung    | 555... | [justin@example.com](mailto:justin@example.com) |
| Jay        | Money     | 555... | [jay@example.com](mailto:jay@example.com)       |

**Note**: Email must be in Column D.

---

## 🚀 Run the Script

```bash
python bronzeleadblast.python.py
```

The first time you run it, a browser will open asking you to authenticate your Gmail account. After that, a `token.json` file will
