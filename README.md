# Derek the Dealer
### Coralstone Capital Group — Deal Harvesting System

Derek monitors the `deals@coralstonecapitalgroup.com` inbox every 10 minutes, extracts all data from wholesale real estate emails using Claude AI, writes every deal to Google Sheets, and pings Caleb and Grant on Telegram when a new deal lands.

---

## How It Works

1. Derek polls the HostGator IMAP inbox every 10 minutes
2. Claude AI identifies wholesale deal emails and filters out noise
3. Claude extracts every piece of deal data into structured JSON
4. Deal gets written to the **Active Deals** tab in Google Sheets
5. After 7 days, deal auto-moves to **Deal Storage**
6. Telegram alert fires for every new deal found

---

## Setup

### 1. Clone and install
```bash
git clone https://github.com/YOUR_USERNAME/derek
cd derek/worker
npm install
```

### 2. Environment variables
Copy `.env.example` to `.env` and fill in:
- `IMAP_HOST` — your HostGator mail server (find in cPanel)
- `IMAP_USER` — deals@coralstonecapitalgroup.com
- `IMAP_PASSWORD` — your email password
- `ANTHROPIC_API_KEY` — from console.anthropic.com
- `GOOGLE_SHEET_ID` — from the Google Sheet URL
- `GOOGLE_SERVICE_ACCOUNT_EMAIL` — from Google Cloud Console
- `GOOGLE_PRIVATE_KEY` — from Google Cloud Console service account JSON
- `TELEGRAM_BOT_TOKEN` — from @BotFather on Telegram
- `TELEGRAM_CHAT_ID` — your group or personal chat ID

### 3. Google Sheets Setup
1. Go to [Google Cloud Console](https://console.cloud.google.com)
2. Create a new project → Enable Google Sheets API
3. Create a Service Account → Download JSON key
4. Share your Google Sheet with the service account email (Editor access)
5. Copy the Sheet ID from the URL: `docs.google.com/spreadsheets/d/SHEET_ID/edit`

### 4. Telegram Setup
1. Message @BotFather on Telegram → `/newbot`
2. Copy the bot token
3. Add the bot to your group or start a DM
4. Get your chat ID: message @userinfobot

### 5. Deploy to Railway
1. Push to GitHub
2. Create new Railway project → Deploy from GitHub repo
3. Add all environment variables in Railway dashboard
4. Railway auto-deploys on every push

---

## Google Sheets Structure

**Active Deals tab** — all deals from the last 7 days
**Deal Storage tab** — all deals older than 7 days (sellable database)

Both tabs share the same column structure so data is consistent.

---

## Telegram Alerts

Every new deal fires a Telegram message with:
- Property address
- Asking price
- Stated ARV and repairs
- Property type, beds/baths/sqft
- Wholesaler info
- Photos included (yes/no)
- Property link

---

Built for Coralstone Capital Group — Tampa Bay, FL
