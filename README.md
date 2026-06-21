# Job Application Tracker

Automatically tracks job applications and rejections from Gmail into a Google Sheet using Google Apps Script.

## How it works

- Scans Gmail every 5 minutes for application confirmation and rejection emails
- Extracts company name and role from the email
- Adds new applications to the sheet with status "Applied"
- Detects rejections even when they're **nested replies** in the original application thread
- Updates the existing row with "Yes" and the rejection date when a rejection arrives
- Tracks which emails were already processed (no duplicates)

## Sheet format

| Company | Role | Date | Rejection | Date Rejected | Interview Stage |
|---------|------|------|-----------|---------------|-----------------|
| Stripe | Senior SDE | 2026-06-10 | | | Applied |
| Meta | SWE | 2026-06-05 | Yes | 2026-06-18 | Applied |
| Google | Staff Eng | 2026-06-12 | | | Phone Screen |

## Setup (one-click)

1. **[Click here to make a copy of the template](https://docs.google.com/spreadsheets/d/1FFnhnMLu72CPy3UhypF7zin0r7JleOVKXGBP7aPQhYs/copy)** — this gives you your own Sheet + script
2. In your copied Sheet, go to **Extensions → Apps Script**
3. Select `setup` from the function dropdown → click ▶️ Run
4. Grant permissions when prompted (Gmail read + Sheets write)
5. Done — runs automatically every 5 minutes against your Gmail

Your data stays completely private in your own Google Drive.

### Alternative: clone from GitHub

If you prefer version control:

1. Create a new Google Sheet → **Extensions → Apps Script** → copy the Script ID from Settings
2. Locally:
   ```bash
   npm install -g @google/clasp
   clasp login
   git clone git@github.com:kai-ion/job-tracker.git
   cd job-tracker
   echo '{"scriptId":"YOUR_SCRIPT_ID","rootDir":"."}' > .clasp.json
   clasp push --force
   ```
3. Back in Apps Script: Run `setup()` → grant permissions

## Detection patterns

### Application emails (subject search)
- "application received"
- "thank you for applying"
- "application confirmation"
- "we received your application"
- "application submitted"
- "thanks for applying"

### Rejection emails (full body search — catches threaded replies)
- "unfortunately"
- "we will not be moving forward"
- "not moving forward"
- "decided to move forward with other"
- "position has been filled"
- "after careful consideration"

### Company extraction (in priority order)
1. Sender email domain (e.g., `noreply@stripe.com` → Stripe)
2. "at/with [Company]" patterns in subject/body
3. Sender display name (e.g., "Meta Careers" → Meta)

### Role extraction
- "for the [Role] position"
- "your application for [Role]"
- "regarding the [Role] role"

### Source detection
LinkedIn, Indeed, Greenhouse, Lever, Wellfound, Glassdoor, Handshake, ZipRecruiter, or Direct

## Custom menu

After setup, the sheet has a **Job Tracker** menu:
- **Add Application Manually** — for applications without email confirmations
- **Backfill Last 30 Days** — scan older emails to catch up
- **Run Check Now** — process emails immediately

## Notion integration

For Notion users, see `Code.notion.gs` — same email detection but writes to a Notion database instead of Google Sheets. Setup requires a Notion integration token and database ID.

## Local development

```bash
# Edit code locally
vim Code.gs

# Deploy to Apps Script
clasp push --force

# Pull changes made in browser
clasp pull

# Open in browser
clasp open
```

## Cost

Free. Google Apps Script has no charges for personal accounts.

## Limits

- 6-minute max runtime per execution (scanning emails takes seconds)
- Runs as frequently as every 1 minute (default: 5 minutes)
- Stores last 500 processed email IDs to prevent reprocessing
