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

## Setup

1. Create a new Google Sheet
2. Go to **Extensions → Apps Script**
3. In the script editor, go to **Settings (gear icon) → Script ID** — copy it
4. Locally:
   ```bash
   npm install -g @google/clasp
   clasp login
   ```
5. Clone and push:
   ```bash
   git clone git@github.com:kai-ion/job-tracker.git
   cd job-tracker
   echo '{"scriptId":"YOUR_SCRIPT_ID","rootDir":"."}' > .clasp.json
   clasp push --force
   ```
6. Back in Apps Script: Run `setup()` → grant permissions
7. Done — runs automatically every 5 minutes

### Alternative: manual paste

If you don't want to use clasp, copy the contents of `Code.gs` into the Apps Script editor and run `setup()`.

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
