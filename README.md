# Job Application Tracker

Automatically tracks job applications and rejections from Gmail into a Google Sheet.

## How it works

- Scans Gmail every 5 minutes for application confirmation and rejection emails
- Extracts company name, role, source (LinkedIn/Indeed/etc) from the email
- Adds new applications to the sheet with status "Applied"
- Updates status to "Rejected" when a rejection email arrives from the same company
- Tracks which emails were already processed (no duplicates)

## Setup (5 minutes)

1. Create a new Google Sheet (this will be your tracker)
2. Go to **Extensions → Apps Script**
3. Delete the default `function myFunction(){}` code
4. Paste the entire contents of `Code.gs`
5. Click **Run → setup** (or select `setup` from the function dropdown and hit ▶️)
6. Grant permissions when prompted (it needs Gmail read + Sheets write)
7. Done — it now runs automatically every 5 minutes

## Features

| Feature | How |
|---------|-----|
| Auto-detect new applications | Scans for "thank you for applying", "application received", etc. |
| Auto-detect rejections | Scans for "unfortunately", "not moving forward", etc. |
| Extracts company name | From sender email domain, "at [Company]" patterns, or display name |
| Extracts role/title | From "for the [Role] position" patterns in subject/body |
| Detects source | LinkedIn, Indeed, Greenhouse, Lever, Wellfound, etc. |
| Backfill | Run `backfill()` to scan last 30 days of emails |
| Manual entry | Use the "Job Tracker" menu in the sheet |
| No duplicates | Tracks processed email IDs |

## Sheet columns

| Column | Description |
|--------|-------------|
| Company | Extracted company name |
| Role | Job title/position |
| Date Applied | When the confirmation email arrived |
| Status | Applied / Rejected / Interview |
| Date Updated | Last status change |
| Email Subject | Original email subject for reference |
| Source | Where you applied (LinkedIn, Direct, etc.) |

## Custom menu

After setup, the sheet has a **Job Tracker** menu:
- **Add Application Manually** — for applications without email confirmations
- **Backfill Last 30 Days** — catch up on older applications
- **Run Check Now** — process emails immediately without waiting

## Customization

Edit the arrays at the top of `Code.gs`:
- `APPLICATION_QUERIES` — Gmail search queries for finding application emails
- `REJECTION_QUERIES` — Gmail search queries for finding rejections
- `APPLICATION_KEYWORDS` — body text that confirms it's an application
- `REJECTION_KEYWORDS` — body text that confirms it's a rejection
- `CHECK_INTERVAL_MINUTES` — how often to scan (default: 5)

## Notes

- Free (Google Apps Script has no cost)
- Runs entirely in Google's cloud (no server needed)
- Max 6-minute runtime per execution (plenty for scanning emails)
- Stores last 500 processed email IDs to prevent reprocessing
