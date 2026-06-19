/**
 * Job Application Tracker — Google Apps Script
 *
 * Automatically scans Gmail for application confirmations and rejections,
 * then logs them to a Google Sheet.
 *
 * Setup:
 * 1. Create a new Google Sheet
 * 2. Go to Extensions → Apps Script
 * 3. Paste this entire file into the editor
 * 4. Run `setup()` once to create headers and the time trigger
 * 5. Done — it runs automatically every 5 minutes
 */

// ============================================================
// CONFIG
// ============================================================

const SHEET_NAME = "Applications";
const CHECK_INTERVAL_MINUTES = 5;

// Gmail search queries for application confirmations
const APPLICATION_QUERIES = [
  'subject:"application received" newer_than:1d',
  'subject:"thank you for applying" newer_than:1d',
  'subject:"application confirmation" newer_than:1d',
  'subject:"we received your application" newer_than:1d',
  'subject:"application submitted" newer_than:1d',
  'subject:"thanks for applying" newer_than:1d',
];

// Gmail search queries for rejections (search body too, not just subject,
// because rejections often come as replies in the application thread)
const REJECTION_QUERIES = [
  '"unfortunately" newer_than:1d',
  '"we will not be moving forward" newer_than:1d',
  '"not moving forward" newer_than:1d',
  '"decided to move forward with other" newer_than:1d',
  '"position has been filled" newer_than:1d',
  '"after careful consideration" newer_than:1d',
  'subject:"update on your application" newer_than:1d',
];

// Keywords in email body that confirm rejection
const REJECTION_KEYWORDS = [
  "unfortunately",
  "not moving forward",
  "decided to pursue other",
  "will not be moving forward",
  "position has been filled",
  "not a match",
  "other candidates",
  "we regret",
  "unable to offer",
  "not selected",
];

// Keywords that confirm an application was received
const APPLICATION_KEYWORDS = [
  "received your application",
  "thank you for applying",
  "application has been submitted",
  "we have received",
  "confirm your application",
  "successfully submitted",
  "thanks for your interest",
];

// ============================================================
// SETUP (run once)
// ============================================================

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }

  // Set headers if empty
  if (sheet.getLastRow() === 0) {
    const headers = ["Company", "Role", "Date", "Rejection", "Date Rejected", "Interview Stage"];
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold");
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 200); // Company
    sheet.setColumnWidth(2, 220); // Role
    sheet.setColumnWidth(3, 110); // Date
    sheet.setColumnWidth(4, 100); // Rejection
    sheet.setColumnWidth(5, 120); // Date Rejected
    sheet.setColumnWidth(6, 150); // Interview Stage
  }

  // Create time trigger (every 5 min)
  deleteExistingTriggers_();
  ScriptApp.newTrigger("processNewEmails")
    .timeBased()
    .everyMinutes(CHECK_INTERVAL_MINUTES)
    .create();

  Logger.log("Setup complete. Trigger created to run every " + CHECK_INTERVAL_MINUTES + " minutes.");
  SpreadsheetApp.getUi().alert("Setup complete! The tracker will check for new emails every 5 minutes.");
}


// ============================================================
// MAIN PROCESSING
// ============================================================

function processNewEmails() {
  const sheet = getOrCreateSheet_();
  const processed = getProcessedMessageIds_(sheet);

  // Process application confirmations
  for (const query of APPLICATION_QUERIES) {
    const threads = GmailApp.search(query, 0, 20);
    for (const thread of threads) {
      for (const msg of thread.getMessages()) {
        if (processed.has(msg.getId())) continue;

        const body = msg.getPlainBody();
        if (isApplicationEmail_(body)) {
          const info = extractApplicationInfo_(msg);
          if (info.company) {
            addApplication_(sheet, info, msg.getId());
          }
        }
      }
    }
  }

  // Process rejections — update existing rows
  for (const query of REJECTION_QUERIES) {
    const threads = GmailApp.search(query, 0, 20);
    for (const thread of threads) {
      for (const msg of thread.getMessages()) {
        if (processed.has(msg.getId())) continue;

        const body = msg.getPlainBody();
        if (isRejectionEmail_(body)) {
          const info = extractApplicationInfo_(msg);
          if (info.company) {
            updateStatus_(sheet, info.company, "Rejected", msg.getId(), msg.getSubject());
          }
        }
      }
    }
  }
}


// ============================================================
// EMAIL PARSING
// ============================================================

function extractApplicationInfo_(message) {
  const subject = message.getSubject();
  const from = message.getFrom();
  const body = message.getPlainBody().substring(0, 2000); // First 2000 chars
  const date = message.getDate();

  let company = extractCompany_(from, subject, body);
  let role = extractRole_(subject, body);
  let source = extractSource_(from, body);

  return { company, role, date, subject, source };
}

function extractCompany_(from, subject, body) {
  // Try from email domain first (e.g., "noreply@stripe.com" → Stripe)
  const domainMatch = from.match(/@([a-z0-9-]+)\./i);
  if (domainMatch) {
    const domain = domainMatch[1].toLowerCase();
    // Skip generic email providers and ATS platforms
    const generic = ["gmail", "yahoo", "outlook", "hotmail", "greenhouse", "lever", "ashby",
                     "smartrecruiters", "workday", "icims", "jobvite", "myworkdayjobs",
                     "successfactors", "taleo", "brassring", "google", "icloud"];
    if (!generic.includes(domain)) {
      return capitalize_(domain);
    }
  }

  const text = subject + "\n" + body;

  // Try "at [Company]" patterns — stop at common sentence continuations
  const atPatterns = [
    /(?:application (?:at|to|for))\s+([A-Z][A-Za-z0-9\s&.-]{1,40}?)(?:\s+for|\s+has|\s+was|\.|,|\n|!|\s+and\b)/,
    /(?:at|with)\s+([A-Z][A-Za-z0-9&.-]+)(?:\s|\.|\,|\n|!|$)/,
    /(?:from)\s+([A-Z][A-Za-z0-9&.-]+?)(?:\s+regarding|\s+about|\.|,|\n|!|$)/,
  ];
  for (const pattern of atPatterns) {
    const match = text.match(pattern);
    if (match) {
      let name = match[1].trim().replace(/[.\s]+$/, ""); // Strip trailing periods/spaces
      if (name.length > 1 && name.length < 40) {
        const reject = ["this", "that", "the", "our", "your", "unfortunately", "time"];
        if (!reject.includes(name.toLowerCase())) {
          return name;
        }
      }
    }
  }

  // Try the display name from the From field
  const nameMatch = from.match(/^"?([^"<]+)"?\s*</);
  if (nameMatch) {
    let name = nameMatch[1].trim();
    // Remove common suffixes like "Careers", "Recruiting", "Talent"
    name = name.replace(/\s*(Careers|Recruiting|Talent|HR|Jobs|Hiring|Team)\s*$/i, "").trim();
    if (name.length > 2 && name.length < 50) {
      return name;
    }
  }

  return "";
}

function extractRole_(subject, body) {
  // Try patterns like "for the [Role] position" or "role: [Role]"
  const patterns = [
    /(?:for the|for our)\s+(.+?)\s+(?:position|role|opening)/i,
    /(?:position|role|job):\s*(.+?)(?:\n|$)/i,
    /applied (?:for|to)\s+(?:the\s+)?(.+?)(?:\s+position|\s+role|\s+at|\s*[.\n])/i,
    /(?:your application for)\s+(.+?)(?:\s+has|\s+was|\s*[.\n])/i,
    /(?:regarding the)\s+(.+?)\s+(?:position|role|opening)/i,
  ];

  for (const pattern of patterns) {
    const match = (subject + "\n" + body).match(pattern);
    if (match && match[1].length < 80) {
      return match[1].trim();
    }
  }

  return "";
}

function extractSource_(from, body) {
  const text = (from + " " + body).toLowerCase();
  if (text.includes("linkedin")) return "LinkedIn";
  if (text.includes("indeed")) return "Indeed";
  if (text.includes("greenhouse")) return "Greenhouse";
  if (text.includes("lever.co")) return "Lever";
  if (text.includes("wellfound") || text.includes("angellist")) return "Wellfound";
  if (text.includes("glassdoor")) return "Glassdoor";
  if (text.includes("handshake")) return "Handshake";
  if (text.includes("ziprecruiter")) return "ZipRecruiter";
  return "Direct";
}


// ============================================================
// SHEET OPERATIONS
// ============================================================

function addApplication_(sheet, info, messageId) {
  const existing = findRow_(sheet, info.company);
  if (existing) return; // Already tracked

  sheet.appendRow([
    info.company,
    info.role,
    Utilities.formatDate(info.date, Session.getScriptTimeZone(), "yyyy-MM-dd"),
    "",          // Rejection
    "",          // Date Rejected
    "Applied",   // Interview Stage
  ]);

  markProcessed_(messageId);
}

function updateStatus_(sheet, company, status, messageId, subject) {
  const data = sheet.getDataRange().getValues();
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
  const companyLower = company.toLowerCase();

  // Find matching row: exact match first, then contains match
  let matchRow = -1;
  for (let i = 1; i < data.length; i++) {
    if (data[i][3]) continue; // Already rejected, skip
    const rowCompany = String(data[i][0]).toLowerCase();
    if (rowCompany === companyLower) {
      matchRow = i;
      break;
    }
    // Fuzzy: one contains the other (handles "TestCompany" vs "Testcompany Inc")
    if (rowCompany.includes(companyLower) || companyLower.includes(rowCompany)) {
      matchRow = i;
      break;
    }
  }

  if (matchRow >= 0 && status === "Rejected") {
    sheet.getRange(matchRow + 1, 4).setValue("Yes");   // Rejection
    sheet.getRange(matchRow + 1, 5).setValue(today);   // Date Rejected
    markProcessed_(messageId);
    return;
  }

  // If no existing row found, add as a new rejected entry
  sheet.appendRow([
    company,
    "",      // Role unknown
    "",      // Date applied unknown
    "Yes",   // Rejection
    today,   // Date Rejected
    "",      // Interview Stage unknown
  ]);
  markProcessed_(messageId);
}

function findRow_(sheet, company) {
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toLowerCase() === company.toLowerCase()) {
      return i + 1;
    }
  }
  return null;
}


// ============================================================
// TRACKING PROCESSED EMAILS
// ============================================================

function getProcessedMessageIds_(sheet) {
  const props = PropertiesService.getScriptProperties();
  const idsJson = props.getProperty("processedIds") || "[]";
  return new Set(JSON.parse(idsJson));
}

function markProcessed_(messageId) {
  const props = PropertiesService.getScriptProperties();
  const idsJson = props.getProperty("processedIds") || "[]";
  const ids = JSON.parse(idsJson);
  ids.push(messageId);
  // Keep only last 500 IDs to avoid hitting property size limits
  if (ids.length > 500) ids.splice(0, ids.length - 500);
  props.setProperty("processedIds", JSON.stringify(ids));
}


// ============================================================
// HELPERS
// ============================================================

function isApplicationEmail_(body) {
  const lower = body.toLowerCase();
  return APPLICATION_KEYWORDS.some(kw => lower.includes(kw));
}

function isRejectionEmail_(body) {
  const lower = body.toLowerCase();
  return REJECTION_KEYWORDS.some(kw => lower.includes(kw));
}

function capitalize_(str) {
  return str.charAt(0).toUpperCase() + str.slice(1);
}

function getOrCreateSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(["Company", "Role", "Date", "Rejection", "Date Rejected", "Interview Stage"]);
    sheet.getRange(1, 1, 1, 6).setFontWeight("bold");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function deleteExistingTriggers_() {
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === "processNewEmails") {
      ScriptApp.deleteTrigger(trigger);
    }
  }
}


// ============================================================
// MANUAL ACTIONS
// ============================================================

/**
 * Run this to backfill from older emails (searches last 30 days).
 * Useful for initial setup to catch applications you already sent.
 */
function backfill() {
  const sheet = getOrCreateSheet_();
  const processed = getProcessedMessageIds_(sheet);
  let count = 0;

  const queries = [
    'subject:"application received" newer_than:30d',
    'subject:"thank you for applying" newer_than:30d',
    'subject:"application confirmation" newer_than:30d',
    'subject:"we received your application" newer_than:30d',
    'subject:"thanks for applying" newer_than:30d',
    'subject:"unfortunately" newer_than:30d',
    'subject:"not moving forward" newer_than:30d',
    'subject:"decided to move forward with other" newer_than:30d',
    'subject:"after careful consideration" newer_than:30d',
  ];

  for (const query of queries) {
    const threads = GmailApp.search(query, 0, 50);
    for (const thread of threads) {
      for (const msg of thread.getMessages()) {
        if (processed.has(msg.getId())) continue;

        const body = msg.getPlainBody();
        const info = extractApplicationInfo_(msg);
        if (!info.company) continue;

        if (isRejectionEmail_(body)) {
          updateStatus_(sheet, info.company, "Rejected", msg.getId(), msg.getSubject());
          count++;
        } else if (isApplicationEmail_(body)) {
          addApplication_(sheet, info, msg.getId());
          count++;
        }
      }
    }
  }

  Logger.log("Backfill complete. Processed " + count + " emails.");
  SpreadsheetApp.getUi().alert("Backfill complete! Processed " + count + " emails.");
}

/**
 * Manually add an application (use from the sheet's custom menu).
 */
function manualAdd() {
  const ui = SpreadsheetApp.getUi();
  const company = ui.prompt("Company name:").getResponseText();
  if (!company) return;
  const role = ui.prompt("Role/Title:").getResponseText();
  const stage = ui.prompt("Interview stage (Applied, Phone Screen, Onsite, etc):").getResponseText() || "Applied";

  const sheet = getOrCreateSheet_();
  sheet.appendRow([
    company,
    role,
    Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd"),
    "",     // Rejection
    "",     // Date Rejected
    stage,
  ]);
}

/**
 * Add a custom menu to the sheet for manual actions.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Job Tracker")
    .addItem("Add Application Manually", "manualAdd")
    .addItem("Backfill Last 30 Days", "backfill")
    .addItem("Run Check Now", "processNewEmails")
    .addSeparator()
    .addItem("Setup (first time)", "setup")
    .toMenu();
}
