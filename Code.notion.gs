/**
 * Job Application Tracker — Notion Integration
 *
 * Same email scanning as Code.gs, but writes to a Notion database instead of Google Sheets.
 *
 * Setup:
 * 1. Create a Notion integration at https://www.notion.so/my-integrations
 *    - Give it "Insert content" and "Read content" capabilities
 *    - Copy the "Internal Integration Secret" (starts with ntn_)
 * 2. Create a Notion database with these properties:
 *    - Company (Title)
 *    - Role (Text)
 *    - Date Applied (Date)
 *    - Status (Select: Applied, Rejected, Interview, Offer)
 *    - Date Updated (Date)
 *    - Email Subject (Text)
 *    - Source (Select: LinkedIn, Indeed, Greenhouse, Lever, Wellfound, Direct, etc.)
 * 3. Share the database with your integration (click ... → Connections → Add your integration)
 * 4. Copy the database ID from the URL: notion.so/{workspace}/{DATABASE_ID}?v=...
 * 5. In Apps Script: Run → setNotionCredentials() and enter your token + database ID
 * 6. Run → setup() to start the trigger
 */

// ============================================================
// NOTION CONFIG
// ============================================================

function getNotionToken_() {
  return PropertiesService.getScriptProperties().getProperty("NOTION_TOKEN") || "";
}

function getNotionDatabaseId_() {
  return PropertiesService.getScriptProperties().getProperty("NOTION_DATABASE_ID") || "";
}

/**
 * Run this once to store your Notion credentials securely.
 */
function setNotionCredentials() {
  const ui = SpreadsheetApp.getUi();
  const token = ui.prompt("Enter your Notion Integration Token (ntn_...):").getResponseText();
  const dbId = ui.prompt("Enter your Notion Database ID:").getResponseText();

  if (!token || !dbId) {
    ui.alert("Both token and database ID are required.");
    return;
  }

  PropertiesService.getScriptProperties().setProperties({
    "NOTION_TOKEN": token.trim(),
    "NOTION_DATABASE_ID": dbId.trim().replace(/-/g, ""),
  });
  ui.alert("Notion credentials saved! Run setup() to start tracking.");
}


// ============================================================
// NOTION API HELPERS
// ============================================================

const NOTION_API = "https://api.notion.com/v1";
const NOTION_VERSION = "2022-06-28";

function notionRequest_(method, endpoint, payload) {
  const token = getNotionToken_();
  if (!token) throw new Error("Notion token not set. Run setNotionCredentials() first.");

  const options = {
    method: method,
    headers: {
      "Authorization": "Bearer " + token,
      "Notion-Version": NOTION_VERSION,
      "Content-Type": "application/json",
    },
    muteHttpExceptions: true,
  };
  if (payload) options.payload = JSON.stringify(payload);

  const resp = UrlFetchApp.fetch(NOTION_API + endpoint, options);
  const code = resp.getResponseCode();
  if (code >= 400) {
    Logger.log("Notion API error " + code + ": " + resp.getContentText().substring(0, 500));
    return null;
  }
  return JSON.parse(resp.getContentText());
}

function notionQuery_(filter) {
  const dbId = getNotionDatabaseId_();
  return notionRequest_("POST", "/databases/" + dbId + "/query", { filter: filter });
}

function notionCreatePage_(properties) {
  const dbId = getNotionDatabaseId_();
  return notionRequest_("POST", "/pages", {
    parent: { database_id: dbId },
    properties: properties,
  });
}

function notionUpdatePage_(pageId, properties) {
  return notionRequest_("PATCH", "/pages/" + pageId, { properties: properties });
}


// ============================================================
// NOTION OPERATIONS (replacements for sheet operations)
// ============================================================

function addApplicationNotion_(info, messageId) {
  // Check if already exists
  const existing = findNotionRow_(info.company, info.role);
  if (existing) return;

  const dateStr = Utilities.formatDate(info.date, Session.getScriptTimeZone(), "yyyy-MM-dd");

  notionCreatePage_({
    "Company": { title: [{ text: { content: info.company } }] },
    "Role": { rich_text: [{ text: { content: info.role || "" } }] },
    "Date Applied": { date: { start: dateStr } },
    "Status": { select: { name: "Applied" } },
    "Date Updated": { date: { start: dateStr } },
    "Email Subject": { rich_text: [{ text: { content: info.subject || "" } }] },
    "Source": { select: { name: info.source || "Direct" } },
  });

  markProcessed_(messageId);
}

function updateStatusNotion_(company, status, messageId, subject) {
  const existing = findNotionRow_(company);
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");

  if (existing) {
    notionUpdatePage_(existing.id, {
      "Status": { select: { name: status } },
      "Date Updated": { date: { start: today } },
    });
  } else {
    // No existing row — create a new rejected entry
    notionCreatePage_({
      "Company": { title: [{ text: { content: company } }] },
      "Role": { rich_text: [{ text: { content: "" } }] },
      "Status": { select: { name: status } },
      "Date Updated": { date: { start: today } },
      "Email Subject": { rich_text: [{ text: { content: subject || "" } }] },
      "Source": { select: { name: "Direct" } },
    });
  }

  markProcessed_(messageId);
}

function findNotionRow_(company, role) {
  const filter = {
    property: "Company",
    title: { equals: company },
  };

  const result = notionQuery_(filter);
  if (!result || !result.results || result.results.length === 0) return null;

  if (role) {
    // Try to match role too
    for (const page of result.results) {
      const pageRole = page.properties["Role"]?.rich_text?.[0]?.text?.content || "";
      if (pageRole.toLowerCase() === role.toLowerCase()) return page;
    }
  }

  return result.results[0]; // Return first match by company
}


// ============================================================
// MAIN PROCESSING (Notion version)
// ============================================================

function processNewEmailsNotion() {
  const processed = getProcessedMessageIds_();

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
            addApplicationNotion_(info, msg.getId());
          }
        }
      }
    }
  }

  // Process rejections
  for (const query of REJECTION_QUERIES) {
    const threads = GmailApp.search(query, 0, 20);
    for (const thread of threads) {
      for (const msg of thread.getMessages()) {
        if (processed.has(msg.getId())) continue;

        const body = msg.getPlainBody();
        if (isRejectionEmail_(body)) {
          const info = extractApplicationInfo_(msg);
          if (info.company) {
            updateStatusNotion_(info.company, "Rejected", msg.getId(), msg.getSubject());
          }
        }
      }
    }
  }
}


// ============================================================
// SETUP (Notion version)
// ============================================================

function setupNotion() {
  const token = getNotionToken_();
  const dbId = getNotionDatabaseId_();

  if (!token || !dbId) {
    SpreadsheetApp.getUi().alert(
      "Notion credentials not set. Run setNotionCredentials() first."
    );
    return;
  }

  // Verify connection
  const result = notionRequest_("GET", "/databases/" + dbId);
  if (!result) {
    SpreadsheetApp.getUi().alert(
      "Could not connect to Notion database. Check your token and database ID, " +
      "and make sure you've shared the database with your integration."
    );
    return;
  }

  // Create time trigger
  deleteExistingTriggersNotion_();
  ScriptApp.newTrigger("processNewEmailsNotion")
    .timeBased()
    .everyMinutes(CHECK_INTERVAL_MINUTES)
    .create();

  SpreadsheetApp.getUi().alert(
    "Notion setup complete! Connected to: " + (result.title?.[0]?.plain_text || "database") +
    "\nTracker will check for new emails every " + CHECK_INTERVAL_MINUTES + " minutes."
  );
}

function deleteExistingTriggersNotion_() {
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === "processNewEmailsNotion") {
      ScriptApp.deleteTrigger(trigger);
    }
  }
}


// ============================================================
// SHARED FUNCTIONS (same as Code.gs)
// ============================================================

const SHEET_NAME = "Applications";
const CHECK_INTERVAL_MINUTES = 5;

const APPLICATION_QUERIES = [
  'subject:"application received" newer_than:1d',
  'subject:"thank you for applying" newer_than:1d',
  'subject:"application confirmation" newer_than:1d',
  'subject:"we received your application" newer_than:1d',
  'subject:"application submitted" newer_than:1d',
  'subject:"thanks for applying" newer_than:1d',
];

const REJECTION_QUERIES = [
  'subject:"unfortunately" newer_than:1d',
  'subject:"we will not be moving forward" newer_than:1d',
  'subject:"not moving forward" newer_than:1d',
  'subject:"decided to move forward with other" newer_than:1d',
  'subject:"position has been filled" newer_than:1d',
  'subject:"after careful consideration" newer_than:1d',
  'subject:"update on your application" newer_than:1d',
];

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

const APPLICATION_KEYWORDS = [
  "received your application",
  "thank you for applying",
  "application has been submitted",
  "we have received",
  "confirm your application",
  "successfully submitted",
  "thanks for your interest",
];

function extractApplicationInfo_(message) {
  const subject = message.getSubject();
  const from = message.getFrom();
  const body = message.getPlainBody().substring(0, 2000);
  const date = message.getDate();

  let company = extractCompany_(from, subject, body);
  let role = extractRole_(subject, body);
  let source = extractSource_(from, body);

  return { company, role, date, subject, source };
}

function extractCompany_(from, subject, body) {
  const domainMatch = from.match(/@([a-z0-9-]+)\./i);
  if (domainMatch) {
    const domain = domainMatch[1].toLowerCase();
    const generic = ["gmail", "yahoo", "outlook", "hotmail", "greenhouse", "lever", "ashby",
                     "smartrecruiters", "workday", "icims", "jobvite", "myworkdayjobs",
                     "successfactors", "taleo", "brassring"];
    if (!generic.includes(domain)) {
      return capitalize_(domain);
    }
  }

  const atMatch = (subject + " " + body).match(/(?:at|with|from|to)\s+([A-Z][A-Za-z0-9\s&.]+?)(?:\s+for|\s+as|\s*[,.\n!])/);
  if (atMatch) return atMatch[1].trim();

  const nameMatch = from.match(/^"?([^"<]+)"?\s*</);
  if (nameMatch) {
    let name = nameMatch[1].trim();
    name = name.replace(/\s*(Careers|Recruiting|Talent|HR|Jobs|Hiring)\s*$/i, "").trim();
    if (name.length > 2 && name.length < 50) return name;
  }

  return "";
}

function extractRole_(subject, body) {
  const patterns = [
    /(?:for the|for our)\s+(.+?)\s+(?:position|role|opening)/i,
    /(?:position|role|job):\s*(.+?)(?:\n|$)/i,
    /applied (?:for|to)\s+(?:the\s+)?(.+?)(?:\s+position|\s+role|\s+at|\s*[.\n])/i,
    /(?:your application for)\s+(.+?)(?:\s+has|\s+was|\s*[.\n])/i,
    /(?:regarding the)\s+(.+?)\s+(?:position|role|opening)/i,
  ];

  for (const pattern of patterns) {
    const match = (subject + "\n" + body).match(pattern);
    if (match && match[1].length < 80) return match[1].trim();
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

function getProcessedMessageIds_() {
  const props = PropertiesService.getScriptProperties();
  const idsJson = props.getProperty("processedIds") || "[]";
  return new Set(JSON.parse(idsJson));
}

function markProcessed_(messageId) {
  const props = PropertiesService.getScriptProperties();
  const idsJson = props.getProperty("processedIds") || "[]";
  const ids = JSON.parse(idsJson);
  ids.push(messageId);
  if (ids.length > 500) ids.splice(0, ids.length - 500);
  props.setProperty("processedIds", JSON.stringify(ids));
}
