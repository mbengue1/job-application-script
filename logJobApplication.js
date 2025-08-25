
/***** CONFIG *****/
const SHEET_NAME = "Internship Tracker Template"; // your main tracker
const LOG_SHEET_NAME = "Log";
const THREAD_HEADER = "Thread ID";               // hidden column we use for de-dupe

// Column indices (1-based) in your sheet
// Based on your description: Progress, Role, Company, Term, Location, Recruiters, Rounds, ThankYou, Offer, DateApplied, Platform, ThreadID
const COL = {
  PROGRESS: 1,      // Progress
  ROLE: 2,          // Role  
  COMPANY: 3,       // Company (moved from end)
  TERM: 4,          // Term
  LOCATION: 5,      // Location
  RECRUITERS: 6,    // @ Recruiters?
  ROUND1: 7,        // First Round
  ROUND2: 8,        // Second Round
  ROUND3: 9,        // Third Round
  THANKYOU: 10,     // Thank You Email?
  OFFER: 11,        // Offer
  DATE_APPLIED: 12, // Date Applied
  PLATFORM: 13,     // Platform
  THREAD_ID: 14,    // Thread ID (hidden)
};

// Broad search for application confirmations (subject OR body, last 2 days)
const GMAIL_QUERY = 'newer_than:2d -label:"Jobs/Processed" (' +
  '"application received" OR "job application" OR "we received your application" OR ' +
  '"thanks for your application" OR "your application is on the way" OR ' +
  '"job application submitted" OR "application submitted" OR "submission confirmation" OR ' +
  '"you applied to" OR "we\'ve received your application" OR "application confirmation" OR ' +
  '"has been submitted" OR "received your submission" OR "application acknowledgment" OR ' +
  '"you\'re in!" OR "thank you for applying" OR "application complete" OR ' +
  '"job interest received" OR "Job Application:" OR ' +
  '"thank you for your application to" OR "we have received your application" OR ' +
  '"your application has been received" OR "application received and reviewed" OR ' +
  '"thank you for applying to" OR "we\'re excited that you are interested" OR ' +
  '"what happens next" OR "we will review your application")';

/***** MAIN *****/
function logJobApplications() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error(`Sheet "${SHEET_NAME}" not found`);

  const logSheet = ss.getSheetByName(LOG_SHEET_NAME) || ss.insertSheet(LOG_SHEET_NAME);

  // Track last processed timestamp to avoid missing emails
  const props = PropertiesService.getScriptProperties();
  const lastProcessed = props.getProperty('lastProcessed');
  const since = lastProcessed ? `after:${lastProcessed} ` : '';
  const query = since + GMAIL_QUERY;

  // Ensure hidden "Thread ID" column exists and is hidden
  const threadColIndex = ensureThreadIdColumn(sheet);

  // Snapshot headers + rows
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  // Read thread IDs already stored to avoid re-adding
  let existingThreadIds = [];
  if (lastRow > 1) {
    existingThreadIds = sheet.getRange(2, threadColIndex, lastRow - 1).getValues().flat().filter(Boolean);
  }

  // Build quick maps for fallback upsert by normalized role+company (for older rows without thread id)
  const existingMap = buildExistingMaps_(sheet, lastRow, lastCol);

  // Search Gmail with timestamp tracking
  const threads = GmailApp.search(query);
  let added = 0;

  // To avoid double handling same thread in one run
  const handledThreads = new Set();

  threads.forEach(thread => {
    const threadId = thread.getId();
    if (handledThreads.has(threadId)) return;
    handledThreads.add(threadId);

    const messages = thread.getMessages();
    const message = messages[messages.length - 1]; // latest
    const body = safePlainText_(message);
    const subject = message.getSubject() || "";
    const from = message.getFrom() || "";
    const appliedDate = message.getDate();

    // Extract fields
    let role = extractRole_(subject, body);          // may be "Unknown"
    role = normalizeRoleName_(role);                 // strip codes/parentheses/noise
    const company = extractCompany_(from, subject, body);
    const term = extractTerm_(subject, body) || "Spring 2026";
    const location = extractLocation_(body) || "Not Specified";

    // Default row (respecting your dropdowns and COL order)
    const appliedDateStr = Utilities.formatDate(appliedDate || new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    
    const rowData = [
      "In Progress",                 // 1  PROGRESS
      role || "Unknown",             // 2  ROLE
      company || "Unknown",          // 3  COMPANY
      term || "Spring 2026",         // 4  TERM
      location || "Not Specified",   // 5  LOCATION
      "No",                          // 6  RECRUITERS
      "No",                          // 7  ROUND1
      "No",                          // 8  ROUND2
      "No",                          // 9  ROUND3
      "No",                          // 10 THANKYOU
      "No",                          // 11 OFFER
      appliedDateStr,                // 12 DATE_APPLIED
      "Email",                       // 13 PLATFORM
      threadId                       // 14 THREAD_ID
    ];

    // Debug logging
    logSheet.appendRow([new Date().toLocaleString(), `🔍 Processing: Role="${role}", Company="${company}", Term="${term}", Location="${location}"`]);

    // 1) Perfect de-dupe: Thread ID already in sheet?
    if (existingThreadIds.includes(threadId)) {
      const rowIndex = findRowByThreadId_(sheet, threadId, threadColIndex, lastRow);
      if (rowIndex) {
        const before = sheet.getRange(rowIndex, 1, 1, lastCol).getValues()[0];
        const changes = upsertRow_(sheet, rowIndex, before, rowData);
        applyColorFormatting(sheet, rowIndex, sheet.getRange(rowIndex, 1, 1, lastCol).getValues()[0]);
        logSheet.appendRow([new Date().toLocaleString(), `📝 Updated by ThreadID ${threadId} — ${changes || "no field changes"}`]);
      } else {
        // Thread ID allegedly exists but row not found (rare) — treat as new append
        appendAndStyle_(sheet, rowData);
        existingThreadIds.push(threadId);
        added++;
        logSheet.appendRow([new Date().toLocaleString(), `✅ Added (recovered) — thread ${threadId}`]);
      }
      // Add processing label to avoid re-scanning
      addProcessingLabel_(thread);
      return;
    }

    // 2) Fallback de-dupe: Match by normalized role + company (if role known)
    const roleKey = normalizeKey_(role || "Unknown");
    const companyKey = normalizeKey_(company || "Unknown");

    let rowIndexToUpdate = null;

    if (role && role.toLowerCase() !== "unknown") {
      const key = roleKey + "|" + companyKey;
      if (existingMap.keyToRow[key]) {
        rowIndexToUpdate = existingMap.keyToRow[key];
      }
    } else {
      // If role unknown, try fallback: company + same applied date (reduces dupes)
      const dateKey = normalizeKey_(rowData[COL.DATE_APPLIED - 1]);
      const keyUnknown = "unknown|" + companyKey + "|" + dateKey;
      if (existingMap.unknownKeyToRow[keyUnknown]) {
        rowIndexToUpdate = existingMap.unknownKeyToRow[keyUnknown];
      }
    }

    if (rowIndexToUpdate) {
      // Update existing row, set Thread ID
      const before = sheet.getRange(rowIndexToUpdate, 1, 1, lastCol).getValues()[0];
      // Always set thread id if missing
      if (!before[COL.THREAD_ID - 1]) {
        sheet.getRange(rowIndexToUpdate, COL.THREAD_ID).setValue(threadId);
        existingThreadIds.push(threadId);
      }
      const changes = upsertRow_(sheet, rowIndexToUpdate, before, rowData);
      applyColorFormatting(sheet, rowIndexToUpdate, sheet.getRange(rowIndexToUpdate, 1, 1, lastCol).getValues()[0]);
      logSheet.appendRow([new Date().toLocaleString(), `📝 Updated by Key (${role || "Unknown"}|${company || "Unknown"}) — ${changes || "no field changes"}`]);
      // Add processing label to avoid re-scanning
      addProcessingLabel_(thread);
      return;
    }

          // 3) If we reach here, it's a brand-new application — append it (even if role is Unknown)
      appendAndStyle_(sheet, rowData);
      existingThreadIds.push(threadId);
      // Update in-memory map so same run won't append twice
      const newLastRow = sheet.getLastRow();
      if (role && role.toLowerCase() !== "unknown") {
        existingMap.keyToRow[roleKey + "|" + companyKey] = newLastRow;
      } else {
        const dateKey = normalizeKey_(rowData[COL.DATE_APPLIED - 1]);
        existingMap.unknownKeyToRow["unknown|" + companyKey + "|" + dateKey] = newLastRow;
      }
      added++;
      logSheet.appendRow([new Date().toLocaleString(), `✅ Added "${rowData[COL.ROLE - 1]}" at ${rowData[COL.COMPANY - 1]} (thread ${threadId})`]);
      // Add processing label to avoid re-scanning
      addProcessingLabel_(thread);
    });

  if (added === 0) {
    logSheet.appendRow([new Date().toLocaleString(), "ℹ️ No new unique applications added (all matched existing by ThreadID or Key)."]);
  }

  // Update last processed timestamp
  props.setProperty('lastProcessed', Math.floor(Date.now()/1000).toString());
}

/***** HELPERS *****/

// Ensure hidden Thread ID column exists; return its index
function ensureThreadIdColumn(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  let idx = headers.indexOf(THREAD_HEADER) + 1;
  if (idx === 0) {
    // Append new header
    const newCol = headers.length + 1;
    sheet.getRange(1, newCol).setValue(THREAD_HEADER);
    // Hide it
    sheet.hideColumns(newCol);
    idx = newCol;
  } else {
    // Make sure it's hidden
    try { sheet.hideColumns(idx); } catch (e) {}
  }
  return idx;
}

// Build maps for dedupe by role|company and unknown role fallback
function buildExistingMaps_(sheet, lastRow, lastCol) {
  const res = { keyToRow: {}, unknownKeyToRow: {} };
  if (lastRow <= 1) return res;

  const values = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

  for (let i = 0; i < values.length; i++) {
    const r = values[i];
    const rowIndex = i + 2;
    const role = normalizeKey_(String(r[COL.ROLE - 1] || "Unknown"));
    const company = normalizeKey_(String(r[COL.COMPANY - 1] || "Unknown"));
    const dateApplied = normalizeKey_(String(r[COL.DATE_APPLIED - 1] || ""));

    if (role !== "unknown") {
      res.keyToRow[role + "|" + company] = rowIndex;
    } else {
      // fallback unknown-role key
      res.unknownKeyToRow["unknown|" + company + "|" + dateApplied] = rowIndex;
    }
  }
  return res;
}

// Find row by threadId in our hidden column
function findRowByThreadId_(sheet, threadId, threadColIndex, lastRow) {
  if (lastRow <= 1) return null;
  const range = sheet.getRange(2, threadColIndex, lastRow - 1);
  const vals = range.getValues().flat();
  const pos = vals.indexOf(threadId);
  return pos >= 0 ? pos + 2 : null;
}

// Update only placeholders/defaults. Return a summary of changes.
function upsertRow_(sheet, rowIndex, before, incoming) {
  const changes = [];
  
  // Progress: ensure proper capitalization using validation helper
  if ((before[COL.PROGRESS - 1] || "").toLowerCase() !== "in progress") {
    const progressValue = coerceProgressValue_(sheet, "In Progress");
    sheet.getRange(rowIndex, COL.PROGRESS).setValue(progressValue);
    changes.push(`Progress→${progressValue}`);
  }

  // Role: Replace role if previous is Unknown OR new one is strictly longer and shares prefix
  const prevRole = String(before[COL.ROLE - 1] || "").trim();
  const newRole = String(incoming[COL.ROLE - 1] || "").trim();
  if (!prevRole || prevRole.toLowerCase() === "unknown" ||
      (newRole && newRole.toLowerCase() !== "unknown" && newRole.length > prevRole.length && newRole.toLowerCase().startsWith(prevRole.toLowerCase()))) {
    sheet.getRange(rowIndex, COL.ROLE).setValue(newRole);
    changes.push(`Role→${newRole}`);
  }

  // Company: if Unknown and we now have a better value
  const prevCompany = String(before[COL.COMPANY - 1] || "").trim();
  const newCompany = String(incoming[COL.COMPANY - 1] || "").trim();
  if (!prevCompany || prevCompany.toLowerCase() === "unknown" ||
      (newCompany && newCompany.toLowerCase() !== "unknown" && newCompany !== prevCompany)) {
    sheet.getRange(rowIndex, COL.COMPANY).setValue(newCompany);
    changes.push(`Company→${newCompany}`);
  }

  // Term: if empty or default
  const prevTerm = String(before[COL.TERM - 1] || "").trim();
  const newTerm = String(incoming[COL.TERM - 1] || "").trim();
  if (!prevTerm || prevTerm === "Spring 2026" || 
      (newTerm && newTerm !== prevTerm && newTerm.toLowerCase() !== "unknown")) {
    sheet.getRange(rowIndex, COL.TERM).setValue(newTerm);
    changes.push(`Term→${newTerm}`);
  }

  // Location: if empty or default
  const prevLocation = String(before[COL.LOCATION - 1] || "").trim();
  const newLocation = String(incoming[COL.LOCATION - 1] || "").trim();
  if (!prevLocation || prevLocation === "Not Specified" ||
      (newLocation && newLocation !== prevLocation && newLocation.toLowerCase() !== "not specified")) {
    sheet.getRange(rowIndex, COL.LOCATION).setValue(newLocation);
    changes.push(`Location→${newLocation}`);
  }

  // Date Applied: if empty or different format
  const prevDate = String(before[COL.DATE_APPLIED - 1] || "").trim();
  const newDate = String(incoming[COL.DATE_APPLIED - 1] || "").trim();
  if (!prevDate || (newDate && newDate !== prevDate)) {
    sheet.getRange(rowIndex, COL.DATE_APPLIED).setValue(newDate);
    changes.push(`Date Applied→${newDate}`);
  }

  // Thread ID: always set if missing
  if (!before[COL.THREAD_ID - 1] && incoming[COL.THREAD_ID - 1]) {
    sheet.getRange(rowIndex, COL.THREAD_ID).setValue(incoming[COL.THREAD_ID - 1]);
    changes.push(`ThreadID set`);
  }

  return changes.join(", ");
}

function appendAndStyle_(sheet, rowData) {
  sheet.appendRow(rowData);
  const rowIndex = sheet.getLastRow();
  applyColorFormatting(sheet, rowIndex, rowData);
}

/***** PARSERS *****/

function safePlainText_(message) {
  try {
    const plain = message.getPlainBody();
    if (plain && plain.trim()) return plain;
    const html = message.getBody() || "";
    return html.replace(/<style[\s\S]*?<\/style>/gi, "")
               .replace(/<script[\s\S]*?<\/script>/gi, "")
               .replace(/<\/?[^>]+>/g, " ")
               .replace(/&nbsp;/g, " ")
               .replace(/\s{2,}/g, " ")
               .trim();
  } catch (e) {
    const html = message.getBody() || "";
    return html.replace(/<\/?[^>]+>/g, " ").trim();
  }
}

function extractRole_(subject, body) {
  const S = subject || "", B = body || "", T = S + "\n" + B;

  let m;
  // "Application received – Company – Role"
  m = S.match(/Application (?:received|submitted)[^\-–—]*[-–—]\s*[^-–—]+[-–—]\s*([^.\n]+)/i);
  if (m) return normalizeRoleName_(m[1]);

  // "Thanks for applying to Company – Role"
  m = S.match(/Thanks for applying to .*?[-–—]\s*([^.\n]+)/i);
  if (m) return normalizeRoleName_(m[1]);

  // "Your application to Company for Role"
  m = T.match(/application to [^\n]+ for ([^.\n]+)\b/i);
  if (m) return normalizeRoleName_(m[1]);

  // "received your (job )?application for Role"
  m = B.match(/received your (?:job )?application for ([^.\n]+)\b/i);
  if (m) return normalizeRoleName_(m[1]);

  // "Position: Role"
  m = T.match(/(?:Position|Job Title)\s*:\s*([^.\n]+)/i);
  if (m) return normalizeRoleName_(m[1]);

  // Workday style: "Job Application: … - 4010 - Software Engineer Co-Op (Jan 2026…)"
  m = T.match(/Job Application:.*?-\s*\d{3,8}\s*-\s*([^(]+)\(/i);
  if (m) return normalizeRoleName_(m[1]);

  // GE Appliances-style "Your application for <role> is on the way"
  m = B.match(/Your application for ([^.\n]+?) is on the way/i);
  if (m) return normalizeRoleName_(m[1]);

  // Subject variants "Application Received: <role>"
  m = S.match(/Application (?:Received|Submitted):\s*([^.\n]+)/i);
  if (m) return normalizeRoleName_(m[1]);

  // Try to find role in subject if it contains common role keywords
  const roleKeywords = /\b(?:intern|co.?op|coop|technician|engineer|developer|analyst|assistant|specialist|coordinator|manager|director|consultant|advisor|representative|associate|assistant|clerk)\b/i;
  if (roleKeywords.test(S)) {
    const words = S.split(/[^a-zA-Z0-9\s\-]/);
    for (const word of words) {
      if (roleKeywords.test(word) && word.length > 5) {
        return normalizeRoleName_(word);
      }
    }
  }

  // Look for role patterns in body
  m = B.match(/\b(?:position|role|job|title)\s*[:\-]\s*([^.\n\r]+)/i);
  if (m) return normalizeRoleName_(m[1]);

  // Try to extract from email signature or footer
  m = B.match(/(?:Best regards|Sincerely|Thanks)[^\n]*\n([^.\n]+)/i);
  if (m && roleKeywords.test(m[1])) return normalizeRoleName_(m[1]);

  // For emails that don't specify role but are application confirmations
  if (B.includes("we have received your application") || B.includes("we've received your application")) {
    // Look for any role-like text in the email
    m = B.match(/\b(?:position|role|job|title|opportunity)\s*(?:of|as)?\s*([^.\n]+)/i);
    if (m) return normalizeRoleName_(m[1]);
    
    // Look for role keywords in context
    m = B.match(/(?:interested in|looking for|applying for)\s+([^.\n]+)/i);
    if (m && roleKeywords.test(m[1])) return normalizeRoleName_(m[1]);
  }

  return "Unknown";
}

function normalizeRoleName_(role) {
  if (!role) return "Unknown";
  let r = String(role).trim();

  // Strip job IDs like “- 4010”
  r = r.replace(/\s[-–—_/.,()+:'’"]/g, "");

  // Remove trailing parenthetical like “(Jan. 2026 Start)”
  r = r.replace(/\s*\([^)]*\)\s*$/g, "");

  // Collapse extra spaces
  r = r.replace(/\s{2,}/g, " ").trim();

  return r || "Unknown";
}

function extractTerm_(subject, body) {
  const txt = subject + "\n" + body;

  // Explicit terms like “Spring 2026”
  let m = txt.match(/\b(Spring|Summer|Fall|Winter)\s?20\d{2}\b/i);
  if (m) return m[0].replace(/\s{2,}/g, " ").trim();

  // Month ranges e.g., "(January 2025 - May 2025)", "Jan. 2026 Start"
  // Map first month to term
  const monthRegex = /(Jan(?:uary)?\.?|Feb(?:ruary)?\.?|Mar(?:ch)?\.?|Apr(?:il)?\.?|May\.?|Jun(?:e)?\.?|Jul(?:y)?\.?|Aug(?:ust)?\.?|Sep(?:t(?:ember)?)?\.?|Oct(?:ober)?\.?|Nov(?:ember)?\.?|Dec(?:ember)?\.?)\s*\.?\s*(20\d{2})/i;
  m = txt.match(monthRegex);
  if (m) {
    const month = m[1].toLowerCase();
    const year = m[2];
    const spring = ["jan", "jan.", "january", "feb", "feb.", "february", "mar", "mar.", "march", "apr", "apr.", "april", "may"];
    const summer = ["jun", "jun.", "june", "jul", "jul.", "july", "aug", "aug.", "august"];
    if (spring.some(k => month.startsWith(k))) return `Spring ${year}`;
    if (summer.some(k => month.startsWith(k))) return `Summer ${year}`;
    return `Fall ${year}`;
  }

  // Year-Round hint
  if (/\byear[-\s]?round\b/i.test(txt)) {
    const y = (txt.match(/20\d{2}/) || [new Date().getFullYear()])[0];
    return `Year-Round ${y}`;
  }

  return null;
}

function extractLocation_(text) {
  const T = text || "";
  let m = T.match(/Location:\s*([^\n]+)/i);
  if (m) return m[1].trim();
  m = T.match(/\bbased in\s+([A-Za-z .\-]+,\s*[A-Z]{2})\b/i);
  if (m) return m[1].trim();
  m = T.match(/\b(?:work location|primary location)\s*:\s*([^\n]+)/i);
  if (m) return m[1].trim();
  return null;
}

function extractCompany_(from, subject, body) {
  const S = subject || "", B = body || "", T = S + "\n" + B;
  const ats = /myworkday|workday|greenhouse|lever|smartrecruiters|successfactors|workable|icims|oraclecloud|ultipro|adp|bamboohr|jazzhr|jobvite|hire/i;
  
  // Skip generic domains that shouldn't be treated as company names
  const genericDomains = /google|gmail|yahoo|outlook|hotmail|icloud|aol|protonmail|zoho|yandex|mail|noreply|no[-_.]?reply|info|support|admin|help|contact|hello|hi/i;

  // 1) Display name in From: "LiveRamp Careers <no-reply@myworkday.com>"
  const nameMatch = from && from.match(/^"?([^"<]+?)"?\s*<[^>]+>/);
  if (nameMatch) {
    const display = nameMatch[1].trim();
    // Try "X Careers", "X Recruiting", "X Talent"
    const nm = display.replace(/(Careers|Recruiting|Talent|HR|Human Resources)\b/i, "").trim();
    if (nm && !/noreply|no[-_.]?reply|info|support/i.test(nm)) return nm;
  }

  // 2) Subject "Thank you for your application to Company"
  let m = S.match(/Thank you for your application to ([^.\n]+)/i);
  if (m) return m[1].trim();

  // 3) Subject "Application received – Company – Role"
  m = S.match(/Application (?:received|submitted)\s*[-–—]\s*([^–—-]+)\s*[-–—]/i);
  if (m) return m[1].trim();

  // 4) "Thanks for applying to Company"
  m = T.match(/Thanks for applying to ([^\n–—-]+?)(?:\s*[-–—]| for\b|\.|\n)/i);
  if (m) return m[1].trim();

  // 5) "Your application to Company for Role"
  m = T.match(/application to ([^\n]+?) for [^\n.()]+/i);
  if (m) return m[1].trim();

  // 6) "We have received your application" + company context
  if (B.includes("we have received your application") || B.includes("we've received your application")) {
    // Look for company name in the email body
    m = B.match(/(?:at|to|for|about)\s+([A-Z][a-z]+(?:\s+[A-Z][a-z]+)*)/i);
    if (m) return m[1].trim();
    
    // Look for company name in signature or context
    m = B.match(/(?:team|recruitment team|all the best)[^.\n]*([A-Z][a-z]+(?:\s+[A-Z][a-z]+)*)/i);
    if (m) return m[1].trim();
  }

  // 7) "Careers at Company"
  m = B.match(/Careers at ([A-Za-z0-9 &\-\.'']+)(?=[,\n]|$)/i);
  if (m) return m[1].trim();

  // 8) Signature block "Sincerely,\nCompany"
  m = B.match(/(?:Thanks|Sincerely|Regards|All the best)[^\n]*\n([A-Za-z0-9 &\-\.'']+)(?=[,\n]|$)/i);
  if (m) return m[1].trim();

  // 9) Look for company names in subject (common patterns)
  const companyPatterns = [
    /(?:at|for|to)\s+([A-Z][a-z]+(?:\s+[A-Z][a-z]+)*)/i,
    /([A-Z][a-z]+(?:\s+[A-Z][a-z]+)*)\s+(?:Intern|Co.?op|Coop|Technician|Engineer)/i
  ];
  
  for (const pattern of companyPatterns) {
    m = S.match(pattern);
    if (m && m[1].length > 2 && m[1].length < 50) {
      return m[1].trim();
    }
  }

  // 10) Fallback from domain if not ATS and not generic
  const dm = from && from.match(/@([a-z0-9\-]+)\.[a-z.]+/i);
  const domain = dm ? dm[1].toLowerCase() : "";
  if (domain && !ats.test(domain) && !genericDomains.test(domain)) {
    if (domain.length > 3 && !/^[a-z]+$/.test(domain)) {
      const cleanDomain = domain.replace(/[-_]/g, " ").replace(/\b\w/g, l => l.toUpperCase());
      return cleanDomain;
    }
  }

  // 11) Fallback local part if looks brand-like (but not generic)
  const local = (from || "").split("@")[0];
  if (local && local.length <= 20 && !genericDomains.test(local) && local.length > 3) {
    const cleanLocal = local.replace(/[-_.]/g, " ").replace(/\b\w/g, l => l.toUpperCase());
    return cleanLocal;
  }

  return "Unknown";
}

/***** FORMATTING *****/
function applyColorFormatting(sheet, rowIndex, rowData) {
  const colors = {
    inProgress: "#FFF9C4", // yellow
    yes: "#C8E6C9"         // green
  };

  // Read current progress as-is (don't overwrite)
  const progress = String(sheet.getRange(rowIndex, COL.PROGRESS).getValue() || "").trim().toLowerCase();

  // Color if "in progress" (typo-safe)
  if (progress === "in progress" || progress === "in progess") {
    sheet.getRange(rowIndex, COL.PROGRESS).setBackground(colors.inProgress);
  } else {
    sheet.getRange(rowIndex, COL.PROGRESS).setBackground(null);
  }

  // YES fields: Recruiters?, Rounds, Thank You?, Offer
  [COL.RECRUITERS, COL.ROUND1, COL.ROUND2, COL.ROUND3, COL.THANKYOU, COL.OFFER].forEach(c => {
    const val = String(sheet.getRange(rowIndex, c).getValue() || "").trim();
    sheet.getRange(rowIndex, c).setBackground(val === "Yes" ? colors.yes : null);
  });
}

/***** UTIL *****/
function normalizeKey_(s) {
  return String(s || "")
    .toLowerCase()
    .replace(/[\s\-–—_/.,()+:'’"]/g, "")
    .trim();
}

function capitalize_(s) {
  if (!s) return s;
  return s.charAt(0).toUpperCase() + s.slice(1);
}

// Helper function to handle data validation mismatches
function coerceProgressValue_(sheet, desired) {
  try {
    // Look at A2's validation rule (first data row)
    const rule = sheet.getRange(2, COL.PROGRESS).getDataValidation();
    if (!rule) return desired;
    const critVals = rule.getCriteriaValues();
    const list = (critVals && critVals[0]) || [];
    // If the sheet only allows "In progess", map to that.
    if (desired === "In Progress" && list.indexOf("In progess") !== -1) return "In progess";
    return list.includes(desired) ? desired : (list[0] || desired);
  } catch (e) {
    // Fallback if validation rule can't be read
    return desired;
  }
}

// Helper function to add processing label to threads
function addProcessingLabel_(thread) {
  try {
    const labelName = "Jobs/Processed";
    let label = GmailApp.getUserLabelByName(labelName);
    if (!label) {
      label = GmailApp.createLabel(labelName);
    }
    thread.addLabel(label);
  } catch (e) {
    // Silently fail if label creation fails
    console.log("Could not add processing label:", e.message);
  }
}

// Function to clean up existing data and fix column mismatches
function cleanupExistingData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error(`Sheet "${SHEET_NAME}" not found`);

  const logSheet = ss.getSheetByName(LOG_SHEET_NAME) || ss.insertSheet(LOG_SHEET_NAME);
  const lastRow = sheet.getLastRow();
  
  if (lastRow <= 1) {
    logSheet.appendRow([new Date().toLocaleString(), "ℹ️ No data rows to clean up"]);
    return;
  }

  let fixed = 0;
  const values = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  
  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    const rowIndex = i + 2;
    let changes = [];
    
    // Check if Company and Term are swapped
    const role = String(row[COL.ROLE - 1] || "").trim();
    const company = String(row[COL.COMPANY - 1] || "").trim();
    const term = String(row[COL.TERM - 1] || "").trim();
    
    // If Term column contains what looks like a company name, swap them
    if (term && (term.includes("LiveRamp") || term.includes("Midmark") || term.includes("Google") || 
                 term.includes("Myworkday") || term.includes("Haier"))) {
      if (company && company.includes("Spring") || company.includes("Summer") || company.includes("Fall")) {
        // Swap Company and Term
        sheet.getRange(rowIndex, COL.COMPANY).setValue(term);
        sheet.getRange(rowIndex, COL.TERM).setValue(company);
        changes.push(`Swapped Company↔Term: "${term}" ↔ "${company}"`);
      }
    }
    
    // Fix date formatting if it's in the old format
    const dateApplied = String(row[COL.DATE_APPLIED - 1] || "").trim();
    if (dateApplied && dateApplied.includes("Sun") || dateApplied.includes("Mon") || 
        dateApplied.includes("Tue") || dateApplied.includes("Wed") || 
        dateApplied.includes("Thu") || dateApplied.includes("Fri") || 
        dateApplied.includes("Sat")) {
      try {
        const date = new Date(dateApplied);
        if (!isNaN(date.getTime())) {
          const formattedDate = Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
          sheet.getRange(rowIndex, COL.DATE_APPLIED).setValue(formattedDate);
          changes.push(`Fixed date format: "${dateApplied}" → "${formattedDate}"`);
        }
      } catch (e) {
        // Ignore date parsing errors
      }
    }
    
    if (changes.length > 0) {
      fixed++;
      logSheet.appendRow([new Date().toLocaleString(), `🔧 Fixed row ${rowIndex}: ${changes.join(", ")}`]);
    }
  }
  
  if (fixed > 0) {
    logSheet.appendRow([new Date().toLocaleString(), `✅ Cleaned up ${fixed} rows with data issues`]);
  } else {
    logSheet.appendRow([new Date().toLocaleString(), "ℹ️ No data issues found to fix"]);
  }
}

// Function to validate column data placement
function validateColumnData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error(`Sheet "${SHEET_NAME}" not found`);

  const logSheet = ss.getSheetByName(LOG_SHEET_NAME) || ss.insertSheet(LOG_SHEET_NAME);
  const lastRow = sheet.getLastRow();
  
  if (lastRow <= 1) {
    logSheet.appendRow([new Date().toLocaleString(), "ℹ️ No data rows to validate"]);
    return;
  }

  let issues = 0;
  const values = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  
  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    const rowIndex = i + 2;
    let rowIssues = [];
    
    // Check Role column (should not contain company names)
    const role = String(row[COL.ROLE - 1] || "").trim();
    if (role && (role.includes("LiveRamp") || role.includes("Midmark") || role.includes("Google") || 
                 role.includes("Myworkday") || role.includes("Haier"))) {
      rowIssues.push(`Role column contains company name: "${role}"`);
    }
    
    // Check Company column (should not contain terms)
    const company = String(row[COL.COMPANY - 1] || "").trim();
    if (company && (company.includes("Spring") || company.includes("Summer") || company.includes("Fall"))) {
      rowIssues.push(`Company column contains term: "${company}"`);
    }
    
    // Check Term column (should not contain company names)
    const term = String(row[COL.TERM - 1] || "").trim();
    if (term && (term.includes("LiveRamp") || term.includes("Midmark") || term.includes("Google") || 
                 term.includes("Myworkday") || term.includes("Haier"))) {
      rowIssues.push(`Term column contains company name: "${term}"`);
    }
    
    // Check Date Applied column (should be in yyyy-MM-dd format)
    const dateApplied = String(row[COL.DATE_APPLIED - 1] || "").trim();
    if (dateApplied && (dateApplied.includes("Sun") || dateApplied.includes("Mon") || 
        dateApplied.includes("Tue") || dateApplied.includes("Wed") || 
        dateApplied.includes("Thu") || dateApplied.includes("Fri") || 
        dateApplied.includes("Sat"))) {
      rowIssues.push(`Date Applied has wrong format: "${dateApplied}"`);
    }
    
    if (rowIssues.length > 0) {
      issues++;
      logSheet.appendRow([new Date().toLocaleString(), `⚠️ Row ${rowIndex} issues: ${rowIssues.join("; ")}`]);
    }
  }
  
  if (issues > 0) {
    logSheet.appendRow([new Date().toLocaleString(), `⚠️ Found ${issues} rows with column data issues`]);
  } else {
    logSheet.appendRow([new Date().toLocaleString(), "✅ All rows have correct column data placement"]);
  }
}

// Function to verify sheet layout and column mapping
function verifySheetLayout() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error(`Sheet "${SHEET_NAME}" not found`);

  const logSheet = ss.getSheetByName(LOG_SHEET_NAME) || ss.insertSheet(LOG_SHEET_NAME);
  
  // Get headers
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  logSheet.appendRow([new Date().toLocaleString(), "🔍 Sheet Layout Verification:"]);
  logSheet.appendRow([new Date().toLocaleString(), `Total columns: ${headers.length}`]);
  
  // Log each column with its position
  headers.forEach((header, index) => {
    const colNum = index + 1;
    logSheet.appendRow([new Date().toLocaleString(), `Column ${colNum}: "${header}"`]);
  });
  
  // Check if our COL constants match the actual layout
  logSheet.appendRow([new Date().toLocaleString(), ""]);
  logSheet.appendRow([new Date().toLocaleString(), "📋 Expected vs Actual Column Mapping:"]);
  
  const expectedColumns = [
    { name: "Progress", expected: COL.PROGRESS },
    { name: "Role", expected: COL.ROLE },
    { name: "Company", expected: COL.COMPANY },
    { name: "Term", expected: COL.TERM },
    { name: "Location", expected: COL.LOCATION },
    { name: "Recruiters", expected: COL.RECRUITERS },
    { name: "First Round", expected: COL.ROUND1 },
    { name: "Second Round", expected: COL.ROUND2 },
    { name: "Third Round", expected: COL.ROUND3 },
    { name: "Thank You Email", expected: COL.THANKYOU },
    { name: "Offer", expected: COL.OFFER },
    { name: "Date Applied", expected: COL.DATE_APPLIED },
    { name: "Platform", expected: COL.PLATFORM }
  ];
  
  expectedColumns.forEach(col => {
    const actualIndex = headers.findIndex(h => 
      h && h.toString().toLowerCase().includes(col.name.toLowerCase())
    );
    const actualCol = actualIndex >= 0 ? actualIndex + 1 : "NOT FOUND";
    const status = actualCol === col.expected ? "✅" : "❌";
    logSheet.appendRow([new Date().toLocaleString(), `${status} ${col.name}: Expected ${col.expected}, Actual ${actualCol}`]);
  });
  
  logSheet.appendRow([new Date().toLocaleString(), ""]);
  logSheet.appendRow([new Date().toLocaleString(), "💡 If you see ❌ marks, update the COL constants to match your sheet layout"]);
}

