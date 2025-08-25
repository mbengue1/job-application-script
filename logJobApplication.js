function logJobApplications() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Internship Tracker Template");
  const logSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Log") || SpreadsheetApp.getActiveSpreadsheet().insertSheet("Log");

  const threads = GmailApp.search(
    'newer_than:2d ("application received" OR "job application" OR "we received your application" OR "thanks for your application" OR "your application is on the way" OR "job application submitted" OR "application submitted" OR "submission confirmation" OR "you applied to" OR "we’ve received your application" OR "application confirmation" OR "has been submitted" OR "received your submission" OR "application acknowledgment" OR "you’re in!" OR "thank you for applying" OR "application complete" OR "job interest received" OR "Job Application:")'
  );

  const lastRow = sheet.getLastRow();
  let existingRoles = [];
  let existingCompanies = [];

  if (lastRow > 1) {
    existingRoles = sheet.getRange(2, 2, lastRow - 1).getValues().flat();  // Column B = Role
    existingCompanies = sheet.getRange(2, 12, lastRow - 1).getValues().flat(); // Column M = Company
  }

  const addedThreadIds = new Set();
  let entriesAdded = 0;

  threads.forEach(thread => {
    const messages = thread.getMessages();
    const message = messages[messages.length - 1];
    const subject = message.getSubject();
    const body = message.getPlainBody();
    const date = message.getDate();
    const threadId = thread.getId();

    if (addedThreadIds.has(threadId)) return;
    addedThreadIds.add(threadId);

    // --- ROLE Extraction ---
    let role = "Unknown";
    const bodyRoleMatch = body.match(/Job Application:.*?-.*?-\s(.+?)\s\(/i);
    if (bodyRoleMatch) role = bodyRoleMatch[1].trim();
    else {
      const altRoleMatch = body.match(/received your job application for (.+?)(\.|\n)/i);
      if (altRoleMatch) role = altRoleMatch[1].trim();
      else {
        const geRoleMatch = body.match(/Your application for (.+?) is on the way/i);
        if (geRoleMatch) role = geRoleMatch[1].trim();
      }
    }

    // --- TERM Extraction ---
    let term = "Spring 2026";
    const rangeMatch = body.match(/\((January|February|March|April|May|June|July|August|September|October|November|December)\s\d{4} -/i);
    if (rangeMatch) {
      const month = rangeMatch[1];
      if (["January", "February", "March", "April", "May"].includes(month)) term = "Spring 2025";
      else if (["June", "July", "August"].includes(month)) term = "Summer 2025";
      else term = "Fall 2025";
    } else {
      const altTerm = body.match(/(Spring|Summer|Fall|Winter)\s?20\d{2}/i);
      if (altTerm) term = altTerm[0].trim();
    }

    // --- LOCATION Extraction ---
    let location = "Not Specified";
    const locationMatch = body.match(/Location: (.+)/i);
    if (locationMatch) location = locationMatch[1].split('\n')[0].trim();

    // --- COMPANY Extraction ---
    const from = message.getFrom();
    let domain = "unknown";
    const domainMatch = from.match(/@([a-z0-9\-]+)\.com/i);
    if (domainMatch) domain = domainMatch[1];

    let company = domain.charAt(0).toUpperCase() + domain.slice(1);
    if (domain === "myworkday" && from.includes("liveramp")) company = "LiveRamp";

    const alreadyExists = existingRoles.includes(role) && existingCompanies.includes(company);
    const defaultData = [
      "In Progress", role, term, location, "No", "No", "No", "No", "No", "No",
      date.toDateString(), company, "Email"
    ];

    if (!alreadyExists) {
      sheet.appendRow(defaultData);
      const rowIndex = sheet.getLastRow();
      applyColorFormatting(sheet, rowIndex, defaultData);
      logSheet.appendRow([new Date().toLocaleString(), `✅ Added "${role}" at ${company}`]);
      entriesAdded++;
    } else {
      const rowIndex = existingRoles.findIndex((r, i) => r === role && existingCompanies[i] === company) + 2;
      const rowRange = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn());
      const row = rowRange.getValues()[0];

      // Update only placeholder/default values
      let updated = false;
      if (row[1] === "Unknown" || !row[1].toLowerCase().includes("software")) {
        sheet.getRange(rowIndex, 2).setValue(role); updated = true;
      }
      if (row[2] === "" || row[2] === "Spring 2026") {
        sheet.getRange(rowIndex, 3).setValue(term); updated = true;
      }
      if (row[3] === "" || row[3] === "Not Specified") {
        sheet.getRange(rowIndex, 4).setValue(location); updated = true;
      }

      const updatedRow = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn()).getValues()[0];
      applyColorFormatting(sheet, rowIndex, updatedRow);

      if (updated) {
        logSheet.appendRow([new Date().toLocaleString(), `📝 Updated "${role}" at ${company}"`]);
      } else {
        logSheet.appendRow([new Date().toLocaleString(), `⚠️ Skipped "${role}" at ${company}" — Already exists with no new data`]);
      }
    }
  });

  if (entriesAdded === 0) {
    logSheet.appendRow([new Date().toLocaleString(), `❌ No new applications added.`]);
  }
}

function applyColorFormatting(sheet, rowIndex, rowData) {
  const colors = {
    "In Progress": "#FFF9C4", // Yellow
    "Yes": "#C8E6C9"          // Green
  };

  const progressCell = sheet.getRange(rowIndex, 1);
  if (rowData[0] === "In Progress") progressCell.setBackground(colors["In Progress"]);

  const yesCols = [5, 6, 7, 8, 9, 10]; // Recruiters?, Rounds, Thank You?, Offer
  yesCols.forEach(col => {
    const val = rowData[col - 1];
    const cell = sheet.getRange(rowIndex, col);
    if (val === "Yes") {
      cell.setBackground(colors["Yes"]);
    } else {
      cell.setBackground(null); // Clear formatting if not "Yes"
    }
  });
}
