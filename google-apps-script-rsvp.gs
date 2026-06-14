/*
  RSVP Web App — Aashritha & Tanmayee's Ranga Pravesha
  =====================================================
  HOW TO SET UP:
  1. Go to https://script.google.com → New project
  2. Paste this entire file, replacing the default code
  3. Create a Google Sheet for responses (or attach this script to one)
  4. If using a standalone script, paste the Sheet ID in SPREADSHEET_ID below
  5. Click "Run" → "installReminderTrigger" once and approve permissions
  6. Click "Deploy" → "New deployment" → Type: Web App
     - Execute as: Me
     - Who has access: Anyone
  7. Copy the deployment URL and paste it into index.html as RSVP_ENDPOINT
*/

const SPREADSHEET_ID = "1GK5mgQ7EGU-QMqbUyWelwpCPUCkhIEtY_DKwswnjUyc";

const CONFIG = {
  sheetName:         "Responses",
  eventTitle:        "Aashritha & Tanmayee's Bharatanatyam Ranga Pravesha",
  eventDateLabel:    "Saturday, July 18, 2026",
  eventTimeLabel:    "From 3:00 PM Onwards",
  doorsLabel:        "Doors open at 2:30 PM",
  eventStartUtc:     "20260718T190000Z",
  eventEndUtc:       "20260718T220000Z",
  eventLocation:     "Shenkman Arts Centre, 245 Centrum Blvd., Ottawa, ON K1E 0A1",
  siteUrl:           "https://aashrithatanmayee.com/",
  googleCalendarUrl: "https://calendar.google.com/calendar/render?action=TEMPLATE&text=Aashritha+%26+Tanmayee%27s+Ranga+Pravesha&dates=20260718T190000Z%2F20260718T220000Z&ctz=America%2FToronto&details=Bharatanatyam+Ranga+Pravesha+of+Aashritha+Kanumuri+%26+Tanmayee+Sista.+Doors+open+at+2%3A30+PM.&location=Shenkman+Arts+Centre%2C+245+Centrum+Blvd.%2C+Ottawa%2C+ON+K1E+0A1",
  reminderAt:        "2026-07-16T10:00:00-04:00",
  familyNames:       "The Kanumuri, Kilari, Sista & Ayyagari Families"
};

const HEADERS = [
  "Timestamp",
  "Name",
  "Email",
  "Phone",
  "Status",
  "Adults",
  "Children",
  "Message",
  "Submitted At",
  "Confirmation Sent At",
  "Reminder Sent At"
];

const HEADER_ALIASES = {
  "Name":         ["Full Name"],
  "Email":        ["Email Address", "E-mail"],
  "Phone":        ["Phone Number", "Mobile"],
  "Status":       ["Attendance", "RSVP"],
  "Adults":       ["Adult Count", "Number of Adults"],
  "Children":     ["Child Count", "Kids", "Number of Children"],
  "Message":      ["Comment", "Comments", "Note", "Notes", "Wish", "Wishes"],
  "Submitted At": ["Date", "Submitted", "Submission Date"]
};

// ── Request handlers ─────────────────────────────────────────────────────────

function doGet(e) {
  const params = e && e.parameter ? e.parameter : {};
  if (params.action === "comments") return publicCommentsResponse_(params);
  return params.callback
    ? jsonp_(params.callback, healthPayload_())
    : json_(healthPayload_());
}

function doPost(e) {
  const data = parseRequest_(e);
  data.status = normaliseStatus_(data.status);

  if (!data.name || !data.email || !data.status) {
    return json_({ ok: false, error: "Name, email, and status are required." });
  }

  const sheet     = getSheet_();
  const headerMap = ensureHeaders_(sheet);
  const rowNumber = upsertRsvp_(sheet, headerMap, data);

  try {
    sendConfirmationEmail_(data);
    sheet.getRange(rowNumber, headerMap["Confirmation Sent At"]).setValue(new Date());
  } catch (err) {
    console.error("Confirmation email failed: " + err.message);
  }

  return json_({ ok: true });
}

// ── Triggers ─────────────────────────────────────────────────────────────────

function installReminderTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === "sendReminderEmails")
    .forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger("sendReminderEmails")
    .timeBased()
    .at(new Date(CONFIG.reminderAt))
    .create();
}

function sendReminderEmails() {
  const sheet     = getSheet_();
  const headerMap = ensureHeaders_(sheet);
  const lastRow   = sheet.getLastRow();
  if (lastRow < 2) return;

  const values = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  values.forEach((row, index) => {
    const rowNumber      = index + 2;
    const status         = normaliseStatus_(row[headerMap.Status - 1]);
    const email          = String(row[headerMap.Email - 1] || "").trim();
    const reminderSentAt = row[headerMap["Reminder Sent At"] - 1];

    if (status !== "yes" || !email || reminderSentAt) return;

    const data = {
      name:     row[headerMap.Name - 1],
      email,
      status,
      adults:   row[headerMap.Adults - 1],
      children: row[headerMap.Children - 1]
    };

    try {
      sendReminderEmail_(data);
      sheet.getRange(rowNumber, headerMap["Reminder Sent At"]).setValue(new Date());
    } catch (err) {
      console.error("Reminder failed for " + email + ": " + err.message);
    }
  });
}

// ── De-duplication utility ────────────────────────────────────────────────────

function dedupeExistingRsvps() {
  const sheet      = getSheet_();
  const headerMap  = ensureHeaders_(sheet);
  const lastRow    = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();
  if (!headerMap.Email || lastRow < 3) return "No duplicate RSVP rows found.";

  const values        = sheet.getRange(2, 1, lastRow - 1, lastColumn).getValues();
  const latestByEmail = {};

  values.forEach((row, index) => {
    const email = String(row[headerMap.Email - 1] || "").trim().toLowerCase();
    if (!email) return;
    const rowNumber = index + 2;
    const previous  = latestByEmail[email];
    if (!previous || rowDateValue_(row, headerMap) >= rowDateValue_(previous.row, headerMap)) {
      latestByEmail[email] = { rowNumber, row };
    }
  });

  const rowsToKeep   = Object.keys(latestByEmail).reduce((memo, email) => { memo[latestByEmail[email].rowNumber] = true; return memo; }, {});
  const rowsToDelete = [];
  values.forEach((row, index) => {
    const rowNumber = index + 2;
    const email     = String(row[headerMap.Email - 1] || "").trim().toLowerCase();
    if (email && !rowsToKeep[rowNumber]) rowsToDelete.push(rowNumber);
  });

  rowsToDelete.sort((a, b) => b - a).forEach(r => sheet.deleteRow(r));
  return "Removed " + rowsToDelete.length + " duplicate RSVP row(s).";
}

// ── Public comments ───────────────────────────────────────────────────────────

function publicCommentsResponse_(params) {
  const requestedLimit = Number(params.limit || 24);
  const limit          = Number.isFinite(requestedLimit) ? Math.max(1, Math.min(requestedLimit, 60)) : 24;
  const payload        = { ok: true, messages: getPublicMessages_(limit) };
  return params.callback ? jsonp_(params.callback, payload) : json_(payload);
}

function getPublicMessages_(limit) {
  const sheet     = getSheet_();
  const headerMap = ensureHeaders_(sheet);
  const lastRow   = sheet.getLastRow();
  if (!headerMap.Message || lastRow < 2) return [];

  return sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn())
    .getValues()
    .map(row => {
      const message = cleanPublicText_(row[headerMap.Message - 1], 700);
      if (!message) return null;
      return {
        name:        publicDisplayName_(row[headerMap.Name - 1]),
        message,
        submittedAt: publicDate_(row, headerMap)
      };
    })
    .filter(Boolean)
    .sort((a, b) => (Date.parse(b.submittedAt) || 0) - (Date.parse(a.submittedAt) || 0))
    .slice(0, limit);
}

// ── Sheet helpers ─────────────────────────────────────────────────────────────

function getSheet_() {
  const scriptId      = PropertiesService.getScriptProperties().getProperty("SPREADSHEET_ID");
  const spreadsheetId = scriptId || SPREADSHEET_ID;
  const spreadsheet   = spreadsheetId
    ? SpreadsheetApp.openById(spreadsheetId)
    : SpreadsheetApp.getActiveSpreadsheet();

  if (!spreadsheet) throw new Error("No spreadsheet found. Set SPREADSHEET_ID or attach this script to the sheet.");

  return spreadsheet.getSheetByName(CONFIG.sheetName)
    || spreadsheet.getSheets()[0]
    || spreadsheet.insertSheet(CONFIG.sheetName);
}

function ensureHeaders_(sheet) {
  const lastColumn = Math.max(sheet.getLastColumn(), HEADERS.length);
  const existing   = sheet.getRange(1, 1, 1, lastColumn).getValues()[0].map(v => String(v || "").trim());
  const hasHeaders = existing.some(Boolean);

  if (!hasHeaders) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sheet.setFrozenRows(1);
    return getHeaderMap_(sheet);
  }

  const canonicalExisting = existing.map(h => canonicalHeader_(h)).filter(Boolean);
  const missing           = HEADERS.filter(h => !canonicalExisting.includes(h));
  if (missing.length) {
    const lastUsed = existing.reduce((last, v, i) => v ? i + 1 : last, 0);
    sheet.getRange(1, lastUsed + 1, 1, missing.length).setValues([missing]);
  }

  sheet.setFrozenRows(1);
  return getHeaderMap_(sheet);
}

function getHeaderMap_(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  return headers.reduce((map, header, index) => {
    const label     = String(header || "").trim();
    const canonical = canonicalHeader_(label);
    if (canonical && !map[canonical]) map[canonical] = index + 1;
    if (label && !map[label]) map[label] = index + 1;
    return map;
  }, {});
}

function canonicalHeader_(header) {
  const label  = String(header || "").trim();
  if (HEADERS.includes(label)) return label;
  const lower  = label.toLowerCase();
  const direct = HEADERS.find(c => c.toLowerCase() === lower);
  if (direct) return direct;
  return HEADERS.find(c => (HEADER_ALIASES[c] || []).some(a => a.toLowerCase() === lower)) || "";
}

function upsertRsvp_(sheet, headerMap, data) {
  const email        = String(data.email || "").trim().toLowerCase();
  const matchingRows = findRowsByEmail_(sheet, headerMap, email);
  const existing     = matchingRows.length
    ? sheet.getRange(matchingRows[0], 1, 1, sheet.getLastColumn()).getValues()[0]
    : new Array(sheet.getLastColumn()).fill("");
  const now = new Date();

  matchingRows.sort((a, b) => b - a).forEach(r => sheet.deleteRow(r));
  const rowNumber = sheet.getLastRow() + 1;

  setCell_(existing, headerMap, "Timestamp",              existing[headerMap.Timestamp - 1] || now);
  setCell_(existing, headerMap, "Name",                   data.name);
  setCell_(existing, headerMap, "Email",                  data.email);
  setCell_(existing, headerMap, "Phone",                  data.phone || "");
  setCell_(existing, headerMap, "Status",                 data.status);
  setCell_(existing, headerMap, "Adults",                 data.status === "yes" ? data.adults   || "0" : "0");
  setCell_(existing, headerMap, "Children",               data.status === "yes" ? data.children || "0" : "0");
  setCell_(existing, headerMap, "Message",                data.message || "");
  setCell_(existing, headerMap, "Submitted At",           data.date || now);
  setCell_(existing, headerMap, "Confirmation Sent At",   "");
  if (data.status !== "yes") setCell_(existing, headerMap, "Reminder Sent At", "");

  sheet.getRange(rowNumber, 1, 1, existing.length).setValues([existing]);
  return rowNumber;
}

function findRowsByEmail_(sheet, headerMap, email) {
  const emailCol = headerMap.Email;
  const lastRow  = sheet.getLastRow();
  if (!emailCol || lastRow < 2) return [];
  const emails = sheet.getRange(2, emailCol, lastRow - 1, 1).getValues();
  return emails.reduce((rows, row, i) => {
    if (String(row[0] || "").trim().toLowerCase() === email) rows.push(i + 2);
    return rows;
  }, []);
}

function setCell_(row, headerMap, header, value) {
  row[headerMap[header] - 1] = value;
}

function rowDateValue_(row, headerMap) {
  const value = row[((headerMap["Submitted At"] || headerMap.Timestamp) - 1)];
  const date  = value instanceof Date ? value : new Date(value);
  const time  = date.getTime();
  return isNaN(time) ? 0 : time;
}

// ── Emails ────────────────────────────────────────────────────────────────────

function sendConfirmationEmail_(data) {
  const attending = data.status === "yes";
  const subject   = attending
    ? "RSVP confirmed — " + CONFIG.eventTitle
    : "RSVP received — " + CONFIG.eventTitle;
  const intro     = attending
    ? "Thank you for your RSVP. We are delighted that you can join us for this special celebration."
    : "Thank you for letting us know that you cannot make it. We will miss celebrating with you.";
  const guestLine = attending ? "<p><strong>Guests:</strong> " + escapeHtml_(formatGuestCount_(data)) + "</p>" : "";
  const calBlock  = attending ? calendarButtonsHtml_() : "";
  const htmlBody  = emailShell_(data.name, intro + guestLine + eventDetailsHtml_() + calBlock);
  const plainBody = attending
    ? "Thank you for your RSVP.\n\n" + eventDetailsText_() + "\n\nAdd to Google Calendar:\n" + CONFIG.googleCalendarUrl
    : "Thank you for letting us know.\n\n" + eventDetailsText_();

  sendEmail_(data.email, subject, plainBody, htmlBody, attending);
}

function sendReminderEmail_(data) {
  const subject   = "Reminder — " + CONFIG.eventTitle + " is this week!";
  const intro     = "We are looking forward to seeing you! Here are the event details again for your calendar.";
  const guestLine = "<p><strong>Guests:</strong> " + escapeHtml_(formatGuestCount_(data)) + "</p>";
  const htmlBody  = emailShell_(data.name || "Guest", intro + guestLine + eventDetailsHtml_() + calendarButtonsHtml_());
  const plainBody = "We look forward to seeing you!\n\n" + eventDetailsText_() + "\n\nAdd to Google Calendar:\n" + CONFIG.googleCalendarUrl;
  sendEmail_(data.email, subject, plainBody, htmlBody, true);
}

function sendEmail_(to, subject, body, htmlBody, attachInvite) {
  const options = { name: "Aashritha & Tanmayee RSVP", htmlBody };
  if (attachInvite) {
    options.attachments = [Utilities.newBlob(buildIcs_(), "text/calendar", "aashritha-tanmayee-ranga-pravesha.ics")];
  }
  MailApp.sendEmail(to, subject, body, options);
}

function eventDetailsHtml_() {
  return [
    "<div style='margin:22px 0;padding:18px;border:1px solid #ead8aa;border-radius:10px;background:#fffaf0;'>",
    "<p style='margin:0 0 8px;'><strong>Date:</strong> "  + escapeHtml_(CONFIG.eventDateLabel)  + "</p>",
    "<p style='margin:0 0 8px;'><strong>Time:</strong> "  + escapeHtml_(CONFIG.eventTimeLabel)  + "</p>",
    "<p style='margin:0 0 8px;'><strong>Doors:</strong> " + escapeHtml_(CONFIG.doorsLabel)      + "</p>",
    "<p style='margin:0;'><strong>Venue:</strong> "       + escapeHtml_(CONFIG.eventLocation)   + "</p>",
    "</div>"
  ].join("");
}

function eventDetailsText_() {
  return [
    CONFIG.eventTitle,
    "Date:  " + CONFIG.eventDateLabel,
    "Time:  " + CONFIG.eventTimeLabel,
    "Doors: " + CONFIG.doorsLabel,
    "Venue: " + CONFIG.eventLocation,
    "Site:  " + CONFIG.siteUrl
  ].join("\n");
}

function calendarButtonsHtml_() {
  return [
    "<p style='margin:22px 0 10px;'>Add the event to your calendar:</p><p>",
    "<a href='" + CONFIG.googleCalendarUrl + "' style='display:inline-block;margin:0 8px 10px 0;padding:11px 18px;border-radius:999px;background:#c9a84c;color:#2a1309;text-decoration:none;font-weight:bold;'>Add to Google Calendar</a>",
    "</p>"
  ].join("");
}

function emailShell_(name, bodyHtml) {
  return [
    "<div style='font-family:Arial,sans-serif;line-height:1.6;color:#2c1408;max-width:620px;'>",
    "<h2 style='font-family:Georgia,serif;color:#3d0c1e;margin-bottom:4px;'>Aashritha &amp; Tanmayee's Ranga Pravesha</h2>",
    "<p style='color:#9a7330;margin-top:0;'>July 18, 2026 &nbsp;&middot;&nbsp; Shenkman Arts Centre, Ottawa</p>",
    "<p>Dear " + escapeHtml_(name) + ",</p>",
    "<div>" + bodyHtml + "</div>",
    "<p style='margin-top:24px;'>With gratitude and joy,<br>" + escapeHtml_(CONFIG.familyNames) + "</p>",
    "</div>"
  ].join("");
}

function formatGuestCount_(data) {
  const adults   = Number(data.adults   || 0);
  const children = Number(data.children || 0);
  const parts    = [];
  if (adults)   parts.push(adults   + " adult"  + (adults   === 1 ? "" : "s"));
  if (children) parts.push(children + " child"  + (children === 1 ? "" : "ren"));
  return parts.length ? parts.join(", ") : "0 guests";
}

// ── iCalendar ─────────────────────────────────────────────────────────────────

function buildIcs_() {
  const lines = [
    "BEGIN:VCALENDAR",
    "VERSION:2.0",
    "PRODID:-//Aashritha and Tanmayee Ranga Pravesha//EN",
    "CALSCALE:GREGORIAN",
    "METHOD:PUBLISH",
    "BEGIN:VEVENT",
    "UID:aashritha-tanmayee-ranga-pravesha-20260718@aashrithatanmayee.com",
    "DTSTAMP:20260101T000000Z",
    "DTSTART:" + CONFIG.eventStartUtc,
    "DTEND:"   + CONFIG.eventEndUtc,
    "SUMMARY:"     + escapeIcs_(CONFIG.eventTitle),
    "LOCATION:"    + escapeIcs_(CONFIG.eventLocation),
    "DESCRIPTION:" + escapeIcs_(CONFIG.doorsLabel + ". You are invited to Aashritha Kanumuri and Tanmayee Sista's Bharatanatyam Ranga Pravesha."),
    "URL:" + CONFIG.siteUrl,
    "END:VEVENT",
    "END:VCALENDAR"
  ];
  return lines.map(foldIcsLine_).join("\r\n") + "\r\n";
}

function foldIcsLine_(line) {
  const chunks = [];
  let remaining = line;
  while (remaining.length > 73) {
    chunks.push(remaining.slice(0, 73));
    remaining = " " + remaining.slice(73);
  }
  chunks.push(remaining);
  return chunks.join("\r\n");
}

// ── Utilities ─────────────────────────────────────────────────────────────────

function normaliseStatus_(status) {
  const v = String(status || "").trim().toLowerCase();
  if (["yes", "y", "attending", "attending!"].includes(v)) return "yes";
  if (["no", "n", "declined", "can't make it", "cant make it"].includes(v)) return "no";
  return v;
}

function publicDisplayName_(name) { return cleanPublicText_(name, 80) || "Guest"; }

function cleanPublicText_(value, maxLength) {
  const text = String(value || "").replace(/\s+/g, " ").trim();
  if (!text) return "";
  return text.length > maxLength ? text.slice(0, maxLength - 1).trim() + "..." : text;
}

function publicDate_(row, headerMap) {
  const value = row[((headerMap["Submitted At"] || headerMap.Timestamp) - 1)];
  const date  = value instanceof Date ? value : new Date(value);
  return isNaN(date.getTime()) ? "" : date.toISOString();
}

function healthPayload_() {
  return { ok: true, message: "Aashritha & Tanmayee RSVP endpoint is running." };
}

function parseRequest_(e) {
  const raw = e && e.postData && e.postData.contents ? e.postData.contents : "{}";
  try {
    return JSON.parse(raw);
  } catch (_) {
    return raw.split("&").reduce((memo, pair) => {
      const parts = pair.split("=");
      if (!parts[0]) return memo;
      memo[decodeURIComponent(parts[0])] = decodeURIComponent((parts[1] || "").replace(/\+/g, " "));
      return memo;
    }, {});
  }
}

function escapeHtml_(value) {
  return String(value || "")
    .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;").replace(/'/g, "&#039;");
}

function escapeIcs_(value) {
  return String(value || "")
    .replace(/\\/g, "\\\\").replace(/;/g, "\\;")
    .replace(/,/g, "\\,").replace(/\n/g, "\\n");
}

function json_(payload) {
  return ContentService.createTextOutput(JSON.stringify(payload)).setMimeType(ContentService.MimeType.JSON);
}

function jsonp_(callback, payload) {
  if (!/^[A-Za-z_$][0-9A-Za-z_$]*(\.[A-Za-z_$][0-9A-Za-z_$]*)*$/.test(String(callback || ""))) {
    return json_({ ok: false, error: "Invalid callback." });
  }
  return ContentService
    .createTextOutput(callback + "(" + JSON.stringify(payload) + ");")
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}
