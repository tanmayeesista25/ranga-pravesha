// ═══════════════════════════════════════════════════════════════════
//  Aashritha & Tanmayee Ranga Pravesha — RSVP Google Apps Script
//  Deploy as: Extensions → Apps Script → Deploy → New Deployment
//  Type: Web App | Execute as: Me | Who has access: Anyone
//  Then paste the Web App URL into index.html → const ENDPOINT = "..."
// ═══════════════════════════════════════════════════════════════════

// 👉 Replace with your Google Sheet ID (from the URL of your sheet)
const SHEET_ID = "PASTE_YOUR_GOOGLE_SHEET_ID_HERE";
const SHEET_NAME = "RSVPs";

function doPost(e) {
  const headers = {
    "Access-Control-Allow-Origin": "*",
    "Access-Control-Allow-Methods": "POST, GET, OPTIONS",
    "Access-Control-Allow-Headers": "Content-Type",
    "Content-Type": "application/json"
  };

  try {
    const data = JSON.parse(e.postData.contents);
    const sheet = getOrCreateSheet();

    // Check if this email already submitted (update instead of duplicate)
    const existing = findRowByEmail(sheet, data.email);
    const row = [
      data.submittedAt || new Date().toISOString(),
      data.name   || "",
      data.email  || "",
      data.phone  || "",
      data.status === "yes" ? "Attending" : "Not Attending",
      data.adults   || "0",
      data.children || "0",
      data.message  || ""
    ];

    if (existing > 0) {
      sheet.getRange(existing, 1, 1, row.length).setValues([row]);
    } else {
      sheet.appendRow(row);
    }

    return ContentService
      .createTextOutput(JSON.stringify({ success: true }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  const action = e.parameter.action;
  const callback = e.parameter.callback;

  if (action === "comments") {
    const sheet = getOrCreateSheet();
    const rows = sheet.getDataRange().getValues();
    const messages = [];

    for (let i = 1; i < rows.length; i++) {
      const msg = rows[i][7]; // message column
      if (msg && msg.toString().trim()) {
        messages.push({
          submittedAt: rows[i][0],
          name: rows[i][1],
          status: rows[i][4],
          message: msg.toString().trim()
        });
      }
    }

    // Return newest first, limit to requested count
    const limit = parseInt(e.parameter.limit) || 24;
    const result = messages.reverse().slice(0, limit);
    const json = JSON.stringify({ success: true, messages: result });

    // JSONP support for cross-origin loading
    if (callback) {
      return ContentService
        .createTextOutput(`${callback}(${json})`)
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    return ContentService
      .createTextOutput(json)
      .setMimeType(ContentService.MimeType.JSON);
  }

  return ContentService
    .createTextOutput(JSON.stringify({ success: false, error: "Unknown action" }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── Helpers ──

function getOrCreateSheet() {
  const ss = SHEET_ID === "PASTE_YOUR_GOOGLE_SHEET_ID_HERE"
    ? SpreadsheetApp.getActiveSpreadsheet()
    : SpreadsheetApp.openById(SHEET_ID);

  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(["Timestamp", "Name", "Email", "Phone", "Status", "Adults", "Children", "Message"]);
    sheet.setFrozenRows(1);
    // Basic formatting
    sheet.getRange(1, 1, 1, 8).setFontWeight("bold").setBackground("#3d0c1e").setFontColor("#e8c96e");
    sheet.setColumnWidth(1, 160);
    sheet.setColumnWidth(2, 160);
    sheet.setColumnWidth(3, 200);
    sheet.setColumnWidth(8, 300);
  }
  return sheet;
}

function findRowByEmail(sheet, email) {
  if (!email) return -1;
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][2] && data[i][2].toString().toLowerCase() === email.toLowerCase()) {
      return i + 1; // 1-indexed row number
    }
  }
  return -1;
}
