/**
 * Merchant Onboarding Document Generator — Sheet Logger
 * ------------------------------------------------------
 * STATUS: SUSPENDED. The app no longer posts to this script — `SHEETS_URL`
 * in app.js is blank, and usage is recorded locally in the browser instead
 * (exportable as CSV from Step 1). This file is retained so logging can be
 * switched back on later: redeploy the web app and paste the /exec URL into
 * SHEETS_URL. Nothing else in app.js needs to change.
 * ------------------------------------------------------
 * This Apps Script receives a log-only submission every time an agent
 * generates a document. It writes exactly ONE row per event with:
 *   Timestamp | Agent Mobile Number | Merchant (Display Name or Merchant ID) | New/Existing
 *
 * IMPORTANT — Workspace domain restriction:
 * Many corporate Google Workspace accounts do not allow deploying a Web App
 * with "Who has access: Anyone" — only "Anyone within [your domain]" is
 * available. Because of that, app.js submits this log via a HIDDEN FORM POST
 * (a real browser navigation inside a hidden iframe), not a fetch() call.
 * A form POST carries the browser's existing Google session cookie, so it
 * succeeds automatically on any device/browser already signed into an
 * @phonepe.com Google account (typical on company-managed phones) — no
 * visible login prompt, no interruption to the agent. If the browser has no
 * active @phonepe.com Google session, the hidden iframe just shows a sign-in
 * page invisibly and the log entry is skipped — this never blocks or delays
 * document generation, which always works regardless of logging status.
 *
 * Because the request arrives as a standard form POST (not JSON), this
 * script reads fields from e.parameter. A JSON-body fallback is kept too,
 * in case you ever call this from a trusted server-side integration instead.
 *
 * SETUP:
 * 1. Open a Google Sheet in your corporate Google Workspace account.
 * 2. Extensions -> Apps Script.
 * 3. Delete any starter code and paste this file's contents.
 * 4. Rename the sheet tab (first tab) to "Log" or update SHEET_NAME below.
 * 5. Click Deploy -> New deployment -> Type: "Web app".
 *      - Execute as: Me
 *      - Who has access: Anyone within [your domain] (this is fine — see above)
 * 6. Copy the deployment URL (ends in /exec).
 * 7. Paste that URL into the SHEETS_URL constant near the top of app.js.
 */

const SHEET_NAME = "Log";

function doPost(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      sheet.appendRow(["Timestamp", "Agent Mobile Number", "Merchant (Display Name / ID)", "Merchant Status"]);
    }

    // Prefer standard form fields (what the hidden-form POST sends).
    // Fall back to a JSON body if this is ever called programmatically instead.
    let data = e.parameter || {};
    if ((!data.agentMobile && !data.merchant) && e.postData && e.postData.contents) {
      try { data = JSON.parse(e.postData.contents); } catch (parseErr) { /* ignore */ }
    }

    sheet.appendRow([
      data.timestamp || new Date().toLocaleString("en-IN", { timeZone: "Asia/Kolkata" }),
      data.agentMobile || "",
      data.merchant || "",
      data.merchantStatus || "",
    ]);

    return ContentService.createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Simple health check so you can verify the deployment URL works
// by opening it directly in a browser.
function doGet(e) {
  return ContentService.createTextOutput("Merchant Onboarding logger is running.");
}
