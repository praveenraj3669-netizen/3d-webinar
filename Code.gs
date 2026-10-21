/**
 * Code.gs — Google Apps Script backend for the LinkedIn audit enquiry form.
 *
 * SETUP:
 * 1. Open the Google Sheet you want responses saved to.
 * 2. Extensions > Apps Script. Paste this file's contents into Code.gs and save.
 * 3. Deploy > New deployment > Type: Web app.
 *      Execute as: Me
 *      Who has access: Anyone
 * 4. Click Deploy, authorize, and copy the Web app URL (ends in /exec).
 *
 * The script writes to the sheet it is bound to (the one you opened Apps Script from).
 */

var SHEET_NAME = 'Responses';
var HEADERS = ['Timestamp', 'Name', 'Email', 'Phone'];

/** Handles GET requests (used by the form via fetch with query params). */
function doGet(e) {
  return handleRequest(e);
}

/** Handles POST requests (same logic, reads from parameters). */
function doPost(e) {
  return handleRequest(e);
}

function handleRequest(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000); // avoid overlapping writes when several people submit at once

    var params = (e && e.parameter) ? e.parameter : {};
    var name = clean(params.name);
    var email = clean(params.email);
    var phone = clean(params.phone);

    var sheet = getOrCreateSheet();
    sheet.appendRow([new Date(), name, email, phone]);

    return jsonResponse({ success: true });
  } catch (err) {
    return jsonResponse({ success: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (ignore) {}
  }
}

/** Returns the "Responses" sheet, creating it with bold headers if missing. */
function getOrCreateSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }

  return sheet;
}

/** Trims input and returns an empty string for missing values. */
function clean(value) {
  return value === undefined || value === null ? '' : String(value).trim();
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
