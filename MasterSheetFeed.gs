/**
 * Lead Control Dashboard — live feed for the November Master Integrated sheet.
 *
 * Paste this into the master sheet: Extensions → Apps Script → replace Code.gs.
 * 1. Change TOKEN below to your own secret (any long random text).
 * 2. Deploy → New deployment → type "Web app"
 *      Execute as: Me
 *      Who has access: Anyone
 * 3. Copy the Web App URL and add  ?token=YOUR_TOKEN  at the end.
 *    Paste that full link into the dashboard (CRM Mapping Mismatch →
 *    "Connect your Google Sheet live").
 *
 * The sheet itself stays private. Only someone who has the full link
 * including the token can read it, and only read-only.
 */
const TOKEN = 'CHANGE-ME-to-a-long-secret-text';

function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.token !== TOKEN) return json_({ error: 'Invalid or missing token' });

  // "range" = a tab name ("Master Sheet") or a tab + range ("Master Sheet!A3:R500")
  const range = p.range || 'Master Sheet';
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let values;
  try {
    if (range.indexOf('!') > -1) {
      values = ss.getRange(range).getDisplayValues();
    } else {
      const sheet = ss.getSheetByName(range);
      if (!sheet) return json_({ error: 'No tab named "' + range + '"' });
      values = sheet.getDataRange().getDisplayValues();
    }
  } catch (err) {
    return json_({ error: String(err) });
  }
  return json_({ values: values, updated: new Date().toISOString() });
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
