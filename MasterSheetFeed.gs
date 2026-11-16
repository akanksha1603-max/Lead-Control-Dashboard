/**
 * Lead Control Dashboard — Apps Script for the November Master Integrated sheet.
 *  • GET  : lets the dashboard read the master sheet live (read-only).
 *  • POST : lets the dashboard email the New & Not Connected report to each
 *           Sales Head from YOUR Gmail, with the Excel attached and people in CC.
 *
 * Setup (once): Extensions → Apps Script → paste this over Code.gs.
 * 1. Change TOKEN below to your own secret (any long random text).
 * 2. Deploy → New deployment → "Web app"
 *      Execute as: Me        Who has access: Anyone
 *    (Already deployed? Deploy → Manage deployments → ✏️ Edit → Version: New version → Deploy,
 *     so the link stays the same.)
 * 3. Allow the permissions Google asks for (Sheets + send email as you).
 * 4. Copy the Web App URL and add  ?token=YOUR_TOKEN  at the end. Paste that full
 *    link into the dashboard (master sheet "connect live" and the report email panel).
 *
 * The sheet stays private: only someone with the full link incl. the token can use it.
 */
const TOKEN = 'CHANGE-ME-to-a-long-secret-text';
const SENDER_NAME = 'Lead Control Dashboard';

function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.token !== TOKEN) return json_({ error: 'Invalid or missing token' });
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

function doPost(e) {
  const p = (e && e.parameter) || {};
  if (p.token !== TOKEN) return json_({ error: 'Invalid or missing token' });
  let d;
  try { d = JSON.parse(e.postData.contents); } catch (err) { return json_({ error: 'Bad request' }); }
  if (d.action !== 'sendReport') return json_({ error: 'Unknown action' });
  if (!d.to) return json_({ error: 'No recipient' });
  try {
    const opts = { htmlBody: d.htmlBody || '', name: SENDER_NAME };
    if (d.cc) opts.cc = d.cc;
    if (d.attachment && d.attachment.base64) {
      opts.attachments = [Utilities.newBlob(Utilities.base64Decode(d.attachment.base64),
                                            d.attachment.mimeType, d.attachment.name)];
    }
    GmailApp.sendEmail(d.to, d.subject || 'Lead report', d.textBody || '', opts);
    return json_({ ok: true, remaining: MailApp.getRemainingDailyQuota() });
  } catch (err) {
    return json_({ error: String(err) });
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
