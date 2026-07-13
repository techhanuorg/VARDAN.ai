/**
 * Paste this into Extensions → Apps Script inside your Google Sheet, then deploy it as a Web App.
 * Set access to "Anyone" and paste the deployment URL into Hospital Settings in VARDAN.ai.
 */
function doPost(e) {
  const body = JSON.parse(e.postData.contents || '{}');
  const sheetName = body.type === 'appointments' ? 'Appointments' : 'Patients';
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName) || SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
  const records = body.records || [];
  if (!records.length) return ContentService.createTextOutput(JSON.stringify({ ok: true }));
  const existingHeaders = sheet.getLastRow() ? sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0] : [];
  const headers = [...new Set([...existingHeaders, ...records.flatMap(Object.keys)])];
  if (headers.join('|') !== existingHeaders.join('|')) sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(sheet.getLastRow() + 1, 1, records.length, headers.length).setValues(records.map(record => headers.map(header => record[header] ?? '')));
  return ContentService.createTextOutput(JSON.stringify({ ok: true, saved: records.length }));
}
