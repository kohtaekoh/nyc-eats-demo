// Eats QA Sync — Google Apps Script
// Deploy as: Web App → Execute as Me → Anyone can access (no auth)
//
// GET  ?qa=KEY        → { removed: ["photo__0__1", ...] }
// POST { action:"saveQA", key:"KEY", removed:[...] }  → { saved: true }

const QA_SHEET = 'EatsQA';

function doGet(e) {
  try {
    const key = e.parameter && e.parameter.qa;
    if (!key) return ok({ error: 'missing ?qa= param' });
    return ok({ removed: loadQA(key) });
  } catch(err) { return ok({ removed: [], error: err.message }); }
}

function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    if (payload.action === 'saveQA') {
      saveQA(payload.key, payload.removed || []);
      return ok({ saved: true });
    }
    return ok({ error: 'unknown action' });
  } catch(err) { return ok({ error: err.message }); }
}

function loadQA(key) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(QA_SHEET);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === key) {
      try { return JSON.parse(data[i][2]); } catch(e) { return []; }
    }
  }
  return [];
}

function saveQA(key, removed) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(QA_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(QA_SHEET);
    sheet.getRange(1,1,1,3).setValues([['key','updated_at','removed_json']]);
    sheet.getRange(1,1,1,3).setFontWeight('bold').setBackground('#e0f2fe');
    sheet.setFrozenRows(1);
  }
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === key) {
      sheet.getRange(i+1,1,1,3).setValues([[key, new Date().toISOString(), JSON.stringify(removed)]]);
      return;
    }
  }
  sheet.appendRow([key, new Date().toISOString(), JSON.stringify(removed)]);
}

function ok(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
