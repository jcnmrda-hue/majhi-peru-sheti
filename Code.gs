// माझी पेरू शेती — Google Apps Script backend
const SHEET_ID = 'PUT_YOUR_GOOGLE_SHEET_ID_HERE';

function doGet() {
  return ContentService.createTextOutput(JSON.stringify({
    ok: true,
    service: 'Guava Farm Sync v4'
  })).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      throw new Error('No POST data received');
    }
    const body = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.openById(SHEET_ID);
    ensureHeader(ss, 'SyncLog', ['Date','Farm ID','Type','Payload']);
    ss.getSheetByName('SyncLog').appendRow([
      new Date(), body.farmId || '', body.type || '', JSON.stringify(body.payload || {})
    ]);
    if (body.type === 'farm_bundle') {
      writeObjects(ss, 'Labour', body.payload?.records || []);
      writeObjects(ss, 'Workers', body.payload?.workers || []);
      writeObjects(ss, 'Tasks', body.payload?.tasks || []);
      writeObjects(ss, 'PlantHealth', body.payload?.weak || []);
    }
    return json({ok:true, message:'Data saved successfully'});
  } catch (err) {
    return json({ok:false, error:String(err)});
  }
}

function ensureHeader(ss, name, headers) {
  let sh = ss.getSheetByName(name) || ss.insertSheet(name);
  if (sh.getLastRow() === 0) sh.appendRow(headers);
}

function writeObjects(ss, name, arr) {
  if (!Array.isArray(arr) || arr.length === 0) return;
  const headers = Object.keys(arr[0]);
  ensureHeader(ss, name, headers);
  const sh = ss.getSheetByName(name);
  arr.forEach(obj => sh.appendRow(headers.map(h => obj[h] ?? '')));
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
