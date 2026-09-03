/**
 * माझी पेरू शेती — Google Apps Script Cloud Sync
 *
 * 1. Google Sheet उघडा → Extensions → Apps Script
 * 2. हा पूर्ण code Code.gs मध्ये paste करा.
 * 3. Deploy → New deployment → Web app
 *    Execute as: Me
 *    Who has access: Anyone
 * 4. Web App URL website मध्ये टाका.
 *
 * Sheet मध्ये एक sheet "FarmData" आपोआप तयार होईल.
 */

const SHEET_NAME = 'FarmData';

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.getRange(1,1,1,4).setValues([['FarmID','UpdatedAt','JSON','UpdatedBy']]);
    sh.setFrozenRows(1);
  }
  return sh;
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  return json_({ok:true, service:'Majhi Peru Sheti Cloud Sync', message:'POST API ready'});
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents || '{}');
    const action = body.action;
    const farmId = String(body.farmId || 'MY-GUAVA-FARM').trim();
    if (!farmId) throw new Error('Farm ID required');

    const sh = getSheet_();
    const last = sh.getLastRow();

    if (action === 'load') {
      if (last < 2) return json_({ok:true, data:null, updatedAt:0});
      const values = sh.getRange(2,1,last-1,4).getValues();
      for (let i=0; i<values.length; i++) {
        if (String(values[i][0]) === farmId) {
          let data = null;
          try { data = JSON.parse(values[i][2]); } catch (_) {}
          return json_({
            ok:true,
            data:data,
            updatedAt:Number(values[i][1]) || 0
          });
        }
      }
      return json_({ok:true, data:null, updatedAt:0});
    }

    if (action === 'save') {
      const incomingTime = Number(body.updatedAt || Date.now());
      const incomingData = body.data;
      if (!incomingData) throw new Error('Data missing');

      let row = -1;
      let existingTime = 0;

      if (last >= 2) {
        const ids = sh.getRange(2,1,last-1,1).getValues();
        const times = sh.getRange(2,2,last-1,1).getValues();
        for (let i=0; i<ids.length; i++) {
          if (String(ids[i][0]) === farmId) {
            row = i + 2;
            existingTime = Number(times[i][0]) || 0;
            break;
          }
        }
      }

      // Last-write-wins protection: an older device cannot overwrite newer cloud data.
      if (existingTime > incomingTime) {
        const oldJson = sh.getRange(row,3).getValue();
        let oldData = null;
        try { oldData = JSON.parse(oldJson); } catch (_) {}
        return json_({
          ok:true,
          data:oldData,
          updatedAt:existingTime,
          conflict:true,
          message:'Cloud data was newer; newer cloud copy returned.'
        });
      }

      const now = Math.max(incomingTime, Date.now());
      const rowValues = [[farmId, now, JSON.stringify(incomingData), Session.getActiveUser().getEmail() || 'web-app']];
      if (row === -1) {
        sh.getRange(last+1,1,1,4).setValues(rowValues);
      } else {
        sh.getRange(row,1,1,4).setValues(rowValues);
      }

      return json_({ok:true, data:incomingData, updatedAt:now, conflict:false});
    }

    throw new Error('Unknown action');
  } catch (err) {
    return json_({ok:false, error:String(err.message || err)});
  }
}
