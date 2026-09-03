/**
 * माझी पेरू शेती — Google Apps Script Cloud Sync V5
 *
 * Google Sheet ID:
 * 1hFKAwIBdwZvGx-It8-qZjz1N4FheYNhjy4UFi5iOmOQ
 *
 * Sheet:
 * FarmData
 *
 * Deployment:
 * Execute as: Me
 * Who has access: Anyone
 */

const SHEET_ID = '1hFKAwIBdwZvGx-It8-qZjz1N4FheYNhjy4UFi5iOmOQ';
const SHEET_NAME = 'FarmData';

function getSheet_() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sh = ss.getSheetByName(SHEET_NAME);

  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.getRange(1, 1, 1, 4).setValues([
      ['FarmID', 'UpdatedAt', 'JSON', 'UpdatedBy']
    ]);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, 4).setFontWeight('bold');
  }

  return sh;
}

function jsonResponse_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet() {
  try {
    const sh = getSheet_();
    return jsonResponse_({
      ok: true,
      service: 'Majhi Peru Sheti Cloud Sync',
      version: 'V5',
      sheet: sh.getName(),
      message: 'Cloud Sync API is working'
    });
  } catch (err) {
    return jsonResponse_({
      ok: false,
      error: String(err.message || err)
    });
  }
}

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      throw new Error('POST data not received');
    }

    const body = JSON.parse(e.postData.contents || '{}');
    const action = String(body.action || '').trim();
    const farmId = String(body.farmId || 'MY-GUAVA-FARM').trim();

    if (!farmId) {
      throw new Error('Farm ID required');
    }

    const sh = getSheet_();
    const lastRow = sh.getLastRow();

    // LOAD
    if (action === 'load') {
      if (lastRow < 2) {
        return jsonResponse_({
          ok: true,
          data: null,
          updatedAt: 0,
          farmId: farmId
        });
      }

      const values = sh.getRange(2, 1, lastRow - 1, 4).getValues();

      for (let i = 0; i < values.length; i++) {
        if (String(values[i][0]).trim() === farmId) {
          let data = null;
          try {
            data = JSON.parse(values[i][2]);
          } catch (_) {
            data = null;
          }

          return jsonResponse_({
            ok: true,
            data: data,
            updatedAt: Number(values[i][1]) || 0,
            farmId: farmId,
            conflict: false
          });
        }
      }

      return jsonResponse_({
        ok: true,
        data: null,
        updatedAt: 0,
        farmId: farmId
      });
    }

    // SAVE
    if (action === 'save') {
      const incomingTime = Number(body.updatedAt || Date.now());
      const incomingData = body.data;

      if (!incomingData || typeof incomingData !== 'object') {
        throw new Error('Farm data missing');
      }

      let existingRow = -1;
      let existingTime = 0;

      if (lastRow >= 2) {
        const values = sh.getRange(2, 1, lastRow - 1, 2).getValues();

        for (let i = 0; i < values.length; i++) {
          if (String(values[i][0]).trim() === farmId) {
            existingRow = i + 2;
            existingTime = Number(values[i][1]) || 0;
            break;
          }
        }
      }

      // Prevent an older device from overwriting newer cloud data.
      if (existingRow !== -1 && existingTime > incomingTime) {
        let cloudData = null;

        try {
          cloudData = JSON.parse(
            sh.getRange(existingRow, 3).getValue()
          );
        } catch (_) {
          cloudData = null;
        }

        return jsonResponse_({
          ok: true,
          data: cloudData,
          updatedAt: existingTime,
          farmId: farmId,
          conflict: true,
          message: 'Cloud data is newer. Cloud copy returned.'
        });
      }

      const saveTime = Math.max(incomingTime, Date.now());

      let userEmail = 'web-app';
      try {
        userEmail = Session.getActiveUser().getEmail() || 'web-app';
      } catch (_) {
        userEmail = 'web-app';
      }

      const rowData = [[
        farmId,
        saveTime,
        JSON.stringify(incomingData),
        userEmail
      ]];

      if (existingRow === -1) {
        sh.getRange(lastRow + 1, 1, 1, 4).setValues(rowData);
      } else {
        sh.getRange(existingRow, 1, 1, 4).setValues(rowData);
      }

      return jsonResponse_({
        ok: true,
        data: incomingData,
        updatedAt: saveTime,
        farmId: farmId,
        conflict: false,
        message: 'Farm data saved successfully'
      });
    }

    throw new Error('Unknown action: ' + action);

  } catch (err) {
    return jsonResponse_({
      ok: false,
      error: String(err.message || err)
    });
  }
}
