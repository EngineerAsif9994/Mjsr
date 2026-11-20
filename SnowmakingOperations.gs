/**
 * MJSR Snowmaking Operations Logger
 * Deploy as a separate Google Apps Script Web App.
 * Uses a dedicated tab in the existing MJSR wet-bulb spreadsheet.
 */
const OPS_SPREADSHEET_ID = '1ajbIAyHrvYgqmD6MiwO-W9x7Is0NmhYNYgTJMjCM50U';
const OPS_SHEET_NAME = 'Snowmaking Operations';
const OPS_EQUIPMENT = ['Pump House', 'Snow Gun 1', 'Snow Gun 2', 'Snow Gun 3'];
const OPS_HEADERS = ['Recorded At (PKT)', 'Event Time (PKT)', 'Equipment', 'Action', 'Runtime Hours'];

function getOpsSheet_() {
  const ss = SpreadsheetApp.openById(OPS_SPREADSHEET_ID);
  let sheet = ss.getSheetByName(OPS_SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(OPS_SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(OPS_HEADERS);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, OPS_HEADERS.length).setFontWeight('bold');
  }
  return sheet;
}

function doPost(e) {
  try {
    const p = (e && e.parameter) || {};
    const equipment = String(p.equipment || '');
    const action = String(p.action || '').toUpperCase();
    const eventTime = new Date(String(p.enteredAt || ''));
    if (OPS_EQUIPMENT.indexOf(equipment) === -1) throw new Error('Unknown equipment');
    if (['START', 'STOP'].indexOf(action) === -1) throw new Error('Invalid action');
    if (isNaN(eventTime.getTime())) throw new Error('Invalid event time');

    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      const sheet = getOpsSheet_();
      const state = buildOpsState_(sheet);
      const current = state[equipment];
      if (action === 'START' && current.status === 'RUNNING') throw new Error(equipment + ' is already running');
      if (action === 'STOP' && current.status !== 'RUNNING') throw new Error(equipment + ' is not running');

      let runtime = '';
      if (action === 'STOP') {
        const start = new Date(current.openStart);
        const hours = (eventTime.getTime() - start.getTime()) / 3600000;
        if (hours < 0) throw new Error('STOP time is earlier than the recorded START time');
        runtime = Math.round(hours * 100) / 100;
      }
      sheet.appendRow([new Date(), eventTime, equipment, action, runtime]);
    } finally {
      lock.releaseLock();
    }
    return ContentService.createTextOutput('OK');
  } catch (err) {
    return ContentService.createTextOutput('ERROR: ' + err.message);
  }
}

function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.action !== 'state') return ContentService.createTextOutput('MJSR Snowmaking Operations Logger');
  const data = { equipment: [] };
  try {
    const sheet = getOpsSheet_();
    const state = buildOpsState_(sheet);
    OPS_EQUIPMENT.forEach(name => {
      const s = state[name];
      data.equipment.push({
        name: name,
        status: s.status,
        lastStart: s.lastStart,
        lastStop: s.lastStop,
        tonight: s.tonight,
        season: s.season,
        openStart: s.openStart
      });
    });
  } catch (err) {
    data.error = err.message;
  }
  const json = JSON.stringify(data);
  const callback = String(p.callback || '');
  if (callback && /^[a-zA-Z_$][0-9a-zA-Z_$\.]*$/.test(callback)) {
    return ContentService.createTextOutput(callback + '(' + json + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

function buildOpsState_(sheet) {
  const result = {};
  OPS_EQUIPMENT.forEach(name => result[name] = {
    status: 'STOPPED', lastStart: '—', lastStop: '—',
    tonight: 0, season: 0, openStart: ''
  });
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return result;
  const rows = sheet.getRange(2, 1, lastRow - 1, OPS_HEADERS.length).getValues();
  rows.sort((a, b) => new Date(a[1]).getTime() - new Date(b[1]).getTime());
  const now = new Date();
  const today = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const nightStart = new Date(today + 'T17:00:00');
  if (now < nightStart) nightStart.setDate(nightStart.getDate() - 1);
  const nightEnd = new Date(nightStart.getTime() + 15 * 3600000);

  rows.forEach(row => {
    const eventTime = row[1] instanceof Date ? row[1] : new Date(row[1]);
    const equipment = String(row[2]);
    const action = String(row[3]).toUpperCase();
    const s = result[equipment];
    if (!s || isNaN(eventTime.getTime())) return;
    if (action === 'START') {
      s.status = 'RUNNING';
      s.openStart = eventTime.toISOString();
      s.lastStart = Utilities.formatDate(eventTime, Session.getScriptTimeZone(), 'dd MMM HH:mm');
    } else if (action === 'STOP') {
      s.status = 'STOPPED';
      s.openStart = '';
      s.lastStop = Utilities.formatDate(eventTime, Session.getScriptTimeZone(), 'dd MMM HH:mm');
      const hours = Number(row[4]) || 0;
      s.season += hours;
      const overlapStart = Math.max(eventTime.getTime() - hours * 3600000, nightStart.getTime());
      const overlapEnd = Math.min(eventTime.getTime(), nightEnd.getTime());
      if (overlapEnd > overlapStart) s.tonight += (overlapEnd - overlapStart) / 3600000;
    }
  });
  OPS_EQUIPMENT.forEach(name => {
    const s = result[name];
    if (s.status === 'RUNNING' && s.openStart) {
      const elapsed = Math.max(0, (now - new Date(s.openStart)) / 3600000);
      s.tonight += elapsed;
      s.season += elapsed;
    }
    s.tonight = Math.round(s.tonight * 100) / 100;
    s.season = Math.round(s.season * 100) / 100;
  });
  return result;
}
