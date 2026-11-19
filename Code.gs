/**
 * FRC Krillion scoreboard — Google Apps Script backend.
 * Paste this into Extensions > Apps Script on a new Google Sheet,
 * then Deploy > New deployment > Web app (Execute as: Me, Who has access: Anyone).
 * The sheet tab "Scores" is created automatically on the first save.
 */
const SHEET_NAME = 'Scores';
const MAX_RETURN = 500;
const TIERS = { 'Plankton': 10, 'Schooler': 30, 'Rare': 60, 'Deep Cut': 85, 'One in a Krillion': 100, 'Miss': 0 };

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(['Saved at', 'Run id', 'Name', 'Total', 'Played at', 'Answers (JSON)']);
    sh.setFrozenRows(1);
  }
  return sh;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function doGet() {
  const sh = sheet_();
  const last = sh.getLastRow();
  if (last < 2) return json_({ runs: [] });
  const start = Math.max(2, last - MAX_RETURN + 1);
  const rows = sh.getRange(start, 1, last - start + 1, 6).getValues();
  const runs = rows.map(r => {
    let ans = [];
    try { ans = JSON.parse(r[5]); } catch (e) {}
    return { id: String(r[1]), name: String(r[2]), total: Number(r[3]) || 0, at: r[4] ? new Date(r[4]).toISOString() : '', ans: ans };
  });
  return json_({ runs: runs });
}

function doPost(e) {
  try {
    const run = JSON.parse(e.postData.contents);
    const name = String(run.name || '').replace(/[\u0000-\u001f]/g, '').trim().slice(0, 30);
    if (!name) return json_({ ok: false, error: 'name' });
    const ans = Array.isArray(run.ans) ? run.ans.slice(0, 7) : [];
    if (ans.length !== 7) return json_({ ok: false, error: 'answers' });
    let total = 0;
    const cleanAns = ans.map(a => {
      const tier = TIERS.hasOwnProperty(a.tier) ? a.tier : 'Miss';
      const p = TIERS[tier];
      total += p;
      return { input: String(a.input || '').slice(0, 60), match: String(a.match || '').slice(0, 80), tier: tier, p: p };
    });
    const id = String(run.id || '').replace(/[^a-z0-9]/gi, '').slice(0, 20);
    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      sheet_().appendRow([new Date(), id, name, total, run.at ? new Date(run.at) : new Date(), JSON.stringify(cleanAns)]);
    } finally {
      lock.releaseLock();
    }
    return json_({ ok: true, total: total });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}
