/**
 * FRC Krillion scoreboard — Google Apps Script backend (one dive per name).
 * Deploy > Manage deployments > Edit > Version: New version, to update without changing the URL.
 *
 * Each player name gets one dive. Starting a dive reserves the name; finishing fills in the score.
 * The leaderboard is only returned to someone who has finished a dive (they send their run id).
 */
const SHEET_NAME = 'Scores';
const MAX_RETURN = 500;
const TIERS = { 'Plankton': 10, 'Schooler': 30, 'Rare': 60, 'Deep Cut': 85, 'One in a Krillion': 100, 'Miss': 0 };
// Columns: A Saved at | B Run id | C Name | D Total | E Played at | F Answers (JSON) | G Status

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(['Saved at', 'Run id', 'Name', 'Total', 'Played at', 'Answers (JSON)', 'Status']);
    sh.setFrozenRows(1);
  }
  if (sh.getRange(1, 7).getValue() !== 'Status') sh.getRange(1, 7).setValue('Status');
  return sh;
}
function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
function key_(name) { return String(name || '').toLowerCase().replace(/[^a-z0-9]/g, ''); }
function cleanName_(name) { return String(name || '').replace(/[\u0000-\u001f]/g, '').trim().slice(0, 30); }
function rows_() {
  const sh = sheet_();
  const last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, 7).getValues().map((r, i) => ({ row: i + 2, r: r }));
}
function nameTaken_(name) {
  const k = key_(name);
  return rows_().some(x => key_(x.r[2]) === k);
}

function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.action === 'check') return json_({ taken: nameTaken_(p.name) });
  if (p.action === 'board') {
    const all = rows_();
    const mine = all.find(x => String(x.r[1]) === String(p.id || ''));
    if (!mine || String(mine.r[6]) !== 'finished') return json_({ error: 'not_played' });
    const done = all.filter(x => String(x.r[6]) === 'finished').slice(-MAX_RETURN);
    return json_({ runs: done.map(x => {
      let ans = [];
      try { ans = JSON.parse(x.r[5]); } catch (err) {}
      return { id: String(x.r[1]), name: String(x.r[2]), total: Number(x.r[3]) || 0, at: x.r[4] ? new Date(x.r[4]).toISOString() : '', ans: ans };
    }) });
  }
  return json_({ ok: true });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const body = JSON.parse(e.postData.contents);
    const id = String(body.id || '').replace(/[^a-z0-9]/gi, '').slice(0, 20);
    if (!id) return json_({ ok: false, error: 'id' });
    const sh = sheet_();

    if (body.type === 'start') {
      const name = cleanName_(body.name);
      if (!name) return json_({ ok: false, error: 'name' });
      if (nameTaken_(name)) return json_({ ok: false, error: 'name_taken' });
      sh.appendRow([new Date(), id, name, '', new Date(), '[]', 'started']);
      return json_({ ok: true });
    }

    if (body.type === 'finish') {
      const mine = rows_().find(x => String(x.r[1]) === id);
      if (!mine) return json_({ ok: false, error: 'no_start' });
      if (String(mine.r[6]) === 'finished') return json_({ ok: false, error: 'already_finished' });
      const ans = Array.isArray(body.ans) ? body.ans.slice(0, 7) : [];
      if (ans.length !== 7) return json_({ ok: false, error: 'answers' });
      let total = 0;
      const cleanAns = ans.map(a => {
        const tier = TIERS.hasOwnProperty(a.tier) ? a.tier : 'Miss';
        total += TIERS[tier];
        return { input: String(a.input || '').slice(0, 60), match: String(a.match || '').slice(0, 80), tier: tier, p: TIERS[tier] };
      });
      sh.getRange(mine.row, 1, 1, 7).setValues([[new Date(), id, mine.r[2], total, mine.r[4], JSON.stringify(cleanAns), 'finished']]);
      return json_({ ok: true, total: total });
    }
    return json_({ ok: false, error: 'type' });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}
