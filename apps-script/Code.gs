// Clark County Home Shortlist: shared picks for Jerson and Melissa.
// Stores one row per (home, person) in the "Picks" tab of this spreadsheet.
const PEOPLE = ['jerson', 'melissa'];
const VOTES = ['', 'yes', 'maybe', 'no'];

function sheet_() {
  const ss = SpreadsheetApp.getActive();
  let s = ss.getSheetByName('Picks');
  if (!s) {
    s = ss.insertSheet('Picks');
    s.appendRow(['home', 'who', 'vote', 'note', 'updated']);
    s.setFrozenRows(1);
  }
  return s;
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function readPicks_() {
  const rows = sheet_().getDataRange().getValues().slice(1);
  const picks = {};
  rows.forEach(r => {
    const id = String(r[0]), who = String(r[1]), v = String(r[2] || ''), n = String(r[3] || '');
    if (!id || PEOPLE.indexOf(who) < 0 || (!v && !n)) return;
    (picks[who] = picks[who] || {})[id] = {v: v, n: n};
  });
  return {picks: picks};
}

function doGet() {
  return json_(readPicks_());
}

function doPost(e) {
  let b;
  try { b = JSON.parse(e.postData.contents); } catch (err) { return json_({error: 'bad json'}); }
  const who = String(b.who || '');
  const id = String(b.id || '').slice(0, 80);
  const v = String(b.v || '');
  let n = String(b.n || '').slice(0, 1000);
  if (PEOPLE.indexOf(who) < 0 || !/^[a-z0-9-]+$/.test(id) || VOTES.indexOf(v) < 0) return json_({error: 'bad input'});
  if (/^[=+\-@]/.test(n)) n = "'" + n;  // keep notes as plain text
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const s = sheet_();
    const vals = s.getDataRange().getValues();
    let row = -1;
    for (let i = 1; i < vals.length; i++) {
      if (String(vals[i][0]) === id && String(vals[i][1]) === who) { row = i + 1; break; }
    }
    const rec = [id, who, v, n, new Date()];
    if (row > 0) s.getRange(row, 1, 1, 5).setValues([rec]);
    else s.appendRow(rec);
  } finally {
    lock.releaseLock();
  }
  return json_({ok: true});
}
