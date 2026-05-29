/**
 * July 4 Game Vote — Google Apps Script backend
 *
 * Deploy this in Google Apps Script bound to a Google Sheet, then paste the
 * resulting Web App URL into index.html (APPS_SCRIPT_URL constant).
 *
 * Setup steps:
 *   1. Create a new blank Google Sheet (rename to "July 4 Block Party Vote" or anything).
 *   2. Extensions → Apps Script. Delete the default Code.gs body.
 *   3. Paste THIS ENTIRE FILE into Code.gs and save.
 *   4. Deploy → New deployment → Type: Web app.
 *        - Description: "July 4 Vote API"
 *        - Execute as: Me
 *        - Who has access: Anyone
 *   5. Authorize when prompted. Copy the Web app URL (ends in /exec).
 *   6. Paste that URL into index.html as APPS_SCRIPT_URL, commit, push.
 *
 * Re-deploying after edits: Deploy → Manage deployments → pencil icon →
 *   Version: New version → Deploy. The URL stays the same.
 */

const SHEET_NAME = 'Votes';
const HEADERS = ['timestamp', 'name', 'email', 'games'];

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(HEADERS);
    sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/** GET — returns all voters so the frontend can tally and render results. */
function doGet(e) {
  const sh = getSheet_();
  const data = sh.getDataRange().getValues();
  const voters = data.slice(1).map(r => ({
    name: String(r[1] || ''),
    games: String(r[3] || '').split(',').map(s => s.trim()).filter(Boolean),
    ts: r[0] instanceof Date ? r[0].getTime() : Date.parse(r[0]) || null,
  })).filter(v => v.name && v.games.length);
  return json_({ ok: true, voters: voters });
}

/** POST — records one vote. Rejects if email already voted. */
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const name  = String(body.name  || '').trim();
    const email = String(body.email || '').trim().toLowerCase();
    const games = Array.isArray(body.games) ? body.games.map(String) : [];

    if (!name)  return json_({ ok: false, error: 'missing_name'  });
    if (!email) return json_({ ok: false, error: 'missing_email' });
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
      return json_({ ok: false, error: 'bad_email' });
    }
    if (!games.length) return json_({ ok: false, error: 'no_games' });

    // Use a lock so two simultaneous votes from the same email can't both pass the dedup check.
    const lock = LockService.getScriptLock();
    lock.waitLock(5000);
    try {
      const sh = getSheet_();
      const data = sh.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][2] || '').trim().toLowerCase() === email) {
          return json_({ ok: false, error: 'already_voted' });
        }
      }
      sh.appendRow([new Date(), name, email, games.join(',')]);
    } finally {
      lock.releaseLock();
    }
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: 'server_error', detail: String(err) });
  }
}
