/**
 * IDoR 2026 · "Safe or Not Safe?" MRI safety game — anonymous results and player ranking
 * Medical Imaging Department, King Abdullah Specialized Hospital
 *
 * Paste this into Extensions → Apps Script of your results Google Sheet (replace everything in Code.gs).
 * The game sends one row per finished round and gets back the player's position, e.g. 2 of 30.
 * Nothing personal is stored: no names, no device details, no location.
 * Ranking: higher score first; equal scores are ordered by the shorter answering time.
 * To restart the ranking (for example after testing), delete the data rows and keep row 1.
 */

const SHEET_NAME = 'Results';
const TIME_ZONE = 'Asia/Riyadh';
const HEADERS = ['Date', 'Time', 'Score', 'Out of', 'Seconds taken', 'Level', 'Language', 'Missed items'];
const LEVELS = ['MRI Safety Champion', 'Safety Star', 'Good Start', 'Keep Learning'];
const MIN_MS = 5000;      // 10 answers in under 5 seconds is not a real round
const MAX_MS = 1800000;   // 30 minutes

// Game item codes → names written to the sheet (only these can ever be written)
const ITEMS = {
  keys: 'Keys', phone: 'Mobile phone', coins: 'Coins', card: 'Bank card',
  scarfpins: 'Headscarf pins', hairpins: 'Hairpins and clips', watch: 'Watch',
  earbuds: 'Wireless earbuds', glasses: 'Glasses', jewellery: 'Earrings and necklace',
  stroller: 'Baby stroller', hearing: 'Hearing aids', wheelchair: 'Own wheelchair',
  pacemaker: 'Pacemaker or heart device', cochlear: 'Cochlear implant',
  eye: 'Old eye injury from metal work', plate: 'Plates or screws',
  pump: 'Insulin pump or sugar sensor', patch: 'Medicine patch',
  tattoo: 'Tattoo or permanent make-up', pregnant: 'Pregnancy',
  gown: 'Hospital gown', earplugs: 'Foam earplugs', headphones: 'MRI team headphones',
  blanket: 'Cotton blanket', scarf: 'Plain cotton headscarf', family: 'Screened family member'
};

function doPost(e) {
  try {
    const data = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (data.app !== 'idor2026-safe-or-not') return json_({ ok: false });

    const score = Number(data.score);
    const total = Number(data.total);
    if (!Number.isInteger(score) || total !== 10 || score < 0 || score > total) return json_({ ok: false });

    const ms = Number(data.timeMs);
    if (!isFinite(ms) || ms < MIN_MS || ms > MAX_MS) return json_({ ok: false });
    const seconds = Math.round(ms / 100) / 10;

    if (LEVELS.indexOf(data.rank) === -1) return json_({ ok: false });
    const language = data.lang === 'ar' ? 'Arabic' : data.lang === 'en' ? 'English' : null;
    if (!language) return json_({ ok: false });

    const missed = (Array.isArray(data.missed) ? data.missed : [])
      .filter(function (id) { return Object.prototype.hasOwnProperty.call(ITEMS, id); })
      .slice(0, 10)
      .map(function (id) { return ITEMS[id]; });

    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      const sheet = getSheet_();

      // Position among everyone recorded so far: count players with a better result
      let better = 0;
      let players = 0;
      const last = sheet.getLastRow();
      if (last > 1) {
        const rows = sheet.getRange(2, 3, last - 1, 3).getValues(); // Score, Out of, Seconds taken
        rows.forEach(function (r) {
          const s = Number(r[0]);
          const t = Number(r[2]);
          if (r[0] === '' || !isFinite(s) || !isFinite(t)) return;
          players++;
          if (s > score || (s === score && t < seconds)) better++;
        });
      }

      const now = new Date();
      sheet.appendRow([
        Utilities.formatDate(now, TIME_ZONE, 'yyyy-MM-dd'),
        Utilities.formatDate(now, TIME_ZONE, 'HH:mm'),
        score, total, seconds, data.rank, language, missed.join(', ')
      ]);

      return json_({ ok: true, rank: better + 1, total: players + 1 });
    } finally {
      lock.releaseLock();
    }
  } catch (err) {
    return json_({ ok: false });
  }
}

// Opening the web app address in a browser shows this line, so you can check it is running.
function doGet() {
  return ContentService.createTextOutput('IDoR 2026 results collector is running.');
}

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
  }
  return sheet;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
