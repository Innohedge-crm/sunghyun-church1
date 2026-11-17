/**
 * 성현교회 주보제작 — 구글시트 저장 서버
 * ─────────────────────────────────────────────
 * 이 스크립트는 「2026년 주보 업데이트」 시트에 붙어서 동작합니다.
 *  · 앱의 전체 데이터(모든 주차, 매달, 고정 정보, 찬송가표)는 「주보앱데이터」 탭에 저장합니다.
 *  · 가장 최근 주의 내용은 첫 번째 탭의 D열에도 field_key(A열) 기준으로 써 넣습니다.
 *    → 행 위치는 바꾸지 않으므로, 홈페이지가 D35·D43 등을 읽는 연동은 그대로 유지됩니다.
 *  · 수식이 들어 있는 칸은 건드리지 않습니다.
 *
 * ▼ 아래 비밀번호만 원하는 것으로 바꾼 뒤 저장하고 배포하세요.
 */
const PASSWORD = '여기에-비밀번호를-적으세요';

const DATA_SHEET = '주보앱데이터';
const MAIN_SHEET_GID = 0;          // 첫 번째(주보 내용) 탭
const CHUNK = 40000;               // 한 칸에 넣는 글자 수 (구글시트 한 칸 최대 50,000자)

function doGet() {
  return json_({ ok: true, app: '성현교회 주보제작 서버', time: new Date().toISOString() });
}

function doPost(e) {
  let req;
  try { req = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: 'bad_request' }); }
  if (!req || req.pw !== PASSWORD) return json_({ ok: false, error: 'auth' });

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) return json_({ ok: false, error: 'busy' });
  try {
    if (req.action === 'ping') return json_({ ok: true });
    if (req.action === 'load') return json_(Object.assign({ ok: true }, readData_()));
    if (req.action === 'save') {
      const cur = readData_();
      if (!req.force && cur.rev && req.baseRev !== cur.rev) {
        return json_({ ok: false, error: 'conflict', rev: cur.rev, savedAt: cur.savedAt });
      }
      const saved = writeData_(req.data);
      let sheetCells = 0;
      if (req.sheet) sheetCells = writeMain_(req.sheet);
      return json_({ ok: true, rev: saved.rev, savedAt: saved.savedAt, sheetCells: sheetCells });
    }
    return json_({ ok: false, error: 'unknown_action' });
  } catch (err) {
    return json_({ ok: false, error: 'server', message: String(err && err.message || err) });
  } finally {
    lock.releaseLock();
  }
}

/* ── 전체 데이터: 「주보앱데이터」 탭 ── */
function dataSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(DATA_SHEET);
  if (!sh) {
    sh = ss.insertSheet(DATA_SHEET);
    sh.getRange('A1:C1').setValues([['rev', 'savedAt', '이 탭은 주보제작 앱이 쓰는 저장 공간입니다. 직접 고치지 마세요.']]);
  }
  return sh;
}
function readData_() {
  const sh = dataSheet_();
  const head = sh.getRange('A2:B2').getValues()[0];
  const last = sh.getLastRow();
  if (last < 3 || !head[0]) return { data: null, rev: '', savedAt: '' };
  const parts = sh.getRange(3, 1, last - 2, 1).getValues().map(r => String(r[0] || ''));
  const text = parts.filter(Boolean).map(p => p.slice(1)).join('');   // 각 칸 앞의 'x' 제거
  let data = null;
  try { data = JSON.parse(text); } catch (err) { data = null; }
  return { data: data, rev: String(head[0]), savedAt: String(head[1]) };
}
function writeData_(data) {
  const sh = dataSheet_();
  const text = JSON.stringify(data);
  const parts = [];
  for (let i = 0; i < text.length; i += CHUNK) parts.push(['x' + text.slice(i, i + CHUNK)]);  // 'x'로 시작 → 수식으로 해석되지 않음
  const last = sh.getLastRow();
  if (last >= 3) sh.getRange(3, 1, last - 2, 1).clearContent();
  if (parts.length) { const r = sh.getRange(3, 1, parts.length, 1); r.setNumberFormat('@'); r.setValues(parts); }
  const rev = Utilities.getUuid().slice(0, 8);
  const savedAt = new Date().toISOString();
  sh.getRange('A2:B2').setNumberFormat('@').setValues([[rev, savedAt]]);
  return { rev: rev, savedAt: savedAt };
}

/* ── 이번 주 내용: 첫 번째 탭 D열 (field_key 기준) ── */
function mainSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheets().filter(s => s.getSheetId() === MAIN_SHEET_GID)[0] || ss.getSheets()[0];
}
function normKey_(s) { return String(s || '').trim().toLowerCase().replace(/\s+/g, '_'); }

/**
 * sheet = { values: { field_key: '값', ... }, titles: { field_key: '찬송 제목' }, wolsak: { label, scripture, preacher } }
 */
function writeMain_(sheet) {
  const sh = mainSheet_();
  const last = sh.getLastRow();
  if (last < 2) return 0;
  const rng = sh.getRange(1, 1, last, 6);
  const vals = rng.getValues();
  const forms = rng.getFormulas();
  let n = 0;
  const put = (r, c, v) => {               // r, c: 0-based
    if (forms[r][c]) return;               // 수식 칸은 그대로 둔다
    if (String(vals[r][c]) === String(v)) return;
    const cell = sh.getRange(r + 1, c + 1);
    cell.setNumberFormat('@');             // 날짜·숫자로 바뀌지 않게
    cell.setValue(v);
    n++;
  };
  const values = sheet.values || {}, titles = sheet.titles || {};
  for (let r = 1; r < vals.length; r++) {
    const key = normKey_(vals[r][0]);
    if (!key) continue;
    if (Object.prototype.hasOwnProperty.call(values, key)) put(r, 3, values[key]);
    if (Object.prototype.hasOwnProperty.call(titles, key)) put(r, 5, titles[key]);
  }
  // 월삭기도회 줄: A열이 비어 있고 C열이 '월삭'으로 시작하거나 F열 안내문에 '월삭기도회'가 있는 줄
  if (sheet.wolsak) {
    for (let r = 1; r < vals.length; r++) {
      if (normKey_(vals[r][0])) continue;
      const c = String(vals[r][2] || ''), f = String(vals[r][5] || '');
      if (c.indexOf('월삭') === 0 || f.indexOf('월삭기도회') >= 0) {
        if (!sheet.wolsak.label && !c) break;   // 월삭기도회가 없는 달 + 이미 비어 있음 → 예시 문구 그대로 둠
        put(r, 2, sheet.wolsak.label || '');
        put(r, 3, sheet.wolsak.scripture || '');
        put(r, 4, sheet.wolsak.preacher || '');
        break;
      }
    }
  }
  return n;
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

/** 배포 전에 한 번 실행해서 권한을 허락하고, 저장 탭이 만들어지는지 확인하는 용도 */
function 처음_실행() {
  dataSheet_();
  Logger.log('준비 완료: 「' + DATA_SHEET + '」 탭이 있습니다. 첫 번째 탭: ' + mainSheet_().getName());
}
