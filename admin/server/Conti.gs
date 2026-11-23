/**
 * 성현교회 찬양콘티 — 악보 내려받기 서버
 * ─────────────────────────────────────────────
 * 홈페이지 「찬양콘티」 메뉴가 이 서버에 비밀번호를 보내 악보 목록과 파일을 받아 갑니다.
 *  · 구글드라이브 폴더는 공개(링크 공유)하지 않아도 됩니다. 이 스크립트가 대신 읽어서 전달합니다.
 *  · 폴더 안에 하위 폴더(예: 「2026.12.06」)를 만들면 홈페이지에서 날짜별로 묶여 보입니다.
 *  · 주보제작 서버(Code.gs)와는 별개의 새 Apps Script 프로젝트로 만드세요.
 *
 * ▼ 아래 두 줄만 바꾼 뒤 저장하고, 웹 앱으로 배포하세요.
 */
const PASSWORD  = '여기에-찬양팀-비밀번호를-적으세요';
const FOLDER_ID = '여기에-악보-폴더-ID를-적으세요';   // 폴더 주소 .../folders/ 뒤의 글자

const MAX_BYTES = 30 * 1024 * 1024;   // 한 파일 최대 30MB

function doGet() {
  return json_({ ok: true, app: '성현교회 찬양콘티 서버', time: new Date().toISOString() });
}

function doPost(e) {
  let req;
  try { req = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: 'bad_request' }); }
  if (!req || req.pw !== PASSWORD) return json_({ ok: false, error: 'auth' });
  try {
    if (req.action === 'ping') return json_({ ok: true });
    if (req.action === 'list') return json_({ ok: true, groups: list_() });
    if (req.action === 'file') return json_(file_(String(req.id || '')));
    return json_({ ok: false, error: 'unknown_action' });
  } catch (err) {
    return json_({ ok: false, error: 'server', message: String(err && err.message || err) });
  }
}

/* 폴더 바로 아래 파일 + 하위 폴더(한 단계)의 파일 목록 */
function list_() {
  const root = DriveApp.getFolderById(FOLDER_ID);
  const groups = [];
  const sub = root.getFolders();
  while (sub.hasNext()) {
    const f = sub.next();
    const files = files_(f);
    if (files.length) groups.push({ name: f.getName(), updated: f.getLastUpdated().toISOString(), files: files });
  }
  // 하위 폴더는 이름 역순(최신 날짜가 위로)
  groups.sort((a, b) => b.name.localeCompare(a.name, 'ko', { numeric: true }));
  const top = files_(root);
  if (top.length) groups.push({ name: groups.length ? '기타 악보' : '', files: top });
  return groups;
}

function files_(folder) {
  const out = [];
  const it = folder.getFiles();
  while (it.hasNext()) {
    const f = it.next();
    out.push({ id: f.getId(), name: f.getName(), mime: f.getMimeType(), size: f.getSize(), updated: f.getLastUpdated().toISOString() });
  }
  out.sort((a, b) => a.name.localeCompare(b.name, 'ko', { numeric: true }));
  return out;
}

/* 파일 한 개를 base64로 전달 (악보 폴더 안의 파일만 허용) */
function file_(id) {
  if (!id) return { ok: false, error: 'no_id' };
  const f = DriveApp.getFileById(id);
  if (!inFolder_(f)) return { ok: false, error: 'forbidden' };
  let blob, name = f.getName();
  const mime = f.getMimeType();
  if (mime.indexOf('application/vnd.google-apps') === 0) {      // 구글 문서·슬라이드 등은 PDF로 변환
    blob = f.getAs('application/pdf');
    if (!/\.pdf$/i.test(name)) name += '.pdf';
  } else {
    blob = f.getBlob();
  }
  const bytes = blob.getBytes();
  if (bytes.length > MAX_BYTES) return { ok: false, error: 'too_big' };
  return { ok: true, name: name, mime: blob.getContentType(), data: Utilities.base64Encode(bytes) };
}

function inFolder_(file) {
  const ps = file.getParents();
  while (ps.hasNext()) {
    const p = ps.next();
    if (p.getId() === FOLDER_ID) return true;
    const pp = p.getParents();
    while (pp.hasNext()) if (pp.next().getId() === FOLDER_ID) return true;
  }
  return false;
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

/** 배포 전에 한 번 실행해서 드라이브 권한을 허락하고, 폴더를 읽을 수 있는지 확인하는 용도 */
function 처음_실행() {
  const g = list_();
  Logger.log('준비 완료: ' + g.length + '개 묶음, 파일 ' + g.reduce((n, x) => n + x.files.length, 0) + '개');
}
