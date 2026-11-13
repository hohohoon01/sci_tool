/**
 * 과학실 교구목록 웹앱 - 구글 시트 연결용 Apps Script
 * 시트 탭 이름: '교구목록'(A~I열 + J열 ID), '설정'(A~D열 드롭다운 목록)
 */
const SHEET = '교구목록';
const SETTINGS = '설정';
const PASSWORD = ''; // 편집 암호. 비워 두면 주소를 아는 누구나 추가·수정·삭제할 수 있습니다.

function doGet() {
  try { return out_(list_()); }
  catch (err) { return out_({ ok: false, error: String(err) }); }
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const req = JSON.parse(e.postData.contents);
    if (PASSWORD && req.pw !== PASSWORD) return out_({ ok: false, error: 'auth' });
    const sh = sheet_();
    const it = req.item || {};
    if (req.action === 'add') {
      const row = firstEmptyRow_(sh);
      const id = Utilities.getUuid().slice(0, 8);
      writeRow_(sh, row, it, id);
      return out_({ ok: true, id: id });
    }
    const row = findRow_(sh, it.id);
    if (!row) return out_({ ok: false, error: '시트에서 해당 교구를 찾지 못했습니다.' });
    if (req.action === 'update') { writeRow_(sh, row, it, it.id); return out_({ ok: true }); }
    if (req.action === 'delete') { sh.deleteRow(row); return out_({ ok: true }); }
    return out_({ ok: false, error: '알 수 없는 요청' });
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function sheet_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET);
  if (!sh) throw new Error("'" + SHEET + "' 탭이 없습니다.");
  if (sh.getRange(1, 10).getValue() !== 'ID') sh.getRange(1, 10).setValue('ID');
  return sh;
}

function list_() {
  const sh = sheet_();
  const last = sh.getLastRow();
  const items = [];
  if (last >= 2) {
    const rows = sh.getRange(2, 1, last - 1, 10).getDisplayValues();
    let changed = false;
    rows.forEach(function (r) {
      const filled = r.slice(0, 8).some(function (v) { return String(v).trim() !== ''; });
      if (!filled) return;
      if (!r[9]) { r[9] = Utilities.getUuid().slice(0, 8); changed = true; }
      items.push({ id: r[9], jang: r[0].trim(), dan: r[1].trim(), lr: r[2].trim(), name: r[3].trim(),
        spec: r[4].trim(), qty: r[5].trim(), unit: r[6].trim(), note: r[7].trim() });
    });
    if (changed) { // 시트에 직접 입력한 줄에 ID를 붙입니다.
      const lock = LockService.getScriptLock();
      lock.waitLock(20000);
      try { sh.getRange(2, 10, rows.length, 1).setValues(rows.map(function (r) { return [r[9]]; })); }
      finally { lock.releaseLock(); }
    }
  }
  return { ok: true, items: items, settings: settings_() };
}

function settings_() {
  const st = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SETTINGS);
  const res = { jang: [], dan: [], lr: [], unit: [] };
  if (!st || st.getLastRow() < 2) return res;
  const keys = ['jang', 'dan', 'lr', 'unit'];
  st.getRange(2, 1, st.getLastRow() - 1, 4).getDisplayValues().forEach(function (r) {
    keys.forEach(function (k, i) { const v = String(r[i]).trim(); if (v) res[k].push(v); });
  });
  return res;
}

function firstEmptyRow_(sh) {
  const last = sh.getLastRow();
  if (last < 2) return 2;
  const rows = sh.getRange(2, 1, last - 1, 8).getDisplayValues();
  for (let i = 0; i < rows.length; i++) {
    if (rows[i].every(function (v) { return String(v).trim() === ''; })) return i + 2;
  }
  return last + 1;
}

function findRow_(sh, id) {
  if (!id || sh.getLastRow() < 2) return 0;
  const ids = sh.getRange(2, 10, sh.getLastRow() - 1, 1).getDisplayValues();
  for (let i = 0; i < ids.length; i++) if (ids[i][0] === String(id)) return i + 2;
  return 0;
}

function writeRow_(sh, row, it, id) {
  const t = function (v) { // 글자로만 저장(수식·날짜 자동 변환 방지)
    v = String(v == null ? '' : v).trim().slice(0, 200);
    return /^[=+]/.test(v) ? "'" + v : v;
  };
  const qty = /^\d+(\.\d+)?$/.test(String(it.qty).trim()) ? Number(it.qty) : t(it.qty);
  sh.getRange(row, 1, 1, 5).setNumberFormat('@');
  sh.getRange(row, 7, 1, 2).setNumberFormat('@');
  sh.getRange(row, 1, 1, 8).setValues([[t(it.jang), t(it.dan), t(it.lr), t(it.name), t(it.spec), qty, t(it.unit), t(it.note)]]);
  sh.getRange(row, 9).setFormula('=IF(A' + row + '="","",A' + row + '&" "&B' + row + '&IF(C' + row + '="",""," "&C' + row + '))');
  sh.getRange(row, 10).setValue(id);
}

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
