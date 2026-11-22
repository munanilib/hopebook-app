/**
 * 📥 이전 사본에서 데이터 가져오기
 *
 * 옛날 방식의 사본(또는 다른 학교 사본)에 쌓인 신청 데이터를 이 사본으로 옮깁니다.
 * - 이전 사본은 읽기만 하고 절대 바꾸지 않습니다.
 * - 옮기는 것: 아래 목록의 데이터 시트(값·수식·숫자 서식)
 * - 못 옮기는 것: 학교 이름·API 키·공개 신청 링크 같은 '스크립트 속성'
 *   → 가져온 뒤 「⚙️ 사본 설정」 메뉴에서 다시 입력해 주세요.
 */
var HR_IMPORT_SHEET_NAMES = [
  '📚 희망도서 목록',
  '📥 엑셀 업로드 목록',
  '🛒 구입 검토함',
  '📦 서점 전달용',
  '🏫 우리학교 소장도서',
  '🔒 신청취소복구',
  '🔒 신청접수백업',
  '🔎 시리즈 누락 점검',
  '📋 시리즈 업무목록',
  '📢 웹 공지',
  '⚙️ 통합 설정'
];
var HR_IMPORT_CHUNK_ROWS = 400;
var HR_IMPORT_TIME_LIMIT_MS = 4.5 * 60 * 1000;

function HR_importFromOldCopy() {
  var ui = SpreadsheetApp.getUi();
  var dest = SpreadsheetApp.getActiveSpreadsheet();

  var ask = ui.prompt(
    '📥 이전 사본에서 데이터 가져오기',
    '이전(옛날) 사본의 주소(URL)를 붙여 넣어 주세요.\n' +
    '※ 이전 사본은 읽기만 하고 바꾸지 않습니다.\n' +
    '※ 그 사본을 열 수 있는 계정으로 로그인되어 있어야 합니다.',
    ui.ButtonSet.OK_CANCEL
  );
  if (ask.getSelectedButton() !== ui.Button.OK) return;

  var url = String(ask.getResponseText() || '').trim();
  if (!url) { ui.alert('주소가 비어 있습니다.'); return; }

  var src;
  try {
    src = SpreadsheetApp.openByUrl(url);
  } catch (err) {
    ui.alert('이전 사본을 열 수 없습니다',
      '주소가 맞는지, 이 계정으로 그 사본을 볼 수 있는지 확인해 주세요.\n\n(' + (err && err.message || err) + ')',
      ui.ButtonSet.OK);
    return;
  }
  if (src.getId() === dest.getId()) {
    ui.alert('지금 열려 있는 사본과 같은 파일입니다. 이전 사본의 주소를 넣어 주세요.');
    return;
  }

  var plan = HR_importPlan_(src, dest);
  if (!plan.length) {
    ui.alert('가져올 데이터 시트를 찾지 못했습니다',
      '이전 사본에 「' + HR_IMPORT_SHEET_NAMES[0] + '」 같은 시트가 없습니다. 주소를 다시 확인해 주세요.',
      ui.ButtonSet.OK);
    return;
  }

  var lines = plan.map(function (p) {
    return '• ' + p.name + ' — 가져올 ' + p.rows + '행' +
      (p.destRows > 1 ? ' (이 사본의 현재 ' + p.destRows + '행은 지워집니다)' : '');
  });
  var ok = ui.alert(
    '이대로 가져올까요?',
    '이전 사본: ' + src.getName() + '\n\n' + lines.join('\n') +
    '\n\n같은 이름의 시트 내용은 이전 사본의 내용으로 바뀝니다.',
    ui.ButtonSet.YES_NO
  );
  if (ok !== ui.Button.YES) return;

  var report = HR_importRun_(src, dest, plan);

  try { if (typeof ensureSelectionControls_ === 'function') ensureSelectionControls_(); } catch (e1) {}
  try { if (typeof styleRequestListSheet === 'function') styleRequestListSheet(); } catch (e2) {}

  ui.alert(
    report.failed.length || report.skipped.length ? '가져오기를 마쳤습니다 (확인 필요)' : '✅ 가져오기를 마쳤습니다',
    report.done.map(function (d) { return '✔ ' + d.name + ' ' + d.rows + '행'; }).join('\n') +
    (report.failed.length ? '\n\n⚠️ 실패:\n' + report.failed.join('\n') : '') +
    (report.skipped.length ? '\n\n⏱ 시간이 부족해 건너뜀(다시 실행하면 이어서 가져옵니다):\n' + report.skipped.join('\n') : '') +
    '\n\n※ 학교 이름·API 키·공개 신청 링크는 옮겨지지 않았습니다.\n' +
    '「⚙️ 사본 설정」 메뉴에서 다시 입력해 주세요.',
    ui.ButtonSet.OK
  );
}

function HR_importPlan_(src, dest) {
  var plan = [];
  HR_IMPORT_SHEET_NAMES.forEach(function (name) {
    var s = src.getSheetByName(name);
    if (!s) return;
    var t = dest.getSheetByName(name);
    plan.push({ name: name, rows: s.getLastRow(), cols: s.getLastColumn(), destRows: t ? t.getLastRow() : 0 });
  });
  return plan;
}

function HR_importRun_(src, dest, plan) {
  var started = Date.now();
  var report = { done: [], failed: [], skipped: [] };
  plan.forEach(function (p) {
    if (Date.now() - started > HR_IMPORT_TIME_LIMIT_MS) { report.skipped.push(p.name); return; }
    try {
      HR_importSheet_(src.getSheetByName(p.name), dest, p.name);
      report.done.push({ name: p.name, rows: p.rows });
    } catch (err) {
      report.failed.push(p.name + ' — ' + (err && err.message || err));
    }
  });
  return report;
}

function HR_importSheet_(s, dest, name) {
  var t = dest.getSheetByName(name);
  var rows = s.getLastRow();
  var cols = s.getLastColumn();

  // 새 사본에 없는 시트면 서식째로 통째로 복사
  if (!t) {
    var copied = s.copyTo(dest);
    copied.setName(name);
    return;
  }
  if (rows < 1 || cols < 1) return; // 이전 사본이 비어 있으면 그대로 둠

  if (t.getMaxRows() < rows) t.insertRowsAfter(t.getMaxRows(), rows - t.getMaxRows());
  if (t.getMaxColumns() < cols) t.insertColumnsAfter(t.getMaxColumns(), cols - t.getMaxColumns());
  t.getDataRange().clearContent();

  for (var r = 1; r <= rows; r += HR_IMPORT_CHUNK_ROWS) {
    var n = Math.min(HR_IMPORT_CHUNK_ROWS, rows - r + 1);
    var from = s.getRange(r, 1, n, cols);
    var values = from.getValues();
    var formulas = from.getFormulas();
    var out = values.map(function (row, i) {
      return row.map(function (v, j) { return formulas[i][j] ? formulas[i][j] : v; });
    });
    var to = t.getRange(r, 1, n, cols);
    to.setNumberFormats(from.getNumberFormats());
    to.setValues(out);
  }
}
