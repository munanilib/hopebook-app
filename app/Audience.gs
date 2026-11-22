/**
 * 👥 신청 구분 (교직원 / 학부모)
 *
 * 한 시트의 「희망도서 목록」 맨 오른쪽에 「신청구분」 칸을 두고, 웹 화면에서 구분별로 갈라서 보여 줍니다.
 * - 「신청구분」이 비어 있는 예전 신청은 모두 교직원으로 봅니다.
 * - 신청 현황·취소·복구는 자기 구분의 신청만 다룰 수 있습니다.
 * - 신청 한도(1인당 권수)와 공지 대상(전체/교직원/학부모)도 구분별로 다룹니다.
 */
var HR_AUDIENCE_HEADER = '신청구분';
var HR_AUDIENCE_LABELS = { staff: '교직원', parent: '학부모' };
var HR_LIMIT_PROPERTIES = { staff: 'HR_LIMIT_STAFF', parent: 'HR_LIMIT_PARENT' };
// 한 번도 정하지 않았을 때의 기본값(0 = 제한 없음). 메뉴에서 정하면 그 값이 우선합니다.
var HR_DEFAULT_LIMITS = { staff: 0, parent: 5 };
var HR_NOTICE_AUDIENCE_HEADER = '대상';
var HR_NOTICE_AUDIENCE_VALUES = ['전체', '교직원', '학부모'];

// 'parent' | 'staff'  (모르거나 비어 있으면 교직원)
function HR_audienceKey_(value) {
  return /parent|학부모/i.test(String(value || '')) ? 'parent' : 'staff';
}

function HR_audienceLabel_(value) {
  return HR_AUDIENCE_LABELS[HR_audienceKey_(value)];
}

// 웹 주소의 ?type= 값을 읽어 'parent' | 'staff' | '' (고르지 않음) 로 바꿉니다.
function HR_audienceFromParameter_(value) {
  var text = String(value || '').trim().toLowerCase();
  if (/^(parent|parents|학부모)$/.test(text)) return 'parent';
  if (/^(staff|teacher|teachers|교직원|교사)$/.test(text)) return 'staff';
  return '';
}

function HR_findAudienceColumn_(headers) {
  return HR_findAnyHeaderColumn_(headers || [], [HR_AUDIENCE_HEADER]);
}

// 「신청구분」 칸이 없으면 맨 오른쪽에 만들고 열 번호를 돌려줍니다.
function HR_ensureAudienceColumn_(sheet) {
  var headers = getHeaders(sheet);
  var column = HR_findAudienceColumn_(headers);
  if (column) return column;

  column = Math.max(sheet.getLastColumn(), headers.length) + 1;
  if (sheet.getMaxColumns() < column) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), column - sheet.getMaxColumns());
  }
  var header = sheet.getRange(HEADER_ROW, column);
  if (column > 1) {
    sheet.getRange(HEADER_ROW, column - 1).copyTo(
      header,
      SpreadsheetApp.CopyPasteType.PASTE_FORMAT,
      false
    );
  }
  header.setValue(HR_AUDIENCE_HEADER);
  sheet.setColumnWidth(column, 90);
  return column;
}

// 학부모 신청자는 「3학년 2반 15번」 모양으로 받습니다. (뒤에 학생 이름이 붙어도 됩니다.)
function HR_parseParentApplicant_(text) {
  var m = String(text || '').match(/(\d{1,2})\s*학년\s*(\d{1,2})\s*반\s*(\d{1,3})\s*번/);
  if (!m) return null;
  return {
    grade: Number(m[1]),
    classNo: Number(m[2]),
    number: Number(m[3]),
    key: Number(m[1]) + '-' + Number(m[2]) + '-' + Number(m[3])
  };
}

// 한도를 세는 기준: 학부모는 학년·반·번호, 교직원은 이름
function HR_applicantKey_(audienceKey, applicant) {
  if (audienceKey === 'parent') {
    var parsed = HR_parseParentApplicant_(applicant);
    return parsed ? parsed.key : '';
  }
  return String(applicant || '').replace(/\s+/g, '').toLowerCase();
}

function HR_getRequestLimit_(audienceKey) {
  var raw = PropertiesService.getScriptProperties().getProperty(HR_LIMIT_PROPERTIES[audienceKey]);
  if (raw === null || String(raw).trim() === '') return HR_DEFAULT_LIMITS[audienceKey] || 0; // 아직 안 정함 → 기본값
  var value = Number(raw);
  return isFinite(value) && value > 0 ? Math.floor(value) : 0; // 0 = 제한 없음(직접 정한 경우)
}

// 지금 목록에서 같은 구분·같은 신청자가 이미 신청한 권수
function HR_countApplicantRequests_(sheet, audienceKey, applicant) {
  var lastRow = sheet.getLastRow();
  if (lastRow < FIRST_DATA_ROW) return 0;

  var headers = getHeaders(sheet);
  var applicantColumn = HR_findHeaderColumn_(headers, '신청자', 11);
  var audienceColumn = HR_findAudienceColumn_(headers);
  var titleColumn = HR_findAnyHeaderColumn_(headers, ['도서명', '책 제목', '책제목', '서명', '제목']);
  var key = HR_applicantKey_(audienceKey, applicant);
  if (!key) return 0;

  var width = Math.max(applicantColumn, audienceColumn, titleColumn);
  var values = sheet
    .getRange(FIRST_DATA_ROW, 1, lastRow - FIRST_DATA_ROW + 1, width)
    .getDisplayValues();
  var count = 0;

  values.forEach(function(row) {
    if (titleColumn && !String(row[titleColumn - 1] || '').trim()) return;
    var rowAudience = audienceColumn ? HR_audienceKey_(row[audienceColumn - 1]) : 'staff';
    if (rowAudience !== audienceKey) return;
    if (HR_applicantKey_(audienceKey, row[applicantColumn - 1]) === key) count++;
  });
  return count;
}

// 한도를 넘으면 안내 글을 돌려주고, 괜찮으면 빈 글자를 돌려줍니다.
function HR_requestLimitMessage_(audienceKey, applicant, addCount) {
  var limit = HR_getRequestLimit_(audienceKey);
  if (!limit) return '';

  // 관리자(사서) 이름으로 신청하면 한도를 적용하지 않습니다.
  var manager = String(
    PropertiesService.getScriptProperties().getProperty(HR_MANAGER_NAME_PROPERTY) || ''
  ).trim();
  if (audienceKey === 'staff' && manager && String(applicant || '').trim().toLowerCase() === manager.toLowerCase()) {
    return '';
  }

  var sheet = HR_getSpreadsheet_().getSheetByName(HR_TARGET_SHEET_NAME);
  if (!sheet) return '';
  var used = HR_countApplicantRequests_(sheet, audienceKey, applicant);
  if (used + addCount <= limit) return '';

  var remain = Math.max(0, limit - used);
  var who = audienceKey === 'parent' ? '한 학생당' : '한 사람당';
  var head = who + ' 최대 ' + limit + '권까지 신청할 수 있어요.';

  if (remain <= 0) {
    return head + ' 이미 ' + used + '권을 모두 신청하셨어요.';
  }
  if (used === 0) {
    return head + ' 지금 ' + addCount + '권을 담으셨어요. ' + limit + '권 이하로 줄여서 다시 신청해 주세요.';
  }
  return head + ' 이미 ' + used + '권을 신청하셨어서 이번에는 ' + remain + '권까지만 가능해요. 담은 ' +
    addCount + '권을 ' + remain + '권 이하로 줄여 주세요.';
}

/* ---------- 공지 대상(전체/교직원/학부모) ---------- */

// 「📢 웹 공지」 탭 맨 오른쪽에 「대상」 칸을 만듭니다. (이미 있으면 그대로 둡니다.)
function HR_ensureNoticeAudienceColumn_(spreadsheet) {
  var sheet = spreadsheet.getSheetByName(HR_WEB_NOTICE_SHEET_NAME);
  if (!sheet) return;
  if (sheet.getMaxColumns() < 6) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), 6 - sheet.getMaxColumns());
  }
  var header = sheet.getRange('F3');
  if (String(header.getDisplayValue() || '').trim() === HR_NOTICE_AUDIENCE_HEADER) return;

  sheet.getRange('E3').copyTo(header, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  header.setValue(HR_NOTICE_AUDIENCE_HEADER);
  sheet.setColumnWidth(6, 90);
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(HR_NOTICE_AUDIENCE_VALUES, true)
    .setAllowInvalid(true)
    .build();
  var rows = Math.max(sheet.getMaxRows() - 3, 1);
  sheet.getRange(4, 6, rows, 1).setDataValidation(rule).setHorizontalAlignment('center');
}

/* ---------- 신청 한도 설정 (메뉴) ---------- */

function HR_showRequestLimitDialog() {
  var ui = SpreadsheetApp.getUi();
  var props = PropertiesService.getScriptProperties();
  var keys = ['staff', 'parent'];
  var result = [];

  for (var i = 0; i < keys.length; i++) {
    var key = keys[i];
    var current = HR_getRequestLimit_(key);
    var answer = ui.prompt(
      '📏 신청 한도 설정 · ' + HR_audienceLabel_(key),
      '1인당 신청(구입)할 수 있는 권수를 숫자로 입력하세요.\n' +
      (key === 'parent' ? '(학부모는 한 학생(학년·반·번호)당입니다.)\n' : '') +
      '0을 넣으면 제한 없음입니다.\n\n지금 설정: ' + (current ? current + '권' : '제한 없음') +
      (props.getProperty(HR_LIMIT_PROPERTIES[key]) === null && HR_DEFAULT_LIMITS[key] ? ' (기본값)' : ''),
      ui.ButtonSet.OK_CANCEL
    );
    if (answer.getSelectedButton() !== ui.Button.OK) return;

    var text = String(answer.getResponseText() || '').trim();
    if (text === '') { result.push(HR_audienceLabel_(key) + ': 그대로 둠'); continue; }
    var value = Number(text);
    if (!isFinite(value) || value < 0 || value > 500) {
      ui.alert('숫자를 0~500 사이로 입력해 주세요.');
      return;
    }
    value = Math.floor(value);
    props.setProperty(HR_LIMIT_PROPERTIES[key], String(value));
    result.push(HR_audienceLabel_(key) + ': ' + (value ? value + '권' : '제한 없음'));
  }

  ui.alert('📏 신청 한도를 저장했어요', result.join('\n'), ui.ButtonSet.OK);
}

// 시트를 열 때 한 번: 필요한 칸이 없으면 만들어 둡니다. (실패해도 조용히 넘어갑니다.)
function HR_prepareAudienceColumns_() {
  var spreadsheet = HR_getSpreadsheet_();
  var list = spreadsheet.getSheetByName(HR_TARGET_SHEET_NAME);
  if (list) HR_ensureAudienceColumn_(list);
  HR_ensureNoticeAudienceColumn_(spreadsheet);
}

/* ---------- 학부모 개인정보 파기 ---------- */

// 학부모가 입력한 학년·반·번호·학생 이름은 구입이 끝나면 함께 파기합니다.
// 「3학년 2반 15번 …」 모양의 신청자 글만 찾아서 「(파기됨)」으로 바꾸고,
// 책 신청 내용과 교직원 이름은 그대로 둡니다.
var HR_SCRUBBED_TEXT = '(파기됨)';

// 한 시트의 한 열에서 학부모 신청자 글을 파기하고 바꾼 칸 수를 돌려줍니다.
function HR_scrubApplicantColumn_(sheet, column, firstRow) {
  if (!sheet || column < 1) return 0;
  var lastRow = sheet.getLastRow();
  if (lastRow < firstRow || sheet.getLastColumn() < column) return 0;

  var range = sheet.getRange(firstRow, column, lastRow - firstRow + 1, 1);
  var values = range.getValues();
  var changed = 0;
  var out = values.map(function(row) {
    if (HR_parseParentApplicant_(row[0])) {
      changed++;
      return [HR_SCRUBBED_TEXT];
    }
    return [row[0]];   // 다른 칸은 원래 값 그대로
  });
  if (changed) range.setValues(out);
  return changed;
}

// 신청 목록·접수 백업·취소 백업·숨김 보관 시트에서 학부모 개인정보를 모두 파기합니다.
function HR_scrubParentApplicants_() {
  var spreadsheet = HR_getSpreadsheet_();
  var result = { list: 0, backup: 0, cancelBackup: 0, archive: 0, total: 0 };

  var list = spreadsheet.getSheetByName(HR_TARGET_SHEET_NAME);
  var listApplicantColumn = 0;
  if (list) {
    listApplicantColumn = HR_findHeaderColumn_(getHeaders(list), '신청자', 11);
    result.list = HR_scrubApplicantColumn_(list, listApplicantColumn, FIRST_DATA_ROW);
  }

  // 접수 백업: 3번째 칸이 신청자
  result.backup = HR_scrubApplicantColumn_(
    spreadsheet.getSheetByName(HR_REQUEST_BACKUP_SHEET_NAME), 3, 2
  );

  // 취소 백업: 8칸의 관리 정보 뒤에 원래 목록의 줄이 그대로 복사되어 있음
  if (listApplicantColumn) {
    result.cancelBackup = HR_scrubApplicantColumn_(
      spreadsheet.getSheetByName(HR_REQUEST_CANCEL_BACKUP_SHEET_NAME), 8 + listApplicantColumn, 2
    );
  }

  // 「새 신청 시작」 때 목록을 통째로 보관한 숨김 시트들(🗂 희망도서 …)
  spreadsheet.getSheets().forEach(function(sheet) {
    if (String(sheet.getName()).indexOf('🗂 희망도서') !== 0) return;
    var column = HR_findHeaderColumn_(getHeaders(sheet), '신청자', 11);
    result.archive += HR_scrubApplicantColumn_(sheet, column, FIRST_DATA_ROW);
  });

  result.total = result.list + result.backup + result.cancelBackup + result.archive;
  return result;
}

// 메뉴: 「🧹 학부모 개인정보 파기」
function HR_destroyParentInfo() {
  var ui = SpreadsheetApp.getUi();
  var answer = ui.alert(
    '🧹 학부모 개인정보 파기',
    '학부모가 입력한 학년·반·번호·학생 이름을 모두 지웁니다.\n' +
    '(신청 목록, 숨김 백업, 새 신청 시작 때 만든 보관 시트)\n\n' +
    '책 신청 내용은 그대로 두고, 신청자 칸만 「(파기됨)」으로 바뀝니다. 교직원 이름은 지우지 않습니다.\n\n' +
    '⚠️ 도서 구입이 끝난 뒤에만 사용하세요. 지운 정보는 되돌릴 수 없고, 학생별 신청 한도도 0권부터 다시 셉니다.\n\n' +
    '계속할까요?',
    ui.ButtonSet.YES_NO
  );
  if (answer !== ui.Button.YES) return;

  var result = HR_scrubParentApplicants_();
  SpreadsheetApp.flush();
  ui.alert(
    '✅ 파기를 마쳤습니다',
    result.total
      ? '학부모 개인정보 ' + result.total + '건을 파기했습니다.\n' +
        '• 신청 목록 ' + result.list + '건\n• 접수 백업 ' + result.backup + '건\n' +
        '• 취소 백업 ' + result.cancelBackup + '건\n• 보관 시트 ' + result.archive + '건'
      : '파기할 학부모 개인정보가 없습니다. (이미 파기했거나 학부모 신청이 없습니다.)',
    ui.ButtonSet.OK
  );
}
