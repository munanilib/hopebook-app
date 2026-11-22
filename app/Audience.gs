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
  return '신청 한도는 1인당 ' + limit + '권이에요. 지금까지 ' + used + '권을 신청했어요.' +
    (remain > 0 ? ' 이번에는 최대 ' + remain + '권까지만 신청할 수 있어요.' : ' 더 신청할 수 없어요.');
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
