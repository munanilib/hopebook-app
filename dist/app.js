/* 자동 생성 파일 — 직접 고치지 마세요. (node tools/build.js)  vreal-1 */
var APP_FACTORY = function (__host) {
  var HtmlService = __host.html;
  var SpreadsheetApp = __host.spreadsheet;

/* ===== ApplyLinks.gs ===== */
/**
 * 🔗 신청 링크 받기
 *
 * 교직원용·학부모용 신청 링크를 한 창에서 보여 주고, 복사하거나 QR 코드를 그림(PNG)으로 저장하게 합니다.
 * QR 은 창 안에서 직접 만들기 때문에 학교 주소가 외부 서비스로 나가지 않습니다.
 */
function HR_showApplyLinksDialog() {
  var props = PropertiesService.getScriptProperties();
  var detected = HR_getWebAppBaseUrl_();
  var saved = String(props.getProperty(HR_PUBLIC_WEB_APP_URL_PROPERTY) || '').trim();

  // 배포된 웹 주소를 자동으로 찾았으면 「공개 링크」로도 저장해 둡니다.
  if (detected && detected !== saved) {
    props.setProperty(HR_PUBLIC_WEB_APP_URL_PROPERTY, detected);
    saved = detected;
  }

  var baseUrl = detected || (/^https:\/\/script\.google\.com\/.+\/exec$/.test(saved) ? saved : '');
  var data = {
    baseUrl: baseUrl,
    detected: !!detected,
    schoolName: String(props.getProperty(HR_SCHOOL_NAME_PROPERTY) || '').trim()
  };

  var template = HtmlService.createTemplateFromFile('ApplyLinks');
  template.dataJson = JSON.stringify(data).replace(/</g, '\\u003c');
  SpreadsheetApp.getUi().showModalDialog(
    template.evaluate().setWidth(780).setHeight(860),
    '🔗 신청 링크 받기'
  );
}


/* ===== Audience.gs ===== */
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
  var value = Number(raw || 0);
  return isFinite(value) && value > 0 ? Math.floor(value) : 0; // 0 = 제한 없음
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
      '1인당 신청할 수 있는 권수를 숫자로 입력하세요.\n' +
      (key === 'parent' ? '(학부모는 한 학생(학년·반·번호)당입니다.)\n' : '') +
      '0을 넣으면 제한 없음입니다.\n\n지금 설정: ' + (current ? current + '권' : '제한 없음'),
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


/* ===== Code.js ===== */
/**
 * ============================================================
 * 희망도서 자동 검색·중복 확인 시스템
 *
 * 검색: YES24 API
 * 분류기호: 국립중앙도서관 API
 * 네이버 API: 사용하지 않음
 * ============================================================
 */

// API 키
const YES24_KEY = '';
const NLK_KEY = '';  // 키는 스크립트 속성(HR_NLK_API_KEY)에 보관합니다. 코드에 적지 마세요.

// 시트 이름
const TARGET_SHEET_NAME = '📚 희망도서 목록';
const LIBRARY_SHEET_NAME = '🏫 우리학교 소장도서';
const BOOKSTORE_EXPORT_SHEET_NAME = '📦 서점 전달용';

// 설정
const HEADER_ROW = 5;
const FIRST_DATA_ROW = 6;
const MAX_EDIT_ROWS = 30;
const API_SLEEP_MS = 300;
const CACHE_SECONDS = 21600;

function getConfiguredYes24Key_() {
  const savedKey = String(
    PropertiesService
      .getScriptProperties()
      .getProperty('HR_YES24_API_KEY') || ''
  ).trim();

  return savedKey || String(YES24_KEY || '').trim();
}

function getConfiguredNlkKey() {
  const savedKey = String(
    PropertiesService
      .getScriptProperties()
      .getProperty('HR_NLK_API_KEY') || ''
  ).trim();

  if (savedKey) return savedKey;

  // 코드에 적혀 있던 기본 키는 처음 쓸 때 스크립트 속성으로 옮깁니다.
  const legacyKey = String(NLK_KEY || '').trim();
  if (legacyKey) {
    PropertiesService.getScriptProperties().setProperty('HR_NLK_API_KEY', legacyKey);
  }
  return legacyKey;
}


/**
 * ============================================================
 * 1. 메뉴 생성
 * ============================================================
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📚 도서 관리')
    .addItem('🔍 기존 데이터 일괄 채우기', 'checkExistingBooks')
        .addItem('🔄 선택한 행 다시 조회', 'reprocessSelectedRows')
    .addItem('🛒 구입 검토함에 접수 도서 추가', 'sendHopeRequestsToReview')
    .addSeparator()
    .addItem('📦 서점 전달용 시트 새로고침', 'buildBookstoreExportSheet')
    .addItem('🎨 희망도서 목록 디자인 다시 적용', 'styleRequestListSheet')
    .addSeparator()
    .addItem('⚙️ 실시간 자동화 권한 승인', 'installEditTrigger')
    .addSeparator()
    .addItem('🔐 관리자 · 새 희망도서 신청 시작', 'startNewRequestRound')
    .addToUi();

  SpreadsheetApp.getUi()
    .createMenu('⭐ 희망도서 신청')
    .addItem('🔗 신청 페이지 열기', 'HR_showHopeRequestDialog')
    .addItem('🧹 새 학기 초기화', 'resetHopeBooksForNewSemester')
    .addToUi();

  




  if (typeof HR_addCopySetupMenu === 'function') {
    HR_addCopySetupMenu();
  }

  if (typeof unifiedAddMenu_ === 'function') {
    unifiedAddMenu_();
  }
  try {
    HR_removeLegacyOpenTrigger_();
  } catch (error) {
    console.warn('옛 열기 트리거 정리 실패: ' + error);
  }
  try {
    HR_prepareAudienceColumns_();
  } catch (error) {
    console.warn('신청구분 칸 준비 실패: ' + error);
  }
  try {
    HR_showUpdateNoticeOnOpen_();
  } catch (error) {
    console.warn('업데이트 소식 표시 실패: ' + error);
  }
  try {
    ensureSelectionControls_();
  } catch (error) {
    console.warn('선택 버튼 표시 실패: ' + error);
  }
}


/**
 * 기존 희망도서 신청 목록을 숨김 보관하고 새 신청 회차를 시작합니다.
 * 구입 검토함, 서점 전달용, 소장도서와 설정은 변경하지 않습니다.
 */
function startNewRequestRound() {
  const ui = SpreadsheetApp.getUi();
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(TARGET_SHEET_NAME);

  if (!sheet) {
    ui.alert('「' + TARGET_SHEET_NAME + '」 시트를 찾지 못했습니다.');
    return;
  }

  const lastRow = sheet.getLastRow();
  const lastColumn = Math.max(sheet.getLastColumn(), 1);
  const headers = sheet
    .getRange(HEADER_ROW, 1, 1, lastColumn)
    .getDisplayValues()[0]
    .map(function(value) {
      return String(value || '').trim();
    });
  let titleIndex = headers.findIndex(function(header) {
    return /^(도서명|책 ?제목|제목)$/.test(header);
  });

  if (titleIndex < 0) {
    titleIndex = Math.min(1, lastColumn - 1);
  }

  const dataValues = lastRow >= FIRST_DATA_ROW
    ? sheet
        .getRange(
          FIRST_DATA_ROW,
          1,
          lastRow - FIRST_DATA_ROW + 1,
          lastColumn
        )
        .getDisplayValues()
    : [];
  const requestCount = dataValues.filter(function(row) {
    return String(row[titleIndex] || '').trim() !== '';
  }).length;
  const response = ui.alert(
    '🆕 새 희망도서 신청 시작',
    '현재 신청 목록 ' + requestCount +
      '건을 숨김 보관한 뒤 새 신청 목록으로 비웁니다.\n\n' +
      '구입 검토함·서점 전달용·우리학교 소장도서·설정은 그대로 유지됩니다.\n' +
      '계속할까요?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    return;
  }

  const archiveName = buildRequestArchiveSheetName_(spreadsheet);
  const archiveSheet = sheet.copyTo(spreadsheet);
  archiveSheet.setName(archiveName);
  archiveSheet.hideSheet();

  if (lastRow >= FIRST_DATA_ROW) {
    sheet
      .getRange(
        FIRST_DATA_ROW,
        1,
        lastRow - FIRST_DATA_ROW + 1,
        lastColumn
      )
      .clearContent();
  }

  if (typeof ensureSelectionControls_ === 'function') {
    ensureSelectionControls_();
  }

  spreadsheet.setActiveSheet(sheet);
  sheet.getRange('A1').activate();
  SpreadsheetApp.flush();

  ui.alert(
    '새 신청 회차를 시작했습니다.',
    '이전 목록 ' + requestCount + '건은 「' + archiveName +
      '」에 숨김 보관했습니다.\n' +
      '필요하면 보기 → 숨겨진 시트에서 다시 확인할 수 있습니다.',
    ui.ButtonSet.OK
  );
}

function buildRequestArchiveSheetName_(spreadsheet) {
  const timeZone =
    spreadsheet.getSpreadsheetTimeZone() || Session.getScriptTimeZone();
  const baseName =
    '🗂 희망도서 ' +
    Utilities.formatDate(new Date(), timeZone, 'yyyy.MM.dd HHmmss');
  let name = baseName;
  let suffix = 2;

  while (spreadsheet.getSheetByName(name)) {
    name = baseName + ' (' + suffix + ')';
    suffix += 1;
  }

  return name.slice(0, 100);
}




/**
 * 서점에 전달할 주문 목록을 개인정보 없이 정리합니다.
 * 취소·오류·예시 행은 제외하고 같은 책은 수량으로 합칩니다.
 */
function styleRequestListSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(TARGET_SHEET_NAME);

  if (!sheet) {
    throw new Error('「' + TARGET_SHEET_NAME + '」 시트를 찾을 수 없습니다.');
  }

  const headers = getHeaders(sheet);
  const lastColumn = Math.max(sheet.getLastColumn(), headers.length, 12);
  const lastRow = Math.max(sheet.getLastRow(), HEADER_ROW);
  const bodyRowCount = Math.max(1, Math.min(sheet.getMaxRows() - FIRST_DATA_ROW + 1, 2000));
  const dataRowCount = Math.max(0, lastRow - FIRST_DATA_ROW + 1);
  const titleColumn = bookstoreFindColumn_(headers, ['도서명', '책제목', '서명', '제목']);
  const isbnColumn = bookstoreFindColumn_(headers, ['ISBN13', 'ISBN']);
  const coverColumn = bookstoreFindColumn_(headers, ['표지', '책표지']);
  const priceColumn = bookstoreFindColumn_(headers, ['정가', '도서정가', '정가원', '가격']);
  const yearColumn = bookstoreFindColumn_(headers, ['출판연도', '발행연도', '출간연도', '연도']);
  const applicantColumn = bookstoreFindColumn_(headers, ['신청자', '신청인']);
  const statusColumn = bookstoreFindColumn_(headers, ['처리상태', '자동입력상태', '상태']);
  const widths = [64,250,150,95,145,190,105,85,220,115,120,110];

  sheet.setHiddenGridlines(true);
  sheet.setFrozenRows(HEADER_ROW);
  sheet.setFrozenColumns(0);

  sheet.getRange(1, 1, 4, lastColumn).breakApart();
  sheet.getRange(1, 1, 4, lastColumn).clearDataValidations();
  sheet.getRange(4, 1, 1, lastColumn).breakApart().clearContent().clearDataValidations();

  sheet.getRange(1, 1, 1, lastColumn)
    .merge()
    .setValue('📚 희망도서 신청 목록')
    .setBackground('#183D35')
    .setFontColor('#FFFFFF')
    .setFontFamily('Arial')
    .setFontSize(23)
    .setFontWeight('bold')
    .setHorizontalAlignment('left')
    .setVerticalAlignment('middle');

  sheet.getRange(2, 1, 1, lastColumn)
    .merge()
    .setValue('웹페이지에서 신청된 도서를 확인하고 구입 검토함으로 모아보세요.')
    .setBackground('#EAF3ED')
    .setFontColor('#34564D')
    .setFontFamily('Arial')
    .setFontSize(12)
    .setVerticalAlignment('middle');

  sheet.getRange(3, 1, 1, lastColumn)
    .merge()
    .setValue('접수된 도서는 중복·취소 상태를 제외한 뒤 구입 검토함에서 최종 확인합니다.')
    .setBackground('#FFF4D6')
    .setFontColor('#765413')
    .setFontFamily('Arial')
    .setFontSize(11)
    .setFontWeight('bold')
    .setVerticalAlignment('middle');

  sheet.getRange(4, 1)
    .insertCheckboxes()
    .setValue(false)
    .setBackground('#D9A441')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  sheet.getRange(4, 2, 1, lastColumn - 1)
    .merge()
    .setValue('접수된 희망도서를 구입 검토함으로 보내기')
    .setBackground('#D9A441')
    .setFontColor('#FFFFFF')
    .setFontFamily('Arial')
    .setFontSize(14)
    .setFontWeight('bold')
    .setHorizontalAlignment('left')
    .setVerticalAlignment('middle');

  sheet.getRange(HEADER_ROW, 1, 1, lastColumn)
    .setBackground('#2E5B4F')
    .setFontColor('#FFFFFF')
    .setFontFamily('Arial')
    .setFontSize(13)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrap(true);

  sheet.getRange(FIRST_DATA_ROW, 1, bodyRowCount, lastColumn)
    .setBackground('#FBFCFB')
    .setFontColor('#2E3A33')
    .setFontFamily('Arial')
    .setFontSize(12)
    .setVerticalAlignment('middle')
    .setWrap(true);

  if (dataRowCount) {
    sheet.getRange(FIRST_DATA_ROW, 1, dataRowCount, lastColumn).setBorder(
      false,false,true,false,false,false,
      '#E1E7E2',SpreadsheetApp.BorderStyle.SOLID
    );
    sheet.setRowHeights(FIRST_DATA_ROW, dataRowCount, 96);
  }

  sheet.setRowHeight(1, 58);
  sheet.setRowHeight(2, 34);
  sheet.setRowHeight(3, 34);
  sheet.setRowHeight(4, 44);
  sheet.setRowHeight(HEADER_ROW, 42);

  sheet.getRange(FIRST_DATA_ROW, 1, bodyRowCount, 1)
    .setHorizontalAlignment('center')
    .setFontWeight('bold')
    .setFontColor('#52635A');

  if (titleColumn) sheet.getRange(FIRST_DATA_ROW, titleColumn, bodyRowCount, 1).setFontWeight('bold').setFontColor('#22372E');
  if (isbnColumn) sheet.getRange(FIRST_DATA_ROW, isbnColumn, bodyRowCount, 1).setNumberFormat('@').setHorizontalAlignment('center');
  if (coverColumn) sheet.getRange(FIRST_DATA_ROW, coverColumn, bodyRowCount, 1).setHorizontalAlignment('center');
  if (priceColumn) sheet.getRange(FIRST_DATA_ROW, priceColumn, bodyRowCount, 1).setNumberFormat('#,##0"원"').setHorizontalAlignment('right');
  if (yearColumn) sheet.getRange(FIRST_DATA_ROW, yearColumn, bodyRowCount, 1).setNumberFormat('0').setHorizontalAlignment('center');
  if (applicantColumn) sheet.getRange(FIRST_DATA_ROW, applicantColumn, bodyRowCount, 1).setBackground('#FFF8E8').setHorizontalAlignment('center');
  if (statusColumn) sheet.getRange(FIRST_DATA_ROW, statusColumn, bodyRowCount, 1).setBackground('#EDF5EF').setFontColor('#315B43').setFontWeight('bold').setHorizontalAlignment('center');

  for (let column = 1; column <= lastColumn; column++) {
    sheet.setColumnWidth(column, widths[column - 1] || 130);
  }

  sheet.getRange(1, 1, 4, lastColumn).setBorder(true,true,true,true,false,false,'#D5DED8',SpreadsheetApp.BorderStyle.SOLID);
}

function buildBookstoreExportSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const source = spreadsheet.getSheetByName(TARGET_SHEET_NAME);

  if (!source) {
    throw new Error('「' + TARGET_SHEET_NAME + '」 시트를 찾을 수 없습니다.');
  }

  const exportTitle = '📦 서점 전달용 도서 주문 목록';
  let target = spreadsheet.getSheetByName(BOOKSTORE_EXPORT_SHEET_NAME);

  if (!target) {
    target = spreadsheet.insertSheet(BOOKSTORE_EXPORT_SHEET_NAME);
  } else {
    const currentTitle = String(target.getRange('A1').getDisplayValue() || '').trim();
    if (
      target.getLastRow() > 1 &&
      currentTitle &&
      currentTitle.indexOf('서점 전달용') === -1
    ) {
      throw new Error(
        '「' + BOOKSTORE_EXPORT_SHEET_NAME +
        '」 시트에 다른 내용이 있어 덮어쓰지 않았습니다.'
      );
    }
  }

  if (target.getFilter()) {
    target.getFilter().remove();
  }
  target.getBandings().forEach(banding => banding.remove());
  target.getRange('A1:J5').breakApart();
  target.clear();
  target.clearConditionalFormatRules();
  target.setHiddenGridlines(true);

  const sourceHeaders = getHeaders(source);
  const columns = {
    sequence: bookstoreFindColumn_(sourceHeaders, ['순번', '번호', 'NO']),
    title: bookstoreFindColumn_(sourceHeaders, ['도서명', '책제목', '서명', '제목']),
    author: bookstoreFindColumn_(sourceHeaders, ['저자', '작가', '글쓴이']),
    publisher: bookstoreFindColumn_(sourceHeaders, ['출판사']),
    year: bookstoreFindColumn_(sourceHeaders, ['출판연도', '발행연도', '출간연도', '연도']),
    isbn: bookstoreFindColumn_(sourceHeaders, ['ISBN13', 'ISBN']),
    price: bookstoreFindColumn_(sourceHeaders, ['정가', '도서정가', '정가원', '가격']),
    status: bookstoreFindColumn_(sourceHeaders, ['처리상태', '자동입력상태', '상태'])
  };

  if (!columns.title && !columns.isbn) {
    throw new Error('희망도서 목록에서 도서명과 ISBN 열을 찾지 못했습니다.');
  }

  const lastRow = source.getLastRow();
  const lastColumn = source.getLastColumn();
  const displayRows = lastRow >= 2
    ? source.getRange(2, 1, lastRow - 1, lastColumn).getDisplayValues()
    : [];
  const rawRows = lastRow >= 2
    ? source.getRange(2, 1, lastRow - 1, lastColumn).getValues()
    : [];
  const grouped = {};
  const order = [];

  displayRows.forEach((displayRow, index) => {
    const rawRow = rawRows[index];
    const sequenceText = columns.sequence
      ? String(displayRow[columns.sequence - 1] || '').trim()
      : String(index + 1);
    const sequenceNumber = Number(sequenceText);
    const title = columns.title
      ? String(displayRow[columns.title - 1] || '').trim()
      : '';
    const isbn = columns.isbn
      ? String(displayRow[columns.isbn - 1] || '').replace(/[^0-9Xx]/g, '')
      : '';
    const author = columns.author
      ? String(displayRow[columns.author - 1] || '').trim()
      : '';
    const publisher = columns.publisher
      ? String(displayRow[columns.publisher - 1] || '').trim()
      : '';
    const year = columns.year
      ? String(displayRow[columns.year - 1] || '').match(/\d{4}/)
      : null;
    const status = columns.status
      ? String(displayRow[columns.status - 1] || '').trim()
      : '';
    let price = columns.price
      ? Number(rawRow[columns.price - 1] || 0)
      : 0;

    if (!price && columns.price) {
      price = Number(
        String(displayRow[columns.price - 1] || '').replace(/[^0-9]/g, '')
      ) || 0;
    }

    if (!isFinite(sequenceNumber) || sequenceNumber <= 0) return;
    if (!title && !isbn) return;
    const isManagerOwnedException = /관리자\s*소장도서\s*예외/.test(status);
    if (/취소|오류|중복/.test(status) || (/소장/.test(status) && !isManagerOwnedException)) return;

    const key = isbn
      ? 'ISBN:' + isbn
      : 'TITLE:' + bookstoreNormalize_(title + '|' + publisher);
    const note = isbn ? '' : 'ISBN 확인 필요';

    if (grouped[key]) {
      grouped[key].quantity += 1;
      grouped[key].amount = grouped[key].price * grouped[key].quantity;
      return;
    }

    grouped[key] = {
      sourceSequence: sequenceNumber,
      title: title,
      author: author,
      publisher: publisher,
      year: year ? year[0] : '',
      isbn: isbn,
      price: price,
      quantity: 1,
      amount: price,
      note: note
    };
    order.push(key);
  });

  const books = order
    .map(key => grouped[key])
    .sort((a, b) => a.sourceSequence - b.sourceSequence);
  const output = books.map((book, index) => [
    index + 1,
    book.title,
    book.author,
    book.publisher,
    book.year,
    book.isbn,
    book.price || '',
    book.quantity,
    book.amount || '',
    book.note
  ]);
  const totalPrice = books.reduce((sum, book) => sum + Number(book.amount || 0), 0);
  const schoolName = String(
    PropertiesService.getScriptProperties().getProperty('HR_SCHOOL_NAME') || ''
  ).trim();
  const madeAt = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone() || 'Asia/Seoul',
    'yyyy. M. d. HH:mm'
  );
  const headers = [
    '순번', '도서명', '저자', '출판사', '출판연도',
    'ISBN13', '정가', '수량', '금액', '비고'
  ];

  target.getRange('A1:J1').merge().setValue(
    '📦 ' + (schoolName ? schoolName + ' ' : '') + '서점 전달용 도서 주문 목록'
  );
  target.getRange('A2:J2').merge().setValue(
    '희망도서 목록에서 취소·오류 도서와 신청자 개인정보를 제외해 정리했습니다.'
  );
  target.getRange('A3:J3').merge().setValue(
    '총 ' + books.length + '종  ·  주문 ' +
    books.reduce((sum, book) => sum + book.quantity, 0) +
    '권  ·  정가 합계 ' +
    Number(totalPrice || 0).toLocaleString('ko-KR') + '원' +
    '  ·  만든 시각 ' + madeAt
  );

  target.getRange('A4')
    .insertCheckboxes()
    .setValue(false)
    .setBackground('#D8A646')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  target.getRange('B4:J4')
    .merge()
    .setValue('최신 신청 목록으로 다시 만들기')
    .setBackground('#D8A646')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold')
    .setFontSize(13)
    .setHorizontalAlignment('left')
    .setVerticalAlignment('middle');

  target.getRange(5, 1, 1, headers.length).setValues([headers]);

  if (output.length) {
    target.getRange(6, 1, output.length, headers.length).setValues(output);
    target.getRange(6, 9, output.length, 1).setFormulaR1C1(
      '=IF(OR(RC[-2]="",RC[-1]=""),"",RC[-2]*RC[-1])'
    );
    target.getRange(6, 6, output.length, 1).setNumberFormat('@');
    target.getRange(6, 7, output.length, 1).setNumberFormat('#,##0"원"');
    target.getRange(6, 8, output.length, 1).setNumberFormat('0');
    target.getRange(6, 9, output.length, 1).setNumberFormat('#,##0"원"');

    target.getRange(5, 1, output.length + 1, headers.length).createFilter();
    target.getRange(6, 1, output.length, headers.length)
      .applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY);

    const totalRow = output.length + 7;
    target.getRange(totalRow, 1, 1, 8).merge().setValue('정가 합계');
    target.getRange(totalRow, 9)
      .setFormula('=SUM(I6:I' + (output.length + 5) + ')')
      .setNumberFormat('#,##0"원"');
    target.getRange(totalRow, 10).setValue('서점 할인·공급가는 별도 협의');
    target.getRange(totalRow, 1, 1, 10)
      .setBackground('#E8EFE9')
      .setFontWeight('bold')
      .setFontColor('#263B31');
  } else {
    target.getRange('A6:J6').merge().setValue(
      '서점에 전달할 신청 도서가 아직 없습니다.'
    );
  }

  bookstoreStyleExportSheet_(target, output.length);
  spreadsheet.setActiveSheet(target);
  target.activate();
  target.setActiveRange(target.getRange('A1'));

  spreadsheet.toast(
    '서점 전달용 목록 ' + books.length + '종을 정리했습니다.',
    '📦 서점 전달용',
    7
  );
}

function onEdit(e) {
  if (!e || !e.range) return;

  const range = e.range;
  const sheet = range.getSheet();
  const sheetName = sheet.getName();
  const valueText = String(e.value || '').toUpperCase();
  const checked = valueText === 'TRUE';
  const checkboxChanged = checked || valueText === 'FALSE';
  const cell = range.getA1Notation();

  // 시리즈 누락 점검: 목록 전체 선택/전체 해제
  if (sheetName === SERIES_FINDER_SHEET && cell === 'A4' && checkboxChanged) {
    toggleSelectionColumn_(sheet, 6, 1, checked);
    return;
  }

  // 구입 검토함: 목록 전체 선택/전체 해제
  if (sheetName === UNIFIED_REVIEW_SHEET && cell === 'A4' && checkboxChanged) {
    toggleSelectionColumn_(sheet, 6, 1, checked);
    return;
  }

  // 파일 업로드 목록: 목록 전체 선택/전체 해제
  if (sheetName === UNIFIED_UPLOAD_SHEET && cell === 'A4' && checkboxChanged) {
    toggleSelectionColumn_(sheet, 6, 1, checked);
    return;
  }

  if (!checked) return;

  if (sheetName === TARGET_SHEET_NAME && cell === 'A4') {
    range.setValue(false);
    sendHopeRequestsToReview();
    return;
  }

  if (sheetName === SERIES_FINDER_SHEET && cell === 'D4') {
    range.setValue(false);
    sendSelectedSeriesToReview();
    return;
  }

  if (sheetName === UNIFIED_UPLOAD_SHEET && cell === 'D4') {
    range.setValue(false);
    sendSelectedUploadItemsToReview();
    return;
  }

  if (sheetName === UNIFIED_REVIEW_SHEET && cell === 'D4') {
    range.setValue(false);
    buildUnifiedBookstoreExportSheet();
    return;
  }

  if (sheetName === BOOKSTORE_EXPORT_SHEET_NAME && cell === 'A4') {
    range.setValue(false);
    buildUnifiedBookstoreExportSheet();
  }
}

function bookstoreStyleExportSheet_(sheet, rowCount) {
  sheet.setFrozenRows(5);
  sheet.setTabColor('#D8A646');

  sheet.setRowHeight(1, 62);
  sheet.setRowHeight(2, 38);
  sheet.setRowHeight(3, 34);
  sheet.setRowHeight(4, 42);
  sheet.setRowHeight(5, 42);

  sheet.getRange('A1:J1')
    .setBackground('#263B31')
    .setFontColor('#FFFFFF')
    .setFontSize(22)
    .setFontWeight('bold')
    .setHorizontalAlignment('left')
    .setVerticalAlignment('middle');
  sheet.getRange('A2:J2')
    .setBackground('#F4EFE5')
    .setFontColor('#58645E')
    .setFontSize(11)
    .setHorizontalAlignment('left')
    .setVerticalAlignment('middle');
  sheet.getRange('A3:J3')
    .setBackground('#EEF3EF')
    .setFontColor('#263B31')
    .setFontWeight('bold')
    .setHorizontalAlignment('left')
    .setVerticalAlignment('middle');
  sheet.getRange('A5:J5')
    .setBackground('#4E6658')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold')
    .setFontSize(11)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  const widths = [60, 280, 170, 150, 90, 145, 95, 65, 105, 210];
  widths.forEach((width, index) => sheet.setColumnWidth(index + 1, width));

  if (rowCount > 0) {
    sheet.getRange(6, 1, rowCount, 10)
      .setVerticalAlignment('middle')
      .setWrap(true)
      .setFontSize(10);
    sheet.getRange(6, 1, rowCount, 1).setHorizontalAlignment('center');
    sheet.getRange(6, 5, rowCount, 5).setHorizontalAlignment('center');

    for (let row = 6; row < 6 + rowCount; row++) {
      sheet.setRowHeight(row, 58);
    }
  }

  sheet.getRange('A1:J5').setFontFamily('Arial');
}

function bookstoreFindColumn_(headers, aliases) {
  const normalizedAliases = aliases.map(alias => bookstoreNormalize_(alias));

  for (let index = 0; index < headers.length; index++) {
    const header = bookstoreNormalize_(headers[index]);
    if (!header) continue;

    for (let aliasIndex = 0; aliasIndex < normalizedAliases.length; aliasIndex++) {
      const alias = normalizedAliases[aliasIndex];
      if (header === alias || header.indexOf(alias) !== -1) {
        return index + 1;
      }
    }
  }

  return 0;
}

function bookstoreNormalize_(value) {
  return String(value || '')
    .replace(/[\s\n\r()（）·._-]/g, '')
    .toUpperCase();
}

/**
 * ============================================================
 * 2. 설치형 편집 트리거 생성
 * ============================================================
 */
function installEditTrigger() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (
      trigger.getHandlerFunction() === 'onEditTriggerAction' &&
      trigger.getEventType() === ScriptApp.EventType.ON_EDIT
    ) {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger('onEditTriggerAction')
    .forSpreadsheet(spreadsheet)
    .onEdit()
    .create();

  SpreadsheetApp.getUi().alert(
    '✅ 실시간 자동 완성 권한이 설정되었습니다.\n\n' +
    '이제 도서명 또는 ISBN을 입력하면 자동으로 검색됩니다.'
  );
}


/**
 * ============================================================
 * 3. 실시간 편집 감지
 * ============================================================
 */
function onEditTriggerAction(e) {
  if (!e || !e.range) return;

  const range = e.range;
  const sheet = range.getSheet();

  if (sheet.getName() !== TARGET_SHEET_NAME) return;
  if (range.getLastRow() < FIRST_DATA_ROW) return;

  const headers = getHeaders(sheet);
  const columns = getColumnMap(headers);

  if (!columns.title && !columns.isbn) return;

  const editedStartColumn = range.getColumn();
  const editedEndColumn = range.getLastColumn();

  const titleEdited =
    columns.title > 0 &&
    columns.title >= editedStartColumn &&
    columns.title <= editedEndColumn;

  const isbnEdited =
    columns.isbn > 0 &&
    columns.isbn >= editedStartColumn &&
    columns.isbn <= editedEndColumn;

  if (!titleEdited && !isbnEdited) return;

  const startRow = Math.max(range.getRow(), FIRST_DATA_ROW);

  const rowCount = Math.min(
    range.getLastRow() - startRow + 1,
    MAX_EDIT_ROWS
  );

  if (rowCount <= 0) return;

  const lock = LockService.getDocumentLock();

  try {
    lock.waitLock(30000);

    for (let offset = 0; offset < rowCount; offset++) {
      const row = startRow + offset;

      processRow(sheet, row, false, headers, columns);

      if (rowCount > 1) {
        Utilities.sleep(API_SLEEP_MS);
      }
    }
  } catch (error) {
    console.error('실시간 처리 오류', error);
  } finally {
    try {
      lock.releaseLock();
    } catch (error) {
      // 잠금을 얻지 못한 경우 무시
    }
  }
}


/**
 * ============================================================
 * 4. 한 행 처리
 * ============================================================
 */
function processRow(sheet, row, force, headers, columns) {
  headers = headers || getHeaders(sheet);
  columns = columns || getColumnMap(headers);

  const rawTitle = columns.title
    ? String(sheet.getRange(row, columns.title).getDisplayValue()).trim()
    : '';

  const rawIsbn = columns.isbn
    ? String(sheet.getRange(row, columns.isbn).getDisplayValue()).trim()
    : '';

  const currentStatus = columns.status
    ? String(sheet.getRange(row, columns.status).getDisplayValue()).trim()
    : '';

  if (!rawTitle && !rawIsbn) {
    if (columns.status) {
      sheet.getRange(row, columns.status).clearContent();
    }

    return false;
  }

  if (!force && shouldSkipStatus(currentStatus)) {
    return false;
  }

  setStatus(sheet, row, columns, '⏳ 조회 중...');
  SpreadsheetApp.flush();

  try {
    const parsedInput = parseTitleAndPublisher(rawTitle);

    // 검색 전에 현재 신청 시트 중복 확인
    let duplicateMessage = findDuplicateRequest(
      sheet,
      row,
      columns,
      parsedInput.title,
      rawIsbn
    );

    // 입력된 ISBN이 있는 경우 학교 소장도서 확인
    const currentStatus = columns.status
      ? String(sheet.getRange(row, columns.status).getDisplayValue() || '').trim()
      : '';
    const managerOwnedBookException = currentStatus.includes('관리자 소장도서 예외');

    if (!duplicateMessage && rawIsbn && !managerOwnedBookException) {
      duplicateMessage = findLibraryDuplicate(
        rawIsbn,
        sheet.getParent()
      );
    }

    const book = smartSearch(
      parsedInput.title,
      rawIsbn,
      parsedInput.publisher
    );

    if (!book) {
      const message = duplicateMessage
        ? duplicateMessage + ' · 정보검색 실패'
        : '❌ 검색실패';

      setStatus(sheet, row, columns, message);

      return false;
    }

    // 도서명만 입력한 경우 검색 결과 ISBN으로 다시 중복 확인
    if (!duplicateMessage) {
      duplicateMessage = findDuplicateRequest(
        sheet,
        row,
        columns,
        book.title,
        book.isbn
      );
    }

    if (!duplicateMessage) {
      duplicateMessage = findLibraryDuplicate(
        book.isbn,
        sheet.getParent()
      );
    }

    fillFinalData(
      sheet,
      row,
      book,
      headers,
      duplicateMessage || '✅ 완료'
    );

    return true;

  } catch (error) {
    console.error(`행 ${row} 처리 오류`, error);

    const errorMessage = error && error.message
      ? `❌ 오류: ${String(error.message).substring(0, 100)}`
      : '❌ 처리 중 오류 발생';

    setStatus(sheet, row, columns, errorMessage);

    return false;
  }
}


/**
 * 완료 또는 중복 처리된 행은 자동 재조회하지 않음
 */
function shouldSkipStatus(status) {
  if (!status) return false;

  return (
    !status.includes('관리자 소장도서 예외') &&
    (
      status.includes('완료') ||
      status.includes('이미 신청된') ||
      status.includes('소장중인')
    )
  );
}


/**
 * ============================================================
 * 5. 스마트 검색
 * ============================================================
 */
function smartSearch(title, isbn, requestedPublisher) {
  const cleanIsbn = normalizeIsbn(isbn);

  let result = null;

  // ISBN 우선 검색
  if (cleanIsbn) {
    result = fetchYes24Api(
      cleanIsbn,
      'Book',
      '',
      true
    );

    if (!result) {
      result = fetchYes24Api(
        cleanIsbn,
        'Foreign',
        '',
        true
      );
    }
  }

  // ISBN 검색 실패 또는 ISBN 미입력 시 제목 검색
  if (!result && title) {
    const searchQuery = requestedPublisher
      ? `${title} ${requestedPublisher}`
      : title;

    result = fetchYes24Api(
      searchQuery,
      'Book',
      requestedPublisher,
      false
    );

    if (!result) {
      result = fetchYes24Api(
        searchQuery,
        'Foreign',
        requestedPublisher,
        false
      );
    }
  }

  if (!result) return null;

  result.callNumber = result.isbn && result.isbn !== '-'
    ? fetchNlkCallNumber(result.isbn)
    : '-';

  return result;
}


/**
 * ============================================================
 * 6. YES24 API
 * ============================================================
 */
function fetchYes24Api(query, searchTarget, requestedPublisher, isIsbnSearch) {
  var key = getConfiguredYes24Key_();
  if (!key) throw new Error('YES24 API 키가 없습니다. 「사본 설정 → API 키 설정」에서 입력해 주세요.');
  var cacheKey = createCacheKey('YES24_V1', query, searchTarget, requestedPublisher, isIsbnSearch);
  var cached = getCachedJson(cacheKey);
  if (cached) return cached;
  var items = yes24SearchItems_(query, key, searchTarget === 'Foreign' ? 'FOREIGN' : 'BOOK', 10, isIsbnSearch);
  var selectedItem = selectBestYes24Item(items, query, requestedPublisher, isIsbnSearch);
  if (!selectedItem) return null;
  var result = {
    title: stripHtml(selectedItem.title) || '-',
    author: stripHtml(selectedItem.author) || '-',
    publisher: stripHtml(selectedItem.publisher) || '-',
    year: extractYear(selectedItem.publishDate),
    priceStd: toNumber(selectedItem.shopPrice),
    category: stripHtml(selectedItem.goodsSortNm || selectedItem.goodsType) || '-',
    isbn: extractBestIsbn(selectedItem.isbn13 || selectedItem.isbn10 || ''),
    cover: selectedItem.cover || '', callNumber: '-'
  };
  putCachedJson(cacheKey, result);
  return result;
}


/**
 * YES24 검색 결과 중 가장 적절한 항목 선택
 */
function selectBestYes24Item(
  items,
  originalQuery,
  requestedPublisher,
  isIsbnSearch
) {
  if (!items || items.length === 0) return null;

  if (isIsbnSearch) {
    return items.find(function(item) { return yes24Isbn13_(item.isbn13 || item.isbn10) === yes24Isbn13_(originalQuery); }) || null;
  }

  const normalizedQuery = normalizeText(originalQuery);
  const normalizedPublisher = normalizeText(requestedPublisher);

  let bestItem = items[0];
  let bestScore = -1;

  items.forEach(item => {
    const itemTitle = normalizeText(
      stripHtml(item.title || '')
    );

    const itemPublisher = normalizeText(
      item.publisher || ''
    );

    let score = 0;

    if (
      normalizedQuery &&
      itemTitle === normalizedQuery
    ) {
      score += 100;

    } else if (
      normalizedQuery &&
      itemTitle.includes(normalizedQuery)
    ) {
      score += 60;

    } else if (
      itemTitle &&
      normalizedQuery.includes(itemTitle)
    ) {
      score += 40;
    }

    if (normalizedPublisher) {
      if (
        itemPublisher === normalizedPublisher
      ) {
        score += 50;

      } else if (
        itemPublisher.includes(normalizedPublisher)
      ) {
        score += 30;
      }
    }

    if (item.isbn13) score += 5;
    if (item.cover) score += 2;

    if (score > bestScore) {
      bestScore = score;
      bestItem = item;
    }
  });

  return bestItem;
}


/**
 * ============================================================
 * 7. 국립중앙도서관 API
 * ============================================================
 */
function fetchNlkCallNumber(isbn) {
  const nlkKey = getConfiguredNlkKey();

  if (!nlkKey) return '-';

  const cleanIsbn = normalizeIsbn(isbn);

  if (!cleanIsbn) return '-';

  const cacheKey = createCacheKey('NLK', cleanIsbn);

  const cachedValue = CacheService
    .getScriptCache()
    .get(cacheKey);

  if (cachedValue !== null) {
    return cachedValue || '-';
  }

  const url =
    'https://www.nl.go.kr/NL/search/openApi/search.do' +
    `?key=${encodeURIComponent(nlkKey)}` +
    '&detailSearch=true' +
    '&isbnOp=isbn' +
    `&isbnCode=${encodeURIComponent(cleanIsbn)}`;

  try {
    const response = UrlFetchApp.fetch(url, {
      method: 'get',
      muteHttpExceptions: true,
      followRedirects: true
    });

    if (response.getResponseCode() !== 200) {
      console.error(
        `국립중앙도서관 API HTTP 오류: ${response.getResponseCode()}`
      );

      return '-';
    }

    const xmlText = response.getContentText('UTF-8');
    const classNumber = parseNlkClassNumber(xmlText);

    CacheService
      .getScriptCache()
      .put(
        cacheKey,
        classNumber,
        CACHE_SECONDS
      );

    return classNumber;

  } catch (error) {
    console.error('국립중앙도서관 조회 오류', error);

    return '-';
  }
}


/**
 * 국립중앙도서관 XML 결과 분석
 */
function parseNlkClassNumber(xmlText) {
  if (!xmlText) return '-';

  try {
    const document = XmlService.parse(xmlText);
    const root = document.getRootElement();

    let classNumber =
      findFirstXmlValue(root, 'class_no') ||
      findFirstXmlValue(root, 'call_no') ||
      '';

    if (!classNumber) return '-';

    classNumber = String(classNumber).trim();

    if (classNumber.includes('-')) {
      classNumber = classNumber
        .split('-')[0]
        .trim();
    }

    return normalizeClassNumber(classNumber);

  } catch (error) {
    console.error('국립중앙도서관 XML 분석 오류', error);

    const classMatch = xmlText.match(
      /<class_no>(?:<!\[CDATA\[)?(.*?)(?:\]\]>)?<\/class_no>/i
    );

    if (classMatch && classMatch[1]) {
      return normalizeClassNumber(
        classMatch[1]
      );
    }

    const callMatch = xmlText.match(
      /<call_no>(?:<!\[CDATA\[)?(.*?)(?:\]\]>)?<\/call_no>/i
    );

    if (callMatch && callMatch[1]) {
      return normalizeClassNumber(
        callMatch[1].split('-')[0]
      );
    }

    return '-';
  }
}


/**
 * XML 트리에서 첫 번째 태그값 찾기
 */
function findFirstXmlValue(element, tagName) {
  if (!element) return '';

  if (element.getName() === tagName) {
    return element.getText();
  }

  const children = element.getChildren();

  for (let i = 0; i < children.length; i++) {
    const value = findFirstXmlValue(
      children[i],
      tagName
    );

    if (value) return value;
  }

  return '';
}


/**
 * 분류기호 정리
 */
function normalizeClassNumber(value) {
  if (!value) return '-';

  let result = String(value)
    .replace(/<!\[CDATA\[|\]\]>/g, '')
    .trim();

  const numberMatch = result.match(
    /\d{3}(?:\.\d+)?/
  );

  if (numberMatch) {
    result = numberMatch[0];
  }

  // 소수점 첫째 자리까지만 사용
  if (result.includes('.')) {
    const parts = result.split('.');

    const decimal = parts[1]
      ? parts[1].substring(0, 1)
      : '';

    result = decimal
      ? `${parts[0]}.${decimal}`
      : parts[0];
  }

  // 학교 운영 기준에 맞춘 아동문학 보정
  if (result === '813.7') {
    result = '813.8';
  }

  return result || '-';
}


/**
 * ============================================================
 * 8. 현재 신청 시트 중복 확인
 * ============================================================
 */
function findDuplicateRequest(
  sheet,
  currentRow,
  columns,
  title,
  isbn
) {
  const lastRow = sheet.getLastRow();

  if (lastRow < FIRST_DATA_ROW) return '';

  const cleanTitle = normalizeText(title);
  const cleanIsbn = normalizeIsbn(isbn);

  if (!cleanTitle && !cleanIsbn) return '';

  const data = sheet
    .getRange(
      FIRST_DATA_ROW,
      1,
      lastRow - FIRST_DATA_ROW + 1,
      sheet.getLastColumn()
    )
    .getDisplayValues();

  for (let index = 0; index < data.length; index++) {
    const sheetRow = FIRST_DATA_ROW + index;

    if (sheetRow === currentRow) continue;

    const existingTitle = columns.title
      ? normalizeText(
          data[index][columns.title - 1]
        )
      : '';

    const existingIsbn = columns.isbn
      ? normalizeIsbn(
          data[index][columns.isbn - 1]
        )
      : '';

    const sameIsbn =
      cleanIsbn &&
      existingIsbn &&
      cleanIsbn === existingIsbn;

    const canCompareTitle =
      !cleanIsbn ||
      !existingIsbn;

    const sameTitle =
      canCompareTitle &&
      cleanTitle &&
      existingTitle &&
      cleanTitle === existingTitle;

    if (sameIsbn || sameTitle) {
      return `❌ 이미 신청된 도서입니다 (${sheetRow}행)`;
    }
  }

  return '';
}


/**
 * ============================================================
 * 9. 학교 소장도서 중복 확인
 * ============================================================
 */
function findLibraryDuplicate(isbn, spreadsheet) {
  const cleanIsbn = normalizeIsbn(isbn);

  if (!cleanIsbn) return '';

  const librarySheet = spreadsheet.getSheetByName(
    LIBRARY_SHEET_NAME
  );

  if (!librarySheet) return '';

  const lastRow = librarySheet.getLastRow();
  const lastColumn = librarySheet.getLastColumn();

  if (
    lastRow < FIRST_DATA_ROW ||
    lastColumn < 1
  ) {
    return '';
  }

  const headers = librarySheet
    .getRange(
      HEADER_ROW,
      1,
      1,
      lastColumn
    )
    .getDisplayValues()[0];

  const columns = getLibraryColumnMap(headers);

  if (!columns.isbn) return '';

  const data = librarySheet
    .getRange(
      FIRST_DATA_ROW,
      1,
      lastRow - FIRST_DATA_ROW + 1,
      lastColumn
    )
    .getDisplayValues();

  for (let index = 0; index < data.length; index++) {
    const libraryIsbn = normalizeIsbn(
      data[index][columns.isbn - 1]
    );

    if (
      libraryIsbn &&
      libraryIsbn === cleanIsbn
    ) {
      const sequence = columns.sequence
        ? String(
            data[index][columns.sequence - 1]
          ).trim()
        : String(FIRST_DATA_ROW + index);

      return `❌ 소장중인 도서 (순번: ${sequence})`;
    }
  }

  return '';
}


/**
 * ============================================================
 * 10. 검색 성공 데이터 입력
 * ============================================================
 */
function fillFinalData(
  sheet,
  row,
  book,
  headers,
  finalStatus
) {
  headers.forEach((header, index) => {
    const normalizedHeader = normalizeHeader(header);

    if (!normalizedHeader) return;

    const cell = sheet.getRange(
      row,
      index + 1
    );

    if (isTitleHeader(normalizedHeader)) {
      HR_setSafeText_(cell, book.title || '-');

    } else if (isIsbnHeader(normalizedHeader)) {
      cell.setNumberFormat('@');

      cell.setValue(
        book.isbn || '-'
      );

    } else if (isCoverHeader(normalizedHeader)) {
      const coverUrl = HR_safeCoverUrl_(book.cover);
      if (coverUrl) {
        cell.setFormula('=IMAGE("' + coverUrl + '",4,120,90)');
      }

    } else if (isPublisherHeader(normalizedHeader)) {
      HR_setSafeText_(cell, book.publisher || '-');

    } else if (isAuthorHeader(normalizedHeader)) {
      HR_setSafeText_(cell, book.author || '-');

    } else if (isPriceHeader(normalizedHeader)) {
      cell.setValue(
        toNumber(book.priceStd)
      );

      cell.setNumberFormat('#,##0');

    } else if (isYearHeader(normalizedHeader)) {
      HR_setSafeText_(cell, book.year || '-');

    } else if (isCategoryHeader(normalizedHeader)) {
      HR_setSafeText_(cell, book.category || '-');

    } else if (isCallNumberHeader(normalizedHeader)) {
      HR_setSafeText_(cell, book.callNumber || '-');

    } else if (isStatusHeader(normalizedHeader)) {
      cell.setValue(
        finalStatus || '✅ 완료'
      );
    }
  });

  try {
    sheet.setRowHeight(row, 120);
  } catch (error) {
    console.error(
      '행 높이 설정 오류',
      error
    );
  }
}


/**
 * ============================================================
 * 11. 기존 데이터 일괄 처리
 * ============================================================
 */
function checkExistingBooks() {
  const ui = SpreadsheetApp.getUi();

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  const sheet = spreadsheet.getSheetByName(
    TARGET_SHEET_NAME
  );

  if (!sheet) {
    ui.alert(
      `❌ '${TARGET_SHEET_NAME}' 시트를 찾을 수 없습니다.`
    );

    return;
  }

  const lastRow = sheet.getLastRow();

  if (lastRow < FIRST_DATA_ROW) {
    ui.alert(
      '처리할 데이터가 없습니다.'
    );

    return;
  }

  const headers = getHeaders(sheet);
  const columns = getColumnMap(headers);

  if (!columns.title && !columns.isbn) {
    ui.alert(
      '❌ 도서명 또는 ISBN 열을 찾을 수 없습니다.'
    );

    return;
  }

  let processedCount = 0;
  let successCount = 0;

  const lock = LockService.getDocumentLock();

  try {
    lock.waitLock(30000);

    for (
      let row = FIRST_DATA_ROW;
      row <= lastRow;
      row++
    ) {
      const title = columns.title
        ? String(
            sheet
              .getRange(row, columns.title)
              .getDisplayValue()
          ).trim()
        : '';

      const isbn = columns.isbn
        ? String(
            sheet
              .getRange(row, columns.isbn)
              .getDisplayValue()
          ).trim()
        : '';

      const status = columns.status
        ? String(
            sheet
              .getRange(row, columns.status)
              .getDisplayValue()
          ).trim()
        : '';

      if (!title && !isbn) continue;
      if (shouldSkipStatus(status)) continue;

      const success = processRow(
        sheet,
        row,
        false,
        headers,
        columns
      );

      processedCount++;

      if (success) {
        successCount++;
      }

      Utilities.sleep(API_SLEEP_MS);
    }

  } catch (error) {
    console.error(
      '일괄 처리 오류',
      error
    );

    ui.alert(
      '❌ 일괄 처리 중 오류가 발생했습니다.\n\n' +
      error.message
    );

    return;

  } finally {
    try {
      lock.releaseLock();
    } catch (error) {
      // 무시
    }
  }

  if (processedCount === 0) {
    ui.alert(
      '💡 새롭게 처리할 데이터가 없습니다.'
    );

    return;
  }

  ui.alert(
    '✅ 일괄 처리 완료',
    [
      `처리한 행: ${processedCount}건`,
      `검색 성공: ${successCount}건`,
      `검색 실패: ${processedCount - successCount}건`
    ].join('\n'),
    ui.ButtonSet.OK
  );
}


/**
 * ============================================================
 * 12. 선택한 행 강제 재조회
 * ============================================================
 */
function reprocessSelectedRows() {
  const ui = SpreadsheetApp.getUi();

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  const sheet = spreadsheet.getActiveSheet();

  if (sheet.getName() !== TARGET_SHEET_NAME) {
    ui.alert(
      `❌ '${TARGET_SHEET_NAME}' 시트에서 실행해 주세요.`
    );

    return;
  }

  const range = sheet.getActiveRange();

  if (!range) {
    ui.alert(
      '다시 조회할 행을 선택해 주세요.'
    );

    return;
  }

  const headers = getHeaders(sheet);
  const columns = getColumnMap(headers);

  const startRow = Math.max(
    range.getRow(),
    FIRST_DATA_ROW
  );

  const endRow = Math.min(
    range.getLastRow(),
    sheet.getLastRow()
  );

  let processedCount = 0;

  const lock = LockService.getDocumentLock();

  try {
    lock.waitLock(30000);

    for (
      let row = startRow;
      row <= endRow;
      row++
    ) {
      const title = columns.title
        ? String(
            sheet
              .getRange(row, columns.title)
              .getDisplayValue()
          ).trim()
        : '';

      const isbn = columns.isbn
        ? String(
            sheet
              .getRange(row, columns.isbn)
              .getDisplayValue()
          ).trim()
        : '';

      if (!title && !isbn) continue;

      processRow(
        sheet,
        row,
        true,
        headers,
        columns
      );

      processedCount++;

      Utilities.sleep(API_SLEEP_MS);
    }

  } catch (error) {
    console.error(
      '재조회 오류',
      error
    );

    ui.alert(
      '❌ 다시 조회 중 오류가 발생했습니다.\n\n' +
      error.message
    );

    return;

  } finally {
    try {
      lock.releaseLock();
    } catch (error) {
      // 무시
    }
  }

  ui.alert(
    processedCount > 0
      ? `✅ 선택한 ${processedCount}개 행을 다시 조회했습니다.`
      : '처리할 도서 정보가 없습니다.'
  );
}


/**
 * ============================================================
 * 13. 헤더 및 열 탐색
 * ============================================================
 */
function getHeaders(sheet) {
  const lastColumn = sheet.getLastColumn();

  if (lastColumn < 1) return [];

  return sheet
    .getRange(
      HEADER_ROW,
      1,
      1,
      lastColumn
    )
    .getDisplayValues()[0];
}


function getColumnMap(headers) {
  const columns = {
    title: 0,
    isbn: 0,
    status: 0
  };

  headers.forEach((header, index) => {
    const normalized = normalizeHeader(header);

    if (isTitleHeader(normalized)) {
      columns.title = index + 1;
    }

    if (isIsbnHeader(normalized)) {
      columns.isbn = index + 1;
    }

    if (isStatusHeader(normalized)) {
      columns.status = index + 1;
    }
  });

  return columns;
}


function getLibraryColumnMap(headers) {
  const columns = {
    isbn: 0,
    sequence: 0
  };

  headers.forEach((header, index) => {
    const normalized = normalizeHeader(header);

    if (isIsbnHeader(normalized)) {
      columns.isbn = index + 1;
    }

    if (
      normalized.includes('순번') ||
      normalized === '번호' ||
      normalized === 'NO' ||
      normalized === 'NUM'
    ) {
      columns.sequence = index + 1;
    }
  });

  return columns;
}


/**
 * ============================================================
 * 14. 상태값 입력
 * ============================================================
 */
function setStatus(
  sheet,
  row,
  columns,
  status
) {
  if (!columns.status) return;

  sheet
    .getRange(
      row,
      columns.status
    )
    .setValue(status);
}


/**
 * ============================================================
 * 15. 제목 / 출판사 입력 분석
 *
 * 예:
 * 긴긴밤/문학동네
 * ============================================================
 */
function parseTitleAndPublisher(input) {
  const value = String(input || '').trim();

  if (!value.includes('/')) {
    return {
      title: value,
      publisher: ''
    };
  }

  const separatorIndex = value.lastIndexOf('/');

  return {
    title: value
      .substring(0, separatorIndex)
      .trim(),

    publisher: value
      .substring(separatorIndex + 1)
      .trim()
  };
}


/**
 * ============================================================
 * 16. ISBN 처리
 * ============================================================
 */
function extractBestIsbn(value) {
  if (!value) return '-';

  const text = String(value).toUpperCase();

  const candidates = text
    .split(/[\s,/]+/)
    .map(item =>
      item.replace(/[^0-9X]/g, '')
    )
    .filter(Boolean);

  const isbn13 = candidates.find(
    item => /^\d{13}$/.test(item)
  );

  if (isbn13) return isbn13;

  const isbn10 = candidates.find(
    item => /^\d{9}[0-9X]$/.test(item)
  );

  if (isbn10) return isbn10;

  const allCharacters = text.replace(
    /[^0-9X]/g,
    ''
  );

  const embeddedIsbn13 = allCharacters.match(
    /\d{13}/
  );

  if (embeddedIsbn13) {
    return embeddedIsbn13[0];
  }

  const embeddedIsbn10 = allCharacters.match(
    /\d{9}[0-9X]/
  );

  if (embeddedIsbn10) {
    return embeddedIsbn10[0];
  }

  return '-';
}


function normalizeIsbn(value) {
  const isbn = extractBestIsbn(value);

  return isbn === '-'
    ? ''
    : isbn;
}


/**
 * ============================================================
 * 17. 문자열 처리
 * ============================================================
 */
function normalizeText(value) {
  return String(value || '')
    .toLowerCase()
    .replace(/<\/?[^>]+>/g, '')
    .replace(
      /[\s\-_:：·ㆍ,.()[\]{}'"“”‘’]/g,
      ''
    )
    .trim();
}


function normalizeHeader(value) {
  return String(value || '')
    .replace(/\s/g, '')
    .toUpperCase()
    .trim();
}


function stripHtml(value) {
  return String(value || '')
    .replace(/<\/?[^>]+>/g, '')
    .replace(/&amp;/g, '&')
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .trim();
}


function extractYear(value) {
  const match = String(value || '').match(
    /\d{4}/
  );

  return match
    ? match[0]
    : '-';
}


function toNumber(value) {
  const number = Number(
    String(value || '')
      .replace(/[^0-9.-]/g, '')
  );

  return Number.isFinite(number)
    ? number
    : 0;
}


/**
 * ============================================================
 * 18. 헤더 이름 판별
 * ============================================================
 */
function isTitleHeader(header) {
  return (
    header.includes('도서명') ||
    header === '책제목' ||
    header === '제목'
  );
}


function isIsbnHeader(header) {
  return header.includes('ISBN');
}


function isStatusHeader(header) {
  return (
    header.includes('상태') ||
    header.includes('처리결과') ||
    header.includes('조회결과')
  );
}


function isCoverHeader(header) {
  return (
    header.includes('표지') ||
    header.includes('이미지')
  );
}


function isPublisherHeader(header) {
  return header.includes('출판사');
}


function isAuthorHeader(header) {
  return (
    header.includes('저자') ||
    header.includes('지은이')
  );
}


function isPriceHeader(header) {
  return (
    header === '정가' ||
    header.includes('도서정가')
  );
}


function isYearHeader(header) {
  return (
    header.includes('출판연도') ||
    header.includes('발행연도') ||
    header === '연도'
  );
}


function isCategoryHeader(header) {
  return (
    header.includes('분류') ||
    header.includes('카테고리')
  );
}


function isCallNumberHeader(header) {
  return (
    header.includes('청구기호') ||
    header.includes('분류기호')
  );
}


/**
 * ============================================================
 * 19. API 결과 캐시
 * ============================================================
 */
function createCacheKey() {
  const rawValue = Array.prototype.slice
    .call(arguments)
    .map(value => String(value || ''))
    .join('|');

  const digest = Utilities.computeDigest(
    Utilities.DigestAlgorithm.MD5,
    rawValue,
    Utilities.Charset.UTF_8
  );

  const digestText = digest
    .map(byte => {
      const adjusted = byte < 0
        ? byte + 256
        : byte;

      return adjusted
        .toString(16)
        .padStart(2, '0');
    })
    .join('');

  return `BOOK_${digestText}`;
}


function getCachedJson(cacheKey) {
  try {
    const value = CacheService
      .getScriptCache()
      .get(cacheKey);

    if (!value) return null;

    return JSON.parse(value);

  } catch (error) {
    return null;
  }
}


function putCachedJson(cacheKey, value) {
  try {
    CacheService
      .getScriptCache()
      .put(
        cacheKey,
        JSON.stringify(value),
        CACHE_SECONDS
      );

  } catch (error) {
    console.error(
      '캐시 저장 오류',
      error
    );
  }
}

function FIX_ACTION_LAYOUT_20260820() {
  var spreadsheet = SpreadsheetApp.getActive();
  styleRequestListSheet();

  var seriesSheet = spreadsheet.getSheetByName('🔎 시리즈 누락 점검');
  if (seriesSheet) {
    seriesAddWorkButton_(seriesSheet);
  }

  var reviewSheet = spreadsheet.getSheetByName('🛒 구입 검토함');
  if (reviewSheet) {
    unifiedEnsureReviewSheet_();
  }

  var exportSheet = spreadsheet.getSheetByName('📦 서점 전달용');
  if (exportSheet) {
    exportSheet.getRange('A4:J4').breakApart().clearContent().clearDataValidations();
    exportSheet.getRange('A4')
      .insertCheckboxes()
      .setValue(false)
      .setBackground('#D9A441')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');
    exportSheet.getRange('B4:J4')
      .merge()
      .setValue('최신 구입 검토 결과로 다시 만들기')
      .setBackground('#D9A441')
      .setFontColor('#FFFFFF')
      .setFontSize(14)
      .setFontWeight('bold')
      .setHorizontalAlignment('left')
      .setVerticalAlignment('middle');
    exportSheet.getRange('A4:J4').setBorder(
      true,true,true,true,false,false,
      '#B88428',SpreadsheetApp.BorderStyle.SOLID_MEDIUM
    );
    exportSheet.setRowHeight(4, 44);
  }

  SpreadsheetApp.flush();
  return '모든 업무 버튼 체크박스를 A4로 정렬했습니다.';
}


function ensureSelectionControls_() {
  const spreadsheet = SpreadsheetApp.getActive();
  const seriesSheet = spreadsheet.getSheetByName(SERIES_FINDER_SHEET);
  if (seriesSheet) {
    seriesAddWorkButton_(seriesSheet);
  }

  const reviewSheet = spreadsheet.getSheetByName(UNIFIED_REVIEW_SHEET);
  if (reviewSheet) {
    unifiedEnsureReviewSheet_();
  }
}


function toggleSelectionColumn_(sheet, startRow, column, checked) {
  const lastRow = sheet.getLastRow();
  if (lastRow < startRow) {
    sheet.getRange(4, 1)
      .setValue(false);
    SpreadsheetApp.getActive().toast('선택할 도서가 없습니다.', '목록 없음', 4);
    return;
  }

  const rowCount = lastRow - startRow + 1;
  const values = Array.from({length: rowCount}, () => [checked]);
  sheet.getRange(startRow, column, rowCount, 1).setValues(values);
  SpreadsheetApp.getActive().toast(
    checked ? rowCount + '권을 모두 선택했습니다.' : '전체 선택을 해제했습니다.',
    checked ? '전체 선택' : '전체 해제',
    4
  );
}


/**
 * 다음 학기 시작 시 이번 학기 접수·검토 데이터를 비우고 서점 전달용 시트를 숨깁니다.
 * 상단 안내/버튼/서식과 예시 행, 소장도서 및 복구·백업 자료는 보존합니다.
 */
function resetHopeBooksForNewSemester() {
  const ui = SpreadsheetApp.getUi();
  const answer = ui.alert(
    '새 학기 초기화 확인',
    '아래 작업을 진행할까요?\n\n' +
      '• 희망도서 목록: 예시 행은 남기고 실제 신청 목록만 비우기\n' +
      '• 시리즈 누락 점검, 엑셀 업로드 목록, 구입 검토함: 데이터 비우기\n' +
      '• 서점 전달용: 내용을 보존하고 시트 숨기기\n\n' +
      '사용방법·신청안내·웹 공지·대시보드·우리학교 소장도서·취소복구·백업은 유지합니다.',
    ui.ButtonSet.YES_NO
  );
  if (answer !== ui.Button.YES) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const resetTargets = [
    { name: '📚 희망도서 목록', firstDataRow: 7, checkboxes: ['A4'] },
    { name: '🔎 시리즈 누락 점검', firstDataRow: 6, checkboxes: ['A4', 'D4'] },
    { name: '📥 엑셀 업로드 목록', firstDataRow: 6, checkboxes: ['A4', 'D4'] },
    { name: '🛒 구입 검토함', firstDataRow: 6, checkboxes: ['A4', 'D4'] }
  ];
  const bookstore = ss.getSheetByName('📦 서점 전달용');
  const missing = resetTargets
    .filter(item => !ss.getSheetByName(item.name))
    .map(item => item.name);
  if (!bookstore) missing.push('📦 서점 전달용');
  if (missing.length) {
    ui.alert('초기화를 중단했습니다', '필요한 시트를 찾지 못했습니다: ' + missing.join(', '), ui.ButtonSet.OK);
    return;
  }

  resetTargets.forEach(item => {
    const sheet = ss.getSheetByName(item.name);
    const lastRow = sheet.getLastRow();
    const lastColumn = sheet.getLastColumn();
    if (lastRow >= item.firstDataRow && lastColumn > 0) {
      sheet.getRange(item.firstDataRow, 1, lastRow - item.firstDataRow + 1, lastColumn).clearContent();
    }
    item.checkboxes.forEach(a1 => sheet.getRange(a1).setValue(false));
  });

  ss.getSheetByName('🔎 시리즈 누락 점검')
    .getRange('A3').setValue('새 학기 준비 완료 · 점검 결과가 비워졌습니다.');
  bookstore.hideSheet();
  ss.toast('신청·검토 데이터 초기화 완료 · 서점 전달용 시트는 숨김 처리했습니다.', '새 학기 준비', 8);
}



/* ===== ImportFromOld.gs ===== */
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


/* ===== RequestSidebarServer.js ===== */
/**
 * 희망도서 신청 사이드바
 *
 * 기존 Code.gs의 processRow(), getHeaders(), getColumnMap()을 그대로 사용해
 * YES24 검색, 소장도서·중복 확인, 청구기호 자동입력을 실행합니다.
 */

var HR_TARGET_SHEET_NAME = '📚 희망도서 목록';
var HR_LEGACY_TARGET_SHEET_NAME = '2차 희망도서';
var HR_LIBRARY_SHEET_NAME = '🏫 우리학교 소장도서';
var HR_COPY_SETUP_SHEET_NAME = '⚙️ 사본 최초설정';
var HR_SIDEBAR_TRIGGER_FUNCTION = 'HR_openSidebarOnSpreadsheetOpen';
var HR_SPREADSHEET_ID_PROPERTY = 'HR_TARGET_SPREADSHEET_ID';
var HR_YES24_KEY_PROPERTY = 'HR_YES24_API_KEY';
var HR_NLK_KEY_PROPERTY = 'HR_NLK_API_KEY';
var HR_SCHOOL_NAME_PROPERTY = 'HR_SCHOOL_NAME';
var HR_MANAGER_NAME_PROPERTY = 'HR_MANAGER_NAME';
var HR_MANAGER_EXTENSION_PROPERTY = 'HR_MANAGER_EXTENSION';
var HR_SETUP_COMPLETE_PROPERTY = 'HR_SETUP_COMPLETE';
var HR_PUBLIC_WEB_APP_URL_PROPERTY = 'HR_PUBLIC_WEB_APP_URL';
var HR_REQUEST_BACKUP_SHEET_NAME = '🔒 신청접수백업';
var HR_REQUEST_CANCEL_BACKUP_SHEET_NAME = '🔒 신청취소복구';
var HR_REQUEST_LOCK_WAIT_MS = 30000;
var HR_YES24_API_MANAGE_URL =
  'https://developers.yes24.com/';

/**
 * 권한 승인 없이 사용하는 공개 신청 웹앱 화면입니다.
 */
function doGet(e) {
  var properties = PropertiesService.getScriptProperties();
  var audience = HR_audienceFromParameter_(e && e.parameter && e.parameter.type);
  var schoolName = String(
    properties.getProperty(HR_SCHOOL_NAME_PROPERTY) || ''
  ).trim();
  var managerName = String(
    properties.getProperty(HR_MANAGER_NAME_PROPERTY) || ''
  ).trim();
  var managerExtension = String(
    properties.getProperty(HR_MANAGER_EXTENSION_PROPERTY) || ''
  ).trim();
  var siteTitle = schoolName
    ? schoolName + ' 희망도서 신청'
    : '희망도서 신청';
  var template = HtmlService.createTemplateFromFile(
    'RequestSidebar'
  );

  template.siteTitle = siteTitle;
  template.audience = audience;
  template.fromGate = String(e && e.parameter && e.parameter.g || '') === '1';
  template.baseUrl = HR_getWebAppBaseUrl_();
  template.managerName = managerName;
  template.managerExtension = managerExtension;

  return template
    .evaluate()
    .setTitle('📚 ' + siteTitle)
    .addMetaTag(
      'viewport',
      'width=device-width, initial-scale=1'
    );
}

// 이 웹 앱의 주소(…/exec). 못 구하면 빈 글자입니다.
function HR_getWebAppBaseUrl_() {
  try {
    var url = String(ScriptApp.getService().getUrl() || '');
    return /^https:\/\/script\.google\.com\/.+\/exec$/.test(url) ? url : '';
  } catch (error) {
    return '';
  }
}

/**
 * 최초 설정에서 저장한 학교 스프레드시트를 사용합니다.
 */
function HR_getSpreadsheet_() {
  
  var properties = PropertiesService.getScriptProperties();
  var spreadsheetId = String(
    properties.getProperty(HR_SPREADSHEET_ID_PROPERTY) || ''
  ).trim();

  if (!spreadsheetId) {
    throw new Error(
      '학교 최초 설정이 필요합니다. ' +
      '⚙️ 사본 최초설정 탭에서 설정을 완료해 주세요.'
    );
  }

  return SpreadsheetApp.openById(spreadsheetId);
}

/**
 * 사용설명서 버튼에는 소유자 전용 /dev 주소가 아니라
 * 누구나 접속 가능한 공개 /exec 주소를 우선 사용합니다.
 */
function HR_getConfiguredPublicWebAppUrl_() {
  var properties = PropertiesService.getScriptProperties();
  var savedUrl = String(
    properties.getProperty(HR_PUBLIC_WEB_APP_URL_PROPERTY) || ''
  ).trim();

  return savedUrl;
}

/**
 * 이전 사본의 시트 이름을 새 이름으로 한 번만 바꿉니다.
 */
function HR_renameLegacyTargetSheet() {
  var spreadsheet = HR_getSetupSpreadsheet_();
  var currentSheet = spreadsheet.getSheetByName(
    HR_TARGET_SHEET_NAME
  );
  var legacySheet;

  if (currentSheet) {
    return {
      success: true,
      message: '희망도서 목록 시트 이름이 이미 적용되어 있습니다.'
    };
  }

  legacySheet = spreadsheet.getSheetByName(
    HR_LEGACY_TARGET_SHEET_NAME
  );

  if (!legacySheet) {
    throw new Error(
      '희망도서 목록으로 바꿀 기존 시트를 찾지 못했습니다.'
    );
  }

  legacySheet.setName(HR_TARGET_SHEET_NAME);

  return {
    success: true,
    message: '시트 이름을 「희망도서 목록」으로 변경했습니다.'
  };
}

/**
 * 학교별로 저장한 YES24 API Key를 가져옵니다.
 * 기존 프로젝트는 Code.gs의 YES24_KEY를 예비값으로 사용합니다.
 */
function HR_getYes24Key_() {
  var properties = PropertiesService.getScriptProperties();
  var savedKey = String(
    properties.getProperty(HR_YES24_KEY_PROPERTY) || ''
  ).trim();

  if (savedKey) {
    return savedKey;
  }

  return (
    typeof YES24_KEY !== 'undefined'
      ? String(YES24_KEY || '').trim()
      : ''
  );
}

/**
 * 국립중앙도서관 API 키를 가져옵니다.
 * KDC 참고용 선택 기능이며, 키가 없으면 빈 값으로 둡니다.
 */
function HR_getNlkKey_() {
  var properties = PropertiesService.getScriptProperties();
  var savedKey = String(
    properties.getProperty(HR_NLK_KEY_PROPERTY) || ''
  ).trim();
  var legacyKey = '';

  if (savedKey) {
    return savedKey;
  }

  // 코드에 적혀 있던 기본 키는 처음 쓸 때 스크립트 속성으로 옮깁니다.
  legacyKey = typeof NLK_KEY !== 'undefined'
    ? String(NLK_KEY || '').trim()
    : '';
  if (legacyKey) {
    properties.setProperty(HR_NLK_KEY_PROPERTY, legacyKey);
  }
  return legacyKey;
}

/**
 * 희망도서 신청 사이드바를 엽니다.
 */
function HR_showRequestSidebar() {
  var html = HtmlService
    .createHtmlOutputFromFile('RequestSidebar')
    .setTitle('📚 희망도서 신청');

  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * 스프레드시트가 열릴 때 실행되는 설치형 트리거입니다.
 * 신청 메뉴를 추가하고 희망도서 목록 탭이면 사이드바를 자동으로 엽니다.
 */
function HR_openSidebarOnSpreadsheetOpen() {
  // 예전에는 시트를 열 때마다 「희망도서 신청 바로가기」 창을 띄웠습니다.
  // 지금은 창을 띄우지 않고, 남아 있는 옛 트리거를 지웁니다.
  HR_removeLegacyOpenTrigger_();
}

/**
 * 옛 사본에 남은 「열 때 신청 창 띄우기」 트리거를 지웁니다.
 */
function HR_removeLegacyOpenTrigger_() {
  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    if (trigger.getHandlerFunction() === HR_SIDEBAR_TRIGGER_FUNCTION) {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}

/**
 * 자동 열기 트리거를 한 번만 설치합니다.
 */
function HR_installSidebarOpenTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  var triggerIndex = 0;

  for (
    triggerIndex = 0;
    triggerIndex < triggers.length;
    triggerIndex++
  ) {
    if (
      triggers[triggerIndex].getHandlerFunction() ===
      HR_SIDEBAR_TRIGGER_FUNCTION
    ) {
      ScriptApp.deleteTrigger(triggers[triggerIndex]);
    }
  }

  HR_addCopySetupMenu();

  SpreadsheetApp.getActiveSpreadsheet().toast(
    '희망도서 신청은 공개 웹페이지에서 이용합니다.',
    '도서 신청',
    5
  );
}

/**
 * 신청자와 제목·ISBN을 새 행에 넣고 기존 자동입력을 실행합니다.
 */
function HR_submitBookRequest(formData) {
  var applicant = HR_cleanInput_(
    formData && formData.applicant,
    40
  );
  var requestAudience = HR_audienceKey_(formData && formData.audience);
  var configuredManagerName = String(
    PropertiesService.getScriptProperties()
      .getProperty(HR_MANAGER_NAME_PROPERTY) || ''
  ).trim();
  var isManagerApplicant = !!configuredManagerName &&
    applicant.toLowerCase() === configuredManagerName.toLowerCase();
  var query = String(
    formData && formData.query || ''
  ).trim();
  var selectedBook =
    formData && formData.book &&
    typeof formData.book === 'object'
      ? formData.book
      : {};
  // 브라우저가 보낸 도서 정보는 믿지 않고, 정리한 뒤 YES24에서 다시 확인한 값으로 바꿉니다.
  selectedBook = HR_prepareSelectedBook_(selectedBook, query);
  var selectedPrice = Number(
    selectedBook.price ||
    selectedBook.priceStandard ||
    selectedBook.priceStd ||
    selectedBook.priceSales ||
    0
  );
  var selectedYear = String(
    selectedBook.year ||
    selectedBook.pubDate ||
    selectedBook.publicationDate ||
    ''
  ).match(/\d{4}/);
  var normalizedIsbn = query.replace(/[^0-9Xx]/g, '');
  var requestId = HR_cleanInput_(
    formData && formData.requestId,
    80
  ) || Utilities.getUuid();
  var isIsbn = (
    normalizedIsbn.length === 10 ||
    normalizedIsbn.length === 13
  );
  var spreadsheet = HR_getSpreadsheet_();
  var sheet = spreadsheet.getSheetByName(HR_TARGET_SHEET_NAME);
  var documentLock = LockService.getScriptLock();
  var headers = [];
  var columns = {};
  var applicantColumn = 11;
  var sequenceColumn = 1;
  var row = 0;
  var previousRow = 0;
  var previousRowHeight = 0;
  var sequenceValues = [];
  var nextSequence = 1;
  var sequenceIndex = 0;
  var ownedBook = null;
  var requestedBook = null;
  var lockAcquired = false;
  var backupRow = 0;
  var fastBook = {};
  var responseResult = {};
  var priceColumn = 0;
  var yearColumn = 0;

  if (!query) {
    throw new Error('책 제목이나 ISBN을 입력해 주세요.');
  }

  if (!sheet) {
    throw new Error(
      HR_TARGET_SHEET_NAME + ' 시트를 찾지 못했습니다.'
    );
  }

  // 신청 전 학교 소장도서 여부를 먼저 확인합니다.
  ownedBook = HR_findOwnedBook_(
    spreadsheet,
    selectedBook
  );

  if (ownedBook.found && !isManagerApplicant) {
    return {
      success: false,
      owned: true,
      message: '이미 소장중인 도서입니다.',
      title: ownedBook.title || selectedBook.title || query,
      isbn: ownedBook.isbn || selectedBook.isbn || normalizedIsbn,
      status: '이미 소장중인 도서입니다.'
    };
  }

  try {
    // 동시 신청 시 행이 겹치지 않도록 저장 구간만 잠급니다.
    lockAcquired = documentLock.tryLock(
      HR_REQUEST_LOCK_WAIT_MS
    );

    if (!lockAcquired) {
      throw new Error(
        '현재 신청자가 많아 잠시 대기 중입니다. 자동으로 다시 시도해 주세요.'
      );
    }

    headers = getHeaders(sheet);
    columns = getColumnMap(headers);

    // 원본 신청 내용은 숨김 백업 시트에 먼저 남깁니다.
    backupRow = HR_appendRequestBackup_(
      spreadsheet,
      {
        requestId: requestId,
        applicant: applicant,
        query: query,
        title: selectedBook.title || '',
        isbn: selectedBook.isbn || normalizedIsbn,
        status: '접수 확인 중'
      }
    );

    requestedBook = HR_findRequestedBook_(
      sheet,
      selectedBook,
      headers,
      columns
    );

    if (requestedBook.found) {
      HR_updateRequestBackup_(
        spreadsheet,
        backupRow,
        '이미 신청됨',
        '',
        ''
      );
      return {
        success: false,
        duplicateRequest: true,
        message: '이미 신청된 도서입니다.',
        title: requestedBook.title || selectedBook.title || query,
        isbn: requestedBook.isbn || selectedBook.isbn || normalizedIsbn,
        status: '이미 신청된 도서입니다.'
      };
    }

    applicantColumn = HR_findHeaderColumn_(
      headers,
      '신청자',
      11
    );
    sequenceColumn = HR_findHeaderColumn_(
      headers,
      '순번',
      1
    );
    row = Math.max(sheet.getLastRow() + 1, FIRST_DATA_ROW);
    previousRow = Math.max(FIRST_DATA_ROW, row - 1);
    previousRowHeight = sheet.getRowHeight(previousRow);

    if (row > FIRST_DATA_ROW) {
      sheet
        .getRange(
          previousRow,
          1,
          1,
          sheet.getLastColumn()
        )
        .copyTo(
          sheet.getRange(
            row,
            1,
            1,
            sheet.getLastColumn()
          ),
          SpreadsheetApp.CopyPasteType.PASTE_FORMAT,
          false
        );

      sheet.setRowHeight(row, previousRowHeight);

      sequenceValues = sheet
        .getRange(FIRST_DATA_ROW, sequenceColumn, row - FIRST_DATA_ROW, 1)
        .getValues()
        .flat();

      for (
        sequenceIndex = 0;
        sequenceIndex < sequenceValues.length;
        sequenceIndex++
      ) {
        var sequenceNumber = Number(
          sequenceValues[sequenceIndex]
        );

        if (
          isFinite(sequenceNumber) &&
          sequenceNumber >= nextSequence
        ) {
          nextSequence = sequenceNumber + 1;
        }
      }
    }

    sheet
      .getRange(row, sequenceColumn)
      .setValue(nextSequence);

    HR_setSafeText_(
      sheet.getRange(row, applicantColumn),
      applicant
    );
    sheet
      .getRange(row, HR_ensureAudienceColumn_(sheet))
      .setValue(HR_audienceLabel_(requestAudience));

    // 검색 화면에서 이미 받은 YES24 결과를 그대로 사용합니다.
    // 같은 책을 다시 API로 조회하지 않아 접수 시간을 줄입니다.
    fastBook = {
      title: selectedBook.title || (!isIsbn ? query : ''),
      author: selectedBook.author || '',
      publisher: selectedBook.publisher || '',
      isbn: selectedBook.isbn || normalizedIsbn,
      cover: selectedBook.cover || '',
      priceStd: selectedPrice,
      year: selectedYear ? selectedYear[0] : '',
      category: selectedBook.category || '',
      callNumber: selectedBook.callNumber || ''
    };

    fillFinalData(
      sheet,
      row,
      fastBook,
      headers,
      '✅ 접수 완료' + (ownedBook.found && isManagerApplicant ? ' · 관리자 소장도서 예외' : '') + HR_unverifiedNote_(selectedBook)
    );

    // 웹 신청에서도 정가와 출판연도가 반드시 시트에 남도록
    // 실제 머리글을 다시 찾아 직접 기록합니다.
    priceColumn = HR_findAnyHeaderColumn_(
      headers,
      ['정가', '도서정가', '정가(원)', '가격']
    );
    yearColumn = HR_findAnyHeaderColumn_(
      headers,
      ['연도', '출판연도', '발행연도', '출간연도']
    );

    if (priceColumn && selectedPrice > 0) {
      sheet
        .getRange(row, priceColumn)
        .setValue(selectedPrice)
        .setNumberFormat('#,##0');
    }

    if (yearColumn && fastBook.year) {
      sheet
        .getRange(row, yearColumn)
        .setValue(fastBook.year)
        .setNumberFormat('0');
    }

    SpreadsheetApp.flush();

    responseResult = {
      success: true,
      row: row,
      title: fastBook.title || query,
      isbn: fastBook.isbn || normalizedIsbn,
      author: fastBook.author,
      publisher: fastBook.publisher,
      price: selectedPrice,
      year: fastBook.year,
      status: '✅ 접수 완료' + (ownedBook.found && isManagerApplicant ? ' · 관리자 소장도서 예외' : '') + HR_unverifiedNote_(selectedBook),
      requestId: requestId
    };

    HR_updateRequestBackup_(
      spreadsheet,
      backupRow,
      '처리 완료' + (ownedBook.found && isManagerApplicant ? ' · 관리자 소장도서 예외' : ''),
      row,
      ''
    );

    if (isManagerApplicant) {
      try {
        sendHopeRequestsToReview(row);
        buildUnifiedBookstoreExportSheet();
      } catch (routeError) {
        console.error('관리자 희망도서 자동 전달 오류', routeError);
      }
    }

    return responseResult;
  } catch (error) {
    if (
      row > 0 &&
      columns.status
    ) {
      sheet
        .getRange(row, columns.status)
        .setValue(
          '❌ 오류: ' +
          String(error.message || error)
        );
    }

    if (backupRow > 0) {
      HR_updateRequestBackup_(
        spreadsheet,
        backupRow,
        '접수 오류',
        row,
        String(error.message || error)
      );
    }

    throw error;
  } finally {
    if (lockAcquired) {
      documentLock.releaseLock();
    }
  }
}

/**
 * 신청 원본을 별도 숨김 시트에 남깁니다.
 * 자동입력 도중 오류가 나도 신청자·제목·ISBN을 복구할 수 있습니다.
 */
function HR_appendRequestBackup_(spreadsheet, requestData) {
  var sheet = spreadsheet.getSheetByName(
    HR_REQUEST_BACKUP_SHEET_NAME
  );
  var headers = [[
    '접수시각',
    '요청ID',
    '신청자',
    '검색어',
    '도서명',
    'ISBN13',
    '처리상태',
    '희망도서행',
    '오류내용'
  ]];
  var row;

  if (!sheet) {
    sheet = spreadsheet.insertSheet(
      HR_REQUEST_BACKUP_SHEET_NAME
    );
    sheet.getRange(1, 1, 1, headers[0].length)
      .setValues(headers)
      .setFontWeight('bold')
      .setBackground('#E8EEF8');
        sheet.setFrozenRows(1);
    sheet.hideSheet();
      }


  row = Math.max(sheet.getLastRow() + 1, 2);
  // 신청자·검색어·도서명·ISBN은 텍스트로 저장해 수식으로 실행되지 않게 합니다.
  sheet.getRange(row, 1, 1, 9).setValues([[
    new Date(),
    '',
    '',
    '',
    '',
    '',
    requestData.status || '접수',
    '',
    ''
  ]]);
  HR_setSafeText_(sheet.getRange(row, 2), requestData.requestId);
  HR_setSafeText_(sheet.getRange(row, 3), requestData.applicant);
  HR_setSafeText_(sheet.getRange(row, 4), requestData.query);
  HR_setSafeText_(sheet.getRange(row, 5), requestData.title);
  HR_setSafeText_(sheet.getRange(row, 6), requestData.isbn);
  sheet.getRange(row, 1).setNumberFormat(
    'yyyy. m. d. HH:mm:ss'
  );
  return row;
}

function HR_updateRequestBackup_(
  spreadsheet,
  backupRow,
  status,
  targetRow,
  errorMessage
) {
  var sheet;
  if (!backupRow) return;
  sheet = spreadsheet.getSheetByName(
    HR_REQUEST_BACKUP_SHEET_NAME
  );
  if (!sheet) return;

  sheet.getRange(backupRow, 7, 1, 2).setValues([[
    status || '',
    targetRow || ''
  ]]);
  HR_setSafeText_(sheet.getRange(backupRow, 9), errorMessage);
}

/**
 * 장바구니에서 고른 여러 권을 순서대로 신청합니다.
 * 한 번에 최대 20권까지 처리하고, 성공·중복·소장 도서를 각각 알려줍니다.
 */
function HR_submitBookRequests(formData) {
  var applicant = String(
    formData && formData.applicant || ''
  ).trim();
  var selectedBooks =
    formData && Array.isArray(formData.books)
      ? formData.books
      : [];
  var requestAudience = HR_audienceKey_(formData && formData.audience);
  var results = [];
  var bookIndex = 0;
  var successCount = 0;
  var duplicateCount = 0;
  var ownedCount = 0;
  var errorCount = 0;
  var requestIdBase = String(
    formData && formData.requestId || Utilities.getUuid()
  ).trim();

  if (!selectedBooks.length) {
    throw new Error('신청할 책을 한 권 이상 선택해 주세요.');
  }

  if (selectedBooks.length > 20) {
    throw new Error(
      '한 번에 최대 20권까지 신청할 수 있습니다.'
    );
  }

  if (requestAudience === 'parent' && !HR_parseParentApplicant_(applicant)) {
    throw new Error('학년·반·번호를 모두 입력해 주세요. (예: 3학년 2반 15번)');
  }

  var limitMessage = HR_requestLimitMessage_(
    requestAudience,
    applicant,
    selectedBooks.length
  );
  if (limitMessage) {
    throw new Error(limitMessage);
  }

  for (
    bookIndex = 0;
    bookIndex < selectedBooks.length;
    bookIndex++
  ) {
    var selectedBook = selectedBooks[bookIndex] || {};
    var result;

    try {
      result = HR_submitBookRequest({
        applicant: applicant,
        audience: requestAudience,
        query: selectedBook.isbn || selectedBook.title,
        book: selectedBook,
        requestId: requestIdBase + '-' + bookIndex
      });

      if (result.success) {
        successCount++;
      } else if (result.owned) {
        ownedCount++;
      } else if (result.duplicateRequest) {
        duplicateCount++;
      } else {
        errorCount++;
      }

      results.push(result);
    } catch (error) {
      errorCount++;
      results.push({
        success: false,
        error: true,
        title: selectedBook.title || '제목 없음',
        isbn: selectedBook.isbn || '',
        message: String(error.message || error)
      });
    }
  }

  return {
    success: successCount > 0,
    successCount: successCount,
    duplicateCount: duplicateCount,
    ownedCount: ownedCount,
    errorCount: errorCount,
    results: results
  };
}

/**
 * YES24에서 책 목록을 검색합니다.
 * 제목 검색은 최대 50권, ISBN 검색은 해당 도서 한 권을 반환합니다.
 */
function HR_searchRequestBooks(query) {
  var keyword = String(query || '').trim();
  if (!keyword) throw new Error('검색할 책 제목이나 ISBN을 입력해 주세요.');
  var isIsbn = /^[0-9Xx\-\s]+$/.test(keyword) && /^(?:[0-9Xx]{10}|[0-9]{13})$/.test(keyword.replace(/[\s-]/g, ''));
  var key = HR_getYes24Key_();
  var items = yes24SearchItems_(keyword, key, 'BOOK', 50, isIsbn);
  var setPattern = /(?:세트|전집|박스(?:세트)?|묶음|전\s*\d+\s*(?:권|종)|총\s*\d+\s*(?:권|종)|\d+\s*(?:~|～|∼|-)\s*\d+\s*권|\d+\s*권\s*(?:세트|완결)|전권|합본)/i;
  function singles(list) {
    return list.filter(function(item) {
      var title = HR_cleanYes24Text_(item && item.title || '');
      var publisher = HR_cleanYes24Text_(item && item.publisher || '');
      var category = HR_cleanYes24Text_(item && item.goodsSortNm || '');
      return !setPattern.test(title) && !/(?:전집|세트)/i.test(publisher + ' ' + category);
    });
  }
  items = singles(items);
  // 권차를 적었다면 해당 단권만 보여 줍니다. 붙여 쓴 시리즈명도 비교합니다.
  var volumeMatch = !isIsbn && keyword.match(/^(.+?)\s+(\d{1,3})\s*(?:권|편|호)?$/);
  if (volumeMatch) {
    var base = volumeMatch[1].trim();
    var baseKey = base.replace(/\s+/g, '').toLowerCase();
    var volume = volumeMatch[2];
    function sameVolume(item) {
      var title = HR_cleanYes24Text_(item && item.title || '').replace(/\s+/g, '').toLowerCase();
      var at = title.indexOf(baseKey);
      return at >= 0 && new RegExp('^' + volume + '(?!\\d)').test(title.slice(at + baseKey.length));
    }
    var matched = items.filter(sameVolume);
    if (!matched.length) {
      matched = singles(yes24SearchItems_(base, key, 'BOOK', 100, false)).filter(sameVolume);
    }
    items = matched;
  } else if (!isIsbn && items.length <= 3) {
    // 띄어 쓰지 않은 시리즈명으로 검색된 단권의 실제 제목을 이용해 다른 권차도 찾습니다.
    var compact = keyword.replace(/\s+/g, '').toLowerCase();
    var sample = items.map(function(item) {
      return HR_cleanYes24Text_(item && item.title || '').match(/^(.+?)\s+\d{1,3}(?:\s|[:：.]|$)/);
    }).filter(Boolean)[0];
    if (sample && sample[1].replace(/\s+/g, '').toLowerCase() === compact && sample[1] !== keyword) {
      var seriesKey = compact;
      items = singles(yes24SearchItems_(sample[1], key, 'BOOK', 100, false)).filter(function(item) {
        var title = HR_cleanYes24Text_(item && item.title || '').replace(/\s+/g, '').toLowerCase();
        return title.indexOf(seriesKey) === 0 && new RegExp('^\\d{1,3}(?!\\d)').test(title.slice(seriesKey.length));
      });
    }
  }
  return items.map(yes24RequestBook_);
}

/**
 * YES24 텍스트를 화면 표시용으로 정리합니다.
 */
function HR_cleanYes24Text_(value) {
  return String(value || '')
    .replace(/<[^>]*>/g, '')
    .replace(/&nbsp;/g, ' ')
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
    .replace(/\s+/g, ' ')
    .trim();
}

/**
 * 검색 결과 표지를 큰 이미지 주소로 정리합니다.
 */
function HR_largeYes24Cover_(value) {
  return HR_safeUrl_(value)
    
    ;
}

/**
 * 외부 링크를 안전한 HTTP 주소만 허용합니다.
 */
function HR_safeUrl_(value) {
  var text = String(value || '').trim();

  return /^https?:\/\//i.test(text)
    ? text
    : '';
}

/**
 * 우리학교 소장도서에서 ISBN을 우선 확인하고,
 * ISBN이 없을 때만 제목으로 중복을 확인합니다.
 */
function HR_findOwnedBook_(spreadsheet, book) {
  var librarySheet = spreadsheet.getSheetByName(
    HR_LIBRARY_SHEET_NAME
  );
  var result = {
    found: false,
    title: '',
    isbn: ''
  };
  var lastRow = 0;
  var lastColumn = 0;
  var headers = [];
  var titleColumn = 0;
  var isbnColumn = 0;
  var values = [];
  var targetTitle = HR_normalizeBookTitle_(
    book && book.title
  );
  var targetIsbn = HR_normalizeIsbn_(
    book && book.isbn
  );
  var rowIndex = 0;

  if (!librarySheet) {
    return result;
  }

  lastRow = librarySheet.getLastRow();
  lastColumn = librarySheet.getLastColumn();

  if (lastRow < 2 || lastColumn < 1) {
    return result;
  }

  headers = librarySheet
    .getRange(1, 1, 1, lastColumn)
    .getDisplayValues()[0];

  titleColumn = HR_findAnyHeaderColumn_(
    headers,
    ['책제목', '도서명', '서명', '제목']
  );
  isbnColumn = HR_findAnyHeaderColumn_(
    headers,
    ['ISBN13', 'ISBN']
  );

  if (!titleColumn && !isbnColumn) {
    return result;
  }

  values = librarySheet
    .getRange(2, 1, lastRow - 1, lastColumn)
    .getDisplayValues();

  for (rowIndex = 0; rowIndex < values.length; rowIndex++) {
    var libraryTitle = titleColumn
      ? String(values[rowIndex][titleColumn - 1] || '').trim()
      : '';
    var libraryIsbn = isbnColumn
      ? HR_normalizeIsbn_(
          values[rowIndex][isbnColumn - 1]
        )
      : '';
    var isbnMatches = Boolean(
      targetIsbn &&
      libraryIsbn &&
      targetIsbn === libraryIsbn
    );
    var titleMatches = Boolean(
      !targetIsbn &&
      targetTitle &&
      libraryTitle &&
      targetTitle === HR_normalizeBookTitle_(libraryTitle)
    );

    if (isbnMatches || titleMatches) {
      return {
        found: true,
        title: libraryTitle,
        isbn: libraryIsbn
      };
    }
  }

  return result;
}

/**
 * 희망도서 목록에 같은 책이 이미 신청되어 있는지 확인합니다.
 * ISBN이 있으면 ISBN을 우선 사용하고, ISBN이 없을 때만 제목을 비교합니다.
 */
function HR_findRequestedBook_(sheet, book, headers, columns) {
  var result = {
    found: false,
    title: '',
    isbn: '',
    row: 0
  };
  var lastRow = sheet.getLastRow();
  var titleColumn = columns && columns.title
    ? columns.title
    : HR_findAnyHeaderColumn_(
        headers || [],
        ['도서명', '책제목', '서명', '제목']
      );
  var isbnColumn = columns && columns.isbn
    ? columns.isbn
    : HR_findAnyHeaderColumn_(
        headers || [],
        ['ISBN13', 'ISBN']
      );
  var targetTitle = HR_normalizeBookTitle_(
    book && book.title
  );
  var targetIsbn = HR_normalizeIsbn_(
    book && book.isbn
  );
  var values = [];
  var rowIndex = 0;

  if (
    lastRow < FIRST_DATA_ROW ||
    (!titleColumn && !isbnColumn)
  ) {
    return result;
  }

  values = sheet
    .getRange(
      FIRST_DATA_ROW,
      1,
      lastRow - FIRST_DATA_ROW + 1,
      sheet.getLastColumn()
    )
    .getDisplayValues();

  for (rowIndex = 0; rowIndex < values.length; rowIndex++) {
    var requestedTitle = titleColumn
      ? String(values[rowIndex][titleColumn - 1] || '').trim()
      : '';
    var requestedIsbn = isbnColumn
      ? HR_normalizeIsbn_(
          values[rowIndex][isbnColumn - 1]
        )
      : '';
    var isbnMatches = Boolean(
      targetIsbn &&
      requestedIsbn &&
      targetIsbn === requestedIsbn
    );
    var titleMatches = Boolean(
      !targetIsbn &&
      targetTitle &&
      requestedTitle &&
      targetTitle === HR_normalizeBookTitle_(requestedTitle)
    );

    if (isbnMatches || titleMatches) {
      return {
        found: true,
        title: requestedTitle,
        isbn: requestedIsbn,
        row: rowIndex + FIRST_DATA_ROW
      };
    }
  }

  return result;
}

function HR_findAnyHeaderColumn_(headers, candidates) {
  var headerIndex = 0;
  var candidateIndex = 0;

  for (
    headerIndex = 0;
    headerIndex < headers.length;
    headerIndex++
  ) {
    var normalizedHeader = String(
      headers[headerIndex] || ''
    )
      .replace(/\s+/g, '')
      .toUpperCase();

    for (
      candidateIndex = 0;
      candidateIndex < candidates.length;
      candidateIndex++
    ) {
      var normalizedCandidate = String(
        candidates[candidateIndex] || ''
      )
        .replace(/\s+/g, '')
        .toUpperCase();

      if (
        normalizedHeader === normalizedCandidate ||
        normalizedHeader.indexOf(normalizedCandidate) !== -1
      ) {
        return headerIndex + 1;
      }
    }
  }

  return 0;
}

function HR_normalizeIsbn_(value) {
  return String(value || '')
    .replace(/[^0-9Xx]/g, '')
    .toUpperCase();
}

function HR_normalizeBookTitle_(value) {
  return String(value || '')
    .replace(/\([^)]*\)/g, '')
    .replace(/\[[^\]]*\]/g, '')
    .replace(/[^0-9A-Za-z가-힣]/g, '')
    .toLowerCase();
}

/**
 * 헤더 이름으로 열 번호를 찾습니다.
 */
function HR_findHeaderColumn_(
  headers,
  targetHeader,
  fallbackColumn
) {
  var headerIndex = 0;
  var normalizedTarget = String(targetHeader || '')
    .replace(/\s+/g, '');

  for (
    headerIndex = 0;
    headerIndex < headers.length;
    headerIndex++
  ) {
    var normalizedHeader = String(
      headers[headerIndex] || ''
    ).replace(/\s+/g, '');

    if (
      normalizedHeader.indexOf(normalizedTarget) !== -1
    ) {
      return headerIndex + 1;
    }
  }

  return fallbackColumn;
}

/**
 * 사용설명서의 핵심 안내를 한 화면에 보이도록 정리합니다.
 */
function HR_formatGuideSheet() {
  var spreadsheet = HR_getSetupSpreadsheet_();
  var sheet = spreadsheet.getSheetByName('📗 희망도서 신청안내');
  var webAppUrl = HR_getConfiguredPublicWebAppUrl_();
  var shortcutUrl = webAppUrl || (
    'https://docs.google.com/spreadsheets/d/' +
    spreadsheet.getId() +
    '/edit'
  );
  var titleColor = '#243B64';
  var accentColor = '#5277C7';
  var textColor = '#27344A';
  var lineColor = '#D8E1F0';
  var steps;

  if (!sheet) {
    throw new Error('사용설명서 시트를 찾지 못했습니다.');
  }

  sheet.setHiddenGridlines(true);
  sheet.setColumnWidth(1, 70);
  sheet.setColumnWidths(2, 2, 52);
  sheet.setColumnWidths(4, 9, 103);
  sheet.setFrozenRows(0);
  sheet.setFrozenColumns(0);
  HR_breakMergedRangesSafely_(sheet);
  sheet.getRange('B3:L50').clear();
  sheet
    .getRange('B3:L50')
    .setBackground('#FFFFFF')
    .setFontFamily('Jua')
    .setFontColor(textColor)
    .setVerticalAlignment('middle')
    .setWrap(true);

  sheet.getRange('B3:L4').merge();
  sheet
    .getRange('B3')
    .setValue('📚 희망도서 신청 안내');
  sheet.getRange('B5:L5').merge();
  sheet
    .getRange('B5')
    .setValue('신청 페이지 열기  →  도서 검색  →  여러 권 선택  →  한 번에 신청');

  steps = [
    {
      badge: 'B7:C8',
      text: 'D7:L8',
      number: '1',
      value: '아래 「희망도서 신청하기」 버튼을 누릅니다.'
    },
    {
      badge: 'B10:C11',
      text: 'D10:L11',
      number: '2',
      value: '책 제목 또는 ISBN을 입력해 원하는 책을 검색합니다.'
    },
    {
      badge: 'B13:C14',
      text: 'D13:L14',
      number: '3',
      value:
        '표지·저자·연도·정가를 확인하고 신청할 책을 여러 권 선택한 뒤 「선택한 도서 신청하기」를 누릅니다.'
    },
    {
      badge: 'B16:C17',
      text: 'D16:L17',
      number: '4',
      value:
        '신청한 도서는 「희망도서 목록」 시트에서 확인할 수 있습니다.'
    }
  ];

  HR_drawGuideSteps_(
    sheet,
    steps,
    accentColor,
    '#F3F6FC',
    lineColor
  );

  sheet.getRange('B19:L20').merge();
  sheet
    .getRange('B19')
    .setValue(
      'ℹ️ 신청자 이름은 선택사항이며, 입력하면 우선 대출 혜택을 받을 수 있습니다. 한 번에 최대 20권까지 신청할 수 있습니다.'
    )
    .setBackground('#EDF2FA')
    .setFontColor(titleColor)
    .setFontSize(14)
    .setHorizontalAlignment('left');

  sheet.getRange('B22:L24').merge();
  sheet
    .getRange('B22')
    .setFormula(
      '=HYPERLINK("' +
      shortcutUrl +
      '","📚 희망도서 신청하기")'
    )
    .setBackground(accentColor)
    .setFontColor('#FFFFFF')
    .setFontSize(19)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  sheet
    .getRange('B3:L4')
    .setBackground(titleColor)
    .setFontColor('#FFFFFF')
    .setFontSize(27)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  sheet
    .getRange('B5:L5')
    .setBackground('#E7EDF8')
    .setFontColor(titleColor)
    .setFontSize(14)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  [3, 4].forEach(function(rowNumber) {
    sheet.setRowHeight(rowNumber, 38);
  });
  sheet.setRowHeight(5, 34);
  [7, 8, 10, 11, 13, 14, 16, 17].forEach(function(rowNumber) {
    sheet.setRowHeight(rowNumber, 34);
  });
  [6, 9, 12, 15, 18, 21, 25, 26]
    .forEach(function(rowNumber) {
      sheet.setRowHeight(rowNumber, 10);
    });
  [19, 20].forEach(function(rowNumber) {
    sheet.setRowHeight(rowNumber, 31);
  });
  [22, 23, 24].forEach(function(rowNumber) {
    sheet.setRowHeight(rowNumber, 34);
  });

  sheet.activate();
  sheet.getRange('A3').activate();
}

function HR_formatGuideSheetAdminFirst_() {
  var spreadsheet = HR_getSetupSpreadsheet_();
  var sheet = spreadsheet.getSheetByName('📗 희망도서 신청안내');
  var webAppUrl = HR_getConfiguredPublicWebAppUrl_();
  var shortcutUrl = webAppUrl || (
    'https://docs.google.com/spreadsheets/d/' +
    spreadsheet.getId() +
    '/edit'
  );
  var titleColor = '#243B64';
  var adminColor = '#7A5128';
  var adminAccent = '#DDA333';
  var userAccent = '#5277C7';
  var textColor = '#27344A';
  var lineColor = '#D8E1F0';
  var adminSteps;
  var userSteps;

  if (!sheet) {
    throw new Error('사용설명서 시트를 찾지 못했습니다.');
  }

  sheet.setHiddenGridlines(true);
  sheet.setColumnWidth(1, 70);
  sheet.setColumnWidths(2, 2, 52);
  sheet.setColumnWidths(4, 9, 103);
  sheet.setFrozenRows(0);
  sheet.setFrozenColumns(0);
  HR_breakMergedRangesSafely_(sheet);
  
  
  sheet.getRange('B3:L50').clear();
  sheet
    .getRange('B3:L50')
    .setBackground('#FFFFFF')
    .setFontFamily('Jua')
    .setFontColor(textColor)
    .setVerticalAlignment('middle')
    .setWrap(true);

  sheet.getRange('B3:L4').merge();
  sheet
    .getRange('B3')
    .setValue('🏫 다른 학교에서 사본으로 사용하는 방법');
  sheet.getRange('B5:L5').merge();
  sheet
    .getRange('B5')
    .setValue('처음 한 번만 1번부터 차례대로 따라 하세요.');

  adminSteps = [
    {
      badge: 'B7:C8',
      text: 'D7:L8',
      number: '1',
      value:
        '원본 시트에서 「파일 → 사본 만들기」를 누르고 새 학교 이름으로 저장합니다.'
    },
    {
      badge: 'B10:C11',
      text: 'D10:L11',
      number: '2',
      value:
        '사본의 「⚙️ 사본 최초설정」 탭에서 API 키 설정을 누릅니다. YES24은 필수, 국립중앙도서관은 선택입니다.'
    },
    {
      badge: 'B13:C14',
      text: 'D13:L14',
      number: '3',
      value:
        '「이 학교 최초 설정」을 누르고 학교 이름을 입력한 뒤 실행합니다.'
    },
    {
      badge: 'B16:C17',
      text: 'D16:L17',
      number: '4',
      value:
        '「우리학교 소장도서」 시트에 책 제목과 ISBN13을 붙여넣습니다.'
    },
    {
      badge: 'B19:C20',
      text: 'D19:L20',
      number: '5',
      value:
        '확장 프로그램 → Apps Script → 배포 → 새 배포 → 웹 앱에서 「나 / 모든 사용자」로 배포합니다. 완료 후 설정 탭은 삭제해도 됩니다.'
    }
  ];

  HR_drawGuideSteps_(
    sheet,
    adminSteps,
    adminAccent,
    '#FFF9EF',
    '#E8D2AB'
  );

  sheet.getRange('B22:L24').merge();
  sheet
    .getRange('B22')
    .setValue(
      '꼭 확인하세요! 「우리학교 소장도서」에 책 제목과 ISBN13을 모두 넣어야 “이미 소장중인 도서입니다.” 안내가 정확히 나옵니다.'
    )
    .setBackground('#FFF0F0')
    .setFontColor('#A33C3C')
    .setFontSize(14)
    .setFontWeight('bold')
    .setHorizontalAlignment('left');

  sheet.getRange('B26:L27').merge();
  sheet
    .getRange('B26')
    .setValue(
      'API 안내  ·  YES24은 도서 검색에 꼭 필요합니다. 국립중앙도서관 API는 KDC 참고용이라 입력하지 않아도 됩니다.'
    )
    .setBackground('#EDF4FF')
    .setFontColor('#284E86')
    .setFontSize(14)
    .setFontWeight('bold')
    .setHorizontalAlignment('left');

  sheet.getRange('B30:L31').merge();
  sheet
    .getRange('B30')
    .setValue('📚 희망도서 신청 방법');
  sheet.getRange('B32:L32').merge();
  sheet
    .getRange('B32')
    .setValue('신청 페이지 열기  →  도서 검색  →  원하는 책 선택');

  userSteps = [
    {
      badge: 'B34:C35',
      text: 'D34:L35',
      number: '1',
      value: '아래 「희망도서 신청하기」 버튼을 누릅니다.'
    },
    {
      badge: 'B37:C38',
      text: 'D37:L38',
      number: '2',
      value: '책 제목 또는 ISBN을 입력해 원하는 책을 검색합니다.'
    },
    {
      badge: 'B40:C41',
      text: 'D40:L41',
      number: '3',
      value:
        '표지·저자·연도·정가를 확인한 뒤 「이 도서 선택하기」를 누릅니다.'
    }
  ];

  HR_drawGuideSteps_(
    sheet,
    userSteps,
    userAccent,
    '#F3F6FC',
    lineColor
  );

  sheet.getRange('B43:L44').merge();
  sheet
    .getRange('B43')
    .setValue(
      'ℹ️ 신청자 이름은 선택사항이며, 입력하면 우선 대출 혜택을 받을 수 있습니다. 시리즈는 여러 권을 차례로 선택할 수 있습니다.'
    )
    .setBackground('#EDF2FA')
    .setFontColor(titleColor)
    .setFontSize(14)
    .setHorizontalAlignment('left');

  sheet.getRange('B46:L48').merge();
  sheet
    .getRange('B46')
    .setFormula(
      '=HYPERLINK("' +
      shortcutUrl +
      '","📚 희망도서 신청하기")'
    )
    .setBackground(userAccent)
    .setFontColor('#FFFFFF')
    .setFontSize(19)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  sheet
    .getRange('B3:L4')
    .setBackground(adminColor)
    .setFontColor('#FFFFFF')
    .setFontSize(27)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  sheet
    .getRange('B5:L5')
    .setBackground('#F3E9DA')
    .setFontColor(adminColor)
    .setFontSize(14)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  sheet
    .getRange('B30:L31')
    .setBackground(titleColor)
    .setFontColor('#FFFFFF')
    .setFontSize(25)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  sheet
    .getRange('B32:L32')
    .setBackground('#E7EDF8')
    .setFontColor(titleColor)
    .setFontSize(14)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  [3, 4, 30, 31].forEach(function(rowNumber) {
    sheet.setRowHeight(rowNumber, 38);
  });
  [5, 32].forEach(function(rowNumber) {
    sheet.setRowHeight(rowNumber, 34);
  });
  [
    7, 8, 10, 11, 13, 14, 16, 17, 19, 20,
    34, 35, 37, 38, 40, 41
  ].forEach(function(rowNumber) {
    sheet.setRowHeight(rowNumber, 34);
  });
  [6, 9, 12, 15, 18, 21, 25, 28, 29, 33, 36, 39, 42, 45, 49, 50]
    .forEach(function(rowNumber) {
      sheet.setRowHeight(rowNumber, 10);
    });
  [22, 23, 24].forEach(function(rowNumber) {
    sheet.setRowHeight(rowNumber, 28);
  });
  [26, 27, 43, 44].forEach(function(rowNumber) {
    sheet.setRowHeight(rowNumber, 31);
  });
  [46, 47, 48].forEach(function(rowNumber) {
    sheet.setRowHeight(rowNumber, 34);
  });

  sheet.activate();
  sheet.getRange('A3').activate();
}

function HR_drawGuideSteps_(
  sheet,
  steps,
  badgeColor,
  textBackground,
  borderColor
) {
  steps.forEach(function(step) {
    sheet.getRange(step.badge).merge();
    sheet.getRange(step.text).merge();
    sheet
      .getRange(step.badge)
      .setBackground(badgeColor)
      .setFontColor('#FFFFFF')
      .setFontSize(24)
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setBorder(
        true,
        true,
        true,
        true,
        false,
        false,
        badgeColor,
        SpreadsheetApp.BorderStyle.SOLID
      );
    sheet
      .getRange(step.badge.split(':')[0])
      .setValue(step.number);
    sheet
      .getRange(step.text)
      .setBackground(textBackground)
      .setFontColor('#27344A')
      .setFontSize(15)
      .setFontWeight('bold')
      .setHorizontalAlignment('left')
      .setBorder(
        true,
        true,
        true,
        true,
        false,
        false,
        borderColor,
        SpreadsheetApp.BorderStyle.SOLID
      );
    sheet
      .getRange(step.text.split(':')[0])
      .setValue(step.value);
  });
}

function HR_formatGuideSheetLegacy_() {
  var spreadsheet = HR_getSpreadsheet_();
  var sheet = spreadsheet.getSheetByName('📗 희망도서 신청안내');
  var webAppUrl = HR_getConfiguredPublicWebAppUrl_();
  var shortcutUrl = webAppUrl || (
    'https://docs.google.com/spreadsheets/d/' +
    spreadsheet.getId() +
    '/edit?gid=0#gid=0'
  );

  if (!sheet) {
    throw new Error('사용설명서 시트를 찾지 못했습니다.');
  }

  var guideRange = sheet.getRange('B3:L46');
  var titleColor = '#243B64';
  var accentColor = '#5277C7';
  var paleColor = '#F3F6FC';
  var textColor = '#27344A';
  var lineColor = '#D8E1F0';

  sheet.setHiddenGridlines(true);
  sheet.setColumnWidth(1, 70);
  sheet.setColumnWidths(2, 2, 52);
  sheet.setColumnWidths(4, 9, 103);

  guideRange.breakApart();
  guideRange.clear();
  guideRange
    .setBackground('#FFFFFF')
    .setFontFamily('Jua')
    .setFontColor(textColor)
    .setVerticalAlignment('middle')
    .setWrap(true);

  sheet.getRange('B3:L4').merge();
  sheet
    .getRange('B3')
    .setValue('📚 희망도서 신청 안내');

  sheet.getRange('B5:L5').merge();
  sheet
    .getRange('B5')
    .setValue('신청 페이지 열기  →  도서 검색  →  원하는 책 선택');

  var steps = [
    {
      badgeRange: 'B7:C8',
      textRange: 'D7:L8',
      number: '1',
      text: '아래 「희망도서 신청하기」 버튼을 누릅니다.'
    },
    {
      badgeRange: 'B10:C11',
      textRange: 'D10:L11',
      number: '2',
      text: '책 제목 또는 ISBN을 입력해 원하는 책을 검색합니다.'
    },
    {
      badgeRange: 'B13:C14',
      textRange: 'D13:L14',
      number: '3',
      text: '표지·저자·연도·정가를 확인한 뒤 「이 도서 선택하기」를 누릅니다.'
    }
  ];

  steps.forEach(function(step) {
    sheet.getRange(step.badgeRange).merge();
    sheet.getRange(step.textRange).merge();

    sheet
      .getRange(step.badgeRange)
      .setBackground(accentColor)
      .setFontColor('#FFFFFF')
      .setFontSize(24)
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setBorder(
        true,
        true,
        true,
        true,
        false,
        false,
        accentColor,
        SpreadsheetApp.BorderStyle.SOLID
      );

    sheet
      .getRange(step.badgeRange.split(':')[0])
      .setValue(step.number);

    sheet
      .getRange(step.textRange)
      .setBackground(paleColor)
      .setFontColor(textColor)
      .setFontSize(15)
      .setFontWeight('bold')
      .setHorizontalAlignment('left')
      .setBorder(
        true,
        true,
        true,
        true,
        false,
        false,
        lineColor,
        SpreadsheetApp.BorderStyle.SOLID
      );

    sheet
      .getRange(step.textRange.split(':')[0])
      .setValue(step.text);
  });

  sheet.getRange('B16:L17').merge();
  sheet
    .getRange('B16')
    .setValue(
      'ℹ️ 신청자 이름은 선택사항이며, 입력하면 우선 대출 혜택을 받을 수 있습니다. 시리즈는 여러 권을 차례로 선택할 수 있습니다.'
    );

  sheet.getRange('B19:L20').merge();
  sheet
    .getRange('B19')
    .setFormula(
      '=HYPERLINK("' +
      shortcutUrl +
      '","📚 희망도서 신청하기")'
    );

  sheet.getRange('B22:L23').merge();
  sheet
    .getRange('B22')
    .setValue('🏫 다른 학교에서 사본으로 사용하는 방법');

  sheet.getRange('B25:C26').merge();
  sheet.getRange('D25:L26').merge();
  sheet.getRange('B28:C29').merge();
  sheet.getRange('D28:L29').merge();
  sheet.getRange('B31:C32').merge();
  sheet.getRange('D31:L32').merge();
  sheet.getRange('B34:C35').merge();
  sheet.getRange('D34:L35').merge();
  sheet.getRange('B37:C38').merge();
  sheet.getRange('D37:L38').merge();

  sheet.getRange('B25').setValue('1');
  sheet
    .getRange('D25')
    .setValue(
      '원본 시트에서 「파일 → 사본 만들기」를 누르고 새 학교 이름으로 저장합니다.'
    );
  sheet.getRange('B28').setValue('2');
  sheet
    .getRange('D28')
    .setValue(
      '사본의 「⚙️ 사본 최초설정」 탭에서 API 키 설정을 누릅니다. YES24은 필수, 국립중앙도서관은 선택입니다.'
    );
  sheet.getRange('B31').setValue('3');
  sheet
    .getRange('D31')
    .setValue(
      '「이 학교 최초 설정」을 누르고 학교 이름을 입력한 뒤 실행합니다.'
    );
  sheet.getRange('B34').setValue('4');
  sheet
    .getRange('D34')
    .setValue(
      '「우리학교 소장도서」 시트에 책 제목과 ISBN13을 붙여넣습니다.'
    );
  sheet.getRange('B37').setValue('5');
  sheet
    .getRange('D37')
    .setValue(
      '확장 프로그램 → Apps Script → 배포 → 새 배포 → 웹 앱에서 「나 / 모든 사용자」로 배포합니다. 완료 후 설정 탭은 삭제해도 됩니다.'
    );

  [
    'B25:C26',
    'B28:C29',
    'B31:C32',
    'B34:C35',
    'B37:C38'
  ].forEach(
    function(rangeA1) {
      sheet
        .getRange(rangeA1)
        .setBackground('#D99A32')
        .setFontColor('#FFFFFF')
        .setFontSize(22)
        .setFontWeight('bold')
        .setHorizontalAlignment('center');
    }
  );

  [
    'D25:L26',
    'D28:L29',
    'D31:L32',
    'D34:L35',
    'D37:L38'
  ].forEach(
    function(rangeA1) {
      sheet
        .getRange(rangeA1)
        .setBackground('#FFF8EC')
        .setFontColor(textColor)
        .setFontSize(14)
        .setFontWeight('bold')
        .setHorizontalAlignment('left')
        .setBorder(
          true,
          true,
          true,
          true,
          false,
          false,
          '#E9D4AE',
          SpreadsheetApp.BorderStyle.SOLID
        );
    }
  );

  sheet.getRange('B41:L43').merge();
  sheet
    .getRange('B41')
    .setValue(
      '꼭 확인하세요! 「우리학교 소장도서」에 책 제목과 ISBN13을 모두 넣어야 “이미 소장중인 도서입니다.” 안내가 정확히 나옵니다.'
    );

  sheet.getRange('B45:L46').merge();
  sheet
    .getRange('B45')
    .setValue(
      'API 안내  ·  YES24 API Key는 도서 검색에 꼭 필요합니다. 국립중앙도서관 API는 KDC 분류번호 참고용이라 입력하지 않아도 신청 기능은 정상 작동합니다.'
    );

  sheet
    .getRange('B3:L4')
    .setBackground(titleColor)
    .setFontColor('#FFFFFF')
    .setFontSize(27)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  sheet
    .getRange('B5:L5')
    .setBackground('#E7EDF8')
    .setFontColor(titleColor)
    .setFontSize(14)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  sheet
    .getRange('B16:L17')
    .setBackground('#EDF2FA')
    .setFontColor(titleColor)
    .setFontSize(14)
    .setFontWeight('normal')
    .setHorizontalAlignment('left')
    .setBorder(
      true,
      true,
      true,
      true,
      false,
      false,
      lineColor,
      SpreadsheetApp.BorderStyle.SOLID
    );

  sheet
    .getRange('B19:L20')
    .setBackground(accentColor)
    .setFontColor('#FFFFFF')
    .setFontSize(17)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  sheet
    .getRange('B22:L23')
    .setBackground('#6F4A25')
    .setFontColor('#FFFFFF')
    .setFontSize(18)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  sheet
    .getRange('B41:L43')
    .setBackground('#FFF0F0')
    .setFontColor('#A33C3C')
    .setFontSize(13)
    .setFontWeight('bold')
    .setHorizontalAlignment('left');

  sheet
    .getRange('B45:L46')
    .setBackground('#EDF4FF')
    .setFontColor('#284E86')
    .setFontSize(13)
    .setFontWeight('bold')
    .setHorizontalAlignment('left');

  [3, 4].forEach(function(rowNumber) {
    sheet.setRowHeight(rowNumber, 38);
  });
  sheet.setRowHeight(5, 34);
  [7, 8, 10, 11, 13, 14].forEach(
    function(rowNumber) {
      sheet.setRowHeight(rowNumber, 32);
    }
  );
  [6, 9, 12, 15, 18].forEach(
    function(rowNumber) {
      sheet.setRowHeight(rowNumber, 12);
    }
  );
  [16, 17, 19, 20].forEach(
    function(rowNumber) {
      sheet.setRowHeight(rowNumber, 31);
    }
  );
  [22, 23].forEach(function(rowNumber) {
    sheet.setRowHeight(rowNumber, 30);
  });
  [
    25, 26,
    28, 29,
    31, 32,
    34, 35,
    37, 38
  ].forEach(
    function(rowNumber) {
      sheet.setRowHeight(rowNumber, 31);
    }
  );
  [21, 24, 27, 30, 33, 36, 39, 40, 44].forEach(
    function(rowNumber) {
      sheet.setRowHeight(rowNumber, 10);
    }
  );
  [41, 42, 43].forEach(function(rowNumber) {
    sheet.setRowHeight(rowNumber, 28);
  });
  [45, 46].forEach(function(rowNumber) {
    sheet.setRowHeight(rowNumber, 31);
  });

  sheet.activate();
  sheet.getRange('A3').activate();
}

/**
 * ============================================================
 * 사본 배포용 학교 최초 설정
 * ============================================================
 */

function HR_addCopySetupMenu() {
  SpreadsheetApp.getUi()
    .createMenu('⚙️ 사본 설정')
    .addItem('✅ 설정 한 번에 점검', 'HR_showSystemCheck')
    .addItem('🔗 신청 링크 받기', 'HR_showApplyLinksDialog')
    .addItem('📢 업데이트 소식 보기', 'HR_showUpdateNotice')
    .addItem('📏 신청 한도 설정', 'HR_showRequestLimitDialog')
    .addSeparator()
    .addItem(
      '🔑 API 키 설정 (YES24 필수)',
      'HR_showYes24KeyDialog'
    )
    .addItem('🏫 학교 최초 설정', 'HR_showSchoolSetupDialog')
    .addItem(
      '🔗 공개 신청 링크 등록',
      'HR_showPublicWebAppUrlDialog'
    )
    .addSeparator()
    .addItem('📥 이전 사본에서 데이터 가져오기', 'HR_importFromOldCopy')
    .addItem('🧰 설정 탭 만들기·새로고침', 'HR_createCopySetupSheet')
    .addItem('🗑️ 설정 완료 후 탭 삭제', 'HR_deleteCopySetupSheet')
    .addToUi();
}

function HR_showYes24KeyDialog() {
  try {
    var template = HtmlService.createTemplateFromFile('CopySetup');
    template.mode = 'api';
    var output = template.evaluate().setWidth(680).setHeight(690);
    SpreadsheetApp.getUi().showModalDialog(
      output,
      '🔑 도서 API 키 설정'
    );
  } catch (error) {
    var ui = SpreadsheetApp.getUi();
    var yes24Response = ui.prompt(
      '🔑 YES24 API 키 설정',
      'YES24 API Key를 붙여 넣으세요.',
      ui.ButtonSet.OK_CANCEL
    );
    if (yes24Response.getSelectedButton() !== ui.Button.OK) return;
    var yes24Key = String(yes24Response.getResponseText() || '').trim();
    if (!yes24Key) {
      ui.alert('YES24 API 키를 입력해 주세요.');
      return;
    }
    var nlkResponse = ui.prompt(
      '국립중앙도서관 API 키 (선택)',
      '없으면 빈칸으로 두고 확인을 누르세요.',
      ui.ButtonSet.OK_CANCEL
    );
    var nlkKey = nlkResponse.getSelectedButton() === ui.Button.OK
      ? String(nlkResponse.getResponseText() || '').trim()
      : '';
    try {
      var result = HR_saveYes24ApiKey({
        yes24Key: yes24Key,
        nlkKey: nlkKey
      });
      ui.alert(result.message || 'API 키를 저장했습니다.');
    } catch (saveError) {
      ui.alert(
        'API 키 저장 오류',
        saveError.message || String(saveError),
        ui.ButtonSet.OK
      );
    }
  }
}

function HR_showSchoolSetupDialog() {
  HR_showCopySetupDialog_(
    'school',
    '🏫 학교 최초 설정',
    700,
    680
  );
}

/**
 * 웹 앱 배포 주소를 한 번만 붙여 넣는 간단한 설정창입니다.
 * 저장하면 사용설명서의 「희망도서 신청하기」 버튼도 자동 갱신됩니다.
 */
function HR_showPublicWebAppUrlDialog() {
    var activeSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  if (activeSheet) {
    activeSheet.getRange('A3').activate();
  }

  var savedUrl = HR_getConfiguredPublicWebAppUrl_();
  var html = HtmlService
    .createHtmlOutput(
      '<!DOCTYPE html>' +
      '<html lang="ko"><head><base target="_top">' +
      '<meta charset="UTF-8">' +
      '<style>' +
      '*{box-sizing:border-box}' +
      'body{margin:0;padding:28px;background:#f5f8ff;color:#253858;' +
      'font-family:-apple-system,BlinkMacSystemFont,"Apple SD Gothic Neo",' +
      '"Noto Sans KR",sans-serif}' +
      'h1{margin:0 0 8px;font-size:24px}' +
      '.sub{margin:0 0 22px;color:#5f6f8b;font-size:15px;line-height:1.65}' +
      '.card{padding:22px;border:1px solid #dbe4f3;border-radius:18px;' +
      'background:#fff;box-shadow:0 10px 28px rgba(54,83,130,.08)}' +
      'label{display:block;margin-bottom:9px;font-size:17px;font-weight:800}' +
      'input{width:100%;height:56px;padding:0 16px;border:2px solid #cfdaf0;' +
      'border-radius:13px;outline:none;font-size:16px}' +
      'input:focus{border-color:#5277c7;box-shadow:0 0 0 4px #e8eefb}' +
      '.help{margin:12px 0 0;color:#66758f;font-size:14px;line-height:1.6}' +
      '.message{display:none;margin-top:15px;padding:12px 14px;' +
      'border-radius:11px;font-size:14px;line-height:1.5}' +
      '.message.show{display:block}.message.error{background:#fff0f0;color:#a23f3f}' +
      '.message.ok{background:#edf8f0;color:#27633b}' +
      '.buttons{display:flex;justify-content:flex-end;gap:10px;margin-top:20px}' +
      'button{height:48px;padding:0 20px;border-radius:12px;font-size:16px;' +
      'font-weight:800;cursor:pointer}.cancel{border:1px solid #d6deeb;' +
      'background:#fff;color:#65738a}.save{border:0;background:#5277c7;color:#fff}' +
      'button:disabled{cursor:wait;opacity:.55}' +
      '</style></head><body>' +
      '<h1>🔗 공개 신청 링크 등록</h1>' +
      '<p class="sub">웹 앱을 배포한 뒤 주소를 아래 칸에 한 번만 붙여 넣으세요.<br>' +
      '사용설명서의 신청 버튼은 자동으로 바뀝니다.</p>' +
      '<div class="card">' +
      '<label for="publicUrl">웹 앱 주소</label>' +
      '<input id="publicUrl" type="url" spellcheck="false" ' +
      'placeholder="https://script.google.com/macros/s/.../exec" value="' +
      HR_escapeHtml_(savedUrl) +
      '">' +
      '<p class="help">Apps Script → 배포 → 배포 관리에서 ' +
      '<strong>/exec로 끝나는 웹 앱 URL</strong>을 복사해 붙여 넣으세요.</p>' +
      '<div id="message" class="message"></div>' +
      '<div class="buttons">' +
      '<button class="cancel" type="button" ' +
      'onclick="google.script.host.close()">닫기</button>' +
      '<button id="saveButton" class="save" type="button" ' +
      'onclick="saveUrl()">링크 저장하기</button>' +
      '</div></div>' +
      '<script>' +
      'function saveUrl(){' +
      'var input=document.getElementById("publicUrl");' +
      'var button=document.getElementById("saveButton");' +
      'var message=document.getElementById("message");' +
      'var url=input.value.trim();' +
      'if(!url){showMessage("웹 앱 주소를 붙여 넣어 주세요.","error");return;}' +
      'button.disabled=true;button.textContent="저장 중…";' +
      'google.script.run.withSuccessHandler(function(result){' +
      'showMessage(result.message,"ok");button.textContent="저장 완료 ✓";' +
      'setTimeout(function(){google.script.host.close();},900);' +
      '}).withFailureHandler(function(error){' +
      'button.disabled=false;button.textContent="링크 저장하기";' +
      'showMessage(error&&error.message?error.message:String(error),"error");' +
      '}).HR_savePublicWebAppUrl({url:url});}' +
      'function showMessage(text,type){var box=document.getElementById("message");' +
      'box.textContent=text;box.className="message show "+type;}' +
      '</script></body></html>'
    )
    .setWidth(650)
    .setHeight(500);

  SpreadsheetApp.getUi().showModalDialog(
    html,
    '공개 신청 링크 등록'
  );
}

/**
 * 공개 /exec 주소를 저장하고 사용설명서 신청 버튼에 즉시 반영합니다.
 */
function HR_savePublicWebAppUrl(formData) {
  var url = String(
    formData && formData.url || ''
  ).trim();

  if (!url) {
    throw new Error('웹 앱 주소를 붙여 넣어 주세요.');
  }

  // 복사 과정에서 들어간 공백·쿼리 문자열을 제거합니다.
  url = url.replace(/\s+/g, '');
  url = url.replace(/[?#].*$/, '');

  // 배포 기본 주소만 붙여 넣어도 공개용 /exec 주소로 자동 보정합니다.
  if (
    /^https:\/\/script\.google\.com\/macros\/s\/[^/]+\/?$/.test(url)
  ) {
    url = url.replace(/\/$/, '') + '/exec';
  }

  url = url
    .replace(/\/dev\/?$/, '/exec')
    .replace(/\/exec\/?$/, '/exec');

  if (
    !/^https:\/\/script\.google\.com\/macros\/s\/[^/]+\/exec$/.test(
      url
    )
  ) {
    throw new Error(
      '주소가 올바르지 않습니다. Apps Script 배포 관리의 웹 앱 주소를 붙여 넣어 주세요.'
    );
  }

  var properties = PropertiesService.getScriptProperties();

  // 주소 검사가 끝난 뒤에만 저장하고 안내 시트를 다시 그립니다.
  // 따라서 잘못된 주소를 입력해도 뒤쪽 시트 디자인은 바뀌지 않습니다.
  properties.setProperty(
    HR_PUBLIC_WEB_APP_URL_PROPERTY,
    url
  );

  var updatedButtons = HR_updateGuidePublicLink_(url);

  return {
    success: true,
    message:
      '링크를 저장했습니다. 희망도서 신청 버튼을 새 주소로 연결했습니다.',
    webAppUrl: url
  };
}

/**
 * 공개 링크 저장 시 안내 시트 전체를 다시 만들지 않고
 * 기존 신청 버튼의 링크만 안전하게 바꿉니다.
 * 병합 셀·고정 행 구조를 건드리지 않으므로 디자인이 깨지지 않습니다.
 */
function HR_updateGuidePublicLink_(url) {
  var spreadsheet = HR_getSetupSpreadsheet_();
  var sheet = spreadsheet.getSheetByName('📗 희망도서 신청안내');
  if (!sheet) {
    throw new Error('희망도서 신청안내 시트를 찾지 못했습니다.');
  }

  var linkText = '📚 희망도서 신청하기';
  var linkValue = SpreadsheetApp
    .newRichTextValue()
    .setText(linkText)
    .setLinkUrl(String(url || ''))
    .build();
  var targets = [];
  var seen = {};

  // 이미 만들어진 신청 버튼이 있으면 링크만 바꿉니다.
  var found = sheet
    .createTextFinder('희망도서 신청하기')
    .matchCase(false)
    .matchEntireCell(false)
    .findAll();

  found.forEach(function(cell) {
    var merged = cell.getMergedRanges();
    var buttonRange = merged.length ? merged[0] : cell;
    var target = buttonRange.getCell(1, 1);
    var key = target.getA1Notation();

    if (!seen[key]) {
      seen[key] = true;
      targets.push({
        cell: target,
        range: buttonRange
      });
    }
  });

  // 현재 통합 안내 양식에는 버튼이 없으므로
  // 본문 아래의 새 행에 버튼을 한 번만 만듭니다.
  if (!targets.length) {
    var buttonStartRow = Math.max(26, sheet.getLastRow() + 2);
    var buttonEndRow = buttonStartRow + 2;

    if (sheet.getMaxRows() < buttonEndRow) {
      sheet.insertRowsAfter(
        sheet.getMaxRows(),
        buttonEndRow - sheet.getMaxRows()
      );
    }

    var newButtonRange = sheet.getRange(
      buttonStartRow,
      1,
      3,
      12
    );

    newButtonRange
      .breakApart()
      .clearContent()
      .merge();

    targets.push({
      cell: newButtonRange.getCell(1, 1),
      range: newButtonRange
    });

    sheet.setRowHeight(buttonStartRow - 1, 10);
    sheet.setRowHeights(buttonStartRow, 3, 30);
  }

  targets.forEach(function(item) {
    item.cell.setRichTextValue(linkValue);

    item.range
      .setBackground('#2F765E')
      .setFontColor('#FFFFFF')
      .setFontFamily('Arial')
      .setFontSize(18)
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle')
      .setBorder(
        true,
        true,
        true,
        true,
        false,
        false,
        '#245B49',
        SpreadsheetApp.BorderStyle.SOLID_THICK
      );
  });

  SpreadsheetApp.flush();
  return targets.length;
}

function HR_showCopySetupDialog_(
  mode,
  title,
  width,
  height
) {
  var template = HtmlService.createTemplateFromFile(
    'CopySetup'
  );

  template.mode = mode;

  SpreadsheetApp.getUi().showModalDialog(
    template
      .evaluate()
      .setWidth(width)
      .setHeight(height),
    title
  );
}

/**
 * 설정 탭의 큰 버튼 영역을 선택하면 설정창을 엽니다.
 * 상단 「⚙️ 사본 설정」 메뉴에서도 같은 창을 열 수 있습니다.
 */
function onSelectionChange(e) {
  if (!e || !e.range) {
    return;
  }

  var sheet = e.range.getSheet();

  if (sheet.getName() !== HR_COPY_SETUP_SHEET_NAME) {
    return;
  }

  var activeCell = e.range.getA1Notation().split(':')[0];

  if (activeCell === 'B61') {
    HR_showYes24KeyDialog();
  } else if (activeCell === 'B65') {
    HR_showSchoolSetupDialog();
  } else if (activeCell === 'B69') {
    HR_showPublicWebAppUrlDialog();
  }
}

function HR_getCopySetupState() {
  var spreadsheet = HR_getSetupSpreadsheet_();
  var properties = PropertiesService.getScriptProperties();
  var libraryStatus = HR_getLibrarySetupStatus_(spreadsheet);
  var configuredSpreadsheetId = String(
    properties.getProperty(HR_SPREADSHEET_ID_PROPERTY) || ''
  ).trim();
  var yes24Key = HR_getYes24Key_();
  var nlkKey = HR_getNlkKey_();
  var webAppUrl = HR_getConfiguredPublicWebAppUrl_();

  return {
    spreadsheetId: spreadsheet.getId(),
    configuredSpreadsheetId: configuredSpreadsheetId,
    spreadsheetName: spreadsheet.getName(),
    schoolName: String(
      properties.getProperty(HR_SCHOOL_NAME_PROPERTY) || ''
    ).trim(),
    managerName: String(
      properties.getProperty(HR_MANAGER_NAME_PROPERTY) || ''
    ).trim(),
    managerExtension: String(
      properties.getProperty(HR_MANAGER_EXTENSION_PROPERTY) || ''
    ).trim(),
    yes24Key: '',
    yes24Ready: Boolean(yes24Key),
    nlkKey: '',
    nlkReady: Boolean(nlkKey),
    targetSheetReady: Boolean(
      spreadsheet.getSheetByName(HR_TARGET_SHEET_NAME)
    ),
    librarySheetReady: libraryStatus.sheetReady,
    libraryHeadersReady: libraryStatus.headersReady,
    libraryTitleColumnReady: libraryStatus.titleColumn > 0,
    libraryIsbnColumnReady: libraryStatus.isbnColumn > 0,
    libraryBookCount: libraryStatus.bookCount,
    webAppUrl: webAppUrl,
    setupComplete:
      String(
        properties.getProperty(HR_SETUP_COMPLETE_PROPERTY) || ''
      ) === 'true' &&
      configuredSpreadsheetId === spreadsheet.getId() &&
      Boolean(yes24Key) &&
      libraryStatus.headersReady
  };
}

function HR_saveYes24ApiKey(formData) {
  var key = String(
    formData && formData.yes24Key || HR_getYes24Key_()
  ).trim();
  var nlkKey = String(
    formData && formData.nlkKey || HR_getNlkKey_()
  ).trim();
  var properties = PropertiesService.getScriptProperties();

  if (!key) {
    throw new Error('YES24 API Key를 입력해 주세요.');
  }

  HR_testYes24Key_(key);

  properties.setProperty(HR_YES24_KEY_PROPERTY, key);

  if (nlkKey) {
    properties.setProperty(HR_NLK_KEY_PROPERTY, nlkKey);
  } else {
    properties.deleteProperty(HR_NLK_KEY_PROPERTY);
  }

  return {
    success: true,
    message:
      'YES24 연결을 확인했습니다.' +
      (nlkKey
        ? ' 국립중앙도서관 키도 저장했습니다.'
        : ' KDC 참고용 키는 비워 두었습니다.')
  };
}

function HR_runSchoolInitialSetup(formData) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var properties = PropertiesService.getScriptProperties();
  var schoolName = String(
    formData && formData.schoolName || ''
  ).trim();
  var managerName = String(
    formData && formData.managerName || ''
  ).trim();
  var managerExtension = String(
    formData && formData.managerExtension || ''
  ).trim();
  var targetSheet;
  var libraryStatus;
  var yes24Ready = Boolean(HR_getYes24Key_());
  var warnings = [];

  if (!spreadsheet) {
    throw new Error(
      '사본 스프레드시트에서 학교 최초 설정창을 열어 주세요.'
    );
  }

  if (!schoolName) {
    throw new Error('학교 이름을 입력해 주세요.');
  }

  targetSheet = spreadsheet.getSheetByName(
    HR_TARGET_SHEET_NAME
  );

  if (!targetSheet) {
    throw new Error(
      '「' + HR_TARGET_SHEET_NAME +
      '」 시트를 찾지 못했습니다.'
    );
  }

  properties.setProperties({
    HR_TARGET_SPREADSHEET_ID: spreadsheet.getId(),
    HR_SCHOOL_NAME: schoolName,
    HR_MANAGER_NAME: managerName,
    HR_MANAGER_EXTENSION: managerExtension,
    HR_SETUP_COMPLETE: 'true'
  });

  libraryStatus = HR_getLibrarySetupStatus_(spreadsheet);

  if (!libraryStatus.headersReady) {
    try {
      HR_ensureLibrarySheet_(spreadsheet);
      libraryStatus = HR_getLibrarySetupStatus_(spreadsheet);
    } catch (libraryError) {
      throw new Error(
        '우리학교 소장도서 시트 준비 중 오류가 발생했습니다: ' +
        String(libraryError && libraryError.message || libraryError)
      );
    }
  }

  try {
    HR_installEditTriggerForSpreadsheet_(spreadsheet);
  } catch (triggerError) {
    warnings.push('자동입력 트리거는 상단 메뉴의 설정 점검에서 다시 설치할 수 있습니다.');
  }

  return {
    success: true,
    message:
      schoolName + ' 설정을 완료했습니다.' +
      (yes24Ready
        ? ''
        : ' YES24 API 키는 「API 키 설정」에서 별도로 입력해 주세요.') +
      (warnings.length ? ' ' + warnings.join(' ') : ''),
    libraryHeadersReady: libraryStatus.headersReady,
    libraryBookCount: libraryStatus.bookCount,
    webAppUrl: HR_getConfiguredPublicWebAppUrl_(),
    warnings: warnings
  };
}

function HR_deleteCopySetupSheet() {
  var spreadsheet = HR_getSetupSpreadsheet_();
  var properties = PropertiesService.getScriptProperties();
  var configuredSpreadsheetId = String(
    properties.getProperty(HR_SPREADSHEET_ID_PROPERTY) || ''
  ).trim();
  var setupSheet = spreadsheet.getSheetByName(
    HR_COPY_SETUP_SHEET_NAME
  );
  var guideSheet = spreadsheet.getSheetByName('📗 희망도서 신청안내');

  if (
    configuredSpreadsheetId !== spreadsheet.getId() ||
    !HR_getYes24Key_()
  ) {
    throw new Error(
      'YES24 API 키와 학교 최초 설정을 먼저 완료해 주세요.'
    );
  }

  if (setupSheet) {
    if (guideSheet) {
      spreadsheet.setActiveSheet(guideSheet);
    }

    spreadsheet.deleteSheet(setupSheet);
  }

  return {
    success: true,
    message: '사본 최초설정 탭을 삭제했습니다.'
  };
}

function HR_createCopySetupSheet() {
  var spreadsheet = HR_getSetupSpreadsheet_();

  HR_createCopySetupSheet_(spreadsheet);

  var setupSheet = spreadsheet.getSheetByName(
    HR_COPY_SETUP_SHEET_NAME
  );

  spreadsheet.setActiveSheet(setupSheet);
  setupSheet.getRange('A3').activate();

  return {
    success: true,
    message: '사본 최초설정 탭을 준비했습니다.'
  };
}

/**
 * 기존 배포용 사본에서 이미지 버튼을 다시 만들지 않고
 * 초보자용 안내 문구와 YES24 바로가기만 빠르게 갱신합니다.
 */
function HR_refreshCopySetupBeginnerGuide() {
  var spreadsheet = HR_getSetupSpreadsheet_();
  HR_createCopySetupSheet_(spreadsheet);

  var sheet = spreadsheet.getSheetByName(
    HR_COPY_SETUP_SHEET_NAME
  );

  spreadsheet.setActiveSheet(sheet);
  sheet.getRange('A3').activate();

  return {
    success: true,
    message: '초보자용 8단계 안내를 맨 위로 옮겼습니다.'
  };
}

/**
 * 초보자용 8단계를 가장 먼저 보이게 배치한 사본 최초설정 화면입니다.
 */
function HR_createCopySetupSheet_(spreadsheet) {
  var sheet = spreadsheet.getSheetByName(
    HR_COPY_SETUP_SHEET_NAME
  );
  var libraryStatus = HR_getLibrarySetupStatus_(spreadsheet);
  var properties = PropertiesService.getScriptProperties();
  var configuredId = String(
    properties.getProperty(HR_SPREADSHEET_ID_PROPERTY) || ''
  ).trim();
  var yes24Ready = Boolean(HR_getYes24Key_());
  var nlkReady = Boolean(HR_getNlkKey_());
  var schoolReady = configuredId === spreadsheet.getId();
  var publicUrlReady = /\/exec(?:$|[?#])/.test(
    HR_getConfiguredPublicWebAppUrl_()
  );
  var detailSteps;

  if (!sheet) {
    sheet = spreadsheet.insertSheet(
      HR_COPY_SETUP_SHEET_NAME,
      0
    );
  }

  
  
    sheet.setFrozenRows(0);
  sheet.setFrozenColumns(0);
  HR_breakMergedRangesSafely_(sheet);
sheet.clear();
  sheet.clearConditionalFormatRules();
  sheet.setHiddenGridlines(true);
  sheet.setColumnWidth(1, 58);
  sheet.setColumnWidths(2, 11, 76);
  sheet
    .getRange('B3:L86')
    .setBackground('#FFFFFF')
    .setFontFamily('Jua')
    .setFontColor('#2E3A33')
    .setVerticalAlignment('middle')
    .setWrap(true);

  sheet.getRange('B3:L4').merge();
  sheet
    .getRange('B3')
    .setValue('⚙️ 다른 학교 사본용 최초 설정');
  sheet.getRange('B5:L5').merge();
  sheet
    .getRange('B5')
    .setValue(
      '처음 한 번만 아래 1번부터 8번까지 차례대로 따라 하세요.'
    );

  sheet.getRange('B7:L8').merge();
  sheet
    .getRange('B7')
    .setValue('👣 컴퓨터가 익숙하지 않아도 괜찮아요! 위에서부터 그대로 따라 하세요')
    .setBackground('#183D35')
    .setFontColor('#FFFFFF')
    .setFontSize(21)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  detailSteps = [
    {
      badge: 'B10:C13',
      text: 'D10:L13',
      number: '1',
      value:
        '사본 만들고 시작 버튼 누르기\n' +
        '① 원본 시트에서 「파일 → 사본 만들기」를 누르고 새 학교 이름으로 저장하세요.\n' +
        '② 새 파일이 열리면 5~10초 기다린 뒤 브라우저 새로고침 버튼을 한 번 누르세요.\n' +
        '③ 위쪽 「⚙️ 시작하기」 → 「✅ 프로그램 사용 시작」을 누르세요. 구글 허용 창이 뜨면 내 계정 → 「고급」 → 「이동(안전하지 않음)」 → 「허용」 순서로 누르세요. 처음 한 번만 하면 되고, 경고가 뜨는 건 정상이에요. 몇 초 뒤 「⚙️ 사본 설정」 메뉴가 나타납니다.'
    },
    {
      badge: 'B15:C18',
      text: 'D15:L18',
      number: '2',
      value:
        'YES24 API 키 받기\n' +
        '바로 아래 「YES24 API 키 발급 사이트 열기」를 누르세요. YES24에 로그인한 뒤 인증키 발급·관리 화면에서 API Key 전체를 복사하세요.'
    },
    {
      badge: 'B21:C24',
      text: 'D21:L24',
      number: '3',
      value:
        '복사한 API 키 붙여 넣기\n' +
        '화면 위쪽 「⚙️ 사본 설정」 → 「🔑 API 키 설정 (YES24 필수)」을 누르세요. YES24 API Key 칸에 붙여 넣고 저장하세요. 국립중앙도서관 키는 KDC 참고용이므로 비워도 됩니다.'
    },
    {
      badge: 'B26:C30',
      text: 'D26:L30',
      number: '4',
      value:
        'DLS에서 우리 학교 소장도서 반출하기\n' +
        'DLS 접속 → 「자료관리」 → 오른쪽 아래 「페이징 권수」를 우리 학교가 가진 전체 권수만큼 입력 → 「칼럼」 → 「칼럼 설정」 → 「ISBN」 체크 → 「반출」을 누르세요.\n' +
        '반출한 파일의 책 제목과 ISBN을 복사해 「우리학교 소장도서」 시트 2행부터 붙여 넣으세요. 이 정보가 있어야 중복 도서를 알려 줍니다.'
    },
    {
      badge: 'B32:C35',
      text: 'D32:L35',
      number: '5',
      value:
        '학교 이름 연결하기\n' +
        '화면 위쪽 「⚙️ 사본 설정」 → 「🏫 학교 최초 설정」을 누르세요. 학교 이름을 입력하고 실행한 뒤 완료 안내가 나올 때까지 기다리세요.'
    },
    {
      badge: 'B37:C42',
      text: 'D37:L42',
      number: '6',
      value:
        '신청 웹페이지 만들기\n' +
        '화면 위쪽 「확장 프로그램」 → 「Apps Script」를 누르세요. 새 화면에서 「배포」 → 「새 배포」 → 톱니바퀴 모양 「유형 선택」 → 「웹 앱」을 고릅니다.\n' +
        '실행 사용자는 「나」, 액세스 권한은 「모든 사용자」로 정하고 「배포」를 누르세요. 마지막에 나오는 /exec 주소를 복사하세요.'
    },
    {
      badge: 'B44:C47',
      text: 'D44:L47',
      number: '7',
      value:
        '공개 신청 링크 저장하기\n' +
        '다시 스프레드시트로 돌아와 「⚙️ 사본 설정」 → 「🔗 공개 신청 링크 등록」을 누르세요. 방금 복사한 /exec 주소를 붙여 넣고 「링크 저장하기」를 누르세요.'
    },
    {
      badge: 'B49:C52',
      text: 'D49:L52',
      number: '8',
      value:
        '마지막으로 시험하기\n' +
        '「사용설명서」 탭의 「📚 희망도서 신청하기」를 눌러 보세요. 책 제목을 하나 검색해 신청해 보고, 「희망도서 목록」에 들어오면 성공입니다. 마지막에는 「⚙️ 사본 설정」 → 「✅ 설정 한 번에 점검」도 눌러 보세요.'
    }
  ];

  HR_drawGuideSteps_(
    sheet,
    detailSteps,
    '#244D42',
    '#EAF3ED',
    '#A8C8BB'
  );

  sheet.getRange('B19:L19').merge();
  sheet
    .getRange('B19')
    .setFormula(
      '=HYPERLINK("' +
      HR_YES24_API_MANAGE_URL +
      '","🔗 YES24 API 키 발급 사이트 열기")'
    )
    .setBackground('#FBF1D6')
    .setFontColor('#7A5A12')
    .setFontSize(15)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setBorder(
      true,
      true,
      true,
      true,
      false,
      false,
      '#E3C47A',
      SpreadsheetApp.BorderStyle.SOLID
    );

  sheet.getRange('B54:L56').merge();
  sheet
    .getRange('B54')
    .setValue(
      '⚠️ 꼭 확인하세요\n' +
      '① 「희망도서 목록」 시트 이름은 변경하지 마세요. 새 목록으로 교체할 때는 기존 시트 이름을 바꾼 뒤, 새 시트를 반드시 「희망도서 목록」으로 만드세요.\n' +
      '② 「우리학교 소장도서」에는 책 제목과 ISBN13을 모두 넣어야 중복 소장 안내가 정확히 작동합니다.'
    )
    .setBackground('#FFF6DA')
    .setFontColor('#7A5A12')
    .setFontSize(14)
    .setFontWeight('bold')
    .setHorizontalAlignment('left')
    .setBorder(
      true,
      true,
      true,
      true,
      false,
      false,
      '#E3C47A',
      SpreadsheetApp.BorderStyle.SOLID
    );

  sheet.getRange('B58:L59').merge();
  sheet
    .getRange('B58')
    .setValue('🧰 설정 버튼 모음  ·  안내를 읽은 뒤 필요한 버튼을 누르세요')
    .setBackground('#183D35')
    .setFontColor('#FFFFFF')
    .setFontSize(20)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  sheet.getRange('B61:L63').merge();
  sheet
    .getRange('B61')
    .setValue(
      (yes24Ready ? '✅ ' : '') +
      '🔑 API 키 설정  ·  YES24 필수 / 국립중앙도서관 선택'
    );
  sheet.getRange('B65:L67').merge();
  sheet
    .getRange('B65')
    .setValue(
      (schoolReady ? '✅ ' : '') +
      '🏫 이 학교 최초 설정'
    );
  sheet.getRange('B69:L71').merge();
  sheet
    .getRange('B69')
    .setValue(
      (publicUrlReady ? '✅ ' : '') +
      '🔗 공개 신청 링크 등록'
    );

  ['B61:L63', 'B65:L67', 'B69:L71'].forEach(
    function(rangeA1) {
      sheet
        .getRange(rangeA1)
        .setBackground('#244D42')
        .setFontColor('#FFFFFF')
        .setFontSize(18)
        .setFontWeight('bold')
        .setHorizontalAlignment('center')
        .setBorder(
          true,
          true,
          true,
          true,
          false,
          false,
          '#B9852A',
          SpreadsheetApp.BorderStyle.SOLID_THICK
        );
    }
  );

  sheet.getRange('B73:L75').merge();
  sheet
    .getRange('B73')
    .setValue(
      '현재 점검  ·  YES24 ' +
      (yes24Ready ? '완료' : '필요') +
      '  ·  KDC 참고 API ' +
      (nlkReady ? '사용 중' : '선택') +
      '  ·  신청서 연결 ' +
      (schoolReady ? '완료' : '필요') +
      '  ·  소장도서 제목/ISBN13 ' +
      (libraryStatus.headersReady ? '완료' : '필요') +
      '  ·  공개 링크 ' +
      (publicUrlReady ? '완료' : '필요')
    )
    .setBackground('#F8F5ED')
    .setFontColor('#34564D')
    .setFontSize(13)
    .setHorizontalAlignment('center');

  sheet.getRange('B77:L79').merge();
  sheet
    .getRange('B77')
    .setValue('🗑️ 모든 설정을 마친 뒤 이 탭 삭제')
    .setBackground('#34564D')
    .setFontColor('#FFFFFF')
    .setFontSize(17)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  sheet.getRange('B81:L82').merge();
  sheet
    .getRange('B81')
    .setValue(
      'YES24은 도서 검색에 꼭 필요합니다. 국립중앙도서관 API는 KDC 참고용이라 없어도 됩니다.\n' +
      '이 시트의 프로그램은 저절로 최신 기능으로 바뀝니다. 따로 업데이트하지 않아도 돼요.'
    )
    .setBackground('#EAF3ED')
    .setFontColor('#244D42')
    .setFontSize(13)
    .setFontWeight('bold')
    .setHorizontalAlignment('left');

  sheet
    .getRange('B3:L4')
    .setBackground('#183D35')
    .setFontColor('#FFFFFF')
    .setFontSize(25)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  sheet
    .getRange('B5:L5')
    .setBackground('#F6EBCF')
    .setFontColor('#183D35')
    .setFontSize(14)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  [3, 4].forEach(function(rowNumber) {
    sheet.setRowHeight(rowNumber, 36);
  });
  [5, 19].forEach(function(rowNumber) {
    sheet.setRowHeight(rowNumber, 34);
  });
  [7, 8, 58, 59].forEach(function(rowNumber) {
    sheet.setRowHeight(rowNumber, 34);
  });
  [
    15, 16, 17, 18,
    21, 22, 23, 24, 32, 33, 34, 35,
    44, 45, 46, 47, 49, 50, 51, 52,
    61, 62, 63, 65, 66, 67, 69, 70, 71,
    73, 74, 75, 77, 78, 79
  ].forEach(function(rowNumber) {
    sheet.setRowHeight(rowNumber, 34);
  });
  [10, 11, 12, 13].forEach(function(rowNumber) {
    sheet.setRowHeight(rowNumber, 62);
  });
  [26, 27, 28, 29, 30].forEach(function(rowNumber) {
    sheet.setRowHeight(rowNumber, 40);
  });
  [37, 38, 39, 40, 41, 42].forEach(function(rowNumber) {
    sheet.setRowHeight(rowNumber, 35);
  });
  [54, 55, 56].forEach(function(rowNumber) {
    sheet.setRowHeight(rowNumber, 42);
  });
  [81, 82].forEach(function(rowNumber) {
    sheet.setRowHeight(rowNumber, 38);
  });
  [
    6, 9, 14, 20, 25, 31, 36, 43, 48, 53,
    57, 60, 64, 68, 72, 76, 80, 83, 84, 85, 86
  ].forEach(function(rowNumber) {
    sheet.setRowHeight(rowNumber, 10);
  });

  HR_addSetupButtonImages_(sheet);
}

function HR_createCopySetupSheetBeforeTopGuide_(spreadsheet) {
  var sheet = spreadsheet.getSheetByName(
    HR_COPY_SETUP_SHEET_NAME
  );
  var libraryStatus = HR_getLibrarySetupStatus_(spreadsheet);
  var properties = PropertiesService.getScriptProperties();
  var configuredId = String(
    properties.getProperty(HR_SPREADSHEET_ID_PROPERTY) || ''
  ).trim();
  var yes24Ready = Boolean(HR_getYes24Key_());
  var nlkReady = Boolean(HR_getNlkKey_());
  var schoolReady = configuredId === spreadsheet.getId();
  var publicUrlReady = /\/exec(?:$|[?#])/.test(
    HR_getConfiguredPublicWebAppUrl_()
  );
  var steps;

  if (!sheet) {
    sheet = spreadsheet.insertSheet(
      HR_COPY_SETUP_SHEET_NAME,
      0
    );
  }

  sheet.getRange('B3:L95').breakApart();
  sheet.clear();
  sheet.clearConditionalFormatRules();
  sheet.setHiddenGridlines(true);
  sheet.setColumnWidth(1, 58);
  sheet.setColumnWidths(2, 11, 76);
  sheet
    .getRange('B3:L95')
    .setBackground('#FFFFFF')
    .setFontFamily('Jua')
    .setFontColor('#27344A')
    .setVerticalAlignment('middle')
    .setWrap(true);

  sheet.getRange('B3:L4').merge();
  sheet
    .getRange('B3')
    .setValue('⚙️ 다른 학교 사본용 최초 설정');
  sheet.getRange('B5:L5').merge();
  sheet
    .getRange('B5')
    .setValue(
      '사본을 만든 뒤 이 화면의 1번부터 차례대로 따라 하세요. 메뉴가 안 보이면 새로고침하세요.'
    );

  steps = [
    {
      badge: 'B7:C8',
      text: 'D7:L8',
      number: '1',
      value:
        '원본에서 「파일 → 사본 만들기」를 누릅니다. 새 파일이 열리면 상단의 「⚙️ 사본 설정」 메뉴가 나타날 때까지 잠시 기다립니다.'
    },
    {
      badge: 'B15:C17',
      text: 'D15:L17',
      number: '3',
      value:
        'DLS에서 책 제목과 ISBN을 반출한 뒤, 「우리학교 소장도서」 시트의 2행부터 붙여 넣습니다.'
    },
    {
      badge: 'B23:C25',
      text: 'D23:L25',
      number: '5',
      value:
        '확장 프로그램 → Apps Script → 배포 → 새 배포 → 웹 앱에서 「나 / 모든 사용자」로 배포합니다.'
    },
    {
      badge: 'B27:C29',
      text: 'D27:L29',
      number: '6',
      value:
        '상단 「⚙️ 사본 설정 → 🔗 공개 신청 링크 등록」을 누르고, 배포한 /exec 주소를 붙여 넣습니다.'
    }
  ];

  HR_drawGuideSteps_(
    sheet,
    steps,
    '#DDA333',
    '#FFF9EF',
    '#E8D2AB'
  );

  sheet.getRange('B11:L13').merge();
  sheet
    .getRange('B11')
    .setValue(
      (yes24Ready ? '✅ 2' : '2') +
      '  API 키 설정  ·  YES24 필수 / 국립중앙도서관 선택'
    );

  sheet.getRange('B19:L21').merge();
  sheet
    .getRange('B19')
    .setValue(
      (schoolReady ? '✅ 4' : '4') +
      '  이 학교 최초 설정'
    );

  sheet.getRange('B14:L14').merge();
  sheet
    .getRange('B14')
    .setFormula(
      '=HYPERLINK("' +
      HR_YES24_API_MANAGE_URL +
      '","🔗 YES24 API 키 발급 사이트 열기")'
    )
    .setBackground('#EAF1FF')
    .setFontColor('#315F9F')
    .setFontSize(14)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setBorder(
      true,
      true,
      true,
      true,
      false,
      false,
      '#BFD0EE',
      SpreadsheetApp.BorderStyle.SOLID
    );

  ['B11:L13', 'B19:L21'].forEach(function(rangeA1) {
    sheet
      .getRange(rangeA1)
      .setBackground('#5277C7')
      .setFontColor('#FFFFFF')
      .setFontSize(18)
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setBorder(
        true,
        true,
        true,
        true,
        false,
        false,
        '#3E61AA',
        SpreadsheetApp.BorderStyle.SOLID_THICK
      );
  });

  sheet.getRange('B31:L33').merge();
  sheet
    .getRange('B31')
    .setValue(
      '⚠️ 주의사항\n' +
      '① 「희망도서 목록」 시트 이름은 변경하지 마세요. 새 목록으로 교체할 때는 기존 시트 이름을 바꾼 뒤, 새 시트를 반드시 「희망도서 목록」으로 만드세요.\n' +
      '② 「우리학교 소장도서」에는 책 제목과 ISBN13을 모두 넣어야 중복 소장 안내가 정확히 작동합니다.'
    )
    .setBackground('#FFF0F0')
    .setFontColor('#A33C3C')
    .setFontSize(14)
    .setFontWeight('bold')
    .setHorizontalAlignment('left')
    .setBorder(
      true,
      true,
      true,
      true,
      false,
      false,
      '#E8B7B7',
      SpreadsheetApp.BorderStyle.SOLID
    );

  sheet.getRange('B35:L37').merge();
  sheet
    .getRange('B35')
    .setValue(
      '현재 점검  ·  YES24 ' +
      (yes24Ready ? '완료' : '필요') +
      '  ·  KDC 참고 API ' +
      (nlkReady ? '사용 중' : '선택') +
      '  ·  신청서 연결 ' +
      (schoolReady ? '완료' : '필요') +
      '  ·  소장도서 제목/ISBN13 ' +
      (libraryStatus.headersReady ? '완료' : '필요') +
      '  ·  공개 링크 ' +
      (publicUrlReady ? '완료' : '필요')
    )
    .setBackground('#F4F6FA')
    .setFontColor('#3B4A64')
    .setFontSize(13)
    .setHorizontalAlignment('center');

  sheet.getRange('B39:L41').merge();
  sheet
    .getRange('B39')
    .setValue('🗑️ 모든 설정을 마친 뒤 이 탭 삭제')
    .setBackground('#6F7785')
    .setFontColor('#FFFFFF')
    .setFontSize(17)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  sheet.getRange('B43:L44').merge();
  sheet
    .getRange('B43')
    .setValue(
      'YES24은 도서 검색에 꼭 필요합니다. 국립중앙도서관 API는 KDC 참고용이라 없어도 됩니다.'
    )
    .setBackground('#EDF4FF')
    .setFontColor('#284E86')
    .setFontSize(13)
    .setFontWeight('bold')
    .setHorizontalAlignment('left');

  sheet
    .getRange('B3:L4')
    .setBackground('#7A5128')
    .setFontColor('#FFFFFF')
    .setFontSize(25)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  sheet
    .getRange('B5:L5')
    .setBackground('#F3E9DA')
    .setFontColor('#7A5128')
    .setFontSize(14)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  [3, 4].forEach(function(rowNumber) {
    sheet.setRowHeight(rowNumber, 36);
  });
  sheet.setRowHeight(5, 34);
  [
    7, 8, 11, 12, 13, 15, 16, 17,
    19, 20, 21, 23, 24, 25, 27, 28, 29,
    35, 36, 37, 39, 40, 41
  ].forEach(function(rowNumber) {
    sheet.setRowHeight(rowNumber, 34);
  });
  [31, 32, 33].forEach(function(rowNumber) {
    sheet.setRowHeight(rowNumber, 42);
  });
  sheet.setRowHeight(14, 32);
  [6, 9, 10, 18, 22, 26, 30, 34, 38, 42, 45]
    .forEach(function(rowNumber) {
      sheet.setRowHeight(rowNumber, 10);
    });
  [43, 44].forEach(function(rowNumber) {
    sheet.setRowHeight(rowNumber, 31);
  });

  HR_writeCopySetupBeginnerDetails_(sheet);
  HR_addSetupButtonImages_(sheet);
}

/**
 * 컴퓨터가 익숙하지 않은 사용자를 위한 상세 설명입니다.
 */
function HR_writeCopySetupBeginnerDetails_(sheet) {
  var detailSteps = [
    {
      badge: 'B50:C53',
      text: 'D50:L53',
      number: '1',
      value:
        '사본을 연 직후\n' +
        '새 파일이 열리면 5~10초 정도 기다리세요. 화면 위쪽에 「⚙️ 사본 설정」 메뉴가 보이면 준비 완료입니다. 안 보이면 브라우저의 새로고침 버튼을 한 번 누르세요.'
    },
    {
      badge: 'B55:C58',
      text: 'D55:L58',
      number: '2',
      value:
        'YES24 API 키 받기\n' +
        '위쪽의 파란 「YES24 API 키 발급 사이트 열기」를 누르세요. YES24에 로그인한 뒤 인증키 발급·관리 화면에서 API Key 전체를 복사하세요.'
    },
    {
      badge: 'B60:C63',
      text: 'D60:L63',
      number: '3',
      value:
        '복사한 API 키 붙여 넣기\n' +
        '다시 이 시트로 돌아와 위쪽 「⚙️ 사본 설정」 → 「🔑 API 키 설정 (YES24 필수)」을 누르세요. YES24 API Key 칸에 붙여 넣고 저장하세요. 국립중앙도서관 키는 비워도 됩니다.'
    },
    {
      badge: 'B65:C68',
      text: 'D65:L68',
      number: '4',
      value:
        'DLS에서 우리 학교 소장도서 반출하기\n' +
        'DLS 접속 → 「자료관리」 → 오른쪽 아래 「페이징 권수」를 우리 학교가 가진 전체 권수만큼 입력 → 「칼럼」 → 「칼럼 설정」 → 「ISBN」 체크 → 「반출」을 누르세요.\n' +
        '반출한 파일의 책 제목과 ISBN을 복사해 「우리학교 소장도서」 시트 2행부터 붙여 넣으세요. 이 정보가 있어야 중복 도서를 알려 줍니다.'
    },
    {
      badge: 'B70:C73',
      text: 'D70:L73',
      number: '5',
      value:
        '학교 이름 연결하기\n' +
        '위쪽 「⚙️ 사본 설정」 → 「🏫 학교 최초 설정」을 누르세요. 학교 이름을 입력하고 실행한 뒤, 완료 안내가 나올 때까지 기다리세요.'
    },
    {
      badge: 'B75:C80',
      text: 'D75:L80',
      number: '6',
      value:
        '신청 웹페이지 만들기\n' +
        '위쪽 「확장 프로그램」 → 「Apps Script」를 누르세요. 새 화면에서 「배포」 → 「새 배포」 → 톱니바퀴 모양 「유형 선택」 → 「웹 앱」을 고릅니다.\n' +
        '실행 사용자는 「나」, 액세스 권한은 「모든 사용자」로 정하고 「배포」를 누르세요. 마지막에 나오는 /exec 주소를 복사하세요.'
    },
    {
      badge: 'B82:C85',
      text: 'D82:L85',
      number: '7',
      value:
        '공개 신청 링크 저장하기\n' +
        '다시 스프레드시트로 돌아와 「⚙️ 사본 설정」 → 「🔗 공개 신청 링크 등록」을 누르세요. 방금 복사한 /exec 주소를 붙여 넣고 「링크 저장하기」를 누르세요.'
    },
    {
      badge: 'B87:C90',
      text: 'D87:L90',
      number: '8',
      value:
        '마지막으로 시험하기\n' +
        '「사용설명서」 탭의 「📚 희망도서 신청하기」를 눌러 보세요. 책 제목을 하나 검색해 신청해 보고, 「희망도서 목록」에 들어오면 성공입니다. 마지막에는 「⚙️ 사본 설정」 → 「✅ 설정 한 번에 점검」도 눌러 보세요.'
    }
  ];

  sheet.getRange('B47:L92').breakApart().clear();
  sheet
    .getRange('B47:L92')
    .setBackground('#FFFFFF')
    .setFontFamily('Jua')
    .setFontColor('#27344A')
    .setVerticalAlignment('middle')
    .setWrap(true);

  sheet.getRange('B47:L48').merge();
  sheet
    .getRange('B47')
    .setValue('👣 컴퓨터가 익숙하지 않다면 아래 8단계를 그대로 따라 하세요')
    .setBackground('#243B64')
    .setFontColor('#FFFFFF')
    .setFontSize(21)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  HR_drawGuideSteps_(
    sheet,
    detailSteps,
    '#5277C7',
    '#F3F6FC',
    '#D8E1F0'
  );

  [47, 48].forEach(function(rowNumber) {
    sheet.setRowHeight(rowNumber, 34);
  });

  [
    50, 51, 52, 53,
    55, 56, 57, 58,
    60, 61, 62, 63,
    65, 66, 67, 68,
    70, 71, 72, 73,
    75, 76, 77, 78, 79, 80,
    82, 83, 84, 85,
    87, 88, 89, 90
  ].forEach(function(rowNumber) {
    sheet.setRowHeight(rowNumber, 31);
  });

  [65, 66, 67, 68].forEach(function(rowNumber) {
    sheet.setRowHeight(rowNumber, 40);
  });

  [49, 54, 59, 64, 69, 74, 81, 86, 91, 92]
    .forEach(function(rowNumber) {
      sheet.setRowHeight(rowNumber, 10);
    });
}

function HR_createCopySetupSheetLegacy_(spreadsheet) {
  var sheet = spreadsheet.getSheetByName(
    HR_COPY_SETUP_SHEET_NAME
  );
  var libraryStatus = HR_getLibrarySetupStatus_(spreadsheet);
  var properties = PropertiesService.getScriptProperties();
  var configuredId = String(
    properties.getProperty(HR_SPREADSHEET_ID_PROPERTY) || ''
  ).trim();
  var yes24Ready = Boolean(HR_getYes24Key_());
  var nlkReady = Boolean(HR_getNlkKey_());
  var schoolReady = configuredId === spreadsheet.getId();

  if (!sheet) {
    sheet = spreadsheet.insertSheet(
      HR_COPY_SETUP_SHEET_NAME,
      0
    );
  }

  sheet.clear();
  sheet.clearConditionalFormatRules();
  sheet.setHiddenGridlines(true);
  sheet.setColumnWidth(1, 58);
  sheet.setColumnWidths(2, 11, 76);

  sheet.getRange('B3:L4').merge();
  sheet
    .getRange('B3')
    .setValue('⚙️ 다른 학교 사본용 최초 설정');

  sheet.getRange('B5:L5').merge();
  sheet
    .getRange('B5')
    .setValue(
      '① API 키 설정  →  ② 학교 최초 설정  →  ③ 소장도서 붙여넣기'
    );

  sheet.getRange('B7:L9').merge();
  sheet
    .getRange('B7')
    .setValue(
      (yes24Ready ? '✅' : '1️⃣') +
      ' 클릭: API 키 설정  ·  YES24 필수 / 국립중앙도서관 선택'
    );

  sheet.getRange('B12:L14').merge();
  sheet
    .getRange('B12')
    .setValue(
      (schoolReady ? '✅' : '2️⃣') +
      ' 클릭: 이 학교 최초 설정'
    );

  sheet.getRange('B17:L19').merge();
  sheet
    .getRange('B17')
    .setValue('🗑️ 설정 완료 후 이 탭 삭제');

  sheet.getRange('B21:L23').merge();
  sheet
    .getRange('B21')
    .setValue(
      '필수: 「우리학교 소장도서」 시트에 책 제목과 ISBN13을 입력해야 중복 도서 확인이 정확히 작동합니다.'
    );

  sheet.getRange('B25:L27').merge();
  sheet
    .getRange('B25')
    .setValue(
      '현재 점검  ·  YES24 API ' +
      (yes24Ready ? '완료' : '필요') +
      '  ·  KDC 참고 API ' +
      (nlkReady ? '사용 중' : '선택') +
      '  ·  신청서 연결 ' +
      (schoolReady ? '완료' : '필요') +
      '  ·  소장도서 제목/ISBN 열 ' +
      (libraryStatus.headersReady ? '완료' : '필요')
    );

  sheet
    .getRange('B3:L27')
    .setFontFamily('Jua')
    .setVerticalAlignment('middle')
    .setWrap(true);

  sheet
    .getRange('B3:L4')
    .setBackground('#243B64')
    .setFontColor('#FFFFFF')
    .setFontSize(25)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  sheet
    .getRange('B5:L5')
    .setBackground('#E9EEF8')
    .setFontColor('#243B64')
    .setFontSize(13)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  ['B7:L9', 'B12:L14'].forEach(function(rangeA1) {
    sheet
      .getRange(rangeA1)
      .setBackground('#5277C7')
      .setFontColor('#FFFFFF')
      .setFontSize(18)
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setBorder(
        true,
        true,
        true,
        true,
        false,
        false,
        '#3E61AA',
        SpreadsheetApp.BorderStyle.SOLID_THICK
      );
  });

  sheet
    .getRange('B17:L19')
    .setBackground('#6F7785')
    .setFontColor('#FFFFFF')
    .setFontSize(17)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  sheet
    .getRange('B21:L23')
    .setBackground('#FFF0F0')
    .setFontColor('#A33C3C')
    .setFontSize(15)
    .setFontWeight('bold')
    .setHorizontalAlignment('left')
    .setBorder(
      true,
      true,
      true,
      true,
      false,
      false,
      '#E8B7B7',
      SpreadsheetApp.BorderStyle.SOLID
    );

  sheet
    .getRange('B25:L27')
    .setBackground('#F4F6FA')
    .setFontColor('#3B4A64')
    .setFontSize(13)
    .setHorizontalAlignment('center');

  [3, 4].forEach(function(rowNumber) {
    sheet.setRowHeight(rowNumber, 36);
  });
  [7, 8, 9, 12, 13, 14, 17, 18, 19].forEach(
    function(rowNumber) {
      sheet.setRowHeight(rowNumber, 34);
    }
  );
  [21, 22, 23, 25, 26, 27].forEach(
    function(rowNumber) {
      sheet.setRowHeight(rowNumber, 31);
    }
  );

  HR_addSetupButtonImages_(sheet);
}

/**
 * 투명한 이미지 버튼을 셀 위에 올려 파란 영역 전체를 클릭할 수 있게 합니다.
 * 사본을 만들면 버튼에 연결된 함수도 함께 복사됩니다.
 */
function HR_prepareTopSetupButtons_(sheet) {
  var lastRow = sheet.getLastRow();
  var lastColumn = Math.max(sheet.getLastColumn(), 15);
  var values = sheet.getRange(1, 1, lastRow, lastColumn).getDisplayValues();
  var startRow = 0;
  var endRow = 0;
  var headerText = '🧰 설정 버튼 모음  ·  필요한 버튼을 순서대로 누르세요';
  var apiReady = false;
  var schoolReady = false;
  var publicReady = false;

  values.forEach(function(row, index) {
    row.forEach(function(cellValue) {
      var cellText = String(cellValue || '').trim();
      if (/설정 버튼 모음/.test(cellText)) {
        if (!startRow) startRow = index + 1;
        headerText = cellText;
      }
      if (/API 키 설정/.test(cellText) && /✅/.test(cellText)) apiReady = true;
      if (/(?:이 학교|학교) 최초 설정/.test(cellText) && /✅/.test(cellText)) schoolReady = true;
      if (/공개 (?:신청 )?링크 등록/.test(cellText)) {
        if (/✅/.test(cellText)) publicReady = true;
        endRow = Math.max(endRow, index + 1);
      }
    });
  });

  if (startRow > 8 && endRow >= startRow) {
    var merged = sheet.getRange(endRow, 1, 1, lastColumn).getMergedRanges();
    merged.forEach(function(range) {
      endRow = Math.max(endRow, range.getLastRow());
    });
    endRow = Math.min(lastRow, endRow + 1);
    sheet.moveRows(
      sheet.getRange(startRow, 1, endRow - startRow + 1, sheet.getMaxColumns()),
      7
    );
  }

  HR_breakMergedRangesSafely_(sheet, 'B7:O20');
  sheet.getRange('B7:O20').clearContent();
  sheet.getRange('B7:L8').merge().setValue(headerText);
  sheet.getRange('B10:D13').merge().setValue((apiReady ? '✅ ' : '') + '🔑 API 키 설정');
  sheet.getRange('F10:H13').merge().setValue((schoolReady ? '✅ ' : '') + '🏫 학교 최초 설정');
  sheet.getRange('J10:L13').merge().setValue((publicReady ? '✅ ' : '') + '🔗 공개 링크 등록');

  sheet.setColumnWidths(2, 3, 100);
  sheet.setColumnWidth(5, 28);
  sheet.setColumnWidths(6, 3, 100);
  sheet.setColumnWidth(9, 28);
  sheet.setColumnWidths(10, 3, 100);

  [10, 11, 12, 13].forEach(function(rowNumber) {
    sheet.setRowHeight(rowNumber, 27);
  });
  sheet.getRange('B14:L20')
    .setBackground('#F8F5ED')
    .setBorder(false, false, false, false, false, false);
  sheet.setRowHeight(9, 8);
  for (var compactRow = 14; compactRow <= 20; compactRow++) {
    sheet.setRowHeight(compactRow, 4);
  }
}

function HR_styleCopySetupModern_(sheet) {
  if (!sheet) return;

  var lastRow = Math.max(sheet.getLastRow(), 82);
  var lastColumn = Math.max(sheet.getLastColumn(), 12);
  var canvas = sheet.getRange(1, 1, lastRow, lastColumn);

  sheet.setHiddenGridlines(true);
  sheet.setFrozenRows(0);
  sheet.setFrozenColumns(0);

  canvas
    .setBackground('#F8F5ED')
    .setFontFamily('Arial')
    .setFontColor('#2E3A33')
    .setFontSize(12)
    .setVerticalAlignment('middle')
    .setWrap(true);

  sheet.setColumnWidth(1, 28);
  sheet.setColumnWidth(2, 78);
  for (var column = 3; column <= lastColumn; column++) {
    sheet.setColumnWidth(column, column <= 12 ? 86 : 72);
  }
  sheet.setColumnWidths(2, 3, 100);
  sheet.setColumnWidth(5, 28);
  sheet.setColumnWidths(6, 3, 100);
  sheet.setColumnWidth(9, 28);
  sheet.setColumnWidths(10, 3, 100);

  canvas.getMergedRanges().forEach(function(range) {
    var text = String(range.getDisplayValue() || '').trim();

    range
      .setBackground('#FFFFFF')
      .setFontColor('#2E3A33')
      .setFontFamily('Arial')
      .setFontSize(12)
      .setVerticalAlignment('middle')
      .setWrap(true)
      .setBorder(
        true, true, true, true, false, false,
        '#D9E2DC', SpreadsheetApp.BorderStyle.SOLID
      );

    if (!text) {
      range
        .setBackground('#F8F5ED')
        .setBorder(false, false, false, false, false, false);
      return;
    }

    if (/다른 학교.*최초 설정|사본용 최초 설정/.test(text)) {
      range
        .setBackground('#183D35')
        .setFontColor('#FFFFFF')
        .setFontSize(25)
        .setFontWeight('bold')
        .setHorizontalAlignment('center');
      return;
    }

    if (/처음 한 번만/.test(text)) {
      range
        .setBackground('#EAF3ED')
        .setFontColor('#34564D')
        .setFontSize(13)
        .setFontWeight('bold')
        .setHorizontalAlignment('center');
      return;
    }

    if (/컴퓨터가 익숙하지 않아도|설정 버튼 모음/.test(text)) {
      range
        .setBackground('#D9A441')
        .setFontColor('#FFFFFF')
        .setFontSize(17)
        .setFontWeight('bold')
        .setHorizontalAlignment('center');
      return;
    }

    if (/^\d+$/.test(text)) {
      range
        .setBackground('#D9A441')
        .setFontColor('#FFFFFF')
        .setFontSize(24)
        .setFontWeight('bold')
        .setHorizontalAlignment('center');
      return;
    }

    if (/API 키 설정|이 학교 최초 설정|학교 최초 설정|공개 신청 링크 등록|공개 링크 등록/.test(text)) {
      range
        .setBackground('#527D6B')
        .setFontColor('#FFFFFF')
        .setFontSize(16)
        .setFontWeight('bold')
        .setHorizontalAlignment('center')
        .setBorder(
          true, true, true, true, false, false,
          '#315D4D', SpreadsheetApp.BorderStyle.SOLID_MEDIUM
        );
      return;
    }

    if (/사이트 열기|메뉴 확인/.test(text)) {
      range
        .setBackground('#E7F0EC')
        .setFontColor('#285244')
        .setFontWeight('bold')
        .setHorizontalAlignment('center');
      return;
    }

    if (/꼭 확인|주의|변경 금지|삭제하지/.test(text)) {
      range
        .setBackground('#FCEDEC')
        .setFontColor('#A33F3F')
        .setFontWeight('bold');
      return;
    }

    if (/현재 점검/.test(text)) {
      range
        .setBackground('#EEF4F1')
        .setFontColor('#34564D')
        .setFontWeight('bold')
        .setHorizontalAlignment('center');
    }
  });

  sheet.activate();
  sheet.setActiveRange(sheet.getRange('B3'));
}

function HR_addSetupButtonImages_(sheet) {
  var transparentPngBase64 =
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJ' +
    'AAAACXBIWXMAAAPoAAAD6AG1e1JrAAAADUlEQVR4nGP4' +
    '//8/MwAI/wMBlb/HmAAAAABJRU5ErkJggg==';
  var buttonBlob = Utilities.newBlob(
    Utilities.base64Decode(transparentPngBase64),
    'image/png',
    'copy-setup-button.png'
  );
  var images = sheet.getImages();
  var imageIndex = 0;

  HR_prepareTopSetupButtons_(sheet);

  for (imageIndex = images.length - 1; imageIndex >= 0; imageIndex--) {
    if (
      String(images[imageIndex].getAltTextTitle() || '')
        .indexOf('HR_COPY_SETUP_BUTTON') === 0
    ) {
      images[imageIndex].remove();
    }
  }

  [
    {
      cell: 'B10',
      width: 300,
      height: 104,
      title: 'API 키 설정',
      functionName: 'HR_showYes24KeyDialog'
    },
    {
      cell: 'F10',
      width: 300,
      height: 104,
      title: '학교 최초 설정',
      functionName: 'HR_showSchoolSetupDialog'
    },
    {
      cell: 'J10',
      width: 300,
      height: 104,
      title: '공개 신청 링크 등록',
      functionName: 'HR_showPublicWebAppUrlDialog'
    },
    {
      cell: 'B77',
      width: 836,
      height: 102,
      title: '설정 탭 삭제',
      functionName: 'HR_deleteCopySetupSheet'
    }
  ].forEach(function(config, index) {
    var anchor = sheet.getRange(config.cell);
    var image = sheet.insertImage(
      buttonBlob,
      anchor.getColumn(),
      anchor.getRow()
    );

    image
      .setWidth(config.width)
      .setHeight(config.height)
      .setAltTextTitle('HR_COPY_SETUP_BUTTON_' + (index + 1))
      .setAltTextDescription(config.title)
      .assignScript(config.functionName);
  });

  HR_styleCopySetupModern_(sheet);
}

function HR_getSetupSpreadsheet_() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  if (spreadsheet) {
    return spreadsheet;
  }

  return HR_getSpreadsheet_();
}

function HR_ensureLibrarySheet_(spreadsheet) {
  var sheet = spreadsheet.getSheetByName(
    HR_LIBRARY_SHEET_NAME
  );
  var lastColumn;
  var headers;
  var titleColumn;
  var isbnColumn;

  if (!sheet) {
    sheet = spreadsheet.insertSheet(HR_LIBRARY_SHEET_NAME);
  }

  lastColumn = Math.max(sheet.getLastColumn(), 1);
  headers = sheet
    .getRange(1, 1, 1, lastColumn)
    .getDisplayValues()[0];
  titleColumn = HR_findAnyHeaderColumn_(
    headers,
    ['책제목', '도서명', '서명', '제목']
  );
  isbnColumn = HR_findAnyHeaderColumn_(
    headers,
    ['ISBN13', 'ISBN']
  );

  if (!titleColumn) {
    lastColumn = Math.max(sheet.getLastColumn(), 0) + 1;
    sheet.getRange(1, lastColumn).setValue('책 제목');
  }

  if (!isbnColumn) {
    lastColumn = Math.max(sheet.getLastColumn(), 0) + 1;
    sheet.getRange(1, lastColumn).setValue('ISBN13');
  }

  sheet
    .getRange(1, 1, 1, sheet.getLastColumn())
    .setBackground('#243B64')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  sheet.setFrozenRows(1);
  sheet.setColumnWidths(1, sheet.getLastColumn(), 180);
}

function HR_getLibrarySetupStatus_(spreadsheet) {
  var sheet = spreadsheet.getSheetByName(
    HR_LIBRARY_SHEET_NAME
  );
  var result = {
    sheetReady: Boolean(sheet),
    headersReady: false,
    titleColumn: 0,
    isbnColumn: 0,
    bookCount: 0
  };
  var lastColumn;
  var headers;

  if (!sheet) {
    return result;
  }

  lastColumn = Math.max(sheet.getLastColumn(), 1);
  headers = sheet
    .getRange(1, 1, 1, lastColumn)
    .getDisplayValues()[0];
  result.titleColumn = HR_findAnyHeaderColumn_(
    headers,
    ['책제목', '도서명', '서명', '제목']
  );
  result.isbnColumn = HR_findAnyHeaderColumn_(
    headers,
    ['ISBN13', 'ISBN']
  );
  result.headersReady = Boolean(
    result.titleColumn && result.isbnColumn
  );
  result.bookCount = Math.max(sheet.getLastRow() - 1, 0);

  return result;
}

function HR_installEditTriggerForSpreadsheet_(spreadsheet) {
  var triggers = ScriptApp.getProjectTriggers();
  var triggerIndex = 0;

  for (
    triggerIndex = 0;
    triggerIndex < triggers.length;
    triggerIndex++
  ) {
    if (
      triggers[triggerIndex].getHandlerFunction() ===
      'onEditTriggerAction'
    ) {
      ScriptApp.deleteTrigger(triggers[triggerIndex]);
    }
  }

  ScriptApp
    .newTrigger('onEditTriggerAction')
    .forSpreadsheet(spreadsheet)
    .onEdit()
    .create();
}

function HR_testYes24Key_(key) {
  yes24Get_('/goods/itemList', {query: '아몬드', category: 'BOOK', pageSize: 1}, key);
}

/**
 * 학교별 필수 설정을 한 번에 확인합니다.
 */
function HR_getSystemCheckState_() {
  var spreadsheet = HR_getSetupSpreadsheet_();
  var properties = PropertiesService.getScriptProperties();
  var targetSheet = spreadsheet.getSheetByName(
    HR_TARGET_SHEET_NAME
  );
  var libraryStatus = HR_getLibrarySetupStatus_(spreadsheet);
  var configuredId = String(
    properties.getProperty(HR_SPREADSHEET_ID_PROPERTY) || ''
  ).trim();
  var publicUrl = HR_getConfiguredPublicWebAppUrl_();
  var checks = [];

  checks.push({
    label: 'YES24 API 키',
    ok: Boolean(HR_getYes24Key_()),
    detail: HR_getYes24Key_() ? '설정됨' : '설정 필요'
  });
  checks.push({
    label: '학교 신청서 연결',
    ok: configuredId === spreadsheet.getId(),
    detail: configuredId === spreadsheet.getId()
      ? '현재 파일에 연결됨'
      : '학교 최초 설정 필요'
  });
  checks.push({
    label: '📚 희망도서 목록',
    ok: Boolean(targetSheet),
    detail: targetSheet ? '정상' : '시트 없음'
  });
  checks.push({
    label: '🏫 우리학교 소장도서',
    ok: Boolean(
      libraryStatus.sheetReady &&
      libraryStatus.headersReady
    ),
    detail: libraryStatus.headersReady
      ? libraryStatus.bookCount + '권 확인'
      : '책 제목·ISBN13 열 확인 필요'
  });
  checks.push({
    label: '공개 신청 링크',
    ok: /\/exec(?:$|\?)/.test(publicUrl),
    detail: /\/exec(?:$|\?)/.test(publicUrl)
      ? '누구나 신청 가능한 주소'
      : '공개 /exec 주소 확인 필요'
  });
  return {
    checks: checks,
    allRequiredReady: checks
      .slice(0, 5)
      .every(function(check) {
        return check.ok;
      })
  };
}

function HR_showSystemCheck() {
  var state = HR_getSystemCheckState_();
  var checkHtml = state.checks
    .map(function(check) {
      return (
        '<div class="item ' +
        (check.ok ? 'ok' : 'warn') +
        '">' +
          '<span class="icon">' +
          (check.ok ? '✓' : '!') +
          '</span>' +
          '<div><strong>' +
          HR_escapeHtml_(check.label) +
          '</strong><p>' +
          HR_escapeHtml_(check.detail) +
          '</p></div>' +
        '</div>'
      );
    })
    .join('');
  var html =
    '<!doctype html><html lang="ko"><head><base target="_top">' +
    '<style>' +
    '*{box-sizing:border-box}body{margin:0;padding:22px;background:#f5f7fb;' +
    'font-family:-apple-system,BlinkMacSystemFont,"Apple SD Gothic Neo",' +
    '"Noto Sans KR",sans-serif;color:#24324a}' +
    'h1{margin:0 0 8px;font-size:24px}.lead{margin:0 0 18px;color:#66758e;' +
    'line-height:1.6}.list{display:grid;gap:10px}.item{display:flex;gap:12px;' +
    'align-items:center;padding:14px;border-radius:14px;background:#fff;' +
    'border:1px solid #dce4f2}.icon{display:flex;width:34px;height:34px;' +
    'flex:0 0 auto;align-items:center;justify-content:center;border-radius:12px;' +
    'font-weight:900}.ok .icon{background:#eaf7ef;color:#237a4b}' +
    '.warn .icon{background:#fff3df;color:#a46608}.item strong{font-size:15px}' +
    '.item p{margin:4px 0 0;color:#6b7688;font-size:13px}.summary{margin-top:16px;' +
    'padding:15px;border-radius:14px;font-weight:900;text-align:center}' +
    '.ready{background:#eaf7ef;color:#237a4b}.check{background:#fff3df;color:#8a5a12}' +
    'button{width:100%;height:48px;margin-top:16px;border:0;border-radius:12px;' +
    'background:#4f78d1;color:white;font-size:16px;font-weight:900;cursor:pointer}' +
    '</style></head><body>' +
    '<h1>✅ 설정 한 번에 점검</h1>' +
    '<p class="lead">학교 신청 시스템의 필수 설정을 한눈에 확인합니다.</p>' +
    '<div class="list">' + checkHtml + '</div>' +
    '<div class="summary ' +
    (state.allRequiredReady ? 'ready' : 'check') +
    '">' +
    (state.allRequiredReady
      ? '필수 설정이 모두 정상입니다.'
      : '표시된 항목을 확인해 주세요.') +
    '</div><button onclick="google.script.host.close()">확인</button>' +
    '</body></html>';

  SpreadsheetApp.getUi().showModalDialog(
    HtmlService
      .createHtmlOutput(html)
      .setWidth(520)
      .setHeight(660),
    '설정 점검'
  );
}

function HR_escapeHtml_(value) {
  return String(value || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

/**
 * ============================================================
 * 공개 신청 웹앱 베스트셀러
 * - 종합 / 청소년 / 어린이
 * - YES24 베스트셀러 API
 * - 24시간 단위 갱신
 * ============================================================
 */

/**
 * 공개 신청 페이지에 최근 희망도서 신청 목록을 제공합니다.
 * 신청자 이름은 개인정보 보호를 위해 절대 반환하지 않습니다.
 */
function HR_getSubmittedRequests(audience) {
  var audienceKey = HR_audienceKey_(audience);
  var spreadsheet = HR_getSpreadsheet_();
  var sheet = spreadsheet.getSheetByName(
    HR_TARGET_SHEET_NAME
  );
  var headers;
  var lastRow;
  var lastColumn;
  var titleColumn;
  var authorColumn;
  var publisherColumn;
  var yearColumn;
  var statusColumn;
  var dateColumn;
  var sequenceColumn;
  var isbnColumn;
  var priceColumn;
  var categoryColumn;
  var pagesColumn;
  var linkColumn;
  var coverColumn;
  var values;
  var formulas;
  var items = [];
  var rowIndex;

  if (!sheet) {
    throw new Error(
      HR_TARGET_SHEET_NAME + ' 시트를 찾지 못했습니다.'
    );
  }

  lastRow = sheet.getLastRow();
  lastColumn = sheet.getLastColumn();

  if (lastRow < FIRST_DATA_ROW || lastColumn < 1) {
    return {
      totalCount: 0,
      shownCount: 0,
      items: [],
      updatedAt: new Date().toISOString()
    };
  }

  headers = getHeaders(sheet);
  titleColumn = HR_findAnyHeaderColumn_(
    headers,
    ['도서명', '책 제목', '책제목', '서명', '제목']
  );
  authorColumn = HR_findAnyHeaderColumn_(
    headers,
    ['저자', '지은이']
  );
  publisherColumn = HR_findAnyHeaderColumn_(
    headers,
    ['출판사', '발행처']
  );
  yearColumn = HR_findAnyHeaderColumn_(
    headers,
    ['연도', '출판연도', '발행연도']
  );
  statusColumn = HR_findAnyHeaderColumn_(
    headers,
    ['처리상태', '처리 상태', '상태']
  );
  dateColumn = HR_findAnyHeaderColumn_(
    headers,
    ['신청일', '신청일자', '접수일', '등록일']
  );
  sequenceColumn = HR_findAnyHeaderColumn_(
    headers,
    ['순번', '번호']
  );
  isbnColumn = HR_findAnyHeaderColumn_(
    headers,
    ['ISBN13', 'ISBN', 'isbn13', 'isbn']
  );
  priceColumn = HR_findAnyHeaderColumn_(
    headers,
    ['정가', '가격', '판매가']
  );
  categoryColumn = HR_findAnyHeaderColumn_(
    headers,
    ['분야', '카테고리', '주제', 'KDC']
  );
  pagesColumn = HR_findAnyHeaderColumn_(
    headers,
    ['쪽수', '페이지', '전체 페이지']
  );
  linkColumn = HR_findAnyHeaderColumn_(
    headers,
    ['YES24 링크', '도서 링크', '링크']
  );
  coverColumn = HR_findAnyHeaderColumn_(
    headers,
    ['표지', '책 표지', '표지 URL', '표지URL', '이미지']
  );

  if (!titleColumn) {
    throw new Error(
      '희망도서 목록에서 도서명 열을 찾지 못했습니다.'
    );
  }

  values = sheet
    .getRange(
      FIRST_DATA_ROW,
      1,
      lastRow - FIRST_DATA_ROW + 1,
      lastColumn
    )
    .getDisplayValues();

  formulas = sheet
    .getRange(
      FIRST_DATA_ROW,
      1,
      lastRow - FIRST_DATA_ROW + 1,
      lastColumn
    )
    .getFormulas();

  var audienceColumn = HR_findAudienceColumn_(headers);

  for (rowIndex = 0; rowIndex < values.length; rowIndex++) {
    var row = values[rowIndex];
    var title = String(
      row[titleColumn - 1] || ''
    ).trim();

    if (!title) {
      continue;
    }

    // 다른 구분(교직원/학부모)의 신청은 보여 주지 않습니다. 구분이 비어 있으면 교직원입니다.
    if (
      (audienceColumn ? HR_audienceKey_(row[audienceColumn - 1]) : 'staff') !==
      audienceKey
    ) {
      continue;
    }

    items.push({
      row: rowIndex + FIRST_DATA_ROW,
      sequence: sequenceColumn
        ? String(row[sequenceColumn - 1] || '').trim()
        : String(rowIndex + 1),
      title: title,
      author: authorColumn
        ? String(row[authorColumn - 1] || '').trim()
        : '',
      publisher: publisherColumn
        ? String(row[publisherColumn - 1] || '').trim()
        : '',
      year: yearColumn
        ? String(row[yearColumn - 1] || '').trim()
        : '',
      status: statusColumn
        ? String(row[statusColumn - 1] || '').trim()
        : '신청 접수',
      requestedAt: dateColumn
        ? String(row[dateColumn - 1] || '').trim()
        : '',
      isbn: isbnColumn
        ? String(row[isbnColumn - 1] || '').trim()
        : '',
      price: priceColumn
        ? String(row[priceColumn - 1] || '').trim()
        : '',
      category: categoryColumn
        ? String(row[categoryColumn - 1] || '').trim()
        : '',
      pages: pagesColumn
        ? String(row[pagesColumn - 1] || '').trim()
        : '',
      link: linkColumn
        ? String(row[linkColumn - 1] || '').trim()
        : '',
      cover: coverColumn
        ? HR_extractRequestCoverUrl_(
            formulas[rowIndex][coverColumn - 1],
            row[coverColumn - 1]
          )
        : ''
    });
  }

  items = items.concat(
    HR_getPendingCancelledItems_(
      spreadsheet,
      headers,
      audienceKey
    )
  );

  return {
    totalCount: items.length,
    shownCount: Math.min(items.length, 100),
    items: items.slice(-100).reverse(),
    updatedAt: new Date().toISOString()
  };
}


function HR_getPendingCancelledItems_(
  spreadsheet,
  sourceHeaders,
  audienceKey
) {
  var backupSheet = spreadsheet.getSheetByName(
    HR_REQUEST_CANCEL_BACKUP_SHEET_NAME
  );
  var headers = Array.isArray(sourceHeaders)
    ? sourceHeaders
    : [];
  var items = [];
  var lastRow;
  var readWidth;
  var values;
  var formulas;
  var titleColumn;
  var authorColumn;
  var publisherColumn;
  var yearColumn;
  var dateColumn;
  var sequenceColumn;
  var isbnColumn;
  var priceColumn;
  var categoryColumn;
  var pagesColumn;
  var linkColumn;
  var coverColumn;
  var rowIndex;

  if (
    !backupSheet ||
    backupSheet.getLastRow() < 2 ||
    backupSheet.getLastColumn() < 8
  ) {
    return items;
  }

  lastRow = backupSheet.getLastRow();
  readWidth = Math.min(
    backupSheet.getLastColumn(),
    8 + headers.length
  );
  values = backupSheet
    .getRange(2, 1, lastRow - 1, readWidth)
    .getDisplayValues();
  formulas = backupSheet
    .getRange(2, 1, lastRow - 1, readWidth)
    .getFormulas();

  titleColumn = HR_findAnyHeaderColumn_(
    headers,
    ['도서명', '책 제목', '책제목', '서명', '제목']
  );
  authorColumn = HR_findAnyHeaderColumn_(
    headers,
    ['저자', '지은이']
  );
  publisherColumn = HR_findAnyHeaderColumn_(
    headers,
    ['출판사', '발행처']
  );
  yearColumn = HR_findAnyHeaderColumn_(
    headers,
    ['연도', '출판연도', '발행연도']
  );
  dateColumn = HR_findAnyHeaderColumn_(
    headers,
    ['신청일', '신청일자', '접수일', '등록일']
  );
  sequenceColumn = HR_findAnyHeaderColumn_(
    headers,
    ['순번', '번호']
  );
  isbnColumn = HR_findAnyHeaderColumn_(
    headers,
    ['ISBN13', 'ISBN', 'isbn13', 'isbn']
  );
  priceColumn = HR_findAnyHeaderColumn_(
    headers,
    ['정가', '가격', '판매가']
  );
  categoryColumn = HR_findAnyHeaderColumn_(
    headers,
    ['분야', '카테고리', '주제', 'KDC']
  );
  pagesColumn = HR_findAnyHeaderColumn_(
    headers,
    ['쪽수', '페이지', '전체 페이지']
  );
  linkColumn = HR_findAnyHeaderColumn_(
    headers,
    ['YES24 링크', '도서 링크', '링크']
  );
  coverColumn = HR_findAnyHeaderColumn_(
    headers,
    ['표지', '책 표지', '표지 URL', '표지URL', '이미지']
  );

  for (
    rowIndex = 0;
    rowIndex < values.length;
    rowIndex++
  ) {
    var backupRow = values[rowIndex];
    var backupFormulaRow = formulas[rowIndex];
    var backupStatus = String(
      backupRow[6] || ''
    ).trim();
    var sourceRow = backupRow.slice(
      8,
      8 + headers.length
    );
    var sourceFormulaRow = backupFormulaRow.slice(
      8,
      8 + headers.length
    );
    var title = String(
      backupRow[2] ||
      (
        titleColumn
          ? sourceRow[titleColumn - 1]
          : ''
      ) ||
      ''
    ).trim();

    if (
      !/^취소/.test(backupStatus) ||
      !title
    ) {
      continue;
    }

    var pendingAudienceColumn = HR_findAudienceColumn_(headers);
    var pendingAudience = pendingAudienceColumn
      ? HR_audienceKey_(sourceRow[pendingAudienceColumn - 1])
      : 'staff';
    if (pendingAudience !== HR_audienceKey_(audienceKey)) {
      continue;
    }

    items.push({
      row: 0,
      sequence: sequenceColumn
        ? String(
          sourceRow[sequenceColumn - 1] || '취소'
        ).trim()
        : '취소',
      title: title,
      author: authorColumn
        ? String(sourceRow[authorColumn - 1] || '').trim()
        : '',
      publisher: publisherColumn
        ? String(sourceRow[publisherColumn - 1] || '').trim()
        : '',
      year: yearColumn
        ? String(sourceRow[yearColumn - 1] || '').trim()
        : '',
      status: '🚫 신청 취소',
      requestedAt: dateColumn
        ? String(sourceRow[dateColumn - 1] || '').trim()
        : String(backupRow[1] || '').trim(),
      isbn: String(
        backupRow[3] ||
        (
          isbnColumn
            ? sourceRow[isbnColumn - 1]
            : ''
        ) ||
        ''
      ).trim(),
      price: priceColumn
        ? String(sourceRow[priceColumn - 1] || '').trim()
        : '',
      category: categoryColumn
        ? String(sourceRow[categoryColumn - 1] || '').trim()
        : '',
      pages: pagesColumn
        ? String(sourceRow[pagesColumn - 1] || '').trim()
        : '',
      link: linkColumn
        ? String(sourceRow[linkColumn - 1] || '').trim()
        : '',
      cover: coverColumn
        ? HR_extractRequestCoverUrl_(
          sourceFormulaRow[coverColumn - 1],
          sourceRow[coverColumn - 1]
        )
        : '',
      cancelId: String(backupRow[0] || '').trim(),
      cancelledBackup: true
    });
  }

  return items;
}

/**
 * 예전 방식으로 상태만 취소 처리된 행을 다시 접수 상태로 바꿉니다.
 */
function HR_restoreLegacyCancelledRequest(requestData) {
  var rowNumber = Number(
    requestData && requestData.row || 0
  );
  var requestedTitle = String(
    requestData && requestData.title || ''
  ).trim();
  var requestedIsbn = HR_normalizeIsbn_(
    requestData && requestData.isbn || ''
  );
  var spreadsheet = HR_getSpreadsheet_();
  var sheet = spreadsheet.getSheetByName(
    HR_TARGET_SHEET_NAME
  );
  var lock = LockService.getScriptLock();
  var lockAcquired = false;
  var headers;
  var titleColumn;
  var isbnColumn;
  var statusColumn;
  var currentTitle;
  var currentIsbn;
  var currentStatus;

  if (!sheet) {
    throw new Error(
      HR_TARGET_SHEET_NAME + ' 시트를 찾지 못했습니다.'
    );
  }

  try {
    lockAcquired = lock.tryLock(
      HR_REQUEST_LOCK_WAIT_MS
    );

    if (!lockAcquired) {
      throw new Error(
        '현재 이용자가 많습니다. 잠시 후 다시 시도해 주세요.'
      );
    }

    if (
      !rowNumber ||
      rowNumber < 2 ||
      rowNumber > sheet.getLastRow()
    ) {
      throw new Error(
        '신청 목록이 변경되었습니다. 새로고침한 뒤 다시 시도해 주세요.'
      );
    }

    headers = getHeaders(sheet);
    titleColumn = HR_findAnyHeaderColumn_(
      headers,
      ['도서명', '책 제목', '책제목', '서명', '제목']
    );
    isbnColumn = HR_findAnyHeaderColumn_(
      headers,
      ['ISBN13', 'ISBN 13', 'ISBN', '국제표준도서번호']
    );
    statusColumn = HR_findAnyHeaderColumn_(
      headers,
      ['처리상태', '처리 상태', '상태']
    );

    if (!titleColumn || !statusColumn) {
      throw new Error(
        '희망도서 목록에서 도서명 또는 처리상태 열을 찾지 못했습니다.'
      );
    }

    currentTitle = String(
      sheet
        .getRange(rowNumber, titleColumn)
        .getDisplayValue() || ''
    ).trim();
    currentIsbn = isbnColumn
      ? HR_normalizeIsbn_(
        sheet
          .getRange(rowNumber, isbnColumn)
          .getDisplayValue()
      )
      : '';
    currentStatus = String(
      sheet
        .getRange(rowNumber, statusColumn)
        .getDisplayValue() || ''
    ).trim();

    if (
      !currentTitle ||
      HR_normalizeBookTitle_(currentTitle) !==
        HR_normalizeBookTitle_(requestedTitle) ||
      (
        requestedIsbn &&
        currentIsbn !== requestedIsbn
      )
    ) {
      throw new Error(
        '신청 목록이 변경되었습니다. 새로고침한 뒤 다시 시도해 주세요.'
      );
    }

    if (!/취소/.test(currentStatus)) {
      return {
        success: true,
        alreadyRestored: true,
        title: currentTitle,
        isbn: currentIsbn,
        status: currentStatus
      };
    }

    sheet
      .getRange(rowNumber, statusColumn)
      .setValue('✅ 접수 완료');

    SpreadsheetApp.flush();

    return {
      success: true,
      restored: true,
      title: currentTitle,
      isbn: currentIsbn,
      row: rowNumber,
      status: '✅ 접수 완료'
    };
  } finally {
    if (lockAcquired) {
      lock.releaseLock();
    }
  }
}

/**
 * 표지 셀의 IMAGE 수식 또는 URL에서 실제 이미지 주소를 꺼냅니다.
 * 저장된 값을 읽기만 하므로 YES24 API를 추가 호출하지 않습니다.
 */
function HR_extractRequestCoverUrl_(
  formula,
  displayValue
) {
  var formulaText = String(formula || '');
  var displayText = String(displayValue || '').trim();
  var match = formulaText.match(
    /IMAGE\("([^"]+)"/i
  );

  if (match && match[1]) {
    return match[1].replace(/""/g, '"');
  }

  return /^https?:\/\//i.test(displayText)
    ? displayText
    : '';
}


function HR_getOrCreateCancelBackup_(
  spreadsheet,
  sourceSheet,
  headers
) {
  var sheet = spreadsheet.getSheetByName(
    HR_REQUEST_CANCEL_BACKUP_SHEET_NAME
  );
  var metadataHeaders = [
    '취소ID',
    '취소시각',
    '도서명',
    'ISBN13',
    '원래행',
    '원본열수',
    '복구상태',
    '복구시각'
  ];
  var sourceHeaders = Array.isArray(headers)
    ? headers
    : [];
  var requiredColumns =
    metadataHeaders.length + sourceHeaders.length;

  if (!sheet) {
    sheet = spreadsheet.insertSheet(
      HR_REQUEST_CANCEL_BACKUP_SHEET_NAME
    );
  }

  if (sheet.getMaxColumns() < requiredColumns) {
    sheet.insertColumnsAfter(
      sheet.getMaxColumns(),
      requiredColumns - sheet.getMaxColumns()
    );
  }

  sheet
    .getRange(1, 1, 1, requiredColumns)
    .setValues([metadataHeaders.concat(sourceHeaders)])
    .setFontWeight('bold')
    .setBackground('#efe8dc');

  sheet.setFrozenRows(1);

  if (!sheet.isSheetHidden()) {
    sheet.hideSheet();
  }

  return sheet;
}

/**
 * 공개 신청 화면에서 해당 도서 한 건만 안전하게 취소합니다.
 * 원본 행은 숨김 백업 시트에 보관한 뒤 목록에서 삭제합니다.
 */
function HR_cancelSubmittedRequest(requestData) {
  var rowNumber = Number(
    requestData && requestData.row || 0
  );
  var requestedTitle = String(
    requestData && requestData.title || ''
  ).trim();
  var requestedIsbn = HR_normalizeIsbn_(
    requestData && requestData.isbn || ''
  );
  var spreadsheet = HR_getSpreadsheet_();
  var sheet = spreadsheet.getSheetByName(
    HR_TARGET_SHEET_NAME
  );
  var lock = LockService.getScriptLock();
  var lockAcquired = false;
  var headers;
  var titleColumn;
  var isbnColumn;
  var statusColumn;
  var currentTitle;
  var currentIsbn;
  var currentStatus;
  var originalColumnCount;
  var backupSheet;
  var backupRow;
  var cancelId;

  if (!sheet) {
    throw new Error(
      HR_TARGET_SHEET_NAME + ' 시트를 찾지 못했습니다.'
    );
  }

  try {
    lockAcquired = lock.tryLock(
      HR_REQUEST_LOCK_WAIT_MS
    );

    if (!lockAcquired) {
      throw new Error(
        '현재 이용자가 많습니다. 잠시 후 다시 시도해 주세요.'
      );
    }

    if (
      !rowNumber ||
      rowNumber < FIRST_DATA_ROW ||
      rowNumber > sheet.getLastRow()
    ) {
      throw new Error(
        '신청 목록이 변경되었습니다. 새로고침한 뒤 다시 시도해 주세요.'
      );
    }

    headers = getHeaders(sheet);
    titleColumn = HR_findAnyHeaderColumn_(
      headers,
      ['도서명', '책 제목', '책제목', '서명', '제목']
    );
    isbnColumn = HR_findAnyHeaderColumn_(
      headers,
      ['ISBN13', 'ISBN 13', 'ISBN', '국제표준도서번호']
    );
    statusColumn = HR_findAnyHeaderColumn_(
      headers,
      ['처리상태', '처리 상태', '상태']
    );

    if (!titleColumn || !statusColumn) {
      throw new Error(
        '희망도서 목록에서 도서명 또는 처리상태 열을 찾지 못했습니다.'
      );
    }

    currentTitle = String(
      sheet
        .getRange(rowNumber, titleColumn)
        .getDisplayValue() || ''
    ).trim();
    currentIsbn = isbnColumn
      ? HR_normalizeIsbn_(
        sheet
          .getRange(rowNumber, isbnColumn)
          .getDisplayValue()
      )
      : '';
    currentStatus = String(
      sheet
        .getRange(rowNumber, statusColumn)
        .getDisplayValue() || ''
    ).trim();

    if (
      !currentTitle ||
      HR_normalizeBookTitle_(currentTitle) !==
        HR_normalizeBookTitle_(requestedTitle) ||
      (
        requestedIsbn &&
        currentIsbn !== requestedIsbn
      )
    ) {
      throw new Error(
        '신청 목록이 변경되었습니다. 새로고침한 뒤 다시 시도해 주세요.'
      );
    }

    var cancelAudienceColumn = HR_findAudienceColumn_(headers);
    var cancelRowAudience = cancelAudienceColumn
      ? HR_audienceKey_(
        sheet.getRange(rowNumber, cancelAudienceColumn).getDisplayValue()
      )
      : 'staff';
    if (cancelRowAudience !== HR_audienceKey_(requestData && requestData.audience)) {
      throw new Error(
        '다른 구분(교직원/학부모)의 신청은 취소할 수 없습니다.'
      );
    }

    if (/구입|입수|선정|반려/.test(currentStatus)) {
      return {
        success: false,
        locked: true,
        title: currentTitle,
        status: currentStatus,
        message:
          '이미 처리 중이거나 완료된 도서라 취소할 수 없습니다.'
      };
    }

    originalColumnCount = sheet.getLastColumn();
    backupSheet = HR_getOrCreateCancelBackup_(
      spreadsheet,
      sheet,
      headers
    );
    backupRow = Math.max(
      backupSheet.getLastRow() + 1,
      2
    );
    cancelId = Utilities.getUuid();

    if (backupSheet.getMaxRows() < backupRow) {
      backupSheet.insertRowsAfter(
        backupSheet.getMaxRows(),
        backupRow - backupSheet.getMaxRows()
      );
    }

    sheet
      .getRange(
        rowNumber,
        1,
        1,
        originalColumnCount
      )
      .copyTo(
        backupSheet.getRange(
          backupRow,
          9,
          1,
          originalColumnCount
        ),
        SpreadsheetApp.CopyPasteType.PASTE_NORMAL,
        false
      );

    backupSheet
      .getRange(backupRow, 1, 1, 8)
      .setValues([[
        cancelId,
        new Date(),
        currentTitle,
        currentIsbn,
        rowNumber,
        originalColumnCount,
        '취소 백업 완료',
        ''
      ]]);

    try {
      sheet.deleteRow(rowNumber);
    } catch (deleteError) {
      backupSheet
        .getRange(backupRow, 7)
        .setValue('삭제 실패 · 원본 유지');
      throw deleteError;
    }

    backupSheet
      .getRange(backupRow, 7)
      .setValue('취소 완료');

    SpreadsheetApp.flush();

    return {
      success: true,
      deleted: true,
      cancelId: cancelId,
      title: currentTitle,
      isbn: currentIsbn,
      status: '신청 취소'
    };
  } finally {
    if (lockAcquired) {
      lock.releaseLock();
    }
  }
}

/**
 * 방금 취소한 신청 한 건을 숨김 백업에서 원래 목록으로 복구합니다.
 */
function HR_undoCancelledRequest(requestData) {
  var cancelId = String(
    requestData && requestData.cancelId || ''
  ).trim();
  var spreadsheet = HR_getSpreadsheet_();
  var sheet = spreadsheet.getSheetByName(
    HR_TARGET_SHEET_NAME
  );
  var backupSheet = spreadsheet.getSheetByName(
    HR_REQUEST_CANCEL_BACKUP_SHEET_NAME
  );
  var lock = LockService.getScriptLock();
  var lockAcquired = false;
  var cancelCell;
  var backupRow;
  var originalRow;
  var originalColumnCount;
  var backupStatus;
  var title;
  var isbn;
  var headers;
  var titleColumn;
  var isbnColumn;
  var currentRows;
  var duplicateRow = 0;
  var rowIndex;
  var targetRow;
  var lastRow;

  if (!cancelId) {
    throw new Error('되돌릴 신청 정보가 없습니다.');
  }

  if (!sheet || !backupSheet) {
    throw new Error(
      '신청 취소 백업을 찾지 못했습니다.'
    );
  }

  try {
    lockAcquired = lock.tryLock(
      HR_REQUEST_LOCK_WAIT_MS
    );

    if (!lockAcquired) {
      throw new Error(
        '현재 이용자가 많습니다. 잠시 후 다시 시도해 주세요.'
      );
    }

    if (backupSheet.getLastRow() < 2) {
      throw new Error(
        '되돌릴 신청 정보를 찾지 못했습니다.'
      );
    }

    cancelCell = backupSheet
      .getRange(
        2,
        1,
        backupSheet.getLastRow() - 1,
        1
      )
      .createTextFinder(cancelId)
      .matchEntireCell(true)
      .findNext();

    if (!cancelCell) {
      throw new Error(
        '되돌릴 신청 정보를 찾지 못했습니다.'
      );
    }

    backupRow = cancelCell.getRow();
    originalRow = Number(
      backupSheet.getRange(backupRow, 5).getValue()
    );
    originalColumnCount = Number(
      backupSheet.getRange(backupRow, 6).getValue()
    );
    backupStatus = String(
      backupSheet.getRange(backupRow, 7).getValue() || ''
    );
    title = String(
      backupSheet.getRange(backupRow, 3).getValue() || ''
    ).trim();
    isbn = HR_normalizeIsbn_(
      backupSheet.getRange(backupRow, 4).getValue()
    );

    if (/^복구 완료|^복구 생략/.test(backupStatus)) {
      return {
        success: true,
        alreadyRestored: true,
        title: title,
        isbn: isbn,
        status: backupStatus
      };
    }

    if (!originalColumnCount) {
      throw new Error(
        '백업된 신청 정보가 올바르지 않습니다.'
      );
    }

    headers = getHeaders(sheet);
    var undoAudienceColumn = HR_findAudienceColumn_(headers);
    var undoRowAudience = (undoAudienceColumn && originalColumnCount >= undoAudienceColumn)
      ? HR_audienceKey_(
        backupSheet.getRange(backupRow, 8 + undoAudienceColumn).getDisplayValue()
      )
      : 'staff';
    if (undoRowAudience !== HR_audienceKey_(requestData && requestData.audience)) {
      throw new Error(
        '다른 구분(교직원/학부모)의 신청은 되돌릴 수 없습니다.'
      );
    }
    titleColumn = HR_findAnyHeaderColumn_(
      headers,
      ['도서명', '책 제목', '책제목', '서명', '제목']
    );
    isbnColumn = HR_findAnyHeaderColumn_(
      headers,
      ['ISBN13', 'ISBN 13', 'ISBN', '국제표준도서번호']
    );

    if (!titleColumn) {
      throw new Error(
        '희망도서 목록에서 도서명 열을 찾지 못했습니다.'
      );
    }

    lastRow = sheet.getLastRow();

    if (lastRow >= FIRST_DATA_ROW) {
      currentRows = sheet
        .getRange(
          FIRST_DATA_ROW,
          1,
          lastRow - FIRST_DATA_ROW + 1,
          sheet.getLastColumn()
        )
        .getDisplayValues();

      for (
        rowIndex = 0;
        rowIndex < currentRows.length;
        rowIndex++
      ) {
        var existingTitle = String(
          currentRows[rowIndex][titleColumn - 1] || ''
        ).trim();
        var existingIsbn = isbnColumn
          ? HR_normalizeIsbn_(
            currentRows[rowIndex][isbnColumn - 1]
          )
          : '';

        if (
          (
            isbn &&
            existingIsbn === isbn
          ) ||
          (
            !isbn &&
            HR_normalizeBookTitle_(existingTitle) ===
              HR_normalizeBookTitle_(title)
          )
        ) {
          duplicateRow = rowIndex + FIRST_DATA_ROW;
          break;
        }
      }
    }

    if (duplicateRow) {
      backupSheet
        .getRange(backupRow, 7, 1, 2)
        .setValues([[
          '복구 생략 · 이미 목록에 있음',
          new Date()
        ]]);

      return {
        success: true,
        alreadyRestored: true,
        title: title,
        isbn: isbn,
        row: duplicateRow,
        status: '이미 목록에 있음'
      };
    }

    if (sheet.getMaxColumns() < originalColumnCount) {
      sheet.insertColumnsAfter(
        sheet.getMaxColumns(),
        originalColumnCount - sheet.getMaxColumns()
      );
    }

    lastRow = sheet.getLastRow();
    targetRow = Math.max(
      FIRST_DATA_ROW,
      Math.min(
        originalRow || lastRow + 1,
        lastRow + 1
      )
    );

    if (targetRow <= lastRow) {
      sheet.insertRowBefore(targetRow);
    } else {
      sheet.insertRowAfter(Math.max(lastRow, HEADER_ROW));
      targetRow = Math.max(lastRow + 1, FIRST_DATA_ROW);
    }

    backupSheet
      .getRange(
        backupRow,
        9,
        1,
        originalColumnCount
      )
      .copyTo(
        sheet.getRange(
          targetRow,
          1,
          1,
          originalColumnCount
        ),
        SpreadsheetApp.CopyPasteType.PASTE_NORMAL,
        false
      );

    backupSheet
      .getRange(backupRow, 7, 1, 2)
      .setValues([[
        '복구 완료',
        new Date()
      ]]);

    SpreadsheetApp.flush();

    return {
      success: true,
      restored: true,
      title: title,
      isbn: isbn,
      row: targetRow,
      status: '신청 접수'
    };
  } finally {
    if (lockAcquired) {
      lock.releaseLock();
    }
  }
}

var HR_BESTSELLER_CACHE_VERSION = 'yes24_v1';
var HR_BESTSELLER_CACHE_MS = 24 * 60 * 60 * 1000;
var HR_BESTSELLER_MAX_RESULTS = 30;
var HR_BESTSELLER_CATEGORIES = {
  all: {
    label: '종합',
    categoryId: ''
  },
  youth: {
    label: '청소년',
    categoryId: ''
  },
  children: {
    label: '어린이',
    categoryId: ''
  },
  picture: {
    label: '그림책',
    categoryId: ''
  }
};

/**
 * 공개 신청 화면에서 세 분류의 베스트셀러를 한 번에 가져옵니다.
 * 같은 날에는 저장해 둔 결과를 재사용하므로 화면을 여러 명이 열어도
 * YES24 API를 반복 호출하지 않습니다.
 */
function HR_getBestsellers() {
  var yes24Key = HR_getYes24Key_();
  var properties = PropertiesService.getScriptProperties();
  var lock = LockService.getScriptLock();
  var categoryKeys = ['all', 'youth', 'children', 'picture'];
  var result = {
    source: 'YES24 인터넷서점',
    cacheHours: 24,
    categories: {}
  };
  var lockAcquired = false;

  if (!yes24Key) {
    throw new Error('YES24 API 키가 설정되지 않았습니다.');
  }

  try {
    lockAcquired = lock.tryLock(15000);

    categoryKeys.forEach(function(categoryKey) {
      result.categories[categoryKey] =
        HR_getBestsellerCategory_(
          categoryKey,
          yes24Key,
          properties,
          lockAcquired
        );
    });
  } finally {
    if (lockAcquired) {
      lock.releaseLock();
    }
  }

  return result;
}

/**
 * 분류 한 개의 저장 결과를 확인하고, 24시간이 지났을 때만 새로 조회합니다.
 */
function HR_getBestsellerCategory_(
  categoryKey,
  yes24Key,
  properties,
  allowRefresh
) {
  var category = HR_BESTSELLER_CATEGORIES[categoryKey];
  var cacheKey =
    'HR_BESTSELLER_' +
    HR_BESTSELLER_CACHE_VERSION +
    '_' +
    categoryKey;
  var timeKey = cacheKey + '_AT';
  var cachedAt = Number(
    properties.getProperty(timeKey) || 0
  );
  var cachedBooks = HR_readBestsellerCache_(
    properties,
    cacheKey
  );
  var isFresh =
    cachedBooks.length > 0 &&
    cachedAt > 0 &&
    Date.now() - cachedAt < HR_BESTSELLER_CACHE_MS;

  if (!category) {
    return {
      label: categoryKey,
      books: [],
      updatedAt: '',
      stale: false
    };
  }

  if (isFresh || !allowRefresh) {
    return {
      label: category.label,
      books: cachedBooks,
      updatedAt: cachedAt
        ? new Date(cachedAt).toISOString()
        : '',
      stale: !isFresh && cachedBooks.length > 0
    };
  }

  try {
    cachedBooks = HR_fetchBestsellerCategory_(
      yes24Key,
      category
    );
    cachedAt = Date.now();

    HR_writeBestsellerCache_(
      properties,
      cacheKey,
      cachedBooks
    );
    properties.setProperty(
      timeKey,
      String(cachedAt)
    );

    return {
      label: category.label,
      books: cachedBooks,
      updatedAt: new Date(cachedAt).toISOString(),
      stale: false
    };
  } catch (error) {
    if (cachedBooks.length > 0) {
      return {
        label: category.label,
        books: cachedBooks,
        updatedAt: cachedAt
          ? new Date(cachedAt).toISOString()
          : '',
        stale: true
      };
    }

    throw error;
  }
}

/**
 * YES24 베스트셀러 API에서 한 분류의 베스트셀러를 조회합니다.
 */
function HR_fetchBestsellerCategory_(yes24Key, category) {
  var categoryId = yes24CategoryId_(category.label, yes24Key);
  var data = yes24Get_('/category/bestseller', {categoryId: categoryId, page: 1, pageSize: HR_BESTSELLER_MAX_RESULTS, detail: 'N'}, yes24Key);
  return data.items.slice(0, HR_BESTSELLER_MAX_RESULTS).map(function(item,index) {
    var book = yes24RequestBook_(item); book.rank = index + 1; return book;
  });
}

/**
 * 베스트셀러 30권을 ScriptProperties 용량 제한에 맞춰 나누어 읽습니다.
 */
function HR_readBestsellerCache_(properties, cacheKey) {
  var count = Number(
    properties.getProperty(cacheKey + '_COUNT') || 0
  );
  var text = '';
  var index;

  if (count > 0) {
    for (index = 0; index < count; index++) {
      text += String(
        properties.getProperty(
          cacheKey + '_PART_' + index
        ) || ''
      );
    }
  } else {
    text = String(
      properties.getProperty(cacheKey) || ''
    );
  }

  return HR_parseBestsellerCache_(text);
}

/**
 * 베스트셀러 결과를 안전한 크기의 여러 속성으로 나누어 저장합니다.
 */
function HR_writeBestsellerCache_(
  properties,
  cacheKey,
  books
) {
  var text = JSON.stringify(books || []);
  var chunkSize = 7000;
  var oldCount = Number(
    properties.getProperty(cacheKey + '_COUNT') || 0
  );
  var count = Math.max(
    1,
    Math.ceil(text.length / chunkSize)
  );
  var values = {};
  var index;

  for (index = 0; index < count; index++) {
    values[cacheKey + '_PART_' + index] =
      text.substring(
        index * chunkSize,
        (index + 1) * chunkSize
      );
  }

  values[cacheKey + '_COUNT'] = String(count);
  properties.setProperties(values, false);
  properties.deleteProperty(cacheKey);

  for (index = count; index < oldCount; index++) {
    properties.deleteProperty(
      cacheKey + '_PART_' + index
    );
  }
}

/**
 * 저장된 베스트셀러 JSON을 안전하게 읽습니다.
 */
function HR_parseBestsellerCache_(text) {
  var value;

  if (!text) {
    return [];
  }

  try {
    value = JSON.parse(text);
  } catch (error) {
    return [];
  }

  return Array.isArray(value)
    ? value
    : [];
}


/**
 * 이미 접수된 도서 중 비어 있는 정가·출판연도만 YES24 정보로 보완합니다.
 * 한 번에 최대 20권까지만 조회해 API 사용량을 제한합니다.
 */
function HR_fillMissingPriceAndYear() {
  var spreadsheet = HR_getSpreadsheet_();
  var sheet = spreadsheet.getSheetByName(HR_TARGET_SHEET_NAME);
  var headers;
  var columns;
  var priceColumn;
  var yearColumn;
  var lastRow;
  var row;
  var checkedCount = 0;
  var updatedCount = 0;
  var failedCount = 0;

  if (!sheet) {
    throw new Error(HR_TARGET_SHEET_NAME + ' 시트를 찾지 못했습니다.');
  }

  headers = getHeaders(sheet);
  columns = getColumnMap(headers);
  priceColumn = HR_findAnyHeaderColumn_(
    headers,
    ['정가', '도서정가', '정가(원)', '가격']
  );
  yearColumn = HR_findAnyHeaderColumn_(
    headers,
    ['연도', '출판연도', '발행연도', '출간연도']
  );
  lastRow = sheet.getLastRow();

  if (!priceColumn || !yearColumn) {
    throw new Error('정가 또는 연도 열을 찾지 못했습니다.');
  }

  for (row = 2; row <= lastRow; row++) {
    var priceValue = Number(
      String(sheet.getRange(row, priceColumn).getDisplayValue() || '')
        .replace(/[^0-9.-]/g, '')
    );
    var yearValue = String(
      sheet.getRange(row, yearColumn).getDisplayValue() || ''
    ).trim();
    var isbnValue = columns.isbn
      ? String(sheet.getRange(row, columns.isbn).getDisplayValue() || '').trim()
      : '';
    var titleValue = columns.title
      ? String(sheet.getRange(row, columns.title).getDisplayValue() || '').trim()
      : '';
    var foundBooks;
    var foundBook;
    var rowUpdated = false;

    if ((priceValue > 0 && yearValue) || (!isbnValue && !titleValue)) {
      continue;
    }

    if (checkedCount >= 20) {
      break;
    }

    checkedCount++;

    try {
      foundBooks = HR_searchRequestBooks(isbnValue || titleValue);
      foundBook = foundBooks && foundBooks.length
        ? foundBooks[0]
        : null;

      if (!foundBook) {
        failedCount++;
        continue;
      }

      if (!(priceValue > 0) && Number(foundBook.price || 0) > 0) {
        sheet
          .getRange(row, priceColumn)
          .setValue(Number(foundBook.price))
          .setNumberFormat('#,##0');
        rowUpdated = true;
      }

      if (!yearValue && foundBook.year) {
        sheet
          .getRange(row, yearColumn)
          .setValue(String(foundBook.year))
          .setNumberFormat('0');
        rowUpdated = true;
      }

      if (rowUpdated) {
        updatedCount++;
      }
    } catch (error) {
      failedCount++;
      console.error('정가·연도 보완 오류 (' + row + '행)', error);
    }

    Utilities.sleep(120);
  }

  SpreadsheetApp.flush();

  return {
    checkedCount: checkedCount,
    updatedCount: updatedCount,
    failedCount: failedCount
  };
}


function HR_breakMergedRangesSafely_(sheet, rangeA1) {
  var targetRange = rangeA1
    ? sheet.getRange(rangeA1)
    : sheet.getRange(
        1,
        1,
        Math.max(sheet.getMaxRows(), 1),
        Math.max(sheet.getMaxColumns(), 1)
      );

  targetRange.getMergedRanges().forEach(function(mergedRange) {
    mergedRange.breakApart();
  });
}



/**
 * 웹페이지에서 해석한 엑셀·CSV 도서 목록을 업무용 시트에 저장합니다.
 * 원본 파일 전체는 전송하지 않고 필요한 서지정보만 받습니다.
 */
function HR_importBookListFile(payload) {
  payload = payload || {};

  var books = Array.isArray(payload.books) ? payload.books : [];
  var fileName = String(payload.fileName || '업로드 파일').trim();

  if (!books.length) {
    throw new Error('파일에서 읽은 도서 정보가 없습니다.');
  }
  if (books.length > 2000) {
    throw new Error('한 번에 최대 1,500권까지 업로드할 수 있습니다.');
  }

  return unifiedReplaceUploadRows_(books, fileName);
}


/**
 * ============================================================
 * 공개 신청 입력 정리·검증
 * - 브라우저에서 온 값은 신뢰하지 않습니다.
 * - 문자열은 정리해 길이를 제한하고, 시트에는 텍스트로만 기록합니다.
 * - 도서 정보는 ISBN으로 YES24에서 다시 확인한 값을 우선합니다.
 * ============================================================
 */
function HR_cleanInput_(value, maxLength) {
  var source = String(value === null || value === undefined ? '' : value);
  var text = '';
  var index = 0;
  var code = 0;

  // 제어문자·보이지 않는 문자는 공백으로 바꿉니다.
  for (index = 0; index < source.length; index++) {
    code = source.charCodeAt(index);
    if (
      code < 32 || code === 127 ||
      (code >= 8203 && code <= 8207) ||
      code === 8232 || code === 8233 || code === 65279
    ) {
      text += ' ';
    } else {
      text += source.charAt(index);
    }
  }

  text = text.replace(/\s+/g, ' ').trim();

  if (maxLength && text.length > maxLength) {
    text = text.slice(0, maxLength);
  }
  return text;
}

/**
 * 텍스트 서식을 먼저 지정한 뒤 값을 넣어, 입력이 수식 문자로 시작해도 실행되지 않게 합니다.
 */
function HR_setSafeText_(cell, value) {
  var text = HR_cleanInput_(value, 500);

  cell.setNumberFormat('@');
  if (!text) {
    cell.setValue('');
    return cell;
  }
  // setValue 는 '=' 로 시작하는 문자열을 수식으로 실행하므로, '=' 로 시작하면
  // 항상 텍스트로 저장되는 서식 있는 텍스트를 사용합니다. 그 밖의 문자열은 텍스트 서식만으로 안전합니다.
  if (text.charAt(0) === '=') {
    cell.setRichTextValue(
      SpreadsheetApp.newRichTextValue().setText(text).build()
    );
  } else {
    cell.setValue(text);
  }
  return cell;
}

function HR_cleanPrice_(value) {
  var number = Number(String(value === null || value === undefined ? '' : value).replace(/[^0-9.]/g, ''));
  return isFinite(number) && number > 0 && number <= 500000
    ? Math.round(number)
    : 0;
}

function HR_cleanCallNumber_(value) {
  var text = HR_cleanInput_(value, 20);
  return /^[0-9]{1,3}(?:\.[0-9]{1,8})?$/.test(text) ? text : '';
}

/**
 * 표지는 YES24 주소만 허용합니다. 수식 안에 들어가므로 따옴표·공백·괄호가 있으면 버립니다.
 */
function HR_safeCoverUrl_(value) {
  var text = String(value || '').trim();

  if (text.length > 300) return '';
  if (/["\\\s()]/.test(text)) return '';
  return /^https?:\/\/(?:[a-z0-9-]+\.)*yes24\.com\//i.test(text)
    ? text
    : '';
}

function HR_prepareSelectedBook_(clientBook, query) {
  var source = clientBook && typeof clientBook === 'object' ? clientBook : {};
  var year = (
    HR_cleanInput_(source.year || source.pubDate || source.publicationDate, 20)
      .match(/\d{4}/) || ['']
  )[0];
  var book = {
    title: HR_cleanInput_(source.title, 200),
    author: HR_cleanInput_(source.author, 150),
    publisher: HR_cleanInput_(source.publisher, 100),
    isbn: yes24Isbn13_(source.isbn) || yes24Isbn13_(query) || '',
    cover: '',
    link: '',
    price: HR_cleanPrice_(
      source.price || source.priceStandard || source.priceStd || source.priceSales
    ),
    year: year && Number(year) >= 1800 && Number(year) <= 2100 ? year : '',
    category: HR_cleanInput_(source.category, 100),
    callNumber: HR_cleanCallNumber_(source.callNumber),
    verified: false
  };
  var items;
  var server;

  if (book.isbn) {
    try {
      items = yes24SearchItems_(book.isbn, HR_getYes24Key_(), 'BOOK', 1, true);

      if (items && items.length) {
        server = yes24RequestBook_(items[0]);
        book.title = HR_cleanInput_(server.title, 200) || book.title;
        book.author = HR_cleanInput_(server.author, 150) || book.author;
        book.publisher = HR_cleanInput_(server.publisher, 100) || book.publisher;
        book.isbn = yes24Isbn13_(server.isbn) || book.isbn;
        book.price = HR_cleanPrice_(server.price) || book.price;
        book.year = server.year || book.year;
        book.category = HR_cleanInput_(server.category, 100) || book.category;
        book.cover = HR_safeCoverUrl_(server.cover);
        book.link = HR_safeUrl_(server.link);
        book.verified = true;
      }
    } catch (error) {
      // YES24 연결 문제로 확인하지 못한 경우에는 정리한 입력값으로 접수하고 '미확인'으로 표시합니다.
      console.error('신청 도서 서버 확인 실패', error);
    }
  }

  return book;
}

function HR_unverifiedNote_(book) {
  return book && book.verified ? '' : ' · 도서정보 미확인';
}


/* ===== ReservationShortcut.js ===== */
/** 시트에 등록된 공개 희망도서 신청 웹앱 바로가기 */
function HR_showHopeRequestDialog() {
  const requestUrl = String(
    PropertiesService.getScriptProperties().getProperty(HR_PUBLIC_WEB_APP_URL_PROPERTY) || ''
  ).trim();

  if (!requestUrl) {
    SpreadsheetApp.getUi().alert(
      '희망도서 신청 링크 없음',
      '먼저 ⚙️ 사본 설정에서 공개 신청 링크를 등록해 주세요.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }

  const html = HtmlService.createHtmlOutput(
    '<!doctype html><html lang="ko"><head><meta charset="utf-8"><style>' +
    'body{margin:0;padding:28px;background:#f8f5ed;font-family:Arial,"Noto Sans KR",sans-serif;color:#183d35;text-align:center}' +
    '.icon{font-size:48px;margin:4px 0 12px}.title{font-size:28px;font-weight:900;margin:0 0 12px}.desc{font-size:16px;color:#5b6f65;line-height:1.6;margin:0 0 26px}' +
    '.open{display:block;padding:17px 20px;border-radius:16px;background:linear-gradient(135deg,#d9a441,#b9852a);color:#fff;text-decoration:none;font-size:21px;font-weight:900;box-shadow:0 10px 24px rgba(33,105,209,.25)}' +
    '.open:hover{filter:brightness(.97);transform:translateY(-1px)}' +
    '</style></head><body><div class="icon">📚</div><div class="title">희망도서 신청 바로가기</div>' +
    '<p class="desc">아래 버튼을 누르면 희망도서 신청 화면이 새 탭으로 열립니다.</p>' +
    '<a class="open" href="' + requestUrl + '" target="_blank" rel="noopener" onclick="setTimeout(function(){google.script.host.close();},300)">희망도서 신청 열기</a>' +
    '</body></html>'
  ).setWidth(520).setHeight(330);

  SpreadsheetApp.getUi().showModalDialog(html, '희망도서 신청 바로가기');
}


/* ===== UnifiedLibrary.js ===== */
/**
 * 보유목록 시리즈 누락·후속권 점검
 * 기존 도서 관리 기능과 독립적으로 동작합니다.
 */

var SERIES_FINDER_SHEET = '🔎 시리즈 누락 점검';
var SERIES_HOLDINGS_SHEET = '🏫 우리학교 소장도서';
var SERIES_WORK_SHEET = '📋 시리즈 업무목록';
var SERIES_MAX_GROUPS = 3000;
var SERIES_MAX_API_REQUESTS = 5000;
var SERIES_DAILY_API_SAFE_LIMIT = 18000;
var SERIES_SEARCH_CHUNK = 5;
var SERIES_MAX_ENRICH_REQUESTS = 20;
var SERIES_EXECUTION_BUDGET_MS = 285000;
var SERIES_FOLLOWUP_RESULTS = 50;
var SERIES_GAP_RESULTS = 50;
var SERIES_MAX_GAP_PAGES = 3;
var SERIES_RESUME_STATE_KEY = 'YES24_SERIES_RESUME_STATE_V1';
var SERIES_AUTO_CONTINUE_HANDLER = 'seriesAutoContinueMissingBooks_';
var SERIES_AUTO_CONTINUE_DELAY_MS = 60000;

function setupSeriesFinder() {
  seriesCancelAutoContinue_();
  seriesClearResumeState_();
  seriesInstallOpenTrigger_();
  seriesInstallWorkButtonTrigger_();
  seriesBuildResultSheet_([], {
    seriesCount: 0,
    searchedCount: 0,
    checkedAt: new Date()
  });
  seriesSpreadsheet_().toast(
    '상단의 「🔎 시리즈 점검」 메뉴에서 후속권·누락권 찾기를 눌러 주세요.',
    '설정 완료',
    8
  );
}

function seriesFinderOnOpen_() {
  seriesAddMenu_();
}

// 시리즈 점검 메뉴는 「🧰 통합 도서업무」 안의 하위 메뉴로 합쳤습니다.
// (예전 설치형 열기 트리거가 이 함수를 불러도 메뉴 줄에 따로 만들지 않습니다.)
function seriesAddMenu_() {
}

function seriesInstallOpenTrigger_() {
  var handler = 'seriesFinderOnOpen_';
  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    if (trigger.getHandlerFunction() === handler) {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  ScriptApp.newTrigger(handler)
    .forSpreadsheet(seriesSpreadsheet_())
    .onOpen()
    .create();
}


function installSeriesWorkButton() {
  seriesInstallWorkButtonTrigger_();

  var sheet = seriesSpreadsheet_().getSheetByName(SERIES_FINDER_SHEET);
  if (!sheet) {
    throw new Error('「' + SERIES_FINDER_SHEET + '」 시트를 찾을 수 없습니다.');
  }

  seriesAddWorkButton_(sheet);
  seriesSpreadsheet_().toast(
    '책을 체크한 뒤 결과 시트 위쪽의 금색 버튼을 누르면 업무목록으로 이동합니다.',
    '업무목록 버튼 설치 완료',
    8
  );
}

function seriesInstallWorkButtonTrigger_() {
  var handler = 'seriesWorkButtonOnEdit_';
  var alreadyInstalled = ScriptApp.getProjectTriggers().some(function(trigger) {
    return trigger.getHandlerFunction() === handler;
  });

  if (!alreadyInstalled) {
    ScriptApp.newTrigger(handler)
      .forSpreadsheet(seriesSpreadsheet_())
      .onEdit()
      .create();
  }
}

function seriesWorkButtonOnEdit_(e) {
    // 상단 체크박스는 Code.gs의 onEdit에서 한 번만 처리합니다.
  return;

  if (!e || !e.range) return;

  var range = e.range;
  var sheet = range.getSheet();
  if (
    sheet.getName() !== SERIES_FINDER_SHEET ||
    range.getA1Notation() !== 'A4' ||
    String(e.value || '').toUpperCase() !== 'TRUE'
  ) {
    return;
  }

  var lock = LockService.getDocumentLock();
  if (!lock.tryLock(3000)) {
    range.setValue(false);
    seriesSpreadsheet_().toast('다른 작업을 처리 중입니다. 잠시 후 다시 눌러 주세요.','잠시만요',5);
    return;
  }

  try {
    sendSelectedSeriesToReview();
  } catch (error) {
    seriesSpreadsheet_().toast(
      String(error && error.message ? error.message : error),
      '구입 검토함 전송 오류',
      8
    );
  } finally {
    range.setValue(false);
    lock.releaseLock();
  }
}

function seriesAddWorkButton_(sheet) {
  sheet.getRange('A4:N4').breakApart();
  sheet.getRange('A4:N4').clearContent().clearDataValidations();

  sheet.getRange('A4')
    .insertCheckboxes()
    .setValue(false)
    .setBackground('#2F5D50')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  sheet.getRange('B4:C4')
    .merge()
    .setValue('전체 선택/해제')
    .setBackground('#2F5D50')
    .setFontColor('#FFFFFF')
    .setFontSize(13)
    .setFontWeight('bold')
    .setHorizontalAlignment('left')
    .setVerticalAlignment('middle');

  sheet.getRange('D4')
    .insertCheckboxes()
    .setValue(false)
    .setBackground('#D9A441')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  sheet.getRange('E4:N4')
    .merge()
    .setValue('선택한 책을 구입 검토함으로 보내기')
    .setBackground('#D9A441')
    .setFontColor('#FFFFFF')
    .setFontSize(14)
    .setFontWeight('bold')
    .setHorizontalAlignment('left')
    .setVerticalAlignment('middle');

  sheet.setRowHeight(4, 44);
  sheet.getRange('A4:C4').setBorder(
    true,true,true,true,false,false,
    '#244A40',SpreadsheetApp.BorderStyle.SOLID_MEDIUM
  );
  sheet.getRange('D4:N4').setBorder(
    true,true,true,true,false,false,
    '#B88428',SpreadsheetApp.BorderStyle.SOLID_MEDIUM
  );
}

function setSeriesYes24Key() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt(
    'YES24 API 키 설정',
    'YES24 API 키를 입력하세요. 기존 프로젝트 키가 있다면 입력하지 않아도 됩니다.',
    ui.ButtonSet.OK_CANCEL
  );
  if (response.getSelectedButton() !== ui.Button.OK) return;

  var key = String(response.getResponseText() || '').trim();
  if (!key) {
    ui.alert('API 키를 입력해 주세요.');
    return;
  }
  HR_testYes24Key_(key);
  PropertiesService.getScriptProperties().setProperty('HR_YES24_API_KEY', key);
  ui.alert('YES24 API 키를 저장했습니다.');
}

function refreshMissingSeriesBooks() {
  var runStartedAt = Date.now();
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) {
    seriesSpreadsheet_().toast(
      '다른 시리즈 점검이 진행 중입니다. 잠시 후 다시 시도해 주세요.',
      '시리즈 점검',
      7
    );
    return;
  }

  try {
    var spreadsheet = seriesSpreadsheet_();
    var key = seriesGetYes24Key_();
    if (!key) {
      throw new Error('YES24 API 키가 없습니다. 「🔎 시리즈 점검 → YES24 API 키 설정」에서 입력해 주세요.');
    }

    spreadsheet.toast('보유목록에서 시리즈와 빈 권차를 찾고 있습니다.', '시리즈 점검', 5);
    var holdings = seriesReadHoldings_();
    var allGroups = seriesGroupHoldings_(holdings.rows).sort(seriesGroupPriority_);
    var groups = allGroups.slice(0, SERIES_MAX_GROUPS);
    var groupLimitReached = allGroups.length > groups.length;
    var fingerprint = seriesResumeFingerprint_(groups);
    var properties = PropertiesService.getScriptProperties();
    var hadSavedResume = Boolean(properties.getProperty(SERIES_RESUME_STATE_KEY));
    var resumeState = seriesLoadResumeState_(fingerprint);
    var resumeOffset = 0;
    var continuing = false;

    if (resumeState && Number(resumeState.offset || 0) > 0) {
      resumeOffset = Number(resumeState.offset || 0);
      continuing = true;
    } else if (!hadSavedResume) {
      resumeOffset = 0;
      continuing = resumeOffset > 0;
    }

    var existingCandidates = continuing ? seriesReadExistingCandidates_() : [];
    var enrichment = seriesEnrichGroupsFromIsbn_(groups, key);

    if (!groups.length) {
      seriesClearResumeState_();
      seriesBuildResultSheet_([], {
        seriesCount: 0,
        searchedCount: 0,
        requestCount: enrichment.requestCount,
        truncated: false,
        progressDone: 0,
        progressTotal: 0,
        checkedAt: new Date()
      });
      spreadsheet.toast('권차가 표시된 시리즈를 찾지 못했습니다.', '점검 완료', 7);
      return;
    }

    spreadsheet.toast(
      (continuing ? '이전 점검 다음부터 ' : '') +
        groups.length + '개 시리즈의 중간 누락과 후속권을 확인합니다.',
      '시리즈 점검',
      8
    );

    var searchResult = seriesSearchYes24_(groups, holdings, key, resumeOffset, runStartedAt);
    var candidates = seriesMergeCandidates_(existingCandidates, searchResult.candidates);
    var incomplete = groupLimitReached || searchResult.truncated;

    if (incomplete) {
      seriesSaveResumeState_({
        fingerprint: fingerprint,
        offset: searchResult.nextOffset,
        total: searchResult.totalSpecCount,
        updatedAt: new Date().toISOString()
      });
      seriesScheduleAutoContinue_();
    } else {
      seriesClearResumeState_();
      seriesCancelAutoContinue_();
    }

    seriesBuildResultSheet_(candidates, {
      seriesCount: groups.length,
      searchedCount: searchResult.searchedGroupCount,
      requestCount: searchResult.requestCount + enrichment.requestCount,
      truncated: incomplete,
      progressDone: searchResult.nextOffset,
      progressTotal: searchResult.totalSpecCount,
      remainingCount: Math.max(0, searchResult.totalSpecCount - searchResult.nextOffset),
      checkedAt: new Date()
    });

    spreadsheet.setActiveSheet(spreadsheet.getSheetByName(SERIES_FINDER_SHEET));
    var percent = searchResult.totalSpecCount
      ? Math.min(100, searchResult.nextOffset / searchResult.totalSpecCount * 100).toFixed(1)
      : '100.0';
    spreadsheet.toast(
      '구매 검토 후보 ' + candidates.length + '권 · 검색 진행 ' +
        searchResult.nextOffset + ' / ' + searchResult.totalSpecCount + '건 (' + percent + '%)' +
        (incomplete ? ' 약 1분 뒤 남은 검색을 자동으로 이어갑니다.' : ' 전체 확인을 마쳤습니다.'),
      incomplete ? '시리즈 점검 저장 완료' : '시리즈 점검 전체 완료',
      10
    );
  } catch (error) {
    seriesSpreadsheet_().toast(
      String(error && error.message ? error.message : error),
      '시리즈 점검 오류',
      12
    );
    throw error;
  } finally {
    lock.releaseLock();
  }
}

function seriesGetYes24Key_() { return getConfiguredYes24Key_() || String(PropertiesService.getScriptProperties().getProperty('SERIES_YES24_KEY') || '').trim(); }

function seriesSpreadsheet_() {
  var active = SpreadsheetApp.getActive();
  if (active) return active;
  if (typeof HR_getSpreadsheet_ === 'function') return HR_getSpreadsheet_();
  throw new Error('연결된 스프레드시트를 찾을 수 없습니다.');
}

function seriesReadHoldings_() {
  var sheet = seriesSpreadsheet_().getSheetByName(SERIES_HOLDINGS_SHEET);
  if (!sheet) throw new Error('「' + SERIES_HOLDINGS_SHEET + '」 시트를 찾을 수 없습니다.');

  var lastRow = sheet.getLastRow();
  var lastColumn = sheet.getLastColumn();
  if (lastRow < 2 || lastColumn < 1) {
    return {rows: [], isbnSet: {}, titleKeySet: {}, parsedRows: []};
  }

  var headerSearchRows = Math.min(10, lastRow);
  var headerGrid = sheet.getRange(1, 1, headerSearchRows, lastColumn).getDisplayValues();
  var headerRow = 1;
  var titleColumn = 0;
  var isbnColumn = 0;
  var publisherColumn = 0;

  for (var r = 0; r < headerGrid.length; r++) {
    for (var c = 0; c < headerGrid[r].length; c++) {
      var header = String(headerGrid[r][c] || '').replace(/\s+/g, '').toUpperCase();
      if (!titleColumn && /^(도서명|책제목|서명|제목)$/.test(header)) titleColumn = c + 1;
      if (!isbnColumn && /^ISBN(13)?$/.test(header)) isbnColumn = c + 1;
      if (!publisherColumn && /^(출판사|발행자|출판자)$/.test(header)) publisherColumn = c + 1;
    }
    if (titleColumn && isbnColumn) {
      headerRow = r + 1;
      break;
    }
  }

  if (!titleColumn || !isbnColumn) {
    throw new Error('「' + SERIES_HOLDINGS_SHEET + '」 시트에서 도서명과 ISBN13 열을 찾지 못했습니다.');
  }

  var count = lastRow - headerRow;
  if (count < 1) return {rows: [], isbnSet: {}, titleKeySet: {}, parsedRows: []};
  var values = sheet.getRange(headerRow + 1, 1, count, lastColumn).getDisplayValues();
  var rows = [];
  var isbnSet = {};
  var titleKeySet = {};
  var parsedRows = [];

  values.forEach(function(row, index) {
    var title = String(row[titleColumn - 1] || '').trim();
    if (!title) return;
    var isbn = seriesCleanIsbn_(row[isbnColumn - 1]);
    var publisher = publisherColumn ? String(row[publisherColumn - 1] || '').trim() : '';
    if (isbn) isbnSet[isbn] = true;
    var titleKey = seriesNormalizeKey_(title);
    if (titleKey) titleKeySet[titleKey] = true;
    var parsed = seriesParseTitle_(title);
    if (parsed) {
      parsedRows.push({
        key: parsed.key,
        volume: parsed.volume,
        publisherKey: seriesNormalizePublisher_(publisher)
      });
    }
    rows.push({
      row: index + headerRow + 1,
      title: title,
      isbn: isbn,
      publisher: publisher
    });
  });
  return {rows: rows, isbnSet: isbnSet, titleKeySet: titleKeySet, parsedRows: parsedRows};
}

function seriesGroupHoldings_(rows) {
  var groupsByKey = {};

  rows.forEach(function(book) {
    var parsed = seriesParseTitle_(book.title);
    if (!parsed) return;

    var publisherKey = seriesNormalizePublisher_(book.publisher) || '출판사미상';
    var formatKey = seriesFormatKey_(book.title);
    var signature = parsed.key + '|' + publisherKey + '|' + formatKey;

    if (!groupsByKey[signature]) {
      groupsByKey[signature] = {
        key: signature,
        seriesKey: parsed.key,
        formatKey: formatKey,
        searchBase: parsed.base,
        aliasKeys: {},
        coreSearchBases: {},
        sampleIsbn: book.isbn || '',
        base: parsed.base,
        volumes: {},
        publishers: {},
        titles: [],
        explicitCount: 0
      };
    }

    var group = groupsByKey[signature];
    var coreParsed = seriesLeadingCore_(book.title);
    if (coreParsed) {
      group.aliasKeys[coreParsed.key] = true;
      group.coreSearchBases[coreParsed.base] = true;
    }
    group.volumes[parsed.volume] = true;
    group.titles.push(book.title);
    if (!group.sampleIsbn && book.isbn) group.sampleIsbn = book.isbn;
    if (parsed.explicit) group.explicitCount++;
    if (publisherKey !== '출판사미상') group.publishers[publisherKey] = true;
  });

  return Object.keys(groupsByKey)
    .map(function(key) {
      var group = groupsByKey[key];
      var coreSearchBases = Object.keys(group.coreSearchBases || {}).sort(function(a, b) {
        return a.length - b.length;
      });
      if (coreSearchBases.length) {
        group.preferredSearchBase = coreSearchBases[0];
        group.searchBase = group.preferredSearchBase;
      }
      group.aliasKeys[group.seriesKey] = true;
      group.volumeList = Object.keys(group.volumes)
        .map(Number)
        .filter(function(value) { return value > 0 && value <= 200; })
        .sort(function(a, b) { return a - b; });
      group.maxVolume = group.volumeList.length
        ? Math.max.apply(null, group.volumeList)
        : 0;
      group.heldText = seriesCompressNumbers_(group.volumeList);
      group.missingVolumes = [];
      for (var volume = 1; volume < group.maxVolume; volume++) {
        if (!group.volumes[volume]) group.missingVolumes.push(volume);
      }
      return group;
    })
    .filter(function(group) {
      if (group.volumeList.length < 1 || group.base.length < 2) return false;
      if (group.explicitCount === 0 && group.volumeList.length === 1 && group.maxVolume > 80) {
        return false;
      }
      return true;
    });
}

function seriesEnrichGroupsFromIsbn_(groups, apiKey) {
  
  var properties = PropertiesService.getScriptProperties();
  var specs = [];
  var cacheWrites = {};

  groups.forEach(function(group) {
    if (!group.sampleIsbn) return;
    var cacheKey = 'YES24_SERIES_EDITION_CACHE_' + group.sampleIsbn;
    var cached = properties.getProperty(cacheKey);
    if (cached) {
      try {
        seriesApplyEditionInfo_(group, JSON.parse(cached));
        return;
      } catch (ignore) {}
    }
    specs.push({group: group, isbn: group.sampleIsbn, cacheKey: cacheKey});
  });

  // 판본 확인은 검색 정확도를 높이는 보조 단계입니다. 실제 후속권 검색이 먼저 끝나도록 일부만 확인합니다.
  specs = specs.slice(0, SERIES_MAX_ENRICH_REQUESTS);
  var requiredToday = specs.length;
  if (seriesApiRemaining_() < requiredToday) {
    throw new Error(
      '오늘 안전하게 사용할 수 있는 YES24 조회량이 부족합니다. 기존 결과는 그대로 보존했습니다. 내일 다시 점검해 주세요.'
    );
  }

  var requestCount = 0;
  for (var start = 0; start < specs.length; start += SERIES_SEARCH_CHUNK) {
    var chunk = specs.slice(start, start + SERIES_SEARCH_CHUNK);
    var requests = chunk.map(function(spec) {
      return {
        url: seriesBuildLookupUrl_(spec.isbn, apiKey),
        method: 'get',
        headers: {'X-Api-Key': apiKey},
        muteHttpExceptions: true,
        followRedirects: true
      };
    });
    var responses = yes24FetchAll_(requests);
    requestCount += responses.length;
    seriesRecordApiUsage_(responses.length);

    responses.forEach(function(response, index) {
      if (response.getResponseCode() !== 200) return;
      var data;
      try {
        data = JSON.parse(response.getContentText('UTF-8'));
      } catch (ignore) {
        return;
      }
      var item = data && data.data && Array.isArray(data.data.items) ? data.data.items[0] : null;
      if (!item) return;

      var info = {
        title: seriesCleanText_(item.title),
        publisher: seriesCleanText_(item.publisher)
      };
      seriesApplyEditionInfo_(chunk[index].group, info);
      cacheWrites[chunk[index].cacheKey] = JSON.stringify(info);
    });

    if (Object.keys(cacheWrites).length) properties.setProperties(cacheWrites, false);
    
    Utilities.sleep(100);
  }

  if (Object.keys(cacheWrites).length) {
    properties.setProperties(cacheWrites, false);
  }
  return {requestCount: requestCount};
}

function seriesApplyEditionInfo_(group, info) {
  var apiTitle = seriesCleanText_(info && info.title);
  var parsed = seriesParseTitle_(apiTitle);
  if (parsed) {
    group.aliasKeys[parsed.key] = true;
    group.searchBase = group.preferredSearchBase || parsed.base || group.searchBase;
    group.formatKey = seriesFormatKey_(apiTitle);
  }
  var publisherKey = seriesNormalizePublisher_(info && info.publisher);
  if (publisherKey) group.publishers[publisherKey] = true;
}

function seriesApiUsageState_() {
  var properties = PropertiesService.getScriptProperties();
  var zone = Session.getScriptTimeZone() || 'Asia/Seoul';
  var today = Utilities.formatDate(new Date(), zone, 'yyyy-MM-dd');
  var savedDate = properties.getProperty('YES24_SERIES_API_USAGE_DATE') || '';
  var used = Number(properties.getProperty('YES24_SERIES_API_USAGE_COUNT') || 0);

  if (savedDate !== today) {
    used = 0;
    properties.setProperties({
      YES24_SERIES_API_USAGE_DATE: today,
      YES24_SERIES_API_USAGE_COUNT: '0'
    }, false);
  }
  return {date: today, used: used};
}

function seriesApiRemaining_() {
  var state = seriesApiUsageState_();
  return Math.max(0, SERIES_DAILY_API_SAFE_LIMIT - state.used);
}

function seriesRecordApiUsage_(count) {
  var properties = PropertiesService.getScriptProperties();
  var state = seriesApiUsageState_();
  properties.setProperty(
    'YES24_SERIES_API_USAGE_COUNT',
    String(state.used + Math.max(0, Number(count || 0)))
  );
}

function seriesBuildLookupUrl_(isbn, apiKey) { return yes24Url_('/goods/itemDetail', {query: yes24Isbn13_(isbn), searchType: 'ISBN13', detail: 'Y'}); }

function seriesGroupPriority_(a, b) {
  var aMulti = a.volumeList.length >= 2 ? 1 : 0;
  var bMulti = b.volumeList.length >= 2 ? 1 : 0;
  if (aMulti !== bMulti) return bMulti - aMulti;
  if (a.explicitCount !== b.explicitCount) return b.explicitCount - a.explicitCount;
  if (a.maxVolume !== b.maxVolume) return b.maxVolume - a.maxVolume;
  return a.base.localeCompare(b.base, 'ko');
}

function seriesResumeFingerprint_(groups) {
  var source = groups.map(function(group) {
    return group.key + '|' + Object.keys(group.volumes || {})
      .map(Number)
      .sort(function(a, b) { return a - b; })
      .join(',');
  }).sort().join('\n');

  var digest = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    source,
    Utilities.Charset.UTF_8
  );
  return digest.map(function(value) {
    var normalized = (Number(value) + 256) % 256;
    return ('0' + normalized.toString(16)).slice(-2);
  }).join('');
}

function seriesLoadResumeState_(fingerprint) {
  var properties = PropertiesService.getScriptProperties();
  var raw = properties.getProperty(SERIES_RESUME_STATE_KEY);
  if (!raw) return null;

  try {
    var state = JSON.parse(raw);
    if (!state || state.fingerprint !== fingerprint) {
      properties.deleteProperty(SERIES_RESUME_STATE_KEY);
      return null;
    }
    return state;
  } catch (ignore) {
    properties.deleteProperty(SERIES_RESUME_STATE_KEY);
    return null;
  }
}

function seriesSaveResumeState_(state) {
  PropertiesService.getScriptProperties().setProperty(
    SERIES_RESUME_STATE_KEY,
    JSON.stringify(state || {})
  );
}

function seriesClearResumeState_() {
  PropertiesService.getScriptProperties().deleteProperty(SERIES_RESUME_STATE_KEY);
}

function resetSeriesFinderProgress() {
  seriesCancelAutoContinue_();
  seriesClearResumeState_();
  seriesSpreadsheet_().toast(
    '저장된 점검 위치를 지웠습니다. 다음 점검은 처음부터 시작합니다.',
    '시리즈 점검 초기화',
    8
  );
}

function seriesAutoContinueMissingBooks_() {
  refreshMissingSeriesBooks();
}

function seriesCancelAutoContinue_() {
  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    if (trigger.getHandlerFunction() === SERIES_AUTO_CONTINUE_HANDLER) {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}

function seriesScheduleAutoContinue_() {
  seriesCancelAutoContinue_();
  ScriptApp.newTrigger(SERIES_AUTO_CONTINUE_HANDLER)
    .timeBased()
    .after(SERIES_AUTO_CONTINUE_DELAY_MS)
    .create();
}

function seriesLegacyResumeOffset_() {
  var sheet = seriesSpreadsheet_().getSheetByName(SERIES_FINDER_SHEET);
  if (!sheet) return 0;
  var summary = String(sheet.getRange('A3').getDisplayValue() || '');
  return summary.indexOf('일부 미확인') >= 0 ? SERIES_MAX_API_REQUESTS : 0;
}

function seriesReadExistingCandidates_() {
  var sheet = seriesSpreadsheet_().getSheetByName(SERIES_FINDER_SHEET);
  if (!sheet || sheet.getLastRow() < 6) return [];

  var rowCount = sheet.getLastRow() - 5;
  var range = sheet.getRange(6, 1, rowCount, 14);
  var values = range.getValues();
  var formulas = range.getFormulas();
  var timezone = Session.getScriptTimeZone();

  return values.map(function(row, index) {
    var title = seriesCleanText_(row[6]);
    var seriesName = seriesCleanText_(row[2]);
    var volumeMatch = String(row[4] || '').match(/(\d{1,3})/);
    if (!title || !seriesName || !volumeMatch) return null;

    var pubDate = row[9];
    if (pubDate instanceof Date && !isNaN(pubDate.getTime())) {
      pubDate = Utilities.formatDate(pubDate, timezone, 'yyyy-MM-dd');
    } else {
      pubDate = seriesCleanText_(pubDate);
    }

    return {
      selected: row[0] === true,
      type: seriesCleanText_(row[1]),
      series: seriesName,
      held: seriesCleanText_(row[3]),
      volume: Number(volumeMatch[1]),
      cover: seriesExtractUrlFromFormula_(formulas[index][5]),
      title: title,
      author: seriesCleanText_(row[7]),
      publisher: seriesCleanText_(row[8]),
      pubDate: pubDate,
      isbn: seriesCleanIsbn_(row[10]),
      price: Number(row[11] || 0),
      link: seriesExtractUrlFromFormula_(formulas[index][12]),
      saleStatus: /^절판(?:\s*·|$)/.test(String(row[13] || '')) ? '절판' : '',
      publisherMatched: false,
      exactKey: false
    };
  }).filter(function(book) { return Boolean(book); });
}

function seriesExtractUrlFromFormula_(formula) {
  var match = String(formula || '').match(/"(https?:\/\/[^"]+)"/i);
  return match && match[1] ? match[1].replace(/""/g, '"') : '';
}

function seriesMergeCandidates_(existing, incoming) {
  var byKey = {};

  function put(book) {
    if (!book) return;
    var key = book.isbn
      ? 'isbn|' + book.isbn
      : 'series|' + seriesNormalizeKey_(book.series) + '|' +
        Number(book.volume || 0) + '|' + seriesNormalizePublisher_(book.publisher);
    var previous = byKey[key];
    if (!previous) {
      byKey[key] = book;
      return;
    }

    var selected = Boolean(previous.selected || book.selected);
    if (seriesCandidateScore_(book) >= seriesCandidateScore_(previous)) {
      byKey[key] = book;
    }
    byKey[key].selected = selected;
  }

  (existing || []).forEach(put);
  (incoming || []).forEach(put);

  return Object.keys(byKey)
    .map(function(key) { return byKey[key]; })
    .sort(function(a, b) {
      var typeOrder = {'중간 누락': 0, '신간 후속권': 1};
      if (typeOrder[a.type] !== typeOrder[b.type]) return typeOrder[a.type] - typeOrder[b.type];
      var baseOrder = String(a.series || '').localeCompare(String(b.series || ''), 'ko');
      return baseOrder || Number(a.volume || 0) - Number(b.volume || 0);
    });
}

function seriesSearchYes24_(groups, holdings, apiKey, resumeOffset, runStartedAt) {
  var followupSpecs = groups.map(function(group) {
    return {
      group: group,
      kind: 'followup',
      query: group.searchBase || group.base,
      page: 1
    };
  });

  var gapGroups = groups.filter(function(group) {
    return group.missingVolumes && group.missingVolumes.length;
  });
  var gapSpecs = [];
  var round = 0;
  var added = true;

  // 모든 중간 누락 권차를 한 권씩 번갈아 대기열에 넣습니다.
  while (added) {
    added = false;
    for (var groupIndex = 0; groupIndex < gapGroups.length; groupIndex++) {
      var gapGroup = gapGroups[groupIndex];
      if (round >= gapGroup.missingVolumes.length) continue;

      var targetVolume = gapGroup.missingVolumes[round];
      gapSpecs.push({
        group: gapGroup,
        kind: 'gap',
        query: (gapGroup.searchBase || gapGroup.base) + ' ' + targetVolume,
        targetVolume: targetVolume,
        page: 1
      });
      added = true;
    }
    round++;
  }

  // 기존 1페이지 작업 뒤에 추가 페이지를 붙여 저장된 진행 위치를 유지합니다.
  var extraSpecs = [];
  gapSpecs.forEach(function(spec) {
    for (var page = 2; page <= SERIES_MAX_GAP_PAGES; page++) {
      extraSpecs.push({group: spec.group, kind: spec.kind, query: spec.query,
        targetVolume: spec.targetVolume, page: page});
    }
  });
  followupSpecs.forEach(function(spec) {
    extraSpecs.push({group: spec.group, kind: spec.kind, query: spec.query, page: 2});
  });
  var allSpecs = followupSpecs.concat(gapSpecs, extraSpecs);
  var totalSpecCount = allSpecs.length;
  var startOffset = Math.max(0, Math.min(Number(resumeOffset || 0), totalSpecCount));
  var remainingJobs = Math.max(0, totalSpecCount - startOffset);
  var availableToday = Math.max(0, seriesApiRemaining_());
  var batchLimit = Math.min(SERIES_MAX_API_REQUESTS, availableToday, remainingJobs);

  if (remainingJobs > 0 && batchLimit < 1) {
    throw new Error(
      '오늘 안전하게 사용할 수 있는 YES24 조회량을 모두 사용했습니다. 기존 결과와 진행 위치는 그대로 보존했습니다. 내일 다시 점검해 주세요.'
    );
  }

  var specs = allSpecs.slice(startOffset, startOffset + batchLimit);
  var responses = [];
  var searchedGroups = {};

  for (var start = 0; start < specs.length; start += SERIES_SEARCH_CHUNK) {
    if (start > 0 && Date.now() - Number(runStartedAt || 0) >= SERIES_EXECUTION_BUDGET_MS) break;
    var chunk = specs.slice(start, start + SERIES_SEARCH_CHUNK);
    var requests = chunk.map(function(spec) {
      searchedGroups[spec.group.key] = true;
      return {
        url: seriesBuildSearchUrl_(spec.query, apiKey, spec.kind, spec.page),
        method: 'get',
        headers: {'X-Api-Key': apiKey},
        muteHttpExceptions: true,
        followRedirects: true
      };
    });

    var batch = yes24FetchAll_(requests);
    seriesRecordApiUsage_(batch.length);
    batch.forEach(function(response, index) {
      responses.push({spec: chunk[index], response: response});
    });

    seriesSpreadsheet_().toast(
      Math.min(startOffset + start + chunk.length, totalSpecCount) +
        ' / ' + totalSpecCount + '건 확인 중',
      'YES24 검색',
      4
    );
    Utilities.sleep(180);
  }

  var bySeriesVolume = {};
  var apiErrors = [];

  responses.forEach(function(entry) {
    var response = entry.response;
    var spec = entry.spec;
    if (response.getResponseCode() !== 200) {
      apiErrors.push(spec.query + ': HTTP ' + response.getResponseCode());
      return;
    }

    var data;
    try {
      data = JSON.parse(response.getContentText('UTF-8'));
    } catch (ignore) {
      apiErrors.push(spec.query + ': 응답 해석 실패');
      return;
    }
    if (data.errorCode) {
      apiErrors.push(spec.query + ': ' + (data.errorMessage || data.errorCode));
      return;
    }

    var items = data.data && Array.isArray(data.data.items) ? data.data.items : [];
    items.forEach(function(item) {
      var candidate = seriesCandidateFromItem_(spec.group, item, holdings);
      if (!candidate) return;
      if (spec.targetVolume && candidate.volume !== spec.targetVolume) return;

      var uniqueKey = spec.group.key + '|' + candidate.volume;
      var previous = bySeriesVolume[uniqueKey];
      if (!previous || seriesCandidateScore_(candidate) > seriesCandidateScore_(previous)) {
        bySeriesVolume[uniqueKey] = candidate;
      }
    });
  });

  if (apiErrors.length > 0) {
    throw new Error('YES24 조회 중 오류가 발생해 결과와 진행 위치를 보존했습니다. API 키와 호출 한도를 확인해 주세요.\n' + apiErrors.slice(0, 3).join('\n'));
  }

  var candidates = Object.keys(bySeriesVolume)
    .map(function(key) { return bySeriesVolume[key]; })
    .sort(function(a, b) {
      var typeOrder = {'중간 누락': 0, '신간 후속권': 1};
      if (typeOrder[a.type] !== typeOrder[b.type]) return typeOrder[a.type] - typeOrder[b.type];
      var baseOrder = a.series.localeCompare(b.series, 'ko');
      return baseOrder || a.volume - b.volume;
    });

  var nextOffset = startOffset + responses.length;
  return {
    candidates: candidates,
    requestCount: responses.length,
    searchedGroupCount: Object.keys(searchedGroups).length,
    truncated: nextOffset < totalSpecCount,
    errorCount: apiErrors.length,
    startOffset: startOffset,
    nextOffset: nextOffset,
    totalSpecCount: totalSpecCount
  };
}

function seriesBuildSearchUrl_(query, apiKey, kind, page) { return yes24Url_('/goods/itemList', {query: query, category: 'BOOK', sort: kind === 'gap' ? 'RELATION' : 'RECENT', page: Math.max(1, Number(page || 1)), pageSize: kind === 'gap' ? SERIES_GAP_RESULTS : SERIES_FOLLOWUP_RESULTS, detail: 'Y'}); }

function seriesCandidateFromItem_(group, item, holdings) {
  var title = seriesCleanText_(item.title);
  if (!title || seriesExcludedTitle_(title)) return null;

  var parsed = seriesParseTitle_(title);
  if (!parsed) return null;

  var publisherKeys = Object.keys(group.publishers || {});
  var publisherMatched = publisherKeys.length === 0 ||
    seriesPublisherMatches_(group.publishers, item.publisher);
  var exactKey = Boolean(group.aliasKeys && group.aliasKeys[parsed.key]);
  var compatibleKey = exactKey;

  if (!compatibleKey && publisherMatched) {
    var aliasKeys = Object.keys(group.aliasKeys || {});
    compatibleKey = aliasKeys.some(function(aliasKey) {
      return seriesKeysCompatible_(parsed.key, aliasKey, true);
    });
  }

  var candidateFormatKey = seriesFormatKey_(title);
  var sameFormat = candidateFormatKey === group.formatKey;

  // ISBN으로 확인한 학교 보유 판본과 같은 시리즈·출판사·판형만 인정합니다.
  if (!compatibleKey || !publisherMatched || !sameFormat) return null;
  if (group.volumes[parsed.volume]) return null;

  var isbn = seriesCleanIsbn_(item.isbn13 || item.isbn10);
  if (!isbn || holdings.isbnSet[isbn]) return null;
  if (seriesHoldingMatchesCandidate_(title, parsed, item.publisher, holdings)) return null;

  var cover = String(item.cover || '')
    
    ;

  return {
    type: parsed.volume > group.maxVolume ? '신간 후속권' : '중간 누락',
    series: group.base,
    seriesKey: parsed.key,
    formatKey: candidateFormatKey,
    publisherKey: seriesNormalizePublisher_(item.publisher),
    held: group.heldText,
    volume: parsed.volume,
    cover: /^https?:\/\//i.test(cover) ? cover : '',
    title: title,
    author: seriesCleanText_(item.author),
    publisher: seriesCleanText_(item.publisher),
    pubDate: seriesCleanText_(item.publishDate),
    isbn: isbn,
    price: Number(item.shopPrice || item.salePrice || 0),
    link: /^https?:\/\//i.test(String(item.link || '')) ? String(item.link) : '',
    saleStatus: /절판/.test(String(item.itemStatus || '')) ? '절판' : '',
    publisherMatched: publisherMatched,
    exactKey: exactKey,
    variantLabel: ''
  };
}

function seriesHoldingMatchesCandidate_(title, parsed, publisher, holdings) {
  // 조회 범위를 줄이지 않습니다. 보유 여부는 정확한 ISBN과 현재 시리즈 권차로만 판단합니다.
  return false;
}

function seriesCandidateScore_(candidate) {
  var score = candidate.publisherMatched ? 1000 : 0;
  score += candidate.exactKey ? 500 : 0;
  score += /^\d{13}$/.test(candidate.isbn) ? 100 : 0;
  score += candidate.cover ? 10 : 0;
  score += candidate.pubDate ? Number(String(candidate.pubDate).replace(/[^0-9]/g, '').slice(0, 8)) || 0 : 0;
  return score;
}

function seriesParseTitle_(title) {
  var clean = seriesCleanText_(title);
  if (!clean) return null;

  clean = seriesReplaceCircledNumerals_(clean);
  clean = seriesReplaceRomanNumerals_(clean);
  clean = clean
    .replace(/[\(\[【]\s*(\d{1,3})\s*[\)\]】]/g, ' $1권 ')
    .replace(/[\(\[【]\s*(개정판|증보판|특별판|양장본?|반양장|리커버|에디션)\s*[\)\]】]/gi, ' ')
    // 괄호 속 시리즈명은 버리지 않고 괄호 기호만 제거합니다.
    .replace(/[\(\)\[\]【】]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();

  var matches = [];
  var current;

  var englishPattern = /(?:vol(?:ume)?|no|season|part|level)\.?\s*(\d{1,3})(?=\s|[.:,_\-]|$)/ig;
  while ((current = englishPattern.exec(clean)) !== null) {
    matches.push({
      volume: Number(current[1]),
      numberIndex: current.index + current[0].lastIndexOf(current[1]),
      tokenIndex: current.index,
      explicit: true
    });
  }

  var explicitPattern = /(?:제\s*)?(\d{1,3})\s*(권|편|호|번째|탄|부|집)(?=\s|[.:,_\-]|$)/g;
  while ((current = explicitPattern.exec(clean)) !== null) {
    matches.push({
      volume: Number(current[1]),
      numberIndex: current.index + current[0].lastIndexOf(current[1]),
      tokenIndex: current.index,
      explicit: true
    });
  }

  if (!matches.length) {
    var standalonePattern = /(?:^|\s|[.:,_\-])(\d{1,3})(?=\s|[.:,_\-]|$)/g;
    while ((current = standalonePattern.exec(clean)) !== null) {
      var possible = Number(current[1]);
      var index = current.index + current[0].lastIndexOf(current[1]);
      if (possible > 0 && possible <= 200 && index >= 2) {
        matches.push({volume: possible, numberIndex: index, tokenIndex: index, explicit: false});
      }
    }
  }

  matches = matches.filter(function(match) {
    return match.volume > 0 && match.volume <= 200;
  });
  if (!matches.length) return null;

  matches.sort(function(a, b) {
    if (a.explicit !== b.explicit) return a.explicit ? 1 : -1;
    return a.numberIndex - b.numberIndex;
  });
  var chosen = matches[matches.length - 1];
  var baseEnd = chosen.explicit && chosen.tokenIndex < chosen.numberIndex
    ? chosen.tokenIndex
    : chosen.numberIndex;
  var base = clean.slice(0, baseEnd)
    .replace(/(?:제\s*)$/, '')
    .replace(/(?:vol(?:ume)?|no|season|part|level)\.?\s*$/i, '')
    .replace(/[\s.:,_\-]+$/, '')
    .trim();
  if (base.length < 2) return null;

  var key = seriesNormalizeKey_(base);
  if (key.length < 2) return null;
  return {base: base, key: key, volume: chosen.volume, explicit: chosen.explicit};
}

function seriesReplaceCircledNumerals_(value) {
  var chars = '①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳';
  return String(value || '').replace(/[①-⑳]/g, function(character) {
    return ' ' + (chars.indexOf(character) + 1) + '권 ';
  });
}

function seriesReplaceRomanNumerals_(value) {
  var romanMap = {
    'ⅩⅩ': 20, 'ⅩⅨ': 19, 'ⅩⅧ': 18, 'ⅩⅦ': 17, 'ⅩⅥ': 16,
    'ⅩⅤ': 15, 'ⅩⅣ': 14, 'ⅩⅢ': 13, 'ⅩⅡ': 12, 'ⅩⅠ': 11,
    'Ⅹ': 10, 'Ⅸ': 9, 'Ⅷ': 8, 'Ⅶ': 7, 'Ⅵ': 6,
    'Ⅴ': 5, 'Ⅳ': 4, 'Ⅲ': 3, 'Ⅱ': 2, 'Ⅰ': 1
  };
  return String(value || '').replace(/[ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩ]+/g, function(token) {
    return romanMap[token] ? ' ' + romanMap[token] + '권 ' : token;
  });
}

function seriesKeysCompatible_(left, right, publisherMatched) {
  var a = String(left || '');
  var b = String(right || '');
  if (!a || !b) return false;
  if (a === b) return true;

  var shorter = a.length <= b.length ? a : b;
  var longer = a.length > b.length ? a : b;
  if (shorter.length < 5 || longer.indexOf(shorter) === -1) return false;

  var ratio = shorter.length / longer.length;
  return ratio >= 0.88 || (publisherMatched && ratio >= 0.72);
}

function seriesNormalizeKey_(value) {
  return String(value || '')
    .toLowerCase()
    .replace(/[^0-9a-z가-힣]/g, '');
}

function seriesNormalizePublisher_(value) {
  return seriesNormalizeKey_(String(value || '')
    .replace(/주식회사|도서출판|출판사|\(주\)|㈜/g, ''));
}

function seriesPublisherMatches_(publisherMap, candidatePublisher) {
  var keys = Object.keys(publisherMap || {});
  if (!keys.length) return false;
  var candidate = seriesNormalizePublisher_(candidatePublisher);
  if (!candidate) return false;
  return keys.some(function(key) {
    return key === candidate ||
      (key.length >= 3 && candidate.indexOf(key) >= 0) ||
      (candidate.length >= 3 && key.indexOf(candidate) >= 0);
  });
}

function seriesExcludedTitle_(title) {
  return /세트|전집|박스|워크북|스티커북|색칠|오디오북|필사노트/i.test(String(title || ''));
}

function seriesCleanIsbn_(value) {
  var digits = String(value || '').replace(/[^0-9Xx]/g, '');
  return digits.length === 10 || digits.length === 13 ? digits.toUpperCase() : '';
}

function seriesCleanText_(value) {
  return String(value || '')
    .replace(/<[^>]*>/g, '')
    .replace(/&nbsp;/g, ' ')
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
    .replace(/\s+/g, ' ')
    .trim();
}

function seriesCompressNumbers_(numbers) {
  if (!numbers.length) return '';
  var parts = [];
  var start = numbers[0];
  var end = numbers[0];
  for (var index = 1; index <= numbers.length; index++) {
    var current = numbers[index];
    if (current === end + 1) {
      end = current;
      continue;
    }
    parts.push(start === end ? String(start) : start + '~' + end);
    start = current;
    end = current;
  }
  return parts.join(', ') + '권';
}

/** 현재 표시된 후보의 판매 상태를 YES24에서 다시 확인합니다. */
function refreshSeriesSaleStatuses() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) {
    seriesSpreadsheet_().toast('시리즈 점검이 진행 중입니다. 완료 후 다시 시도해 주세요.', '절판 확인', 7);
    return;
  }
  try {
    seriesCancelAutoContinue_();
    var sheet = seriesSpreadsheet_().getSheetByName(SERIES_FINDER_SHEET);
    if (!sheet || sheet.getLastRow() < 6) return;
    var key = seriesGetYes24Key_();
    if (!key) throw new Error('YES24 API 키가 없습니다.');
    var count = sheet.getLastRow() - 5;
    var rows = sheet.getRange(6, 1, count, 14).getDisplayValues();
    var jobs = rows.map(function(row, index) {
      var isbn = seriesCleanIsbn_(row[10]);
      return isbn ? {index: index, isbn: isbn} : null;
    }).filter(Boolean);
    if (seriesApiRemaining_() < jobs.length) throw new Error('오늘 남은 YES24 API 호출량이 부족합니다. 내일 다시 확인해 주세요.');
    var changed = 0;
    var startTime = Date.now();
    for (var start = 0; start < jobs.length; start += SERIES_SEARCH_CHUNK) {
      if (Date.now() - startTime >= SERIES_EXECUTION_BUDGET_MS) break;
      var chunk = jobs.slice(start, start + SERIES_SEARCH_CHUNK);
      var requests = chunk.map(function(job) {
        return {url: yes24Url_('/goods/itemDetail', {searchType: 'ISBN13', query: job.isbn, detail: 'N'}),
          method: 'get', headers: {'X-Api-Key': key}, muteHttpExceptions: true, followRedirects: true};
      });
      var responses = yes24FetchAll_(requests);
      seriesRecordApiUsage_(responses.length);
      responses.forEach(function(response, i) {
        if (response.getResponseCode() !== 200) return;
        var data;
        try { data = JSON.parse(response.getContentText('UTF-8')); } catch (ignore) { return; }
        var items = data && data.data && data.data.items;
        if (!data.success || !Array.isArray(items)) return;
        var book = items.filter(function(item) { return seriesCleanIsbn_(item.isbn13 || item.isbn10) === chunk[i].isbn; })[0];
        if (!book) return;
        var rowNumber = chunk[i].index + 6;
        var noteCell = sheet.getRange(rowNumber, 14);
        var note = String(noteCell.getValue() || '').replace(/^절판\s*·\s*/, '');
        var updatedNote = (/절판/.test(String(book.itemStatus || '')) ? '절판 · ' : '') + note;
        if (updatedNote !== String(noteCell.getValue() || '')) {
          noteCell.setValue(updatedNote);
          changed++;
        }
      });
      Utilities.sleep(180);
    }
    seriesApplyApiGauge_(sheet, sheet.getRange('A3').getDisplayValue(), sheet.getRange('E3').getDisplayValue());
    seriesSpreadsheet_().toast('절판 여부 확인 완료 · 표시 변경 ' + changed + '권', '시리즈 점검', 8);
  } finally {
    if (PropertiesService.getScriptProperties().getProperty(SERIES_RESUME_STATE_KEY)) {
      seriesScheduleAutoContinue_();
    }
    lock.releaseLock();
  }
}

function seriesBuildResultSheet_(candidates, meta) {
  var spreadsheet = seriesSpreadsheet_();
  var sheet = spreadsheet.getSheetByName(SERIES_FINDER_SHEET) || spreadsheet.insertSheet(SERIES_FINDER_SHEET);

  if (sheet.getFilter()) sheet.getFilter().remove();
  sheet.getRange(1, 1, Math.min(sheet.getMaxRows(), 8), Math.min(sheet.getMaxColumns(), 14)).breakApart();
  sheet.clear();
  sheet.clearConditionalFormatRules();
  sheet.setHiddenGridlines(true);

  var headers = [
    '선택', '구분', '시리즈명', '현재 보유 권차', '찾은 권차', '표지', '도서명',
    '저자', '출판사', '출간일', 'ISBN13', '정가', 'YES24', '점검 메모'
  ];

  sheet.getRange('A1:N1').merge().setValue('🔎 우리 학교 시리즈 빈칸 찾기');
  sheet.getRange('A2:N2').merge().setValue(
    '보유목록과 같은 시리즈·출판사·판형을 기준으로 빠진 권차와 새 후속권만 찾습니다.'
  );

  var checkedAt = Utilities.formatDate(meta.checkedAt || new Date(), Session.getScriptTimeZone(), 'yyyy. M. d. HH:mm');
  var progressTotal = Number(meta.progressTotal || 0);
  var progressDone = Math.min(progressTotal, Number(meta.progressDone || 0));
  var progressPercent = progressTotal ? (progressDone / progressTotal * 100).toFixed(1) : '100.0';
  var progressText = progressTotal
    ? '  ·  검색 진행 ' + progressDone + '/' + progressTotal + '건 (' + progressPercent + '%)' +
      (meta.truncated ? '  ·  남은 검색 ' + Math.max(0, progressTotal - progressDone) + '건' : '')
    : '';
  var summaryText =
    '확인 시리즈 ' + Number(meta.seriesCount || 0) + '개  ·  구매 검토 후보 ' +
    candidates.length + '권' + progressText +
    (meta.truncated ? '  ·  ⚠ 이어서 점검 필요' : '  ·  ✅ 전체 확인 완료');
  var lastCheckedText = '마지막 확인\n' + checkedAt;
  sheet.getRange(5, 1, 1, headers.length).setValues([headers]);

  if (candidates.length) {
    var rows = candidates.map(function(book) {
      var coverFormula = book.cover
        ? '=IFERROR(IMAGE("' + seriesEscapeFormula_(book.cover) + '",4,105,72),"")'
        : '';
      var linkFormula = book.link
        ? '=HYPERLINK("' + seriesEscapeFormula_(book.link) + '","YES24에서 보기")'
        : '';
      return [
        Boolean(book.selected),
        book.type,
        book.series,
        book.held,
        book.volume + '권',
        coverFormula,
        book.title,
        book.author,
        book.publisher,
        book.pubDate,
        book.isbn,
        book.price || '',
        linkFormula,
        (book.saleStatus === '절판' ? '절판 · ' : '') +
          (book.type === '신간 후속권'
            ? '현재 마지막 권 다음에 나온 책'
            : '보유 권차 사이에서 빠진 책')
      ];
    });
    sheet.getRange(6, 1, rows.length, headers.length).setValues(rows);
    sheet.getRange(6, 1, rows.length, 1).insertCheckboxes();
    sheet.getRange(6, 1, rows.length, headers.length).setVerticalAlignment('middle').setWrap(true);
    sheet.getRange(6, 12, rows.length, 1).setNumberFormat('#,##0"원"');
    sheet.setRowHeights(6, rows.length, 112);
    sheet.getRange(5, 1, rows.length + 1, headers.length).createFilter();
  } else {
    sheet.getRange('A6:N7').merge().setValue(
      Number(meta.searchedCount || 0) > 0
        ? '현재 YES24 검색 결과에서 추가 구매 후보를 찾지 못했습니다.'
        : '상단 메뉴 「🔎 시리즈 점검 → 후속권·누락권 이어서 점검」를 누르면 결과가 채워집니다.'
    );
    sheet.getRange('A6:N7').setHorizontalAlignment('center').setVerticalAlignment('middle');
  }

  seriesStyleResultSheet_(sheet, candidates.length);
  seriesApplyApiGauge_(sheet, summaryText, lastCheckedText);
  seriesAddWorkButton_(sheet);
}


/**
 * 시리즈 점검 상단에 API 사용량 막대를 표시합니다.
 * 표시를 새로고침하는 과정에서는 YES24 API를 호출하지 않습니다.
 */
function seriesApplyApiGauge_(sheet, summaryText, lastCheckedText) {
  var apiUsage = seriesApiUsageState_();
  var apiUsed = Math.max(0, Number(apiUsage.used || 0));
  var apiRemaining = Math.max(0, SERIES_DAILY_API_SAFE_LIMIT - apiUsed);
  var usageRatio = SERIES_DAILY_API_SAFE_LIMIT ? apiUsed / SERIES_DAILY_API_SAFE_LIMIT : 0;
  var apiPercent = (Math.min(1, usageRatio) * 100).toFixed(1);
  var barColor = usageRatio >= 0.9 ? '#C85B52' : usageRatio >= 0.7 ? '#D9A441' : '#2D6A50';
  var panelBackground = usageRatio >= 0.9 ? '#FDECEC' : usageRatio >= 0.7 ? '#FFF4D9' : '#EAF5EE';
  var panelText = usageRatio >= 0.9 ? '#A33B36' : usageRatio >= 0.7 ? '#8A6518' : '#2D6A50';

  sheet.getRange('A3:N3').breakApart().clearContent().clearNote();

  sheet.getRange('A3:D3')
    .merge()
    .setValue(summaryText)
    .setBackground('#F7F4EA')
    .setFontColor('#765A25')
    .setFontSize(10)
    .setFontWeight('bold')
    .setWrap(true)
    .setHorizontalAlignment('left')
    .setVerticalAlignment('middle');

  sheet.getRange('E3:F3')
    .merge()
    .setValue(lastCheckedText || '마지막 확인\n-')
    .setBackground('#EEF2EF')
    .setFontColor('#4E6357')
    .setFontSize(9)
    .setFontWeight('bold')
    .setWrap(true)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  sheet.getRange('G3:H3')
    .merge()
    .setFormula(
      '=SPARKLINE(' + apiUsed +
      ',{"charttype","bar";"max",' + SERIES_DAILY_API_SAFE_LIMIT +
      ';"color1","' + barColor + '"})'
    )
    .setBackground(panelBackground)
    .setVerticalAlignment('middle');

  sheet.getRange('I3')
    .setValue('API ' + apiUsed + '/' + SERIES_DAILY_API_SAFE_LIMIT + '\n남음 ' + apiRemaining + '회')
    .setNote('오늘 API 사용률 ' + apiPercent + '% · 이 표시를 확인하거나 새로고침하는 것 자체는 API를 사용하지 않습니다.')
    .setBackground(panelBackground)
    .setFontColor(panelText)
    .setFontSize(9)
    .setFontWeight('bold')
    .setWrap(true)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  sheet.getRange('J3:N3')
    .merge()
    .setValue('사용량 표시 새로고침은 API 미사용')
    .setBackground('#F6F8F7')
    .setFontColor('#78867F')
    .setFontSize(9)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  sheet.getRange('E3:I3').setBorder(
    true, true, true, true, false, false,
    usageRatio >= 0.9 ? '#E3A8A4' : usageRatio >= 0.7 ? '#E6CA87' : '#B7D7C4',
    SpreadsheetApp.BorderStyle.SOLID
  );
}

/**
 * API를 추가로 호출하지 않고 현재 사용량 막대만 다시 표시합니다.
 */
function refreshSeriesApiGauge() {
  var sheet = seriesSpreadsheet_().getSheetByName(SERIES_FINDER_SHEET);
  if (!sheet) {
    seriesSpreadsheet_().toast('시리즈 누락 점검 시트가 없습니다.', 'API 사용량', 5);
    return;
  }

  var currentText = String(sheet.getRange('A3').getDisplayValue() || '').trim();
  var lastMatch = currentText.match(
    /마지막 확인\s+(\d{4}\.\s*\d{1,2}\.\s*\d{1,2}\.\s*\d{1,2}:\d{2})/
  );
  var lastCheckedText = String(sheet.getRange('E3').getDisplayValue() || '').trim();

  if (lastMatch) {
    lastCheckedText = '마지막 확인\n' + lastMatch[1];
    currentText = currentText.replace(
      /\s*·\s*마지막 확인\s+\d{4}\.\s*\d{1,2}\.\s*\d{1,2}\.\s*\d{1,2}:\d{2}/,
      ''
    );
  }
  currentText = currentText
    .replace(/\s*·\s*오늘 API\s*\d+\/\d+회\s*\([^)]*\)\s*·\s*남음\s*\d+회/g, '')
    .trim();

  if (!currentText) currentText = '시리즈 점검 결과';
  if (!lastCheckedText) lastCheckedText = '마지막 확인\n-';

  seriesApplyApiGauge_(sheet, currentText, lastCheckedText);
  seriesSpreadsheet_().toast(
    'API를 호출하지 않고 사용량 표시만 새로고침했습니다.',
    'API 사용량',
    4
  );
}

function seriesStyleResultSheet_(sheet, rowCount) {
  sheet.setFrozenRows(5);
  sheet.setRowHeight(1, 62);
  sheet.setRowHeight(2, 42);
  sheet.setRowHeight(3, 46);
  sheet.setRowHeight(4, 14);
  sheet.setRowHeight(5, 42);

  sheet.getRange('A1:N1')
    .setBackground('#20382D')
    .setFontColor('#FFFFFF')
    .setFontSize(24)
    .setFontWeight('bold')
    .setHorizontalAlignment('left')
    .setVerticalAlignment('middle');
  sheet.getRange('A2:N2')
    .setBackground('#E9F1EC')
    .setFontColor('#40594D')
    .setFontSize(11)
    .setVerticalAlignment('middle');
  sheet.getRange('A3:D3')
    .setBackground('#F7F4EA')
    .setFontColor('#765A25')
    .setFontSize(10)
    .setFontWeight('bold')
    .setHorizontalAlignment('left')
    .setVerticalAlignment('middle');
  sheet.getRange('E3:N3')
    .setFontSize(9)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sheet.getRange('A5:N5')
    .setBackground('#526C5D')
    .setFontColor('#FFFFFF')
    .setFontSize(11)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  var widths = [56, 96, 190, 130, 82, 84, 280, 180, 125, 92, 125, 92, 122, 190];
  widths.forEach(function(width, index) { sheet.setColumnWidth(index + 1, width); });

  if (rowCount) {
    sheet.getRange(6, 1, rowCount, 14)
      .setFontSize(11)
      .setBorder(false, false, true, false, false, false, '#E2E8E4', SpreadsheetApp.BorderStyle.SOLID);
    sheet.getRange(6, 1, rowCount, 6).setHorizontalAlignment('center');
    sheet.getRange(6, 7, rowCount, 1).setFontWeight('bold');
    sheet.getRange(6, 13, rowCount, 1).setFontColor('#2D6A50').setFontWeight('bold');

    var dataRange = sheet.getRange(6, 1, rowCount, 14);
    sheet.setConditionalFormatRules([
      SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied('=$B6="중간 누락"')
        .setBackground('#FFF3EE')
        .setRanges([dataRange])
        .build(),
      SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied('=$B6="신간 후속권"')
        .setBackground('#EEF7F2')
        .setRanges([dataRange])
        .build(),
      SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied('=$A6=TRUE')
        .setBackground('#FFF0B8')
        .setFontColor('#3E3516')
        .setRanges([dataRange])
        .build()
    ]);
  }
  sheet.activate();
  seriesSpreadsheet_().setActiveRange(sheet.getRange('A1'));
}

function setupSeriesWorkSheet() {
  var spreadsheet = seriesSpreadsheet_();
  var sheet = spreadsheet.getSheetByName(SERIES_WORK_SHEET);

  if (!sheet) {
    sheet = spreadsheet.insertSheet(SERIES_WORK_SHEET);
  }

  seriesEnsureWorkSheet_(sheet);
  spreadsheet.setActiveSheet(sheet);
  spreadsheet.setActiveRange(sheet.getRange('A1'));
  spreadsheet.toast(
    '검색 결과에서 필요한 책을 체크한 뒤 「🔎 시리즈 점검 → 체크한 도서 업무목록으로 보내기」를 누르세요.',
    '업무목록 준비 완료',
    8
  );
}

function exportSelectedSeriesBooks() {
  var spreadsheet = seriesSpreadsheet_();
  var source = spreadsheet.getSheetByName(SERIES_FINDER_SHEET);
  if (!source) throw new Error('「' + SERIES_FINDER_SHEET + '」 시트를 찾을 수 없습니다.');

  var lastRow = source.getLastRow();
  if (lastRow < 6) {
    spreadsheet.toast('먼저 시리즈 누락 점검을 실행해 주세요.', '보낼 도서 없음', 6);
    return;
  }

  var values = source.getRange(6, 1, lastRow - 5, 14).getValues();
  var display = source.getRange(6, 1, lastRow - 5, 14).getDisplayValues();
  var formulas = source.getRange(6, 1, lastRow - 5, 14).getFormulas();
  var selectedIndexes = [];

  values.forEach(function(row, index) {
    if (row[0] === true) selectedIndexes.push(index);
  });

  if (!selectedIndexes.length) {
    spreadsheet.toast('업무목록으로 보낼 책의 「선택」 칸을 체크해 주세요.', '선택 필요', 7);
    return;
  }

  var target = spreadsheet.getSheetByName(SERIES_WORK_SHEET);
  if (!target) target = spreadsheet.insertSheet(SERIES_WORK_SHEET);
  seriesEnsureWorkSheet_(target);

  var existingKeys = {};
  var targetLastRow = target.getLastRow();
  if (targetLastRow >= 6) {
    var existing = target.getRange(6, 7, targetLastRow - 5, 5).getDisplayValues();
    existing.forEach(function(row) {
      var titleKey = seriesNormalizeKey_(row[0]);
      var isbn = seriesCleanIsbn_(row[4]);
      if (isbn) existingKeys['isbn:' + isbn] = true;
      if (titleKey) existingKeys['title:' + titleKey] = true;
    });
  }

  var output = [];
  var duplicateCount = 0;

  selectedIndexes.forEach(function(index) {
    var rowDisplay = display[index];
    var rowValues = values[index];
    var rowFormulas = formulas[index];
    var isbn = seriesCleanIsbn_(rowDisplay[10]);
    var titleKey = seriesNormalizeKey_(rowDisplay[6]);
    var duplicate =
      (isbn && existingKeys['isbn:' + isbn]) ||
      (titleKey && existingKeys['title:' + titleKey]);

    if (duplicate) {
      duplicateCount++;
      return;
    }

    var coverFormula = rowFormulas[5] || '';
    var linkFormula = rowFormulas[12] || '';
    var price = Number(rowValues[11] || 0) || '';

    output.push([
      '검토중',
      rowDisplay[1],
      rowDisplay[2],
      rowDisplay[3],
      rowDisplay[4],
      coverFormula,
      rowDisplay[6],
      rowDisplay[7],
      rowDisplay[8],
      rowDisplay[9],
      rowDisplay[10],
      price,
      linkFormula,
      new Date(),
      ''
    ]);

    if (isbn) existingKeys['isbn:' + isbn] = true;
    if (titleKey) existingKeys['title:' + titleKey] = true;
  });

  if (output.length) {
    var startRow = Math.max(target.getLastRow() + 1, 6);
    target.getRange(startRow, 1, output.length, 15).setValues(output);
    target.getRange(startRow, 12, output.length, 1).setNumberFormat('#,##0"원"');
    target.getRange(startRow, 14, output.length, 1).setNumberFormat('yyyy. m. d. HH:mm');
    target.getRange(startRow, 1, output.length, 15)
      .setVerticalAlignment('middle')
      .setWrap(true);
    for (var rowNumber = startRow; rowNumber < startRow + output.length; rowNumber++) {
      target.setRowHeight(rowNumber, 112);
    }
    seriesStyleWorkSheet_(target);
  }

  selectedIndexes.forEach(function(index) {
    source.getRange(index + 6, 1).setValue(false);
  });

  spreadsheet.setActiveSheet(target);
  spreadsheet.setActiveRange(
    output.length ? target.getRange(Math.max(target.getLastRow() - output.length + 1, 6), 1) : target.getRange('A1')
  );

  spreadsheet.toast(
    '업무목록에 ' + output.length + '권을 추가했습니다.' +
      (duplicateCount ? ' 이미 있는 ' + duplicateCount + '권은 제외했습니다.' : ''),
    '전송 완료',
    9
  );
}

function seriesEnsureWorkSheet_(sheet) {
  var headers = [
    '처리상태', '구분', '시리즈명', '현재 보유 권차', '구입 권차', '표지', '도서명',
    '저자', '출판사', '출간일', 'ISBN13', '정가', 'YES24', '추출일', '비고'
  ];
  var compatible = sheet.getLastRow() >= 5 &&
    String(sheet.getRange(5, 1).getDisplayValue()) === '처리상태';
  var empty = sheet.getLastRow() <= 1 && !String(sheet.getRange('A1').getDisplayValue()).trim();

  if (!compatible && !empty) {
    throw new Error('「' + SERIES_WORK_SHEET + '」 시트의 기존 형식이 달라 자동으로 덮어쓰지 않았습니다.');
  }

  if (empty) {
    if (sheet.getMaxColumns() < 15) {
      sheet.insertColumnsAfter(sheet.getMaxColumns(), 15 - sheet.getMaxColumns());
    }
    if (sheet.getMaxRows() < 1000) {
      sheet.insertRowsAfter(sheet.getMaxRows(), 1000 - sheet.getMaxRows());
    }
    sheet.clear();
    sheet.clearConditionalFormatRules();
    sheet.setHiddenGridlines(true);
    sheet.getRange('A1:O1').merge().setValue('📋 시리즈 구입 업무목록');
    sheet.getRange('A2:O2').merge().setValue(
      '시리즈 누락 점검 결과에서 실제 구입을 검토할 책만 모아 관리합니다. 같은 ISBN이나 도서명은 중복으로 들어오지 않습니다.'
    );
    sheet.getRange('A3:O3').merge().setValue(
      '업무 흐름  ·  검토중 → 구입결정 → 주문완료   |   구입하지 않을 책은 보류로 바꾸세요.'
    );
    sheet.getRange(5, 1, 1, headers.length).setValues([headers]);
  }

  var validation = SpreadsheetApp.newDataValidation()
    .requireValueInList(['검토중', '구입결정', '보류', '주문완료'], true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(6, 1, Math.max(1, sheet.getMaxRows() - 5), 1).setDataValidation(validation);
  seriesStyleWorkSheet_(sheet);
}

function seriesStyleWorkSheet_(sheet) {
  sheet.setHiddenGridlines(true);
  sheet.setFrozenRows(5);
  sheet.setTabColor('#D9A441');

  sheet.setRowHeight(1, 64);
  sheet.setRowHeight(2, 42);
  sheet.setRowHeight(3, 38);
  sheet.setRowHeight(4, 14);
  sheet.setRowHeight(5, 42);

  sheet.getRange('A1:O1')
    .setBackground('#253C32')
    .setFontColor('#FFFFFF')
    .setFontSize(24)
    .setFontWeight('bold')
    .setHorizontalAlignment('left')
    .setVerticalAlignment('middle');
  sheet.getRange('A2:O2')
    .setBackground('#EAF2ED')
    .setFontColor('#40594D')
    .setFontSize(11)
    .setVerticalAlignment('middle');
  sheet.getRange('A3:O3')
    .setBackground('#FFF5D8')
    .setFontColor('#73551D')
    .setFontSize(11)
    .setFontWeight('bold')
    .setVerticalAlignment('middle');
  sheet.getRange('A5:O5')
    .setBackground('#587064')
    .setFontColor('#FFFFFF')
    .setFontSize(11)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  var widths = [105, 95, 185, 125, 88, 86, 280, 175, 125, 95, 125, 95, 120, 125, 220];
  widths.forEach(function(width, index) {
    sheet.setColumnWidth(index + 1, width);
  });

  var lastRow = sheet.getLastRow();
  if (lastRow >= 6) {
    var rowCount = lastRow - 5;
    sheet.getRange(6, 1, rowCount, 15)
      .setFontSize(11)
      .setVerticalAlignment('middle')
      .setWrap(true)
      .setBorder(false, false, true, false, false, false, '#E2E8E4', SpreadsheetApp.BorderStyle.SOLID);
    sheet.getRange(6, 1, rowCount, 6).setHorizontalAlignment('center');
    sheet.getRange(6, 7, rowCount, 1).setFontWeight('bold');
    sheet.getRange(6, 12, rowCount, 1).setNumberFormat('#,##0"원"');
    sheet.getRange(6, 14, rowCount, 1).setNumberFormat('yyyy. m. d. HH:mm');
    for (var row = 6; row <= lastRow; row++) sheet.setRowHeight(row, 112);
  }

  if (sheet.getFilter()) sheet.getFilter().remove();
  if (sheet.getLastRow() >= 5) {
    sheet.getRange(5, 1, Math.max(1, sheet.getLastRow() - 4), 15).createFilter();
  }

  var dataRange = sheet.getRange(6, 1, Math.max(1, sheet.getMaxRows() - 5), 15);
  sheet.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$A6="검토중"')
      .setBackground('#FFF9E8')
      .setRanges([dataRange])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$A6="구입결정"')
      .setBackground('#E9F5EC')
      .setRanges([dataRange])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$A6="보류"')
      .setBackground('#F2F2F2')
      .setFontColor('#777777')
      .setRanges([dataRange])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$A6="주문완료"')
      .setBackground('#EAF0FA')
      .setRanges([dataRange])
      .build()
  ]);
}

function seriesEscapeFormula_(value) {
  return String(value || '').replace(/"/g, '""');
}



var UNIFIED_DASHBOARD_SHEET = '🏠 업무 대시보드';
var UNIFIED_REVIEW_SHEET = '🛒 구입 검토함';
var UNIFIED_UPLOAD_SHEET = '📥 엑셀 업로드 목록';
var UNIFIED_SETTINGS_SHEET = '⚙️ 통합 설정';
var UNIFIED_REQUEST_SHEET = '📚 희망도서 목록';
var UNIFIED_EXPORT_SHEET = '📦 서점 전달용';

function unifiedAddMenu_() {
  SpreadsheetApp.getUi()
    .createMenu('🧰 통합 도서업무')
    .addItem('🏠 전체 화면 정리·새로고침', 'setupUnifiedLibraryWorkflow')
    .addItem('🎨 전체 디자인 다시 적용', 'applyUnifiedLibraryDesign')
    .addSeparator()
    .addItem('📚 희망도서를 구입 검토함으로 보내기', 'sendHopeRequestsToReview')
    .addItem('🔎 시리즈 누락·후속권 이어서 점검', 'refreshMissingSeriesBooks')
    .addItem('🛒 체크한 시리즈를 구입 검토함으로 보내기', 'sendSelectedSeriesToReview')
    .addItem('📥 업로드 목록을 열기', 'openUnifiedUploadSheet')
    .addSeparator()
    .addItem('📦 중복 확인 후 서점 전달용으로 보내기', 'buildUnifiedBookstoreExportSheet')
    .addSeparator()
    .addSubMenu(
      SpreadsheetApp.getUi()
        .createMenu('🔎 시리즈 점검 도구')
        .addItem('현재 결과 절판 여부 확인', 'refreshSeriesSaleStatuses')
        .addItem('결과 시트 만들기·복구', 'setupSeriesFinder')
        .addItem('저장된 점검 위치 초기화', 'resetSeriesFinderProgress')
        .addSeparator()
        .addItem('YES24 API 키 설정', 'setSeriesYes24Key')
    )
    .addToUi();
}

function setupUnifiedLibraryWorkflow() {
  var ss = seriesSpreadsheet_();

  unifiedMigrateSheetNamesAndLayout_();

  if (!ss.getSheetByName(SERIES_FINDER_SHEET)) {
    seriesBuildResultSheet_([], {
      seriesCount: 0,
      searchedCount: 0,
      requestCount: 0,
      checkedAt: new Date()
    });
  }

  unifiedEnsureReviewSheet_();
  createUnifiedGuides_();
  sendHopeRequestsToReview();
  applyUnifiedLibraryDesign();
  unifiedRefreshDashboard_();

  ss.setActiveSheet(ss.getSheetByName(UNIFIED_DASHBOARD_SHEET));
  ss.toast('희망도서 → 구입 검토함 → 서점 전달용 흐름으로 정리했습니다.', '통합 설정 완료', 8);
}

function openUnifiedUploadSheet() {
  var ss = seriesSpreadsheet_();
  var sheet = unifiedEnsureUploadSheet_();
  var rowCount = Math.max(0, sheet.getLastRow() - 5);

  unifiedStyleUploadSheet_(sheet, rowCount);
  ss.setActiveSheet(sheet);
}

function refreshUnifiedPurchaseReview() {
  return unifiedRefreshPurchaseReview_([]);
}

function unifiedReadRequestItems_(sheet) {
  if (sheet.getLastRow() < FIRST_DATA_ROW) return [];

  var headers = sheet.getRange(HEADER_ROW, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  var map = unifiedHeaderMap_(headers);
  var rowCount = sheet.getLastRow() - FIRST_DATA_ROW + 1;
  var display = sheet.getRange(FIRST_DATA_ROW, 1, rowCount, sheet.getLastColumn()).getDisplayValues();
  var raw = sheet.getRange(FIRST_DATA_ROW, 1, rowCount, sheet.getLastColumn()).getValues();

  return display.map(function(row, index) {
    var status = map.status ? String(row[map.status - 1] || '') : '';
    if (/취소|오류|중복|소장/.test(status)) return null;
    if (map.sequence && !/^\d+$/.test(String(row[map.sequence - 1] || '').trim())) return null;

    var title = map.title ? String(row[map.title - 1] || '').trim() : '';
    var isbn = map.isbn ? String(row[map.isbn - 1] || '') : '';
    if (!title && !isbn) return null;

    var price = map.price
      ? Number(raw[index][map.price - 1] || String(row[map.price - 1] || '').replace(/[^0-9]/g, ''))
      : 0;

    return {
      source: '희망도서',
      title: title,
      author: map.author ? row[map.author - 1] : '',
      publisher: map.publisher ? row[map.publisher - 1] : '',
      year: map.year ? (String(row[map.year - 1] || '').match(/\d{4}/) || [''])[0] : '',
      isbn: isbn,
      price: price || 0,
      quantity: 1
    };
  }).filter(function(item) { return item; });
}

function unifiedReadSeriesItems_(sheet) {
  if (sheet.getLastRow() < 6) return [];
  var headers = sheet.getRange(5, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  var map = unifiedHeaderMap_(headers);
  var display = sheet.getRange(6, 1, sheet.getLastRow() - 5, sheet.getLastColumn()).getDisplayValues();
  var raw = sheet.getRange(6, 1, sheet.getLastRow() - 5, sheet.getLastColumn()).getValues();
  return display.map(function(row, index) {
    var status = map.status ? String(row[map.status - 1] || '') : '';
    if (/보류|주문완료/.test(status)) return null;
    var title = map.title ? String(row[map.title - 1] || '').trim() : '';
    var isbn = map.isbn ? String(row[map.isbn - 1] || '') : '';
    if (!title && !isbn) return null;
    return {
      source: '시리즈 누락',
      title: title,
      author: map.author ? row[map.author - 1] : '',
      publisher: map.publisher ? row[map.publisher - 1] : '',
      year: map.year ? (String(row[map.year - 1] || '').match(/\d{4}/) || [''])[0] : '',
      isbn: isbn,
      price: map.price ? Number(raw[index][map.price - 1] || String(row[map.price - 1] || '').replace(/[^0-9]/g, '')) : 0,
      quantity: 1
    };
  }).filter(function(item) { return item; });
}

function unifiedHeaderMap_(headers) {
  var map = {};
  headers.forEach(function(value, index) {
    var h = String(value || '').replace(/\s+/g, '').toUpperCase().replace(/\([^)]*\)/g, '');
    if (!map.sequence && /^(순번|번호|NO)$/.test(h)) map.sequence = index + 1;
    if (!map.title && /^(도서명|책제목|서명|제목)$/.test(h)) map.title = index + 1;
    if (!map.author && /^(저자|작가|글쓴이)$/.test(h)) map.author = index + 1;
    if (!map.publisher && /^(출판사|발행자|출판자)$/.test(h)) map.publisher = index + 1;
    if (!map.year && /^(출간일|출판연도|발행연도|출간연도|연도)$/.test(h)) map.year = index + 1;
    if (!map.isbn && /^ISBN(13)?$/.test(h)) map.isbn = index + 1;
    if (!map.price && /^(정가|도서정가|가격)$/.test(h)) map.price = index + 1;
    if (!map.status && /^(처리상태|자동입력상태|상태)$/.test(h)) map.status = index + 1;
  });
  return map;
}

function unifiedReadReviewManual_(sheet) {
  var result = {};
  if (sheet.getLastRow() < 6) return result;
  var rows = sheet.getRange(6, 1, sheet.getLastRow() - 5, 12).getValues();
  rows.forEach(function(row) {
    var isbn = seriesCleanIsbn_(row[7]);
    var titleKey = seriesNormalizeKey_(row[3]);
    var key = isbn ? 'isbn:' + isbn : 'title:' + titleKey;
    if (!key || key === 'title:') return;
    result[key] = {selected: row[0] === true, status: row[2], quantity: row[9], note: row[11]};
  });
  return result;
}

function unifiedEnsureReviewSheet_() {
  var ss = seriesSpreadsheet_();
  var sheet = ss.getSheetByName(UNIFIED_REVIEW_SHEET) || ss.insertSheet(UNIFIED_REVIEW_SHEET);

  if (sheet.getMaxColumns() < 12) sheet.insertColumnsAfter(sheet.getMaxColumns(), 12 - sheet.getMaxColumns());
  if (sheet.getMaxRows() < 1000) sheet.insertRowsAfter(sheet.getMaxRows(), 1000 - sheet.getMaxRows());

  var compatible = sheet.getLastRow() >= 5 && String(sheet.getRange('A5').getDisplayValue()) === '선택';
  var empty = sheet.getLastRow() <= 1 && !String(sheet.getRange('A1').getDisplayValue()).trim();
  if (!compatible && !empty) throw new Error('「' + UNIFIED_REVIEW_SHEET + '」에 다른 내용이 있어 덮어쓰지 않았습니다.');

  if (empty) {
    sheet.setHiddenGridlines(true);
    sheet.getRange('A1:L1').merge().setValue('🛒 구입 검토함');
    sheet.getRange('A2:L2').merge().setValue('희망도서와 선택한 시리즈 누락 도서를 한곳에서 최종 검토합니다.');
    sheet.getRange('A3:L3').merge().setValue('구입할 책을 체크한 뒤 아래 금색 버튼을 누르면 소장도서를 다시 확인하고 서점 전달용 목록을 만듭니다.');
    sheet.getRange(5, 1, 1, 12).setValues([[
      '선택','출처','처리상태','도서명','저자','출판사','출판연도','ISBN13','정가','수량','금액','비고'
    ]]);
  }

  sheet.getRange('A3:L3')
    .breakApart()
    .merge()
    .setValue('구입할 도서의 「선택」 체크박스를 누르면 「서점 전달용」 시트에 반영됩니다.');

  sheet.getRange('A4:L4').breakApart().clearContent().clearDataValidations();
  sheet.getRange('A4')
    .insertCheckboxes()
    .setValue(false)
    .setBackground('#2F5D50')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  sheet.getRange('B4:C4')
    .merge()
    .setValue('전체 선택/해제')
    .setBackground('#2F5D50')
    .setFontColor('#FFFFFF')
    .setFontSize(13)
    .setFontWeight('bold')
    .setHorizontalAlignment('left')
    .setVerticalAlignment('middle');

  sheet.getRange('D4')
    .insertCheckboxes()
    .setValue(false)
    .setBackground('#D9A441')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  sheet.getRange('E4:L4')
    .merge()
    .setValue('중복 확인 후 서점 전달용으로 보내기')
    .setBackground('#D9A441')
    .setFontColor('#FFFFFF')
    .setFontSize(14)
    .setFontWeight('bold')
    .setHorizontalAlignment('left')
    .setVerticalAlignment('middle');

  sheet.getRange('B4:C4')
    .merge()
    .setValue('전체 선택/해제')
    .setBackground('#2F5D50')
    .setFontColor('#FFFFFF')
    .setFontSize(13)
    .setFontWeight('bold')
    .setHorizontalAlignment('left')
    .setVerticalAlignment('middle');

  sheet.getRange('D4')
    .insertCheckboxes()
    .setValue(false)
    .setBackground('#D9A441')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  sheet.getRange('E4:L4')
    .merge()
    .setValue('중복 확인 후 서점 전달용으로 보내기')
    .setBackground('#D9A441')
    .setFontColor('#FFFFFF')
    .setFontSize(14)
    .setFontWeight('bold')
    .setHorizontalAlignment('left')
    .setVerticalAlignment('middle');

  unifiedStyleReviewSheet_(sheet);
  return sheet;
}

function unifiedStyleReviewSheet_(sheet) {
  sheet.setHiddenGridlines(true);
  sheet.setFrozenRows(5);

  sheet.getRange('A1:L1').setBackground('#183D35').setFontColor('#FFFFFF').setFontFamily('Arial').setFontSize(23).setFontWeight('bold').setHorizontalAlignment('left');
  sheet.getRange('A2:L2').setBackground('#EAF3ED').setFontColor('#34564D').setFontFamily('Arial').setFontSize(12);
  sheet.getRange('A3:L3').setBackground('#FFF4D6').setFontColor('#765413').setFontFamily('Arial').setFontSize(11).setFontWeight('bold');
  sheet.getRange('A4:L4').setBorder(true,true,true,true,false,false,'#B88428',SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sheet.getRange('A5:L5').setBackground('#2E5B4F').setFontColor('#FFFFFF').setFontFamily('Arial').setFontWeight('bold').setHorizontalAlignment('center');

  sheet.setRowHeight(1, 58);
  sheet.setRowHeight(2, 34);
  sheet.setRowHeight(3, 34);
  sheet.setRowHeight(4, 44);
  sheet.setRowHeight(5, 38);

  var widths = [65,110,105,260,170,135,95,135,100,75,110,220];
  widths.forEach(function(width, i) { sheet.setColumnWidth(i + 1, width); });

  sheet.getRange(6, 1, Math.max(1, sheet.getMaxRows() - 5), 12)
    .setWrap(true)
    .setVerticalAlignment('middle')
    .setFontFamily('Arial')
    .setFontSize(11)
    .setBackground('#FBFCFB');

  var formatRowCount = Math.max(1, sheet.getMaxRows() - 5);
  sheet.getRange(6, 9, formatRowCount, 1).setNumberFormat('#,##0"원"');
  sheet.getRange(6, 10, formatRowCount, 1).setNumberFormat('0');
  sheet.getRange(6, 11, formatRowCount, 1).setNumberFormat('#,##0"원"');
}

function buildUnifiedBookstoreExportSheet() {
  var ss = seriesSpreadsheet_();
  var source = ss.getSheetByName(UNIFIED_REVIEW_SHEET);

  if (!source || source.getLastRow() < 6) {
    throw new Error('먼저 희망도서 또는 시리즈 도서를 구입 검토함으로 보내 주세요.');
  }

  var rowCount = source.getLastRow() - 5;
  var rows = source.getRange(6, 1, rowCount, 12).getValues();
  var holdings = seriesReadHoldings_();
  var books = [];
  var ownedCount = 0;

  rows.forEach(function(row, index) {
    var selected = row[0] === true || String(row[2]) === '구입결정';
    if (!selected) return;

    var isbn = seriesCleanIsbn_(row[7]);
    var titleKey = seriesNormalizeKey_(row[3]);
    var alreadyOwned = !String(row[1] || '').includes('관리자') && Boolean(
      (isbn && holdings.isbnSet[isbn]) ||
      (titleKey && holdings.titleKeySet[titleKey])
    );

    if (alreadyOwned) {
      source.getRange(index + 6, 3).setValue('소장중');
      source.getRange(index + 6, 1).setValue(false);
      ownedCount++;
      return;
    }

    books.push(row);
  });

  if (!books.length) {
    throw new Error(
      ownedCount
        ? '선택한 도서는 모두 우리학교 소장도서입니다. 검토함의 상태를 확인해 주세요.'
        : '구입할 책의 선택 칸을 체크하거나 처리상태를 「구입결정」으로 바꿔 주세요.'
    );
  }

  var target = ss.getSheetByName(UNIFIED_EXPORT_SHEET) || ss.insertSheet(UNIFIED_EXPORT_SHEET);
  if (target.getFilter()) target.getFilter().remove();
  target.getBandings().forEach(function(b) { b.remove(); });
  target.getRange(1, 1, Math.min(target.getMaxRows(), 5), Math.min(target.getMaxColumns(), 10)).breakApart();
  target.clear();
  target.clearConditionalFormatRules();
  target.setHiddenGridlines(true);

  target.getRange('A1:J1').merge().setValue('📦 서점 전달용 도서 주문 목록');
  target.getRange('A2:J2').merge().setValue('구입 검토를 마친 희망도서와 시리즈 누락 도서입니다.');
  target.getRange('A3:J3').merge().setValue(
    ownedCount
      ? '우리학교 소장도서 ' + ownedCount + '권을 제외했습니다.'
      : '우리학교 소장도서와 다시 대조해 중복이 없는 도서만 담았습니다.'
  );
  target.getRange('A4')
    .insertCheckboxes()
    .setValue(false)
    .setBackground('#D9A441')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  target.getRange('B4:J4')
    .merge()
    .setValue('최신 구입 검토 결과로 다시 만들기')
    .setBackground('#D9A441')
    .setFontColor('#FFFFFF')
    .setFontSize(14)
    .setFontWeight('bold')
    .setHorizontalAlignment('left')
    .setVerticalAlignment('middle');
  target.getRange('A5:J5').setValues([[
    '번호','도서명','저자','출판사','출판연도','ISBN13','정가','수량','금액','출처·비고'
  ]]);

  var output = books.map(function(row, index) {
    var quantity = Math.max(1, Number(row[9] || 1));
    var price = Number(row[8] || 0);
    return [
      index + 1,row[3],row[4],row[5],row[6],row[7],
      price || '',quantity,price ? price * quantity : '',
      row[1] + (row[11] ? ' · ' + row[11] : '')
    ];
  });

  target.getRange(6, 1, output.length, 10).setValues(output);
  target.getRange(6, 7, output.length, 1).setNumberFormat('#,##0"원"');
  target.getRange(6, 8, output.length, 1).setNumberFormat('0');
  target.getRange(6, 9, output.length, 1).setNumberFormat('#,##0"원"');
  target.getRange(5, 1, output.length + 1, 10).createFilter();
  unifiedStyleExportSheet_(target, output.length);

  ss.setActiveSheet(target);
  unifiedRefreshDashboard_();
  ss.toast(
    '서점 전달용 ' + output.length + '권을 만들었습니다.' +
    (ownedCount ? ' 소장도서 ' + ownedCount + '권은 제외했습니다.' : ''),
    '중복 확인 완료',
    8
  );

  return {success:true, exported:output.length, owned:ownedCount};
}

function unifiedStyleExportSheet_(sheet, count) {
  sheet.setFrozenRows(5);
  sheet.getRange('A1:J1').setBackground('#183D35').setFontColor('#FFFFFF').setFontSize(24).setFontWeight('bold');
  sheet.getRange('A2:J2').setBackground('#F7E8BF').setFontColor('#6B4D11').setFontSize(12);
  sheet.getRange('A4:J4').setBorder(true,true,true,true,false,false,'#B88428',SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sheet.getRange('A5:J5').setBackground('#B77A2A').setFontColor('#FFFFFF').setFontWeight('bold').setHorizontalAlignment('center');
  sheet.setRowHeight(1, 60); sheet.setRowHeight(2, 34); sheet.setRowHeight(4, 44); sheet.setRowHeight(5, 40);
  var widths = [55,285,180,140,95,135,100,70,110,220];
  widths.forEach(function(width, i) { sheet.setColumnWidth(i + 1, width); });
  if (count) {
    sheet.getRange(6, 1, count, 10).setWrap(true).setVerticalAlignment('middle').setFontSize(11);
    sheet.getRange(6, 7, count, 1).setNumberFormat('#,##0"원"');
    sheet.getRange(6, 8, count, 1).setNumberFormat('0');
    sheet.getRange(6, 9, count, 1).setNumberFormat('#,##0"원"');
    sheet.setRowHeights(6, count, 60);
  }
}

function unifiedEnsureSettingsSheet_() {
  var ss = seriesSpreadsheet_();
  var sheet = ss.getSheetByName(UNIFIED_SETTINGS_SHEET) || ss.insertSheet(UNIFIED_SETTINGS_SHEET);
  sheet.clear();
  sheet.setHiddenGridlines(true);
  sheet.getRange('A1:H1').merge().setValue('⚙️ 통합 도서업무 설정');
  sheet.getRange('A2:H2').merge().setValue('희망도서 신청과 시리즈 누락 점검은 같은 YES24 API 키와 우리학교 소장도서 목록을 사용합니다.');
  sheet.getRange('A4:B8').setValues([
    ['확인 항목','현재 설정'],
    ['희망도서 원본 시트',UNIFIED_REQUEST_SHEET],
    ['소장도서 원본 시트',SERIES_HOLDINGS_SHEET],
    ['시리즈 결과 시트',SERIES_FINDER_SHEET],
    ['통합 검토 시트',UNIFIED_REVIEW_SHEET]
  ]);
  var key = seriesGetYes24Key_();
  sheet.getRange('D4:E6').setValues([
    ['API 상태',''],
    ['YES24 API',key ? '설정됨' : '설정 필요'],
    ['안내','API 키는 상단 「⚙️ 사본 설정」에서 입력합니다.']
  ]);
  sheet.getRange('A1:H1').setBackground('#183D35').setFontColor('#FFFFFF').setFontSize(23).setFontWeight('bold');
  sheet.getRange('A2:H2').setBackground('#EAF3ED').setFontColor('#34564D');
  sheet.getRange('A4:B4').setBackground('#2E5B4F').setFontColor('#FFFFFF').setFontWeight('bold');
  sheet.getRange('D4:E4').setBackground('#B77A2A').setFontColor('#FFFFFF').setFontWeight('bold');
  sheet.setColumnWidth(1, 190); sheet.setColumnWidth(2, 230); sheet.setColumnWidth(3, 30); sheet.setColumnWidth(4, 170); sheet.setColumnWidth(5, 330);
  return sheet;
}

function unifiedRefreshDashboard_() {
  var ss = seriesSpreadsheet_();
  var sheet = ss.getSheetByName(UNIFIED_DASHBOARD_SHEET) || ss.insertSheet(UNIFIED_DASHBOARD_SHEET, 0);

  sheet.getDataRange().getMergedRanges().forEach(function(range) {
    range.breakApart();
  });
  sheet.clear();
  sheet.setHiddenGridlines(true);

  var requestSource = ss.getSheetByName(UNIFIED_REQUEST_SHEET);
  var requestCount = requestSource ? unifiedReadRequestItems_(requestSource).length : 0;
  var seriesCount = unifiedCountData_(ss.getSheetByName(SERIES_FINDER_SHEET), 6);
  var uploadCount = unifiedCountData_(ss.getSheetByName(UNIFIED_UPLOAD_SHEET), 6);
  var review = ss.getSheetByName(UNIFIED_REVIEW_SHEET);
  var reviewCount = unifiedCountData_(review, 6);
  var exportCount = unifiedCountData_(ss.getSheetByName(UNIFIED_EXPORT_SHEET), 6);
  var decided = 0;
  var amount = 0;

  if (review && review.getLastRow() >= 6) {
    var rows = review.getRange(6, 1, review.getLastRow() - 5, 12).getValues();
    rows.forEach(function(row) {
      if (row[0] === true || String(row[2]) === '구입결정') {
        decided++;
        amount += Number(row[10] || 0);
      }
    });
  }

  sheet.getRange('A1:L1').merge().setValue('🏠 학교도서관 통합 구입 업무');
  sheet.getRange('A2:L2').merge().setValue(
    '희망도서·시리즈 후보·업로드 목록을 모아 중복을 확인하고 서점 전달용 목록으로 정리합니다.'
  );

  var cards = [
    ['A4:C4','희망도서 신청','A5:C7',requestCount + '권'],
    ['D4:F4','시리즈 후보','D5:F7',seriesCount + '권'],
    ['G4:I4','파일 업로드','G5:I7',uploadCount + '권'],
    ['J4:L4','통합 검토','J5:L7',reviewCount + '권']
  ];
  cards.forEach(function(card) {
    sheet.getRange(card[0]).merge().setValue(card[1]);
    sheet.getRange(card[2]).merge().setValue(card[3]);
  });

  sheet.getRange('A9:D9').merge().setValue('구입 결정  ' + decided + '권');
  sheet.getRange('E9:H9').merge().setValue('서점 전달용  ' + exportCount + '권');
  sheet.getRange('I9:L9').merge().setValue(
    '예상 금액  ' + String(amount).replace(/\B(?=(\d{3})+(?!\d))/g, ',') + '원'
  );

  sheet.getRange('A11:L11').merge().setValue('업무 순서');
  sheet.getRange('A12:L16').setValues([
    ['1. 자료 모으기','','','2. 구입 검토','','','3. 중복 확인','','','4. 서점 전달','',''],
    ['희망도서·시리즈·엑셀 목록','','','구입할 책 체크·수량 확인','','','소장도서와 자동 대조','','','정가×수량 확인 후 전달','',''],
    ['📚  🔎  📥','','','🛒 구입 검토함','','','🏫 소장도서 확인','','','📦 서점 전달용','',''],
    ['','','','','','','','','','','',''],
    ['필요한 기능은 상단 「🧰 통합 도서업무」 메뉴에서 실행할 수 있습니다.','','','','','','','','','','','']
  ]);

  sheet.getRange('A1:L1').setBackground('#183D35').setFontColor('#FFFFFF').setFontSize(28).setFontWeight('bold');
  sheet.getRange('A2:L2').setBackground('#183D35').setFontColor('#DCEAE4').setFontSize(12);

  ['A4:C4','D4:F4','G4:I4','J4:L4'].forEach(function(address) {
    sheet.getRange(address)
      .setBackground('#DCEBE5')
      .setFontColor('#34564D')
      .setFontWeight('bold')
      .setHorizontalAlignment('center');
  });
  ['A5:C7','D5:F7','G5:I7','J5:L7'].forEach(function(address) {
    sheet.getRange(address)
      .setBackground('#FAFBF8')
      .setFontColor('#183D35')
      .setFontSize(25)
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');
  });

  ['A9:D9','E9:H9','I9:L9'].forEach(function(address) {
    sheet.getRange(address)
      .setBackground('#F7E8BF')
      .setFontColor('#6B4D11')
      .setFontSize(15)
      .setFontWeight('bold')
      .setHorizontalAlignment('center');
  });

  sheet.getRange('A11:L11').setBackground('#2E5B4F').setFontColor('#FFFFFF').setFontWeight('bold');
  sheet.getRange('A12:L16')
    .setBackground('#F8F5ED')
    .setFontColor('#2F3D38')
    .setFontSize(12)
    .setWrap(true)
    .setVerticalAlignment('middle');

  for (var column = 1; column <= 12; column++) {
    sheet.setColumnWidth(column, 96);
  }
  sheet.setRowHeight(1, 68);
  sheet.setRowHeight(2, 34);
  sheet.setRowHeight(9, 50);
  sheet.setRowHeight(11, 38);
  sheet.setRowHeight(12, 42);
  sheet.setRowHeight(13, 58);
  sheet.setRowHeight(14, 42);
  sheet.setRowHeight(16, 42);
  sheet.setFrozenRows(2);
}

function unifiedCountData_(sheet, startRow) {
  if (!sheet || sheet.getLastRow() < startRow) return 0;
  return sheet.getRange(startRow, 1, sheet.getLastRow() - startRow + 1, Math.max(1, sheet.getLastColumn()))
    .getDisplayValues().filter(function(row) {
      return row.some(function(value) { return String(value || '').trim(); });
    }).length;
}


function unifiedAuditDataSources() {
  var ss = seriesSpreadsheet_();
  var requestSheet = ss.getSheetByName(UNIFIED_REQUEST_SHEET);
  var seriesSheet = ss.getSheetByName(SERIES_WORK_SHEET);
  var requestItems = requestSheet ? unifiedReadRequestItems_(requestSheet) : [];
  var seriesItems = seriesSheet ? unifiedReadSeriesItems_(seriesSheet) : [];
  var holdings = seriesReadHoldings_();
  console.log(JSON.stringify({
    requestLastRow: requestSheet ? requestSheet.getLastRow() : 0,
    requestHead: requestSheet ? requestSheet.getRange(1, 1, Math.min(4, requestSheet.getLastRow()), requestSheet.getLastColumn()).getDisplayValues() : [],
    requestItems: requestItems.length,
    seriesItems: seriesItems.length,
    holdings: holdings.rows.length
  }));
}


// 원본(예전) 시리즈 시트의 ID는 코드에 두지 않고 스크립트 속성 LEGACY_SERIES_SOURCE_ID 에 입력해 사용합니다.
function importLegacySeriesSheetsOnce() {
  var targetBook = seriesSpreadsheet_();
  var legacyId = String(
    PropertiesService.getScriptProperties().getProperty('LEGACY_SERIES_SOURCE_ID') || ''
  ).trim();
  if (!legacyId) {
    throw new Error('가져올 원본 시트 ID가 없습니다. 프로젝트 설정 → 스크립트 속성에 LEGACY_SERIES_SOURCE_ID 를 입력해 주세요.');
  }
  var sourceBook = SpreadsheetApp.openById(legacyId);
  var names = [SERIES_FINDER_SHEET, SERIES_WORK_SHEET];
  var imported = [];

  names.forEach(function(name) {
    var source = sourceBook.getSheetByName(name);
    if (!source) return;
    var current = targetBook.getSheetByName(name);
    var hasData = current && (
      (name === SERIES_FINDER_SHEET && current.getLastRow() >= 6 && String(current.getRange('G6').getDisplayValue()).trim()) ||
      (name === SERIES_WORK_SHEET && current.getLastRow() >= 6 && String(current.getRange('G6').getDisplayValue()).trim())
    );
    if (hasData) return;

    var copied = source.copyTo(targetBook);
    copied.setName('__통합중__' + name);
    var safeSheet = targetBook.getSheetByName(UNIFIED_DASHBOARD_SHEET) || targetBook.getSheets()[0];
    targetBook.setActiveSheet(safeSheet);
    if (current) targetBook.deleteSheet(current);
    copied.setName(name);
    imported.push(name);
  });

  refreshUnifiedPurchaseReview();
  unifiedRefreshDashboard_();
  targetBook.toast(
    imported.length ? imported.join(', ') + '의 기존 데이터를 가져왔습니다.' : '이미 통합된 시리즈 데이터가 있어 그대로 두었습니다.',
    '기존 자료 통합',
    8
  );
}


function unifiedReadExistingSeriesReview_(sheet) {
  var items = [];
  if (!sheet || sheet.getLastRow() < 6) return items;

  var rows = sheet.getRange(6, 1, sheet.getLastRow() - 5, 12).getValues();
  rows.forEach(function(row) {
    var source = String(row[1] || '').trim();
    if (/희망도서/.test(source)) return;
    if (/주문완료|소장중/.test(String(row[2] || ''))) return;
    if (!row[3] && !row[7]) return;

    items.push({
      source: source || '기존 검토',
      title: row[3],
      author: row[4],
      publisher: row[5],
      year: row[6],
      isbn: row[7],
      price: Number(row[8] || 0),
      quantity: Math.max(1, Number(row[9] || 1))
    });
  });

  return items;
}

function unifiedReadSelectedSeriesItems_(sheet) {
  if (!sheet || sheet.getLastRow() < 6) return [];

  var headers = sheet.getRange(5, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  var map = unifiedHeaderMap_(headers);
  var rowCount = sheet.getLastRow() - 5;
  var display = sheet.getRange(6, 1, rowCount, sheet.getLastColumn()).getDisplayValues();
  var raw = sheet.getRange(6, 1, rowCount, sheet.getLastColumn()).getValues();

  return display.map(function(row, index) {
    if (raw[index][0] !== true) return null;

    var title = map.title ? String(row[map.title - 1] || '').trim() : '';
    var isbn = map.isbn ? String(row[map.isbn - 1] || '') : '';
    if (!title && !isbn) return null;

    return {
      source: '시리즈 누락',
      title: title,
      author: map.author ? row[map.author - 1] : '',
      publisher: map.publisher ? row[map.publisher - 1] : '',
      year: map.year ? (String(row[map.year - 1] || '').match(/\d{4}/) || [''])[0] : '',
      isbn: isbn,
      price: map.price
        ? Number(raw[index][map.price - 1] || String(row[map.price - 1] || '').replace(/[^0-9]/g, ''))
        : 0,
      quantity: 1
    };
  }).filter(function(item) { return item; });
}

function unifiedRefreshPurchaseReview_(newSeriesItems) {
  var ss = seriesSpreadsheet_();
  var review = unifiedEnsureReviewSheet_();
  var oldManual = unifiedReadReviewManual_(review);
  var preservedSeries = unifiedReadExistingSeriesReview_(review);
  var holdings = seriesReadHoldings_();
  var items = preservedSeries.concat(Array.isArray(newSeriesItems) ? newSeriesItems : []);

  var requestSheet = ss.getSheetByName(UNIFIED_REQUEST_SHEET);
  if (requestSheet) {
    items = unifiedReadRequestItems_(requestSheet).concat(items);
  }

  var grouped = {};
  var order = [];
  var ownedSkipped = 0;

  items.forEach(function(item) {
    var isbn = seriesCleanIsbn_(item.isbn || item.isbn13 || item.isbn10);
    var titleKey = seriesNormalizeKey_(item.title);

    if ((isbn && holdings.isbnSet[isbn]) || (titleKey && holdings.titleKeySet[titleKey])) {
      ownedSkipped++;
      return;
    }

    var key = isbn ? 'isbn:' + isbn : 'title:' + titleKey;
    if (!key || key === 'title:') return;

    if (!grouped[key]) {
      grouped[key] = item;
      grouped[key].isbn = isbn;
      grouped[key].quantity = Math.max(1, Number(item.quantity || 1));
      grouped[key].sources = {};
      order.push(key);
    } else {
      grouped[key].quantity = Math.max(
        grouped[key].quantity,
        Math.max(1, Number(item.quantity || 1))
      );
      if (!grouped[key].author && item.author) grouped[key].author = item.author;
      if (!grouped[key].publisher && item.publisher) grouped[key].publisher = item.publisher;
      if (!grouped[key].year && item.year) grouped[key].year = item.year;
      if (!grouped[key].price && item.price) grouped[key].price = item.price;
    }

    grouped[key].sources[item.source] = true;
  });

  var output = order.map(function(key) {
    var item = grouped[key];
    var manual = oldManual[key] || oldManual['title:' + seriesNormalizeKey_(item.title)] || {};
    var quantity = Number(manual.quantity || item.quantity || 1);
    var price = Number(item.price || 0);

    return [
      manual.selected === true,
      Object.keys(item.sources).join(' + '),
      manual.status || '검토중',
      item.title,
      item.author,
      item.publisher,
      item.year,
      item.isbn,
      price || '',
      quantity,
      price ? price * quantity : '',
      manual.note || item.note || ''
    ];
  });

  if (review.getFilter()) review.getFilter().remove();
  review.getRange(6, 1, Math.max(1, review.getMaxRows() - 5), 12).clearContent().clearDataValidations();

  if (output.length) {
    review.getRange(6, 1, output.length, 12).setValues(output);
    review.getRange(6, 1, output.length, 1).insertCheckboxes();
    review.getRange(6, 3, output.length, 1).setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInList(['검토중','구입결정','보류','주문완료','소장중'], true)
        .setAllowInvalid(false)
        .build()
    );
    review.getRange(6, 9, output.length, 1).setNumberFormat('#,##0"원"');
    review.getRange(6, 10, output.length, 1).setNumberFormat('0');
    review.getRange(6, 11, output.length, 1).setNumberFormat('#,##0"원"');
    review.getRange(5, 1, output.length + 1, 12).createFilter();
    for (var rowNumber = 6; rowNumber < 6 + output.length; rowNumber++) {
      review.setRowHeight(rowNumber, 62);
    }
  }

  unifiedStyleReviewSheet_(review);
  unifiedRefreshDashboard_();

  return {
    count: output.length,
    ownedSkipped: ownedSkipped
  };
}

function sendHopeRequestsToReview() {
  var ss = seriesSpreadsheet_();
  var result = unifiedRefreshPurchaseReview_([]);
  var review = ss.getSheetByName(UNIFIED_REVIEW_SHEET);
  if (review) ss.setActiveSheet(review);

  ss.toast(
    '희망도서를 구입 검토함에 모았습니다. 현재 ' + result.count + '권입니다.',
    '전송 완료',
    7
  );

  return result;
}

function sendSelectedSeriesToReview() {
  var ss = seriesSpreadsheet_();
  var sheet = ss.getSheetByName(SERIES_FINDER_SHEET);
  if (!sheet) throw new Error('「' + SERIES_FINDER_SHEET + '」 시트를 찾을 수 없습니다.');

  var selected = unifiedReadSelectedSeriesItems_(sheet);
  if (!selected.length) {
    throw new Error('구입 검토함으로 보낼 시리즈 도서를 먼저 체크해 주세요.');
  }

  var result = unifiedRefreshPurchaseReview_(selected);
  sheet.getRange(6, 1, Math.max(1, sheet.getLastRow() - 5), 1).setValue(false);

  var review = ss.getSheetByName(UNIFIED_REVIEW_SHEET);
  if (review) ss.setActiveSheet(review);

  ss.toast(
    '선택한 시리즈 ' + selected.length + '권을 구입 검토함으로 보냈습니다.',
    '전송 완료',
    7
  );

  return result;
}



/**
 * ============================================================
 * 파일 업로드 목록
 * ============================================================
 */
function unifiedEnsureUploadSheet_() {
  var ss = seriesSpreadsheet_();
  var sheet = ss.getSheetByName(UNIFIED_UPLOAD_SHEET);

  if (!sheet) {
    sheet = ss.insertSheet(UNIFIED_UPLOAD_SHEET);
  }

  if (sheet.getMaxColumns() < 11) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), 11 - sheet.getMaxColumns());
  }
  if (sheet.getMaxRows() < 100) {
    sheet.insertRowsAfter(sheet.getMaxRows(), 100 - sheet.getMaxRows());
  }

  return sheet;
}

function unifiedReplaceUploadRows_(books, fileName) {
  var lock = LockService.getDocumentLock();
  lock.waitLock(30000);

  try {
    if (!Array.isArray(books) || !books.length) {
      throw new Error('업로드할 도서 정보가 없습니다.');
    }

    var limited = books.slice(0, 1500);
    var unique = [];
    var seen = {};
    var skipped = 0;

    limited.forEach(function(book) {
      book = book || {};

      var title = String(book.title || '').trim();
      var isbn = seriesCleanIsbn_(book.isbn || '');
      var publisher = String(book.publisher || '').trim();
      var yearMatch = String(book.year || '').match(/\d{4}/);
      var year = yearMatch ? yearMatch[0] : '';
      var price = Number(String(book.price || '').replace(/[^0-9.-]/g, '')) || 0;
      var quantity = Math.max(1, Number(book.quantity || 1));
      var seriesName = String(book.series || '').trim();

      if (!title && !isbn) {
        skipped++;
        return;
      }

      var key = isbn
        ? 'isbn:' + isbn
        : 'title:' + seriesNormalizeKey_(title) + '|' + seriesNormalizeKey_(publisher) + '|' + year;

      if (seen[key]) {
        skipped++;
        return;
      }
      seen[key] = true;

      unique.push({
        title: title,
        author: String(book.author || '').trim(),
        publisher: publisher,
        year: year,
        isbn: isbn,
        price: price,
        quantity: quantity,
        series: seriesName,
        note: String(book.note || '').trim()
      });
    });

    if (!unique.length) {
      throw new Error('도서명 또는 ISBN이 있는 행을 찾지 못했습니다.');
    }

    var sheet = unifiedEnsureUploadSheet_();
    var requiredRows = unique.length + 5;
    if (sheet.getMaxRows() < requiredRows) {
      sheet.insertRowsAfter(sheet.getMaxRows(), requiredRows - sheet.getMaxRows());
    }

    if (sheet.getFilter()) {
      sheet.getFilter().remove();
    }

    // 기존 시트에 남아 있는 고정 열과 병합 범위를 먼저 해제해야
    // 제목 영역을 다시 만들 때 병합/고정 충돌이 발생하지 않습니다.
    sheet.setFrozenRows(0);
    sheet.setFrozenColumns(0);
    sheet.getRange(1, 1, Math.min(5, sheet.getMaxRows()), 11)
      .getMergedRanges()
      .forEach(function(range) {
        range.breakApart();
      });

    sheet.clear();
    sheet.clearConditionalFormatRules();

    sheet.getRange('A1:K1').merge().setValue('📥 파일 업로드 도서 목록');
    sheet.getRange('A2:K2').merge().setValue(
      '엑셀·CSV에서 필요한 서지정보만 가져왔습니다. 구입할 책을 선택한 뒤 구입 검토함으로 보내세요.'
    );
    sheet.getRange('A3:K3').merge().setValue(
      '원본 파일: ' + String(fileName || '업로드 파일') + '  ·  불러온 도서 ' + unique.length + '권'
    );

    sheet.getRange('A4').insertCheckboxes().setValue(false);
    sheet.getRange('B4:C4').merge().setValue('목록 전체 선택 / 전체 해제');
    sheet.getRange('D4').insertCheckboxes().setValue(false);
    sheet.getRange('E4:K4').merge().setValue('선택한 책을 구입 검토함으로 보내기 →');

    var headers = [
      '선택', '도서명', '저자', '출판사', '출판연도',
      'ISBN13', '정가', '수량', '금액', '시리즈명', '비고'
    ];
    sheet.getRange(5, 1, 1, headers.length).setValues([headers]);

    var rows = unique.map(function(book) {
      return [
        false,
        book.title,
        book.author,
        book.publisher,
        book.year,
        book.isbn,
        book.price || '',
        book.quantity,
        book.price ? book.price * book.quantity : '',
        book.series,
        book.note
      ];
    });

    sheet.getRange(6, 1, rows.length, 11).setValues(rows);
    sheet.getRange(6, 1, rows.length, 1).insertCheckboxes();
    sheet.getRange(6, 6, rows.length, 1).setNumberFormat('@');
    sheet.getRange(6, 7, rows.length, 1).setNumberFormat('#,##0"원"');
    sheet.getRange(6, 8, rows.length, 1).setNumberFormat('0');
    sheet.getRange(6, 9, rows.length, 1).setNumberFormat('#,##0"원"');
    sheet.getRange(5, 1, rows.length + 1, 11).createFilter();

    unifiedStyleUploadSheet_(sheet, rows.length);
    SpreadsheetApp.flush();

    return {
      importedCount: rows.length,
      skippedCount: skipped + Math.max(0, books.length - limited.length),
      limited: books.length > 1500,
      sheetName: UNIFIED_UPLOAD_SHEET
    };
  } finally {
    lock.releaseLock();
  }
}

function unifiedReadSelectedUploadItems_(sheet) {
  if (!sheet || sheet.getLastRow() < 6) return [];

  var rowCount = sheet.getLastRow() - 5;
  var raw = sheet.getRange(6, 1, rowCount, 11).getValues();
  var display = sheet.getRange(6, 1, rowCount, 11).getDisplayValues();

  return raw.map(function(row, index) {
    if (row[0] !== true) return null;

    var title = String(display[index][1] || '').trim();
    var isbn = seriesCleanIsbn_(display[index][5] || '');
    if (!title && !isbn) return null;

    var seriesName = String(display[index][9] || '').trim();
    var note = String(display[index][10] || '').trim();

    return {
      source: '파일 업로드',
      title: title,
      author: display[index][2],
      publisher: display[index][3],
      year: (String(display[index][4] || '').match(/\d{4}/) || [''])[0],
      isbn: isbn,
      price: Number(row[6] || String(display[index][6] || '').replace(/[^0-9]/g, '')) || 0,
      quantity: Math.max(1, Number(row[7] || 1)),
      note: [seriesName ? '시리즈: ' + seriesName : '', note].filter(function(value) {
        return value;
      }).join(' · ')
    };
  }).filter(function(item) {
    return item;
  });
}

function sendSelectedUploadItemsToReview() {
  var ss = seriesSpreadsheet_();
  var sheet = ss.getSheetByName(UNIFIED_UPLOAD_SHEET);
  if (!sheet) {
    throw new Error('「' + UNIFIED_UPLOAD_SHEET + '」 시트를 찾을 수 없습니다.');
  }

  var selected = unifiedReadSelectedUploadItems_(sheet);
  if (!selected.length) {
    throw new Error('구입 검토함으로 보낼 책을 먼저 체크해 주세요.');
  }

  var result = unifiedRefreshPurchaseReview_(selected);
  sheet.getRange(6, 1, Math.max(1, sheet.getLastRow() - 5), 1).setValue(false);
  sheet.getRange('A4').setValue(false);
  sheet.getRange('D4').setValue(false);

  var review = ss.getSheetByName(UNIFIED_REVIEW_SHEET);
  if (review) ss.setActiveSheet(review);

  ss.toast(
    '선택한 ' + selected.length + '권을 구입 검토함으로 보냈습니다.',
    '전송 완료',
    7
  );

  return {
    sentCount: selected.length,
    reviewCount: result.count,
    ownedSkipped: result.ownedSkipped
  };
}

function unifiedStyleUploadSheet_(sheet, rowCount) {
  if (!sheet) return;

  sheet.setHiddenGridlines(true);

  // 제목과 안내 행은 여러 열에 병합되어 있으므로 열 고정은 사용하지 않습니다.
  // 행만 고정하면 병합 셀 오류 없이 제목과 헤더가 계속 보입니다.
  sheet.setFrozenColumns(0);
  sheet.setFrozenRows(5);
  sheet.setTabColor('#2F7258');

  sheet.getRange('A1:K1')
    .setBackground('#173F35')
    .setFontColor('#FFFFFF')
    .setFontFamily('Arial')
    .setFontSize(22)
    .setFontWeight('bold')
    .setHorizontalAlignment('left')
    .setVerticalAlignment('middle');
  sheet.getRange('A2:K2')
    .setBackground('#EAF3EE')
    .setFontColor('#456258')
    .setFontFamily('Arial')
    .setFontSize(11)
    .setFontWeight('bold')
    .setHorizontalAlignment('left')
    .setVerticalAlignment('middle');
  sheet.getRange('A3:K3')
    .setBackground('#FBF2DE')
    .setFontColor('#7A591E')
    .setFontFamily('Arial')
    .setFontSize(10)
    .setFontWeight('bold')
    .setHorizontalAlignment('left')
    .setVerticalAlignment('middle');

  sheet.getRange('A4:C4')
    .setBackground('#DDEDE4')
    .setFontColor('#24553F')
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setBorder(true, true, true, true, false, false, '#B8D2C3', SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange('D4:K4')
    .setBackground('#D9A943')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setBorder(true, true, true, true, false, false, '#B88727', SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange('A5:K5')
    .setBackground('#456354')
    .setFontColor('#FFFFFF')
    .setFontFamily('Arial')
    .setFontSize(11)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  sheet.setRowHeight(1, 58);
  sheet.setRowHeight(2, 36);
  sheet.setRowHeight(3, 32);
  sheet.setRowHeight(4, 42);
  sheet.setRowHeight(5, 40);

  var widths = [62, 285, 180, 150, 95, 145, 95, 65, 105, 190, 210];
  widths.forEach(function(width, index) {
    sheet.setColumnWidth(index + 1, width);
  });

  if (rowCount > 0) {
    var dataRange = sheet.getRange(6, 1, rowCount, 11);
    dataRange
      .setFontFamily('Arial')
      .setFontSize(10)
      .setFontColor('#263A32')
      .setVerticalAlignment('middle')
      .setWrap(true);

    sheet.getRange(6, 1, rowCount, 1).setHorizontalAlignment('center');
    sheet.getRange(6, 5, rowCount, 1).setHorizontalAlignment('center');
    sheet.getRange(6, 6, rowCount, 1)
      .setHorizontalAlignment('center')
      .setNumberFormat('@');
    sheet.getRange(6, 7, rowCount, 1)
      .setHorizontalAlignment('right')
      .setNumberFormat('#,##0"원"');
    sheet.getRange(6, 8, rowCount, 1)
      .setHorizontalAlignment('center')
      .setNumberFormat('0');
    sheet.getRange(6, 9, rowCount, 1)
      .setHorizontalAlignment('right')
      .setNumberFormat('#,##0"원"');

    sheet.setRowHeights(6, rowCount, 48);

    var backgrounds = [];
    for (var rowIndex = 0; rowIndex < rowCount; rowIndex++) {
      var rowColor = rowIndex % 2 === 1 ? '#F4F8F5' : '#FFFFFF';
      backgrounds.push(new Array(11).fill(rowColor));
    }
    dataRange.setBackgrounds(backgrounds);
    dataRange.setBorder(
      false, false, true, false, false, false,
      '#DCE6DF', SpreadsheetApp.BorderStyle.SOLID
    );
  }
}


function unifiedMigrateSheetNamesAndLayout_() {
  var ss = seriesSpreadsheet_();

  var requestOld = ss.getSheetByName('희망도서 목록');
  var requestNew = ss.getSheetByName(UNIFIED_REQUEST_SHEET);
  if (requestOld && !requestNew) {
    requestOld.setName(UNIFIED_REQUEST_SHEET);
    requestNew = requestOld;
  }

  var holdingsOld = ss.getSheetByName('우리학교 소장도서');
  var holdingsNew = ss.getSheetByName(SERIES_HOLDINGS_SHEET);
  if (holdingsOld && !holdingsNew) {
    holdingsOld.setName(SERIES_HOLDINGS_SHEET);
    holdingsNew = holdingsOld;
  }

  var guideOld = ss.getSheetByName('사용설명서');
  var applicantGuide = ss.getSheetByName('📗 희망도서 신청안내');
  if (guideOld && !applicantGuide) {
    guideOld.setName('📗 희망도서 신청안내');
  }

  requestNew = ss.getSheetByName(UNIFIED_REQUEST_SHEET);
  if (requestNew) {
    var rowOne = requestNew.getRange(1, 1, 1, Math.max(1, requestNew.getLastColumn())).getDisplayValues()[0];
    var rowFive = requestNew.getMaxRows() >= 5
      ? requestNew.getRange(5, 1, 1, Math.max(1, requestNew.getLastColumn())).getDisplayValues()[0]
      : [];
    var rowOneLooksLikeHeader = rowOne.some(function(value) {
      return /순번|도서명|책제목|ISBN|처리상태/.test(String(value || ''));
    });
    var rowFiveLooksLikeHeader = rowFive.some(function(value) {
      return /순번|도서명|책제목|ISBN|처리상태/.test(String(value || ''));
    });

    if (rowOneLooksLikeHeader && !rowFiveLooksLikeHeader) {
      requestNew.insertRowsBefore(1, 4);
    }
  }
}

function unifiedStyleHoldingsSheet_(sheet) {
  if (!sheet || sheet.getLastColumn() < 1) return;

  var lastColumn = sheet.getLastColumn();
  var lastRow = Math.max(sheet.getLastRow(), 1);
  var bodyRows = Math.max(1, sheet.getMaxRows() - 1);

  sheet.setHiddenGridlines(true);
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, lastColumn)
    .setBackground('#2E5B4F')
    .setFontColor('#FFFFFF')
    .setFontFamily('Arial')
    .setFontSize(13)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sheet.setRowHeight(1, 42);

  sheet.getRange(2, 1, bodyRows, lastColumn)
    .setBackground('#FBFCFB')
    .setFontColor('#2E3A33')
    .setFontFamily('Arial')
    .setFontSize(11)
    .setVerticalAlignment('middle')
    .setWrap(true);

  if (lastRow > 1) {
    sheet.setRowHeights(2, lastRow - 1, 44);
  }
}

function unifiedStyleCopySetupSheet_(sheet) {
  if (!sheet) return;

  // 설정 버튼은 이미 상단에 있으므로 병합된 행을 다시 이동하지 않습니다.
  unifiedArrangeCopySetupButtons_(sheet);

  var lastRow = Math.max(sheet.getLastRow(), 1);
  var lastColumn = Math.max(sheet.getLastColumn(), 10);
  var canvas = sheet.getRange(1, 1, lastRow, lastColumn);
  var mergedRanges = canvas.getMergedRanges();

  sheet.setHiddenGridlines(true);
  sheet.setFrozenRows(0);
  sheet.setFrozenColumns(0);

  canvas
    .setBackground('#F8F5ED')
    .setFontFamily('Arial')
    .setFontColor('#2E3A33')
    .setFontSize(12)
    .setVerticalAlignment('middle')
    .setWrap(true);

  sheet.setColumnWidth(1, 28);
  sheet.setColumnWidth(2, 78);
  for (var column = 3; column <= lastColumn; column++) {
    sheet.setColumnWidth(column, column <= 10 ? 86 : 72);
  }
  sheet.setColumnWidths(2, 3, 100);
  sheet.setColumnWidth(5, 28);
  sheet.setColumnWidths(6, 3, 100);
  sheet.setColumnWidth(9, 28);
  sheet.setColumnWidths(10, 3, 100);

  sheet.setRowHeights(1, lastRow, 28);

  mergedRanges.forEach(function(range) {
    var text = String(range.getDisplayValue() || '').trim();

    range
      .setBackground('#FFFFFF')
      .setFontColor('#2E3A33')
      .setFontFamily('Arial')
      .setFontSize(12)
      .setVerticalAlignment('middle')
      .setWrap(true)
      .setBorder(
        true,true,true,true,false,false,
        '#D9E2DC',SpreadsheetApp.BorderStyle.SOLID
      );

    if (!text) {
      range.setBackground('#F8F5ED').setBorder(false,false,false,false,false,false);
      return;
    }

    if (/다른 학교.*최초 설정|사본용 최초 설정/.test(text)) {
      range
        .setBackground('#183D35')
        .setFontColor('#FFFFFF')
        .setFontSize(25)
        .setFontWeight('bold')
        .setHorizontalAlignment('center');
      return;
    }

    if (/처음 한 번만/.test(text)) {
      range
        .setBackground('#EAF3ED')
        .setFontColor('#34564D')
        .setFontSize(13)
        .setFontWeight('bold')
        .setHorizontalAlignment('center');
      return;
    }

    if (/컴퓨터가 익숙하지 않아도/.test(text)) {
      range
        .setBackground('#D9A441')
        .setFontColor('#FFFFFF')
        .setFontSize(17)
        .setFontWeight('bold')
        .setHorizontalAlignment('center');
      return;
    }

    if (/^\d+$/.test(text)) {
      range
        .setBackground('#D9A441')
        .setFontColor('#FFFFFF')
        .setFontSize(24)
        .setFontWeight('bold')
        .setHorizontalAlignment('center');
      return;
    }

    if (/사이트 열기|API 키 설정|학교 최초 설정|공개 신청 링크 등록|버튼 모음|메뉴 확인|이 탭 삭제/.test(text)) {
      range
        .setBackground('#DCEBE5')
        .setFontColor('#244D42')
        .setFontSize(13)
        .setFontWeight('bold')
        .setHorizontalAlignment('center')
        .setBorder(
          true,true,true,true,false,false,
          '#A8C8BB',SpreadsheetApp.BorderStyle.SOLID_MEDIUM
        );
      return;
    }

    if (/꼭 확인|주의|변경 금지|삭제하지/.test(text)) {
      range
        .setBackground('#FCEDEC')
        .setFontColor('#A33F3F')
        .setFontWeight('bold');
      return;
    }

    range
      .setBackground('#FFFFFF')
      .setFontColor('#2E3A33')
      .setFontSize(12);
  });

  for (var row = 1; row <= lastRow; row++) {
    var rowText = sheet.getRange(row, 1, 1, lastColumn).getDisplayValues()[0].join(' ').trim();
    if (rowText) sheet.setRowHeight(row, Math.max(sheet.getRowHeight(row), 32));
  }

  sheet.activate();
  sheet.setActiveRange(sheet.getRange('B3'));
}

function createUnifiedGuides_() {
  var applicantSteps = [
    ['1', '신청 페이지 열기', '학교에서 안내한 희망도서 신청 웹페이지를 엽니다.'],
    ['2', '신청자 이름 입력(선택)', '이름을 입력하면 우선 대출 혜택을 받을 수 있습니다. 입력하지 않아도 신청할 수 있습니다.'],
    ['3', '책 검색', '책 제목 또는 ISBN을 입력하고 「책 검색하기」를 누릅니다.'],
    ['4', '책 확인·장바구니 담기', '표지·저자·출판사·연도를 확인한 뒤 원하는 책을 장바구니에 담습니다. 다른 책도 이어서 검색할 수 있습니다.'],
    ['5', '선택한 도서 신청', '장바구니를 확인하고 「선택한 도서 신청하기」를 누릅니다.'],
    ['6', '신청 확인·취소', '「신청 목록」에서 접수 결과를 확인할 수 있고, 필요하면 신청 취소와 취소 되돌리기를 할 수 있습니다.']
  ];

  var staffSteps = [
    ['1', '희망도서 확인', '「📚 희망도서 목록」에서 접수 도서를 확인합니다.'],
    ['2', '시리즈 누락 점검', '「🔎 시리즈 누락 점검」에서 후속권·중간 누락권을 찾고 필요한 책만 체크합니다.'],
    ['3', '파일 목록 업로드(선택)', '웹페이지의 「📥 목록 업로드」에서 엑셀·CSV를 불러옵니다. 최대 1,500권을 전체 미리보기로 확인한 뒤 시트에 넣을 수 있습니다.'],
    ['4', '구입 검토함으로 모으기', '희망도서·시리즈 후보·업로드 목록에서 구입할 책을 체크하고 금색 버튼을 눌러 「🛒 구입 검토함」에 모읍니다.'],
    ['5', '구입 수량 확인', '검토함에서 구입할 책의 「선택」을 체크하고 수량과 비고를 확인합니다. 수량은 권수만 숫자로 입력합니다.'],
    ['6', '중복 확인·서점 전달', '검토함의 금색 버튼을 누르면 우리학교 소장도서와 다시 대조한 뒤 중복이 없는 책만 기존 「📦 서점 전달용」 시트에 정리됩니다.'],
    ['7', '금액 확인', '서점 전달용에서 정가×수량 금액이 자동 계산되는지 확인합니다.'],
    ['8', '엑셀로 전달', '서점 전달용 시트를 확인한 뒤 파일 → 다운로드 → Microsoft Excel(.xlsx)로 저장해 서점에 전달합니다.'],
    ['9', '기초자료 관리', '「🏫 우리학교 소장도서」의 도서명과 ISBN13을 최신 상태로 유지해야 중복 확인과 시리즈 점검이 정확하게 작동합니다.']
  ];

  unifiedBuildGuideSheet_(
    '📗 희망도서 신청안내',
    '📗 희망도서 신청안내',
    '신청자가 책을 찾고 여러 권을 장바구니에 담아 신청하는 방법입니다.',
    applicantSteps,
    '신청자 이름은 선택사항이며, 입력하면 우선 대출 혜택을 받을 수 있습니다.'
  );

  if (
    typeof HR_getConfiguredPublicWebAppUrl_ === 'function' &&
    typeof HR_updateGuidePublicLink_ === 'function'
  ) {
    var savedPublicUrl = HR_getConfiguredPublicWebAppUrl_();
    if (savedPublicUrl) {
      HR_updateGuidePublicLink_(savedPublicUrl);
    }
  }

  unifiedBuildGuideSheet_(
    '📘 전체 사용방법',
    '📘 학교도서관 구입 업무 전체 사용방법',
    '희망도서·시리즈 후보·엑셀 목록을 한곳에서 검토해 서점 전달용 목록으로 만드는 순서입니다.',
    staffSteps,
    '업무 흐름: 자료 모으기 → 구입 검토함 → 소장 중복 확인 → 서점 전달용'
  );
}

function unifiedBuildGuideSheet_(name, title, subtitle, steps, footer) {
  var ss = seriesSpreadsheet_();
  var sheet = ss.getSheetByName(name) || ss.insertSheet(name);
  var maxColumns = 12;
  var requiredRows = 6 + steps.length * 3;

  if (sheet.getFilter()) sheet.getFilter().remove();
  sheet.getBandings().forEach(function(b) { b.remove(); });
  sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).breakApart();
  sheet.clear();
  sheet.clearConditionalFormatRules();

  if (sheet.getMaxColumns() < maxColumns) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), maxColumns - sheet.getMaxColumns());
  }
  if (sheet.getMaxRows() < requiredRows) {
    sheet.insertRowsAfter(sheet.getMaxRows(), requiredRows - sheet.getMaxRows());
  }

  sheet.setHiddenGridlines(true);
  sheet.getRange('A1:L2')
    .merge()
    .setValue(title)
    .setBackground('#183D35')
    .setFontColor('#FFFFFF')
    .setFontFamily('Arial')
    .setFontSize(25)
    .setFontWeight('bold')
    .setHorizontalAlignment('left')
    .setVerticalAlignment('middle');

  sheet.getRange('A3:L3')
    .merge()
    .setValue(subtitle)
    .setBackground('#EAF3ED')
    .setFontColor('#34564D')
    .setFontFamily('Arial')
    .setFontSize(12)
    .setVerticalAlignment('middle');

  var startRow = 5;
  steps.forEach(function(step, index) {
    var row = startRow + index * 3;

    sheet.getRange(row, 1, 2, 2)
      .merge()
      .setValue(step[0])
      .setBackground('#D9A441')
      .setFontColor('#FFFFFF')
      .setFontFamily('Arial')
      .setFontSize(22)
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');

    sheet.getRange(row, 3, 1, 10)
      .merge()
      .setValue(step[1])
      .setBackground('#FFF4D6')
      .setFontColor('#684C18')
      .setFontFamily('Arial')
      .setFontSize(14)
      .setFontWeight('bold')
      .setVerticalAlignment('middle');

    sheet.getRange(row + 1, 3, 1, 10)
      .merge()
      .setValue(step[2])
      .setBackground('#FBFCFB')
      .setFontColor('#2E3A33')
      .setFontFamily('Arial')
      .setFontSize(12)
      .setWrap(true)
      .setVerticalAlignment('middle');

    sheet.getRange(row, 1, 2, 12).setBorder(
      true,true,true,true,false,false,
      '#E0D5B8',SpreadsheetApp.BorderStyle.SOLID
    );

    sheet.setRowHeight(row, 34);
    sheet.setRowHeight(row + 1, 52);
    sheet.setRowHeight(row + 2, 10);
  });

  var footerRow = startRow + steps.length * 3;
  sheet.getRange(footerRow, 1, 2, 12)
    .merge()
    .setValue(footer)
    .setBackground('#DCEBE5')
    .setFontColor('#244D42')
    .setFontFamily('Arial')
    .setFontSize(13)
    .setFontWeight('bold')
    .setWrap(true)
    .setVerticalAlignment('middle');

  for (var column = 1; column <= maxColumns; column++) {
    sheet.setColumnWidth(column, column <= 2 ? 54 : 96);
  }

  sheet.setRowHeight(1, 38);
  sheet.setRowHeight(2, 38);
  sheet.setRowHeight(3, 42);
  sheet.setFrozenRows(3);
}

function applyUnifiedLibraryDesign() {
  var ss = seriesSpreadsheet_();

  unifiedMigrateSheetNamesAndLayout_();
  createUnifiedGuides_();

  var request = ss.getSheetByName(UNIFIED_REQUEST_SHEET);
  if (request) styleRequestListSheet();

  var series = ss.getSheetByName(SERIES_FINDER_SHEET);
  if (series) {
    seriesStyleResultSheet_(series);
    seriesAddWorkButton_(series);
  }

  var review = unifiedEnsureReviewSheet_();
  unifiedStyleReviewSheet_(review);

  var uploadSheet = ss.getSheetByName(UNIFIED_UPLOAD_SHEET);
  if (uploadSheet) {
    unifiedStyleUploadSheet_(uploadSheet, Math.max(0, uploadSheet.getLastRow() - 5));
  }

  var exportSheet = ss.getSheetByName(UNIFIED_EXPORT_SHEET);
  if (exportSheet && exportSheet.getLastRow() >= 4) {
    unifiedStyleExportSheet_(exportSheet, Math.max(0, exportSheet.getLastRow() - 5));
  }

  unifiedStyleHoldingsSheet_(ss.getSheetByName(SERIES_HOLDINGS_SHEET));
  unifiedStyleCopySetupSheet_(ss.getSheetByName('⚙️ 사본 최초설정'));
  unifiedRefreshDashboard_();

  ss.toast('전체 시트의 제목·색상·글꼴을 같은 디자인으로 정리했습니다.','디자인 적용 완료',7);
}

function removeLegacyUnifiedSheetsAfterConfirmation() {
  var ss = seriesSpreadsheet_();
  ['⚙️ 통합 설정','📋 시리즈 업무목록'].forEach(function(name) {
    var sheet = ss.getSheetByName(name);
    if (sheet) ss.deleteSheet(sheet);
  });
  ss.toast('통합 설정과 시리즈 업무목록 시트를 삭제했습니다.','정리 완료',6);
}


function unifiedArrangeCopySetupButtons_(sheet) {
  if (!sheet || sheet.getLastRow() < 14) return;

  var lastColumn = Math.max(sheet.getLastColumn(), 15);
  var scanRows = Math.min(sheet.getLastRow(), 30);
  var values = sheet.getRange(1, 1, scanRows, lastColumn).getDisplayValues();
  var headerText = '🧰 설정 버튼 모음  ·  필요한 버튼을 순서대로 누르세요';
  var apiReady = false;
  var schoolReady = false;
  var publicReady = false;

  values.forEach(function(row) {
    row.forEach(function(cellValue) {
      var cellText = String(cellValue || '').trim();
      if (/설정 버튼 모음/.test(cellText)) headerText = cellText;
      if (/API 키 설정/.test(cellText) && /✅/.test(cellText)) apiReady = true;
      if (/(?:이 학교|학교) 최초 설정/.test(cellText) && /✅/.test(cellText)) schoolReady = true;
      if (/공개 (?:신청 )?링크 등록/.test(cellText) && /✅/.test(cellText)) publicReady = true;
    });
  });

  sheet.getRange('B7:O20').breakApart().clearContent();
  sheet.getRange('B7:L8').merge().setValue(headerText);
  sheet.getRange('B10:D13').merge().setValue((apiReady ? '✅ ' : '') + '🔑 API 키 설정');
  sheet.getRange('F10:H13').merge().setValue((schoolReady ? '✅ ' : '') + '🏫 학교 최초 설정');
  sheet.getRange('J10:L13').merge().setValue((publicReady ? '✅ ' : '') + '🔗 공개 링크 등록');

  sheet.setColumnWidths(2, 3, 100);
  sheet.setColumnWidth(5, 28);
  sheet.setColumnWidths(6, 3, 100);
  sheet.setColumnWidth(9, 28);
  sheet.setColumnWidths(10, 3, 100);

  [10, 11, 12, 13].forEach(function(rowNumber) {
    sheet.setRowHeight(rowNumber, 27);
  });
  sheet.getRange('B14:L20')
    .setBackground('#F8F5ED')
    .setBorder(false, false, false, false, false, false);
  sheet.setRowHeight(9, 8);
  for (var compactRow = 14; compactRow <= 20; compactRow++) {
    sheet.setRowHeight(compactRow, 4);
  }

  sheet.getImages().forEach(function(image) {
    var title = String(image.getAltTextTitle() || '');
    var description = String(image.getAltTextDescription() || '');
    if (title.indexOf('HR_COPY_SETUP_BUTTON') !== 0) return;

    if (/API 키/.test(description)) {
      image.setAnchorCell(sheet.getRange('B10')).setWidth(300).setHeight(104);
    } else if (/학교 최초 설정/.test(description)) {
      image.setAnchorCell(sheet.getRange('F10')).setWidth(300).setHeight(104);
    } else if (/공개 신청 링크|공개 링크/.test(description)) {
      image.setAnchorCell(sheet.getRange('J10')).setWidth(300).setHeight(104);
    }
  });
}

function unifiedMoveCopySetupButtonsToTop_(sheet) {
  if (!sheet || sheet.getLastRow() < 8) return;

  var lastRow = sheet.getLastRow();
  var lastColumn = Math.max(sheet.getLastColumn(), 1);
  var values = sheet.getRange(1, 1, lastRow, lastColumn).getDisplayValues();
  var startRow = 0;
  var endRow = 0;

  values.forEach(function(row, index) {
    var text = row.join(' ').trim();

    if (!startRow && /설정 버튼 모음/.test(text)) {
      startRow = index + 1;
    }

    if (/공개 신청 링크 등록/.test(text)) {
      endRow = Math.max(endRow, index + 1);
    }
  });

  if (!startRow || !endRow || endRow < startRow || startRow <= 8) {
    return;
  }

  var endMerged = sheet
    .getRange(endRow, 1, 1, lastColumn)
    .getMergedRanges();

  endMerged.forEach(function(range) {
    endRow = Math.max(endRow, range.getLastRow());
  });

  endRow = Math.min(lastRow, endRow + 1);

  sheet.moveRows(
    sheet.getRange(
      startRow,
      1,
      endRow - startRow + 1,
      sheet.getMaxColumns()
    ),
    7
  );
}


function seriesLeadingCore_(value) {
  var clean = seriesCleanText_(value);
  if (!clean) return null;

  var stripped = clean.replace(/^\s*(?:[\(\[【][^)\]】]{2,60}[\)\]】]\s*)+/, '').trim();
  if (!stripped || stripped === clean) return null;

  var parsed = seriesParseTitle_(stripped);
  if (!parsed || parsed.key.length < 2) return null;
  return parsed;
}


function seriesFamilySearchText_(value) {
  var text = seriesCleanText_(value);
  if (!text) return '';

  text = text
    .replace(/(?:학습\s*)?만화로\s*(?:보는|읽는|배우는)/gi, ' ')
    .replace(/그림으로\s*(?:보는|읽는|배우는)/gi, ' ')
    .replace(/(?:어린이|초등학생|청소년)을?\s*위한/gi, ' ')
    .replace(/학습만화|그래픽\s*노블|그림책|코믹스?|양장본?|반양장|보드북|팝업북|리커버|개정판|특별판/gi, ' ')
    .replace(/\s+/g, ' ')
    .trim();

  return seriesNormalizeKey_(text).length >= 4 ? text : seriesCleanText_(value);
}


function seriesFamilyKey_(value) {
  return seriesNormalizeKey_(seriesFamilySearchText_(value));
}


function seriesFormatKey_(value) {
  var key = seriesNormalizeKey_(value);
  var tags = [];

  if (/학습만화/.test(key)) tags.push('학습만화');
  else if (/만화로보는|만화로읽는|코믹|만화/.test(key)) tags.push('만화');

  if (/그림으로보는|그림으로읽는|그림책|그래픽노블/.test(key)) tags.push('그림');
  if (/양장|하드커버/.test(key)) tags.push('양장');
  if (/반양장|소프트커버/.test(key)) tags.push('반양장');
  if (/보드북/.test(key)) tags.push('보드북');
  if (/팝업북/.test(key)) tags.push('팝업북');
  if (/큰글자|대활자/.test(key)) tags.push('큰글자');
  if (/리커버/.test(key)) tags.push('리커버');
  if (/개정판|개정증보/.test(key)) tags.push('개정판');
  if (/컬러|올컬러/.test(key)) tags.push('컬러');

  return tags.length ? tags.join('+') : '일반';
}


/* ===== UpdateNotice.gs ===== */
/**
 * 📢 업데이트 소식
 *
 * 개발자가 「📢 업데이트 소식 관리」 구글 시트에 소식을 쓰면, 사본이 열릴 때 한 번 안내 창으로 보여 줍니다.
 * - 시트의 「상태」 칸이 「🟢 보이기」(또는 TRUE)인 줄 중 가장 위에 있는 소식 하나를 보여 줍니다.
 * - 한 번 본 소식은 다시 안 뜹니다. 제목이나 내용을 바꾸면 새 소식으로 보고 다시 보여 줍니다.
 * - 시작일·종료일은 비워도 됩니다. 기간 밖이면 보여 주지 않습니다.
 * - 인터넷이 안 되거나 시트를 못 읽으면 조용히 건너뜁니다.
 */
var HR_NOTICE_SHEET_ID = '1aR1O-pNTpSotlsJ5uCvRNHioUdMNhC2vKVOlsluITdM';
var HR_NOTICE_URL =
  'https://docs.google.com/spreadsheets/d/' + HR_NOTICE_SHEET_ID +
  '/export?format=csv';
var HR_NOTICE_SEEN_PROPERTY = 'HR_NOTICE_SEEN_ID';

// CSV 한 덩어리를 2차원 배열로 바꿉니다(따옴표·줄바꿈 처리).
function HR_parseCsv_(text) {
  var rows = [];
  var row = [];
  var field = '';
  var inQuotes = false;
  for (var i = 0; i < text.length; i++) {
    var ch = text.charAt(i);
    if (inQuotes) {
      if (ch === '"') {
        if (text.charAt(i + 1) === '"') { field += '"'; i++; } else { inQuotes = false; }
      } else {
        field += ch;
      }
    } else if (ch === '"') {
      inQuotes = true;
    } else if (ch === ',') {
      row.push(field); field = '';
    } else if (ch === '\n' || ch === '\r') {
      if (ch === '\r' && text.charAt(i + 1) === '\n') i++;
      row.push(field); field = '';
      rows.push(row); row = [];
    } else {
      field += ch;
    }
  }
  if (field !== '' || row.length) { row.push(field); rows.push(row); }
  return rows;
}

function HR_noticeIsOn_(value) {
  // 「🟢 보이기」 같은 그림 글자는 떼고 글자만 비교합니다.
  var text = String(value || '').replace(/[^0-9A-Za-z가-힣]/g, '').toLowerCase();
  return ['true', 'o', 'y', 'yes', 'on', '1', '예', '표시', '체크', '켜짐', '보이기'].indexOf(text) >= 0;
}

// 2026-12-01 · 2026. 12. 1 · 2026/12/1 · 2026.12.1 → 2026-12-01 (못 읽으면 빈 글자)
function HR_updateNoticeDate_(value) {
  var m = String(value || '').trim().match(/^(\d{4})\D+(\d{1,2})\D+(\d{1,2})/);
  if (!m) return '';
  return m[1] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[3]).slice(-2);
}

function HR_noticeHash_(text) {
  var hash = 5381;
  for (var i = 0; i < text.length; i++) {
    hash = ((hash * 33) ^ text.charCodeAt(i)) >>> 0;
  }
  return 'n' + hash.toString(36);
}

function HR_fetchNotice_() {
  var response = UrlFetchApp.fetch(HR_NOTICE_URL, { muteHttpExceptions: true });
  if (response.getResponseCode() !== 200) return null;
  var rows = HR_parseCsv_(response.getContentText('UTF-8').replace(/^\uFEFF/, ''));

  var headerIndex = -1;
  for (var r = 0; r < rows.length; r++) {
    var head = String(rows[r][0] || '').replace(/\s+/g, ' ').trim();
    if (head.indexOf('상태') === 0 || head === '표시') { headerIndex = r; break; }
  }
  if (headerIndex < 0) return null; // 시트를 못 읽었거나 모양이 다름

  var today = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd');
  for (var i = headerIndex + 1; i < rows.length; i++) {
    var line = rows[i];
    if (!HR_noticeIsOn_(line[0])) continue;
    var title = String(line[1] || '').trim();
    var body = String(line[2] || '').trim();
    if (!body) continue;
    var start = HR_updateNoticeDate_(line[3]);
    var end = HR_updateNoticeDate_(line[4]);
    if (start && today < start) continue;
    if (end && today > end) continue;
    return {
      id: HR_noticeHash_(title + '|' + body),
      title: title || '업데이트 소식',
      body: body
    };
  }
  return null;
}

function HR_escapeNoticeHtml_(text) {
  return String(text)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}

function HR_showNoticeDialog_(notice) {
  var html = HtmlService.createHtmlOutput(
    '<!doctype html><html lang="ko"><head><meta charset="utf-8"><style>' +
    'body{margin:0;padding:24px;background:#f8f5ed;font-family:Arial,"Noto Sans KR",sans-serif;color:#2e3a33}' +
    '.badge{display:inline-block;background:#d9a441;color:#fff;font-weight:700;font-size:13px;padding:4px 12px;border-radius:999px;margin-bottom:12px}' +
    '.title{font-size:22px;font-weight:800;color:#183d35;margin:0 0 14px}' +
    '.body{font-size:15px;line-height:1.65;background:#fff;border:1px solid #d9e2dc;border-radius:12px;padding:16px;max-height:230px;overflow:auto}' +
    '.ok{display:block;width:100%;margin-top:16px;padding:13px;border:0;border-radius:12px;background:#244d42;color:#fff;font-size:16px;font-weight:700;cursor:pointer}' +
    '.ok:hover{background:#183d35}' +
    '</style></head><body><div class="badge">📢 업데이트 소식</div>' +
    '<h1 class="title">' + HR_escapeNoticeHtml_(notice.title) + '</h1>' +
    '<div class="body">' + HR_escapeNoticeHtml_(notice.body).replace(/\n/g, '<br>') + '</div>' +
    '<button class="ok" onclick="google.script.host.close()">확인했어요</button>' +
    '</body></html>'
  ).setWidth(460).setHeight(430);
  SpreadsheetApp.getUi().showModelessDialog(html, '📢 업데이트 소식');
}

// 시트를 열 때: 새 소식이면 한 번만 보여 줍니다.
function HR_showUpdateNoticeOnOpen_() {
  var notice = HR_fetchNotice_();
  if (!notice) return;
  var props = PropertiesService.getUserProperties();
  if (props.getProperty(HR_NOTICE_SEEN_PROPERTY) === notice.id) return;
  HR_showNoticeDialog_(notice);
  props.setProperty(HR_NOTICE_SEEN_PROPERTY, notice.id);
}

// 메뉴에서: 언제든 다시 봅니다.
function HR_showUpdateNotice() {
  var ui = SpreadsheetApp.getUi();
  var notice;
  try {
    notice = HR_fetchNotice_();
  } catch (error) {
    ui.alert('소식을 불러오지 못했습니다', '인터넷 연결을 확인하고 잠시 뒤 다시 시도해 주세요.', ui.ButtonSet.OK);
    return;
  }
  if (!notice) {
    ui.alert('📢 업데이트 소식', '지금 새 소식이 없습니다.', ui.ButtonSet.OK);
    return;
  }
  HR_showNoticeDialog_(notice);
}


/* ===== WebNotice.js ===== */
var HR_WEB_NOTICE_SHEET_NAME = '📢 웹 공지';

/* HR_NOTICE_SETUP_V2 */
function HR_upgradeNoticeDatePickersAndGuide() {
  // 추가로 갱신할 시트가 있으면 스크립트 속성 HR_EXTRA_SPREADSHEET_IDS 에 쉼표로 구분해 입력합니다.
  var targetIds = [HR_getSpreadsheet_().getId()].concat(
    String(
      PropertiesService.getScriptProperties().getProperty('HR_EXTRA_SPREADSHEET_IDS') || ''
    ).split(',').map(function(id) { return id.trim(); })
  );
  var visited = {};
  var results = [];

  targetIds.forEach(function(spreadsheetId) {
    if (!spreadsheetId || visited[spreadsheetId]) {
      return;
    }
    visited[spreadsheetId] = true;
    var spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    HR_applyNoticeDatePickersAndGuide_(spreadsheet);
    HR_ensureNoticeAudienceColumn_(spreadsheet);
    results.push(spreadsheet.getName());
  });

  SpreadsheetApp.flush();
  return '완료: ' + results.join(', ');
}

function HR_applyNoticeDatePickersAndGuide_(spreadsheet) {
  var notice = spreadsheet.getSheetByName(HR_WEB_NOTICE_SHEET_NAME);
  var isNew = !notice;

  if (!notice) {
    notice = spreadsheet.insertSheet(HR_WEB_NOTICE_SHEET_NAME);
  }

  if (notice.getMaxRows() < 1000) {
    notice.insertRowsAfter(notice.getMaxRows(), 1000 - notice.getMaxRows());
  }
  if (notice.getMaxColumns() < 5) {
    notice.insertColumnsAfter(notice.getMaxColumns(), 5 - notice.getMaxColumns());
  }

  if (isNew || !notice.getRange('A1').getDisplayValue()) {
    notice.getRange('A1:E1').merge().setValue('📢 웹페이지 공지 전광판');
    notice.getRange('A2:E2').merge().setValue(
      '표시에 체크하면 웹페이지 상단에 공지가 올라갑니다. 시작일·종료일은 비워도 됩니다.'
    );
    notice.getRange('A3:E3').setValues([[
      '표시', '종류', '공지 내용', '시작일', '종료일'
    ]]);
  }

  notice.getRange('A2').setValue(
    '표시에 체크하면 웹페이지 상단에 공지가 올라갑니다. ' +
    '시작일·종료일 셀을 더블 클릭하면 달력이 열립니다. 날짜는 비워도 됩니다.'
  );

  var dateRule = SpreadsheetApp.newDataValidation()
    .requireDate()
    .setAllowInvalid(false)
    .setHelpText('셀을 눌러 달력에서 날짜를 선택하세요.')
    .build();

  var typeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['일반', '긴급'], true)
    .setAllowInvalid(false)
    .build();

  notice.getRange('A4:A1000').insertCheckboxes();
  notice.getRange('B4:B1000').setDataValidation(typeRule);
  notice.getRange('D4:E1000')
    .setDataValidation(dateRule)
    .setNumberFormat('yyyy. m. d.');

  notice.setHiddenGridlines(true);
  notice.setFrozenRows(3);
  notice.setColumnWidth(1, 70);
  notice.setColumnWidth(2, 100);
  notice.setColumnWidth(3, 620);
  notice.setColumnWidth(4, 120);
  notice.setColumnWidth(5, 120);
  notice.setRowHeight(1, 58);
  notice.setRowHeight(2, 38);
  notice.setRowHeight(3, 36);

  notice.getRange('A1:E1')
    .setBackground('#244f42')
    .setFontColor('#ffffff')
    .setFontSize(20)
    .setFontWeight('bold')
    .setVerticalAlignment('middle');

  notice.getRange('A2:E2')
    .setBackground('#eef5f1')
    .setFontColor('#52685f')
    .setFontSize(11)
    .setWrap(true)
    .setVerticalAlignment('middle');

  notice.getRange('A3:E3')
    .setBackground('#d7a843')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  notice.getRange('A4:B1000').setHorizontalAlignment('center');
  notice.getRange('D4:E1000').setHorizontalAlignment('center');
  notice.getRange('A4:E1000').setVerticalAlignment('middle');

  var guide = spreadsheet.getSheetByName('📘 전체 사용방법');
  if (!guide) {
    return;
  }

  var guideTitle = '📢 웹 공지 등록 및 기간 설정';
  var titleCell = guide.createTextFinder(guideTitle)
    .matchEntireCell(true)
    .findNext();
  var titleRow = titleCell ? titleCell.getRow() : guide.getLastRow() + 2;

  if (!titleCell) {
    guide.getRange(titleRow, 1, 1, 12).merge();
    guide.getRange(titleRow + 1, 1, 2, 12).merge();
  }

  guide.getRange(titleRow, 1)
    .setValue(guideTitle);

  guide.getRange(titleRow + 1, 1)
    .setValue(
      '「📢 웹 공지」 시트에서 표시를 체크하고 공지 내용을 입력하세요. ' +
      '시작일·종료일 셀을 더블 클릭하면 달력이 열리며, 원하는 날짜를 선택할 수 있습니다. ' +
      '날짜를 비우면 바로 표시되고, 종료일을 비우면 직접 체크를 해제할 때까지 계속 표시됩니다.'
    );

  guide.setRowHeight(titleRow, 38);
  guide.setRowHeights(titleRow + 1, 2, 34);

  guide.getRange(titleRow, 1, 1, 12)
    .setBackground('#244f42')
    .setFontColor('#ffffff')
    .setFontSize(13)
    .setFontWeight('bold')
    .setVerticalAlignment('middle')
    .setBorder(true, true, true, true, false, false, '#b9cec5', SpreadsheetApp.BorderStyle.SOLID);

  guide.getRange(titleRow + 1, 1, 2, 12)
    .setBackground('#f7faf8')
    .setFontColor('#3e5149')
    .setFontSize(11)
    .setWrap(true)
    .setVerticalAlignment('middle')
    .setBorder(true, true, true, true, false, false, '#d6e2dc', SpreadsheetApp.BorderStyle.SOLID);
}


function HR_getActiveWebNotices(audience) {
  var audienceKey = HR_audienceKey_(audience);
  var spreadsheet = HR_getSpreadsheet_();
  var sheet = spreadsheet.getSheetByName(HR_WEB_NOTICE_SHEET_NAME);

  if (!sheet || sheet.getLastRow() < 4) {
    return [];
  }

  var values = sheet.getRange(4, 1, sheet.getLastRow() - 3, 6).getValues();
  var now = new Date();
  now.setHours(0, 0, 0, 0);

  return values.reduce(function(result, row) {
    var enabled = row[0] === true;
    var type = String(row[1] || '일반').trim();
    var message = String(row[2] || '').trim();
    var start = HR_noticeDate_(row[3], false);
    var end = HR_noticeDate_(row[4], true);

    var target = String(row[5] || '전체').trim();

    if (!enabled || !message) {
      return result;
    }
    // 「대상」이 교직원/학부모면 해당 구분에만 보여 줍니다. 비어 있거나 전체면 모두에게 보여 줍니다.
    if (
      (target === '교직원' && audienceKey !== 'staff') ||
      (target === '학부모' && audienceKey !== 'parent')
    ) {
      return result;
    }
    if (start && now.getTime() < start.getTime()) {
      return result;
    }
    if (end && now.getTime() > end.getTime()) {
      return result;
    }

    result.push({
      type: type === '긴급' ? '긴급' : '일반',
      message: message
    });
    return result;
  }, []).slice(0, 10);
}

function HR_noticeDate_(value, endOfDay) {
  if (!value) {
    return null;
  }

  var date = value instanceof Date
    ? new Date(value.getTime())
    : new Date(value);

  if (isNaN(date.getTime())) {
    return null;
  }

  date.setHours(
    endOfDay ? 23 : 0,
    endOfDay ? 59 : 0,
    endOfDay ? 59 : 0,
    0
  );
  return date;
}

function HR_setupWebNoticeSheet() {
  var spreadsheet = HR_getSpreadsheet_();
  var sheet = spreadsheet.getSheetByName(HR_WEB_NOTICE_SHEET_NAME);

  if (sheet) {
    return {
      success: true,
      created: false,
      sheetName: HR_WEB_NOTICE_SHEET_NAME
    };
  }

  sheet = spreadsheet.insertSheet(HR_WEB_NOTICE_SHEET_NAME);
  sheet.setHiddenGridlines(true);
  sheet.setTabColor('#d9a13b');

  sheet.getRange('A1:E1').merge()
    .setValue('📢 웹페이지 공지 전광판')
    .setBackground('#284b3f')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setFontSize(20)
    .setHorizontalAlignment('left')
    .setVerticalAlignment('middle');
  sheet.setRowHeight(1, 56);

  sheet.getRange('A2:E2').merge()
    .setValue('표시에 체크하면 웹페이지 상단에서 공지가 흘러갑니다. 시작일·종료일은 비워도 됩니다.')
    .setBackground('#eef5f1')
    .setFontColor('#4d665d')
    .setFontSize(11)
    .setWrap(true)
    .setVerticalAlignment('middle');
  sheet.setRowHeight(2, 42);

  sheet.getRange('A3:E3')
    .setValues([['표시', '종류', '공지 내용', '시작일', '종료일']])
    .setBackground('#d9a13b')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sheet.setRowHeight(3, 36);

  sheet.getRange('A4:A23').insertCheckboxes();

  var typeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['일반', '긴급'], true)
    .setAllowInvalid(false)
    .build();

  sheet.getRange('B4:B23')
    .setDataValidation(typeRule)
    .setValue('일반');
  sheet.getRange('C4:C23').setWrap(true);
  sheet.getRange('D4:E23').setNumberFormat('yyyy. m. d.');

  sheet.setColumnWidth(1, 74);
  sheet.setColumnWidth(2, 90);
  sheet.setColumnWidth(3, 620);
  sheet.setColumnWidth(4, 120);
  sheet.setColumnWidth(5, 120);
  sheet.setFrozenRows(3);

  sheet.getRange('A4:E23')
    .setVerticalAlignment('middle')
    .setBorder(
      false,
      false,
      true,
      false,
      false,
      false,
      '#e6ebe8',
      SpreadsheetApp.BorderStyle.SOLID
    );

  sheet.getRange('C4')
    .setValue('예: 희망도서 신청은 8월 31일까지입니다.');
  sheet.getRange('A4').setValue(false);

  return {
    success: true,
    created: true,
    sheetName: HR_WEB_NOTICE_SHEET_NAME
  };
}


/* ===== Yes24Api.js ===== */
/** YES24 공식 Open API 공통 연결. 키는 ScriptProperties에만 보관합니다. */
function yes24Url_(path, params) {
  return 'https://apis.yes24.com/v1' + path + '?' + Object.keys(params || {}).map(function(k) {
    return encodeURIComponent(k) + '=' + encodeURIComponent(params[k]);
  }).join('&');
}

function yes24Isbn13_(value) {
  var isbn = String(value || '').replace(/[^0-9Xx]/g, '').toUpperCase();
  if (/^\d{9}[\dX]$/.test(isbn)) {
    var base = '978' + isbn.slice(0, 9), sum = 0;
    for (var i = 0; i < 12; i++) sum += Number(base[i]) * (i % 2 ? 3 : 1);
    return base + ((10 - sum % 10) % 10);
  }
  return /^\d{13}$/.test(isbn) ? isbn : '';
}

function yes24Decode_(response) {
  var status = response.getResponseCode(), payload;
  try { payload = JSON.parse(response.getContentText('UTF-8')); }
  catch (ignore) { throw new Error('YES24 응답을 읽지 못했습니다. HTTP ' + status); }
  var errorCode = String(payload.errorCode || '');
  if (status === 404 && /^(SEARCH_001|GOODS_001|GOODS_002|BEST_001)$/.test(errorCode)) {
    return {items: [], totalCount: 0};
  }
  if (status !== 200 || payload.success !== true) {
    var message = errorCode === 'RATE_002' ? 'YES24 일일 호출 한도에 도달했습니다. 내일 다시 실행해 주세요.' :
      errorCode === 'RATE_001' ? 'YES24 요청이 많습니다. 잠시 후 다시 실행해 주세요.' :
      status === 401 ? 'YES24 API 키를 확인해 주세요.' : 'YES24 연결 오류입니다. 잠시 후 다시 실행해 주세요.';
    throw new Error(message + ' (' + (errorCode || status) + ')');
  }
  if (!payload.data || typeof payload.data !== 'object') throw new Error('YES24 응답에 데이터가 없습니다.');
  return payload.data;
}

function yes24Fetch_(url, key) {
  if (!key) throw new Error('YES24 API 키가 없습니다. 「사본 설정 → API 키 설정」에서 입력해 주세요.');
  var response;
  for (var attempt = 0; attempt < 3; attempt++) {
    response = UrlFetchApp.fetch(url, {method: 'get', headers: {'X-Api-Key': key}, muteHttpExceptions: true, followRedirects: false});
    var code = response.getResponseCode(), payload = {};
    try { payload = JSON.parse(response.getContentText('UTF-8')); } catch (ignore) {}
    if ((code === 429 && payload.errorCode !== 'RATE_002') || code >= 500) {
      if (attempt < 2) { Utilities.sleep(1200 * (attempt + 1)); continue; }
    }
    break;
  }
  return response;
}

function yes24Get_(path, params, key) { return yes24Decode_(yes24Fetch_(yes24Url_(path, params), key)); }

function yes24SearchItems_(query, key, category, pageSize, isIsbn) {
  var isbn = isIsbn ? yes24Isbn13_(query) : '';
  if (isIsbn && !isbn) return [];
  var data = isbn ? yes24Get_('/goods/itemDetail', {searchType: 'ISBN13', query: isbn, detail: 'Y'}, key) :
    yes24Get_('/goods/itemList', {query: query, category: category || 'BOOK', sort: 'RELATION', page: 1, pageSize: pageSize || 50, detail: 'Y'}, key);
  if (!Array.isArray(data.items)) throw new Error('YES24 도서 목록을 읽지 못했습니다.');
  return data.items.filter(function(item) { return !isbn || yes24Isbn13_(item.isbn13 || item.isbn10) === isbn; });
}

function yes24RequestBook_(item) {
  var date = HR_cleanYes24Text_(item.publishDate);
  return {
    title: HR_cleanYes24Text_(item.title), author: HR_cleanYes24Text_(item.author),
    publisher: HR_cleanYes24Text_(item.publisher), pubDate: date,
    year: (date.match(/\d{4}/) || [''])[0], price: Number(item.shopPrice || 0),
    category: HR_cleanYes24Text_(item.goodsSortNm || item.goodsType),
    isbn: HR_cleanYes24Text_(item.isbn13 || item.isbn10),
    cover: HR_safeUrl_(item.cover), link: HR_safeUrl_(item.link)
  };
}

function yes24CategoryId_(label, key) {
  if (label === '종합') return '001';
  var cache = CacheService.getScriptCache(), cacheKey = 'YES24_CATEGORY_' + label;
  var cached = cache.get(cacheKey);
  if (cached) return cached;
  var data = yes24Get_('/category/list', {}, key);
  if (!Array.isArray(data.data)) throw new Error('YES24 분류 목록을 읽지 못했습니다.');
  var matches = data.data.filter(function(row) {
    var name = String(row.categoryName || '').trim();
    var path = String(row.categoryFullPath || '');
    return String(row.categoryId || '').indexOf('001') === 0 &&
      (name === label || path.indexOf(label) >= 0);
  }).sort(function(a, b) { return String(a.categoryId).length - String(b.categoryId).length || String(a.categoryId).localeCompare(String(b.categoryId)); });
  if (!matches.length) throw new Error('YES24에서 ' + label + ' 분류를 찾지 못했습니다.');
  cache.put(cacheKey, String(matches[0].categoryId), 21600);
  return String(matches[0].categoryId);
}

/** 소규모 묶음과 재시도로 초당 제한을 준수합니다. 실패는 호출자까지 전달합니다. */
function yes24FetchAll_(requests) {
  var responses = UrlFetchApp.fetchAll(requests.map(function(request) {
    request.followRedirects = false; return request;
  }));
  responses = responses.map(function(response, index) {
    var status = response.getResponseCode(), payload = {};
    try { payload = JSON.parse(response.getContentText('UTF-8')); } catch (ignore) {}
    if ((status === 429 && payload.errorCode !== 'RATE_002') || status >= 500) {
      Utilities.sleep(1200);
      response = yes24Fetch_(requests[index].url, requests[index].headers['X-Api-Key']);
    }
    var data = yes24Decode_(response);
    if (!Array.isArray(data.items)) throw new Error('YES24 도서 목록을 읽지 못했습니다.');
    return {getResponseCode: function() {return 200;}, getContentText: function() {return JSON.stringify({success:true, data:data});}};
  });
  Utilities.sleep(550);
  return responses;
}

/** 시트 데이터를 변경하지 않는 연결 검사. 공개 도서 정보만 로그에 기록합니다. */
function verifyYes24Connection() {
  var key = getConfiguredYes24Key_();
  var books = HR_searchRequestBooks('아몬드');
  if (!books.length) throw new Error('제목 검색 결과가 없습니다.');
  var isbn = books.filter(function(b) {return /^\d{13}$/.test(b.isbn);})[0].isbn;
  var exact = HR_searchRequestBooks(isbn);
  if (!exact.length || exact[0].isbn !== isbn) throw new Error('ISBN 조회 검증 실패');
  var summary = {titleSearch: books.length, isbnLookup: true, categories: {}};
  Object.keys(HR_BESTSELLER_CATEGORIES).forEach(function(name) {
    var category = HR_BESTSELLER_CATEGORIES[name];
    var rows = HR_fetchBestsellerCategory_(key, category);
    summary.categories[category.label] = rows.length;
  });
  Logger.log(JSON.stringify(summary));
  return summary;
}


  var __fns = {
    "HR_showApplyLinksDialog": HR_showApplyLinksDialog,
    "HR_audienceKey_": HR_audienceKey_,
    "HR_audienceLabel_": HR_audienceLabel_,
    "HR_audienceFromParameter_": HR_audienceFromParameter_,
    "HR_findAudienceColumn_": HR_findAudienceColumn_,
    "HR_ensureAudienceColumn_": HR_ensureAudienceColumn_,
    "HR_parseParentApplicant_": HR_parseParentApplicant_,
    "HR_applicantKey_": HR_applicantKey_,
    "HR_getRequestLimit_": HR_getRequestLimit_,
    "HR_countApplicantRequests_": HR_countApplicantRequests_,
    "HR_requestLimitMessage_": HR_requestLimitMessage_,
    "HR_ensureNoticeAudienceColumn_": HR_ensureNoticeAudienceColumn_,
    "HR_showRequestLimitDialog": HR_showRequestLimitDialog,
    "HR_prepareAudienceColumns_": HR_prepareAudienceColumns_,
    "getConfiguredYes24Key_": getConfiguredYes24Key_,
    "getConfiguredNlkKey": getConfiguredNlkKey,
    "onOpen": onOpen,
    "startNewRequestRound": startNewRequestRound,
    "buildRequestArchiveSheetName_": buildRequestArchiveSheetName_,
    "styleRequestListSheet": styleRequestListSheet,
    "buildBookstoreExportSheet": buildBookstoreExportSheet,
    "onEdit": onEdit,
    "bookstoreStyleExportSheet_": bookstoreStyleExportSheet_,
    "bookstoreFindColumn_": bookstoreFindColumn_,
    "bookstoreNormalize_": bookstoreNormalize_,
    "installEditTrigger": installEditTrigger,
    "onEditTriggerAction": onEditTriggerAction,
    "processRow": processRow,
    "shouldSkipStatus": shouldSkipStatus,
    "smartSearch": smartSearch,
    "fetchYes24Api": fetchYes24Api,
    "selectBestYes24Item": selectBestYes24Item,
    "fetchNlkCallNumber": fetchNlkCallNumber,
    "parseNlkClassNumber": parseNlkClassNumber,
    "findFirstXmlValue": findFirstXmlValue,
    "normalizeClassNumber": normalizeClassNumber,
    "findDuplicateRequest": findDuplicateRequest,
    "findLibraryDuplicate": findLibraryDuplicate,
    "fillFinalData": fillFinalData,
    "checkExistingBooks": checkExistingBooks,
    "reprocessSelectedRows": reprocessSelectedRows,
    "getHeaders": getHeaders,
    "getColumnMap": getColumnMap,
    "getLibraryColumnMap": getLibraryColumnMap,
    "setStatus": setStatus,
    "parseTitleAndPublisher": parseTitleAndPublisher,
    "extractBestIsbn": extractBestIsbn,
    "normalizeIsbn": normalizeIsbn,
    "normalizeText": normalizeText,
    "normalizeHeader": normalizeHeader,
    "stripHtml": stripHtml,
    "extractYear": extractYear,
    "toNumber": toNumber,
    "isTitleHeader": isTitleHeader,
    "isIsbnHeader": isIsbnHeader,
    "isStatusHeader": isStatusHeader,
    "isCoverHeader": isCoverHeader,
    "isPublisherHeader": isPublisherHeader,
    "isAuthorHeader": isAuthorHeader,
    "isPriceHeader": isPriceHeader,
    "isYearHeader": isYearHeader,
    "isCategoryHeader": isCategoryHeader,
    "isCallNumberHeader": isCallNumberHeader,
    "createCacheKey": createCacheKey,
    "getCachedJson": getCachedJson,
    "putCachedJson": putCachedJson,
    "FIX_ACTION_LAYOUT_20260820": FIX_ACTION_LAYOUT_20260820,
    "ensureSelectionControls_": ensureSelectionControls_,
    "toggleSelectionColumn_": toggleSelectionColumn_,
    "resetHopeBooksForNewSemester": resetHopeBooksForNewSemester,
    "HR_importFromOldCopy": HR_importFromOldCopy,
    "HR_importPlan_": HR_importPlan_,
    "HR_importRun_": HR_importRun_,
    "HR_importSheet_": HR_importSheet_,
    "doGet": doGet,
    "HR_getWebAppBaseUrl_": HR_getWebAppBaseUrl_,
    "HR_getSpreadsheet_": HR_getSpreadsheet_,
    "HR_getConfiguredPublicWebAppUrl_": HR_getConfiguredPublicWebAppUrl_,
    "HR_renameLegacyTargetSheet": HR_renameLegacyTargetSheet,
    "HR_getYes24Key_": HR_getYes24Key_,
    "HR_getNlkKey_": HR_getNlkKey_,
    "HR_showRequestSidebar": HR_showRequestSidebar,
    "HR_openSidebarOnSpreadsheetOpen": HR_openSidebarOnSpreadsheetOpen,
    "HR_removeLegacyOpenTrigger_": HR_removeLegacyOpenTrigger_,
    "HR_installSidebarOpenTrigger": HR_installSidebarOpenTrigger,
    "HR_submitBookRequest": HR_submitBookRequest,
    "HR_appendRequestBackup_": HR_appendRequestBackup_,
    "HR_updateRequestBackup_": HR_updateRequestBackup_,
    "HR_submitBookRequests": HR_submitBookRequests,
    "HR_searchRequestBooks": HR_searchRequestBooks,
    "HR_cleanYes24Text_": HR_cleanYes24Text_,
    "HR_largeYes24Cover_": HR_largeYes24Cover_,
    "HR_safeUrl_": HR_safeUrl_,
    "HR_findOwnedBook_": HR_findOwnedBook_,
    "HR_findRequestedBook_": HR_findRequestedBook_,
    "HR_findAnyHeaderColumn_": HR_findAnyHeaderColumn_,
    "HR_normalizeIsbn_": HR_normalizeIsbn_,
    "HR_normalizeBookTitle_": HR_normalizeBookTitle_,
    "HR_findHeaderColumn_": HR_findHeaderColumn_,
    "HR_formatGuideSheet": HR_formatGuideSheet,
    "HR_formatGuideSheetAdminFirst_": HR_formatGuideSheetAdminFirst_,
    "HR_drawGuideSteps_": HR_drawGuideSteps_,
    "HR_formatGuideSheetLegacy_": HR_formatGuideSheetLegacy_,
    "HR_addCopySetupMenu": HR_addCopySetupMenu,
    "HR_showYes24KeyDialog": HR_showYes24KeyDialog,
    "HR_showSchoolSetupDialog": HR_showSchoolSetupDialog,
    "HR_showPublicWebAppUrlDialog": HR_showPublicWebAppUrlDialog,
    "HR_savePublicWebAppUrl": HR_savePublicWebAppUrl,
    "HR_updateGuidePublicLink_": HR_updateGuidePublicLink_,
    "HR_showCopySetupDialog_": HR_showCopySetupDialog_,
    "onSelectionChange": onSelectionChange,
    "HR_getCopySetupState": HR_getCopySetupState,
    "HR_saveYes24ApiKey": HR_saveYes24ApiKey,
    "HR_runSchoolInitialSetup": HR_runSchoolInitialSetup,
    "HR_deleteCopySetupSheet": HR_deleteCopySetupSheet,
    "HR_createCopySetupSheet": HR_createCopySetupSheet,
    "HR_refreshCopySetupBeginnerGuide": HR_refreshCopySetupBeginnerGuide,
    "HR_createCopySetupSheet_": HR_createCopySetupSheet_,
    "HR_createCopySetupSheetBeforeTopGuide_": HR_createCopySetupSheetBeforeTopGuide_,
    "HR_writeCopySetupBeginnerDetails_": HR_writeCopySetupBeginnerDetails_,
    "HR_createCopySetupSheetLegacy_": HR_createCopySetupSheetLegacy_,
    "HR_prepareTopSetupButtons_": HR_prepareTopSetupButtons_,
    "HR_styleCopySetupModern_": HR_styleCopySetupModern_,
    "HR_addSetupButtonImages_": HR_addSetupButtonImages_,
    "HR_getSetupSpreadsheet_": HR_getSetupSpreadsheet_,
    "HR_ensureLibrarySheet_": HR_ensureLibrarySheet_,
    "HR_getLibrarySetupStatus_": HR_getLibrarySetupStatus_,
    "HR_installEditTriggerForSpreadsheet_": HR_installEditTriggerForSpreadsheet_,
    "HR_testYes24Key_": HR_testYes24Key_,
    "HR_getSystemCheckState_": HR_getSystemCheckState_,
    "HR_showSystemCheck": HR_showSystemCheck,
    "HR_escapeHtml_": HR_escapeHtml_,
    "HR_getSubmittedRequests": HR_getSubmittedRequests,
    "HR_getPendingCancelledItems_": HR_getPendingCancelledItems_,
    "HR_restoreLegacyCancelledRequest": HR_restoreLegacyCancelledRequest,
    "HR_extractRequestCoverUrl_": HR_extractRequestCoverUrl_,
    "HR_getOrCreateCancelBackup_": HR_getOrCreateCancelBackup_,
    "HR_cancelSubmittedRequest": HR_cancelSubmittedRequest,
    "HR_undoCancelledRequest": HR_undoCancelledRequest,
    "HR_getBestsellers": HR_getBestsellers,
    "HR_getBestsellerCategory_": HR_getBestsellerCategory_,
    "HR_fetchBestsellerCategory_": HR_fetchBestsellerCategory_,
    "HR_readBestsellerCache_": HR_readBestsellerCache_,
    "HR_writeBestsellerCache_": HR_writeBestsellerCache_,
    "HR_parseBestsellerCache_": HR_parseBestsellerCache_,
    "HR_fillMissingPriceAndYear": HR_fillMissingPriceAndYear,
    "HR_breakMergedRangesSafely_": HR_breakMergedRangesSafely_,
    "HR_importBookListFile": HR_importBookListFile,
    "HR_cleanInput_": HR_cleanInput_,
    "HR_setSafeText_": HR_setSafeText_,
    "HR_cleanPrice_": HR_cleanPrice_,
    "HR_cleanCallNumber_": HR_cleanCallNumber_,
    "HR_safeCoverUrl_": HR_safeCoverUrl_,
    "HR_prepareSelectedBook_": HR_prepareSelectedBook_,
    "HR_unverifiedNote_": HR_unverifiedNote_,
    "HR_showHopeRequestDialog": HR_showHopeRequestDialog,
    "setupSeriesFinder": setupSeriesFinder,
    "seriesFinderOnOpen_": seriesFinderOnOpen_,
    "seriesAddMenu_": seriesAddMenu_,
    "seriesInstallOpenTrigger_": seriesInstallOpenTrigger_,
    "installSeriesWorkButton": installSeriesWorkButton,
    "seriesInstallWorkButtonTrigger_": seriesInstallWorkButtonTrigger_,
    "seriesWorkButtonOnEdit_": seriesWorkButtonOnEdit_,
    "seriesAddWorkButton_": seriesAddWorkButton_,
    "setSeriesYes24Key": setSeriesYes24Key,
    "refreshMissingSeriesBooks": refreshMissingSeriesBooks,
    "seriesGetYes24Key_": seriesGetYes24Key_,
    "seriesSpreadsheet_": seriesSpreadsheet_,
    "seriesReadHoldings_": seriesReadHoldings_,
    "seriesGroupHoldings_": seriesGroupHoldings_,
    "seriesEnrichGroupsFromIsbn_": seriesEnrichGroupsFromIsbn_,
    "seriesApplyEditionInfo_": seriesApplyEditionInfo_,
    "seriesApiUsageState_": seriesApiUsageState_,
    "seriesApiRemaining_": seriesApiRemaining_,
    "seriesRecordApiUsage_": seriesRecordApiUsage_,
    "seriesBuildLookupUrl_": seriesBuildLookupUrl_,
    "seriesGroupPriority_": seriesGroupPriority_,
    "seriesResumeFingerprint_": seriesResumeFingerprint_,
    "seriesLoadResumeState_": seriesLoadResumeState_,
    "seriesSaveResumeState_": seriesSaveResumeState_,
    "seriesClearResumeState_": seriesClearResumeState_,
    "resetSeriesFinderProgress": resetSeriesFinderProgress,
    "seriesAutoContinueMissingBooks_": seriesAutoContinueMissingBooks_,
    "seriesCancelAutoContinue_": seriesCancelAutoContinue_,
    "seriesScheduleAutoContinue_": seriesScheduleAutoContinue_,
    "seriesLegacyResumeOffset_": seriesLegacyResumeOffset_,
    "seriesReadExistingCandidates_": seriesReadExistingCandidates_,
    "seriesExtractUrlFromFormula_": seriesExtractUrlFromFormula_,
    "seriesMergeCandidates_": seriesMergeCandidates_,
    "seriesSearchYes24_": seriesSearchYes24_,
    "seriesBuildSearchUrl_": seriesBuildSearchUrl_,
    "seriesCandidateFromItem_": seriesCandidateFromItem_,
    "seriesHoldingMatchesCandidate_": seriesHoldingMatchesCandidate_,
    "seriesCandidateScore_": seriesCandidateScore_,
    "seriesParseTitle_": seriesParseTitle_,
    "seriesReplaceCircledNumerals_": seriesReplaceCircledNumerals_,
    "seriesReplaceRomanNumerals_": seriesReplaceRomanNumerals_,
    "seriesKeysCompatible_": seriesKeysCompatible_,
    "seriesNormalizeKey_": seriesNormalizeKey_,
    "seriesNormalizePublisher_": seriesNormalizePublisher_,
    "seriesPublisherMatches_": seriesPublisherMatches_,
    "seriesExcludedTitle_": seriesExcludedTitle_,
    "seriesCleanIsbn_": seriesCleanIsbn_,
    "seriesCleanText_": seriesCleanText_,
    "seriesCompressNumbers_": seriesCompressNumbers_,
    "refreshSeriesSaleStatuses": refreshSeriesSaleStatuses,
    "seriesBuildResultSheet_": seriesBuildResultSheet_,
    "seriesApplyApiGauge_": seriesApplyApiGauge_,
    "refreshSeriesApiGauge": refreshSeriesApiGauge,
    "seriesStyleResultSheet_": seriesStyleResultSheet_,
    "setupSeriesWorkSheet": setupSeriesWorkSheet,
    "exportSelectedSeriesBooks": exportSelectedSeriesBooks,
    "seriesEnsureWorkSheet_": seriesEnsureWorkSheet_,
    "seriesStyleWorkSheet_": seriesStyleWorkSheet_,
    "seriesEscapeFormula_": seriesEscapeFormula_,
    "unifiedAddMenu_": unifiedAddMenu_,
    "setupUnifiedLibraryWorkflow": setupUnifiedLibraryWorkflow,
    "openUnifiedUploadSheet": openUnifiedUploadSheet,
    "refreshUnifiedPurchaseReview": refreshUnifiedPurchaseReview,
    "unifiedReadRequestItems_": unifiedReadRequestItems_,
    "unifiedReadSeriesItems_": unifiedReadSeriesItems_,
    "unifiedHeaderMap_": unifiedHeaderMap_,
    "unifiedReadReviewManual_": unifiedReadReviewManual_,
    "unifiedEnsureReviewSheet_": unifiedEnsureReviewSheet_,
    "unifiedStyleReviewSheet_": unifiedStyleReviewSheet_,
    "buildUnifiedBookstoreExportSheet": buildUnifiedBookstoreExportSheet,
    "unifiedStyleExportSheet_": unifiedStyleExportSheet_,
    "unifiedEnsureSettingsSheet_": unifiedEnsureSettingsSheet_,
    "unifiedRefreshDashboard_": unifiedRefreshDashboard_,
    "unifiedCountData_": unifiedCountData_,
    "unifiedAuditDataSources": unifiedAuditDataSources,
    "importLegacySeriesSheetsOnce": importLegacySeriesSheetsOnce,
    "unifiedReadExistingSeriesReview_": unifiedReadExistingSeriesReview_,
    "unifiedReadSelectedSeriesItems_": unifiedReadSelectedSeriesItems_,
    "unifiedRefreshPurchaseReview_": unifiedRefreshPurchaseReview_,
    "sendHopeRequestsToReview": sendHopeRequestsToReview,
    "sendSelectedSeriesToReview": sendSelectedSeriesToReview,
    "unifiedEnsureUploadSheet_": unifiedEnsureUploadSheet_,
    "unifiedReplaceUploadRows_": unifiedReplaceUploadRows_,
    "unifiedReadSelectedUploadItems_": unifiedReadSelectedUploadItems_,
    "sendSelectedUploadItemsToReview": sendSelectedUploadItemsToReview,
    "unifiedStyleUploadSheet_": unifiedStyleUploadSheet_,
    "unifiedMigrateSheetNamesAndLayout_": unifiedMigrateSheetNamesAndLayout_,
    "unifiedStyleHoldingsSheet_": unifiedStyleHoldingsSheet_,
    "unifiedStyleCopySetupSheet_": unifiedStyleCopySetupSheet_,
    "createUnifiedGuides_": createUnifiedGuides_,
    "unifiedBuildGuideSheet_": unifiedBuildGuideSheet_,
    "applyUnifiedLibraryDesign": applyUnifiedLibraryDesign,
    "removeLegacyUnifiedSheetsAfterConfirmation": removeLegacyUnifiedSheetsAfterConfirmation,
    "unifiedArrangeCopySetupButtons_": unifiedArrangeCopySetupButtons_,
    "unifiedMoveCopySetupButtonsToTop_": unifiedMoveCopySetupButtonsToTop_,
    "seriesLeadingCore_": seriesLeadingCore_,
    "seriesFamilySearchText_": seriesFamilySearchText_,
    "seriesFamilyKey_": seriesFamilyKey_,
    "seriesFormatKey_": seriesFormatKey_,
    "HR_parseCsv_": HR_parseCsv_,
    "HR_noticeIsOn_": HR_noticeIsOn_,
    "HR_updateNoticeDate_": HR_updateNoticeDate_,
    "HR_noticeHash_": HR_noticeHash_,
    "HR_fetchNotice_": HR_fetchNotice_,
    "HR_escapeNoticeHtml_": HR_escapeNoticeHtml_,
    "HR_showNoticeDialog_": HR_showNoticeDialog_,
    "HR_showUpdateNoticeOnOpen_": HR_showUpdateNoticeOnOpen_,
    "HR_showUpdateNotice": HR_showUpdateNotice,
    "HR_upgradeNoticeDatePickersAndGuide": HR_upgradeNoticeDatePickersAndGuide,
    "HR_applyNoticeDatePickersAndGuide_": HR_applyNoticeDatePickersAndGuide_,
    "HR_getActiveWebNotices": HR_getActiveWebNotices,
    "HR_noticeDate_": HR_noticeDate_,
    "HR_setupWebNoticeSheet": HR_setupWebNoticeSheet,
    "yes24Url_": yes24Url_,
    "yes24Isbn13_": yes24Isbn13_,
    "yes24Decode_": yes24Decode_,
    "yes24Fetch_": yes24Fetch_,
    "yes24Get_": yes24Get_,
    "yes24SearchItems_": yes24SearchItems_,
    "yes24RequestBook_": yes24RequestBook_,
    "yes24CategoryId_": yes24CategoryId_,
    "yes24FetchAll_": yes24FetchAll_,
    "verifyYes24Connection": verifyYes24Connection
  };
  return {
    version: "real-1",
    files: {"ApplyLinks":"<!DOCTYPE html>\n<html lang=\"ko\">\n<head>\n<base target=\"_top\">\n<meta charset=\"utf-8\">\n<meta name=\"viewport\" content=\"width=device-width, initial-scale=1\">\n<style>\n  :root{--forest:#2c5949;--forest-dark:#203f35;--gold:#d3a248;--gold-dark:#ad762b;--cream:#fbf4e5;--ink:#25352f;--sub:#68766f;--line:#dbe4dc;--paper:#fffdf8;--soft:#e8f2ec}\n  *{box-sizing:border-box}\n  body{margin:0;padding:22px 24px 26px;background:#f4f3eb;font-family:Arial,\"Noto Sans KR\",\"Malgun Gothic\",sans-serif;color:var(--ink)}\n  h1{margin:0 0 6px;font-size:30px;color:var(--forest-dark)}\n  .lead{margin:0 0 18px;color:var(--sub);line-height:1.6;font-size:17px}\n  .tabs{display:grid;grid-template-columns:1fr 1fr;gap:12px;margin-bottom:16px}\n  .tab{padding:16px 10px;border:2px solid var(--line);border-radius:18px;background:#fff;font-size:21px;font-weight:800;color:var(--sub);cursor:pointer}\n  .tab small{display:block;font-size:13px;font-weight:500;margin-top:3px}\n  .tab.active.staff{background:linear-gradient(135deg,#3f8068,#1f4f3f);border-color:#1f4f3f;color:#fff}\n  .tab.active.parent{background:linear-gradient(135deg,var(--gold),var(--gold-dark));border-color:var(--gold-dark);color:#fff}\n  .panel{background:var(--paper);border:1px solid var(--line);border-radius:22px;padding:22px 18px;text-align:center;box-shadow:0 8px 24px rgba(32,63,53,.07)}\n  .panel[hidden]{display:none}\n  .panel .why{margin:0 0 12px;color:var(--sub);font-size:16px;line-height:1.5}\n  canvas.qr{width:300px;height:300px;max-width:100%;border:1px solid var(--line);border-radius:14px;background:#fff;display:block;margin:0 auto 18px}\n  .actions{display:grid;gap:10px;max-width:460px;margin:0 auto}\n  .act{padding:16px 12px;border:0;border-radius:15px;font-size:19px;font-weight:800;cursor:pointer;color:#fff;background:var(--forest)}\n  .act.gold{background:var(--gold-dark)}\n  .act.ghost{background:#fff;color:var(--forest-dark);border:2px solid var(--line);font-size:16px;padding:12px}\n  .act:hover{filter:brightness(.95)}\n  details{margin:16px auto 0;max-width:460px;text-align:left}\n  summary{cursor:pointer;color:var(--sub);font-size:14px;padding:6px 0}\n  .linkbox input{width:100%;padding:10px 12px;border:1px solid var(--line);border-radius:10px;font-size:12px;color:var(--sub);background:#fff}\n  .foot{margin-top:16px;text-align:center}\n  .foot button{padding:12px 36px;border:2px solid var(--line);border-radius:14px;background:#fff;font-size:16px;font-weight:700;color:var(--forest-dark);cursor:pointer}\n  .warn{background:#fff6da;border:1px solid #e3c47a;color:#7a5a12;border-radius:16px;padding:18px 20px;line-height:1.8;font-size:17px}\n  .warn b{color:#5b430c;font-size:19px}\n  .steps{margin:12px 0 0;padding-left:22px}\n  .steps li{margin-bottom:6px}\n  .toast{position:fixed;left:50%;bottom:24px;transform:translateX(-50%);background:var(--forest-dark);color:#fff;padding:12px 22px;border-radius:999px;font-size:16px;opacity:0;transition:opacity .2s;pointer-events:none;z-index:20}\n  .toast.show{opacity:1}\n  .preview{position:fixed;inset:0;background:rgba(32,63,53,.93);display:none;align-items:center;justify-content:center;padding:14px;z-index:10}\n  .preview.show{display:flex}\n  .preview .box{background:#fff;border-radius:20px;padding:18px;max-width:560px;width:100%;max-height:96vh;overflow:auto;text-align:center}\n  .preview h2{margin:0 0 6px;color:var(--forest-dark);font-size:22px}\n  .preview .tip{margin:0 0 12px;color:var(--sub);font-size:15px;line-height:1.5}\n  .preview img{width:100%;max-width:400px;border:1px solid var(--line);border-radius:10px}\n  .preview .actions{margin-top:14px}\n</style>\n</head>\n<body>\n  <h1>🔗 신청 링크 받기</h1>\n  <p class=\"lead\">누구에게 보낼 링크인지 골라서, <b>복사</b>하거나 <b>QR 그림</b>으로 저장하세요.</p>\n\n  <div id=\"needDeploy\" class=\"warn\" hidden>\n    <b>먼저 웹 앱을 한 번 배포해야 링크가 만들어져요.</b>\n    <ol class=\"steps\">\n      <li>위쪽 「확장 프로그램」 → 「Apps Script」를 누르세요.</li>\n      <li>새 화면 오른쪽 위 「배포」 → 「새 배포」 → 톱니바퀴(유형 선택) → 「웹 앱」을 고르세요.</li>\n      <li>실행 사용자는 「나」, 액세스 권한은 「모든 사용자」로 하고 「배포」를 누르세요.</li>\n      <li>끝나면 이 창을 닫고, 「⚙️ 사본 설정 → 🔗 신청 링크 받기」를 다시 누르세요.</li>\n    </ol>\n    <p style=\"margin:12px 0 0;font-size:15px\">이미 배포했는데도 이 안내가 나오면, 「⚙️ 사본 설정 → 🔗 공개 링크 등록」에 /exec 주소를 붙여 넣은 뒤 다시 눌러 보세요.</p>\n  </div>\n\n  <div id=\"main\" hidden>\n    <div class=\"tabs\" role=\"tablist\">\n      <button type=\"button\" class=\"tab parent active\" data-audience=\"parent\" role=\"tab\">👨‍👩‍👧 학부모용<small>가정통신문·알림장</small></button>\n      <button type=\"button\" class=\"tab staff\" data-audience=\"staff\" role=\"tab\">🧑‍🏫 교직원용<small>선생님께 안내</small></button>\n    </div>\n\n    <section class=\"panel\" data-audience=\"parent\" role=\"tabpanel\">\n      <p class=\"why\">학부모님께 보내는 링크예요. 이 링크로 들어가면 <b>교직원 화면은 보이지 않아요.</b></p>\n      <canvas class=\"qr\" width=\"600\" height=\"600\"></canvas>\n      <div class=\"actions\">\n        <button type=\"button\" class=\"act\" data-act=\"copy\">📋 링크 복사</button>\n        <button type=\"button\" class=\"act\" data-act=\"message\">💬 안내 문구까지 복사 (문자·알림장용)</button>\n        <button type=\"button\" class=\"act gold\" data-act=\"save\">⬇️ QR 그림 저장 (인쇄용)</button>\n        <button type=\"button\" class=\"act ghost\" data-act=\"open\">🔎 링크 열어 보기</button>\n      </div>\n      <details><summary>링크 전체 보기</summary><div class=\"linkbox\"><input type=\"text\" readonly></div></details>\n    </section>\n\n    <section class=\"panel\" data-audience=\"staff\" role=\"tabpanel\" hidden>\n      <p class=\"why\">선생님·교직원께 보내는 링크예요.</p>\n      <canvas class=\"qr\" width=\"600\" height=\"600\"></canvas>\n      <div class=\"actions\">\n        <button type=\"button\" class=\"act\" data-act=\"copy\">📋 링크 복사</button>\n        <button type=\"button\" class=\"act\" data-act=\"message\">💬 안내 문구까지 복사 (문자·알림장용)</button>\n        <button type=\"button\" class=\"act gold\" data-act=\"save\">⬇️ QR 그림 저장 (인쇄용)</button>\n        <button type=\"button\" class=\"act ghost\" data-act=\"open\">🔎 링크 열어 보기</button>\n      </div>\n      <details><summary>링크 전체 보기</summary><div class=\"linkbox\"><input type=\"text\" readonly></div></details>\n    </section>\n  </div>\n\n  <div class=\"foot\"><button type=\"button\" onclick=\"google.script.host.close()\">닫기</button></div>\n\n  <div id=\"preview\" class=\"preview\" role=\"dialog\" aria-modal=\"true\">\n    <div class=\"box\">\n      <h2>QR 그림이 만들어졌어요 ✅</h2>\n      <p class=\"tip\">저장이 시작되지 않으면 아래 그림을 <b>오른쪽 클릭(휴대폰은 길게 누르기)</b> → 「이미지를 다른 이름으로 저장」을 누르세요.</p>\n      <img id=\"previewImg\" alt=\"QR 그림\">\n      <div class=\"actions\">\n        <a id=\"downloadLink\" class=\"act\" style=\"text-decoration:none;display:block\" href=\"#\" download>⬇️ 다시 저장하기</a>\n        <button type=\"button\" class=\"act ghost\" id=\"closePreview\">닫기</button>\n      </div>\n    </div>\n  </div>\n  <div id=\"toast\" class=\"toast\"></div>\n\n<script>\n/* QR-BEGIN */\n// 외부 서비스 없이 QR 코드를 만듭니다. (바이트 모드 · 오류정정 M · 버전 1~10, 최대 213바이트)\n// 알고리즘은 QR 코드 표준(ISO/IEC 18004)을 따릅니다.\nvar QrCore = (function () {\n  'use strict';\n\n  // 버전별: [전체 코드워드, 블록 수, 블록당 오류정정 코드워드]  (오류정정 M)\n  var TABLE = {\n    1: [26, 1, 10], 2: [44, 1, 16], 3: [70, 1, 26], 4: [100, 2, 18], 5: [134, 2, 24],\n    6: [172, 4, 16], 7: [196, 4, 18], 8: [242, 4, 22], 9: [292, 5, 22], 10: [346, 5, 26]\n  };\n  var MAX_VERSION = 10;\n\n  // --- 갈루아체 GF(256) ---\n  var EXP = [], LOG = [];\n  (function () {\n    var x = 1, i;\n    for (i = 0; i < 255; i++) { EXP[i] = x; LOG[x] = i; x <<= 1; if (x & 0x100) x ^= 0x11d; }\n    for (i = 255; i < 512; i++) EXP[i] = EXP[i - 255];\n  })();\n  function gfMul(a, b) { return (a === 0 || b === 0) ? 0 : EXP[LOG[a] + LOG[b]]; }\n\n  function rsGenerator(degree) {\n    var poly = [1], i, j, next;\n    for (i = 0; i < degree; i++) {\n      next = [];\n      for (j = 0; j <= poly.length; j++) next[j] = 0;\n      for (j = 0; j < poly.length; j++) {\n        next[j] ^= poly[j];\n        next[j + 1] ^= gfMul(poly[j], EXP[i]);\n      }\n      poly = next;\n    }\n    return poly; // 최고차항 계수 1 포함, 길이 degree+1\n  }\n  function rsRemainder(data, gen) {\n    var degree = gen.length - 1, rem = [], i, j, factor;\n    for (i = 0; i < degree; i++) rem[i] = 0;\n    for (i = 0; i < data.length; i++) {\n      factor = data[i] ^ rem[0];\n      rem.shift(); rem.push(0);\n      for (j = 0; j < degree; j++) rem[j] ^= gfMul(gen[j + 1], factor);\n    }\n    return rem;\n  }\n\n  function utf8Bytes(text) {\n    var out = [], i, c, d;\n    var s = unescape(encodeURIComponent(String(text)));\n    for (i = 0; i < s.length; i++) out.push(s.charCodeAt(i));\n    return out;\n  }\n\n  function numAlign(ver) { return ver === 1 ? 0 : Math.floor(ver / 7) + 2; }\n  function rawModules(ver) {\n    var n = (16 * ver + 128) * ver + 64;\n    if (ver >= 2) { var a = numAlign(ver); n -= (25 * a - 10) * a - 55; if (ver >= 7) n -= 36; }\n    return n;\n  }\n  function alignPositions(ver) {\n    if (ver === 1) return [];\n    var n = numAlign(ver), size = ver * 4 + 17;\n    var step = Math.ceil((ver * 4 + 4) / (n * 2 - 2)) * 2;\n    var res = [6], pos;\n    for (pos = size - 7; res.length < n; pos -= step) res.splice(1, 0, pos);\n    return res;\n  }\n\n  function bit(v, i) { return ((v >>> i) & 1) !== 0; }\n\n  function pickVersion(byteLen) {\n    var ver, t, dataCw, capBits, need;\n    for (ver = 1; ver <= MAX_VERSION; ver++) {\n      t = TABLE[ver];\n      dataCw = t[0] - t[1] * t[2];\n      capBits = dataCw * 8;\n      need = 4 + (ver <= 9 ? 8 : 16) + byteLen * 8;\n      if (need <= capBits) return ver;\n    }\n    throw new Error('주소가 너무 길어서 QR을 만들 수 없습니다.');\n  }\n\n  function encode(text) {\n    var bytes = utf8Bytes(text);\n    var ver = pickVersion(bytes.length);\n    var t = TABLE[ver], raw = Math.floor(rawModules(ver) / 8);\n    var numBlocks = t[1], eccLen = t[2], dataCw = raw - numBlocks * eccLen;\n    var size = ver * 4 + 17;\n    var i, j, k;\n\n    // 1) 데이터 비트열\n    var bits = [];\n    function push(val, len) { for (var b = len - 1; b >= 0; b--) bits.push((val >>> b) & 1); }\n    push(4, 4);                               // 바이트 모드\n    push(bytes.length, ver <= 9 ? 8 : 16);    // 글자 수\n    for (i = 0; i < bytes.length; i++) push(bytes[i], 8);\n    var capBits = dataCw * 8;\n    for (i = 0; i < 4 && bits.length < capBits; i++) bits.push(0); // 종료 문자\n    while (bits.length % 8 !== 0) bits.push(0);\n    var data = [];\n    for (i = 0; i < bits.length; i += 8) {\n      var v = 0; for (j = 0; j < 8; j++) v = (v << 1) | bits[i + j]; data.push(v);\n    }\n    for (k = 0; data.length < dataCw; k++) data.push(k % 2 === 0 ? 0xEC : 0x11);\n\n    // 2) 블록 나누기 + 오류정정 + 끼워 넣기\n    var shortLen = Math.floor(raw / numBlocks), numShort = numBlocks - raw % numBlocks;\n    var gen = rsGenerator(eccLen), blocks = [], offset = 0;\n    for (i = 0; i < numBlocks; i++) {\n      var len = shortLen - eccLen + (i < numShort ? 0 : 1);\n      var dat = data.slice(offset, offset + len); offset += len;\n      var ecc = rsRemainder(dat, gen);\n      if (i < numShort) dat.push(0);          // 짧은 블록은 자리 맞춤용 빈칸\n      blocks.push(dat.concat(ecc));\n    }\n    var all = [];\n    for (i = 0; i < blocks[0].length; i++) {\n      for (j = 0; j < blocks.length; j++) {\n        if (i !== shortLen - eccLen || j >= numShort) all.push(blocks[j][i]);\n      }\n    }\n\n    // 3) 모듈 배치\n    var modules = [], isFn = [];\n    for (i = 0; i < size; i++) { modules.push([]); isFn.push([]); for (j = 0; j < size; j++) { modules[i].push(false); isFn[i].push(false); } }\n    function setFn(x, y, dark) { modules[y][x] = !!dark; isFn[y][x] = true; }\n\n    for (i = 0; i < size; i++) { setFn(6, i, i % 2 === 0); setFn(i, 6, i % 2 === 0); }\n    function finder(cx, cy) {\n      for (var dy = -4; dy <= 4; dy++) for (var dx = -4; dx <= 4; dx++) {\n        var dist = Math.max(Math.abs(dx), Math.abs(dy)), xx = cx + dx, yy = cy + dy;\n        if (xx >= 0 && xx < size && yy >= 0 && yy < size) setFn(xx, yy, dist !== 2 && dist !== 4);\n      }\n    }\n    finder(3, 3); finder(size - 4, 3); finder(3, size - 4);\n    var ap = alignPositions(ver), n = ap.length;\n    for (i = 0; i < n; i++) for (j = 0; j < n; j++) {\n      if ((i === 0 && j === 0) || (i === 0 && j === n - 1) || (i === n - 1 && j === 0)) continue;\n      for (var dy = -2; dy <= 2; dy++) for (var dx = -2; dx <= 2; dx++)\n        setFn(ap[i] + dx, ap[j] + dy, Math.max(Math.abs(dx), Math.abs(dy)) !== 1);\n    }\n    function drawFormat(mask) {\n      var d = (0 << 3) | mask;                // 오류정정 M = 00\n      var rem = d;\n      for (var q = 0; q < 10; q++) rem = (rem << 1) ^ ((rem >>> 9) * 0x537);\n      var fb = ((d << 10) | rem) ^ 0x5412;\n      for (var a = 0; a <= 5; a++) setFn(8, a, bit(fb, a));\n      setFn(8, 7, bit(fb, 6)); setFn(8, 8, bit(fb, 7)); setFn(7, 8, bit(fb, 8));\n      for (var b = 9; b < 15; b++) setFn(14 - b, 8, bit(fb, b));\n      for (var c = 0; c < 8; c++) setFn(size - 1 - c, 8, bit(fb, c));\n      for (var e = 8; e < 15; e++) setFn(8, size - 15 + e, bit(fb, e));\n      setFn(8, size - 8, true);\n    }\n    drawFormat(0);\n    if (ver >= 7) {\n      var r = ver;\n      for (i = 0; i < 12; i++) r = (r << 1) ^ ((r >>> 11) * 0x1F25);\n      var vb = (ver << 12) | r;\n      for (i = 0; i < 18; i++) {\n        var col = bit(vb, i), a2 = size - 11 + (i % 3), b2 = Math.floor(i / 3);\n        setFn(a2, b2, col); setFn(b2, a2, col);\n      }\n    }\n\n    // 데이터 비트 지그재그 배치\n    var idx = 0;\n    for (var right = size - 1; right >= 1; right -= 2) {\n      if (right === 6) right = 5;\n      for (var vert = 0; vert < size; vert++) {\n        for (j = 0; j < 2; j++) {\n          var x = right - j, upward = ((right + 1) & 2) === 0, y = upward ? size - 1 - vert : vert;\n          if (!isFn[y][x] && idx < all.length * 8) { modules[y][x] = bit(all[idx >>> 3], 7 - (idx & 7)); idx++; }\n        }\n      }\n    }\n\n    // 4) 마스크 8가지 중 벌점이 가장 낮은 것 고르기\n    function applyMask(mask) {\n      for (var y = 0; y < size; y++) for (var x = 0; x < size; x++) {\n        var inv;\n        switch (mask) {\n          case 0: inv = (x + y) % 2 === 0; break;\n          case 1: inv = y % 2 === 0; break;\n          case 2: inv = x % 3 === 0; break;\n          case 3: inv = (x + y) % 3 === 0; break;\n          case 4: inv = (Math.floor(x / 3) + Math.floor(y / 2)) % 2 === 0; break;\n          case 5: inv = x * y % 2 + x * y % 3 === 0; break;\n          case 6: inv = (x * y % 2 + x * y % 3) % 2 === 0; break;\n          default: inv = ((x + y) % 2 + x * y % 3) % 2 === 0;\n        }\n        if (!isFn[y][x] && inv) modules[y][x] = !modules[y][x];\n      }\n    }\n    function penalty() {\n      var score = 0, y, x, run, c, dark = 0;\n      for (y = 0; y < size; y++) {\n        run = 1;\n        for (x = 1; x < size; x++) {\n          if (modules[y][x] === modules[y][x - 1]) { run++; if (run === 5) score += 3; else if (run > 5) score++; } else run = 1;\n        }\n      }\n      for (x = 0; x < size; x++) {\n        run = 1;\n        for (y = 1; y < size; y++) {\n          if (modules[y][x] === modules[y - 1][x]) { run++; if (run === 5) score += 3; else if (run > 5) score++; } else run = 1;\n        }\n      }\n      for (y = 0; y < size - 1; y++) for (x = 0; x < size - 1; x++) {\n        c = modules[y][x];\n        if (c === modules[y][x + 1] && c === modules[y + 1][x] && c === modules[y + 1][x + 1]) score += 3;\n      }\n      var pat = [true, false, true, true, true, false, true, false, false, false, false];\n      function matches(get) {\n        var k2, ok1 = true, ok2 = true;\n        for (k2 = 0; k2 < 11; k2++) { if (get(k2) !== pat[k2]) ok1 = false; if (get(k2) !== pat[10 - k2]) ok2 = false; }\n        return ok1 || ok2;\n      }\n      for (y = 0; y < size; y++) for (x = 0; x <= size - 11; x++) { if (matches(function (k2) { return modules[y][x + k2]; })) score += 40; }\n      for (x = 0; x < size; x++) for (y = 0; y <= size - 11; y++) { if (matches(function (k2) { return modules[y + k2][x]; })) score += 40; }\n      for (y = 0; y < size; y++) for (x = 0; x < size; x++) if (modules[y][x]) dark++;\n      score += Math.floor(Math.abs(dark * 20 - size * size * 10) / (size * size)) * 10;\n      return score;\n    }\n    var bestMask = 0, bestScore = Infinity, m;\n    for (m = 0; m < 8; m++) {\n      applyMask(m); drawFormat(m);\n      var sc = penalty();\n      if (sc < bestScore) { bestScore = sc; bestMask = m; }\n      applyMask(m); // 되돌리기\n    }\n    applyMask(bestMask); drawFormat(bestMask);\n\n    return { version: ver, size: size, modules: modules };\n  }\n\n  return { encode: encode };\n})();\n/* QR-END */\n\n(function () {\n  var DATA = <?!= dataJson ?>;\n  var LABEL = { staff: '교직원용', parent: '학부모용' };\n  var COLOR = { staff: '#203f35', parent: '#ad762b' };\n\n  function el(id) { return document.getElementById(id); }\n  function toast(text) {\n    var t = el('toast'); t.textContent = text; t.classList.add('show');\n    setTimeout(function () { t.classList.remove('show'); }, 2000);\n  }\n  function linkFor(audience) {\n    var base = DATA.baseUrl;\n    return base + (base.indexOf('?') < 0 ? '?' : '&') + 'type=' + audience;\n  }\n\n  // QR 을 캔버스에 그립니다. (바깥 여백 4칸)\n  function drawQr(canvas, text, size) {\n    var qr = QrCore.encode(text), quiet = 4, n = qr.size;\n    var cell = Math.floor(size / (n + quiet * 2));\n    var total = cell * (n + quiet * 2);\n    var ctx = canvas.getContext('2d');\n    ctx.fillStyle = '#fff'; ctx.fillRect(0, 0, size, size);\n    var offset = Math.floor((size - total) / 2) + quiet * cell;\n    ctx.fillStyle = '#000';\n    for (var y = 0; y < n; y++) for (var x = 0; x < n; x++) {\n      if (qr.modules[y][x]) ctx.fillRect(offset + x * cell, offset + y * cell, cell, cell);\n    }\n  }\n\n  function wrapText(ctx, text, maxWidth) {\n    var lines = [], line = '';\n    for (var i = 0; i < text.length; i++) {\n      var test = line + text.charAt(i);\n      if (ctx.measureText(test).width > maxWidth && line) { lines.push(line); line = text.charAt(i); } else { line = test; }\n    }\n    if (line) lines.push(line);\n    return lines;\n  }\n\n  // 인쇄해서 쓸 수 있는 그림(제목 + QR + 안내 + 주소)을 만듭니다.\n  function buildPoster(audience) {\n    var url = linkFor(audience);\n    var W = 1000, H = 1240, c = document.createElement('canvas');\n    c.width = W; c.height = H;\n    var ctx = c.getContext('2d');\n    ctx.fillStyle = '#fff'; ctx.fillRect(0, 0, W, H);\n    ctx.fillStyle = COLOR[audience]; ctx.fillRect(0, 0, W, 18);\n    ctx.textAlign = 'center'; ctx.textBaseline = 'alphabetic';\n\n    var school = DATA.schoolName ? DATA.schoolName : '';\n    ctx.fillStyle = '#25352f';\n    var y = 120;\n    if (school) { ctx.font = '700 44px \"Noto Sans KR\",\"Malgun Gothic\",Arial,sans-serif'; ctx.fillText(school, W / 2, y); y += 78; }\n    ctx.fillStyle = '#203f35';\n    ctx.font = '900 76px \"Noto Sans KR\",\"Malgun Gothic\",Arial,sans-serif';\n    ctx.fillText('희망도서 신청', W / 2, y);\n\n    y += 40;\n    var badge = LABEL[audience] + ' ' + (audience === 'parent' ? '👨‍👩‍👧' : '🧑‍🏫');\n    ctx.font = '800 42px \"Noto Sans KR\",\"Malgun Gothic\",Arial,sans-serif';\n    var bw = ctx.measureText(badge).width + 70;\n    ctx.fillStyle = COLOR[audience];\n    var bx = (W - bw) / 2, by = y, bh = 70, r = 35;\n    ctx.beginPath(); ctx.moveTo(bx + r, by); ctx.arcTo(bx + bw, by, bx + bw, by + bh, r); ctx.arcTo(bx + bw, by + bh, bx, by + bh, r); ctx.arcTo(bx, by + bh, bx, by, r); ctx.arcTo(bx, by, bx + bw, by, r); ctx.closePath(); ctx.fill();\n    ctx.fillStyle = '#fff'; ctx.fillText(badge, W / 2, by + 50);\n\n    var qrSize = 720, qc = document.createElement('canvas');\n    qc.width = qrSize; qc.height = qrSize;\n    drawQr(qc, url, qrSize);\n    var qy = by + bh + 40;\n    ctx.fillStyle = '#fff'; ctx.strokeStyle = '#dbe4dc'; ctx.lineWidth = 3;\n    ctx.fillRect((W - qrSize) / 2 - 6, qy - 6, qrSize + 12, qrSize + 12);\n    ctx.strokeRect((W - qrSize) / 2 - 6, qy - 6, qrSize + 12, qrSize + 12);\n    ctx.drawImage(qc, (W - qrSize) / 2, qy);\n\n    var ty = qy + qrSize + 70;\n    ctx.fillStyle = '#25352f';\n    ctx.font = '700 38px \"Noto Sans KR\",\"Malgun Gothic\",Arial,sans-serif';\n    ctx.fillText('스마트폰 카메라로 QR을 비춰 주세요', W / 2, ty);\n\n    ctx.fillStyle = '#68766f';\n    ctx.font = '400 21px Arial,sans-serif';\n    var lines = wrapText(ctx, url, W - 160), ly = ty + 52;\n    for (var i = 0; i < lines.length && i < 3; i++) { ctx.fillText(lines[i], W / 2, ly); ly += 28; }\n    return c;\n  }\n\n  function fileName(audience) {\n    var school = (DATA.schoolName || '').replace(/[\\\\\\/:*?\"<>|]/g, '').trim();\n    return (school ? school + '_' : '') + (audience === 'parent' ? '학부모용' : '교직원용') + '_희망도서_QR.png';\n  }\n\n  function savePoster(audience) {\n    var canvas = buildPoster(audience), dataUrl = canvas.toDataURL('image/png');\n    var name = fileName(audience);\n    el('previewImg').src = dataUrl;\n    var link = el('downloadLink'); link.href = dataUrl; link.setAttribute('download', name);\n    el('preview').classList.add('show');\n    try { link.click(); toast('저장을 시작했어요'); } catch (e) {}\n  }\n\n\n  function fileName(audience) {\n    var school = (DATA.schoolName || '').replace(/[\\\\\\/:*?\"<>|]/g, '').trim();\n    return (school ? school + '_' : '') + (audience === 'parent' ? '학부모용' : '교직원용') + '_희망도서_QR.png';\n  }\n\n  function savePoster(audience) {\n    var canvas = buildPoster(audience), dataUrl = canvas.toDataURL('image/png');\n    el('previewImg').src = dataUrl;\n    var link = el('downloadLink'); link.href = dataUrl; link.setAttribute('download', fileName(audience));\n    el('preview').classList.add('show');\n    try { link.click(); toast('저장을 시작했어요'); } catch (e) {}\n  }\n\n  function copyValue(text, okMessage, inputForFallback) {\n    function fallback() {\n      var ta = inputForFallback || document.createElement('textarea');\n      if (!inputForFallback) { ta.value = text; ta.style.position = 'fixed'; ta.style.opacity = '0'; document.body.appendChild(ta); }\n      ta.focus(); ta.select();\n      var ok = false;\n      try { ok = document.execCommand('copy'); } catch (e) {}\n      if (!inputForFallback) document.body.removeChild(ta);\n      toast(ok ? okMessage : '복사가 안 돼요. 「링크 전체 보기」에서 길게 눌러 복사해 주세요.');\n    }\n    if (navigator.clipboard && navigator.clipboard.writeText) {\n      navigator.clipboard.writeText(text).then(function () { toast(okMessage); }, fallback);\n    } else {\n      fallback();\n    }\n  }\n\n  function messageFor(audience) {\n    var school = DATA.schoolName ? DATA.schoolName + ' ' : '';\n    return school + '희망도서 신청 안내\\n' +\n      (audience === 'parent' ? '읽고 싶은 책을 아래 링크에서 신청해 주세요.\\n' : '읽고 싶은 책을 아래 링크에서 신청해 주세요.\\n') +\n      linkFor(audience);\n  }\n\n  if (!DATA.baseUrl) {\n    el('needDeploy').hidden = false;\n  } else {\n    el('main').hidden = false;\n    var panels = {};\n    Array.prototype.forEach.call(document.querySelectorAll('.panel'), function (panel) {\n      var audience = panel.getAttribute('data-audience'), input = panel.querySelector('input');\n      panels[audience] = panel;\n      input.value = linkFor(audience);\n      try { drawQr(panel.querySelector('canvas'), input.value, 600); }\n      catch (e) { panel.querySelector('canvas').insertAdjacentHTML('afterend', '<p class=\"why\">QR을 만들지 못했어요: ' + String(e.message || e) + '</p>'); }\n      panel.addEventListener('click', function (event) {\n        var act = event.target && event.target.getAttribute && event.target.getAttribute('data-act');\n        if (act === 'copy') copyValue(input.value, '링크를 복사했어요 ✅', input);\n        else if (act === 'message') copyValue(messageFor(audience), '안내 문구를 복사했어요 ✅');\n        else if (act === 'open') window.open(input.value, '_blank');\n        else if (act === 'save') savePoster(audience);\n      });\n    });\n    Array.prototype.forEach.call(document.querySelectorAll('.tab'), function (tab) {\n      tab.addEventListener('click', function () {\n        var audience = tab.getAttribute('data-audience');\n        Array.prototype.forEach.call(document.querySelectorAll('.tab'), function (t) { t.classList.toggle('active', t === tab); });\n        Object.keys(panels).forEach(function (key) { panels[key].hidden = key !== audience; });\n      });\n    });\n  }\n  el('closePreview').addEventListener('click', function () { el('preview').classList.remove('show'); });\n})();\n</script>\n</body>\n</html>\n","CopySetup":"<!DOCTYPE html>\n<html lang=\"ko\">\n<head>\n  <base target=\"_top\">\n  <meta charset=\"UTF-8\">\n\n  <style>\n    :root {\n      --navy: #243b64;\n      --blue: #5277c7;\n      --pale: #f3f6fc;\n      --line: #d8e1f0;\n      --text: #27344a;\n      --muted: #69758a;\n      --success: #28734d;\n      --danger: #a33c3c;\n      --warning: #9a6418;\n    }\n\n    * {\n      box-sizing: border-box;\n    }\n\n    body {\n      margin: 0;\n      padding: 24px;\n      background: #f6f8fc;\n      color: var(--text);\n      font-family:\n        -apple-system,\n        BlinkMacSystemFont,\n        \"Apple SD Gothic Neo\",\n        \"Noto Sans KR\",\n        sans-serif;\n      font-size: 16px;\n    }\n\n    .header {\n      margin-bottom: 18px;\n    }\n\n    h1 {\n      margin: 0;\n      color: var(--navy);\n      font-size: 25px;\n      letter-spacing: -0.04em;\n    }\n\n    .subtitle {\n      margin: 8px 0 0;\n      color: var(--muted);\n      line-height: 1.6;\n    }\n\n    .card {\n      margin-bottom: 16px;\n      padding: 20px;\n      border: 1px solid var(--line);\n      border-radius: 16px;\n      background: #fff;\n      box-shadow: 0 8px 24px rgba(36, 59, 100, 0.06);\n    }\n\n    label {\n      display: block;\n      margin-bottom: 8px;\n      color: var(--navy);\n      font-size: 16px;\n      font-weight: 800;\n    }\n\n    input {\n      width: 100%;\n      min-height: 50px;\n      padding: 0 14px;\n      border: 1px solid var(--line);\n      border-radius: 11px;\n      outline: none;\n      color: var(--text);\n      font: inherit;\n    }\n\n    input:focus {\n      border-color: var(--blue);\n      box-shadow: 0 0 0 4px rgba(82, 119, 199, 0.12);\n    }\n\n    .help {\n      margin: 9px 0 0;\n      color: var(--muted);\n      font-size: 13px;\n      line-height: 1.55;\n    }\n\n    .api-divider {\n      height: 1px;\n      margin: 20px 0;\n      background: var(--line);\n    }\n\n    .required,\n    .optional {\n      display: inline-flex;\n      margin-left: 7px;\n      padding: 3px 8px;\n      border-radius: 999px;\n      font-size: 12px;\n      vertical-align: 2px;\n    }\n\n    .required {\n      background: #fff0f0;\n      color: var(--danger);\n    }\n\n    .optional {\n      background: #edf4ff;\n      color: #315f9f;\n    }\n\n    .status-list {\n      display: grid;\n      gap: 9px;\n      margin: 0;\n      padding: 0;\n      list-style: none;\n    }\n\n    .status-item {\n      display: flex;\n      align-items: center;\n      gap: 9px;\n      min-height: 42px;\n      padding: 9px 12px;\n      border-radius: 10px;\n      background: var(--pale);\n      font-size: 14px;\n      font-weight: 700;\n    }\n\n    .status-item.ok {\n      background: #edf8f1;\n      color: var(--success);\n    }\n\n    .status-item.need {\n      background: #fff6e8;\n      color: var(--warning);\n    }\n\n    .notice {\n      margin-top: 14px;\n      padding: 13px 14px;\n      border: 1px solid #edc5c5;\n      border-radius: 11px;\n      background: #fff1f1;\n      color: var(--danger);\n      font-size: 14px;\n      font-weight: 800;\n      line-height: 1.6;\n    }\n\n    .message {\n      display: none;\n      margin: 14px 0;\n      padding: 12px 14px;\n      border-radius: 10px;\n      font-size: 14px;\n      line-height: 1.55;\n    }\n\n    .message.show {\n      display: block;\n    }\n\n    .message.success {\n      background: #edf8f1;\n      color: var(--success);\n    }\n\n    .message.error {\n      background: #fff1f1;\n      color: var(--danger);\n    }\n\n    .buttons {\n      display: flex;\n      flex-wrap: wrap;\n      justify-content: flex-end;\n      gap: 9px;\n      margin-top: 18px;\n    }\n\n    button {\n      min-height: 46px;\n      padding: 0 18px;\n      border-radius: 10px;\n      font: inherit;\n      font-weight: 800;\n      cursor: pointer;\n    }\n\n    button:disabled {\n      cursor: wait;\n      opacity: 0.55;\n    }\n\n    .primary {\n      border: 0;\n      background: var(--blue);\n      color: #fff;\n    }\n\n    .secondary {\n      border: 1px solid var(--line);\n      background: #fff;\n      color: var(--navy);\n    }\n\n    .danger {\n      border: 1px solid #efc7c7;\n      background: #fff5f5;\n      color: var(--danger);\n    }\n\n    .hidden {\n      display: none;\n    }\n    .contact-divider {\n      height: 1px;\n      margin: 22px 0;\n      background: var(--line);\n    }\n\n    .contact-grid {\n      display: grid;\n      grid-template-columns: 1fr 1fr;\n      gap: 14px;\n    }\n\n    .contact-grid label {\n      display: block;\n      margin-bottom: 8px;\n    }\n\n    .contact-help {\n      margin-top: 12px;\n      color: #667085;\n      font-size: 13px;\n      line-height: 1.6;\n    }\n\n    @media (max-width: 560px) {\n      .contact-grid {\n        grid-template-columns: 1fr;\n      }\n    }\n  </style>\n</head>\n\n<body>\n  <header class=\"header\">\n    <h1 id=\"pageTitle\"></h1>\n    <p id=\"pageSubtitle\" class=\"subtitle\"></p>\n  </header>\n\n  <section id=\"apiPanel\" class=\"card hidden\">\n    <label for=\"yes24Key\">\n      YES24 API Key\n      <span class=\"required\">필수</span>\n    </label>\n    <input\n      id=\"yes24Key\"\n      type=\"text\"\n      autocomplete=\"off\"\n      placeholder=\"발급받은 YES24 API Key\"\n    >\n    <p class=\"help\">\n      저장할 때 실제 도서 검색을 한 번 실행해 키가 정상인지 확인합니다.\n    </p>\n\n    <div class=\"api-divider\"></div>\n\n    <label for=\"nlkKey\">\n      국립중앙도서관 API Key\n      <span class=\"optional\">선택</span>\n    </label>\n    <input\n      id=\"nlkKey\"\n      type=\"text\"\n      autocomplete=\"off\"\n      placeholder=\"없으면 비워 두세요\"\n    >\n    <p class=\"help\">\n      KDC 분류번호를 참고할 때만 사용합니다.\n      입력하지 않아도 도서 검색과 신청은 정상 작동합니다.\n    </p>\n  </section>\n\n  <section id=\"schoolPanel\" class=\"hidden\">\n    <div class=\"card\">\n      <label for=\"schoolName\">학교명</label>\n      <input\n        id=\"schoolName\"\n        type=\"text\"\n        autocomplete=\"off\"\n        placeholder=\"예: ***초등학교\"\n      >\n      <p id=\"spreadsheetName\" class=\"help\"></p>\n\n      <div class=\"contact-divider\"></div>\n\n      <div class=\"contact-grid\">\n        <div>\n          <label for=\"managerName\">도서관 담당자 이름</label>\n          <input\n            id=\"managerName\"\n            type=\"text\"\n            autocomplete=\"off\"\n            placeholder=\"예: 홍길동\"\n          >\n        </div>\n\n        <div>\n          <label for=\"managerExtension\">내선번호</label>\n          <input\n            id=\"managerExtension\"\n            type=\"text\"\n            inputmode=\"tel\"\n            autocomplete=\"off\"\n            placeholder=\"예: 1234\"\n          >\n        </div>\n      </div>\n\n      <p class=\"contact-help\">\n        입력한 담당자 이름과 내선번호는 공개 희망도서 신청 페이지 맨 아래에 표시됩니다.\n        필요하지 않으면 비워 두어도 됩니다.\n      </p>\n    </div>\n\n    <div class=\"card\">\n      <label>설정 점검</label>\n      <ul id=\"statusList\" class=\"status-list\"></ul>\n\n      <div class=\"notice\">\n        「우리학교 소장도서」 시트에는 책 제목과 ISBN13을 반드시 입력해야 합니다.\n        ISBN13이 없으면 “이미 소장중인 도서” 확인이 정확히 작동하지 않습니다.\n      </div>\n    </div>\n\n    <div class=\"card\">\n      <label>공개 신청 링크</label>\n      <p id=\"webAppStatus\" class=\"help\"></p>\n      <p class=\"help\">\n        웹 앱을 새로 배포한 뒤 이 창에서 학교 최초 설정을 다시 실행하면\n        사용설명서의 신청 버튼이 새 공개 링크로 연결됩니다.\n      </p>\n    </div>\n  </section>\n\n  <div id=\"message\" class=\"message\"></div>\n\n  <div class=\"buttons\">\n    <button\n      class=\"secondary\"\n      type=\"button\"\n      onclick=\"google.script.host.close()\"\n    >\n      닫기\n    </button>\n\n    <button\n      id=\"saveButton\"\n      class=\"primary\"\n      type=\"button\"\n      onclick=\"saveCurrentMode()\"\n    ></button>\n  </div>\n\n  <script>\n    const MODE = '<?= mode ?>';\n\n    const pageTitle =\n      document.getElementById('pageTitle');\n    const pageSubtitle =\n      document.getElementById('pageSubtitle');\n    const apiPanel =\n      document.getElementById('apiPanel');\n    const schoolPanel =\n      document.getElementById('schoolPanel');\n    const yes24Key =\n      document.getElementById('yes24Key');\n    const nlkKey =\n      document.getElementById('nlkKey');\n    const schoolName =\n      document.getElementById('schoolName');\n    const managerName =\n      document.getElementById('managerName');\n    const managerExtension =\n      document.getElementById('managerExtension');\n    const spreadsheetName =\n      document.getElementById('spreadsheetName');\n    const statusList =\n      document.getElementById('statusList');\n    const webAppStatus =\n      document.getElementById('webAppStatus');\n    const message =\n      document.getElementById('message');\n    const saveButton =\n      document.getElementById('saveButton');\n\n    window.addEventListener('load', function() {\n      if (MODE === 'api') {\n        pageTitle.textContent = '🔑 도서 API 키 설정';\n        pageSubtitle.textContent =\n          'YES24은 꼭 입력하고, KDC 참고용 키는 필요할 때만 입력하세요.';\n        apiPanel.classList.remove('hidden');\n        saveButton.textContent = '저장하고 YES24 연결 확인';\n      } else {\n        pageTitle.textContent = '🏫 학교 최초 설정';\n        pageSubtitle.textContent =\n          '학교명과 담당자 연락처를 저장하고 신청서와 소장도서 목록에 연결합니다.';\n        schoolPanel.classList.remove('hidden');\n        saveButton.textContent = '이 학교 최초 설정 실행';\n      }\n\n      loadState();\n    });\n\n    function loadState() {\n      setBusy(true);\n\n      google.script.run\n        .withSuccessHandler(function(state) {\n          setBusy(false);\n          yes24Key.value = state.yes24Key || '';\n          nlkKey.value = state.nlkKey || '';\n          schoolName.value = state.schoolName || '';\n          managerName.value = state.managerName || '';\n          managerExtension.value = state.managerExtension || '';\n          spreadsheetName.textContent =\n            '현재 사본: ' + (state.spreadsheetName || '');\n          renderStatus(state);\n        })\n        .withFailureHandler(function(error) {\n          setBusy(false);\n          showMessage(getErrorMessage(error), 'error');\n        })\n        .HR_getCopySetupState();\n    }\n\n    function renderStatus(state) {\n      const items = [\n        {\n          ok: state.yes24Ready,\n          text: 'YES24 API 키 (필수)'\n        },\n        {\n          ok: state.nlkReady,\n          optional: true,\n          text: '국립중앙도서관 API 키 (KDC 참고용·선택)'\n        },\n        {\n          ok: state.targetSheetReady,\n          text: '희망도서 목록 시트'\n        },\n        {\n          ok: state.librarySheetReady,\n          text: '우리학교 소장도서 시트'\n        },\n        {\n          ok: state.libraryHeadersReady,\n          text: '소장도서 책 제목·ISBN13 열'\n        },\n        {\n          ok: state.configuredSpreadsheetId === state.spreadsheetId,\n          text: '이 학교 신청서 연결'\n        }\n      ];\n\n      statusList.innerHTML = '';\n\n      items.forEach(function(item) {\n        const li = document.createElement('li');\n\n        li.className =\n          'status-item ' +\n          (item.ok || item.optional ? 'ok' : 'need');\n        li.textContent =\n          (\n            item.ok\n              ? '✅ '\n              : item.optional\n                ? '➖ '\n                : '⚠️ '\n          ) +\n          item.text +\n          (\n            item.optional && !item.ok\n              ? ' · 없어도 됨'\n              : ''\n          );\n        statusList.appendChild(li);\n      });\n\n      webAppStatus.textContent = state.webAppUrl\n        ? '✅ 공개 신청 링크가 있습니다.'\n        : '⚠️ 아직 공개 신청 링크가 없습니다. 웹 앱 배포가 필요합니다.';\n    }\n\n    function saveCurrentMode() {\n      if (MODE === 'api') {\n        saveApiKey();\n      } else {\n        runSchoolSetup();\n      }\n    }\n\n    function saveApiKey() {\n      const key = yes24Key.value.trim();\n\n      if (!key) {\n        showMessage(\n          'YES24 API Key를 입력해 주세요.',\n          'error'\n        );\n        return;\n      }\n\n      setBusy(true);\n      showMessage(\n        'YES24 검색 연결을 확인하고 있습니다…',\n        'success'\n      );\n\n      google.script.run\n        .withSuccessHandler(function(response) {\n          setBusy(false);\n          showMessage(response.message, 'success');\n        })\n        .withFailureHandler(function(error) {\n          setBusy(false);\n          showMessage(getErrorMessage(error), 'error');\n        })\n        .HR_saveYes24ApiKey({\n          yes24Key: key,\n          nlkKey: nlkKey.value.trim()\n        });\n    }\n\n    function runSchoolSetup() {\n      setBusy(true);\n      showMessage(\n        '현재 학교 사본을 연결하고 점검하고 있습니다…',\n        'success'\n      );\n\n      google.script.run\n        .withSuccessHandler(function(response) {\n          setBusy(false);\n          showMessage(response.message, 'success');\n          loadState();\n        })\n        .withFailureHandler(function(error) {\n          setBusy(false);\n          showMessage(getErrorMessage(error), 'error');\n        })\n        .HR_runSchoolInitialSetup({\n          schoolName: schoolName.value.trim(),\n          managerName: managerName.value.trim(),\n          managerExtension: managerExtension.value.trim()\n        });\n    }\n\n    function setBusy(isBusy) {\n      saveButton.disabled = isBusy;\n    }\n\n    function showMessage(text, type) {\n      message.textContent = text;\n      message.className =\n        'message show ' + (type || 'success');\n    }\n\n    function getErrorMessage(error) {\n      return error && error.message\n        ? error.message\n        : String(\n          error ||\n          '처리 중 오류가 발생했습니다.'\n        );\n    }\n  </script>\n</body>\n</html>\n","RequestSidebar":"<!DOCTYPE html>\n<html lang=\"ko\">\n<head>\n  <base target=\"_top\">\n  <meta charset=\"UTF-8\">\n\n  <style>\n    :root {\n      --blue: #4f78d1;\n      --blue-dark: #3159ac;\n      --blue-soft: #edf3ff;\n      --ink: #24324a;\n      --sub: #68748a;\n      --line: #dce4f2;\n      --success: #237a4b;\n      --success-soft: #eaf7ef;\n      --danger: #b33a3a;\n      --danger-soft: #fff0ef;\n      --paper: #ffffff;\n      --background: #f5f7fb;\n    }\n\n    * {\n      box-sizing: border-box;\n    }\n\n    body {\n      margin: 0;\n      padding: 17px;\n      background: var(--background);\n      color: var(--ink);\n      font-family:\n        -apple-system,\n        BlinkMacSystemFont,\n        \"Apple SD Gothic Neo\",\n        \"Noto Sans KR\",\n        sans-serif;\n    }\n\n    .header {\n      width: 100%;\n      max-width: 760px;\n      margin-left: auto;\n      margin-right: auto;\n      margin-bottom: 16px;\n    }\n\n    .eyebrow {\n      margin: 0 0 7px;\n      color: var(--blue);\n      font-size: 12px;\n      font-weight: 900;\n      letter-spacing: 0.12em;\n    }\n\n    h1 {\n      margin: 0;\n      font-size: 27px;\n      line-height: 1.25;\n      letter-spacing: -0.04em;\n    }\n\n    .description {\n      margin: 10px 0 0;\n      color: #536077;\n      font-size: 14px;\n      line-height: 1.7;\n    }\n\n    .card {\n      width: 100%;\n      max-width: 760px;\n      margin: 0 auto;\n      padding: 16px;\n      border: 1px solid var(--line);\n      border-radius: 16px;\n      background: var(--paper);\n      box-shadow: 0 8px 25px rgba(50, 76, 125, 0.08);\n    }\n\n    .field {\n      margin-bottom: 17px;\n    }\n\n    .field:last-of-type {\n      margin-bottom: 0;\n    }\n\n    label {\n      display: block;\n      margin-bottom: 8px;\n      color: #273b5e;\n      font-size: 15px;\n      font-weight: 850;\n    }\n\n    input {\n      width: 100%;\n      height: 54px;\n      padding: 0 14px;\n      border: 1px solid var(--line);\n      border-radius: 12px;\n      outline: none;\n      background: #fff;\n      color: var(--ink);\n      font-size: 17px;\n      font-weight: 650;\n    }\n\n    input:focus {\n      border-color: var(--blue);\n      box-shadow: 0 0 0 3px rgba(79, 120, 209, 0.12);\n    }\n\n    .help {\n      margin: 8px 1px 0;\n      color: #68748a;\n      font-size: 12px;\n      line-height: 1.6;\n    }\n\n    .submit-button {\n      width: 100%;\n      min-height: 55px;\n      margin-top: 18px;\n      border: 0;\n      border-radius: 12px;\n      background: var(--blue);\n      color: #fff;\n      font-size: 17px;\n      font-weight: 900;\n      cursor: pointer;\n      transition: background 0.15s, transform 0.15s;\n    }\n\n    .submit-button:hover {\n      background: var(--blue-dark);\n      transform: translateY(-1px);\n    }\n\n    .submit-button:disabled {\n      cursor: wait;\n      opacity: 0.58;\n      transform: none;\n    }\n\n    .message {\n      display: none;\n      margin-top: 15px;\n      padding: 14px;\n      border-radius: 11px;\n      font-size: 14px;\n      line-height: 1.65;\n    }\n\n    .message.show {\n      display: block;\n    }\n\n    .message.info {\n      background: var(--blue-soft);\n      color: var(--blue-dark);\n    }\n\n    .message.success {\n      background: var(--success-soft);\n      color: var(--success);\n    }\n\n    .message.error {\n      background: var(--danger-soft);\n      color: var(--danger);\n    }\n\n    .search-results {\n      display: none;\n      margin-top: 15px;\n    }\n\n    .search-results.show {\n      display: block;\n    }\n\n    .search-results-title {\n      margin: 0 0 9px;\n      color: #273b5e;\n      font-size: 15px;\n      font-weight: 900;\n    }\n\n    .book-list {\n      display: grid;\n      gap: 10px;\n    }\n\n    .book-card {\n      display: grid;\n      grid-template-columns: 76px 1fr;\n      gap: 12px;\n      padding: 12px;\n      border: 1px solid var(--line);\n      border-radius: 13px;\n      background: #fff;\n      box-shadow: 0 5px 16px rgba(50, 76, 125, 0.06);\n    }\n\n    .book-cover {\n      display: flex;\n      width: 76px;\n      height: 110px;\n      overflow: hidden;\n      align-items: center;\n      justify-content: center;\n      border-radius: 7px;\n      background: #eef1f6;\n      color: #8c96a8;\n      font-size: 24px;\n    }\n\n    .book-cover img {\n      width: 100%;\n      height: 100%;\n      object-fit: cover;\n    }\n\n    .book-info {\n      display: flex;\n      min-width: 0;\n      flex-direction: column;\n    }\n\n    .book-title {\n      margin: 0;\n      color: var(--ink);\n      font-size: 15px;\n      font-weight: 900;\n      line-height: 1.45;\n    }\n\n    .book-meta {\n      margin: 7px 0 0;\n      color: #5d6980;\n      font-size: 12px;\n      line-height: 1.55;\n    }\n\n    .select-book-button {\n      width: 100%;\n      min-height: 43px;\n      margin-top: auto;\n      padding: 9px 8px;\n      border: 0;\n      border-radius: 9px;\n      background: var(--blue-soft);\n      color: var(--blue-dark);\n      font-size: 14px;\n      font-weight: 900;\n      cursor: pointer;\n    }\n\n    .select-book-button:hover {\n      background: #dce8ff;\n    }\n\n    .select-book-button:disabled {\n      cursor: wait;\n      opacity: 0.55;\n    }\n\n    .select-book-button.done {\n      background: var(--success-soft);\n      color: var(--success);\n      opacity: 1;\n    }\n\n    .book-card.done {\n      border-color: #b9dcc7;\n      background: #fbfffc;\n    }\n\n    .book-card.selected {\n      border: 2px solid var(--blue);\n      background: #f7faff;\n      box-shadow: 0 8px 22px rgba(79, 120, 209, 0.16);\n    }\n\n    .select-book-button.selected {\n      background: var(--blue);\n      color: #fff;\n    }\n\n    .cart-bar {\n      position: sticky;\n      bottom: 12px;\n      z-index: 30;\n      display: none;\n      width: 100%;\n      max-width: 760px;\n      margin: 16px auto 0;\n      padding: 12px;\n      align-items: center;\n      gap: 12px;\n      border: 1px solid #cbd9f5;\n      border-radius: 16px;\n      background: rgba(255, 255, 255, 0.96);\n      box-shadow: 0 12px 36px rgba(37, 61, 105, 0.19);\n      backdrop-filter: blur(10px);\n    }\n\n    .cart-bar.show {\n      display: flex;\n    }\n\n    .cart-count {\n      flex: 0 0 auto;\n      color: var(--blue-dark);\n      font-size: 15px;\n      font-weight: 900;\n      white-space: nowrap;\n    }\n\n    .cart-submit {\n      min-height: 50px;\n      flex: 1;\n      border: 0;\n      border-radius: 12px;\n      background: var(--blue);\n      color: #fff;\n      font-size: 16px;\n      font-weight: 900;\n      cursor: pointer;\n    }\n\n    .cart-submit:disabled {\n      cursor: wait;\n      opacity: 0.55;\n    }\n\n    .cart-toggle {\n      display: flex;\n      min-height: 50px;\n      flex: 0 0 auto;\n      padding: 8px 14px;\n      align-items: center;\n      gap: 10px;\n      border: 0;\n      border-radius: 12px;\n      background: #eef4ff;\n      color: var(--blue-dark);\n      cursor: pointer;\n    }\n\n    .cart-count strong {\n      color: var(--blue);\n      font-size: 19px;\n    }\n\n    .cart-open-hint {\n      color: #637492;\n      font-size: 12px;\n      font-weight: 800;\n    }\n\n    .cart-drawer {\n      position: fixed;\n      inset: 0;\n      z-index: 95;\n      display: none;\n    }\n\n    .cart-drawer.show {\n      display: block;\n    }\n\n    .cart-backdrop {\n      position: absolute;\n      inset: 0;\n      width: 100%;\n      height: 100%;\n      border: 0;\n      background: rgba(25, 36, 56, 0.5);\n      backdrop-filter: blur(4px);\n      cursor: pointer;\n    }\n\n    .cart-panel {\n      position: absolute;\n      top: 0;\n      right: 0;\n      display: flex;\n      width: min(430px, 100%);\n      height: 100%;\n      padding: 24px;\n      flex-direction: column;\n      background: #f8faff;\n      box-shadow: -18px 0 60px rgba(28, 44, 73, 0.24);\n      animation: cart-slide-in 0.22s ease-out;\n    }\n\n    .cart-panel-header {\n      display: flex;\n      align-items: flex-start;\n      justify-content: space-between;\n      gap: 16px;\n    }\n\n    .cart-panel-header p {\n      margin: 0 0 4px;\n      color: #7990b8;\n      font-size: 11px;\n      font-weight: 950;\n      letter-spacing: 0.12em;\n    }\n\n    .cart-panel-header h2 {\n      margin: 0;\n      color: var(--ink);\n      font-size: 25px;\n      letter-spacing: -0.04em;\n    }\n\n    .cart-close {\n      width: 44px;\n      height: 44px;\n      flex: 0 0 auto;\n      border: 0;\n      border-radius: 14px;\n      background: #e7edf9;\n      color: #435675;\n      font-size: 28px;\n      line-height: 1;\n      cursor: pointer;\n    }\n\n    .cart-guide {\n      margin: 14px 0 16px;\n      padding: 11px 13px;\n      border-radius: 12px;\n      background: #edf3ff;\n      color: #526580;\n      font-size: 13px;\n      font-weight: 750;\n      line-height: 1.55;\n    }\n\n    .cart-list {\n      min-height: 0;\n      flex: 1;\n      overflow-y: auto;\n      padding-right: 4px;\n    }\n\n    .cart-empty {\n      display: flex;\n      min-height: 220px;\n      align-items: center;\n      justify-content: center;\n      color: #75829a;\n      font-size: 14px;\n      font-weight: 800;\n      text-align: center;\n    }\n\n    .cart-item {\n      display: grid;\n      grid-template-columns: 58px minmax(0, 1fr) auto;\n      gap: 12px;\n      margin-bottom: 10px;\n      padding: 12px;\n      align-items: center;\n      border: 1px solid #dfe7f5;\n      border-radius: 15px;\n      background: #fff;\n      box-shadow: 0 7px 20px rgba(49, 73, 112, 0.07);\n    }\n\n    .cart-item-cover {\n      display: flex;\n      width: 58px;\n      height: 78px;\n      overflow: hidden;\n      align-items: center;\n      justify-content: center;\n      border-radius: 7px;\n      background: #edf1f7;\n      font-size: 24px;\n    }\n\n    .cart-item-cover img {\n      width: 100%;\n      height: 100%;\n      object-fit: cover;\n    }\n\n    .cart-item-title {\n      margin: 0;\n      color: var(--ink);\n      font-size: 14px;\n      font-weight: 900;\n      line-height: 1.4;\n    }\n\n    .cart-item-meta {\n      margin: 5px 0 0;\n      color: #748098;\n      font-size: 11px;\n      line-height: 1.45;\n    }\n\n    .cart-remove {\n      min-width: 54px;\n      min-height: 38px;\n      border: 1px solid #efc9c2;\n      border-radius: 10px;\n      background: #fff7f5;\n      color: #b34f3f;\n      font-size: 12px;\n      font-weight: 900;\n      cursor: pointer;\n    }\n\n    .cart-panel-actions {\n      display: grid;\n      grid-template-columns: auto minmax(0, 1fr);\n      gap: 10px;\n      padding-top: 14px;\n      border-top: 1px solid #dde5f2;\n    }\n\n    .cart-clear {\n      min-height: 50px;\n      padding: 0 16px;\n      border: 1px solid #ccd7e9;\n      border-radius: 12px;\n      background: #fff;\n      color: #5e6e88;\n      font-size: 13px;\n      font-weight: 900;\n      cursor: pointer;\n    }\n\n    .cart-drawer-submit {\n      width: 100%;\n    }\n\n    @keyframes cart-slide-in {\n      from {\n        opacity: 0.4;\n        transform: translateX(30px);\n      }\n\n      to {\n        opacity: 1;\n        transform: translateX(0);\n      }\n    }\n\n    @media (max-width: 620px) {\n      .cart-bar {\n        bottom: 8px;\n        padding: 9px;\n        gap: 8px;\n      }\n\n      .cart-toggle {\n        min-width: 0;\n        padding: 8px 10px;\n      }\n\n      .cart-open-hint {\n        display: none;\n      }\n\n      .cart-submit {\n        min-height: 48px;\n        font-size: 14px;\n      }\n\n      .cart-panel {\n        top: auto;\n        bottom: 0;\n        width: 100%;\n        height: min(78vh, 680px);\n        padding: 20px 16px 16px;\n        border-radius: 24px 24px 0 0;\n        animation-name: cart-rise-in;\n      }\n\n      .cart-item {\n        grid-template-columns: 52px minmax(0, 1fr) auto;\n        padding: 10px;\n      }\n\n      .cart-item-cover {\n        width: 52px;\n        height: 70px;\n      }\n    }\n\n    @keyframes cart-rise-in {\n      from {\n        opacity: 0.4;\n        transform: translateY(30px);\n      }\n\n      to {\n        opacity: 1;\n        transform: translateY(0);\n      }\n    }\n\n    .footer {\n      width: 100%;\n      max-width: 760px;\n      margin: 16px auto 0;\n      color: #68748a;\n      font-size: 12px;\n      line-height: 1.65;\n      text-align: center;\n    }\n\n    .creator-watermark {\n      position: fixed;\n      bottom: 10px;\n      left: 12px;\n      z-index: 20;\n      padding: 5px 8px;\n      border-radius: 7px;\n      background: rgba(255, 255, 255, 0.78);\n      color: #4b5870;\n      font-size: 11px;\n      font-weight: 750;\n      letter-spacing: -0.02em;\n      opacity: 0.56;\n      pointer-events: none;\n      backdrop-filter: blur(4px);\n    }\n\n    .notice-modal {\n      position: fixed;\n      inset: 0;\n      z-index: 100;\n      display: none;\n      padding: 22px;\n      align-items: center;\n      justify-content: center;\n      background: rgba(27, 38, 58, 0.48);\n      backdrop-filter: blur(5px);\n    }\n\n    .notice-modal.show {\n      display: flex;\n    }\n\n    .notice-dialog {\n      width: min(500px, 100%);\n      overflow: hidden;\n      border: 1px solid rgba(255, 255, 255, 0.72);\n      border-radius: 24px;\n      background: #fff;\n      box-shadow: 0 24px 70px rgba(21, 34, 58, 0.28);\n      animation: notice-in 0.2s ease-out;\n    }\n\n    .notice-visual {\n      display: flex;\n      min-height: 130px;\n      align-items: center;\n      justify-content: center;\n      background:\n        radial-gradient(circle at 30% 20%, #ffffff 0, transparent 32%),\n        linear-gradient(135deg, #eef4ff, #dfe9ff);\n    }\n\n    .notice-modal.owned .notice-visual {\n      background:\n        radial-gradient(circle at 30% 20%, #ffffff 0, transparent 32%),\n        linear-gradient(135deg, #fff3eb, #ffe1d4);\n    }\n\n    .notice-icon {\n      display: flex;\n      width: 76px;\n      height: 76px;\n      align-items: center;\n      justify-content: center;\n      border-radius: 24px;\n      background: #fff;\n      box-shadow: 0 10px 30px rgba(58, 90, 158, 0.16);\n      font-size: 38px;\n    }\n\n    .notice-content {\n      padding: 24px 24px 22px;\n      text-align: center;\n    }\n\n    .notice-title {\n      margin: 0;\n      color: var(--ink);\n      font-size: 22px;\n      font-weight: 950;\n      line-height: 1.4;\n      letter-spacing: -0.04em;\n    }\n\n    .notice-book {\n      margin: 13px 0 0;\n      color: #344c75;\n      font-size: 16px;\n      font-weight: 850;\n      line-height: 1.55;\n      word-break: keep-all;\n    }\n\n    .notice-description {\n      max-height: 300px;\n      overflow-y: auto;\n      margin: 16px 0 0;\n      padding: 14px 16px;\n      border: 1px solid #e1e8f4;\n      border-radius: 14px;\n      background: #f7f9fd;\n      color: #536179;\n      font-size: 14px;\n      font-weight: 700;\n      line-height: 1.72;\n      text-align: left;\n      word-break: keep-all;\n      white-space: pre-line;\n    }\n\n    .notice-close {\n      width: 100%;\n      min-height: 52px;\n      margin-top: 20px;\n      border: 0;\n      border-radius: 13px;\n      background: var(--blue);\n      color: #fff;\n      font-size: 16px;\n      font-weight: 900;\n      cursor: pointer;\n    }\n\n    .notice-close:hover {\n      background: var(--blue-dark);\n    }\n\n    .notice-actions {\n      display: grid;\n      grid-template-columns: 1fr 1fr;\n      gap: 10px;\n      margin-top: 20px;\n    }\n\n    .notice-actions.single {\n      grid-template-columns: 1fr;\n    }\n\n    .notice-actions .notice-close {\n      margin-top: 0;\n    }\n\n    .notice-undo {\n      min-height: 52px;\n      border: 0;\n      border-radius: 13px;\n      background: #f1ae45;\n      color: #49320e;\n      font-size: 15px;\n      font-weight: 950;\n      cursor: pointer;\n    }\n\n    .notice-undo:hover {\n      background: #e59c2f;\n    }\n\n    .notice-undo:disabled {\n      cursor: wait;\n      opacity: 0.58;\n    }\n\n    .notice-undo[hidden] {\n      display: none;\n    }\n\n    .notice-modal.owned .notice-close {\n      background: #d96545;\n    }\n\n    @keyframes notice-in {\n      from {\n        opacity: 0;\n        transform: translateY(12px) scale(0.97);\n      }\n\n      to {\n        opacity: 1;\n        transform: translateY(0) scale(1);\n      }\n    }\n\n    @media (min-width: 720px) {\n      body {\n        padding: 36px 24px;\n      }\n\n      .book-list {\n        grid-template-columns: repeat(2, minmax(0, 1fr));\n      }\n    }\n\n    @media (max-width: 420px) {\n      body {\n        padding: 17px;\n      }\n    }\n  \n    .bs-section {\n      width: 100%;\n      max-width: 760px;\n      margin: 18px auto 0;\n      padding: 17px;\n      border: 1px solid var(--line);\n      border-radius: 18px;\n      background: var(--paper);\n      box-shadow: 0 8px 25px rgba(50, 76, 125, 0.08);\n    }\n\n    .bs-heading {\n      display: flex;\n      align-items: flex-start;\n      justify-content: space-between;\n      gap: 12px;\n      margin-bottom: 14px;\n    }\n\n    .bs-title {\n      margin: 0;\n      font-size: 21px;\n      line-height: 1.35;\n      letter-spacing: -0.03em;\n    }\n\n    .bs-help {\n      margin: 5px 0 0;\n      color: var(--sub);\n      font-size: 13px;\n      line-height: 1.55;\n    }\n\n    .bs-badge {\n      flex: 0 0 auto;\n      padding: 6px 9px;\n      border-radius: 999px;\n      background: var(--blue-soft);\n      color: var(--blue-dark);\n      font-size: 11px;\n      font-weight: 850;\n    }\n\n    .bs-tabs {\n      display: grid;\n      grid-template-columns: repeat(4, 1fr);\n      gap: 7px;\n      margin-bottom: 14px;\n      padding: 5px;\n      border-radius: 13px;\n      background: #eef2f8;\n    }\n\n    .bs-tab {\n      min-height: 43px;\n      padding: 0 8px;\n      border: 0;\n      border-radius: 10px;\n      background: transparent;\n      color: #526078;\n      font-size: 15px;\n      font-weight: 850;\n      cursor: pointer;\n    }\n\n    .bs-tab.active {\n      background: #fff;\n      color: var(--blue-dark);\n      box-shadow: 0 3px 10px rgba(49, 89, 172, 0.13);\n    }\n\n    .bs-status {\n      display: none;\n      padding: 24px 10px;\n      color: var(--sub);\n      font-size: 14px;\n      text-align: center;\n    }\n\n    .bs-status.show {\n      display: block;\n    }\n\n    .bs-list {\n      display: grid;\n      grid-auto-flow: column;\n      grid-auto-columns: minmax(205px, 215px);\n      gap: 12px;\n      overflow-x: auto;\n      padding: 2px 2px 12px;\n      scroll-snap-type: x proximity;\n    }\n\n    .bs-card {\n      position: relative;\n      display: flex;\n      min-height: 330px;\n      flex-direction: column;\n      padding: 12px;\n      border: 1px solid #e1e7f2;\n      border-radius: 15px;\n      background: #fff;\n      scroll-snap-align: start;\n    }\n\n    .bs-rank {\n      position: absolute;\n      top: 9px;\n      left: 9px;\n      z-index: 1;\n      min-width: 31px;\n      height: 31px;\n      padding: 0 7px;\n      border-radius: 10px;\n      background: rgba(36, 50, 74, 0.92);\n      color: #fff;\n      font-size: 14px;\n      font-weight: 900;\n      line-height: 31px;\n      text-align: center;\n    }\n\n    .bs-cover {\n      display: flex;\n      width: 112px;\n      height: 158px;\n      margin: 0 auto 11px;\n      align-items: center;\n      justify-content: center;\n      overflow: hidden;\n      border-radius: 8px;\n      background: #eef1f6;\n      color: #7c8799;\n      font-size: 34px;\n      box-shadow: 0 7px 16px rgba(36, 50, 74, 0.14);\n    }\n\n    .bs-cover img {\n      width: 100%;\n      height: 100%;\n      object-fit: cover;\n    }\n\n    .bs-book-title {\n      display: -webkit-box;\n      overflow: hidden;\n      margin: 0;\n      color: var(--ink);\n      font-size: 15px;\n      font-weight: 900;\n      line-height: 1.45;\n      -webkit-box-orient: vertical;\n      -webkit-line-clamp: 2;\n    }\n\n    .bs-meta {\n      display: -webkit-box;\n      overflow: hidden;\n      min-height: 38px;\n      margin: 6px 0 10px;\n      color: var(--sub);\n      font-size: 12px;\n      line-height: 1.55;\n      -webkit-box-orient: vertical;\n      -webkit-line-clamp: 2;\n    }\n\n    .bs-actions {\n      display: grid;\n      grid-template-columns: 1fr auto;\n      gap: 7px;\n      margin-top: auto;\n    }\n\n    .bs-request {\n      min-height: 40px;\n      padding: 0 10px;\n      border: 0;\n      border-radius: 10px;\n      background: var(--blue);\n      color: #fff;\n      font-size: 12px;\n      font-weight: 900;\n      cursor: pointer;\n    }\n\n    .bs-link {\n      display: inline-flex;\n      min-height: 40px;\n      padding: 0 10px;\n      align-items: center;\n      justify-content: center;\n      border: 1px solid var(--line);\n      border-radius: 10px;\n      color: #5b6880;\n      font-size: 12px;\n      font-weight: 800;\n      text-decoration: none;\n    }\n\n    .bs-source {\n      margin: 4px 1px 0;\n      color: #7d8798;\n      font-size: 11px;\n      text-align: right;\n    }\n\n    .bs-source a {\n      color: var(--blue-dark);\n    }\n\n    @media (max-width: 520px) {\n      .bs-heading {\n        display: block;\n      }\n\n      .bs-badge {\n        display: inline-block;\n        margin-top: 8px;\n      }\n\n      .bs-list {\n        grid-auto-columns: minmax(190px, 78vw);\n      }\n    }\n\n  \n    .bs-more-wrap {\n      margin-top: 6px;\n    }\n\n    .bs-more {\n      display: flex;\n      width: 100%;\n      min-height: 48px;\n      align-items: center;\n      justify-content: center;\n      border: 1px solid #c9d7f2;\n      border-radius: 12px;\n      background: var(--blue-soft);\n      color: var(--blue-dark);\n      font-size: 15px;\n      font-weight: 900;\n      text-decoration: none;\n      transition: background 0.18s, transform 0.18s;\n    }\n\n    .bs-more:hover {\n      background: #dfe9ff;\n      transform: translateY(-1px);\n    }\n\n  \n    .view-tabs {\n      position: sticky;\n      top: 10px;\n      z-index: 30;\n      display: grid;\n      width: 100%;\n      max-width: 760px;\n      margin: 16px auto;\n      padding: 6px;\n      grid-template-columns: repeat(4, 1fr);\n      gap: 7px;\n      border: 1px solid var(--line);\n      border-radius: 16px;\n      background: rgba(255, 255, 255, 0.96);\n      box-shadow: 0 8px 24px rgba(50, 76, 125, 0.11);\n      backdrop-filter: blur(12px);\n    }\n\n    .view-tab {\n      min-height: 49px;\n      padding: 0 12px;\n      border: 0;\n      border-radius: 11px;\n      background: transparent;\n      color: #59677e;\n      font-size: 16px;\n      font-weight: 900;\n      cursor: pointer;\n      transition: background 0.18s, color 0.18s, transform 0.18s;\n    }\n\n    .view-tab:hover {\n      background: #f1f5fd;\n    }\n\n    .view-tab.active {\n      background: var(--blue);\n      color: #fff;\n      box-shadow: 0 5px 13px rgba(79, 120, 209, 0.25);\n    }\n\n    .page-panel {\n      display: none;\n    }\n\n    .page-panel.active {\n      display: block;\n    }\n\n    .bs-carousel {\n      position: relative;\n    }\n\n    .bs-list {\n      scroll-behavior: smooth;\n      scrollbar-width: thin;\n      scrollbar-color: #aebedd #eef2f8;\n      overscroll-behavior-inline: contain;\n      padding-left: 44px;\n      padding-right: 44px;\n    }\n\n    .bs-arrow {\n      position: absolute;\n      top: 50%;\n      z-index: 5;\n      display: flex;\n      width: 48px;\n      height: 58px;\n      margin-top: -29px;\n      align-items: center;\n      justify-content: center;\n      border: 1px solid #cad7ee;\n      border-radius: 16px;\n      background: rgba(255, 255, 255, 0.95);\n      color: var(--blue-dark);\n      font-family: Arial, sans-serif;\n      font-size: 40px;\n      font-weight: 700;\n      line-height: 1;\n      box-shadow: 0 7px 20px rgba(36, 50, 74, 0.18);\n      cursor: pointer;\n      transition: transform 0.18s, background 0.18s;\n    }\n\n    .bs-arrow:hover {\n      background: var(--blue-soft);\n      transform: scale(1.05);\n    }\n\n    .bs-arrow-previous {\n      left: -6px;\n    }\n\n    .bs-arrow-next {\n      right: -6px;\n    }\n\n    .requests-section,\n    .guide-section {\n      width: 100%;\n      max-width: 760px;\n      margin: 0 auto;\n      padding: 22px;\n      border: 1px solid var(--line);\n      border-radius: 20px;\n      background: #fff;\n      box-shadow: 0 10px 28px rgba(50, 76, 125, 0.09);\n    }\n\n    .panel-heading {\n      display: flex;\n      align-items: flex-start;\n      justify-content: space-between;\n      gap: 16px;\n      margin-bottom: 18px;\n    }\n\n    .panel-eyebrow {\n      margin: 0 0 5px !important;\n      color: var(--blue);\n      font-size: 11px !important;\n      font-weight: 900;\n      letter-spacing: 0.12em;\n    }\n\n    .panel-heading h2 {\n      margin: 0;\n      color: var(--ink);\n      font-size: 24px;\n      letter-spacing: -0.035em;\n    }\n\n    .panel-heading p:not(.panel-eyebrow) {\n      margin: 7px 0 0;\n      color: var(--sub);\n      font-size: 14px;\n      line-height: 1.6;\n    }\n\n    .refresh-button {\n      flex: 0 0 auto;\n      min-height: 42px;\n      padding: 0 15px;\n      border: 1px solid #cbd8ef;\n      border-radius: 11px;\n      background: var(--blue-soft);\n      color: var(--blue-dark);\n      font-size: 14px;\n      font-weight: 900;\n      cursor: pointer;\n    }\n\n    .request-filter-tabs {\n      display: grid;\n      grid-template-columns: 1fr 1fr;\n      gap: 10px;\n      margin: 18px 0 13px;\n      padding: 6px;\n      border: 1px solid #dce5df;\n      border-radius: 16px;\n      background: #f3f6f3;\n    }\n\n    .request-filter-tab {\n      display: flex;\n      min-height: 54px;\n      padding: 0 18px;\n      border: 0;\n      border-radius: 12px;\n      align-items: center;\n      justify-content: space-between;\n      gap: 12px;\n      background: transparent;\n      color: #63726a;\n      font-size: 15px;\n      font-weight: 900;\n      cursor: pointer;\n      transition:\n        background 0.16s ease,\n        color 0.16s ease,\n        box-shadow 0.16s ease;\n    }\n\n    .request-filter-tab strong {\n      display: inline-flex;\n      min-width: 32px;\n      min-height: 28px;\n      padding: 0 8px;\n      border-radius: 999px;\n      align-items: center;\n      justify-content: center;\n      background: #e3e9e5;\n      color: #526258;\n      font-size: 13px;\n    }\n\n    .request-filter-tab.active {\n      background: #2f604e;\n      color: #fff;\n      box-shadow: 0 7px 18px rgba(47, 96, 78, 0.18);\n    }\n\n    .request-filter-tab.active strong {\n      background: #fff;\n      color: #2f604e;\n    }\n\n    .request-filter-tab[data-request-filter=\"cancelled\"].active {\n      background: #a84d45;\n      box-shadow: 0 7px 18px rgba(168, 77, 69, 0.17);\n    }\n\n    .request-filter-tab[data-request-filter=\"cancelled\"].active strong {\n      color: #a84d45;\n    }\n\n    @media (max-width: 520px) {\n      .request-filter-tab {\n        min-height: 50px;\n        padding: 0 12px;\n        font-size: 14px;\n      }\n    }\n\n    .request-summary {\n      display: none;\n      margin-bottom: 13px;\n      padding: 11px 14px;\n      border-radius: 11px;\n      background: #f0f5ff;\n      color: #3b5eaa;\n      font-size: 13px;\n      font-weight: 850;\n    }\n\n    .request-summary.show {\n      display: block;\n    }\n\n    .request-list-status {\n      padding: 34px 14px;\n      border: 1px dashed #cfd9ea;\n      border-radius: 14px;\n      color: var(--sub);\n      text-align: center;\n    }\n\n    .request-list-status.hide {\n      display: none;\n    }\n\n    .request-list {\n      display: grid;\n      grid-template-columns: repeat(2, minmax(0, 1fr));\n      gap: 12px;\n    }\n\n    .request-item {\n      display: flex;\n      min-height: 132px;\n      flex-direction: column;\n      padding: 15px;\n      border: 1px solid #e0e7f2;\n      border-radius: 15px;\n      background: #fbfcff;\n    }\n\n    .request-item-top {\n      display: flex;\n      align-items: center;\n      justify-content: space-between;\n      gap: 8px;\n      margin-bottom: 9px;\n    }\n\n    .request-sequence {\n      color: #70809a;\n      font-size: 12px;\n      font-weight: 900;\n    }\n\n    .request-status-badge {\n      max-width: 62%;\n      padding: 5px 8px;\n      overflow: hidden;\n      border-radius: 999px;\n      background: #fff4d8;\n      color: #936214;\n      font-size: 11px;\n      font-weight: 900;\n      text-overflow: ellipsis;\n      white-space: nowrap;\n    }\n\n    .request-status-badge.done {\n      background: #e6f6eb;\n      color: #247347;\n    }\n\n    .request-status-badge.issue {\n      background: #fff0ef;\n      color: #a5413d;\n    }\n\n    .request-item h3 {\n      margin: 0;\n      color: var(--ink);\n      font-size: 16px;\n      font-weight: 900;\n      line-height: 1.45;\n    }\n\n    .request-item p {\n      margin: 7px 0 0;\n      color: var(--sub);\n      font-size: 13px;\n      line-height: 1.55;\n    }\n\n    .request-date {\n      margin-top: auto !important;\n      padding-top: 9px;\n      color: #8691a3 !important;\n      font-size: 11px !important;\n    }\n\n    .guide-steps {\n      display: grid;\n      gap: 12px;\n      margin: 0;\n      padding: 0;\n      list-style: none;\n    }\n\n    .guide-steps li {\n      display: grid;\n      grid-template-columns: 48px 1fr;\n      gap: 13px;\n      padding: 16px;\n      border: 1px solid #e0e7f2;\n      border-radius: 15px;\n      background: #fbfcff;\n    }\n\n    .guide-number {\n      display: flex;\n      width: 45px;\n      height: 45px;\n      align-items: center;\n      justify-content: center;\n      border-radius: 14px;\n      background: var(--blue);\n      color: #fff;\n      font-size: 20px;\n      font-weight: 950;\n    }\n\n    .guide-steps strong {\n      color: var(--ink);\n      font-size: 16px;\n      line-height: 1.5;\n    }\n\n    .guide-steps p {\n      margin: 6px 0 0;\n      color: var(--sub);\n      font-size: 14px;\n      line-height: 1.65;\n    }\n\n    .guide-start-button {\n      width: 100%;\n      min-height: 54px;\n      margin-top: 17px;\n      border: 0;\n      border-radius: 13px;\n      background: var(--blue);\n      color: #fff;\n      font-size: 17px;\n      font-weight: 900;\n      cursor: pointer;\n    }\n\n    @media (max-width: 620px) {\n      .view-tab {\n        min-height: 46px;\n        padding: 0 5px;\n        font-size: 13px;\n      }\n\n      .request-list {\n        grid-template-columns: 1fr;\n      }\n\n      .requests-section,\n      .guide-section {\n        padding: 17px;\n      }\n\n      .panel-heading {\n        display: block;\n      }\n\n      .refresh-button {\n        width: 100%;\n        margin-top: 12px;\n      }\n\n      .bs-arrow {\n        width: 43px;\n        height: 54px;\n        font-size: 35px;\n      }\n    }\n\n\n    .guide-hover {\n      position: relative;\n      min-width: 0;\n    }\n\n    .guide-hover-trigger {\n      width: 100%;\n      min-height: 49px;\n      padding: 0 12px;\n      border: 0;\n      border-radius: 11px;\n      background: #fff5d9;\n      color: #7b5813;\n      font-size: 16px;\n      font-weight: 900;\n      cursor: help;\n      transition: background 0.18s, transform 0.18s;\n    }\n\n    .guide-hover-trigger:hover,\n    .guide-hover-trigger:focus {\n      background: #ffeab2;\n      transform: translateY(-1px);\n      outline: none;\n    }\n\n    .guide-popover {\n      position: absolute;\n      top: calc(100% + 10px);\n      right: 0;\n      z-index: 100;\n      width: min(440px, calc(100vw - 34px));\n      max-height: calc(100vh - 120px);\n      padding: 18px;\n      overflow-y: auto;\n      border: 1px solid #ead79c;\n      border-radius: 17px;\n      background: #fffdf6;\n      box-shadow: 0 16px 40px rgba(36, 50, 74, 0.22);\n      opacity: 0;\n      visibility: hidden;\n      transform: translateY(-7px);\n      pointer-events: none;\n      transition: opacity 0.16s, transform 0.16s, visibility 0.16s;\n    }\n\n    .guide-hover:hover .guide-popover,\n    .guide-hover:focus-within .guide-popover,\n    .guide-hover.open .guide-popover {\n      opacity: 1;\n      visibility: visible;\n      transform: translateY(0);\n      pointer-events: auto;\n    }\n\n    .guide-popover::before {\n      position: absolute;\n      top: -7px;\n      right: 34px;\n      width: 13px;\n      height: 13px;\n      border-top: 1px solid #ead79c;\n      border-left: 1px solid #ead79c;\n      background: #fffdf6;\n      content: '';\n      transform: rotate(45deg);\n    }\n\n    .guide-popover h2 {\n      margin: 0 0 12px;\n      color: var(--ink);\n      font-size: 20px;\n      letter-spacing: -0.03em;\n    }\n\n    .guide-popover ol {\n      display: grid;\n      gap: 8px;\n      margin: 0;\n      padding: 0;\n      list-style: none;\n    }\n\n    .guide-popover li {\n      display: grid;\n      grid-template-columns: 108px 1fr;\n      gap: 9px;\n      padding: 10px;\n      border-radius: 11px;\n      background: #fff8e7;\n    }\n\n    .guide-popover strong {\n      color: #725214;\n      font-size: 13px;\n    }\n\n    .guide-popover span {\n      color: #4f5b6d;\n      font-size: 13px;\n      line-height: 1.5;\n    }\n\n    .guide-popover p {\n      margin: 12px 0 0;\n      padding: 11px;\n      border-radius: 11px;\n      background: #eef4ff;\n      color: #3f5f9d;\n      font-size: 12px;\n      font-weight: 800;\n      line-height: 1.55;\n    }\n\n    .request-card-actions {\n      display: flex;\n      justify-content: flex-end;\n      margin-top: auto;\n      padding-top: 12px;\n    }\n\n    .request-cancel-button {\n      min-height: 37px;\n      padding: 0 13px;\n      border: 1px solid #edc5c0;\n      border-radius: 10px;\n      background: #fff5f3;\n      color: #a2463e;\n      font-size: 12px;\n      font-weight: 900;\n      cursor: pointer;\n    }\n\n    .request-cancel-button:hover {\n      background: #ffe6e2;\n    }\n\n    .request-cancel-button.restore {\n      border-color: #b9d9c2;\n      background: #edf8f0;\n      color: #27643b;\n    }\n\n    .request-cancel-button.restore:hover {\n      background: #dff1e4;\n    }\n\n    .request-cancel-button:disabled {\n      border-color: #d9dee7;\n      background: #f1f3f6;\n      color: #8892a2;\n      cursor: default;\n    }\n\n    .cancel-modal {\n      position: fixed;\n      inset: 0;\n      z-index: 120;\n      display: none;\n      padding: 22px;\n      align-items: center;\n      justify-content: center;\n      background: rgba(25, 34, 51, 0.58);\n      backdrop-filter: blur(5px);\n    }\n\n    .cancel-modal.show {\n      display: flex;\n    }\n\n    .cancel-dialog {\n      width: min(430px, 100%);\n      padding: 25px;\n      border-radius: 22px;\n      background: #fff;\n      box-shadow: 0 22px 60px rgba(24, 34, 52, 0.3);\n      text-align: center;\n      animation: notice-in 0.2s ease-out;\n    }\n\n    .cancel-icon {\n      margin-bottom: 8px;\n      font-size: 42px;\n    }\n\n    .cancel-dialog h2 {\n      margin: 0;\n      color: var(--ink);\n      font-size: 22px;\n      letter-spacing: -0.035em;\n    }\n\n    .cancel-book {\n      margin: 12px 0 0;\n      color: var(--blue-dark);\n      font-size: 16px;\n      font-weight: 900;\n      line-height: 1.5;\n    }\n\n    .cancel-description {\n      margin: 10px 0 18px;\n      color: var(--sub);\n      font-size: 13px;\n      line-height: 1.6;\n    }\n\n    .cancel-actions {\n      display: grid;\n      grid-template-columns: 1fr 1fr;\n      gap: 9px;\n    }\n\n    .cancel-actions button {\n      min-height: 48px;\n      border-radius: 12px;\n      font-size: 14px;\n      font-weight: 900;\n      cursor: pointer;\n    }\n\n    .cancel-keep {\n      border: 1px solid #d8e0ed;\n      background: #fff;\n      color: #59677e;\n    }\n\n    .cancel-confirm {\n      border: 0;\n      background: #d95d51;\n      color: #fff;\n    }\n\n    .cancel-confirm:disabled {\n      opacity: 0.58;\n      cursor: wait;\n    }\n\n    @media (max-width: 620px) {\n      .guide-hover-trigger {\n        min-height: 46px;\n        padding: 0 5px;\n        font-size: 13px;\n      }\n\n      .guide-popover {\n        position: fixed;\n        top: 82px;\n        right: 17px;\n        left: 17px;\n        width: auto;\n      }\n\n      .guide-popover::before {\n        display: none;\n      }\n\n      .guide-popover li {\n        grid-template-columns: 1fr;\n        gap: 3px;\n      }\n    }\n\n\n    .request-item {\n      cursor: pointer;\n      transition:\n        border-color 0.18s ease,\n        box-shadow 0.18s ease,\n        transform 0.18s ease;\n    }\n\n    .request-item:hover,\n    .request-item:focus-visible,\n    .request-item.expanded {\n      border-color: #aebfe8;\n      box-shadow: 0 12px 28px rgba(48, 78, 143, 0.13);\n    }\n\n    .request-biblio {\n      display: none;\n      gap: 8px;\n      margin-top: 14px;\n      padding: 14px;\n      border: 1px solid #dfe6f5;\n      border-radius: 14px;\n      background: #f8faff;\n    }\n\n    .request-item.expanded .request-biblio {\n      display: grid;\n    }\n\n    .request-biblio-row {\n      display: grid;\n      grid-template-columns: 84px minmax(0, 1fr);\n      gap: 10px;\n      align-items: start;\n      color: #34445e;\n      font-size: 15px;\n      line-height: 1.5;\n    }\n\n    .request-biblio-row strong {\n      color: #173567;\n      font-weight: 800;\n    }\n\n    .request-biblio-row span,\n    .request-biblio-row a {\n      min-width: 0;\n      overflow-wrap: anywhere;\n    }\n\n    .request-biblio-row a {\n      color: #315fbd;\n      font-weight: 800;\n    }\n\n    .request-biblio-empty {\n      margin: 0;\n      color: #68738a;\n      font-size: 14px;\n    }\n\n    .request-card-actions {\n      display: flex;\n      gap: 10px;\n      margin-top: 14px;\n    }\n\n    .request-detail-button,\n    .request-cancel-button {\n      min-height: 42px;\n      padding: 0 15px;\n      border-radius: 12px;\n      font-size: 14px;\n      font-weight: 850;\n    }\n\n    .request-detail-button {\n      flex: 1;\n      border: 1px solid #cdd9f2;\n      background: #eef3ff;\n      color: #234b98;\n    }\n\n    .request-cancel-button {\n      border: 1px solid #ebc6c2;\n      background: #fff5f4;\n      color: #a03e37;\n    }\n\n    .request-cancel-button:disabled {\n      border-color: #e1e4ea;\n      background: #f4f5f7;\n      color: #8a909b;\n      cursor: default;\n    }\n\n    @media (max-width: 560px) {\n      .request-biblio-row {\n        grid-template-columns: 74px minmax(0, 1fr);\n        font-size: 14px;\n      }\n\n      .request-card-actions {\n        flex-direction: column;\n      }\n\n      .request-detail-button,\n      .request-cancel-button {\n        width: 100%;\n      }\n    }\n\n\n  \n    /* HR_WEB_NOTICE_TICKER */\n    .hr-notice-ticker[hidden] {\n      display: none !important;\n    }\n\n    .hr-notice-ticker {\n      width: min(1160px, calc(100% - 34px));\n      margin: 14px auto 0;\n      overflow: hidden;\n      border: 1px solid #ead39a;\n      border-radius: 15px;\n      background: linear-gradient(90deg, #fff5d7 0%, #fffaf0 100%);\n      color: #6f4b12;\n      box-shadow: 0 8px 24px rgba(102, 78, 29, 0.09);\n    }\n\n    .hr-notice-ticker.urgent {\n      border-color: #efb2aa;\n      background: linear-gradient(90deg, #fff0ed 0%, #fff8f4 100%);\n      color: #9b332d;\n    }\n\n    .hr-notice-inner {\n      display: flex;\n      min-height: 52px;\n      align-items: stretch;\n    }\n\n    .hr-notice-label {\n      display: flex;\n      flex: 0 0 auto;\n      align-items: center;\n      justify-content: center;\n      padding: 0 20px;\n      background: #d9a13b;\n      color: #fff;\n      font-size: 15px;\n      font-weight: 900;\n      letter-spacing: .02em;\n    }\n\n    .hr-notice-ticker.urgent .hr-notice-label {\n      background: #c44c43;\n    }\n\n    .hr-notice-window {\n      min-width: 0;\n      flex: 1;\n      overflow: hidden;\n      mask-image: linear-gradient(90deg, transparent, #000 4%, #000 96%, transparent);\n      -webkit-mask-image: linear-gradient(90deg, transparent, #000 4%, #000 96%, transparent);\n    }\n\n    .hr-notice-track {\n      display: inline-block;\n      min-width: max-content;\n      padding: 14px 0 13px 100%;\n      white-space: nowrap;\n      font-size: 17px;\n      font-weight: 850;\n      line-height: 1.45;\n      animation: hrTickerMove var(--ticker-duration, 26s) linear infinite;\n      will-change: transform;\n    }\n\n    .hr-notice-ticker:hover .hr-notice-track,\n    .hr-notice-ticker:focus-within .hr-notice-track {\n      animation-play-state: paused;\n    }\n\n    @keyframes hrTickerMove {\n      from { transform: translateX(0); }\n      to { transform: translateX(-100%); }\n    }\n\n    @media (max-width: 640px) {\n      .hr-notice-ticker {\n        width: calc(100% - 20px);\n        margin-top: 10px;\n        border-radius: 12px;\n      }\n\n      .hr-notice-label {\n        padding: 0 13px;\n        font-size: 13px;\n      }\n\n      .hr-notice-track {\n        font-size: 15px;\n      }\n    }\n\n    @media (prefers-reduced-motion: reduce) {\n      .hr-notice-track {\n        min-width: 100%;\n        padding-left: 16px;\n        padding-right: 16px;\n        white-space: normal;\n        animation: none;\n      }\n    }\n\n  \n    /* HR_WEB_NOTICE_TICKER_V2 */\n    .hr-notice-ticker {\n      width: min(920px, calc(100% - 24px));\n      margin: 14px auto 12px;\n      border: 1px solid rgba(205, 157, 65, 0.42);\n      border-radius: 18px;\n      background: rgba(255, 253, 248, 0.98);\n      color: #694713;\n      box-shadow: 0 10px 28px rgba(45, 70, 58, 0.11);\n    }\n\n    .hr-notice-ticker + .view-tabs {\n      margin-top: 0;\n    }\n\n    .hr-notice-inner {\n      min-height: 48px;\n      align-items: stretch;\n    }\n\n    .hr-notice-label {\n      gap: 5px;\n      padding: 0 16px;\n      border-radius: 17px 0 0 17px;\n      background: linear-gradient(135deg, #dda842 0%, #c98d2e 100%);\n      font-size: 14px;\n      letter-spacing: 0;\n    }\n\n    .hr-notice-window {\n      display: flex;\n      align-items: center;\n    }\n\n    .hr-notice-track {\n      padding: 12px 0 11px 100%;\n      font-size: 15px;\n      font-weight: 800;\n      line-height: 1.45;\n    }\n\n    .hr-notice-ticker.urgent {\n      border-color: rgba(194, 76, 67, 0.42);\n      background: #fff8f5;\n      color: #95352f;\n    }\n\n    .hr-notice-ticker.urgent .hr-notice-label {\n      background: linear-gradient(135deg, #d45b52 0%, #b63d36 100%);\n    }\n\n    @media (max-width: 640px) {\n      .hr-notice-ticker {\n        width: calc(100% - 20px);\n        margin: 10px auto 10px;\n        border-radius: 14px;\n      }\n\n      .hr-notice-ticker + .view-tabs {\n        margin-top: 0;\n      }\n\n      .hr-notice-label {\n        padding: 0 12px;\n        border-radius: 13px 0 0 13px;\n        font-size: 12px;\n      }\n\n      .hr-notice-track {\n        padding-top: 11px;\n        padding-bottom: 10px;\n        font-size: 14px;\n      }\n    }\n\n    /* HR_WEB_NOTICE_TICKER_HEIGHT_FIX_V3\n       모션 줄이기 환경에서 기존 규칙과 V2 규칙이 겹치면\n       공지 문구가 한 글자씩 줄바꿈되며 전광판이 길어지는 문제를 방지합니다. */\n    .hr-notice-ticker {\n      height: 50px;\n      min-height: 50px;\n      max-height: 50px;\n      overflow: hidden;\n    }\n\n    .hr-notice-inner,\n    .hr-notice-label,\n    .hr-notice-window {\n      height: 48px;\n      min-height: 48px;\n      max-height: 48px;\n    }\n\n    .hr-notice-inner,\n    .hr-notice-window {\n      overflow: hidden;\n    }\n\n    .hr-notice-label {\n      align-self: flex-start;\n      white-space: nowrap;\n    }\n\n    .hr-notice-track {\n      height: 48px;\n      max-height: 48px;\n      overflow: hidden;\n      box-sizing: border-box;\n      white-space: nowrap;\n    }\n\n    @media (prefers-reduced-motion: reduce) {\n      .hr-notice-track {\n        width: auto;\n        min-width: max-content;\n        padding: 12px 0 11px 100%;\n        overflow: hidden;\n        white-space: nowrap;\n        text-overflow: clip;\n        animation: hrTickerMove var(--ticker-duration, 26s) linear infinite;\n      }\n    }\n  </style>\n\n  <style id=\"visualRefreshV20\">\n    :root {\n      --blue: #2f6653;\n      --blue-dark: #234d40;\n      --blue-soft: #e8f2ec;\n      --ink: #25352f;\n      --sub: #68766f;\n      --line: #dbe4dc;\n      --paper: #fffdf8;\n      --background: #f4f3eb;\n      --gold: #d3a248;\n      --gold-dark: #ad762b;\n      --cream: #fbf4e5;\n      --forest: #2c5949;\n      --forest-dark: #203f35;\n    }\n\n    body {\n      min-height: 100vh;\n      padding: 28px 18px 64px;\n      background:\n        radial-gradient(circle at 7% 5%, rgba(234, 213, 165, 0.52), transparent 27%),\n        radial-gradient(circle at 96% 13%, rgba(188, 215, 198, 0.58), transparent 26%),\n        linear-gradient(180deg, #faf7ef 0%, #eef3ee 100%);\n      color: var(--ink);\n    }\n\n    .header {\n      position: relative;\n      width: 100%;\n      max-width: 820px;\n      margin: 0 auto;\n      padding: 34px 38px 50px;\n      overflow: hidden;\n      border: 1px solid rgba(255, 255, 255, 0.32);\n      border-radius: 30px;\n      background:\n        linear-gradient(135deg, rgba(255,255,255,0.06), transparent 44%),\n        linear-gradient(135deg, #23483b 0%, #386a56 100%);\n      color: #fff;\n      box-shadow: 0 20px 45px rgba(38, 76, 62, 0.2);\n    }\n\n    .header::before {\n      position: absolute;\n      top: -85px;\n      right: -45px;\n      width: 230px;\n      height: 230px;\n      border: 42px solid rgba(255, 255, 255, 0.055);\n      border-radius: 50%;\n      content: \"\";\n    }\n\n    .header::after {\n      position: absolute;\n      right: 34px;\n      bottom: 18px;\n      content: \"📚\";\n      font-size: 78px;\n      line-height: 1;\n      opacity: 0.2;\n      transform: rotate(-4deg);\n    }\n\n    .header > * {\n      position: relative;\n      z-index: 1;\n    }\n\n    .eyebrow {\n      margin-bottom: 10px;\n      color: #f2d58f;\n      font-size: 12px;\n      letter-spacing: 0.18em;\n    }\n\n    h1 {\n      color: #fff;\n      font-size: clamp(30px, 5vw, 40px);\n      letter-spacing: -0.045em;\n    }\n\n    .description {\n      max-width: 500px;\n      margin-top: 12px;\n      color: #e6efe9;\n      font-size: 15px;\n      line-height: 1.7;\n    }\n\n    .view-tabs {\n      top: 10px;\n      max-width: 920px;\n      margin: -25px auto 25px;\n      padding: 7px;\n      grid-template-columns: repeat(4, minmax(0, 1fr));\n      gap: 7px;\n      border: 1px solid rgba(49, 91, 73, 0.14);\n      border-radius: 20px;\n      background: rgba(255, 253, 248, 0.96);\n      box-shadow: 0 14px 34px rgba(52, 75, 63, 0.16);\n    }\n\n    .view-tab,\n    .guide-hover-trigger {\n      min-height: 51px;\n      border-radius: 14px;\n      color: #52655c;\n      font-size: 15px;\n    }\n\n    .view-tab:hover {\n      background: #edf4ef;\n    }\n\n    .view-tab.active {\n      background: linear-gradient(135deg, #d8ae58 0%, #bd8433 100%);\n      color: #fff;\n      box-shadow: 0 7px 18px rgba(177, 121, 43, 0.26);\n    }\n\n    .guide-hover-trigger {\n      background: #edf4ef;\n      color: #315c4c;\n    }\n\n    .guide-hover-trigger:hover,\n    .guide-hover-trigger:focus {\n      background: #dfede4;\n      color: #244b3d;\n    }\n\n    .guide-popover {\n      border-color: #d7e1d8;\n      border-radius: 20px;\n      background: #fffdf8;\n      box-shadow: 0 20px 48px rgba(37, 67, 54, 0.2);\n    }\n\n    .guide-popover h2 {\n      color: #294f41;\n    }\n\n    .card,\n    .bs-section,\n    .requests-section {\n      max-width: 820px;\n      border: 1px solid rgba(56, 92, 76, 0.12);\n      border-radius: 26px;\n      background: rgba(255, 253, 248, 0.96);\n      box-shadow: 0 15px 38px rgba(46, 72, 59, 0.1);\n    }\n\n    .card {\n      padding: 24px;\n    }\n\n    .form-intro {\n      display: flex;\n      align-items: center;\n      gap: 14px;\n      margin: 0 0 20px;\n      padding: 0 2px 18px;\n      border-bottom: 1px dashed #d9e2da;\n    }\n\n    .form-intro-icon {\n      display: grid;\n      flex: 0 0 48px;\n      width: 48px;\n      height: 48px;\n      place-items: center;\n      border-radius: 16px;\n      background: #e8f2ec;\n      font-size: 25px;\n    }\n\n    .form-intro h2 {\n      margin: 0;\n      color: #284b3e;\n      font-size: 20px;\n      letter-spacing: -0.03em;\n    }\n\n    .form-intro p {\n      margin: 4px 0 0;\n      color: var(--sub);\n      font-size: 13px;\n      line-height: 1.5;\n    }\n\n    label {\n      color: #314e43;\n      font-size: 16px;\n    }\n\n    input {\n      height: 57px;\n      border-color: #d6dfd7;\n      border-radius: 15px;\n      background: #fff;\n      color: #263831;\n      font-size: 17px;\n    }\n\n    input:focus {\n      border-color: #4f806d;\n      box-shadow: 0 0 0 4px rgba(68, 121, 98, 0.13);\n    }\n\n    .submit-button {\n      min-height: 58px;\n      border-radius: 15px;\n      background: linear-gradient(135deg, #356b57 0%, #274f42 100%);\n      box-shadow: 0 9px 20px rgba(39, 79, 66, 0.2);\n    }\n\n    .submit-button:hover {\n      background: linear-gradient(135deg, #2d604e 0%, #1f4338 100%);\n    }\n\n    .bs-section {\n      margin-top: 20px;\n      padding: 23px;\n    }\n\n    .bs-title {\n      color: #294d40;\n      font-size: 23px;\n    }\n\n    .bs-badge {\n      background: #f6ead1;\n      color: #90621f;\n    }\n\n    .bs-tabs {\n      gap: 8px;\n      padding: 6px;\n      border: 1px solid #dce5dd;\n      border-radius: 16px;\n      background: #edf3ee;\n    }\n\n    .bs-tab {\n      min-height: 45px;\n      border-radius: 12px;\n      color: #5b6c64;\n    }\n\n    .bs-tab.active {\n      color: #274e40;\n      box-shadow: 0 5px 13px rgba(50, 83, 68, 0.13);\n    }\n\n    .bs-card {\n      border-color: #e3e4d9;\n      border-radius: 20px;\n      background: #fffefb;\n      box-shadow: 0 8px 20px rgba(70, 76, 61, 0.08);\n    }\n\n    .bs-card:hover {\n      border-color: #c8d7cb;\n      box-shadow: 0 13px 28px rgba(51, 79, 65, 0.15);\n    }\n\n    .bs-rank {\n      background: #d4a64f;\n      box-shadow: 0 5px 12px rgba(179, 126, 42, 0.25);\n    }\n\n    .bs-request {\n      background: #356b57;\n    }\n\n    .bs-request:hover {\n      background: #284f42;\n    }\n\n    .bs-link,\n    .bs-more {\n      border-color: #cadbcf;\n      background: #edf5ef;\n      color: #2e5c4b;\n    }\n\n    .bs-more:hover {\n      background: #dfece3;\n    }\n\n    .bs-arrow {\n      border-color: rgba(65, 100, 83, 0.2);\n      background: rgba(255, 253, 248, 0.94);\n      color: #315e4d;\n      box-shadow: 0 8px 22px rgba(47, 77, 62, 0.16);\n    }\n\n    .requests-section {\n      padding: 25px;\n    }\n\n    .panel-eyebrow {\n      color: #bf842f;\n    }\n\n    .panel-heading h2 {\n      color: #294e40;\n    }\n\n    .request-item {\n      border-color: #dde5de;\n      border-radius: 20px;\n      background: #fffefb;\n    }\n\n    .request-item:hover,\n    .request-item:focus-visible,\n    .request-item.expanded {\n      border-color: #acc6b5;\n      box-shadow: 0 13px 30px rgba(43, 78, 61, 0.14);\n    }\n\n    .request-biblio {\n      border-color: #dce6de;\n      background: #f5f8f5;\n    }\n\n    .refresh-button,\n    .request-detail-button {\n      border-color: #cadbcf;\n      background: #edf5ef;\n      color: #2d5a49;\n    }\n\n    .cart-bar {\n      max-width: 820px;\n      border-color: #d7e1d8;\n      border-radius: 20px;\n      background: rgba(255, 253, 248, 0.97);\n      box-shadow: 0 15px 36px rgba(44, 74, 59, 0.18);\n    }\n\n    .cart-submit {\n      background: #356b57;\n    }\n\n    .footer {\n      color: #728078;\n    }\n\n    .creator-watermark {\n      color: rgba(45, 77, 63, 0.55);\n    }\n\n    @media (min-width: 760px) {\n      #requestForm {\n        display: grid;\n        grid-template-columns: 0.82fr 1.18fr 170px;\n        gap: 14px;\n        align-items: end;\n      }\n\n      #requestForm .field {\n        margin-bottom: 0;\n      }\n\n      #requestForm .help {\n        min-height: 39px;\n      }\n\n      #requestForm .submit-button {\n        width: 170px;\n        margin-top: 0;\n        margin-bottom: 1px;\n      }\n    }\n\n    @media (max-width: 620px) {\n      body {\n        padding: 14px 12px 54px;\n      }\n\n      .header {\n        padding: 27px 24px 44px;\n        border-radius: 24px;\n      }\n\n      .header::after {\n        right: 17px;\n        bottom: 15px;\n        font-size: 55px;\n      }\n\n      .view-tabs {\n        margin: -22px 8px 20px;\n        grid-template-columns: 1fr 1fr 0.78fr;\n        border-radius: 17px;\n      }\n\n      .view-tab,\n      .guide-hover-trigger {\n        min-height: 48px;\n        padding: 0 5px;\n        font-size: 13px;\n      }\n\n      .card,\n      .bs-section,\n      .requests-section {\n        border-radius: 22px;\n      }\n\n      .card,\n      .bs-section {\n        padding: 18px;\n      }\n\n      .form-intro {\n        align-items: flex-start;\n      }\n    }\n  </style>\n\n\n  <style id=\"requestCoverV22\">\n    .header {\n      min-height: 0;\n    }\n\n    .request-biblio-cover {\n      display: flex;\n      grid-column: 1 / -1;\n      justify-content: center;\n      margin: 0 0 5px;\n      padding: 5px 0 12px;\n      border-bottom: 1px dashed #d7e2d9;\n    }\n\n    .request-biblio-cover img {\n      width: 112px;\n      height: 156px;\n      object-fit: cover;\n      border: 1px solid #d7dfd8;\n      border-radius: 12px;\n      background: #fff;\n      box-shadow: 0 9px 22px rgba(43, 67, 55, 0.17);\n    }\n\n    @media (max-width: 720px) {\n      .header {\n        min-height: 0;\n      }\n\n      .request-biblio-cover img {\n        width: 100px;\n        height: 140px;\n      }\n    }\n    .school-contact-footer {\n      width: min(1180px, calc(100% - 32px));\n      margin: 42px auto 118px;\n      padding: 24px 28px;\n      border-radius: 24px;\n      background: linear-gradient(135deg, #263b33 0%, #334d42 100%);\n      color: #ffffff;\n      box-shadow: 0 16px 38px rgba(27, 54, 43, 0.16);\n    }\n\n    .school-contact-row {\n      display: flex;\n      flex-wrap: wrap;\n      align-items: center;\n      gap: 12px 24px;\n    }\n\n    .school-contact-badge {\n      display: inline-flex;\n      align-items: center;\n      min-height: 38px;\n      padding: 0 15px;\n      border-radius: 999px;\n      background: #e0ad4d;\n      color: #253a32;\n      font-size: 14px;\n      font-weight: 900;\n      white-space: nowrap;\n    }\n\n    .school-contact-detail {\n      display: inline-flex;\n      align-items: baseline;\n      gap: 8px;\n      font-size: 16px;\n      line-height: 1.6;\n    }\n\n    .school-contact-detail strong {\n      color: #f4d58c;\n      font-size: 13px;\n      font-weight: 900;\n      letter-spacing: .04em;\n    }\n\n    .site-credit {\n      margin-top: 16px;\n      padding-top: 15px;\n      border-top: 1px solid rgba(255, 255, 255, 0.16);\n      color: rgba(255, 255, 255, 0.68);\n      font-size: 12px;\n      text-align: right;\n    }\n\n    @media (max-width: 640px) {\n      .school-contact-footer {\n        padding: 22px 20px;\n      }\n\n      .school-contact-row,\n      .school-contact-detail {\n        align-items: flex-start;\n      }\n\n      .school-contact-row {\n        flex-direction: column;\n        gap: 10px;\n      }\n\n      .site-credit {\n        text-align: left;\n      }\n    }\n  \n    .upload-section {\n      width: 100%;\n      max-width: 920px;\n      margin: 18px auto 0;\n      padding: 30px;\n      border: 1px solid var(--line);\n      border-radius: 26px;\n      background: rgba(255, 253, 248, 0.98);\n      box-shadow: 0 18px 42px rgba(35, 77, 64, 0.12);\n    }\n\n    .upload-heading {\n      display: flex;\n      align-items: flex-start;\n      justify-content: space-between;\n      gap: 18px;\n      margin-bottom: 22px;\n    }\n\n    .upload-heading h2 {\n      margin: 3px 0 7px;\n      color: var(--forest-dark);\n      font-size: 28px;\n      line-height: 1.25;\n    }\n\n    .upload-heading p {\n      margin: 0;\n      color: var(--sub);\n      font-size: 15px;\n      line-height: 1.7;\n    }\n\n    .upload-badge {\n      flex: 0 0 auto;\n      padding: 9px 13px;\n      border-radius: 999px;\n      background: #fbf0d7;\n      color: #8a6321;\n      font-size: 13px;\n      font-weight: 900;\n    }\n\n    .upload-card {\n      padding: 22px;\n      border: 1px solid #d6e1d9;\n      border-radius: 20px;\n      background: #fff;\n    }\n\n    .upload-dropzone {\n      display: flex;\n      min-height: 185px;\n      align-items: center;\n      justify-content: center;\n      padding: 28px;\n      border: 2px dashed #9ebbaa;\n      border-radius: 18px;\n      background: #f3f8f4;\n      text-align: center;\n      cursor: pointer;\n      transition: border-color 0.18s, background 0.18s, transform 0.18s;\n    }\n\n    .upload-dropzone:hover,\n    .upload-dropzone.dragging {\n      border-color: var(--forest);\n      background: #e8f3ec;\n      transform: translateY(-1px);\n    }\n\n    .upload-dropzone input {\n      position: absolute;\n      width: 1px;\n      height: 1px;\n      overflow: hidden;\n      opacity: 0;\n      pointer-events: none;\n    }\n\n    .upload-icon {\n      display: block;\n      margin-bottom: 8px;\n      font-size: 44px;\n    }\n\n    .upload-dropzone strong {\n      display: block;\n      color: var(--forest-dark);\n      font-size: 20px;\n      line-height: 1.4;\n    }\n\n    .upload-dropzone span {\n      display: block;\n      margin-top: 8px;\n      color: var(--sub);\n      font-size: 14px;\n      line-height: 1.6;\n    }\n\n    .upload-format-note {\n      margin: 14px 2px 0;\n      color: #6d786f;\n      font-size: 13px;\n      line-height: 1.65;\n    }\n\n    .upload-status {\n      display: none;\n      margin-top: 16px;\n      padding: 14px 16px;\n      border-radius: 14px;\n      font-size: 14px;\n      font-weight: 800;\n      line-height: 1.6;\n    }\n\n    .upload-status.show {\n      display: block;\n    }\n\n    .upload-status.info {\n      background: #edf4fb;\n      color: #315f86;\n    }\n\n    .upload-status.success {\n      background: #e9f5ed;\n      color: #286348;\n    }\n\n    .upload-status.error {\n      background: #fff0ef;\n      color: #a9433c;\n    }\n\n    .upload-preview {\n      display: none;\n      margin-top: 18px;\n    }\n\n    .upload-preview.show {\n      display: block;\n    }\n\n    .upload-preview-head {\n      display: flex;\n      align-items: center;\n      justify-content: space-between;\n      gap: 12px;\n      margin-bottom: 10px;\n    }\n\n    .upload-preview-head strong {\n      color: var(--forest-dark);\n      font-size: 17px;\n    }\n\n    .upload-preview-head span {\n      color: var(--sub);\n      font-size: 13px;\n    }\n\n    .upload-table-wrap {\n      max-height: 560px;\n      overflow: auto;\n      scrollbar-gutter: stable;\n      border: 1px solid #d9e2dc;\n      border-radius: 14px;\n      background: #fff;\n    }\n\n    .upload-table {\n      width: 100%;\n      min-width: 760px;\n      border-collapse: collapse;\n      font-size: 13px;\n    }\n\n    .upload-table th,\n    .upload-table td {\n      padding: 11px 12px;\n      border-bottom: 1px solid #e4e9e5;\n      text-align: left;\n      vertical-align: middle;\n    }\n\n    .upload-table th {\n      position: sticky;\n      top: 0;\n      z-index: 1;\n      background: #456354;\n      color: #fff;\n      font-weight: 900;\n      white-space: nowrap;\n    }\n\n    .upload-table td:nth-child(1) {\n      width: 44px;\n      color: #7a847d;\n      text-align: center;\n    }\n\n    .upload-table td:nth-child(2) {\n      min-width: 230px;\n      color: #23372f;\n      font-weight: 800;\n    }\n\n    .upload-actions {\n      display: flex;\n      justify-content: flex-end;\n      gap: 10px;\n      margin-top: 18px;\n    }\n\n    .upload-reset,\n    .upload-save {\n      min-height: 50px;\n      padding: 0 22px;\n      border-radius: 13px;\n      font-size: 15px;\n      font-weight: 900;\n      cursor: pointer;\n    }\n\n    .upload-reset {\n      border: 1px solid #d6dfd8;\n      background: #fff;\n      color: #65716a;\n    }\n\n    .upload-save {\n      border: 0;\n      background: var(--forest);\n      color: #fff;\n    }\n\n    .upload-save:disabled,\n    .upload-reset:disabled {\n      cursor: default;\n      opacity: 0.48;\n    }\n\n    @media (max-width: 720px) {\n      .view-tabs {\n        grid-template-columns: repeat(2, minmax(0, 1fr));\n      }\n\n      .upload-section {\n        padding: 20px 14px;\n      }\n\n      .upload-heading {\n        display: block;\n      }\n\n      .upload-badge {\n        display: inline-block;\n        margin-top: 12px;\n      }\n\n      .upload-actions {\n        flex-direction: column;\n      }\n\n      .upload-reset,\n      .upload-save {\n        width: 100%;\n      }\n    }\n\n\n    /* 모든 주요 영역의 가로폭을 하나로 통일 */\n    .header,\n    .view-tabs,\n    .card,\n    .bs-section,\n    .requests-section,\n    .guide-section,\n    .cart-bar,\n    .footer,\n    .upload-section,\n    .school-contact-footer {\n      width: calc(100% - 36px);\n      max-width: 920px;\n      margin-left: auto;\n      margin-right: auto;\n    }\n\n@media (max-width: 700px) {\n  .header,\n  .view-tabs,\n  .card,\n  .bs-section,\n  .requests-section,\n  .guide-section,\n  .cart-bar,\n  .footer,\n  .upload-section,\n  .school-contact-footer {\n    width: calc(100% - 20px);\n  }\n}\n\n\n    /* ===== 교직원 / 학부모 구분 ===== */\n    .audience-gate{position:fixed;inset:0;z-index:9999;display:flex;align-items:center;justify-content:center;padding:28px 20px;overflow:auto;background:radial-gradient(1200px 700px at 20% 0%,#3a7360 0%,transparent 60%),linear-gradient(160deg,var(--forest),var(--forest-dark))}\n    .audience-gate[hidden]{display:none}\n    .audience-card{width:100%;max-width:980px;text-align:center}\n    .audience-card .eyebrow{margin:0 0 6px;color:var(--gold);letter-spacing:.28em;font-weight:800;font-size:14px}\n    .audience-card h2{margin:6px 0 10px;color:#fff;font-size:clamp(30px,6vw,52px);line-height:1.2}\n    .audience-card p.audience-desc{color:rgba(255,255,255,.82);margin:0 0 30px;line-height:1.6;font-size:clamp(15px,2.4vw,19px)}\n    .audience-choices{display:grid;grid-template-columns:repeat(auto-fit,minmax(260px,1fr));gap:20px}\n    .audience-btn{display:flex;flex-direction:column;align-items:center;justify-content:center;gap:8px;min-height:230px;padding:26px 18px;border:0;border-radius:26px;font-size:clamp(26px,4vw,38px);font-weight:900;cursor:pointer;color:#fff;text-decoration:none;background:linear-gradient(135deg,#3f8068,#1f4f3f);box-shadow:0 16px 40px rgba(0,0,0,.3);transition:transform .15s ease,box-shadow .15s ease}\n    .audience-btn.parent{background:linear-gradient(135deg,var(--gold),var(--gold-dark))}\n    .audience-btn:hover,.audience-btn:focus-visible{transform:translateY(-4px);box-shadow:0 22px 50px rgba(0,0,0,.38);outline:none}\n    .audience-btn .audience-icon{font-size:clamp(46px,7vw,72px);line-height:1}\n    .audience-btn span.audience-sub{font-size:clamp(14px,2.2vw,18px);font-weight:600;opacity:.95}\n    .audience-bar{display:flex;align-items:center;justify-content:center;gap:10px;margin:12px auto 0;font-size:14px}\n    .audience-bar[hidden]{display:none}\n    .audience-badge{padding:6px 15px;border-radius:999px;background:var(--blue-soft);color:var(--forest-dark);font-weight:800}\n    .audience-bar .audience-switch{border:1px solid var(--line);background:#fff;border-radius:999px;padding:6px 13px;cursor:pointer;color:var(--sub);font-size:13px;text-decoration:none}\n    .audience-bar .audience-switch[hidden]{display:none}\n    .parent-fields input,.parent-fields select,#pName{width:100%;box-sizing:border-box;padding:12px;border:1px solid var(--line);border-radius:12px;font-size:16px;background:#fff;color:var(--ink)}\n    .parent-fields label{font-size:13px;color:var(--sub);display:block}\n    .parent-fields input[type=number],.parent-fields select{text-align:center}\n\n    /* ===== 학부모/교직원 화면 정리 ===== */\n    .parent-fields{display:grid;grid-template-columns:repeat(3,1fr);gap:10px}\n    .parent-fields .pname{grid-column:1/-1}\n    .parent-fields label{margin:0 0 5px}\n    .parent-hint{font-weight:400;font-size:13px;color:var(--sub)}\n    #parentApplicantField[hidden],#staffApplicantField[hidden]{display:none}\n    .header .audience-bar{position:absolute;top:20px;right:22px;margin:0;gap:8px;z-index:2}\n    .header .audience-badge{background:rgba(255,255,255,.16);color:#fff;border:1px solid rgba(255,255,255,.3);padding:6px 14px}\n    .header .audience-bar .audience-switch{background:transparent;color:#fff;border:1px solid rgba(255,255,255,.45)}\n    .header .audience-bar .audience-switch:hover{background:rgba(255,255,255,.14)}\n    body.aud-parent .view-tabs{grid-template-columns:repeat(3,minmax(0,1fr))}\n    @media (min-width:760px){\n      body.aud-parent #requestForm{grid-template-columns:1fr 170px}\n      body.aud-parent #parentApplicantField{grid-column:1/-1}\n      .parent-fields{grid-template-columns:repeat(3,minmax(0,.55fr)) minmax(0,1.4fr)}\n      .parent-fields .pname{grid-column:auto}\n    }\n    @media (max-width:620px){\n      .header .audience-bar{position:static;margin:14px 0 0;justify-content:flex-start;flex-wrap:wrap}\n    }\n</style>\n\n  <script src=\"https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js\"></script>\n</head>\n\n<body>\n  <div id=\"audienceGate\" class=\"audience-gate\" hidden>\n    <div class=\"audience-card\" role=\"region\" aria-labelledby=\"audienceGateTitle\">\n      <p class=\"eyebrow\">BOOK REQUEST</p>\n      <h2 id=\"audienceGateTitle\">희망도서 신청</h2>\n      <p class=\"audience-desc\">신청하시는 분을 골라 주세요.<br>교직원과 학부모의 신청은 따로 관리돼요.</p>\n      <div class=\"audience-choices\">\n        <a class=\"audience-btn\" href=\"#\" data-audience=\"staff\"><span class=\"audience-icon\">🧑‍🏫</span>교직원<span class=\"audience-sub\">선생님·교직원 희망도서 신청</span></a>\n        <a class=\"audience-btn parent\" href=\"#\" data-audience=\"parent\"><span class=\"audience-icon\">👨‍👩‍👧</span>학부모<span class=\"audience-sub\">학부모님 희망도서 신청</span></a>\n      </div>\n    </div>\n  </div>\n  <header class=\"header\">\n    <p class=\"eyebrow\">BOOK REQUEST</p>\n    <h1>📚 <?= siteTitle ?></h1>\n\n    <p class=\"description\">\n      책 제목 또는 ISBN을 입력하세요.<br>\n      책 정보와 중복 여부는 자동으로 확인됩니다.\n    </p>\n\n    <div id=\"audienceBar\" class=\"audience-bar\" hidden>\n      <span id=\"audienceBadge\" class=\"audience-badge\"></span>\n      <a id=\"audienceSwitch\" class=\"audience-switch\" href=\"#\" target=\"_top\" hidden>← 처음 화면</a>\n    </div>\n  </header>\n\n\n  <section id=\"hrNoticeTicker\" class=\"hr-notice-ticker\" aria-label=\"도서관 공지 전광판\" aria-live=\"polite\" hidden>\n    <div class=\"hr-notice-inner\">\n      <span class=\"hr-notice-label\">📢 알림</span>\n      <div class=\"hr-notice-window\">\n        <div id=\"hrNoticeTrack\" class=\"hr-notice-track\"></div>\n      </div>\n    </div>\n  </section>\n\n  <nav class=\"view-tabs\" aria-label=\"희망도서 페이지 메뉴\">\n    <button class=\"view-tab active\" type=\"button\" data-panel=\"applyPanel\" aria-selected=\"true\">\n      🔎 도서 신청\n    </button>\n    <button class=\"view-tab\" type=\"button\" data-panel=\"requestsPanel\" aria-selected=\"false\">\n      📋 신청 목록\n    </button>\n    <button class=\"view-tab\" type=\"button\" data-panel=\"uploadPanel\" aria-selected=\"false\">\n      📥 목록 업로드\n    </button>\n    <div id=\"guideHover\" class=\"guide-hover\">\n      <button\n        id=\"guideHoverTrigger\"\n        class=\"guide-hover-trigger\"\n        type=\"button\"\n        aria-expanded=\"false\"\n        aria-describedby=\"guidePopover\"\n      >\n        ❓ 신청 방법\n      </button>\n\n      <div id=\"guidePopover\" class=\"guide-popover\" role=\"tooltip\">\n        <h2>이렇게 신청하세요</h2>\n        <ol>\n          <li><strong>1. 책 검색</strong><span>책 제목 또는 ISBN을 입력하고 「책 검색하기」를 누릅니다.</span></li>\n          <li><strong>2. 신청바구니에 담기</strong><span>표지·저자·출판사·연도·정가를 확인하고 「신청바구니에 담기」를 누릅니다.</span></li>\n          <li><strong>3. 다른 책도 계속 담기</strong><span>검색어를 바꿔도 담은 책은 유지되며, 한 번에 최대 20권까지 담을 수 있습니다.</span></li>\n          <li><strong>4. 한꺼번에 신청</strong><span>화면 아래 신청바구니에서 「선택한 도서 신청하기」를 누릅니다.</span></li>\n          <li><strong>5. 결과 확인·취소</strong><span>「신청 목록」에서 접수 결과와 서지정보를 확인하고, 필요하면 신청을 취소하거나 되돌릴 수 있습니다.</span></li>\n        </ol>\n        <p>희망도서 신청 시 우선 대출 혜택을 받을 수 있습니다.</p>\n      </div>\n    </div>\n  </nav>\n\n  <div id=\"applyPanel\" class=\"page-panel active\">\n\n  <main class=\"card\">\n    <div class=\"form-intro\">\n      <div class=\"form-intro-icon\">📖</div>\n      <div>\n        <h2>어떤 책을 만나고 싶나요?</h2>\n        <p>책 제목이나 ISBN으로 찾고, 표지를 확인한 뒤 신청하세요.</p>\n      </div>\n    </div>\n\n    <form id=\"requestForm\">\n      <div class=\"field\" id=\"staffApplicantField\">\n        <label for=\"applicant\">신청자</label>\n\n        <input\n          id=\"applicant\"\n          name=\"requestApplicant\"\n          type=\"search\"\n          autocomplete=\"off\"\n          autocapitalize=\"off\"\n          spellcheck=\"false\"\n          data-lpignore=\"true\"\n          data-1p-ignore=\"true\"\n          placeholder=\"예: 홍길동\"\n        >\n\n        <p class=\"help\">\n          희망도서 신청 시 우선 대출 혜택을 받을 수 있습니다.\n        </p>\n      </div>\n\n      <div class=\"field\" id=\"parentApplicantField\" hidden>\n        <label>우리 아이 정보 <span class=\"parent-hint\">(학년·반·번호는 꼭 입력해 주세요)</span></label>\n        <div class=\"parent-fields\">\n          <div>\n            <label for=\"pGrade\">학년</label>\n            <select id=\"pGrade\"><option value=\"\">선택</option><option value=\"1\">1</option><option value=\"2\">2</option><option value=\"3\">3</option><option value=\"4\">4</option><option value=\"5\">5</option><option value=\"6\">6</option></select>\n          </div>\n          <div>\n            <label for=\"pClass\">반</label>\n            <input id=\"pClass\" type=\"number\" min=\"1\" max=\"30\" inputmode=\"numeric\" placeholder=\"2\">\n          </div>\n          <div>\n            <label for=\"pNumber\">번호</label>\n            <input id=\"pNumber\" type=\"number\" min=\"1\" max=\"60\" inputmode=\"numeric\" placeholder=\"15\">\n          </div>\n          <div class=\"pname\">\n            <label for=\"pName\">학생 이름 (선택)</label>\n            <input id=\"pName\" type=\"text\" autocomplete=\"off\" placeholder=\"예: 홍길동\">\n          </div>\n        </div>\n        <p class=\"help\">\n          입력한 정보는 신청 확인에만 쓰이고, 신청 목록에는 보이지 않아요.\n        </p>\n      </div>\n\n      <div class=\"field\">\n        <label for=\"query\">책 제목 또는 ISBN</label>\n\n        <input\n          id=\"query\"\n          type=\"search\"\n          autocomplete=\"off\"\n          placeholder=\"예: 아몬드 또는 978...\"\n          required\n        >\n\n        <p class=\"help\">\n          같은 제목의 책이 많다면 ISBN 13자리를 권장합니다.\n        </p>\n      </div>\n\n      <button\n        id=\"submitButton\"\n        class=\"submit-button\"\n        type=\"submit\"\n      >\n        책 검색하기\n      </button>\n    </form>\n\n    <div id=\"message\" class=\"message\"></div>\n\n    <section id=\"searchResults\" class=\"search-results\">\n      <h2 class=\"search-results-title\">\n        신청할 책을 여러 권 선택해 주세요.<br>\n        선택을 마치면 아래 신청 버튼을 누르세요.\n      </h2>\n\n      <div id=\"bookList\" class=\"book-list\"></div>\n    </section>\n  </main>\n\n  <section class=\"bs-section\" aria-labelledby=\"bsTitle\">\n    <div class=\"bs-heading\">\n      <div>\n        <h2 id=\"bsTitle\" class=\"bs-title\">🔥 지금 많이 읽는 책</h2>\n        <p class=\"bs-help\">관심 있는 책을 누르면 정확한 도서를 다시 확인한 뒤 신청할 수 있습니다.</p>\n      </div>\n      <span class=\"bs-badge\">하루 1회 갱신</span>\n    </div>\n\n    <div class=\"bs-tabs\" role=\"tablist\" aria-label=\"베스트셀러 분류\">\n      <button class=\"bs-tab active\" type=\"button\" data-bs-category=\"all\" aria-selected=\"true\">종합</button>\n      <button class=\"bs-tab\" type=\"button\" data-bs-category=\"youth\" aria-selected=\"false\">청소년</button>\n      <button class=\"bs-tab\" type=\"button\" data-bs-category=\"children\" aria-selected=\"false\">어린이</button>\n      <button class=\"bs-tab\" type=\"button\" data-bs-category=\"picture\" aria-selected=\"false\">그림책</button>\n    </div>\n\n    <div id=\"bsStatus\" class=\"bs-status show\">베스트셀러를 불러오고 있습니다…</div>\n    <div class=\"bs-carousel\">\n      <button id=\"bsPrevious\" class=\"bs-arrow bs-arrow-previous\" type=\"button\" aria-label=\"이전 베스트셀러\">‹</button>\n      <div id=\"bsList\" class=\"bs-list\"></div>\n      <button id=\"bsNext\" class=\"bs-arrow bs-arrow-next\" type=\"button\" aria-label=\"다음 베스트셀러\">›</button>\n    </div>\n\n\n    <div class=\"bs-more-wrap\">\n      <a\n        class=\"bs-more\"\n        href=\"https://www.yes24.com/Product/Category/BestSeller\"\n        target=\"_blank\"\n        rel=\"noopener noreferrer\"\n      >\n        YES24에서 베스트셀러 더 보기 →\n      </a>\n    </div>\n\n    <p class=\"bs-source\">\n      도서 정보 제공:\n      <a href=\"https://www.yes24.com/Product/Category/BestSeller\" target=\"_blank\" rel=\"noopener noreferrer\">YES24 인터넷서점</a>\n    </p>\n  </section>\n\n  <section id=\"cartBar\" class=\"cart-bar\" aria-label=\"신청바구니\">\n    <button\n      id=\"cartToggle\"\n      class=\"cart-toggle\"\n      type=\"button\"\n      aria-expanded=\"false\"\n      aria-controls=\"cartDrawer\"\n    >\n      <span class=\"cart-count\">\n        🛒 신청바구니 <strong id=\"cartCount\">0</strong>/20권\n      </span>\n      <span class=\"cart-open-hint\">담은 책 보기</span>\n    </button>\n\n    <button\n      id=\"cartSubmit\"\n      class=\"cart-submit\"\n      type=\"button\"\n      disabled\n    >\n      선택한 도서 신청하기\n    </button>\n  </section>\n\n  <div\n    id=\"cartDrawer\"\n    class=\"cart-drawer\"\n    role=\"dialog\"\n    aria-modal=\"true\"\n    aria-labelledby=\"cartDrawerTitle\"\n  >\n    <button\n      class=\"cart-backdrop\"\n      type=\"button\"\n      data-cart-close\n      aria-label=\"신청바구니 닫기\"\n    ></button>\n\n    <section class=\"cart-panel\">\n      <div class=\"cart-panel-header\">\n        <div>\n          <p>BOOK CART</p>\n          <h2 id=\"cartDrawerTitle\">🛒 신청바구니</h2>\n        </div>\n        <button\n          id=\"cartClose\"\n          class=\"cart-close\"\n          type=\"button\"\n          aria-label=\"신청바구니 닫기\"\n        >×</button>\n      </div>\n\n      <p class=\"cart-guide\">\n        다른 책을 다시 검색해도 그대로 유지됩니다. 한 번에 최대 20권까지 담을 수 있습니다.\n      </p>\n\n      <div id=\"cartList\" class=\"cart-list\"></div>\n\n      <div class=\"cart-panel-actions\">\n        <button id=\"cartClear\" class=\"cart-clear\" type=\"button\">\n          전체 비우기\n        </button>\n        <button\n          id=\"cartDrawerSubmit\"\n          class=\"cart-submit cart-drawer-submit\"\n          type=\"button\"\n          disabled\n        >\n          담은 도서 신청하기\n        </button>\n      </div>\n    </section>\n  </div>\n\n  </div>\n\n  <section\n    id=\"requestsPanel\"\n    class=\"page-panel requests-section\"\n    aria-labelledby=\"requestsTitle\"\n  >\n    <div class=\"panel-heading\">\n      <div>\n        <p class=\"panel-eyebrow\">REQUEST LIST</p>\n        <h2 id=\"requestsTitle\">📋 신청된 희망도서</h2>\n        <p>신청자 이름은 표시하지 않고 도서 정보만 보여줍니다.</p>\n      </div>\n      <button id=\"requestRefresh\" class=\"refresh-button\" type=\"button\">\n        새로고침\n      </button>\n    </div>\n\n    <div\n      class=\"request-filter-tabs\"\n      role=\"tablist\"\n      aria-label=\"신청 목록 구분\"\n    >\n      <button\n        class=\"request-filter-tab active\"\n        type=\"button\"\n        data-request-filter=\"active\"\n        role=\"tab\"\n        aria-selected=\"true\"\n      >\n        <span>📚 신청 도서</span>\n        <strong id=\"requestActiveCount\">0</strong>\n      </button>\n      <button\n        class=\"request-filter-tab\"\n        type=\"button\"\n        data-request-filter=\"cancelled\"\n        role=\"tab\"\n        aria-selected=\"false\"\n      >\n        <span>↩️ 신청 취소</span>\n        <strong id=\"requestCancelledCount\">0</strong>\n      </button>\n    </div>\n\n    <div id=\"requestSummary\" class=\"request-summary\"></div>\n    <div id=\"requestListStatus\" class=\"request-list-status\">\n      신청 목록 탭을 누르면 최신 목록을 불러옵니다.\n    </div>\n    <div id=\"requestList\" class=\"request-list\"></div>\n  </section>\n\n\n  <section\n    id=\"uploadPanel\"\n    class=\"page-panel upload-section\"\n    aria-labelledby=\"uploadTitle\"\n  >\n    <div class=\"upload-heading\">\n      <div>\n        <p class=\"panel-eyebrow\">BOOK LIST IMPORT</p>\n        <h2 id=\"uploadTitle\">📥 도서 목록 파일 업로드</h2>\n        <p>\n          도서 목록 엑셀·CSV 파일을 선택하면<br>\n          도서명과 ISBN 등의 열을 자동으로 찾아 업무용 목록으로 정리합니다.\n        </p>\n      </div>\n      <span class=\"upload-badge\">사서 업무용</span>\n    </div>\n\n    <div class=\"upload-card\">\n      <label id=\"uploadDropzone\" class=\"upload-dropzone\" for=\"uploadFile\">\n        <input\n          id=\"uploadFile\"\n          type=\"file\"\n          accept=\".xls,.xlsx,.xlsm,.csv,.tsv,.html,.htm\"\n        >\n        <span>\n          <span class=\"upload-icon\">📄</span>\n          <strong>엑셀·CSV 파일을 선택하거나 여기에 놓으세요</strong>\n          <span>.xls · .xlsx · .xlsm · .csv · .tsv / 한 번에 최대 1,500권</span>\n        </span>\n      </label>\n\n      <p class=\"upload-format-note\">\n        도서명 또는 ISBN 열이 있어야 합니다. 저자·출판사·출판연도·정가·수량·시리즈명은\n        파일에 있는 항목만 가져오며, 표지와 파일 작성자 정보는 가져오지 않습니다.\n      </p>\n\n      <div id=\"uploadStatus\" class=\"upload-status\" role=\"status\"></div>\n\n      <div id=\"uploadPreview\" class=\"upload-preview\">\n        <div class=\"upload-preview-head\">\n          <strong>가져올 내용 미리보기</strong>\n          <span id=\"uploadPreviewCount\"></span>\n        </div>\n        <div class=\"upload-table-wrap\">\n          <table class=\"upload-table\">\n            <thead>\n              <tr>\n                <th>번호</th>\n                <th>도서명</th>\n                <th>저자</th>\n                <th>출판사</th>\n                <th>연도</th>\n                <th>ISBN13</th>\n                <th>정가</th>\n              </tr>\n            </thead>\n            <tbody id=\"uploadPreviewBody\"></tbody>\n          </table>\n        </div>\n      </div>\n\n      <div class=\"upload-actions\">\n        <button id=\"uploadReset\" class=\"upload-reset\" type=\"button\" disabled>\n          다시 선택\n        </button>\n        <button id=\"uploadSave\" class=\"upload-save\" type=\"button\" disabled>\n          이 목록을 시트에 넣기\n        </button>\n      </div>\n    </div>\n  </section>\n\n\n  <p class=\"footer\">\n    처리에는 몇 초 정도 걸릴 수 있습니다.<br>\n    창을 닫지 말고 완료 안내를 확인해 주세요.\n  </p>\n\n  <footer class=\"school-contact-footer\">\n    <? if (managerName || managerExtension) { ?>\n      <div class=\"school-contact-row\">\n        <span class=\"school-contact-badge\">📚 도서관 문의</span>\n\n        <? if (managerName) { ?>\n          <span class=\"school-contact-detail\">\n            <strong>담당자</strong>\n            <span><?= managerName ?></span>\n          </span>\n        <? } ?>\n\n        <? if (managerExtension) { ?>\n          <span class=\"school-contact-detail\">\n            <strong>내선</strong>\n            <span><?= managerExtension ?></span>\n          </span>\n        <? } ?>\n      </div>\n    <? } ?>\n\n    <div class=\"site-credit\">\n      학교도서관 도서구입 지원 도구    </div>\n  </footer>\n\n  <div\n    id=\"noticeModal\"\n    class=\"notice-modal\"\n    role=\"dialog\"\n    aria-modal=\"true\"\n    aria-labelledby=\"noticeTitle\"\n  >\n    <div class=\"notice-dialog\">\n      <div class=\"notice-visual\">\n        <div id=\"noticeIcon\" class=\"notice-icon\">📚</div>\n      </div>\n\n      <div class=\"notice-content\">\n        <h2 id=\"noticeTitle\" class=\"notice-title\"></h2>\n        <p id=\"noticeBook\" class=\"notice-book\"></p>\n        <p id=\"noticeDescription\" class=\"notice-description\"></p>\n\n        <div\n          id=\"noticeActions\"\n          class=\"notice-actions single\"\n        >\n          <button\n            id=\"noticeUndo\"\n            class=\"notice-undo\"\n            type=\"button\"\n            hidden\n          >\n            신청 취소 되돌리기\n          </button>\n          <button\n            id=\"noticeClose\"\n            class=\"notice-close\"\n            type=\"button\"\n          >\n            확인\n          </button>\n        </div>\n      </div>\n    </div>\n  </div>\n\n\n  <div\n    id=\"cancelModal\"\n    class=\"cancel-modal\"\n    role=\"dialog\"\n    aria-modal=\"true\"\n    aria-labelledby=\"cancelTitle\"\n  >\n    <div class=\"cancel-dialog\">\n      <div class=\"cancel-icon\">↩️</div>\n      <h2 id=\"cancelTitle\">희망도서 신청을 취소할까요?</h2>\n      <p id=\"cancelBook\" class=\"cancel-book\"></p>\n      <p class=\"cancel-description\">\n        취소하면 신청 목록에서 이 도서 한 권만 삭제됩니다. 바로 되돌릴 수 있도록 안전하게 보관합니다.\n      </p>\n      <div class=\"cancel-actions\">\n        <button id=\"cancelKeep\" class=\"cancel-keep\" type=\"button\">\n          그대로 두기\n        </button>\n        <button id=\"cancelConfirm\" class=\"cancel-confirm\" type=\"button\">\n          신청 취소하기\n        </button>\n      </div>\n    </div>\n  </div>\n\n  <script>\n    const form =\n      document.getElementById('requestForm');\n    const applicant =\n      document.getElementById('applicant');\n    const query =\n      document.getElementById('query');\n    const submitButton =\n      document.getElementById('submitButton');\n    const message =\n      document.getElementById('message');\n    const searchResults =\n      document.getElementById('searchResults');\n    const bookList =\n      document.getElementById('bookList');\n    const noticeModal =\n      document.getElementById('noticeModal');\n    const noticeIcon =\n      document.getElementById('noticeIcon');\n    const noticeTitle =\n      document.getElementById('noticeTitle');\n    const noticeBook =\n      document.getElementById('noticeBook');\n    const noticeDescription =\n      document.getElementById('noticeDescription');\n    const noticeActions =\n      document.getElementById('noticeActions');\n    const noticeUndo =\n      document.getElementById('noticeUndo');\n    const noticeClose =\n      document.getElementById('noticeClose');\n    const cartBar =\n      document.getElementById('cartBar');\n    const cartCount =\n      document.getElementById('cartCount');\n    const cartSubmit =\n      document.getElementById('cartSubmit');\n    const viewTabs =\n      Array.from(document.querySelectorAll('.view-tab'));\n    const pagePanels =\n      Array.from(document.querySelectorAll('.page-panel'));\n    const applyPanel =\n      document.getElementById('applyPanel');\n    const requestRefresh =\n      document.getElementById('requestRefresh');\n    const requestSummary =\n      document.getElementById('requestSummary');\n    const requestListStatus =\n      document.getElementById('requestListStatus');\n    const requestList =\n      document.getElementById('requestList');\n    const requestFilterTabs =\n      Array.from(\n        document.querySelectorAll('.request-filter-tab')\n      );\n    const requestActiveCount =\n      document.getElementById('requestActiveCount');\n    const requestCancelledCount =\n      document.getElementById('requestCancelledCount');\n    const uploadFile =\n      document.getElementById('uploadFile');\n    const uploadDropzone =\n      document.getElementById('uploadDropzone');\n    const uploadStatus =\n      document.getElementById('uploadStatus');\n    const uploadPreview =\n      document.getElementById('uploadPreview');\n    const uploadPreviewCount =\n      document.getElementById('uploadPreviewCount');\n    const uploadPreviewBody =\n      document.getElementById('uploadPreviewBody');\n    const uploadReset =\n      document.getElementById('uploadReset');\n    const uploadSave =\n      document.getElementById('uploadSave');\n    let uploadParsedBooks = [];\n    let uploadFileName = '';\n    const bsListElement =\n      document.getElementById('bsList');\n    const bsPrevious =\n      document.getElementById('bsPrevious');\n    const bsNext =\n      document.getElementById('bsNext');\n    const guideHover =\n      document.getElementById('guideHover');\n    const guideHoverTrigger =\n      document.getElementById('guideHoverTrigger');\n    const cancelModal =\n      document.getElementById('cancelModal');\n    const cancelBook =\n      document.getElementById('cancelBook');\n    const cancelKeep =\n      document.getElementById('cancelKeep');\n    const cancelConfirm =\n      document.getElementById('cancelConfirm');\n    const cartToggle =\n      document.getElementById('cartToggle');\n    const cartDrawer =\n      document.getElementById('cartDrawer');\n    const cartClose =\n      document.getElementById('cartClose');\n    const cartList =\n      document.getElementById('cartList');\n    const cartClear =\n      document.getElementById('cartClear');\n    const cartDrawerSubmit =\n      document.getElementById('cartDrawerSubmit');\n    const cartBackdrop =\n      document.querySelector('[data-cart-close]');\n    let books = [];\n    let cartBooks = new Map();\n    let submissionInProgress = false;\n    let submittedRequestsLoaded = false;\n    let submittedRequestItems = [];\n    let requestListFilter = 'active';\n    let bsAutoTimer = null;\n    let cancelCandidate = null;\n    let undoCandidate = null;\n\n    noticeClose.addEventListener('click', closeNoticeModal);\n    noticeUndo.addEventListener('click', undoCancelledRequest);\n    cartSubmit.addEventListener('click', submitCart);\n    cartDrawerSubmit.addEventListener('click', submitCart);\n    cartToggle.addEventListener('click', openCartDrawer);\n    cartClose.addEventListener('click', closeCartDrawer);\n    cartBackdrop.addEventListener('click', closeCartDrawer);\n    cartClear.addEventListener('click', clearCart);\n\n    document.addEventListener('keydown', function(event) {\n      if (event.key === 'Escape') {\n        closeCartDrawer();\n      }\n    });\n\n    viewTabs.forEach(function(tab) {\n      tab.addEventListener('click', function() {\n        showPagePanel(tab.dataset.panel || 'applyPanel');\n      });\n    });\n\n    requestRefresh.addEventListener('click', function() {\n      loadSubmittedRequests(true);\n    });\n\n    requestFilterTabs.forEach(function(tab) {\n      tab.addEventListener('click', function() {\n        requestListFilter =\n          tab.dataset.requestFilter || 'active';\n\n        renderSubmittedRequests({\n          items: submittedRequestItems\n        });\n      });\n    });\n\n    guideHoverTrigger.addEventListener('click', function(event) {\n      event.stopPropagation();\n      const opened = guideHover.classList.toggle('open');\n      guideHoverTrigger.setAttribute(\n        'aria-expanded',\n        opened ? 'true' : 'false'\n      );\n    });\n\n    document.addEventListener('click', function(event) {\n      if (!guideHover.contains(event.target)) {\n        closeGuidePopover();\n      }\n    });\n\n    cancelKeep.addEventListener('click', closeCancelModal);\n    cancelConfirm.addEventListener(\n      'click',\n      confirmCancelRequest\n    );\n\n    cancelModal.addEventListener('click', function(event) {\n      if (event.target === cancelModal) {\n        closeCancelModal();\n      }\n    });\n\n    bsPrevious.addEventListener('click', function() {\n      scrollBestsellerCards(-1);\n      startBestsellerAuto();\n    });\n\n    bsNext.addEventListener('click', function() {\n      scrollBestsellerCards(1);\n      startBestsellerAuto();\n    });\n\n    bsListElement.addEventListener(\n      'mouseenter',\n      stopBestsellerAuto\n    );\n    bsListElement.addEventListener(\n      'mouseleave',\n      startBestsellerAuto\n    );\n    bsListElement.addEventListener(\n      'focusin',\n      stopBestsellerAuto\n    );\n    bsListElement.addEventListener(\n      'focusout',\n      startBestsellerAuto\n    );\n    bsListElement.addEventListener(\n      'touchstart',\n      stopBestsellerAuto,\n      { passive: true }\n    );\n    bsListElement.addEventListener(\n      'touchend',\n      function() {\n        window.setTimeout(startBestsellerAuto, 1200);\n      },\n      { passive: true }\n    );\n\n    document.addEventListener('visibilitychange', function() {\n      if (document.hidden) {\n        stopBestsellerAuto();\n      } else {\n        startBestsellerAuto();\n      }\n    });\n\n    noticeModal.addEventListener('click', function(event) {\n      if (event.target === noticeModal) {\n        closeNoticeModal();\n      }\n    });\n\n    document.addEventListener('keydown', function(event) {\n      if (event.key !== 'Escape') {\n        return;\n      }\n\n      if (noticeModal.classList.contains('show')) {\n        closeNoticeModal();\n      }\n\n      if (cancelModal.classList.contains('show')) {\n        closeCancelModal();\n      }\n\n      closeGuidePopover();\n    });\n\n\n\n    function setUploadStatus(message, type) {\n      uploadStatus.textContent = String(message || '');\n      uploadStatus.className = 'upload-status show ' + (type || 'info');\n    }\n\n    function resetUploadPanel() {\n      uploadParsedBooks = [];\n      uploadFileName = '';\n      uploadFile.value = '';\n      uploadPreviewBody.innerHTML = '';\n      uploadPreviewCount.textContent = '';\n      uploadPreview.classList.remove('show');\n      uploadStatus.className = 'upload-status';\n      uploadStatus.textContent = '';\n      uploadSave.disabled = true;\n      uploadReset.disabled = true;\n    }\n\n    function normalizeUploadHeader(value) {\n      return String(value == null ? '' : value)\n        .replace(/<[^>]*>/g, '')\n        .replace(/[\\s\\n\\r()（）\\[\\]{}·._\\-\\/]/g, '')\n        .toUpperCase();\n    }\n\n    function findUploadHeaderKey(value) {\n      const normalized = normalizeUploadHeader(value);\n      const aliases = {\n        title: ['상품명','도서명','책제목','서명','제목'],\n        isbn: ['ISBN13','ISBN','국제표준도서번호'],\n        author: ['저자아티스트','저자','작가','글쓴이','지은이'],\n        publisher: ['출판사제작사','출판사','발행자','출판자'],\n        year: ['출간일','출판연도','발행연도','출간연도','연도'],\n        price: ['정가','도서정가','정가원','가격'],\n        quantity: ['수량','주문수량','권수'],\n        series: ['시리즈명','시리즈'],\n        note: ['비고','메모','구분','분류']\n      };\n\n      for (const key in aliases) {\n        if (aliases[key].indexOf(normalized) !== -1) {\n          return key;\n        }\n      }\n      return '';\n    }\n\n    function cleanUploadIsbn(value) {\n      let text = String(value == null ? '' : value).trim();\n\n      if (/^[0-9.]+E\\+?[0-9]+$/i.test(text)) {\n        const numeric = Number(text);\n        if (Number.isFinite(numeric)) {\n          text = numeric.toFixed(0);\n        }\n      }\n\n      return text.replace(/[^0-9Xx]/g, '').toUpperCase();\n    }\n\n    function cleanUploadYear(value) {\n      const match = String(value == null ? '' : value).match(/(?:19|20)\\d{2}/);\n      return match ? match[0] : '';\n    }\n\n    function cleanUploadNumber(value, fallback) {\n      const numeric = Number(\n        String(value == null ? '' : value)\n          .replace(/[^0-9.-]/g, '')\n      );\n      return Number.isFinite(numeric) && numeric > 0\n        ? numeric\n        : (fallback || 0);\n    }\n\n    function decodeUploadText(buffer) {\n      const sample = new TextDecoder('latin1').decode(buffer.slice(0, 4096));\n      const charsetMatch = sample.match(/charset\\s*=\\s*[\"']?([^\"'\\s;>]+)/i);\n      const charset = charsetMatch ? charsetMatch[1].toLowerCase() : 'utf-8';\n\n      try {\n        return new TextDecoder(charset).decode(buffer);\n      } catch (error) {\n        return new TextDecoder('utf-8').decode(buffer);\n      }\n    }\n\n    function htmlMatricesFromText(text) {\n      if (!/<table[\\s>]/i.test(text)) {\n        return [];\n      }\n\n      const documentNode = new DOMParser().parseFromString(text, 'text/html');\n      return Array.from(documentNode.querySelectorAll('table')).map(function(table) {\n        return Array.from(table.querySelectorAll('tr')).map(function(row) {\n          return Array.from(row.querySelectorAll('th,td')).map(function(cell) {\n            return cell.textContent.replace(/\\u00a0/g, ' ').trim();\n          });\n        });\n      });\n    }\n\n    function delimitedMatrixFromText(text, delimiter) {\n      const rows = [];\n      let row = [];\n      let cell = '';\n      let quoted = false;\n\n      for (let index = 0; index < text.length; index++) {\n        const character = text[index];\n\n        if (quoted) {\n          if (character === '\"' && text[index + 1] === '\"') {\n            cell += '\"';\n            index++;\n          } else if (character === '\"') {\n            quoted = false;\n          } else {\n            cell += character;\n          }\n          continue;\n        }\n\n        if (character === '\"') {\n          quoted = true;\n        } else if (character === delimiter) {\n          row.push(cell);\n          cell = '';\n        } else if (character === '\\n') {\n          row.push(cell.replace(/\\r$/, ''));\n          rows.push(row);\n          row = [];\n          cell = '';\n        } else {\n          cell += character;\n        }\n      }\n\n      if (cell || row.length) {\n        row.push(cell.replace(/\\r$/, ''));\n        rows.push(row);\n      }\n      return rows;\n    }\n\n    function workbookMatricesFromBuffer(file, buffer) {\n      const extension = (file.name.split('.').pop() || '').toLowerCase();\n      const text = decodeUploadText(buffer);\n      const htmlMatrices = htmlMatricesFromText(text);\n\n      if (htmlMatrices.length) {\n        return htmlMatrices;\n      }\n\n      if (extension === 'csv' || extension === 'tsv') {\n        return [\n          delimitedMatrixFromText(\n            text.replace(/^\\uFEFF/, ''),\n            extension === 'tsv' ? '\\t' : ','\n          )\n        ];\n      }\n\n      if (typeof XLSX === 'undefined') {\n        throw new Error(\n          '엑셀 읽기 도구를 불러오지 못했습니다. 인터넷 연결을 확인한 뒤 다시 시도해 주세요.'\n        );\n      }\n\n      const workbook = XLSX.read(buffer, {\n        type: 'array',\n        raw: false,\n        cellDates: false\n      });\n\n      return workbook.SheetNames.map(function(sheetName) {\n        return XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], {\n          header: 1,\n          defval: '',\n          raw: false,\n          blankrows: false\n        });\n      });\n    }\n\n    function locateUploadTable(matrices) {\n      let best = null;\n\n      matrices.forEach(function(matrix) {\n        const rowLimit = Math.min(matrix.length, 50);\n\n        for (let rowIndex = 0; rowIndex < rowLimit; rowIndex++) {\n          const row = Array.isArray(matrix[rowIndex]) ? matrix[rowIndex] : [];\n          const map = {};\n          let score = 0;\n\n          row.forEach(function(value, columnIndex) {\n            const key = findUploadHeaderKey(value);\n            if (key && (map[key] === undefined || (key === 'isbn' && normalizeUploadHeader(value) === 'ISBN13'))) {\n              map[key] = columnIndex;\n              score++;\n            }\n          });\n\n          const essential = map.title !== undefined || map.isbn !== undefined;\n          if (essential && score >= 2 && (!best || score > best.score)) {\n            best = {\n              matrix: matrix,\n              headerRow: rowIndex,\n              map: map,\n              score: score\n            };\n          }\n        }\n      });\n\n      if (!best) {\n        throw new Error(\n          '도서명 또는 ISBN 열을 찾지 못했습니다. 첫 행의 열 이름을 확인해 주세요.'\n        );\n      }\n      return best;\n    }\n\n    function uploadBooksFromTable(table) {\n      const books = [];\n      const seen = {};\n      let skipped = 0;\n\n      for (\n        let rowIndex = table.headerRow + 1;\n        rowIndex < table.matrix.length;\n        rowIndex++\n      ) {\n        const row = Array.isArray(table.matrix[rowIndex])\n          ? table.matrix[rowIndex]\n          : [];\n\n        const valueAt = function(key) {\n          const columnIndex = table.map[key];\n          return columnIndex === undefined ? '' : row[columnIndex];\n        };\n\n        const title = String(valueAt('title') || '')\n          .replace(/<[^>]*>/g, '')\n          .replace(/\\s+/g, ' ')\n          .trim();\n        const isbn = cleanUploadIsbn(valueAt('isbn'));\n\n        if (!title && !isbn) {\n          continue;\n        }\n\n        const publisher = String(valueAt('publisher') || '').trim();\n        const year = cleanUploadYear(valueAt('year'));\n        const key = isbn\n          ? 'isbn:' + isbn\n          : 'title:' +\n            title.replace(/\\s+/g, '').toUpperCase() + '|' +\n            publisher.replace(/\\s+/g, '').toUpperCase() + '|' +\n            year;\n\n        if (seen[key]) {\n          skipped++;\n          continue;\n        }\n        seen[key] = true;\n\n        books.push({\n          title: title,\n          author: String(valueAt('author') || '').trim(),\n          publisher: publisher,\n          year: year,\n          isbn: isbn,\n          price: cleanUploadNumber(valueAt('price'), 0),\n          quantity: cleanUploadNumber(valueAt('quantity'), 1),\n          series: String(valueAt('series') || '').trim(),\n          note: String(valueAt('note') || '').trim()\n        });\n      }\n\n      return {\n        books: books.slice(0, 1500),\n        skipped: skipped + Math.max(0, books.length - 1500),\n        limited: books.length > 1500\n      };\n    }\n\n    function renderUploadPreview(result) {\n      uploadPreviewBody.innerHTML = '';\n\n      const fragment = document.createDocumentFragment();\n\n      result.books.forEach(function(book, index) {\n        const row = document.createElement('tr');\n        const cells = [\n          String(index + 1),\n          book.title || '(제목 없음)',\n          book.author || '',\n          book.publisher || '',\n          book.year || '',\n          book.isbn || '',\n          book.price ? book.price.toLocaleString('ko-KR') + '원' : ''\n        ];\n\n        cells.forEach(function(value) {\n          const cell = document.createElement('td');\n          cell.textContent = value;\n          row.appendChild(cell);\n        });\n\n        fragment.appendChild(row);\n      });\n\n      uploadPreviewBody.appendChild(fragment);\n      uploadPreviewCount.textContent =\n        '총 ' + result.books.length + '권 · 전체 ' + result.books.length + '권 표시';\n      uploadPreview.classList.add('show');\n    }\n\n    async function handleUploadFile(file) {\n      resetUploadPanel();\n\n      if (!file) {\n        return;\n      }\n\n      if (file.size > 15 * 1024 * 1024) {\n        setUploadStatus('파일 크기는 15MB 이하만 가능합니다.', 'error');\n        return;\n      }\n\n      uploadFileName = file.name;\n      uploadReset.disabled = false;\n      setUploadStatus('파일을 읽고 열 이름을 확인하고 있습니다…', 'info');\n\n      try {\n        const buffer = await file.arrayBuffer();\n        const matrices = workbookMatricesFromBuffer(file, buffer);\n        const table = locateUploadTable(matrices);\n        const result = uploadBooksFromTable(table);\n\n        if (!result.books.length) {\n          throw new Error('도서명 또는 ISBN이 있는 도서 행을 찾지 못했습니다.');\n        }\n\n        uploadParsedBooks = result.books;\n        renderUploadPreview(result);\n        uploadSave.disabled = false;\n\n        let message =\n          file.name + '에서 ' + result.books.length + '권을 찾았습니다.';\n        if (result.skipped) {\n          message += ' 중복·빈 행 ' + result.skipped + '개는 제외했습니다.';\n        }\n        if (result.limited) {\n          message += ' 한 번에 최대 1,500권까지만 표시합니다.';\n        }\n        setUploadStatus(message, 'success');\n      } catch (error) {\n        uploadParsedBooks = [];\n        uploadSave.disabled = true;\n        setUploadStatus(\n          error && error.message\n            ? error.message\n            : '파일을 읽지 못했습니다.',\n          'error'\n        );\n      }\n    }\n\n    uploadFile.addEventListener('change', function() {\n      handleUploadFile(uploadFile.files && uploadFile.files[0]);\n    });\n\n    uploadDropzone.addEventListener('dragover', function(event) {\n      event.preventDefault();\n      uploadDropzone.classList.add('dragging');\n    });\n\n    uploadDropzone.addEventListener('dragleave', function() {\n      uploadDropzone.classList.remove('dragging');\n    });\n\n    uploadDropzone.addEventListener('drop', function(event) {\n      event.preventDefault();\n      uploadDropzone.classList.remove('dragging');\n      const file = event.dataTransfer.files && event.dataTransfer.files[0];\n      handleUploadFile(file);\n    });\n\n    uploadReset.addEventListener('click', function() {\n      resetUploadPanel();\n      uploadFile.click();\n    });\n\n    uploadSave.addEventListener('click', function() {\n      if (!uploadParsedBooks.length) {\n        setUploadStatus('먼저 파일을 선택해 주세요.', 'error');\n        return;\n      }\n\n      uploadSave.disabled = true;\n      uploadReset.disabled = true;\n      uploadSave.textContent = '시트에 넣는 중…';\n      setUploadStatus(\n        '업무용 업로드 목록 시트를 만들고 있습니다. 창을 닫지 마세요.',\n        'info'\n      );\n\n      google.script.run\n        .withSuccessHandler(function(result) {\n          uploadSave.textContent = '시트 입력 완료';\n          uploadReset.disabled = false;\n          setUploadStatus(\n            '완료: ' + result.importedCount +\n            '권을 「' + result.sheetName +\n            '」 시트에 넣었습니다. 스프레드시트에서 구입할 책을 체크한 뒤 구입 검토함으로 보내세요.',\n            'success'\n          );\n        })\n        .withFailureHandler(function(error) {\n          uploadSave.disabled = false;\n          uploadReset.disabled = false;\n          uploadSave.textContent = '이 목록을 시트에 넣기';\n          setUploadStatus(\n            error && error.message\n              ? error.message\n              : '시트에 저장하지 못했습니다.',\n            'error'\n          );\n        })\n        .HR_importBookListFile({\n          fileName: uploadFileName,\n          books: uploadParsedBooks\n        });\n    });\n\n    function showPagePanel(panelId) {\n      pagePanels.forEach(function(panel) {\n        panel.classList.toggle(\n          'active',\n          panel.id === panelId\n        );\n      });\n\n      viewTabs.forEach(function(tab) {\n        const active = tab.dataset.panel === panelId;\n\n        tab.classList.toggle('active', active);\n        tab.setAttribute(\n          'aria-selected',\n          active ? 'true' : 'false'\n        );\n      });\n\n      if (panelId === 'requestsPanel') {\n        stopBestsellerAuto();\n        loadSubmittedRequests(false);\n      } else if (panelId === 'applyPanel') {\n        startBestsellerAuto();\n      } else {\n        stopBestsellerAuto();\n      }\n\n      document.querySelector('.view-tabs').scrollIntoView({\n        behavior: 'smooth',\n        block: 'start'\n      });\n    }\n\n    form.addEventListener('submit', function(event) {\n      event.preventDefault();\n\n      const applicantValue = applicant.value.trim();\n      const queryValue = query.value.trim();\n\n      if (!queryValue) {\n        showMessage('책 제목이나 ISBN을 입력해 주세요.', 'error');\n        query.focus();\n        return;\n      }\n\n      searchResults.classList.remove('show');\n      bookList.innerHTML = '';\n      books = [];\n      updateCartBar();\n      setSearchLoading(true);\n\n      showMessage(\n        'YES24에서 책을 찾고 있습니다…',\n        'info'\n      );\n\n      google.script.run\n        .withSuccessHandler(function(responseBooks) {\n          setSearchLoading(false);\n\n          books = Array.isArray(responseBooks)\n            ? responseBooks\n            : [];\n\n          if (!books.length) {\n            showMessage(\n              '검색 결과가 없습니다. 제목이나 ISBN을 다시 확인해 주세요.',\n              'error'\n            );\n            return;\n          }\n\n          renderBooks();\n          searchResults.classList.add('show');\n\n          showMessage(\n            '검색 결과에서 신청할 책을 확인해 주세요.',\n            'info'\n          );\n        })\n        .withFailureHandler(function(error) {\n          setSearchLoading(false);\n          showMessage(getErrorMessage(error), 'error');\n        })\n        .HR_searchRequestBooks(queryValue);\n    });\n\n    function getBookKey(book) {\n      const isbn = String(book && book.isbn || '')\n        .replace(/[^0-9Xx]/g, '')\n        .toUpperCase();\n\n      if (isbn) {\n        return 'isbn:' + isbn;\n      }\n\n      return 'book:' + [\n        book && book.title,\n        book && book.publisher,\n        book && (book.year || book.pubDate)\n      ].map(function(value) {\n        return String(value || '')\n          .trim()\n          .toLowerCase()\n          .replace(/\\s+/g, ' ');\n      }).join('|');\n    }\n\n    function findBookButton(bookKey) {\n      return Array.from(\n        document.querySelectorAll('.select-book-button')\n      ).find(function(button) {\n        return button.dataset.bookKey === bookKey;\n      }) || null;\n    }\n\n    function setVisibleBookSelection(button, selected) {\n      if (!button || button.dataset.done === 'true') {\n        return;\n      }\n\n      button.textContent = selected\n        ? '장바구니에 담김 ✓'\n        : '신청바구니에 담기';\n      button.classList.toggle('selected', selected);\n\n      if (button.closest('.book-card')) {\n        button.closest('.book-card')\n          .classList.toggle('selected', selected);\n      }\n    }\n\n    function renderBooks() {\n      bookList.innerHTML = '';\n\n      books.forEach(function(book, index) {\n        const card = document.createElement('article');\n        const cover = document.createElement('div');\n        const info = document.createElement('div');\n        const title = document.createElement('h3');\n        const meta = document.createElement('p');\n        const button = document.createElement('button');\n        const coverUrl = safeUrl(book.cover);\n        const bookKey = getBookKey(book);\n        const isInCart = cartBooks.has(bookKey);\n\n        card.className = 'book-card';\n        cover.className = 'book-cover';\n        info.className = 'book-info';\n        title.className = 'book-title';\n        meta.className = 'book-meta';\n        button.className = 'select-book-button';\n        button.type = 'button';\n        button.dataset.index = String(index);\n        button.dataset.bookKey = bookKey;\n\n        if (coverUrl) {\n          const image = document.createElement('img');\n\n          image.src = coverUrl;\n          image.alt = (book.title || '책') + ' 표지';\n          image.loading = 'lazy';\n          image.addEventListener('error', function() {\n            cover.textContent = '📖';\n          });\n          cover.appendChild(image);\n        } else {\n          cover.textContent = '📖';\n        }\n\n        title.textContent = book.title || '제목 없음';\n        meta.textContent = [\n          book.author,\n          book.publisher,\n          book.year ? book.year + '년' : book.pubDate,\n          formatPrice(book.price),\n          book.isbn\n        ].filter(Boolean).join(' · ');\n        button.textContent = isInCart\n          ? '장바구니에 담김 ✓'\n          : '신청바구니에 담기';\n        button.classList.toggle('selected', isInCart);\n        card.classList.toggle('selected', isInCart);\n        button.addEventListener('click', function() {\n          toggleCart(index, button);\n        });\n\n        info.appendChild(title);\n        info.appendChild(meta);\n        info.appendChild(button);\n        card.appendChild(cover);\n        card.appendChild(info);\n        bookList.appendChild(card);\n      });\n    }\n\n    function toggleCart(index, selectedButton) {\n      const book = books[index];\n      const bookKey = getBookKey(book);\n\n      if (\n        !book ||\n        !bookKey ||\n        selectedButton.dataset.done === 'true'\n      ) {\n        return;\n      }\n\n      if (cartBooks.has(bookKey)) {\n        cartBooks.delete(bookKey);\n        setVisibleBookSelection(selectedButton, false);\n      } else {\n        if (cartBooks.size >= 20) {\n          showMessage(\n            '한 번에 최대 20권까지 장바구니에 담을 수 있습니다.',\n            'error'\n          );\n          return;\n        }\n\n        cartBooks.set(bookKey, book);\n        setVisibleBookSelection(selectedButton, true);\n        showMessage(\n          '장바구니에 담았습니다. 다른 책을 계속 검색해도 그대로 유지됩니다.',\n          'success'\n        );\n      }\n\n      updateCartBar();\n    }\n\n    function updateCartBar() {\n      const count = cartBooks.size;\n\n      cartCount.textContent = String(count);\n      cartSubmit.disabled =\n        submissionInProgress || count === 0;\n      cartDrawerSubmit.disabled =\n        submissionInProgress || count === 0;\n      cartClear.disabled =\n        submissionInProgress || count === 0;\n      cartBar.classList.toggle('show', count > 0);\n      renderCartDrawer();\n    }\n\n    function openCartDrawer() {\n      if (!cartBooks.size) {\n        return;\n      }\n\n      renderCartDrawer();\n      cartDrawer.classList.add('show');\n      cartToggle.setAttribute('aria-expanded', 'true');\n      document.body.style.overflow = 'hidden';\n    }\n\n    function closeCartDrawer() {\n      cartDrawer.classList.remove('show');\n      cartToggle.setAttribute('aria-expanded', 'false');\n      document.body.style.overflow = '';\n    }\n\n    function clearCart() {\n      if (submissionInProgress) {\n        return;\n      }\n\n      cartBooks = new Map();\n\n      document\n        .querySelectorAll('.select-book-button')\n        .forEach(function(button) {\n          setVisibleBookSelection(button, false);\n        });\n\n      closeCartDrawer();\n      updateCartBar();\n      showMessage('신청바구니를 비웠습니다.', 'info');\n    }\n\n    function removeCartBook(bookKey) {\n      if (submissionInProgress) {\n        return;\n      }\n\n      cartBooks.delete(bookKey);\n      setVisibleBookSelection(\n        findBookButton(bookKey),\n        false\n      );\n      updateCartBar();\n\n      if (!cartBooks.size) {\n        closeCartDrawer();\n      }\n    }\n\n    function renderCartDrawer() {\n      cartList.innerHTML = '';\n\n      if (!cartBooks.size) {\n        const empty = document.createElement('div');\n        empty.className = 'cart-empty';\n        empty.textContent =\n          '아직 담은 책이 없습니다.\\n검색 결과에서 원하는 책을 담아 주세요.';\n        cartList.appendChild(empty);\n        return;\n      }\n\n      cartBooks.forEach(function(book, bookKey) {\n        const item = document.createElement('article');\n        const cover = document.createElement('div');\n        const info = document.createElement('div');\n        const title = document.createElement('h3');\n        const meta = document.createElement('p');\n        const remove = document.createElement('button');\n        const coverUrl = safeUrl(book.cover);\n\n        item.className = 'cart-item';\n        cover.className = 'cart-item-cover';\n        title.className = 'cart-item-title';\n        meta.className = 'cart-item-meta';\n        remove.className = 'cart-remove';\n        remove.type = 'button';\n        remove.textContent = '빼기';\n\n        if (coverUrl) {\n          const image = document.createElement('img');\n          image.src = coverUrl;\n          image.alt = (book.title || '책') + ' 표지';\n          image.loading = 'lazy';\n          image.addEventListener('error', function() {\n            cover.textContent = '📖';\n          });\n          cover.appendChild(image);\n        } else {\n          cover.textContent = '📖';\n        }\n\n        title.textContent = book.title || '제목 없음';\n        meta.textContent = [\n          book.author,\n          book.publisher,\n          book.year ? book.year + '년' : book.pubDate\n        ].filter(Boolean).join(' · ');\n\n        remove.addEventListener('click', function() {\n          removeCartBook(bookKey);\n        });\n\n        info.appendChild(title);\n        info.appendChild(meta);\n        item.appendChild(cover);\n        item.appendChild(info);\n        item.appendChild(remove);\n        cartList.appendChild(item);\n      });\n    }\n\n    function submitCart() {\n      const selectedEntries =\n        Array.from(cartBooks.entries());\n      const selectedKeys = selectedEntries.map(function(entry) {\n        return entry[0];\n      });\n      const selectedBooks = selectedEntries.map(function(entry) {\n        return entry[1];\n      });\n\n      if (!selectedBooks.length) {\n        showMessage(\n          '신청할 책을 한 권 이상 장바구니에 담아 주세요.',\n          'error'\n        );\n        return;\n      }\n\n      if (submissionInProgress) {\n        return;\n      }\n\n      submissionInProgress = true;\n      setCartLoading(true);\n      showMessage(\n        '장바구니의 ' +\n        selectedBooks.length +\n        '권을 안전하게 접수하고 있습니다…',\n        'info'\n      );\n\n      const requestPayload = {\n        applicant: applicant.value.trim(),\n        audience: window.HR_AUDIENCE || 'staff',\n        books: selectedBooks,\n        requestId: createRequestId()\n      };\n\n      sendCartRequest(\n        requestPayload,\n        selectedKeys,\n        selectedBooks,\n        0\n      );\n    }\n\n    function sendCartRequest(\n      requestPayload,\n      selectedKeys,\n      selectedBooks,\n      retryCount\n    ) {\n      google.script.run\n        .withSuccessHandler(function(response) {\n          submissionInProgress = false;\n          setCartLoading(false);\n          handleCartResponse(\n            response || {},\n            selectedKeys,\n            selectedBooks\n          );\n        })\n        .withFailureHandler(function(error) {\n          if (\n            retryCount < 2 &&\n            isTemporarySubmitError(error)\n          ) {\n            const waitTime =\n              900 +\n              retryCount * 1300 +\n              Math.floor(Math.random() * 700);\n\n            showMessage(\n              '신청이 몰려 순서대로 처리하고 있습니다. 잠시만 기다려 주세요…',\n              'info'\n            );\n\n            window.setTimeout(function() {\n              sendCartRequest(\n                requestPayload,\n                selectedKeys,\n                selectedBooks,\n                retryCount + 1\n              );\n            }, waitTime);\n            return;\n          }\n\n          submissionInProgress = false;\n          setCartLoading(false);\n          showMessage(getErrorMessage(error), 'error');\n        })\n        .HR_submitBookRequests(requestPayload);\n    }\n\n    function isTemporarySubmitError(error) {\n      const text = getErrorMessage(error).toLowerCase();\n\n      return (\n        text.includes('현재 신청자가 많') ||\n        text.includes('too many') ||\n        text.includes('timeout') ||\n        text.includes('timed out') ||\n        text.includes('lock') ||\n        text.includes('service invoked') ||\n        text.includes('일시적') ||\n        text.includes('시간 초과')\n      );\n    }\n\n    function createRequestId() {\n      if (\n        window.crypto &&\n        typeof window.crypto.randomUUID === 'function'\n      ) {\n        return window.crypto.randomUUID();\n      }\n\n      return (\n        String(Date.now()) + '-' +\n        Math.random().toString(36).slice(2)\n      );\n    }\n\n        function handleCartResponse(\n      response,\n      selectedKeys,\n      selectedBooks\n    ) {\n      const responseItems = Array.isArray(response.results)\n        ? response.results\n        : [];\n      const successTitles = [];\n      const duplicateTitles = [];\n      const ownedTitles = [];\n      const errorTitles = [];\n\n      responseItems.forEach(function(item, itemIndex) {\n        const book = selectedBooks[itemIndex] || {};\n        const bookKey =\n          selectedKeys[itemIndex] || getBookKey(book);\n        const button = findBookButton(bookKey);\n        const displayTitle =\n          item.title || book.title || '제목 없음';\n\n        if (item.success) {\n          successTitles.push(displayTitle);\n          cartBooks.delete(bookKey);\n          markBookCardDone(button, '신청 완료 ✓');\n        } else if (item.owned) {\n          ownedTitles.push(displayTitle);\n          cartBooks.delete(bookKey);\n          markBookCardDone(button, '소장 중');\n        } else if (item.duplicateRequest) {\n          duplicateTitles.push(displayTitle);\n          cartBooks.delete(bookKey);\n          markBookCardDone(button, '이미 신청됨');\n        } else {\n          errorTitles.push(displayTitle);\n          setVisibleBookSelection(button, true);\n        }\n      });\n\n      if (successTitles.length) {\n        submittedRequestsLoaded = false;\n      }\n\n      updateCartBar();\n\n      if (!cartBooks.size) {\n        closeCartDrawer();\n      }\n\n      const descriptionLines = [];\n\n      function addResultGroup(icon, label, titles) {\n        if (!titles.length) {\n          return;\n        }\n\n        descriptionLines.push(\n          icon + ' ' + label + ' (' + titles.length + '권)\\n' +\n          titles.map(function(title) {\n            return '• ' + title;\n          }).join('\\n')\n        );\n      }\n\n      addResultGroup('✅', '신청 완료', successTitles);\n      addResultGroup('🔁', '이미 신청된 도서', duplicateTitles);\n      addResultGroup('🏫', '학교에서 소장 중', ownedTitles);\n      addResultGroup('⚠️', '장바구니에 남겨 둔 도서', errorTitles);\n\n      showNoticeModal({\n        type: successTitles.length ? 'success' : 'duplicate',\n        icon: successTitles.length ? '✅' : '📚',\n        title: successTitles.length\n          ? '신청 완료 · ' + successTitles.length + '권'\n          : '신청 결과를 확인해 주세요.',\n        description: descriptionLines.join('\\n\\n')\n      });\n\n      showMessage(\n        successTitles.length\n          ? successTitles.length +\n            '권 신청을 완료했습니다.'\n          : '새로 신청된 책이 없습니다.',\n        successTitles.length ? 'success' : 'info'\n      );\n    }\n\n    function markBookCardDone(button, label) {\n      if (!button) {\n        return;\n      }\n\n      button.dataset.done = 'true';\n      button.textContent = label;\n      button.disabled = true;\n      button.classList.remove('selected');\n      button.classList.add('done');\n\n      if (button.closest('.book-card')) {\n        button.closest('.book-card')\n          .classList.remove('selected');\n        button.closest('.book-card')\n          .classList.add('done');\n      }\n    }\n\n    function setSearchLoading(isLoading) {\n      submitButton.disabled = isLoading;\n      submitButton.textContent =\n        isLoading\n          ? '검색하고 있습니다…'\n          : '책 검색하기';\n    }\n\n    function setCartLoading(isLoading) {\n      const buttons =\n        document.querySelectorAll('.select-book-button');\n\n      buttons.forEach(function(button) {\n        button.disabled =\n          isLoading ||\n          button.dataset.done === 'true';\n      });\n\n      query.disabled = isLoading;\n      applicant.disabled = isLoading;\n      submitButton.disabled = isLoading;\n      cartSubmit.disabled =\n        isLoading || cartBooks.size === 0;\n      cartDrawerSubmit.disabled =\n        isLoading || cartBooks.size === 0;\n      cartClear.disabled = isLoading;\n      cartSubmit.textContent = isLoading\n        ? '안전하게 접수하는 중…'\n        : '선택한 도서 신청하기';\n    }\n\n\n    function loadSubmittedRequests(forceRefresh) {\n      if (\n        submittedRequestsLoaded &&\n        !forceRefresh\n      ) {\n        return;\n      }\n\n      requestRefresh.disabled = true;\n      requestRefresh.textContent = '불러오는 중…';\n      requestList.innerHTML = '';\n      requestSummary.classList.remove('show');\n      requestListStatus.textContent =\n        '신청된 희망도서를 불러오고 있습니다…';\n      requestListStatus.classList.remove('hide');\n\n      google.script.run\n        .withSuccessHandler(function(response) {\n          requestRefresh.disabled = false;\n          requestRefresh.textContent = '새로고침';\n          submittedRequestsLoaded = true;\n          renderSubmittedRequests(response || {});\n        })\n        .withFailureHandler(function(error) {\n          requestRefresh.disabled = false;\n          requestRefresh.textContent = '다시 불러오기';\n          requestListStatus.textContent =\n            getErrorMessage(error);\n          requestListStatus.classList.remove('hide');\n        })\n        .HR_getSubmittedRequests(window.HR_AUDIENCE || 'staff');\n    }\n\n    function renderSubmittedRequests(response) {\n      const responseItems = Array.isArray(response.items)\n        ? response.items\n        : [];\n      const activeItems = responseItems.filter(function(item) {\n        return !/취소/.test(\n          String(item.status || '')\n        );\n      });\n      const cancelledItems = responseItems.filter(function(item) {\n        return /취소/.test(\n          String(item.status || '')\n        );\n      });\n      const items = requestListFilter === 'cancelled'\n        ? cancelledItems\n        : activeItems;\n      const totalCount = items.length;\n      const shownCount = items.length;\n\n      submittedRequestItems = responseItems;\n      requestActiveCount.textContent =\n        String(activeItems.length);\n      requestCancelledCount.textContent =\n        String(cancelledItems.length);\n\n      requestFilterTabs.forEach(function(tab) {\n        const selected =\n          tab.dataset.requestFilter ===\n          requestListFilter;\n\n        tab.classList.toggle('active', selected);\n        tab.setAttribute(\n          'aria-selected',\n          selected ? 'true' : 'false'\n        );\n      });\n\n      requestList.innerHTML = '';\n      requestSummary.textContent =\n        requestListFilter === 'cancelled'\n          ? '신청 취소 ' + totalCount + '권'\n          : '신청 도서 ' + totalCount + '권';\n      requestSummary.classList.add('show');\n\n      if (!items.length) {\n        requestListStatus.textContent =\n          requestListFilter === 'cancelled'\n            ? '현재 취소된 신청이 없습니다.'\n            : '아직 신청된 희망도서가 없습니다.';\n        requestListStatus.classList.remove('hide');\n        return;\n      }\n\n      requestListStatus.classList.add('hide');\n\n      items.forEach(function(item, index) {\n        const card = document.createElement('article');\n        const top = document.createElement('div');\n        const sequence = document.createElement('span');\n        const status = document.createElement('span');\n        const title = document.createElement('h3');\n        const meta = document.createElement('p');\n        const date = document.createElement('p');\n        const details = document.createElement('div');\n        const coverUrl = safeUrl(item.cover);\n        const detailHint = document.createElement('button');\n        const actions = document.createElement('div');\n        const cancelButton = document.createElement('button');\n        const statusText =\n          String(item.status || '신청 접수').trim();\n\n        card.className = 'request-item';\n        card.tabIndex = 0;\n        card.setAttribute('role', 'button');\n        card.setAttribute('aria-expanded', 'false');\n        top.className = 'request-item-top';\n        sequence.className = 'request-sequence';\n        status.className =\n          'request-status-badge ' +\n          getRequestStatusClass(statusText);\n        date.className = 'request-date';\n        details.className = 'request-biblio';\n        detailHint.className = 'request-detail-button';\n        detailHint.type = 'button';\n        detailHint.textContent = '서지정보 보기';\n        actions.className = 'request-card-actions';\n        cancelButton.className = 'request-cancel-button';\n        cancelButton.type = 'button';\n\n        sequence.textContent =\n          item.sequence\n            ? '#' + item.sequence\n            : '#' + (totalCount - index);\n        status.textContent = statusText;\n        title.textContent = item.title || '제목 없음';\n        meta.textContent = [\n          item.author,\n          item.publisher,\n          item.year ? item.year + '년' : ''\n        ].filter(Boolean).join(' · ') ||\n          '도서 정보를 확인하고 있습니다.';\n        date.textContent = item.requestedAt\n          ? '신청일 ' + item.requestedAt\n          : '';\n\n        if (coverUrl) {\n          const coverBox = document.createElement('div');\n          const coverImage = document.createElement('img');\n\n          coverBox.className = 'request-biblio-cover';\n          coverImage.src = coverUrl;\n          coverImage.alt =\n            (item.title || '책') + ' 표지';\n          coverImage.loading = 'lazy';\n          coverImage.addEventListener('error', function() {\n            coverBox.remove();\n          });\n          coverBox.appendChild(coverImage);\n          details.appendChild(coverBox);\n        }\n\n        [\n          ['저자', item.author],\n          ['출판사', item.publisher],\n          ['출판연도', item.year ? item.year + '년' : ''],\n          ['ISBN13', item.isbn],\n          ['정가', item.price\n            ? String(item.price).replace(/\\s*원$/, '') + '원'\n            : ''],\n          ['분야', item.category],\n          ['쪽수', item.pages]\n        ].forEach(function(entry) {\n          if (!entry[1]) {\n            return;\n          }\n\n          const row = document.createElement('div');\n          const label = document.createElement('strong');\n          const value = document.createElement('span');\n\n          row.className = 'request-biblio-row';\n          label.textContent = entry[0];\n          value.textContent = entry[1];\n          row.appendChild(label);\n          row.appendChild(value);\n          details.appendChild(row);\n        });\n\n        if (safeUrl(item.link)) {\n          const row = document.createElement('div');\n          const label = document.createElement('strong');\n          const link = document.createElement('a');\n\n          row.className = 'request-biblio-row';\n          label.textContent = '도서 정보';\n          link.href = safeUrl(item.link);\n          link.target = '_blank';\n          link.rel = 'noopener noreferrer';\n          link.textContent = 'YES24에서 보기 ↗';\n          link.addEventListener('click', function(event) {\n            event.stopPropagation();\n          });\n          row.appendChild(label);\n          row.appendChild(link);\n          details.appendChild(row);\n        }\n\n        if (!details.children.length) {\n          const empty = document.createElement('p');\n          empty.className = 'request-biblio-empty';\n          empty.textContent =\n            '추가 서지정보를 확인하고 있습니다.';\n          details.appendChild(empty);\n        }\n\n        function toggleDetails(event) {\n          if (\n            event &&\n            event.type === 'keydown' &&\n            event.key !== 'Enter' &&\n            event.key !== ' '\n          ) {\n            return;\n          }\n\n          if (event && event.type === 'keydown') {\n            event.preventDefault();\n          }\n\n          const expanded =\n            card.classList.toggle('expanded');\n          card.setAttribute(\n            'aria-expanded',\n            expanded ? 'true' : 'false'\n          );\n          detailHint.textContent = expanded\n            ? '서지정보 접기'\n            : '서지정보 보기';\n        }\n\n        card.addEventListener('click', toggleDetails);\n        card.addEventListener('keydown', toggleDetails);\n        detailHint.addEventListener('click', function(event) {\n          event.stopPropagation();\n          toggleDetails();\n        });\n\n        if (/취소/.test(statusText)) {\n          cancelButton.classList.add('restore');\n          cancelButton.textContent = '취소 되돌리기';\n          cancelButton.addEventListener('click', function(event) {\n            event.stopPropagation();\n            undoRequestFromList(\n              item,\n              cancelButton\n            );\n          });\n        } else if (/구입|입수|선정|반려/.test(statusText)) {\n          cancelButton.disabled = true;\n          cancelButton.textContent = '취소 불가';\n        } else {\n          cancelButton.textContent = '신청 취소';\n          cancelButton.addEventListener('click', function(event) {\n            event.stopPropagation();\n            openCancelRequest(\n              item,\n              cancelButton,\n              status\n            );\n          });\n        }\n\n        top.appendChild(sequence);\n        top.appendChild(status);\n        card.appendChild(top);\n        card.appendChild(title);\n        card.appendChild(meta);\n\n        if (date.textContent) {\n          card.appendChild(date);\n        }\n\n        card.appendChild(details);\n        actions.appendChild(detailHint);\n        actions.appendChild(cancelButton);\n        card.appendChild(actions);\n        requestList.appendChild(card);\n      });\n    }\n\n    function getRequestStatusClass(statusText) {\n      if (/오류|실패|취소|반려/.test(statusText)) {\n        return 'issue';\n      }\n\n      if (/완료|선정|구입|입수|처리 완료/.test(statusText)) {\n        return 'done';\n      }\n\n      return 'pending';\n    }\n\n    function openCancelRequest(\n      item,\n      button,\n      statusElement\n    ) {\n      cancelCandidate = {\n        item: item,\n        button: button,\n        statusElement: statusElement\n      };\n      cancelBook.textContent =\n        '『' + String(item.title || '제목 없음') + '』';\n      cancelConfirm.disabled = false;\n      cancelConfirm.textContent = '신청 취소하기';\n      cancelModal.classList.add('show');\n\n      window.setTimeout(function() {\n        cancelKeep.focus();\n      }, 30);\n    }\n\n    function closeCancelModal() {\n      cancelModal.classList.remove('show');\n      cancelCandidate = null;\n      cancelConfirm.disabled = false;\n      cancelConfirm.textContent = '신청 취소하기';\n    }\n\n    function confirmCancelRequest() {\n      if (!cancelCandidate) {\n        return;\n      }\n\n      const candidate = cancelCandidate;\n      cancelConfirm.disabled = true;\n      cancelConfirm.textContent = '취소 중…';\n\n      google.script.run\n        .withSuccessHandler(function(result) {\n          if (!result || result.success === false) {\n            const reason =\n              result && result.message\n                ? result.message\n                : '이 도서는 현재 취소할 수 없습니다.';\n            closeCancelModal();\n            showNoticeModal({\n              type: 'duplicate',\n              icon: 'ℹ️',\n              title: '취소할 수 없습니다',\n              book: candidate.item.title,\n              description: reason\n            });\n            return;\n          }\n\n          submittedRequestsLoaded = false;\n          closeCancelModal();\n          loadSubmittedRequests(true);\n\n          showNoticeModal({\n            type: 'success',\n            icon: '✅',\n            title: '신청이 취소되었습니다',\n            book: candidate.item.title,\n            description:\n              '신청 목록에서는 이 도서 한 권만 삭제했습니다. 취소한 내용은 안전하게 보관되어 바로 되돌릴 수 있습니다.',\n            undoId: result.cancelId\n          });\n        })\n        .withFailureHandler(function(error) {\n          const reason = getErrorMessage(error);\n          closeCancelModal();\n          showNoticeModal({\n            type: 'duplicate',\n            icon: '⚠️',\n            title: '취소하지 못했습니다',\n            book: candidate.item.title,\n            description: reason\n          });\n        })\n        .HR_cancelSubmittedRequest({\n          audience: window.HR_AUDIENCE || 'staff',\n          row: candidate.item.row,\n          title: candidate.item.title,\n          isbn: candidate.item.isbn\n        });\n    }\n\n    function undoRequestFromList(item, button) {\n      if (!item || !button) {\n        return;\n      }\n\n      button.disabled = true;\n      button.textContent = '복구 중…';\n\n      const onSuccess = function(result) {\n        if (!result || result.success === false) {\n          const reason =\n            result && result.message\n              ? result.message\n              : '신청 취소를 되돌리지 못했습니다.';\n\n          button.disabled = false;\n          button.textContent = '취소 되돌리기';\n          showNoticeModal({\n            type: 'duplicate',\n            icon: '⚠️',\n            title: '되돌리지 못했습니다',\n            book: item.title,\n            description: reason\n          });\n          return;\n        }\n\n        submittedRequestsLoaded = false;\n        loadSubmittedRequests(true);\n\n        showNoticeModal({\n          type: 'success',\n          icon: '↩️',\n          title: result.alreadyRestored\n            ? '이미 신청 목록에 있습니다'\n            : '신청 취소를 되돌렸습니다',\n          book: item.title,\n          description: result.alreadyRestored\n            ? '같은 도서가 이미 신청 목록에 있어 중복으로 넣지 않았습니다.'\n            : '취소 전 신청 내용이 희망도서 목록에 다시 들어갔습니다.'\n        });\n      };\n\n      const onFailure = function(error) {\n        button.disabled = false;\n        button.textContent = '취소 되돌리기';\n        showNoticeModal({\n          type: 'duplicate',\n          icon: '⚠️',\n          title: '되돌리지 못했습니다',\n          book: item.title,\n          description: getErrorMessage(error)\n        });\n      };\n\n      if (item.cancelId) {\n        google.script.run\n          .withSuccessHandler(onSuccess)\n          .withFailureHandler(onFailure)\n          .HR_undoCancelledRequest({\n            audience: window.HR_AUDIENCE || 'staff',\n            cancelId: item.cancelId\n          });\n        return;\n      }\n\n      google.script.run\n        .withSuccessHandler(onSuccess)\n        .withFailureHandler(onFailure)\n        .HR_restoreLegacyCancelledRequest({\n          row: item.row,\n          title: item.title,\n          isbn: item.isbn\n        });\n    }\n\n    function undoCancelledRequest() {\n      if (!undoCandidate) {\n        return;\n      }\n\n      const candidate = undoCandidate;\n      noticeUndo.disabled = true;\n      noticeUndo.textContent = '복구 중…';\n\n      google.script.run\n        .withSuccessHandler(function(result) {\n          if (!result || result.success === false) {\n            const reason =\n              result && result.message\n                ? result.message\n                : '신청을 되돌리지 못했습니다.';\n\n            showNoticeModal({\n              type: 'duplicate',\n              icon: '⚠️',\n              title: '되돌리지 못했습니다',\n              book: candidate.title,\n              description: reason,\n              undoId: candidate.cancelId\n            });\n            return;\n          }\n\n          submittedRequestsLoaded = false;\n          closeNoticeModal();\n          loadSubmittedRequests(true);\n\n          showNoticeModal({\n            type: 'success',\n            icon: '↩️',\n            title: result.alreadyRestored\n              ? '이미 신청 목록에 있습니다'\n              : '신청 취소를 되돌렸습니다',\n            book: candidate.title,\n            description: result.alreadyRestored\n              ? '같은 도서가 이미 신청 목록에 있어 중복으로 넣지 않았습니다.'\n              : '취소 전 신청 내용이 희망도서 목록에 다시 들어갔습니다.'\n          });\n        })\n        .withFailureHandler(function(error) {\n          showNoticeModal({\n            type: 'duplicate',\n            icon: '⚠️',\n            title: '되돌리지 못했습니다',\n            book: candidate.title,\n            description: getErrorMessage(error),\n            undoId: candidate.cancelId\n          });\n        })\n        .HR_undoCancelledRequest({\n          audience: window.HR_AUDIENCE || 'staff',\n          cancelId: candidate.cancelId\n        });\n    }\n\n    function closeGuidePopover() {\n      guideHover.classList.remove('open');\n      guideHoverTrigger.setAttribute(\n        'aria-expanded',\n        'false'\n      );\n    }\n\n    function safeUrl(value) {\n      const text = String(value || '').trim();\n\n      return /^https?:\\/\\//i.test(text)\n        ? text\n        : '';\n    }\n\n    function formatPrice(value) {\n      const number = Number(value || 0);\n\n      return number > 0\n        ? number.toLocaleString('ko-KR') + '원'\n        : '';\n    }\n\n    function showMessage(text, type) {\n      message.textContent = text;\n      message.className =\n        'message show ' + (type || 'info');\n    }\n\n    function showNoticeModal(options) {\n      const notice = options || {};\n\n      noticeModal.className =\n        'notice-modal show ' + (notice.type || 'duplicate');\n      noticeIcon.textContent = notice.icon || '📚';\n      noticeTitle.textContent =\n        notice.title || '안내';\n      noticeBook.textContent =\n        notice.book ? '『' + notice.book + '』' : '';\n      noticeDescription.textContent =\n        notice.description || '';\n\n      undoCandidate = notice.undoId\n        ? {\n          cancelId: notice.undoId,\n          title: notice.book || ''\n        }\n        : null;\n\n      noticeUndo.hidden = !undoCandidate;\n      noticeUndo.disabled = false;\n      noticeUndo.textContent = '신청 취소 되돌리기';\n      noticeActions.classList.toggle(\n        'single',\n        !undoCandidate\n      );\n\n      window.setTimeout(function() {\n        if (undoCandidate) {\n          noticeUndo.focus();\n        } else {\n          noticeClose.focus();\n        }\n      }, 30);\n    }\n\n    function closeNoticeModal() {\n      noticeModal.className = 'notice-modal';\n      undoCandidate = null;\n      noticeUndo.hidden = true;\n      noticeUndo.disabled = false;\n      noticeUndo.textContent = '신청 취소 되돌리기';\n      noticeActions.classList.add('single');\n    }\n\n    function getErrorMessage(error) {\n      return error && error.message\n        ? error.message\n        : String(\n          error ||\n          '처리 중 오류가 발생했습니다.'\n        );\n    }\n\n    let bsData = {};\n    let bsActiveCategory = 'all';\n\n    function loadBestsellers() {\n      const status = document.getElementById('bsStatus');\n\n      status.textContent = '베스트셀러를 불러오고 있습니다…';\n      status.classList.add('show');\n\n      google.script.run\n        .withSuccessHandler(function(response) {\n          bsData = response && response.categories\n            ? response.categories\n            : {};\n          renderBestsellers(bsActiveCategory);\n        })\n        .withFailureHandler(function() {\n          status.textContent =\n            '베스트셀러를 잠시 불러오지 못했습니다. 책 검색은 그대로 이용할 수 있습니다.';\n          status.classList.add('show');\n        })\n        .HR_getBestsellers();\n    }\n\n    function renderBestsellers(categoryKey) {\n      const list = document.getElementById('bsList');\n      const status = document.getElementById('bsStatus');\n      const category = bsData[categoryKey] || {};\n      const categoryBooks = Array.isArray(category.books)\n        ? category.books\n        : [];\n\n      stopBestsellerAuto();\n      list.innerHTML = '';\n\n      document.querySelectorAll('.bs-tab').forEach(function(tab) {\n        const active =\n          tab.dataset.bsCategory === categoryKey;\n\n        tab.classList.toggle('active', active);\n        tab.setAttribute(\n          'aria-selected',\n          active ? 'true' : 'false'\n        );\n      });\n\n      if (!categoryBooks.length) {\n        status.textContent =\n          '이 분류의 베스트셀러가 아직 없습니다.';\n        status.classList.add('show');\n        return;\n      }\n\n      status.classList.remove('show');\n\n      categoryBooks.forEach(function(book, index) {\n        const card = document.createElement('article');\n        const rank = document.createElement('span');\n        const cover = document.createElement('div');\n        const title = document.createElement('h3');\n        const meta = document.createElement('p');\n        const actions = document.createElement('div');\n        const requestButton = document.createElement('button');\n        const infoLink = document.createElement('a');\n        const coverUrl = safeUrl(book.cover);\n        const itemLink = safeUrl(book.link);\n\n        card.className = 'bs-card';\n        rank.className = 'bs-rank';\n        cover.className = 'bs-cover';\n        title.className = 'bs-book-title';\n        meta.className = 'bs-meta';\n        actions.className = 'bs-actions';\n        requestButton.className = 'bs-request';\n        infoLink.className = 'bs-link';\n\n        rank.textContent = String(book.rank || index + 1);\n        title.textContent = book.title || '제목 없음';\n        meta.textContent = [\n          book.author,\n          book.publisher,\n          book.year ? book.year + '년' : '',\n          formatPrice(book.price)\n        ].filter(Boolean).join(' · ');\n\n        if (coverUrl) {\n          const image = document.createElement('img');\n\n          image.src = coverUrl;\n          image.alt = (book.title || '책') + ' 표지';\n          image.loading = 'lazy';\n          image.addEventListener('error', function() {\n            cover.textContent = '📖';\n          });\n          cover.appendChild(image);\n        } else {\n          cover.textContent = '📖';\n        }\n\n        requestButton.type = 'button';\n        requestButton.textContent = '이 책 신청하기';\n        requestButton.addEventListener('click', function() {\n          searchBestsellerBook(book);\n        });\n\n        if (itemLink) {\n          infoLink.href = itemLink;\n          infoLink.target = '_blank';\n          infoLink.rel = 'noopener noreferrer';\n        } else {\n          infoLink.href = '#';\n          infoLink.addEventListener('click', function(event) {\n            event.preventDefault();\n          });\n        }\n        infoLink.textContent = '정보';\n\n        actions.appendChild(requestButton);\n        actions.appendChild(infoLink);\n        card.appendChild(rank);\n        card.appendChild(cover);\n        card.appendChild(title);\n        card.appendChild(meta);\n        card.appendChild(actions);\n        list.appendChild(card);\n      });\n\n      const originalCards =\n        Array.from(list.querySelectorAll('.bs-card'));\n\n      originalCards.forEach(function(originalCard, index) {\n        const cloneCard = originalCard.cloneNode(true);\n        const cloneRequest =\n          cloneCard.querySelector('.bs-request');\n\n        cloneCard.classList.add('bs-card-clone');\n\n        if (cloneRequest) {\n          cloneRequest.addEventListener('click', function() {\n            searchBestsellerBook(categoryBooks[index]);\n          });\n        }\n\n        list.appendChild(cloneCard);\n      });\n\n      list.dataset.bookCount =\n        String(categoryBooks.length);\n      bsListElement.scrollLeft = 0;\n      startBestsellerAuto();\n    }\n\n    function getBestsellerScrollStep() {\n      const firstCard =\n        bsListElement.querySelector('.bs-card');\n\n      if (!firstCard) {\n        return 220;\n      }\n\n      return firstCard.getBoundingClientRect().width + 12;\n    }\n\n    function scrollBestsellerCards(direction) {\n      const step = getBestsellerScrollStep();\n      const bookCount = Number(\n        bsListElement.dataset.bookCount || 0\n      );\n      const cycleWidth = step * bookCount;\n\n      if (!bookCount || cycleWidth <= 1) {\n        return;\n      }\n\n      if (direction > 0) {\n        if (\n          bsListElement.scrollLeft >=\n          cycleWidth - 8\n        ) {\n          bsListElement.scrollLeft =\n            Math.max(\n              0,\n              bsListElement.scrollLeft -\n              cycleWidth\n            );\n        }\n\n        bsListElement.scrollBy({\n          left: step,\n          behavior: 'smooth'\n        });\n        return;\n      }\n\n      if (bsListElement.scrollLeft <= 8) {\n        bsListElement.scrollLeft = cycleWidth;\n      }\n\n      bsListElement.scrollBy({\n        left: -step,\n        behavior: 'smooth'\n      });\n    }\n\n    function stopBestsellerAuto() {\n      if (bsAutoTimer) {\n        window.clearInterval(bsAutoTimer);\n        bsAutoTimer = null;\n      }\n    }\n\n    function startBestsellerAuto() {\n      stopBestsellerAuto();\n\n      if (\n        !applyPanel.classList.contains('active') ||\n        !bsListElement.querySelector('.bs-card')\n      ) {\n        return;\n      }\n\n      bsAutoTimer = window.setInterval(function() {\n        if (!document.hidden) {\n          scrollBestsellerCards(1);\n        }\n      }, 2000);\n    }\n\n    function searchBestsellerBook(book) {\n      const searchValue =\n        String(book.isbn || book.title || '').trim();\n\n      if (!searchValue) {\n        showMessage('이 책의 검색 정보가 없습니다.', 'error');\n        return;\n      }\n\n      query.value = searchValue;\n      window.scrollTo({ top: 0, behavior: 'smooth' });\n      showMessage('선택한 책을 정확히 확인하고 있습니다…', 'info');\n\n      if (typeof form.requestSubmit === 'function') {\n        form.requestSubmit();\n      } else {\n        submitButton.click();\n      }\n    }\n\n    document.querySelectorAll('.bs-tab').forEach(function(tab) {\n      tab.addEventListener('click', function() {\n        bsActiveCategory = tab.dataset.bsCategory || 'all';\n        renderBestsellers(bsActiveCategory);\n      });\n    });\n\n    loadBestsellers();\n\n  </script>\n\n  <script>\n    (function () {\n      const noticeWrap = document.getElementById('hrNoticeTicker');\n      const noticeTrack = document.getElementById('hrNoticeTrack');\n\n      function renderWebNotices(items) {\n        const notices = Array.isArray(items)\n          ? items.filter(function (item) {\n              return item && String(item.message || '').trim();\n            })\n          : [];\n\n        if (!noticeWrap || !noticeTrack || notices.length === 0) {\n          if (noticeWrap) noticeWrap.hidden = true;\n          return;\n        }\n\n        const urgent = notices.some(function (item) {\n          return String(item.type || '') === '긴급';\n        });\n\n        const message = notices.map(function (item) {\n          const prefix = String(item.type || '') === '긴급' ? '[긴급] ' : '';\n          return prefix + String(item.message || '').trim();\n        }).join('　◆　');\n\n        noticeTrack.textContent = message;\n        noticeTrack.style.setProperty(\n          '--ticker-duration',\n          Math.max(18, Math.min(70, message.length * 0.24)) + 's'\n        );\n        noticeWrap.classList.toggle('urgent', urgent);\n        noticeWrap.setAttribute('aria-label', '도서관 공지: ' + message);\n        noticeWrap.hidden = false;\n      }\n\n      function loadWebNotices() {\n        if (\n          typeof google === 'undefined' ||\n          !google.script ||\n          !google.script.run\n        ) {\n          if (noticeWrap) noticeWrap.hidden = true;\n          return;\n        }\n\n        google.script.run\n          .withSuccessHandler(renderWebNotices)\n          .withFailureHandler(function () {\n            if (noticeWrap) noticeWrap.hidden = true;\n          })\n          .HR_getActiveWebNotices(window.HR_AUDIENCE || 'staff');\n      }\n\n      if (document.readyState === 'loading') {\n        document.addEventListener('DOMContentLoaded', loadWebNotices);\n      } else {\n        loadWebNotices();\n      }\n\n      window.HR_reloadWebNotices = loadWebNotices;\n      window.setInterval(loadWebNotices, 300000);\n    })();\n  </script>\n\n  <script>\n    (function () {\n      var INITIAL = '<?= audience ?>';\n      var BASE_URL = '<?= baseUrl ?>';\n      var FROM_GATE = '<?= fromGate ?>' === 'true';\n      var gate = document.getElementById('audienceGate');\n      var bar = document.getElementById('audienceBar');\n      var badge = document.getElementById('audienceBadge');\n      var switchButton = document.getElementById('audienceSwitch');\n      var staffField = document.getElementById('staffApplicantField');\n      var parentField = document.getElementById('parentApplicantField');\n      var gradeInput = document.getElementById('pGrade');\n      var classInput = document.getElementById('pClass');\n      var numberInput = document.getElementById('pNumber');\n      var nameInput = document.getElementById('pName');\n      var uploadTab = document.querySelector('[data-panel=\"uploadPanel\"]');\n      var applyTab = document.querySelector('[data-panel=\"applyPanel\"]');\n      var applicantInput = document.getElementById('applicant');\n\n      window.HR_AUDIENCE = '';\n\n      function syncParentApplicant() {\n        var grade = gradeInput.value;\n        var classNo = String(classInput.value || '').trim();\n        var number = String(numberInput.value || '').trim();\n        var name = String(nameInput.value || '').trim();\n        applicantInput.value = (grade && classNo && number)\n          ? grade + '학년 ' + classNo + '반 ' + number + '번' + (name ? ' ' + name : '')\n          : '';\n      }\n\n      [gradeInput, classInput, numberInput, nameInput].forEach(function (el) {\n        el.addEventListener('input', syncParentApplicant);\n        el.addEventListener('change', syncParentApplicant);\n      });\n\n      function applyAudience(audience) {\n        var isParent = audience === 'parent';\n        window.HR_AUDIENCE = audience;\n        document.body.classList.toggle('aud-parent', isParent);\n        gate.hidden = true;\n        bar.hidden = false;\n        badge.textContent = isParent ? '👨‍👩‍👧 학부모 신청' : '🧑‍🏫 교직원 신청';\n        staffField.hidden = isParent;\n        parentField.hidden = !isParent;\n        if (uploadTab) uploadTab.hidden = isParent;      // 목록 업로드는 교직원 전용\n        if (isParent) {\n          applicantInput.value = '';\n          syncParentApplicant();\n          if (applyTab) applyTab.click();\n        }\n        if (typeof window.HR_reloadWebNotices === 'function') window.HR_reloadWebNotices();\n        if (typeof loadSubmittedRequests === 'function') {\n          try { loadSubmittedRequests(true); } catch (e) {}\n        }\n      }\n\n      function linkFor(audience) {\n        return BASE_URL + (BASE_URL.indexOf('?') < 0 ? '?' : '&') + 'type=' + audience + '&g=1';\n      }\n\n      // 주소를 알면 실제로 페이지를 이동하고(새 페이지로 꽉 차게), 모르면 같은 페이지에서 바꿉니다.\n      Array.prototype.forEach.call(gate.querySelectorAll('[data-audience]'), function (button) {\n        var audience = button.getAttribute('data-audience');\n        if (BASE_URL) {\n          button.setAttribute('href', linkFor(audience));\n          button.setAttribute('target', '_top');\n        } else {\n          button.addEventListener('click', function (event) {\n            event.preventDefault();\n            applyAudience(audience);\n          });\n        }\n      });\n\n      if (BASE_URL) {\n        switchButton.setAttribute('href', BASE_URL);\n      } else {\n        switchButton.addEventListener('click', function (event) {\n          event.preventDefault();\n          window.location.reload();\n        });\n      }\n\n      if (INITIAL === 'parent' || INITIAL === 'staff') {\n        switchButton.hidden = !FROM_GATE;   // 전용 링크로 바로 들어온 경우엔 처음 화면으로 못 가게 숨깁니다.\n        applyAudience(INITIAL);\n      } else {\n        gate.hidden = false;\n        if (typeof window.HR_reloadWebNotices === 'function') window.HR_reloadWebNotices();\n      }\n    })();\n  </script>\n</body>\n</html>\n"},
    menus: (typeof APP_MENUS !== 'undefined') ? APP_MENUS : [],
    rpc: ["HR_showApplyLinksDialog","HR_showRequestLimitDialog","getConfiguredNlkKey","startNewRequestRound","styleRequestListSheet","buildBookstoreExportSheet","installEditTrigger","onEditTriggerAction","processRow","shouldSkipStatus","smartSearch","fetchYes24Api","selectBestYes24Item","fetchNlkCallNumber","parseNlkClassNumber","findFirstXmlValue","normalizeClassNumber","findDuplicateRequest","findLibraryDuplicate","fillFinalData","checkExistingBooks","reprocessSelectedRows","getHeaders","getColumnMap","getLibraryColumnMap","setStatus","parseTitleAndPublisher","extractBestIsbn","normalizeIsbn","normalizeText","normalizeHeader","stripHtml","extractYear","toNumber","isTitleHeader","isIsbnHeader","isStatusHeader","isCoverHeader","isPublisherHeader","isAuthorHeader","isPriceHeader","isYearHeader","isCategoryHeader","isCallNumberHeader","createCacheKey","getCachedJson","putCachedJson","FIX_ACTION_LAYOUT_20260820","resetHopeBooksForNewSemester","HR_importFromOldCopy","HR_renameLegacyTargetSheet","HR_showRequestSidebar","HR_openSidebarOnSpreadsheetOpen","HR_installSidebarOpenTrigger","HR_submitBookRequest","HR_submitBookRequests","HR_searchRequestBooks","HR_formatGuideSheet","HR_addCopySetupMenu","HR_showYes24KeyDialog","HR_showSchoolSetupDialog","HR_showPublicWebAppUrlDialog","HR_savePublicWebAppUrl","HR_getCopySetupState","HR_saveYes24ApiKey","HR_runSchoolInitialSetup","HR_deleteCopySetupSheet","HR_createCopySetupSheet","HR_refreshCopySetupBeginnerGuide","HR_showSystemCheck","HR_getSubmittedRequests","HR_restoreLegacyCancelledRequest","HR_cancelSubmittedRequest","HR_undoCancelledRequest","HR_getBestsellers","HR_fillMissingPriceAndYear","HR_importBookListFile","HR_showHopeRequestDialog","setupSeriesFinder","installSeriesWorkButton","setSeriesYes24Key","refreshMissingSeriesBooks","resetSeriesFinderProgress","refreshSeriesSaleStatuses","refreshSeriesApiGauge","setupSeriesWorkSheet","exportSelectedSeriesBooks","setupUnifiedLibraryWorkflow","openUnifiedUploadSheet","refreshUnifiedPurchaseReview","buildUnifiedBookstoreExportSheet","unifiedAuditDataSources","importLegacySeriesSheetsOnce","sendHopeRequestsToReview","sendSelectedSeriesToReview","sendSelectedUploadItemsToReview","applyUnifiedLibraryDesign","removeLegacyUnifiedSheetsAfterConfirmation","HR_showUpdateNotice","HR_upgradeNoticeDatePickersAndGuide","HR_getActiveWebNotices","HR_setupWebNoticeSheet","verifyYes24Connection"],
    has: function (n) { return typeof __fns[n] === 'function'; },
    call: function (n, a) {
      if (typeof __fns[n] !== 'function') throw new Error('함수를 찾을 수 없습니다: ' + n);
      return __fns[n].apply(null, a || []);
    }
  };
};
