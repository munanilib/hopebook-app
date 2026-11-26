/**
 * 👥 같은 책을 원하는 사람 수 (신청 인원)
 *
 * - 같은 책이 이미 신청돼 있을 때 다른 사람이 또 신청하면, 책은 새로 늘리지 않고 「신청 인원」만 올립니다.
 * - 같은 사람이 같은 책을 여러 번 눌러도 한 명으로만 셉니다.
 * - 구입 검토함에는 비고에 「👥 N명 신청」으로 보이고, 많이 신청한 책이 위로 올라옵니다.
 * - 검색 결과에는 「우리 학교에 이미 있어요」「이미 N명이 신청했어요」를 미리 보여 줍니다.
 */
var HR_REQUESTERS_HEADER = '신청 인원';

function HR_findRequestersColumn_(headers) {
  return HR_findAnyHeaderColumn_(headers || [], [HR_REQUESTERS_HEADER]);
}

// 「신청 인원」 칸이 없으면 맨 오른쪽에 만들고 열 번호를 돌려줍니다.
function HR_ensureRequestersColumn_(sheet) {
  var headers = getHeaders(sheet);
  var column = HR_findRequestersColumn_(headers);
  if (column) return column;

  column = Math.max(sheet.getLastColumn(), headers.length) + 1;
  if (sheet.getMaxColumns() < column) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), column - sheet.getMaxColumns());
  }
  var header = sheet.getRange(HEADER_ROW, column);
  if (column > 1) {
    sheet.getRange(HEADER_ROW, column - 1).copyTo(header, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  }
  header.setValue(HR_REQUESTERS_HEADER);
  sheet.setColumnWidth(column, 80);
  return column;
}

function HR_normalizeApplicantText_(text) {
  return String(text || '').replace(/\s+/g, '').toLowerCase();
}

// 숨김 접수 기록에서, 이 책을 「이미 신청됨」으로 다시 신청한 서로 다른 사람 수(원래 신청자 제외)
function HR_countExtraRequesters_(spreadsheet, book, originalApplicant) {
  var backup = spreadsheet.getSheetByName(HR_REQUEST_BACKUP_SHEET_NAME);
  if (!backup || backup.getLastRow() < 2) return 0;
  var targetIsbn = HR_normalizeIsbn_(book && book.isbn);
  var targetTitle = HR_normalizeBookTitle_(book && book.title);
  var rows = backup.getRange(2, 3, backup.getLastRow() - 1, 5).getDisplayValues(); // C신청자 D검색어 E도서명 F ISBN G처리상태
  var seen = {};
  var original = HR_normalizeApplicantText_(originalApplicant);
  var count = 0;
  rows.forEach(function(row) {
    if (String(row[4]) !== '이미 신청됨') return;
    var isbn = HR_normalizeIsbn_(row[3]);
    var same = targetIsbn ? (isbn && isbn === targetIsbn) : (targetTitle && HR_normalizeBookTitle_(row[2]) === targetTitle);
    if (!same) return;
    var who = HR_normalizeApplicantText_(row[0]);
    if (!who || who === original || seen[who]) return;
    seen[who] = true;
    count++;
  });
  return count;
}

// 이미 신청된 책을 또 신청했을 때: 신청 인원을 다시 세어 희망도서 목록에 적습니다. 총 인원을 돌려줍니다.
function HR_recordExtraRequester_(spreadsheet, sheet, existingRow, book, applicant) {
  if (!existingRow) return 1;
  var headers = getHeaders(sheet);
  var applicantColumn = HR_findHeaderColumn_(headers, '신청자', 11);
  var original = applicantColumn ? String(sheet.getRange(existingRow, applicantColumn).getDisplayValue() || '') : '';
  var total = 1 + HR_countExtraRequesters_(spreadsheet, book, original);
  var column = HR_ensureRequestersColumn_(sheet);
  sheet.getRange(existingRow, column).setValue(total);
  return {
    total: total,
    sameApplicant: HR_normalizeApplicantText_(original) === HR_normalizeApplicantText_(applicant)
  };
}

// 검색 결과 각 책에 「이미 소장」「이미 신청(몇 명)」 표시를 붙입니다. 시트는 한 번씩만 읽습니다.
function HR_annotateBooksStatus_(books) {
  var list = Array.isArray(books) ? books : [];
  if (!list.length) return list;
  try {
    var spreadsheet = HR_getSpreadsheet_();
    var ownedIsbn = {}, ownedTitle = {};
    var library = spreadsheet.getSheetByName(HR_LIBRARY_SHEET_NAME);
    if (library && library.getLastRow() >= 2 && library.getLastColumn() >= 1) {
      var lh = library.getRange(1, 1, 1, library.getLastColumn()).getDisplayValues()[0];
      var lt = HR_findAnyHeaderColumn_(lh, ['책제목', '도서명', '서명', '제목']);
      var li = HR_findAnyHeaderColumn_(lh, ['ISBN13', 'ISBN']);
      library.getRange(2, 1, library.getLastRow() - 1, library.getLastColumn()).getDisplayValues().forEach(function(r) {
        var isbn = li ? HR_normalizeIsbn_(r[li - 1]) : '';
        if (isbn) ownedIsbn[isbn] = true;
        else if (lt) { var k = HR_normalizeBookTitle_(r[lt - 1]); if (k) ownedTitle[k] = true; }
      });
    }
    var requestedIsbn = {}, requestedTitle = {};
    var sheet = spreadsheet.getSheetByName(HR_TARGET_SHEET_NAME);
    if (sheet && sheet.getLastRow() >= FIRST_DATA_ROW) {
      var headers = getHeaders(sheet);
      var columns = getColumnMap(headers);
      var rc = HR_findRequestersColumn_(headers);
      sheet.getRange(FIRST_DATA_ROW, 1, sheet.getLastRow() - FIRST_DATA_ROW + 1, sheet.getLastColumn()).getDisplayValues().forEach(function(r) {
        var isbn = columns.isbn ? HR_normalizeIsbn_(r[columns.isbn - 1]) : '';
        var title = columns.title ? HR_normalizeBookTitle_(r[columns.title - 1]) : '';
        if (!isbn && !title) return;
        var n = rc ? Math.max(1, Number(r[rc - 1]) || 1) : 1;
        if (isbn) requestedIsbn[isbn] = n; else requestedTitle[title] = n;
      });
    }
    list.forEach(function(book) {
      var isbn = HR_normalizeIsbn_(book && book.isbn);
      var title = HR_normalizeBookTitle_(book && book.title);
      if (isbn ? ownedIsbn[isbn] : (title && ownedTitle[title])) book.alreadyOwned = true;
      var n = isbn ? requestedIsbn[isbn] : (title ? requestedTitle[title] : 0);
      if (n) book.alreadyRequested = n;
    });
  } catch (error) {
    console.warn('검색 결과 표시 붙이기 실패: ' + error);
  }
  return list;
}

// 구입 검토함 비고에 「👥 N명 신청」을 붙이거나 갱신합니다(이전 표시는 지움).
function HR_composeRequesterNote_(note, count) {
  var text = String(note || '').replace(/👥\s*\d+명 신청\s*(?:·\s*)?/g, '').trim();
  var n = Number(count) || 0;
  if (n < 2) return text;
  var tag = '👥 ' + n + '명 신청';
  return text ? tag + ' · ' + text : tag;
}
