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
    ensureSelectionControls_();
  } catch (error) {
    console.warn('선택 버튼 표시 실패: ' + error);
  }
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

