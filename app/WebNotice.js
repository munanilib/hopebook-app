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


function HR_getActiveWebNotices() {
  var spreadsheet = HR_getSpreadsheet_();
  var sheet = spreadsheet.getSheetByName(HR_WEB_NOTICE_SHEET_NAME);

  if (!sheet || sheet.getLastRow() < 4) {
    return [];
  }

  var values = sheet.getRange(4, 1, sheet.getLastRow() - 3, 5).getValues();
  var now = new Date();
  now.setHours(0, 0, 0, 0);

  return values.reduce(function(result, row) {
    var enabled = row[0] === true;
    var type = String(row[1] || '일반').trim();
    var message = String(row[2] || '').trim();
    var start = HR_noticeDate_(row[3], false);
    var end = HR_noticeDate_(row[4], true);

    if (!enabled || !message) {
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
