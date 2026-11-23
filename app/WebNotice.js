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
    HR_styleWebNoticeSheet_(spreadsheet.getSheetByName(HR_WEB_NOTICE_SHEET_NAME));
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


// 「📢 웹 공지」 시트를 6칸(표시·종류·공지 내용·시작일·종료일·대상)으로 한 덩어리처럼 예쁘게 다시 꾸밉니다.
// 이미 적어 둔 공지 내용은 그대로 두고 모양만 바꿉니다. 여러 번 불러도 같은 결과입니다.
function HR_styleWebNoticeSheet_(sheet) {
  if (!sheet) return;
  var LAST = 6;
  if (sheet.getMaxColumns() < LAST) sheet.insertColumnsAfter(sheet.getMaxColumns(), LAST - sheet.getMaxColumns());
  if (sheet.getMaxRows() < 1000) sheet.insertRowsAfter(sheet.getMaxRows(), 1000 - sheet.getMaxRows());

  // 머리말: 제목·안내는 여섯 칸 전체에 걸쳐 놓고, 칸 이름도 여섯 개를 한 줄로 맞춥니다.
  sheet.getRange('A1:F1').breakApart();
  sheet.getRange('A2:F2').breakApart();
  sheet.getRange('A1:F1').merge();
  sheet.getRange('A2:F2').merge();
  sheet.getRange('A1').setValue('📢 웹페이지 공지 전광판');
  sheet.getRange('A2').setValue(
    '「표시」에 체크하면 웹페이지 맨 위에 공지가 올라갑니다.  ' +
    '「대상」에서 전체 · 교직원 · 학부모 중 누구에게 보일지 고를 수 있습니다. (비워 두면 전체)  ' +
    '시작일·종료일은 비워도 됩니다.'
  );
  sheet.getRange('A3:F3').setValues([['표시', '종류', '공지 내용', '시작일', '종료일', '대상']]);

  sheet.setHiddenGridlines(true);
  sheet.setFrozenRows(3);
  sheet.setTabColor('#9AA6A0');
  [80, 90, 560, 120, 120, 110].forEach(function(w, i) { sheet.setColumnWidth(i + 1, w); });
  sheet.setRowHeight(1, 60);
  sheet.setRowHeight(2, 48);
  sheet.setRowHeight(3, 38);

  sheet.getRange('A1:F1')
    .setBackground('#244f42').setFontColor('#ffffff').setFontSize(20).setFontWeight('bold')
    .setHorizontalAlignment('left').setVerticalAlignment('middle');
  sheet.getRange('A2:F2')
    .setBackground('#eef5f1').setFontColor('#52685f').setFontSize(11).setFontWeight('normal')
    .setWrap(true).setHorizontalAlignment('left').setVerticalAlignment('middle');
  sheet.getRange('A3:F3')
    .setBackground('#d7a843').setFontColor('#ffffff').setFontSize(12).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');

  // 본문: 여섯 칸을 같은 높이·같은 줄무늬·같은 선으로 맞춥니다.
  var body = sheet.getRange('A4:F1000');
  body.setFontSize(11).setFontColor('#2e3a33').setBackground('#ffffff').setVerticalAlignment('middle')
    .setBorder(false, false, true, false, false, false, '#e3e9e5', SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange('A4:B1000').setHorizontalAlignment('center');
  sheet.getRange('C4:C1000').setHorizontalAlignment('left').setWrap(true);
  sheet.getRange('D4:F1000').setHorizontalAlignment('center');

  // 칸 모양: 체크박스 · 종류 · 날짜 · 대상(드롭다운)
  sheet.getRange('A4:A1000').insertCheckboxes();
  sheet.getRange('B4:B1000').setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(['일반', '긴급'], true).setAllowInvalid(false).build()
  );
  sheet.getRange('D4:E1000')
    .setDataValidation(SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(false)
      .setHelpText('셀을 눌러 달력에서 날짜를 선택하세요.').build())
    .setNumberFormat('yyyy. m. d.');
  sheet.getRange('F4:F1000').setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(HR_NOTICE_AUDIENCE_VALUES, true).setAllowInvalid(true)
      .setHelpText('전체 · 교직원 · 학부모 중에서 고르세요. 비워 두면 전체입니다.').build()
  );

  // 색으로 한눈에: 긴급은 빨간 글자, 대상은 칩처럼 색을 칠합니다.
  function rule(rangeA1, text, bg, fg) {
    return SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(text).setBackground(bg).setFontColor(fg).setBold(true)
      .setRanges([sheet.getRange(rangeA1)]).build();
  }
  sheet.setConditionalFormatRules([
    rule('B4:B1000', '긴급', '#fdecea', '#b3261e'),
    rule('F4:F1000', '학부모', '#fdf0d0', '#8a5a00'),
    rule('F4:F1000', '교직원', '#dcefe6', '#1f5c4a'),
    rule('F4:F1000', '전체', '#eceff0', '#55616a')
  ]);
}


function HR_getActiveWebNotices(audience) {
  // 공지는 없어도 신청에는 지장이 없으므로, 어떤 문제가 생겨도 오류를 내지 않고 빈 목록을 돌려줍니다.
  try {
    return HR_readActiveWebNotices_(audience);
  } catch (error) {
    console.warn('웹 공지를 읽지 못했습니다: ' + error);
    return [];
  }
}

function HR_readActiveWebNotices_(audience) {
  var audienceKey = HR_audienceKey_(audience);
  var spreadsheet = HR_getSpreadsheet_();
  var sheet = spreadsheet.getSheetByName(HR_WEB_NOTICE_SHEET_NAME);

  if (!sheet || sheet.getLastRow() < 4) {
    return [];
  }

  // 옛 시트는 6번째 칸(대상)이 아직 없을 수 있어서, 있는 칸까지만 읽습니다.
  var columnCount = Math.min(6, sheet.getMaxColumns());
  var values = sheet.getRange(4, 1, sheet.getLastRow() - 3, columnCount).getValues();
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
  HR_styleWebNoticeSheet_(sheet);

  return {
    success: true,
    created: true,
    sheetName: HR_WEB_NOTICE_SHEET_NAME
  };
}
