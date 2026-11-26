/**
 * 🎛️ 클릭 막대 (체크박스 대신)
 *
 * 각 시트 맨 위(4행)의 「전체 선택」「구입 검토함으로 보내기」 같은 버튼을 체크박스가 아니라
 * 막대 전체를 누르면 동작하는 큰 버튼으로 만듭니다.
 * - 칸을 선택하는 순간 동작합니다. (onSelectionChange)
 * - 보내기·만들기처럼 되돌리기 번거로운 동작은 「정말 할까요?」를 한 번 물어봅니다.
 * - 「전체 선택/해제」는 막대 앞의 ☐/☑ 로 지금 상태를 보여 줍니다.
 */
var UNIFIED_BARS_VERSION = '2';
// 전체 선택 막대의 두 상태 글자 (앞 그림이 상태를 알려 줍니다)
var BAR_SELECT_ON = '✅  전체 선택';
var BAR_SELECT_OFF = '⬜  전체 해제';
var UNIFIED_BARS_VERSION_PROPERTY = 'UNIFIED_BARS_VERSION';

// 시트 이름 → 막대 모양. (kind: select=전체 선택 / send=보내기 / check=점검 시작 / rebuild=다시 만들기)
function unifiedBarSpec_(sheet) {
  var name = sheet.getName();
  var SEND = '📤  선택한 책 구입 검토함으로 보내기  ›';
  var SELECT = BAR_SELECT_ON;

  if (name === TARGET_SHEET_NAME) {
    var last = Math.max(sheet.getLastColumn(), 12);
    return { last: last, bars: [
      { id: 'hope-send', kind: 'send', c1: 1, c2: last, label: '📥  신청 목록을 구입 검토함으로 보내기  ›', confirm: '접수된 희망도서를 구입 검토함으로 보낼까요?' }
    ] };
  }
  if (name === SERIES_FINDER_SHEET) {
    return { last: 14, bars: [
      { id: 'series-select', kind: 'select', c1: 1, c2: 3, label: SELECT },
      { id: 'series-check', kind: 'check', c1: 4, c2: 8, label: '🔎  시리즈 누락 점검 시작  ›' },
      { id: 'series-send', kind: 'send', c1: 9, c2: 14, label: SEND, confirm: '체크한 시리즈 도서를 구입 검토함으로 보낼까요?' }
    ] };
  }
  if (name === UNIFIED_UPLOAD_SHEET) {
    return { last: 11, bars: [
      { id: 'upload-select', kind: 'select', c1: 1, c2: 3, label: SELECT },
      { id: 'upload-send', kind: 'send', c1: 4, c2: 11, label: SEND, confirm: '체크한 책을 구입 검토함으로 보낼까요?' }
    ] };
  }
  if (name === UNIFIED_REVIEW_SHEET) {
    return { last: 12, bars: [
      { id: 'review-select', kind: 'select', c1: 1, c2: 3, label: SELECT },
      { id: 'review-export', kind: 'send', c1: 4, c2: 12, label: '📦  서점 전달용으로 보내기 (중복 확인)  ›', confirm: '체크한 책의 소장 중복을 확인하고 서점 전달용 목록을 만들까요?' }
    ] };
  }
  if (name === UNIFIED_EXPORT_SHEET) {
    return { last: 10, bars: [
      { id: 'export-rebuild', kind: 'send', c1: 1, c2: 10, label: '🔄  서점 전달용 다시 만들기  ›', confirm: '구입 검토함의 최신 선택으로 서점 전달용 목록을 다시 만들까요?' }
    ] };
  }
  return null;
}

// 모든 막대는 같은 높이·같은 글자 크기·같은 모양(그림 + 글 + ›)이고, 역할에 따라 색만 다릅니다.
var UNIFIED_BAR_STYLE = {
  select:  { bg: '#CFE6DA', fg: '#14493A', border: '#2E7D5B', size: 16 },
  send:    { bg: '#B9770E', fg: '#FFFFFF', border: '#7A4A06', size: 16 },
  check:   { bg: '#1F5C4A', fg: '#FFFFFF', border: '#0F3A2E', size: 16 }
};

// 4행에 막대를 그립니다. (이미 있던 체크박스·병합은 지우고 새로 만듭니다. 여러 번 불러도 같은 결과입니다.)
function unifiedDrawActionBars_(sheet) {
  try {
    var spec = sheet ? unifiedBarSpec_(sheet) : null;
    if (!spec) return;
    if (sheet.getMaxColumns() < spec.last) sheet.insertColumnsAfter(sheet.getMaxColumns(), spec.last - sheet.getMaxColumns());

    var row = sheet.getRange(4, 1, 1, spec.last);
    row.breakApart();
    row.clearContent();
    row.clearDataValidations();

    spec.bars.forEach(function(bar) {
      var style = UNIFIED_BAR_STYLE[bar.kind] || UNIFIED_BAR_STYLE.send;
      var range = sheet.getRange(4, bar.c1, 1, bar.c2 - bar.c1 + 1);
      if (bar.c2 > bar.c1) range.merge();
      range
        .setValue(bar.label)
        .setBackground(style.bg)
        .setFontColor(style.fg)
        .setFontFamily('Arial')
        .setFontSize(style.size)
        .setFontWeight('bold')
        .setHorizontalAlignment('center')
        .setVerticalAlignment('middle')
        .setBorder(true, true, true, true, false, false, style.border, SpreadsheetApp.BorderStyle.SOLID_THICK);
    });
    sheet.setRowHeight(4, 56);
  } catch (error) {
    console.warn('클릭 막대 그리기 실패(' + (sheet && sheet.getName()) + '): ' + error);
  }
}

// 전체 선택 막대의 상태 표시를 ☐ 로 되돌립니다. (보낸 뒤 체크를 모두 해제했을 때)
function unifiedResetBarState_(sheet) {
  try {
    var spec = unifiedBarSpec_(sheet);
    if (!spec) return;
    spec.bars.forEach(function(bar) {
      if (bar.kind === 'select') sheet.getRange(4, bar.c1).setValue(BAR_SELECT_ON);
    });
  } catch (error) {
    console.warn('막대 상태 되돌리기 실패: ' + error);
  }
}

// 모든 대상 시트에 막대를 한 번에 그립니다. (이미 쓰던 사본을 새 모양으로 바꿀 때)
function unifiedApplyActionBarsAll_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  [TARGET_SHEET_NAME, SERIES_FINDER_SHEET, UNIFIED_UPLOAD_SHEET, UNIFIED_REVIEW_SHEET, UNIFIED_EXPORT_SHEET].forEach(function(name) {
    var sheet = ss.getSheetByName(name);
    if (sheet) unifiedDrawActionBars_(sheet);
  });
}

// 선택한 칸이 막대 안인지 찾습니다. (막대 하나 안에 들어올 때만. 4행 전체를 고르는 등은 무시)
function unifiedBarAt_(sheet, range) {
  if (!range || range.getRow() !== 4 || range.getNumRows() !== 1) return null;
  var spec = unifiedBarSpec_(sheet);
  if (!spec) return null;
  var first = range.getColumn(), last = range.getLastColumn();
  for (var i = 0; i < spec.bars.length; i++) {
    if (first >= spec.bars[i].c1 && last <= spec.bars[i].c2) return spec.bars[i];
  }
  return null;
}

// 막대를 눌렀을 때 동작. 막대였으면 true 를 돌려줍니다. (onSelectionChange 에서 부릅니다)
function unifiedHandleBarClick_(sheet, range) {
  var bar = unifiedBarAt_(sheet, range);
  if (!bar) return false;

  // 같은 클릭이 두 번 처리되지 않게 잠깐 표시를 남깁니다.
  var cache = CacheService.getDocumentCache();
  var busyKey = 'BAR_BUSY_' + bar.id;
  if (cache && cache.get(busyKey)) return true;
  if (cache) cache.put(busyKey, '1', 6);

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  try {
    if (bar.kind === 'select') {
      var cell = sheet.getRange(4, bar.c1);
      var checked = String(cell.getValue() || '').indexOf('⬜') !== 0;   // 지금 「전체 선택」 상태가 아니면 모두 선택
      toggleSelectionColumn_(sheet, 6, 1, checked);
      cell.setValue(checked ? BAR_SELECT_OFF : BAR_SELECT_ON);
      if (sheet.getName() === UNIFIED_REVIEW_SHEET) unifiedUpdateDashboardNumbers_();
    } else if (bar.id === 'series-check') {
      HR_openSeriesCheckDialog_(true);   // 누르면 창이 열리면서 바로 점검을 시작합니다.
    } else {
      var ui = SpreadsheetApp.getUi();
      var answer = ui.alert('확인', bar.confirm, ui.ButtonSet.YES_NO);
      if (answer === ui.Button.YES) {
        if (bar.id === 'hope-send') sendHopeRequestsToReview();
        else if (bar.id === 'series-send') sendSelectedSeriesToReview();
        else if (bar.id === 'upload-send') sendSelectedUploadItemsToReview();
        else buildUnifiedBookstoreExportSheet();   // review-export · export-rebuild
      }
    }
  } catch (error) {
    ss.toast(String(error && error.message ? error.message : error), '처리하지 못했습니다', 8);
  } finally {
    // 같은 막대를 다시 눌러도 동작하도록, 선택을 막대 밖(목록 첫 칸)으로 옮깁니다.
    try { sheet.setActiveRange(sheet.getRange(6, 2)); } catch (error) {}
  }
  return true;
}

/* ---------- 🔎 시리즈 누락 점검 창 ---------- */

function HR_showSeriesCheckDialog() { HR_openSeriesCheckDialog_(false); }

// autostart 이면 창이 열리자마자 점검(또는 이어서 점검)을 시작합니다.
function HR_openSeriesCheckDialog_(autostart) {
  var template = HtmlService.createTemplateFromFile('SeriesCheck');
  template.autostart = autostart ? 'true' : 'false';
  SpreadsheetApp.getUi().showModalDialog(
    template.evaluate().setWidth(520).setHeight(620),
    '🔎 시리즈 누락 점검'
  );
}

// 창이 열릴 때와 점검이 끝날 때마다 현재 상태를 알려 줍니다.
function HR_seriesCheckState() {
  var props = PropertiesService.getScriptProperties();
  var resume = null;
  try {
    var parsed = JSON.parse(props.getProperty(SERIES_RESUME_STATE_KEY) || 'null');
    if (parsed && Number(parsed.total) > 0) resume = { offset: Number(parsed.offset || 0), total: Number(parsed.total || 0) };
  } catch (error) {}
  var ss = seriesSpreadsheet_();
  var result = ss.getSheetByName(SERIES_FINDER_SHEET);
  var holdings = ss.getSheetByName(SERIES_HOLDINGS_SHEET);
  return {
    hasKey: Boolean(seriesGetYes24Key_()),
    resume: resume,
    candidates: result ? unifiedCountData_(result, 6) : 0,
    holdings: holdings ? Math.max(0, holdings.getLastRow() - 1) : 0
  };
}

// 「점검 시작」: fresh 이면 저장된 위치를 지우고 처음부터, 아니면 이어서 점검합니다.
function HR_seriesCheckStart(fresh) {
  if (fresh) {
    seriesCancelAutoContinue_();
    seriesClearResumeState_();
  }
  refreshMissingSeriesBooks();
  return HR_seriesCheckState();
}
