/**
 * 🎛️ 클릭 막대 (체크박스 대신)
 *
 * 각 시트 맨 위(4행)의 「전체 선택」「구입 검토함으로 보내기」 같은 버튼을 체크박스가 아니라
 * 막대 전체를 누르면 동작하는 큰 버튼으로 만듭니다.
 * - 「전체 선택」은 칸을 선택하는 순간 동작합니다. (onSelectionChange)
 * - 확인창·점검창이 필요한 막대는 맨 앞 ☐ 칸을 체크하면 동작합니다. (설치형 트리거 onEditTriggerAction)
 *   단순 트리거(onSelectionChange)는 권한이 없어 showModalDialog·alert 를 띄울 수 없기 때문입니다.
 * - 보내기·만들기처럼 되돌리기 번거로운 동작은 「정말 할까요?」를 한 번 물어봅니다.
 * - 「전체 선택/해제」는 막대 앞의 ☐/☑ 로 지금 상태를 보여 줍니다.
 */
var UNIFIED_BARS_VERSION = '3';
// 전체 선택 막대의 두 상태 글자 (앞 그림이 상태를 알려 줍니다)
var BAR_SELECT_ON = '✅  전체 선택';
var BAR_SELECT_OFF = '⬜  전체 해제';
var UNIFIED_BARS_VERSION_PROPERTY = 'UNIFIED_BARS_VERSION';

// 시트 이름 → 막대 모양. 막대의 열 범위는 화면 폭에 맞춰 그릴 때 정해 저장해 두고, 누름 판정도 그 저장값을 씁니다. (kind: select=전체 선택 / send=보내기 / check=점검 시작 / rebuild=다시 만들기)
function unifiedBarSpec_(sheet) {
  var spec = unifiedBarSpecBase_(sheet);
  if (!spec) return null;
  var stored = unifiedStoredBarLayout_(sheet);
  if (stored) {
    spec.bars.forEach(function(bar) {
      var range = stored[bar.id];
      if (range) { bar.c1 = range[0]; bar.c2 = range[1]; }
    });
  }
  return spec;
}

function unifiedBarSpecBase_(sheet) {
  var name = sheet.getName();
  var SEND = '📤  선택한 책 구입 검토함으로 보내기  ›';
  var SELECT = BAR_SELECT_ON;

  if (name === TARGET_SHEET_NAME) {
    var last = Math.max(sheet.getLastColumn(), 12);
    return { last: last, fractions: [1], bars: [
      { id: 'hope-send', kind: 'send', c1: 1, c2: last, label: '📥  신청 목록을 구입 검토함으로 보내기  ›', confirm: '접수된 희망도서를 구입 검토함으로 보낼까요?' }
    ] };
  }
  if (name === SERIES_FINDER_SHEET) {
    return { last: 14, fractions: [0.26, 0.34, 0.40], bars: [
      { id: 'series-select', kind: 'select', c1: 1, c2: 3, label: SELECT },
      { id: 'series-check', kind: 'check', c1: 4, c2: 8, label: '🔎  시리즈 누락 점검 시작  ›' },
      { id: 'series-send', kind: 'send', c1: 9, c2: 14, label: SEND, confirm: '체크한 시리즈 도서를 구입 검토함으로 보낼까요?' }
    ] };
  }
  if (name === UNIFIED_UPLOAD_SHEET) {
    return { last: 11, fractions: [0.30, 0.70], bars: [
      { id: 'upload-select', kind: 'select', c1: 1, c2: 3, label: SELECT },
      { id: 'upload-send', kind: 'send', c1: 4, c2: 11, label: SEND, confirm: '체크한 책을 구입 검토함으로 보낼까요?' }
    ] };
  }
  if (name === UNIFIED_REVIEW_SHEET) {
    return { last: 12, fractions: [0.30, 0.70], bars: [
      { id: 'review-select', kind: 'select', c1: 1, c2: 3, label: SELECT },
      { id: 'review-export', kind: 'send', c1: 4, c2: 12, label: '📦  서점 전달용으로 보내기 (중복 확인)  ›', confirm: '체크한 책의 소장 중복을 확인하고 서점 전달용 목록을 만들까요?' }
    ] };
  }
  if (name === UNIFIED_EXPORT_SHEET) {
    return { last: 10, fractions: [0.42, 0.58], bars: [
      { id: 'export-rebuild', kind: 'send', c1: 1, c2: 5, label: '🔄  서점 전달용 다시 만들기  ›', confirm: '구입 검토함의 최신 선택으로 서점 전달용 목록을 다시 만들까요?' },
      { id: 'export-confirm', kind: 'check', c1: 6, c2: 10, label: '✅  이번 회차 구입 확정  ›' }
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

var UNIFIED_BARS_VIEW_WIDTH = 1400;   // 한 화면에 보이는 폭(픽셀)으로 보고 막대를 나눕니다.

function unifiedBarsPropertyKey_(sheet) { return 'UNIFIED_BARS_' + sheet.getSheetId(); }

function unifiedStoredBarLayout_(sheet) {
  try {
    var raw = PropertiesService.getDocumentProperties().getProperty(unifiedBarsPropertyKey_(sheet));
    return raw ? JSON.parse(raw) : null;
  } catch (error) {
    return null;
  }
}

// 열 폭(픽셀)을 보고, 막대들이 화면에서 비율대로 보이게 각 막대의 시작·끝 열을 정합니다.
// 마지막 막대가 한 화면(약 1400픽셀)을 넘지 않게 하고, 남는 열은 막대를 그리지 않고 비워 둡니다.
function unifiedComputeBarColumns_(sheet, spec) {
  var widths = [], total = 0, column;
  for (column = 1; column <= spec.last; column++) {
    var width = Number(sheet.getColumnWidth(column)) || 100;
    widths.push(width);
    total += width;
  }
  var view = Math.min(total, UNIFIED_BARS_VIEW_WIDTH);
  var layout = {};
  var start = 1;
  var done = 0;      // 이미 쓴 폭
  var shared = 0;    // 지금까지의 비율 합
  spec.bars.forEach(function(bar, index) {
    shared += spec.fractions[index] || 0;
    var target = view * shared;
    var left = spec.bars.length - index - 1;   // 뒤에 남은 막대 수만큼 열을 남겨 둡니다.
    var end = start, best = start, bestGap = Infinity, accumulated = done;
    for (column = start; column <= spec.last - left; column++) {
      accumulated += widths[column - 1];
      var gap = Math.abs(accumulated - target);
      if (gap < bestGap) { bestGap = gap; best = column; }
      if (accumulated > target) break;
    }
    end = best;
    layout[bar.id] = [start, end];
    for (column = start; column <= end; column++) done += widths[column - 1];
    start = end + 1;
  });
  return layout;
}

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
    row.setBackground('#F4F8F5').setBorder(false, false, false, false, false, false);

    // 막대마다 차지할 열을 화면 폭 기준으로 정하고 저장합니다.
    var base = unifiedBarSpecBase_(sheet);
    var layout = unifiedComputeBarColumns_(sheet, base);
    PropertiesService.getDocumentProperties().setProperty(unifiedBarsPropertyKey_(sheet), JSON.stringify(layout));
    spec.bars.forEach(function(bar) { bar.c1 = layout[bar.id][0]; bar.c2 = layout[bar.id][1]; });

    spec.bars.forEach(function(bar) {
      var style = UNIFIED_BAR_STYLE[bar.kind] || UNIFIED_BAR_STYLE.send;
      var range = sheet.getRange(4, bar.c1, 1, bar.c2 - bar.c1 + 1);
      var label = bar.label;
      // 확인창·점검창이 필요한 막대는 맨 앞 칸에 ☐ 체크칸을 둡니다. (체크하면 설치형 트리거가 받아서 창을 띄웁니다)
      if (unifiedBarNeedsCheckbox_(bar)) {
        var box = sheet.getRange(4, bar.c1);
        box.insertCheckboxes().setValue(false);
        range
          .setBackground(style.bg)
          .setFontColor(style.fg)
          .setFontSize(20)
          .setHorizontalAlignment('center')
          .setVerticalAlignment('middle')
          .setBorder(true, true, true, true, false, false, style.border, SpreadsheetApp.BorderStyle.SOLID_THICK);
        if (bar.c2 === bar.c1) { box.setNote(bar.label + '\n(☐ 를 체크하면 실행됩니다)'); return; }
        range = sheet.getRange(4, bar.c1 + 1, 1, bar.c2 - bar.c1);
        label = '  ' + bar.label;
      }
      if (range.getNumColumns() > 1) range.merge();
      range
        .setValue(label)
        .setBackground(style.bg)
        .setFontColor(style.fg)
        .setFontFamily('Arial')
        .setFontSize(style.size)
        .setFontWeight('bold')
        .setHorizontalAlignment('center')
        .setVerticalAlignment('middle')
        .setBorder(true, true, true, true, false, false, style.border, SpreadsheetApp.BorderStyle.SOLID_THICK);
      if (unifiedBarNeedsCheckbox_(bar) && bar.c2 > bar.c1) {
        // ☐ 칸과 글자가 한 버튼으로 보이도록: 사이 테두리를 없애고, ☐ 는 오른쪽·글자는 왼쪽으로 붙입니다.
        sheet.getRange(4, bar.c1).setHorizontalAlignment('right').setBorder(null, null, null, false, null, null);
        range.setHorizontalAlignment('left').setBorder(null, false, null, null, null, null);
      }
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
  // 확인창·점검창이 필요한 막대는 칸 선택(단순 트리거)으로는 창을 띄울 수 없어서, 맨 앞 ☐ 체크칸으로 안내만 합니다.
  if (unifiedBarNeedsCheckbox_(bar)) return unifiedGuideBarCheckbox_(sheet, range, bar);
  return unifiedRunBarAction_(sheet, bar);
}

// 막대의 실제 동작. (전체 선택은 칸 선택에서, 나머지는 체크칸을 체크했을 때 설치형 트리거에서 부릅니다)
function unifiedRunBarAction_(sheet, bar) {
  // 같은 클릭이 두 번 처리되지 않게 잠깐 표시를 남깁니다.
  var cache = CacheService.getDocumentCache();
  var busyKey = 'BAR_BUSY_' + bar.id;
  if (cache && cache.get(busyKey)) return true;
  if (cache) cache.put(busyKey, '1', 6);

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  // 눌렀다는 것을 바로 알려 줍니다. (확인 창이 뜨기 전에도 반응이 보이도록)
  try { ss.toast('누르신 버튼을 처리하고 있습니다…', bar.label.replace(/\s*›\s*$/, ''), 3); } catch (error) {}
  try {
    if (bar.kind === 'select') {
      var cell = sheet.getRange(4, bar.c1);
      var checked = String(cell.getValue() || '').indexOf('⬜') !== 0;   // 지금 「전체 선택」 상태가 아니면 모두 선택
      toggleSelectionColumn_(sheet, 6, 1, checked);
      cell.setValue(checked ? BAR_SELECT_OFF : BAR_SELECT_ON);
      if (sheet.getName() === UNIFIED_REVIEW_SHEET) unifiedUpdateDashboardNumbers_();
    } else if (bar.id === 'export-confirm') {
      unifiedConfirmPurchaseRound_();   // 안에서 직접 확인을 묻고 기록합니다.
    } else if (bar.id === 'series-check') {
      unifiedStartSeriesCheckFromBar_();   // 창 없이 바로 점검합니다. (진행 상황은 시트 오른쪽 아래 알림으로)
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

/* ---------- ☐ 체크칸으로 실행하는 막대 ---------- */

// 확인창·점검창을 띄우는 막대인지 (전체 선택만 칸 선택으로 바로 동작합니다)
function unifiedBarNeedsCheckbox_(bar) {
  return !!bar && bar.kind !== 'select';
}

function unifiedBarHasCheckbox_(sheet, bar) {
  try {
    var rule = sheet.getRange(4, bar.c1).getDataValidation();
    return !!rule && rule.getCriteriaType() === SpreadsheetApp.DataValidationCriteria.CHECKBOX;
  } catch (error) {
    return false;
  }
}

// 막대 글자 부분을 눌렀을 때(단순 트리거): 맨 앞 ☐ 칸으로 안내합니다.
function unifiedGuideBarCheckbox_(sheet, range, bar) {
  if (range.getColumn() === bar.c1 && range.getLastColumn() === bar.c1) return true;   // 체크칸 자체를 고른 것
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var title = bar.label.replace(/\s*›\s*$/, '');
  if (!unifiedBarHasCheckbox_(sheet, bar)) {
    try { ss.toast('버튼 모양이 바뀌었습니다. 시트를 닫았다가 다시 열면 막대 앞에 ☐ 칸이 생깁니다. 그동안은 위쪽 메뉴를 이용해 주세요.', title, 8); } catch (error) {}
    return true;
  }
  try { ss.toast('막대 맨 앞의 ☐ 칸을 체크하면 실행됩니다.', title, 5); } catch (error) {}
  try { sheet.getRange(4, bar.c1).activate(); } catch (error) {}
  return true;
}

// 막대 맨 앞 ☐ 칸을 체크했을 때. 설치형 트리거(onEditTriggerAction)에서 부르므로 확인창·점검창을 띄울 수 있습니다. 처리했으면 true.
function unifiedHandleBarEdit_(e) {
  var range = e && e.range;
  if (!range || range.getRow() !== 4 || range.getNumRows() !== 1 || range.getNumColumns() !== 1) return false;
  var sheet = range.getSheet();
  var bar = unifiedBarAt_(sheet, range);
  if (!bar || !unifiedBarNeedsCheckbox_(bar) || range.getColumn() !== bar.c1) return false;
  if (String(e.value || '').toUpperCase() !== 'TRUE') return true;
  range.setValue(false);
  unifiedRunBarAction_(sheet, bar);
  return true;
}

// 「🔎 시리즈 누락 점검 시작」 ☐ 를 체크했을 때: 창 없이 바로 점검합니다.
// 진행 상황·결과는 refreshMissingSeriesBooks 가 시트 오른쪽 아래 알림으로 보여 주고,
// 오래 걸리면 중간에 멈췄다가 약 1분 뒤 저절로 이어서 합니다. (창으로 보고 싶으면 메뉴의 「시리즈 누락 점검 (쉬운 창)」)
function unifiedStartSeriesCheckFromBar_() {
  var ss = seriesSpreadsheet_();
  if (!seriesGetYes24Key_()) {
    ss.toast('YES24 API 키가 아직 없습니다. 「⚙️ 사본 설정 → 🔑 API 키 설정」에서 먼저 입력해 주세요.', '🔎 시리즈 누락 점검', 10);
    return;
  }
  ss.toast('점검을 시작했습니다. 다른 작업을 하셔도 되고, 끝나면 여기에 알림이 뜹니다.', '🔎 시리즈 누락 점검', 8);
  refreshMissingSeriesBooks();
}
