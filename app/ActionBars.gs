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
var UNIFIED_BARS_VERSION = '5';
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

// 모든 막대는 같은 높이·같은 모양이고, 역할에 따라 색만 다릅니다.
// tab: ☐ 체크칸 바탕(막대보다 한 톤 진하게 — 「여기를 체크」가 한눈에 보이도록)
var UNIFIED_BAR_STYLE = {
  select:  { bg: '#CFE6DA', tab: '#B7D9C7', fg: '#14493A', border: '#2E7D5B', size: 16 },
  send:    { bg: '#B9770E', tab: '#955F0A', fg: '#FFFFFF', border: '#7A4A06', size: 16 },
  check:   { bg: '#1F5C4A', tab: '#16463A', fg: '#FFFFFF', border: '#0F3A2E', size: 16 }
};
UNIFIED_BAR_STYLE.pending = { bg: '#B3261E', fg: '#FFFFFF' };   // 「한 번 더 누르면 실행합니다」 상태
var UNIFIED_BAR_MIN_FONT = 11;

// 글자 폭 어림(px). 한글 1em, 영문·숫자 0.6em, 띄어쓰기 0.3em, 그림 문자 1.25em (굵은 글꼴 기준)
function unifiedBarTextWidth_(text, size) {
  var em = 0;
  String(text || '').replace(/[\uD800-\uDBFF][\uDC00-\uDFFF]|[\s\S]/g, function(ch) {
    if (ch.length === 2 || /[☀-➿⬀-⯿]/.test(ch)) em += 1.25;
    else if (/\s/.test(ch)) em += 0.3;
    else if (/[ᄀ-ᇿ　-鿿가-힯＀-￯]/.test(ch)) em += 1;
    else em += 0.6;
    return ch;
  });
  return em * size * 4 / 3;
}

// 막대 글자 칸 폭에 맞는 가장 큰 글자 크기 (max 에서 시작해 줄여 갑니다)
function unifiedBarFitFontSize_(sheet, c1, c2, label, max) {
  var width = 0;
  for (var c = c1; c <= c2; c++) width += Number(sheet.getColumnWidth(c)) || 100;
  width -= 28;   // 양쪽 여백
  for (var size = max; size > UNIFIED_BAR_MIN_FONT; size--) {
    if (unifiedBarTextWidth_(label, size) <= width) return size;
  }
  return UNIFIED_BAR_MIN_FONT;
}

var UNIFIED_BARS_VIEW_WIDTH = 1480;   // 한 화면에 보이는 폭(픽셀)으로 보고 막대를 나눕니다. (마지막 막대 글자가 잘리지 않을 만큼)

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

    // 글자 칸 범위를 먼저 정하고, 한 시트 안의 막대는 모두 같은 글자 크기(가장 좁은 막대에 맞춤)로 씁니다.
    var fontSize = 16;
    spec.bars.forEach(function(bar) {
      var style = UNIFIED_BAR_STYLE[bar.kind] || UNIFIED_BAR_STYLE.send;
      bar.t1 = unifiedBarNeedsCheckbox_(bar) && bar.c2 > bar.c1 ? bar.c1 + 1 : bar.c1;
      fontSize = Math.min(fontSize, unifiedBarFitFontSize_(sheet, bar.t1, bar.c2, '  ' + bar.label, style.size));
    });

    var EDGE = SpreadsheetApp.BorderStyle.SOLID_MEDIUM;
    spec.bars.forEach(function(bar) {
      var style = UNIFIED_BAR_STYLE[bar.kind] || UNIFIED_BAR_STYLE.send;
      var whole = sheet.getRange(4, bar.c1, 1, bar.c2 - bar.c1 + 1);
      whole
        .setBackground(style.bg)
        .setFontColor(style.fg)
        .setVerticalAlignment('middle')
        .setBorder(true, true, true, true, false, false, style.border, EDGE);

      if (unifiedBarNeedsCheckbox_(bar)) {
        // 맨 앞 칸: 한 톤 진한 바탕, 큰 ☐ 를 가운데에 (체크하면 설치형 트리거가 받아서 창을 띄웁니다)
        var box = sheet.getRange(4, bar.c1);
        box.insertCheckboxes().setValue(false);
        box.setBackground(style.tab).setFontColor(style.fg).setFontSize(22).setHorizontalAlignment('center');
        box.setNote(null);
        if (bar.c2 === bar.c1) { box.setNote(bar.label.replace(/\s*›\s*$/, '').trim() + '\n(☐ 를 체크하면 실행됩니다)'); return; }
      }

      var text = sheet.getRange(4, bar.t1, 1, bar.c2 - bar.t1 + 1);
      if (text.getNumColumns() > 1) text.merge();
      text
        .setValue('  ' + bar.label)   // 왼쪽 정렬 + 앞 여백
        .setFontFamily('Arial')
        .setFontSize(fontSize)
        .setFontWeight('bold')
        .setHorizontalAlignment(unifiedBarNeedsCheckbox_(bar) ? 'left' : 'center')   // ☐ 가 있는 막대는 ☐ 만 가운데, 글자는 왼쪽
        .setWrap(false);
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
      if (bar.kind === 'select') sheet.getRange(4, bar.c1).setValue('  ' + BAR_SELECT_ON);
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
// 막대는 체크칸 없이 그냥 누르면 됩니다. 되돌리기 번거로운 동작은 막대 글자가 「한 번 더 누르면 실행합니다」로 바뀌고,
// 한 번 더 누르면 실행됩니다. (칸 선택 이벤트에서는 확인 창을 띄울 수 없어서 창 대신 이렇게 확인합니다)
// 예외: 「시리즈 누락 점검 시작」은 YES24 에 접속해야 해서 칸 선택 이벤트로는 할 수 없어, 앞의 ☐ 를 체크하는 방식입니다.
function unifiedHandleBarClick_(sheet, range) {
  var bar = unifiedBarAt_(sheet, range);
  unifiedRevertPendingBars_(sheet, bar);   // 다른 곳을 누르면 「한 번 더」 상태를 취소합니다.
  if (!bar) return false;
  if (unifiedBarNeedsCheckbox_(bar)) return unifiedGuideBarCheckbox_(sheet, range, bar);
  if (bar.kind === 'select') return unifiedRunBarAction_(sheet, bar);
  return unifiedClickWithConfirm_(sheet, bar);
}

var UNIFIED_BAR_PENDING_SECONDS = 25;   // 「한 번 더」 확인을 기다리는 시간
var UNIFIED_BAR_NO_DIALOG = false;      // 막대 클릭으로 실행 중에는 확인 창(ui.alert)을 쓰지 않습니다.

function unifiedBarPendingKey_(sheet, bar) { return 'BAR_PENDING_' + sheet.getSheetId() + '_' + bar.id; }

// 첫 번째 누름: 막대를 「한 번 더」 상태로 바꿉니다. 두 번째 누름: 실행합니다.
function unifiedClickWithConfirm_(sheet, bar) {
  var cache = CacheService.getDocumentCache();
  var key = unifiedBarPendingKey_(sheet, bar);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var pendingAt = cache ? Number(cache.get(key) || 0) : 0;
  var title = bar.label.replace(/\s*›\s*$/, '').trim();

  if (pendingAt && Date.now() - pendingAt < UNIFIED_BAR_PENDING_SECONDS * 1000) {
    cache.remove(key);
    unifiedRestoreBarLook_(sheet, bar);
    UNIFIED_BAR_NO_DIALOG = true;
    try { return unifiedRunBarAction_(sheet, bar); } finally { UNIFIED_BAR_NO_DIALOG = false; }
  }

  if (cache) cache.put(key, String(Date.now()), UNIFIED_BAR_PENDING_SECONDS + 5);
  var style = UNIFIED_BAR_STYLE.pending;
  var cell = sheet.getRange(4, bar.c1);
  cell.setValue('⚠  한 번 더 누르면 실행합니다').setBackground(style.bg).setFontColor(style.fg);
  var detail = unifiedBarConfirmDetail_(bar);
  try { ss.toast(detail + '\n(그만두려면 다른 칸을 누르세요)', title, 8); } catch (error) {}
  // 같은 막대를 또 누르면 다시 이벤트가 오도록 선택을 막대 밖으로 옮깁니다.
  try { sheet.setActiveRange(sheet.getRange(6, 2)); } catch (error) {}
  return true;
}

// 첫 번째 누름에서 알려 줄 내용(무엇을 하는지, 주의할 점)
function unifiedBarConfirmDetail_(bar) {
  var text = bar.confirm || '이 작업을 실행할까요?';
  try {
    if (bar.id === 'review-export' || bar.id === 'export-rebuild') {
      var review = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(UNIFIED_REVIEW_SHEET);
      if (review && review.getLastRow() >= 6) {
        var rows = review.getRange(6, 1, review.getLastRow() - 5, 12).getValues();
        var workbooks = 0;
        rows.forEach(function(row) {
          var selected = row[0] === true || String(row[2]) === '구입결정';
          if (selected && unifiedFlagsFromNote_(row[11]).workbook) workbooks++;
        });
        if (workbooks) text += '\n⚠ 문제집·학습지로 보이는 책이 ' + workbooks + '권 있습니다. (그대로 포함됩니다)';
      }
    } else if (bar.id === 'export-confirm') {
      var round = unifiedPurchaseLogStats_(SpreadsheetApp.getActiveSpreadsheet()).rounds + 1;
      text = '서점 전달용 목록을 ' + round + '회차 구입으로 기록하고, 구입 검토함의 해당 책을 「주문완료」로 바꿉니다.';
    }
  } catch (error) {}
  return text;
}

// 막대를 원래 글자·색으로 되돌립니다.
function unifiedRestoreBarLook_(sheet, bar) {
  var style = UNIFIED_BAR_STYLE[bar.kind] || UNIFIED_BAR_STYLE.send;
  sheet.getRange(4, bar.c1)
    .setValue('  ' + bar.label)
    .setBackground(style.bg)
    .setFontColor(style.fg);
}

// 「한 번 더」 상태로 남은 막대가 있으면 되돌립니다. (방금 누른 막대와, 막 설정한 직후의 이벤트는 건드리지 않습니다)
function unifiedRevertPendingBars_(sheet, exceptBar) {
  try {
    var cache = CacheService.getDocumentCache();
    if (!cache) return;
    var spec = unifiedBarSpec_(sheet);
    if (!spec) return;
    spec.bars.forEach(function(bar) {
      if (exceptBar && exceptBar.id === bar.id) return;
      if (bar.kind === 'select' || unifiedBarNeedsCheckbox_(bar)) return;
      var key = unifiedBarPendingKey_(sheet, bar);
      var at = Number(cache.get(key) || 0);
      if (!at) return;
      if (Date.now() - at < 2000) return;   // 선택을 막대 밖으로 옮긴 직후에 오는 이벤트
      cache.remove(key);
      unifiedRestoreBarLook_(sheet, bar);
    });
  } catch (error) {
    console.warn('막대 되돌리기 실패: ' + error);
  }
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
      var checked = String(cell.getValue() || '').trim().indexOf('⬜') !== 0;   // 지금 「전체 선택」 상태가 아니면 모두 선택
      toggleSelectionColumn_(sheet, 6, 1, checked);
      cell.setValue('  ' + (checked ? BAR_SELECT_OFF : BAR_SELECT_ON));
      if (sheet.getName() === UNIFIED_REVIEW_SHEET) unifiedUpdateDashboardNumbers_();
    } else if (bar.id === 'export-confirm') {
      unifiedConfirmPurchaseRound_();   // 안에서 직접 확인을 묻고 기록합니다.
    } else if (bar.id === 'series-check') {
      HR_openSeriesCheckDialog_(false);   // ☐ 체크로 부르므로 창을 띄울 수 있습니다. 창에서 「점검 시작/이어서 점검」을 누릅니다.
    } else {
      var answer = true;
      if (!UNIFIED_BAR_NO_DIALOG) {   // 메뉴·☐ 체크로 부른 경우에만 확인 창을 띄웁니다.
        var ui = SpreadsheetApp.getUi();
        answer = ui.alert('확인', bar.confirm, ui.ButtonSet.YES_NO) === ui.Button.YES;
      }
      if (answer) {
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
  try { seriesEnsureAutoContinue_(); } catch (error) { console.warn('이어가기 예약 점검 실패: ' + error); }
  try { seriesShowStateStatus_(); } catch (error) {}
  var resume = null;
  try {
    var parsed = JSON.parse(props.getProperty(SERIES_RESUME_STATE_KEY) || 'null');
    if (parsed && Number(parsed.total) > 0) resume = { offset: Number(parsed.offset || 0), total: Number(parsed.total || 0) };
  } catch (error) {}
  var ss = seriesSpreadsheet_();
  var result = ss.getSheetByName(SERIES_FINDER_SHEET);
  var holdings = ss.getSheetByName(SERIES_HOLDINGS_SHEET);
  var quotaNext = null;
  try {
    var parsedNext = JSON.parse(props.getProperty('YES24_SERIES_QUOTA_NEXT') || 'null');
    if (parsedNext && parsedNext.at) quotaNext = { at: Number(parsedNext.at), text: String(parsedNext.text || '') };
  } catch (error) {}
  return {
    hasKey: Boolean(seriesGetYes24Key_()),
    quotaNext: quotaNext,
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

// ☐ 체크칸이 필요한 막대인지. YES24 에 접속하는 「시리즈 누락 점검 시작」만 그렇습니다.
// (칸 선택 이벤트에서는 외부 접속·확인 창이 막혀 있어서, 체크 이벤트로만 할 수 있습니다. 나머지 막대는 그냥 누르면 됩니다.)
function unifiedBarNeedsCheckbox_(bar) {
  return !!bar && bar.id === 'series-check';
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
