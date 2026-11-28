/* 자동 생성 파일 — 직접 고치지 마세요. (node tools/build.js)  vreal-1 */
var APP_FACTORY = function (__host) {
  var HtmlService = __host.html;
  var SpreadsheetApp = __host.spreadsheet;

/* ===== ActionBars.gs ===== */
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


/* ===== ApplyLinks.gs ===== */
/**
 * 🔗 신청 링크 받기
 *
 * 교직원용·학부모용 신청 링크를 한 창에서 보여 주고, 복사하거나 QR 코드를 그림(PNG)으로 저장하게 합니다.
 * QR 은 창 안에서 직접 만들기 때문에 학교 주소가 외부 서비스로 나가지 않습니다.
 */
function HR_showApplyLinksDialog() {
  var props = PropertiesService.getScriptProperties();
  var registered = HR_getRegisteredWebAppUrl_();
  var detected = HR_getDetectedWebAppUrl_();

  // 등록한 주소가 없을 때만, 찾은 주소를 「공개 링크」로 저장해 둡니다. (등록한 주소는 절대 덮어쓰지 않습니다.)
  if (!registered && detected) {
    props.setProperty(HR_PUBLIC_WEB_APP_URL_PROPERTY, detected);
  }

  var baseUrl = registered || detected;
  var problem = '';
  var switched = false;

  // 등록한 주소가 실제로 열리는지 확인합니다.
  if (baseUrl) {
    var status = HR_probeWebAppUrl_(baseUrl);
    if (status === 'notfound' || status === 'restricted') {
      // 등록한 주소는 안 열리는데 구글이 알려 주는 주소가 열리면, 그 주소로 바꿔서 보여 줍니다.
      if (detected && detected !== baseUrl && HR_probeWebAppUrl_(detected) === 'ok') {
        baseUrl = detected;
        switched = true;
        props.setProperty(HR_PUBLIC_WEB_APP_URL_PROPERTY, detected);
      } else {
        problem = HR_probeProblemMessage_(status);
      }
    }
  }

  var data = {
    baseUrl: problem ? '' : baseUrl,
    detected: !!baseUrl,
    mismatch: !!(registered && detected && registered !== detected && !switched),
    switched: switched,
    problem: problem,
    badUrl: problem ? baseUrl : '',
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
  var head = who + ' 최대 ' + limit + '권까지 신청할 수 있습니다.';

  if (remain <= 0) {
    return head + ' 이미 ' + used + '권을 모두 신청하셨습니다.';
  }
  if (used === 0) {
    return head + ' 지금 ' + addCount + '권을 담으셨습니다. ' + limit + '권 이하로 줄여서 다시 신청해 주십시오.';
  }
  return head + ' 이미 ' + used + '권을 신청하셨으므로 이번에는 ' + remain + '권까지만 가능합니다. 담은 ' +
    addCount + '권을 ' + remain + '권 이하로 줄여 주십시오.';
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
      '1인당 신청(구입)할 수 있는 권수를 숫자로 입력하십시오.\n' +
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
      ui.alert('숫자를 0~500 사이로 입력해 주십시오.');
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

// 웹 화면에서 부릅니다: 학부모가 고른 학년·반·번호의 신청 한도와 남은 권수를 알려 줍니다.
// (신청한 책 제목은 돌려주지 않고 권수만 돌려줍니다.)
function HR_getParentQuota(grade, classNo, number) {
  var limit = HR_getRequestLimit_('parent');
  var g = Number(grade), c = Number(classNo), n = Number(number);
  var result = { limit: limit, used: 0, remain: limit, known: false };
  if (!limit || !g || !c || !n) return result;

  var sheet = HR_getSpreadsheet_().getSheetByName(HR_TARGET_SHEET_NAME);
  if (!sheet) return result;
  var used = HR_countApplicantRequests_(sheet, 'parent', g + '학년 ' + c + '반 ' + n + '번');
  result.used = used;
  result.remain = Math.max(0, limit - used);
  result.known = true;
  return result;
}

// 학부모 화면의 「반」 선택칸에 보여 줄 최대 반 수 (학교 설정에서 정함, 안 정했으면 20)
var HR_MAX_CLASSES_PROPERTY = 'HR_MAX_CLASSES';
function HR_getMaxClasses_() {
  var raw = Number(PropertiesService.getScriptProperties().getProperty(HR_MAX_CLASSES_PROPERTY));
  return isFinite(raw) && raw >= 1 && raw <= 30 ? Math.floor(raw) : 20;
}


/* ===== BookFlags.gs ===== */
/**
 * 🏷️ 책 종류 표시: 만화 · 문제집·학습지
 *
 * 구입 검토함 · 서점 전달용 · 시리즈 후보 · 업로드 목록의 「비고」 맨 앞에 표시를 붙입니다.
 *   📙 만화                   : 만화·학습만화·웹툰 등 (읽는 책이라 사도 되지만 구분해서 보기 위해)
 *   ⚠ 문제집·학습지 의심       : 문제집·학습지·워크북·기출·수능·자습서 등 (실수로 사지 않도록 경고)
 * 서점 전달용을 만들기 전에는 문제집·학습지로 보이는 책이 있으면 한 번 더 확인을 묻습니다.
 * 제목(그리고 있으면 YES24 분야)만 보고 판단하는 추정이라, 틀릴 수 있으면 「비고」에서 표시를 지우면 됩니다.
 */
var BOOK_FLAG_COMIC = '📙 만화';
var BOOK_FLAG_WORKBOOK = '⚠ 문제집·학습지 의심';

function unifiedBookFlags_(title, category) {
  var text = String(title || '');
  var cat = String(category || '');
  var compact = text.replace(/\s+/g, '');

  var comic = /만화|코믹|웹툰|카툰|그래픽\s*노블/.test(text) ||
    /만화|웹툰|코믹/.test(cat) ||
    /^(Why\?|와이|살아남기|마법천자문|수학도둑|과학도둑|먼나라이웃나라|한국사대모험)/.test(compact.replace(/^\(.*?\)/, ''));

  var strong = /문제집|학습지|워크북|익힘책|자습서|평가문제|단원평가|기출|모의고사|수능\s*(?:특강|완성|대비|연계|개념|문제|국어|영어|수학|한국사|탐구|영역|독해)|내신\s*(?:대비|문제|완성|만점|필수|\d\s*등급)|오답\s*노트|받아쓰기|따라\s*쓰기|쓰기\s*노트|연산\s*(?:문제|훈련|노트|학습)|EBS|개념\s*원리|디딤돌|우공비|마더텅|완자|기본서|급수\s*한자|한자\s*쓰기/i.test(text) ||
    /(?:^|[\s\[(])쎈(?:[\s\])]|$)/.test(text) ||
    /참고서|수험서|자격증|문제집|학습지|교재/.test(cat);
  // 「초등 3학년 수학」처럼 학년·과목만 있는 제목은 이야기책일 수도 있어서, 읽을거리 말이 없을 때만 의심합니다.
  var gradeSubject = /(?:초등|중등|중학|고등)?\s*\d\s*학년\s*(?:\d\s*학기\s*)?(?:국어|수학|영어|과학|사회)/.test(text) &&
    !/이야기|동화|소설|위인|탐험|모험|그림책|만화|인물|여행/.test(text);

  return { comic: Boolean(comic), workbook: Boolean(strong || gradeSubject) };
}

// 비고 글 맨 앞에 붙은 표시들을 읽습니다.
function unifiedFlagsFromNote_(note) {
  var text = String(note || '');
  return { comic: text.indexOf(BOOK_FLAG_COMIC) >= 0, workbook: text.indexOf(BOOK_FLAG_WORKBOOK) >= 0 };
}

function unifiedFlagTagText_(flags) {
  var tags = [];
  if (flags && flags.workbook) tags.push(BOOK_FLAG_WORKBOOK);
  if (flags && flags.comic) tags.push(BOOK_FLAG_COMIC);
  return tags.join(' · ');
}

// 비고에 없는 표시만 맨 앞에 붙입니다. (이미 붙은 표시는 그대로 두고, 선생님이 지운 표시는 다시 붙입니다.)
function unifiedComposeNote_(note, flags) {
  var current = String(note || '').trim();
  var have = unifiedFlagsFromNote_(current);
  var add = unifiedFlagTagText_({ workbook: flags.workbook && !have.workbook, comic: flags.comic && !have.comic });
  if (!add) return current;
  return current ? add + ' · ' + current : add;
}

// 구입 검토함 비고 열(L)에 표시를 붙입니다. 바뀐 줄이 있을 때만 씁니다. 문제집·학습지 의심 권수를 돌려줍니다.
function unifiedApplyBookFlagsToReview_(sheet) {
  if (!sheet || sheet.getLastRow() < 6) return 0;
  var count = sheet.getLastRow() - 5;
  var rows = sheet.getRange(6, 1, count, 12).getValues();
  var notes = [], changed = false, workbooks = 0;
  rows.forEach(function(row) {
    var title = String(row[3] || '').trim();
    var note = String(row[11] || '');
    if (!title) { notes.push([note]); return; }
    var next = unifiedComposeNote_(note, unifiedBookFlags_(title, ''));
    if (next !== note) changed = true;
    if (unifiedFlagsFromNote_(next).workbook) workbooks++;
    notes.push([next]);
  });
  if (changed) sheet.getRange(6, 12, count, 1).setValues(notes);
  return workbooks;
}

// 서점 전달용을 만들기 전 확인: 문제집·학습지로 보이는 책을 보여 주고 그래도 포함할지 묻습니다.
// 돌려주는 값: true=계속, false=멈춤
function unifiedConfirmWorkbooks_(books) {
  var titles = [];
  (books || []).forEach(function(row) {
    var flags = unifiedBookFlags_(row[3], '');
    var noted = unifiedFlagsFromNote_(row[11]);
    if (flags.workbook || noted.workbook) titles.push(String(row[3] || '').trim());
  });
  if (!titles.length) return true;

  if (typeof UNIFIED_BAR_NO_DIALOG !== 'undefined' && UNIFIED_BAR_NO_DIALOG) return true;   // 막대 클릭으로 실행 중: 창 없이 계속(첫 번째 누름에서 알려 줬습니다)
  var ui;
  try { ui = SpreadsheetApp.getUi(); } catch (error) { return true; }
  var list = titles.slice(0, 10).map(function(title) { return '· ' + title; }).join('\n') +
    (titles.length > 10 ? '\n… 외 ' + (titles.length - 10) + '권' : '');
  var answer = ui.alert(
    '⚠ 문제집·학습지로 보이는 책이 ' + titles.length + '권 있습니다',
    list + '\n\n그래도 서점 전달용에 포함할까요?\n(아니오를 누르면 만들기를 멈춥니다. 구입 검토함에서 체크를 풀거나 「보류」로 바꿔 주세요.\n문제집이 아닌 책이면 비고의 「⚠ 문제집·학습지 의심」 표시를 지우면 됩니다.)',
    ui.ButtonSet.YES_NO
  );
  return answer === ui.Button.YES;
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
    if (typeof seriesEnsureAutoContinue_ === 'function') seriesEnsureAutoContinue_();
  } catch (error) {
    console.warn('시리즈 이어가기 예약 확인 실패: ' + error);
  }
  try {
    if (typeof seriesShowStateStatus_ === 'function') seriesShowStateStatus_();
  } catch (error) {}
  try {
    HR_refreshGuidesOnce_();
  } catch (error) {
    console.warn('안내 시트 새 내용 반영 실패: ' + error);
  }
  try {
    HR_arrangeTabsOnce_();
  } catch (error) {
    console.warn('시트 탭 정리 실패: ' + error);
  }
  try {
    HR_showShortcutOnOpen_();
  } catch (error) {
    console.warn('신청 바로가기 창 표시 실패: ' + error);
  }
  try {
    HR_showUpdateNoticeOnOpen_();
  } catch (error) {
    console.warn('업데이트 소식 표시 실패: ' + error);
  }
  try {
    HR_notifyInquiryAnswersOnOpen_();
  } catch (error) {
    console.warn('문의 답변 확인 실패: ' + error);
  }
  try {
    ensureSelectionControls_();
  } catch (error) {
    console.warn('선택 버튼 표시 실패: ' + error);
  }
  try {
    HR_ensureChangeTrigger_();
  } catch (error) {
    console.warn('변경 트리거 준비 실패: ' + error);
  }
  try {
    unifiedUpdateDashboardNumbers_(true);   // 열 때마다 대시보드 숫자를 최신으로
  } catch (error) {
    console.warn('대시보드 숫자 갱신 실패: ' + error);
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

  unifiedDrawActionBars_(sheet);   // 체크박스 대신 클릭 막대
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

// 고친 칸이 6행 이하이고 지정한 열(번호) 중 하나에 걸치는지
function unifiedEditTouches_(range, columns) {
  if (range.getLastRow() < 6) return false;
  var first = range.getColumn(), last = range.getLastColumn();
  return columns.some(function(column) { return column >= first && column <= last; });
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

  // 맨 위 막대의 ☐ 칸은 설치형 트리거(onEditTriggerAction)가 처리합니다. (예전 A4·D4 체크박스 동작과 겹치지 않게)
  if (range.getRow() === 4 && unifiedBarNeedsCheckbox_(unifiedBarAt_(sheet, range))) return;

  // 대시보드의 「총 예산」 칸에 직접 입력
  if (sheetName === UNIFIED_DASHBOARD_SHEET && cell === 'A17') {
    try {
      unifiedBudgetFromDashboard_(range);
    } catch (error) {
      console.warn('예산 입력 처리 실패: ' + error);
    }
    return;
  }

  // 대시보드 숫자 바로 갱신: 검토함 체크·금액·수량을 바꿨을 때, 서점 전달용 정가·수량·금액을 고쳤을 때
  // (A4 전체선택은 아래에서 처리한 직후에 갱신합니다.)
  if (cell !== 'A4' || !checkboxChanged) {
    try {
      if (sheetName === UNIFIED_REVIEW_SHEET && unifiedEditTouches_(range, [1, 3, 9, 10, 11, 12])) unifiedUpdateDashboardNumbers_();
      else if (sheetName === UNIFIED_EXPORT_SHEET && unifiedEditTouches_(range, [7, 8, 9])) unifiedUpdateDashboardNumbers_();
      else if (
        (sheetName === TARGET_SHEET_NAME || sheetName === UNIFIED_UPLOAD_SHEET || sheetName === UNIFIED_REVIEW_SHEET) &&
        range.getLastRow() >= 6 && unifiedEditTouches_(range, sheetName === UNIFIED_REVIEW_SHEET ? [4] : [2])
      ) unifiedUpdateDashboardNumbers_(true);   // 도서명을 쓰거나 지웠을 때 모인 책 권수도 다시 셉니다.
      else if (sheetName === UNIFIED_PURCHASE_LOG_SHEET && range.getLastRow() >= 5 && range.getColumn() <= 4 && range.getLastColumn() >= 3) unifiedUpdateDashboardNumbers_();
    } catch (error) {
      console.warn('대시보드 숫자 갱신 실패: ' + error);
    }
  }

  // 시리즈 누락 점검: 목록 전체 선택/전체 해제
  if (sheetName === SERIES_FINDER_SHEET && cell === 'A4' && checkboxChanged) {
    toggleSelectionColumn_(sheet, 6, 1, checked);
    return;
  }

  // 구입 검토함: 목록 전체 선택/전체 해제
  if (sheetName === UNIFIED_REVIEW_SHEET && cell === 'A4' && checkboxChanged) {
    toggleSelectionColumn_(sheet, 6, 1, checked);
    try {
      unifiedUpdateDashboardNumbers_();
    } catch (error) {
      console.warn('대시보드 숫자 갱신 실패: ' + error);
    }
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

  unifiedDrawActionBars_(sheet);   // 체크박스 대신 클릭 막대
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
// 시트 구조가 바뀌었을 때(줄 삭제·추가 등) 대시보드의 책 권수를 다시 셉니다.
// 줄을 지우는 것은 「편집」이 아니라서 onEdit 에는 오지 않으므로, 같은 입구(onEditTriggerAction)에 「변경」 트리거를 따로 겁니다.
function unifiedHandleStructureChange_(e) {
  var type = String(e && e.changeType || '');
  if (!/REMOVE_ROW|INSERT_ROW|REMOVE_COLUMN|INSERT_COLUMN|OTHER/.test(type)) return;
  try {
    unifiedUpdateDashboardNumbers_(true);
  } catch (error) {
    console.warn('대시보드 갱신 실패(구조 변경): ' + error);
  }
}

// 「변경」 트리거가 없으면 만듭니다. (시트를 열 때 한 번 확인)
function HR_ensureChangeTrigger_() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var exists = ScriptApp.getProjectTriggers().some(function(trigger) {
    return trigger.getHandlerFunction() === 'onEditTriggerAction' &&
      String(trigger.getEventType()) === String(ScriptApp.EventType.ON_CHANGE);
  });
  if (exists) return false;
  ScriptApp.newTrigger('onEditTriggerAction')
    .forSpreadsheet(spreadsheet)
    .onChange()
    .create();
  return true;
}

function onEditTriggerAction(e) {
  if (e && e.changeType) {
    unifiedHandleStructureChange_(e);
    return;
  }
  if (!e || !e.range) return;

  // 맨 위 막대의 ☐ 칸 (확인창·점검창은 여기서만 띄울 수 있습니다)
  if (unifiedHandleBarEdit_(e)) return;

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
    unifiedDrawActionBars_(exportSheet);
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
    try {
      unifiedApplyBookFlagsToReview_(reviewSheet);   // 이미 모여 있던 책에도 만화·문제집 표시를 붙입니다.
    } catch (error) {
      console.warn('책 종류 표시 실패: ' + error);
    }
  }

  // 업로드 목록 시트는 파일을 올리기 전에도 미리 만들어 둡니다.
  try {
    unifiedEnsureUploadSheetReady_();
  } catch (error) {
    console.warn('업로드 목록 시트 준비 실패: ' + error);
  }
}


function toggleSelectionColumn_(sheet, startRow, column, checked) {
  const lastRow = sheet.getLastRow();
  if (lastRow < startRow) {
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
    unifiedDrawActionBars_(sheet);   // 클릭 막대를 처음 모양으로
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


/* ===== Inquiry.gs ===== */
/**
 * 💬 개발자에게 문의하기
 *
 * 사본을 쓰는 담당 선생님이 개발자에게 글을 남기고, 개발자의 답을 사본 안에서 받아 보는 기능입니다.
 * - 메뉴 「⚙️ 사본 설정 → 💬 개발자에게 문의하기 / 💬 내 문의 보기」
 * - 글은 개발자의 「문의함」 시트에 쌓입니다. 학교 이름과 앱 버전이 함께 갑니다. (학생 자료는 보내지 않습니다.)
 * - 보낼 때 받은 접수번호(무작위 번호)를 사본에 보관해 두었다가, 그 번호로만 자기 문의의 답을 읽어 옵니다.
 */
var HR_INQUIRY_URL = 'https://script.google.com/macros/s/AKfycbwGkSnEsgRFvV1DznOsTJveZ3tNWr8YPDZJ0wms_WQ3bj9_1sCi_y7NOQKXqOZ4Vigd/exec';
var HR_INQUIRY_TICKETS_PROPERTY = 'HR_INQUIRY_TICKETS';
var HR_INQUIRY_NAME_PROPERTY = 'HR_INQUIRY_NAME';
var HR_INQUIRY_KINDS = ['사용 질문', '오류 신고', '기능 요청', '기타'];

function HR_showInquiryDialog() { HR_openInquiryDialog_('send'); }
function HR_showMyInquiriesDialog() { HR_openInquiryDialog_('mine'); }

function HR_openInquiryDialog_(tab) {
  var props = PropertiesService.getScriptProperties();
  var template = HtmlService.createTemplateFromFile('Inquiry');
  template.tab = tab;
  template.school = String(props.getProperty('HR_SCHOOL_NAME') || '');
  template.savedName = String(props.getProperty(HR_INQUIRY_NAME_PROPERTY) || '');
  template.kinds = JSON.stringify(HR_INQUIRY_KINDS);
  SpreadsheetApp.getUi().showModalDialog(
    template.evaluate().setWidth(520).setHeight(640),
    '💬 개발자에게 문의'
  );
}

function HR_readInquiryTickets_() {
  try {
    var list = JSON.parse(PropertiesService.getScriptProperties().getProperty(HR_INQUIRY_TICKETS_PROPERTY) || '[]');
    return Array.isArray(list) ? list : [];
  } catch (error) {
    return [];
  }
}

// 문의 보내기 (창의 「보내기」 버튼)
function HR_sendInquiry(form) {
  var props = PropertiesService.getScriptProperties();
  var school = String(props.getProperty('HR_SCHOOL_NAME') || '').trim();
  if (!school) {
    return { ok: false, message: '학교 이름이 아직 없습니다. 「🏫 학교 최초 설정」을 먼저 해 주세요.' };
  }
  var kind = String(form && form.kind || '');
  var name = String(form && form.name || '').trim().slice(0, 40);
  var body = String(form && form.body || '').trim().slice(0, 2000);
  var email = String(form && form.email || '').trim().slice(0, 100);
  if (HR_INQUIRY_KINDS.indexOf(kind) < 0) kind = '기타';
  if (body.length < 5) return { ok: false, message: '문의 내용을 5글자 이상 적어 주세요.' };

  var response;
  var result;
  try {
    response = UrlFetchApp.fetch(HR_INQUIRY_URL, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({
        school: school, name: name, kind: kind, body: body, email: email,
        ver: typeof HR_BUILD_ID !== 'undefined' ? HR_BUILD_ID : '', website: String(form && form.website || '')
      }),
      muteHttpExceptions: true
    });
  } catch (error) {
    // 대표적인 원인: 사본에 「외부 서비스 연결」 권한이 아직 없을 때
    return { ok: false, message: '문의를 보내지 못했습니다. (' + String(error && error.message || error).slice(0, 160) + ')' };
  }
  try {
    result = JSON.parse(response.getContentText());
  } catch (error) {
    return {
      ok: false,
      message: '받는 곳의 응답을 읽지 못했습니다. (응답 코드 ' + response.getResponseCode() + ') 잠시 뒤 다시 시도해 주세요.'
    };
  }
  if (!result || !result.ok) {
    return { ok: false, message: (result && result.message) || '접수하지 못했습니다. 잠시 뒤 다시 시도해 주세요.' };
  }

  if (name) props.setProperty(HR_INQUIRY_NAME_PROPERTY, name);
  var tickets = HR_readInquiryTickets_();
  tickets.push({
    id: result.ticket,
    at: Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm'),
    kind: kind,
    preview: body.replace(/\s+/g, ' ').slice(0, 60)
  });
  props.setProperty(HR_INQUIRY_TICKETS_PROPERTY, JSON.stringify(tickets.slice(-30)));
  return { ok: true };
}

// 내 문의와 답변 (창의 「내 문의」 탭)
function HR_getMyInquiries() {
  var tickets = HR_readInquiryTickets_();
  if (!tickets.length) return { ok: true, items: [] };
  var answers = {};
  try {
    var response = UrlFetchApp.fetch(
      HR_INQUIRY_URL + '?action=answers&tickets=' + encodeURIComponent(tickets.map(function(t) { return t.id; }).join(',')),
      { muteHttpExceptions: true }
    );
    var data = JSON.parse(response.getContentText());
    if (data && data.ok) answers = data.items || {};
  } catch (error) {
    return { ok: false, message: '답변을 불러오지 못했습니다. (' + String(error && error.message || error).slice(0, 160) + ')' };
  }
  // 답이 온 문의는 직접 확인했으니, 시트를 열 때 다시 알려 주지 않게 표시합니다.
  var changed = false;
  tickets.forEach(function(t) {
    var a = answers[t.id];
    if (a && a.status === '답변완료' && !t.done) { t.done = true; changed = true; }
  });
  if (changed) PropertiesService.getScriptProperties().setProperty(HR_INQUIRY_TICKETS_PROPERTY, JSON.stringify(tickets));
  return {
    ok: true,
    items: tickets.slice().reverse().map(function(t) {
      var a = answers[t.id] || {};
      return {
        at: t.at, kind: t.kind, preview: t.preview,
        status: a.status || '대기', answer: a.answer || '', answeredAt: a.answeredAt || ''
      };
    })
  };
}

// 시트를 열 때: 아직 답을 못 본 문의가 있으면 가볍게 확인해서, 답이 도착했으면 알림을 한 번 띄웁니다.
// - 보낸 문의가 없거나 모두 확인했으면 인터넷 요청을 하지 않습니다.
// - 같은 답은 한 번만 알립니다.
function HR_notifyInquiryAnswersOnOpen_() {
  var tickets = HR_readInquiryTickets_();
  var waiting = tickets.filter(function(t) { return !t.done; });
  if (!waiting.length) return;

  var response = UrlFetchApp.fetch(
    HR_INQUIRY_URL + '?action=answers&tickets=' + encodeURIComponent(waiting.map(function(t) { return t.id; }).join(',')),
    { muteHttpExceptions: true }
  );
  var data = JSON.parse(response.getContentText());
  if (!data || !data.ok) return;

  var arrived = 0;
  tickets.forEach(function(t) {
    var a = (data.items || {})[t.id];
    if (!t.done && a && a.status === '답변완료') { t.done = true; arrived++; }
  });
  if (!arrived) return;

  PropertiesService.getScriptProperties().setProperty(HR_INQUIRY_TICKETS_PROPERTY, JSON.stringify(tickets));
  SpreadsheetApp.getActiveSpreadsheet().toast(
    '개발자가 문의 ' + arrived + '건에 답변을 남겼습니다. 메뉴 「⚙️ 사본 설정 → 📬 내 문의·답변 보기」에서 확인하십시오.',
    '💬 답변이 도착했습니다',
    15
  );
}


/* ===== Manual.gs ===== */
/**
 * 📘 전체 사용방법 (새 버전)
 *
 * 신청 받기부터 서점 전달까지, 새로 생긴 메뉴(교직원/학부모 신청, 링크·QR, 신청 한도, 새 신청 시작,
 * 업데이트 소식, 데이터 가져오기)까지 한 장에 정리합니다.
 * 안내 시트는 원래 교체해도 되는 시트라서, 새 버전이 나오면 처음 열 때 한 번 자동으로 새 내용으로 바꿉니다.
 */
var HR_GUIDE_VERSION = '2026-11-28-1';
var HR_GUIDE_VERSION_PROPERTY = 'HR_GUIDE_VERSION';

var HR_MANUAL_TITLE = '📘 학교도서관 희망도서·구입 업무 전체 사용방법';
var HR_MANUAL_SUBTITLE =
  '신청 받기부터 서점에 주문서를 보낼 때까지 한 장으로 정리했습니다. 처음이라면 맨 위부터 차례로 읽어 보십시오.';
var HR_MANUAL_FOOTER =
  '업무 흐름: 신청 받기 → 구입 검토함 → 소장 중복 확인 → 서점 전달용  ·  막히면 맨 아래 「자주 묻는 질문」을 먼저 보십시오.';

var HR_MANUAL_SECTIONS = [
  {
    title: '🧭 한눈에 보기 — 업무는 4단계입니다',
    items: [
      ['① 신청 받기',
        '교직원과 학부모가 웹 페이지(링크·QR)로 희망도서를 신청합니다. 신청은 「📚 희망도서 목록」에 자동으로 쌓이고, 표지·저자·출판사·정가·ISBN 같은 책 정보는 저절로 채워집니다.'],
      ['② 후보 모으기',
        '「📚 희망도서 목록」의 신청을 「🛒 구입 검토함」으로 보냅니다. 시리즈 누락·후속권 점검 결과(🔎)와 엑셀 업로드(📥) 후보도 같은 검토함에 모입니다.'],
      ['③ 구입 결정',
        '「🛒 구입 검토함」에서 살 책의 「선택」을 체크하고 수량·메모를 확인합니다. 우리 학교에 이미 있는 책은 자동으로 걸러집니다. 만화는 비고에 「📙 만화」, 문제집·학습지로 보이는 책은 「⚠ 문제집·학습지 의심」(줄이 연한 빨강)으로 표시되고, 서점 전달용을 만들기 전에 한 번 더 물어봅니다. 표시가 틀리면 비고에서 지우면 됩니다.'],
      ['④ 서점에 전달',
        '검토함의 금색 버튼을 누르면 「📦 서점 전달용」이 만들어집니다. 개인정보 없이 같은 책은 수량으로 합치고 금액도 계산해 줍니다. 파일 → 다운로드 → Microsoft Excel(.xlsx)로 저장해 서점에 보내십시오.']
    ]
  },
  {
    title: '🔗 신청 링크 나눠 주기 — 교직원용 / 학부모용',
    items: [
      ['링크 받기',
        '「⚙️ 사본 설정 → 🔗 신청 링크 받기」를 누르십시오. 교직원용·학부모용 링크가 나오고, 복사하거나 QR 그림으로 저장해 가정통신문에 넣을 수 있습니다. 문자·알림장에 붙일 안내 문구도 한 번에 복사됩니다.'],
      ['처음 한 번: 웹 앱 배포',
        '링크가 만들어지려면 Apps Script에서 웹 앱을 한 번 배포해야 합니다. (확장 프로그램 → Apps Script → 배포 → 새 배포 → 웹 앱, 실행 사용자는 「나」, 액세스는 「모든 사용자」) 「링크 받기」 창에 순서가 크게 나옵니다.'],
      ['교직원 / 학부모',
        '두 링크로 들어가면 각자 자기 신청 목록만 보입니다. 서로의 신청은 보이지도, 취소할 수도 없습니다. 한 시트의 「신청구분」 칸에서 사서 선생님이 함께 관리합니다.'],
      ['학부모 신청 정보',
        '학년·반·번호는 꼭 골라야 하고, 학생 이름은 선택입니다. 입력한 정보는 신청 확인에만 쓰이고 신청 목록 화면에는 보이지 않습니다.'],
      ['교직원 인증번호',
        '「⚙️ 사본 설정 → 🔑 교직원 인증번호 설정」에서 켜고 끌 수 있습니다. 켜 두면 웹 페이지에서 「교직원」으로 들어갈 때 인증번호를 입력해야 합니다. 학부모 화면에는 인증번호가 없습니다. 번호를 10번 틀리면 5분 동안 입력이 막힙니다. 끄면 번호 없이 들어갈 수 있고, 다시 켜면 저장해 둔 번호를 그대로 씁니다.'],
      ['신청 한도',
        '「⚙️ 사본 설정 → 📏 신청 한도 설정」에서 1인당 신청할 수 있는 권수를 정합니다. 학부모는 한 학생(학년·반·번호)당 5권이 기본이고, 교직원은 제한이 없습니다. 0을 넣으면 제한 없음입니다. 한도를 넘으면 안내 창이 뜹니다.'],
      ['링크가 열리지 않습니다',
        '「공개 신청 링크 등록」에 넣은 주소가 틀리면 「신청 링크 받기」 창에 원인과 해결 방법이 나옵니다. 「배포 → 배포 관리」에서 웹 앱 주소를 처음부터 끝까지 다시 복사해 붙여 넣으십시오.']
    ]
  },
  {
    title: '📋 메뉴 지도 — 위쪽 메뉴를 언제 쓰십니까?',
    items: [
      ['📚 도서 관리',
        '「희망도서 목록」을 다루는 메뉴입니다.\n' +
        '• 🔍 기존 데이터 일괄 채우기 — 비어 있는 책 정보를 한꺼번에 채워요\n' +
        '• 🔄 선택한 행 다시 조회 — 고른 줄의 책 정보를 다시 가져와요\n' +
        '• 🛒 구입 검토함에 접수 도서 추가 · 📦 서점 전달용 시트 새로고침\n' +
        '• 🎨 희망도서 목록 디자인 다시 적용 · ⚙️ 실시간 자동화 권한 승인\n' +
        '• 🔐 관리자 · 새 희망도서 신청 시작 — 아래 「새 학기·새 신청 회차」 참고'],
      ['⭐ 희망도서 신청',
        '• 🔗 신청 페이지 열기 — 등록된 신청 링크로 바로 이동해요\n• 🧹 새 학기 초기화 — 아래 「새 학기·새 신청 회차」 참고'],
      ['⚙️ 사본 설정',
        '처음 설정과 관리 메뉴입니다.\n' +
        '• ✅ 설정 한 번에 점검 · 🔑 API 키 설정 (YES24 필수) · 🏫 학교 최초 설정 · 🔗 공개 신청 링크 등록\n' +
        '• 🔗 신청 링크 받기 — 교직원/학부모 링크와 QR 그림\n' +
        '• 📏 신청 한도 설정 · 🔑 교직원 인증번호 설정 · 📢 업데이트 소식 보기\n' +
        '• 💬 개발자에게 문의하기 · 📬 내 문의·답변 보기\n' +
        '• 📥 이전 사본에서 데이터 가져오기 — 옛 사본의 신청 데이터를 이 사본으로 옮겨요\n' +
        '• 🧰 설정 탭 만들기·새로고침 · 🗑️ 설정 완료 후 탭 삭제'],
      ['🧰 통합 도서업무',
        '구입 업무 전체를 이끄는 메뉴입니다.\n' +
        '• 🏠 전체 화면 정리·새로고침 · 🎨 전체 디자인 다시 적용 · 🧭 시트 탭 정리\n' +
        '• 📚 희망도서를 구입 검토함으로 보내기 · 🛒 체크한 시리즈를 구입 검토함으로 보내기\n' +
        '• 🔎 시리즈 누락 점검(쉬운 창) · 🔎 시리즈 누락·후속권 이어서 점검 · 📥 업로드 목록을 열기 · 💵 예산 설정(대시보드의 노란 「총 예산」 칸에 직접 써도 됩니다) · ✅ 이번 회차 구입 확정 · 🧾 구입 기록 보기/숨기기(기록 시트는 평소 숨겨져 있고, 대시보드의 「지금까지 확정」 칸을 눌러도 열립니다)\n' +
        '• 📦 중복 확인 후 서점 전달용으로 보내기\n' +
        '• 🔎 시리즈 점검 도구(하위 메뉴) — 현재 결과 절판 여부 확인 · 결과 시트 만들기·복구 · 저장된 점검 위치 초기화 · YES24 API 키 설정']
    ]
  },
  {
    title: '🗓 새 학기·새 신청 회차',
    items: [
      ['🔐 새 희망도서 신청 시작',
        '지금 신청 목록을 숨김 시트(🗂 희망도서 날짜)로 보관한 뒤 목록을 비우고 새 신청을 받습니다. 구입 검토함·서점 전달용·소장도서·설정은 그대로입니다. 학생별 신청 한도도 새 목록 기준으로 다시 0권부터 셉니다.'],
      ['🧹 새 학기 초기화',
        '이번 학기 접수·검토 데이터를 비우고 서점 전달용 시트를 숨깁니다. 소장도서와 백업 자료는 남깁니다. 한 학기가 끝났을 때만 쓰십시오.'],
      ['누르기 전에',
        '둘 다 목록을 비우는 메뉴입니다. 누르기 전에 서점 전달용을 엑셀(.xlsx)로 내려받아 두십시오.']
    ]
  },
  {
    title: '📢 소식과 공지',
    items: [
      ['업데이트 소식',
        '새 소식이 있으면 시트를 열 때마다 안내 창이 뜹니다. 「이 소식 다시 보지 않기」를 누르면 그 소식은 더 뜨지 않습니다. 다시 보려면 「⚙️ 사본 설정 → 📢 업데이트 소식 보기」를 누르십시오.'],
      ['웹 공지 (신청 화면 위쪽)',
        '「📢 웹 공지」 탭에서 「표시」를 체크하면 신청 페이지 위쪽에 공지가 뜹니다. 「대상」 칸에서 전체·교직원·학부모를 고를 수 있습니다. 시작일·종료일은 비워도 됩니다.']
    ]
  },
  {
    title: '🛠 자주 묻는 질문',
    items: [
      ['책이 신청이 되지 않습니다',
        '이미 학교에 있는 책(소장도서)이거나 이미 신청된 책입니다. 「🏫 우리학교 소장도서」에 도서명과 ISBN13을 최신으로 넣어 두면 중복을 정확히 알려 줍니다.'],
      ['한도를 넘었다고 나옵니다',
        '1인당(학부모는 한 학생당) 신청 한도 때문입니다. 신청을 취소하면 그 권수는 다시 신청할 수 있고, 한도는 「📏 신청 한도 설정」에서 바꿉니다.'],
      ['웹 화면이 옛날 모양입니다',
        '열어 둔 웹 페이지 탭을 새로고침(F5)하십시오. 화면 맨 아래의 「화면 버전」으로 최신인지 확인할 수 있고, 새 버전이 나오면 위쪽에 안내가 뜹니다.'],
      ['메뉴가 안 보입니다',
        '처음 열면 「⚙️ 시작하기」만 보입니다. 「프로그램 사용 시작」을 눌러 권한을 한 번 허용하고 새로고침하십시오.'],
      ['탭이 너무 많아서 찾기 어렵습니다',
        '「🧰 통합 도서업무 → 🧭 시트 탭 정리」를 누르면 업무 순서로 놓고 색으로 묶습니다. (초록: 업무 · 금색: 자료 · 회색: 안내·설정)']
    ]
  }
];

// 글자 수와 줄바꿈으로 칸 높이를 어림합니다. (합친 칸은 자동으로 높이가 안 맞춰져서)
function HR_manualRowHeight_(text, charsPerLine) {
  var lines = 0;
  String(text).split('\n').forEach(function(part) {
    lines += Math.max(1, Math.ceil(part.length / charsPerLine));
  });
  return Math.max(36, lines * 21 + 16);
}

// 구역(section) 형태의 안내 시트를 만듭니다.
function HR_buildManualSheet_(name, title, subtitle, sections, footer) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(name) || ss.insertSheet(name);
  var maxColumns = 12;
  var rowsNeeded = 8;
  sections.forEach(function(section) { rowsNeeded += section.items.length + 2; });

  if (sheet.getFilter()) sheet.getFilter().remove();
  sheet.getBandings().forEach(function(b) { b.remove(); });
  sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).breakApart();
  sheet.clear();
  sheet.clearConditionalFormatRules();
  if (sheet.getMaxColumns() < maxColumns) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), maxColumns - sheet.getMaxColumns());
  }
  if (sheet.getMaxRows() < rowsNeeded) {
    sheet.insertRowsAfter(sheet.getMaxRows(), rowsNeeded - sheet.getMaxRows());
  }

  sheet.setHiddenGridlines(true);
  sheet.setFrozenRows(0);
  sheet.setFrozenColumns(0);
  sheet.setTabColor('#9AA6A0');
  sheet.setColumnWidths(1, 3, 62);   // 제목 칸(A~C)
  sheet.setColumnWidths(4, 9, 74);   // 설명 칸(D~L)

  sheet.getRange('A1:L2').merge().setValue(title)
    .setBackground('#183D35').setFontColor('#FFFFFF').setFontFamily('Arial').setFontSize(24)
    .setFontWeight('bold').setHorizontalAlignment('left').setVerticalAlignment('middle');
  sheet.setRowHeight(1, 34);
  sheet.setRowHeight(2, 34);

  sheet.getRange('A3:L3').merge().setValue(subtitle)
    .setBackground('#EAF3ED').setFontColor('#34564D').setFontFamily('Arial').setFontSize(13)
    .setWrap(true).setVerticalAlignment('middle');
  sheet.setRowHeight(3, 46);
  sheet.setRowHeight(4, 14);

  var row = 5;
  sections.forEach(function(section) {
    sheet.getRange(row, 1, 1, maxColumns).merge().setValue(section.title)
      .setBackground('#2F5D50').setFontColor('#FFFFFF').setFontFamily('Arial').setFontSize(16)
      .setFontWeight('bold').setHorizontalAlignment('left').setVerticalAlignment('middle');
    sheet.setRowHeight(row, 40);
    row++;

    section.items.forEach(function(item) {
      sheet.getRange(row, 1, 1, 3).merge().setValue(item[0])
        .setBackground('#FFF4D6').setFontColor('#684C18').setFontFamily('Arial').setFontSize(13)
        .setFontWeight('bold').setHorizontalAlignment('left').setVerticalAlignment('middle').setWrap(true);
      sheet.getRange(row, 4, 1, 9).merge().setValue(item[1])
        .setBackground('#FBFCFB').setFontColor('#2E3A33').setFontFamily('Arial').setFontSize(12)
        .setHorizontalAlignment('left').setVerticalAlignment('middle').setWrap(true);
      sheet.getRange(row, 1, 1, maxColumns).setBorder(
        true, true, true, true, false, false, '#E0D5B8', SpreadsheetApp.BorderStyle.SOLID
      );
      sheet.setRowHeight(row, HR_manualRowHeight_(item[1], 50));
      row++;
    });

    sheet.setRowHeight(row, 16);   // 구역 사이 빈 줄
    row++;
  });

  sheet.getRange(row, 1, 1, maxColumns).merge().setValue(footer)
    .setBackground('#DCEBE5').setFontColor('#244D42').setFontFamily('Arial').setFontSize(13)
    .setFontWeight('bold').setHorizontalAlignment('left').setVerticalAlignment('middle').setWrap(true);
  sheet.setRowHeight(row, 46);

  return sheet;
}

// 안내 시트가 새 버전이 아니면, 처음 열 때 한 번 새 내용으로 바꿉니다.
function HR_refreshGuidesOnce_() {
  var props = PropertiesService.getScriptProperties();
  if (props.getProperty(HR_GUIDE_VERSION_PROPERTY) === HR_GUIDE_VERSION) return;
  var lock = LockService.getScriptLock();
  var locked = false;
  try { locked = lock.tryLock(3000); } catch (error) {}
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var original = spreadsheet.getActiveSheet();   // 안내 시트를 그린 뒤 원래 보던 시트로 돌아오기 위해
  try {
    if (props.getProperty(HR_GUIDE_VERSION_PROPERTY) === HR_GUIDE_VERSION) return;
    createUnifiedGuides_();
    try {
      HR_styleWebNoticeSheet_(spreadsheet.getSheetByName(HR_WEB_NOTICE_SHEET_NAME));   // 웹 공지 시트도 새 모양으로 한 번 다시 꾸밉니다.
    } catch (error) {
      console.warn('웹 공지 시트 꾸미기 실패: ' + error);
    }
    try {
      HR_writeUpdateLogSheet_();   // 「업데이트 내용」 탭을 최신 내용으로 다시 씁니다.
    } catch (error) {
      console.warn('업데이트 내용 탭 쓰기 실패: ' + error);
    }
    try {
      unifiedApplyActionBarsAll_();   // 체크박스였던 맨 위 버튼을 클릭 막대로 한 번 바꿉니다.
    } catch (error) {
      console.warn('클릭 막대 적용 실패: ' + error);
    }
    try {
      unifiedRefreshDashboard_();   // 대시보드도 새 디자인으로 한 번 다시 그립니다.
    } catch (error) {
      console.warn('대시보드 갱신 실패: ' + error);
    }
    props.setProperty(HR_GUIDE_VERSION_PROPERTY, HR_GUIDE_VERSION);
    try {
      if (original && !original.isSheetHidden()) spreadsheet.setActiveSheet(original);
    } catch (error) {}
  } finally {
    if (locked) lock.releaseLock();
  }
}


/* ===== PurchaseLog.gs ===== */
/**
 * 🧾 구입 기록 (회차별)
 *
 * 책은 한 번만 사는 것이 아니므로, 서점에 주문한 뒤 「✅ 이번 회차 구입 확정」을 누르면
 * 그 회차의 날짜·권수·금액을 「🧾 구입 기록」 시트에 한 줄로 남깁니다.
 * 대시보드의 남은 예산 = 총 예산 − 지금까지 확정한 모든 회차 금액 − 지금 고른(아직 확정 전) 금액
 * - 확정한 책은 구입 검토함에서 「주문완료」로 바뀌고 체크가 풀려서, 다음 회차와 섞이지 않습니다.
 * - 기록 시트에서 금액·메모를 고치거나 줄을 지우면 예산이 그대로 다시 계산됩니다.
 * - 기록 시트는 평소에는 숨겨져 있습니다. 대시보드의 「✅ 지금까지 확정」 칸을 누르거나 메뉴 「🧾 구입 기록 보기」로 엽니다.
 */
var UNIFIED_PURCHASE_LOG_SHEET = '🧾 구입 기록';
var PURCHASE_LOG_FIRST_ROW = 5;

// 기록 시트가 없으면 「📦 서점 전달용」 바로 오른쪽에 만듭니다.
function unifiedEnsurePurchaseLogSheet_() {
  var ss = seriesSpreadsheet_();
  var sheet = ss.getSheetByName(UNIFIED_PURCHASE_LOG_SHEET);
  if (sheet) return sheet;

  var anchor = ss.getSheetByName(UNIFIED_EXPORT_SHEET);
  sheet = anchor ? ss.insertSheet(UNIFIED_PURCHASE_LOG_SHEET, anchor.getIndex()) : ss.insertSheet(UNIFIED_PURCHASE_LOG_SHEET);
  if (sheet.getMaxColumns() < 5) sheet.insertColumnsAfter(sheet.getMaxColumns(), 5 - sheet.getMaxColumns());
  sheet.setHiddenGridlines(true);
  sheet.setTabColor('#B9770E');

  sheet.getRange('A1:E1').merge().setValue('🧾 구입 기록')
    .setBackground('#183D35').setFontColor('#FFFFFF').setFontSize(22).setFontWeight('bold')
    .setHorizontalAlignment('left').setVerticalAlignment('middle');
  sheet.getRange('A2:E2').merge().setValue('서점 전달용 시트의 「✅ 이번 회차 구입 확정」을 누르면 한 회차가 한 줄씩 쌓입니다. 대시보드의 남은 예산이 이 합계로 계산됩니다.')
    .setBackground('#EAF3ED').setFontColor('#34564D').setFontSize(11).setWrap(true).setVerticalAlignment('middle');
  sheet.getRange('A3:E3').merge().setValue('금액·메모는 직접 고칠 수 있고, 잘못 확정했다면 그 줄을 지우면 예산이 되돌아갑니다.')
    .setBackground('#FFF4D6').setFontColor('#765413').setFontSize(10).setFontWeight('bold').setVerticalAlignment('middle');
  sheet.getRange('A4:E4').setValues([['회차', '확정일', '권수', '금액(원)', '메모']])
    .setBackground('#2E5B4F').setFontColor('#FFFFFF').setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  [70, 120, 80, 140, 420].forEach(function(w, i) { sheet.setColumnWidth(i + 1, w); });
  sheet.setRowHeight(1, 54); sheet.setRowHeight(2, 36); sheet.setRowHeight(3, 28); sheet.setRowHeight(4, 34);
  sheet.setFrozenRows(4);
  sheet.getRange(PURCHASE_LOG_FIRST_ROW, 1, 500, 5).setVerticalAlignment('middle').setFontSize(11);
  sheet.getRange(PURCHASE_LOG_FIRST_ROW, 1, 500, 3).setHorizontalAlignment('center');
  sheet.getRange(PURCHASE_LOG_FIRST_ROW, 4, 500, 1).setNumberFormat('#,##0"원"').setHorizontalAlignment('right');
  sheet.getRange(PURCHASE_LOG_FIRST_ROW, 2, 500, 1).setNumberFormat('yyyy-mm-dd');
  sheet.getRange('A4:E504').setBorder(true, true, true, true, true, true, '#D3DDD7', SpreadsheetApp.BorderStyle.SOLID);
  try { sheet.hideSheet(); } catch (error) {}   // 평소에는 숨겨 두고, 필요할 때만 엽니다.
  return sheet;
}

// 지금까지 확정한 회차 수와 금액·권수 합계
function unifiedPurchaseLogStats_(ss) {
  var result = { rounds: 0, amount: 0, quantity: 0 };
  var sheet = ss.getSheetByName(UNIFIED_PURCHASE_LOG_SHEET);
  if (!sheet || sheet.getLastRow() < PURCHASE_LOG_FIRST_ROW) return result;
  sheet.getRange(PURCHASE_LOG_FIRST_ROW, 1, sheet.getLastRow() - PURCHASE_LOG_FIRST_ROW + 1, 4).getValues().forEach(function(row) {
    var amount = Number(row[3] || 0);
    if (!amount && !Number(row[2] || 0)) return;   // 빈 줄
    result.rounds++;
    result.amount += amount;
    result.quantity += Number(row[2] || 0);
  });
  return result;
}

// 「✅ 이번 회차 구입 확정」 (서점 전달용 시트의 막대·메뉴)
function unifiedConfirmPurchaseRound() { unifiedConfirmPurchaseRound_(); }

function unifiedConfirmPurchaseRound_() {
  var ss = seriesSpreadsheet_();
  var ui = SpreadsheetApp.getUi();
  var exportSheet = ss.getSheetByName(UNIFIED_EXPORT_SHEET);
  if (!exportSheet || exportSheet.getLastRow() < 6) {
    ss.toast('확정할 서점 전달용 목록이 없습니다. 먼저 구입 검토함에서 서점 전달용을 만들어 주세요.', '구입 확정', 8);
    return { success: false };
  }

  // 서점 전달용에 담긴 책(= 이번 회차에 주문한 책)
  var books = [];
  var quantity = 0;
  var amount = 0;
  var keys = {};
  exportSheet.getRange(6, 1, exportSheet.getLastRow() - 5, 10).getValues().forEach(function(row) {
    if (!String(row[1] || '').trim()) return;
    var count = Math.max(1, Number(row[7] || 1));
    books.push(row);
    quantity += count;
    amount += Number(row[8] || 0) || Number(row[6] || 0) * count;
    var key = unifiedBookKey_(row[5], row[1]);
    if (key) keys[key] = true;
  });
  if (!books.length) {
    ss.toast('서점 전달용에 책이 없습니다.', '구입 확정', 6);
    return { success: false };
  }

  var log = unifiedEnsurePurchaseLogSheet_();
  var stats = unifiedPurchaseLogStats_(ss);
  var round = stats.rounds + 1;
  var skipDialog = typeof UNIFIED_BAR_NO_DIALOG !== 'undefined' && UNIFIED_BAR_NO_DIALOG;   // 막대 클릭: 첫 번째 누름에서 확인을 마쳤습니다
  var answer = skipDialog ? ui.Button.YES : ui.alert(
    '✅ ' + round + '회차 구입 확정',
    '서점 전달용의 ' + quantity + '권 · ' + unifiedMoney_(amount) + '을 ' + round + '회차로 기록할까요?\\n\\n' +
      '· 구입 검토함에서 이 책들은 「주문완료」로 바뀌고 체크가 풀립니다.\\n' +
      '· 기록은 「' + UNIFIED_PURCHASE_LOG_SHEET + '」 시트에서 고치거나 지울 수 있습니다.',
    ui.ButtonSet.YES_NO
  );
  if (answer !== ui.Button.YES) return { success: false, cancelled: true };

  // 막대 클릭(칸 선택 이벤트)에서는 잠금을 못 쓸 수 있어서, 안 되면 잠금 없이 진행합니다.
  var lock = null;
  try { lock = LockService.getDocumentLock(); lock.waitLock(20000); } catch (lockError) { lock = null; }
  try {
    var row = Math.max(log.getLastRow() + 1, PURCHASE_LOG_FIRST_ROW);
    log.getRange(row, 1, 1, 5).setValues([[
      round, new Date(), quantity, amount, ''
    ]]);
    log.getRange(row, 2).setNumberFormat('yyyy-mm-dd');
    log.getRange(row, 4).setNumberFormat('#,##0"원"');

    // 확정한 책은 구입 검토함에서 「주문완료」로 바꾸고 체크를 풉니다.
    var review = ss.getSheetByName(UNIFIED_REVIEW_SHEET);
    var marked = 0;
    if (review && review.getLastRow() >= 6) {
      var count = review.getLastRow() - 5;
      var rows = review.getRange(6, 1, count, 12).getValues();
      var checks = [], statuses = [];
      rows.forEach(function(r) {
        var key = unifiedBookKey_(r[7], r[3]);
        var hit = key && keys[key] && String(r[2]) !== '주문완료' && String(r[2]) !== '소장중';
        if (hit) marked++;
        checks.push([hit ? false : r[0]]);
        statuses.push([hit ? '주문완료' : r[2]]);
      });
      review.getRange(6, 1, count, 1).setValues(checks);
      review.getRange(6, 3, count, 1).setValues(statuses);
    }
    SpreadsheetApp.flush();
    try { unifiedUpdateDashboardNumbers_(); } catch (error) { console.warn('대시보드 갱신 실패: ' + error); }
    ss.toast(
      round + '회차 ' + quantity + '권 · ' + unifiedMoney_(amount) + '을 기록했습니다. (구입 검토함 ' + marked + '권을 주문완료로 바꿨습니다)',
      '✅ 구입 확정', 9
    );
    return { success: true, round: round, quantity: quantity, amount: amount, marked: marked };
  } finally {
    if (lock) { try { lock.releaseLock(); } catch (releaseError) {} }
  }
}

// 숨겨 둔 「🧾 구입 기록」 시트를 열어서 보여 줍니다.
function unifiedShowPurchaseLog() {
  var ss = seriesSpreadsheet_();
  var sheet = unifiedEnsurePurchaseLogSheet_();
  sheet.showSheet();
  ss.setActiveSheet(sheet);
  ss.toast('다 보셨으면 메뉴 「🧰 통합 도서업무 → 🧾 구입 기록 숨기기」로 다시 숨길 수 있습니다.', '🧾 구입 기록', 7);
}

function unifiedHidePurchaseLog() {
  var ss = seriesSpreadsheet_();
  var sheet = ss.getSheetByName(UNIFIED_PURCHASE_LOG_SHEET);
  if (!sheet) return;
  var dashboard = ss.getSheetByName(UNIFIED_DASHBOARD_SHEET);
  if (dashboard) ss.setActiveSheet(dashboard);
  sheet.hideSheet();
  ss.toast('구입 기록을 숨겼습니다.', '🧾 구입 기록', 4);
}

// 대시보드의 「✅ 지금까지 확정」 칸(D16:F16)을 누르면 기록 시트를 엽니다. 처리했으면 true
function unifiedHandleDashboardClick_(sheet, range) {
  if (sheet.getName() !== UNIFIED_DASHBOARD_SHEET) return false;
  if (range.getRow() !== 16 || range.getNumRows() !== 1 || range.getColumn() < 4 || range.getLastColumn() > 6) return false;
  try {
    unifiedShowPurchaseLog();
  } catch (error) {
    SpreadsheetApp.getActiveSpreadsheet().toast(String(error && error.message || error), '구입 기록을 열지 못했습니다', 7);
  }
  return true;
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
  // 주소가 열리는지 확인하는 요청(?probe=1)에는 화면을 만들지 않고 짧게 답합니다.
  if (e && e.parameter && e.parameter.probe === '1') {
    return ContentService.createTextOutput('ok');
  }
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
  template.buildId = HR_BUILD_ID;
  template.limitStaff = HR_getRequestLimit_('staff');
  template.limitParent = HR_getRequestLimit_('parent');
  template.staffCodeRequired = HR_isStaffCodeEnabled_();
  template.maxClasses = HR_getMaxClasses_();
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

// 이 코드가 만들어진 때(빌드가 채워 줍니다). 화면에 보여 주고, 열어 둔 화면이 옛 버전인지 알아보는 데 씁니다.
var HR_BUILD_ID = 'v0d3dfab';

function HR_getBuildId() {
  return HR_BUILD_ID;
}

// 웹 앱 주소가 실제로 열리는지 확인합니다.
// 'ok' 열림 · 'notfound' 없는 주소(404) · 'restricted' 로그인이 필요함(공개 아님) · 'unknown' 확인 못 함
function HR_probeWebAppUrl_(url) {
  try {
    var response = UrlFetchApp.fetch(String(url) + '?probe=1', {
      muteHttpExceptions: true,
      followRedirects: false
    });
    var code = response.getResponseCode();
    if (code === 404) return 'notfound';
    if (code === 401 || code === 403) return 'restricted';
    if (code >= 300 && code < 400) {
      var headers = response.getHeaders() || {};
      var location = String(headers.Location || headers.location || '');
      return /accounts\.google\.com/.test(location) ? 'restricted' : 'ok';
    }
    return code >= 200 && code < 300 ? 'ok' : 'unknown';
  } catch (error) {
    return 'unknown';
  }
}

// 확인 결과를 사용자에게 보여 줄 말로 바꿉니다. (열리면 빈 글자)
function HR_probeProblemMessage_(status) {
  if (status === 'notfound') {
    return '이 주소로는 접속이 안 돼요(페이지를 찾을 수 없음). 주소가 중간에 잘렸거나 지워진 배포일 수 있어요. ' +
      'Apps Script 「배포 → 배포 관리」에서 현재 웹 앱 주소를 처음부터 끝까지 다시 복사해 주세요.';
  }
  if (status === 'restricted') {
    return '이 주소는 로그인이 필요해요. Apps Script 「배포 → 배포 관리」에서 웹 앱의 액세스 권한을 「모든 사용자」로 바꿔 주세요.';
  }
  return '';
}

// 웹 앱 주소(…/exec) 모양인지 확인합니다.
function HR_isWebAppExecUrl_(url) {
  return /^https:\/\/script\.google\.com\/(?:a\/macros\/[^/]+|macros)\/s\/[^/]+\/exec$/.test(String(url || ''));
}

// 「공개 신청 링크 등록」에 저장한 주소(검사를 통과한 것만). 없으면 빈 글자입니다.
function HR_getRegisteredWebAppUrl_() {
  var saved = String(
    PropertiesService.getScriptProperties().getProperty(HR_PUBLIC_WEB_APP_URL_PROPERTY) || ''
  ).trim().replace(/[?#].*$/, '');
  return HR_isWebAppExecUrl_(saved) ? saved : '';
}

// 구글이 알려 주는 배포 주소. 못 구하면 빈 글자입니다.
function HR_getDetectedWebAppUrl_() {
  try {
    var url = String(ScriptApp.getService().getUrl() || '');
    return HR_isWebAppExecUrl_(url) ? url : '';
  } catch (error) {
    return '';
  }
}

// 이 웹 앱의 주소. 「공개 신청 링크 등록」에 넣은 주소가 항상 먼저이고, 없을 때만 구글이 알려 주는 주소를 씁니다.
function HR_getWebAppBaseUrl_() {
  return HR_getRegisteredWebAppUrl_() || HR_getDetectedWebAppUrl_();
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
  HR_requireStaffAccess_(requestAudience, formData && formData.staffCode);
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
      var requesterInfo = HR_recordExtraRequester_(spreadsheet, sheet, requestedBook.row, selectedBook, applicant);
      var requesterTotal = requesterInfo && requesterInfo.total ? requesterInfo.total : 1;
      var duplicateMessage = requesterInfo && requesterInfo.sameApplicant
        ? '이미 신청하신 도서입니다.'
        : '이미 신청된 도서입니다. 함께 원하는 분으로 기록했어요. (현재 ' + requesterTotal + '명)';
      return {
        success: false,
        duplicateRequest: true,
        requesters: requesterTotal,
        message: duplicateMessage,
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
  HR_requireStaffAccess_(requestAudience, formData && formData.staffCode);
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
    throw new Error('학년·반·번호를 모두 골라 주십시오. (예: 3학년 2반 15번)');
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

  // 학부모는 신청이 끝난 지금 기준의 남은 권수를 함께 돌려줘서 화면이 바로 바뀌게 합니다.
  var quota = null;
  if (requestAudience === 'parent') {
    try {
      var parsedApplicant = HR_parseParentApplicant_(applicant);
      quota = HR_getParentQuota(parsedApplicant.grade, parsedApplicant.classNo, parsedApplicant.number);
    } catch (quotaError) {
      quota = null;
    }
  }

  return {
    success: successCount > 0,
    successCount: successCount,
    duplicateCount: duplicateCount,
    ownedCount: ownedCount,
    errorCount: errorCount,
    results: results,
    quota: quota
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
  return HR_annotateBooksStatus_(items.map(yes24RequestBook_));
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
        '사본의 「⚙️ 사본 최초설정」 탭에서 API 키 설정을 누릅니다. YES24은 필수입니다.'
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
      'API 안내  ·  YES24은 도서 검색에 꼭 필요합니다.'
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
      '사본의 「⚙️ 사본 최초설정」 탭에서 API 키 설정을 누릅니다. YES24은 필수입니다.'
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
      'API 안내  ·  YES24 API Key는 도서 검색에 꼭 필요합니다.'
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
    .addItem('🔑 교직원 인증번호 설정', 'HR_showStaffCodeDialog')
    .addItem('💬 개발자에게 문의하기', 'HR_showInquiryDialog')
    .addItem('📬 내 문의·답변 보기', 'HR_showMyInquiriesDialog')
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
    try {
      var result = HR_saveYes24ApiKey({
        yes24Key: yes24Key
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

  // 실제로 열리는 주소인지 확인합니다. 안 열리는 주소는 저장하지 않고 이유를 알려 줍니다.
  var problem = HR_probeProblemMessage_(HR_probeWebAppUrl_(url));
  if (problem) {
    throw new Error(problem);
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

  // 시트 맨 위의 클릭 막대(전체 선택·보내기·점검 시작 등)와 대시보드의 「지금까지 확정」 칸
  try {
    if (unifiedHandleDashboardClick_(sheet, e.range)) {
      return;
    }
    if (unifiedHandleBarClick_(sheet, e.range)) {
      return;
    }
  } catch (error) {
    console.warn('클릭 막대 처리 실패: ' + error);
    try { SpreadsheetApp.getActiveSpreadsheet().toast(String(error && error.message || error).slice(0, 150), '버튼 처리 오류', 8); } catch (e2) {}
  }

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
    maxClasses: String(properties.getProperty(HR_MAX_CLASSES_PROPERTY) || '').trim(),
    yes24Key: '',
    yes24Ready: Boolean(yes24Key),
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
  var properties = PropertiesService.getScriptProperties();

  if (!key) {
    throw new Error('YES24 API Key를 입력해 주세요.');
  }

  HR_testYes24Key_(key);

  properties.setProperty(HR_YES24_KEY_PROPERTY, key);

  return {
    success: true,
    message: 'YES24 연결을 확인했습니다.'
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
  var maxClassesText = String(
    formData && formData.maxClasses || ''
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

  if (maxClassesText && !(/^\d{1,2}$/.test(maxClassesText) && Number(maxClassesText) >= 1 && Number(maxClassesText) <= 30)) {
    throw new Error('한 학년 최대 반 수는 1~30 사이의 숫자로 입력해 주세요. (비워 두면 20반까지 보입니다.)');
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
    HR_MAX_CLASSES: maxClassesText,
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
        '화면 위쪽 「⚙️ 사본 설정」 → 「🔑 API 키 설정 (YES24 필수)」을 누르세요. YES24 API Key 칸에 붙여 넣고 저장하세요.'
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
      '🔑 API 키 설정  ·  YES24 필수'
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
      'YES24은 도서 검색에 꼭 필요합니다.\n' +
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
      '  API 키 설정  ·  YES24 필수'
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
      'YES24은 도서 검색에 꼭 필요합니다.'
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
        '다시 이 시트로 돌아와 위쪽 「⚙️ 사본 설정」 → 「🔑 API 키 설정 (YES24 필수)」을 누르세요. YES24 API Key 칸에 붙여 넣고 저장하세요.'
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
      ' 클릭: API 키 설정  ·  YES24 필수'
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
  sheet.getRange('J10:L13').merge().setValue((publicReady ? '✅ ' : '') + '🔗 공개 신청 링크 등록');

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

  // 시리즈 누락 점검이 출판사로 시리즈를 구분하므로 「출판사」 칸도 둡니다.
  if (!HR_findAnyHeaderColumn_(headers, ['출판사', '발행자', '출판자', '발행처'])) {
    lastColumn = Math.max(sheet.getLastColumn(), 0) + 1;
    sheet.getRange(1, lastColumn).setValue('출판사');
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
function HR_getSubmittedRequests(audience, staffCode, applicant) {
  var audienceKey = HR_audienceKey_(audience);
  HR_requireStaffAccess_(audienceKey, staffCode);
  // 신청한 본인의 신청만 돌려줍니다. (학부모: 학년·반·번호, 교직원: 이름)
  var myKey = HR_applicantKey_(audienceKey, applicant);
  if (!myKey) {
    return { totalCount: 0, shownCount: 0, items: [], needApplicant: true, updatedAt: new Date().toISOString() };
  }
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
  var applicantColumn = HR_findHeaderColumn_(headers, '신청자', 11);

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

    if (HR_applicantKey_(audienceKey, row[applicantColumn - 1]) !== myKey) {
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
      audienceKey,
      myKey
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
  audienceKey,
  myKey
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
    var pendingApplicantColumn = HR_findHeaderColumn_(headers, '신청자', 11);
    if (HR_applicantKey_(audienceKey, sourceRow[pendingApplicantColumn - 1]) !== myKey) {
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
  HR_requireStaffAccess_(HR_audienceKey_(requestData && requestData.audience), requestData && requestData.staffCode);
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

    // 신청한 본인만 취소할 수 있습니다.
    var cancelApplicantColumn = HR_findHeaderColumn_(headers, '신청자', 11);
    var cancelMyKey = HR_applicantKey_(cancelRowAudience, requestData && requestData.applicant);
    var cancelRowKey = HR_applicantKey_(
      cancelRowAudience,
      sheet.getRange(rowNumber, cancelApplicantColumn).getDisplayValue()
    );
    if (!cancelMyKey || cancelMyKey !== cancelRowKey) {
      throw new Error('신청한 본인만 취소할 수 있습니다.');
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
  HR_requireStaffAccess_(HR_audienceKey_(requestData && requestData.audience), requestData && requestData.staffCode);
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
    // 신청한 본인만 되돌릴 수 있습니다.
    var undoApplicantColumn = HR_findHeaderColumn_(headers, '신청자', 11);
    var undoMyKey = HR_applicantKey_(undoRowAudience, requestData && requestData.applicant);
    var undoRowKey = (originalColumnCount >= undoApplicantColumn)
      ? HR_applicantKey_(
        undoRowAudience,
        backupSheet.getRange(backupRow, 8 + undoApplicantColumn).getDisplayValue()
      )
      : '';
    if (!undoMyKey || undoMyKey !== undoRowKey) {
      throw new Error('신청한 본인만 되돌릴 수 있습니다.');
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


/* ===== Requesters.gs ===== */
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


/* ===== ReservationShortcut.js ===== */
/** 시트를 열 때마다: 공개 신청 링크가 등록돼 있으면 바로가기 창을 띄웁니다(링크가 없으면 조용히 건너뜀). */
function HR_showShortcutOnOpen_() {
  var url = String(
    PropertiesService.getScriptProperties().getProperty(HR_PUBLIC_WEB_APP_URL_PROPERTY) || ''
  ).trim();
  if (!url) return;
  HR_showHopeRequestDialog();
}

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


/* ===== SeriesStop.gs ===== */
/**
 * ⏹ 시리즈 점검 멈추기 — 돌고 있는 점검을 멈춥니다.
 *   · UnifiedLibrary.js: 메뉴 항목 / seriesSearchYes24_ 반복문 / refreshMissingSeriesBooks 점검 직후에서 이 파일의 함수를 부릅니다.
 *
 * 돌고 있는 점검은 한 묶음이 끝나는 대로(몇 초 안에) 멈추고, 1분 뒤 자동 이어가기도 지웁니다.
 * 멈춘 점검의 이번 회 결과는 시트에 쓰지 않고, 저장된 점검 위치도 지웁니다. (이전에 쓰인 결과 시트는 그대로 남습니다)
 */
var SERIES_STOP_PROPERTY = 'SERIES_STOP_REQUESTED_AT';

function HR_seriesStopCheck() {
  var properties = PropertiesService.getScriptProperties();
  properties.setProperty(SERIES_STOP_PROPERTY, String(Date.now()));
  ['YES24_SERIES_QUOTA_NEXT', 'YES24_SERIES_QUOTA_RETRY', 'YES24_SERIES_ERROR_NEXT', 'YES24_SERIES_ERROR_RETRY'].forEach(function(key) { properties.deleteProperty(key); });
  try { seriesCancelAutoContinue_(); } catch (error) { console.warn('자동 이어가기 지우기 실패: ' + error); }
  try { seriesClearResumeState_(); } catch (error) { console.warn('점검 위치 지우기 실패: ' + error); }
  try { seriesWriteStatus_('stop', '⏹ 멈추기를 눌렀습니다 · ' + seriesNowText_() + ' (다음 점검은 처음부터 시작합니다)'); } catch (error) {}
  try {
    seriesSpreadsheet_().toast(
      '점검을 멈추도록 요청했습니다. 돌고 있는 점검은 몇 초 안에 멈추고, 1분 뒤 자동 이어가기는 취소했습니다. 다음 점검은 처음부터 시작합니다.',
      '⏹ 시리즈 점검 멈추기',
      10
    );
  } catch (error) {}
}

// 이 점검(runStartedAt 에 시작)이 시작된 뒤에 멈추기를 눌렀는지. 멈추기 전에 이미 끝난 요청은 영향 없습니다.
function seriesStopRequested_(runStartedAt) {
  var requestedAt = Number(PropertiesService.getScriptProperties().getProperty(SERIES_STOP_PROPERTY) || 0);
  return requestedAt > Number(runStartedAt || 0);
}


/* ===== StaffCode.gs ===== */
/**
 * 🔑 교직원 인증번호
 *
 * 켜 두면 웹 신청 페이지에서 「교직원」으로 들어갈 때 인증번호를 입력해야 합니다.
 * - 학부모 화면에는 인증번호가 없습니다.
 * - 번호는 이 시트의 스크립트 속성에만 저장되고, 웹 페이지 코드에는 들어가지 않습니다.
 * - 신청·취소·되돌리기·목록 보기는 서버에서도 번호를 다시 확인합니다. (화면만 가린 것이 아닙니다.)
 * - 번호를 여러 번 틀리면 잠시 막습니다.
 */
var HR_STAFF_CODE_ENABLED_PROPERTY = 'HR_STAFF_CODE_ENABLED';
var HR_STAFF_CODE_PROPERTY = 'HR_STAFF_CODE';
var HR_STAFF_CODE_FAIL_KEY = 'HR_STAFF_CODE_FAILS';
var HR_STAFF_CODE_MAX_FAILS = 10;      // 이만큼 틀리면
var HR_STAFF_CODE_LOCK_SECONDS = 300;  // 5분 동안 입력을 막습니다.

function HR_getStaffCode_() {
  return String(PropertiesService.getScriptProperties().getProperty(HR_STAFF_CODE_PROPERTY) || '').trim();
}

// 켜져 있고 번호도 정해져 있을 때만 true
function HR_isStaffCodeEnabled_() {
  var props = PropertiesService.getScriptProperties();
  return props.getProperty(HR_STAFF_CODE_ENABLED_PROPERTY) === '1' && !!HR_getStaffCode_();
}

function HR_staffCodeFailCount_() {
  return Number(CacheService.getScriptCache().get(HR_STAFF_CODE_FAIL_KEY) || 0);
}

// 번호를 확인합니다. 맞으면 true, 틀리면 false, 너무 많이 틀렸으면 오류를 냅니다.
function HR_checkStaffCode_(code) {
  if (!HR_isStaffCodeEnabled_()) return true;
  var cache = CacheService.getScriptCache();
  if (HR_staffCodeFailCount_() >= HR_STAFF_CODE_MAX_FAILS) {
    throw new Error('인증번호를 여러 번 틀려서 잠시 막혔습니다. 5분 뒤에 다시 시도해 주십시오.');
  }
  if (String(code || '').trim() === HR_getStaffCode_()) return true;
  cache.put(HR_STAFF_CODE_FAIL_KEY, String(HR_staffCodeFailCount_() + 1), HR_STAFF_CODE_LOCK_SECONDS);
  return false;
}

// 교직원 신청·취소·목록에서 부릅니다. 교직원이 아니거나 인증번호가 꺼져 있으면 그냥 지나갑니다.
function HR_requireStaffAccess_(audienceKey, code) {
  if (audienceKey !== 'staff') return;
  if (!HR_checkStaffCode_(code)) {
    throw new Error('교직원 인증번호가 맞지 않습니다. 처음 화면에서 다시 들어와 인증번호를 입력해 주십시오.');
  }
}

// 웹 화면에서 부릅니다: 입력한 번호가 맞는지 알려 줍니다.
function HR_verifyStaffCode(code) {
  if (!HR_isStaffCodeEnabled_()) return { ok: true };
  var ok = HR_checkStaffCode_(code);
  return { ok: ok };
}

// 메뉴: 「🔑 교직원 인증번호 설정」
function HR_showStaffCodeDialog() {
  var ui = SpreadsheetApp.getUi();
  var enabled = HR_isStaffCodeEnabled_();
  var answer = ui.alert(
    '🔑 교직원 인증번호 설정',
    '웹 신청 페이지에서 「교직원」으로 들어갈 때 인증번호를 입력하게 합니다. (학부모는 해당 없음)\n\n' +
    '현재 상태: ' + (enabled ? '🟢 켜짐 (번호: ' + HR_getStaffCode_() + ')' : '⚪ 꺼짐') + '\n\n' +
    '• 「예」 = ' + (enabled ? '번호 바꾸기' : '켜기(번호 정하기)') + '\n' +
    '• 「아니요」 = ' + (enabled ? '끄기' : '그대로 두기') + '\n' +
    '• 「취소」 = 닫기',
    ui.ButtonSet.YES_NO_CANCEL
  );
  var props = PropertiesService.getScriptProperties();

  if (answer === ui.Button.NO) {
    if (enabled) {
      props.setProperty(HR_STAFF_CODE_ENABLED_PROPERTY, '0');
      CacheService.getScriptCache().remove(HR_STAFF_CODE_FAIL_KEY);
      ui.alert('인증번호를 껐습니다', '이제 교직원도 번호 없이 들어갈 수 있습니다.\n(저장해 둔 번호는 지워지지 않아서, 다시 켜면 같은 번호를 쓸 수 있습니다.)', ui.ButtonSet.OK);
    }
    return;
  }
  if (answer !== ui.Button.YES) return;

  var prompt = ui.prompt(
    '인증번호 정하기',
    '교직원에게 알려 줄 번호를 입력해 주십시오.\n(4~12자, 영문·숫자만 쓸 수 있습니다. 예: 2468)',
    ui.ButtonSet.OK_CANCEL
  );
  if (prompt.getSelectedButton() !== ui.Button.OK) return;
  var code = String(prompt.getResponseText() || '').trim();
  if (!/^[0-9A-Za-z]{4,12}$/.test(code)) {
    ui.alert('번호를 확인해 주십시오', '4~12자의 영문·숫자만 쓸 수 있습니다. 다시 시도해 주십시오.', ui.ButtonSet.OK);
    return;
  }
  props.setProperty(HR_STAFF_CODE_PROPERTY, code);
  props.setProperty(HR_STAFF_CODE_ENABLED_PROPERTY, '1');
  CacheService.getScriptCache().remove(HR_STAFF_CODE_FAIL_KEY);
  ui.alert(
    '✅ 인증번호를 켰습니다',
    '교직원 인증번호: ' + code + '\n\n이 번호를 선생님들께 알려 주십시오. 열려 있던 웹 페이지는 새로고침하면 인증 창이 나옵니다.',
    ui.ButtonSet.OK
  );
}


/* ===== TabArrange.gs ===== */
/**
 * 🧭 시트 탭 정리
 *
 * 탭이 많아서 찾기 어려운 문제를 줄이기 위해 업무 순서대로 놓고 색으로 묶습니다.
 * - 탭 이름과 내용은 바꾸지 않습니다. (코드와 기존 데이터가 이름으로 연결되어 있기 때문입니다.)
 * - 초록: 매일 쓰는 업무 / 금색: 자료·기초 / 회색: 안내·설정
 * - 목록에 없는 시트(선생님이 직접 만든 시트 등)는 맨 뒤에 그대로 둡니다.
 * - 숨겨진 시트는 건드리지 않습니다.
 * - 사용방법이 맨 앞입니다. 처음 설정이 필요한 사본은 「사본 최초설정」이 바로 그 옆에 있습니다.
 * - 새 업데이트 후 처음 열 때 한 번만 자동으로 정리하고, 그 뒤에는 메뉴를 눌렀을 때만 정리합니다.
 */
var HR_TAB_PLAN = [
  ['📘 전체 사용방법', '#9AA6A0'],
  ['⚙️ 사본 최초설정', '#9AA6A0'],      // 설정이 끝나기 전에만 있는 탭이라, 있으면 사용방법 바로 옆에 둡니다.
  ['📗 희망도서 신청안내', '#9AA6A0'],
  ['🏠 업무 대시보드', '#2F7258'],
  ['📚 희망도서 목록', '#2F7258'],
  ['🛒 구입 검토함', '#2F7258'],
  ['📦 서점 전달용', '#2F7258'],
  ['🔎 시리즈 누락 점검', '#D9A441'],
  ['📋 시리즈 업무목록', '#D9A441'],
  ['📥 엑셀 업로드 목록', '#D9A441'],
  ['🏫 우리학교 소장도서', '#D9A441'],
  ['📢 웹 공지', '#9AA6A0'],
  ['업데이트 내용', '#9AA6A0'],
  ['⚙️ 통합 설정', '#9AA6A0']
];
var HR_TAB_HIDE = ['🔒 신청접수백업', '🔒 신청취소복구'];
var HR_TABS_ARRANGED_PROPERTY = 'HR_TABS_ARRANGED';
// 탭 순서 규칙을 바꾸면 이 번호를 올립니다. 번호가 다르면 처음 열 때 한 번 새 순서로 정리합니다.
var HR_TABS_ARRANGE_VERSION = '2';

// 탭을 업무 순서로 놓고 색을 칠합니다. 옮긴 시트 수를 돌려줍니다.
function HR_arrangeTabsNow_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var original = ss.getActiveSheet();
  var position = 1;
  var moved = 0;

  HR_TAB_PLAN.forEach(function(item) {
    var sheet = ss.getSheetByName(item[0]);
    if (!sheet || sheet.isSheetHidden()) return;
    sheet.setTabColor(item[1]);
    ss.setActiveSheet(sheet);
    ss.moveActiveSheet(position);
    position++;
    moved++;
  });

  // 백업용 탭은 숨깁니다. (보이는 시트가 하나뿐이면 숨기지 않습니다.)
  HR_TAB_HIDE.forEach(function(name) {
    var sheet = ss.getSheetByName(name);
    if (!sheet || sheet.isSheetHidden()) return;
    var visibleCount = ss.getSheets().filter(function(s) { return !s.isSheetHidden(); }).length;
    if (visibleCount > 1) sheet.hideSheet();
  });

  // 원래 보던 시트로 돌아갑니다.
  try {
    if (original && !original.isSheetHidden()) ss.setActiveSheet(original);
  } catch (error) {
    console.warn('원래 시트로 돌아가지 못했습니다: ' + error);
  }
  return moved;
}

// 메뉴: 「🧭 시트 탭 정리」
function HR_arrangeTabs() {
  var moved = HR_arrangeTabsNow_();
  PropertiesService.getScriptProperties().setProperty(HR_TABS_ARRANGED_PROPERTY, HR_TABS_ARRANGE_VERSION);
  SpreadsheetApp.getActiveSpreadsheet().toast(
    '시트 탭 ' + moved + '개를 업무 순서로 정리하고 색으로 묶었어요. (초록: 업무 · 금색: 자료 · 회색: 안내·설정)',
    '🧭 탭 정리 완료',
    8
  );
}

// 시트를 열 때: 아직 한 번도 정리하지 않았다면 한 번만 정리합니다.
function HR_arrangeTabsOnce_() {
  var props = PropertiesService.getScriptProperties();
  if (props.getProperty(HR_TABS_ARRANGED_PROPERTY) === HR_TABS_ARRANGE_VERSION) return;
  var lock = LockService.getScriptLock();
  var locked = false;
  try { locked = lock.tryLock(3000); } catch (error) {}
  try {
    if (props.getProperty(HR_TABS_ARRANGED_PROPERTY) === HR_TABS_ARRANGE_VERSION) return;
    HR_arrangeTabsNow_();
    props.setProperty(HR_TABS_ARRANGED_PROPERTY, HR_TABS_ARRANGE_VERSION);
  } finally {
    if (locked) lock.releaseLock();
  }
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
var SERIES_ENRICH_BUDGET_MS = 150000;      // 준비 단계(출판사·판본 확인)에 쓰는 시간. 다 못 하면 다음 알람에서 이어서 합니다.
// 한 번에 점검하는 시간. 이 뒤에 결과 시트를 쓰는 시간(수십 초)이 더 필요하고 전체 한도는 6분이라, 여유를 넉넉히 둡니다.
var SERIES_EXECUTION_BUDGET_MS = 210000;
var SERIES_FOLLOWUP_RESULTS = 50;
var SERIES_GAP_RESULTS = 50;
var SERIES_MAX_GAP_PAGES = 3;
var SERIES_ALGORITHM_VERSION = 'v3-prepare-first';
var SERIES_RESUME_STATE_KEY = 'YES24_SERIES_RESUME_STATE_V1';
// 소장도서 탭에 적어 둔 「출판사」 칸은 묶음 나누기에 쓰지 않습니다. (일부 줄만 적거나 표기가 서로 달라 같은 시리즈가 갈라져 있는 책이 누락으로 나오는 일이 있었습니다)
// 출판사는 YES24 에서 ISBN 으로 자동으로 확인한 값(seriesEnrichGroupsFromIsbn_)만 씁니다.
var SERIES_USE_REGISTERED_PUBLISHER = false;
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
    '시리즈 시트 위쪽의 「🔎 시리즈 누락 점검 시작 ›」 막대 맨 앞 ☐ 칸을 체크해 주세요.',
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

  unifiedDrawActionBars_(sheet);   // 체크박스 대신 클릭 막대
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

    PropertiesService.getScriptProperties().setProperty(SERIES_RUN_AT_PROPERTY, String(Date.now()));
    seriesWriteStatus_('run', '🟢 점검 중입니다 · ' + seriesNowText_() + ' 시작 (끝나면 이 줄이 바뀝니다)');
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
    var enrichment = seriesEnrichGroupsFromIsbn_(groups, key, runStartedAt);
    // 준비 단계를 다 못 끝냈으면 검색은 다음 알람으로 미루고, 지금까지 확인한 만큼은 기억해 둡니다.
    if (enrichment.pending > 0) {
      if (typeof seriesStopRequested_ === 'function' && seriesStopRequested_(runStartedAt)) return;
      seriesSaveResumeState_({
        fingerprint: fingerprint,
        offset: resumeOffset,
        total: Math.max(Number(resumeState && resumeState.total) || 0, groups.length),
        updatedAt: new Date().toISOString()
      });
      seriesScheduleAutoContinue_();
      seriesWriteStatus_('wait', '⏳ 준비 중 · 시리즈 출판사·판본 확인 ' + seriesProgressText_(enrichment.total - enrichment.pending, enrichment.total).replace(/건\)$/, '개)') + ' · 약 1분 뒤 저절로 이어갑니다');
      spreadsheet.toast('시리즈 ' + enrichment.pending + '개의 출판사·판본을 더 확인해야 합니다. 약 1분 뒤 저절로 이어갑니다.', '시리즈 점검 준비 중', 8);
      return;
    }

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
    // 시리즈 점검 멈추기: 멈추기를 눌렀으면 이번 결과는 쓰지 않고 이어가기도 하지 않습니다.
    if (typeof seriesStopRequested_ === 'function' && seriesStopRequested_(runStartedAt)) {
      seriesClearResumeState_();
      seriesCancelAutoContinue_();
      seriesWriteStatus_('stop', '⏹ 점검을 멈췄습니다 · ' + seriesNowText_() + ' (다음 점검은 처음부터 시작합니다)');
      spreadsheet.toast('점검을 멈췄습니다. 이번 점검의 결과는 쓰지 않았습니다. 다음 점검은 처음부터 시작합니다.', '⏹ 시리즈 점검 멈춤', 8);
      return;
    }
    var candidates = seriesMergeCandidates_(existingCandidates, searchResult.candidates);
    // 3000개를 넘는 시리즈는 이번 점검 대상이 아닙니다. (예전에는 이것 때문에 다 끝나도 「미완료」로 남아 1분마다 영원히 반복됐습니다.)
    var incomplete = searchResult.truncated;

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
    // 결과 시트를 다 쓴 뒤에야 진행 위치를 앞으로 옮깁니다. (시트를 쓰다 시간이 끝나도 이번 구간을 다시 하면 되고, 결과가 빠지지 않습니다)
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

    PropertiesService.getScriptProperties().deleteProperty(SERIES_ERROR_RETRY_PROPERTY);
    PropertiesService.getScriptProperties().deleteProperty(SERIES_ERROR_NEXT_PROPERTY);
    PropertiesService.getScriptProperties().deleteProperty(SERIES_QUOTA_RETRY_PROPERTY);   // 정상 진행되면 한도 재시도 횟수 초기화
    PropertiesService.getScriptProperties().deleteProperty(SERIES_QUOTA_NEXT_PROPERTY);
    if (incomplete) {
      seriesWriteStatus_('wait', '⏳ 대기 중 · ' + seriesProgressText_(searchResult.nextOffset, searchResult.totalSpecCount) + ' · ' + seriesNowText_() + ' 저장 · 약 1분 뒤 저절로 이어서 점검합니다');
    } else {
      seriesWriteStatus_('done', '✅ 전체 확인 완료 · ' + seriesProgressText_(searchResult.totalSpecCount, searchResult.totalSpecCount) + ' · ' + seriesNowText_());
    }

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
    if (searchResult.errorCount) {
      spreadsheet.toast(
        '일부 검색 ' + searchResult.errorCount + '건은 YES24 오류로 건너뛰었습니다. (예: ' + (searchResult.skippedQueries[0] || '') + ') 다음 점검 때 다시 시도됩니다.',
        '시리즈 점검 참고', 12
      );
    }
    if (groupLimitReached) {
      spreadsheet.toast(
        '소장도서의 시리즈가 ' + allGroups.length + '개로 한 번에 점검하는 ' + SERIES_MAX_GROUPS + '개를 넘어, 우선순위가 낮은 ' + (allGroups.length - groups.length) + '개는 점검하지 않았습니다.',
        '시리즈 점검 참고', 12
      );
    }
  } catch (error) {
    var errorText = String(error && error.message ? error.message : error);
    // 하루 조회 한도에 닿아 멈춘 경우: 오류로 끝내지 않고, 한도가 풀리는 시각에 저절로 이어서 점검하도록 예약합니다.
    if (seriesIsQuotaError_(errorText)) {
      try {
        var resumeInfo = seriesScheduleQuotaResume_(errorText);
        seriesWriteStatus_('wait', '⏳ YES24 하루 조회 한도로 쉬는 중 · ' + resumeInfo.whenText + ' 저절로 이어서 점검합니다 (진행 위치는 보존됨)');
        seriesSpreadsheet_().toast(
          'YES24 하루 조회 한도에 닿아 잠시 멈췄습니다. ' + resumeInfo.whenText + '에 저절로 이어서 점검합니다. (진행 위치는 보존되어 있습니다. 시트를 열어 두지 않아도 됩니다)',
          '시리즈 점검 일시 정지',
          15
        );
        return;
      } catch (scheduleError) {
        console.warn('한도 뒤 이어가기 예약 실패: ' + scheduleError);
      }
    }
    // 키 문제가 아닌 일시적인 오류(YES24 서버 불안정·요청 많음 등)는 멈추지 않고 5분 뒤 다시 시도합니다.
    if (!/API 키/.test(errorText)) {
      try {
        var retryInfo = seriesScheduleErrorRetry_(errorText);
        if (retryInfo) {
          seriesSpreadsheet_().toast('일시적인 오류로 잠시 멈췄습니다. ' + retryInfo.whenText + '에 다시 시도합니다.', '시리즈 점검 일시 정지', 12);
          return;
        }
      } catch (retryError) {
        console.warn('오류 뒤 다시 시도 예약 실패: ' + retryError);
      }
    }
    try { seriesCancelAutoContinue_(); } catch (ignore) {}   // 더 시도해도 소용없는 오류(키 문제 등)는 알람을 끕니다
    seriesWriteStatus_('error', '⚠ 오류로 멈췄습니다 · ' + seriesNowText_() + ' · ' + errorText.split('\n')[0].slice(0, 120));
    seriesSpreadsheet_().toast(errorText, '시리즈 점검 오류', 12);
    throw error;
  } finally {
    try { PropertiesService.getScriptProperties().deleteProperty(SERIES_RUN_AT_PROPERTY); } catch (ignore) {}
    lock.releaseLock();
  }
}

// 오늘 조회 한도에 닿아서 멈춘 오류인지 (키 오류·일반 서버 오류와 구분합니다)
function seriesIsQuotaError_(message) {
  return /오늘 안전하게 사용할 수 있는|HTTP 429|RATE_002|일일 호출 한도|한도.{0,8}(초과|도달)|(초과|도달).{0,8}한도/.test(String(message || ''));
}

var SERIES_QUOTA_RETRY_PROPERTY = 'YES24_SERIES_QUOTA_RETRY';
var SERIES_QUOTA_NEXT_PROPERTY = 'YES24_SERIES_QUOTA_NEXT';   // 한도 때문에 예약한 다음 점검 시각
var SERIES_QUOTA_MAX_RETRIES = 6;

// 한도가 풀릴 때 다시 시작할 시각. 우리 쪽 안전 한도(날짜가 바뀌면 초기화)에 닿았으면 다음 날 00:10,
// YES24 가 직접 한도를 알려 왔으면(아직 풀리지 않았을 수 있어) 1시간 뒤에 다시 시도합니다.
function seriesQuotaResumeTime_(now, yes24Rejected, timeZone) {
  if (yes24Rejected) return new Date(now.getTime() + 60 * 60 * 1000);
  var zone = timeZone || 'Asia/Seoul';
  var y = Number(Utilities.formatDate(now, zone, 'yyyy'));
  var m = Number(Utilities.formatDate(now, zone, 'MM'));
  var d = Number(Utilities.formatDate(now, zone, 'dd'));
  return new Date(y, m - 1, d + 1, 0, 10, 0);
}

// 한도 뒤 자동 이어가기를 예약합니다. 너무 여러 번 실패하면(6번) 예약을 멈춥니다.
function seriesScheduleQuotaResume_(message) {
  var properties = PropertiesService.getScriptProperties();
  var retries = Number(properties.getProperty(SERIES_QUOTA_RETRY_PROPERTY) || 0);
  if (retries >= SERIES_QUOTA_MAX_RETRIES) {
    properties.deleteProperty(SERIES_QUOTA_RETRY_PROPERTY);
    throw new Error('한도 때문에 여러 번 다시 시도했지만 계속 막혀 있습니다.');
  }
  properties.setProperty(SERIES_QUOTA_RETRY_PROPERTY, String(retries + 1));

  var yes24Rejected = !/오늘 안전하게 사용할 수 있는/.test(String(message || ''));
  var when = seriesQuotaResumeTime_(new Date(), yes24Rejected, Session.getScriptTimeZone());
  seriesScheduleAutoContinue_(10);   // 10분마다 확인하고, 쉬는 시간이 지났을 때 이어갑니다
  var zone = Session.getScriptTimeZone() || 'Asia/Seoul';
  var whenText = Utilities.formatDate(when, zone, 'M월 d일 HH:mm') + ' 무렵';
  properties.setProperty(SERIES_QUOTA_NEXT_PROPERTY, JSON.stringify({ at: when.getTime(), text: whenText }));
  return { when: when, whenText: whenText };
}

var SERIES_ERROR_RETRY_PROPERTY = 'YES24_SERIES_ERROR_RETRY';
var SERIES_ERROR_NEXT_PROPERTY = 'YES24_SERIES_ERROR_NEXT';
var SERIES_ERROR_RETRY_MAX = 6;
var SERIES_ERROR_RETRY_DELAY_MS = 5 * 60 * 1000;
// 일시적인 오류 뒤 5분 뒤에 다시 시도하도록 예약합니다. 연달아 6번 실패하면 예약하지 않고 null 을 돌려줍니다.
function seriesScheduleErrorRetry_(message) {
  var properties = PropertiesService.getScriptProperties();
  var retries = Number(properties.getProperty(SERIES_ERROR_RETRY_PROPERTY) || 0);
  if (retries >= SERIES_ERROR_RETRY_MAX) {
    properties.deleteProperty(SERIES_ERROR_RETRY_PROPERTY);
    return null;
  }
  properties.setProperty(SERIES_ERROR_RETRY_PROPERTY, String(retries + 1));
  seriesScheduleAutoContinue_(5);
  var when = new Date(Date.now() + SERIES_ERROR_RETRY_DELAY_MS);
  properties.setProperty(SERIES_ERROR_NEXT_PROPERTY, JSON.stringify({ at: when.getTime() }));
  var whenText = Utilities.formatDate(when, Session.getScriptTimeZone() || 'Asia/Seoul', 'HH:mm') + ' 무렵';
  seriesWriteStatus_('wait', '⏳ 일시적인 오류로 잠시 쉬는 중 · ' + whenText + ' 다시 시도합니다 (' + (retries + 1) + '/' + SERIES_ERROR_RETRY_MAX + '번째) · ' + String(message || '').split('\n')[0].slice(0, 80));
  return { whenText: whenText };
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
    var publisher = publisherColumn && SERIES_USE_REGISTERED_PUBLISHER ? String(row[publisherColumn - 1] || '').trim() : '';
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

  // 같은 시리즈·판형에서 출판사가 적힌 줄이 하나라도 있으면, 출판사가 비어 있는 줄은 그 출판사들의 묶음에 함께 넣습니다.
  // (출판사를 일부 줄에만 적었을 때 같은 시리즈가 갈라져서, 갈라진 쪽마다 없는 권을 누락으로 잡던 문제를 막습니다.)
  var knownPublishers = {};
  var prepared = [];
  rows.forEach(function(book) {
    var parsed = seriesParseTitle_(book.title);
    if (!parsed) return;
    var publisherKey = seriesNormalizePublisher_(book.publisher) || '';
    var formatKey = seriesFormatKey_(book.title);
    prepared.push({ book: book, parsed: parsed, publisherKey: publisherKey, formatKey: formatKey });
    if (publisherKey) {
      var seriesFormat = parsed.key + '|' + formatKey;
      if (!knownPublishers[seriesFormat]) knownPublishers[seriesFormat] = {};
      knownPublishers[seriesFormat][publisherKey] = true;
    }
  });

  function addToGroup(book, parsed, publisherKey, formatKey) {
    var groupPublisher = publisherKey || '출판사미상';
    var signature = parsed.key + '|' + groupPublisher + '|' + formatKey;

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
    if (publisherKey) group.publishers[publisherKey] = true;
  }

  prepared.forEach(function(entry) {
    var known = knownPublishers[entry.parsed.key + '|' + entry.formatKey];
    if (!entry.publisherKey && known) {
      Object.keys(known).forEach(function(publisherKey) {
        addToGroup(entry.book, entry.parsed, publisherKey, entry.formatKey);
      });
      return;
    }
    addToGroup(entry.book, entry.parsed, entry.publisherKey, entry.formatKey);
  });

  var finalGroups = Object.keys(groupsByKey)
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

  // 같은 시리즈·판형인데 출판사 표기가 달라(예: 「고릴라박스」와 「비룡소」) 따로 묶인 곳의 보유 권도 알아 둡니다.
  // 후보를 고를 때 「다른 묶음에 이미 있는 권」을 누락으로 잡지 않기 위해서입니다. (순환 참조를 피하려고 복사본만 담습니다)
  finalGroups.forEach(function(group) {
    group.siblingHeld = finalGroups
      .filter(function(other) {
        return other !== group && other.seriesKey === group.seriesKey && other.formatKey === group.formatKey;
      })
      .map(function(other) {
        return { publishers: other.publishers, volumes: other.volumes };
      });
  });
  return finalGroups;
}

function seriesEnrichGroupsFromIsbn_(groups, apiKey, runStartedAt) {
  
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

  // 준비 단계: 검색을 시작하기 전에 모든 시리즈의 판본·출판사를 먼저 확인합니다.
  // (예전에는 점검 한 번에 20개씩만 확인해서, 확인 전에 검색된 시리즈는 출판사 필터 없이 다른 출판사 책까지 후보로 걸렸습니다.)
  // 시간이 모자라면 확인한 만큼 저장해 두고 나머지는 다음 알람에서 이어서 합니다. 확인한 결과는 기억해 두어 다시 묻지 않습니다.
  var totalSpecs = specs.length;
  var allowed = Math.min(totalSpecs, seriesApiRemaining_());
  if (totalSpecs > 0 && allowed < 1) {
    throw new Error(
      '오늘 안전하게 사용할 수 있는 YES24 조회량이 부족합니다. 기존 결과는 그대로 보존했습니다. 내일 다시 점검해 주세요.'
    );
  }
  specs = specs.slice(0, allowed);

  var requestCount = 0;
  var processed = 0;
  for (var start = 0; start < specs.length; start += SERIES_SEARCH_CHUNK) {
    if (start > 0 && Date.now() - Number(runStartedAt || 0) >= SERIES_ENRICH_BUDGET_MS) break;
    if (typeof seriesStopRequested_ === 'function' && seriesStopRequested_(runStartedAt)) break;
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
    processed += chunk.length;
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

    // 판본 기억 저장 공간이 가득 차도 점검은 계속합니다. (기억은 속도를 위한 보조 기능입니다.)
    if (Object.keys(cacheWrites).length) {
      try { properties.setProperties(cacheWrites, false); } catch (ignore) { cacheWrites = {}; }
    }

    Utilities.sleep(100);
  }

  if (Object.keys(cacheWrites).length) {
    try { properties.setProperties(cacheWrites, false); } catch (ignore) {}
  }
  return {requestCount: requestCount, total: totalSpecs, done: processed, pending: Math.max(0, totalSpecs - processed)};
}

// 두 시리즈 이름(정규화한 글)이 사실상 같은 시리즈인지: 글자 두 개씩 겹치는 비율(Dice)이 0.6 이상이거나,
// 한쪽이 다른 쪽을 품고 길이 비율이 0.6 이상이면 같은 것으로 봅니다.
function seriesNameKeysSimilar_(left, right) {
  var a = String(left || ''), b = String(right || '');
  if (!a || !b) return false;
  if (a === b) return true;
  var shorter = a.length <= b.length ? a : b;
  var longer = a.length > b.length ? a : b;
  if (longer.indexOf(shorter) >= 0 && shorter.length / longer.length >= 0.6) return true;
  function bigrams(text) {
    var list = [];
    for (var i = 0; i < text.length - 1; i++) list.push(text.substr(i, 2));
    return list;
  }
  var x = bigrams(a), y = bigrams(b);
  if (!x.length || !y.length) return false;
  var pool = {};
  y.forEach(function(gram) { pool[gram] = (pool[gram] || 0) + 1; });
  var common = 0;
  x.forEach(function(gram) { if (pool[gram] > 0) { pool[gram]--; common++; } });
  return (2 * common) / (x.length + y.length) >= 0.6;
}

function seriesApplyEditionInfo_(group, info) {
  var apiTitle = seriesCleanText_(info && info.title);
  var parsed = seriesParseTitle_(apiTitle);
  // YES24가 그 ISBN을 학교 소장도서와 전혀 다른 이름으로 돌려주면(예: 소장 「처음 읽는 삼국지」 ↔ YES24 「삼국지 1」)
  // 그 정보를 믿지 않고 무시합니다. 믿으면 짧고 흔한 이름(「삼국지」)이 별칭이 되어 이름이 다른 책이 후보로 걸립니다.
  if (parsed && !seriesNameKeysSimilar_(parsed.key, group.seriesKey)) return;
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
  // 점검 방식을 크게 바꿀 때 올립니다. 올리면 예전 방식으로 하던 중간 진행·결과를 이어가지 않고 처음부터 다시 합니다.
  source = SERIES_ALGORITHM_VERSION + '\n' + source;

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

// 이어가기 예약이 울릴 때마다 불립니다. 예약은 「한 번만 울리는 알람」이 아니라 「계속 울리는 알람」이라서,
// 한 번 헛돌거나 오류가 나도 다음 알람에서 다시 이어집니다. 점검이 다 끝나면 알람을 지웁니다.
// 여기서 「할 일이 있는지」를 확인하고, 쉬어야 하는 때(한도 대기·오류 대기)나 이미 돌고 있을 때는 조용히 넘어갑니다.
function seriesAutoContinueMissingBooks_() {
  var properties = PropertiesService.getScriptProperties();
  var pending = properties.getProperty(SERIES_RESUME_STATE_KEY) ||
    properties.getProperty(SERIES_QUOTA_NEXT_PROPERTY) || properties.getProperty(SERIES_ERROR_NEXT_PROPERTY);
  if (!pending) { seriesCancelAutoContinue_(); return; }
  var now = Date.now();
  var runAt = Number(properties.getProperty(SERIES_RUN_AT_PROPERTY) || 0);
  if (runAt && now - runAt < 7 * 60 * 1000) return;   // 지금 돌고 있는 점검이 있습니다
  var waiting = [SERIES_QUOTA_NEXT_PROPERTY, SERIES_ERROR_NEXT_PROPERTY].some(function(key) {
    try {
      var info = JSON.parse(properties.getProperty(key) || 'null');
      return info && Number(info.at) > now;
    } catch (ignore) { return false; }
  });
  if (waiting) return;   // 한도·오류로 쉬는 시간이 아직 안 지났습니다
  refreshMissingSeriesBooks();
}

// 점검이 중간에 멈춰 진행 위치만 남았는데 이어가기 예약이 하나도 없으면(예: 예전 버전에서 한도에 닿아 멈춘 경우) 스스로 예약을 다시 만듭니다.
// 오늘 한도가 남아 있으면 1분 뒤, 다 썼으면 한도가 풀리는 시각(다음 날 00:10)입니다.
function seriesEnsureAutoContinue_() {
  var properties = PropertiesService.getScriptProperties();
  if (!properties.getProperty(SERIES_RESUME_STATE_KEY)) return false;
  var already = ScriptApp.getProjectTriggers().some(function(trigger) {
    return trigger.getHandlerFunction() === SERIES_AUTO_CONTINUE_HANDLER;
  });
  if (already) return false;
  if (seriesApiRemaining_() > 0) {
    seriesScheduleAutoContinue_();
    seriesWriteStatus_('wait', '⏳ 점검이 중간에 멈춰 있어 다시 예약했습니다 · 약 1분 뒤 저절로 이어서 점검합니다 (' + seriesNowText_() + ' 확인)');
    return true;
  }
  var when = seriesQuotaResumeTime_(new Date(), false, Session.getScriptTimeZone());
  seriesScheduleAutoContinue_(10);
  var zone = Session.getScriptTimeZone() || 'Asia/Seoul';
  properties.setProperty(SERIES_QUOTA_NEXT_PROPERTY, JSON.stringify({
    at: when.getTime(), text: Utilities.formatDate(when, zone, 'M월 d일 HH:mm') + ' 무렵'
  }));
  seriesWriteStatus_('wait', '⏳ 오늘 YES24 조회 한도를 다 써서 쉬는 중 · ' + Utilities.formatDate(when, zone, 'M월 d일 HH:mm') + ' 무렵 저절로 이어서 점검합니다');
  return true;
}

// 시리즈 점검 결과 시트 맨 위(2번째 줄)에 지금 상태를 크게 보여 줍니다. (점검 중 / 이어가기 대기 / 한도 대기 / 완료 / 멈춤 / 오류)
var SERIES_STATUS_STYLES = {
  run:   { bg: '#E4F1FB', fg: '#1F5A8A' },
  wait:  { bg: '#FFF4D9', fg: '#8A6518' },
  done:  { bg: '#EAF5EE', fg: '#2D6A50' },
  stop:  { bg: '#EEEEEE', fg: '#555555' },
  error: { bg: '#FDECEC', fg: '#A33B36' }
};
// 진행 막대: [██████░░░░░░░░░░░░░░] 30.0% (4000/13741건)
function seriesProgressText_(done, total) {
  total = Number(total) || 0;
  if (total <= 0) return '';
  done = Math.max(0, Math.min(total, Number(done) || 0));
  var percent = done / total * 100;
  var filled = Math.round(percent / 5);
  return '[' + new Array(filled + 1).join('█') + new Array(20 - filled + 1).join('░') + '] ' + percent.toFixed(1) + '% (' + done + '/' + total + '건)';
}
function seriesNowText_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'Asia/Seoul', 'M월 d일 HH:mm');
}
function seriesWriteStatus_(kind, text) {
  try {
    var sheet = seriesSpreadsheet_().getSheetByName(SERIES_FINDER_SHEET);
    if (!sheet) return;
    var style = SERIES_STATUS_STYLES[kind] || SERIES_STATUS_STYLES.stop;
    sheet.getRange('A2:N2').setBackground(style.bg).setFontColor(style.fg).setFontWeight('bold').setFontSize(14)
      .setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true);
    sheet.setRowHeight(2, 44);
    sheet.getRange('A2').setValue(text);
    SpreadsheetApp.flush();
  } catch (error) {
    console.warn('점검 상태 표시 실패: ' + error);
  }
}

// 시트를 열 때/창을 열 때, 저장된 점검 상태를 맨 위 상태 줄에 맞춰 보여 줍니다. (예약을 만들지는 않고 보여 주기만 합니다)
var SERIES_RUN_AT_PROPERTY = 'YES24_SERIES_RUN_AT';
function seriesShowStateStatus_() {
  var properties = PropertiesService.getScriptProperties();
  var resume = null;
  try { resume = JSON.parse(properties.getProperty(SERIES_RESUME_STATE_KEY) || 'null'); } catch (ignore) {}
  if (!resume || !(Number(resume.total) > 0)) return;
  var runAt = Number(properties.getProperty(SERIES_RUN_AT_PROPERTY) || 0);
  if (runAt && Date.now() - runAt < 7 * 60 * 1000) return;   // 지금 돌고 있는 점검이 쓴 「점검 중」 줄을 덮지 않습니다
  var progress = seriesProgressText_(Number(resume.offset || 0), Number(resume.total));
  var hasTrigger = false;
  try {
    hasTrigger = ScriptApp.getProjectTriggers().some(function(trigger) {
      return trigger.getHandlerFunction() === SERIES_AUTO_CONTINUE_HANDLER;
    });
  } catch (ignore) { return; }
  var next = null;
  try { next = JSON.parse(properties.getProperty(SERIES_QUOTA_NEXT_PROPERTY) || 'null'); } catch (ignore) {}
  if (hasTrigger && next && Number(next.at) > Date.now()) {
    seriesWriteStatus_('wait', '⏳ YES24 하루 조회 한도로 쉬는 중 · ' + next.text + ' 저절로 이어서 점검합니다 · ' + progress);
  } else if (hasTrigger) {
    seriesWriteStatus_('wait', '⏳ 대기 중 · ' + progress + ' · 곧 저절로 이어서 점검합니다');
  } else {
    seriesWriteStatus_('error', '⚠ 점검이 멈춰 있습니다 · ' + progress + ' · 위의 ☐ 를 체크해 점검 창을 열면 다시 예약됩니다');
  }
}

function seriesCancelAutoContinue_() {
  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    if (trigger.getHandlerFunction() === SERIES_AUTO_CONTINUE_HANDLER) {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}

// 1분마다 울리는 이어가기 알람을 만듭니다. (한 번만 울리는 알람은 헛돌거나 오류가 나면 거기서 끊겼습니다.)
function seriesScheduleAutoContinue_(everyMinutes) {
  seriesCancelAutoContinue_();
  ScriptApp.newTrigger(SERIES_AUTO_CONTINUE_HANDLER)
    .timeBased()
    .everyMinutes(everyMinutes || 1)
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
    if (typeof seriesStopRequested_ === 'function' && seriesStopRequested_(runStartedAt)) break;
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

    // 한 묶음(5건) 중 한 건이라도 YES24 가 일시적으로 거절하면 yes24FetchAll_ 이 오류를 던집니다.
    // 예전에는 그 오류 하나로 지금까지 몇 분간 한 검색이 전부 버려지고 점검이 그대로 멈췄습니다.
    // 일시적인 오류는 잠깐 쉬었다 다시 시도하고, 그래도 안 되면 지금까지 한 만큼만 저장하고 이어서 점검하게 합니다.
    var batch = null;
    var fetchFailure = null;
    for (var attempt = 1; attempt <= 3; attempt++) {
      try {
        batch = yes24FetchAll_(requests);
        break;
      } catch (fetchError) {
        fetchFailure = fetchError;
        var fetchMessage = String(fetchError && fetchError.message ? fetchError.message : fetchError);
        if (seriesIsQuotaError_(fetchMessage) || /API 키/.test(fetchMessage)) break;   // 한도·키 문제는 다시 해도 소용없습니다
        Utilities.sleep(3000 * attempt);
      }
    }
    if (!batch) {
      if (!responses.length) throw fetchFailure;   // 이번에 한 검색이 하나도 없으면 오류로 알립니다
      break;                                       // 한 만큼은 저장하고, 나머지는 이어서 점검합니다
    }
    seriesRecordApiUsage_(batch.length);
    batch.forEach(function(response, index) {
      responses.push({spec: chunk[index], response: response});
    });

    if ((start / SERIES_SEARCH_CHUNK) % 10 === 0) {
      try {
        seriesWriteStatus_('run', '🟢 점검 중 · ' + seriesProgressText_(Math.min(startOffset + start + chunk.length, totalSpecCount), totalSpecCount) + ' · ' + seriesNowText_() + ' 갱신');
      } catch (ignore) {}
    }
    seriesSpreadsheet_().toast(
      Math.min(startOffset + start + chunk.length, totalSpecCount) +
        ' / ' + totalSpecCount + '건 확인 중',
      'YES24 검색',
      4
    );
    Utilities.sleep(180);
  }

  var bySeriesVolume = {};
  var apiErrors = [];     // 키·한도 문제처럼 전체를 멈춰야 하는 오류
  var skippedQueries = [];  // 그 검색 한 건만 건너뛰어도 되는 오류

  responses.forEach(function(entry) {
    var response = entry.response;
    var spec = entry.spec;
    var status = response.getResponseCode();
    if (status !== 200) {
      // 인증·한도 오류(401·403·429)는 전부 멈추고, 그 밖의 오류(400·404·5xx)는 그 검색만 건너뜁니다.
      // (예전에는 한 건만 실패해도 점검 전체가 같은 자리에서 영원히 멈췄습니다.)
      (/^(401|403|429)$/.test(String(status)) ? apiErrors : skippedQueries).push(spec.query + ': HTTP ' + status);
      return;
    }

    var data;
    try {
      data = JSON.parse(response.getContentText('UTF-8'));
    } catch (ignore) {
      skippedQueries.push(spec.query + ': 응답 해석 실패');
      return;
    }
    if (data.errorCode) {
      apiErrors.push(spec.query + ': ' + data.errorCode + ' ' + (data.errorMessage || ''));
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

  // 절반 이상(3건 이상)이 실패했다면 키·서버 문제이므로 멈추고, 몇 건만 실패했으면 건너뛰고 계속합니다.
  if (apiErrors.length > 0 || (skippedQueries.length >= 3 && skippedQueries.length * 2 >= responses.length) || (responses.length > 0 && skippedQueries.length === responses.length)) {
    throw new Error('YES24 조회 중 오류가 발생해 결과와 진행 위치를 보존했습니다. API 키와 호출 한도를 확인해 주세요.\n' + apiErrors.concat(skippedQueries).slice(0, 3).join('\n'));
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
    errorCount: skippedQueries.length,
    skippedQueries: skippedQueries.slice(0, 5),
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
  // 출판사 표기가 달라 따로 묶인 같은 시리즈의 묶음에 이미 있는 권이고, 그 묶음의 출판사가 이 책의 출판사와 맞으면 가진 책입니다.
  if ((group.siblingHeld || []).some(function(other) {
    return other.volumes[parsed.volume] && seriesPublisherMatches_(other.publishers, item.publisher);
  })) return null;

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
        unifiedComposeNote_(
          (book.saleStatus === '절판' ? '절판 · ' : '') +
            (book.type === '신간 후속권'
              ? '현재 마지막 권 다음에 나온 책'
              : '보유 권차 사이에서 빠진 책'),
          unifiedBookFlags_(book.title, '')
        )
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
        : '위쪽의 「🔎 시리즈 누락 점검 시작 ›」 막대를 누르면 결과가 채워집니다.'
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
    '검색 결과에서 필요한 책을 체크한 뒤 위쪽의 「📤 선택한 책 구입 검토함으로 보내기 ›」 막대를 누르세요.',
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
    .addItem('🧭 시트 탭 정리', 'HR_arrangeTabs')
    .addSeparator()
    .addItem('📚 희망도서를 구입 검토함으로 보내기', 'sendHopeRequestsToReview')
    .addItem('🔎 시리즈 누락 점검 (쉬운 창)', 'HR_showSeriesCheckDialog')
    .addItem('⏹ 시리즈 점검 멈추기', 'HR_seriesStopCheck')
    .addItem('🔎 시리즈 누락·후속권 이어서 점검', 'refreshMissingSeriesBooks')
    .addItem('🛒 체크한 시리즈를 구입 검토함으로 보내기', 'sendSelectedSeriesToReview')
    .addItem('📥 업로드 목록을 열기', 'openUnifiedUploadSheet')
    .addItem('💵 예산 설정', 'unifiedSetBudget')
    .addItem('✅ 이번 회차 구입 확정', 'unifiedConfirmPurchaseRound')
    .addItem('🧾 구입 기록 보기', 'unifiedShowPurchaseLog')
    .addItem('🧾 구입 기록 숨기기', 'unifiedHidePurchaseLog')
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
  unifiedEnsureUploadSheetReady_();
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
  var requestersColumn = headers.map(function(h) { return String(h || '').replace(/\s+/g, ''); }).indexOf('신청인원') + 1;

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
      category: map.category ? String(row[map.category - 1] || '') : '',
      price: price || 0,
      quantity: 1,
      requesters: requestersColumn ? Math.max(1, Number(row[requestersColumn - 1]) || 1) : 1
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
    if (!map.category && /^(분야|카테고리|주제|KDC)$/.test(h)) map.category = index + 1;
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

  // 문제집·학습지 의심은 줄 전체를 연한 빨강으로, 만화는 도서명 칸을 연한 보라로 표시합니다.
  sheet.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=REGEXMATCH($L6&"","문제집·학습지")')
      .setBackground('#FDECEA')
      .setRanges([sheet.getRange(6, 1, Math.max(1, sheet.getMaxRows() - 5), 12)])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=REGEXMATCH($L6&"","만화")')
      .setBackground('#EFE6F7')
      .setRanges([sheet.getRange(6, 4, Math.max(1, sheet.getMaxRows() - 5), 1)])
      .build()
  ]);

  unifiedDrawActionBars_(sheet);   // 체크박스 대신 클릭 막대
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

  // 문제집·학습지로 보이는 책이 있으면 한 번 더 확인합니다.
  if (!unifiedConfirmWorkbooks_(books)) {
    ss.toast('서점 전달용 만들기를 멈췄습니다. 구입 검토함에서 해당 책을 확인해 주세요.', '문제집·학습지 확인', 8);
    return { success: false, cancelled: true };
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
  PropertiesService.getDocumentProperties().setProperty(UNIFIED_EXPORT_BUILT_AT_PROPERTY, String(Date.now()));   // 만든 시각 (대시보드에 표시)
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

  unifiedDrawActionBars_(sheet);   // 체크박스 대신 클릭 막대
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

// 대시보드 모양을 바꾸면 이 번호를 올립니다. 번호가 다르면 숫자만 고치는 대신 전체를 다시 그립니다.
var UNIFIED_DASHBOARD_LAYOUT = '5';
var UNIFIED_BUDGET_PROPERTY = 'UNIFIED_BUDGET';
var UNIFIED_DASHBOARD_LAYOUT_PROPERTY = 'UNIFIED_DASHBOARD_LAYOUT';
var UNIFIED_EXPORT_BUILT_AT_PROPERTY = 'UNIFIED_EXPORT_BUILT_AT';

// 학교 예산(원). 설정하지 않았으면 0
function unifiedGetBudget_() {
  var raw = PropertiesService.getScriptProperties().getProperty(UNIFIED_BUDGET_PROPERTY);
  var value = Number(raw);
  return isFinite(value) && value > 0 ? Math.round(value) : 0;
}

// 입력한 글을 예산(원)으로 바꿉니다. 비어 있으면 0(예산 끔), 숫자로 읽을 수 없으면 null.
// 예: 3000000 · 3,000,000원 · 300만원 · 2.5만
function unifiedParseBudget_(input) {
  var text = String(input == null ? '' : input).replace(/[,\s원₩]/g, '');
  if (!text || /여기에입력/.test(text)) return 0;
  var man = text.match(/^(\d+(?:\.\d+)?)만$/);
  var value = man ? Math.round(Number(man[1]) * 10000) : Number(text);
  return isFinite(value) && value > 0 ? Math.round(value) : null;
}

function unifiedSaveBudget_(value) {
  var props = PropertiesService.getScriptProperties();
  if (value) props.setProperty(UNIFIED_BUDGET_PROPERTY, String(Math.round(value)));
  else props.deleteProperty(UNIFIED_BUDGET_PROPERTY);
}

// 대시보드의 「총 예산」 칸(A17)에 직접 입력했을 때: 저장하고 대시보드를 다시 계산합니다.
function unifiedBudgetFromDashboard_(range) {
  var value = unifiedParseBudget_(range.getValue());
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (value === null) {
    ss.toast('예산은 숫자로 입력해 주세요. 예: 3000000 또는 300만원', '💵 예산', 7);
  } else {
    unifiedSaveBudget_(value);
    ss.toast(value ? '총 예산을 ' + unifiedMoney_(value) + '으로 정했습니다.' : '예산 표시를 껐습니다.', '💵 예산', 5);
  }
  unifiedUpdateDashboardNumbers_();   // 잘못 입력했으면 원래 값으로 되돌려 보여 줍니다.
}

// 메뉴: 「💵 예산 설정」 (대시보드 칸에 직접 입력하는 것이 더 쉽지만, 메뉴로도 입력할 수 있습니다)
function unifiedSetBudget() {
  var ui = SpreadsheetApp.getUi();
  var current = unifiedGetBudget_();
  var response = ui.prompt(
    '💵 예산 설정',
    '이번 도서 구입 총 예산(원)을 숫자로 입력해 주세요. 예: 3000000 또는 300만원\n' +
      (current ? '지금 설정: ' + unifiedMoney_(current) + '\n' : '') +
      '비워 두고 확인을 누르면 예산 표시를 끕니다.\n' +
      '(대시보드의 「총 예산」 칸에 바로 입력해도 됩니다.)',
    ui.ButtonSet.OK_CANCEL
  );
  if (response.getSelectedButton() !== ui.Button.OK) return;

  var value = unifiedParseBudget_(response.getResponseText());
  if (value === null) {
    ui.alert('예산을 숫자로 입력해 주세요. (예: 3000000, 300만원)');
    return;
  }
  unifiedSaveBudget_(value);
  unifiedUpdateDashboardNumbers_();
  SpreadsheetApp.getActiveSpreadsheet().toast(
    value ? '예산을 ' + unifiedMoney_(value) + '으로 설정했습니다.' : '예산 표시를 껐습니다.',
    '💵 예산', 6
  );
}

// 책 한 권을 가리키는 열쇠: ISBN이 있으면 ISBN, 없으면 도서명
function unifiedBookKey_(isbn, title) {
  var cleanIsbn = seriesCleanIsbn_(isbn);
  if (cleanIsbn) return 'i:' + cleanIsbn;
  var titleKey = seriesNormalizeKey_(title);
  return titleKey ? 't:' + titleKey : '';
}

function unifiedMoney_(amount) {
  return String(Math.round(Number(amount) || 0)).replace(/\B(?=(\d{3})+(?!\d))/g, ',') + '원';
}

// 대시보드에 쓸 숫자를 한 번에 셉니다.
// onlyPurchase 이면 구입 진행 현황(결정·서점 전달)만 셉니다. (체크를 바꿀 때마다 빠르게 고치기 위해)
function unifiedDashboardStats_(ss, onlyPurchase) {
  var stats = {};

  if (!onlyPurchase) {
    var requestSource = ss.getSheetByName(UNIFIED_REQUEST_SHEET);
    stats.request = requestSource ? unifiedReadRequestItems_(requestSource).length : 0;
    stats.series = unifiedCountData_(ss.getSheetByName(SERIES_FINDER_SHEET), 6);
    stats.upload = unifiedCountData_(ss.getSheetByName(UNIFIED_UPLOAD_SHEET), 6);
    stats.review = unifiedCountData_(ss.getSheetByName(UNIFIED_REVIEW_SHEET), 6);
    stats.audience = unifiedAudienceCounts_(ss);
  }

  // 구입 검토함에서 체크한(또는 구입결정인) 책
  var review = ss.getSheetByName(UNIFIED_REVIEW_SHEET);
  var decidedKeys = {};
  var decided = 0;
  var decidedAmount = 0;
  var workbookDecided = 0;
  var comicDecided = 0;
  if (review && review.getLastRow() >= 6) {
    review.getRange(6, 1, review.getLastRow() - 5, 12).getValues().forEach(function(row) {
      if (!String(row[3] || '').trim()) return;                       // 도서명이 없는 줄은 책이 아닙니다.
      if (!(row[0] === true || String(row[2]) === '구입결정')) return;
      decided++;
      var decidedFlags = unifiedBookFlags_(row[3], '');
      var decidedNoted = unifiedFlagsFromNote_(row[11]);
      if (decidedFlags.workbook || decidedNoted.workbook) workbookDecided++;
      if (decidedFlags.comic || decidedNoted.comic) comicDecided++;
      var price = Number(row[8] || 0);
      var quantity = Math.max(1, Number(row[9] || 1));
      decidedAmount += Number(row[10] || 0) || price * quantity;
      var key = unifiedBookKey_(row[7], row[3]);
      if (key) decidedKeys[key] = true;
    });
  }
  stats.decided = decided;
  stats.workbookDecided = workbookDecided;
  stats.comicDecided = comicDecided;
  stats.decidedAmount = decidedAmount;
  stats.budget = unifiedGetBudget_();
  var purchaseLog = unifiedPurchaseLogStats_(ss);
  stats.confirmedAmount = purchaseLog.amount;
  stats.confirmedRounds = purchaseLog.rounds;
  stats.confirmedQuantity = purchaseLog.quantity;

  // 서점 전달용에 만들어 둔 책
  var exportSheet = ss.getSheetByName(UNIFIED_EXPORT_SHEET);
  var exportKeys = {};
  var exportCount = 0;
  var exportAmount = 0;
  if (exportSheet && exportSheet.getLastRow() >= 6) {
    exportSheet.getRange(6, 1, exportSheet.getLastRow() - 5, 10).getValues().forEach(function(row) {
      if (!String(row[1] || '').trim()) return;
      exportCount++;
      var price = Number(row[6] || 0);
      var quantity = Math.max(1, Number(row[7] || 1));
      exportAmount += Number(row[8] || 0) || price * quantity;
      var key = unifiedBookKey_(row[5], row[1]);
      if (key) exportKeys[key] = true;
    });
  }
  stats.exportCount = exportCount;
  stats.exportAmount = exportAmount;

  // 서점 전달용 상태: none(아직 안 만듦) · same(검토함 선택과 같음) · differ(다름) · stale(체크가 0권인데 목록이 있음)
  var builtAt = Number(PropertiesService.getDocumentProperties().getProperty(UNIFIED_EXPORT_BUILT_AT_PROPERTY) || 0);
  stats.builtAtText = builtAt
    ? Utilities.formatDate(new Date(builtAt), 'Asia/Seoul', 'M월 d일 H:mm')
    : '';
  if (!exportCount) {
    stats.exportState = 'none';
  } else if (!decided) {
    stats.exportState = 'stale';
  } else {
    var a = Object.keys(decidedKeys), b = Object.keys(exportKeys);
    var same = a.length === b.length && a.every(function(key) { return exportKeys[key]; });
    stats.exportState = same ? 'same' : 'differ';
  }
  return stats;
}

// 숫자와 설명 글만 다시 채웁니다. (모양은 그대로)
function unifiedFillDashboardNumbers_(sheet, stats, onlyPurchase) {
  var GREEN_MID = '#2E5B4F', MUTED = '#6C7F77', RED = '#B3261E';

  if (!onlyPurchase) {
    sheet.getRange('A2').setValue(
      '희망도서·시리즈 후보·업로드 목록을 모아 중복을 확인하고 서점 전달용 목록으로 정리합니다.   ·   갱신 ' +
      Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy년 M월 d일 H:mm')
    );
    sheet.getRange('A6').setValue(stats.request + '권');
    sheet.getRange('D6').setValue(stats.series + '권');
    sheet.getRange('G6').setValue(stats.upload + '권');
    sheet.getRange('J6').setValue(stats.review + '권');
    sheet.getRange('A8').setValue(
      stats.audience ? '교직원 ' + stats.audience.staff + '권 · 학부모 ' + stats.audience.parent + '권' : '교직원·학부모 신청'
    );
  }

  // 구입 진행 현황
  sheet.getRange('A12').setValue(stats.decided + '권');
  var decidedNote = stats.decided
    ? '검토함에서 체크한 책 · ' + unifiedMoney_(stats.decidedAmount)
    : '아직 체크한 책이 없습니다';
  if (stats.workbookDecided) decidedNote += '\n⚠ 문제집·학습지로 보이는 책 ' + stats.workbookDecided + '권 포함';
  else if (stats.comicDecided) decidedNote += ' · 📙 만화 ' + stats.comicDecided + '권';
  sheet.getRange('A13').setValue(decidedNote)
    .setFontColor(stats.workbookDecided ? RED : MUTED).setFontWeight(stats.workbookDecided ? 'bold' : 'normal');

  sheet.getRange('E12').setValue(stats.exportCount + '권');
  var note;
  var noteColor = MUTED, noteWeight = 'normal';
  if (stats.exportState === 'none') {
    note = '아직 만들지 않았습니다';
  } else if (stats.exportState === 'same') {
    note = '✔ 검토함 선택과 같습니다' + (stats.builtAtText ? ' · ' + stats.builtAtText + ' 만듦' : '');
    noteColor = GREEN_MID;
  } else if (stats.exportState === 'differ') {
    note = '⚠ 검토함 선택과 다릅니다 · 다시 만들어 주세요';
    noteColor = RED; noteWeight = 'bold';
  } else {
    note = '이전에 만든 목록입니다' + (stats.builtAtText ? ' · ' + stats.builtAtText : '');
  }
  sheet.getRange('E13').setValue(note).setFontColor(noteColor).setFontWeight(noteWeight);

  // 예산 현황: 총 예산 − (지금까지 확정한 회차 금액 + 이번 회차 예정 금액)
  var budget = Number(stats.budget || 0);
  var confirmed = Number(stats.confirmedAmount || 0);
  var planned = Number(stats.decidedAmount || 0);
  var used = confirmed + planned;
  var remain = budget - used;
  sheet.getRange('A17').setValue(budget ? unifiedMoney_(budget) : '✏️  여기에 입력');
  sheet.getRange('D17').setValue(unifiedMoney_(confirmed));
  sheet.getRange('D18').setValue(stats.confirmedRounds ? stats.confirmedRounds + '회 확정 · ' + (stats.confirmedQuantity || 0) + '권' : '아직 확정한 회차가 없습니다');
  sheet.getRange('G17').setValue(unifiedMoney_(planned));
  sheet.getRange('G18').setValue(stats.decided ? '검토함에서 체크한 ' + stats.decided + '권' : '체크한 책 없음');
  sheet.getRange('J17').setValue(budget ? (remain < 0 ? '-' + unifiedMoney_(-remain) : unifiedMoney_(remain)) : '-')
    .setFontColor(budget && remain < 0 ? RED : '#1F5C4A');
  var bar = sheet.getRange('A19');
  if (budget) {
    bar.setFormula('=SPARKLINE(' + Math.min(used, budget) + ',{"charttype","bar";"max",' + budget + ';"color1","' + (remain < 0 ? '#B3261E' : '#2E5B4F') + '";"color2","#E3EFE9"})');
  } else {
    bar.setValue('');
  }
  var percent = budget ? Math.round(used / budget * 100) : 0;
  sheet.getRange('A20').setValue(
    !budget ? '총 예산을 아직 정하지 않았습니다. 위의 노란 「총 예산」 칸에 금액을 쓰고 Enter를 누르세요.'
      : remain < 0 ? '⚠ 예산을 ' + unifiedMoney_(-remain) + ' 넘었습니다 · 사용률 ' + percent + '%'
      : '사용률 ' + percent + '% · 아직 ' + unifiedMoney_(remain) + ' 쓸 수 있습니다 (확정 ' + unifiedMoney_(confirmed) + ' + 이번 회차 예정 ' + unifiedMoney_(planned) + ')'
  ).setFontColor(budget && remain < 0 ? RED : MUTED).setFontWeight(budget && remain < 0 ? 'bold' : 'normal');

  sheet.getRange('I12').setValue(unifiedMoney_(stats.exportAmount));
  sheet.getRange('I13').setValue(stats.exportCount ? '서점 전달용 정가 × 수량 합계' : '서점 전달용을 만들면 표시됩니다')
    .setFontColor(MUTED).setFontWeight('normal');
}

// 체크를 바꾸거나 서점 전달용 금액·수량을 고칠 때 숫자만 바로 고칩니다.
// full 이면 모인 책 현황(희망도서·시리즈·업로드·검토 권수)까지 모두 다시 셉니다.
// 모양이 예전 것이면(처음 한 번) 전체를 다시 그립니다.
function unifiedUpdateDashboardNumbers_(full) {
  var ss = seriesSpreadsheet_();
  var sheet = ss.getSheetByName(UNIFIED_DASHBOARD_SHEET);
  if (!sheet) return;
  var layout = PropertiesService.getDocumentProperties().getProperty(UNIFIED_DASHBOARD_LAYOUT_PROPERTY);
  if (layout !== UNIFIED_DASHBOARD_LAYOUT) {
    unifiedRefreshDashboard_();
    return;
  }
  unifiedFillDashboardNumbers_(sheet, unifiedDashboardStats_(ss, !full), !full);
}

// 대시보드 모양을 그립니다. 숫자는 unifiedFillDashboardNumbers_ 가 채웁니다.
function unifiedRefreshDashboard_() {
  var ss = seriesSpreadsheet_();
  var sheet = ss.getSheetByName(UNIFIED_DASHBOARD_SHEET) || ss.insertSheet(UNIFIED_DASHBOARD_SHEET, 0);

  sheet.getDataRange().getMergedRanges().forEach(function(range) {
    range.breakApart();
  });
  sheet.clear();
  sheet.setHiddenGridlines(true);

  var GREEN = '#183D35', GREEN_MID = '#2E5B4F', GREEN_SOFT = '#E3EFE9', GOLD = '#D9A441', GOLD_SOFT = '#FBF1D4';
  var INK = '#26352F', MUTED = '#6C7F77', LINE = '#CFE0D8';

  // 모든 줄 높이를 먼저 기본으로 돌려 둡니다. (이전 모양이 남지 않게)
  sheet.setRowHeights(1, 28, 24);
  for (var column = 1; column <= 12; column++) sheet.setColumnWidth(column, 96);
  sheet.getRange('A1:L28').setVerticalAlignment('middle').setFontColor(INK);

  // 1) 제목 띠
  sheet.getRange('A1:L1').merge().setValue('🏠 학교도서관 통합 구입 업무')
    .setBackground(GREEN).setFontColor('#FFFFFF').setFontSize(26).setFontWeight('bold')
    .setHorizontalAlignment('left').setBorder(false, false, false, false, false, false);
  sheet.getRange('A2:L2').merge().setValue('')
    .setBackground(GREEN).setFontColor('#BFD6CC').setFontSize(11).setHorizontalAlignment('left');
  sheet.getRange('A3:L3').setBackground(GOLD);
  sheet.setRowHeight(1, 64);
  sheet.setRowHeight(2, 32);
  sheet.setRowHeight(3, 5);

  // 2) 모인 책 현황 카드 4개 (제목 / 숫자 / 설명)
  sheet.getRange('A4:L4').merge().setValue('📊  모인 책 현황')
    .setFontSize(13).setFontWeight('bold').setFontColor(GREEN_MID).setHorizontalAlignment('left');
  sheet.setRowHeight(4, 38);
  var cards = [
    ['A', 'C', '📚 희망도서 신청', '교직원·학부모 신청'],
    ['D', 'F', '🔎 시리즈 후보', '누락·후속권 점검 결과'],
    ['G', 'I', '📥 파일 업로드', '엑셀·CSV로 올린 목록'],
    ['J', 'L', '🛒 통합 검토', '구입 검토함에 모인 책']
  ];
  cards.forEach(function(card) {
    sheet.getRange(card[0] + '5:' + card[1] + '5').merge().setValue(card[2])
      .setFontSize(11).setFontWeight('bold').setFontColor(GREEN_MID).setBackground(GREEN_SOFT).setHorizontalAlignment('center');
    sheet.getRange(card[0] + '6:' + card[1] + '7').merge().setValue('0권')
      .setFontSize(30).setFontWeight('bold').setFontColor(GREEN).setBackground('#FFFFFF').setHorizontalAlignment('center');
    sheet.getRange(card[0] + '8:' + card[1] + '8').merge().setValue(card[3])
      .setFontSize(10).setFontColor(MUTED).setBackground('#FFFFFF').setHorizontalAlignment('center');
    sheet.getRange(card[0] + '5:' + card[1] + '8')
      .setBorder(true, true, true, true, false, false, LINE, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    sheet.getRange(card[0] + '5:' + card[1] + '5')
      .setBorder(true, true, false, true, false, false, GOLD, SpreadsheetApp.BorderStyle.SOLID_THICK);
  });
  sheet.setRowHeight(5, 32);
  sheet.setRowHeight(6, 30);
  sheet.setRowHeight(7, 30);
  sheet.setRowHeight(8, 28);

  // 3) 구입 진행 현황 (금색): 제목(11) / 숫자(12) / 설명(13)
  sheet.getRange('A10:L10').merge().setValue('💰  구입 진행 현황')
    .setFontSize(13).setFontWeight('bold').setFontColor(GREEN_MID).setHorizontalAlignment('left');
  sheet.setRowHeight(9, 12);
  sheet.setRowHeight(10, 38);
  var kpis = [
    ['A', 'D', '✅ 구입 결정'],
    ['E', 'H', '📦 서점 전달용'],
    ['I', 'L', '💵 서점 전달 금액']
  ];
  kpis.forEach(function(kpi) {
    sheet.getRange(kpi[0] + '11:' + kpi[1] + '11').merge().setValue(kpi[2])
      .setFontSize(11).setFontWeight('bold').setFontColor('#7A5A12').setHorizontalAlignment('center');
    sheet.getRange(kpi[0] + '12:' + kpi[1] + '12').merge().setValue('0')
      .setFontSize(22).setFontWeight('bold').setFontColor('#5B410A').setHorizontalAlignment('center');
    sheet.getRange(kpi[0] + '13:' + kpi[1] + '13').merge().setValue('')
      .setFontSize(10).setFontColor(MUTED).setHorizontalAlignment('center').setWrap(true);
    sheet.getRange(kpi[0] + '11:' + kpi[1] + '13').setBackground(GOLD_SOFT)
      .setBorder(true, true, true, true, false, false, GOLD, SpreadsheetApp.BorderStyle.SOLID);
  });
  sheet.setRowHeight(11, 28);
  sheet.setRowHeight(12, 44);
  sheet.setRowHeight(13, 38);

  // 3-2) 예산 현황: 제목(15) / 이름(16) / 값(17) / 설명(18) / 막대(19) / 요약(20)
  var logSheet = unifiedEnsurePurchaseLogSheet_();
  // 예전에 보이는 채로 만들어진 기록 시트는 한 번 숨깁니다.
  var docProps = PropertiesService.getDocumentProperties();
  if (docProps.getProperty('UNIFIED_LOG_HIDDEN_ONCE') !== '1') {
    try { if (!logSheet.isSheetHidden()) logSheet.hideSheet(); } catch (error) {}
    docProps.setProperty('UNIFIED_LOG_HIDDEN_ONCE', '1');
  }
  sheet.getRange('A15:L15').merge().setValue('🧾  예산 현황')
    .setFontSize(13).setFontWeight('bold').setFontColor(GREEN_MID).setHorizontalAlignment('left');
  sheet.setRowHeight(14, 12);
  sheet.setRowHeight(15, 38);
  [['A', 'C', '💰 총 예산 (눌러서 입력)'], ['D', 'F', '✅ 지금까지 확정 (눌러서 보기)'], ['G', 'I', '🛒 이번 회차 예정'], ['J', 'L', '🟢 남은 예산']].forEach(function(item) {
    var label = sheet.getRange(item[0] + '16:' + item[1] + '16').merge();
    label.setValue(item[2]);   // 「지금까지 확정」 칸은 누르면 숨겨 둔 기록 시트가 열립니다.
    label.setFontSize(11).setFontWeight('bold').setFontColor(GREEN_MID).setHorizontalAlignment('center').setBackground(GREEN_SOFT);
    sheet.getRange(item[0] + '17:' + item[1] + '17').merge().setValue('-')
      .setFontSize(20).setFontWeight('bold').setFontColor('#1F5C4A').setHorizontalAlignment('center').setBackground('#FFFFFF');
    sheet.getRange(item[0] + '18:' + item[1] + '18').merge().setValue('')
      .setFontSize(9).setFontColor(MUTED).setHorizontalAlignment('center').setBackground('#FFFFFF');
    sheet.getRange(item[0] + '16:' + item[1] + '18')
      .setBorder(true, true, true, true, false, false, LINE, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  });
  sheet.getRange('A19:L19').merge().setBackground('#FFFFFF').setVerticalAlignment('middle');
  sheet.getRange('A20:L20').merge().setValue('')
    .setFontSize(10).setFontColor(MUTED).setHorizontalAlignment('center').setBackground('#FFFFFF');
  sheet.getRange('A19:L20').setBorder(true, true, true, true, false, false, LINE, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  // 총 예산 칸은 직접 쓰는 입력칸입니다. (눈에 띄게 노란 바탕·굵은 금색 테두리)
  sheet.getRange('A17:C17').setBackground(GOLD_SOFT).setFontColor('#7A5A12')
    .setBorder(true, true, true, true, false, false, GOLD, SpreadsheetApp.BorderStyle.SOLID_THICK);
  sheet.getRange('A17').setNote('여기에 총 예산을 쓰고 Enter를 누르세요.\n예: 3000000 · 3,000,000원 · 300만원\n지우면 예산 표시가 꺼집니다.');
  sheet.setRowHeight(16, 28);
  sheet.setRowHeight(17, 42);
  sheet.setRowHeight(18, 22);
  sheet.setRowHeight(19, 24);
  sheet.setRowHeight(20, 28);

  // 4) 업무 순서 (바로가기 포함): 제목(22) / 단계(23) / 설명(24) / 바로가기(25)
  sheet.getRange('A22:L22').merge().setValue('🧭  업무 순서')
    .setFontSize(13).setFontWeight('bold').setFontColor(GREEN_MID).setHorizontalAlignment('left');
  sheet.setRowHeight(21, 12);
  sheet.setRowHeight(22, 38);
  var steps = [
    ['A', 'C', '1', '자료 모으기', '희망도서·시리즈·엑셀 목록', UNIFIED_REQUEST_SHEET, '📚 희망도서 목록'],
    ['D', 'F', '2', '구입 검토', '구입할 책 체크·수량 확인', UNIFIED_REVIEW_SHEET, '🛒 구입 검토함'],
    ['G', 'I', '3', '중복 확인', '소장도서와 자동 대조', null, '🏫 소장도서 확인'],
    ['J', 'L', '4', '서점 전달·확정', '주문 후 「구입 확정」으로 회차 기록', UNIFIED_EXPORT_SHEET, '📦 서점 전달용']
  ];
  steps.forEach(function(step) {
    var c1 = step[0], c2 = step[1];
    sheet.getRange(c1 + '23:' + c2 + '23').merge().setValue(step[2] + '  ' + step[3])
      .setBackground(GREEN_MID).setFontColor('#FFFFFF').setFontSize(12).setFontWeight('bold').setHorizontalAlignment('center');
    sheet.getRange(c1 + '24:' + c2 + '24').merge().setValue(step[4])
      .setBackground('#FFFFFF').setFontSize(10).setFontColor(MUTED).setHorizontalAlignment('center').setWrap(true);
    var link = sheet.getRange(c1 + '25:' + c2 + '25').merge().setBackground('#FFFFFF')
      .setFontSize(11).setFontWeight('bold').setFontColor(GREEN_MID).setHorizontalAlignment('center');
    var target = step[5] ? ss.getSheetByName(step[5]) : (ss.getSheetByName('🏫 우리학교 소장도서') || null);
    if (target) {
      link.setFormula('=HYPERLINK("#gid=' + target.getSheetId() + '","' + step[6] + ' ›")');
    } else {
      link.setValue(step[6]);
    }
    sheet.getRange(c1 + '23:' + c2 + '25')
      .setBorder(true, true, true, true, false, false, LINE, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  });
  sheet.setRowHeight(23, 34);
  sheet.setRowHeight(24, 30);
  sheet.setRowHeight(25, 32);

  // 5) 안내
  sheet.getRange('A27:L27').merge()
    .setValue('💡 필요한 기능은 위쪽 「🧰 통합 도서업무」 메뉴에서 실행하고, 위 업무 순서의 시트 이름을 누르면 해당 시트로 바로 이동합니다.')
    .setBackground(GOLD_SOFT).setFontSize(10).setFontColor('#7A5A12').setHorizontalAlignment('left').setWrap(true);
  sheet.setRowHeight(26, 12);
  sheet.setRowHeight(27, 34);
  sheet.setFrozenRows(3);

  // 숫자 채우기와 모양 번호 기록
  unifiedFillDashboardNumbers_(sheet, unifiedDashboardStats_(ss, false), false);
  PropertiesService.getDocumentProperties().setProperty(UNIFIED_DASHBOARD_LAYOUT_PROPERTY, UNIFIED_DASHBOARD_LAYOUT);
}

// 희망도서 목록에서 교직원/학부모 신청 권수를 셉니다. (열이 없거나 문제가 있으면 null)
function unifiedAudienceCounts_(ss) {
  try {
    var list = ss.getSheetByName(UNIFIED_REQUEST_SHEET);
    if (!list || list.getLastRow() < 6 || list.getLastColumn() < 1) return null;
    var headers = list.getRange(5, 1, 1, list.getLastColumn()).getDisplayValues()[0];
    var col = 0;
    for (var i = 0; i < headers.length; i++) {
      if (String(headers[i]).replace(/\s+/g, '') === '신청구분') { col = i + 1; break; }
    }
    var titleCol = 2;
    var rows = list.getRange(6, 1, list.getLastRow() - 5, list.getLastColumn()).getDisplayValues();
    var staff = 0, parent = 0;
    rows.forEach(function(row) {
      if (!String(row[titleCol - 1] || '').trim()) return;
      if (col && String(row[col - 1]).trim() === '학부모') parent++; else staff++;
    });
    return { staff: staff, parent: parent };
  } catch (error) {
    return null;
  }
}

function unifiedCountData_(sheet, startRow) {
  if (!sheet || sheet.getLastRow() < startRow) return 0;
  return sheet.getRange(startRow, 1, sheet.getLastRow() - startRow + 1, Math.max(1, sheet.getLastColumn()))
    .getDisplayValues().filter(function(row) {
      // 빈 체크박스(FALSE)만 들어 있는 줄은 책이 아니므로 세지 않습니다.
      return row.some(function(value) {
        var text = String(value || '').trim();
        return text && text.toUpperCase() !== 'FALSE';
      });
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
      grouped[key].requesters = Math.max(Number(grouped[key].requesters) || 1, Number(item.requesters) || 1);
    }

    grouped[key].sources[item.source] = true;
  });

  // 많이 신청한 책이 위로 오도록 정렬합니다(같은 인원이면 들어온 순서 그대로).
  order = order.map(function(key, index) { return { key: key, index: index }; }).sort(function(a, b) {
    return (Number(grouped[b.key].requesters) || 1) - (Number(grouped[a.key].requesters) || 1) || a.index - b.index;
  }).map(function(entry) { return entry.key; });

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
      HR_composeRequesterNote_(unifiedComposeNote_(manual.note || item.note || '', unifiedBookFlags_(item.title, item.category)), item.requesters)
    ];
  });
  var workbookCount = output.filter(function(row) { return unifiedFlagsFromNote_(row[11]).workbook; }).length;

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
    ownedSkipped: ownedSkipped,
    workbookCount: workbookCount
  };
}

function sendHopeRequestsToReview() {
  var ss = seriesSpreadsheet_();
  var result = unifiedRefreshPurchaseReview_([]);
  var review = ss.getSheetByName(UNIFIED_REVIEW_SHEET);
  if (review) ss.setActiveSheet(review);

  ss.toast(
    '희망도서를 구입 검토함에 모았습니다. 현재 ' + result.count + '권입니다.' +
      (result.workbookCount ? ' ⚠ 문제집·학습지로 보이는 책이 ' + result.workbookCount + '권 있습니다. 비고를 확인해 주세요.' : ''),
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
    // 처음 만들 때는 「희망도서 목록」 바로 오른쪽에 놓습니다. (목록이 없으면 맨 뒤)
    var listSheet = ss.getSheetByName(UNIFIED_REQUEST_SHEET);
    sheet = listSheet
      ? ss.insertSheet(UNIFIED_UPLOAD_SHEET, listSheet.getIndex())
      : ss.insertSheet(UNIFIED_UPLOAD_SHEET);
  }

  if (sheet.getMaxColumns() < 11) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), 11 - sheet.getMaxColumns());
  }
  if (sheet.getMaxRows() < 100) {
    sheet.insertRowsAfter(sheet.getMaxRows(), 100 - sheet.getMaxRows());
  }

  return sheet;
}

// 「📥 엑셀 업로드 목록」 시트를 미리 만들어 둡니다. (웹에서 파일을 올려야 그제서야 생기지 않도록)
// 이미 있으면 아무것도 바꾸지 않습니다. 「희망도서 목록」 바로 오른쪽에 놓습니다.
function unifiedEnsureUploadSheetReady_() {
  var ss = seriesSpreadsheet_();
  if (ss.getSheetByName(UNIFIED_UPLOAD_SHEET)) return false;

  var sheet = unifiedEnsureUploadSheet_();
  sheet.setHiddenGridlines(true);
  sheet.getRange('A1:K1').merge().setValue('📥 파일 업로드 도서 목록');
  sheet.getRange('A2:K2').merge().setValue(
    '웹페이지의 「📥 목록 업로드」에서 엑셀·CSV를 올리면 필요한 서지정보만 여기에 채워집니다. 구입할 책을 고른 뒤 구입 검토함으로 보내세요.'
  );
  sheet.getRange('A3:K3').merge().setValue('아직 올린 파일이 없습니다.');
  sheet.getRange(5, 1, 1, 11).setValues([[
    '선택', '도서명', '저자', '출판사', '출판연도', 'ISBN13', '정가', '수량', '금액', '시리즈명', '비고'
  ]]);
  unifiedStyleUploadSheet_(sheet, 0);   // 모양과 클릭 막대
  return true;
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
        unifiedComposeNote_(book.note, unifiedBookFlags_(book.title, ''))
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
  unifiedResetBarState_(sheet);

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

  unifiedDrawActionBars_(sheet);   // 체크박스 대신 클릭 막대
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
    ['1', '신청 링크로 들어가기', '선생님은 교직원용 링크, 학부모님은 학부모용 링크(또는 QR)로 들어옵니다. 링크는 「⚙️ 사본 설정 → 🔗 신청 링크 받기」에서 받을 수 있습니다.'],
    ['2', '신청하는 분 고르기', '링크 하나로 들어왔다면 첫 화면에서 「교직원」 또는 「학부모」를 누르십시오. 전용 링크로 들어왔다면 이 단계는 건너뜁니다. 교직원 인증번호를 켜 둔 학교라면 「교직원」을 누른 뒤 인증번호를 입력합니다.'],
    ['3', '신청자 정보 쓰기', '교직원은 이름을 입력합니다. 이름을 입력하면 우선 대출 혜택을 받을 수 있습니다. 학부모는 학년·반·번호를 꼭 고르고, 학생 이름은 선택입니다. 학부모 정보는 신청 목록에 보이지 않습니다.'],
    ['4', '책 검색', '책 제목 또는 ISBN을 입력하고 「책 검색하기」를 누릅니다.'],
    ['5', '책 확인·신청바구니 담기', '표지·저자·출판사·연도를 확인한 뒤 「신청바구니에 담기」를 누릅니다. 여러 권을 담을 수 있고, 한 번에 최대 20권까지 신청할 수 있습니다.'],
    ['6', '선택한 도서 신청', '신청바구니에서 「선택한 도서 신청하기」를 누릅니다. 학교에서 정한 신청 한도(예: 학부모는 한 학생당 5권)를 넘으면 안내 창이 뜹니다.'],
    ['7', '신청 확인·취소', '「신청 목록」에서 접수 결과를 확인합니다. 잘못 신청했다면 취소하고 바로 되돌릴 수도 있습니다. 교직원과 학부모의 신청 목록은 따로 보입니다.']
  ];

  unifiedBuildGuideSheet_(
    '📗 희망도서 신청안내',
    '📗 희망도서 신청안내',
    '교직원과 학부모가 책을 찾고 여러 권을 신청바구니에 담아 신청하는 방법입니다.',
    applicantSteps,
    '신청한 책이 이미 학교에 있거나 이미 신청된 책이면 신청되지 않습니다. 신청 한도는 학교에서 정합니다.'
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

  HR_buildManualSheet_(
    '📘 전체 사용방법',
    HR_MANUAL_TITLE,
    HR_MANUAL_SUBTITLE,
    HR_MANUAL_SECTIONS,
    HR_MANUAL_FOOTER
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

/* ---------- 🏫 소장도서 「출판사」 칸 ---------- */

// 소장도서 탭 머리줄에 「출판사」 칸이 없으면 바로 오른쪽 빈 칸에 만듭니다. (이미 있으면 아무것도 하지 않습니다)
// 시리즈 누락 점검(seriesReadHoldings_)이 이 칸을 읽어 같은 이름의 다른 출판사 시리즈를 구분합니다.
function unifiedEnsureHoldingsPublisherColumn_(ss) {
  ss = ss || seriesSpreadsheet_();
  var sheet = ss.getSheetByName(SERIES_HOLDINGS_SHEET);
  if (!sheet) return false;
  var lastRow = sheet.getLastRow();
  var lastColumn = sheet.getLastColumn();
  if (lastRow < 1 || lastColumn < 1) return false;

  // 머리줄 찾기: 맨 위 10줄 안에서 ISBN 칸이 있는 줄 (seriesReadHoldings_ 와 같은 방식)
  var grid = sheet.getRange(1, 1, Math.min(10, lastRow), lastColumn).getDisplayValues();
  var headerRow = 0;
  var lastHeaderColumn = 0;
  for (var r = 0; r < grid.length && !headerRow; r++) {
    var hasIsbn = grid[r].some(function(value) { return /^ISBN(13)?$/.test(String(value || '').replace(/\s+/g, '').toUpperCase()); });
    if (!hasIsbn) continue;
    headerRow = r + 1;
    for (var c = 0; c < grid[r].length; c++) {
      var header = String(grid[r][c] || '').replace(/\s+/g, '');
      if (/^(출판사|발행자|출판자|발행처)$/.test(header)) return false;   // 이미 있음
      if (header) lastHeaderColumn = c + 1;
    }
  }
  if (!headerRow || !lastHeaderColumn) return false;

  var column = lastHeaderColumn + 1;
  if (sheet.getMaxColumns() < column) sheet.insertColumnsAfter(sheet.getMaxColumns(), column - sheet.getMaxColumns());
  var target = sheet.getRange(headerRow, column);
  if (String(target.getDisplayValue() || '').trim()) return false;   // 그 칸에 다른 글이 있으면 건드리지 않습니다.

  sheet.getRange(headerRow, lastHeaderColumn).copyFormatToRange(sheet, column, column, headerRow, headerRow);
  target.setValue('출판사').setNote('DLS 소장 목록의 출판사를 붙여 넣어 주세요. (도서명·ISBN과 같은 줄 순서) 시리즈 누락 점검이 출판사로 시리즈를 구분합니다.');
  try { sheet.setColumnWidth(column, Math.max(sheet.getColumnWidth(lastHeaderColumn), 160)); } catch (error) {}

  // 머리줄에 필터가 걸려 있으면 새 칸까지 넓힙니다.
  try {
    var filter = sheet.getFilter();
    if (filter) {
      var range = filter.getRange();
      if (range.getRow() === headerRow && range.getColumn() <= lastHeaderColumn && range.getLastColumn() < column) {
        var top = range.getRow(), left = range.getColumn(), rows = Math.max(range.getNumRows(), sheet.getLastRow() - top + 1);
        filter.remove();
        sheet.getRange(top, left, rows, column - left + 1).createFilter();
      }
    }
  } catch (error) {
    console.warn('소장도서 필터 넓히기 실패: ' + error);
  }
  return true;
}


/* ===== UpdateLog.gs ===== */
/**
 * 📜 업데이트 내용 (사본 안의 「업데이트 내용」 탭)
 *
 * 이제까지 새로 생기거나 고쳐진 기능을 한 곳에 모아 보여 줍니다.
 * - 내용은 아래 HR_UPDATE_LOG 에 적고(최신 것을 맨 위에), 안내 시트 버전(HR_GUIDE_VERSION)을 올리면
 *   모든 사본이 다음에 열릴 때 이 탭을 새 내용으로 다시 씁니다.
 * - 선생님이 이 탭에 직접 적은 글은 덮어써질 수 있습니다. (안내용 탭입니다.)
 */
var HR_UPDATE_LOG_SHEET = '업데이트 내용';

// [날짜, 분류, 제목, 내용]  — 최신 것이 위
var HR_UPDATE_LOG = [
  ['2026-11-28', '시리즈 점검', '점검 순서를 바꿈: 준비 → 검색 → 저장 (예전 방식으로 하던 중간 점검은 처음부터 다시 합니다)', '① 준비: 검색을 시작하기 전에 모든 시리즈의 출판사·판본을 YES24 에서 먼저 확인합니다. 예전에는 점검 한 번에 20개씩만 확인해서, 확인 전에 검색된 시리즈는 출판사 필터 없이 다른 출판사 책까지 후보로 걸렸습니다. ② 검색: 그 뒤에 후속권·중간 누락을 검색합니다. ③ 저장: 결과 시트를 다 쓴 뒤에야 진행 위치를 앞으로 옮깁니다(시트를 쓰다 시간이 끝나도 결과가 빠지지 않습니다). 한 번에 점검하는 시간도 줄여 여유를 두었습니다. 예전 방식으로 하다 만 점검은 이어가지 않고 처음부터 다시 합니다.'],
  ['2026-11-28', '시리즈 점검', '이어가기 알람을 「계속 울리는 알람」으로 바꿈', '예전에는 「1분 뒤에 한 번만 울리는 알람」으로 이어갔기 때문에, 그 한 번이 헛돌거나 오류로 끝나면 거기서 이어가기가 끊겨 점검이 멈춰 있었습니다. 이제 점검이 끝날 때까지 1분마다(한도·오류로 쉬는 중에는 10분·5분마다) 알람이 계속 울려서, 한 번 실패해도 다음 알람에서 저절로 이어집니다. 쉬는 시간이 안 지났거나 이미 돌고 있으면 조용히 넘어가고, 다 끝나면 알람을 지웁니다.'],
  ['2026-11-28', '시리즈 점검', '자꾸 멈추던 문제 고침 (일시적 YES24 거절에 점검이 통째로 버려지던 문제)', 'YES24 가 한 묶음 중 한 건만 일시적으로 거절해도 점검이 오류로 끝나고, 그때까지 몇 분간 한 검색이 모두 버려지던 문제를 고쳤습니다. 이제 잠깐 쉬었다 다시 시도하고, 그래도 안 되면 지금까지 한 만큼은 저장한 채 이어서 점검합니다. 키 문제가 아닌 오류로 끝나도 5분 뒤 저절로 다시 시도합니다(연달아 6번까지). 시트 맨 위 상태 줄에 이유가 보입니다.'],
  ['2026-11-28', '시리즈 점검', '점검 상태를 시트 맨 위에 크게 표시', '시리즈 점검 결과 시트 맨 위 둘째 줄에 지금 상태가 색깔로 보입니다. 🟢 점검 중(시작 시각) · ⏳ 대기 중(약 1분 뒤 자동으로 이어감) · ⏳ 한도로 쉬는 중(이어갈 시각) · ✅ 전체 확인 완료 · ⏹ 멈춤 · ⚠ 오류. 점검이 중간에 멈춘 채 예약이 없으면 시트나 점검 창을 열 때 스스로 다시 예약합니다.'],
  ['2026-11-28', '시리즈 점검', '한도 뒤 자동 이어가기 보강', 'YES24 응답에 들어 있는 한도 오류 문구도 알아보도록 넓혔고, 시리즈 점검 창에 「한도 때문에 ○월 ○일 ○○:○○ 무렵에 저절로 이어서 점검하도록 예약되어 있습니다」가 보이게 했습니다. 예약이 안 되어 있으면 이 문구가 보이지 않습니다.'],
  ['2026-11-27', '시리즈 점검', 'YES24 하루 한도에 닿으면 풀릴 때 저절로 이어서 점검', '점검 도중 YES24 하루 조회 한도에 닿으면 오류로 끝나지 않고 잠시 멈춥니다. 우리 쪽 안전 한도에 닿았으면 다음 날 00:10 무렵, YES24가 직접 한도를 알려 왔으면 1시간 뒤에 저절로 이어서 점검합니다(시트를 열어 두지 않아도 됩니다). 진행 위치는 그대로 보존됩니다. 계속 막혀 있으면 6번까지만 다시 시도하고 알려 드립니다. 멈추려면 「⏹ 시리즈 점검 멈추기」를 누르세요.'],
  ['2026-11-27', '시리즈 점검', '출판사는 자동으로 확인 (소장도서에 적지 않아도 됩니다)', '소장도서 탭에 직접 적은 「출판사」 칸은 더 이상 쓰지 않습니다. 일부 줄만 적거나 표기가 서로 달라 같은 시리즈가 갈라지면 가진 책이 누락으로 나오는 일이 있어서, 이제 출판사는 소장도서의 ISBN 으로 YES24 에서 자동으로 확인한 값만 씁니다. 이미 적어 둔 칸은 그대로 두셔도 되고 지우셔도 됩니다. 고친 뒤에는 「시리즈 누락 점검」을 처음부터 다시 실행해 주세요.'],
  ['2026-11-27', '시리즈 점검', '출판사 표기가 달라도 이미 있는 권은 누락으로 나오지 않게', '같은 시리즈인데 소장도서의 출판사 표기가 서로 달라(예: 「고릴라박스」와 「비룡소」) 두 묶음으로 갈라지면, 한쪽에 있는 권을 다른 쪽에서 「중간 누락」으로 잡던 문제를 고쳤습니다. 이제 다른 묶음에 이미 있는 권이고 출판사가 맞으면 가진 책으로 봅니다. 서로 다른 출판사의 같은 이름 시리즈는 예전처럼 따로 점검합니다. 고친 뒤에는 「시리즈 누락 점검」을 처음부터 다시 실행해 주세요.'],
  ['2026-11-27', '시리즈 점검', '시리즈 점검 멈추기', '「🧰 통합 도서업무 → ⏹ 시리즈 점검 멈추기」를 누르면 돌고 있는 시리즈 점검이 몇 초 안에 멈추고, 1분 뒤 자동 이어가기도 취소됩니다. 멈춘 점검의 이번 결과는 쓰지 않고 저장된 점검 위치도 지워서 다음 점검은 처음부터 시작합니다. 이전에 만들어진 결과 시트는 그대로 남습니다.'],
  ['2026-11-27', '시리즈 점검', '점검 막대를 체크하면 점검 창이 뜹니다', '시리즈 점검 시트의 「🔎 시리즈 누락 점검 시작」 ☐ 를 체크하면 「시리즈 누락 점검」 창이 뜹니다. 창에서 「점검 시작」(또는 중간에 멈춘 점검이 있으면 「이어서 점검」·「처음부터 다시 점검하기」)을 누르면 됩니다.'],
  ['2026-11-27', '시리즈 점검', '출판사를 일부만 넣어도 있는 책이 누락으로 나오지 않게', '소장도서의 「출판사」 칸을 일부 줄에만 넣으면 같은 시리즈가 출판사별로 갈라져서, 이미 가진 권이 「누락」으로 나오던 문제를 고쳤습니다. 이제 출판사가 비어 있는 줄은 같은 시리즈의 출판사 묶음에 함께 들어가 보유로 인정됩니다. 서로 다른 출판사의 같은 이름 시리즈는 예전처럼 따로 점검합니다. 고친 뒤에는 「시리즈 누락 점검」을 처음부터 다시 실행해 주세요.'],
  ['2026-11-27', '시트 화면', '막대는 체크칸 없이 그냥 클릭', '☐ 체크칸을 없애고 막대를 그냥 누르면 동작하게 되돌렸습니다. 보내기·서점 전달용 만들기·구입 확정처럼 되돌리기 번거로운 동작은 막대 글자가 「⚠ 한 번 더 누르면 실행합니다」로 바뀌고, 한 번 더 누르면 실행됩니다(다른 칸을 누르면 취소). 확인 창은 뜨지 않습니다. 시리즈 누락 점검 시작만 YES24 접속이 필요해서 앞의 ☐ 를 체크하는 방식을 유지합니다.'],
  ['2026-11-27', '시트 화면', '막대 버튼 디자인 정리', '☐ 체크칸과 막대 글자를 왼쪽으로 붙이고, ☐ 칸은 막대보다 한 톤 진한 바탕에 크게 두어 「여기를 체크」가 한눈에 보이게 하고, 막대 글자가 칸 폭에 맞춰 잘리지 않도록 크기를 맞췄습니다. 한 시트 안의 막대는 모두 같은 글자 크기로 보입니다.'],
  ['2026-11-27', '시리즈 점검', '점검 막대는 창 없이 바로 점검', '시리즈 점검 시트의 「🔎 시리즈 누락 점검 시작」 ☐ 를 체크하면 창 없이 바로 점검을 시작하고, 진행 상황과 결과는 시트 오른쪽 아래 알림으로 알려 줍니다. 메뉴의 「쉬운 창」으로 점검할 때도 시작한 뒤에는 창을 닫아도 됩니다.'],
  ['2026-11-27', '시리즈 점검', '소장도서 탭에 「출판사」 칸', '「🏫 우리학교 소장도서」 탭에 「출판사」 칸이 생겼습니다. DLS 소장 목록의 출판사를 이 칸에 붙여 넣으면(도서명·ISBN과 같은 줄 순서로) 시리즈 누락 점검이 같은 이름의 다른 출판사 시리즈를 섞지 않아 엉뚱한 후보가 줄어듭니다. 비워 두어도 예전처럼 동작합니다.'],
  ['2026-11-27', '시트 화면', '막대 버튼 앞에 ☐ 체크칸', '「시리즈 누락 점검 시작」「구입 검토함으로 보내기」처럼 창이 뜨는 막대를 눌렀을 때 「권한이 없어 창을 띄울 수 없다」는 오류가 나던 문제를 고쳤습니다. 이제 막대 맨 앞의 ☐ 칸을 체크하면 확인창·점검창이 뜹니다. 「전체 선택」은 예전처럼 막대를 누르면 됩니다.'],
  ['2026-11-26', '웹 신청', '이미 있는 책·이미 신청된 책 미리 알림, N명 신청', '검색 결과에서 「우리 학교에 이미 있는 책이에요」「이미 N명이 신청한 책이에요」를 미리 보여 줍니다. 이미 신청된 책을 다른 사람이 또 신청하면 책이 늘지 않고 희망도서 목록의 「신청 인원」이 올라가며(같은 사람은 한 번만), 구입 검토함에는 비고에 「👥 N명 신청」이 붙고 많이 신청한 책이 위로 올라옵니다.'],
  ['2026-11-26', '구입 검토', '만화·문제집 표시와 경고', '구입 검토함·시리즈 후보·엑셀 업로드 목록의 비고에 「📙 만화」가 붙고, 문제집·학습지로 보이는 책은 「⚠ 문제집·학습지 의심」과 함께 줄이 연한 빨강으로 보입니다. 서점 전달용을 만들기 전에 해당 책을 보여 주고 한 번 더 묻습니다. 대시보드에도 경고가 나옵니다. (제목·분야로 추정하므로 틀리면 비고에서 지우면 됩니다.)'],
  ['2026-11-26', '대시보드', '예산 현황과 회차별 구입 기록', '대시보드의 노란 「총 예산」 칸에 금액을 쓰면 남은 예산과 사용률이 나옵니다. 서점에 주문한 뒤 「서점 전달용」의 「✅ 이번 회차 구입 확정」을 누르면 회차·날짜·권수·금액이 기록되고, 남은 예산은 확정한 모든 회차를 합쳐서 계산합니다. 기록은 숨겨진 「🧾 구입 기록」 시트에 쌓이며 대시보드의 「지금까지 확정」 칸으로 열 수 있습니다.'],
  ['2026-11-26', '대시보드', '서점 전달용 일치 표시와 금액', '서점 전달용 목록이 구입 검토함에서 고른 책과 같은지(같음/다름/아직 안 만듦)를 알려 주고, 서점 전달 금액을 표시합니다. 체크를 바꾸거나 목록 줄을 지우거나 시트를 열 때 숫자가 바로 바뀌며, 빈 체크박스만 있는 줄이 권수를 부풀리던 오류도 고쳤습니다.'],
  ['2026-11-26', '시트 화면', '체크박스 대신 클릭 막대', '희망도서 목록·시리즈 점검·업로드 목록·구입 검토함·서점 전달용 맨 위의 버튼을 체크박스가 아니라 막대 전체를 누르면 동작하는 큰 버튼으로 바꾸었습니다. 보내기·만들기는 누르면 한 번 더 확인을 묻고, 막대 비율은 화면 폭에 맞춰집니다.'],
  ['2026-11-26', '시리즈 점검', '쉬운 점검 창과 정확도 개선', '시리즈 점검 시트의 「🔎 시리즈 누락 점검 시작」 막대를 누르면 창이 열리면서 바로 점검을 시작하고, 중간에 멈추면 이어서 합니다. 이름이 전혀 다른 책(예: 「처음 읽는 삼국지」와 「삼국지」)이 후보로 걸리던 원인을 고쳤고, 검색 한 건의 오류로 점검 전체가 멈추거나 반복되던 문제도 고쳤습니다. 소장도서에 출판사를 적어 두면 더 정확해집니다.'],
  ['2026-11-26', '시트 구성', '엑셀 업로드 목록을 미리 만들어 둠', '파일을 올려야 그제서야 생기던 「📥 엑셀 업로드 목록」 시트를 처음부터 희망도서 목록 오른쪽에 만들어 둡니다.'],
  ['2026-11-23', '학부모·교직원 신청', '본인 신청만 보기·취소, 표지 표시', '웹 「신청 목록」에는 지금 입력한 신청자(학부모는 학년·반·번호, 교직원은 이름)의 신청만 표지와 함께 보이고, 취소와 되돌리기도 신청한 본인만 할 수 있습니다.'],
  ['2026-11-23', '학부모·교직원 신청', '안내 화면 다듬기', '학부모 화면에서 도서관 담당자·내선번호를 숨기고, 교사/학부모 선택 화면 맨 위에 「○○초등학교 희망도서 신청」처럼 학교 이름을 보여 줍니다. 검색창과 「책 검색하기」 버튼 높이도 맞췄습니다.'],
  ['2026-11-23', '개발자 문의', '개발자에게 문의하기 · 답변 알림', '메뉴 「⚙️ 사본 설정 → 💬 개발자에게 문의하기」로 질문이나 오류를 보낼 수 있고, 「📬 내 문의·답변 보기」에서 답을 봅니다. 답이 달리면 시트를 열 때 알림이 한 번 뜹니다. (학생 이름·번호 같은 개인정보는 적지 마십시오.)'],
  ['2026-11-23', '안내', '업데이트 소식 창', '새 소식이 있으면 시트를 열 때마다 안내 창이 뜨고, 「이 소식 다시 보지 않기」를 누르면 그 소식은 더 뜨지 않습니다. 메뉴 「📢 업데이트 소식 보기」로 언제든 다시 볼 수 있습니다.'],
  ['2026-11-23', '웹 공지', '웹 공지 시트 디자인과 대상 나누기', '「📢 웹 공지」 시트를 칸마다 색을 넣어 보기 좋게 바꾸고, 공지를 전체·교직원·학부모 중 누구에게 보일지 「대상」 칸에서 고르게 했습니다.'],
  ['2026-11-23', '학부모 신청', '남은 신청 권수 실시간 표시', '학부모가 학년·반·번호를 고르면 「한 학생당 최대 N권 중 M권 더 신청할 수 있습니다」가 나오고, 신청하는 즉시 바뀝니다.'],
  ['2026-11-23', '안정성', '사본이 많아져도 안정적으로 자동 업데이트', '사본이 5분마다 아주 작은 버전 번호만 확인하고, 바뀌었을 때만 새 코드를 받도록 개선했습니다.'],
  ['2026-11-22', '학부모·교직원 신청', '교직원/학부모 구분 신청과 신청 한도', '신청 링크를 교직원용·학부모용으로 나눠 따로 관리하고, 한 사람(학생)당 신청 권수 한도를 정할 수 있습니다. 한 학년 최대 반 수도 학교 설정에서 정합니다. 학부모에게는 입력한 학년·반·번호가 구입 처리에만 쓰인다는 안내를 보여 줍니다.'],
  ['2026-11-22', '보안', '교직원 인증번호', '메뉴 「🔑 교직원 인증번호 설정」으로 번호를 켜 두면, 번호를 아는 교직원만 교직원 신청 화면에 들어갈 수 있습니다. 언제든 끌 수 있습니다.'],
  ['2026-11-22', '대시보드', '업무 대시보드 새 디자인', '희망도서·시리즈·업로드·검토 권수, 구입 진행 현황, 업무 순서 바로가기를 한 장에 모은 대시보드를 새로 꾸몄습니다.'],
  ['2026-11-22', '신청 링크', '신청 링크 받기 · QR 그림', '메뉴 「🔗 신청 링크 받기」에서 학부모용·교직원용 링크를 복사하고, 인쇄할 수 있는 QR 그림을 저장할 수 있습니다.'],
  ['2026-11-22', '안내', '전체 사용방법 새로 작성', '「📘 전체 사용방법」과 「📗 희망도서 신청안내」를 새로 썼습니다. 안내 시트는 업데이트될 때 자동으로 새 내용으로 바뀝니다.'],
  ['2026-11-22', '시트 구성', '시트 탭 정리·이전 사본 데이터 가져오기', '탭을 업무 순서로 정리하고 색으로 묶었으며, 이전 사본의 신청 데이터를 새 사본으로 옮기는 메뉴를 추가했습니다.'],
  ['2026-11-22', '자동 업데이트', '모든 사본 자동 업데이트', '이제 새 기능은 사본을 다시 만들지 않아도 시트를 열면 알아서 최신 버전으로 반영됩니다.']
];

var HR_UPDATE_LOG_COLORS = {
  '웹 신청': '#E3EFE9', '학부모·교직원 신청': '#E3EFE9', '학부모 신청': '#E3EFE9', '신청 링크': '#E3EFE9',
  '구입 검토': '#FBF1D4', '대시보드': '#FBF1D4', '시리즈 점검': '#E8F0FA',
  '시트 화면': '#F1EBF8', '시트 구성': '#F1EBF8', '웹 공지': '#F1EBF8',
  '개발자 문의': '#FDECEA', '안내': '#EEF2F0', '보안': '#FDECEA', '안정성': '#EEF2F0', '자동 업데이트': '#EEF2F0'
};

// 「업데이트 내용」 탭을 HR_UPDATE_LOG 로 다시 씁니다. (없으면 「웹 공지」 오른쪽에 만듭니다.)
function HR_writeUpdateLogSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(HR_UPDATE_LOG_SHEET);
  if (!sheet) {
    var anchor = ss.getSheetByName('📢 웹 공지');
    sheet = anchor ? ss.insertSheet(HR_UPDATE_LOG_SHEET, anchor.getIndex()) : ss.insertSheet(HR_UPDATE_LOG_SHEET);
  }
  sheet.getDataRange().getMergedRanges().forEach(function(range) { range.breakApart(); });
  sheet.clear();
  sheet.clearConditionalFormatRules();
  if (sheet.getMaxColumns() < 4) sheet.insertColumnsAfter(sheet.getMaxColumns(), 4 - sheet.getMaxColumns());
  sheet.setHiddenGridlines(true);
  sheet.setTabColor('#9AA6A0');
  [110, 130, 260, 640].forEach(function(width, i) { sheet.setColumnWidth(i + 1, width); });

  sheet.getRange('A1:D1').merge().setValue('📜 업데이트 내용')
    .setBackground('#183D35').setFontColor('#FFFFFF').setFontSize(22).setFontWeight('bold')
    .setHorizontalAlignment('left').setVerticalAlignment('middle');
  sheet.getRange('A2:D2').merge().setValue('지금까지 새로 생기거나 고쳐진 기능입니다. 최신 내용이 맨 위에 있고, 사본은 시트를 열 때 자동으로 최신 버전으로 바뀝니다.')
    .setBackground('#EAF3ED').setFontColor('#34564D').setFontSize(11).setWrap(true).setVerticalAlignment('middle');
  sheet.setRowHeight(1, 56);
  sheet.setRowHeight(2, 36);

  sheet.getRange(3, 1, 1, 4).setValues([['날짜', '분류', '제목', '내용']])
    .setBackground('#D9A441').setFontColor('#FFFFFF').setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  sheet.setRowHeight(3, 34);
  sheet.setFrozenRows(3);

  var rows = HR_UPDATE_LOG.map(function(entry) { return [entry[0], entry[1], entry[2], entry[3]]; });
  var body = sheet.getRange(4, 1, rows.length, 4);
  body.setNumberFormat('@');   // 날짜가 달력 날짜로 바뀌지 않게 글자로 둡니다.
  body.setValues(rows);
  body.setVerticalAlignment('top').setWrap(true).setFontSize(11).setFontColor('#2E3A33');
  sheet.getRange(4, 1, rows.length, 2).setHorizontalAlignment('center');
  sheet.getRange(4, 3, rows.length, 1).setFontWeight('bold').setFontColor('#244D42');
  sheet.getRange(4, 1, rows.length, 4).setBorder(true, true, true, true, true, true, '#D3DDD7', SpreadsheetApp.BorderStyle.SOLID);

  // 분류 칸은 종류별로 색을 다르게 합니다.
  var colors = rows.map(function(row) { return [HR_UPDATE_LOG_COLORS[row[1]] || '#EEF2F0']; });
  sheet.getRange(4, 2, rows.length, 1).setBackgrounds(colors).setFontWeight('bold');
  // 같은 날짜가 이어질 때 첫 줄에만 굵은 구분선을 둡니다.
  rows.forEach(function(row, index) {
    if (index === 0 || rows[index - 1][0] !== row[0]) {
      sheet.getRange(4 + index, 1, 1, 4).setBorder(true, null, null, null, null, null, '#B9770E', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    }
  });
  return rows.length;
}


/* ===== UpdateNotice.gs ===== */
/**
 * 📢 업데이트 소식
 *
 * 개발자가 「📢 업데이트 소식 관리」 구글 시트에 소식을 쓰면, 사본이 열릴 때마다 안내 창으로 보여 줍니다.
 * - 시트의 「상태」 칸이 「🟢 보이기」(또는 TRUE)인 줄 중 가장 위에 있는 소식 하나를 보여 줍니다.
 * - 상태가 「🟢 보이기」인 동안에는 열 때마다 뜹니다. 보는 사람은 「이 소식 다시 보지 않기」로 끌 수 있고, 개발자는 상태를 끄면 모두에게 멈춥니다.
 * - 시작일·종료일은 비워도 됩니다. 기간 밖이면 보여 주지 않습니다.
 * - 인터넷이 안 되거나 시트를 못 읽으면 조용히 건너뜁니다.
 */
var HR_NOTICE_SHEET_ID = '1aR1O-pNTpSotlsJ5uCvRNHioUdMNhC2vKVOlsluITdM';
var HR_NOTICE_URL =
  'https://docs.google.com/spreadsheets/d/' + HR_NOTICE_SHEET_ID +
  '/export?format=csv';
// 「다시 보지 않기」를 누른 소식 번호. (예전에 한 번 보여 준 기록 HR_NOTICE_SEEN_ID 는 더 이상 쓰지 않습니다.)
var HR_NOTICE_SEEN_PROPERTY = 'HR_NOTICE_DISMISSED_ID';

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
    '.skip{display:block;width:100%;margin-top:8px;padding:11px;border:1px solid #c9d3cc;border-radius:12px;background:#fff;color:#4b5a52;font-size:14px;font-weight:700;cursor:pointer}' +
    '.skip:hover{background:#eef3ef}' +
    '</style></head><body><div class="badge">📢 업데이트 소식</div>' +
    '<h1 class="title">' + HR_escapeNoticeHtml_(notice.title) + '</h1>' +
    '<div class="body">' + HR_escapeNoticeHtml_(notice.body).replace(/\n/g, '<br>') + '</div>' +
    '<button class="ok" onclick="google.script.host.close()">확인했어요</button>' +
    '<button class="skip" id="skip" onclick="var b=this;b.disabled=true;b.textContent=\'처리 중…\';' +
    'google.script.run.withSuccessHandler(function(){google.script.host.close()})' +
    '.withFailureHandler(function(e){b.disabled=false;b.textContent=\'저장하지 못했어요. 다시 눌러 주세요\';})' +
    '.HR_dismissUpdateNotice(\'' + notice.id + '\')">이 소식 다시 보지 않기</button>' +
    '</body></html>'
  ).setWidth(460).setHeight(430);
  SpreadsheetApp.getUi().showModelessDialog(html, '📢 업데이트 소식');
}

// 시트를 열 때마다: 「🟢 보이기」인 소식이 있으면 보여 줍니다. (그만 보여 주려면 소식 시트에서 상태를 끄십시오.)
function HR_showUpdateNoticeOnOpen_() {
  var notice = HR_fetchNotice_();
  if (!notice) return;
  // 「다시 보지 않기」를 누른 소식은 열 때 띄우지 않습니다. (제목이나 내용이 바뀌면 새 소식으로 다시 뜹니다.)
  if (HR_noticeDismissed_(notice.id)) return;
  HR_showNoticeDialog_(notice);
}

// 안내 창의 「이 소식 다시 보지 않기」 버튼
function HR_dismissUpdateNotice(id) {
  var text = String(id || '');
  if (!/^n[0-9a-z]{1,12}$/.test(text)) return;
  // 사용자 저장소(열 때 실행하는 계정과 다를 수 있음)와 시트 저장소 둘 다에 적어 둡니다.
  var saved = false;
  try { PropertiesService.getUserProperties().setProperty(HR_NOTICE_SEEN_PROPERTY, text); saved = true; } catch (e1) {}
  try { PropertiesService.getDocumentProperties().setProperty(HR_NOTICE_SEEN_PROPERTY, text); saved = true; } catch (e2) {}
  if (!saved) throw new Error('저장하지 못했습니다.');
}

function HR_noticeDismissed_(id) {
  var found = false;
  try { if (PropertiesService.getUserProperties().getProperty(HR_NOTICE_SEEN_PROPERTY) === id) found = true; } catch (e1) {}
  try { if (PropertiesService.getDocumentProperties().getProperty(HR_NOTICE_SEEN_PROPERTY) === id) found = true; } catch (e2) {}
  return found;
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
    .setFontColor('#ffffff').setFontSize(12).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  // 칸마다 색을 달리합니다: [칸 이름 색, 본문 연한 색]
  var COLUMN_COLORS = [
    ['#2f7258', '#eaf4ef'],   // 표시   - 초록
    ['#4a78b5', '#e8f0fa'],   // 종류   - 파랑
    ['#d7a843', '#fdf6e3'],   // 공지 내용 - 금색
    ['#3a8f8f', '#e6f4f4'],   // 시작일 - 청록
    ['#c9694f', '#fbece7'],   // 종료일 - 산호
    ['#7a5aa6', '#f1ebf8']    // 대상   - 보라
  ];
  COLUMN_COLORS.forEach(function(color, i) {
    sheet.getRange(3, i + 1).setBackground(color[0]);
    sheet.getRange(4, i + 1, 997, 1).setBackground(color[1]);
  });

  // 본문: 여섯 칸을 같은 높이·같은 줄무늬·같은 선으로 맞춥니다.
  var body = sheet.getRange('A4:F1000');
  body.setFontSize(11).setFontColor('#2e3a33').setVerticalAlignment('middle');
  // 칸칸마다 또렷한 선: 머리글~본문 전체에 가로·세로 선을 긋고, 표 바깥 테두리는 진하게 합니다.
  sheet.getRange('A3:F1000')
    .setBorder(true, true, true, true, true, true, '#b9cec5', SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange('A3:F1000')
    .setBorder(true, true, true, true, false, false, '#244f42', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sheet.getRange('A3:F3')
    .setBorder(null, null, true, null, true, null, '#244f42', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
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
    "unifiedBarSpec_": unifiedBarSpec_,
    "unifiedBarSpecBase_": unifiedBarSpecBase_,
    "unifiedBarTextWidth_": unifiedBarTextWidth_,
    "unifiedBarFitFontSize_": unifiedBarFitFontSize_,
    "unifiedBarsPropertyKey_": unifiedBarsPropertyKey_,
    "unifiedStoredBarLayout_": unifiedStoredBarLayout_,
    "unifiedComputeBarColumns_": unifiedComputeBarColumns_,
    "unifiedDrawActionBars_": unifiedDrawActionBars_,
    "unifiedResetBarState_": unifiedResetBarState_,
    "unifiedApplyActionBarsAll_": unifiedApplyActionBarsAll_,
    "unifiedBarAt_": unifiedBarAt_,
    "unifiedHandleBarClick_": unifiedHandleBarClick_,
    "unifiedBarPendingKey_": unifiedBarPendingKey_,
    "unifiedClickWithConfirm_": unifiedClickWithConfirm_,
    "unifiedBarConfirmDetail_": unifiedBarConfirmDetail_,
    "unifiedRestoreBarLook_": unifiedRestoreBarLook_,
    "unifiedRevertPendingBars_": unifiedRevertPendingBars_,
    "unifiedRunBarAction_": unifiedRunBarAction_,
    "HR_showSeriesCheckDialog": HR_showSeriesCheckDialog,
    "HR_openSeriesCheckDialog_": HR_openSeriesCheckDialog_,
    "HR_seriesCheckState": HR_seriesCheckState,
    "HR_seriesCheckStart": HR_seriesCheckStart,
    "unifiedBarNeedsCheckbox_": unifiedBarNeedsCheckbox_,
    "unifiedBarHasCheckbox_": unifiedBarHasCheckbox_,
    "unifiedGuideBarCheckbox_": unifiedGuideBarCheckbox_,
    "unifiedHandleBarEdit_": unifiedHandleBarEdit_,
    "unifiedStartSeriesCheckFromBar_": unifiedStartSeriesCheckFromBar_,
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
    "HR_getParentQuota": HR_getParentQuota,
    "HR_getMaxClasses_": HR_getMaxClasses_,
    "unifiedBookFlags_": unifiedBookFlags_,
    "unifiedFlagsFromNote_": unifiedFlagsFromNote_,
    "unifiedFlagTagText_": unifiedFlagTagText_,
    "unifiedComposeNote_": unifiedComposeNote_,
    "unifiedApplyBookFlagsToReview_": unifiedApplyBookFlagsToReview_,
    "unifiedConfirmWorkbooks_": unifiedConfirmWorkbooks_,
    "getConfiguredYes24Key_": getConfiguredYes24Key_,
    "getConfiguredNlkKey": getConfiguredNlkKey,
    "onOpen": onOpen,
    "startNewRequestRound": startNewRequestRound,
    "buildRequestArchiveSheetName_": buildRequestArchiveSheetName_,
    "styleRequestListSheet": styleRequestListSheet,
    "buildBookstoreExportSheet": buildBookstoreExportSheet,
    "unifiedEditTouches_": unifiedEditTouches_,
    "onEdit": onEdit,
    "bookstoreStyleExportSheet_": bookstoreStyleExportSheet_,
    "bookstoreFindColumn_": bookstoreFindColumn_,
    "bookstoreNormalize_": bookstoreNormalize_,
    "installEditTrigger": installEditTrigger,
    "unifiedHandleStructureChange_": unifiedHandleStructureChange_,
    "HR_ensureChangeTrigger_": HR_ensureChangeTrigger_,
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
    "HR_showInquiryDialog": HR_showInquiryDialog,
    "HR_showMyInquiriesDialog": HR_showMyInquiriesDialog,
    "HR_openInquiryDialog_": HR_openInquiryDialog_,
    "HR_readInquiryTickets_": HR_readInquiryTickets_,
    "HR_sendInquiry": HR_sendInquiry,
    "HR_getMyInquiries": HR_getMyInquiries,
    "HR_notifyInquiryAnswersOnOpen_": HR_notifyInquiryAnswersOnOpen_,
    "HR_manualRowHeight_": HR_manualRowHeight_,
    "HR_buildManualSheet_": HR_buildManualSheet_,
    "HR_refreshGuidesOnce_": HR_refreshGuidesOnce_,
    "unifiedEnsurePurchaseLogSheet_": unifiedEnsurePurchaseLogSheet_,
    "unifiedPurchaseLogStats_": unifiedPurchaseLogStats_,
    "unifiedConfirmPurchaseRound": unifiedConfirmPurchaseRound,
    "unifiedConfirmPurchaseRound_": unifiedConfirmPurchaseRound_,
    "unifiedShowPurchaseLog": unifiedShowPurchaseLog,
    "unifiedHidePurchaseLog": unifiedHidePurchaseLog,
    "unifiedHandleDashboardClick_": unifiedHandleDashboardClick_,
    "doGet": doGet,
    "HR_getBuildId": HR_getBuildId,
    "HR_probeWebAppUrl_": HR_probeWebAppUrl_,
    "HR_probeProblemMessage_": HR_probeProblemMessage_,
    "HR_isWebAppExecUrl_": HR_isWebAppExecUrl_,
    "HR_getRegisteredWebAppUrl_": HR_getRegisteredWebAppUrl_,
    "HR_getDetectedWebAppUrl_": HR_getDetectedWebAppUrl_,
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
    "HR_findRequestersColumn_": HR_findRequestersColumn_,
    "HR_ensureRequestersColumn_": HR_ensureRequestersColumn_,
    "HR_normalizeApplicantText_": HR_normalizeApplicantText_,
    "HR_countExtraRequesters_": HR_countExtraRequesters_,
    "HR_recordExtraRequester_": HR_recordExtraRequester_,
    "HR_annotateBooksStatus_": HR_annotateBooksStatus_,
    "HR_composeRequesterNote_": HR_composeRequesterNote_,
    "HR_showShortcutOnOpen_": HR_showShortcutOnOpen_,
    "HR_showHopeRequestDialog": HR_showHopeRequestDialog,
    "HR_seriesStopCheck": HR_seriesStopCheck,
    "seriesStopRequested_": seriesStopRequested_,
    "HR_getStaffCode_": HR_getStaffCode_,
    "HR_isStaffCodeEnabled_": HR_isStaffCodeEnabled_,
    "HR_staffCodeFailCount_": HR_staffCodeFailCount_,
    "HR_checkStaffCode_": HR_checkStaffCode_,
    "HR_requireStaffAccess_": HR_requireStaffAccess_,
    "HR_verifyStaffCode": HR_verifyStaffCode,
    "HR_showStaffCodeDialog": HR_showStaffCodeDialog,
    "HR_arrangeTabsNow_": HR_arrangeTabsNow_,
    "HR_arrangeTabs": HR_arrangeTabs,
    "HR_arrangeTabsOnce_": HR_arrangeTabsOnce_,
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
    "seriesIsQuotaError_": seriesIsQuotaError_,
    "seriesQuotaResumeTime_": seriesQuotaResumeTime_,
    "seriesScheduleQuotaResume_": seriesScheduleQuotaResume_,
    "seriesScheduleErrorRetry_": seriesScheduleErrorRetry_,
    "seriesGetYes24Key_": seriesGetYes24Key_,
    "seriesSpreadsheet_": seriesSpreadsheet_,
    "seriesReadHoldings_": seriesReadHoldings_,
    "seriesGroupHoldings_": seriesGroupHoldings_,
    "seriesEnrichGroupsFromIsbn_": seriesEnrichGroupsFromIsbn_,
    "seriesNameKeysSimilar_": seriesNameKeysSimilar_,
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
    "seriesEnsureAutoContinue_": seriesEnsureAutoContinue_,
    "seriesProgressText_": seriesProgressText_,
    "seriesNowText_": seriesNowText_,
    "seriesWriteStatus_": seriesWriteStatus_,
    "seriesShowStateStatus_": seriesShowStateStatus_,
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
    "unifiedGetBudget_": unifiedGetBudget_,
    "unifiedParseBudget_": unifiedParseBudget_,
    "unifiedSaveBudget_": unifiedSaveBudget_,
    "unifiedBudgetFromDashboard_": unifiedBudgetFromDashboard_,
    "unifiedSetBudget": unifiedSetBudget,
    "unifiedBookKey_": unifiedBookKey_,
    "unifiedMoney_": unifiedMoney_,
    "unifiedDashboardStats_": unifiedDashboardStats_,
    "unifiedFillDashboardNumbers_": unifiedFillDashboardNumbers_,
    "unifiedUpdateDashboardNumbers_": unifiedUpdateDashboardNumbers_,
    "unifiedRefreshDashboard_": unifiedRefreshDashboard_,
    "unifiedAudienceCounts_": unifiedAudienceCounts_,
    "unifiedCountData_": unifiedCountData_,
    "unifiedAuditDataSources": unifiedAuditDataSources,
    "importLegacySeriesSheetsOnce": importLegacySeriesSheetsOnce,
    "unifiedReadExistingSeriesReview_": unifiedReadExistingSeriesReview_,
    "unifiedReadSelectedSeriesItems_": unifiedReadSelectedSeriesItems_,
    "unifiedRefreshPurchaseReview_": unifiedRefreshPurchaseReview_,
    "sendHopeRequestsToReview": sendHopeRequestsToReview,
    "sendSelectedSeriesToReview": sendSelectedSeriesToReview,
    "unifiedEnsureUploadSheet_": unifiedEnsureUploadSheet_,
    "unifiedEnsureUploadSheetReady_": unifiedEnsureUploadSheetReady_,
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
    "unifiedEnsureHoldingsPublisherColumn_": unifiedEnsureHoldingsPublisherColumn_,
    "HR_writeUpdateLogSheet_": HR_writeUpdateLogSheet_,
    "HR_parseCsv_": HR_parseCsv_,
    "HR_noticeIsOn_": HR_noticeIsOn_,
    "HR_updateNoticeDate_": HR_updateNoticeDate_,
    "HR_noticeHash_": HR_noticeHash_,
    "HR_fetchNotice_": HR_fetchNotice_,
    "HR_escapeNoticeHtml_": HR_escapeNoticeHtml_,
    "HR_showNoticeDialog_": HR_showNoticeDialog_,
    "HR_showUpdateNoticeOnOpen_": HR_showUpdateNoticeOnOpen_,
    "HR_dismissUpdateNotice": HR_dismissUpdateNotice,
    "HR_noticeDismissed_": HR_noticeDismissed_,
    "HR_showUpdateNotice": HR_showUpdateNotice,
    "HR_upgradeNoticeDatePickersAndGuide": HR_upgradeNoticeDatePickersAndGuide,
    "HR_applyNoticeDatePickersAndGuide_": HR_applyNoticeDatePickersAndGuide_,
    "HR_styleWebNoticeSheet_": HR_styleWebNoticeSheet_,
    "HR_getActiveWebNotices": HR_getActiveWebNotices,
    "HR_readActiveWebNotices_": HR_readActiveWebNotices_,
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
    files: {"ApplyLinks":"<!DOCTYPE html>\n<html lang=\"ko\">\n<head>\n<base target=\"_top\">\n<meta charset=\"utf-8\">\n<meta name=\"viewport\" content=\"width=device-width, initial-scale=1\">\n<style>\n  :root{--forest:#2c5949;--forest-dark:#203f35;--gold:#d3a248;--gold-dark:#ad762b;--cream:#fbf4e5;--ink:#25352f;--sub:#68766f;--line:#dbe4dc;--paper:#fffdf8;--soft:#e8f2ec}\n  *{box-sizing:border-box}\n  body{margin:0;padding:22px 24px 26px;background:#f4f3eb;font-family:Arial,\"Noto Sans KR\",\"Malgun Gothic\",sans-serif;color:var(--ink)}\n  h1{margin:0 0 6px;font-size:30px;color:var(--forest-dark)}\n  .lead{margin:0 0 18px;color:var(--sub);line-height:1.6;font-size:17px}\n  .tabs{display:grid;grid-template-columns:1fr 1fr;gap:12px;margin-bottom:16px}\n  .tab{padding:16px 10px;border:2px solid var(--line);border-radius:18px;background:#fff;font-size:21px;font-weight:800;color:var(--sub);cursor:pointer}\n  .tab small{display:block;font-size:13px;font-weight:500;margin-top:3px}\n  .tab.active.staff{background:linear-gradient(135deg,#3f8068,#1f4f3f);border-color:#1f4f3f;color:#fff}\n  .tab.active.parent{background:linear-gradient(135deg,var(--gold),var(--gold-dark));border-color:var(--gold-dark);color:#fff}\n  .panel{background:var(--paper);border:1px solid var(--line);border-radius:22px;padding:22px 18px;text-align:center;box-shadow:0 8px 24px rgba(32,63,53,.07)}\n  .panel[hidden]{display:none}\n  .panel .why{margin:0 0 12px;color:var(--sub);font-size:16px;line-height:1.5}\n  canvas.qr{width:300px;height:300px;max-width:100%;border:1px solid var(--line);border-radius:14px;background:#fff;display:block;margin:0 auto 18px}\n  .actions{display:grid;gap:10px;max-width:460px;margin:0 auto}\n  .act{padding:16px 12px;border:0;border-radius:15px;font-size:19px;font-weight:800;cursor:pointer;color:#fff;background:var(--forest)}\n  .act.gold{background:var(--gold-dark)}\n  .act.ghost{background:#fff;color:var(--forest-dark);border:2px solid var(--line);font-size:16px;padding:12px}\n  .act:hover{filter:brightness(.95)}\n  details{margin:16px auto 0;max-width:460px;text-align:left}\n  summary{cursor:pointer;color:var(--sub);font-size:14px;padding:6px 0}\n  .linkbox input{width:100%;padding:10px 12px;border:1px solid var(--line);border-radius:10px;font-size:12px;color:var(--sub);background:#fff}\n  .foot{margin-top:16px;text-align:center}\n  .foot button{padding:12px 36px;border:2px solid var(--line);border-radius:14px;background:#fff;font-size:16px;font-weight:700;color:var(--forest-dark);cursor:pointer}\n  .warn{background:#fff6da;border:1px solid #e3c47a;color:#7a5a12;border-radius:16px;padding:18px 20px;line-height:1.8;font-size:17px}\n  .warn b{color:#5b430c;font-size:19px}\n  .steps{margin:12px 0 0;padding-left:22px}\n  .steps li{margin-bottom:6px}\n  .toast{position:fixed;left:50%;bottom:24px;transform:translateX(-50%);background:var(--forest-dark);color:#fff;padding:12px 22px;border-radius:999px;font-size:16px;opacity:0;transition:opacity .2s;pointer-events:none;z-index:20}\n  .toast.show{opacity:1}\n  .preview{position:fixed;inset:0;background:rgba(32,63,53,.93);display:none;align-items:center;justify-content:center;padding:14px;z-index:10}\n  .preview.show{display:flex}\n  .preview .box{background:#fff;border-radius:20px;padding:18px;max-width:560px;width:100%;max-height:96vh;overflow:auto;text-align:center}\n  .preview h2{margin:0 0 6px;color:var(--forest-dark);font-size:22px}\n  .preview .tip{margin:0 0 12px;color:var(--sub);font-size:15px;line-height:1.5}\n  .preview img{width:100%;max-width:400px;border:1px solid var(--line);border-radius:10px}\n  .preview .actions{margin-top:14px}\n</style>\n</head>\n<body>\n  <h1>🔗 신청 링크 받기</h1>\n  <p class=\"lead\">누구에게 보낼 링크인지 골라서, <b>복사</b>하거나 <b>QR 그림</b>으로 저장하세요.</p>\n\n  <div id=\"needDeploy\" class=\"warn\" hidden>\n    <b>먼저 웹 앱을 한 번 배포해야 링크가 만들어져요.</b>\n    <ol class=\"steps\">\n      <li>위쪽 「확장 프로그램」 → 「Apps Script」를 누르세요.</li>\n      <li>새 화면 오른쪽 위 「배포」 → 「새 배포」 → 톱니바퀴(유형 선택) → 「웹 앱」을 고르세요.</li>\n      <li>실행 사용자는 「나」, 액세스 권한은 「모든 사용자」로 하고 「배포」를 누르세요.</li>\n      <li>끝나면 이 창을 닫고, 「⚙️ 사본 설정 → 🔗 신청 링크 받기」를 다시 누르세요.</li>\n    </ol>\n    <p style=\"margin:12px 0 0;font-size:15px\">이미 배포했는데도 이 안내가 나오면, 「⚙️ 사본 설정 → 🔗 공개 신청 링크 등록」에 /exec 주소를 붙여 넣은 뒤 다시 눌러 보세요.</p>\n  </div>\n\n  <div id=\"problemNote\" class=\"warn\" hidden style=\"margin-bottom:14px\">\n    <b>⚠️ 지금 등록된 주소로는 접속이 안 돼요</b>\n    <p id=\"problemText\" style=\"margin:10px 0 0;font-size:16px;line-height:1.7\"></p>\n    <p style=\"margin:12px 0 0;font-size:14px;line-height:1.6\">현재 등록된 주소:<br><span id=\"badUrl\" style=\"word-break:break-all;color:#5b430c\"></span></p>\n    <ol class=\"steps\">\n      <li>위쪽 「확장 프로그램」 → 「Apps Script」 → 「배포」 → 「배포 관리」를 여세요.</li>\n      <li>웹 앱의 주소(…/exec)를 <b>처음부터 끝까지</b> 복사하세요.</li>\n      <li>시트에서 「⚙️ 사본 설정 → 🔗 공개 신청 링크 등록」에 붙여 넣고 저장하세요.</li>\n      <li>다시 「🔗 신청 링크 받기」를 누르면 QR이 만들어져요.</li>\n    </ol>\n  </div>\n\n  <div id=\"switchedNote\" class=\"warn\" hidden style=\"margin-bottom:14px;font-size:15px;line-height:1.6;background:#eaf3ed;border-color:#a8c8bb;color:#203f35\">\n    ✅ 등록돼 있던 주소가 열리지 않아서, 지금 배포된 웹 앱 주소로 자동으로 바꿨어요.\n  </div>\n\n  <div id=\"mismatchNote\" class=\"warn\" hidden style=\"margin-bottom:14px;font-size:15px;line-height:1.6\">\n    ⚠️ 「공개 신청 링크 등록」에 저장된 주소를 쓰고 있어요. 이 주소가 지금 배포한 웹 앱과 다르면 링크가 안 열릴 수 있어요.\n    그럴 때는 Apps Script 「배포 → 배포 관리」에서 현재 웹 앱 주소를 복사해 「공개 신청 링크 등록」에 다시 붙여 넣어 주세요.\n  </div>\n\n  <div id=\"main\" hidden>\n    <div class=\"tabs\" role=\"tablist\">\n      <button type=\"button\" class=\"tab parent active\" data-audience=\"parent\" role=\"tab\">👨‍👩‍👧 학부모용<small>가정통신문·알림장</small></button>\n      <button type=\"button\" class=\"tab staff\" data-audience=\"staff\" role=\"tab\">🧑‍🏫 교직원용<small>선생님께 안내</small></button>\n    </div>\n\n    <section class=\"panel\" data-audience=\"parent\" role=\"tabpanel\">\n      <p class=\"why\">학부모님께 보내는 링크예요. 이 링크로 들어가면 <b>교직원 화면은 보이지 않아요.</b></p>\n      <canvas class=\"qr\" width=\"600\" height=\"600\"></canvas>\n      <div class=\"actions\">\n        <button type=\"button\" class=\"act\" data-act=\"copy\">📋 링크 복사</button>\n        <button type=\"button\" class=\"act\" data-act=\"message\">💬 안내 문구까지 복사 (문자·알림장용)</button>\n        <button type=\"button\" class=\"act gold\" data-act=\"save\">⬇️ QR 그림 저장 (인쇄용)</button>\n        <button type=\"button\" class=\"act ghost\" data-act=\"open\">🔎 링크 열어 보기</button>\n      </div>\n      <details><summary>링크 전체 보기</summary><div class=\"linkbox\"><input type=\"text\" readonly></div></details>\n    </section>\n\n    <section class=\"panel\" data-audience=\"staff\" role=\"tabpanel\" hidden>\n      <p class=\"why\">선생님·교직원께 보내는 링크예요.</p>\n      <canvas class=\"qr\" width=\"600\" height=\"600\"></canvas>\n      <div class=\"actions\">\n        <button type=\"button\" class=\"act\" data-act=\"copy\">📋 링크 복사</button>\n        <button type=\"button\" class=\"act\" data-act=\"message\">💬 안내 문구까지 복사 (문자·알림장용)</button>\n        <button type=\"button\" class=\"act gold\" data-act=\"save\">⬇️ QR 그림 저장 (인쇄용)</button>\n        <button type=\"button\" class=\"act ghost\" data-act=\"open\">🔎 링크 열어 보기</button>\n      </div>\n      <details><summary>링크 전체 보기</summary><div class=\"linkbox\"><input type=\"text\" readonly></div></details>\n    </section>\n  </div>\n\n  <div class=\"foot\"><button type=\"button\" onclick=\"google.script.host.close()\">닫기</button></div>\n\n  <div id=\"preview\" class=\"preview\" role=\"dialog\" aria-modal=\"true\">\n    <div class=\"box\">\n      <h2>QR 그림이 만들어졌어요 ✅</h2>\n      <p class=\"tip\">저장이 시작되지 않으면 아래 그림을 <b>오른쪽 클릭(휴대폰은 길게 누르기)</b> → 「이미지를 다른 이름으로 저장」을 누르세요.</p>\n      <img id=\"previewImg\" alt=\"QR 그림\">\n      <div class=\"actions\">\n        <a id=\"downloadLink\" class=\"act\" style=\"text-decoration:none;display:block\" href=\"#\" download>⬇️ 다시 저장하기</a>\n        <button type=\"button\" class=\"act ghost\" id=\"closePreview\">닫기</button>\n      </div>\n    </div>\n  </div>\n  <div id=\"toast\" class=\"toast\"></div>\n\n<script>\n/* QR-BEGIN */\n// 외부 서비스 없이 QR 코드를 만듭니다. (바이트 모드 · 오류정정 M · 버전 1~10, 최대 213바이트)\n// 알고리즘은 QR 코드 표준(ISO/IEC 18004)을 따릅니다.\nvar QrCore = (function () {\n  'use strict';\n\n  // 버전별: [전체 코드워드, 블록 수, 블록당 오류정정 코드워드]  (오류정정 M)\n  var TABLE = {\n    1: [26, 1, 10], 2: [44, 1, 16], 3: [70, 1, 26], 4: [100, 2, 18], 5: [134, 2, 24],\n    6: [172, 4, 16], 7: [196, 4, 18], 8: [242, 4, 22], 9: [292, 5, 22], 10: [346, 5, 26]\n  };\n  var MAX_VERSION = 10;\n\n  // --- 갈루아체 GF(256) ---\n  var EXP = [], LOG = [];\n  (function () {\n    var x = 1, i;\n    for (i = 0; i < 255; i++) { EXP[i] = x; LOG[x] = i; x <<= 1; if (x & 0x100) x ^= 0x11d; }\n    for (i = 255; i < 512; i++) EXP[i] = EXP[i - 255];\n  })();\n  function gfMul(a, b) { return (a === 0 || b === 0) ? 0 : EXP[LOG[a] + LOG[b]]; }\n\n  function rsGenerator(degree) {\n    var poly = [1], i, j, next;\n    for (i = 0; i < degree; i++) {\n      next = [];\n      for (j = 0; j <= poly.length; j++) next[j] = 0;\n      for (j = 0; j < poly.length; j++) {\n        next[j] ^= poly[j];\n        next[j + 1] ^= gfMul(poly[j], EXP[i]);\n      }\n      poly = next;\n    }\n    return poly; // 최고차항 계수 1 포함, 길이 degree+1\n  }\n  function rsRemainder(data, gen) {\n    var degree = gen.length - 1, rem = [], i, j, factor;\n    for (i = 0; i < degree; i++) rem[i] = 0;\n    for (i = 0; i < data.length; i++) {\n      factor = data[i] ^ rem[0];\n      rem.shift(); rem.push(0);\n      for (j = 0; j < degree; j++) rem[j] ^= gfMul(gen[j + 1], factor);\n    }\n    return rem;\n  }\n\n  function utf8Bytes(text) {\n    var out = [], i, c, d;\n    var s = unescape(encodeURIComponent(String(text)));\n    for (i = 0; i < s.length; i++) out.push(s.charCodeAt(i));\n    return out;\n  }\n\n  function numAlign(ver) { return ver === 1 ? 0 : Math.floor(ver / 7) + 2; }\n  function rawModules(ver) {\n    var n = (16 * ver + 128) * ver + 64;\n    if (ver >= 2) { var a = numAlign(ver); n -= (25 * a - 10) * a - 55; if (ver >= 7) n -= 36; }\n    return n;\n  }\n  function alignPositions(ver) {\n    if (ver === 1) return [];\n    var n = numAlign(ver), size = ver * 4 + 17;\n    var step = Math.ceil((ver * 4 + 4) / (n * 2 - 2)) * 2;\n    var res = [6], pos;\n    for (pos = size - 7; res.length < n; pos -= step) res.splice(1, 0, pos);\n    return res;\n  }\n\n  function bit(v, i) { return ((v >>> i) & 1) !== 0; }\n\n  function pickVersion(byteLen) {\n    var ver, t, dataCw, capBits, need;\n    for (ver = 1; ver <= MAX_VERSION; ver++) {\n      t = TABLE[ver];\n      dataCw = t[0] - t[1] * t[2];\n      capBits = dataCw * 8;\n      need = 4 + (ver <= 9 ? 8 : 16) + byteLen * 8;\n      if (need <= capBits) return ver;\n    }\n    throw new Error('주소가 너무 길어서 QR을 만들 수 없습니다.');\n  }\n\n  function encode(text) {\n    var bytes = utf8Bytes(text);\n    var ver = pickVersion(bytes.length);\n    var t = TABLE[ver], raw = Math.floor(rawModules(ver) / 8);\n    var numBlocks = t[1], eccLen = t[2], dataCw = raw - numBlocks * eccLen;\n    var size = ver * 4 + 17;\n    var i, j, k;\n\n    // 1) 데이터 비트열\n    var bits = [];\n    function push(val, len) { for (var b = len - 1; b >= 0; b--) bits.push((val >>> b) & 1); }\n    push(4, 4);                               // 바이트 모드\n    push(bytes.length, ver <= 9 ? 8 : 16);    // 글자 수\n    for (i = 0; i < bytes.length; i++) push(bytes[i], 8);\n    var capBits = dataCw * 8;\n    for (i = 0; i < 4 && bits.length < capBits; i++) bits.push(0); // 종료 문자\n    while (bits.length % 8 !== 0) bits.push(0);\n    var data = [];\n    for (i = 0; i < bits.length; i += 8) {\n      var v = 0; for (j = 0; j < 8; j++) v = (v << 1) | bits[i + j]; data.push(v);\n    }\n    for (k = 0; data.length < dataCw; k++) data.push(k % 2 === 0 ? 0xEC : 0x11);\n\n    // 2) 블록 나누기 + 오류정정 + 끼워 넣기\n    var shortLen = Math.floor(raw / numBlocks), numShort = numBlocks - raw % numBlocks;\n    var gen = rsGenerator(eccLen), blocks = [], offset = 0;\n    for (i = 0; i < numBlocks; i++) {\n      var len = shortLen - eccLen + (i < numShort ? 0 : 1);\n      var dat = data.slice(offset, offset + len); offset += len;\n      var ecc = rsRemainder(dat, gen);\n      if (i < numShort) dat.push(0);          // 짧은 블록은 자리 맞춤용 빈칸\n      blocks.push(dat.concat(ecc));\n    }\n    var all = [];\n    for (i = 0; i < blocks[0].length; i++) {\n      for (j = 0; j < blocks.length; j++) {\n        if (i !== shortLen - eccLen || j >= numShort) all.push(blocks[j][i]);\n      }\n    }\n\n    // 3) 모듈 배치\n    var modules = [], isFn = [];\n    for (i = 0; i < size; i++) { modules.push([]); isFn.push([]); for (j = 0; j < size; j++) { modules[i].push(false); isFn[i].push(false); } }\n    function setFn(x, y, dark) { modules[y][x] = !!dark; isFn[y][x] = true; }\n\n    for (i = 0; i < size; i++) { setFn(6, i, i % 2 === 0); setFn(i, 6, i % 2 === 0); }\n    function finder(cx, cy) {\n      for (var dy = -4; dy <= 4; dy++) for (var dx = -4; dx <= 4; dx++) {\n        var dist = Math.max(Math.abs(dx), Math.abs(dy)), xx = cx + dx, yy = cy + dy;\n        if (xx >= 0 && xx < size && yy >= 0 && yy < size) setFn(xx, yy, dist !== 2 && dist !== 4);\n      }\n    }\n    finder(3, 3); finder(size - 4, 3); finder(3, size - 4);\n    var ap = alignPositions(ver), n = ap.length;\n    for (i = 0; i < n; i++) for (j = 0; j < n; j++) {\n      if ((i === 0 && j === 0) || (i === 0 && j === n - 1) || (i === n - 1 && j === 0)) continue;\n      for (var dy = -2; dy <= 2; dy++) for (var dx = -2; dx <= 2; dx++)\n        setFn(ap[i] + dx, ap[j] + dy, Math.max(Math.abs(dx), Math.abs(dy)) !== 1);\n    }\n    function drawFormat(mask) {\n      var d = (0 << 3) | mask;                // 오류정정 M = 00\n      var rem = d;\n      for (var q = 0; q < 10; q++) rem = (rem << 1) ^ ((rem >>> 9) * 0x537);\n      var fb = ((d << 10) | rem) ^ 0x5412;\n      for (var a = 0; a <= 5; a++) setFn(8, a, bit(fb, a));\n      setFn(8, 7, bit(fb, 6)); setFn(8, 8, bit(fb, 7)); setFn(7, 8, bit(fb, 8));\n      for (var b = 9; b < 15; b++) setFn(14 - b, 8, bit(fb, b));\n      for (var c = 0; c < 8; c++) setFn(size - 1 - c, 8, bit(fb, c));\n      for (var e = 8; e < 15; e++) setFn(8, size - 15 + e, bit(fb, e));\n      setFn(8, size - 8, true);\n    }\n    drawFormat(0);\n    if (ver >= 7) {\n      var r = ver;\n      for (i = 0; i < 12; i++) r = (r << 1) ^ ((r >>> 11) * 0x1F25);\n      var vb = (ver << 12) | r;\n      for (i = 0; i < 18; i++) {\n        var col = bit(vb, i), a2 = size - 11 + (i % 3), b2 = Math.floor(i / 3);\n        setFn(a2, b2, col); setFn(b2, a2, col);\n      }\n    }\n\n    // 데이터 비트 지그재그 배치\n    var idx = 0;\n    for (var right = size - 1; right >= 1; right -= 2) {\n      if (right === 6) right = 5;\n      for (var vert = 0; vert < size; vert++) {\n        for (j = 0; j < 2; j++) {\n          var x = right - j, upward = ((right + 1) & 2) === 0, y = upward ? size - 1 - vert : vert;\n          if (!isFn[y][x] && idx < all.length * 8) { modules[y][x] = bit(all[idx >>> 3], 7 - (idx & 7)); idx++; }\n        }\n      }\n    }\n\n    // 4) 마스크 8가지 중 벌점이 가장 낮은 것 고르기\n    function applyMask(mask) {\n      for (var y = 0; y < size; y++) for (var x = 0; x < size; x++) {\n        var inv;\n        switch (mask) {\n          case 0: inv = (x + y) % 2 === 0; break;\n          case 1: inv = y % 2 === 0; break;\n          case 2: inv = x % 3 === 0; break;\n          case 3: inv = (x + y) % 3 === 0; break;\n          case 4: inv = (Math.floor(x / 3) + Math.floor(y / 2)) % 2 === 0; break;\n          case 5: inv = x * y % 2 + x * y % 3 === 0; break;\n          case 6: inv = (x * y % 2 + x * y % 3) % 2 === 0; break;\n          default: inv = ((x + y) % 2 + x * y % 3) % 2 === 0;\n        }\n        if (!isFn[y][x] && inv) modules[y][x] = !modules[y][x];\n      }\n    }\n    function penalty() {\n      var score = 0, y, x, run, c, dark = 0;\n      for (y = 0; y < size; y++) {\n        run = 1;\n        for (x = 1; x < size; x++) {\n          if (modules[y][x] === modules[y][x - 1]) { run++; if (run === 5) score += 3; else if (run > 5) score++; } else run = 1;\n        }\n      }\n      for (x = 0; x < size; x++) {\n        run = 1;\n        for (y = 1; y < size; y++) {\n          if (modules[y][x] === modules[y - 1][x]) { run++; if (run === 5) score += 3; else if (run > 5) score++; } else run = 1;\n        }\n      }\n      for (y = 0; y < size - 1; y++) for (x = 0; x < size - 1; x++) {\n        c = modules[y][x];\n        if (c === modules[y][x + 1] && c === modules[y + 1][x] && c === modules[y + 1][x + 1]) score += 3;\n      }\n      var pat = [true, false, true, true, true, false, true, false, false, false, false];\n      function matches(get) {\n        var k2, ok1 = true, ok2 = true;\n        for (k2 = 0; k2 < 11; k2++) { if (get(k2) !== pat[k2]) ok1 = false; if (get(k2) !== pat[10 - k2]) ok2 = false; }\n        return ok1 || ok2;\n      }\n      for (y = 0; y < size; y++) for (x = 0; x <= size - 11; x++) { if (matches(function (k2) { return modules[y][x + k2]; })) score += 40; }\n      for (x = 0; x < size; x++) for (y = 0; y <= size - 11; y++) { if (matches(function (k2) { return modules[y + k2][x]; })) score += 40; }\n      for (y = 0; y < size; y++) for (x = 0; x < size; x++) if (modules[y][x]) dark++;\n      score += Math.floor(Math.abs(dark * 20 - size * size * 10) / (size * size)) * 10;\n      return score;\n    }\n    var bestMask = 0, bestScore = Infinity, m;\n    for (m = 0; m < 8; m++) {\n      applyMask(m); drawFormat(m);\n      var sc = penalty();\n      if (sc < bestScore) { bestScore = sc; bestMask = m; }\n      applyMask(m); // 되돌리기\n    }\n    applyMask(bestMask); drawFormat(bestMask);\n\n    return { version: ver, size: size, modules: modules };\n  }\n\n  return { encode: encode };\n})();\n/* QR-END */\n\n(function () {\n  var DATA = <?!= dataJson ?>;\n  var LABEL = { staff: '교직원용', parent: '학부모용' };\n  var COLOR = { staff: '#203f35', parent: '#ad762b' };\n\n  function el(id) { return document.getElementById(id); }\n  function toast(text) {\n    var t = el('toast'); t.textContent = text; t.classList.add('show');\n    setTimeout(function () { t.classList.remove('show'); }, 2000);\n  }\n  function linkFor(audience) {\n    var base = DATA.baseUrl;\n    return base + (base.indexOf('?') < 0 ? '?' : '&') + 'type=' + audience;\n  }\n\n  // QR 을 캔버스에 그립니다. (바깥 여백 4칸)\n  function drawQr(canvas, text, size) {\n    var qr = QrCore.encode(text), quiet = 4, n = qr.size;\n    var cell = Math.floor(size / (n + quiet * 2));\n    var total = cell * (n + quiet * 2);\n    var ctx = canvas.getContext('2d');\n    ctx.fillStyle = '#fff'; ctx.fillRect(0, 0, size, size);\n    var offset = Math.floor((size - total) / 2) + quiet * cell;\n    ctx.fillStyle = '#000';\n    for (var y = 0; y < n; y++) for (var x = 0; x < n; x++) {\n      if (qr.modules[y][x]) ctx.fillRect(offset + x * cell, offset + y * cell, cell, cell);\n    }\n  }\n\n  function wrapText(ctx, text, maxWidth) {\n    var lines = [], line = '';\n    for (var i = 0; i < text.length; i++) {\n      var test = line + text.charAt(i);\n      if (ctx.measureText(test).width > maxWidth && line) { lines.push(line); line = text.charAt(i); } else { line = test; }\n    }\n    if (line) lines.push(line);\n    return lines;\n  }\n\n  // 인쇄해서 쓸 수 있는 그림(제목 + QR + 안내 + 주소)을 만듭니다.\n  function buildPoster(audience) {\n    var url = linkFor(audience);\n    var W = 1000, H = 1240, c = document.createElement('canvas');\n    c.width = W; c.height = H;\n    var ctx = c.getContext('2d');\n    ctx.fillStyle = '#fff'; ctx.fillRect(0, 0, W, H);\n    ctx.fillStyle = COLOR[audience]; ctx.fillRect(0, 0, W, 18);\n    ctx.textAlign = 'center'; ctx.textBaseline = 'alphabetic';\n\n    var school = DATA.schoolName ? DATA.schoolName : '';\n    ctx.fillStyle = '#25352f';\n    var y = 120;\n    if (school) { ctx.font = '700 44px \"Noto Sans KR\",\"Malgun Gothic\",Arial,sans-serif'; ctx.fillText(school, W / 2, y); y += 78; }\n    ctx.fillStyle = '#203f35';\n    ctx.font = '900 76px \"Noto Sans KR\",\"Malgun Gothic\",Arial,sans-serif';\n    ctx.fillText('희망도서 신청', W / 2, y);\n\n    y += 40;\n    var badge = LABEL[audience] + ' ' + (audience === 'parent' ? '👨‍👩‍👧' : '🧑‍🏫');\n    ctx.font = '800 42px \"Noto Sans KR\",\"Malgun Gothic\",Arial,sans-serif';\n    var bw = ctx.measureText(badge).width + 70;\n    ctx.fillStyle = COLOR[audience];\n    var bx = (W - bw) / 2, by = y, bh = 70, r = 35;\n    ctx.beginPath(); ctx.moveTo(bx + r, by); ctx.arcTo(bx + bw, by, bx + bw, by + bh, r); ctx.arcTo(bx + bw, by + bh, bx, by + bh, r); ctx.arcTo(bx, by + bh, bx, by, r); ctx.arcTo(bx, by, bx + bw, by, r); ctx.closePath(); ctx.fill();\n    ctx.fillStyle = '#fff'; ctx.fillText(badge, W / 2, by + 50);\n\n    var qrSize = 720, qc = document.createElement('canvas');\n    qc.width = qrSize; qc.height = qrSize;\n    drawQr(qc, url, qrSize);\n    var qy = by + bh + 40;\n    ctx.fillStyle = '#fff'; ctx.strokeStyle = '#dbe4dc'; ctx.lineWidth = 3;\n    ctx.fillRect((W - qrSize) / 2 - 6, qy - 6, qrSize + 12, qrSize + 12);\n    ctx.strokeRect((W - qrSize) / 2 - 6, qy - 6, qrSize + 12, qrSize + 12);\n    ctx.drawImage(qc, (W - qrSize) / 2, qy);\n\n    var ty = qy + qrSize + 70;\n    ctx.fillStyle = '#25352f';\n    ctx.font = '700 38px \"Noto Sans KR\",\"Malgun Gothic\",Arial,sans-serif';\n    ctx.fillText('스마트폰 카메라로 QR을 비춰 주세요', W / 2, ty);\n\n    ctx.fillStyle = '#68766f';\n    ctx.font = '400 21px Arial,sans-serif';\n    var lines = wrapText(ctx, url, W - 160), ly = ty + 52;\n    for (var i = 0; i < lines.length && i < 3; i++) { ctx.fillText(lines[i], W / 2, ly); ly += 28; }\n    return c;\n  }\n\n  function fileName(audience) {\n    var school = (DATA.schoolName || '').replace(/[\\\\\\/:*?\"<>|]/g, '').trim();\n    return (school ? school + '_' : '') + (audience === 'parent' ? '학부모용' : '교직원용') + '_희망도서_QR.png';\n  }\n\n  function savePoster(audience) {\n    var canvas = buildPoster(audience), dataUrl = canvas.toDataURL('image/png');\n    var name = fileName(audience);\n    el('previewImg').src = dataUrl;\n    var link = el('downloadLink'); link.href = dataUrl; link.setAttribute('download', name);\n    el('preview').classList.add('show');\n    try { link.click(); toast('저장을 시작했어요'); } catch (e) {}\n  }\n\n\n  function fileName(audience) {\n    var school = (DATA.schoolName || '').replace(/[\\\\\\/:*?\"<>|]/g, '').trim();\n    return (school ? school + '_' : '') + (audience === 'parent' ? '학부모용' : '교직원용') + '_희망도서_QR.png';\n  }\n\n  function savePoster(audience) {\n    var canvas = buildPoster(audience), dataUrl = canvas.toDataURL('image/png');\n    el('previewImg').src = dataUrl;\n    var link = el('downloadLink'); link.href = dataUrl; link.setAttribute('download', fileName(audience));\n    el('preview').classList.add('show');\n    try { link.click(); toast('저장을 시작했어요'); } catch (e) {}\n  }\n\n  function copyValue(text, okMessage, inputForFallback) {\n    function fallback() {\n      var ta = inputForFallback || document.createElement('textarea');\n      if (!inputForFallback) { ta.value = text; ta.style.position = 'fixed'; ta.style.opacity = '0'; document.body.appendChild(ta); }\n      ta.focus(); ta.select();\n      var ok = false;\n      try { ok = document.execCommand('copy'); } catch (e) {}\n      if (!inputForFallback) document.body.removeChild(ta);\n      toast(ok ? okMessage : '복사가 안 돼요. 「링크 전체 보기」에서 길게 눌러 복사해 주세요.');\n    }\n    if (navigator.clipboard && navigator.clipboard.writeText) {\n      navigator.clipboard.writeText(text).then(function () { toast(okMessage); }, fallback);\n    } else {\n      fallback();\n    }\n  }\n\n  function messageFor(audience) {\n    var school = DATA.schoolName ? DATA.schoolName + ' ' : '';\n    return school + '희망도서 신청 안내\\n' +\n      (audience === 'parent' ? '읽고 싶은 책을 아래 링크에서 신청해 주세요.\\n' : '읽고 싶은 책을 아래 링크에서 신청해 주세요.\\n') +\n      linkFor(audience);\n  }\n\n  if (DATA.problem) {\n    el('problemNote').hidden = false;\n    el('problemText').textContent = DATA.problem;\n    el('badUrl').textContent = DATA.badUrl || '';\n  } else if (!DATA.baseUrl) {\n    el('needDeploy').hidden = false;\n  } else {\n    if (DATA.switched) el('switchedNote').hidden = false;\n    el('main').hidden = false;\n    if (DATA.mismatch) el('mismatchNote').hidden = false;\n    var panels = {};\n    Array.prototype.forEach.call(document.querySelectorAll('.panel'), function (panel) {\n      var audience = panel.getAttribute('data-audience'), input = panel.querySelector('input');\n      panels[audience] = panel;\n      input.value = linkFor(audience);\n      try { drawQr(panel.querySelector('canvas'), input.value, 600); }\n      catch (e) { panel.querySelector('canvas').insertAdjacentHTML('afterend', '<p class=\"why\">QR을 만들지 못했어요: ' + String(e.message || e) + '</p>'); }\n      panel.addEventListener('click', function (event) {\n        var act = event.target && event.target.getAttribute && event.target.getAttribute('data-act');\n        if (act === 'copy') copyValue(input.value, '링크를 복사했어요 ✅', input);\n        else if (act === 'message') copyValue(messageFor(audience), '안내 문구를 복사했어요 ✅');\n        else if (act === 'open') window.open(input.value, '_blank');\n        else if (act === 'save') savePoster(audience);\n      });\n    });\n    Array.prototype.forEach.call(document.querySelectorAll('.tab'), function (tab) {\n      tab.addEventListener('click', function () {\n        var audience = tab.getAttribute('data-audience');\n        Array.prototype.forEach.call(document.querySelectorAll('.tab'), function (t) { t.classList.toggle('active', t === tab); });\n        Object.keys(panels).forEach(function (key) { panels[key].hidden = key !== audience; });\n      });\n    });\n  }\n  el('closePreview').addEventListener('click', function () { el('preview').classList.remove('show'); });\n})();\n</script>\n</body>\n</html>\n","CopySetup":"<!DOCTYPE html>\n<html lang=\"ko\">\n<head>\n  <base target=\"_top\">\n  <meta charset=\"UTF-8\">\n\n  <style>\n    :root {\n      --navy: #243b64;\n      --blue: #5277c7;\n      --pale: #f3f6fc;\n      --line: #d8e1f0;\n      --text: #27344a;\n      --muted: #69758a;\n      --success: #28734d;\n      --danger: #a33c3c;\n      --warning: #9a6418;\n    }\n\n    * {\n      box-sizing: border-box;\n    }\n\n    body {\n      margin: 0;\n      padding: 24px;\n      background: #f6f8fc;\n      color: var(--text);\n      font-family:\n        -apple-system,\n        BlinkMacSystemFont,\n        \"Apple SD Gothic Neo\",\n        \"Noto Sans KR\",\n        sans-serif;\n      font-size: 16px;\n    }\n\n    .header {\n      margin-bottom: 18px;\n    }\n\n    h1 {\n      margin: 0;\n      color: var(--navy);\n      font-size: 25px;\n      letter-spacing: -0.04em;\n    }\n\n    .subtitle {\n      margin: 8px 0 0;\n      color: var(--muted);\n      line-height: 1.6;\n    }\n\n    .card {\n      margin-bottom: 16px;\n      padding: 20px;\n      border: 1px solid var(--line);\n      border-radius: 16px;\n      background: #fff;\n      box-shadow: 0 8px 24px rgba(36, 59, 100, 0.06);\n    }\n\n    label {\n      display: block;\n      margin-bottom: 8px;\n      color: var(--navy);\n      font-size: 16px;\n      font-weight: 800;\n    }\n\n    input {\n      width: 100%;\n      min-height: 50px;\n      padding: 0 14px;\n      border: 1px solid var(--line);\n      border-radius: 11px;\n      outline: none;\n      color: var(--text);\n      font: inherit;\n    }\n\n    input:focus {\n      border-color: var(--blue);\n      box-shadow: 0 0 0 4px rgba(82, 119, 199, 0.12);\n    }\n\n    .help {\n      margin: 9px 0 0;\n      color: var(--muted);\n      font-size: 13px;\n      line-height: 1.55;\n    }\n\n    .api-divider {\n      height: 1px;\n      margin: 20px 0;\n      background: var(--line);\n    }\n\n    .required,\n    .optional {\n      display: inline-flex;\n      margin-left: 7px;\n      padding: 3px 8px;\n      border-radius: 999px;\n      font-size: 12px;\n      vertical-align: 2px;\n    }\n\n    .required {\n      background: #fff0f0;\n      color: var(--danger);\n    }\n\n    .optional {\n      background: #edf4ff;\n      color: #315f9f;\n    }\n\n    .status-list {\n      display: grid;\n      gap: 9px;\n      margin: 0;\n      padding: 0;\n      list-style: none;\n    }\n\n    .status-item {\n      display: flex;\n      align-items: center;\n      gap: 9px;\n      min-height: 42px;\n      padding: 9px 12px;\n      border-radius: 10px;\n      background: var(--pale);\n      font-size: 14px;\n      font-weight: 700;\n    }\n\n    .status-item.ok {\n      background: #edf8f1;\n      color: var(--success);\n    }\n\n    .status-item.need {\n      background: #fff6e8;\n      color: var(--warning);\n    }\n\n    .notice {\n      margin-top: 14px;\n      padding: 13px 14px;\n      border: 1px solid #edc5c5;\n      border-radius: 11px;\n      background: #fff1f1;\n      color: var(--danger);\n      font-size: 14px;\n      font-weight: 800;\n      line-height: 1.6;\n    }\n\n    .message {\n      display: none;\n      margin: 14px 0;\n      padding: 12px 14px;\n      border-radius: 10px;\n      font-size: 14px;\n      line-height: 1.55;\n    }\n\n    .message.show {\n      display: block;\n    }\n\n    .message.success {\n      background: #edf8f1;\n      color: var(--success);\n    }\n\n    .message.error {\n      background: #fff1f1;\n      color: var(--danger);\n    }\n\n    .buttons {\n      display: flex;\n      flex-wrap: wrap;\n      justify-content: flex-end;\n      gap: 9px;\n      margin-top: 18px;\n    }\n\n    button {\n      min-height: 46px;\n      padding: 0 18px;\n      border-radius: 10px;\n      font: inherit;\n      font-weight: 800;\n      cursor: pointer;\n    }\n\n    button:disabled {\n      cursor: wait;\n      opacity: 0.55;\n    }\n\n    .primary {\n      border: 0;\n      background: var(--blue);\n      color: #fff;\n    }\n\n    .secondary {\n      border: 1px solid var(--line);\n      background: #fff;\n      color: var(--navy);\n    }\n\n    .danger {\n      border: 1px solid #efc7c7;\n      background: #fff5f5;\n      color: var(--danger);\n    }\n\n    .hidden {\n      display: none;\n    }\n    .contact-divider {\n      height: 1px;\n      margin: 22px 0;\n      background: var(--line);\n    }\n\n    .contact-grid {\n      display: grid;\n      grid-template-columns: 1fr 1fr;\n      gap: 14px;\n    }\n\n    .contact-grid label {\n      display: block;\n      margin-bottom: 8px;\n    }\n\n    .contact-help {\n      margin-top: 12px;\n      color: #667085;\n      font-size: 13px;\n      line-height: 1.6;\n    }\n\n    @media (max-width: 560px) {\n      .contact-grid {\n        grid-template-columns: 1fr;\n      }\n    }\n  </style>\n</head>\n\n<body>\n  <header class=\"header\">\n    <h1 id=\"pageTitle\"></h1>\n    <p id=\"pageSubtitle\" class=\"subtitle\"></p>\n  </header>\n\n  <section id=\"apiPanel\" class=\"card hidden\">\n    <label for=\"yes24Key\">\n      YES24 API Key\n      <span class=\"required\">필수</span>\n    </label>\n    <input\n      id=\"yes24Key\"\n      type=\"text\"\n      autocomplete=\"off\"\n      placeholder=\"발급받은 YES24 API Key\"\n    >\n    <p class=\"help\">\n      저장할 때 실제 도서 검색을 한 번 실행해 키가 정상인지 확인합니다.\n    </p>\n\n  </section>\n\n  <section id=\"schoolPanel\" class=\"hidden\">\n    <div class=\"card\">\n      <label for=\"schoolName\">학교명</label>\n      <input\n        id=\"schoolName\"\n        type=\"text\"\n        autocomplete=\"off\"\n        placeholder=\"예: ***초등학교\"\n      >\n      <p id=\"spreadsheetName\" class=\"help\"></p>\n\n      <div class=\"contact-divider\"></div>\n\n      <div class=\"contact-grid\">\n        <div>\n          <label for=\"managerName\">도서관 담당자 이름</label>\n          <input\n            id=\"managerName\"\n            type=\"text\"\n            autocomplete=\"off\"\n            placeholder=\"예: 홍길동\"\n          >\n        </div>\n\n        <div>\n          <label for=\"managerExtension\">내선번호</label>\n          <input\n            id=\"managerExtension\"\n            type=\"text\"\n            inputmode=\"tel\"\n            autocomplete=\"off\"\n            placeholder=\"예: 1234\"\n          >\n        </div>\n      </div>\n\n      <p class=\"contact-help\">\n        입력한 담당자 이름과 내선번호는 공개 희망도서 신청 페이지 맨 아래에 표시됩니다.\n        필요하지 않으면 비워 두어도 됩니다.\n      </p>\n\n      <div class=\"contact-divider\"></div>\n\n      <label for=\"maxClasses\">한 학년 최대 반 수</label>\n      <input\n        id=\"maxClasses\"\n        type=\"number\"\n        min=\"1\"\n        max=\"30\"\n        inputmode=\"numeric\"\n        autocomplete=\"off\"\n        placeholder=\"예: 8\"\n      >\n      <p class=\"contact-help\">\n        학부모가 신청할 때 「반」 선택칸에 이 숫자까지만 보입니다. 비워 두면 20반까지 보입니다.\n      </p>\n    </div>\n\n    <div class=\"card\">\n      <label>설정 점검</label>\n      <ul id=\"statusList\" class=\"status-list\"></ul>\n\n      <div class=\"notice\">\n        「우리학교 소장도서」 시트에는 책 제목과 ISBN13을 반드시 입력해야 합니다.\n        ISBN13이 없으면 “이미 소장중인 도서” 확인이 정확히 작동하지 않습니다.\n      </div>\n    </div>\n\n    <div class=\"card\">\n      <label>공개 신청 링크</label>\n      <p id=\"webAppStatus\" class=\"help\"></p>\n      <p class=\"help\">\n        웹 앱을 새로 배포한 뒤 이 창에서 학교 최초 설정을 다시 실행하면\n        사용설명서의 신청 버튼이 새 공개 링크로 연결됩니다.\n      </p>\n    </div>\n  </section>\n\n  <div id=\"message\" class=\"message\"></div>\n\n  <div class=\"buttons\">\n    <button\n      class=\"secondary\"\n      type=\"button\"\n      onclick=\"google.script.host.close()\"\n    >\n      닫기\n    </button>\n\n    <button\n      id=\"saveButton\"\n      class=\"primary\"\n      type=\"button\"\n      onclick=\"saveCurrentMode()\"\n    ></button>\n  </div>\n\n  <script>\n    const MODE = '<?= mode ?>';\n\n    const pageTitle =\n      document.getElementById('pageTitle');\n    const pageSubtitle =\n      document.getElementById('pageSubtitle');\n    const apiPanel =\n      document.getElementById('apiPanel');\n    const schoolPanel =\n      document.getElementById('schoolPanel');\n    const yes24Key =\n      document.getElementById('yes24Key');\n    const schoolName =\n      document.getElementById('schoolName');\n    const managerName =\n      document.getElementById('managerName');\n    const managerExtension =\n      document.getElementById('managerExtension');\n    const maxClasses =\n      document.getElementById('maxClasses');\n    const spreadsheetName =\n      document.getElementById('spreadsheetName');\n    const statusList =\n      document.getElementById('statusList');\n    const webAppStatus =\n      document.getElementById('webAppStatus');\n    const message =\n      document.getElementById('message');\n    const saveButton =\n      document.getElementById('saveButton');\n\n    window.addEventListener('load', function() {\n      if (MODE === 'api') {\n        pageTitle.textContent = '🔑 도서 API 키 설정';\n        pageSubtitle.textContent =\n          'YES24은 꼭 입력하고, KDC 참고용 키는 필요할 때만 입력하세요.';\n        apiPanel.classList.remove('hidden');\n        saveButton.textContent = '저장하고 YES24 연결 확인';\n      } else {\n        pageTitle.textContent = '🏫 학교 최초 설정';\n        pageSubtitle.textContent =\n          '학교명과 담당자 연락처를 저장하고 신청서와 소장도서 목록에 연결합니다.';\n        schoolPanel.classList.remove('hidden');\n        saveButton.textContent = '이 학교 최초 설정 실행';\n      }\n\n      loadState();\n    });\n\n    function loadState() {\n      setBusy(true);\n\n      google.script.run\n        .withSuccessHandler(function(state) {\n          setBusy(false);\n          yes24Key.value = state.yes24Key || '';\n          schoolName.value = state.schoolName || '';\n          managerName.value = state.managerName || '';\n          managerExtension.value = state.managerExtension || '';\n          maxClasses.value = state.maxClasses || '';\n          spreadsheetName.textContent =\n            '현재 사본: ' + (state.spreadsheetName || '');\n          renderStatus(state);\n        })\n        .withFailureHandler(function(error) {\n          setBusy(false);\n          showMessage(getErrorMessage(error), 'error');\n        })\n        .HR_getCopySetupState();\n    }\n\n    function renderStatus(state) {\n      const items = [\n        {\n          ok: state.yes24Ready,\n          text: 'YES24 API 키 (필수)'\n        },\n        {\n          ok: state.targetSheetReady,\n          text: '희망도서 목록 시트'\n        },\n        {\n          ok: state.librarySheetReady,\n          text: '우리학교 소장도서 시트'\n        },\n        {\n          ok: state.libraryHeadersReady,\n          text: '소장도서 책 제목·ISBN13 열'\n        },\n        {\n          ok: state.configuredSpreadsheetId === state.spreadsheetId,\n          text: '이 학교 신청서 연결'\n        }\n      ];\n\n      statusList.innerHTML = '';\n\n      items.forEach(function(item) {\n        const li = document.createElement('li');\n\n        li.className =\n          'status-item ' +\n          (item.ok || item.optional ? 'ok' : 'need');\n        li.textContent =\n          (\n            item.ok\n              ? '✅ '\n              : item.optional\n                ? '➖ '\n                : '⚠️ '\n          ) +\n          item.text +\n          (\n            item.optional && !item.ok\n              ? ' · 없어도 됨'\n              : ''\n          );\n        statusList.appendChild(li);\n      });\n\n      webAppStatus.textContent = state.webAppUrl\n        ? '✅ 공개 신청 링크가 있습니다.'\n        : '⚠️ 아직 공개 신청 링크가 없습니다. 웹 앱 배포가 필요합니다.';\n    }\n\n    function saveCurrentMode() {\n      if (MODE === 'api') {\n        saveApiKey();\n      } else {\n        runSchoolSetup();\n      }\n    }\n\n    function saveApiKey() {\n      const key = yes24Key.value.trim();\n\n      if (!key) {\n        showMessage(\n          'YES24 API Key를 입력해 주세요.',\n          'error'\n        );\n        return;\n      }\n\n      setBusy(true);\n      showMessage(\n        'YES24 검색 연결을 확인하고 있습니다…',\n        'success'\n      );\n\n      google.script.run\n        .withSuccessHandler(function(response) {\n          setBusy(false);\n          showMessage(response.message, 'success');\n        })\n        .withFailureHandler(function(error) {\n          setBusy(false);\n          showMessage(getErrorMessage(error), 'error');\n        })\n        .HR_saveYes24ApiKey({\n          yes24Key: key\n        });\n    }\n\n    function runSchoolSetup() {\n      setBusy(true);\n      showMessage(\n        '현재 학교 사본을 연결하고 점검하고 있습니다…',\n        'success'\n      );\n\n      google.script.run\n        .withSuccessHandler(function(response) {\n          setBusy(false);\n          showMessage(response.message, 'success');\n          loadState();\n        })\n        .withFailureHandler(function(error) {\n          setBusy(false);\n          showMessage(getErrorMessage(error), 'error');\n        })\n        .HR_runSchoolInitialSetup({\n          schoolName: schoolName.value.trim(),\n          managerName: managerName.value.trim(),\n          managerExtension: managerExtension.value.trim(),\n          maxClasses: maxClasses.value.trim()\n        });\n    }\n\n    function setBusy(isBusy) {\n      saveButton.disabled = isBusy;\n    }\n\n    function showMessage(text, type) {\n      message.textContent = text;\n      message.className =\n        'message show ' + (type || 'success');\n    }\n\n    function getErrorMessage(error) {\n      return error && error.message\n        ? error.message\n        : String(\n          error ||\n          '처리 중 오류가 발생했습니다.'\n        );\n    }\n  </script>\n</body>\n</html>\n","Inquiry":"<!doctype html>\n<html lang=\"ko\">\n<head>\n<base target=\"_top\">\n<meta charset=\"utf-8\">\n<style>\n  *{box-sizing:border-box}\n  body{margin:0;padding:18px 20px 22px;background:#f8f5ed;font-family:Arial,\"Noto Sans KR\",sans-serif;color:#2e3a33}\n  .tabs{display:flex;gap:8px;margin-bottom:14px}\n  .tab{flex:1;padding:11px;border:1px solid #c9d3cc;border-radius:12px;background:#fff;font-size:15px;font-weight:700;color:#4b5a52;cursor:pointer}\n  .tab.on{background:#244d42;border-color:#244d42;color:#fff}\n  label{display:block;margin:12px 0 5px;font-size:13px;font-weight:700;color:#314e43}\n  input,select,textarea{width:100%;padding:10px 12px;border:1px solid #cfd8d2;border-radius:10px;font-size:15px;font-family:inherit;background:#fff}\n  textarea{height:150px;resize:vertical;line-height:1.5}\n  .hint{margin:6px 2px 0;font-size:12px;color:#6a786f;line-height:1.5}\n  .warn{margin-top:12px;padding:10px 12px;border-radius:10px;background:#fdf1d6;color:#6b4d0a;font-size:12.5px;line-height:1.5}\n  .school{display:inline-block;margin-bottom:2px;padding:3px 10px;border-radius:999px;background:#e3efe9;color:#1f5c4a;font-size:12.5px;font-weight:700}\n  button.go{width:100%;margin-top:16px;padding:13px;border:0;border-radius:12px;background:#244d42;color:#fff;font-size:16px;font-weight:800;cursor:pointer}\n  button.go:disabled{opacity:.55;cursor:default}\n  .msg{margin-top:12px;padding:10px 12px;border-radius:10px;font-size:14px;line-height:1.5}\n  .msg.ok{background:#e3efe9;color:#1f5c4a}.msg.err{background:#fdecea;color:#8a2a1f}\n  .item{margin-bottom:12px;padding:13px 14px;border:1px solid #d9e2dc;border-radius:12px;background:#fff}\n  .item .top{display:flex;justify-content:space-between;align-items:center;font-size:12.5px;color:#6a786f}\n  .badge{padding:2px 9px;border-radius:999px;font-weight:700;font-size:12px}\n  .badge.wait{background:#eceff0;color:#55616a}.badge.done{background:#dcefe6;color:#1f5c4a}\n  .q{margin:8px 0 0;font-size:14.5px;font-weight:700;color:#244d42}\n  .a{margin:9px 0 0;padding:10px 12px;border-radius:9px;background:#fdf6e3;font-size:14px;line-height:1.55;white-space:pre-wrap}\n  .empty{padding:30px 10px;text-align:center;color:#6a786f;font-size:14px}\n  .hp{position:absolute;left:-9999px;opacity:0;height:0;width:0}\n  [hidden]{display:none!important}\n</style>\n</head>\n<body>\n<div class=\"tabs\">\n  <button type=\"button\" class=\"tab\" id=\"tabSend\">✏️ 문의하기</button>\n  <button type=\"button\" class=\"tab\" id=\"tabMine\">📬 내 문의</button>\n</div>\n\n<div id=\"paneSend\">\n  <span class=\"school\" id=\"school\"></span>\n  <label for=\"kind\">종류</label>\n  <select id=\"kind\"></select>\n  <label for=\"name\">이름 (선택)</label>\n  <input id=\"name\" type=\"text\" maxlength=\"40\" autocomplete=\"off\" placeholder=\"예: 홍길동 선생님\" value=\"<?= savedName ?>\">\n  <label for=\"body\">문의 내용</label>\n  <textarea id=\"body\" maxlength=\"2000\" placeholder=\"어떤 화면에서 무엇을 하셨을 때 어떻게 되었는지 적어 주시면 빨리 도와드릴 수 있습니다.\"></textarea>\n  <label for=\"email\">답장 받을 이메일 (선택)</label>\n  <input id=\"email\" type=\"email\" maxlength=\"100\" autocomplete=\"off\" placeholder=\"비워 두셔도 「내 문의」 탭에서 답을 볼 수 있습니다.\">\n  <input class=\"hp\" id=\"website\" type=\"text\" tabindex=\"-1\" autocomplete=\"off\" aria-hidden=\"true\">\n  <div class=\"warn\">🔒 학생 이름·학년반번호 같은 개인정보는 적지 마십시오. 학교 이름과 앱 버전만 함께 전달됩니다.</div>\n  <button class=\"go\" id=\"send\" type=\"button\">보내기</button>\n  <div id=\"sendMsg\" class=\"msg\" hidden></div>\n</div>\n\n<div id=\"paneMine\" hidden>\n  <div id=\"mineList\"></div>\n  <button class=\"go\" id=\"refresh\" type=\"button\">새로고침</button>\n</div>\n\n<script>\n  var KINDS = <?!= kinds ?>;\n  var SCHOOL = '<?= school ?>';\n  var START = '<?= tab ?>';\n  var $ = function (id) { return document.getElementById(id); };\n\n  $('school').textContent = SCHOOL ? '🏫 ' + SCHOOL : '🏫 학교 이름 미설정';\n  KINDS.forEach(function (k) { var o = document.createElement('option'); o.value = k; o.textContent = k; $('kind').appendChild(o); });\n\n  function show(tab) {\n    $('paneSend').hidden = tab !== 'send';\n    $('paneMine').hidden = tab !== 'mine';\n    $('tabSend').className = 'tab' + (tab === 'send' ? ' on' : '');\n    $('tabMine').className = 'tab' + (tab === 'mine' ? ' on' : '');\n    if (tab === 'mine') loadMine();\n  }\n  $('tabSend').onclick = function () { show('send'); };\n  $('tabMine').onclick = function () { show('mine'); };\n\n  function msg(el, text, ok) { el.textContent = text; el.className = 'msg ' + (ok ? 'ok' : 'err'); el.hidden = false; }\n\n  $('send').onclick = function () {\n    var btn = $('send');\n    var body = $('body').value.trim();\n    if (body.length < 5) { msg($('sendMsg'), '문의 내용을 5글자 이상 적어 주세요.', false); return; }\n    btn.disabled = true; btn.textContent = '보내는 중…'; $('sendMsg').hidden = true;\n    google.script.run\n      .withSuccessHandler(function (r) {\n        btn.disabled = false; btn.textContent = '보내기';\n        if (r && r.ok) {\n          $('body').value = '';\n          msg($('sendMsg'), '✅ 문의를 보냈습니다. 답이 달리면 「📬 내 문의」 탭에서 볼 수 있습니다.', true);\n        } else {\n          msg($('sendMsg'), (r && r.message) || '보내지 못했습니다.', false);\n        }\n      })\n      .withFailureHandler(function (e) {\n        btn.disabled = false; btn.textContent = '보내기';\n        msg($('sendMsg'), '보내지 못했습니다. (' + String((e && e.message) || e || '알 수 없는 오류').slice(0, 200) + ')', false);\n      })\n      .HR_sendInquiry({ kind: $('kind').value, name: $('name').value, body: body, email: $('email').value, website: $('website').value });\n  };\n\n  function esc(t) { var d = document.createElement('div'); d.textContent = String(t == null ? '' : t); return d.innerHTML; }\n\n  function loadMine() {\n    var list = $('mineList');\n    list.innerHTML = '<div class=\"empty\">불러오는 중…</div>';\n    google.script.run\n      .withSuccessHandler(function (r) {\n        if (!r || !r.ok) { list.innerHTML = '<div class=\"empty\">' + esc((r && r.message) || '불러오지 못했습니다.') + '</div>'; return; }\n        if (!r.items.length) { list.innerHTML = '<div class=\"empty\">아직 보낸 문의가 없습니다.</div>'; return; }\n        list.innerHTML = r.items.map(function (it) {\n          var done = it.status === '답변완료';\n          return '<div class=\"item\"><div class=\"top\"><span>' + esc(it.at) + ' · ' + esc(it.kind) + '</span>' +\n            '<span class=\"badge ' + (done ? 'done' : 'wait') + '\">' + (done ? '답변 도착' : '답변 대기') + '</span></div>' +\n            '<p class=\"q\">' + esc(it.preview) + '</p>' +\n            (done && it.answer ? '<div class=\"a\">💬 ' + esc(it.answer) + (it.answeredAt ? '\\n(' + esc(it.answeredAt) + ')' : '') + '</div>' : '') +\n            '</div>';\n        }).join('');\n      })\n      .withFailureHandler(function () { list.innerHTML = '<div class=\"empty\">불러오지 못했습니다. 잠시 뒤 다시 시도해 주세요.</div>'; })\n      .HR_getMyInquiries();\n  }\n  $('refresh').onclick = loadMine;\n  show(START === 'mine' ? 'mine' : 'send');\n</script>\n</body>\n</html>\n","RequestSidebar":"<!DOCTYPE html>\n<html lang=\"ko\">\n<head>\n  <base target=\"_top\">\n  <meta charset=\"UTF-8\">\n\n  <style>\n    :root {\n      --blue: #4f78d1;\n      --blue-dark: #3159ac;\n      --blue-soft: #edf3ff;\n      --ink: #24324a;\n      --sub: #68748a;\n      --line: #dce4f2;\n      --success: #237a4b;\n      --success-soft: #eaf7ef;\n      --danger: #b33a3a;\n      --danger-soft: #fff0ef;\n      --paper: #ffffff;\n      --background: #f5f7fb;\n    }\n\n    * {\n      box-sizing: border-box;\n    }\n\n    body {\n      margin: 0;\n      padding: 17px;\n      background: var(--background);\n      color: var(--ink);\n      font-family:\n        -apple-system,\n        BlinkMacSystemFont,\n        \"Apple SD Gothic Neo\",\n        \"Noto Sans KR\",\n        sans-serif;\n    }\n\n    .header {\n      width: 100%;\n      max-width: 760px;\n      margin-left: auto;\n      margin-right: auto;\n      margin-bottom: 16px;\n    }\n\n    .eyebrow {\n      margin: 0 0 7px;\n      color: var(--blue);\n      font-size: 12px;\n      font-weight: 900;\n      letter-spacing: 0.12em;\n    }\n\n    h1 {\n      margin: 0;\n      font-size: 27px;\n      line-height: 1.25;\n      letter-spacing: -0.04em;\n    }\n\n    .description {\n      margin: 10px 0 0;\n      color: #536077;\n      font-size: 14px;\n      line-height: 1.7;\n    }\n\n    .card {\n      width: 100%;\n      max-width: 760px;\n      margin: 0 auto;\n      padding: 16px;\n      border: 1px solid var(--line);\n      border-radius: 16px;\n      background: var(--paper);\n      box-shadow: 0 8px 25px rgba(50, 76, 125, 0.08);\n    }\n\n    .field {\n      margin-bottom: 17px;\n    }\n\n    .field:last-of-type {\n      margin-bottom: 0;\n    }\n\n    label {\n      display: block;\n      margin-bottom: 8px;\n      color: #273b5e;\n      font-size: 15px;\n      font-weight: 850;\n    }\n\n    input {\n      width: 100%;\n      height: 54px;\n      padding: 0 14px;\n      border: 1px solid var(--line);\n      border-radius: 12px;\n      outline: none;\n      background: #fff;\n      color: var(--ink);\n      font-size: 17px;\n      font-weight: 650;\n    }\n\n    input:focus {\n      border-color: var(--blue);\n      box-shadow: 0 0 0 3px rgba(79, 120, 209, 0.12);\n    }\n\n    .help {\n      margin: 8px 1px 0;\n      color: #68748a;\n      font-size: 12px;\n      line-height: 1.6;\n    }\n\n    .submit-button {\n      width: 100%;\n      min-height: 55px;\n      margin-top: 18px;\n      border: 0;\n      border-radius: 12px;\n      background: var(--blue);\n      color: #fff;\n      font-size: 17px;\n      font-weight: 900;\n      cursor: pointer;\n      transition: background 0.15s, transform 0.15s;\n    }\n\n    .submit-button:hover {\n      background: var(--blue-dark);\n      transform: translateY(-1px);\n    }\n\n    .submit-button:disabled {\n      cursor: wait;\n      opacity: 0.58;\n      transform: none;\n    }\n\n    .message {\n      display: none;\n      margin-top: 15px;\n      padding: 14px;\n      border-radius: 11px;\n      font-size: 14px;\n      line-height: 1.65;\n    }\n\n    .message.show {\n      display: block;\n    }\n\n    .message.info {\n      background: var(--blue-soft);\n      color: var(--blue-dark);\n    }\n\n    .message.success {\n      background: var(--success-soft);\n      color: var(--success);\n    }\n\n    .message.error {\n      background: var(--danger-soft);\n      color: var(--danger);\n    }\n\n    .search-results {\n      display: none;\n      margin-top: 15px;\n    }\n\n    .search-results.show {\n      display: block;\n    }\n\n    .search-results-title {\n      margin: 0 0 9px;\n      color: #273b5e;\n      font-size: 15px;\n      font-weight: 900;\n    }\n\n    .book-list {\n      display: grid;\n      gap: 10px;\n    }\n\n    .book-card {\n      display: grid;\n      grid-template-columns: 76px 1fr;\n      gap: 12px;\n      padding: 12px;\n      border: 1px solid var(--line);\n      border-radius: 13px;\n      background: #fff;\n      box-shadow: 0 5px 16px rgba(50, 76, 125, 0.06);\n    }\n\n    .book-cover {\n      display: flex;\n      width: 76px;\n      height: 110px;\n      overflow: hidden;\n      align-items: center;\n      justify-content: center;\n      border-radius: 7px;\n      background: #eef1f6;\n      color: #8c96a8;\n      font-size: 24px;\n    }\n\n    .book-cover img {\n      width: 100%;\n      height: 100%;\n      object-fit: cover;\n    }\n\n    .book-info {\n      display: flex;\n      min-width: 0;\n      flex-direction: column;\n    }\n\n    .book-title {\n      margin: 0;\n      color: var(--ink);\n      font-size: 15px;\n      font-weight: 900;\n      line-height: 1.45;\n    }\n\n    .book-meta {\n      margin: 7px 0 0;\n      color: #5d6980;\n      font-size: 12px;\n      line-height: 1.55;\n    }\n\n    .select-book-button {\n      width: 100%;\n      min-height: 43px;\n      margin-top: auto;\n      padding: 9px 8px;\n      border: 0;\n      border-radius: 9px;\n      background: var(--blue-soft);\n      color: var(--blue-dark);\n      font-size: 14px;\n      font-weight: 900;\n      cursor: pointer;\n    }\n\n    .select-book-button:hover {\n      background: #dce8ff;\n    }\n\n    .select-book-button:disabled {\n      cursor: wait;\n      opacity: 0.55;\n    }\n\n    .select-book-button.done {\n      background: var(--success-soft);\n      color: var(--success);\n      opacity: 1;\n    }\n\n    .book-card.done {\n      border-color: #b9dcc7;\n      background: #fbfffc;\n    }\n\n    .book-card.selected {\n      border: 2px solid var(--blue);\n      background: #f7faff;\n      box-shadow: 0 8px 22px rgba(79, 120, 209, 0.16);\n    }\n\n    .select-book-button.selected {\n      background: var(--blue);\n      color: #fff;\n    }\n\n    .cart-bar {\n      position: sticky;\n      bottom: 12px;\n      z-index: 30;\n      display: none;\n      width: 100%;\n      max-width: 760px;\n      margin: 16px auto 0;\n      padding: 12px;\n      align-items: center;\n      gap: 12px;\n      border: 1px solid #cbd9f5;\n      border-radius: 16px;\n      background: rgba(255, 255, 255, 0.96);\n      box-shadow: 0 12px 36px rgba(37, 61, 105, 0.19);\n      backdrop-filter: blur(10px);\n    }\n\n    .cart-bar.show {\n      display: flex;\n    }\n\n    .cart-count {\n      flex: 0 0 auto;\n      color: var(--blue-dark);\n      font-size: 15px;\n      font-weight: 900;\n      white-space: nowrap;\n    }\n\n    .cart-submit {\n      min-height: 50px;\n      flex: 1;\n      border: 0;\n      border-radius: 12px;\n      background: var(--blue);\n      color: #fff;\n      font-size: 16px;\n      font-weight: 900;\n      cursor: pointer;\n    }\n\n    .cart-submit:disabled {\n      cursor: wait;\n      opacity: 0.55;\n    }\n\n    .cart-toggle {\n      display: flex;\n      min-height: 50px;\n      flex: 0 0 auto;\n      padding: 8px 14px;\n      align-items: center;\n      gap: 10px;\n      border: 0;\n      border-radius: 12px;\n      background: #eef4ff;\n      color: var(--blue-dark);\n      cursor: pointer;\n    }\n\n    .cart-count strong {\n      color: var(--blue);\n      font-size: 19px;\n    }\n\n    .cart-open-hint {\n      color: #637492;\n      font-size: 12px;\n      font-weight: 800;\n    }\n\n    .cart-drawer {\n      position: fixed;\n      inset: 0;\n      z-index: 95;\n      display: none;\n    }\n\n    .cart-drawer.show {\n      display: block;\n    }\n\n    .cart-backdrop {\n      position: absolute;\n      inset: 0;\n      width: 100%;\n      height: 100%;\n      border: 0;\n      background: rgba(25, 36, 56, 0.5);\n      backdrop-filter: blur(4px);\n      cursor: pointer;\n    }\n\n    .cart-panel {\n      position: absolute;\n      top: 0;\n      right: 0;\n      display: flex;\n      width: min(430px, 100%);\n      height: 100%;\n      padding: 24px;\n      flex-direction: column;\n      background: #f8faff;\n      box-shadow: -18px 0 60px rgba(28, 44, 73, 0.24);\n      animation: cart-slide-in 0.22s ease-out;\n    }\n\n    .cart-panel-header {\n      display: flex;\n      align-items: flex-start;\n      justify-content: space-between;\n      gap: 16px;\n    }\n\n    .cart-panel-header p {\n      margin: 0 0 4px;\n      color: #7990b8;\n      font-size: 11px;\n      font-weight: 950;\n      letter-spacing: 0.12em;\n    }\n\n    .cart-panel-header h2 {\n      margin: 0;\n      color: var(--ink);\n      font-size: 25px;\n      letter-spacing: -0.04em;\n    }\n\n    .cart-close {\n      width: 44px;\n      height: 44px;\n      flex: 0 0 auto;\n      border: 0;\n      border-radius: 14px;\n      background: #e7edf9;\n      color: #435675;\n      font-size: 28px;\n      line-height: 1;\n      cursor: pointer;\n    }\n\n    .cart-guide {\n      margin: 14px 0 16px;\n      padding: 11px 13px;\n      border-radius: 12px;\n      background: #edf3ff;\n      color: #526580;\n      font-size: 13px;\n      font-weight: 750;\n      line-height: 1.55;\n    }\n\n    .cart-list {\n      min-height: 0;\n      flex: 1;\n      overflow-y: auto;\n      padding-right: 4px;\n    }\n\n    .cart-empty {\n      display: flex;\n      min-height: 220px;\n      align-items: center;\n      justify-content: center;\n      color: #75829a;\n      font-size: 14px;\n      font-weight: 800;\n      text-align: center;\n    }\n\n    .cart-item {\n      display: grid;\n      grid-template-columns: 58px minmax(0, 1fr) auto;\n      gap: 12px;\n      margin-bottom: 10px;\n      padding: 12px;\n      align-items: center;\n      border: 1px solid #dfe7f5;\n      border-radius: 15px;\n      background: #fff;\n      box-shadow: 0 7px 20px rgba(49, 73, 112, 0.07);\n    }\n\n    .cart-item-cover {\n      display: flex;\n      width: 58px;\n      height: 78px;\n      overflow: hidden;\n      align-items: center;\n      justify-content: center;\n      border-radius: 7px;\n      background: #edf1f7;\n      font-size: 24px;\n    }\n\n    .cart-item-cover img {\n      width: 100%;\n      height: 100%;\n      object-fit: cover;\n    }\n\n    .cart-item-title {\n      margin: 0;\n      color: var(--ink);\n      font-size: 14px;\n      font-weight: 900;\n      line-height: 1.4;\n    }\n\n    .cart-item-meta {\n      margin: 5px 0 0;\n      color: #748098;\n      font-size: 11px;\n      line-height: 1.45;\n    }\n\n    .cart-remove {\n      min-width: 54px;\n      min-height: 38px;\n      border: 1px solid #efc9c2;\n      border-radius: 10px;\n      background: #fff7f5;\n      color: #b34f3f;\n      font-size: 12px;\n      font-weight: 900;\n      cursor: pointer;\n    }\n\n    .cart-panel-actions {\n      display: grid;\n      grid-template-columns: auto minmax(0, 1fr);\n      gap: 10px;\n      padding-top: 14px;\n      border-top: 1px solid #dde5f2;\n    }\n\n    .cart-clear {\n      min-height: 50px;\n      padding: 0 16px;\n      border: 1px solid #ccd7e9;\n      border-radius: 12px;\n      background: #fff;\n      color: #5e6e88;\n      font-size: 13px;\n      font-weight: 900;\n      cursor: pointer;\n    }\n\n    .cart-drawer-submit {\n      width: 100%;\n    }\n\n    @keyframes cart-slide-in {\n      from {\n        opacity: 0.4;\n        transform: translateX(30px);\n      }\n\n      to {\n        opacity: 1;\n        transform: translateX(0);\n      }\n    }\n\n    @media (max-width: 620px) {\n      .cart-bar {\n        bottom: 8px;\n        padding: 9px;\n        gap: 8px;\n      }\n\n      .cart-toggle {\n        min-width: 0;\n        padding: 8px 10px;\n      }\n\n      .cart-open-hint {\n        display: none;\n      }\n\n      .cart-submit {\n        min-height: 48px;\n        font-size: 14px;\n      }\n\n      .cart-panel {\n        top: auto;\n        bottom: 0;\n        width: 100%;\n        height: min(78vh, 680px);\n        padding: 20px 16px 16px;\n        border-radius: 24px 24px 0 0;\n        animation-name: cart-rise-in;\n      }\n\n      .cart-item {\n        grid-template-columns: 52px minmax(0, 1fr) auto;\n        padding: 10px;\n      }\n\n      .cart-item-cover {\n        width: 52px;\n        height: 70px;\n      }\n    }\n\n    @keyframes cart-rise-in {\n      from {\n        opacity: 0.4;\n        transform: translateY(30px);\n      }\n\n      to {\n        opacity: 1;\n        transform: translateY(0);\n      }\n    }\n\n    .footer {\n      width: 100%;\n      max-width: 760px;\n      margin: 16px auto 0;\n      color: #68748a;\n      font-size: 12px;\n      line-height: 1.65;\n      text-align: center;\n    }\n\n    .creator-watermark {\n      position: fixed;\n      bottom: 10px;\n      left: 12px;\n      z-index: 20;\n      padding: 5px 8px;\n      border-radius: 7px;\n      background: rgba(255, 255, 255, 0.78);\n      color: #4b5870;\n      font-size: 11px;\n      font-weight: 750;\n      letter-spacing: -0.02em;\n      opacity: 0.56;\n      pointer-events: none;\n      backdrop-filter: blur(4px);\n    }\n\n    .notice-modal {\n      position: fixed;\n      inset: 0;\n      z-index: 100;\n      display: none;\n      padding: 22px;\n      align-items: center;\n      justify-content: center;\n      background: rgba(27, 38, 58, 0.48);\n      backdrop-filter: blur(5px);\n    }\n\n    .notice-modal.show {\n      display: flex;\n    }\n\n    .notice-dialog {\n      width: min(500px, 100%);\n      overflow: hidden;\n      border: 1px solid rgba(255, 255, 255, 0.72);\n      border-radius: 24px;\n      background: #fff;\n      box-shadow: 0 24px 70px rgba(21, 34, 58, 0.28);\n      animation: notice-in 0.2s ease-out;\n    }\n\n    .notice-visual {\n      display: flex;\n      min-height: 130px;\n      align-items: center;\n      justify-content: center;\n      background:\n        radial-gradient(circle at 30% 20%, #ffffff 0, transparent 32%),\n        linear-gradient(135deg, #eef4ff, #dfe9ff);\n    }\n\n    .notice-modal.owned .notice-visual {\n      background:\n        radial-gradient(circle at 30% 20%, #ffffff 0, transparent 32%),\n        linear-gradient(135deg, #fff3eb, #ffe1d4);\n    }\n\n    .notice-icon {\n      display: flex;\n      width: 76px;\n      height: 76px;\n      align-items: center;\n      justify-content: center;\n      border-radius: 24px;\n      background: #fff;\n      box-shadow: 0 10px 30px rgba(58, 90, 158, 0.16);\n      font-size: 38px;\n    }\n\n    .notice-content {\n      padding: 24px 24px 22px;\n      text-align: center;\n    }\n\n    .notice-title {\n      margin: 0;\n      color: var(--ink);\n      font-size: 22px;\n      font-weight: 950;\n      line-height: 1.4;\n      letter-spacing: -0.04em;\n    }\n\n    .notice-book {\n      margin: 13px 0 0;\n      color: #344c75;\n      font-size: 16px;\n      font-weight: 850;\n      line-height: 1.55;\n      word-break: keep-all;\n    }\n\n    .notice-description {\n      max-height: 300px;\n      overflow-y: auto;\n      margin: 16px 0 0;\n      padding: 14px 16px;\n      border: 1px solid #e1e8f4;\n      border-radius: 14px;\n      background: #f7f9fd;\n      color: #536179;\n      font-size: 14px;\n      font-weight: 700;\n      line-height: 1.72;\n      text-align: left;\n      word-break: keep-all;\n      white-space: pre-line;\n    }\n\n    .notice-close {\n      width: 100%;\n      min-height: 52px;\n      margin-top: 20px;\n      border: 0;\n      border-radius: 13px;\n      background: var(--blue);\n      color: #fff;\n      font-size: 16px;\n      font-weight: 900;\n      cursor: pointer;\n    }\n\n    .notice-close:hover {\n      background: var(--blue-dark);\n    }\n\n    .notice-actions {\n      display: grid;\n      grid-template-columns: 1fr 1fr;\n      gap: 10px;\n      margin-top: 20px;\n    }\n\n    .notice-actions.single {\n      grid-template-columns: 1fr;\n    }\n\n    .notice-actions .notice-close {\n      margin-top: 0;\n    }\n\n    .notice-undo {\n      min-height: 52px;\n      border: 0;\n      border-radius: 13px;\n      background: #f1ae45;\n      color: #49320e;\n      font-size: 15px;\n      font-weight: 950;\n      cursor: pointer;\n    }\n\n    .notice-undo:hover {\n      background: #e59c2f;\n    }\n\n    .notice-undo:disabled {\n      cursor: wait;\n      opacity: 0.58;\n    }\n\n    .notice-undo[hidden] {\n      display: none;\n    }\n\n    .notice-modal.owned .notice-close {\n      background: #d96545;\n    }\n\n    @keyframes notice-in {\n      from {\n        opacity: 0;\n        transform: translateY(12px) scale(0.97);\n      }\n\n      to {\n        opacity: 1;\n        transform: translateY(0) scale(1);\n      }\n    }\n\n    @media (min-width: 720px) {\n      body {\n        padding: 36px 24px;\n      }\n\n      .book-list {\n        grid-template-columns: repeat(2, minmax(0, 1fr));\n      }\n    }\n\n    @media (max-width: 420px) {\n      body {\n        padding: 17px;\n      }\n    }\n  \n    .bs-section {\n      width: 100%;\n      max-width: 760px;\n      margin: 18px auto 0;\n      padding: 17px;\n      border: 1px solid var(--line);\n      border-radius: 18px;\n      background: var(--paper);\n      box-shadow: 0 8px 25px rgba(50, 76, 125, 0.08);\n    }\n\n    .bs-heading {\n      display: flex;\n      align-items: flex-start;\n      justify-content: space-between;\n      gap: 12px;\n      margin-bottom: 14px;\n    }\n\n    .bs-title {\n      margin: 0;\n      font-size: 21px;\n      line-height: 1.35;\n      letter-spacing: -0.03em;\n    }\n\n    .bs-help {\n      margin: 5px 0 0;\n      color: var(--sub);\n      font-size: 13px;\n      line-height: 1.55;\n    }\n\n    .bs-badge {\n      flex: 0 0 auto;\n      padding: 6px 9px;\n      border-radius: 999px;\n      background: var(--blue-soft);\n      color: var(--blue-dark);\n      font-size: 11px;\n      font-weight: 850;\n    }\n\n    .bs-tabs {\n      display: grid;\n      grid-template-columns: repeat(4, 1fr);\n      gap: 7px;\n      margin-bottom: 14px;\n      padding: 5px;\n      border-radius: 13px;\n      background: #eef2f8;\n    }\n\n    .bs-tab {\n      min-height: 43px;\n      padding: 0 8px;\n      border: 0;\n      border-radius: 10px;\n      background: transparent;\n      color: #526078;\n      font-size: 15px;\n      font-weight: 850;\n      cursor: pointer;\n    }\n\n    .bs-tab.active {\n      background: #fff;\n      color: var(--blue-dark);\n      box-shadow: 0 3px 10px rgba(49, 89, 172, 0.13);\n    }\n\n    .bs-status {\n      display: none;\n      padding: 24px 10px;\n      color: var(--sub);\n      font-size: 14px;\n      text-align: center;\n    }\n\n    .bs-status.show {\n      display: block;\n    }\n\n    .bs-list {\n      display: grid;\n      grid-auto-flow: column;\n      grid-auto-columns: minmax(205px, 215px);\n      gap: 12px;\n      overflow-x: auto;\n      padding: 2px 2px 12px;\n      scroll-snap-type: x proximity;\n    }\n\n    .bs-card {\n      position: relative;\n      display: flex;\n      min-height: 330px;\n      flex-direction: column;\n      padding: 12px;\n      border: 1px solid #e1e7f2;\n      border-radius: 15px;\n      background: #fff;\n      scroll-snap-align: start;\n    }\n\n    .bs-rank {\n      position: absolute;\n      top: 9px;\n      left: 9px;\n      z-index: 1;\n      min-width: 31px;\n      height: 31px;\n      padding: 0 7px;\n      border-radius: 10px;\n      background: rgba(36, 50, 74, 0.92);\n      color: #fff;\n      font-size: 14px;\n      font-weight: 900;\n      line-height: 31px;\n      text-align: center;\n    }\n\n    .bs-cover {\n      display: flex;\n      width: 112px;\n      height: 158px;\n      margin: 0 auto 11px;\n      align-items: center;\n      justify-content: center;\n      overflow: hidden;\n      border-radius: 8px;\n      background: #eef1f6;\n      color: #7c8799;\n      font-size: 34px;\n      box-shadow: 0 7px 16px rgba(36, 50, 74, 0.14);\n    }\n\n    .bs-cover img {\n      width: 100%;\n      height: 100%;\n      object-fit: cover;\n    }\n\n    .bs-book-title {\n      display: -webkit-box;\n      overflow: hidden;\n      margin: 0;\n      color: var(--ink);\n      font-size: 15px;\n      font-weight: 900;\n      line-height: 1.45;\n      -webkit-box-orient: vertical;\n      -webkit-line-clamp: 2;\n    }\n\n    .bs-meta {\n      display: -webkit-box;\n      overflow: hidden;\n      min-height: 38px;\n      margin: 6px 0 10px;\n      color: var(--sub);\n      font-size: 12px;\n      line-height: 1.55;\n      -webkit-box-orient: vertical;\n      -webkit-line-clamp: 2;\n    }\n\n    .bs-actions {\n      display: grid;\n      grid-template-columns: 1fr auto;\n      gap: 7px;\n      margin-top: auto;\n    }\n\n    .bs-request {\n      min-height: 40px;\n      padding: 0 10px;\n      border: 0;\n      border-radius: 10px;\n      background: var(--blue);\n      color: #fff;\n      font-size: 12px;\n      font-weight: 900;\n      cursor: pointer;\n    }\n\n    .bs-link {\n      display: inline-flex;\n      min-height: 40px;\n      padding: 0 10px;\n      align-items: center;\n      justify-content: center;\n      border: 1px solid var(--line);\n      border-radius: 10px;\n      color: #5b6880;\n      font-size: 12px;\n      font-weight: 800;\n      text-decoration: none;\n    }\n\n    .bs-source {\n      margin: 4px 1px 0;\n      color: #7d8798;\n      font-size: 11px;\n      text-align: right;\n    }\n\n    .bs-source a {\n      color: var(--blue-dark);\n    }\n\n    @media (max-width: 520px) {\n      .bs-heading {\n        display: block;\n      }\n\n      .bs-badge {\n        display: inline-block;\n        margin-top: 8px;\n      }\n\n      .bs-list {\n        grid-auto-columns: minmax(190px, 78vw);\n      }\n    }\n\n  \n    .bs-more-wrap {\n      margin-top: 6px;\n    }\n\n    .bs-more {\n      display: flex;\n      width: 100%;\n      min-height: 48px;\n      align-items: center;\n      justify-content: center;\n      border: 1px solid #c9d7f2;\n      border-radius: 12px;\n      background: var(--blue-soft);\n      color: var(--blue-dark);\n      font-size: 15px;\n      font-weight: 900;\n      text-decoration: none;\n      transition: background 0.18s, transform 0.18s;\n    }\n\n    .bs-more:hover {\n      background: #dfe9ff;\n      transform: translateY(-1px);\n    }\n\n  \n    .view-tabs {\n      position: sticky;\n      top: 10px;\n      z-index: 30;\n      display: grid;\n      width: 100%;\n      max-width: 760px;\n      margin: 16px auto;\n      padding: 6px;\n      grid-template-columns: repeat(4, 1fr);\n      gap: 7px;\n      border: 1px solid var(--line);\n      border-radius: 16px;\n      background: rgba(255, 255, 255, 0.96);\n      box-shadow: 0 8px 24px rgba(50, 76, 125, 0.11);\n      backdrop-filter: blur(12px);\n    }\n\n    .view-tab {\n      min-height: 49px;\n      padding: 0 12px;\n      border: 0;\n      border-radius: 11px;\n      background: transparent;\n      color: #59677e;\n      font-size: 16px;\n      font-weight: 900;\n      cursor: pointer;\n      transition: background 0.18s, color 0.18s, transform 0.18s;\n    }\n\n    .view-tab:hover {\n      background: #f1f5fd;\n    }\n\n    .view-tab.active {\n      background: var(--blue);\n      color: #fff;\n      box-shadow: 0 5px 13px rgba(79, 120, 209, 0.25);\n    }\n\n    .page-panel {\n      display: none;\n    }\n\n    .page-panel.active {\n      display: block;\n    }\n\n    .bs-carousel {\n      position: relative;\n    }\n\n    .bs-list {\n      scroll-behavior: smooth;\n      scrollbar-width: thin;\n      scrollbar-color: #aebedd #eef2f8;\n      overscroll-behavior-inline: contain;\n      padding-left: 44px;\n      padding-right: 44px;\n    }\n\n    .bs-arrow {\n      position: absolute;\n      top: 50%;\n      z-index: 5;\n      display: flex;\n      width: 48px;\n      height: 58px;\n      margin-top: -29px;\n      align-items: center;\n      justify-content: center;\n      border: 1px solid #cad7ee;\n      border-radius: 16px;\n      background: rgba(255, 255, 255, 0.95);\n      color: var(--blue-dark);\n      font-family: Arial, sans-serif;\n      font-size: 40px;\n      font-weight: 700;\n      line-height: 1;\n      box-shadow: 0 7px 20px rgba(36, 50, 74, 0.18);\n      cursor: pointer;\n      transition: transform 0.18s, background 0.18s;\n    }\n\n    .bs-arrow:hover {\n      background: var(--blue-soft);\n      transform: scale(1.05);\n    }\n\n    .bs-arrow-previous {\n      left: -6px;\n    }\n\n    .bs-arrow-next {\n      right: -6px;\n    }\n\n    .requests-section,\n    .guide-section {\n      width: 100%;\n      max-width: 760px;\n      margin: 0 auto;\n      padding: 22px;\n      border: 1px solid var(--line);\n      border-radius: 20px;\n      background: #fff;\n      box-shadow: 0 10px 28px rgba(50, 76, 125, 0.09);\n    }\n\n    .panel-heading {\n      display: flex;\n      align-items: flex-start;\n      justify-content: space-between;\n      gap: 16px;\n      margin-bottom: 18px;\n    }\n\n    .panel-eyebrow {\n      margin: 0 0 5px !important;\n      color: var(--blue);\n      font-size: 11px !important;\n      font-weight: 900;\n      letter-spacing: 0.12em;\n    }\n\n    .panel-heading h2 {\n      margin: 0;\n      color: var(--ink);\n      font-size: 24px;\n      letter-spacing: -0.035em;\n    }\n\n    .panel-heading p:not(.panel-eyebrow) {\n      margin: 7px 0 0;\n      color: var(--sub);\n      font-size: 14px;\n      line-height: 1.6;\n    }\n\n    .refresh-button {\n      flex: 0 0 auto;\n      min-height: 42px;\n      padding: 0 15px;\n      border: 1px solid #cbd8ef;\n      border-radius: 11px;\n      background: var(--blue-soft);\n      color: var(--blue-dark);\n      font-size: 14px;\n      font-weight: 900;\n      cursor: pointer;\n    }\n\n    .request-filter-tabs {\n      display: grid;\n      grid-template-columns: 1fr 1fr;\n      gap: 10px;\n      margin: 18px 0 13px;\n      padding: 6px;\n      border: 1px solid #dce5df;\n      border-radius: 16px;\n      background: #f3f6f3;\n    }\n\n    .request-filter-tab {\n      display: flex;\n      min-height: 54px;\n      padding: 0 18px;\n      border: 0;\n      border-radius: 12px;\n      align-items: center;\n      justify-content: space-between;\n      gap: 12px;\n      background: transparent;\n      color: #63726a;\n      font-size: 15px;\n      font-weight: 900;\n      cursor: pointer;\n      transition:\n        background 0.16s ease,\n        color 0.16s ease,\n        box-shadow 0.16s ease;\n    }\n\n    .request-filter-tab strong {\n      display: inline-flex;\n      min-width: 32px;\n      min-height: 28px;\n      padding: 0 8px;\n      border-radius: 999px;\n      align-items: center;\n      justify-content: center;\n      background: #e3e9e5;\n      color: #526258;\n      font-size: 13px;\n    }\n\n    .request-filter-tab.active {\n      background: #2f604e;\n      color: #fff;\n      box-shadow: 0 7px 18px rgba(47, 96, 78, 0.18);\n    }\n\n    .request-filter-tab.active strong {\n      background: #fff;\n      color: #2f604e;\n    }\n\n    .request-filter-tab[data-request-filter=\"cancelled\"].active {\n      background: #a84d45;\n      box-shadow: 0 7px 18px rgba(168, 77, 69, 0.17);\n    }\n\n    .request-filter-tab[data-request-filter=\"cancelled\"].active strong {\n      color: #a84d45;\n    }\n\n    @media (max-width: 520px) {\n      .request-filter-tab {\n        min-height: 50px;\n        padding: 0 12px;\n        font-size: 14px;\n      }\n    }\n\n    .request-summary {\n      display: none;\n      margin-bottom: 13px;\n      padding: 11px 14px;\n      border-radius: 11px;\n      background: #f0f5ff;\n      color: #3b5eaa;\n      font-size: 13px;\n      font-weight: 850;\n    }\n\n    .request-summary.show {\n      display: block;\n    }\n\n    .request-list-status {\n      padding: 34px 14px;\n      border: 1px dashed #cfd9ea;\n      border-radius: 14px;\n      color: var(--sub);\n      text-align: center;\n    }\n\n    .request-list-status.hide {\n      display: none;\n    }\n\n    .request-list {\n      display: grid;\n      grid-template-columns: repeat(2, minmax(0, 1fr));\n      gap: 12px;\n    }\n\n    .request-item {\n      display: flex;\n      min-height: 132px;\n      flex-direction: column;\n      padding: 15px;\n      border: 1px solid #e0e7f2;\n      border-radius: 15px;\n      background: #fbfcff;\n    }\n\n    .request-thumb {\n      display: block;\n      width: 64px;\n      height: 90px;\n      object-fit: cover;\n      border-radius: 6px;\n      margin: 4px 0 8px;\n      box-shadow: 0 2px 8px rgba(0,0,0,.25);\n    }\n\n    .request-item-top {\n      display: flex;\n      align-items: center;\n      justify-content: space-between;\n      gap: 8px;\n      margin-bottom: 9px;\n    }\n\n    .request-sequence {\n      color: #70809a;\n      font-size: 12px;\n      font-weight: 900;\n    }\n\n    .request-status-badge {\n      max-width: 62%;\n      padding: 5px 8px;\n      overflow: hidden;\n      border-radius: 999px;\n      background: #fff4d8;\n      color: #936214;\n      font-size: 11px;\n      font-weight: 900;\n      text-overflow: ellipsis;\n      white-space: nowrap;\n    }\n\n    .request-status-badge.done {\n      background: #e6f6eb;\n      color: #247347;\n    }\n\n    .request-status-badge.issue {\n      background: #fff0ef;\n      color: #a5413d;\n    }\n\n    .request-item h3 {\n      margin: 0;\n      color: var(--ink);\n      font-size: 16px;\n      font-weight: 900;\n      line-height: 1.45;\n    }\n\n    .request-item p {\n      margin: 7px 0 0;\n      color: var(--sub);\n      font-size: 13px;\n      line-height: 1.55;\n    }\n\n    .request-date {\n      margin-top: auto !important;\n      padding-top: 9px;\n      color: #8691a3 !important;\n      font-size: 11px !important;\n    }\n\n    .guide-steps {\n      display: grid;\n      gap: 12px;\n      margin: 0;\n      padding: 0;\n      list-style: none;\n    }\n\n    .guide-steps li {\n      display: grid;\n      grid-template-columns: 48px 1fr;\n      gap: 13px;\n      padding: 16px;\n      border: 1px solid #e0e7f2;\n      border-radius: 15px;\n      background: #fbfcff;\n    }\n\n    .guide-number {\n      display: flex;\n      width: 45px;\n      height: 45px;\n      align-items: center;\n      justify-content: center;\n      border-radius: 14px;\n      background: var(--blue);\n      color: #fff;\n      font-size: 20px;\n      font-weight: 950;\n    }\n\n    .guide-steps strong {\n      color: var(--ink);\n      font-size: 16px;\n      line-height: 1.5;\n    }\n\n    .guide-steps p {\n      margin: 6px 0 0;\n      color: var(--sub);\n      font-size: 14px;\n      line-height: 1.65;\n    }\n\n    .guide-start-button {\n      width: 100%;\n      min-height: 54px;\n      margin-top: 17px;\n      border: 0;\n      border-radius: 13px;\n      background: var(--blue);\n      color: #fff;\n      font-size: 17px;\n      font-weight: 900;\n      cursor: pointer;\n    }\n\n    @media (max-width: 620px) {\n      .view-tab {\n        min-height: 46px;\n        padding: 0 5px;\n        font-size: 13px;\n      }\n\n      .request-list {\n        grid-template-columns: 1fr;\n      }\n\n      .requests-section,\n      .guide-section {\n        padding: 17px;\n      }\n\n      .panel-heading {\n        display: block;\n      }\n\n      .refresh-button {\n        width: 100%;\n        margin-top: 12px;\n      }\n\n      .bs-arrow {\n        width: 43px;\n        height: 54px;\n        font-size: 35px;\n      }\n    }\n\n\n    .guide-hover {\n      position: relative;\n      min-width: 0;\n    }\n\n    .guide-hover-trigger {\n      width: 100%;\n      min-height: 49px;\n      padding: 0 12px;\n      border: 0;\n      border-radius: 11px;\n      background: #fff5d9;\n      color: #7b5813;\n      font-size: 16px;\n      font-weight: 900;\n      cursor: help;\n      transition: background 0.18s, transform 0.18s;\n    }\n\n    .guide-hover-trigger:hover,\n    .guide-hover-trigger:focus {\n      background: #ffeab2;\n      transform: translateY(-1px);\n      outline: none;\n    }\n\n    .guide-popover {\n      position: absolute;\n      top: calc(100% + 10px);\n      right: 0;\n      z-index: 100;\n      width: min(440px, calc(100vw - 34px));\n      max-height: calc(100vh - 120px);\n      padding: 18px;\n      overflow-y: auto;\n      border: 1px solid #ead79c;\n      border-radius: 17px;\n      background: #fffdf6;\n      box-shadow: 0 16px 40px rgba(36, 50, 74, 0.22);\n      opacity: 0;\n      visibility: hidden;\n      transform: translateY(-7px);\n      pointer-events: none;\n      transition: opacity 0.16s, transform 0.16s, visibility 0.16s;\n    }\n\n    .guide-hover:hover .guide-popover,\n    .guide-hover:focus-within .guide-popover,\n    .guide-hover.open .guide-popover {\n      opacity: 1;\n      visibility: visible;\n      transform: translateY(0);\n      pointer-events: auto;\n    }\n\n    .guide-popover::before {\n      position: absolute;\n      top: -7px;\n      right: 34px;\n      width: 13px;\n      height: 13px;\n      border-top: 1px solid #ead79c;\n      border-left: 1px solid #ead79c;\n      background: #fffdf6;\n      content: '';\n      transform: rotate(45deg);\n    }\n\n    .guide-popover h2 {\n      margin: 0 0 12px;\n      color: var(--ink);\n      font-size: 20px;\n      letter-spacing: -0.03em;\n    }\n\n    .guide-popover ol {\n      display: grid;\n      gap: 8px;\n      margin: 0;\n      padding: 0;\n      list-style: none;\n    }\n\n    .guide-popover li {\n      display: grid;\n      grid-template-columns: 108px 1fr;\n      gap: 9px;\n      padding: 10px;\n      border-radius: 11px;\n      background: #fff8e7;\n    }\n\n    .guide-popover strong {\n      color: #725214;\n      font-size: 13px;\n    }\n\n    .guide-popover span {\n      color: #4f5b6d;\n      font-size: 13px;\n      line-height: 1.5;\n    }\n\n    .guide-popover p {\n      margin: 12px 0 0;\n      padding: 11px;\n      border-radius: 11px;\n      background: #eef4ff;\n      color: #3f5f9d;\n      font-size: 12px;\n      font-weight: 800;\n      line-height: 1.55;\n    }\n\n    .request-card-actions {\n      display: flex;\n      justify-content: flex-end;\n      margin-top: auto;\n      padding-top: 12px;\n    }\n\n    .request-cancel-button {\n      min-height: 37px;\n      padding: 0 13px;\n      border: 1px solid #edc5c0;\n      border-radius: 10px;\n      background: #fff5f3;\n      color: #a2463e;\n      font-size: 12px;\n      font-weight: 900;\n      cursor: pointer;\n    }\n\n    .request-cancel-button:hover {\n      background: #ffe6e2;\n    }\n\n    .request-cancel-button.restore {\n      border-color: #b9d9c2;\n      background: #edf8f0;\n      color: #27643b;\n    }\n\n    .request-cancel-button.restore:hover {\n      background: #dff1e4;\n    }\n\n    .request-cancel-button:disabled {\n      border-color: #d9dee7;\n      background: #f1f3f6;\n      color: #8892a2;\n      cursor: default;\n    }\n\n    .cancel-modal {\n      position: fixed;\n      inset: 0;\n      z-index: 120;\n      display: none;\n      padding: 22px;\n      align-items: center;\n      justify-content: center;\n      background: rgba(25, 34, 51, 0.58);\n      backdrop-filter: blur(5px);\n    }\n\n    .cancel-modal.show {\n      display: flex;\n    }\n\n    .cancel-dialog {\n      width: min(430px, 100%);\n      padding: 25px;\n      border-radius: 22px;\n      background: #fff;\n      box-shadow: 0 22px 60px rgba(24, 34, 52, 0.3);\n      text-align: center;\n      animation: notice-in 0.2s ease-out;\n    }\n\n    .cancel-icon {\n      margin-bottom: 8px;\n      font-size: 42px;\n    }\n\n    .cancel-dialog h2 {\n      margin: 0;\n      color: var(--ink);\n      font-size: 22px;\n      letter-spacing: -0.035em;\n    }\n\n    .cancel-book {\n      margin: 12px 0 0;\n      color: var(--blue-dark);\n      font-size: 16px;\n      font-weight: 900;\n      line-height: 1.5;\n    }\n\n    .cancel-description {\n      margin: 10px 0 18px;\n      color: var(--sub);\n      font-size: 13px;\n      line-height: 1.6;\n    }\n\n    .cancel-actions {\n      display: grid;\n      grid-template-columns: 1fr 1fr;\n      gap: 9px;\n    }\n\n    .cancel-actions button {\n      min-height: 48px;\n      border-radius: 12px;\n      font-size: 14px;\n      font-weight: 900;\n      cursor: pointer;\n    }\n\n    .cancel-keep {\n      border: 1px solid #d8e0ed;\n      background: #fff;\n      color: #59677e;\n    }\n\n    .cancel-confirm {\n      border: 0;\n      background: #d95d51;\n      color: #fff;\n    }\n\n    .cancel-confirm:disabled {\n      opacity: 0.58;\n      cursor: wait;\n    }\n\n    @media (max-width: 620px) {\n      .guide-hover-trigger {\n        min-height: 46px;\n        padding: 0 5px;\n        font-size: 13px;\n      }\n\n      .guide-popover {\n        position: fixed;\n        top: 82px;\n        right: 17px;\n        left: 17px;\n        width: auto;\n      }\n\n      .guide-popover::before {\n        display: none;\n      }\n\n      .guide-popover li {\n        grid-template-columns: 1fr;\n        gap: 3px;\n      }\n    }\n\n\n    .request-item {\n      cursor: pointer;\n      transition:\n        border-color 0.18s ease,\n        box-shadow 0.18s ease,\n        transform 0.18s ease;\n    }\n\n    .request-item:hover,\n    .request-item:focus-visible,\n    .request-item.expanded {\n      border-color: #aebfe8;\n      box-shadow: 0 12px 28px rgba(48, 78, 143, 0.13);\n    }\n\n    .request-biblio {\n      display: none;\n      gap: 8px;\n      margin-top: 14px;\n      padding: 14px;\n      border: 1px solid #dfe6f5;\n      border-radius: 14px;\n      background: #f8faff;\n    }\n\n    .request-item.expanded .request-biblio {\n      display: grid;\n    }\n\n    .request-biblio-row {\n      display: grid;\n      grid-template-columns: 84px minmax(0, 1fr);\n      gap: 10px;\n      align-items: start;\n      color: #34445e;\n      font-size: 15px;\n      line-height: 1.5;\n    }\n\n    .request-biblio-row strong {\n      color: #173567;\n      font-weight: 800;\n    }\n\n    .request-biblio-row span,\n    .request-biblio-row a {\n      min-width: 0;\n      overflow-wrap: anywhere;\n    }\n\n    .request-biblio-row a {\n      color: #315fbd;\n      font-weight: 800;\n    }\n\n    .request-biblio-empty {\n      margin: 0;\n      color: #68738a;\n      font-size: 14px;\n    }\n\n    .request-card-actions {\n      display: flex;\n      gap: 10px;\n      margin-top: 14px;\n    }\n\n    .request-detail-button,\n    .request-cancel-button {\n      min-height: 42px;\n      padding: 0 15px;\n      border-radius: 12px;\n      font-size: 14px;\n      font-weight: 850;\n    }\n\n    .request-detail-button {\n      flex: 1;\n      border: 1px solid #cdd9f2;\n      background: #eef3ff;\n      color: #234b98;\n    }\n\n    .request-cancel-button {\n      border: 1px solid #ebc6c2;\n      background: #fff5f4;\n      color: #a03e37;\n    }\n\n    .request-cancel-button:disabled {\n      border-color: #e1e4ea;\n      background: #f4f5f7;\n      color: #8a909b;\n      cursor: default;\n    }\n\n    @media (max-width: 560px) {\n      .request-biblio-row {\n        grid-template-columns: 74px minmax(0, 1fr);\n        font-size: 14px;\n      }\n\n      .request-card-actions {\n        flex-direction: column;\n      }\n\n      .request-detail-button,\n      .request-cancel-button {\n        width: 100%;\n      }\n    }\n\n\n  \n    /* HR_WEB_NOTICE_TICKER */\n    .hr-notice-ticker[hidden] {\n      display: none !important;\n    }\n\n    .hr-notice-ticker {\n      width: min(1160px, calc(100% - 34px));\n      margin: 14px auto 0;\n      overflow: hidden;\n      border: 1px solid #ead39a;\n      border-radius: 15px;\n      background: linear-gradient(90deg, #fff5d7 0%, #fffaf0 100%);\n      color: #6f4b12;\n      box-shadow: 0 8px 24px rgba(102, 78, 29, 0.09);\n    }\n\n    .hr-notice-ticker.urgent {\n      border-color: #efb2aa;\n      background: linear-gradient(90deg, #fff0ed 0%, #fff8f4 100%);\n      color: #9b332d;\n    }\n\n    .hr-notice-inner {\n      display: flex;\n      min-height: 52px;\n      align-items: stretch;\n    }\n\n    .hr-notice-label {\n      display: flex;\n      flex: 0 0 auto;\n      align-items: center;\n      justify-content: center;\n      padding: 0 20px;\n      background: #d9a13b;\n      color: #fff;\n      font-size: 15px;\n      font-weight: 900;\n      letter-spacing: .02em;\n    }\n\n    .hr-notice-ticker.urgent .hr-notice-label {\n      background: #c44c43;\n    }\n\n    .hr-notice-window {\n      min-width: 0;\n      flex: 1;\n      overflow: hidden;\n      mask-image: linear-gradient(90deg, transparent, #000 4%, #000 96%, transparent);\n      -webkit-mask-image: linear-gradient(90deg, transparent, #000 4%, #000 96%, transparent);\n    }\n\n    .hr-notice-track {\n      display: inline-block;\n      min-width: max-content;\n      padding: 14px 0 13px 100%;\n      white-space: nowrap;\n      font-size: 17px;\n      font-weight: 850;\n      line-height: 1.45;\n      animation: hrTickerMove var(--ticker-duration, 26s) linear infinite;\n      will-change: transform;\n    }\n\n    .hr-notice-ticker:hover .hr-notice-track,\n    .hr-notice-ticker:focus-within .hr-notice-track {\n      animation-play-state: paused;\n    }\n\n    @keyframes hrTickerMove {\n      from { transform: translateX(0); }\n      to { transform: translateX(-100%); }\n    }\n\n    @media (max-width: 640px) {\n      .hr-notice-ticker {\n        width: calc(100% - 20px);\n        margin-top: 10px;\n        border-radius: 12px;\n      }\n\n      .hr-notice-label {\n        padding: 0 13px;\n        font-size: 13px;\n      }\n\n      .hr-notice-track {\n        font-size: 15px;\n      }\n    }\n\n    @media (prefers-reduced-motion: reduce) {\n      .hr-notice-track {\n        min-width: 100%;\n        padding-left: 16px;\n        padding-right: 16px;\n        white-space: normal;\n        animation: none;\n      }\n    }\n\n  \n    /* HR_WEB_NOTICE_TICKER_V2 */\n    .hr-notice-ticker {\n      width: min(920px, calc(100% - 24px));\n      margin: 14px auto 12px;\n      border: 1px solid rgba(205, 157, 65, 0.42);\n      border-radius: 18px;\n      background: rgba(255, 253, 248, 0.98);\n      color: #694713;\n      box-shadow: 0 10px 28px rgba(45, 70, 58, 0.11);\n    }\n\n    .hr-notice-ticker + .view-tabs {\n      margin-top: 0;\n    }\n\n    .hr-notice-inner {\n      min-height: 48px;\n      align-items: stretch;\n    }\n\n    .hr-notice-label {\n      gap: 5px;\n      padding: 0 16px;\n      border-radius: 17px 0 0 17px;\n      background: linear-gradient(135deg, #dda842 0%, #c98d2e 100%);\n      font-size: 14px;\n      letter-spacing: 0;\n    }\n\n    .hr-notice-window {\n      display: flex;\n      align-items: center;\n    }\n\n    .hr-notice-track {\n      padding: 12px 0 11px 100%;\n      font-size: 15px;\n      font-weight: 800;\n      line-height: 1.45;\n    }\n\n    .hr-notice-ticker.urgent {\n      border-color: rgba(194, 76, 67, 0.42);\n      background: #fff8f5;\n      color: #95352f;\n    }\n\n    .hr-notice-ticker.urgent .hr-notice-label {\n      background: linear-gradient(135deg, #d45b52 0%, #b63d36 100%);\n    }\n\n    @media (max-width: 640px) {\n      .hr-notice-ticker {\n        width: calc(100% - 20px);\n        margin: 10px auto 10px;\n        border-radius: 14px;\n      }\n\n      .hr-notice-ticker + .view-tabs {\n        margin-top: 0;\n      }\n\n      .hr-notice-label {\n        padding: 0 12px;\n        border-radius: 13px 0 0 13px;\n        font-size: 12px;\n      }\n\n      .hr-notice-track {\n        padding-top: 11px;\n        padding-bottom: 10px;\n        font-size: 14px;\n      }\n    }\n\n    /* HR_WEB_NOTICE_TICKER_HEIGHT_FIX_V3\n       모션 줄이기 환경에서 기존 규칙과 V2 규칙이 겹치면\n       공지 문구가 한 글자씩 줄바꿈되며 전광판이 길어지는 문제를 방지합니다. */\n    .hr-notice-ticker {\n      height: 50px;\n      min-height: 50px;\n      max-height: 50px;\n      overflow: hidden;\n    }\n\n    .hr-notice-inner,\n    .hr-notice-label,\n    .hr-notice-window {\n      height: 48px;\n      min-height: 48px;\n      max-height: 48px;\n    }\n\n    .hr-notice-inner,\n    .hr-notice-window {\n      overflow: hidden;\n    }\n\n    .hr-notice-label {\n      align-self: flex-start;\n      white-space: nowrap;\n    }\n\n    .hr-notice-track {\n      height: 48px;\n      max-height: 48px;\n      overflow: hidden;\n      box-sizing: border-box;\n      white-space: nowrap;\n    }\n\n    @media (prefers-reduced-motion: reduce) {\n      .hr-notice-track {\n        width: auto;\n        min-width: max-content;\n        padding: 12px 0 11px 100%;\n        overflow: hidden;\n        white-space: nowrap;\n        text-overflow: clip;\n        animation: hrTickerMove var(--ticker-duration, 26s) linear infinite;\n      }\n    }\n        .book-status-badge{margin:8px 0 0;padding:7px 11px;border-radius:10px;background:#fff4dc;color:#7a5410;font-size:13px;font-weight:800;line-height:1.45}\n      .book-status-badge.owned{background:#e8eef8;color:#35527a}\n</style>\n\n  <style id=\"visualRefreshV20\">\n    :root {\n      --blue: #2f6653;\n      --blue-dark: #234d40;\n      --blue-soft: #e8f2ec;\n      --ink: #25352f;\n      --sub: #68766f;\n      --line: #dbe4dc;\n      --paper: #fffdf8;\n      --background: #f4f3eb;\n      --gold: #d3a248;\n      --gold-dark: #ad762b;\n      --cream: #fbf4e5;\n      --forest: #2c5949;\n      --forest-dark: #203f35;\n    }\n\n    body {\n      min-height: 100vh;\n      padding: 28px 18px 64px;\n      background:\n        radial-gradient(circle at 7% 5%, rgba(234, 213, 165, 0.52), transparent 27%),\n        radial-gradient(circle at 96% 13%, rgba(188, 215, 198, 0.58), transparent 26%),\n        linear-gradient(180deg, #faf7ef 0%, #eef3ee 100%);\n      color: var(--ink);\n    }\n\n    .header {\n      position: relative;\n      width: 100%;\n      max-width: 820px;\n      margin: 0 auto;\n      padding: 34px 38px 50px;\n      overflow: hidden;\n      border: 1px solid rgba(255, 255, 255, 0.32);\n      border-radius: 30px;\n      background:\n        linear-gradient(135deg, rgba(255,255,255,0.06), transparent 44%),\n        linear-gradient(135deg, #23483b 0%, #386a56 100%);\n      color: #fff;\n      box-shadow: 0 20px 45px rgba(38, 76, 62, 0.2);\n    }\n\n    .header::before {\n      position: absolute;\n      top: -85px;\n      right: -45px;\n      width: 230px;\n      height: 230px;\n      border: 42px solid rgba(255, 255, 255, 0.055);\n      border-radius: 50%;\n      content: \"\";\n    }\n\n    .header::after {\n      position: absolute;\n      right: 34px;\n      bottom: 18px;\n      content: \"📚\";\n      font-size: 78px;\n      line-height: 1;\n      opacity: 0.2;\n      transform: rotate(-4deg);\n    }\n\n    .header > * {\n      position: relative;\n      z-index: 1;\n    }\n\n    .eyebrow {\n      margin-bottom: 10px;\n      color: #f2d58f;\n      font-size: 12px;\n      letter-spacing: 0.18em;\n    }\n\n    h1 {\n      color: #fff;\n      font-size: clamp(30px, 5vw, 40px);\n      letter-spacing: -0.045em;\n    }\n\n    .description {\n      max-width: 500px;\n      margin-top: 12px;\n      color: #e6efe9;\n      font-size: 15px;\n      line-height: 1.7;\n    }\n\n    .view-tabs {\n      top: 10px;\n      max-width: 920px;\n      margin: -25px auto 25px;\n      padding: 7px;\n      grid-template-columns: repeat(4, minmax(0, 1fr));\n      gap: 7px;\n      border: 1px solid rgba(49, 91, 73, 0.14);\n      border-radius: 20px;\n      background: rgba(255, 253, 248, 0.96);\n      box-shadow: 0 14px 34px rgba(52, 75, 63, 0.16);\n    }\n\n    .view-tab,\n    .guide-hover-trigger {\n      min-height: 51px;\n      border-radius: 14px;\n      color: #52655c;\n      font-size: 15px;\n    }\n\n    .view-tab:hover {\n      background: #edf4ef;\n    }\n\n    .view-tab.active {\n      background: linear-gradient(135deg, #d8ae58 0%, #bd8433 100%);\n      color: #fff;\n      box-shadow: 0 7px 18px rgba(177, 121, 43, 0.26);\n    }\n\n    .guide-hover-trigger {\n      background: #edf4ef;\n      color: #315c4c;\n    }\n\n    .guide-hover-trigger:hover,\n    .guide-hover-trigger:focus {\n      background: #dfede4;\n      color: #244b3d;\n    }\n\n    .guide-popover {\n      border-color: #d7e1d8;\n      border-radius: 20px;\n      background: #fffdf8;\n      box-shadow: 0 20px 48px rgba(37, 67, 54, 0.2);\n    }\n\n    .guide-popover h2 {\n      color: #294f41;\n    }\n\n    .card,\n    .bs-section,\n    .requests-section {\n      max-width: 820px;\n      border: 1px solid rgba(56, 92, 76, 0.12);\n      border-radius: 26px;\n      background: rgba(255, 253, 248, 0.96);\n      box-shadow: 0 15px 38px rgba(46, 72, 59, 0.1);\n    }\n\n    .card {\n      padding: 24px;\n    }\n\n    .form-intro {\n      display: flex;\n      align-items: center;\n      gap: 14px;\n      margin: 0 0 20px;\n      padding: 0 2px 18px;\n      border-bottom: 1px dashed #d9e2da;\n    }\n\n    .form-intro-icon {\n      display: grid;\n      flex: 0 0 48px;\n      width: 48px;\n      height: 48px;\n      place-items: center;\n      border-radius: 16px;\n      background: #e8f2ec;\n      font-size: 25px;\n    }\n\n    .form-intro h2 {\n      margin: 0;\n      color: #284b3e;\n      font-size: 20px;\n      letter-spacing: -0.03em;\n    }\n\n    .form-intro p {\n      margin: 4px 0 0;\n      color: var(--sub);\n      font-size: 13px;\n      line-height: 1.5;\n    }\n\n    label {\n      color: #314e43;\n      font-size: 16px;\n    }\n\n    input {\n      height: 57px;\n      border-color: #d6dfd7;\n      border-radius: 15px;\n      background: #fff;\n      color: #263831;\n      font-size: 17px;\n    }\n\n    input:focus {\n      border-color: #4f806d;\n      box-shadow: 0 0 0 4px rgba(68, 121, 98, 0.13);\n    }\n\n    .submit-button {\n      min-height: 58px;\n      border-radius: 15px;\n      background: linear-gradient(135deg, #356b57 0%, #274f42 100%);\n      box-shadow: 0 9px 20px rgba(39, 79, 66, 0.2);\n    }\n\n    .submit-button:hover {\n      background: linear-gradient(135deg, #2d604e 0%, #1f4338 100%);\n    }\n\n    .bs-section {\n      margin-top: 20px;\n      padding: 23px;\n    }\n\n    .bs-title {\n      color: #294d40;\n      font-size: 23px;\n    }\n\n    .bs-badge {\n      background: #f6ead1;\n      color: #90621f;\n    }\n\n    .bs-tabs {\n      gap: 8px;\n      padding: 6px;\n      border: 1px solid #dce5dd;\n      border-radius: 16px;\n      background: #edf3ee;\n    }\n\n    .bs-tab {\n      min-height: 45px;\n      border-radius: 12px;\n      color: #5b6c64;\n    }\n\n    .bs-tab.active {\n      color: #274e40;\n      box-shadow: 0 5px 13px rgba(50, 83, 68, 0.13);\n    }\n\n    .bs-card {\n      border-color: #e3e4d9;\n      border-radius: 20px;\n      background: #fffefb;\n      box-shadow: 0 8px 20px rgba(70, 76, 61, 0.08);\n    }\n\n    .bs-card:hover {\n      border-color: #c8d7cb;\n      box-shadow: 0 13px 28px rgba(51, 79, 65, 0.15);\n    }\n\n    .bs-rank {\n      background: #d4a64f;\n      box-shadow: 0 5px 12px rgba(179, 126, 42, 0.25);\n    }\n\n    .bs-request {\n      background: #356b57;\n    }\n\n    .bs-request:hover {\n      background: #284f42;\n    }\n\n    .bs-link,\n    .bs-more {\n      border-color: #cadbcf;\n      background: #edf5ef;\n      color: #2e5c4b;\n    }\n\n    .bs-more:hover {\n      background: #dfece3;\n    }\n\n    .bs-arrow {\n      border-color: rgba(65, 100, 83, 0.2);\n      background: rgba(255, 253, 248, 0.94);\n      color: #315e4d;\n      box-shadow: 0 8px 22px rgba(47, 77, 62, 0.16);\n    }\n\n    .requests-section {\n      padding: 25px;\n    }\n\n    .panel-eyebrow {\n      color: #bf842f;\n    }\n\n    .panel-heading h2 {\n      color: #294e40;\n    }\n\n    .request-item {\n      border-color: #dde5de;\n      border-radius: 20px;\n      background: #fffefb;\n    }\n\n    .request-item:hover,\n    .request-item:focus-visible,\n    .request-item.expanded {\n      border-color: #acc6b5;\n      box-shadow: 0 13px 30px rgba(43, 78, 61, 0.14);\n    }\n\n    .request-biblio {\n      border-color: #dce6de;\n      background: #f5f8f5;\n    }\n\n    .refresh-button,\n    .request-detail-button {\n      border-color: #cadbcf;\n      background: #edf5ef;\n      color: #2d5a49;\n    }\n\n    .cart-bar {\n      max-width: 820px;\n      border-color: #d7e1d8;\n      border-radius: 20px;\n      background: rgba(255, 253, 248, 0.97);\n      box-shadow: 0 15px 36px rgba(44, 74, 59, 0.18);\n    }\n\n    .cart-submit {\n      background: #356b57;\n    }\n\n    .footer {\n      color: #728078;\n    }\n\n    .creator-watermark {\n      color: rgba(45, 77, 63, 0.55);\n    }\n\n    @media (min-width: 760px) {\n      #requestForm {\n        display: grid;\n        grid-template-columns: 0.82fr 1.18fr 170px;\n        gap: 14px;\n        align-items: start;\n      }\n\n      #requestForm .field {\n        margin-bottom: 0;\n      }\n\n      #requestForm .help {\n        min-height: 39px;\n      }\n\n      /* 입력칸과 「책 검색하기」 버튼의 윗선·높이를 맞춥니다. (이름표 한 줄 = 24px + 8px) */\n      #requestForm .field > label {\n        line-height: 24px;\n        margin-bottom: 8px;\n      }\n\n      #requestForm .submit-button {\n        width: 170px;\n        height: 57px;\n        min-height: 57px;\n        margin-top: 32px;\n        margin-bottom: 0;\n      }\n    }\n\n    @media (max-width: 620px) {\n      body {\n        padding: 14px 12px 54px;\n      }\n\n      .header {\n        padding: 27px 24px 44px;\n        border-radius: 24px;\n      }\n\n      .header::after {\n        right: 17px;\n        bottom: 15px;\n        font-size: 55px;\n      }\n\n      .view-tabs {\n        margin: -22px 8px 20px;\n        grid-template-columns: 1fr 1fr 0.78fr;\n        border-radius: 17px;\n      }\n\n      .view-tab,\n      .guide-hover-trigger {\n        min-height: 48px;\n        padding: 0 5px;\n        font-size: 13px;\n      }\n\n      .card,\n      .bs-section,\n      .requests-section {\n        border-radius: 22px;\n      }\n\n      .card,\n      .bs-section {\n        padding: 18px;\n      }\n\n      .form-intro {\n        align-items: flex-start;\n      }\n    }\n  </style>\n\n\n  <style id=\"requestCoverV22\">\n    .header {\n      min-height: 0;\n    }\n\n    .request-biblio-cover {\n      display: flex;\n      grid-column: 1 / -1;\n      justify-content: center;\n      margin: 0 0 5px;\n      padding: 5px 0 12px;\n      border-bottom: 1px dashed #d7e2d9;\n    }\n\n    .request-biblio-cover img {\n      width: 112px;\n      height: 156px;\n      object-fit: cover;\n      border: 1px solid #d7dfd8;\n      border-radius: 12px;\n      background: #fff;\n      box-shadow: 0 9px 22px rgba(43, 67, 55, 0.17);\n    }\n\n    @media (max-width: 720px) {\n      .header {\n        min-height: 0;\n      }\n\n      .request-biblio-cover img {\n        width: 100px;\n        height: 140px;\n      }\n    }\n    .school-contact-footer {\n      width: min(1180px, calc(100% - 32px));\n      margin: 42px auto 118px;\n      padding: 24px 28px;\n      border-radius: 24px;\n      background: linear-gradient(135deg, #263b33 0%, #334d42 100%);\n      color: #ffffff;\n      box-shadow: 0 16px 38px rgba(27, 54, 43, 0.16);\n    }\n\n    .school-contact-row {\n      display: flex;\n      flex-wrap: wrap;\n      align-items: center;\n      gap: 12px 24px;\n    }\n\n    .school-contact-badge {\n      display: inline-flex;\n      align-items: center;\n      min-height: 38px;\n      padding: 0 15px;\n      border-radius: 999px;\n      background: #e0ad4d;\n      color: #253a32;\n      font-size: 14px;\n      font-weight: 900;\n      white-space: nowrap;\n    }\n\n    .school-contact-detail {\n      display: inline-flex;\n      align-items: baseline;\n      gap: 8px;\n      font-size: 16px;\n      line-height: 1.6;\n    }\n\n    .school-contact-detail strong {\n      color: #f4d58c;\n      font-size: 13px;\n      font-weight: 900;\n      letter-spacing: .04em;\n    }\n\n    .site-credit {\n      margin-top: 16px;\n      padding-top: 15px;\n      border-top: 1px solid rgba(255, 255, 255, 0.16);\n      color: rgba(255, 255, 255, 0.68);\n      font-size: 12px;\n      text-align: right;\n    }\n\n    @media (max-width: 640px) {\n      .school-contact-footer {\n        padding: 22px 20px;\n      }\n\n      .school-contact-row,\n      .school-contact-detail {\n        align-items: flex-start;\n      }\n\n      .school-contact-row {\n        flex-direction: column;\n        gap: 10px;\n      }\n\n      .site-credit {\n        text-align: left;\n      }\n    }\n  \n    .upload-section {\n      width: 100%;\n      max-width: 920px;\n      margin: 18px auto 0;\n      padding: 30px;\n      border: 1px solid var(--line);\n      border-radius: 26px;\n      background: rgba(255, 253, 248, 0.98);\n      box-shadow: 0 18px 42px rgba(35, 77, 64, 0.12);\n    }\n\n    .upload-heading {\n      display: flex;\n      align-items: flex-start;\n      justify-content: space-between;\n      gap: 18px;\n      margin-bottom: 22px;\n    }\n\n    .upload-heading h2 {\n      margin: 3px 0 7px;\n      color: var(--forest-dark);\n      font-size: 28px;\n      line-height: 1.25;\n    }\n\n    .upload-heading p {\n      margin: 0;\n      color: var(--sub);\n      font-size: 15px;\n      line-height: 1.7;\n    }\n\n    .upload-badge {\n      flex: 0 0 auto;\n      padding: 9px 13px;\n      border-radius: 999px;\n      background: #fbf0d7;\n      color: #8a6321;\n      font-size: 13px;\n      font-weight: 900;\n    }\n\n    .upload-card {\n      padding: 22px;\n      border: 1px solid #d6e1d9;\n      border-radius: 20px;\n      background: #fff;\n    }\n\n    .upload-dropzone {\n      display: flex;\n      min-height: 185px;\n      align-items: center;\n      justify-content: center;\n      padding: 28px;\n      border: 2px dashed #9ebbaa;\n      border-radius: 18px;\n      background: #f3f8f4;\n      text-align: center;\n      cursor: pointer;\n      transition: border-color 0.18s, background 0.18s, transform 0.18s;\n    }\n\n    .upload-dropzone:hover,\n    .upload-dropzone.dragging {\n      border-color: var(--forest);\n      background: #e8f3ec;\n      transform: translateY(-1px);\n    }\n\n    .upload-dropzone input {\n      position: absolute;\n      width: 1px;\n      height: 1px;\n      overflow: hidden;\n      opacity: 0;\n      pointer-events: none;\n    }\n\n    .upload-icon {\n      display: block;\n      margin-bottom: 8px;\n      font-size: 44px;\n    }\n\n    .upload-dropzone strong {\n      display: block;\n      color: var(--forest-dark);\n      font-size: 20px;\n      line-height: 1.4;\n    }\n\n    .upload-dropzone span {\n      display: block;\n      margin-top: 8px;\n      color: var(--sub);\n      font-size: 14px;\n      line-height: 1.6;\n    }\n\n    .upload-format-note {\n      margin: 14px 2px 0;\n      color: #6d786f;\n      font-size: 13px;\n      line-height: 1.65;\n    }\n      .upload-format-note span{display:block;word-break:keep-all}\n\n\n    .upload-status {\n      display: none;\n      margin-top: 16px;\n      padding: 14px 16px;\n      border-radius: 14px;\n      font-size: 14px;\n      font-weight: 800;\n      line-height: 1.6;\n    }\n\n    .upload-status.show {\n      display: block;\n    }\n\n    .upload-status.info {\n      background: #edf4fb;\n      color: #315f86;\n    }\n\n    .upload-status.success {\n      background: #e9f5ed;\n      color: #286348;\n    }\n\n    .upload-status.error {\n      background: #fff0ef;\n      color: #a9433c;\n    }\n\n    .upload-preview {\n      display: none;\n      margin-top: 18px;\n    }\n\n    .upload-preview.show {\n      display: block;\n    }\n\n    .upload-preview-head {\n      display: flex;\n      align-items: center;\n      justify-content: space-between;\n      gap: 12px;\n      margin-bottom: 10px;\n    }\n\n    .upload-preview-head strong {\n      color: var(--forest-dark);\n      font-size: 17px;\n    }\n\n    .upload-preview-head span {\n      color: var(--sub);\n      font-size: 13px;\n    }\n\n    .upload-table-wrap {\n      max-height: 560px;\n      overflow: auto;\n      scrollbar-gutter: stable;\n      border: 1px solid #d9e2dc;\n      border-radius: 14px;\n      background: #fff;\n    }\n\n    .upload-table {\n      width: 100%;\n      min-width: 760px;\n      border-collapse: collapse;\n      font-size: 13px;\n    }\n\n    .upload-table th,\n    .upload-table td {\n      padding: 11px 12px;\n      border-bottom: 1px solid #e4e9e5;\n      text-align: left;\n      vertical-align: middle;\n    }\n\n    .upload-table th {\n      position: sticky;\n      top: 0;\n      z-index: 1;\n      background: #456354;\n      color: #fff;\n      font-weight: 900;\n      white-space: nowrap;\n    }\n\n    .upload-table td:nth-child(1) {\n      width: 44px;\n      color: #7a847d;\n      text-align: center;\n    }\n\n    .upload-table td:nth-child(2) {\n      min-width: 230px;\n      color: #23372f;\n      font-weight: 800;\n    }\n\n    .upload-actions {\n      display: flex;\n      justify-content: flex-end;\n      gap: 10px;\n      margin-top: 18px;\n    }\n\n    .upload-reset,\n    .upload-save {\n      min-height: 50px;\n      padding: 0 22px;\n      border-radius: 13px;\n      font-size: 15px;\n      font-weight: 900;\n      cursor: pointer;\n    }\n\n    .upload-reset {\n      border: 1px solid #d6dfd8;\n      background: #fff;\n      color: #65716a;\n    }\n\n    .upload-save {\n      border: 0;\n      background: var(--forest);\n      color: #fff;\n    }\n\n    .upload-save:disabled,\n    .upload-reset:disabled {\n      cursor: default;\n      opacity: 0.48;\n    }\n\n    @media (max-width: 720px) {\n      .view-tabs {\n        grid-template-columns: repeat(2, minmax(0, 1fr));\n      }\n\n      .upload-section {\n        padding: 20px 14px;\n      }\n\n      .upload-heading {\n        display: block;\n      }\n\n      .upload-badge {\n        display: inline-block;\n        margin-top: 12px;\n      }\n\n      .upload-actions {\n        flex-direction: column;\n      }\n\n      .upload-reset,\n      .upload-save {\n        width: 100%;\n      }\n    }\n\n\n    /* 모든 주요 영역의 가로폭을 하나로 통일 */\n    .header,\n    .view-tabs,\n    .card,\n    .bs-section,\n    .requests-section,\n    .guide-section,\n    .cart-bar,\n    .footer,\n    .upload-section,\n    .school-contact-footer {\n      width: calc(100% - 36px);\n      max-width: 920px;\n      margin-left: auto;\n      margin-right: auto;\n    }\n\n@media (max-width: 700px) {\n  .header,\n  .view-tabs,\n  .card,\n  .bs-section,\n  .requests-section,\n  .guide-section,\n  .cart-bar,\n  .footer,\n  .upload-section,\n  .school-contact-footer {\n    width: calc(100% - 20px);\n  }\n}\n\n\n    /* ===== 교직원 / 학부모 구분 ===== */\n    .audience-gate{position:fixed;inset:0;z-index:9999;display:flex;align-items:center;justify-content:center;padding:28px 20px;overflow:auto;background:radial-gradient(1200px 700px at 20% 0%,#3a7360 0%,transparent 60%),linear-gradient(160deg,var(--forest),var(--forest-dark))}\n    .audience-gate[hidden]{display:none}\n    .audience-card{width:100%;max-width:980px;text-align:center}\n    .audience-card .eyebrow{margin:0 0 6px;color:var(--gold);letter-spacing:.28em;font-weight:800;font-size:14px}\n    .audience-card h2{margin:6px 0 10px;color:#fff;font-size:clamp(30px,6vw,52px);line-height:1.2}\n    .audience-card p.audience-desc{color:rgba(255,255,255,.82);margin:0 0 30px;line-height:1.6;font-size:clamp(15px,2.4vw,19px)}\n    .audience-choices{display:grid;grid-template-columns:repeat(auto-fit,minmax(260px,1fr));gap:20px}\n    .audience-btn{display:flex;flex-direction:column;align-items:center;justify-content:center;gap:8px;min-height:230px;padding:26px 18px;border:0;border-radius:26px;font-size:clamp(26px,4vw,38px);font-weight:900;cursor:pointer;color:#fff;text-decoration:none;background:linear-gradient(135deg,#3f8068,#1f4f3f);box-shadow:0 16px 40px rgba(0,0,0,.3);transition:transform .15s ease,box-shadow .15s ease}\n    .audience-btn.parent{background:linear-gradient(135deg,var(--gold),var(--gold-dark))}\n    .audience-btn:hover,.audience-btn:focus-visible{transform:translateY(-4px);box-shadow:0 22px 50px rgba(0,0,0,.38);outline:none}\n    .audience-btn .audience-icon{font-size:clamp(46px,7vw,72px);line-height:1}\n    .audience-btn span.audience-sub{font-size:clamp(14px,2.2vw,18px);font-weight:600;opacity:.95}\n    .code-card{max-width:420px;margin:0 auto}\n    .code-input{width:100%;height:60px;box-sizing:border-box;padding:0 16px;border:0;border-radius:16px;font-size:24px;text-align:center;letter-spacing:.2em;background:#fff;color:var(--ink)}\n    .code-message{margin:12px 0 0;color:#ffd9d2;font-weight:700;font-size:15px}\n    .code-message[hidden]{display:none}\n    .code-submit{margin-top:16px;width:100%;height:56px;border:0;border-radius:16px;background:linear-gradient(135deg,var(--gold),var(--gold-dark));color:#fff;font-size:20px;font-weight:900;cursor:pointer}\n    .code-submit[disabled]{opacity:.6;cursor:wait}\n    .code-back{display:inline-block;margin-top:18px;color:rgba(255,255,255,.8);font-size:14px;text-decoration:underline}\n    .audience-bar{display:flex;align-items:center;justify-content:center;gap:10px;margin:12px auto 0;font-size:14px}\n    .audience-bar[hidden]{display:none}\n    .audience-badge{padding:6px 15px;border-radius:999px;background:var(--blue-soft);color:var(--forest-dark);font-weight:800}\n    .audience-bar .audience-switch{border:1px solid var(--line);background:#fff;border-radius:999px;padding:6px 13px;cursor:pointer;color:var(--sub);font-size:13px;text-decoration:none}\n    .audience-bar .audience-switch[hidden]{display:none}\n    .parent-fields input,.parent-fields select,#pName{width:100%;height:57px;box-sizing:border-box;padding:0 14px;border:1px solid var(--line);border-radius:15px;font-size:17px;line-height:1.2;background:#fff;color:var(--ink)}\n    .parent-fields select{-webkit-appearance:none;appearance:none;padding-right:38px;background:#fff url(\"data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='14' height='9' viewBox='0 0 14 9'%3E%3Cpath d='M1 1.5l6 6 6-6' fill='none' stroke='%2368766f' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'/%3E%3C/svg%3E\") no-repeat right 14px center}\n    .parent-fields label{font-size:13px;color:var(--sub);display:block}\n    .parent-fields select{text-align:center}\n\n    /* ===== 학부모/교직원 화면 정리 ===== */\n    .parent-fields{display:grid;grid-template-columns:repeat(3,1fr);gap:10px}\n    .parent-fields .pname{grid-column:1/-1}\n    .parent-quota{margin:12px 0 4px;padding:13px 16px;border-radius:15px;background:#fbf1d4;border:1px solid #ecd48f;color:#5b410a;font-size:15px;line-height:1.55;text-align:center}\n    .parent-quota strong{font-size:20px;color:#2e5b4f}\n    .parent-quota.full{background:#fdecea;border-color:#f0b8b0;color:#8a2a1f}\n    .parent-quota.full strong{color:#b3261e}\n    .parent-quota[hidden]{display:none}\n    .parent-fields label{margin:0 0 5px}\n    .parent-hint{font-weight:400;font-size:13px;color:var(--sub)}\n    #parentApplicantField[hidden],#staffApplicantField[hidden]{display:none}\n    .header .audience-bar{position:absolute;top:20px;right:22px;margin:0;gap:8px;z-index:2}\n    .header .audience-badge{background:rgba(255,255,255,.16);color:#fff;border:1px solid rgba(255,255,255,.3);padding:6px 14px}\n    .header .audience-bar .audience-switch{background:transparent;color:#fff;border:1px solid rgba(255,255,255,.45)}\n    .header .audience-bar .audience-switch:hover{background:rgba(255,255,255,.14)}\n    body.aud-parent .view-tabs{grid-template-columns:repeat(3,minmax(0,1fr))}\n    @media (min-width:760px){\n      body.aud-parent #requestForm{grid-template-columns:1fr 170px}\n      /* 학부모 화면: 검색창과 「책 검색하기」 버튼의 높이를 맞춥니다. (버튼을 입력칸과 같은 줄에 둡니다.) */\n      body.aud-parent #requestForm{align-items:start}\n      body.aud-parent #requestForm label[for=\"query\"]{line-height:24px;margin-bottom:8px}\n      body.aud-parent #requestForm .submit-button{margin-top:32px;margin-bottom:0;min-height:57px;height:57px}\n      body.aud-parent #parentApplicantField{grid-column:1/-1}\n      .parent-fields{grid-template-columns:repeat(3,minmax(0,.55fr)) minmax(0,1.4fr)}\n      .parent-fields .pname{grid-column:auto}\n    }\n    @media (max-width:620px){\n      .header .audience-bar{position:static;margin:14px 0 0;justify-content:flex-start;flex-wrap:wrap}\n      .parent-fields{gap:8px}\n      .parent-fields input,.parent-fields select,#pName{padding:0 8px;font-size:16px}\n      .parent-fields select{padding-right:24px;background-position:right 8px center;text-align-last:center}\n    }\n\n    /* ===== 신청 방법: 가운데 큰 안내창 ===== */\n    .guide-hover:hover .guide-popover,\n    .guide-hover:focus-within .guide-popover{opacity:0;visibility:hidden;pointer-events:none}\n    .guide-popover{position:fixed;top:50%;left:50%;right:auto;z-index:10001;width:min(620px,calc(100vw - 28px));max-height:calc(100vh - 40px);padding:26px 22px 22px;border:0;border-radius:26px;background:var(--paper);box-shadow:0 24px 70px rgba(0,0,0,.35);transform:translate(-50%,-46%)}\n    .guide-popover::before{display:none}\n    body.guide-open .guide-popover{opacity:1;visibility:visible;pointer-events:auto;transform:translate(-50%,-50%)}\n    #guideBackdrop{position:fixed;inset:0;z-index:10000;background:rgba(32,63,53,.62);opacity:0;visibility:hidden;transition:opacity .16s}\n    body.guide-open #guideBackdrop{opacity:1;visibility:visible}\n    .guide-popover h2{margin:0 0 16px;font-size:26px;color:var(--forest-dark)}\n    .guide-popover ol{gap:10px}\n    .guide-popover li{grid-template-columns:1fr;gap:4px;padding:14px 16px;border-radius:16px;background:var(--cream);border:1px solid #efe2bd}\n    .guide-popover strong{font-size:18px;color:var(--gold-dark)}\n    .guide-popover span{font-size:16px;line-height:1.6;color:var(--ink)}\n    .guide-popover p{margin:14px 0 0;padding:13px 15px;border-radius:14px;background:var(--blue-soft);color:var(--forest-dark);font-size:15px;font-weight:700;line-height:1.6}\n    .guide-popover p.guide-limit{margin-top:10px;background:#fff6da;color:#7a5a12}\n    .guide-close{display:block;width:100%;margin-top:16px;padding:15px;border:0;border-radius:15px;background:linear-gradient(135deg,var(--forest),var(--forest-dark));color:#fff;font-size:18px;font-weight:800;cursor:pointer}\n    @media (max-width:620px){ .guide-popover{padding:22px 16px 18px} .guide-popover h2{font-size:22px} .guide-popover strong{font-size:17px} .guide-popover span{font-size:15px} }\n\n    /* ===== 화면 버전 · 새 버전 안내 ===== */\n    .build-id{margin:6px auto 0;text-align:center;font-size:11px;color:#9aa69f}\n    .new-version-banner{position:fixed;top:0;left:0;right:0;z-index:10050;padding:12px 16px;background:var(--gold);color:#fff;text-align:center;font-size:16px;font-weight:800;box-shadow:0 6px 18px rgba(0,0,0,.2)}\n    .new-version-banner[hidden]{display:none}\n</style>\n\n  <script src=\"https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js\"></script>\n</head>\n\n<body>\n  <div id=\"newVersionBanner\" class=\"new-version-banner\" hidden>🔄 새 버전이 나왔습니다. 키보드 <b>F5</b>(또는 새로고침 버튼)를 눌러 주십시오.</div>\n  <div id=\"audienceGate\" class=\"audience-gate\" hidden>\n    <div class=\"audience-card\" role=\"region\" aria-labelledby=\"audienceGateTitle\">\n      <p class=\"eyebrow\">BOOK REQUEST</p>\n      <h2 id=\"audienceGateTitle\"><?= siteTitle ?></h2>\n      <p class=\"audience-desc\">신청하시는 분을 골라 주십시오.<br>교직원과 학부모의 신청은 따로 관리됩니다.</p>\n      <div class=\"audience-choices\">\n        <a class=\"audience-btn\" href=\"#\" data-audience=\"staff\"><span class=\"audience-icon\">🧑‍🏫</span>교직원<span class=\"audience-sub\">교직원 희망도서 신청</span></a>\n        <a class=\"audience-btn parent\" href=\"#\" data-audience=\"parent\"><span class=\"audience-icon\">👨‍👩‍👧</span>학부모<span class=\"audience-sub\">학부모님 희망도서 신청</span></a>\n      </div>\n    </div>\n  </div>\n  <div id=\"staffCodeGate\" class=\"audience-gate\" hidden>\n    <form class=\"audience-card code-card\" id=\"staffCodeForm\" autocomplete=\"off\">\n      <p class=\"eyebrow\">STAFF ONLY</p>\n      <h2>🔑 교직원 인증</h2>\n      <p class=\"audience-desc\">학교에서 안내한 인증번호를 입력해 주십시오.</p>\n      <input id=\"staffCodeInput\" class=\"code-input\" type=\"password\" inputmode=\"text\" autocomplete=\"off\" maxlength=\"12\" placeholder=\"인증번호\" aria-label=\"교직원 인증번호\">\n      <p id=\"staffCodeMessage\" class=\"code-message\" role=\"alert\" hidden></p>\n      <button type=\"submit\" id=\"staffCodeSubmit\" class=\"code-submit\">확인</button>\n      <a id=\"staffCodeBack\" class=\"code-back\" href=\"#\">← 처음 화면</a>\n    </form>\n  </div>\n  <header class=\"header\">\n    <p class=\"eyebrow\">BOOK REQUEST</p>\n    <h1>📚 <?= siteTitle ?></h1>\n\n    <p class=\"description\">\n      책 제목 또는 ISBN을 입력하십시오.<br>\n      책 정보와 중복 여부는 자동으로 확인됩니다.\n    </p>\n\n    <div id=\"audienceBar\" class=\"audience-bar\" hidden>\n      <span id=\"audienceBadge\" class=\"audience-badge\"></span>\n      <a id=\"audienceSwitch\" class=\"audience-switch\" href=\"#\" target=\"_top\" hidden>← 처음 화면</a>\n    </div>\n  </header>\n\n\n  <section id=\"hrNoticeTicker\" class=\"hr-notice-ticker\" aria-label=\"도서관 공지 전광판\" aria-live=\"polite\" hidden>\n    <div class=\"hr-notice-inner\">\n      <span class=\"hr-notice-label\">📢 알림</span>\n      <div class=\"hr-notice-window\">\n        <div id=\"hrNoticeTrack\" class=\"hr-notice-track\"></div>\n      </div>\n    </div>\n  </section>\n\n  <nav class=\"view-tabs\" aria-label=\"희망도서 페이지 메뉴\">\n    <button class=\"view-tab active\" type=\"button\" data-panel=\"applyPanel\" aria-selected=\"true\">\n      🔎 도서 신청\n    </button>\n    <button class=\"view-tab\" type=\"button\" data-panel=\"requestsPanel\" aria-selected=\"false\">\n      📋 신청 목록\n    </button>\n    <button class=\"view-tab\" type=\"button\" data-panel=\"uploadPanel\" aria-selected=\"false\">\n      📥 목록 업로드\n    </button>\n    <div id=\"guideHover\" class=\"guide-hover\">\n      <button\n        id=\"guideHoverTrigger\"\n        class=\"guide-hover-trigger\"\n        type=\"button\"\n        aria-expanded=\"false\"\n        aria-describedby=\"guidePopover\"\n      >\n        ❓ 신청 방법\n      </button>\n\n      <div id=\"guidePopover\" class=\"guide-popover\" role=\"dialog\" aria-modal=\"true\" aria-labelledby=\"guideTitle\">\n        <h2 id=\"guideTitle\">이렇게 신청하십시오</h2>\n        <ol id=\"guideSteps\">\n          <li><strong>1. 책 검색</strong><span>책 제목 또는 ISBN을 입력하고 「책 검색하기」를 누릅니다.</span></li>\n          <li><strong>2. 신청바구니에 담기</strong><span>표지·저자·출판사·연도·정가를 확인하고 「신청바구니에 담기」를 누릅니다.</span></li>\n          <li><strong>3. 다른 책도 계속 담기</strong><span>검색어를 바꿔도 담은 책은 유지되며, 한 번에 최대 20권까지 담을 수 있습니다.</span></li>\n          <li><strong>4. 한꺼번에 신청</strong><span>화면 아래 신청바구니에서 「선택한 도서 신청하기」를 누릅니다.</span></li>\n          <li><strong>5. 결과 확인·취소</strong><span>「신청 목록」에서 접수 결과와 서지정보를 확인하고, 필요하면 신청을 취소하거나 되돌릴 수 있습니다.</span></li>\n        </ol>\n        <p id=\"guideNote\">희망도서 신청 시 우선 대출 혜택을 받을 수 있습니다.</p>\n        <button type=\"button\" class=\"guide-close\" id=\"guideClose\">확인했습니다</button>\n      </div>\n    </div>\n  </nav>\n\n  <div id=\"applyPanel\" class=\"page-panel active\">\n\n  <main class=\"card\">\n    <div class=\"form-intro\">\n      <div class=\"form-intro-icon\">📖</div>\n      <div>\n        <h2>어떤 책을 신청하시겠습니까?</h2>\n        <p>책 제목이나 ISBN으로 찾고, 표지를 확인한 뒤 신청하십시오.</p>\n      </div>\n    </div>\n\n    <form id=\"requestForm\">\n      <div class=\"field\" id=\"staffApplicantField\">\n        <label for=\"applicant\">신청자</label>\n\n        <input\n          id=\"applicant\"\n          name=\"requestApplicant\"\n          type=\"search\"\n          autocomplete=\"off\"\n          autocapitalize=\"off\"\n          spellcheck=\"false\"\n          data-lpignore=\"true\"\n          data-1p-ignore=\"true\"\n          placeholder=\"예: 홍길동\"\n        >\n\n        <p class=\"help\">\n          희망도서 신청 시 우선 대출 혜택을 받을 수 있습니다.\n        </p>\n      </div>\n\n      <div class=\"field\" id=\"parentApplicantField\" hidden>\n        <label>우리 아이 정보 <span class=\"parent-hint\">(학년·반·번호를 눌러서 골라 주십시오)</span></label>\n        <div class=\"parent-fields\">\n          <div>\n            <label for=\"pGrade\">학년</label>\n            <select id=\"pGrade\"><option value=\"\">선택</option><option value=\"1\">1</option><option value=\"2\">2</option><option value=\"3\">3</option><option value=\"4\">4</option><option value=\"5\">5</option><option value=\"6\">6</option></select>\n          </div>\n          <div>\n            <label for=\"pClass\">반</label>\n            <select id=\"pClass\"><option value=\"\">선택</option></select>\n          </div>\n          <div>\n            <label for=\"pNumber\">번호</label>\n            <select id=\"pNumber\"><option value=\"\">선택</option></select>\n          </div>\n          <div class=\"pname\">\n            <label for=\"pName\">학생 이름 (선택)</label>\n            <input id=\"pName\" type=\"text\" autocomplete=\"off\" placeholder=\"예: 홍길동\">\n          </div>\n        </div>\n        <div id=\"parentQuota\" class=\"parent-quota\" hidden aria-live=\"polite\"></div>\n        <p class=\"help\">\n          🔒 입력하신 학년·반·번호(와 학생 이름)는 희망도서 신청 확인과 도서 구입 처리에만 쓰이며, 도서 구입이 끝나면 함께 파기됩니다. 신청 목록 화면에는 보이지 않습니다.\n        </p>\n      </div>\n\n      <div class=\"field\">\n        <label for=\"query\">책 제목 또는 ISBN</label>\n\n        <input\n          id=\"query\"\n          type=\"search\"\n          autocomplete=\"off\"\n          placeholder=\"예: 아몬드 또는 978...\"\n          required\n        >\n\n        <p class=\"help\">\n          같은 제목의 책이 많다면 ISBN 13자리를 권장합니다.\n        </p>\n      </div>\n\n      <button\n        id=\"submitButton\"\n        class=\"submit-button\"\n        type=\"submit\"\n      >\n        책 검색하기\n      </button>\n    </form>\n\n\n    <div id=\"message\" class=\"message\"></div>\n\n    <section id=\"searchResults\" class=\"search-results\">\n      <h2 class=\"search-results-title\">\n        신청할 책을 여러 권 선택해 주십시오.<br>\n        선택을 마치면 아래 신청 버튼을 누르십시오.\n      </h2>\n\n      <div id=\"bookList\" class=\"book-list\"></div>\n    </section>\n  </main>\n\n  <section class=\"bs-section\" aria-labelledby=\"bsTitle\">\n    <div class=\"bs-heading\">\n      <div>\n        <h2 id=\"bsTitle\" class=\"bs-title\">🔥 지금 많이 읽는 책</h2>\n        <p class=\"bs-help\">관심 있는 책을 누르면 정확한 도서를 다시 확인한 뒤 신청할 수 있습니다.</p>\n      </div>\n      <span class=\"bs-badge\">하루 1회 갱신</span>\n    </div>\n\n    <div class=\"bs-tabs\" role=\"tablist\" aria-label=\"베스트셀러 분류\">\n      <button class=\"bs-tab active\" type=\"button\" data-bs-category=\"all\" aria-selected=\"true\">종합</button>\n      <button class=\"bs-tab\" type=\"button\" data-bs-category=\"youth\" aria-selected=\"false\">청소년</button>\n      <button class=\"bs-tab\" type=\"button\" data-bs-category=\"children\" aria-selected=\"false\">어린이</button>\n      <button class=\"bs-tab\" type=\"button\" data-bs-category=\"picture\" aria-selected=\"false\">그림책</button>\n    </div>\n\n    <div id=\"bsStatus\" class=\"bs-status show\">베스트셀러를 불러오고 있습니다…</div>\n    <div class=\"bs-carousel\">\n      <button id=\"bsPrevious\" class=\"bs-arrow bs-arrow-previous\" type=\"button\" aria-label=\"이전 베스트셀러\">‹</button>\n      <div id=\"bsList\" class=\"bs-list\"></div>\n      <button id=\"bsNext\" class=\"bs-arrow bs-arrow-next\" type=\"button\" aria-label=\"다음 베스트셀러\">›</button>\n    </div>\n\n\n    <div class=\"bs-more-wrap\">\n      <a\n        class=\"bs-more\"\n        href=\"https://www.yes24.com/Product/Category/BestSeller\"\n        target=\"_blank\"\n        rel=\"noopener noreferrer\"\n      >\n        YES24에서 베스트셀러 더 보기 →\n      </a>\n    </div>\n\n    <p class=\"bs-source\">\n      도서 정보 제공:\n      <a href=\"https://www.yes24.com/Product/Category/BestSeller\" target=\"_blank\" rel=\"noopener noreferrer\">YES24 인터넷서점</a>\n    </p>\n  </section>\n\n  <section id=\"cartBar\" class=\"cart-bar\" aria-label=\"신청바구니\">\n    <button\n      id=\"cartToggle\"\n      class=\"cart-toggle\"\n      type=\"button\"\n      aria-expanded=\"false\"\n      aria-controls=\"cartDrawer\"\n    >\n      <span class=\"cart-count\">\n        🛒 신청바구니 <strong id=\"cartCount\">0</strong>/20권\n      </span>\n      <span class=\"cart-open-hint\">담은 책 보기</span>\n    </button>\n\n    <button\n      id=\"cartSubmit\"\n      class=\"cart-submit\"\n      type=\"button\"\n      disabled\n    >\n      선택한 도서 신청하기\n    </button>\n  </section>\n\n  <div\n    id=\"cartDrawer\"\n    class=\"cart-drawer\"\n    role=\"dialog\"\n    aria-modal=\"true\"\n    aria-labelledby=\"cartDrawerTitle\"\n  >\n    <button\n      class=\"cart-backdrop\"\n      type=\"button\"\n      data-cart-close\n      aria-label=\"신청바구니 닫기\"\n    ></button>\n\n    <section class=\"cart-panel\">\n      <div class=\"cart-panel-header\">\n        <div>\n          <p>BOOK CART</p>\n          <h2 id=\"cartDrawerTitle\">🛒 신청바구니</h2>\n        </div>\n        <button\n          id=\"cartClose\"\n          class=\"cart-close\"\n          type=\"button\"\n          aria-label=\"신청바구니 닫기\"\n        >×</button>\n      </div>\n\n      <p class=\"cart-guide\">\n        다른 책을 다시 검색해도 그대로 유지됩니다. 한 번에 최대 20권까지 담을 수 있습니다.\n      </p>\n\n      <div id=\"cartList\" class=\"cart-list\"></div>\n\n      <div class=\"cart-panel-actions\">\n        <button id=\"cartClear\" class=\"cart-clear\" type=\"button\">\n          전체 비우기\n        </button>\n        <button\n          id=\"cartDrawerSubmit\"\n          class=\"cart-submit cart-drawer-submit\"\n          type=\"button\"\n          disabled\n        >\n          담은 도서 신청하기\n        </button>\n      </div>\n    </section>\n  </div>\n\n  </div>\n\n  <section\n    id=\"requestsPanel\"\n    class=\"page-panel requests-section\"\n    aria-labelledby=\"requestsTitle\"\n  >\n    <div class=\"panel-heading\">\n      <div>\n        <p class=\"panel-eyebrow\">REQUEST LIST</p>\n        <h2 id=\"requestsTitle\">📋 신청된 희망도서</h2>\n        <p>내가 신청한 도서만 표시되며, 취소도 신청한 본인만 할 수 있습니다.</p>\n      </div>\n      <button id=\"requestRefresh\" class=\"refresh-button\" type=\"button\">\n        새로고침\n      </button>\n    </div>\n\n    <div\n      class=\"request-filter-tabs\"\n      role=\"tablist\"\n      aria-label=\"신청 목록 구분\"\n    >\n      <button\n        class=\"request-filter-tab active\"\n        type=\"button\"\n        data-request-filter=\"active\"\n        role=\"tab\"\n        aria-selected=\"true\"\n      >\n        <span>📚 신청 도서</span>\n        <strong id=\"requestActiveCount\">0</strong>\n      </button>\n      <button\n        class=\"request-filter-tab\"\n        type=\"button\"\n        data-request-filter=\"cancelled\"\n        role=\"tab\"\n        aria-selected=\"false\"\n      >\n        <span>↩️ 신청 취소</span>\n        <strong id=\"requestCancelledCount\">0</strong>\n      </button>\n    </div>\n\n    <div id=\"requestSummary\" class=\"request-summary\"></div>\n    <div id=\"requestListStatus\" class=\"request-list-status\">\n      신청 목록 탭을 누르면 최신 목록을 불러옵니다.\n    </div>\n    <div id=\"requestList\" class=\"request-list\"></div>\n  </section>\n\n\n  <section\n    id=\"uploadPanel\"\n    class=\"page-panel upload-section\"\n    aria-labelledby=\"uploadTitle\"\n  >\n    <div class=\"upload-heading\">\n      <div>\n        <p class=\"panel-eyebrow\">BOOK LIST IMPORT</p>\n        <h2 id=\"uploadTitle\">📥 도서 목록 파일 업로드</h2>\n        <p>\n          도서 목록 엑셀·CSV 파일을 선택하면<br>\n          도서명과 ISBN 등의 열을 자동으로 찾아 업무용 목록으로 정리합니다.\n        </p>\n      </div>\n      <span class=\"upload-badge\">사서 업무용</span>\n    </div>\n\n    <div class=\"upload-card\">\n      <label id=\"uploadDropzone\" class=\"upload-dropzone\" for=\"uploadFile\">\n        <input\n          id=\"uploadFile\"\n          type=\"file\"\n          accept=\".xls,.xlsx,.xlsm,.csv,.tsv,.html,.htm\"\n        >\n        <span>\n          <span class=\"upload-icon\">📄</span>\n          <strong>엑셀·CSV 파일을 선택하거나 여기에 놓으십시오</strong>\n          <span>.xls · .xlsx · .xlsm · .csv · .tsv / 한 번에 최대 1,500권</span>\n        </span>\n      </label>\n\n      <p class=\"upload-format-note\">\n        <span>도서명 또는 ISBN 열이 있어야 합니다.</span>\n        <span>저자·출판사·출판연도·정가·수량·시리즈명은 파일에 있는 항목만 가져옵니다.</span>\n        <span>표지와 파일 작성자 정보는 가져오지 않습니다.</span>\n      </p>\n\n      <div id=\"uploadStatus\" class=\"upload-status\" role=\"status\"></div>\n\n      <div id=\"uploadPreview\" class=\"upload-preview\">\n        <div class=\"upload-preview-head\">\n          <strong>가져올 내용 미리보기</strong>\n          <span id=\"uploadPreviewCount\"></span>\n        </div>\n        <div class=\"upload-table-wrap\">\n          <table class=\"upload-table\">\n            <thead>\n              <tr>\n                <th>번호</th>\n                <th>도서명</th>\n                <th>저자</th>\n                <th>출판사</th>\n                <th>연도</th>\n                <th>ISBN13</th>\n                <th>정가</th>\n              </tr>\n            </thead>\n            <tbody id=\"uploadPreviewBody\"></tbody>\n          </table>\n        </div>\n      </div>\n\n      <div class=\"upload-actions\">\n        <button id=\"uploadReset\" class=\"upload-reset\" type=\"button\" disabled>\n          다시 선택\n        </button>\n        <button id=\"uploadSave\" class=\"upload-save\" type=\"button\" disabled>\n          이 목록을 시트에 넣기\n        </button>\n      </div>\n    </div>\n  </section>\n\n\n  <p class=\"footer\">\n    처리에는 몇 초 정도 걸릴 수 있습니다.<br>\n    창을 닫지 말고 완료 안내를 확인해 주십시오.\n  </p>\n  <p class=\"build-id\">화면 버전 <?= buildId ?></p>\n\n  <footer class=\"school-contact-footer\">\n    <? if (managerName || managerExtension) { ?>\n      <div class=\"school-contact-row\">\n        <span class=\"school-contact-badge\">📚 도서관 문의</span>\n\n        <? if (managerName) { ?>\n          <span class=\"school-contact-detail\">\n            <strong>담당자</strong>\n            <span><?= managerName ?></span>\n          </span>\n        <? } ?>\n\n        <? if (managerExtension) { ?>\n          <span class=\"school-contact-detail\">\n            <strong>내선</strong>\n            <span><?= managerExtension ?></span>\n          </span>\n        <? } ?>\n      </div>\n    <? } ?>\n\n    <div class=\"site-credit\">\n      학교도서관 도서구입 지원 도구    </div>\n  </footer>\n\n  <div\n    id=\"noticeModal\"\n    class=\"notice-modal\"\n    role=\"dialog\"\n    aria-modal=\"true\"\n    aria-labelledby=\"noticeTitle\"\n  >\n    <div class=\"notice-dialog\">\n      <div class=\"notice-visual\">\n        <div id=\"noticeIcon\" class=\"notice-icon\">📚</div>\n      </div>\n\n      <div class=\"notice-content\">\n        <h2 id=\"noticeTitle\" class=\"notice-title\"></h2>\n        <p id=\"noticeBook\" class=\"notice-book\"></p>\n        <p id=\"noticeDescription\" class=\"notice-description\"></p>\n\n        <div\n          id=\"noticeActions\"\n          class=\"notice-actions single\"\n        >\n          <button\n            id=\"noticeUndo\"\n            class=\"notice-undo\"\n            type=\"button\"\n            hidden\n          >\n            신청 취소 되돌리기\n          </button>\n          <button\n            id=\"noticeClose\"\n            class=\"notice-close\"\n            type=\"button\"\n          >\n            확인\n          </button>\n        </div>\n      </div>\n    </div>\n  </div>\n\n\n  <div\n    id=\"cancelModal\"\n    class=\"cancel-modal\"\n    role=\"dialog\"\n    aria-modal=\"true\"\n    aria-labelledby=\"cancelTitle\"\n  >\n    <div class=\"cancel-dialog\">\n      <div class=\"cancel-icon\">↩️</div>\n      <h2 id=\"cancelTitle\">희망도서 신청을 취소하시겠습니까?</h2>\n      <p id=\"cancelBook\" class=\"cancel-book\"></p>\n      <p class=\"cancel-description\">\n        취소하면 신청 목록에서 이 도서 한 권만 삭제됩니다. 바로 되돌릴 수 있도록 안전하게 보관합니다.\n      </p>\n      <div class=\"cancel-actions\">\n        <button id=\"cancelKeep\" class=\"cancel-keep\" type=\"button\">\n          그대로 두기\n        </button>\n        <button id=\"cancelConfirm\" class=\"cancel-confirm\" type=\"button\">\n          신청 취소하기\n        </button>\n      </div>\n    </div>\n  </div>\n\n  <script>\n    const form =\n      document.getElementById('requestForm');\n    const applicant =\n      document.getElementById('applicant');\n    const query =\n      document.getElementById('query');\n    const submitButton =\n      document.getElementById('submitButton');\n    const message =\n      document.getElementById('message');\n    const searchResults =\n      document.getElementById('searchResults');\n    const bookList =\n      document.getElementById('bookList');\n    const noticeModal =\n      document.getElementById('noticeModal');\n    const noticeIcon =\n      document.getElementById('noticeIcon');\n    const noticeTitle =\n      document.getElementById('noticeTitle');\n    const noticeBook =\n      document.getElementById('noticeBook');\n    const noticeDescription =\n      document.getElementById('noticeDescription');\n    const noticeActions =\n      document.getElementById('noticeActions');\n    const noticeUndo =\n      document.getElementById('noticeUndo');\n    const noticeClose =\n      document.getElementById('noticeClose');\n    const cartBar =\n      document.getElementById('cartBar');\n    const cartCount =\n      document.getElementById('cartCount');\n    const cartSubmit =\n      document.getElementById('cartSubmit');\n    const viewTabs =\n      Array.from(document.querySelectorAll('.view-tab'));\n    const pagePanels =\n      Array.from(document.querySelectorAll('.page-panel'));\n    const applyPanel =\n      document.getElementById('applyPanel');\n    const requestRefresh =\n      document.getElementById('requestRefresh');\n    const requestSummary =\n      document.getElementById('requestSummary');\n    const requestListStatus =\n      document.getElementById('requestListStatus');\n    const requestList =\n      document.getElementById('requestList');\n    const requestFilterTabs =\n      Array.from(\n        document.querySelectorAll('.request-filter-tab')\n      );\n    const requestActiveCount =\n      document.getElementById('requestActiveCount');\n    const requestCancelledCount =\n      document.getElementById('requestCancelledCount');\n    const uploadFile =\n      document.getElementById('uploadFile');\n    const uploadDropzone =\n      document.getElementById('uploadDropzone');\n    const uploadStatus =\n      document.getElementById('uploadStatus');\n    const uploadPreview =\n      document.getElementById('uploadPreview');\n    const uploadPreviewCount =\n      document.getElementById('uploadPreviewCount');\n    const uploadPreviewBody =\n      document.getElementById('uploadPreviewBody');\n    const uploadReset =\n      document.getElementById('uploadReset');\n    const uploadSave =\n      document.getElementById('uploadSave');\n    let uploadParsedBooks = [];\n    let uploadFileName = '';\n    const bsListElement =\n      document.getElementById('bsList');\n    const bsPrevious =\n      document.getElementById('bsPrevious');\n    const bsNext =\n      document.getElementById('bsNext');\n    const guideHover =\n      document.getElementById('guideHover');\n    const guideHoverTrigger =\n      document.getElementById('guideHoverTrigger');\n    const cancelModal =\n      document.getElementById('cancelModal');\n    const cancelBook =\n      document.getElementById('cancelBook');\n    const cancelKeep =\n      document.getElementById('cancelKeep');\n    const cancelConfirm =\n      document.getElementById('cancelConfirm');\n    const cartToggle =\n      document.getElementById('cartToggle');\n    const cartDrawer =\n      document.getElementById('cartDrawer');\n    const cartClose =\n      document.getElementById('cartClose');\n    const cartList =\n      document.getElementById('cartList');\n    const cartClear =\n      document.getElementById('cartClear');\n    const cartDrawerSubmit =\n      document.getElementById('cartDrawerSubmit');\n    const cartBackdrop =\n      document.querySelector('[data-cart-close]');\n    let books = [];\n    let cartBooks = new Map();\n    let submissionInProgress = false;\n    let submittedRequestsLoaded = false;\n    let submittedRequestItems = [];\n    let requestListFilter = 'active';\n    let bsAutoTimer = null;\n    let cancelCandidate = null;\n    let undoCandidate = null;\n\n    noticeClose.addEventListener('click', closeNoticeModal);\n    noticeUndo.addEventListener('click', undoCancelledRequest);\n    cartSubmit.addEventListener('click', submitCart);\n    cartDrawerSubmit.addEventListener('click', submitCart);\n    cartToggle.addEventListener('click', openCartDrawer);\n    cartClose.addEventListener('click', closeCartDrawer);\n    cartBackdrop.addEventListener('click', closeCartDrawer);\n    cartClear.addEventListener('click', clearCart);\n\n    document.addEventListener('keydown', function(event) {\n      if (event.key === 'Escape') {\n        closeCartDrawer();\n      }\n    });\n\n    viewTabs.forEach(function(tab) {\n      tab.addEventListener('click', function() {\n        showPagePanel(tab.dataset.panel || 'applyPanel');\n      });\n    });\n\n    requestRefresh.addEventListener('click', function() {\n      loadSubmittedRequests(true);\n    });\n\n    requestFilterTabs.forEach(function(tab) {\n      tab.addEventListener('click', function() {\n        requestListFilter =\n          tab.dataset.requestFilter || 'active';\n\n        renderSubmittedRequests({\n          items: submittedRequestItems\n        });\n      });\n    });\n\n    guideHoverTrigger.addEventListener('click', function(event) {\n      event.stopPropagation();\n      const opened = guideHover.classList.toggle('open');\n      guideHoverTrigger.setAttribute(\n        'aria-expanded',\n        opened ? 'true' : 'false'\n      );\n    });\n\n    document.addEventListener('click', function(event) {\n      if (!guideHover.contains(event.target)) {\n        closeGuidePopover();\n      }\n    });\n\n    cancelKeep.addEventListener('click', closeCancelModal);\n    cancelConfirm.addEventListener(\n      'click',\n      confirmCancelRequest\n    );\n\n    cancelModal.addEventListener('click', function(event) {\n      if (event.target === cancelModal) {\n        closeCancelModal();\n      }\n    });\n\n    bsPrevious.addEventListener('click', function() {\n      scrollBestsellerCards(-1);\n      startBestsellerAuto();\n    });\n\n    bsNext.addEventListener('click', function() {\n      scrollBestsellerCards(1);\n      startBestsellerAuto();\n    });\n\n    bsListElement.addEventListener(\n      'mouseenter',\n      stopBestsellerAuto\n    );\n    bsListElement.addEventListener(\n      'mouseleave',\n      startBestsellerAuto\n    );\n    bsListElement.addEventListener(\n      'focusin',\n      stopBestsellerAuto\n    );\n    bsListElement.addEventListener(\n      'focusout',\n      startBestsellerAuto\n    );\n    bsListElement.addEventListener(\n      'touchstart',\n      stopBestsellerAuto,\n      { passive: true }\n    );\n    bsListElement.addEventListener(\n      'touchend',\n      function() {\n        window.setTimeout(startBestsellerAuto, 1200);\n      },\n      { passive: true }\n    );\n\n    document.addEventListener('visibilitychange', function() {\n      if (document.hidden) {\n        stopBestsellerAuto();\n      } else {\n        startBestsellerAuto();\n      }\n    });\n\n    noticeModal.addEventListener('click', function(event) {\n      if (event.target === noticeModal) {\n        closeNoticeModal();\n      }\n    });\n\n    document.addEventListener('keydown', function(event) {\n      if (event.key !== 'Escape') {\n        return;\n      }\n\n      if (noticeModal.classList.contains('show')) {\n        closeNoticeModal();\n      }\n\n      if (cancelModal.classList.contains('show')) {\n        closeCancelModal();\n      }\n\n      closeGuidePopover();\n    });\n\n\n\n    function setUploadStatus(message, type) {\n      uploadStatus.textContent = String(message || '');\n      uploadStatus.className = 'upload-status show ' + (type || 'info');\n    }\n\n    function resetUploadPanel() {\n      uploadParsedBooks = [];\n      uploadFileName = '';\n      uploadFile.value = '';\n      uploadPreviewBody.innerHTML = '';\n      uploadPreviewCount.textContent = '';\n      uploadPreview.classList.remove('show');\n      uploadStatus.className = 'upload-status';\n      uploadStatus.textContent = '';\n      uploadSave.disabled = true;\n      uploadReset.disabled = true;\n    }\n\n    function normalizeUploadHeader(value) {\n      return String(value == null ? '' : value)\n        .replace(/<[^>]*>/g, '')\n        .replace(/[\\s\\n\\r()（）\\[\\]{}·._\\-\\/]/g, '')\n        .toUpperCase();\n    }\n\n    function findUploadHeaderKey(value) {\n      const normalized = normalizeUploadHeader(value);\n      const aliases = {\n        title: ['상품명','도서명','책제목','서명','제목'],\n        isbn: ['ISBN13','ISBN','국제표준도서번호'],\n        author: ['저자아티스트','저자','작가','글쓴이','지은이'],\n        publisher: ['출판사제작사','출판사','발행자','출판자'],\n        year: ['출간일','출판연도','발행연도','출간연도','연도'],\n        price: ['정가','도서정가','정가원','가격'],\n        quantity: ['수량','주문수량','권수'],\n        series: ['시리즈명','시리즈'],\n        note: ['비고','메모','구분','분류']\n      };\n\n      for (const key in aliases) {\n        if (aliases[key].indexOf(normalized) !== -1) {\n          return key;\n        }\n      }\n      return '';\n    }\n\n    function cleanUploadIsbn(value) {\n      let text = String(value == null ? '' : value).trim();\n\n      if (/^[0-9.]+E\\+?[0-9]+$/i.test(text)) {\n        const numeric = Number(text);\n        if (Number.isFinite(numeric)) {\n          text = numeric.toFixed(0);\n        }\n      }\n\n      return text.replace(/[^0-9Xx]/g, '').toUpperCase();\n    }\n\n    function cleanUploadYear(value) {\n      const match = String(value == null ? '' : value).match(/(?:19|20)\\d{2}/);\n      return match ? match[0] : '';\n    }\n\n    function cleanUploadNumber(value, fallback) {\n      const numeric = Number(\n        String(value == null ? '' : value)\n          .replace(/[^0-9.-]/g, '')\n      );\n      return Number.isFinite(numeric) && numeric > 0\n        ? numeric\n        : (fallback || 0);\n    }\n\n    function decodeUploadText(buffer) {\n      const sample = new TextDecoder('latin1').decode(buffer.slice(0, 4096));\n      const charsetMatch = sample.match(/charset\\s*=\\s*[\"']?([^\"'\\s;>]+)/i);\n      const charset = charsetMatch ? charsetMatch[1].toLowerCase() : 'utf-8';\n\n      try {\n        return new TextDecoder(charset).decode(buffer);\n      } catch (error) {\n        return new TextDecoder('utf-8').decode(buffer);\n      }\n    }\n\n    function htmlMatricesFromText(text) {\n      if (!/<table[\\s>]/i.test(text)) {\n        return [];\n      }\n\n      const documentNode = new DOMParser().parseFromString(text, 'text/html');\n      return Array.from(documentNode.querySelectorAll('table')).map(function(table) {\n        return Array.from(table.querySelectorAll('tr')).map(function(row) {\n          return Array.from(row.querySelectorAll('th,td')).map(function(cell) {\n            return cell.textContent.replace(/\\u00a0/g, ' ').trim();\n          });\n        });\n      });\n    }\n\n    function delimitedMatrixFromText(text, delimiter) {\n      const rows = [];\n      let row = [];\n      let cell = '';\n      let quoted = false;\n\n      for (let index = 0; index < text.length; index++) {\n        const character = text[index];\n\n        if (quoted) {\n          if (character === '\"' && text[index + 1] === '\"') {\n            cell += '\"';\n            index++;\n          } else if (character === '\"') {\n            quoted = false;\n          } else {\n            cell += character;\n          }\n          continue;\n        }\n\n        if (character === '\"') {\n          quoted = true;\n        } else if (character === delimiter) {\n          row.push(cell);\n          cell = '';\n        } else if (character === '\\n') {\n          row.push(cell.replace(/\\r$/, ''));\n          rows.push(row);\n          row = [];\n          cell = '';\n        } else {\n          cell += character;\n        }\n      }\n\n      if (cell || row.length) {\n        row.push(cell.replace(/\\r$/, ''));\n        rows.push(row);\n      }\n      return rows;\n    }\n\n    function workbookMatricesFromBuffer(file, buffer) {\n      const extension = (file.name.split('.').pop() || '').toLowerCase();\n      const text = decodeUploadText(buffer);\n      const htmlMatrices = htmlMatricesFromText(text);\n\n      if (htmlMatrices.length) {\n        return htmlMatrices;\n      }\n\n      if (extension === 'csv' || extension === 'tsv') {\n        return [\n          delimitedMatrixFromText(\n            text.replace(/^\\uFEFF/, ''),\n            extension === 'tsv' ? '\\t' : ','\n          )\n        ];\n      }\n\n      if (typeof XLSX === 'undefined') {\n        throw new Error(\n          '엑셀 읽기 도구를 불러오지 못했습니다. 인터넷 연결을 확인한 뒤 다시 시도해 주십시오.'\n        );\n      }\n\n      const workbook = XLSX.read(buffer, {\n        type: 'array',\n        raw: false,\n        cellDates: false\n      });\n\n      return workbook.SheetNames.map(function(sheetName) {\n        return XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], {\n          header: 1,\n          defval: '',\n          raw: false,\n          blankrows: false\n        });\n      });\n    }\n\n    function locateUploadTable(matrices) {\n      let best = null;\n\n      matrices.forEach(function(matrix) {\n        const rowLimit = Math.min(matrix.length, 50);\n\n        for (let rowIndex = 0; rowIndex < rowLimit; rowIndex++) {\n          const row = Array.isArray(matrix[rowIndex]) ? matrix[rowIndex] : [];\n          const map = {};\n          let score = 0;\n\n          row.forEach(function(value, columnIndex) {\n            const key = findUploadHeaderKey(value);\n            if (key && (map[key] === undefined || (key === 'isbn' && normalizeUploadHeader(value) === 'ISBN13'))) {\n              map[key] = columnIndex;\n              score++;\n            }\n          });\n\n          const essential = map.title !== undefined || map.isbn !== undefined;\n          if (essential && score >= 2 && (!best || score > best.score)) {\n            best = {\n              matrix: matrix,\n              headerRow: rowIndex,\n              map: map,\n              score: score\n            };\n          }\n        }\n      });\n\n      if (!best) {\n        throw new Error(\n          '도서명 또는 ISBN 열을 찾지 못했습니다. 첫 행의 열 이름을 확인해 주십시오.'\n        );\n      }\n      return best;\n    }\n\n    function uploadBooksFromTable(table) {\n      const books = [];\n      const seen = {};\n      let skipped = 0;\n\n      for (\n        let rowIndex = table.headerRow + 1;\n        rowIndex < table.matrix.length;\n        rowIndex++\n      ) {\n        const row = Array.isArray(table.matrix[rowIndex])\n          ? table.matrix[rowIndex]\n          : [];\n\n        const valueAt = function(key) {\n          const columnIndex = table.map[key];\n          return columnIndex === undefined ? '' : row[columnIndex];\n        };\n\n        const title = String(valueAt('title') || '')\n          .replace(/<[^>]*>/g, '')\n          .replace(/\\s+/g, ' ')\n          .trim();\n        const isbn = cleanUploadIsbn(valueAt('isbn'));\n\n        if (!title && !isbn) {\n          continue;\n        }\n\n        const publisher = String(valueAt('publisher') || '').trim();\n        const year = cleanUploadYear(valueAt('year'));\n        const key = isbn\n          ? 'isbn:' + isbn\n          : 'title:' +\n            title.replace(/\\s+/g, '').toUpperCase() + '|' +\n            publisher.replace(/\\s+/g, '').toUpperCase() + '|' +\n            year;\n\n        if (seen[key]) {\n          skipped++;\n          continue;\n        }\n        seen[key] = true;\n\n        books.push({\n          title: title,\n          author: String(valueAt('author') || '').trim(),\n          publisher: publisher,\n          year: year,\n          isbn: isbn,\n          price: cleanUploadNumber(valueAt('price'), 0),\n          quantity: cleanUploadNumber(valueAt('quantity'), 1),\n          series: String(valueAt('series') || '').trim(),\n          note: String(valueAt('note') || '').trim()\n        });\n      }\n\n      return {\n        books: books.slice(0, 1500),\n        skipped: skipped + Math.max(0, books.length - 1500),\n        limited: books.length > 1500\n      };\n    }\n\n    function renderUploadPreview(result) {\n      uploadPreviewBody.innerHTML = '';\n\n      const fragment = document.createDocumentFragment();\n\n      result.books.forEach(function(book, index) {\n        const row = document.createElement('tr');\n        const cells = [\n          String(index + 1),\n          book.title || '(제목 없음)',\n          book.author || '',\n          book.publisher || '',\n          book.year || '',\n          book.isbn || '',\n          book.price ? book.price.toLocaleString('ko-KR') + '원' : ''\n        ];\n\n        cells.forEach(function(value) {\n          const cell = document.createElement('td');\n          cell.textContent = value;\n          row.appendChild(cell);\n        });\n\n        fragment.appendChild(row);\n      });\n\n      uploadPreviewBody.appendChild(fragment);\n      uploadPreviewCount.textContent =\n        '총 ' + result.books.length + '권 · 전체 ' + result.books.length + '권 표시';\n      uploadPreview.classList.add('show');\n    }\n\n    async function handleUploadFile(file) {\n      resetUploadPanel();\n\n      if (!file) {\n        return;\n      }\n\n      if (file.size > 15 * 1024 * 1024) {\n        setUploadStatus('파일 크기는 15MB 이하만 가능합니다.', 'error');\n        return;\n      }\n\n      uploadFileName = file.name;\n      uploadReset.disabled = false;\n      setUploadStatus('파일을 읽고 열 이름을 확인하고 있습니다…', 'info');\n\n      try {\n        const buffer = await file.arrayBuffer();\n        const matrices = workbookMatricesFromBuffer(file, buffer);\n        const table = locateUploadTable(matrices);\n        const result = uploadBooksFromTable(table);\n\n        if (!result.books.length) {\n          throw new Error('도서명 또는 ISBN이 있는 도서 행을 찾지 못했습니다.');\n        }\n\n        uploadParsedBooks = result.books;\n        renderUploadPreview(result);\n        uploadSave.disabled = false;\n\n        let message =\n          file.name + '에서 ' + result.books.length + '권을 찾았습니다.';\n        if (result.skipped) {\n          message += ' 중복·빈 행 ' + result.skipped + '개는 제외했습니다.';\n        }\n        if (result.limited) {\n          message += ' 한 번에 최대 1,500권까지만 표시합니다.';\n        }\n        setUploadStatus(message, 'success');\n      } catch (error) {\n        uploadParsedBooks = [];\n        uploadSave.disabled = true;\n        setUploadStatus(\n          error && error.message\n            ? error.message\n            : '파일을 읽지 못했습니다.',\n          'error'\n        );\n      }\n    }\n\n    uploadFile.addEventListener('change', function() {\n      handleUploadFile(uploadFile.files && uploadFile.files[0]);\n    });\n\n    uploadDropzone.addEventListener('dragover', function(event) {\n      event.preventDefault();\n      uploadDropzone.classList.add('dragging');\n    });\n\n    uploadDropzone.addEventListener('dragleave', function() {\n      uploadDropzone.classList.remove('dragging');\n    });\n\n    uploadDropzone.addEventListener('drop', function(event) {\n      event.preventDefault();\n      uploadDropzone.classList.remove('dragging');\n      const file = event.dataTransfer.files && event.dataTransfer.files[0];\n      handleUploadFile(file);\n    });\n\n    uploadReset.addEventListener('click', function() {\n      resetUploadPanel();\n      uploadFile.click();\n    });\n\n    uploadSave.addEventListener('click', function() {\n      if (!uploadParsedBooks.length) {\n        setUploadStatus('먼저 파일을 선택해 주십시오.', 'error');\n        return;\n      }\n\n      uploadSave.disabled = true;\n      uploadReset.disabled = true;\n      uploadSave.textContent = '시트에 넣는 중…';\n      setUploadStatus(\n        '업무용 업로드 목록 시트를 만들고 있습니다. 창을 닫지 마십시오.',\n        'info'\n      );\n\n      google.script.run\n        .withSuccessHandler(function(result) {\n          uploadSave.textContent = '시트 입력 완료';\n          uploadReset.disabled = false;\n          setUploadStatus(\n            '완료: ' + result.importedCount +\n            '권을 「' + result.sheetName +\n            '」 시트에 넣었습니다. 스프레드시트에서 구입할 책을 체크한 뒤 구입 검토함으로 보내십시오.',\n            'success'\n          );\n        })\n        .withFailureHandler(function(error) {\n          uploadSave.disabled = false;\n          uploadReset.disabled = false;\n          uploadSave.textContent = '이 목록을 시트에 넣기';\n          setUploadStatus(\n            error && error.message\n              ? error.message\n              : '시트에 저장하지 못했습니다.',\n            'error'\n          );\n        })\n        .HR_importBookListFile({\n          fileName: uploadFileName,\n          books: uploadParsedBooks\n        });\n    });\n\n    function showPagePanel(panelId) {\n      pagePanels.forEach(function(panel) {\n        panel.classList.toggle(\n          'active',\n          panel.id === panelId\n        );\n      });\n\n      viewTabs.forEach(function(tab) {\n        const active = tab.dataset.panel === panelId;\n\n        tab.classList.toggle('active', active);\n        tab.setAttribute(\n          'aria-selected',\n          active ? 'true' : 'false'\n        );\n      });\n\n      if (panelId === 'requestsPanel') {\n        stopBestsellerAuto();\n        loadSubmittedRequests(false);\n      } else if (panelId === 'applyPanel') {\n        startBestsellerAuto();\n      } else {\n        stopBestsellerAuto();\n      }\n\n      document.querySelector('.view-tabs').scrollIntoView({\n        behavior: 'smooth',\n        block: 'start'\n      });\n    }\n\n    form.addEventListener('submit', function(event) {\n      event.preventDefault();\n\n      const applicantValue = applicant.value.trim();\n      const queryValue = query.value.trim();\n\n      if (!queryValue) {\n        showMessage('책 제목이나 ISBN을 입력해 주십시오.', 'error');\n        query.focus();\n        return;\n      }\n\n      searchResults.classList.remove('show');\n      bookList.innerHTML = '';\n      books = [];\n      updateCartBar();\n      setSearchLoading(true);\n\n      showMessage(\n        'YES24에서 책을 찾고 있습니다…',\n        'info'\n      );\n\n      google.script.run\n        .withSuccessHandler(function(responseBooks) {\n          setSearchLoading(false);\n\n          books = Array.isArray(responseBooks)\n            ? responseBooks\n            : [];\n\n          if (!books.length) {\n            showMessage(\n              '검색 결과가 없습니다. 제목이나 ISBN을 다시 확인해 주십시오.',\n              'error'\n            );\n            return;\n          }\n\n          renderBooks();\n          searchResults.classList.add('show');\n\n          showMessage(\n            '검색 결과에서 신청할 책을 확인해 주십시오.',\n            'info'\n          );\n        })\n        .withFailureHandler(function(error) {\n          setSearchLoading(false);\n          showMessage(getErrorMessage(error), 'error');\n        })\n        .HR_searchRequestBooks(queryValue);\n    });\n\n    function getBookKey(book) {\n      const isbn = String(book && book.isbn || '')\n        .replace(/[^0-9Xx]/g, '')\n        .toUpperCase();\n\n      if (isbn) {\n        return 'isbn:' + isbn;\n      }\n\n      return 'book:' + [\n        book && book.title,\n        book && book.publisher,\n        book && (book.year || book.pubDate)\n      ].map(function(value) {\n        return String(value || '')\n          .trim()\n          .toLowerCase()\n          .replace(/\\s+/g, ' ');\n      }).join('|');\n    }\n\n    function findBookButton(bookKey) {\n      return Array.from(\n        document.querySelectorAll('.select-book-button')\n      ).find(function(button) {\n        return button.dataset.bookKey === bookKey;\n      }) || null;\n    }\n\n    function setVisibleBookSelection(button, selected) {\n      if (!button || button.dataset.done === 'true') {\n        return;\n      }\n\n      button.textContent = selected\n        ? '장바구니에 담김 ✓'\n        : '신청바구니에 담기';\n      button.classList.toggle('selected', selected);\n\n      if (button.closest('.book-card')) {\n        button.closest('.book-card')\n          .classList.toggle('selected', selected);\n      }\n    }\n\n    function renderBooks() {\n      bookList.innerHTML = '';\n\n      books.forEach(function(book, index) {\n        const card = document.createElement('article');\n        const cover = document.createElement('div');\n        const info = document.createElement('div');\n        const title = document.createElement('h3');\n        const meta = document.createElement('p');\n        const button = document.createElement('button');\n        const coverUrl = safeUrl(book.cover);\n        const bookKey = getBookKey(book);\n        const isInCart = cartBooks.has(bookKey);\n\n        card.className = 'book-card';\n        cover.className = 'book-cover';\n        info.className = 'book-info';\n        title.className = 'book-title';\n        meta.className = 'book-meta';\n        button.className = 'select-book-button';\n        button.type = 'button';\n        button.dataset.index = String(index);\n        button.dataset.bookKey = bookKey;\n\n        if (coverUrl) {\n          const image = document.createElement('img');\n\n          image.src = coverUrl;\n          image.alt = (book.title || '책') + ' 표지';\n          image.loading = 'lazy';\n          image.addEventListener('error', function() {\n            cover.textContent = '📖';\n          });\n          cover.appendChild(image);\n        } else {\n          cover.textContent = '📖';\n        }\n\n        title.textContent = book.title || '제목 없음';\n        meta.textContent = [\n          book.author,\n          book.publisher,\n          book.year ? book.year + '년' : book.pubDate,\n          formatPrice(book.price),\n          book.isbn\n        ].filter(Boolean).join(' · ');\n        button.textContent = isInCart\n          ? '장바구니에 담김 ✓'\n          : '신청바구니에 담기';\n        button.classList.toggle('selected', isInCart);\n        card.classList.toggle('selected', isInCart);\n        button.addEventListener('click', function() {\n          toggleCart(index, button);\n        });\n\n        info.appendChild(title);\n        info.appendChild(meta);\n        if (book.alreadyOwned || book.alreadyRequested) {\n          const badge = document.createElement('p');\n          badge.className = 'book-status-badge' + (book.alreadyOwned ? ' owned' : '');\n          badge.textContent = book.alreadyOwned\n            ? '📚 우리 학교에 이미 있는 책이에요'\n            : '⏳ 이미 ' + (book.alreadyRequested > 1 ? book.alreadyRequested + '명이 ' : '') + '신청한 책이에요. 담으면 함께 원하는 분으로 기록돼요';\n          info.appendChild(badge);\n        }\n        info.appendChild(button);\n        card.appendChild(cover);\n        card.appendChild(info);\n        bookList.appendChild(card);\n      });\n    }\n\n    function toggleCart(index, selectedButton) {\n      const book = books[index];\n      const bookKey = getBookKey(book);\n\n      if (\n        !book ||\n        !bookKey ||\n        selectedButton.dataset.done === 'true'\n      ) {\n        return;\n      }\n\n      if (cartBooks.has(bookKey)) {\n        cartBooks.delete(bookKey);\n        setVisibleBookSelection(selectedButton, false);\n      } else {\n        if (cartBooks.size >= 20) {\n          showMessage(\n            '한 번에 최대 20권까지 장바구니에 담을 수 있습니다.',\n            'error'\n          );\n          return;\n        }\n\n        cartBooks.set(bookKey, book);\n        setVisibleBookSelection(selectedButton, true);\n        showMessage(\n          '장바구니에 담았습니다. 다른 책을 계속 검색해도 그대로 유지됩니다.',\n          'success'\n        );\n      }\n\n      updateCartBar();\n    }\n\n    function updateCartBar() {\n      const count = cartBooks.size;\n\n      cartCount.textContent = String(count);\n      cartSubmit.disabled =\n        submissionInProgress || count === 0;\n      cartDrawerSubmit.disabled =\n        submissionInProgress || count === 0;\n      cartClear.disabled =\n        submissionInProgress || count === 0;\n      cartBar.classList.toggle('show', count > 0);\n      renderCartDrawer();\n    }\n\n    function openCartDrawer() {\n      if (!cartBooks.size) {\n        return;\n      }\n\n      renderCartDrawer();\n      cartDrawer.classList.add('show');\n      cartToggle.setAttribute('aria-expanded', 'true');\n      document.body.style.overflow = 'hidden';\n    }\n\n    function closeCartDrawer() {\n      cartDrawer.classList.remove('show');\n      cartToggle.setAttribute('aria-expanded', 'false');\n      document.body.style.overflow = '';\n    }\n\n    function clearCart() {\n      if (submissionInProgress) {\n        return;\n      }\n\n      cartBooks = new Map();\n\n      document\n        .querySelectorAll('.select-book-button')\n        .forEach(function(button) {\n          setVisibleBookSelection(button, false);\n        });\n\n      closeCartDrawer();\n      updateCartBar();\n      showMessage('신청바구니를 비웠습니다.', 'info');\n    }\n\n    function removeCartBook(bookKey) {\n      if (submissionInProgress) {\n        return;\n      }\n\n      cartBooks.delete(bookKey);\n      setVisibleBookSelection(\n        findBookButton(bookKey),\n        false\n      );\n      updateCartBar();\n\n      if (!cartBooks.size) {\n        closeCartDrawer();\n      }\n    }\n\n    function renderCartDrawer() {\n      cartList.innerHTML = '';\n\n      if (!cartBooks.size) {\n        const empty = document.createElement('div');\n        empty.className = 'cart-empty';\n        empty.textContent =\n          '아직 담은 책이 없습니다.\\n검색 결과에서 원하는 책을 담아 주십시오.';\n        cartList.appendChild(empty);\n        return;\n      }\n\n      cartBooks.forEach(function(book, bookKey) {\n        const item = document.createElement('article');\n        const cover = document.createElement('div');\n        const info = document.createElement('div');\n        const title = document.createElement('h3');\n        const meta = document.createElement('p');\n        const remove = document.createElement('button');\n        const coverUrl = safeUrl(book.cover);\n\n        item.className = 'cart-item';\n        cover.className = 'cart-item-cover';\n        title.className = 'cart-item-title';\n        meta.className = 'cart-item-meta';\n        remove.className = 'cart-remove';\n        remove.type = 'button';\n        remove.textContent = '빼기';\n\n        if (coverUrl) {\n          const image = document.createElement('img');\n          image.src = coverUrl;\n          image.alt = (book.title || '책') + ' 표지';\n          image.loading = 'lazy';\n          image.addEventListener('error', function() {\n            cover.textContent = '📖';\n          });\n          cover.appendChild(image);\n        } else {\n          cover.textContent = '📖';\n        }\n\n        title.textContent = book.title || '제목 없음';\n        meta.textContent = [\n          book.author,\n          book.publisher,\n          book.year ? book.year + '년' : book.pubDate\n        ].filter(Boolean).join(' · ');\n\n        remove.addEventListener('click', function() {\n          removeCartBook(bookKey);\n        });\n\n        info.appendChild(title);\n        info.appendChild(meta);\n        item.appendChild(cover);\n        item.appendChild(info);\n        item.appendChild(remove);\n        cartList.appendChild(item);\n      });\n    }\n\n    function submitCart() {\n      const selectedEntries =\n        Array.from(cartBooks.entries());\n      const selectedKeys = selectedEntries.map(function(entry) {\n        return entry[0];\n      });\n      const selectedBooks = selectedEntries.map(function(entry) {\n        return entry[1];\n      });\n\n      if (!selectedBooks.length) {\n        showMessage(\n          '신청할 책을 한 권 이상 장바구니에 담아 주십시오.',\n          'error'\n        );\n        return;\n      }\n\n      if (submissionInProgress) {\n        return;\n      }\n\n      submissionInProgress = true;\n      setCartLoading(true);\n      showMessage(\n        '장바구니의 ' +\n        selectedBooks.length +\n        '권을 안전하게 접수하고 있습니다…',\n        'info'\n      );\n\n      const requestPayload = {\n        applicant: applicant.value.trim(),\n        audience: window.HR_AUDIENCE || 'staff',\n        staffCode: window.HR_STAFF_CODE || '',\n        books: selectedBooks,\n        requestId: createRequestId()\n      };\n\n      sendCartRequest(\n        requestPayload,\n        selectedKeys,\n        selectedBooks,\n        0\n      );\n    }\n\n    function sendCartRequest(\n      requestPayload,\n      selectedKeys,\n      selectedBooks,\n      retryCount\n    ) {\n      google.script.run\n        .withSuccessHandler(function(response) {\n          submissionInProgress = false;\n          setCartLoading(false);\n          handleCartResponse(\n            response || {},\n            selectedKeys,\n            selectedBooks\n          );\n        })\n        .withFailureHandler(function(error) {\n          if (\n            retryCount < 2 &&\n            isTemporarySubmitError(error)\n          ) {\n            const waitTime =\n              900 +\n              retryCount * 1300 +\n              Math.floor(Math.random() * 700);\n\n            showMessage(\n              '신청이 몰려 순서대로 처리하고 있습니다. 잠시만 기다려 주십시오…',\n              'info'\n            );\n\n            window.setTimeout(function() {\n              sendCartRequest(\n                requestPayload,\n                selectedKeys,\n                selectedBooks,\n                retryCount + 1\n              );\n            }, waitTime);\n            return;\n          }\n\n          submissionInProgress = false;\n          setCartLoading(false);\n          const failText = getErrorMessage(error);\n\n          // 꼭 알아야 하는 안내(신청 한도, 학부모 정보 빠짐)는 눈에 띄는 팝업으로 보여 줍니다.\n          if (/최대 \\d+권까지 신청할 수 있습니다/.test(failText)) {\n            message.className = 'message';\n            showNoticeModal({\n              type: 'duplicate',\n              icon: '✋',\n              title: '신청 한도를 넘었습니다',\n              description: failText\n            });\n          } else if (/학년·반·번호를 모두 골라/.test(failText)) {\n            message.className = 'message';\n            showNoticeModal({\n              type: 'duplicate',\n              icon: '🧒',\n              title: '우리 아이 정보를 골라 주십시오',\n              description: failText\n            });\n          } else {\n            showMessage(failText, 'error');\n          }\n        })\n        .HR_submitBookRequests(requestPayload);\n    }\n\n    function isTemporarySubmitError(error) {\n      const text = getErrorMessage(error).toLowerCase();\n\n      return (\n        text.includes('현재 신청자가 많') ||\n        text.includes('too many') ||\n        text.includes('timeout') ||\n        text.includes('timed out') ||\n        text.includes('lock') ||\n        text.includes('service invoked') ||\n        text.includes('일시적') ||\n        text.includes('시간 초과')\n      );\n    }\n\n    function createRequestId() {\n      if (\n        window.crypto &&\n        typeof window.crypto.randomUUID === 'function'\n      ) {\n        return window.crypto.randomUUID();\n      }\n\n      return (\n        String(Date.now()) + '-' +\n        Math.random().toString(36).slice(2)\n      );\n    }\n\n        function handleCartResponse(\n      response,\n      selectedKeys,\n      selectedBooks\n    ) {\n      // 신청이 끝났으니 남은 신청 가능 권수를 다시 계산해서 보여 줍니다.\n      submittedRequestsLoaded = false;\n      if (response && response.quota && typeof window.HR_setQuota === 'function') {\n        window.HR_setQuota(response.quota);          // 서버가 방금 센 값을 바로 보여 줍니다.\n      } else if (typeof window.HR_refreshQuota === 'function') {\n        window.HR_refreshQuota();\n      }\n      const responseItems = Array.isArray(response.results)\n        ? response.results\n        : [];\n      const successTitles = [];\n      const duplicateTitles = [];\n      const ownedTitles = [];\n      const errorTitles = [];\n\n      responseItems.forEach(function(item, itemIndex) {\n        const book = selectedBooks[itemIndex] || {};\n        const bookKey =\n          selectedKeys[itemIndex] || getBookKey(book);\n        const button = findBookButton(bookKey);\n        const displayTitle =\n          item.title || book.title || '제목 없음';\n\n        if (item.success) {\n          successTitles.push(displayTitle);\n          cartBooks.delete(bookKey);\n          markBookCardDone(button, '신청 완료 ✓');\n        } else if (item.owned) {\n          ownedTitles.push(displayTitle);\n          cartBooks.delete(bookKey);\n          markBookCardDone(button, '소장 중');\n        } else if (item.duplicateRequest) {\n          duplicateTitles.push(displayTitle);\n          cartBooks.delete(bookKey);\n          markBookCardDone(button, '이미 신청됨');\n        } else {\n          errorTitles.push(displayTitle);\n          setVisibleBookSelection(button, true);\n        }\n      });\n\n      if (successTitles.length) {\n        submittedRequestsLoaded = false;\n      }\n\n      updateCartBar();\n\n      if (!cartBooks.size) {\n        closeCartDrawer();\n      }\n\n      const descriptionLines = [];\n\n      function addResultGroup(icon, label, titles) {\n        if (!titles.length) {\n          return;\n        }\n\n        descriptionLines.push(\n          icon + ' ' + label + ' (' + titles.length + '권)\\n' +\n          titles.map(function(title) {\n            return '• ' + title;\n          }).join('\\n')\n        );\n      }\n\n      addResultGroup('✅', '신청 완료', successTitles);\n      addResultGroup('🔁', '이미 신청된 도서', duplicateTitles);\n      addResultGroup('🏫', '학교에서 소장 중', ownedTitles);\n      addResultGroup('⚠️', '장바구니에 남겨 둔 도서', errorTitles);\n\n      showNoticeModal({\n        type: successTitles.length ? 'success' : 'duplicate',\n        icon: successTitles.length ? '✅' : '📚',\n        title: successTitles.length\n          ? '신청 완료 · ' + successTitles.length + '권'\n          : '신청 결과를 확인해 주십시오.',\n        description: descriptionLines.join('\\n\\n')\n      });\n\n      showMessage(\n        successTitles.length\n          ? successTitles.length +\n            '권 신청을 완료했습니다.'\n          : '새로 신청된 책이 없습니다.',\n        successTitles.length ? 'success' : 'info'\n      );\n    }\n\n    function markBookCardDone(button, label) {\n      if (!button) {\n        return;\n      }\n\n      button.dataset.done = 'true';\n      button.textContent = label;\n      button.disabled = true;\n      button.classList.remove('selected');\n      button.classList.add('done');\n\n      if (button.closest('.book-card')) {\n        button.closest('.book-card')\n          .classList.remove('selected');\n        button.closest('.book-card')\n          .classList.add('done');\n      }\n    }\n\n    function setSearchLoading(isLoading) {\n      submitButton.disabled = isLoading;\n      submitButton.textContent =\n        isLoading\n          ? '검색하고 있습니다…'\n          : '책 검색하기';\n    }\n\n    function setCartLoading(isLoading) {\n      const buttons =\n        document.querySelectorAll('.select-book-button');\n\n      buttons.forEach(function(button) {\n        button.disabled =\n          isLoading ||\n          button.dataset.done === 'true';\n      });\n\n      query.disabled = isLoading;\n      applicant.disabled = isLoading;\n      submitButton.disabled = isLoading;\n      cartSubmit.disabled =\n        isLoading || cartBooks.size === 0;\n      cartDrawerSubmit.disabled =\n        isLoading || cartBooks.size === 0;\n      cartClear.disabled = isLoading;\n      cartSubmit.textContent = isLoading\n        ? '안전하게 접수하는 중…'\n        : '선택한 도서 신청하기';\n    }\n\n\n    function loadSubmittedRequests(forceRefresh) {\n      if (\n        submittedRequestsLoaded &&\n        !forceRefresh\n      ) {\n        return;\n      }\n\n      requestRefresh.disabled = true;\n      requestRefresh.textContent = '불러오는 중…';\n      requestList.innerHTML = '';\n      requestSummary.classList.remove('show');\n      requestListStatus.textContent =\n        '신청된 희망도서를 불러오고 있습니다…';\n      requestListStatus.classList.remove('hide');\n\n      google.script.run\n        .withSuccessHandler(function(response) {\n          requestRefresh.disabled = false;\n          requestRefresh.textContent = '새로고침';\n          submittedRequestsLoaded = true;\n          renderSubmittedRequests(response || {});\n        })\n        .withFailureHandler(function(error) {\n          requestRefresh.disabled = false;\n          requestRefresh.textContent = '다시 불러오기';\n          requestListStatus.textContent =\n            getErrorMessage(error);\n          requestListStatus.classList.remove('hide');\n        })\n        .HR_getSubmittedRequests(window.HR_AUDIENCE || 'staff', window.HR_STAFF_CODE || '', applicant.value.trim());\n    }\n\n    function renderSubmittedRequests(response) {\n      if (typeof window.HR_refreshQuota === 'function') window.HR_refreshQuota();\n      if (response.needApplicant) {\n        submittedRequestItems = [];\n        requestActiveCount.textContent = '0';\n        requestCancelledCount.textContent = '0';\n        requestList.innerHTML = '';\n        requestSummary.classList.remove('show');\n        requestListStatus.textContent = (window.HR_AUDIENCE === 'parent')\n          ? '「도서 신청」 화면에서 학년·반·번호를 먼저 선택해 주십시오. 선택하신 학생의 신청만 보여 드립니다.'\n          : '「도서 신청」 화면에서 신청자 이름을 먼저 입력해 주십시오. 입력하신 분의 신청만 보여 드립니다.';\n        requestListStatus.classList.remove('hide');\n        return;\n      }\n      const responseItems = Array.isArray(response.items)\n        ? response.items\n        : [];\n      const activeItems = responseItems.filter(function(item) {\n        return !/취소/.test(\n          String(item.status || '')\n        );\n      });\n      const cancelledItems = responseItems.filter(function(item) {\n        return /취소/.test(\n          String(item.status || '')\n        );\n      });\n      const items = requestListFilter === 'cancelled'\n        ? cancelledItems\n        : activeItems;\n      const totalCount = items.length;\n      const shownCount = items.length;\n\n      submittedRequestItems = responseItems;\n      requestActiveCount.textContent =\n        String(activeItems.length);\n      requestCancelledCount.textContent =\n        String(cancelledItems.length);\n\n      requestFilterTabs.forEach(function(tab) {\n        const selected =\n          tab.dataset.requestFilter ===\n          requestListFilter;\n\n        tab.classList.toggle('active', selected);\n        tab.setAttribute(\n          'aria-selected',\n          selected ? 'true' : 'false'\n        );\n      });\n\n      requestList.innerHTML = '';\n      requestSummary.textContent =\n        requestListFilter === 'cancelled'\n          ? '신청 취소 ' + totalCount + '권'\n          : '신청 도서 ' + totalCount + '권';\n      requestSummary.classList.add('show');\n\n      if (!items.length) {\n        requestListStatus.textContent =\n          requestListFilter === 'cancelled'\n            ? '현재 취소된 신청이 없습니다.'\n            : '아직 신청된 희망도서가 없습니다.';\n        requestListStatus.classList.remove('hide');\n        return;\n      }\n\n      requestListStatus.classList.add('hide');\n\n      items.forEach(function(item, index) {\n        const card = document.createElement('article');\n        const top = document.createElement('div');\n        const sequence = document.createElement('span');\n        const status = document.createElement('span');\n        const title = document.createElement('h3');\n        const meta = document.createElement('p');\n        const date = document.createElement('p');\n        const details = document.createElement('div');\n        const coverUrl = safeUrl(item.cover);\n        const detailHint = document.createElement('button');\n        const actions = document.createElement('div');\n        const cancelButton = document.createElement('button');\n        const statusText =\n          String(item.status || '신청 접수').trim();\n\n        card.className = 'request-item';\n        card.tabIndex = 0;\n        card.setAttribute('role', 'button');\n        card.setAttribute('aria-expanded', 'false');\n        top.className = 'request-item-top';\n        sequence.className = 'request-sequence';\n        status.className =\n          'request-status-badge ' +\n          getRequestStatusClass(statusText);\n        date.className = 'request-date';\n        details.className = 'request-biblio';\n        detailHint.className = 'request-detail-button';\n        detailHint.type = 'button';\n        detailHint.textContent = '서지정보 보기';\n        actions.className = 'request-card-actions';\n        cancelButton.className = 'request-cancel-button';\n        cancelButton.type = 'button';\n\n        sequence.textContent =\n          item.sequence\n            ? '#' + item.sequence\n            : '#' + (totalCount - index);\n        status.textContent = statusText;\n        title.textContent = item.title || '제목 없음';\n        meta.textContent = [\n          item.author,\n          item.publisher,\n          item.year ? item.year + '년' : ''\n        ].filter(Boolean).join(' · ') ||\n          '도서 정보를 확인하고 있습니다.';\n        date.textContent = item.requestedAt\n          ? '신청일 ' + item.requestedAt\n          : '';\n\n        if (coverUrl) {\n          const coverBox = document.createElement('div');\n          const coverImage = document.createElement('img');\n\n          coverBox.className = 'request-biblio-cover';\n          coverImage.src = coverUrl;\n          coverImage.alt =\n            (item.title || '책') + ' 표지';\n          coverImage.loading = 'lazy';\n          coverImage.addEventListener('error', function() {\n            coverBox.remove();\n          });\n          coverBox.appendChild(coverImage);\n          details.appendChild(coverBox);\n        }\n\n        [\n          ['저자', item.author],\n          ['출판사', item.publisher],\n          ['출판연도', item.year ? item.year + '년' : ''],\n          ['ISBN13', item.isbn],\n          ['정가', item.price\n            ? String(item.price).replace(/\\s*원$/, '') + '원'\n            : ''],\n          ['분야', item.category],\n          ['쪽수', item.pages]\n        ].forEach(function(entry) {\n          if (!entry[1]) {\n            return;\n          }\n\n          const row = document.createElement('div');\n          const label = document.createElement('strong');\n          const value = document.createElement('span');\n\n          row.className = 'request-biblio-row';\n          label.textContent = entry[0];\n          value.textContent = entry[1];\n          row.appendChild(label);\n          row.appendChild(value);\n          details.appendChild(row);\n        });\n\n        if (safeUrl(item.link)) {\n          const row = document.createElement('div');\n          const label = document.createElement('strong');\n          const link = document.createElement('a');\n\n          row.className = 'request-biblio-row';\n          label.textContent = '도서 정보';\n          link.href = safeUrl(item.link);\n          link.target = '_blank';\n          link.rel = 'noopener noreferrer';\n          link.textContent = 'YES24에서 보기 ↗';\n          link.addEventListener('click', function(event) {\n            event.stopPropagation();\n          });\n          row.appendChild(label);\n          row.appendChild(link);\n          details.appendChild(row);\n        }\n\n        if (!details.children.length) {\n          const empty = document.createElement('p');\n          empty.className = 'request-biblio-empty';\n          empty.textContent =\n            '추가 서지정보를 확인하고 있습니다.';\n          details.appendChild(empty);\n        }\n\n        function toggleDetails(event) {\n          if (\n            event &&\n            event.type === 'keydown' &&\n            event.key !== 'Enter' &&\n            event.key !== ' '\n          ) {\n            return;\n          }\n\n          if (event && event.type === 'keydown') {\n            event.preventDefault();\n          }\n\n          const expanded =\n            card.classList.toggle('expanded');\n          card.setAttribute(\n            'aria-expanded',\n            expanded ? 'true' : 'false'\n          );\n          detailHint.textContent = expanded\n            ? '서지정보 접기'\n            : '서지정보 보기';\n        }\n\n        card.addEventListener('click', toggleDetails);\n        card.addEventListener('keydown', toggleDetails);\n        detailHint.addEventListener('click', function(event) {\n          event.stopPropagation();\n          toggleDetails();\n        });\n\n        if (/취소/.test(statusText)) {\n          cancelButton.classList.add('restore');\n          cancelButton.textContent = '취소 되돌리기';\n          cancelButton.addEventListener('click', function(event) {\n            event.stopPropagation();\n            undoRequestFromList(\n              item,\n              cancelButton\n            );\n          });\n        } else if (/구입|입수|선정|반려/.test(statusText)) {\n          cancelButton.disabled = true;\n          cancelButton.textContent = '취소 불가';\n        } else {\n          cancelButton.textContent = '신청 취소';\n          cancelButton.addEventListener('click', function(event) {\n            event.stopPropagation();\n            openCancelRequest(\n              item,\n              cancelButton,\n              status\n            );\n          });\n        }\n\n        top.appendChild(sequence);\n        top.appendChild(status);\n        card.appendChild(top);\n        if (coverUrl) {\n          const thumb = document.createElement('img');\n          thumb.className = 'request-thumb';\n          thumb.src = coverUrl;\n          thumb.alt = (item.title || '책') + ' 표지';\n          thumb.loading = 'lazy';\n          thumb.addEventListener('error', function() { thumb.remove(); });\n          card.appendChild(thumb);\n        }\n        card.appendChild(title);\n        card.appendChild(meta);\n\n        if (date.textContent) {\n          card.appendChild(date);\n        }\n\n        card.appendChild(details);\n        actions.appendChild(detailHint);\n        actions.appendChild(cancelButton);\n        card.appendChild(actions);\n        requestList.appendChild(card);\n      });\n    }\n\n    function getRequestStatusClass(statusText) {\n      if (/오류|실패|취소|반려/.test(statusText)) {\n        return 'issue';\n      }\n\n      if (/완료|선정|구입|입수|처리 완료/.test(statusText)) {\n        return 'done';\n      }\n\n      return 'pending';\n    }\n\n    function openCancelRequest(\n      item,\n      button,\n      statusElement\n    ) {\n      cancelCandidate = {\n        item: item,\n        button: button,\n        statusElement: statusElement\n      };\n      cancelBook.textContent =\n        '『' + String(item.title || '제목 없음') + '』';\n      cancelConfirm.disabled = false;\n      cancelConfirm.textContent = '신청 취소하기';\n      cancelModal.classList.add('show');\n\n      window.setTimeout(function() {\n        cancelKeep.focus();\n      }, 30);\n    }\n\n    function closeCancelModal() {\n      cancelModal.classList.remove('show');\n      cancelCandidate = null;\n      cancelConfirm.disabled = false;\n      cancelConfirm.textContent = '신청 취소하기';\n    }\n\n    function confirmCancelRequest() {\n      if (!cancelCandidate) {\n        return;\n      }\n\n      const candidate = cancelCandidate;\n      cancelConfirm.disabled = true;\n      cancelConfirm.textContent = '취소 중…';\n\n      google.script.run\n        .withSuccessHandler(function(result) {\n          if (!result || result.success === false) {\n            const reason =\n              result && result.message\n                ? result.message\n                : '이 도서는 현재 취소할 수 없습니다.';\n            closeCancelModal();\n            showNoticeModal({\n              type: 'duplicate',\n              icon: 'ℹ️',\n              title: '취소할 수 없습니다',\n              book: candidate.item.title,\n              description: reason\n            });\n            return;\n          }\n\n          submittedRequestsLoaded = false;\n          closeCancelModal();\n          loadSubmittedRequests(true);\n\n          showNoticeModal({\n            type: 'success',\n            icon: '✅',\n            title: '신청이 취소되었습니다',\n            book: candidate.item.title,\n            description:\n              '신청 목록에서는 이 도서 한 권만 삭제했습니다. 취소한 내용은 안전하게 보관되어 바로 되돌릴 수 있습니다.',\n            undoId: result.cancelId\n          });\n        })\n        .withFailureHandler(function(error) {\n          const reason = getErrorMessage(error);\n          closeCancelModal();\n          showNoticeModal({\n            type: 'duplicate',\n            icon: '⚠️',\n            title: '취소하지 못했습니다',\n            book: candidate.item.title,\n            description: reason\n          });\n        })\n        .HR_cancelSubmittedRequest({\n          audience: window.HR_AUDIENCE || 'staff',\n          staffCode: window.HR_STAFF_CODE || '',\n          applicant: applicant.value.trim(),\n          row: candidate.item.row,\n          title: candidate.item.title,\n          isbn: candidate.item.isbn\n        });\n    }\n\n    function undoRequestFromList(item, button) {\n      if (!item || !button) {\n        return;\n      }\n\n      button.disabled = true;\n      button.textContent = '복구 중…';\n\n      const onSuccess = function(result) {\n        if (!result || result.success === false) {\n          const reason =\n            result && result.message\n              ? result.message\n              : '신청 취소를 되돌리지 못했습니다.';\n\n          button.disabled = false;\n          button.textContent = '취소 되돌리기';\n          showNoticeModal({\n            type: 'duplicate',\n            icon: '⚠️',\n            title: '되돌리지 못했습니다',\n            book: item.title,\n            description: reason\n          });\n          return;\n        }\n\n        submittedRequestsLoaded = false;\n        loadSubmittedRequests(true);\n\n        showNoticeModal({\n          type: 'success',\n          icon: '↩️',\n          title: result.alreadyRestored\n            ? '이미 신청 목록에 있습니다'\n            : '신청 취소를 되돌렸습니다',\n          book: item.title,\n          description: result.alreadyRestored\n            ? '같은 도서가 이미 신청 목록에 있어 중복으로 넣지 않았습니다.'\n            : '취소 전 신청 내용이 희망도서 목록에 다시 들어갔습니다.'\n        });\n      };\n\n      const onFailure = function(error) {\n        button.disabled = false;\n        button.textContent = '취소 되돌리기';\n        showNoticeModal({\n          type: 'duplicate',\n          icon: '⚠️',\n          title: '되돌리지 못했습니다',\n          book: item.title,\n          description: getErrorMessage(error)\n        });\n      };\n\n      if (item.cancelId) {\n        google.script.run\n          .withSuccessHandler(onSuccess)\n          .withFailureHandler(onFailure)\n          .HR_undoCancelledRequest({\n            audience: window.HR_AUDIENCE || 'staff',\n            staffCode: window.HR_STAFF_CODE || '',\n            applicant: applicant.value.trim(),\n            cancelId: item.cancelId\n          });\n        return;\n      }\n\n      google.script.run\n        .withSuccessHandler(onSuccess)\n        .withFailureHandler(onFailure)\n        .HR_restoreLegacyCancelledRequest({\n          row: item.row,\n          title: item.title,\n          isbn: item.isbn\n        });\n    }\n\n    function undoCancelledRequest() {\n      if (!undoCandidate) {\n        return;\n      }\n\n      const candidate = undoCandidate;\n      noticeUndo.disabled = true;\n      noticeUndo.textContent = '복구 중…';\n\n      google.script.run\n        .withSuccessHandler(function(result) {\n          if (!result || result.success === false) {\n            const reason =\n              result && result.message\n                ? result.message\n                : '신청을 되돌리지 못했습니다.';\n\n            showNoticeModal({\n              type: 'duplicate',\n              icon: '⚠️',\n              title: '되돌리지 못했습니다',\n              book: candidate.title,\n              description: reason,\n              undoId: candidate.cancelId\n            });\n            return;\n          }\n\n          submittedRequestsLoaded = false;\n          closeNoticeModal();\n          loadSubmittedRequests(true);\n\n          showNoticeModal({\n            type: 'success',\n            icon: '↩️',\n            title: result.alreadyRestored\n              ? '이미 신청 목록에 있습니다'\n              : '신청 취소를 되돌렸습니다',\n            book: candidate.title,\n            description: result.alreadyRestored\n              ? '같은 도서가 이미 신청 목록에 있어 중복으로 넣지 않았습니다.'\n              : '취소 전 신청 내용이 희망도서 목록에 다시 들어갔습니다.'\n          });\n        })\n        .withFailureHandler(function(error) {\n          showNoticeModal({\n            type: 'duplicate',\n            icon: '⚠️',\n            title: '되돌리지 못했습니다',\n            book: candidate.title,\n            description: getErrorMessage(error),\n            undoId: candidate.cancelId\n          });\n        })\n        .HR_undoCancelledRequest({\n          audience: window.HR_AUDIENCE || 'staff',\n          staffCode: window.HR_STAFF_CODE || '',\n          applicant: applicant.value.trim(),\n          cancelId: candidate.cancelId\n        });\n    }\n\n    function closeGuidePopover() {\n      guideHover.classList.remove('open');\n      guideHoverTrigger.setAttribute(\n        'aria-expanded',\n        'false'\n      );\n    }\n\n    function safeUrl(value) {\n      const text = String(value || '').trim();\n\n      return /^https?:\\/\\//i.test(text)\n        ? text\n        : '';\n    }\n\n    function formatPrice(value) {\n      const number = Number(value || 0);\n\n      return number > 0\n        ? number.toLocaleString('ko-KR') + '원'\n        : '';\n    }\n\n    function showMessage(text, type) {\n      message.textContent = text;\n      message.className =\n        'message show ' + (type || 'info');\n    }\n\n    function showNoticeModal(options) {\n      const notice = options || {};\n\n      noticeModal.className =\n        'notice-modal show ' + (notice.type || 'duplicate');\n      noticeIcon.textContent = notice.icon || '📚';\n      noticeTitle.textContent =\n        notice.title || '안내';\n      noticeBook.textContent =\n        notice.book ? '『' + notice.book + '』' : '';\n      noticeDescription.textContent =\n        notice.description || '';\n\n      undoCandidate = notice.undoId\n        ? {\n          cancelId: notice.undoId,\n          title: notice.book || ''\n        }\n        : null;\n\n      noticeUndo.hidden = !undoCandidate;\n      noticeUndo.disabled = false;\n      noticeUndo.textContent = '신청 취소 되돌리기';\n      noticeActions.classList.toggle(\n        'single',\n        !undoCandidate\n      );\n\n      window.setTimeout(function() {\n        if (undoCandidate) {\n          noticeUndo.focus();\n        } else {\n          noticeClose.focus();\n        }\n      }, 30);\n    }\n\n    function closeNoticeModal() {\n      noticeModal.className = 'notice-modal';\n      undoCandidate = null;\n      noticeUndo.hidden = true;\n      noticeUndo.disabled = false;\n      noticeUndo.textContent = '신청 취소 되돌리기';\n      noticeActions.classList.add('single');\n    }\n\n    function getErrorMessage(error) {\n      return error && error.message\n        ? String(error.message).replace(/^(Error|Exception):\\s*/, '')\n        : String(\n          error ||\n          '처리 중 오류가 발생했습니다.'\n        );\n    }\n\n    let bsData = {};\n    let bsActiveCategory = 'all';\n\n    function loadBestsellers() {\n      const status = document.getElementById('bsStatus');\n\n      status.textContent = '베스트셀러를 불러오고 있습니다…';\n      status.classList.add('show');\n\n      google.script.run\n        .withSuccessHandler(function(response) {\n          bsData = response && response.categories\n            ? response.categories\n            : {};\n          renderBestsellers(bsActiveCategory);\n        })\n        .withFailureHandler(function() {\n          status.textContent =\n            '베스트셀러를 잠시 불러오지 못했습니다. 책 검색은 그대로 이용할 수 있습니다.';\n          status.classList.add('show');\n        })\n        .HR_getBestsellers();\n    }\n\n    function renderBestsellers(categoryKey) {\n      const list = document.getElementById('bsList');\n      const status = document.getElementById('bsStatus');\n      const category = bsData[categoryKey] || {};\n      const categoryBooks = Array.isArray(category.books)\n        ? category.books\n        : [];\n\n      stopBestsellerAuto();\n      list.innerHTML = '';\n\n      document.querySelectorAll('.bs-tab').forEach(function(tab) {\n        const active =\n          tab.dataset.bsCategory === categoryKey;\n\n        tab.classList.toggle('active', active);\n        tab.setAttribute(\n          'aria-selected',\n          active ? 'true' : 'false'\n        );\n      });\n\n      if (!categoryBooks.length) {\n        status.textContent =\n          '이 분류의 베스트셀러가 아직 없습니다.';\n        status.classList.add('show');\n        return;\n      }\n\n      status.classList.remove('show');\n\n      categoryBooks.forEach(function(book, index) {\n        const card = document.createElement('article');\n        const rank = document.createElement('span');\n        const cover = document.createElement('div');\n        const title = document.createElement('h3');\n        const meta = document.createElement('p');\n        const actions = document.createElement('div');\n        const requestButton = document.createElement('button');\n        const infoLink = document.createElement('a');\n        const coverUrl = safeUrl(book.cover);\n        const itemLink = safeUrl(book.link);\n\n        card.className = 'bs-card';\n        rank.className = 'bs-rank';\n        cover.className = 'bs-cover';\n        title.className = 'bs-book-title';\n        meta.className = 'bs-meta';\n        actions.className = 'bs-actions';\n        requestButton.className = 'bs-request';\n        infoLink.className = 'bs-link';\n\n        rank.textContent = String(book.rank || index + 1);\n        title.textContent = book.title || '제목 없음';\n        meta.textContent = [\n          book.author,\n          book.publisher,\n          book.year ? book.year + '년' : '',\n          formatPrice(book.price)\n        ].filter(Boolean).join(' · ');\n\n        if (coverUrl) {\n          const image = document.createElement('img');\n\n          image.src = coverUrl;\n          image.alt = (book.title || '책') + ' 표지';\n          image.loading = 'lazy';\n          image.addEventListener('error', function() {\n            cover.textContent = '📖';\n          });\n          cover.appendChild(image);\n        } else {\n          cover.textContent = '📖';\n        }\n\n        requestButton.type = 'button';\n        requestButton.textContent = '이 책 신청하기';\n        requestButton.addEventListener('click', function() {\n          searchBestsellerBook(book);\n        });\n\n        if (itemLink) {\n          infoLink.href = itemLink;\n          infoLink.target = '_blank';\n          infoLink.rel = 'noopener noreferrer';\n        } else {\n          infoLink.href = '#';\n          infoLink.addEventListener('click', function(event) {\n            event.preventDefault();\n          });\n        }\n        infoLink.textContent = '정보';\n\n        actions.appendChild(requestButton);\n        actions.appendChild(infoLink);\n        card.appendChild(rank);\n        card.appendChild(cover);\n        card.appendChild(title);\n        card.appendChild(meta);\n        card.appendChild(actions);\n        list.appendChild(card);\n      });\n\n      const originalCards =\n        Array.from(list.querySelectorAll('.bs-card'));\n\n      originalCards.forEach(function(originalCard, index) {\n        const cloneCard = originalCard.cloneNode(true);\n        const cloneRequest =\n          cloneCard.querySelector('.bs-request');\n\n        cloneCard.classList.add('bs-card-clone');\n\n        if (cloneRequest) {\n          cloneRequest.addEventListener('click', function() {\n            searchBestsellerBook(categoryBooks[index]);\n          });\n        }\n\n        list.appendChild(cloneCard);\n      });\n\n      list.dataset.bookCount =\n        String(categoryBooks.length);\n      bsListElement.scrollLeft = 0;\n      startBestsellerAuto();\n    }\n\n    function getBestsellerScrollStep() {\n      const firstCard =\n        bsListElement.querySelector('.bs-card');\n\n      if (!firstCard) {\n        return 220;\n      }\n\n      return firstCard.getBoundingClientRect().width + 12;\n    }\n\n    function scrollBestsellerCards(direction) {\n      const step = getBestsellerScrollStep();\n      const bookCount = Number(\n        bsListElement.dataset.bookCount || 0\n      );\n      const cycleWidth = step * bookCount;\n\n      if (!bookCount || cycleWidth <= 1) {\n        return;\n      }\n\n      if (direction > 0) {\n        if (\n          bsListElement.scrollLeft >=\n          cycleWidth - 8\n        ) {\n          bsListElement.scrollLeft =\n            Math.max(\n              0,\n              bsListElement.scrollLeft -\n              cycleWidth\n            );\n        }\n\n        bsListElement.scrollBy({\n          left: step,\n          behavior: 'smooth'\n        });\n        return;\n      }\n\n      if (bsListElement.scrollLeft <= 8) {\n        bsListElement.scrollLeft = cycleWidth;\n      }\n\n      bsListElement.scrollBy({\n        left: -step,\n        behavior: 'smooth'\n      });\n    }\n\n    function stopBestsellerAuto() {\n      if (bsAutoTimer) {\n        window.clearInterval(bsAutoTimer);\n        bsAutoTimer = null;\n      }\n    }\n\n    function startBestsellerAuto() {\n      stopBestsellerAuto();\n\n      if (\n        !applyPanel.classList.contains('active') ||\n        !bsListElement.querySelector('.bs-card')\n      ) {\n        return;\n      }\n\n      bsAutoTimer = window.setInterval(function() {\n        if (!document.hidden) {\n          scrollBestsellerCards(1);\n        }\n      }, 2000);\n    }\n\n    function searchBestsellerBook(book) {\n      const searchValue =\n        String(book.isbn || book.title || '').trim();\n\n      if (!searchValue) {\n        showMessage('이 책의 검색 정보가 없습니다.', 'error');\n        return;\n      }\n\n      query.value = searchValue;\n      window.scrollTo({ top: 0, behavior: 'smooth' });\n      showMessage('선택한 책을 정확히 확인하고 있습니다…', 'info');\n\n      if (typeof form.requestSubmit === 'function') {\n        form.requestSubmit();\n      } else {\n        submitButton.click();\n      }\n    }\n\n    document.querySelectorAll('.bs-tab').forEach(function(tab) {\n      tab.addEventListener('click', function() {\n        bsActiveCategory = tab.dataset.bsCategory || 'all';\n        renderBestsellers(bsActiveCategory);\n      });\n    });\n\n    loadBestsellers();\n\n  </script>\n\n  <script>\n    (function () {\n      const noticeWrap = document.getElementById('hrNoticeTicker');\n      const noticeTrack = document.getElementById('hrNoticeTrack');\n\n      function renderWebNotices(items) {\n        const notices = Array.isArray(items)\n          ? items.filter(function (item) {\n              return item && String(item.message || '').trim();\n            })\n          : [];\n\n        if (!noticeWrap || !noticeTrack || notices.length === 0) {\n          if (noticeWrap) noticeWrap.hidden = true;\n          return;\n        }\n\n        const urgent = notices.some(function (item) {\n          return String(item.type || '') === '긴급';\n        });\n\n        const message = notices.map(function (item) {\n          const prefix = String(item.type || '') === '긴급' ? '[긴급] ' : '';\n          return prefix + String(item.message || '').trim();\n        }).join('　◆　');\n\n        noticeTrack.textContent = message;\n        noticeTrack.style.setProperty(\n          '--ticker-duration',\n          Math.max(18, Math.min(70, message.length * 0.24)) + 's'\n        );\n        noticeWrap.classList.toggle('urgent', urgent);\n        noticeWrap.setAttribute('aria-label', '도서관 공지: ' + message);\n        noticeWrap.hidden = false;\n      }\n\n      function loadWebNotices(retried) {\n        if (\n          typeof google === 'undefined' ||\n          !google.script ||\n          !google.script.run\n        ) {\n          if (noticeWrap) noticeWrap.hidden = true;\n          return;\n        }\n\n        google.script.run\n          .withSuccessHandler(renderWebNotices)\n          .withFailureHandler(function () {\n            // 일시적인 오류일 수 있어 한 번 더 시도한 뒤에도 안 되면 조용히 숨깁니다.\n            if (retried) {\n              if (noticeWrap) noticeWrap.hidden = true;\n              return;\n            }\n            window.setTimeout(function () { loadWebNotices(true); }, 4000);\n          })\n          .HR_getActiveWebNotices(window.HR_AUDIENCE || 'staff');\n      }\n\n      if (document.readyState === 'loading') {\n        document.addEventListener(\"DOMContentLoaded\", function () { loadWebNotices(false); });\n      } else {\n        loadWebNotices();\n      }\n\n      window.HR_reloadWebNotices = function () { loadWebNotices(false); };\n      window.setInterval(function () { loadWebNotices(false); }, 300000);\n    })();\n  </script>\n\n\n\n  <script>\n    // 화면을 오래 열어 두는 동안 새 버전이 나오면 알려 줍니다.\n    (function () {\n      var BUILD = '<?= buildId ?>';\n      var banner = document.getElementById('newVersionBanner');\n      if (!BUILD || !banner || typeof google === 'undefined' || !google.script || !google.script.run) return;\n      function check() {\n        google.script.run\n          .withSuccessHandler(function (serverBuild) {\n            if (serverBuild && serverBuild !== BUILD) banner.hidden = false;\n          })\n          .withFailureHandler(function () {})\n          .HR_getBuildId();\n      }\n      window.setTimeout(check, 5000);\n      window.setInterval(check, 300000);\n    })();\n  </script>\n  <script>\n    // 신청 방법 안내: 교직원용/학부모용을 따로 보여 주고, 신청 한도가 있으면 같이 알려 줍니다.\n    (function () {\n      var LIMITS = { staff: Number('<?= limitStaff ?>') || 0, parent: Number('<?= limitParent ?>') || 0 };\n      var hover = document.getElementById('guideHover');\n      var pop = document.getElementById('guidePopover');\n      var title = document.getElementById('guideTitle');\n      var steps = document.getElementById('guideSteps');\n      var note = document.getElementById('guideNote');\n      if (!hover || !pop) return;\n\n      // 위쪽 메뉴 줄의 흐림 효과 때문에 창이 잘리지 않도록 화면 맨 바깥으로 옮깁니다.\n      var backdrop = document.createElement('div');\n      backdrop.id = 'guideBackdrop';\n      document.body.appendChild(backdrop);\n      document.body.appendChild(pop);\n      pop.addEventListener('click', function (event) { event.stopPropagation(); });\n      function close() { if (typeof closeGuidePopover === 'function') closeGuidePopover(); }\n      backdrop.addEventListener('click', close);\n      document.getElementById('guideClose').addEventListener('click', close);\n      document.addEventListener('keydown', function (event) { if (event.key === 'Escape') close(); });\n      new MutationObserver(function () {\n        document.body.classList.toggle('guide-open', hover.classList.contains('open'));\n      }).observe(hover, { attributes: true, attributeFilter: ['class'] });\n\n      var COMMON = [\n        ['2. 신청바구니에 담기', '표지·저자·출판사·연도·정가를 확인하고 「신청바구니에 담기」를 누릅니다. 여러 권을 담을 수 있고, 한 번에 최대 20권까지 신청할 수 있습니다.'],\n        ['3. 신청하기', '화면 아래 신청바구니에서 「선택한 도서 신청하기」를 누릅니다.'],\n        ['4. 결과 확인·취소', '「신청 목록」에서 접수 결과를 확인하십시오. 잘못 신청했다면 취소할 수 있고, 바로 되돌릴 수도 있습니다.']\n      ];\n      var CONTENT = {\n        staff: {\n          title: '선생님, 이렇게 신청하십시오',\n          steps: [\n            ['1. 신청자 이름 쓰기', '「신청자」 칸에 이름을 입력하십시오. (예: 홍길동)'],\n            ['2. 책 검색하기', '책 제목 또는 ISBN을 입력하고 「책 검색하기」를 누릅니다.'],\n            ['3. 신청바구니에 담기', COMMON[0][1]],\n            ['4. 신청하기', COMMON[1][1]],\n            ['5. 결과 확인·취소', COMMON[2][1]]\n          ],\n          note: '희망도서 신청 시 우선 대출 혜택을 받을 수 있습니다.'\n        },\n        parent: {\n          title: '학부모님, 이렇게 신청하십시오',\n          steps: [\n            ['1. 우리 아이 정보 고르기', '학년·반·번호를 눌러서 골라 주십시오. 학생 이름은 선택입니다.'],\n            ['2. 책 검색하기', '읽고 싶은 책의 제목 또는 ISBN을 입력하고 「책 검색하기」를 누릅니다.'],\n            ['3. 신청바구니에 담기', COMMON[0][1]],\n            ['4. 신청하기', COMMON[1][1]],\n            ['5. 결과 확인·취소', COMMON[2][1]]\n          ],\n          note: '입력하신 학년·반·번호(와 학생 이름)는 희망도서 신청 확인과 도서 구입 처리에만 쓰이며, 도서 구입이 끝나면 함께 파기됩니다. 신청 목록 화면에는 보이지 않습니다. 이미 학교에 있는 책이나 이미 신청된 책은 신청되지 않습니다.'\n        }\n      };\n\n      function esc(text) {\n        return String(text).replace(/[&<>\"']/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '\"': '&quot;', \"'\": '&#39;' }[c]; });\n      }\n\n      window.HR_renderGuide = function (audience) {\n        var key = audience === 'parent' ? 'parent' : 'staff';\n        var data = CONTENT[key], limit = LIMITS[key];\n        title.textContent = data.title;\n        steps.innerHTML = data.steps.map(function (s) {\n          return '<li><strong>' + esc(s[0]) + '</strong><span>' + esc(s[1]) + '</span></li>';\n        }).join('');\n        note.textContent = data.note;\n        var old = pop.querySelector('.guide-limit');\n        if (old) old.parentNode.removeChild(old);\n        if (limit > 0) {\n          var p = document.createElement('p');\n          p.className = 'guide-limit';\n          p.textContent = (key === 'parent' ? '한 학생당' : '한 사람당') + ' 최대 ' + limit + '권까지 신청할 수 있습니다.';\n          note.parentNode.insertBefore(p, note.nextSibling);\n        }\n      };\n      window.HR_renderGuide('staff');\n    })();\n  </script>\n  <script>\n    (function () {\n      var INITIAL = '<?= audience ?>';\n      var BASE_URL = '<?= baseUrl ?>';\n      var FROM_GATE = '<?= fromGate ?>' === 'true';\n      var STAFF_CODE_REQUIRED = '<?= staffCodeRequired ?>' === 'true';\n      var gate = document.getElementById('audienceGate');\n      var bar = document.getElementById('audienceBar');\n      var badge = document.getElementById('audienceBadge');\n      var switchButton = document.getElementById('audienceSwitch');\n      var staffField = document.getElementById('staffApplicantField');\n      var parentField = document.getElementById('parentApplicantField');\n      var gradeInput = document.getElementById('pGrade');\n      var classInput = document.getElementById('pClass');\n      var numberInput = document.getElementById('pNumber');\n      var nameInput = document.getElementById('pName');\n      var uploadTab = document.querySelector('[data-panel=\"uploadPanel\"]');\n      var applyTab = document.querySelector('[data-panel=\"applyPanel\"]');\n      var applicantInput = document.getElementById('applicant');\n\n      window.HR_AUDIENCE = '';\n\n      function fillOptions(select, max) {\n        for (var i = 1; i <= max; i++) {\n          var option = document.createElement('option');\n          option.value = String(i);\n          option.textContent = String(i);\n          select.appendChild(option);\n        }\n      }\n      fillOptions(classInput, Number('<?= maxClasses ?>') || 20);    // 반: 학교 설정에서 정한 반 수까지(안 정했으면 20)\n      fillOptions(numberInput, 40);   // 번호: 1~40\n\n      function syncParentApplicant() {\n        var grade = gradeInput.value;\n        var classNo = String(classInput.value || '').trim();\n        var number = String(numberInput.value || '').trim();\n        var name = String(nameInput.value || '').trim();\n        applicantInput.value = (grade && classNo && number)\n          ? grade + '학년 ' + classNo + '반 ' + number + '번' + (name ? ' ' + name : '')\n          : '';\n      }\n\n      [gradeInput, classInput, numberInput, nameInput].forEach(function (el) {\n        el.addEventListener('input', syncParentApplicant);\n        el.addEventListener('change', syncParentApplicant);\n      });\n\n      // ----- 교직원 인증번호 -----\n      var codeGate = document.getElementById('staffCodeGate');\n      var codeForm = document.getElementById('staffCodeForm');\n      var codeInput = document.getElementById('staffCodeInput');\n      var codeMessage = document.getElementById('staffCodeMessage');\n      var codeSubmit = document.getElementById('staffCodeSubmit');\n      var codeBack = document.getElementById('staffCodeBack');\n      window.HR_STAFF_CODE = '';\n      try { window.HR_STAFF_CODE = window.sessionStorage.getItem('HR_STAFF_CODE') || ''; } catch (e) {}\n\n      function showCodeMessage(text) {\n        codeMessage.textContent = text || '';\n        codeMessage.hidden = !text;\n      }\n\n      function askStaffCode() {\n        gate.hidden = true;\n        bar.hidden = true;\n        codeGate.hidden = false;\n        showCodeMessage('');\n        codeInput.value = '';\n        window.setTimeout(function () { try { codeInput.focus(); } catch (e) {} }, 50);\n      }\n\n      // 이미 입력해 둔 번호가 있으면 서버에 한 번 더 물어보고, 맞으면 바로 들어갑니다.\n      function checkSavedCode(onOk, onNeed) {\n        if (!window.HR_STAFF_CODE || typeof google === 'undefined' || !google.script || !google.script.run) {\n          onNeed();\n          return;\n        }\n        google.script.run\n          .withSuccessHandler(function (result) {\n            if (result && result.ok) { onOk(); } else { window.HR_STAFF_CODE = ''; onNeed(); }\n          })\n          .withFailureHandler(function () { onNeed(); })\n          .HR_verifyStaffCode(window.HR_STAFF_CODE);\n      }\n\n      if (codeForm) {\n        codeForm.addEventListener('submit', function (event) {\n          event.preventDefault();\n          var typed = String(codeInput.value || '').trim();\n          if (!typed) { showCodeMessage('인증번호를 입력해 주십시오.'); return; }\n          codeSubmit.disabled = true;\n          showCodeMessage('');\n          google.script.run\n            .withSuccessHandler(function (result) {\n              codeSubmit.disabled = false;\n              if (result && result.ok) {\n                window.HR_STAFF_CODE = typed;\n                try { window.sessionStorage.setItem('HR_STAFF_CODE', typed); } catch (e) {}\n                codeGate.hidden = true;\n                applyAudience('staff', true);\n              } else {\n                codeInput.value = '';\n                showCodeMessage('인증번호가 맞지 않습니다. 다시 입력해 주십시오.');\n                codeInput.focus();\n              }\n            })\n            .withFailureHandler(function (error) {\n              codeSubmit.disabled = false;\n              showCodeMessage(typeof getErrorMessage === 'function' ? getErrorMessage(error) : '확인하지 못했습니다. 잠시 뒤 다시 시도해 주십시오.');\n            })\n            .HR_verifyStaffCode(typed);\n        });\n        codeBack.setAttribute('href', BASE_URL || '#');\n        if (BASE_URL) {\n          codeBack.setAttribute('target', '_top');\n        } else {\n          codeBack.addEventListener('click', function (event) { event.preventDefault(); window.location.reload(); });\n        }\n        codeBack.hidden = !FROM_GATE;\n      }\n\n      // ----- 남은 신청 가능 권수 안내 (학부모) -----\n      var quotaBox = document.getElementById('parentQuota');\n      var LIMIT_PARENT = Number('<?= limitParent ?>') || 0;\n      var quotaTimer = null;\n\n      function renderQuota(info) {\n        if (!quotaBox) return;\n        if (!LIMIT_PARENT) { quotaBox.hidden = true; return; }\n        quotaBox.classList.remove('full');\n        if (!info || !info.known) {\n          quotaBox.innerHTML = '한 학생당 최대 <strong>' + LIMIT_PARENT + '권</strong>까지 신청할 수 있습니다.<br>학년·반·번호를 고르면 남은 권수를 알려 드립니다.';\n        } else if (info.remain <= 0) {\n          quotaBox.classList.add('full');\n          quotaBox.innerHTML = '이미 <strong>' + info.used + '권</strong>을 신청하여 더 신청할 수 없습니다. (최대 ' + info.limit + '권)';\n        } else {\n          quotaBox.innerHTML = '한 학생당 최대 ' + info.limit + '권 중 <strong>' + info.remain + '권</strong> 더 신청할 수 있습니다.' +\n            (info.used ? '<br>이미 ' + info.used + '권을 신청했습니다.' : '');\n        }\n        quotaBox.hidden = false;\n      }\n\n      function refreshQuota() {\n        if (!LIMIT_PARENT || window.HR_AUDIENCE !== 'parent') { if (quotaBox) quotaBox.hidden = true; return; }\n        if (!gradeInput.value || !classInput.value || !numberInput.value) { renderQuota(null); return; }\n        if (typeof google === 'undefined' || !google.script || !google.script.run) { renderQuota(null); return; }\n        google.script.run\n          .withSuccessHandler(renderQuota)\n          .withFailureHandler(function () { renderQuota(null); })\n          .HR_getParentQuota(gradeInput.value, classInput.value, numberInput.value);\n      }\n      window.HR_refreshQuota = refreshQuota;\n      window.HR_setQuota = renderQuota;\n      // 열어 둔 채로 다른 쪽에서 신청이 늘어도 맞게 보이도록 30초마다 다시 확인합니다.\n      window.setInterval(function () { if (!document.hidden) refreshQuota(); }, 30000);\n\n      // 신청자가 바뀌면 신청 목록을 그 사람 기준으로 다시 불러옵니다.\n      var listTimer = null;\n      function reloadListForApplicant() {\n        window.clearTimeout(listTimer);\n        listTimer = window.setTimeout(function () {\n          submittedRequestsLoaded = false;\n          if (typeof loadSubmittedRequests === 'function') { try { loadSubmittedRequests(true); } catch (e) {} }\n        }, 600);\n      }\n      [gradeInput, classInput, numberInput, nameInput, applicantInput].forEach(function (el) {\n        el.addEventListener('change', reloadListForApplicant);\n      });\n      applicantInput.addEventListener('input', reloadListForApplicant);\n\n      [gradeInput, classInput, numberInput].forEach(function (el) {\n        el.addEventListener('change', function () {\n          window.clearTimeout(quotaTimer);\n          quotaTimer = window.setTimeout(refreshQuota, 150);\n        });\n      });\n\n      function applyAudience(audience, codeChecked) {\n        var isParent = audience === 'parent';\n        if (!isParent && STAFF_CODE_REQUIRED && !codeChecked) {\n          checkSavedCode(function () { applyAudience(audience, true); }, askStaffCode);\n          return;\n        }\n        window.HR_AUDIENCE = audience;\n        document.body.classList.toggle('aud-parent', isParent);\n        if (window.HR_renderGuide) window.HR_renderGuide(audience);\n        gate.hidden = true;\n        bar.hidden = false;\n        badge.textContent = isParent ? '👨‍👩‍👧 학부모 신청' : '🧑‍🏫 교직원 신청';\n        staffField.hidden = isParent;\n        parentField.hidden = !isParent;\n        if (uploadTab) uploadTab.hidden = isParent;      // 목록 업로드는 교직원 전용\n        if (isParent) {\n          // 학부모에게는 도서관 담당자·내선번호를 보여 주지 않습니다.\n          Array.prototype.forEach.call(document.querySelectorAll('.school-contact-row'), function (node) { node.remove(); });\n          applicantInput.value = '';\n          syncParentApplicant();\n          if (applyTab) applyTab.click();\n        }\n        if (typeof window.HR_reloadWebNotices === 'function') window.HR_reloadWebNotices();\n        refreshQuota();\n        if (typeof loadSubmittedRequests === 'function') {\n          try { loadSubmittedRequests(true); } catch (e) {}\n        }\n      }\n\n      function linkFor(audience) {\n        return BASE_URL + (BASE_URL.indexOf('?') < 0 ? '?' : '&') + 'type=' + audience + '&g=1';\n      }\n\n      // 주소를 알면 실제로 페이지를 이동하고(새 페이지로 꽉 차게), 모르면 같은 페이지에서 바꿉니다.\n      Array.prototype.forEach.call(gate.querySelectorAll('[data-audience]'), function (button) {\n        var audience = button.getAttribute('data-audience');\n        if (BASE_URL) {\n          button.setAttribute('href', linkFor(audience));\n          button.setAttribute('target', '_top');\n        } else {\n          button.addEventListener('click', function (event) {\n            event.preventDefault();\n            applyAudience(audience);\n          });\n        }\n      });\n\n      if (BASE_URL) {\n        switchButton.setAttribute('href', BASE_URL);\n      } else {\n        switchButton.addEventListener('click', function (event) {\n          event.preventDefault();\n          window.location.reload();\n        });\n      }\n\n      if (INITIAL === 'parent' || INITIAL === 'staff') {\n        switchButton.hidden = !FROM_GATE;   // 전용 링크로 바로 들어온 경우엔 처음 화면으로 못 가게 숨깁니다.\n        applyAudience(INITIAL);\n      } else {\n        gate.hidden = false;\n        if (typeof window.HR_reloadWebNotices === 'function') window.HR_reloadWebNotices();\n      }\n    })();\n  </script>\n</body>\n</html>\n","SeriesCheck":"<!doctype html>\n<html lang=\"ko\">\n<head>\n<base target=\"_top\">\n<meta charset=\"utf-8\">\n<style>\n  *{box-sizing:border-box}\n  body{margin:0;padding:20px 22px 22px;background:#f8f5ed;font-family:Arial,\"Noto Sans KR\",sans-serif;color:#2e3a33}\n  h1{margin:0 0 6px;font-size:22px;color:#183d35}\n  .lead{margin:0 0 14px;font-size:14px;line-height:1.55;color:#52685f}\n  .steps{margin:0 0 14px;padding:0;list-style:none}\n  .steps li{display:flex;gap:12px;align-items:flex-start;margin-bottom:9px;padding:11px 13px;border:1px solid #d9e2dc;border-radius:12px;background:#fff;font-size:14px;line-height:1.45}\n  .num{flex:none;width:26px;height:26px;border-radius:50%;background:#d9a441;color:#fff;font-weight:800;display:flex;align-items:center;justify-content:center;font-size:14px}\n  .steps b{display:block;color:#244d42}\n  .status{margin:0 0 12px;padding:11px 13px;border-radius:12px;background:#e3efe9;color:#1f5c4a;font-size:14px;line-height:1.5}\n  .status.warn{background:#fdecea;color:#8a2a1f}\n  .status.wait{background:#fdf1d6;color:#6b4d0a}\n  .bar{height:12px;border-radius:999px;background:#dfe7e2;overflow:hidden;margin-top:8px}\n  .bar i{display:block;height:100%;width:0;background:#2e5b4f;border-radius:999px;transition:width .4s}\n  button{width:100%;border:0;border-radius:14px;font-family:inherit;cursor:pointer}\n  .go{padding:16px;background:#2e5b4f;color:#fff;font-size:18px;font-weight:800}\n  .go:disabled{opacity:.55;cursor:default}\n  .again{margin-top:9px;padding:11px;background:#fff;border:1px solid #c9d3cc;color:#4b5a52;font-size:13.5px;font-weight:700}\n  .msg{margin-top:12px;padding:10px 12px;border-radius:10px;font-size:14px;line-height:1.5}\n  .msg.ok{background:#e3efe9;color:#1f5c4a}.msg.err{background:#fdecea;color:#8a2a1f}\n  .note{margin-top:12px;font-size:12.5px;color:#6a786f;line-height:1.5}\n  [hidden]{display:none!important}\n</style>\n</head>\n<body>\n  <h1>🔎 시리즈 누락 점검</h1>\n  <p class=\"lead\">우리 학교 소장도서에서 시리즈를 찾아, <b>빠진 권</b>과 <b>새로 나온 후속권</b>을 알려 드립니다.</p>\n\n  <ol class=\"steps\">\n    <li><span class=\"num\">1</span><div><b>아래 「점검 시작」을 누릅니다</b>소장도서의 시리즈를 찾고 YES24에서 빠진 권을 확인합니다.</div></li>\n    <li><span class=\"num\">2</span><div><b>잠시 기다립니다</b>오래 걸리면 중간에 멈췄다가 약 1분 뒤 저절로 이어서 점검합니다. 창을 닫아도 됩니다.</div></li>\n    <li><span class=\"num\">3</span><div><b>「시리즈 누락 점검」 표에서 고릅니다</b>사고 싶은 책을 체크하고 「구입 검토함으로 보내기」를 누르세요.</div></li>\n  </ol>\n\n  <div id=\"status\" class=\"status\">상태를 확인하는 중…</div>\n  <div id=\"progress\" hidden><div class=\"bar\"><i id=\"fill\"></i></div></div>\n\n  <button class=\"go\" id=\"go\" type=\"button\" disabled>▶ 점검 시작</button>\n  <button class=\"again\" id=\"again\" type=\"button\" hidden>처음부터 다시 점검하기</button>\n  <div id=\"msg\" class=\"msg\" hidden></div>\n  <p class=\"note\">YES24 API 키가 필요합니다. 없으면 「⚙️ 사본 설정 → 🔑 API 키 설정」에서 먼저 입력해 주세요.</p>\n  <p class=\"note\">💡 출판사는 적지 않아도 됩니다. 소장도서의 ISBN 으로 YES24 에서 자동으로 확인해 같은 이름의 다른 출판사 책이 걸리는 일을 줄입니다.</p>\n\n<script>\n  var $ = function (id) { return document.getElementById(id); };\n  var busy = false;\n  var AUTO = '<?= autostart ?>' === 'true';   // 시트의 점검 막대를 눌러서 열렸으면 바로 시작\n  var autoDone = false;\n\n  function showMsg(text, ok) { var m = $('msg'); m.textContent = text; m.className = 'msg ' + (ok ? 'ok' : 'err'); m.hidden = false; }\n\n  function render(s) {\n    var status = $('status'), go = $('go'), again = $('again');\n    if (AUTO && !autoDone && s.hasKey && !busy) { autoDone = true; start(false); return; }\n    $('progress').hidden = true;\n    again.hidden = true;\n    if (!s.hasKey) {\n      status.className = 'status warn';\n      status.textContent = '⚠ YES24 API 키가 아직 없습니다. 「⚙️ 사본 설정 → 🔑 API 키 설정」에서 입력한 뒤 다시 열어 주세요.';\n      go.disabled = true; go.textContent = '▶ 점검 시작';\n      return;\n    }\n    go.disabled = busy;\n    if (s.resume) {\n      var pct = Math.min(100, Math.round(s.resume.offset / s.resume.total * 100));\n      status.className = 'status wait';\n      status.textContent = '이전 점검이 중간에서 멈췄습니다. (' + s.resume.offset + ' / ' + s.resume.total + '건, ' + pct + '%) 이어서 점검할 수 있습니다.';\n      if (s.quotaNext && s.quotaNext.at > Date.now()) status.textContent += ' · YES24 하루 한도 때문에 ' + s.quotaNext.text + '에 저절로 이어서 점검하도록 예약되어 있습니다.';\n      $('progress').hidden = false; $('fill').style.width = pct + '%';\n      go.textContent = '▶ 이어서 점검';\n      again.hidden = false;\n    } else {\n      status.className = 'status';\n      status.textContent = '소장도서 ' + s.holdings + '권 · 지금 표에 있는 점검 결과 ' + s.candidates + '권' + (s.candidates ? '' : ' · 아직 점검한 적이 없거나 결과가 없습니다.');\n      go.textContent = s.candidates ? '▶ 다시 점검' : '▶ 점검 시작';\n    }\n  }\n\n  function load() {\n    google.script.run\n      .withSuccessHandler(render)\n      .withFailureHandler(function () { $('status').className = 'status warn'; $('status').textContent = '상태를 불러오지 못했습니다. 잠시 뒤 다시 열어 주세요.'; })\n      .HR_seriesCheckState();\n  }\n\n  function start(fresh) {\n    if (busy) return;\n    busy = true; $('go').disabled = true; $('again').hidden = true; $('msg').hidden = true;\n    $('go').textContent = '점검하는 중… (몇 분 걸릴 수 있습니다)';\n    $('status').className = 'status wait';\n    $('status').textContent = '🔎 점검을 시작했습니다. 이 창은 지금 닫아도 됩니다. 점검은 계속 진행되고, 끝나면 시트 오른쪽 아래에 알림이 뜹니다.';\n    google.script.run\n      .withSuccessHandler(function (s) {\n        busy = false; render(s);\n        if (s.resume) showMsg('한 번에 다 끝내지 못해서 저장했습니다. 약 1분 뒤 저절로 이어서 점검합니다. (바로 이어서 하려면 다시 눌러도 됩니다)', true);\n        else showMsg('✅ 점검을 마쳤습니다. 「시리즈 누락 점검」 표에서 결과를 확인하세요.', true);\n      })\n      .withFailureHandler(function (e) {\n        busy = false; load();\n        showMsg('점검하지 못했습니다. (' + String((e && e.message) || e || '알 수 없는 오류').slice(0, 200) + ')', false);\n      })\n      .HR_seriesCheckStart(!!fresh);\n  }\n\n  $('go').onclick = function () { start(false); };\n  $('again').onclick = function () { start(true); };\n  load();\n</script>\n</body>\n</html>\n"},
    menus: (typeof APP_MENUS !== 'undefined') ? APP_MENUS : [],
    rpc: ["HR_showSeriesCheckDialog","HR_seriesCheckState","HR_seriesCheckStart","HR_showApplyLinksDialog","HR_showRequestLimitDialog","HR_getParentQuota","getConfiguredNlkKey","startNewRequestRound","styleRequestListSheet","buildBookstoreExportSheet","installEditTrigger","onEditTriggerAction","processRow","shouldSkipStatus","smartSearch","fetchYes24Api","selectBestYes24Item","fetchNlkCallNumber","parseNlkClassNumber","findFirstXmlValue","normalizeClassNumber","findDuplicateRequest","findLibraryDuplicate","fillFinalData","checkExistingBooks","reprocessSelectedRows","getHeaders","getColumnMap","getLibraryColumnMap","setStatus","parseTitleAndPublisher","extractBestIsbn","normalizeIsbn","normalizeText","normalizeHeader","stripHtml","extractYear","toNumber","isTitleHeader","isIsbnHeader","isStatusHeader","isCoverHeader","isPublisherHeader","isAuthorHeader","isPriceHeader","isYearHeader","isCategoryHeader","isCallNumberHeader","createCacheKey","getCachedJson","putCachedJson","FIX_ACTION_LAYOUT_20260820","resetHopeBooksForNewSemester","HR_importFromOldCopy","HR_showInquiryDialog","HR_showMyInquiriesDialog","HR_sendInquiry","HR_getMyInquiries","unifiedConfirmPurchaseRound","unifiedShowPurchaseLog","unifiedHidePurchaseLog","HR_getBuildId","HR_renameLegacyTargetSheet","HR_showRequestSidebar","HR_openSidebarOnSpreadsheetOpen","HR_installSidebarOpenTrigger","HR_submitBookRequest","HR_submitBookRequests","HR_searchRequestBooks","HR_formatGuideSheet","HR_addCopySetupMenu","HR_showYes24KeyDialog","HR_showSchoolSetupDialog","HR_showPublicWebAppUrlDialog","HR_savePublicWebAppUrl","HR_getCopySetupState","HR_saveYes24ApiKey","HR_runSchoolInitialSetup","HR_deleteCopySetupSheet","HR_createCopySetupSheet","HR_refreshCopySetupBeginnerGuide","HR_showSystemCheck","HR_getSubmittedRequests","HR_restoreLegacyCancelledRequest","HR_cancelSubmittedRequest","HR_undoCancelledRequest","HR_getBestsellers","HR_fillMissingPriceAndYear","HR_importBookListFile","HR_showHopeRequestDialog","HR_seriesStopCheck","HR_verifyStaffCode","HR_showStaffCodeDialog","HR_arrangeTabs","setupSeriesFinder","installSeriesWorkButton","setSeriesYes24Key","refreshMissingSeriesBooks","resetSeriesFinderProgress","refreshSeriesSaleStatuses","refreshSeriesApiGauge","setupSeriesWorkSheet","exportSelectedSeriesBooks","setupUnifiedLibraryWorkflow","openUnifiedUploadSheet","refreshUnifiedPurchaseReview","buildUnifiedBookstoreExportSheet","unifiedSetBudget","unifiedAuditDataSources","importLegacySeriesSheetsOnce","sendHopeRequestsToReview","sendSelectedSeriesToReview","sendSelectedUploadItemsToReview","applyUnifiedLibraryDesign","removeLegacyUnifiedSheetsAfterConfirmation","HR_dismissUpdateNotice","HR_showUpdateNotice","HR_upgradeNoticeDatePickersAndGuide","HR_getActiveWebNotices","HR_setupWebNoticeSheet","verifyYes24Connection"],
    has: function (n) { return typeof __fns[n] === 'function'; },
    call: function (n, a) {
      if (typeof __fns[n] !== 'function') throw new Error('함수를 찾을 수 없습니다: ' + n);
      return __fns[n].apply(null, a || []);
    }
  };
};
