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
