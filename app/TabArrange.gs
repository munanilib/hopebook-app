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
