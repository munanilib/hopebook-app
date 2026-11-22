/* 자동 생성 파일 — 직접 고치지 마세요. (node tools/build.js)  v2 */
var APP_FACTORY = function (__host) {
  var HtmlService = __host.html;

/* ===== Main.gs ===== */
/**
 * 🧪 자동 업데이트 시험판 앱
 * 이 파일을 고쳐서 GitHub(stable)에 올리면 모든 사본에 반영됩니다.
 */
const TEST_MESSAGE = '자동 업데이트 성공! 두 번째 버전입니다';

var APP_MENUS = [
  { title: '🧪 시험판', items: [
    ['💬 버전 확인', 'showVersion'],
    ['📋 사이드바 열기', 'showSidebar'],
    '-',
    ['✍️ 시트에 기록 남기기', 'writeTestRow']
  ]}
];

function onOpen(e) {
  SpreadsheetApp.getActive().toast(TEST_MESSAGE, '🧪 시험판 v' + appVersion_(), 5);
}

function appVersion_() {
  return '1';
}

function showVersion() {
  SpreadsheetApp.getUi().alert('🧪 시험판 v' + appVersion_() + '\n\n' + TEST_MESSAGE);
}

function showSidebar() {
  const html = HtmlService.createTemplateFromFile('Sidebar');
  html.message = TEST_MESSAGE;
  SpreadsheetApp.getUi().showSidebar(html.evaluate().setTitle('🧪 시험판'));
}

function writeTestRow() {
  const sh = SpreadsheetApp.getActiveSheet();
  sh.appendRow([new Date(), 'v' + appVersion_(), TEST_MESSAGE]);
}

// 화면(사이드바·웹 페이지)에서 부르는 함수
function getServerInfo(who) {
  return {
    version: appVersion_(),
    message: TEST_MESSAGE,
    who: who || '',
    time: Utilities.formatDate(new Date(), 'Asia/Seoul', 'HH:mm:ss')
  };
}

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Web')
    .setTitle('🧪 시험판 웹 페이지')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}


  var __fns = {
    "onOpen": onOpen,
    "appVersion_": appVersion_,
    "showVersion": showVersion,
    "showSidebar": showSidebar,
    "writeTestRow": writeTestRow,
    "getServerInfo": getServerInfo,
    "doGet": doGet
  };
  return {
    version: "2",
    files: {"Sidebar":"<!DOCTYPE html>\n<html><head><base target=\"_top\">\n<style>body{font-family:sans-serif;padding:12px}button{padding:8px 12px}#out{margin-top:12px;color:#1a7f37}</style>\n</head><body>\n  <h3>🧪 사이드바</h3>\n  <p>템플릿 값: <b><?= message ?></b></p>\n  <button onclick=\"ask()\">서버에 물어보기</button>\n  <div id=\"out\"></div>\n<script>\nfunction ask(){\n  document.getElementById('out').textContent = '물어보는 중…';\n  google.script.run\n    .withSuccessHandler(function(r){ document.getElementById('out').textContent = 'v' + r.version + ' · ' + r.message + ' · ' + r.time; })\n    .withFailureHandler(function(e){ document.getElementById('out').textContent = '오류: ' + e.message; })\n    .getServerInfo('sidebar');\n}\n</script>\n</body></html>\n","Web":"<!DOCTYPE html>\n<html><head><base target=\"_top\">\n<style>body{font-family:sans-serif;padding:24px;max-width:480px;margin:auto}#out{margin-top:16px;font-size:18px;color:#1a7f37}</style>\n</head><body>\n  <h2>🧪 시험판 웹 페이지</h2>\n  <p>로그인 없이 열리는 페이지입니다.</p>\n  <div id=\"out\">불러오는 중…</div>\n<script>\ngoogle.script.run\n  .withSuccessHandler(function(r){ document.getElementById('out').textContent = 'v' + r.version + ' · ' + r.message + ' (' + r.time + ')'; })\n  .withFailureHandler(function(e){ document.getElementById('out').textContent = '오류: ' + e.message; })\n  .getServerInfo('web');\n</script>\n</body></html>\n"},
    menus: (typeof APP_MENUS !== 'undefined') ? APP_MENUS : [],
    rpc: ["showVersion","showSidebar","writeTestRow","getServerInfo"],
    has: function (n) { return typeof __fns[n] === 'function'; },
    call: function (n, a) {
      if (typeof __fns[n] !== 'function') throw new Error('함수를 찾을 수 없습니다: ' + n);
      return __fns[n].apply(null, a || []);
    }
  };
};
