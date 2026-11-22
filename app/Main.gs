/**
 * 🧪 자동 업데이트 시험판 앱
 * 이 파일을 고쳐서 GitHub(stable)에 올리면 모든 사본에 반영됩니다.
 */
const TEST_MESSAGE = '🎉🎉 자동 업데이트 성공! 세 번째 버전입니다 🎉🎉';

var APP_MENUS = [
  { title: '🚀 시험판 v3 (자동 업데이트됨!)', items: [
    ['💬 버전 확인', 'showVersion'],
    ['📋 사이드바 열기', 'showSidebar'],
    '-',
    ['✍️ 시트에 기록 남기기', 'writeTestRow'],
    '-',
    ['🎁 새로 생긴 기능', 'newFeature']
  ]}
];

function onOpen(e) {
  SpreadsheetApp.getActive().toast(TEST_MESSAGE, '🧪 시험판 v' + appVersion_(), 5);
}

function appVersion_() {
  return '3';
}

function showVersion() {
  SpreadsheetApp.getUi().alert('🧪 시험판 v' + appVersion_() + '\n\n' + TEST_MESSAGE);
}

function showSidebar() {
  const html = HtmlService.createTemplateFromFile('Sidebar');
  html.message = TEST_MESSAGE;
  SpreadsheetApp.getUi().showSidebar(html.evaluate().setTitle('🧪 시험판'));
}

function newFeature() {
  SpreadsheetApp.getUi().alert('🎁 새 기능', '이 메뉴는 v3에서 새로 생겼어요.\n시트의 스크립트는 하나도 바꾸지 않았습니다!', SpreadsheetApp.getUi().ButtonSet.OK);
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
