/**
 * 📘 전체 사용방법 (새 버전)
 *
 * 신청 받기부터 서점 전달까지, 새로 생긴 메뉴(교직원/학부모 신청, 링크·QR, 신청 한도, 새 신청 시작,
 * 업데이트 소식, 데이터 가져오기)까지 한 장에 정리합니다.
 * 안내 시트는 원래 교체해도 되는 시트라서, 새 버전이 나오면 처음 열 때 한 번 자동으로 새 내용으로 바꿉니다.
 */
var HR_GUIDE_VERSION = '2026-11-23-1';
var HR_GUIDE_VERSION_PROPERTY = 'HR_GUIDE_VERSION';

var HR_MANUAL_TITLE = '📘 학교도서관 희망도서·구입 업무 전체 사용방법';
var HR_MANUAL_SUBTITLE =
  '신청 받기부터 서점에 주문서를 보낼 때까지 한 장으로 정리했어요. 처음이라면 맨 위부터 차례로 읽어 보세요.';
var HR_MANUAL_FOOTER =
  '업무 흐름: 신청 받기 → 구입 검토함 → 소장 중복 확인 → 서점 전달용  ·  막히면 맨 아래 「자주 묻는 질문」을 먼저 보세요.';

var HR_MANUAL_SECTIONS = [
  {
    title: '🧭 한눈에 보기 — 업무는 4단계예요',
    items: [
      ['① 신청 받기',
        '교직원과 학부모가 웹 페이지(링크·QR)로 희망도서를 신청해요. 신청은 「📚 희망도서 목록」에 자동으로 쌓이고, 표지·저자·출판사·정가·ISBN 같은 책 정보는 저절로 채워져요.'],
      ['② 후보 모으기',
        '「📚 희망도서 목록」의 신청을 「🛒 구입 검토함」으로 보내요. 시리즈 누락·후속권 점검 결과(🔎)와 엑셀 업로드(📥) 후보도 같은 검토함에 모여요.'],
      ['③ 구입 결정',
        '「🛒 구입 검토함」에서 살 책의 「선택」을 체크하고 수량·메모를 확인해요. 우리 학교에 이미 있는 책은 자동으로 걸러져요.'],
      ['④ 서점에 전달',
        '검토함의 금색 버튼을 누르면 「📦 서점 전달용」이 만들어져요. 개인정보 없이 같은 책은 수량으로 합치고 금액도 계산해 줘요. 파일 → 다운로드 → Microsoft Excel(.xlsx)로 저장해 서점에 보내세요.']
    ]
  },
  {
    title: '🔗 신청 링크 나눠 주기 — 교직원용 / 학부모용',
    items: [
      ['링크 받기',
        '「⚙️ 사본 설정 → 🔗 신청 링크 받기」를 누르세요. 교직원용·학부모용 링크가 나오고, 복사하거나 QR 그림으로 저장해 가정통신문에 넣을 수 있어요. 문자·알림장에 붙일 안내 문구도 한 번에 복사돼요.'],
      ['처음 한 번: 웹 앱 배포',
        '링크가 만들어지려면 Apps Script에서 웹 앱을 한 번 배포해야 해요. (확장 프로그램 → Apps Script → 배포 → 새 배포 → 웹 앱, 실행 사용자는 「나」, 액세스는 「모든 사용자」) 「링크 받기」 창에 순서가 크게 나와요.'],
      ['교직원 / 학부모',
        '두 링크로 들어가면 각자 자기 신청 목록만 보여요. 서로의 신청은 보이지도, 취소할 수도 없어요. 한 시트의 「신청구분」 칸에서 사서 선생님이 함께 관리해요.'],
      ['학부모 신청 정보',
        '학년·반·번호는 꼭 골라야 하고, 학생 이름은 안 써도 돼요. 입력한 정보는 신청 확인에만 쓰이고 신청 목록 화면에는 보이지 않아요.'],
      ['신청 한도',
        '「⚙️ 사본 설정 → 📏 신청 한도 설정」에서 1인당 신청할 수 있는 권수를 정해요. 학부모는 한 학생(학년·반·번호)당 5권이 기본이고, 교직원은 제한이 없어요. 0을 넣으면 제한 없음이에요. 한도를 넘으면 안내 창이 떠요.'],
      ['링크가 안 열려요',
        '「공개 신청 링크 등록」에 넣은 주소가 틀리면 「신청 링크 받기」 창에 원인과 해결 방법이 나와요. 「배포 → 배포 관리」에서 웹 앱 주소를 처음부터 끝까지 다시 복사해 붙여 넣으세요.']
    ]
  },
  {
    title: '📋 메뉴 지도 — 위쪽 메뉴를 언제 쓰나요?',
    items: [
      ['📚 도서 관리',
        '「희망도서 목록」을 다루는 메뉴예요.\n' +
        '• 🔍 기존 데이터 일괄 채우기 — 비어 있는 책 정보를 한꺼번에 채워요\n' +
        '• 🔄 선택한 행 다시 조회 — 고른 줄의 책 정보를 다시 가져와요\n' +
        '• 🛒 구입 검토함에 접수 도서 추가 · 📦 서점 전달용 시트 새로고침\n' +
        '• 🎨 희망도서 목록 디자인 다시 적용 · ⚙️ 실시간 자동화 권한 승인\n' +
        '• 🔐 관리자 · 새 희망도서 신청 시작 — 아래 「새 학기·새 신청 회차」 참고'],
      ['⭐ 희망도서 신청',
        '• 🔗 신청 페이지 열기 — 등록된 신청 링크로 바로 이동해요\n• 🧹 새 학기 초기화 — 아래 「새 학기·새 신청 회차」 참고'],
      ['⚙️ 사본 설정',
        '처음 설정과 관리 메뉴예요.\n' +
        '• ✅ 설정 한 번에 점검 · 🔑 API 키 설정 (YES24 필수) · 🏫 학교 최초 설정 · 🔗 공개 신청 링크 등록\n' +
        '• 🔗 신청 링크 받기 — 교직원/학부모 링크와 QR 그림\n' +
        '• 📏 신청 한도 설정 · 📢 업데이트 소식 보기\n' +
        '• 📥 이전 사본에서 데이터 가져오기 — 옛 사본의 신청 데이터를 이 사본으로 옮겨요\n' +
        '• 🧰 설정 탭 만들기·새로고침 · 🗑️ 설정 완료 후 탭 삭제'],
      ['🧰 통합 도서업무',
        '구입 업무 전체를 이끄는 메뉴예요.\n' +
        '• 🏠 전체 화면 정리·새로고침 · 🎨 전체 디자인 다시 적용 · 🧭 시트 탭 정리\n' +
        '• 📚 희망도서를 구입 검토함으로 보내기 · 🛒 체크한 시리즈를 구입 검토함으로 보내기\n' +
        '• 🔎 시리즈 누락·후속권 이어서 점검 · 📥 업로드 목록을 열기\n' +
        '• 📦 중복 확인 후 서점 전달용으로 보내기\n' +
        '• 🔎 시리즈 점검 도구(하위 메뉴) — 현재 결과 절판 여부 확인 · 결과 시트 만들기·복구 · 저장된 점검 위치 초기화 · YES24 API 키 설정']
    ]
  },
  {
    title: '🗓 새 학기·새 신청 회차',
    items: [
      ['🔐 새 희망도서 신청 시작',
        '지금 신청 목록을 숨김 시트(🗂 희망도서 날짜)로 보관한 뒤 목록을 비우고 새 신청을 받아요. 구입 검토함·서점 전달용·소장도서·설정은 그대로예요. 학생별 신청 한도도 새 목록 기준으로 다시 0권부터 세요.'],
      ['🧹 새 학기 초기화',
        '이번 학기 접수·검토 데이터를 비우고 서점 전달용 시트를 숨겨요. 소장도서와 백업 자료는 남겨요. 한 학기가 끝났을 때만 쓰세요.'],
      ['누르기 전에',
        '둘 다 목록을 비우는 메뉴예요. 누르기 전에 서점 전달용을 엑셀(.xlsx)로 내려받아 두세요.']
    ]
  },
  {
    title: '📢 소식과 공지',
    items: [
      ['업데이트 소식',
        '새 기능이 생기면 시트를 열 때 안내 창이 한 번 떠요. 다시 보려면 「⚙️ 사본 설정 → 📢 업데이트 소식 보기」를 누르세요.'],
      ['웹 공지 (신청 화면 위쪽)',
        '「📢 웹 공지」 탭에서 「표시」를 체크하면 신청 페이지 위쪽에 공지가 떠요. 「대상」 칸에서 전체·교직원·학부모를 고를 수 있어요. 시작일·종료일은 비워도 돼요.']
    ]
  },
  {
    title: '🛠 자주 묻는 질문',
    items: [
      ['책이 신청이 안 돼요',
        '이미 학교에 있는 책(소장도서)이거나 이미 신청된 책이에요. 「🏫 우리학교 소장도서」에 도서명과 ISBN13을 최신으로 넣어 두면 중복을 정확히 알려 줘요.'],
      ['한도를 넘었다고 나와요',
        '1인당(학부모는 한 학생당) 신청 한도 때문이에요. 신청을 취소하면 그 권수는 다시 신청할 수 있고, 한도는 「📏 신청 한도 설정」에서 바꿔요.'],
      ['웹 화면이 옛날 모양이에요',
        '열어 둔 웹 페이지 탭을 새로고침(F5)하세요. 화면 맨 아래의 「화면 버전」으로 최신인지 확인할 수 있고, 새 버전이 나오면 위쪽에 안내가 떠요.'],
      ['메뉴가 안 보여요',
        '처음 열면 「⚙️ 시작하기」만 보여요. 「프로그램 사용 시작」을 눌러 권한을 한 번 허용하고 새로고침하세요.'],
      ['탭이 너무 많아서 찾기 어려워요',
        '「🧰 통합 도서업무 → 🧭 시트 탭 정리」를 누르면 업무 순서로 놓고 색으로 묶어요. (초록: 업무 · 금색: 자료 · 회색: 안내·설정)']
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
    props.setProperty(HR_GUIDE_VERSION_PROPERTY, HR_GUIDE_VERSION);
    try {
      if (original && !original.isSheetHidden()) spreadsheet.setActiveSheet(original);
    } catch (error) {}
  } finally {
    if (locked) lock.releaseLock();
  }
}
