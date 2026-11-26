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
