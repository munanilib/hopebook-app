/**
 * 📢 업데이트 소식
 *
 * 개발자가 GitHub 의 notice.json 에 소식을 쓰면, 사본이 열릴 때 한 번 안내 창으로 보여 줍니다.
 * - id 가 비어 있으면 아무것도 보여 주지 않습니다.
 * - 한 번 본 소식(같은 id)은 다시 안 뜹니다. 새 소식은 id 를 바꾸면 됩니다.
 * - start·end(YYYY-MM-DD)는 비워도 됩니다. 기간 밖이면 보여 주지 않습니다.
 * - 인터넷에 연결되지 않으면 조용히 건너뜁니다.
 */
var HR_NOTICE_URL = 'https://raw.githubusercontent.com/munanilib/hopebook-app/refs/heads/stable/notice.json';
var HR_NOTICE_SEEN_PROPERTY = 'HR_NOTICE_SEEN_ID';

function HR_fetchNotice_() {
  var response = UrlFetchApp.fetch(HR_NOTICE_URL + '?t=' + Date.now(), { muteHttpExceptions: true });
  if (response.getResponseCode() !== 200) return null;
  var data = JSON.parse(response.getContentText('UTF-8'));
  var id = String(data && data.id || '').trim();
  var body = String(data && data.body || '').trim();
  if (!id || !body) return null;

  var today = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd');
  var start = String(data.start || '').trim();
  var end = String(data.end || '').trim();
  if (start && today < start) return null;
  if (end && today > end) return null;

  return { id: id, title: String(data.title || '업데이트 소식').trim(), body: body };
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
    '</style></head><body><div class="badge">📢 업데이트 소식</div>' +
    '<h1 class="title">' + HR_escapeNoticeHtml_(notice.title) + '</h1>' +
    '<div class="body">' + HR_escapeNoticeHtml_(notice.body).replace(/\n/g, '<br>') + '</div>' +
    '<button class="ok" onclick="google.script.host.close()">확인했어요</button>' +
    '</body></html>'
  ).setWidth(460).setHeight(430);
  SpreadsheetApp.getUi().showModelessDialog(html, '📢 업데이트 소식');
}

// 시트를 열 때: 새 소식이면 한 번만 보여 줍니다.
function HR_showUpdateNoticeOnOpen_() {
  var notice = HR_fetchNotice_();
  if (!notice) return;
  var props = PropertiesService.getUserProperties();
  if (props.getProperty(HR_NOTICE_SEEN_PROPERTY) === notice.id) return;
  HR_showNoticeDialog_(notice);
  props.setProperty(HR_NOTICE_SEEN_PROPERTY, notice.id);
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
