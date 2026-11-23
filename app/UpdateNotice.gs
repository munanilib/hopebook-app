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
    '.withFailureHandler(function(){google.script.host.close()})' +
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
  if (PropertiesService.getUserProperties().getProperty(HR_NOTICE_SEEN_PROPERTY) === notice.id) return;
  HR_showNoticeDialog_(notice);
}

// 안내 창의 「이 소식 다시 보지 않기」 버튼
function HR_dismissUpdateNotice(id) {
  var text = String(id || '');
  if (!/^n[0-9a-z]{1,12}$/.test(text)) return;
  PropertiesService.getUserProperties().setProperty(HR_NOTICE_SEEN_PROPERTY, text);
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
