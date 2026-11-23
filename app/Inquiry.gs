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
