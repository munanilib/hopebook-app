/**
 * 🔗 신청 링크 받기
 *
 * 교직원용·학부모용 신청 링크를 한 창에서 보여 주고, 복사하거나 QR 코드를 그림(PNG)으로 저장하게 합니다.
 * QR 은 창 안에서 직접 만들기 때문에 학교 주소가 외부 서비스로 나가지 않습니다.
 */
function HR_showApplyLinksDialog() {
  var props = PropertiesService.getScriptProperties();
  var detected = HR_getWebAppBaseUrl_();
  var saved = String(props.getProperty(HR_PUBLIC_WEB_APP_URL_PROPERTY) || '').trim();

  // 배포된 웹 주소를 자동으로 찾았으면 「공개 링크」로도 저장해 둡니다.
  if (detected && detected !== saved) {
    props.setProperty(HR_PUBLIC_WEB_APP_URL_PROPERTY, detected);
    saved = detected;
  }

  var baseUrl = detected || (/^https:\/\/script\.google\.com\/.+\/exec$/.test(saved) ? saved : '');
  var data = {
    baseUrl: baseUrl,
    detected: !!detected,
    schoolName: String(props.getProperty(HR_SCHOOL_NAME_PROPERTY) || '').trim()
  };

  var template = HtmlService.createTemplateFromFile('ApplyLinks');
  template.dataJson = JSON.stringify(data).replace(/</g, '\\u003c');
  SpreadsheetApp.getUi().showModalDialog(
    template.evaluate().setWidth(780).setHeight(860),
    '🔗 신청 링크 받기'
  );
}
