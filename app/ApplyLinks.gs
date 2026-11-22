/**
 * 🔗 신청 링크 받기
 *
 * 교직원용·학부모용 신청 링크를 한 창에서 보여 주고, 복사하거나 QR 코드를 그림(PNG)으로 저장하게 합니다.
 * QR 은 창 안에서 직접 만들기 때문에 학교 주소가 외부 서비스로 나가지 않습니다.
 */
function HR_showApplyLinksDialog() {
  var props = PropertiesService.getScriptProperties();
  var registered = HR_getRegisteredWebAppUrl_();
  var detected = HR_getDetectedWebAppUrl_();

  // 등록한 주소가 없을 때만, 찾은 주소를 「공개 링크」로 저장해 둡니다. (등록한 주소는 절대 덮어쓰지 않습니다.)
  if (!registered && detected) {
    props.setProperty(HR_PUBLIC_WEB_APP_URL_PROPERTY, detected);
  }

  var baseUrl = registered || detected;
  var problem = '';
  var switched = false;

  // 등록한 주소가 실제로 열리는지 확인합니다.
  if (baseUrl) {
    var status = HR_probeWebAppUrl_(baseUrl);
    if (status === 'notfound' || status === 'restricted') {
      // 등록한 주소는 안 열리는데 구글이 알려 주는 주소가 열리면, 그 주소로 바꿔서 보여 줍니다.
      if (detected && detected !== baseUrl && HR_probeWebAppUrl_(detected) === 'ok') {
        baseUrl = detected;
        switched = true;
        props.setProperty(HR_PUBLIC_WEB_APP_URL_PROPERTY, detected);
      } else {
        problem = HR_probeProblemMessage_(status);
      }
    }
  }

  var data = {
    baseUrl: problem ? '' : baseUrl,
    detected: !!baseUrl,
    mismatch: !!(registered && detected && registered !== detected && !switched),
    switched: switched,
    problem: problem,
    badUrl: problem ? baseUrl : '',
    schoolName: String(props.getProperty(HR_SCHOOL_NAME_PROPERTY) || '').trim()
  };

  var template = HtmlService.createTemplateFromFile('ApplyLinks');
  template.dataJson = JSON.stringify(data).replace(/</g, '\\u003c');
  SpreadsheetApp.getUi().showModalDialog(
    template.evaluate().setWidth(780).setHeight(860),
    '🔗 신청 링크 받기'
  );
}
