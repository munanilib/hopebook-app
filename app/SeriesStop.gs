/**
 * ⏹ 시리즈 점검 멈추기 — 돌고 있는 점검을 멈춥니다.
 *   · UnifiedLibrary.js: 메뉴 항목 / seriesSearchYes24_ 반복문 / refreshMissingSeriesBooks 점검 직후에서 이 파일의 함수를 부릅니다.
 *
 * 돌고 있는 점검은 한 묶음이 끝나는 대로(몇 초 안에) 멈추고, 1분 뒤 자동 이어가기도 지웁니다.
 * 멈춘 점검의 이번 회 결과는 시트에 쓰지 않고, 저장된 점검 위치도 지웁니다. (이전에 쓰인 결과 시트는 그대로 남습니다)
 */
var SERIES_STOP_PROPERTY = 'SERIES_STOP_REQUESTED_AT';

function HR_seriesStopCheck() {
  var properties = PropertiesService.getScriptProperties();
  properties.setProperty(SERIES_STOP_PROPERTY, String(Date.now()));
  try { seriesCancelAutoContinue_(); } catch (error) { console.warn('자동 이어가기 지우기 실패: ' + error); }
  try { seriesClearResumeState_(); } catch (error) { console.warn('점검 위치 지우기 실패: ' + error); }
  try {
    seriesSpreadsheet_().toast(
      '점검을 멈추도록 요청했습니다. 돌고 있는 점검은 몇 초 안에 멈추고, 1분 뒤 자동 이어가기는 취소했습니다. 다음 점검은 처음부터 시작합니다.',
      '⏹ 시리즈 점검 멈추기',
      10
    );
  } catch (error) {}
}

// 이 점검(runStartedAt 에 시작)이 시작된 뒤에 멈추기를 눌렀는지. 멈추기 전에 이미 끝난 요청은 영향 없습니다.
function seriesStopRequested_(runStartedAt) {
  var requestedAt = Number(PropertiesService.getScriptProperties().getProperty(SERIES_STOP_PROPERTY) || 0);
  return requestedAt > Number(runStartedAt || 0);
}
