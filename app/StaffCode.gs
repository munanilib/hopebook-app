/**
 * 🔑 교직원 인증번호
 *
 * 켜 두면 웹 신청 페이지에서 「교직원」으로 들어갈 때 인증번호를 입력해야 합니다.
 * - 학부모 화면에는 인증번호가 없습니다.
 * - 번호는 이 시트의 스크립트 속성에만 저장되고, 웹 페이지 코드에는 들어가지 않습니다.
 * - 신청·취소·되돌리기·목록 보기는 서버에서도 번호를 다시 확인합니다. (화면만 가린 것이 아닙니다.)
 * - 번호를 여러 번 틀리면 잠시 막습니다.
 */
var HR_STAFF_CODE_ENABLED_PROPERTY = 'HR_STAFF_CODE_ENABLED';
var HR_STAFF_CODE_PROPERTY = 'HR_STAFF_CODE';
var HR_STAFF_CODE_FAIL_KEY = 'HR_STAFF_CODE_FAILS';
var HR_STAFF_CODE_MAX_FAILS = 10;      // 이만큼 틀리면
var HR_STAFF_CODE_LOCK_SECONDS = 300;  // 5분 동안 입력을 막습니다.

function HR_getStaffCode_() {
  return String(PropertiesService.getScriptProperties().getProperty(HR_STAFF_CODE_PROPERTY) || '').trim();
}

// 켜져 있고 번호도 정해져 있을 때만 true
function HR_isStaffCodeEnabled_() {
  var props = PropertiesService.getScriptProperties();
  return props.getProperty(HR_STAFF_CODE_ENABLED_PROPERTY) === '1' && !!HR_getStaffCode_();
}

function HR_staffCodeFailCount_() {
  return Number(CacheService.getScriptCache().get(HR_STAFF_CODE_FAIL_KEY) || 0);
}

// 번호를 확인합니다. 맞으면 true, 틀리면 false, 너무 많이 틀렸으면 오류를 냅니다.
function HR_checkStaffCode_(code) {
  if (!HR_isStaffCodeEnabled_()) return true;
  var cache = CacheService.getScriptCache();
  if (HR_staffCodeFailCount_() >= HR_STAFF_CODE_MAX_FAILS) {
    throw new Error('인증번호를 여러 번 틀려서 잠시 막혔습니다. 5분 뒤에 다시 시도해 주십시오.');
  }
  if (String(code || '').trim() === HR_getStaffCode_()) return true;
  cache.put(HR_STAFF_CODE_FAIL_KEY, String(HR_staffCodeFailCount_() + 1), HR_STAFF_CODE_LOCK_SECONDS);
  return false;
}

// 교직원 신청·취소·목록에서 부릅니다. 교직원이 아니거나 인증번호가 꺼져 있으면 그냥 지나갑니다.
function HR_requireStaffAccess_(audienceKey, code) {
  if (audienceKey !== 'staff') return;
  if (!HR_checkStaffCode_(code)) {
    throw new Error('교직원 인증번호가 맞지 않습니다. 처음 화면에서 다시 들어와 인증번호를 입력해 주십시오.');
  }
}

// 웹 화면에서 부릅니다: 입력한 번호가 맞는지 알려 줍니다.
function HR_verifyStaffCode(code) {
  if (!HR_isStaffCodeEnabled_()) return { ok: true };
  var ok = HR_checkStaffCode_(code);
  return { ok: ok };
}

// 메뉴: 「🔑 교직원 인증번호 설정」
function HR_showStaffCodeDialog() {
  var ui = SpreadsheetApp.getUi();
  var enabled = HR_isStaffCodeEnabled_();
  var answer = ui.alert(
    '🔑 교직원 인증번호 설정',
    '웹 신청 페이지에서 「교직원」으로 들어갈 때 인증번호를 입력하게 합니다. (학부모는 해당 없음)\n\n' +
    '현재 상태: ' + (enabled ? '🟢 켜짐 (번호: ' + HR_getStaffCode_() + ')' : '⚪ 꺼짐') + '\n\n' +
    '• 「예」 = ' + (enabled ? '번호 바꾸기' : '켜기(번호 정하기)') + '\n' +
    '• 「아니요」 = ' + (enabled ? '끄기' : '그대로 두기') + '\n' +
    '• 「취소」 = 닫기',
    ui.ButtonSet.YES_NO_CANCEL
  );
  var props = PropertiesService.getScriptProperties();

  if (answer === ui.Button.NO) {
    if (enabled) {
      props.setProperty(HR_STAFF_CODE_ENABLED_PROPERTY, '0');
      CacheService.getScriptCache().remove(HR_STAFF_CODE_FAIL_KEY);
      ui.alert('인증번호를 껐습니다', '이제 교직원도 번호 없이 들어갈 수 있습니다.\n(저장해 둔 번호는 지워지지 않아서, 다시 켜면 같은 번호를 쓸 수 있습니다.)', ui.ButtonSet.OK);
    }
    return;
  }
  if (answer !== ui.Button.YES) return;

  var prompt = ui.prompt(
    '인증번호 정하기',
    '교직원에게 알려 줄 번호를 입력해 주십시오.\n(4~12자, 영문·숫자만 쓸 수 있습니다. 예: 2468)',
    ui.ButtonSet.OK_CANCEL
  );
  if (prompt.getSelectedButton() !== ui.Button.OK) return;
  var code = String(prompt.getResponseText() || '').trim();
  if (!/^[0-9A-Za-z]{4,12}$/.test(code)) {
    ui.alert('번호를 확인해 주십시오', '4~12자의 영문·숫자만 쓸 수 있습니다. 다시 시도해 주십시오.', ui.ButtonSet.OK);
    return;
  }
  props.setProperty(HR_STAFF_CODE_PROPERTY, code);
  props.setProperty(HR_STAFF_CODE_ENABLED_PROPERTY, '1');
  CacheService.getScriptCache().remove(HR_STAFF_CODE_FAIL_KEY);
  ui.alert(
    '✅ 인증번호를 켰습니다',
    '교직원 인증번호: ' + code + '\n\n이 번호를 선생님들께 알려 주십시오. 열려 있던 웹 페이지는 새로고침하면 인증 창이 나옵니다.',
    ui.ButtonSet.OK
  );
}
