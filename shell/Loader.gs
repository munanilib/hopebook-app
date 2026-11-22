/**
 * =====================================================================
 * 🧩 껍데기 (Loader.gs)
 *
 *  이 파일은 각 학교 사본에 들어가며, 거의 바뀌지 않습니다.
 *  실제 기능 코드는 실행할 때마다 GitHub의 app.js 에서 읽어 옵니다.
 *  → 원본 저장소(stable 브랜치)만 고치면 모든 사본에 자동 반영됩니다.
 *
 *  ⚠️ 이 파일을 고치지 마세요. 기능 수정은 GitHub 저장소에서 합니다.
 * =====================================================================
 */

var SHELL_VERSION = '1';
var APP_CODE_URL = 'https://raw.githubusercontent.com/OWNER/REPO/stable/dist/app.js';

var APP_CACHE_SECONDS  = 300;    // 5분마다 GitHub에서 새 코드 확인
var APP_BACKUP_SECONDS = 21600;  // GitHub에 못 닿을 때 쓸 예비본 보관 시간(최대 6시간)
var APP_CHUNK_CHARS    = 25000;  // 캐시 한 칸 크기 (한글 3바이트 × 25000 < 100KB)
var MENU_SLOT_COUNT    = 60;     // 메뉴 항목 최대 개수

var __APP = null;


/* ------------------------------------------------------------------ */
/* 1. 코드 읽어 오기 (캐시 → GitHub → 예비본)                           */
/* ------------------------------------------------------------------ */
function shellApp_() {
  if (__APP) return __APP;
  var code = shellReadCode_();
  var factory = (0, eval)(code + '\n;APP_FACTORY');
  __APP = factory(shellHost_());
  return __APP;
}

function shellReadCode_(forceFresh) {
  var cache = CacheService.getScriptCache();
  if (!forceFresh) {
    var fresh = shellCacheGet_(cache, 'APP_FRESH');
    if (fresh) return fresh;
  }
  try {
    var res = UrlFetchApp.fetch(APP_CODE_URL + '?t=' + Date.now(), { muteHttpExceptions: true });
    if (res.getResponseCode() !== 200) throw new Error('HTTP ' + res.getResponseCode());
    var code = res.getContentText('UTF-8');
    if (code.indexOf('APP_FACTORY') < 0) throw new Error('올바른 app.js 가 아닙니다.');
    shellCachePut_(cache, 'APP_FRESH', code, APP_CACHE_SECONDS);
    shellCachePut_(cache, 'APP_BACKUP', code, APP_BACKUP_SECONDS);
    return code;
  } catch (e) {
    var backup = shellCacheGet_(cache, 'APP_BACKUP');
    if (backup) {
      console.warn('GitHub 연결 실패, 예비본 사용: ' + e);
      return backup;
    }
    throw new Error('프로그램 코드를 불러오지 못했습니다. 인터넷 연결을 확인하고 잠시 뒤 다시 시도해 주세요.\n(' + e + ')');
  }
}

function shellCachePut_(cache, key, text, seconds) {
  var parts = {};
  var n = Math.ceil(text.length / APP_CHUNK_CHARS);
  for (var i = 0; i < n; i++) parts[key + '_' + i] = text.substr(i * APP_CHUNK_CHARS, APP_CHUNK_CHARS);
  parts[key + '_n'] = String(n);
  cache.putAll(parts, seconds);
}

function shellCacheGet_(cache, key) {
  var n = Number(cache.get(key + '_n') || 0);
  if (!n) return null;
  var keys = [];
  for (var i = 0; i < n; i++) keys.push(key + '_' + i);
  var got = cache.getAll(keys);
  var out = '';
  for (var j = 0; j < n; j++) {
    if (got[keys[j]] == null) return null;
    out += got[keys[j]];
  }
  return out;
}


/* ------------------------------------------------------------------ */
/* 2. 앱이 쓰는 도구 (HTML 파일 · google.script.run 연결)                */
/* ------------------------------------------------------------------ */
function shellHost_() {
  return {
    shellVersion: SHELL_VERSION,
    html: shellHtmlService_(),
    spreadsheet: shellSpreadsheetApp_()
  };
}

// 앱 코드 안의 HtmlService 를 이 객체로 바꿔 끼웁니다.
// createHtmlOutputFromFile('이름') 등이 GitHub 에서 받아 온 HTML 을 쓰도록 합니다.
function shellHtmlService_() {
  var real = HtmlService;
  var wrap = {
    SandboxMode: real.SandboxMode,
    XFrameOptionsMode: real.XFrameOptionsMode,
    getUserAgent: function () { return real.getUserAgent(); }
  };
  wrap.createHtmlOutputFromFile = function (name) {
    return real.createHtmlOutput(shellPrepareHtml_(shellFile_(name)));
  };
  wrap.createTemplateFromFile = function (name) {
    return real.createTemplate(shellPrepareHtml_(shellFile_(name)));
  };
  wrap.createHtmlOutput = function (x) {
    return typeof x === 'string' ? real.createHtmlOutput(shellPrepareHtml_(x)) : real.createHtmlOutput(x);
  };
  wrap.createTemplate = function (x) {
    return typeof x === 'string' ? real.createTemplate(shellPrepareHtml_(x)) : real.createTemplate(x);
  };
  return wrap;
}

function shellFile_(name) {
  var files = shellApp_().files || {};
  var key = String(name).replace(/\.html$/, '');
  if (!(key in files)) throw new Error('HTML 파일을 찾을 수 없습니다: ' + name);
  return files[key];
}

// HTML 속 google.script.run.함수() 호출을 껍데기의 shellRpc 로 이어 줍니다.
var SHELL_CLIENT_SHIM_ =
  '<script>(function(){' +
  'function mk(r){return new Proxy({},{get:function(_,p){' +
  'if(p==="withSuccessHandler"||p==="withFailureHandler"||p==="withUserObject")' +
  '{return function(){return mk(r[p].apply(r,arguments));};}' +
  'return function(){return r.shellRpc(String(p),Array.prototype.slice.call(arguments));};' +
  '}});}' +
  'window.__gsRun=mk(google.script.run);' +
  'try{Object.defineProperty(window,"__gsRunReady",{value:true});}catch(e){}' +
  '})();</script>';

function shellPrepareHtml_(html) {
  if (html.indexOf('google.script.run') < 0) return html;
  html = html.replace(/google\.script\.run/g, '__gsRun');
  return SHELL_CLIENT_SHIM_ + html;
}


/* ------------------------------------------------------------------ */
/* 3. 입구들 (Apps Script 가 이름으로 부르는 함수)                       */
/* ------------------------------------------------------------------ */

// 화면(HTML)에서 부르는 서버 함수. 앱이 허용한 이름만 실행합니다.
function shellRpc(name, args) {
  shellRpcLog_('받음 ' + name);
  try {
    var app = shellApp_();
    if (!app.rpc || app.rpc.indexOf(name) < 0) throw new Error('허용되지 않은 호출입니다: ' + name);
    var out = app.call(name, args || []);
    shellRpcLog_('성공 ' + name + ' → ' + JSON.stringify(out));
    return out;
  } catch (err) {
    shellRpcLog_('오류 ' + name + ' → ' + String(err && err.stack || err));
    throw err;
  }
}

function shellRpcLog_(msg) {
  try {
    var ss = SpreadsheetApp.getActive();
    var sh = ss.getSheetByName('_rpc') || ss.insertSheet('_rpc');
    sh.appendRow([new Date(), msg]);
  } catch (e) {}
}

function doGet(e)  { return shellApp_().call('doGet', [e]); }
function doPost(e) { return shellApp_().call('doPost', [e]); }

// 시트를 열 때: 권한 승인 전이면 「시작하기」 메뉴만, 승인 후에는 설치형 트리거가 메뉴를 만듭니다.
function onOpen(e) {
  if (PropertiesService.getScriptProperties().getProperty('SHELL_READY') === '1') return;
  SpreadsheetApp.getUi()
    .createMenu('⚙️ 시작하기')
    .addItem('✅ 프로그램 사용 시작 (권한 승인)', 'shellSetup')
    .addItem('🩺 진단 (문제 확인)', 'shellDiag')
    .addToUi();
}

// 단계별로 시험해서 결과를 '_진단' 시트에 적습니다.
function shellDiag() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName('_진단') || ss.insertSheet('_진단');
  sh.clear();
  var rows = [['단계', '결과']];
  function step(name, fn) {
    try { rows.push([name, 'OK ' + (fn() || '')]); }
    catch (err) { rows.push([name, 'ERROR ' + String(err && err.stack || err)]); }
  }
  var props = PropertiesService.getScriptProperties();
  step('1 속성 읽기', function () { return props.getProperty('SHELL_READY'); });
  step('2 속성 쓰기', function () { props.setProperty('DIAG', '1'); return 'ok'; });
  step('3 트리거 목록', function () {
    return ScriptApp.getProjectTriggers().map(function (t) { return t.getHandlerFunction() + ':' + t.getEventType(); }).join(', ');
  });
  step('4 코드 받기', function () { return shellReadCode_(true).length + '자'; });
  step('5 eval', function () { __APP = null; return typeof shellApp_(); });
  step('6 메뉴 목록', function () { return JSON.stringify(shellApp_().menus); });
  step('7 메뉴맵 저장값', function () { return props.getProperty('SHELL_MENU_MAP'); });
  step('8 마지막 오류', function () { return props.getProperty('SHELL_LAST_ERROR'); });
  step('9 m01 실행', function () { return shellMenu_('m01'); });
  sh.getRange(1, 1, rows.length, 2).setValues(rows);
  sh.setColumnWidth(2, 700);
  ss.setActiveSheet(sh);
}

function shellSetup() {
  try {
    shellSetupInner_();
  } catch (err) {
    SpreadsheetApp.getUi().alert('설정 중 오류', String(err && err.stack || err), SpreadsheetApp.getUi().ButtonSet.OK);
    throw err;
  }
}

function shellSetupInner_() {
  var ss = SpreadsheetApp.getActive();
  var has = ScriptApp.getProjectTriggers().some(function (t) {
    return t.getHandlerFunction() === 'shellOnOpen';
  });
  if (!has) ScriptApp.newTrigger('shellOnOpen').forSpreadsheet(ss).onOpen().create();
  PropertiesService.getScriptProperties().setProperty('SHELL_READY', '1');
  shellReadCode_(true);
  shellOnOpen({ source: ss });
  ss.toast('이제 시트를 열 때마다 최신 기능이 자동으로 적용됩니다.', '✅ 준비 완료', 8);
}

// 설치형 열기 트리거: 최신 코드로 메뉴를 만들고 앱의 onOpen 을 실행합니다.
function shellOnOpen(e) {
  try {
    var app = shellApp_();
    shellBuildMenus_(app);
    PropertiesService.getScriptProperties().deleteProperty('SHELL_LAST_ERROR');
    if (app.has('onOpen')) app.call('onOpen', [e]);
  } catch (err) {
    PropertiesService.getScriptProperties().setProperty('SHELL_LAST_ERROR', String(err && err.stack || err));
    SpreadsheetApp.getUi()
      .createMenu('⚠️ 시작하기 (오류)')
      .addItem('🔍 오류 내용 보기', 'shellShowError')
      .addItem('🔧 다시 설정', 'shellSetup')
      .addToUi();
  }
}

function shellShowError() {
  var msg = PropertiesService.getScriptProperties().getProperty('SHELL_LAST_ERROR') || '(기록된 오류 없음)';
  SpreadsheetApp.getUi().alert('열기 중 오류', msg, SpreadsheetApp.getUi().ButtonSet.OK);
}

// ---- 메뉴 슬롯: 함수 이름 → m01~m60 (이름별로 고정해서 저장) ----
// 앱 코드가 createMenu(...).addItem('이름', '함수이름') 으로 직접 메뉴를 만들어도
// 껍데기가 가로채서 슬롯(m01~)에 연결하므로 원본 코드를 고치지 않아도 됩니다.
var __SLOTS = null;

function shellSlots_() {
  if (!__SLOTS) __SLOTS = JSON.parse(PropertiesService.getScriptProperties().getProperty('SHELL_SLOTS') || '{}');
  return __SLOTS;
}

function shellSlot_(fnName) {
  var slots = shellSlots_();
  if (slots[fnName]) return slots[fnName];
  var lock = LockService.getScriptLock();
  var locked = false;
  try { locked = lock.tryLock(5000); } catch (e) {}
  try {
    __SLOTS = null;
    slots = shellSlots_();
    if (!slots[fnName]) {
      var used = {};
      Object.keys(slots).forEach(function (k) { used[slots[k]] = true; });
      var id = null;
      for (var i = 1; i <= MENU_SLOT_COUNT; i++) {
        var cand = 'm' + (i < 10 ? '0' : '') + i;
        if (!used[cand]) { id = cand; break; }
      }
      if (!id) throw new Error('메뉴 항목이 너무 많습니다(최대 ' + MENU_SLOT_COUNT + '개).');
      slots[fnName] = id;
      PropertiesService.getScriptProperties().setProperty('SHELL_SLOTS', JSON.stringify(slots));
    }
    return slots[fnName];
  } finally {
    if (locked) lock.releaseLock();
  }
}

function shellMenuBuilder_(menu) {
  var b = {
    __real: menu,
    addItem: function (label, fn) { menu.addItem(label, shellSlot_(fn)); return b; },
    addSeparator: function () { menu.addSeparator(); return b; },
    addSubMenu: function (sub) { menu.addSubMenu(sub && sub.__real ? sub.__real : sub); return b; },
    addToUi: function () { menu.addToUi(); return b; },
    addToSpreadsheet: function () { return menu.addToUi(); }
  };
  return b;
}

function shellWrapUi_(ui) {
  return new Proxy(ui, {
    get: function (t, p) {
      if (p === 'createMenu') return function (title) { return shellMenuBuilder_(t.createMenu(title)); };
      var v = t[p];
      return typeof v === 'function' ? v.bind(t) : v;
    }
  });
}

// 앱 안의 SpreadsheetApp 을 이것으로 바꿔 끼웁니다 (getUi().createMenu 만 가로챔).
function shellSpreadsheetApp_() {
  var real = SpreadsheetApp;
  return new Proxy(real, {
    get: function (t, p) {
      if (p === 'getUi') return function () { return shellWrapUi_(t.getUi()); };
      var v = t[p];
      return typeof v === 'function' ? v.bind(t) : v;
    }
  });
}

// 앱이 목록(APP_MENUS)으로 정한 메뉴
function shellBuildMenus_(app) {
  var ui = shellWrapUi_(SpreadsheetApp.getUi());
  (app.menus || []).forEach(function (menu) {
    var m = ui.createMenu(menu.title);
    (menu.items || []).forEach(function (it) {
      if (it === '-') m.addSeparator(); else m.addItem(it[0], it[1]);
    });
    m.addToUi();
  });
}

function shellMenu_(id) {
  try {
    var slots = shellSlots_();
    var name = null;
    Object.keys(slots).forEach(function (k) { if (slots[k] === id) name = k; });
    if (!name) { __SLOTS = null; slots = shellSlots_(); Object.keys(slots).forEach(function (k) { if (slots[k] === id) name = k; }); }
    if (!name) throw new Error('메뉴가 오래되었습니다. 시트를 새로고침해 주세요. (' + id + ')');
    return shellApp_().call(name, []);
  } catch (err) {
    shellLogError_('메뉴 ' + id, err);
    try { SpreadsheetApp.getUi().alert('⚠️ 오류 (' + id + ')', String(err && err.stack || err), SpreadsheetApp.getUi().ButtonSet.OK); } catch (e3) {}
    throw err;
  }
}

// 오류를 '_오류' 시트에 적어 둡니다 (원인 확인용)
function shellLogError_(where, err) {
  try {
    var ss = SpreadsheetApp.getActive();
    var sh = ss.getSheetByName('_오류') || ss.insertSheet('_오류');
    sh.appendRow([new Date(), where, String(err && err.stack || err)]);
  } catch (e2) {}
}

// 설치형 트리거 등 이름으로 불리는 앱 함수들 (빌드가 목록을 만들어 줍니다)
function shellTrigger_(name, e) { return shellApp_().call(name, [e]); }

// 단순 트리거(onEdit 등)는 인터넷을 쓸 수 없어서, 저장해 둔 코드가 있을 때만 실행합니다.
function shellSimple_(name, e) {
  try {
    if (!__APP) {
      var cache = CacheService.getScriptCache();
      var code = shellCacheGet_(cache, 'APP_FRESH') || shellCacheGet_(cache, 'APP_BACKUP');
      if (!code) return;
      __APP = (0, eval)(code + '\n;APP_FACTORY')(shellHost_());
    }
    if (__APP.has(name)) return __APP.call(name, [e]);
  } catch (err) {
    console.warn(name + ' 실행 실패: ' + err);
  }
}

/* ---- 메뉴 슬롯 ---- */
function m01(){return shellMenu_('m01');} function m02(){return shellMenu_('m02');} function m03(){return shellMenu_('m03');}
function m04(){return shellMenu_('m04');} function m05(){return shellMenu_('m05');} function m06(){return shellMenu_('m06');}
function m07(){return shellMenu_('m07');} function m08(){return shellMenu_('m08');} function m09(){return shellMenu_('m09');}
function m10(){return shellMenu_('m10');} function m11(){return shellMenu_('m11');} function m12(){return shellMenu_('m12');}
function m13(){return shellMenu_('m13');} function m14(){return shellMenu_('m14');} function m15(){return shellMenu_('m15');}
function m16(){return shellMenu_('m16');} function m17(){return shellMenu_('m17');} function m18(){return shellMenu_('m18');}
function m19(){return shellMenu_('m19');} function m20(){return shellMenu_('m20');} function m21(){return shellMenu_('m21');}
function m22(){return shellMenu_('m22');} function m23(){return shellMenu_('m23');} function m24(){return shellMenu_('m24');}
function m25(){return shellMenu_('m25');} function m26(){return shellMenu_('m26');} function m27(){return shellMenu_('m27');}
function m28(){return shellMenu_('m28');} function m29(){return shellMenu_('m29');} function m30(){return shellMenu_('m30');}
function m31(){return shellMenu_('m31');} function m32(){return shellMenu_('m32');} function m33(){return shellMenu_('m33');}
function m34(){return shellMenu_('m34');} function m35(){return shellMenu_('m35');} function m36(){return shellMenu_('m36');}
function m37(){return shellMenu_('m37');} function m38(){return shellMenu_('m38');} function m39(){return shellMenu_('m39');}
function m40(){return shellMenu_('m40');} function m41(){return shellMenu_('m41');} function m42(){return shellMenu_('m42');}
function m43(){return shellMenu_('m43');} function m44(){return shellMenu_('m44');} function m45(){return shellMenu_('m45');}
function m46(){return shellMenu_('m46');} function m47(){return shellMenu_('m47');} function m48(){return shellMenu_('m48');}
function m49(){return shellMenu_('m49');} function m50(){return shellMenu_('m50');} function m51(){return shellMenu_('m51');}
function m52(){return shellMenu_('m52');} function m53(){return shellMenu_('m53');} function m54(){return shellMenu_('m54');}
function m55(){return shellMenu_('m55');} function m56(){return shellMenu_('m56');} function m57(){return shellMenu_('m57');}
function m58(){return shellMenu_('m58');} function m59(){return shellMenu_('m59');} function m60(){return shellMenu_('m60');}

/* ---- 트리거·버튼 입구 (빌드가 아래에 추가합니다) ---- */
//@@ENTRY_POINTS@@
