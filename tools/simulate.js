// Apps Script 흉내 환경에서 껍데기 + app.js 동작 확인
const fs = require('fs'); const vm = require('vm');
let appCode = fs.readFileSync('dist/app.js', 'utf8');
const store = {}; const props = {}; const triggers = []; const log = [];
const ctx = {
  console,
  CacheService: { getScriptCache: () => ({
    get: k => store[k] ?? null, getAll: ks => Object.fromEntries(ks.map(k => [k, store[k]])),
    putAll: (o) => Object.assign(store, o) }) },
  UrlFetchApp: { fetch: () => ({ getResponseCode: () => 200, getContentText: () => appCode }) },
  PropertiesService: { getScriptProperties: () => ({ getProperty: k => props[k] ?? null, setProperty: (k, v) => { props[k] = v; } }) },
  ScriptApp: { getProjectTriggers: () => triggers,
    newTrigger: n => ({ forSpreadsheet: () => ({ onOpen: () => ({ create: () => triggers.push({ getHandlerFunction: () => n }) }) }) }) },
  HtmlService: {
    SandboxMode: {}, XFrameOptionsMode: {}, getUserAgent: () => 'x',
    createHtmlOutput: h => ({ html: h, setTitle() { return this; }, addMetaTag() { return this; } }),
    createTemplate: h => ({ tpl: h, evaluate() { let out = h; for (const k in this) if (typeof this[k] === 'string' && k !== 'tpl') out = out.replace(new RegExp('<\\?=\\s*' + k + '\\s*\\?>', 'g'), this[k]); return { html: out, setTitle() { return this; } }; } })
  },
  SpreadsheetApp: {
    getActive: () => ({ toast: (m, t) => log.push('toast:' + t + ':' + m) }),
    getActiveSheet: () => ({ appendRow: r => log.push('row:' + r.slice(1).join('|')) }),
    getUi: () => ({
      createMenu: t => { const m = { t, items: [], addItem(a, f) { this.items.push(a + '→' + f); return this; }, addSeparator() { return this; }, addToUi() { log.push('menu:' + t + ' ' + this.items.join(', ')); } }; return m; },
      alert: m => log.push('alert:' + m), showSidebar: h => log.push('sidebar:' + h.html.length + 'B')
    })
  },
  Utilities: { formatDate: () => '12:00:00' }
};
vm.createContext(ctx);
vm.runInContext(fs.readFileSync('dist/Loader.gs', 'utf8'), ctx);
const run = (s) => vm.runInContext(s, ctx);

run('onOpen({})'); console.log('1) 첫 열기:', log.splice(0));
run('shellSetup()'); console.log('2) 시작하기:', log.splice(0));
run('m01()'); console.log('3) 메뉴 m01:', log.splice(0));
run('m02()'); console.log('4) 사이드바:', log.splice(0));
run('m03()'); console.log('5) 기록:', log.splice(0));
console.log('6) RPC:', JSON.stringify(run("shellRpc('getServerInfo',['web'])")));
try { run("shellRpc('appVersion_',[])"); } catch (e) { console.log('7) 비공개 함수 차단:', e.message); }
const page = run('doGet({})'); console.log('8) 웹 페이지 shim 주입:', page.html.startsWith('<script>'), page.html.includes('__gsRun') && !page.html.includes('google.script.run.withSuccessHandler'));

// GitHub 코드가 바뀐 상황: 캐시 만료 + 새 코드
appCode = appCode.replace("const TEST_MESSAGE = '첫 번째 버전입니다'", "const TEST_MESSAGE = '두 번째 버전!'");
for (const k of Object.keys(store)) if (k.startsWith('APP_FRESH')) delete store[k];
run('__APP = null'); run('shellOnOpen({})'); console.log('9) 코드 변경 후 열기:', log.splice(0));
