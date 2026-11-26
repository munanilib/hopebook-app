#!/usr/bin/env node
/**
 * 앱 소스(app/*.gs, app/*.html)를 하나의 dist/app.js 로 묶고,
 * 사본에 넣을 껍데기 dist/Loader.gs 를 만듭니다.
 *
 *   node tools/build.js
 *
 * app/app.config.json
 *   { "codeUrl": "https://raw.githubusercontent.com/OWNER/REPO/stable/dist/app.js",
 *     "version": "2026.11.22-1",
 *     "extraEntryPoints": ["버튼에 연결한 함수 이름", ...] }
 */
const fs = require('fs');
const path = require('path');

const ROOT = path.resolve(__dirname, '..');
const APP = path.join(ROOT, 'app');
const DIST = path.join(ROOT, 'dist');
const cfg = JSON.parse(fs.readFileSync(path.join(APP, 'app.config.json'), 'utf8'));

const SIMPLE_TRIGGERS = ['onEdit', 'onSelectionChange', 'onChange'];
const SHELL_OWNED = new Set(['onOpen', 'doGet', 'doPost', ...SIMPLE_TRIGGERS]);

const names = fs.readdirSync(APP).sort();
const scripts = names.filter(n => /\.(gs|js)$/.test(n));
const htmls = names.filter(n => /\.html$/.test(n));

let code = '';
for (const f of scripts) code += `\n/* ===== ${f} ===== */\n` + fs.readFileSync(path.join(APP, f), 'utf8') + '\n';

// 빌드 번호: 소스 내용으로 만든 짧은 번호입니다. (시간이 아니라 내용 기준이라, 같은 소스면 언제 빌드해도 같습니다.
// 그래서 배포 검사에서 「dist 가 최신인가」를 정확히 비교할 수 있고, 화면 맨 아래에서 버전이 같은지 확인할 수 있습니다.)
// 웹 화면(HTML)이 바뀌어도 번호가 바뀌어야 합니다. (사본의 껍데기는 이 번호가 같으면 새 코드를 받지 않습니다.)
const htmlForHash = htmls.map(f => fs.readFileSync(path.join(APP, f), 'utf8')).join('\n');
const buildId = 'v' + require('crypto').createHash('sha1').update(code + '\n/*HTML*/\n' + htmlForHash).digest('hex').slice(0, 7);
code = code.split('__BUILD_ID__').join(buildId);

// 최상위 함수 이름
const fnNames = [...new Set([...code.matchAll(/^function\s+([A-Za-z_$][\w$]*)\s*\(/gm)].map(m => m[1]))];

// 트리거로 등록되는 함수 이름: newTrigger('이름') 또는 newTrigger(상수)
const triggerNames = new Set();
for (const m of code.matchAll(/newTrigger\(\s*(['"])([\w$]+)\1\s*\)/g)) triggerNames.add(m[2]);
for (const m of code.matchAll(/newTrigger\(\s*([A-Za-z_$][\w$]*)\s*\)/g)) {
  const re = new RegExp(`(?:const|let|var)\\s+${m[1]}\\s*=\\s*(['"])([\\w$]+)\\1`);
  const c = code.match(re);
  if (c) triggerNames.add(c[2]);
  else console.warn(`⚠️  newTrigger(${m[1]}) 의 함수 이름을 찾지 못했습니다. extraEntryPoints 에 직접 넣어 주세요.`);
}
for (const n of cfg.extraEntryPoints || []) triggerNames.add(n);

const missing = [...triggerNames].filter(n => !fnNames.includes(n));
if (missing.length) { console.error('❌ 앱에 없는 입구 함수:', missing.join(', ')); process.exit(1); }

// HTML 파일
const files = {};
for (const f of htmls) files[f.replace(/\.html$/, '')] = fs.readFileSync(path.join(APP, f), 'utf8');

// google.script.run 으로 부를 수 있는 함수 = 이름이 _ 로 끝나지 않는 함수 (Apps Script 규칙과 같음)
const rpc = fnNames.filter(n => !n.endsWith('_') && !SHELL_OWNED.has(n));

const bundle = `/* 자동 생성 파일 — 직접 고치지 마세요. (node tools/build.js)  v${cfg.version} */
var APP_FACTORY = function (__host) {
  var HtmlService = __host.html;
  var SpreadsheetApp = __host.spreadsheet;
${code}
  var __fns = {${fnNames.map(n => `\n    ${JSON.stringify(n)}: ${n}`).join(',')}
  };
  return {
    version: ${JSON.stringify(cfg.version)},
    files: ${JSON.stringify(files)},
    menus: (typeof APP_MENUS !== 'undefined') ? APP_MENUS : [],
    rpc: ${JSON.stringify(rpc)},
    has: function (n) { return typeof __fns[n] === 'function'; },
    call: function (n, a) {
      if (typeof __fns[n] !== 'function') throw new Error('함수를 찾을 수 없습니다: ' + n);
      return __fns[n].apply(null, a || []);
    }
  };
};
`;

fs.mkdirSync(DIST, { recursive: true });
fs.writeFileSync(path.join(DIST, 'app.js'), bundle);
// 아주 작은 버전 번호 파일: 사본의 껍데기가 이것만 먼저 확인해서, 바뀌지 않았으면 큰 파일을 다시 받지 않습니다.
fs.writeFileSync(path.join(DIST, 'version.txt'), buildId + '\n');

// 껍데기
let shell = fs.readFileSync(path.join(ROOT, 'shell', 'Loader.gs'), 'utf8');
shell = shell.replace(/var APP_CODE_URL = '[^']*';/, `var APP_CODE_URL = '${cfg.codeUrl}';`);
const entries = [];
for (const n of SIMPLE_TRIGGERS) if (fnNames.includes(n)) entries.push(`function ${n}(e) { return shellSimple_('${n}', e); }`);
for (const n of [...triggerNames].sort()) entries.push(`function ${n}(e) { return shellTrigger_('${n}', e); }`);
shell = shell.replace('//@@ENTRY_POINTS@@', entries.join('\n') || '// (없음)');
fs.writeFileSync(path.join(DIST, 'Loader.gs'), shell);

// 사본 매니페스트 (권한 범위는 껍데기에서 고정)
fs.writeFileSync(path.join(DIST, 'appsscript.json'), JSON.stringify(cfg.manifest, null, 2) + '\n');

console.log(`✅ v${cfg.version}  함수 ${fnNames.length}개 · HTML ${htmls.length}개 · 입구 ${entries.length}개 · ${(bundle.length / 1024).toFixed(0)}KB`);
