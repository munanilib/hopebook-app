/**
 * 🏷️ 책 종류 표시: 만화 · 문제집·학습지
 *
 * 구입 검토함 · 서점 전달용 · 시리즈 후보 · 업로드 목록의 「비고」 맨 앞에 표시를 붙입니다.
 *   📙 만화                   : 만화·학습만화·웹툰 등 (읽는 책이라 사도 되지만 구분해서 보기 위해)
 *   ⚠ 문제집·학습지 의심       : 문제집·학습지·워크북·기출·수능·자습서 등 (실수로 사지 않도록 경고)
 * 서점 전달용을 만들기 전에는 문제집·학습지로 보이는 책이 있으면 한 번 더 확인을 묻습니다.
 * 제목(그리고 있으면 YES24 분야)만 보고 판단하는 추정이라, 틀릴 수 있으면 「비고」에서 표시를 지우면 됩니다.
 */
var BOOK_FLAG_COMIC = '📙 만화';
var BOOK_FLAG_WORKBOOK = '⚠ 문제집·학습지 의심';

function unifiedBookFlags_(title, category) {
  var text = String(title || '');
  var cat = String(category || '');
  var compact = text.replace(/\s+/g, '');

  var comic = /만화|코믹|웹툰|카툰|그래픽\s*노블/.test(text) ||
    /만화|웹툰|코믹/.test(cat) ||
    /^(Why\?|와이|살아남기|마법천자문|수학도둑|과학도둑|먼나라이웃나라|한국사대모험)/.test(compact.replace(/^\(.*?\)/, ''));

  var strong = /문제집|학습지|워크북|익힘책|자습서|평가문제|단원평가|기출|모의고사|수능\s*(?:특강|완성|대비|연계|개념|문제|국어|영어|수학|한국사|탐구|영역|독해)|내신\s*(?:대비|문제|완성|만점|필수|\d\s*등급)|오답\s*노트|받아쓰기|따라\s*쓰기|쓰기\s*노트|연산\s*(?:문제|훈련|노트|학습)|EBS|개념\s*원리|디딤돌|우공비|마더텅|완자|기본서|급수\s*한자|한자\s*쓰기/i.test(text) ||
    /(?:^|[\s\[(])쎈(?:[\s\])]|$)/.test(text) ||
    /참고서|수험서|자격증|문제집|학습지|교재/.test(cat);
  // 「초등 3학년 수학」처럼 학년·과목만 있는 제목은 이야기책일 수도 있어서, 읽을거리 말이 없을 때만 의심합니다.
  var gradeSubject = /(?:초등|중등|중학|고등)?\s*\d\s*학년\s*(?:\d\s*학기\s*)?(?:국어|수학|영어|과학|사회)/.test(text) &&
    !/이야기|동화|소설|위인|탐험|모험|그림책|만화|인물|여행/.test(text);

  return { comic: Boolean(comic), workbook: Boolean(strong || gradeSubject) };
}

// 비고 글 맨 앞에 붙은 표시들을 읽습니다.
function unifiedFlagsFromNote_(note) {
  var text = String(note || '');
  return { comic: text.indexOf(BOOK_FLAG_COMIC) >= 0, workbook: text.indexOf(BOOK_FLAG_WORKBOOK) >= 0 };
}

function unifiedFlagTagText_(flags) {
  var tags = [];
  if (flags && flags.workbook) tags.push(BOOK_FLAG_WORKBOOK);
  if (flags && flags.comic) tags.push(BOOK_FLAG_COMIC);
  return tags.join(' · ');
}

// 비고에 없는 표시만 맨 앞에 붙입니다. (이미 붙은 표시는 그대로 두고, 선생님이 지운 표시는 다시 붙입니다.)
function unifiedComposeNote_(note, flags) {
  var current = String(note || '').trim();
  var have = unifiedFlagsFromNote_(current);
  var add = unifiedFlagTagText_({ workbook: flags.workbook && !have.workbook, comic: flags.comic && !have.comic });
  if (!add) return current;
  return current ? add + ' · ' + current : add;
}

// 구입 검토함 비고 열(L)에 표시를 붙입니다. 바뀐 줄이 있을 때만 씁니다. 문제집·학습지 의심 권수를 돌려줍니다.
function unifiedApplyBookFlagsToReview_(sheet) {
  if (!sheet || sheet.getLastRow() < 6) return 0;
  var count = sheet.getLastRow() - 5;
  var rows = sheet.getRange(6, 1, count, 12).getValues();
  var notes = [], changed = false, workbooks = 0;
  rows.forEach(function(row) {
    var title = String(row[3] || '').trim();
    var note = String(row[11] || '');
    if (!title) { notes.push([note]); return; }
    var next = unifiedComposeNote_(note, unifiedBookFlags_(title, ''));
    if (next !== note) changed = true;
    if (unifiedFlagsFromNote_(next).workbook) workbooks++;
    notes.push([next]);
  });
  if (changed) sheet.getRange(6, 12, count, 1).setValues(notes);
  return workbooks;
}

// 서점 전달용을 만들기 전 확인: 문제집·학습지로 보이는 책을 보여 주고 그래도 포함할지 묻습니다.
// 돌려주는 값: true=계속, false=멈춤
function unifiedConfirmWorkbooks_(books) {
  var titles = [];
  (books || []).forEach(function(row) {
    var flags = unifiedBookFlags_(row[3], '');
    var noted = unifiedFlagsFromNote_(row[11]);
    if (flags.workbook || noted.workbook) titles.push(String(row[3] || '').trim());
  });
  if (!titles.length) return true;

  if (typeof UNIFIED_BAR_NO_DIALOG !== 'undefined' && UNIFIED_BAR_NO_DIALOG) return true;   // 막대 클릭으로 실행 중: 창 없이 계속(첫 번째 누름에서 알려 줬습니다)
  var ui;
  try { ui = SpreadsheetApp.getUi(); } catch (error) { return true; }
  var list = titles.slice(0, 10).map(function(title) { return '· ' + title; }).join('\n') +
    (titles.length > 10 ? '\n… 외 ' + (titles.length - 10) + '권' : '');
  var answer = ui.alert(
    '⚠ 문제집·학습지로 보이는 책이 ' + titles.length + '권 있습니다',
    list + '\n\n그래도 서점 전달용에 포함할까요?\n(아니오를 누르면 만들기를 멈춥니다. 구입 검토함에서 체크를 풀거나 「보류」로 바꿔 주세요.\n문제집이 아닌 책이면 비고의 「⚠ 문제집·학습지 의심」 표시를 지우면 됩니다.)',
    ui.ButtonSet.YES_NO
  );
  return answer === ui.Button.YES;
}
