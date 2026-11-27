/**
 * 🎭 장르로 찾기 (웹 신청 화면) — 도서관 정보나루 API 키를 넣어야 켜집니다.
 *
 * 책 제목을 모를 때 「공포」「로맨스」 같은 장르 버튼으로 책을 찾아 신청바구니에 담을 수 있게 합니다.
 *  ① 도서관 정보나루(국립중앙도서관) 도서 검색으로 장르를 나타내는 말(예: 공포 · 호러 · 무서운 이야기)을
 *     「많이 빌려 간 순서」로 가져옵니다. 말마다 번갈아 섞어서 한 가지 말에 치우치지 않게 합니다.
 *  ② 가져온 책마다 YES24에서 ISBN으로 표지·정가를 찾아 신청에 쓸 수 있는 모양으로 만듭니다.
 * 세트·전집, 문제집·학습지로 보이는 책, 우리 학교에 이미 있는 책(소장도서)은 뺍니다. 같은 장르·연령은 6시간 저장해 호출을 아낍니다.
 * 키가 없으면 이 기능은 꺼져 있고, 웹 화면에 장르 칸이 보이지 않습니다.
 */
var HR_GENRES = {
  horror:     { label: '공포',          queries: ['공포', '호러', '무서운 이야기', '귀신 이야기'] },
  romance:    { label: '로맨스',        queries: ['로맨스', '연애', '첫사랑', '설렘'] },
  mystery:    { label: '추리·미스터리', queries: ['추리', '미스터리', '탐정', '수수께끼'] },
  fantasy:    { label: '판타지',        queries: ['판타지', '마법', '마법사', '용'] },
  sf:         { label: 'SF',            queries: ['SF', '과학소설', '우주', '로봇'] },
  adventure:  { label: '모험',          queries: ['모험', '탐험', '보물', '여행 이야기'] },
  humor:      { label: '유머',          queries: ['유머', '웃긴 이야기', '개그', '엉뚱한'] },
  friendship: { label: '우정·학교',     queries: ['우정', '친구', '학교 이야기', '짝꿍'] },
  growth:     { label: '성장',          queries: ['성장 소설', '성장', '사춘기', '꿈'] },
  history:    { label: '역사',          queries: ['역사 이야기', '한국사', '세계사', '위인'] },
  science:    { label: '과학',          queries: ['과학 이야기', '과학', '지구 우주', '발명'] },
  animal:     { label: '동물',          queries: ['동물 이야기', '동물', '반려동물', '곤충'] },
  sports:     { label: '스포츠',        queries: ['스포츠', '축구', '야구', '운동'] },
  poem:       { label: '시·동시',       queries: ['동시', '시집', '동시집', '시 그림책'] }
};
var HR_GENRE_AGES = { all: '', children: '어린이', youth: '청소년' };
var HR_GENRE_CACHE_SECONDS = 21600;
var HR_GENRE_MAX_RESULTS = 40;       // 화면에 보여 줄 최대 권수
var HR_GENRE_CANDIDATES = 60;        // 정보나루에서 모을 후보 수(세트·문제집 등을 빼고 남길 여유)
var HR_D4L_BASE = 'https://data4library.kr/api/';

// 웹 화면이 켜질 때 부릅니다: 정보나루 키가 있어 「장르로 찾기」를 쓸 수 있는지.
function HR_genreEnabled() {
  return Boolean(HR_getD4lKey_());
}

function HR_genreAllowed_(item) {
  var setPattern = /(?:세트|전집|박스(?:세트)?|묶음|전\s*\d+\s*(?:권|종)|총\s*\d+\s*(?:권|종)|\d+\s*(?:~|～|∼|-)\s*\d+\s*권|\d+\s*권\s*(?:세트|완결)|전권|합본)/i;
  var title = HR_cleanYes24Text_(item && item.title || '');
  var category = HR_cleanYes24Text_(item && item.goodsSortNm || '');
  var publisher = HR_cleanYes24Text_(item && item.publisher || '');
  if (!title) return false;
  if (setPattern.test(title) || /(?:전집|세트)/i.test(publisher + ' ' + category)) return false;
  if (unifiedBookFlags_(title, category).workbook) return false;   // 문제집·학습지로 보이는 책은 뺍니다.
  return true;
}

// 정보나루 응답에서 책 목록을 꺼냅니다. (docs: [{doc:{...}}] 모양)
function HR_d4lParseDocs_(text) {
  var data;
  try { data = JSON.parse(text); } catch (error) { throw new Error('정보나루 응답을 읽지 못했습니다.'); }
  var root = data && data.response ? data.response : data;
  if (root && root.error) {
    throw new Error('정보나루: ' + String(root.error));
  }
  var docs = root && Array.isArray(root.docs) ? root.docs : [];
  return docs.map(function(entry) { return entry && entry.doc ? entry.doc : entry; }).filter(Boolean);
}

// 정보나루 도서 검색 한 쪽: 많이 빌려 간 순서(sort=loan)
function HR_d4lSearchBooks_(keyword, page, key) {
  var url = HR_D4L_BASE + 'srchBooks?authKey=' + encodeURIComponent(key) +
    '&keyword=' + encodeURIComponent(keyword) +
    '&pageNo=' + page + '&pageSize=50&sort=loan&order=desc&format=json';
  var response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  var code = response.getResponseCode();
  var text = response.getContentText('UTF-8');
  if (code === 401 || code === 403) throw new Error('정보나루 키를 확인해 주세요. (권한 오류)');
  if (code !== 200) throw new Error('정보나루에 연결하지 못했습니다. (' + code + ')');
  return HR_d4lParseDocs_(text);
}

// ISBN13 목록 → YES24 책 정보(표지·정가 등). 5개씩 묶어서 부릅니다. 못 찾은 책은 건너뜁니다.
function HR_genreYes24ByIsbn_(isbns, key) {
  var found = {};
  for (var start = 0; start < isbns.length; start += 5) {
    var chunk = isbns.slice(start, start + 5);
    var requests = chunk.map(function(isbn) {
      return {
        url: yes24Url_('/goods/itemList', { query: isbn, category: 'BOOK', sort: 'RELATION', page: 1, pageSize: 5, detail: 'Y' }),
        method: 'get', headers: { 'X-Api-Key': key }, muteHttpExceptions: true, followRedirects: true
      };
    });
    var responses = yes24FetchAll_(requests);
    responses.forEach(function(response, i) {
      var data;
      try { data = JSON.parse(response.getContentText('UTF-8')); } catch (ignore) { return; }
      var items = data && data.data && Array.isArray(data.data.items) ? data.data.items : [];
      if (items.length) found[chunk[i]] = items[0];
    });
  }
  return found;
}

// 웹 화면에서 부릅니다: 장르 버튼을 눌렀을 때 책 목록을 돌려줍니다.
function HR_searchGenre(genreKey, ageKey) {
  var genre = HR_GENRES[String(genreKey || '')];
  if (!genre) throw new Error('알 수 없는 장르입니다.');
  var age = Object.prototype.hasOwnProperty.call(HR_GENRE_AGES, String(ageKey || '')) ? String(ageKey) : 'all';
  var d4lKey = HR_getD4lKey_();
  if (!d4lKey) throw new Error('장르로 찾기는 아직 켜져 있지 않습니다.');

  var cache = CacheService.getScriptCache();
  var cacheKey = 'HR_GENRE_D4L_v1_' + genreKey + '_' + age;
  var cached = cache.get(cacheKey);
  if (cached) {
    try { return JSON.parse(cached); } catch (ignore) {}
  }

  // ① 정보나루: 말마다 가져와 번갈아 섞습니다. (한 말이 실패해도 나머지로 계속)
  var ageWord = HR_GENRE_AGES[age];
  var lists = [];
  var firstError = null;
  genre.queries.forEach(function(word) {
    try {
      lists.push(HR_d4lSearchBooks_((ageWord ? ageWord + ' ' : '') + word, 1, d4lKey));
    } catch (error) {
      if (!firstError) firstError = error;
      lists.push([]);
    }
  });
  var totalFound = lists.reduce(function(sum, list) { return sum + list.length; }, 0);
  if (!totalFound && firstError) throw firstError;

  var merged = [];
  var longest = Math.max.apply(null, lists.map(function(list) { return list.length; }).concat([0]));
  for (var rank = 0; rank < longest; rank++) {
    lists.forEach(function(list) { if (rank < list.length) merged.push(list[rank]); });
  }
  var seen = {};
  var isbns = [];
  merged.forEach(function(doc) {
    var isbn = String(doc.isbn13 || doc.isbn || '').replace(/[^0-9Xx]/g, '');
    if (isbn.length !== 13 || seen[isbn] || isbns.length >= HR_GENRE_CANDIDATES) return;
    seen[isbn] = true;
    isbns.push(isbn);
  });

  // ② YES24에서 표지·정가를 찾아 신청에 쓰는 모양으로 만듭니다. 세트·문제집은 뺍니다.
  var yes24 = HR_genreYes24ByIsbn_(isbns, HR_getYes24Key_());
  var books = [];
  isbns.forEach(function(isbn) {
    var item = yes24[isbn];
    if (!item || !HR_genreAllowed_(item)) return;
    var book = yes24RequestBook_(item);
    if (book && book.title) books.push(book);
  });
  // 우리 학교에 이미 있는 책은 뺍니다. (이미 신청된 책은 「N명 신청」 표시와 함께 그대로 보여 줍니다.)
  books = HR_annotateBooksStatus_(books).filter(function(book) { return !book.alreadyOwned; });
  books = books.slice(0, HR_GENRE_MAX_RESULTS);

  try { if (books.length) cache.put(cacheKey, JSON.stringify(books), HR_GENRE_CACHE_SECONDS); } catch (ignore) {}
  return books;
}
