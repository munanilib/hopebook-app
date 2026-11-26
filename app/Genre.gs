/**
 * 🎭 장르로 찾기 (웹 신청 화면)
 *
 * 책 제목을 모를 때 「공포」「로맨스」 같은 장르 버튼으로 책을 찾아 신청바구니에 담을 수 있게 합니다.
 * - YES24 키워드 검색으로 찾고(분류 번호를 따로 쓰지 않아 어느 사본에서도 같게 동작합니다),
 *   세트·전집과 문제집·학습지로 보이는 책은 결과에서 뺍니다.
 * - 같은 장르·연령 조합은 6시간 동안 저장해 두어, 여러 명이 눌러도 YES24 호출을 아낍니다.
 * - 연령(전체 / 어린이 / 청소년)은 검색어에 붙여서 좁힙니다. (정확한 구분은 아니고 추천 방식입니다.)
 */
var HR_GENRES = {
  horror:     { label: '공포',        query: '공포' },
  romance:    { label: '로맨스',      query: '로맨스' },
  mystery:    { label: '추리·미스터리', query: '추리 미스터리' },
  fantasy:    { label: '판타지',      query: '판타지' },
  sf:         { label: 'SF',          query: 'SF 과학소설' },
  adventure:  { label: '모험',        query: '모험' },
  humor:      { label: '유머',        query: '유머 웃긴 이야기' },
  friendship: { label: '우정·학교',   query: '우정 학교 이야기' },
  growth:     { label: '성장',        query: '성장 소설' },
  history:    { label: '역사',        query: '역사 이야기' },
  science:    { label: '과학',        query: '과학 이야기' },
  animal:     { label: '동물',        query: '동물 이야기' },
  sports:     { label: '스포츠',      query: '스포츠' },
  poem:       { label: '시·동시',     query: '동시 시집' }
};
var HR_GENRE_AGES = { all: '', children: '어린이', youth: '청소년' };
var HR_GENRE_CACHE_SECONDS = 21600;
var HR_GENRE_MAX_RESULTS = 30;

// 웹 화면에서 부릅니다: 장르 버튼을 눌렀을 때 책 목록을 돌려줍니다.
function HR_searchGenre(genreKey, ageKey) {
  var genre = HR_GENRES[String(genreKey || '')];
  if (!genre) throw new Error('알 수 없는 장르입니다.');
  var age = Object.prototype.hasOwnProperty.call(HR_GENRE_AGES, String(ageKey || '')) ? String(ageKey) : 'all';

  var cache = CacheService.getScriptCache();
  var cacheKey = 'HR_GENRE_v1_' + genreKey + '_' + age;
  var cached = cache.get(cacheKey);
  if (cached) {
    try { return JSON.parse(cached); } catch (ignore) {}
  }

  var query = (HR_GENRE_AGES[age] ? HR_GENRE_AGES[age] + ' ' : '') + genre.query;
  var items = yes24SearchItems_(query, HR_getYes24Key_(), 'BOOK', 50, false);

  var setPattern = /(?:세트|전집|박스(?:세트)?|묶음|전\s*\d+\s*(?:권|종)|총\s*\d+\s*(?:권|종)|\d+\s*(?:~|～|∼|-)\s*\d+\s*권|\d+\s*권\s*(?:세트|완결)|전권|합본)/i;
  var seen = {};
  var books = [];
  items.forEach(function(item) {
    var title = HR_cleanYes24Text_(item && item.title || '');
    var category = HR_cleanYes24Text_(item && item.goodsSortNm || '');
    var publisher = HR_cleanYes24Text_(item && item.publisher || '');
    if (setPattern.test(title) || /(?:전집|세트)/i.test(publisher + ' ' + category)) return;
    if (unifiedBookFlags_(title, category).workbook) return;   // 문제집·학습지로 보이는 책은 뺍니다.
    var book = yes24RequestBook_(item);
    var id = book.isbn || book.title;
    if (!book.title || seen[id]) return;
    seen[id] = true;
    books.push(book);
  });
  books = books.slice(0, HR_GENRE_MAX_RESULTS);

  try { if (books.length) cache.put(cacheKey, JSON.stringify(books), HR_GENRE_CACHE_SECONDS); } catch (ignore) {}
  return books;
}
