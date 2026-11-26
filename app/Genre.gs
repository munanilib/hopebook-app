/**
 * 🎭 장르로 찾기 (웹 신청 화면)
 *
 * 책 제목을 모를 때 「공포」「로맨스」 같은 장르 버튼으로 책을 찾아 신청바구니에 담을 수 있게 합니다.
 * 책을 많이, 다양하게 가져오기 위해 두 가지 방법을 합칩니다.
 *   ① YES24 분류별 인기도서: 분류 목록에서 장르에 맞는 분류를 찾아 그 분류의 인기도서를 가져옵니다.
 *   ② 키워드 검색: 장르를 나타내는 말 여러 개(예: 공포 · 호러 · 무서운 이야기 · 귀신)로 각각 두 쪽씩 검색해
 *      결과를 번갈아 섞어서 한 가지 말에 치우치지 않게 합니다.
 * 세트·전집과 문제집·학습지로 보이는 책은 뺍니다. 같은 장르·연령은 6시간 저장해 YES24 호출을 아낍니다.
 */
var HR_GENRES = {
  horror:     { label: '공포',          queries: ['공포', '호러', '무서운 이야기', '귀신 이야기'], cats: ['공포', '호러'] },
  romance:    { label: '로맨스',        queries: ['로맨스', '연애', '첫사랑', '설렘'],           cats: ['로맨스'] },
  mystery:    { label: '추리·미스터리', queries: ['추리', '미스터리', '탐정', '수수께끼'],       cats: ['추리', '미스터리'] },
  fantasy:    { label: '판타지',        queries: ['판타지', '마법', '마법사', '용'],             cats: ['판타지'] },
  sf:         { label: 'SF',            queries: ['SF', '과학소설', '우주', '로봇'],             cats: ['SF', '과학소설'] },
  adventure:  { label: '모험',          queries: ['모험', '탐험', '보물', '여행 이야기'],        cats: ['모험'] },
  humor:      { label: '유머',          queries: ['유머', '웃긴 이야기', '개그', '엉뚱한'],       cats: ['유머'] },
  friendship: { label: '우정·학교',     queries: ['우정', '친구', '학교 이야기', '짝꿍'],        cats: ['우정'] },
  growth:     { label: '성장',          queries: ['성장 소설', '성장', '사춘기', '꿈'],          cats: ['성장'] },
  history:    { label: '역사',          queries: ['역사 이야기', '한국사', '세계사', '위인'],     cats: ['역사'] },
  science:    { label: '과학',          queries: ['과학 이야기', '과학', '지구 우주', '발명'],    cats: ['과학'] },
  animal:     { label: '동물',          queries: ['동물 이야기', '동물', '반려동물', '곤충'],     cats: ['동물'] },
  sports:     { label: '스포츠',        queries: ['스포츠', '축구', '야구', '운동'],             cats: ['스포츠'] },
  poem:       { label: '시·동시',       queries: ['동시', '시집', '동시집', '시 그림책'],         cats: ['동시'] }
};
var HR_GENRE_AGES = { all: '', children: '어린이', youth: '청소년' };
var HR_GENRE_CACHE_SECONDS = 21600;
var HR_GENRE_MAX_RESULTS = 60;
var HR_GENRE_PAGES = 2;

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

// ① 장르에 맞는 YES24 분류를 찾아 그 분류의 인기도서를 가져옵니다. (못 찾거나 오류가 나면 빈 목록)
function HR_genreCategoryItems_(genre, genreKey, age, key) {
  try {
    var cache = CacheService.getScriptCache();
    var cacheKey = 'HR_GENRE_CAT_v1_' + genreKey + '_' + age;
    var cached = cache.get(cacheKey);
    var ids;
    if (cached) {
      ids = JSON.parse(cached);
    } else {
      var ageWord = HR_GENRE_AGES[age];
      var data = yes24Get_('/category/list', {}, key);
      var rows = Array.isArray(data.data) ? data.data : [];
      var matches = rows.filter(function(row) {
        var id = String(row.categoryId || '');
        var name = String(row.categoryName || '');
        var path = String(row.categoryFullPath || '');
        if (id.indexOf('001') !== 0) return false;
        if (ageWord && path.indexOf(ageWord) < 0) return false;
        return genre.cats.some(function(word) { return name.indexOf(word) >= 0 || path.indexOf(word) >= 0; });
      }).sort(function(a, b) {
        return String(a.categoryId).length - String(b.categoryId).length || String(a.categoryId).localeCompare(String(b.categoryId));
      });
      ids = matches.slice(0, 2).map(function(row) { return String(row.categoryId); });
      try { cache.put(cacheKey, JSON.stringify(ids), HR_GENRE_CACHE_SECONDS); } catch (ignore) {}
    }
    var items = [];
    ids.forEach(function(id) {
      var best = yes24Get_('/category/bestseller', { categoryId: id, page: 1, pageSize: 30, detail: 'N' }, key);
      if (Array.isArray(best.items)) items = items.concat(best.items);
    });
    return items;
  } catch (error) {
    return [];
  }
}

// ② 장르를 나타내는 말 여러 개로 검색해, 결과를 번갈아 섞어서 돌려줍니다.
function HR_genreKeywordItems_(genre, age, key) {
  var ageWord = HR_GENRE_AGES[age];
  var requests = [];
  genre.queries.forEach(function(word, queryIndex) {
    for (var page = 1; page <= HR_GENRE_PAGES; page++) {
      requests.push({
        queryIndex: queryIndex,
        request: {
          url: yes24Url_('/goods/itemList', { query: (ageWord ? ageWord + ' ' : '') + word, category: 'BOOK', sort: 'RELATION', page: page, pageSize: 50, detail: 'Y' }),
          method: 'get', headers: { 'X-Api-Key': key }, muteHttpExceptions: true, followRedirects: true
        }
      });
    }
  });

  var perQuery = genre.queries.map(function() { return []; });
  for (var start = 0; start < requests.length; start += 5) {
    var chunk = requests.slice(start, start + 5);
    var responses = yes24FetchAll_(chunk.map(function(entry) { return entry.request; }));
    responses.forEach(function(response, i) {
      var data;
      try { data = JSON.parse(response.getContentText('UTF-8')); } catch (ignore) { return; }
      var items = data && data.data && Array.isArray(data.data.items) ? data.data.items : [];
      perQuery[chunk[i].queryIndex] = perQuery[chunk[i].queryIndex].concat(items);
    });
  }

  // 말마다 1등, 2등… 순서로 번갈아 섞습니다.
  var merged = [];
  var longest = Math.max.apply(null, perQuery.map(function(list) { return list.length; }).concat([0]));
  for (var rank = 0; rank < longest; rank++) {
    perQuery.forEach(function(list) { if (rank < list.length) merged.push(list[rank]); });
  }
  return merged;
}

// 웹 화면에서 부릅니다: 장르 버튼을 눌렀을 때 책 목록을 돌려줍니다.
function HR_searchGenre(genreKey, ageKey) {
  var genre = HR_GENRES[String(genreKey || '')];
  if (!genre) throw new Error('알 수 없는 장르입니다.');
  var age = Object.prototype.hasOwnProperty.call(HR_GENRE_AGES, String(ageKey || '')) ? String(ageKey) : 'all';

  var cache = CacheService.getScriptCache();
  var cacheKey = 'HR_GENRE_v2_' + genreKey + '_' + age;
  var cached = cache.get(cacheKey);
  if (cached) {
    try { return JSON.parse(cached); } catch (ignore) {}
  }

  var key = HR_getYes24Key_();
  var items = HR_genreCategoryItems_(genre, String(genreKey), age, key);   // 인기도서 먼저
  items = items.concat(HR_genreKeywordItems_(genre, age, key));            // 키워드 검색 결과를 이어서

  var seen = {};
  var books = [];
  items.forEach(function(item) {
    if (!HR_genreAllowed_(item)) return;
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
