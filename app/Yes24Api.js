/** YES24 공식 Open API 공통 연결. 키는 ScriptProperties에만 보관합니다. */
function yes24Url_(path, params) {
  return 'https://apis.yes24.com/v1' + path + '?' + Object.keys(params || {}).map(function(k) {
    return encodeURIComponent(k) + '=' + encodeURIComponent(params[k]);
  }).join('&');
}

function yes24Isbn13_(value) {
  var isbn = String(value || '').replace(/[^0-9Xx]/g, '').toUpperCase();
  if (/^\d{9}[\dX]$/.test(isbn)) {
    var base = '978' + isbn.slice(0, 9), sum = 0;
    for (var i = 0; i < 12; i++) sum += Number(base[i]) * (i % 2 ? 3 : 1);
    return base + ((10 - sum % 10) % 10);
  }
  return /^\d{13}$/.test(isbn) ? isbn : '';
}

function yes24Decode_(response) {
  var status = response.getResponseCode(), payload;
  try { payload = JSON.parse(response.getContentText('UTF-8')); }
  catch (ignore) { throw new Error('YES24 응답을 읽지 못했습니다. HTTP ' + status); }
  var errorCode = String(payload.errorCode || '');
  if (status === 404 && /^(SEARCH_001|GOODS_001|GOODS_002|BEST_001)$/.test(errorCode)) {
    return {items: [], totalCount: 0};
  }
  if (status !== 200 || payload.success !== true) {
    var message = errorCode === 'RATE_002' ? 'YES24 일일 호출 한도에 도달했습니다. 내일 다시 실행해 주세요.' :
      errorCode === 'RATE_001' ? 'YES24 요청이 많습니다. 잠시 후 다시 실행해 주세요.' :
      status === 401 ? 'YES24 API 키를 확인해 주세요.' : 'YES24 연결 오류입니다. 잠시 후 다시 실행해 주세요.';
    throw new Error(message + ' (' + (errorCode || status) + ')');
  }
  if (!payload.data || typeof payload.data !== 'object') throw new Error('YES24 응답에 데이터가 없습니다.');
  return payload.data;
}

function yes24Fetch_(url, key) {
  if (!key) throw new Error('YES24 API 키가 없습니다. 「사본 설정 → API 키 설정」에서 입력해 주세요.');
  var response;
  for (var attempt = 0; attempt < 3; attempt++) {
    response = UrlFetchApp.fetch(url, {method: 'get', headers: {'X-Api-Key': key}, muteHttpExceptions: true, followRedirects: false});
    var code = response.getResponseCode(), payload = {};
    try { payload = JSON.parse(response.getContentText('UTF-8')); } catch (ignore) {}
    if ((code === 429 && payload.errorCode !== 'RATE_002') || code >= 500) {
      if (attempt < 2) { Utilities.sleep(1200 * (attempt + 1)); continue; }
    }
    break;
  }
  return response;
}

function yes24Get_(path, params, key) { return yes24Decode_(yes24Fetch_(yes24Url_(path, params), key)); }

function yes24SearchItems_(query, key, category, pageSize, isIsbn) {
  var isbn = isIsbn ? yes24Isbn13_(query) : '';
  if (isIsbn && !isbn) return [];
  var data = isbn ? yes24Get_('/goods/itemDetail', {searchType: 'ISBN13', query: isbn, detail: 'Y'}, key) :
    yes24Get_('/goods/itemList', {query: query, category: category || 'BOOK', sort: 'RELATION', page: 1, pageSize: pageSize || 50, detail: 'Y'}, key);
  if (!Array.isArray(data.items)) throw new Error('YES24 도서 목록을 읽지 못했습니다.');
  return data.items.filter(function(item) { return !isbn || yes24Isbn13_(item.isbn13 || item.isbn10) === isbn; });
}

function yes24RequestBook_(item) {
  var date = HR_cleanYes24Text_(item.publishDate);
  return {
    title: HR_cleanYes24Text_(item.title), author: HR_cleanYes24Text_(item.author),
    publisher: HR_cleanYes24Text_(item.publisher), pubDate: date,
    year: (date.match(/\d{4}/) || [''])[0], price: Number(item.shopPrice || 0),
    category: HR_cleanYes24Text_(item.goodsSortNm || item.goodsType),
    isbn: HR_cleanYes24Text_(item.isbn13 || item.isbn10),
    cover: HR_safeUrl_(item.cover), link: HR_safeUrl_(item.link)
  };
}

function yes24CategoryId_(label, key) {
  if (label === '종합') return '001';
  var cache = CacheService.getScriptCache(), cacheKey = 'YES24_CATEGORY_' + label;
  var cached = cache.get(cacheKey);
  if (cached) return cached;
  var data = yes24Get_('/category/list', {}, key);
  if (!Array.isArray(data.data)) throw new Error('YES24 분류 목록을 읽지 못했습니다.');
  var matches = data.data.filter(function(row) {
    var name = String(row.categoryName || '').trim();
    var path = String(row.categoryFullPath || '');
    return String(row.categoryId || '').indexOf('001') === 0 &&
      (name === label || path.indexOf(label) >= 0);
  }).sort(function(a, b) { return String(a.categoryId).length - String(b.categoryId).length || String(a.categoryId).localeCompare(String(b.categoryId)); });
  if (!matches.length) throw new Error('YES24에서 ' + label + ' 분류를 찾지 못했습니다.');
  cache.put(cacheKey, String(matches[0].categoryId), 21600);
  return String(matches[0].categoryId);
}

/** 소규모 묶음과 재시도로 초당 제한을 준수합니다. 실패는 호출자까지 전달합니다. */
function yes24FetchAll_(requests) {
  var responses = UrlFetchApp.fetchAll(requests.map(function(request) {
    request.followRedirects = false; return request;
  }));
  responses = responses.map(function(response, index) {
    var status = response.getResponseCode(), payload = {};
    try { payload = JSON.parse(response.getContentText('UTF-8')); } catch (ignore) {}
    if ((status === 429 && payload.errorCode !== 'RATE_002') || status >= 500) {
      Utilities.sleep(1200);
      response = yes24Fetch_(requests[index].url, requests[index].headers['X-Api-Key']);
    }
    var data = yes24Decode_(response);
    if (!Array.isArray(data.items)) throw new Error('YES24 도서 목록을 읽지 못했습니다.');
    return {getResponseCode: function() {return 200;}, getContentText: function() {return JSON.stringify({success:true, data:data});}};
  });
  Utilities.sleep(550);
  return responses;
}

/** 시트 데이터를 변경하지 않는 연결 검사. 공개 도서 정보만 로그에 기록합니다. */
function verifyYes24Connection() {
  var key = getConfiguredYes24Key_();
  var books = HR_searchRequestBooks('아몬드');
  if (!books.length) throw new Error('제목 검색 결과가 없습니다.');
  var isbn = books.filter(function(b) {return /^\d{13}$/.test(b.isbn);})[0].isbn;
  var exact = HR_searchRequestBooks(isbn);
  if (!exact.length || exact[0].isbn !== isbn) throw new Error('ISBN 조회 검증 실패');
  var summary = {titleSearch: books.length, isbnLookup: true, categories: {}};
  Object.keys(HR_BESTSELLER_CATEGORIES).forEach(function(name) {
    var category = HR_BESTSELLER_CATEGORIES[name];
    var rows = HR_fetchBestsellerCategory_(key, category);
    summary.categories[category.label] = rows.length;
  });
  Logger.log(JSON.stringify(summary));
  return summary;
}
