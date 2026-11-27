/**
 * 보유목록 시리즈 누락·후속권 점검
 * 기존 도서 관리 기능과 독립적으로 동작합니다.
 */

var SERIES_FINDER_SHEET = '🔎 시리즈 누락 점검';
var SERIES_HOLDINGS_SHEET = '🏫 우리학교 소장도서';
var SERIES_WORK_SHEET = '📋 시리즈 업무목록';
var SERIES_MAX_GROUPS = 3000;
var SERIES_MAX_API_REQUESTS = 5000;
var SERIES_DAILY_API_SAFE_LIMIT = 18000;
var SERIES_SEARCH_CHUNK = 5;
var SERIES_MAX_ENRICH_REQUESTS = 20;
var SERIES_EXECUTION_BUDGET_MS = 285000;
var SERIES_FOLLOWUP_RESULTS = 50;
var SERIES_GAP_RESULTS = 50;
var SERIES_MAX_GAP_PAGES = 3;
var SERIES_RESUME_STATE_KEY = 'YES24_SERIES_RESUME_STATE_V1';
var SERIES_AUTO_CONTINUE_HANDLER = 'seriesAutoContinueMissingBooks_';
var SERIES_AUTO_CONTINUE_DELAY_MS = 60000;

function setupSeriesFinder() {
  seriesCancelAutoContinue_();
  seriesClearResumeState_();
  seriesInstallOpenTrigger_();
  seriesInstallWorkButtonTrigger_();
  seriesBuildResultSheet_([], {
    seriesCount: 0,
    searchedCount: 0,
    checkedAt: new Date()
  });
  seriesSpreadsheet_().toast(
    '시리즈 시트 위쪽의 「🔎 시리즈 누락 점검 시작 ›」 막대 맨 앞 ☐ 칸을 체크해 주세요.',
    '설정 완료',
    8
  );
}

function seriesFinderOnOpen_() {
  seriesAddMenu_();
}

// 시리즈 점검 메뉴는 「🧰 통합 도서업무」 안의 하위 메뉴로 합쳤습니다.
// (예전 설치형 열기 트리거가 이 함수를 불러도 메뉴 줄에 따로 만들지 않습니다.)
function seriesAddMenu_() {
}

function seriesInstallOpenTrigger_() {
  var handler = 'seriesFinderOnOpen_';
  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    if (trigger.getHandlerFunction() === handler) {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  ScriptApp.newTrigger(handler)
    .forSpreadsheet(seriesSpreadsheet_())
    .onOpen()
    .create();
}


function installSeriesWorkButton() {
  seriesInstallWorkButtonTrigger_();

  var sheet = seriesSpreadsheet_().getSheetByName(SERIES_FINDER_SHEET);
  if (!sheet) {
    throw new Error('「' + SERIES_FINDER_SHEET + '」 시트를 찾을 수 없습니다.');
  }

  seriesAddWorkButton_(sheet);
  seriesSpreadsheet_().toast(
    '책을 체크한 뒤 결과 시트 위쪽의 금색 버튼을 누르면 업무목록으로 이동합니다.',
    '업무목록 버튼 설치 완료',
    8
  );
}

function seriesInstallWorkButtonTrigger_() {
  var handler = 'seriesWorkButtonOnEdit_';
  var alreadyInstalled = ScriptApp.getProjectTriggers().some(function(trigger) {
    return trigger.getHandlerFunction() === handler;
  });

  if (!alreadyInstalled) {
    ScriptApp.newTrigger(handler)
      .forSpreadsheet(seriesSpreadsheet_())
      .onEdit()
      .create();
  }
}

function seriesWorkButtonOnEdit_(e) {
    // 상단 체크박스는 Code.gs의 onEdit에서 한 번만 처리합니다.
  return;

  if (!e || !e.range) return;

  var range = e.range;
  var sheet = range.getSheet();
  if (
    sheet.getName() !== SERIES_FINDER_SHEET ||
    range.getA1Notation() !== 'A4' ||
    String(e.value || '').toUpperCase() !== 'TRUE'
  ) {
    return;
  }

  var lock = LockService.getDocumentLock();
  if (!lock.tryLock(3000)) {
    range.setValue(false);
    seriesSpreadsheet_().toast('다른 작업을 처리 중입니다. 잠시 후 다시 눌러 주세요.','잠시만요',5);
    return;
  }

  try {
    sendSelectedSeriesToReview();
  } catch (error) {
    seriesSpreadsheet_().toast(
      String(error && error.message ? error.message : error),
      '구입 검토함 전송 오류',
      8
    );
  } finally {
    range.setValue(false);
    lock.releaseLock();
  }
}

function seriesAddWorkButton_(sheet) {
  sheet.getRange('A4:N4').breakApart();
  sheet.getRange('A4:N4').clearContent().clearDataValidations();

  sheet.getRange('A4')
    .insertCheckboxes()
    .setValue(false)
    .setBackground('#2F5D50')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  sheet.getRange('B4:C4')
    .merge()
    .setValue('전체 선택/해제')
    .setBackground('#2F5D50')
    .setFontColor('#FFFFFF')
    .setFontSize(13)
    .setFontWeight('bold')
    .setHorizontalAlignment('left')
    .setVerticalAlignment('middle');

  sheet.getRange('D4')
    .insertCheckboxes()
    .setValue(false)
    .setBackground('#D9A441')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  sheet.getRange('E4:N4')
    .merge()
    .setValue('선택한 책을 구입 검토함으로 보내기')
    .setBackground('#D9A441')
    .setFontColor('#FFFFFF')
    .setFontSize(14)
    .setFontWeight('bold')
    .setHorizontalAlignment('left')
    .setVerticalAlignment('middle');

  sheet.setRowHeight(4, 44);
  sheet.getRange('A4:C4').setBorder(
    true,true,true,true,false,false,
    '#244A40',SpreadsheetApp.BorderStyle.SOLID_MEDIUM
  );
  sheet.getRange('D4:N4').setBorder(
    true,true,true,true,false,false,
    '#B88428',SpreadsheetApp.BorderStyle.SOLID_MEDIUM
  );

  unifiedDrawActionBars_(sheet);   // 체크박스 대신 클릭 막대
}

function setSeriesYes24Key() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt(
    'YES24 API 키 설정',
    'YES24 API 키를 입력하세요. 기존 프로젝트 키가 있다면 입력하지 않아도 됩니다.',
    ui.ButtonSet.OK_CANCEL
  );
  if (response.getSelectedButton() !== ui.Button.OK) return;

  var key = String(response.getResponseText() || '').trim();
  if (!key) {
    ui.alert('API 키를 입력해 주세요.');
    return;
  }
  HR_testYes24Key_(key);
  PropertiesService.getScriptProperties().setProperty('HR_YES24_API_KEY', key);
  ui.alert('YES24 API 키를 저장했습니다.');
}

function refreshMissingSeriesBooks() {
  var runStartedAt = Date.now();
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) {
    seriesSpreadsheet_().toast(
      '다른 시리즈 점검이 진행 중입니다. 잠시 후 다시 시도해 주세요.',
      '시리즈 점검',
      7
    );
    return;
  }

  try {
    var spreadsheet = seriesSpreadsheet_();
    var key = seriesGetYes24Key_();
    if (!key) {
      throw new Error('YES24 API 키가 없습니다. 「🔎 시리즈 점검 → YES24 API 키 설정」에서 입력해 주세요.');
    }

    spreadsheet.toast('보유목록에서 시리즈와 빈 권차를 찾고 있습니다.', '시리즈 점검', 5);
    var holdings = seriesReadHoldings_();
    var allGroups = seriesGroupHoldings_(holdings.rows).sort(seriesGroupPriority_);
    var groups = allGroups.slice(0, SERIES_MAX_GROUPS);
    var groupLimitReached = allGroups.length > groups.length;
    var fingerprint = seriesResumeFingerprint_(groups);
    var properties = PropertiesService.getScriptProperties();
    var hadSavedResume = Boolean(properties.getProperty(SERIES_RESUME_STATE_KEY));
    var resumeState = seriesLoadResumeState_(fingerprint);
    var resumeOffset = 0;
    var continuing = false;

    if (resumeState && Number(resumeState.offset || 0) > 0) {
      resumeOffset = Number(resumeState.offset || 0);
      continuing = true;
    } else if (!hadSavedResume) {
      resumeOffset = 0;
      continuing = resumeOffset > 0;
    }

    var existingCandidates = continuing ? seriesReadExistingCandidates_() : [];
    var enrichment = seriesEnrichGroupsFromIsbn_(groups, key);

    if (!groups.length) {
      seriesClearResumeState_();
      seriesBuildResultSheet_([], {
        seriesCount: 0,
        searchedCount: 0,
        requestCount: enrichment.requestCount,
        truncated: false,
        progressDone: 0,
        progressTotal: 0,
        checkedAt: new Date()
      });
      spreadsheet.toast('권차가 표시된 시리즈를 찾지 못했습니다.', '점검 완료', 7);
      return;
    }

    spreadsheet.toast(
      (continuing ? '이전 점검 다음부터 ' : '') +
        groups.length + '개 시리즈의 중간 누락과 후속권을 확인합니다.',
      '시리즈 점검',
      8
    );

    var searchResult = seriesSearchYes24_(groups, holdings, key, resumeOffset, runStartedAt);
    // 시리즈 점검 멈추기: 멈추기를 눌렀으면 이번 결과는 쓰지 않고 이어가기도 하지 않습니다.
    if (typeof seriesStopRequested_ === 'function' && seriesStopRequested_(runStartedAt)) {
      seriesClearResumeState_();
      seriesCancelAutoContinue_();
      spreadsheet.toast('점검을 멈췄습니다. 이번 점검의 결과는 쓰지 않았습니다. 다음 점검은 처음부터 시작합니다.', '⏹ 시리즈 점검 멈춤', 8);
      return;
    }
    var candidates = seriesMergeCandidates_(existingCandidates, searchResult.candidates);
    // 3000개를 넘는 시리즈는 이번 점검 대상이 아닙니다. (예전에는 이것 때문에 다 끝나도 「미완료」로 남아 1분마다 영원히 반복됐습니다.)
    var incomplete = searchResult.truncated;

    if (incomplete) {
      seriesSaveResumeState_({
        fingerprint: fingerprint,
        offset: searchResult.nextOffset,
        total: searchResult.totalSpecCount,
        updatedAt: new Date().toISOString()
      });
      seriesScheduleAutoContinue_();
    } else {
      seriesClearResumeState_();
      seriesCancelAutoContinue_();
    }

    seriesBuildResultSheet_(candidates, {
      seriesCount: groups.length,
      searchedCount: searchResult.searchedGroupCount,
      requestCount: searchResult.requestCount + enrichment.requestCount,
      truncated: incomplete,
      progressDone: searchResult.nextOffset,
      progressTotal: searchResult.totalSpecCount,
      remainingCount: Math.max(0, searchResult.totalSpecCount - searchResult.nextOffset),
      checkedAt: new Date()
    });

    spreadsheet.setActiveSheet(spreadsheet.getSheetByName(SERIES_FINDER_SHEET));
    var percent = searchResult.totalSpecCount
      ? Math.min(100, searchResult.nextOffset / searchResult.totalSpecCount * 100).toFixed(1)
      : '100.0';
    spreadsheet.toast(
      '구매 검토 후보 ' + candidates.length + '권 · 검색 진행 ' +
        searchResult.nextOffset + ' / ' + searchResult.totalSpecCount + '건 (' + percent + '%)' +
        (incomplete ? ' 약 1분 뒤 남은 검색을 자동으로 이어갑니다.' : ' 전체 확인을 마쳤습니다.'),
      incomplete ? '시리즈 점검 저장 완료' : '시리즈 점검 전체 완료',
      10
    );
    if (searchResult.errorCount) {
      spreadsheet.toast(
        '일부 검색 ' + searchResult.errorCount + '건은 YES24 오류로 건너뛰었습니다. (예: ' + (searchResult.skippedQueries[0] || '') + ') 다음 점검 때 다시 시도됩니다.',
        '시리즈 점검 참고', 12
      );
    }
    if (groupLimitReached) {
      spreadsheet.toast(
        '소장도서의 시리즈가 ' + allGroups.length + '개로 한 번에 점검하는 ' + SERIES_MAX_GROUPS + '개를 넘어, 우선순위가 낮은 ' + (allGroups.length - groups.length) + '개는 점검하지 않았습니다.',
        '시리즈 점검 참고', 12
      );
    }
  } catch (error) {
    seriesSpreadsheet_().toast(
      String(error && error.message ? error.message : error),
      '시리즈 점검 오류',
      12
    );
    throw error;
  } finally {
    lock.releaseLock();
  }
}

function seriesGetYes24Key_() { return getConfiguredYes24Key_() || String(PropertiesService.getScriptProperties().getProperty('SERIES_YES24_KEY') || '').trim(); }

function seriesSpreadsheet_() {
  var active = SpreadsheetApp.getActive();
  if (active) return active;
  if (typeof HR_getSpreadsheet_ === 'function') return HR_getSpreadsheet_();
  throw new Error('연결된 스프레드시트를 찾을 수 없습니다.');
}

function seriesReadHoldings_() {
  var sheet = seriesSpreadsheet_().getSheetByName(SERIES_HOLDINGS_SHEET);
  if (!sheet) throw new Error('「' + SERIES_HOLDINGS_SHEET + '」 시트를 찾을 수 없습니다.');

  var lastRow = sheet.getLastRow();
  var lastColumn = sheet.getLastColumn();
  if (lastRow < 2 || lastColumn < 1) {
    return {rows: [], isbnSet: {}, titleKeySet: {}, parsedRows: []};
  }

  var headerSearchRows = Math.min(10, lastRow);
  var headerGrid = sheet.getRange(1, 1, headerSearchRows, lastColumn).getDisplayValues();
  var headerRow = 1;
  var titleColumn = 0;
  var isbnColumn = 0;
  var publisherColumn = 0;

  for (var r = 0; r < headerGrid.length; r++) {
    for (var c = 0; c < headerGrid[r].length; c++) {
      var header = String(headerGrid[r][c] || '').replace(/\s+/g, '').toUpperCase();
      if (!titleColumn && /^(도서명|책제목|서명|제목)$/.test(header)) titleColumn = c + 1;
      if (!isbnColumn && /^ISBN(13)?$/.test(header)) isbnColumn = c + 1;
      if (!publisherColumn && /^(출판사|발행자|출판자)$/.test(header)) publisherColumn = c + 1;
    }
    if (titleColumn && isbnColumn) {
      headerRow = r + 1;
      break;
    }
  }

  if (!titleColumn || !isbnColumn) {
    throw new Error('「' + SERIES_HOLDINGS_SHEET + '」 시트에서 도서명과 ISBN13 열을 찾지 못했습니다.');
  }

  var count = lastRow - headerRow;
  if (count < 1) return {rows: [], isbnSet: {}, titleKeySet: {}, parsedRows: []};
  var values = sheet.getRange(headerRow + 1, 1, count, lastColumn).getDisplayValues();
  var rows = [];
  var isbnSet = {};
  var titleKeySet = {};
  var parsedRows = [];

  values.forEach(function(row, index) {
    var title = String(row[titleColumn - 1] || '').trim();
    if (!title) return;
    var isbn = seriesCleanIsbn_(row[isbnColumn - 1]);
    var publisher = publisherColumn ? String(row[publisherColumn - 1] || '').trim() : '';
    if (isbn) isbnSet[isbn] = true;
    var titleKey = seriesNormalizeKey_(title);
    if (titleKey) titleKeySet[titleKey] = true;
    var parsed = seriesParseTitle_(title);
    if (parsed) {
      parsedRows.push({
        key: parsed.key,
        volume: parsed.volume,
        publisherKey: seriesNormalizePublisher_(publisher)
      });
    }
    rows.push({
      row: index + headerRow + 1,
      title: title,
      isbn: isbn,
      publisher: publisher
    });
  });
  return {rows: rows, isbnSet: isbnSet, titleKeySet: titleKeySet, parsedRows: parsedRows};
}

function seriesGroupHoldings_(rows) {
  var groupsByKey = {};

  // 같은 시리즈·판형에서 출판사가 적힌 줄이 하나라도 있으면, 출판사가 비어 있는 줄은 그 출판사들의 묶음에 함께 넣습니다.
  // (출판사를 일부 줄에만 적었을 때 같은 시리즈가 갈라져서, 갈라진 쪽마다 없는 권을 누락으로 잡던 문제를 막습니다.)
  var knownPublishers = {};
  var prepared = [];
  rows.forEach(function(book) {
    var parsed = seriesParseTitle_(book.title);
    if (!parsed) return;
    var publisherKey = seriesNormalizePublisher_(book.publisher) || '';
    var formatKey = seriesFormatKey_(book.title);
    prepared.push({ book: book, parsed: parsed, publisherKey: publisherKey, formatKey: formatKey });
    if (publisherKey) {
      var seriesFormat = parsed.key + '|' + formatKey;
      if (!knownPublishers[seriesFormat]) knownPublishers[seriesFormat] = {};
      knownPublishers[seriesFormat][publisherKey] = true;
    }
  });

  function addToGroup(book, parsed, publisherKey, formatKey) {
    var groupPublisher = publisherKey || '출판사미상';
    var signature = parsed.key + '|' + groupPublisher + '|' + formatKey;

    if (!groupsByKey[signature]) {
      groupsByKey[signature] = {
        key: signature,
        seriesKey: parsed.key,
        formatKey: formatKey,
        searchBase: parsed.base,
        aliasKeys: {},
        coreSearchBases: {},
        sampleIsbn: book.isbn || '',
        base: parsed.base,
        volumes: {},
        publishers: {},
        titles: [],
        explicitCount: 0
      };
    }

    var group = groupsByKey[signature];
    var coreParsed = seriesLeadingCore_(book.title);
    if (coreParsed) {
      group.aliasKeys[coreParsed.key] = true;
      group.coreSearchBases[coreParsed.base] = true;
    }
    group.volumes[parsed.volume] = true;
    group.titles.push(book.title);
    if (!group.sampleIsbn && book.isbn) group.sampleIsbn = book.isbn;
    if (parsed.explicit) group.explicitCount++;
    if (publisherKey) group.publishers[publisherKey] = true;
  }

  prepared.forEach(function(entry) {
    var known = knownPublishers[entry.parsed.key + '|' + entry.formatKey];
    if (!entry.publisherKey && known) {
      Object.keys(known).forEach(function(publisherKey) {
        addToGroup(entry.book, entry.parsed, publisherKey, entry.formatKey);
      });
      return;
    }
    addToGroup(entry.book, entry.parsed, entry.publisherKey, entry.formatKey);
  });

  var finalGroups = Object.keys(groupsByKey)
    .map(function(key) {
      var group = groupsByKey[key];
      var coreSearchBases = Object.keys(group.coreSearchBases || {}).sort(function(a, b) {
        return a.length - b.length;
      });
      if (coreSearchBases.length) {
        group.preferredSearchBase = coreSearchBases[0];
        group.searchBase = group.preferredSearchBase;
      }
      group.aliasKeys[group.seriesKey] = true;
      group.volumeList = Object.keys(group.volumes)
        .map(Number)
        .filter(function(value) { return value > 0 && value <= 200; })
        .sort(function(a, b) { return a - b; });
      group.maxVolume = group.volumeList.length
        ? Math.max.apply(null, group.volumeList)
        : 0;
      group.heldText = seriesCompressNumbers_(group.volumeList);
      group.missingVolumes = [];
      for (var volume = 1; volume < group.maxVolume; volume++) {
        if (!group.volumes[volume]) group.missingVolumes.push(volume);
      }
      return group;
    })
    .filter(function(group) {
      if (group.volumeList.length < 1 || group.base.length < 2) return false;
      if (group.explicitCount === 0 && group.volumeList.length === 1 && group.maxVolume > 80) {
        return false;
      }
      return true;
    });

  // 같은 시리즈·판형인데 출판사 표기가 달라(예: 「고릴라박스」와 「비룡소」) 따로 묶인 곳의 보유 권도 알아 둡니다.
  // 후보를 고를 때 「다른 묶음에 이미 있는 권」을 누락으로 잡지 않기 위해서입니다. (순환 참조를 피하려고 복사본만 담습니다)
  finalGroups.forEach(function(group) {
    group.siblingHeld = finalGroups
      .filter(function(other) {
        return other !== group && other.seriesKey === group.seriesKey && other.formatKey === group.formatKey;
      })
      .map(function(other) {
        return { publishers: other.publishers, volumes: other.volumes };
      });
  });
  return finalGroups;
}

function seriesEnrichGroupsFromIsbn_(groups, apiKey) {
  
  var properties = PropertiesService.getScriptProperties();
  var specs = [];
  var cacheWrites = {};

  groups.forEach(function(group) {
    if (!group.sampleIsbn) return;
    var cacheKey = 'YES24_SERIES_EDITION_CACHE_' + group.sampleIsbn;
    var cached = properties.getProperty(cacheKey);
    if (cached) {
      try {
        seriesApplyEditionInfo_(group, JSON.parse(cached));
        return;
      } catch (ignore) {}
    }
    specs.push({group: group, isbn: group.sampleIsbn, cacheKey: cacheKey});
  });

  // 판본 확인은 검색 정확도를 높이는 보조 단계입니다. 실제 후속권 검색이 먼저 끝나도록 일부만 확인합니다.
  specs = specs.slice(0, SERIES_MAX_ENRICH_REQUESTS);
  var requiredToday = specs.length;
  if (seriesApiRemaining_() < requiredToday) {
    throw new Error(
      '오늘 안전하게 사용할 수 있는 YES24 조회량이 부족합니다. 기존 결과는 그대로 보존했습니다. 내일 다시 점검해 주세요.'
    );
  }

  var requestCount = 0;
  for (var start = 0; start < specs.length; start += SERIES_SEARCH_CHUNK) {
    var chunk = specs.slice(start, start + SERIES_SEARCH_CHUNK);
    var requests = chunk.map(function(spec) {
      return {
        url: seriesBuildLookupUrl_(spec.isbn, apiKey),
        method: 'get',
        headers: {'X-Api-Key': apiKey},
        muteHttpExceptions: true,
        followRedirects: true
      };
    });
    var responses = yes24FetchAll_(requests);
    requestCount += responses.length;
    seriesRecordApiUsage_(responses.length);

    responses.forEach(function(response, index) {
      if (response.getResponseCode() !== 200) return;
      var data;
      try {
        data = JSON.parse(response.getContentText('UTF-8'));
      } catch (ignore) {
        return;
      }
      var item = data && data.data && Array.isArray(data.data.items) ? data.data.items[0] : null;
      if (!item) return;

      var info = {
        title: seriesCleanText_(item.title),
        publisher: seriesCleanText_(item.publisher)
      };
      seriesApplyEditionInfo_(chunk[index].group, info);
      cacheWrites[chunk[index].cacheKey] = JSON.stringify(info);
    });

    // 판본 기억 저장 공간이 가득 차도 점검은 계속합니다. (기억은 속도를 위한 보조 기능입니다.)
    if (Object.keys(cacheWrites).length) {
      try { properties.setProperties(cacheWrites, false); } catch (ignore) { cacheWrites = {}; }
    }

    Utilities.sleep(100);
  }

  if (Object.keys(cacheWrites).length) {
    try { properties.setProperties(cacheWrites, false); } catch (ignore) {}
  }
  return {requestCount: requestCount};
}

// 두 시리즈 이름(정규화한 글)이 사실상 같은 시리즈인지: 글자 두 개씩 겹치는 비율(Dice)이 0.6 이상이거나,
// 한쪽이 다른 쪽을 품고 길이 비율이 0.6 이상이면 같은 것으로 봅니다.
function seriesNameKeysSimilar_(left, right) {
  var a = String(left || ''), b = String(right || '');
  if (!a || !b) return false;
  if (a === b) return true;
  var shorter = a.length <= b.length ? a : b;
  var longer = a.length > b.length ? a : b;
  if (longer.indexOf(shorter) >= 0 && shorter.length / longer.length >= 0.6) return true;
  function bigrams(text) {
    var list = [];
    for (var i = 0; i < text.length - 1; i++) list.push(text.substr(i, 2));
    return list;
  }
  var x = bigrams(a), y = bigrams(b);
  if (!x.length || !y.length) return false;
  var pool = {};
  y.forEach(function(gram) { pool[gram] = (pool[gram] || 0) + 1; });
  var common = 0;
  x.forEach(function(gram) { if (pool[gram] > 0) { pool[gram]--; common++; } });
  return (2 * common) / (x.length + y.length) >= 0.6;
}

function seriesApplyEditionInfo_(group, info) {
  var apiTitle = seriesCleanText_(info && info.title);
  var parsed = seriesParseTitle_(apiTitle);
  // YES24가 그 ISBN을 학교 소장도서와 전혀 다른 이름으로 돌려주면(예: 소장 「처음 읽는 삼국지」 ↔ YES24 「삼국지 1」)
  // 그 정보를 믿지 않고 무시합니다. 믿으면 짧고 흔한 이름(「삼국지」)이 별칭이 되어 이름이 다른 책이 후보로 걸립니다.
  if (parsed && !seriesNameKeysSimilar_(parsed.key, group.seriesKey)) return;
  if (parsed) {
    group.aliasKeys[parsed.key] = true;
    group.searchBase = group.preferredSearchBase || parsed.base || group.searchBase;
    group.formatKey = seriesFormatKey_(apiTitle);
  }
  var publisherKey = seriesNormalizePublisher_(info && info.publisher);
  if (publisherKey) group.publishers[publisherKey] = true;
}

function seriesApiUsageState_() {
  var properties = PropertiesService.getScriptProperties();
  var zone = Session.getScriptTimeZone() || 'Asia/Seoul';
  var today = Utilities.formatDate(new Date(), zone, 'yyyy-MM-dd');
  var savedDate = properties.getProperty('YES24_SERIES_API_USAGE_DATE') || '';
  var used = Number(properties.getProperty('YES24_SERIES_API_USAGE_COUNT') || 0);

  if (savedDate !== today) {
    used = 0;
    properties.setProperties({
      YES24_SERIES_API_USAGE_DATE: today,
      YES24_SERIES_API_USAGE_COUNT: '0'
    }, false);
  }
  return {date: today, used: used};
}

function seriesApiRemaining_() {
  var state = seriesApiUsageState_();
  return Math.max(0, SERIES_DAILY_API_SAFE_LIMIT - state.used);
}

function seriesRecordApiUsage_(count) {
  var properties = PropertiesService.getScriptProperties();
  var state = seriesApiUsageState_();
  properties.setProperty(
    'YES24_SERIES_API_USAGE_COUNT',
    String(state.used + Math.max(0, Number(count || 0)))
  );
}

function seriesBuildLookupUrl_(isbn, apiKey) { return yes24Url_('/goods/itemDetail', {query: yes24Isbn13_(isbn), searchType: 'ISBN13', detail: 'Y'}); }

function seriesGroupPriority_(a, b) {
  var aMulti = a.volumeList.length >= 2 ? 1 : 0;
  var bMulti = b.volumeList.length >= 2 ? 1 : 0;
  if (aMulti !== bMulti) return bMulti - aMulti;
  if (a.explicitCount !== b.explicitCount) return b.explicitCount - a.explicitCount;
  if (a.maxVolume !== b.maxVolume) return b.maxVolume - a.maxVolume;
  return a.base.localeCompare(b.base, 'ko');
}

function seriesResumeFingerprint_(groups) {
  var source = groups.map(function(group) {
    return group.key + '|' + Object.keys(group.volumes || {})
      .map(Number)
      .sort(function(a, b) { return a - b; })
      .join(',');
  }).sort().join('\n');

  var digest = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    source,
    Utilities.Charset.UTF_8
  );
  return digest.map(function(value) {
    var normalized = (Number(value) + 256) % 256;
    return ('0' + normalized.toString(16)).slice(-2);
  }).join('');
}

function seriesLoadResumeState_(fingerprint) {
  var properties = PropertiesService.getScriptProperties();
  var raw = properties.getProperty(SERIES_RESUME_STATE_KEY);
  if (!raw) return null;

  try {
    var state = JSON.parse(raw);
    if (!state || state.fingerprint !== fingerprint) {
      properties.deleteProperty(SERIES_RESUME_STATE_KEY);
      return null;
    }
    return state;
  } catch (ignore) {
    properties.deleteProperty(SERIES_RESUME_STATE_KEY);
    return null;
  }
}

function seriesSaveResumeState_(state) {
  PropertiesService.getScriptProperties().setProperty(
    SERIES_RESUME_STATE_KEY,
    JSON.stringify(state || {})
  );
}

function seriesClearResumeState_() {
  PropertiesService.getScriptProperties().deleteProperty(SERIES_RESUME_STATE_KEY);
}

function resetSeriesFinderProgress() {
  seriesCancelAutoContinue_();
  seriesClearResumeState_();
  seriesSpreadsheet_().toast(
    '저장된 점검 위치를 지웠습니다. 다음 점검은 처음부터 시작합니다.',
    '시리즈 점검 초기화',
    8
  );
}

function seriesAutoContinueMissingBooks_() {
  refreshMissingSeriesBooks();
}

function seriesCancelAutoContinue_() {
  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    if (trigger.getHandlerFunction() === SERIES_AUTO_CONTINUE_HANDLER) {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}

function seriesScheduleAutoContinue_() {
  seriesCancelAutoContinue_();
  ScriptApp.newTrigger(SERIES_AUTO_CONTINUE_HANDLER)
    .timeBased()
    .after(SERIES_AUTO_CONTINUE_DELAY_MS)
    .create();
}

function seriesLegacyResumeOffset_() {
  var sheet = seriesSpreadsheet_().getSheetByName(SERIES_FINDER_SHEET);
  if (!sheet) return 0;
  var summary = String(sheet.getRange('A3').getDisplayValue() || '');
  return summary.indexOf('일부 미확인') >= 0 ? SERIES_MAX_API_REQUESTS : 0;
}

function seriesReadExistingCandidates_() {
  var sheet = seriesSpreadsheet_().getSheetByName(SERIES_FINDER_SHEET);
  if (!sheet || sheet.getLastRow() < 6) return [];

  var rowCount = sheet.getLastRow() - 5;
  var range = sheet.getRange(6, 1, rowCount, 14);
  var values = range.getValues();
  var formulas = range.getFormulas();
  var timezone = Session.getScriptTimeZone();

  return values.map(function(row, index) {
    var title = seriesCleanText_(row[6]);
    var seriesName = seriesCleanText_(row[2]);
    var volumeMatch = String(row[4] || '').match(/(\d{1,3})/);
    if (!title || !seriesName || !volumeMatch) return null;

    var pubDate = row[9];
    if (pubDate instanceof Date && !isNaN(pubDate.getTime())) {
      pubDate = Utilities.formatDate(pubDate, timezone, 'yyyy-MM-dd');
    } else {
      pubDate = seriesCleanText_(pubDate);
    }

    return {
      selected: row[0] === true,
      type: seriesCleanText_(row[1]),
      series: seriesName,
      held: seriesCleanText_(row[3]),
      volume: Number(volumeMatch[1]),
      cover: seriesExtractUrlFromFormula_(formulas[index][5]),
      title: title,
      author: seriesCleanText_(row[7]),
      publisher: seriesCleanText_(row[8]),
      pubDate: pubDate,
      isbn: seriesCleanIsbn_(row[10]),
      price: Number(row[11] || 0),
      link: seriesExtractUrlFromFormula_(formulas[index][12]),
      saleStatus: /^절판(?:\s*·|$)/.test(String(row[13] || '')) ? '절판' : '',
      publisherMatched: false,
      exactKey: false
    };
  }).filter(function(book) { return Boolean(book); });
}

function seriesExtractUrlFromFormula_(formula) {
  var match = String(formula || '').match(/"(https?:\/\/[^"]+)"/i);
  return match && match[1] ? match[1].replace(/""/g, '"') : '';
}

function seriesMergeCandidates_(existing, incoming) {
  var byKey = {};

  function put(book) {
    if (!book) return;
    var key = book.isbn
      ? 'isbn|' + book.isbn
      : 'series|' + seriesNormalizeKey_(book.series) + '|' +
        Number(book.volume || 0) + '|' + seriesNormalizePublisher_(book.publisher);
    var previous = byKey[key];
    if (!previous) {
      byKey[key] = book;
      return;
    }

    var selected = Boolean(previous.selected || book.selected);
    if (seriesCandidateScore_(book) >= seriesCandidateScore_(previous)) {
      byKey[key] = book;
    }
    byKey[key].selected = selected;
  }

  (existing || []).forEach(put);
  (incoming || []).forEach(put);

  return Object.keys(byKey)
    .map(function(key) { return byKey[key]; })
    .sort(function(a, b) {
      var typeOrder = {'중간 누락': 0, '신간 후속권': 1};
      if (typeOrder[a.type] !== typeOrder[b.type]) return typeOrder[a.type] - typeOrder[b.type];
      var baseOrder = String(a.series || '').localeCompare(String(b.series || ''), 'ko');
      return baseOrder || Number(a.volume || 0) - Number(b.volume || 0);
    });
}

function seriesSearchYes24_(groups, holdings, apiKey, resumeOffset, runStartedAt) {
  var followupSpecs = groups.map(function(group) {
    return {
      group: group,
      kind: 'followup',
      query: group.searchBase || group.base,
      page: 1
    };
  });

  var gapGroups = groups.filter(function(group) {
    return group.missingVolumes && group.missingVolumes.length;
  });
  var gapSpecs = [];
  var round = 0;
  var added = true;

  // 모든 중간 누락 권차를 한 권씩 번갈아 대기열에 넣습니다.
  while (added) {
    added = false;
    for (var groupIndex = 0; groupIndex < gapGroups.length; groupIndex++) {
      var gapGroup = gapGroups[groupIndex];
      if (round >= gapGroup.missingVolumes.length) continue;

      var targetVolume = gapGroup.missingVolumes[round];
      gapSpecs.push({
        group: gapGroup,
        kind: 'gap',
        query: (gapGroup.searchBase || gapGroup.base) + ' ' + targetVolume,
        targetVolume: targetVolume,
        page: 1
      });
      added = true;
    }
    round++;
  }

  // 기존 1페이지 작업 뒤에 추가 페이지를 붙여 저장된 진행 위치를 유지합니다.
  var extraSpecs = [];
  gapSpecs.forEach(function(spec) {
    for (var page = 2; page <= SERIES_MAX_GAP_PAGES; page++) {
      extraSpecs.push({group: spec.group, kind: spec.kind, query: spec.query,
        targetVolume: spec.targetVolume, page: page});
    }
  });
  followupSpecs.forEach(function(spec) {
    extraSpecs.push({group: spec.group, kind: spec.kind, query: spec.query, page: 2});
  });
  var allSpecs = followupSpecs.concat(gapSpecs, extraSpecs);
  var totalSpecCount = allSpecs.length;
  var startOffset = Math.max(0, Math.min(Number(resumeOffset || 0), totalSpecCount));
  var remainingJobs = Math.max(0, totalSpecCount - startOffset);
  var availableToday = Math.max(0, seriesApiRemaining_());
  var batchLimit = Math.min(SERIES_MAX_API_REQUESTS, availableToday, remainingJobs);

  if (remainingJobs > 0 && batchLimit < 1) {
    throw new Error(
      '오늘 안전하게 사용할 수 있는 YES24 조회량을 모두 사용했습니다. 기존 결과와 진행 위치는 그대로 보존했습니다. 내일 다시 점검해 주세요.'
    );
  }

  var specs = allSpecs.slice(startOffset, startOffset + batchLimit);
  var responses = [];
  var searchedGroups = {};

  for (var start = 0; start < specs.length; start += SERIES_SEARCH_CHUNK) {
    if (typeof seriesStopRequested_ === 'function' && seriesStopRequested_(runStartedAt)) break;
    if (start > 0 && Date.now() - Number(runStartedAt || 0) >= SERIES_EXECUTION_BUDGET_MS) break;
    var chunk = specs.slice(start, start + SERIES_SEARCH_CHUNK);
    var requests = chunk.map(function(spec) {
      searchedGroups[spec.group.key] = true;
      return {
        url: seriesBuildSearchUrl_(spec.query, apiKey, spec.kind, spec.page),
        method: 'get',
        headers: {'X-Api-Key': apiKey},
        muteHttpExceptions: true,
        followRedirects: true
      };
    });

    var batch = yes24FetchAll_(requests);
    seriesRecordApiUsage_(batch.length);
    batch.forEach(function(response, index) {
      responses.push({spec: chunk[index], response: response});
    });

    seriesSpreadsheet_().toast(
      Math.min(startOffset + start + chunk.length, totalSpecCount) +
        ' / ' + totalSpecCount + '건 확인 중',
      'YES24 검색',
      4
    );
    Utilities.sleep(180);
  }

  var bySeriesVolume = {};
  var apiErrors = [];     // 키·한도 문제처럼 전체를 멈춰야 하는 오류
  var skippedQueries = [];  // 그 검색 한 건만 건너뛰어도 되는 오류

  responses.forEach(function(entry) {
    var response = entry.response;
    var spec = entry.spec;
    var status = response.getResponseCode();
    if (status !== 200) {
      // 인증·한도 오류(401·403·429)는 전부 멈추고, 그 밖의 오류(400·404·5xx)는 그 검색만 건너뜁니다.
      // (예전에는 한 건만 실패해도 점검 전체가 같은 자리에서 영원히 멈췄습니다.)
      (/^(401|403|429)$/.test(String(status)) ? apiErrors : skippedQueries).push(spec.query + ': HTTP ' + status);
      return;
    }

    var data;
    try {
      data = JSON.parse(response.getContentText('UTF-8'));
    } catch (ignore) {
      skippedQueries.push(spec.query + ': 응답 해석 실패');
      return;
    }
    if (data.errorCode) {
      apiErrors.push(spec.query + ': ' + (data.errorMessage || data.errorCode));
      return;
    }

    var items = data.data && Array.isArray(data.data.items) ? data.data.items : [];
    items.forEach(function(item) {
      var candidate = seriesCandidateFromItem_(spec.group, item, holdings);
      if (!candidate) return;
      if (spec.targetVolume && candidate.volume !== spec.targetVolume) return;

      var uniqueKey = spec.group.key + '|' + candidate.volume;
      var previous = bySeriesVolume[uniqueKey];
      if (!previous || seriesCandidateScore_(candidate) > seriesCandidateScore_(previous)) {
        bySeriesVolume[uniqueKey] = candidate;
      }
    });
  });

  // 절반 이상(3건 이상)이 실패했다면 키·서버 문제이므로 멈추고, 몇 건만 실패했으면 건너뛰고 계속합니다.
  if (apiErrors.length > 0 || (skippedQueries.length >= 3 && skippedQueries.length * 2 >= responses.length) || (responses.length > 0 && skippedQueries.length === responses.length)) {
    throw new Error('YES24 조회 중 오류가 발생해 결과와 진행 위치를 보존했습니다. API 키와 호출 한도를 확인해 주세요.\n' + apiErrors.concat(skippedQueries).slice(0, 3).join('\n'));
  }

  var candidates = Object.keys(bySeriesVolume)
    .map(function(key) { return bySeriesVolume[key]; })
    .sort(function(a, b) {
      var typeOrder = {'중간 누락': 0, '신간 후속권': 1};
      if (typeOrder[a.type] !== typeOrder[b.type]) return typeOrder[a.type] - typeOrder[b.type];
      var baseOrder = a.series.localeCompare(b.series, 'ko');
      return baseOrder || a.volume - b.volume;
    });

  var nextOffset = startOffset + responses.length;
  return {
    candidates: candidates,
    requestCount: responses.length,
    searchedGroupCount: Object.keys(searchedGroups).length,
    truncated: nextOffset < totalSpecCount,
    errorCount: skippedQueries.length,
    skippedQueries: skippedQueries.slice(0, 5),
    startOffset: startOffset,
    nextOffset: nextOffset,
    totalSpecCount: totalSpecCount
  };
}

function seriesBuildSearchUrl_(query, apiKey, kind, page) { return yes24Url_('/goods/itemList', {query: query, category: 'BOOK', sort: kind === 'gap' ? 'RELATION' : 'RECENT', page: Math.max(1, Number(page || 1)), pageSize: kind === 'gap' ? SERIES_GAP_RESULTS : SERIES_FOLLOWUP_RESULTS, detail: 'Y'}); }

function seriesCandidateFromItem_(group, item, holdings) {
  var title = seriesCleanText_(item.title);
  if (!title || seriesExcludedTitle_(title)) return null;

  var parsed = seriesParseTitle_(title);
  if (!parsed) return null;

  var publisherKeys = Object.keys(group.publishers || {});
  var publisherMatched = publisherKeys.length === 0 ||
    seriesPublisherMatches_(group.publishers, item.publisher);
  var exactKey = Boolean(group.aliasKeys && group.aliasKeys[parsed.key]);
  var compatibleKey = exactKey;

  if (!compatibleKey && publisherMatched) {
    var aliasKeys = Object.keys(group.aliasKeys || {});
    compatibleKey = aliasKeys.some(function(aliasKey) {
      return seriesKeysCompatible_(parsed.key, aliasKey, true);
    });
  }

  var candidateFormatKey = seriesFormatKey_(title);
  var sameFormat = candidateFormatKey === group.formatKey;

  // ISBN으로 확인한 학교 보유 판본과 같은 시리즈·출판사·판형만 인정합니다.
  if (!compatibleKey || !publisherMatched || !sameFormat) return null;
  if (group.volumes[parsed.volume]) return null;
  // 출판사 표기가 달라 따로 묶인 같은 시리즈의 묶음에 이미 있는 권이고, 그 묶음의 출판사가 이 책의 출판사와 맞으면 가진 책입니다.
  if ((group.siblingHeld || []).some(function(other) {
    return other.volumes[parsed.volume] && seriesPublisherMatches_(other.publishers, item.publisher);
  })) return null;

  var isbn = seriesCleanIsbn_(item.isbn13 || item.isbn10);
  if (!isbn || holdings.isbnSet[isbn]) return null;
  if (seriesHoldingMatchesCandidate_(title, parsed, item.publisher, holdings)) return null;

  var cover = String(item.cover || '')
    
    ;

  return {
    type: parsed.volume > group.maxVolume ? '신간 후속권' : '중간 누락',
    series: group.base,
    seriesKey: parsed.key,
    formatKey: candidateFormatKey,
    publisherKey: seriesNormalizePublisher_(item.publisher),
    held: group.heldText,
    volume: parsed.volume,
    cover: /^https?:\/\//i.test(cover) ? cover : '',
    title: title,
    author: seriesCleanText_(item.author),
    publisher: seriesCleanText_(item.publisher),
    pubDate: seriesCleanText_(item.publishDate),
    isbn: isbn,
    price: Number(item.shopPrice || item.salePrice || 0),
    link: /^https?:\/\//i.test(String(item.link || '')) ? String(item.link) : '',
    saleStatus: /절판/.test(String(item.itemStatus || '')) ? '절판' : '',
    publisherMatched: publisherMatched,
    exactKey: exactKey,
    variantLabel: ''
  };
}

function seriesHoldingMatchesCandidate_(title, parsed, publisher, holdings) {
  // 조회 범위를 줄이지 않습니다. 보유 여부는 정확한 ISBN과 현재 시리즈 권차로만 판단합니다.
  return false;
}

function seriesCandidateScore_(candidate) {
  var score = candidate.publisherMatched ? 1000 : 0;
  score += candidate.exactKey ? 500 : 0;
  score += /^\d{13}$/.test(candidate.isbn) ? 100 : 0;
  score += candidate.cover ? 10 : 0;
  score += candidate.pubDate ? Number(String(candidate.pubDate).replace(/[^0-9]/g, '').slice(0, 8)) || 0 : 0;
  return score;
}

function seriesParseTitle_(title) {
  var clean = seriesCleanText_(title);
  if (!clean) return null;

  clean = seriesReplaceCircledNumerals_(clean);
  clean = seriesReplaceRomanNumerals_(clean);
  clean = clean
    .replace(/[\(\[【]\s*(\d{1,3})\s*[\)\]】]/g, ' $1권 ')
    .replace(/[\(\[【]\s*(개정판|증보판|특별판|양장본?|반양장|리커버|에디션)\s*[\)\]】]/gi, ' ')
    // 괄호 속 시리즈명은 버리지 않고 괄호 기호만 제거합니다.
    .replace(/[\(\)\[\]【】]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();

  var matches = [];
  var current;

  var englishPattern = /(?:vol(?:ume)?|no|season|part|level)\.?\s*(\d{1,3})(?=\s|[.:,_\-]|$)/ig;
  while ((current = englishPattern.exec(clean)) !== null) {
    matches.push({
      volume: Number(current[1]),
      numberIndex: current.index + current[0].lastIndexOf(current[1]),
      tokenIndex: current.index,
      explicit: true
    });
  }

  var explicitPattern = /(?:제\s*)?(\d{1,3})\s*(권|편|호|번째|탄|부|집)(?=\s|[.:,_\-]|$)/g;
  while ((current = explicitPattern.exec(clean)) !== null) {
    matches.push({
      volume: Number(current[1]),
      numberIndex: current.index + current[0].lastIndexOf(current[1]),
      tokenIndex: current.index,
      explicit: true
    });
  }

  if (!matches.length) {
    var standalonePattern = /(?:^|\s|[.:,_\-])(\d{1,3})(?=\s|[.:,_\-]|$)/g;
    while ((current = standalonePattern.exec(clean)) !== null) {
      var possible = Number(current[1]);
      var index = current.index + current[0].lastIndexOf(current[1]);
      if (possible > 0 && possible <= 200 && index >= 2) {
        matches.push({volume: possible, numberIndex: index, tokenIndex: index, explicit: false});
      }
    }
  }

  matches = matches.filter(function(match) {
    return match.volume > 0 && match.volume <= 200;
  });
  if (!matches.length) return null;

  matches.sort(function(a, b) {
    if (a.explicit !== b.explicit) return a.explicit ? 1 : -1;
    return a.numberIndex - b.numberIndex;
  });
  var chosen = matches[matches.length - 1];
  var baseEnd = chosen.explicit && chosen.tokenIndex < chosen.numberIndex
    ? chosen.tokenIndex
    : chosen.numberIndex;
  var base = clean.slice(0, baseEnd)
    .replace(/(?:제\s*)$/, '')
    .replace(/(?:vol(?:ume)?|no|season|part|level)\.?\s*$/i, '')
    .replace(/[\s.:,_\-]+$/, '')
    .trim();
  if (base.length < 2) return null;

  var key = seriesNormalizeKey_(base);
  if (key.length < 2) return null;
  return {base: base, key: key, volume: chosen.volume, explicit: chosen.explicit};
}

function seriesReplaceCircledNumerals_(value) {
  var chars = '①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳';
  return String(value || '').replace(/[①-⑳]/g, function(character) {
    return ' ' + (chars.indexOf(character) + 1) + '권 ';
  });
}

function seriesReplaceRomanNumerals_(value) {
  var romanMap = {
    'ⅩⅩ': 20, 'ⅩⅨ': 19, 'ⅩⅧ': 18, 'ⅩⅦ': 17, 'ⅩⅥ': 16,
    'ⅩⅤ': 15, 'ⅩⅣ': 14, 'ⅩⅢ': 13, 'ⅩⅡ': 12, 'ⅩⅠ': 11,
    'Ⅹ': 10, 'Ⅸ': 9, 'Ⅷ': 8, 'Ⅶ': 7, 'Ⅵ': 6,
    'Ⅴ': 5, 'Ⅳ': 4, 'Ⅲ': 3, 'Ⅱ': 2, 'Ⅰ': 1
  };
  return String(value || '').replace(/[ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩ]+/g, function(token) {
    return romanMap[token] ? ' ' + romanMap[token] + '권 ' : token;
  });
}

function seriesKeysCompatible_(left, right, publisherMatched) {
  var a = String(left || '');
  var b = String(right || '');
  if (!a || !b) return false;
  if (a === b) return true;

  var shorter = a.length <= b.length ? a : b;
  var longer = a.length > b.length ? a : b;
  if (shorter.length < 5 || longer.indexOf(shorter) === -1) return false;

  var ratio = shorter.length / longer.length;
  return ratio >= 0.88 || (publisherMatched && ratio >= 0.72);
}

function seriesNormalizeKey_(value) {
  return String(value || '')
    .toLowerCase()
    .replace(/[^0-9a-z가-힣]/g, '');
}

function seriesNormalizePublisher_(value) {
  return seriesNormalizeKey_(String(value || '')
    .replace(/주식회사|도서출판|출판사|\(주\)|㈜/g, ''));
}

function seriesPublisherMatches_(publisherMap, candidatePublisher) {
  var keys = Object.keys(publisherMap || {});
  if (!keys.length) return false;
  var candidate = seriesNormalizePublisher_(candidatePublisher);
  if (!candidate) return false;
  return keys.some(function(key) {
    return key === candidate ||
      (key.length >= 3 && candidate.indexOf(key) >= 0) ||
      (candidate.length >= 3 && key.indexOf(candidate) >= 0);
  });
}

function seriesExcludedTitle_(title) {
  return /세트|전집|박스|워크북|스티커북|색칠|오디오북|필사노트/i.test(String(title || ''));
}

function seriesCleanIsbn_(value) {
  var digits = String(value || '').replace(/[^0-9Xx]/g, '');
  return digits.length === 10 || digits.length === 13 ? digits.toUpperCase() : '';
}

function seriesCleanText_(value) {
  return String(value || '')
    .replace(/<[^>]*>/g, '')
    .replace(/&nbsp;/g, ' ')
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
    .replace(/\s+/g, ' ')
    .trim();
}

function seriesCompressNumbers_(numbers) {
  if (!numbers.length) return '';
  var parts = [];
  var start = numbers[0];
  var end = numbers[0];
  for (var index = 1; index <= numbers.length; index++) {
    var current = numbers[index];
    if (current === end + 1) {
      end = current;
      continue;
    }
    parts.push(start === end ? String(start) : start + '~' + end);
    start = current;
    end = current;
  }
  return parts.join(', ') + '권';
}

/** 현재 표시된 후보의 판매 상태를 YES24에서 다시 확인합니다. */
function refreshSeriesSaleStatuses() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) {
    seriesSpreadsheet_().toast('시리즈 점검이 진행 중입니다. 완료 후 다시 시도해 주세요.', '절판 확인', 7);
    return;
  }
  try {
    seriesCancelAutoContinue_();
    var sheet = seriesSpreadsheet_().getSheetByName(SERIES_FINDER_SHEET);
    if (!sheet || sheet.getLastRow() < 6) return;
    var key = seriesGetYes24Key_();
    if (!key) throw new Error('YES24 API 키가 없습니다.');
    var count = sheet.getLastRow() - 5;
    var rows = sheet.getRange(6, 1, count, 14).getDisplayValues();
    var jobs = rows.map(function(row, index) {
      var isbn = seriesCleanIsbn_(row[10]);
      return isbn ? {index: index, isbn: isbn} : null;
    }).filter(Boolean);
    if (seriesApiRemaining_() < jobs.length) throw new Error('오늘 남은 YES24 API 호출량이 부족합니다. 내일 다시 확인해 주세요.');
    var changed = 0;
    var startTime = Date.now();
    for (var start = 0; start < jobs.length; start += SERIES_SEARCH_CHUNK) {
      if (Date.now() - startTime >= SERIES_EXECUTION_BUDGET_MS) break;
      var chunk = jobs.slice(start, start + SERIES_SEARCH_CHUNK);
      var requests = chunk.map(function(job) {
        return {url: yes24Url_('/goods/itemDetail', {searchType: 'ISBN13', query: job.isbn, detail: 'N'}),
          method: 'get', headers: {'X-Api-Key': key}, muteHttpExceptions: true, followRedirects: true};
      });
      var responses = yes24FetchAll_(requests);
      seriesRecordApiUsage_(responses.length);
      responses.forEach(function(response, i) {
        if (response.getResponseCode() !== 200) return;
        var data;
        try { data = JSON.parse(response.getContentText('UTF-8')); } catch (ignore) { return; }
        var items = data && data.data && data.data.items;
        if (!data.success || !Array.isArray(items)) return;
        var book = items.filter(function(item) { return seriesCleanIsbn_(item.isbn13 || item.isbn10) === chunk[i].isbn; })[0];
        if (!book) return;
        var rowNumber = chunk[i].index + 6;
        var noteCell = sheet.getRange(rowNumber, 14);
        var note = String(noteCell.getValue() || '').replace(/^절판\s*·\s*/, '');
        var updatedNote = (/절판/.test(String(book.itemStatus || '')) ? '절판 · ' : '') + note;
        if (updatedNote !== String(noteCell.getValue() || '')) {
          noteCell.setValue(updatedNote);
          changed++;
        }
      });
      Utilities.sleep(180);
    }
    seriesApplyApiGauge_(sheet, sheet.getRange('A3').getDisplayValue(), sheet.getRange('E3').getDisplayValue());
    seriesSpreadsheet_().toast('절판 여부 확인 완료 · 표시 변경 ' + changed + '권', '시리즈 점검', 8);
  } finally {
    if (PropertiesService.getScriptProperties().getProperty(SERIES_RESUME_STATE_KEY)) {
      seriesScheduleAutoContinue_();
    }
    lock.releaseLock();
  }
}

function seriesBuildResultSheet_(candidates, meta) {
  var spreadsheet = seriesSpreadsheet_();
  var sheet = spreadsheet.getSheetByName(SERIES_FINDER_SHEET) || spreadsheet.insertSheet(SERIES_FINDER_SHEET);

  if (sheet.getFilter()) sheet.getFilter().remove();
  sheet.getRange(1, 1, Math.min(sheet.getMaxRows(), 8), Math.min(sheet.getMaxColumns(), 14)).breakApart();
  sheet.clear();
  sheet.clearConditionalFormatRules();
  sheet.setHiddenGridlines(true);

  var headers = [
    '선택', '구분', '시리즈명', '현재 보유 권차', '찾은 권차', '표지', '도서명',
    '저자', '출판사', '출간일', 'ISBN13', '정가', 'YES24', '점검 메모'
  ];

  sheet.getRange('A1:N1').merge().setValue('🔎 우리 학교 시리즈 빈칸 찾기');
  sheet.getRange('A2:N2').merge().setValue(
    '보유목록과 같은 시리즈·출판사·판형을 기준으로 빠진 권차와 새 후속권만 찾습니다.'
  );

  var checkedAt = Utilities.formatDate(meta.checkedAt || new Date(), Session.getScriptTimeZone(), 'yyyy. M. d. HH:mm');
  var progressTotal = Number(meta.progressTotal || 0);
  var progressDone = Math.min(progressTotal, Number(meta.progressDone || 0));
  var progressPercent = progressTotal ? (progressDone / progressTotal * 100).toFixed(1) : '100.0';
  var progressText = progressTotal
    ? '  ·  검색 진행 ' + progressDone + '/' + progressTotal + '건 (' + progressPercent + '%)' +
      (meta.truncated ? '  ·  남은 검색 ' + Math.max(0, progressTotal - progressDone) + '건' : '')
    : '';
  var summaryText =
    '확인 시리즈 ' + Number(meta.seriesCount || 0) + '개  ·  구매 검토 후보 ' +
    candidates.length + '권' + progressText +
    (meta.truncated ? '  ·  ⚠ 이어서 점검 필요' : '  ·  ✅ 전체 확인 완료');
  var lastCheckedText = '마지막 확인\n' + checkedAt;
  sheet.getRange(5, 1, 1, headers.length).setValues([headers]);

  if (candidates.length) {
    var rows = candidates.map(function(book) {
      var coverFormula = book.cover
        ? '=IFERROR(IMAGE("' + seriesEscapeFormula_(book.cover) + '",4,105,72),"")'
        : '';
      var linkFormula = book.link
        ? '=HYPERLINK("' + seriesEscapeFormula_(book.link) + '","YES24에서 보기")'
        : '';
      return [
        Boolean(book.selected),
        book.type,
        book.series,
        book.held,
        book.volume + '권',
        coverFormula,
        book.title,
        book.author,
        book.publisher,
        book.pubDate,
        book.isbn,
        book.price || '',
        linkFormula,
        unifiedComposeNote_(
          (book.saleStatus === '절판' ? '절판 · ' : '') +
            (book.type === '신간 후속권'
              ? '현재 마지막 권 다음에 나온 책'
              : '보유 권차 사이에서 빠진 책'),
          unifiedBookFlags_(book.title, '')
        )
      ];
    });
    sheet.getRange(6, 1, rows.length, headers.length).setValues(rows);
    sheet.getRange(6, 1, rows.length, 1).insertCheckboxes();
    sheet.getRange(6, 1, rows.length, headers.length).setVerticalAlignment('middle').setWrap(true);
    sheet.getRange(6, 12, rows.length, 1).setNumberFormat('#,##0"원"');
    sheet.setRowHeights(6, rows.length, 112);
    sheet.getRange(5, 1, rows.length + 1, headers.length).createFilter();
  } else {
    sheet.getRange('A6:N7').merge().setValue(
      Number(meta.searchedCount || 0) > 0
        ? '현재 YES24 검색 결과에서 추가 구매 후보를 찾지 못했습니다.'
        : '위쪽의 「🔎 시리즈 누락 점검 시작 ›」 막대를 누르면 결과가 채워집니다.'
    );
    sheet.getRange('A6:N7').setHorizontalAlignment('center').setVerticalAlignment('middle');
  }

  seriesStyleResultSheet_(sheet, candidates.length);
  seriesApplyApiGauge_(sheet, summaryText, lastCheckedText);
  seriesAddWorkButton_(sheet);
}


/**
 * 시리즈 점검 상단에 API 사용량 막대를 표시합니다.
 * 표시를 새로고침하는 과정에서는 YES24 API를 호출하지 않습니다.
 */
function seriesApplyApiGauge_(sheet, summaryText, lastCheckedText) {
  var apiUsage = seriesApiUsageState_();
  var apiUsed = Math.max(0, Number(apiUsage.used || 0));
  var apiRemaining = Math.max(0, SERIES_DAILY_API_SAFE_LIMIT - apiUsed);
  var usageRatio = SERIES_DAILY_API_SAFE_LIMIT ? apiUsed / SERIES_DAILY_API_SAFE_LIMIT : 0;
  var apiPercent = (Math.min(1, usageRatio) * 100).toFixed(1);
  var barColor = usageRatio >= 0.9 ? '#C85B52' : usageRatio >= 0.7 ? '#D9A441' : '#2D6A50';
  var panelBackground = usageRatio >= 0.9 ? '#FDECEC' : usageRatio >= 0.7 ? '#FFF4D9' : '#EAF5EE';
  var panelText = usageRatio >= 0.9 ? '#A33B36' : usageRatio >= 0.7 ? '#8A6518' : '#2D6A50';

  sheet.getRange('A3:N3').breakApart().clearContent().clearNote();

  sheet.getRange('A3:D3')
    .merge()
    .setValue(summaryText)
    .setBackground('#F7F4EA')
    .setFontColor('#765A25')
    .setFontSize(10)
    .setFontWeight('bold')
    .setWrap(true)
    .setHorizontalAlignment('left')
    .setVerticalAlignment('middle');

  sheet.getRange('E3:F3')
    .merge()
    .setValue(lastCheckedText || '마지막 확인\n-')
    .setBackground('#EEF2EF')
    .setFontColor('#4E6357')
    .setFontSize(9)
    .setFontWeight('bold')
    .setWrap(true)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  sheet.getRange('G3:H3')
    .merge()
    .setFormula(
      '=SPARKLINE(' + apiUsed +
      ',{"charttype","bar";"max",' + SERIES_DAILY_API_SAFE_LIMIT +
      ';"color1","' + barColor + '"})'
    )
    .setBackground(panelBackground)
    .setVerticalAlignment('middle');

  sheet.getRange('I3')
    .setValue('API ' + apiUsed + '/' + SERIES_DAILY_API_SAFE_LIMIT + '\n남음 ' + apiRemaining + '회')
    .setNote('오늘 API 사용률 ' + apiPercent + '% · 이 표시를 확인하거나 새로고침하는 것 자체는 API를 사용하지 않습니다.')
    .setBackground(panelBackground)
    .setFontColor(panelText)
    .setFontSize(9)
    .setFontWeight('bold')
    .setWrap(true)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  sheet.getRange('J3:N3')
    .merge()
    .setValue('사용량 표시 새로고침은 API 미사용')
    .setBackground('#F6F8F7')
    .setFontColor('#78867F')
    .setFontSize(9)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  sheet.getRange('E3:I3').setBorder(
    true, true, true, true, false, false,
    usageRatio >= 0.9 ? '#E3A8A4' : usageRatio >= 0.7 ? '#E6CA87' : '#B7D7C4',
    SpreadsheetApp.BorderStyle.SOLID
  );
}

/**
 * API를 추가로 호출하지 않고 현재 사용량 막대만 다시 표시합니다.
 */
function refreshSeriesApiGauge() {
  var sheet = seriesSpreadsheet_().getSheetByName(SERIES_FINDER_SHEET);
  if (!sheet) {
    seriesSpreadsheet_().toast('시리즈 누락 점검 시트가 없습니다.', 'API 사용량', 5);
    return;
  }

  var currentText = String(sheet.getRange('A3').getDisplayValue() || '').trim();
  var lastMatch = currentText.match(
    /마지막 확인\s+(\d{4}\.\s*\d{1,2}\.\s*\d{1,2}\.\s*\d{1,2}:\d{2})/
  );
  var lastCheckedText = String(sheet.getRange('E3').getDisplayValue() || '').trim();

  if (lastMatch) {
    lastCheckedText = '마지막 확인\n' + lastMatch[1];
    currentText = currentText.replace(
      /\s*·\s*마지막 확인\s+\d{4}\.\s*\d{1,2}\.\s*\d{1,2}\.\s*\d{1,2}:\d{2}/,
      ''
    );
  }
  currentText = currentText
    .replace(/\s*·\s*오늘 API\s*\d+\/\d+회\s*\([^)]*\)\s*·\s*남음\s*\d+회/g, '')
    .trim();

  if (!currentText) currentText = '시리즈 점검 결과';
  if (!lastCheckedText) lastCheckedText = '마지막 확인\n-';

  seriesApplyApiGauge_(sheet, currentText, lastCheckedText);
  seriesSpreadsheet_().toast(
    'API를 호출하지 않고 사용량 표시만 새로고침했습니다.',
    'API 사용량',
    4
  );
}

function seriesStyleResultSheet_(sheet, rowCount) {
  sheet.setFrozenRows(5);
  sheet.setRowHeight(1, 62);
  sheet.setRowHeight(2, 42);
  sheet.setRowHeight(3, 46);
  sheet.setRowHeight(4, 14);
  sheet.setRowHeight(5, 42);

  sheet.getRange('A1:N1')
    .setBackground('#20382D')
    .setFontColor('#FFFFFF')
    .setFontSize(24)
    .setFontWeight('bold')
    .setHorizontalAlignment('left')
    .setVerticalAlignment('middle');
  sheet.getRange('A2:N2')
    .setBackground('#E9F1EC')
    .setFontColor('#40594D')
    .setFontSize(11)
    .setVerticalAlignment('middle');
  sheet.getRange('A3:D3')
    .setBackground('#F7F4EA')
    .setFontColor('#765A25')
    .setFontSize(10)
    .setFontWeight('bold')
    .setHorizontalAlignment('left')
    .setVerticalAlignment('middle');
  sheet.getRange('E3:N3')
    .setFontSize(9)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sheet.getRange('A5:N5')
    .setBackground('#526C5D')
    .setFontColor('#FFFFFF')
    .setFontSize(11)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  var widths = [56, 96, 190, 130, 82, 84, 280, 180, 125, 92, 125, 92, 122, 190];
  widths.forEach(function(width, index) { sheet.setColumnWidth(index + 1, width); });

  if (rowCount) {
    sheet.getRange(6, 1, rowCount, 14)
      .setFontSize(11)
      .setBorder(false, false, true, false, false, false, '#E2E8E4', SpreadsheetApp.BorderStyle.SOLID);
    sheet.getRange(6, 1, rowCount, 6).setHorizontalAlignment('center');
    sheet.getRange(6, 7, rowCount, 1).setFontWeight('bold');
    sheet.getRange(6, 13, rowCount, 1).setFontColor('#2D6A50').setFontWeight('bold');

    var dataRange = sheet.getRange(6, 1, rowCount, 14);
    sheet.setConditionalFormatRules([
      SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied('=$B6="중간 누락"')
        .setBackground('#FFF3EE')
        .setRanges([dataRange])
        .build(),
      SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied('=$B6="신간 후속권"')
        .setBackground('#EEF7F2')
        .setRanges([dataRange])
        .build(),
      SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied('=$A6=TRUE')
        .setBackground('#FFF0B8')
        .setFontColor('#3E3516')
        .setRanges([dataRange])
        .build()
    ]);
  }
  sheet.activate();
  seriesSpreadsheet_().setActiveRange(sheet.getRange('A1'));
}

function setupSeriesWorkSheet() {
  var spreadsheet = seriesSpreadsheet_();
  var sheet = spreadsheet.getSheetByName(SERIES_WORK_SHEET);

  if (!sheet) {
    sheet = spreadsheet.insertSheet(SERIES_WORK_SHEET);
  }

  seriesEnsureWorkSheet_(sheet);
  spreadsheet.setActiveSheet(sheet);
  spreadsheet.setActiveRange(sheet.getRange('A1'));
  spreadsheet.toast(
    '검색 결과에서 필요한 책을 체크한 뒤 위쪽의 「📤 선택한 책 구입 검토함으로 보내기 ›」 막대를 누르세요.',
    '업무목록 준비 완료',
    8
  );
}

function exportSelectedSeriesBooks() {
  var spreadsheet = seriesSpreadsheet_();
  var source = spreadsheet.getSheetByName(SERIES_FINDER_SHEET);
  if (!source) throw new Error('「' + SERIES_FINDER_SHEET + '」 시트를 찾을 수 없습니다.');

  var lastRow = source.getLastRow();
  if (lastRow < 6) {
    spreadsheet.toast('먼저 시리즈 누락 점검을 실행해 주세요.', '보낼 도서 없음', 6);
    return;
  }

  var values = source.getRange(6, 1, lastRow - 5, 14).getValues();
  var display = source.getRange(6, 1, lastRow - 5, 14).getDisplayValues();
  var formulas = source.getRange(6, 1, lastRow - 5, 14).getFormulas();
  var selectedIndexes = [];

  values.forEach(function(row, index) {
    if (row[0] === true) selectedIndexes.push(index);
  });

  if (!selectedIndexes.length) {
    spreadsheet.toast('업무목록으로 보낼 책의 「선택」 칸을 체크해 주세요.', '선택 필요', 7);
    return;
  }

  var target = spreadsheet.getSheetByName(SERIES_WORK_SHEET);
  if (!target) target = spreadsheet.insertSheet(SERIES_WORK_SHEET);
  seriesEnsureWorkSheet_(target);

  var existingKeys = {};
  var targetLastRow = target.getLastRow();
  if (targetLastRow >= 6) {
    var existing = target.getRange(6, 7, targetLastRow - 5, 5).getDisplayValues();
    existing.forEach(function(row) {
      var titleKey = seriesNormalizeKey_(row[0]);
      var isbn = seriesCleanIsbn_(row[4]);
      if (isbn) existingKeys['isbn:' + isbn] = true;
      if (titleKey) existingKeys['title:' + titleKey] = true;
    });
  }

  var output = [];
  var duplicateCount = 0;

  selectedIndexes.forEach(function(index) {
    var rowDisplay = display[index];
    var rowValues = values[index];
    var rowFormulas = formulas[index];
    var isbn = seriesCleanIsbn_(rowDisplay[10]);
    var titleKey = seriesNormalizeKey_(rowDisplay[6]);
    var duplicate =
      (isbn && existingKeys['isbn:' + isbn]) ||
      (titleKey && existingKeys['title:' + titleKey]);

    if (duplicate) {
      duplicateCount++;
      return;
    }

    var coverFormula = rowFormulas[5] || '';
    var linkFormula = rowFormulas[12] || '';
    var price = Number(rowValues[11] || 0) || '';

    output.push([
      '검토중',
      rowDisplay[1],
      rowDisplay[2],
      rowDisplay[3],
      rowDisplay[4],
      coverFormula,
      rowDisplay[6],
      rowDisplay[7],
      rowDisplay[8],
      rowDisplay[9],
      rowDisplay[10],
      price,
      linkFormula,
      new Date(),
      ''
    ]);

    if (isbn) existingKeys['isbn:' + isbn] = true;
    if (titleKey) existingKeys['title:' + titleKey] = true;
  });

  if (output.length) {
    var startRow = Math.max(target.getLastRow() + 1, 6);
    target.getRange(startRow, 1, output.length, 15).setValues(output);
    target.getRange(startRow, 12, output.length, 1).setNumberFormat('#,##0"원"');
    target.getRange(startRow, 14, output.length, 1).setNumberFormat('yyyy. m. d. HH:mm');
    target.getRange(startRow, 1, output.length, 15)
      .setVerticalAlignment('middle')
      .setWrap(true);
    for (var rowNumber = startRow; rowNumber < startRow + output.length; rowNumber++) {
      target.setRowHeight(rowNumber, 112);
    }
    seriesStyleWorkSheet_(target);
  }

  selectedIndexes.forEach(function(index) {
    source.getRange(index + 6, 1).setValue(false);
  });

  spreadsheet.setActiveSheet(target);
  spreadsheet.setActiveRange(
    output.length ? target.getRange(Math.max(target.getLastRow() - output.length + 1, 6), 1) : target.getRange('A1')
  );

  spreadsheet.toast(
    '업무목록에 ' + output.length + '권을 추가했습니다.' +
      (duplicateCount ? ' 이미 있는 ' + duplicateCount + '권은 제외했습니다.' : ''),
    '전송 완료',
    9
  );
}

function seriesEnsureWorkSheet_(sheet) {
  var headers = [
    '처리상태', '구분', '시리즈명', '현재 보유 권차', '구입 권차', '표지', '도서명',
    '저자', '출판사', '출간일', 'ISBN13', '정가', 'YES24', '추출일', '비고'
  ];
  var compatible = sheet.getLastRow() >= 5 &&
    String(sheet.getRange(5, 1).getDisplayValue()) === '처리상태';
  var empty = sheet.getLastRow() <= 1 && !String(sheet.getRange('A1').getDisplayValue()).trim();

  if (!compatible && !empty) {
    throw new Error('「' + SERIES_WORK_SHEET + '」 시트의 기존 형식이 달라 자동으로 덮어쓰지 않았습니다.');
  }

  if (empty) {
    if (sheet.getMaxColumns() < 15) {
      sheet.insertColumnsAfter(sheet.getMaxColumns(), 15 - sheet.getMaxColumns());
    }
    if (sheet.getMaxRows() < 1000) {
      sheet.insertRowsAfter(sheet.getMaxRows(), 1000 - sheet.getMaxRows());
    }
    sheet.clear();
    sheet.clearConditionalFormatRules();
    sheet.setHiddenGridlines(true);
    sheet.getRange('A1:O1').merge().setValue('📋 시리즈 구입 업무목록');
    sheet.getRange('A2:O2').merge().setValue(
      '시리즈 누락 점검 결과에서 실제 구입을 검토할 책만 모아 관리합니다. 같은 ISBN이나 도서명은 중복으로 들어오지 않습니다.'
    );
    sheet.getRange('A3:O3').merge().setValue(
      '업무 흐름  ·  검토중 → 구입결정 → 주문완료   |   구입하지 않을 책은 보류로 바꾸세요.'
    );
    sheet.getRange(5, 1, 1, headers.length).setValues([headers]);
  }

  var validation = SpreadsheetApp.newDataValidation()
    .requireValueInList(['검토중', '구입결정', '보류', '주문완료'], true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(6, 1, Math.max(1, sheet.getMaxRows() - 5), 1).setDataValidation(validation);
  seriesStyleWorkSheet_(sheet);
}

function seriesStyleWorkSheet_(sheet) {
  sheet.setHiddenGridlines(true);
  sheet.setFrozenRows(5);
  sheet.setTabColor('#D9A441');

  sheet.setRowHeight(1, 64);
  sheet.setRowHeight(2, 42);
  sheet.setRowHeight(3, 38);
  sheet.setRowHeight(4, 14);
  sheet.setRowHeight(5, 42);

  sheet.getRange('A1:O1')
    .setBackground('#253C32')
    .setFontColor('#FFFFFF')
    .setFontSize(24)
    .setFontWeight('bold')
    .setHorizontalAlignment('left')
    .setVerticalAlignment('middle');
  sheet.getRange('A2:O2')
    .setBackground('#EAF2ED')
    .setFontColor('#40594D')
    .setFontSize(11)
    .setVerticalAlignment('middle');
  sheet.getRange('A3:O3')
    .setBackground('#FFF5D8')
    .setFontColor('#73551D')
    .setFontSize(11)
    .setFontWeight('bold')
    .setVerticalAlignment('middle');
  sheet.getRange('A5:O5')
    .setBackground('#587064')
    .setFontColor('#FFFFFF')
    .setFontSize(11)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  var widths = [105, 95, 185, 125, 88, 86, 280, 175, 125, 95, 125, 95, 120, 125, 220];
  widths.forEach(function(width, index) {
    sheet.setColumnWidth(index + 1, width);
  });

  var lastRow = sheet.getLastRow();
  if (lastRow >= 6) {
    var rowCount = lastRow - 5;
    sheet.getRange(6, 1, rowCount, 15)
      .setFontSize(11)
      .setVerticalAlignment('middle')
      .setWrap(true)
      .setBorder(false, false, true, false, false, false, '#E2E8E4', SpreadsheetApp.BorderStyle.SOLID);
    sheet.getRange(6, 1, rowCount, 6).setHorizontalAlignment('center');
    sheet.getRange(6, 7, rowCount, 1).setFontWeight('bold');
    sheet.getRange(6, 12, rowCount, 1).setNumberFormat('#,##0"원"');
    sheet.getRange(6, 14, rowCount, 1).setNumberFormat('yyyy. m. d. HH:mm');
    for (var row = 6; row <= lastRow; row++) sheet.setRowHeight(row, 112);
  }

  if (sheet.getFilter()) sheet.getFilter().remove();
  if (sheet.getLastRow() >= 5) {
    sheet.getRange(5, 1, Math.max(1, sheet.getLastRow() - 4), 15).createFilter();
  }

  var dataRange = sheet.getRange(6, 1, Math.max(1, sheet.getMaxRows() - 5), 15);
  sheet.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$A6="검토중"')
      .setBackground('#FFF9E8')
      .setRanges([dataRange])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$A6="구입결정"')
      .setBackground('#E9F5EC')
      .setRanges([dataRange])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$A6="보류"')
      .setBackground('#F2F2F2')
      .setFontColor('#777777')
      .setRanges([dataRange])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$A6="주문완료"')
      .setBackground('#EAF0FA')
      .setRanges([dataRange])
      .build()
  ]);
}

function seriesEscapeFormula_(value) {
  return String(value || '').replace(/"/g, '""');
}



var UNIFIED_DASHBOARD_SHEET = '🏠 업무 대시보드';
var UNIFIED_REVIEW_SHEET = '🛒 구입 검토함';
var UNIFIED_UPLOAD_SHEET = '📥 엑셀 업로드 목록';
var UNIFIED_SETTINGS_SHEET = '⚙️ 통합 설정';
var UNIFIED_REQUEST_SHEET = '📚 희망도서 목록';
var UNIFIED_EXPORT_SHEET = '📦 서점 전달용';

function unifiedAddMenu_() {
  SpreadsheetApp.getUi()
    .createMenu('🧰 통합 도서업무')
    .addItem('🏠 전체 화면 정리·새로고침', 'setupUnifiedLibraryWorkflow')
    .addItem('🎨 전체 디자인 다시 적용', 'applyUnifiedLibraryDesign')
    .addItem('🧭 시트 탭 정리', 'HR_arrangeTabs')
    .addSeparator()
    .addItem('📚 희망도서를 구입 검토함으로 보내기', 'sendHopeRequestsToReview')
    .addItem('🔎 시리즈 누락 점검 (쉬운 창)', 'HR_showSeriesCheckDialog')
    .addItem('⏹ 시리즈 점검 멈추기', 'HR_seriesStopCheck')
    .addItem('🔎 시리즈 누락·후속권 이어서 점검', 'refreshMissingSeriesBooks')
    .addItem('🛒 체크한 시리즈를 구입 검토함으로 보내기', 'sendSelectedSeriesToReview')
    .addItem('📥 업로드 목록을 열기', 'openUnifiedUploadSheet')
    .addItem('💵 예산 설정', 'unifiedSetBudget')
    .addItem('✅ 이번 회차 구입 확정', 'unifiedConfirmPurchaseRound')
    .addItem('🧾 구입 기록 보기', 'unifiedShowPurchaseLog')
    .addItem('🧾 구입 기록 숨기기', 'unifiedHidePurchaseLog')
    .addSeparator()
    .addItem('📦 중복 확인 후 서점 전달용으로 보내기', 'buildUnifiedBookstoreExportSheet')
    .addSeparator()
    .addSubMenu(
      SpreadsheetApp.getUi()
        .createMenu('🔎 시리즈 점검 도구')
        .addItem('현재 결과 절판 여부 확인', 'refreshSeriesSaleStatuses')
        .addItem('결과 시트 만들기·복구', 'setupSeriesFinder')
        .addItem('저장된 점검 위치 초기화', 'resetSeriesFinderProgress')
        .addSeparator()
        .addItem('YES24 API 키 설정', 'setSeriesYes24Key')
    )
    .addToUi();
}

function setupUnifiedLibraryWorkflow() {
  var ss = seriesSpreadsheet_();

  unifiedMigrateSheetNamesAndLayout_();

  if (!ss.getSheetByName(SERIES_FINDER_SHEET)) {
    seriesBuildResultSheet_([], {
      seriesCount: 0,
      searchedCount: 0,
      requestCount: 0,
      checkedAt: new Date()
    });
  }

  unifiedEnsureReviewSheet_();
  unifiedEnsureUploadSheetReady_();
  createUnifiedGuides_();
  sendHopeRequestsToReview();
  applyUnifiedLibraryDesign();
  unifiedRefreshDashboard_();

  ss.setActiveSheet(ss.getSheetByName(UNIFIED_DASHBOARD_SHEET));
  ss.toast('희망도서 → 구입 검토함 → 서점 전달용 흐름으로 정리했습니다.', '통합 설정 완료', 8);
}

function openUnifiedUploadSheet() {
  var ss = seriesSpreadsheet_();
  var sheet = unifiedEnsureUploadSheet_();
  var rowCount = Math.max(0, sheet.getLastRow() - 5);

  unifiedStyleUploadSheet_(sheet, rowCount);
  ss.setActiveSheet(sheet);
}

function refreshUnifiedPurchaseReview() {
  return unifiedRefreshPurchaseReview_([]);
}

function unifiedReadRequestItems_(sheet) {
  if (sheet.getLastRow() < FIRST_DATA_ROW) return [];

  var headers = sheet.getRange(HEADER_ROW, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  var map = unifiedHeaderMap_(headers);
  var rowCount = sheet.getLastRow() - FIRST_DATA_ROW + 1;
  var display = sheet.getRange(FIRST_DATA_ROW, 1, rowCount, sheet.getLastColumn()).getDisplayValues();
  var raw = sheet.getRange(FIRST_DATA_ROW, 1, rowCount, sheet.getLastColumn()).getValues();
  var requestersColumn = headers.map(function(h) { return String(h || '').replace(/\s+/g, ''); }).indexOf('신청인원') + 1;

  return display.map(function(row, index) {
    var status = map.status ? String(row[map.status - 1] || '') : '';
    if (/취소|오류|중복|소장/.test(status)) return null;
    if (map.sequence && !/^\d+$/.test(String(row[map.sequence - 1] || '').trim())) return null;

    var title = map.title ? String(row[map.title - 1] || '').trim() : '';
    var isbn = map.isbn ? String(row[map.isbn - 1] || '') : '';
    if (!title && !isbn) return null;

    var price = map.price
      ? Number(raw[index][map.price - 1] || String(row[map.price - 1] || '').replace(/[^0-9]/g, ''))
      : 0;

    return {
      source: '희망도서',
      title: title,
      author: map.author ? row[map.author - 1] : '',
      publisher: map.publisher ? row[map.publisher - 1] : '',
      year: map.year ? (String(row[map.year - 1] || '').match(/\d{4}/) || [''])[0] : '',
      isbn: isbn,
      category: map.category ? String(row[map.category - 1] || '') : '',
      price: price || 0,
      quantity: 1,
      requesters: requestersColumn ? Math.max(1, Number(row[requestersColumn - 1]) || 1) : 1
    };
  }).filter(function(item) { return item; });
}

function unifiedReadSeriesItems_(sheet) {
  if (sheet.getLastRow() < 6) return [];
  var headers = sheet.getRange(5, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  var map = unifiedHeaderMap_(headers);
  var display = sheet.getRange(6, 1, sheet.getLastRow() - 5, sheet.getLastColumn()).getDisplayValues();
  var raw = sheet.getRange(6, 1, sheet.getLastRow() - 5, sheet.getLastColumn()).getValues();
  return display.map(function(row, index) {
    var status = map.status ? String(row[map.status - 1] || '') : '';
    if (/보류|주문완료/.test(status)) return null;
    var title = map.title ? String(row[map.title - 1] || '').trim() : '';
    var isbn = map.isbn ? String(row[map.isbn - 1] || '') : '';
    if (!title && !isbn) return null;
    return {
      source: '시리즈 누락',
      title: title,
      author: map.author ? row[map.author - 1] : '',
      publisher: map.publisher ? row[map.publisher - 1] : '',
      year: map.year ? (String(row[map.year - 1] || '').match(/\d{4}/) || [''])[0] : '',
      isbn: isbn,
      price: map.price ? Number(raw[index][map.price - 1] || String(row[map.price - 1] || '').replace(/[^0-9]/g, '')) : 0,
      quantity: 1
    };
  }).filter(function(item) { return item; });
}

function unifiedHeaderMap_(headers) {
  var map = {};
  headers.forEach(function(value, index) {
    var h = String(value || '').replace(/\s+/g, '').toUpperCase().replace(/\([^)]*\)/g, '');
    if (!map.sequence && /^(순번|번호|NO)$/.test(h)) map.sequence = index + 1;
    if (!map.title && /^(도서명|책제목|서명|제목)$/.test(h)) map.title = index + 1;
    if (!map.author && /^(저자|작가|글쓴이)$/.test(h)) map.author = index + 1;
    if (!map.publisher && /^(출판사|발행자|출판자)$/.test(h)) map.publisher = index + 1;
    if (!map.year && /^(출간일|출판연도|발행연도|출간연도|연도)$/.test(h)) map.year = index + 1;
    if (!map.isbn && /^ISBN(13)?$/.test(h)) map.isbn = index + 1;
    if (!map.price && /^(정가|도서정가|가격)$/.test(h)) map.price = index + 1;
    if (!map.status && /^(처리상태|자동입력상태|상태)$/.test(h)) map.status = index + 1;
    if (!map.category && /^(분야|카테고리|주제|KDC)$/.test(h)) map.category = index + 1;
  });
  return map;
}

function unifiedReadReviewManual_(sheet) {
  var result = {};
  if (sheet.getLastRow() < 6) return result;
  var rows = sheet.getRange(6, 1, sheet.getLastRow() - 5, 12).getValues();
  rows.forEach(function(row) {
    var isbn = seriesCleanIsbn_(row[7]);
    var titleKey = seriesNormalizeKey_(row[3]);
    var key = isbn ? 'isbn:' + isbn : 'title:' + titleKey;
    if (!key || key === 'title:') return;
    result[key] = {selected: row[0] === true, status: row[2], quantity: row[9], note: row[11]};
  });
  return result;
}

function unifiedEnsureReviewSheet_() {
  var ss = seriesSpreadsheet_();
  var sheet = ss.getSheetByName(UNIFIED_REVIEW_SHEET) || ss.insertSheet(UNIFIED_REVIEW_SHEET);

  if (sheet.getMaxColumns() < 12) sheet.insertColumnsAfter(sheet.getMaxColumns(), 12 - sheet.getMaxColumns());
  if (sheet.getMaxRows() < 1000) sheet.insertRowsAfter(sheet.getMaxRows(), 1000 - sheet.getMaxRows());

  var compatible = sheet.getLastRow() >= 5 && String(sheet.getRange('A5').getDisplayValue()) === '선택';
  var empty = sheet.getLastRow() <= 1 && !String(sheet.getRange('A1').getDisplayValue()).trim();
  if (!compatible && !empty) throw new Error('「' + UNIFIED_REVIEW_SHEET + '」에 다른 내용이 있어 덮어쓰지 않았습니다.');

  if (empty) {
    sheet.setHiddenGridlines(true);
    sheet.getRange('A1:L1').merge().setValue('🛒 구입 검토함');
    sheet.getRange('A2:L2').merge().setValue('희망도서와 선택한 시리즈 누락 도서를 한곳에서 최종 검토합니다.');
    sheet.getRange('A3:L3').merge().setValue('구입할 책을 체크한 뒤 아래 금색 버튼을 누르면 소장도서를 다시 확인하고 서점 전달용 목록을 만듭니다.');
    sheet.getRange(5, 1, 1, 12).setValues([[
      '선택','출처','처리상태','도서명','저자','출판사','출판연도','ISBN13','정가','수량','금액','비고'
    ]]);
  }

  sheet.getRange('A3:L3')
    .breakApart()
    .merge()
    .setValue('구입할 도서의 「선택」 체크박스를 누르면 「서점 전달용」 시트에 반영됩니다.');

  sheet.getRange('A4:L4').breakApart().clearContent().clearDataValidations();
  sheet.getRange('A4')
    .insertCheckboxes()
    .setValue(false)
    .setBackground('#2F5D50')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  sheet.getRange('B4:C4')
    .merge()
    .setValue('전체 선택/해제')
    .setBackground('#2F5D50')
    .setFontColor('#FFFFFF')
    .setFontSize(13)
    .setFontWeight('bold')
    .setHorizontalAlignment('left')
    .setVerticalAlignment('middle');

  sheet.getRange('D4')
    .insertCheckboxes()
    .setValue(false)
    .setBackground('#D9A441')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  sheet.getRange('E4:L4')
    .merge()
    .setValue('중복 확인 후 서점 전달용으로 보내기')
    .setBackground('#D9A441')
    .setFontColor('#FFFFFF')
    .setFontSize(14)
    .setFontWeight('bold')
    .setHorizontalAlignment('left')
    .setVerticalAlignment('middle');

  sheet.getRange('B4:C4')
    .merge()
    .setValue('전체 선택/해제')
    .setBackground('#2F5D50')
    .setFontColor('#FFFFFF')
    .setFontSize(13)
    .setFontWeight('bold')
    .setHorizontalAlignment('left')
    .setVerticalAlignment('middle');

  sheet.getRange('D4')
    .insertCheckboxes()
    .setValue(false)
    .setBackground('#D9A441')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  sheet.getRange('E4:L4')
    .merge()
    .setValue('중복 확인 후 서점 전달용으로 보내기')
    .setBackground('#D9A441')
    .setFontColor('#FFFFFF')
    .setFontSize(14)
    .setFontWeight('bold')
    .setHorizontalAlignment('left')
    .setVerticalAlignment('middle');

  unifiedStyleReviewSheet_(sheet);
  return sheet;
}

function unifiedStyleReviewSheet_(sheet) {
  sheet.setHiddenGridlines(true);
  sheet.setFrozenRows(5);

  sheet.getRange('A1:L1').setBackground('#183D35').setFontColor('#FFFFFF').setFontFamily('Arial').setFontSize(23).setFontWeight('bold').setHorizontalAlignment('left');
  sheet.getRange('A2:L2').setBackground('#EAF3ED').setFontColor('#34564D').setFontFamily('Arial').setFontSize(12);
  sheet.getRange('A3:L3').setBackground('#FFF4D6').setFontColor('#765413').setFontFamily('Arial').setFontSize(11).setFontWeight('bold');
  sheet.getRange('A4:L4').setBorder(true,true,true,true,false,false,'#B88428',SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sheet.getRange('A5:L5').setBackground('#2E5B4F').setFontColor('#FFFFFF').setFontFamily('Arial').setFontWeight('bold').setHorizontalAlignment('center');

  sheet.setRowHeight(1, 58);
  sheet.setRowHeight(2, 34);
  sheet.setRowHeight(3, 34);
  sheet.setRowHeight(4, 44);
  sheet.setRowHeight(5, 38);

  var widths = [65,110,105,260,170,135,95,135,100,75,110,220];
  widths.forEach(function(width, i) { sheet.setColumnWidth(i + 1, width); });

  sheet.getRange(6, 1, Math.max(1, sheet.getMaxRows() - 5), 12)
    .setWrap(true)
    .setVerticalAlignment('middle')
    .setFontFamily('Arial')
    .setFontSize(11)
    .setBackground('#FBFCFB');

  var formatRowCount = Math.max(1, sheet.getMaxRows() - 5);
  sheet.getRange(6, 9, formatRowCount, 1).setNumberFormat('#,##0"원"');
  sheet.getRange(6, 10, formatRowCount, 1).setNumberFormat('0');
  sheet.getRange(6, 11, formatRowCount, 1).setNumberFormat('#,##0"원"');

  // 문제집·학습지 의심은 줄 전체를 연한 빨강으로, 만화는 도서명 칸을 연한 보라로 표시합니다.
  sheet.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=REGEXMATCH($L6&"","문제집·학습지")')
      .setBackground('#FDECEA')
      .setRanges([sheet.getRange(6, 1, Math.max(1, sheet.getMaxRows() - 5), 12)])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=REGEXMATCH($L6&"","만화")')
      .setBackground('#EFE6F7')
      .setRanges([sheet.getRange(6, 4, Math.max(1, sheet.getMaxRows() - 5), 1)])
      .build()
  ]);

  unifiedDrawActionBars_(sheet);   // 체크박스 대신 클릭 막대
}

function buildUnifiedBookstoreExportSheet() {
  var ss = seriesSpreadsheet_();
  var source = ss.getSheetByName(UNIFIED_REVIEW_SHEET);

  if (!source || source.getLastRow() < 6) {
    throw new Error('먼저 희망도서 또는 시리즈 도서를 구입 검토함으로 보내 주세요.');
  }

  var rowCount = source.getLastRow() - 5;
  var rows = source.getRange(6, 1, rowCount, 12).getValues();
  var holdings = seriesReadHoldings_();
  var books = [];
  var ownedCount = 0;

  rows.forEach(function(row, index) {
    var selected = row[0] === true || String(row[2]) === '구입결정';
    if (!selected) return;

    var isbn = seriesCleanIsbn_(row[7]);
    var titleKey = seriesNormalizeKey_(row[3]);
    var alreadyOwned = !String(row[1] || '').includes('관리자') && Boolean(
      (isbn && holdings.isbnSet[isbn]) ||
      (titleKey && holdings.titleKeySet[titleKey])
    );

    if (alreadyOwned) {
      source.getRange(index + 6, 3).setValue('소장중');
      source.getRange(index + 6, 1).setValue(false);
      ownedCount++;
      return;
    }

    books.push(row);
  });

  if (!books.length) {
    throw new Error(
      ownedCount
        ? '선택한 도서는 모두 우리학교 소장도서입니다. 검토함의 상태를 확인해 주세요.'
        : '구입할 책의 선택 칸을 체크하거나 처리상태를 「구입결정」으로 바꿔 주세요.'
    );
  }

  // 문제집·학습지로 보이는 책이 있으면 한 번 더 확인합니다.
  if (!unifiedConfirmWorkbooks_(books)) {
    ss.toast('서점 전달용 만들기를 멈췄습니다. 구입 검토함에서 해당 책을 확인해 주세요.', '문제집·학습지 확인', 8);
    return { success: false, cancelled: true };
  }

  var target = ss.getSheetByName(UNIFIED_EXPORT_SHEET) || ss.insertSheet(UNIFIED_EXPORT_SHEET);
  if (target.getFilter()) target.getFilter().remove();
  target.getBandings().forEach(function(b) { b.remove(); });
  target.getRange(1, 1, Math.min(target.getMaxRows(), 5), Math.min(target.getMaxColumns(), 10)).breakApart();
  target.clear();
  target.clearConditionalFormatRules();
  target.setHiddenGridlines(true);

  target.getRange('A1:J1').merge().setValue('📦 서점 전달용 도서 주문 목록');
  target.getRange('A2:J2').merge().setValue('구입 검토를 마친 희망도서와 시리즈 누락 도서입니다.');
  target.getRange('A3:J3').merge().setValue(
    ownedCount
      ? '우리학교 소장도서 ' + ownedCount + '권을 제외했습니다.'
      : '우리학교 소장도서와 다시 대조해 중복이 없는 도서만 담았습니다.'
  );
  target.getRange('A4')
    .insertCheckboxes()
    .setValue(false)
    .setBackground('#D9A441')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  target.getRange('B4:J4')
    .merge()
    .setValue('최신 구입 검토 결과로 다시 만들기')
    .setBackground('#D9A441')
    .setFontColor('#FFFFFF')
    .setFontSize(14)
    .setFontWeight('bold')
    .setHorizontalAlignment('left')
    .setVerticalAlignment('middle');
  target.getRange('A5:J5').setValues([[
    '번호','도서명','저자','출판사','출판연도','ISBN13','정가','수량','금액','출처·비고'
  ]]);

  var output = books.map(function(row, index) {
    var quantity = Math.max(1, Number(row[9] || 1));
    var price = Number(row[8] || 0);
    return [
      index + 1,row[3],row[4],row[5],row[6],row[7],
      price || '',quantity,price ? price * quantity : '',
      row[1] + (row[11] ? ' · ' + row[11] : '')
    ];
  });

  target.getRange(6, 1, output.length, 10).setValues(output);
  target.getRange(6, 7, output.length, 1).setNumberFormat('#,##0"원"');
  target.getRange(6, 8, output.length, 1).setNumberFormat('0');
  target.getRange(6, 9, output.length, 1).setNumberFormat('#,##0"원"');
  target.getRange(5, 1, output.length + 1, 10).createFilter();
  unifiedStyleExportSheet_(target, output.length);

  ss.setActiveSheet(target);
  PropertiesService.getDocumentProperties().setProperty(UNIFIED_EXPORT_BUILT_AT_PROPERTY, String(Date.now()));   // 만든 시각 (대시보드에 표시)
  unifiedRefreshDashboard_();
  ss.toast(
    '서점 전달용 ' + output.length + '권을 만들었습니다.' +
    (ownedCount ? ' 소장도서 ' + ownedCount + '권은 제외했습니다.' : ''),
    '중복 확인 완료',
    8
  );

  return {success:true, exported:output.length, owned:ownedCount};
}

function unifiedStyleExportSheet_(sheet, count) {
  sheet.setFrozenRows(5);
  sheet.getRange('A1:J1').setBackground('#183D35').setFontColor('#FFFFFF').setFontSize(24).setFontWeight('bold');
  sheet.getRange('A2:J2').setBackground('#F7E8BF').setFontColor('#6B4D11').setFontSize(12);
  sheet.getRange('A4:J4').setBorder(true,true,true,true,false,false,'#B88428',SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sheet.getRange('A5:J5').setBackground('#B77A2A').setFontColor('#FFFFFF').setFontWeight('bold').setHorizontalAlignment('center');
  sheet.setRowHeight(1, 60); sheet.setRowHeight(2, 34); sheet.setRowHeight(4, 44); sheet.setRowHeight(5, 40);
  var widths = [55,285,180,140,95,135,100,70,110,220];
  widths.forEach(function(width, i) { sheet.setColumnWidth(i + 1, width); });
  if (count) {
    sheet.getRange(6, 1, count, 10).setWrap(true).setVerticalAlignment('middle').setFontSize(11);
    sheet.getRange(6, 7, count, 1).setNumberFormat('#,##0"원"');
    sheet.getRange(6, 8, count, 1).setNumberFormat('0');
    sheet.getRange(6, 9, count, 1).setNumberFormat('#,##0"원"');
    sheet.setRowHeights(6, count, 60);
  }

  unifiedDrawActionBars_(sheet);   // 체크박스 대신 클릭 막대
}

function unifiedEnsureSettingsSheet_() {
  var ss = seriesSpreadsheet_();
  var sheet = ss.getSheetByName(UNIFIED_SETTINGS_SHEET) || ss.insertSheet(UNIFIED_SETTINGS_SHEET);
  sheet.clear();
  sheet.setHiddenGridlines(true);
  sheet.getRange('A1:H1').merge().setValue('⚙️ 통합 도서업무 설정');
  sheet.getRange('A2:H2').merge().setValue('희망도서 신청과 시리즈 누락 점검은 같은 YES24 API 키와 우리학교 소장도서 목록을 사용합니다.');
  sheet.getRange('A4:B8').setValues([
    ['확인 항목','현재 설정'],
    ['희망도서 원본 시트',UNIFIED_REQUEST_SHEET],
    ['소장도서 원본 시트',SERIES_HOLDINGS_SHEET],
    ['시리즈 결과 시트',SERIES_FINDER_SHEET],
    ['통합 검토 시트',UNIFIED_REVIEW_SHEET]
  ]);
  var key = seriesGetYes24Key_();
  sheet.getRange('D4:E6').setValues([
    ['API 상태',''],
    ['YES24 API',key ? '설정됨' : '설정 필요'],
    ['안내','API 키는 상단 「⚙️ 사본 설정」에서 입력합니다.']
  ]);
  sheet.getRange('A1:H1').setBackground('#183D35').setFontColor('#FFFFFF').setFontSize(23).setFontWeight('bold');
  sheet.getRange('A2:H2').setBackground('#EAF3ED').setFontColor('#34564D');
  sheet.getRange('A4:B4').setBackground('#2E5B4F').setFontColor('#FFFFFF').setFontWeight('bold');
  sheet.getRange('D4:E4').setBackground('#B77A2A').setFontColor('#FFFFFF').setFontWeight('bold');
  sheet.setColumnWidth(1, 190); sheet.setColumnWidth(2, 230); sheet.setColumnWidth(3, 30); sheet.setColumnWidth(4, 170); sheet.setColumnWidth(5, 330);
  return sheet;
}

// 대시보드 모양을 바꾸면 이 번호를 올립니다. 번호가 다르면 숫자만 고치는 대신 전체를 다시 그립니다.
var UNIFIED_DASHBOARD_LAYOUT = '5';
var UNIFIED_BUDGET_PROPERTY = 'UNIFIED_BUDGET';
var UNIFIED_DASHBOARD_LAYOUT_PROPERTY = 'UNIFIED_DASHBOARD_LAYOUT';
var UNIFIED_EXPORT_BUILT_AT_PROPERTY = 'UNIFIED_EXPORT_BUILT_AT';

// 학교 예산(원). 설정하지 않았으면 0
function unifiedGetBudget_() {
  var raw = PropertiesService.getScriptProperties().getProperty(UNIFIED_BUDGET_PROPERTY);
  var value = Number(raw);
  return isFinite(value) && value > 0 ? Math.round(value) : 0;
}

// 입력한 글을 예산(원)으로 바꿉니다. 비어 있으면 0(예산 끔), 숫자로 읽을 수 없으면 null.
// 예: 3000000 · 3,000,000원 · 300만원 · 2.5만
function unifiedParseBudget_(input) {
  var text = String(input == null ? '' : input).replace(/[,\s원₩]/g, '');
  if (!text || /여기에입력/.test(text)) return 0;
  var man = text.match(/^(\d+(?:\.\d+)?)만$/);
  var value = man ? Math.round(Number(man[1]) * 10000) : Number(text);
  return isFinite(value) && value > 0 ? Math.round(value) : null;
}

function unifiedSaveBudget_(value) {
  var props = PropertiesService.getScriptProperties();
  if (value) props.setProperty(UNIFIED_BUDGET_PROPERTY, String(Math.round(value)));
  else props.deleteProperty(UNIFIED_BUDGET_PROPERTY);
}

// 대시보드의 「총 예산」 칸(A17)에 직접 입력했을 때: 저장하고 대시보드를 다시 계산합니다.
function unifiedBudgetFromDashboard_(range) {
  var value = unifiedParseBudget_(range.getValue());
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (value === null) {
    ss.toast('예산은 숫자로 입력해 주세요. 예: 3000000 또는 300만원', '💵 예산', 7);
  } else {
    unifiedSaveBudget_(value);
    ss.toast(value ? '총 예산을 ' + unifiedMoney_(value) + '으로 정했습니다.' : '예산 표시를 껐습니다.', '💵 예산', 5);
  }
  unifiedUpdateDashboardNumbers_();   // 잘못 입력했으면 원래 값으로 되돌려 보여 줍니다.
}

// 메뉴: 「💵 예산 설정」 (대시보드 칸에 직접 입력하는 것이 더 쉽지만, 메뉴로도 입력할 수 있습니다)
function unifiedSetBudget() {
  var ui = SpreadsheetApp.getUi();
  var current = unifiedGetBudget_();
  var response = ui.prompt(
    '💵 예산 설정',
    '이번 도서 구입 총 예산(원)을 숫자로 입력해 주세요. 예: 3000000 또는 300만원\n' +
      (current ? '지금 설정: ' + unifiedMoney_(current) + '\n' : '') +
      '비워 두고 확인을 누르면 예산 표시를 끕니다.\n' +
      '(대시보드의 「총 예산」 칸에 바로 입력해도 됩니다.)',
    ui.ButtonSet.OK_CANCEL
  );
  if (response.getSelectedButton() !== ui.Button.OK) return;

  var value = unifiedParseBudget_(response.getResponseText());
  if (value === null) {
    ui.alert('예산을 숫자로 입력해 주세요. (예: 3000000, 300만원)');
    return;
  }
  unifiedSaveBudget_(value);
  unifiedUpdateDashboardNumbers_();
  SpreadsheetApp.getActiveSpreadsheet().toast(
    value ? '예산을 ' + unifiedMoney_(value) + '으로 설정했습니다.' : '예산 표시를 껐습니다.',
    '💵 예산', 6
  );
}

// 책 한 권을 가리키는 열쇠: ISBN이 있으면 ISBN, 없으면 도서명
function unifiedBookKey_(isbn, title) {
  var cleanIsbn = seriesCleanIsbn_(isbn);
  if (cleanIsbn) return 'i:' + cleanIsbn;
  var titleKey = seriesNormalizeKey_(title);
  return titleKey ? 't:' + titleKey : '';
}

function unifiedMoney_(amount) {
  return String(Math.round(Number(amount) || 0)).replace(/\B(?=(\d{3})+(?!\d))/g, ',') + '원';
}

// 대시보드에 쓸 숫자를 한 번에 셉니다.
// onlyPurchase 이면 구입 진행 현황(결정·서점 전달)만 셉니다. (체크를 바꿀 때마다 빠르게 고치기 위해)
function unifiedDashboardStats_(ss, onlyPurchase) {
  var stats = {};

  if (!onlyPurchase) {
    var requestSource = ss.getSheetByName(UNIFIED_REQUEST_SHEET);
    stats.request = requestSource ? unifiedReadRequestItems_(requestSource).length : 0;
    stats.series = unifiedCountData_(ss.getSheetByName(SERIES_FINDER_SHEET), 6);
    stats.upload = unifiedCountData_(ss.getSheetByName(UNIFIED_UPLOAD_SHEET), 6);
    stats.review = unifiedCountData_(ss.getSheetByName(UNIFIED_REVIEW_SHEET), 6);
    stats.audience = unifiedAudienceCounts_(ss);
  }

  // 구입 검토함에서 체크한(또는 구입결정인) 책
  var review = ss.getSheetByName(UNIFIED_REVIEW_SHEET);
  var decidedKeys = {};
  var decided = 0;
  var decidedAmount = 0;
  var workbookDecided = 0;
  var comicDecided = 0;
  if (review && review.getLastRow() >= 6) {
    review.getRange(6, 1, review.getLastRow() - 5, 12).getValues().forEach(function(row) {
      if (!String(row[3] || '').trim()) return;                       // 도서명이 없는 줄은 책이 아닙니다.
      if (!(row[0] === true || String(row[2]) === '구입결정')) return;
      decided++;
      var decidedFlags = unifiedBookFlags_(row[3], '');
      var decidedNoted = unifiedFlagsFromNote_(row[11]);
      if (decidedFlags.workbook || decidedNoted.workbook) workbookDecided++;
      if (decidedFlags.comic || decidedNoted.comic) comicDecided++;
      var price = Number(row[8] || 0);
      var quantity = Math.max(1, Number(row[9] || 1));
      decidedAmount += Number(row[10] || 0) || price * quantity;
      var key = unifiedBookKey_(row[7], row[3]);
      if (key) decidedKeys[key] = true;
    });
  }
  stats.decided = decided;
  stats.workbookDecided = workbookDecided;
  stats.comicDecided = comicDecided;
  stats.decidedAmount = decidedAmount;
  stats.budget = unifiedGetBudget_();
  var purchaseLog = unifiedPurchaseLogStats_(ss);
  stats.confirmedAmount = purchaseLog.amount;
  stats.confirmedRounds = purchaseLog.rounds;
  stats.confirmedQuantity = purchaseLog.quantity;

  // 서점 전달용에 만들어 둔 책
  var exportSheet = ss.getSheetByName(UNIFIED_EXPORT_SHEET);
  var exportKeys = {};
  var exportCount = 0;
  var exportAmount = 0;
  if (exportSheet && exportSheet.getLastRow() >= 6) {
    exportSheet.getRange(6, 1, exportSheet.getLastRow() - 5, 10).getValues().forEach(function(row) {
      if (!String(row[1] || '').trim()) return;
      exportCount++;
      var price = Number(row[6] || 0);
      var quantity = Math.max(1, Number(row[7] || 1));
      exportAmount += Number(row[8] || 0) || price * quantity;
      var key = unifiedBookKey_(row[5], row[1]);
      if (key) exportKeys[key] = true;
    });
  }
  stats.exportCount = exportCount;
  stats.exportAmount = exportAmount;

  // 서점 전달용 상태: none(아직 안 만듦) · same(검토함 선택과 같음) · differ(다름) · stale(체크가 0권인데 목록이 있음)
  var builtAt = Number(PropertiesService.getDocumentProperties().getProperty(UNIFIED_EXPORT_BUILT_AT_PROPERTY) || 0);
  stats.builtAtText = builtAt
    ? Utilities.formatDate(new Date(builtAt), 'Asia/Seoul', 'M월 d일 H:mm')
    : '';
  if (!exportCount) {
    stats.exportState = 'none';
  } else if (!decided) {
    stats.exportState = 'stale';
  } else {
    var a = Object.keys(decidedKeys), b = Object.keys(exportKeys);
    var same = a.length === b.length && a.every(function(key) { return exportKeys[key]; });
    stats.exportState = same ? 'same' : 'differ';
  }
  return stats;
}

// 숫자와 설명 글만 다시 채웁니다. (모양은 그대로)
function unifiedFillDashboardNumbers_(sheet, stats, onlyPurchase) {
  var GREEN_MID = '#2E5B4F', MUTED = '#6C7F77', RED = '#B3261E';

  if (!onlyPurchase) {
    sheet.getRange('A2').setValue(
      '희망도서·시리즈 후보·업로드 목록을 모아 중복을 확인하고 서점 전달용 목록으로 정리합니다.   ·   갱신 ' +
      Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy년 M월 d일 H:mm')
    );
    sheet.getRange('A6').setValue(stats.request + '권');
    sheet.getRange('D6').setValue(stats.series + '권');
    sheet.getRange('G6').setValue(stats.upload + '권');
    sheet.getRange('J6').setValue(stats.review + '권');
    sheet.getRange('A8').setValue(
      stats.audience ? '교직원 ' + stats.audience.staff + '권 · 학부모 ' + stats.audience.parent + '권' : '교직원·학부모 신청'
    );
  }

  // 구입 진행 현황
  sheet.getRange('A12').setValue(stats.decided + '권');
  var decidedNote = stats.decided
    ? '검토함에서 체크한 책 · ' + unifiedMoney_(stats.decidedAmount)
    : '아직 체크한 책이 없습니다';
  if (stats.workbookDecided) decidedNote += '\n⚠ 문제집·학습지로 보이는 책 ' + stats.workbookDecided + '권 포함';
  else if (stats.comicDecided) decidedNote += ' · 📙 만화 ' + stats.comicDecided + '권';
  sheet.getRange('A13').setValue(decidedNote)
    .setFontColor(stats.workbookDecided ? RED : MUTED).setFontWeight(stats.workbookDecided ? 'bold' : 'normal');

  sheet.getRange('E12').setValue(stats.exportCount + '권');
  var note;
  var noteColor = MUTED, noteWeight = 'normal';
  if (stats.exportState === 'none') {
    note = '아직 만들지 않았습니다';
  } else if (stats.exportState === 'same') {
    note = '✔ 검토함 선택과 같습니다' + (stats.builtAtText ? ' · ' + stats.builtAtText + ' 만듦' : '');
    noteColor = GREEN_MID;
  } else if (stats.exportState === 'differ') {
    note = '⚠ 검토함 선택과 다릅니다 · 다시 만들어 주세요';
    noteColor = RED; noteWeight = 'bold';
  } else {
    note = '이전에 만든 목록입니다' + (stats.builtAtText ? ' · ' + stats.builtAtText : '');
  }
  sheet.getRange('E13').setValue(note).setFontColor(noteColor).setFontWeight(noteWeight);

  // 예산 현황: 총 예산 − (지금까지 확정한 회차 금액 + 이번 회차 예정 금액)
  var budget = Number(stats.budget || 0);
  var confirmed = Number(stats.confirmedAmount || 0);
  var planned = Number(stats.decidedAmount || 0);
  var used = confirmed + planned;
  var remain = budget - used;
  sheet.getRange('A17').setValue(budget ? unifiedMoney_(budget) : '✏️  여기에 입력');
  sheet.getRange('D17').setValue(unifiedMoney_(confirmed));
  sheet.getRange('D18').setValue(stats.confirmedRounds ? stats.confirmedRounds + '회 확정 · ' + (stats.confirmedQuantity || 0) + '권' : '아직 확정한 회차가 없습니다');
  sheet.getRange('G17').setValue(unifiedMoney_(planned));
  sheet.getRange('G18').setValue(stats.decided ? '검토함에서 체크한 ' + stats.decided + '권' : '체크한 책 없음');
  sheet.getRange('J17').setValue(budget ? (remain < 0 ? '-' + unifiedMoney_(-remain) : unifiedMoney_(remain)) : '-')
    .setFontColor(budget && remain < 0 ? RED : '#1F5C4A');
  var bar = sheet.getRange('A19');
  if (budget) {
    bar.setFormula('=SPARKLINE(' + Math.min(used, budget) + ',{"charttype","bar";"max",' + budget + ';"color1","' + (remain < 0 ? '#B3261E' : '#2E5B4F') + '";"color2","#E3EFE9"})');
  } else {
    bar.setValue('');
  }
  var percent = budget ? Math.round(used / budget * 100) : 0;
  sheet.getRange('A20').setValue(
    !budget ? '총 예산을 아직 정하지 않았습니다. 위의 노란 「총 예산」 칸에 금액을 쓰고 Enter를 누르세요.'
      : remain < 0 ? '⚠ 예산을 ' + unifiedMoney_(-remain) + ' 넘었습니다 · 사용률 ' + percent + '%'
      : '사용률 ' + percent + '% · 아직 ' + unifiedMoney_(remain) + ' 쓸 수 있습니다 (확정 ' + unifiedMoney_(confirmed) + ' + 이번 회차 예정 ' + unifiedMoney_(planned) + ')'
  ).setFontColor(budget && remain < 0 ? RED : MUTED).setFontWeight(budget && remain < 0 ? 'bold' : 'normal');

  sheet.getRange('I12').setValue(unifiedMoney_(stats.exportAmount));
  sheet.getRange('I13').setValue(stats.exportCount ? '서점 전달용 정가 × 수량 합계' : '서점 전달용을 만들면 표시됩니다')
    .setFontColor(MUTED).setFontWeight('normal');
}

// 체크를 바꾸거나 서점 전달용 금액·수량을 고칠 때 숫자만 바로 고칩니다.
// full 이면 모인 책 현황(희망도서·시리즈·업로드·검토 권수)까지 모두 다시 셉니다.
// 모양이 예전 것이면(처음 한 번) 전체를 다시 그립니다.
function unifiedUpdateDashboardNumbers_(full) {
  var ss = seriesSpreadsheet_();
  var sheet = ss.getSheetByName(UNIFIED_DASHBOARD_SHEET);
  if (!sheet) return;
  var layout = PropertiesService.getDocumentProperties().getProperty(UNIFIED_DASHBOARD_LAYOUT_PROPERTY);
  if (layout !== UNIFIED_DASHBOARD_LAYOUT) {
    unifiedRefreshDashboard_();
    return;
  }
  unifiedFillDashboardNumbers_(sheet, unifiedDashboardStats_(ss, !full), !full);
}

// 대시보드 모양을 그립니다. 숫자는 unifiedFillDashboardNumbers_ 가 채웁니다.
function unifiedRefreshDashboard_() {
  var ss = seriesSpreadsheet_();
  var sheet = ss.getSheetByName(UNIFIED_DASHBOARD_SHEET) || ss.insertSheet(UNIFIED_DASHBOARD_SHEET, 0);

  sheet.getDataRange().getMergedRanges().forEach(function(range) {
    range.breakApart();
  });
  sheet.clear();
  sheet.setHiddenGridlines(true);

  var GREEN = '#183D35', GREEN_MID = '#2E5B4F', GREEN_SOFT = '#E3EFE9', GOLD = '#D9A441', GOLD_SOFT = '#FBF1D4';
  var INK = '#26352F', MUTED = '#6C7F77', LINE = '#CFE0D8';

  // 모든 줄 높이를 먼저 기본으로 돌려 둡니다. (이전 모양이 남지 않게)
  sheet.setRowHeights(1, 28, 24);
  for (var column = 1; column <= 12; column++) sheet.setColumnWidth(column, 96);
  sheet.getRange('A1:L28').setVerticalAlignment('middle').setFontColor(INK);

  // 1) 제목 띠
  sheet.getRange('A1:L1').merge().setValue('🏠 학교도서관 통합 구입 업무')
    .setBackground(GREEN).setFontColor('#FFFFFF').setFontSize(26).setFontWeight('bold')
    .setHorizontalAlignment('left').setBorder(false, false, false, false, false, false);
  sheet.getRange('A2:L2').merge().setValue('')
    .setBackground(GREEN).setFontColor('#BFD6CC').setFontSize(11).setHorizontalAlignment('left');
  sheet.getRange('A3:L3').setBackground(GOLD);
  sheet.setRowHeight(1, 64);
  sheet.setRowHeight(2, 32);
  sheet.setRowHeight(3, 5);

  // 2) 모인 책 현황 카드 4개 (제목 / 숫자 / 설명)
  sheet.getRange('A4:L4').merge().setValue('📊  모인 책 현황')
    .setFontSize(13).setFontWeight('bold').setFontColor(GREEN_MID).setHorizontalAlignment('left');
  sheet.setRowHeight(4, 38);
  var cards = [
    ['A', 'C', '📚 희망도서 신청', '교직원·학부모 신청'],
    ['D', 'F', '🔎 시리즈 후보', '누락·후속권 점검 결과'],
    ['G', 'I', '📥 파일 업로드', '엑셀·CSV로 올린 목록'],
    ['J', 'L', '🛒 통합 검토', '구입 검토함에 모인 책']
  ];
  cards.forEach(function(card) {
    sheet.getRange(card[0] + '5:' + card[1] + '5').merge().setValue(card[2])
      .setFontSize(11).setFontWeight('bold').setFontColor(GREEN_MID).setBackground(GREEN_SOFT).setHorizontalAlignment('center');
    sheet.getRange(card[0] + '6:' + card[1] + '7').merge().setValue('0권')
      .setFontSize(30).setFontWeight('bold').setFontColor(GREEN).setBackground('#FFFFFF').setHorizontalAlignment('center');
    sheet.getRange(card[0] + '8:' + card[1] + '8').merge().setValue(card[3])
      .setFontSize(10).setFontColor(MUTED).setBackground('#FFFFFF').setHorizontalAlignment('center');
    sheet.getRange(card[0] + '5:' + card[1] + '8')
      .setBorder(true, true, true, true, false, false, LINE, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    sheet.getRange(card[0] + '5:' + card[1] + '5')
      .setBorder(true, true, false, true, false, false, GOLD, SpreadsheetApp.BorderStyle.SOLID_THICK);
  });
  sheet.setRowHeight(5, 32);
  sheet.setRowHeight(6, 30);
  sheet.setRowHeight(7, 30);
  sheet.setRowHeight(8, 28);

  // 3) 구입 진행 현황 (금색): 제목(11) / 숫자(12) / 설명(13)
  sheet.getRange('A10:L10').merge().setValue('💰  구입 진행 현황')
    .setFontSize(13).setFontWeight('bold').setFontColor(GREEN_MID).setHorizontalAlignment('left');
  sheet.setRowHeight(9, 12);
  sheet.setRowHeight(10, 38);
  var kpis = [
    ['A', 'D', '✅ 구입 결정'],
    ['E', 'H', '📦 서점 전달용'],
    ['I', 'L', '💵 서점 전달 금액']
  ];
  kpis.forEach(function(kpi) {
    sheet.getRange(kpi[0] + '11:' + kpi[1] + '11').merge().setValue(kpi[2])
      .setFontSize(11).setFontWeight('bold').setFontColor('#7A5A12').setHorizontalAlignment('center');
    sheet.getRange(kpi[0] + '12:' + kpi[1] + '12').merge().setValue('0')
      .setFontSize(22).setFontWeight('bold').setFontColor('#5B410A').setHorizontalAlignment('center');
    sheet.getRange(kpi[0] + '13:' + kpi[1] + '13').merge().setValue('')
      .setFontSize(10).setFontColor(MUTED).setHorizontalAlignment('center').setWrap(true);
    sheet.getRange(kpi[0] + '11:' + kpi[1] + '13').setBackground(GOLD_SOFT)
      .setBorder(true, true, true, true, false, false, GOLD, SpreadsheetApp.BorderStyle.SOLID);
  });
  sheet.setRowHeight(11, 28);
  sheet.setRowHeight(12, 44);
  sheet.setRowHeight(13, 38);

  // 3-2) 예산 현황: 제목(15) / 이름(16) / 값(17) / 설명(18) / 막대(19) / 요약(20)
  var logSheet = unifiedEnsurePurchaseLogSheet_();
  // 예전에 보이는 채로 만들어진 기록 시트는 한 번 숨깁니다.
  var docProps = PropertiesService.getDocumentProperties();
  if (docProps.getProperty('UNIFIED_LOG_HIDDEN_ONCE') !== '1') {
    try { if (!logSheet.isSheetHidden()) logSheet.hideSheet(); } catch (error) {}
    docProps.setProperty('UNIFIED_LOG_HIDDEN_ONCE', '1');
  }
  sheet.getRange('A15:L15').merge().setValue('🧾  예산 현황')
    .setFontSize(13).setFontWeight('bold').setFontColor(GREEN_MID).setHorizontalAlignment('left');
  sheet.setRowHeight(14, 12);
  sheet.setRowHeight(15, 38);
  [['A', 'C', '💰 총 예산 (눌러서 입력)'], ['D', 'F', '✅ 지금까지 확정 (눌러서 보기)'], ['G', 'I', '🛒 이번 회차 예정'], ['J', 'L', '🟢 남은 예산']].forEach(function(item) {
    var label = sheet.getRange(item[0] + '16:' + item[1] + '16').merge();
    label.setValue(item[2]);   // 「지금까지 확정」 칸은 누르면 숨겨 둔 기록 시트가 열립니다.
    label.setFontSize(11).setFontWeight('bold').setFontColor(GREEN_MID).setHorizontalAlignment('center').setBackground(GREEN_SOFT);
    sheet.getRange(item[0] + '17:' + item[1] + '17').merge().setValue('-')
      .setFontSize(20).setFontWeight('bold').setFontColor('#1F5C4A').setHorizontalAlignment('center').setBackground('#FFFFFF');
    sheet.getRange(item[0] + '18:' + item[1] + '18').merge().setValue('')
      .setFontSize(9).setFontColor(MUTED).setHorizontalAlignment('center').setBackground('#FFFFFF');
    sheet.getRange(item[0] + '16:' + item[1] + '18')
      .setBorder(true, true, true, true, false, false, LINE, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  });
  sheet.getRange('A19:L19').merge().setBackground('#FFFFFF').setVerticalAlignment('middle');
  sheet.getRange('A20:L20').merge().setValue('')
    .setFontSize(10).setFontColor(MUTED).setHorizontalAlignment('center').setBackground('#FFFFFF');
  sheet.getRange('A19:L20').setBorder(true, true, true, true, false, false, LINE, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  // 총 예산 칸은 직접 쓰는 입력칸입니다. (눈에 띄게 노란 바탕·굵은 금색 테두리)
  sheet.getRange('A17:C17').setBackground(GOLD_SOFT).setFontColor('#7A5A12')
    .setBorder(true, true, true, true, false, false, GOLD, SpreadsheetApp.BorderStyle.SOLID_THICK);
  sheet.getRange('A17').setNote('여기에 총 예산을 쓰고 Enter를 누르세요.\n예: 3000000 · 3,000,000원 · 300만원\n지우면 예산 표시가 꺼집니다.');
  sheet.setRowHeight(16, 28);
  sheet.setRowHeight(17, 42);
  sheet.setRowHeight(18, 22);
  sheet.setRowHeight(19, 24);
  sheet.setRowHeight(20, 28);

  // 4) 업무 순서 (바로가기 포함): 제목(22) / 단계(23) / 설명(24) / 바로가기(25)
  sheet.getRange('A22:L22').merge().setValue('🧭  업무 순서')
    .setFontSize(13).setFontWeight('bold').setFontColor(GREEN_MID).setHorizontalAlignment('left');
  sheet.setRowHeight(21, 12);
  sheet.setRowHeight(22, 38);
  var steps = [
    ['A', 'C', '1', '자료 모으기', '희망도서·시리즈·엑셀 목록', UNIFIED_REQUEST_SHEET, '📚 희망도서 목록'],
    ['D', 'F', '2', '구입 검토', '구입할 책 체크·수량 확인', UNIFIED_REVIEW_SHEET, '🛒 구입 검토함'],
    ['G', 'I', '3', '중복 확인', '소장도서와 자동 대조', null, '🏫 소장도서 확인'],
    ['J', 'L', '4', '서점 전달·확정', '주문 후 「구입 확정」으로 회차 기록', UNIFIED_EXPORT_SHEET, '📦 서점 전달용']
  ];
  steps.forEach(function(step) {
    var c1 = step[0], c2 = step[1];
    sheet.getRange(c1 + '23:' + c2 + '23').merge().setValue(step[2] + '  ' + step[3])
      .setBackground(GREEN_MID).setFontColor('#FFFFFF').setFontSize(12).setFontWeight('bold').setHorizontalAlignment('center');
    sheet.getRange(c1 + '24:' + c2 + '24').merge().setValue(step[4])
      .setBackground('#FFFFFF').setFontSize(10).setFontColor(MUTED).setHorizontalAlignment('center').setWrap(true);
    var link = sheet.getRange(c1 + '25:' + c2 + '25').merge().setBackground('#FFFFFF')
      .setFontSize(11).setFontWeight('bold').setFontColor(GREEN_MID).setHorizontalAlignment('center');
    var target = step[5] ? ss.getSheetByName(step[5]) : (ss.getSheetByName('🏫 우리학교 소장도서') || null);
    if (target) {
      link.setFormula('=HYPERLINK("#gid=' + target.getSheetId() + '","' + step[6] + ' ›")');
    } else {
      link.setValue(step[6]);
    }
    sheet.getRange(c1 + '23:' + c2 + '25')
      .setBorder(true, true, true, true, false, false, LINE, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  });
  sheet.setRowHeight(23, 34);
  sheet.setRowHeight(24, 30);
  sheet.setRowHeight(25, 32);

  // 5) 안내
  sheet.getRange('A27:L27').merge()
    .setValue('💡 필요한 기능은 위쪽 「🧰 통합 도서업무」 메뉴에서 실행하고, 위 업무 순서의 시트 이름을 누르면 해당 시트로 바로 이동합니다.')
    .setBackground(GOLD_SOFT).setFontSize(10).setFontColor('#7A5A12').setHorizontalAlignment('left').setWrap(true);
  sheet.setRowHeight(26, 12);
  sheet.setRowHeight(27, 34);
  sheet.setFrozenRows(3);

  // 숫자 채우기와 모양 번호 기록
  unifiedFillDashboardNumbers_(sheet, unifiedDashboardStats_(ss, false), false);
  PropertiesService.getDocumentProperties().setProperty(UNIFIED_DASHBOARD_LAYOUT_PROPERTY, UNIFIED_DASHBOARD_LAYOUT);
}

// 희망도서 목록에서 교직원/학부모 신청 권수를 셉니다. (열이 없거나 문제가 있으면 null)
function unifiedAudienceCounts_(ss) {
  try {
    var list = ss.getSheetByName(UNIFIED_REQUEST_SHEET);
    if (!list || list.getLastRow() < 6 || list.getLastColumn() < 1) return null;
    var headers = list.getRange(5, 1, 1, list.getLastColumn()).getDisplayValues()[0];
    var col = 0;
    for (var i = 0; i < headers.length; i++) {
      if (String(headers[i]).replace(/\s+/g, '') === '신청구분') { col = i + 1; break; }
    }
    var titleCol = 2;
    var rows = list.getRange(6, 1, list.getLastRow() - 5, list.getLastColumn()).getDisplayValues();
    var staff = 0, parent = 0;
    rows.forEach(function(row) {
      if (!String(row[titleCol - 1] || '').trim()) return;
      if (col && String(row[col - 1]).trim() === '학부모') parent++; else staff++;
    });
    return { staff: staff, parent: parent };
  } catch (error) {
    return null;
  }
}

function unifiedCountData_(sheet, startRow) {
  if (!sheet || sheet.getLastRow() < startRow) return 0;
  return sheet.getRange(startRow, 1, sheet.getLastRow() - startRow + 1, Math.max(1, sheet.getLastColumn()))
    .getDisplayValues().filter(function(row) {
      // 빈 체크박스(FALSE)만 들어 있는 줄은 책이 아니므로 세지 않습니다.
      return row.some(function(value) {
        var text = String(value || '').trim();
        return text && text.toUpperCase() !== 'FALSE';
      });
    }).length;
}


function unifiedAuditDataSources() {
  var ss = seriesSpreadsheet_();
  var requestSheet = ss.getSheetByName(UNIFIED_REQUEST_SHEET);
  var seriesSheet = ss.getSheetByName(SERIES_WORK_SHEET);
  var requestItems = requestSheet ? unifiedReadRequestItems_(requestSheet) : [];
  var seriesItems = seriesSheet ? unifiedReadSeriesItems_(seriesSheet) : [];
  var holdings = seriesReadHoldings_();
  console.log(JSON.stringify({
    requestLastRow: requestSheet ? requestSheet.getLastRow() : 0,
    requestHead: requestSheet ? requestSheet.getRange(1, 1, Math.min(4, requestSheet.getLastRow()), requestSheet.getLastColumn()).getDisplayValues() : [],
    requestItems: requestItems.length,
    seriesItems: seriesItems.length,
    holdings: holdings.rows.length
  }));
}


// 원본(예전) 시리즈 시트의 ID는 코드에 두지 않고 스크립트 속성 LEGACY_SERIES_SOURCE_ID 에 입력해 사용합니다.
function importLegacySeriesSheetsOnce() {
  var targetBook = seriesSpreadsheet_();
  var legacyId = String(
    PropertiesService.getScriptProperties().getProperty('LEGACY_SERIES_SOURCE_ID') || ''
  ).trim();
  if (!legacyId) {
    throw new Error('가져올 원본 시트 ID가 없습니다. 프로젝트 설정 → 스크립트 속성에 LEGACY_SERIES_SOURCE_ID 를 입력해 주세요.');
  }
  var sourceBook = SpreadsheetApp.openById(legacyId);
  var names = [SERIES_FINDER_SHEET, SERIES_WORK_SHEET];
  var imported = [];

  names.forEach(function(name) {
    var source = sourceBook.getSheetByName(name);
    if (!source) return;
    var current = targetBook.getSheetByName(name);
    var hasData = current && (
      (name === SERIES_FINDER_SHEET && current.getLastRow() >= 6 && String(current.getRange('G6').getDisplayValue()).trim()) ||
      (name === SERIES_WORK_SHEET && current.getLastRow() >= 6 && String(current.getRange('G6').getDisplayValue()).trim())
    );
    if (hasData) return;

    var copied = source.copyTo(targetBook);
    copied.setName('__통합중__' + name);
    var safeSheet = targetBook.getSheetByName(UNIFIED_DASHBOARD_SHEET) || targetBook.getSheets()[0];
    targetBook.setActiveSheet(safeSheet);
    if (current) targetBook.deleteSheet(current);
    copied.setName(name);
    imported.push(name);
  });

  refreshUnifiedPurchaseReview();
  unifiedRefreshDashboard_();
  targetBook.toast(
    imported.length ? imported.join(', ') + '의 기존 데이터를 가져왔습니다.' : '이미 통합된 시리즈 데이터가 있어 그대로 두었습니다.',
    '기존 자료 통합',
    8
  );
}


function unifiedReadExistingSeriesReview_(sheet) {
  var items = [];
  if (!sheet || sheet.getLastRow() < 6) return items;

  var rows = sheet.getRange(6, 1, sheet.getLastRow() - 5, 12).getValues();
  rows.forEach(function(row) {
    var source = String(row[1] || '').trim();
    if (/희망도서/.test(source)) return;
    if (/주문완료|소장중/.test(String(row[2] || ''))) return;
    if (!row[3] && !row[7]) return;

    items.push({
      source: source || '기존 검토',
      title: row[3],
      author: row[4],
      publisher: row[5],
      year: row[6],
      isbn: row[7],
      price: Number(row[8] || 0),
      quantity: Math.max(1, Number(row[9] || 1))
    });
  });

  return items;
}

function unifiedReadSelectedSeriesItems_(sheet) {
  if (!sheet || sheet.getLastRow() < 6) return [];

  var headers = sheet.getRange(5, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  var map = unifiedHeaderMap_(headers);
  var rowCount = sheet.getLastRow() - 5;
  var display = sheet.getRange(6, 1, rowCount, sheet.getLastColumn()).getDisplayValues();
  var raw = sheet.getRange(6, 1, rowCount, sheet.getLastColumn()).getValues();

  return display.map(function(row, index) {
    if (raw[index][0] !== true) return null;

    var title = map.title ? String(row[map.title - 1] || '').trim() : '';
    var isbn = map.isbn ? String(row[map.isbn - 1] || '') : '';
    if (!title && !isbn) return null;

    return {
      source: '시리즈 누락',
      title: title,
      author: map.author ? row[map.author - 1] : '',
      publisher: map.publisher ? row[map.publisher - 1] : '',
      year: map.year ? (String(row[map.year - 1] || '').match(/\d{4}/) || [''])[0] : '',
      isbn: isbn,
      price: map.price
        ? Number(raw[index][map.price - 1] || String(row[map.price - 1] || '').replace(/[^0-9]/g, ''))
        : 0,
      quantity: 1
    };
  }).filter(function(item) { return item; });
}

function unifiedRefreshPurchaseReview_(newSeriesItems) {
  var ss = seriesSpreadsheet_();
  var review = unifiedEnsureReviewSheet_();
  var oldManual = unifiedReadReviewManual_(review);
  var preservedSeries = unifiedReadExistingSeriesReview_(review);
  var holdings = seriesReadHoldings_();
  var items = preservedSeries.concat(Array.isArray(newSeriesItems) ? newSeriesItems : []);

  var requestSheet = ss.getSheetByName(UNIFIED_REQUEST_SHEET);
  if (requestSheet) {
    items = unifiedReadRequestItems_(requestSheet).concat(items);
  }

  var grouped = {};
  var order = [];
  var ownedSkipped = 0;

  items.forEach(function(item) {
    var isbn = seriesCleanIsbn_(item.isbn || item.isbn13 || item.isbn10);
    var titleKey = seriesNormalizeKey_(item.title);

    if ((isbn && holdings.isbnSet[isbn]) || (titleKey && holdings.titleKeySet[titleKey])) {
      ownedSkipped++;
      return;
    }

    var key = isbn ? 'isbn:' + isbn : 'title:' + titleKey;
    if (!key || key === 'title:') return;

    if (!grouped[key]) {
      grouped[key] = item;
      grouped[key].isbn = isbn;
      grouped[key].quantity = Math.max(1, Number(item.quantity || 1));
      grouped[key].sources = {};
      order.push(key);
    } else {
      grouped[key].quantity = Math.max(
        grouped[key].quantity,
        Math.max(1, Number(item.quantity || 1))
      );
      if (!grouped[key].author && item.author) grouped[key].author = item.author;
      if (!grouped[key].publisher && item.publisher) grouped[key].publisher = item.publisher;
      if (!grouped[key].year && item.year) grouped[key].year = item.year;
      if (!grouped[key].price && item.price) grouped[key].price = item.price;
      grouped[key].requesters = Math.max(Number(grouped[key].requesters) || 1, Number(item.requesters) || 1);
    }

    grouped[key].sources[item.source] = true;
  });

  // 많이 신청한 책이 위로 오도록 정렬합니다(같은 인원이면 들어온 순서 그대로).
  order = order.map(function(key, index) { return { key: key, index: index }; }).sort(function(a, b) {
    return (Number(grouped[b.key].requesters) || 1) - (Number(grouped[a.key].requesters) || 1) || a.index - b.index;
  }).map(function(entry) { return entry.key; });

  var output = order.map(function(key) {
    var item = grouped[key];
    var manual = oldManual[key] || oldManual['title:' + seriesNormalizeKey_(item.title)] || {};
    var quantity = Number(manual.quantity || item.quantity || 1);
    var price = Number(item.price || 0);

    return [
      manual.selected === true,
      Object.keys(item.sources).join(' + '),
      manual.status || '검토중',
      item.title,
      item.author,
      item.publisher,
      item.year,
      item.isbn,
      price || '',
      quantity,
      price ? price * quantity : '',
      HR_composeRequesterNote_(unifiedComposeNote_(manual.note || item.note || '', unifiedBookFlags_(item.title, item.category)), item.requesters)
    ];
  });
  var workbookCount = output.filter(function(row) { return unifiedFlagsFromNote_(row[11]).workbook; }).length;

  if (review.getFilter()) review.getFilter().remove();
  review.getRange(6, 1, Math.max(1, review.getMaxRows() - 5), 12).clearContent().clearDataValidations();

  if (output.length) {
    review.getRange(6, 1, output.length, 12).setValues(output);
    review.getRange(6, 1, output.length, 1).insertCheckboxes();
    review.getRange(6, 3, output.length, 1).setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInList(['검토중','구입결정','보류','주문완료','소장중'], true)
        .setAllowInvalid(false)
        .build()
    );
    review.getRange(6, 9, output.length, 1).setNumberFormat('#,##0"원"');
    review.getRange(6, 10, output.length, 1).setNumberFormat('0');
    review.getRange(6, 11, output.length, 1).setNumberFormat('#,##0"원"');
    review.getRange(5, 1, output.length + 1, 12).createFilter();
    for (var rowNumber = 6; rowNumber < 6 + output.length; rowNumber++) {
      review.setRowHeight(rowNumber, 62);
    }
  }

  unifiedStyleReviewSheet_(review);
  unifiedRefreshDashboard_();

  return {
    count: output.length,
    ownedSkipped: ownedSkipped,
    workbookCount: workbookCount
  };
}

function sendHopeRequestsToReview() {
  var ss = seriesSpreadsheet_();
  var result = unifiedRefreshPurchaseReview_([]);
  var review = ss.getSheetByName(UNIFIED_REVIEW_SHEET);
  if (review) ss.setActiveSheet(review);

  ss.toast(
    '희망도서를 구입 검토함에 모았습니다. 현재 ' + result.count + '권입니다.' +
      (result.workbookCount ? ' ⚠ 문제집·학습지로 보이는 책이 ' + result.workbookCount + '권 있습니다. 비고를 확인해 주세요.' : ''),
    '전송 완료',
    7
  );

  return result;
}

function sendSelectedSeriesToReview() {
  var ss = seriesSpreadsheet_();
  var sheet = ss.getSheetByName(SERIES_FINDER_SHEET);
  if (!sheet) throw new Error('「' + SERIES_FINDER_SHEET + '」 시트를 찾을 수 없습니다.');

  var selected = unifiedReadSelectedSeriesItems_(sheet);
  if (!selected.length) {
    throw new Error('구입 검토함으로 보낼 시리즈 도서를 먼저 체크해 주세요.');
  }

  var result = unifiedRefreshPurchaseReview_(selected);
  sheet.getRange(6, 1, Math.max(1, sheet.getLastRow() - 5), 1).setValue(false);

  var review = ss.getSheetByName(UNIFIED_REVIEW_SHEET);
  if (review) ss.setActiveSheet(review);

  ss.toast(
    '선택한 시리즈 ' + selected.length + '권을 구입 검토함으로 보냈습니다.',
    '전송 완료',
    7
  );

  return result;
}



/**
 * ============================================================
 * 파일 업로드 목록
 * ============================================================
 */
function unifiedEnsureUploadSheet_() {
  var ss = seriesSpreadsheet_();
  var sheet = ss.getSheetByName(UNIFIED_UPLOAD_SHEET);

  if (!sheet) {
    // 처음 만들 때는 「희망도서 목록」 바로 오른쪽에 놓습니다. (목록이 없으면 맨 뒤)
    var listSheet = ss.getSheetByName(UNIFIED_REQUEST_SHEET);
    sheet = listSheet
      ? ss.insertSheet(UNIFIED_UPLOAD_SHEET, listSheet.getIndex())
      : ss.insertSheet(UNIFIED_UPLOAD_SHEET);
  }

  if (sheet.getMaxColumns() < 11) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), 11 - sheet.getMaxColumns());
  }
  if (sheet.getMaxRows() < 100) {
    sheet.insertRowsAfter(sheet.getMaxRows(), 100 - sheet.getMaxRows());
  }

  return sheet;
}

// 「📥 엑셀 업로드 목록」 시트를 미리 만들어 둡니다. (웹에서 파일을 올려야 그제서야 생기지 않도록)
// 이미 있으면 아무것도 바꾸지 않습니다. 「희망도서 목록」 바로 오른쪽에 놓습니다.
function unifiedEnsureUploadSheetReady_() {
  var ss = seriesSpreadsheet_();
  if (ss.getSheetByName(UNIFIED_UPLOAD_SHEET)) return false;

  var sheet = unifiedEnsureUploadSheet_();
  sheet.setHiddenGridlines(true);
  sheet.getRange('A1:K1').merge().setValue('📥 파일 업로드 도서 목록');
  sheet.getRange('A2:K2').merge().setValue(
    '웹페이지의 「📥 목록 업로드」에서 엑셀·CSV를 올리면 필요한 서지정보만 여기에 채워집니다. 구입할 책을 고른 뒤 구입 검토함으로 보내세요.'
  );
  sheet.getRange('A3:K3').merge().setValue('아직 올린 파일이 없습니다.');
  sheet.getRange(5, 1, 1, 11).setValues([[
    '선택', '도서명', '저자', '출판사', '출판연도', 'ISBN13', '정가', '수량', '금액', '시리즈명', '비고'
  ]]);
  unifiedStyleUploadSheet_(sheet, 0);   // 모양과 클릭 막대
  return true;
}

function unifiedReplaceUploadRows_(books, fileName) {
  var lock = LockService.getDocumentLock();
  lock.waitLock(30000);

  try {
    if (!Array.isArray(books) || !books.length) {
      throw new Error('업로드할 도서 정보가 없습니다.');
    }

    var limited = books.slice(0, 1500);
    var unique = [];
    var seen = {};
    var skipped = 0;

    limited.forEach(function(book) {
      book = book || {};

      var title = String(book.title || '').trim();
      var isbn = seriesCleanIsbn_(book.isbn || '');
      var publisher = String(book.publisher || '').trim();
      var yearMatch = String(book.year || '').match(/\d{4}/);
      var year = yearMatch ? yearMatch[0] : '';
      var price = Number(String(book.price || '').replace(/[^0-9.-]/g, '')) || 0;
      var quantity = Math.max(1, Number(book.quantity || 1));
      var seriesName = String(book.series || '').trim();

      if (!title && !isbn) {
        skipped++;
        return;
      }

      var key = isbn
        ? 'isbn:' + isbn
        : 'title:' + seriesNormalizeKey_(title) + '|' + seriesNormalizeKey_(publisher) + '|' + year;

      if (seen[key]) {
        skipped++;
        return;
      }
      seen[key] = true;

      unique.push({
        title: title,
        author: String(book.author || '').trim(),
        publisher: publisher,
        year: year,
        isbn: isbn,
        price: price,
        quantity: quantity,
        series: seriesName,
        note: String(book.note || '').trim()
      });
    });

    if (!unique.length) {
      throw new Error('도서명 또는 ISBN이 있는 행을 찾지 못했습니다.');
    }

    var sheet = unifiedEnsureUploadSheet_();
    var requiredRows = unique.length + 5;
    if (sheet.getMaxRows() < requiredRows) {
      sheet.insertRowsAfter(sheet.getMaxRows(), requiredRows - sheet.getMaxRows());
    }

    if (sheet.getFilter()) {
      sheet.getFilter().remove();
    }

    // 기존 시트에 남아 있는 고정 열과 병합 범위를 먼저 해제해야
    // 제목 영역을 다시 만들 때 병합/고정 충돌이 발생하지 않습니다.
    sheet.setFrozenRows(0);
    sheet.setFrozenColumns(0);
    sheet.getRange(1, 1, Math.min(5, sheet.getMaxRows()), 11)
      .getMergedRanges()
      .forEach(function(range) {
        range.breakApart();
      });

    sheet.clear();
    sheet.clearConditionalFormatRules();

    sheet.getRange('A1:K1').merge().setValue('📥 파일 업로드 도서 목록');
    sheet.getRange('A2:K2').merge().setValue(
      '엑셀·CSV에서 필요한 서지정보만 가져왔습니다. 구입할 책을 선택한 뒤 구입 검토함으로 보내세요.'
    );
    sheet.getRange('A3:K3').merge().setValue(
      '원본 파일: ' + String(fileName || '업로드 파일') + '  ·  불러온 도서 ' + unique.length + '권'
    );

    sheet.getRange('A4').insertCheckboxes().setValue(false);
    sheet.getRange('B4:C4').merge().setValue('목록 전체 선택 / 전체 해제');
    sheet.getRange('D4').insertCheckboxes().setValue(false);
    sheet.getRange('E4:K4').merge().setValue('선택한 책을 구입 검토함으로 보내기 →');

    var headers = [
      '선택', '도서명', '저자', '출판사', '출판연도',
      'ISBN13', '정가', '수량', '금액', '시리즈명', '비고'
    ];
    sheet.getRange(5, 1, 1, headers.length).setValues([headers]);

    var rows = unique.map(function(book) {
      return [
        false,
        book.title,
        book.author,
        book.publisher,
        book.year,
        book.isbn,
        book.price || '',
        book.quantity,
        book.price ? book.price * book.quantity : '',
        book.series,
        unifiedComposeNote_(book.note, unifiedBookFlags_(book.title, ''))
      ];
    });

    sheet.getRange(6, 1, rows.length, 11).setValues(rows);
    sheet.getRange(6, 1, rows.length, 1).insertCheckboxes();
    sheet.getRange(6, 6, rows.length, 1).setNumberFormat('@');
    sheet.getRange(6, 7, rows.length, 1).setNumberFormat('#,##0"원"');
    sheet.getRange(6, 8, rows.length, 1).setNumberFormat('0');
    sheet.getRange(6, 9, rows.length, 1).setNumberFormat('#,##0"원"');
    sheet.getRange(5, 1, rows.length + 1, 11).createFilter();

    unifiedStyleUploadSheet_(sheet, rows.length);
    SpreadsheetApp.flush();

    return {
      importedCount: rows.length,
      skippedCount: skipped + Math.max(0, books.length - limited.length),
      limited: books.length > 1500,
      sheetName: UNIFIED_UPLOAD_SHEET
    };
  } finally {
    lock.releaseLock();
  }
}

function unifiedReadSelectedUploadItems_(sheet) {
  if (!sheet || sheet.getLastRow() < 6) return [];

  var rowCount = sheet.getLastRow() - 5;
  var raw = sheet.getRange(6, 1, rowCount, 11).getValues();
  var display = sheet.getRange(6, 1, rowCount, 11).getDisplayValues();

  return raw.map(function(row, index) {
    if (row[0] !== true) return null;

    var title = String(display[index][1] || '').trim();
    var isbn = seriesCleanIsbn_(display[index][5] || '');
    if (!title && !isbn) return null;

    var seriesName = String(display[index][9] || '').trim();
    var note = String(display[index][10] || '').trim();

    return {
      source: '파일 업로드',
      title: title,
      author: display[index][2],
      publisher: display[index][3],
      year: (String(display[index][4] || '').match(/\d{4}/) || [''])[0],
      isbn: isbn,
      price: Number(row[6] || String(display[index][6] || '').replace(/[^0-9]/g, '')) || 0,
      quantity: Math.max(1, Number(row[7] || 1)),
      note: [seriesName ? '시리즈: ' + seriesName : '', note].filter(function(value) {
        return value;
      }).join(' · ')
    };
  }).filter(function(item) {
    return item;
  });
}

function sendSelectedUploadItemsToReview() {
  var ss = seriesSpreadsheet_();
  var sheet = ss.getSheetByName(UNIFIED_UPLOAD_SHEET);
  if (!sheet) {
    throw new Error('「' + UNIFIED_UPLOAD_SHEET + '」 시트를 찾을 수 없습니다.');
  }

  var selected = unifiedReadSelectedUploadItems_(sheet);
  if (!selected.length) {
    throw new Error('구입 검토함으로 보낼 책을 먼저 체크해 주세요.');
  }

  var result = unifiedRefreshPurchaseReview_(selected);
  sheet.getRange(6, 1, Math.max(1, sheet.getLastRow() - 5), 1).setValue(false);
  unifiedResetBarState_(sheet);

  var review = ss.getSheetByName(UNIFIED_REVIEW_SHEET);
  if (review) ss.setActiveSheet(review);

  ss.toast(
    '선택한 ' + selected.length + '권을 구입 검토함으로 보냈습니다.',
    '전송 완료',
    7
  );

  return {
    sentCount: selected.length,
    reviewCount: result.count,
    ownedSkipped: result.ownedSkipped
  };
}

function unifiedStyleUploadSheet_(sheet, rowCount) {
  if (!sheet) return;

  sheet.setHiddenGridlines(true);

  // 제목과 안내 행은 여러 열에 병합되어 있으므로 열 고정은 사용하지 않습니다.
  // 행만 고정하면 병합 셀 오류 없이 제목과 헤더가 계속 보입니다.
  sheet.setFrozenColumns(0);
  sheet.setFrozenRows(5);
  sheet.setTabColor('#2F7258');

  sheet.getRange('A1:K1')
    .setBackground('#173F35')
    .setFontColor('#FFFFFF')
    .setFontFamily('Arial')
    .setFontSize(22)
    .setFontWeight('bold')
    .setHorizontalAlignment('left')
    .setVerticalAlignment('middle');
  sheet.getRange('A2:K2')
    .setBackground('#EAF3EE')
    .setFontColor('#456258')
    .setFontFamily('Arial')
    .setFontSize(11)
    .setFontWeight('bold')
    .setHorizontalAlignment('left')
    .setVerticalAlignment('middle');
  sheet.getRange('A3:K3')
    .setBackground('#FBF2DE')
    .setFontColor('#7A591E')
    .setFontFamily('Arial')
    .setFontSize(10)
    .setFontWeight('bold')
    .setHorizontalAlignment('left')
    .setVerticalAlignment('middle');

  sheet.getRange('A4:C4')
    .setBackground('#DDEDE4')
    .setFontColor('#24553F')
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setBorder(true, true, true, true, false, false, '#B8D2C3', SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange('D4:K4')
    .setBackground('#D9A943')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setBorder(true, true, true, true, false, false, '#B88727', SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange('A5:K5')
    .setBackground('#456354')
    .setFontColor('#FFFFFF')
    .setFontFamily('Arial')
    .setFontSize(11)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  sheet.setRowHeight(1, 58);
  sheet.setRowHeight(2, 36);
  sheet.setRowHeight(3, 32);
  sheet.setRowHeight(4, 42);
  sheet.setRowHeight(5, 40);

  var widths = [62, 285, 180, 150, 95, 145, 95, 65, 105, 190, 210];
  widths.forEach(function(width, index) {
    sheet.setColumnWidth(index + 1, width);
  });

  if (rowCount > 0) {
    var dataRange = sheet.getRange(6, 1, rowCount, 11);
    dataRange
      .setFontFamily('Arial')
      .setFontSize(10)
      .setFontColor('#263A32')
      .setVerticalAlignment('middle')
      .setWrap(true);

    sheet.getRange(6, 1, rowCount, 1).setHorizontalAlignment('center');
    sheet.getRange(6, 5, rowCount, 1).setHorizontalAlignment('center');
    sheet.getRange(6, 6, rowCount, 1)
      .setHorizontalAlignment('center')
      .setNumberFormat('@');
    sheet.getRange(6, 7, rowCount, 1)
      .setHorizontalAlignment('right')
      .setNumberFormat('#,##0"원"');
    sheet.getRange(6, 8, rowCount, 1)
      .setHorizontalAlignment('center')
      .setNumberFormat('0');
    sheet.getRange(6, 9, rowCount, 1)
      .setHorizontalAlignment('right')
      .setNumberFormat('#,##0"원"');

    sheet.setRowHeights(6, rowCount, 48);

    var backgrounds = [];
    for (var rowIndex = 0; rowIndex < rowCount; rowIndex++) {
      var rowColor = rowIndex % 2 === 1 ? '#F4F8F5' : '#FFFFFF';
      backgrounds.push(new Array(11).fill(rowColor));
    }
    dataRange.setBackgrounds(backgrounds);
    dataRange.setBorder(
      false, false, true, false, false, false,
      '#DCE6DF', SpreadsheetApp.BorderStyle.SOLID
    );
  }

  unifiedDrawActionBars_(sheet);   // 체크박스 대신 클릭 막대
}


function unifiedMigrateSheetNamesAndLayout_() {
  var ss = seriesSpreadsheet_();

  var requestOld = ss.getSheetByName('희망도서 목록');
  var requestNew = ss.getSheetByName(UNIFIED_REQUEST_SHEET);
  if (requestOld && !requestNew) {
    requestOld.setName(UNIFIED_REQUEST_SHEET);
    requestNew = requestOld;
  }

  var holdingsOld = ss.getSheetByName('우리학교 소장도서');
  var holdingsNew = ss.getSheetByName(SERIES_HOLDINGS_SHEET);
  if (holdingsOld && !holdingsNew) {
    holdingsOld.setName(SERIES_HOLDINGS_SHEET);
    holdingsNew = holdingsOld;
  }

  var guideOld = ss.getSheetByName('사용설명서');
  var applicantGuide = ss.getSheetByName('📗 희망도서 신청안내');
  if (guideOld && !applicantGuide) {
    guideOld.setName('📗 희망도서 신청안내');
  }

  requestNew = ss.getSheetByName(UNIFIED_REQUEST_SHEET);
  if (requestNew) {
    var rowOne = requestNew.getRange(1, 1, 1, Math.max(1, requestNew.getLastColumn())).getDisplayValues()[0];
    var rowFive = requestNew.getMaxRows() >= 5
      ? requestNew.getRange(5, 1, 1, Math.max(1, requestNew.getLastColumn())).getDisplayValues()[0]
      : [];
    var rowOneLooksLikeHeader = rowOne.some(function(value) {
      return /순번|도서명|책제목|ISBN|처리상태/.test(String(value || ''));
    });
    var rowFiveLooksLikeHeader = rowFive.some(function(value) {
      return /순번|도서명|책제목|ISBN|처리상태/.test(String(value || ''));
    });

    if (rowOneLooksLikeHeader && !rowFiveLooksLikeHeader) {
      requestNew.insertRowsBefore(1, 4);
    }
  }
}

function unifiedStyleHoldingsSheet_(sheet) {
  if (!sheet || sheet.getLastColumn() < 1) return;

  var lastColumn = sheet.getLastColumn();
  var lastRow = Math.max(sheet.getLastRow(), 1);
  var bodyRows = Math.max(1, sheet.getMaxRows() - 1);

  sheet.setHiddenGridlines(true);
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, lastColumn)
    .setBackground('#2E5B4F')
    .setFontColor('#FFFFFF')
    .setFontFamily('Arial')
    .setFontSize(13)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sheet.setRowHeight(1, 42);

  sheet.getRange(2, 1, bodyRows, lastColumn)
    .setBackground('#FBFCFB')
    .setFontColor('#2E3A33')
    .setFontFamily('Arial')
    .setFontSize(11)
    .setVerticalAlignment('middle')
    .setWrap(true);

  if (lastRow > 1) {
    sheet.setRowHeights(2, lastRow - 1, 44);
  }
}

function unifiedStyleCopySetupSheet_(sheet) {
  if (!sheet) return;

  // 설정 버튼은 이미 상단에 있으므로 병합된 행을 다시 이동하지 않습니다.
  unifiedArrangeCopySetupButtons_(sheet);

  var lastRow = Math.max(sheet.getLastRow(), 1);
  var lastColumn = Math.max(sheet.getLastColumn(), 10);
  var canvas = sheet.getRange(1, 1, lastRow, lastColumn);
  var mergedRanges = canvas.getMergedRanges();

  sheet.setHiddenGridlines(true);
  sheet.setFrozenRows(0);
  sheet.setFrozenColumns(0);

  canvas
    .setBackground('#F8F5ED')
    .setFontFamily('Arial')
    .setFontColor('#2E3A33')
    .setFontSize(12)
    .setVerticalAlignment('middle')
    .setWrap(true);

  sheet.setColumnWidth(1, 28);
  sheet.setColumnWidth(2, 78);
  for (var column = 3; column <= lastColumn; column++) {
    sheet.setColumnWidth(column, column <= 10 ? 86 : 72);
  }
  sheet.setColumnWidths(2, 3, 100);
  sheet.setColumnWidth(5, 28);
  sheet.setColumnWidths(6, 3, 100);
  sheet.setColumnWidth(9, 28);
  sheet.setColumnWidths(10, 3, 100);

  sheet.setRowHeights(1, lastRow, 28);

  mergedRanges.forEach(function(range) {
    var text = String(range.getDisplayValue() || '').trim();

    range
      .setBackground('#FFFFFF')
      .setFontColor('#2E3A33')
      .setFontFamily('Arial')
      .setFontSize(12)
      .setVerticalAlignment('middle')
      .setWrap(true)
      .setBorder(
        true,true,true,true,false,false,
        '#D9E2DC',SpreadsheetApp.BorderStyle.SOLID
      );

    if (!text) {
      range.setBackground('#F8F5ED').setBorder(false,false,false,false,false,false);
      return;
    }

    if (/다른 학교.*최초 설정|사본용 최초 설정/.test(text)) {
      range
        .setBackground('#183D35')
        .setFontColor('#FFFFFF')
        .setFontSize(25)
        .setFontWeight('bold')
        .setHorizontalAlignment('center');
      return;
    }

    if (/처음 한 번만/.test(text)) {
      range
        .setBackground('#EAF3ED')
        .setFontColor('#34564D')
        .setFontSize(13)
        .setFontWeight('bold')
        .setHorizontalAlignment('center');
      return;
    }

    if (/컴퓨터가 익숙하지 않아도/.test(text)) {
      range
        .setBackground('#D9A441')
        .setFontColor('#FFFFFF')
        .setFontSize(17)
        .setFontWeight('bold')
        .setHorizontalAlignment('center');
      return;
    }

    if (/^\d+$/.test(text)) {
      range
        .setBackground('#D9A441')
        .setFontColor('#FFFFFF')
        .setFontSize(24)
        .setFontWeight('bold')
        .setHorizontalAlignment('center');
      return;
    }

    if (/사이트 열기|API 키 설정|학교 최초 설정|공개 신청 링크 등록|버튼 모음|메뉴 확인|이 탭 삭제/.test(text)) {
      range
        .setBackground('#DCEBE5')
        .setFontColor('#244D42')
        .setFontSize(13)
        .setFontWeight('bold')
        .setHorizontalAlignment('center')
        .setBorder(
          true,true,true,true,false,false,
          '#A8C8BB',SpreadsheetApp.BorderStyle.SOLID_MEDIUM
        );
      return;
    }

    if (/꼭 확인|주의|변경 금지|삭제하지/.test(text)) {
      range
        .setBackground('#FCEDEC')
        .setFontColor('#A33F3F')
        .setFontWeight('bold');
      return;
    }

    range
      .setBackground('#FFFFFF')
      .setFontColor('#2E3A33')
      .setFontSize(12);
  });

  for (var row = 1; row <= lastRow; row++) {
    var rowText = sheet.getRange(row, 1, 1, lastColumn).getDisplayValues()[0].join(' ').trim();
    if (rowText) sheet.setRowHeight(row, Math.max(sheet.getRowHeight(row), 32));
  }

  sheet.activate();
  sheet.setActiveRange(sheet.getRange('B3'));
}

function createUnifiedGuides_() {
  var applicantSteps = [
    ['1', '신청 링크로 들어가기', '선생님은 교직원용 링크, 학부모님은 학부모용 링크(또는 QR)로 들어옵니다. 링크는 「⚙️ 사본 설정 → 🔗 신청 링크 받기」에서 받을 수 있습니다.'],
    ['2', '신청하는 분 고르기', '링크 하나로 들어왔다면 첫 화면에서 「교직원」 또는 「학부모」를 누르십시오. 전용 링크로 들어왔다면 이 단계는 건너뜁니다. 교직원 인증번호를 켜 둔 학교라면 「교직원」을 누른 뒤 인증번호를 입력합니다.'],
    ['3', '신청자 정보 쓰기', '교직원은 이름을 입력합니다. 이름을 입력하면 우선 대출 혜택을 받을 수 있습니다. 학부모는 학년·반·번호를 꼭 고르고, 학생 이름은 선택입니다. 학부모 정보는 신청 목록에 보이지 않습니다.'],
    ['4', '책 검색', '책 제목 또는 ISBN을 입력하고 「책 검색하기」를 누릅니다.'],
    ['5', '책 확인·신청바구니 담기', '표지·저자·출판사·연도를 확인한 뒤 「신청바구니에 담기」를 누릅니다. 여러 권을 담을 수 있고, 한 번에 최대 20권까지 신청할 수 있습니다.'],
    ['6', '선택한 도서 신청', '신청바구니에서 「선택한 도서 신청하기」를 누릅니다. 학교에서 정한 신청 한도(예: 학부모는 한 학생당 5권)를 넘으면 안내 창이 뜹니다.'],
    ['7', '신청 확인·취소', '「신청 목록」에서 접수 결과를 확인합니다. 잘못 신청했다면 취소하고 바로 되돌릴 수도 있습니다. 교직원과 학부모의 신청 목록은 따로 보입니다.']
  ];

  unifiedBuildGuideSheet_(
    '📗 희망도서 신청안내',
    '📗 희망도서 신청안내',
    '교직원과 학부모가 책을 찾고 여러 권을 신청바구니에 담아 신청하는 방법입니다.',
    applicantSteps,
    '신청한 책이 이미 학교에 있거나 이미 신청된 책이면 신청되지 않습니다. 신청 한도는 학교에서 정합니다.'
  );

  if (
    typeof HR_getConfiguredPublicWebAppUrl_ === 'function' &&
    typeof HR_updateGuidePublicLink_ === 'function'
  ) {
    var savedPublicUrl = HR_getConfiguredPublicWebAppUrl_();
    if (savedPublicUrl) {
      HR_updateGuidePublicLink_(savedPublicUrl);
    }
  }

  HR_buildManualSheet_(
    '📘 전체 사용방법',
    HR_MANUAL_TITLE,
    HR_MANUAL_SUBTITLE,
    HR_MANUAL_SECTIONS,
    HR_MANUAL_FOOTER
  );
}

function unifiedBuildGuideSheet_(name, title, subtitle, steps, footer) {
  var ss = seriesSpreadsheet_();
  var sheet = ss.getSheetByName(name) || ss.insertSheet(name);
  var maxColumns = 12;
  var requiredRows = 6 + steps.length * 3;

  if (sheet.getFilter()) sheet.getFilter().remove();
  sheet.getBandings().forEach(function(b) { b.remove(); });
  sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).breakApart();
  sheet.clear();
  sheet.clearConditionalFormatRules();

  if (sheet.getMaxColumns() < maxColumns) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), maxColumns - sheet.getMaxColumns());
  }
  if (sheet.getMaxRows() < requiredRows) {
    sheet.insertRowsAfter(sheet.getMaxRows(), requiredRows - sheet.getMaxRows());
  }

  sheet.setHiddenGridlines(true);
  sheet.getRange('A1:L2')
    .merge()
    .setValue(title)
    .setBackground('#183D35')
    .setFontColor('#FFFFFF')
    .setFontFamily('Arial')
    .setFontSize(25)
    .setFontWeight('bold')
    .setHorizontalAlignment('left')
    .setVerticalAlignment('middle');

  sheet.getRange('A3:L3')
    .merge()
    .setValue(subtitle)
    .setBackground('#EAF3ED')
    .setFontColor('#34564D')
    .setFontFamily('Arial')
    .setFontSize(12)
    .setVerticalAlignment('middle');

  var startRow = 5;
  steps.forEach(function(step, index) {
    var row = startRow + index * 3;

    sheet.getRange(row, 1, 2, 2)
      .merge()
      .setValue(step[0])
      .setBackground('#D9A441')
      .setFontColor('#FFFFFF')
      .setFontFamily('Arial')
      .setFontSize(22)
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');

    sheet.getRange(row, 3, 1, 10)
      .merge()
      .setValue(step[1])
      .setBackground('#FFF4D6')
      .setFontColor('#684C18')
      .setFontFamily('Arial')
      .setFontSize(14)
      .setFontWeight('bold')
      .setVerticalAlignment('middle');

    sheet.getRange(row + 1, 3, 1, 10)
      .merge()
      .setValue(step[2])
      .setBackground('#FBFCFB')
      .setFontColor('#2E3A33')
      .setFontFamily('Arial')
      .setFontSize(12)
      .setWrap(true)
      .setVerticalAlignment('middle');

    sheet.getRange(row, 1, 2, 12).setBorder(
      true,true,true,true,false,false,
      '#E0D5B8',SpreadsheetApp.BorderStyle.SOLID
    );

    sheet.setRowHeight(row, 34);
    sheet.setRowHeight(row + 1, 52);
    sheet.setRowHeight(row + 2, 10);
  });

  var footerRow = startRow + steps.length * 3;
  sheet.getRange(footerRow, 1, 2, 12)
    .merge()
    .setValue(footer)
    .setBackground('#DCEBE5')
    .setFontColor('#244D42')
    .setFontFamily('Arial')
    .setFontSize(13)
    .setFontWeight('bold')
    .setWrap(true)
    .setVerticalAlignment('middle');

  for (var column = 1; column <= maxColumns; column++) {
    sheet.setColumnWidth(column, column <= 2 ? 54 : 96);
  }

  sheet.setRowHeight(1, 38);
  sheet.setRowHeight(2, 38);
  sheet.setRowHeight(3, 42);
  sheet.setFrozenRows(3);
}

function applyUnifiedLibraryDesign() {
  var ss = seriesSpreadsheet_();

  unifiedMigrateSheetNamesAndLayout_();
  createUnifiedGuides_();

  var request = ss.getSheetByName(UNIFIED_REQUEST_SHEET);
  if (request) styleRequestListSheet();

  var series = ss.getSheetByName(SERIES_FINDER_SHEET);
  if (series) {
    seriesStyleResultSheet_(series);
    seriesAddWorkButton_(series);
  }

  var review = unifiedEnsureReviewSheet_();
  unifiedStyleReviewSheet_(review);

  var uploadSheet = ss.getSheetByName(UNIFIED_UPLOAD_SHEET);
  if (uploadSheet) {
    unifiedStyleUploadSheet_(uploadSheet, Math.max(0, uploadSheet.getLastRow() - 5));
  }

  var exportSheet = ss.getSheetByName(UNIFIED_EXPORT_SHEET);
  if (exportSheet && exportSheet.getLastRow() >= 4) {
    unifiedStyleExportSheet_(exportSheet, Math.max(0, exportSheet.getLastRow() - 5));
  }

  unifiedStyleHoldingsSheet_(ss.getSheetByName(SERIES_HOLDINGS_SHEET));
  unifiedStyleCopySetupSheet_(ss.getSheetByName('⚙️ 사본 최초설정'));
  unifiedRefreshDashboard_();

  ss.toast('전체 시트의 제목·색상·글꼴을 같은 디자인으로 정리했습니다.','디자인 적용 완료',7);
}

function removeLegacyUnifiedSheetsAfterConfirmation() {
  var ss = seriesSpreadsheet_();
  ['⚙️ 통합 설정','📋 시리즈 업무목록'].forEach(function(name) {
    var sheet = ss.getSheetByName(name);
    if (sheet) ss.deleteSheet(sheet);
  });
  ss.toast('통합 설정과 시리즈 업무목록 시트를 삭제했습니다.','정리 완료',6);
}


function unifiedArrangeCopySetupButtons_(sheet) {
  if (!sheet || sheet.getLastRow() < 14) return;

  var lastColumn = Math.max(sheet.getLastColumn(), 15);
  var scanRows = Math.min(sheet.getLastRow(), 30);
  var values = sheet.getRange(1, 1, scanRows, lastColumn).getDisplayValues();
  var headerText = '🧰 설정 버튼 모음  ·  필요한 버튼을 순서대로 누르세요';
  var apiReady = false;
  var schoolReady = false;
  var publicReady = false;

  values.forEach(function(row) {
    row.forEach(function(cellValue) {
      var cellText = String(cellValue || '').trim();
      if (/설정 버튼 모음/.test(cellText)) headerText = cellText;
      if (/API 키 설정/.test(cellText) && /✅/.test(cellText)) apiReady = true;
      if (/(?:이 학교|학교) 최초 설정/.test(cellText) && /✅/.test(cellText)) schoolReady = true;
      if (/공개 (?:신청 )?링크 등록/.test(cellText) && /✅/.test(cellText)) publicReady = true;
    });
  });

  sheet.getRange('B7:O20').breakApart().clearContent();
  sheet.getRange('B7:L8').merge().setValue(headerText);
  sheet.getRange('B10:D13').merge().setValue((apiReady ? '✅ ' : '') + '🔑 API 키 설정');
  sheet.getRange('F10:H13').merge().setValue((schoolReady ? '✅ ' : '') + '🏫 학교 최초 설정');
  sheet.getRange('J10:L13').merge().setValue((publicReady ? '✅ ' : '') + '🔗 공개 링크 등록');

  sheet.setColumnWidths(2, 3, 100);
  sheet.setColumnWidth(5, 28);
  sheet.setColumnWidths(6, 3, 100);
  sheet.setColumnWidth(9, 28);
  sheet.setColumnWidths(10, 3, 100);

  [10, 11, 12, 13].forEach(function(rowNumber) {
    sheet.setRowHeight(rowNumber, 27);
  });
  sheet.getRange('B14:L20')
    .setBackground('#F8F5ED')
    .setBorder(false, false, false, false, false, false);
  sheet.setRowHeight(9, 8);
  for (var compactRow = 14; compactRow <= 20; compactRow++) {
    sheet.setRowHeight(compactRow, 4);
  }

  sheet.getImages().forEach(function(image) {
    var title = String(image.getAltTextTitle() || '');
    var description = String(image.getAltTextDescription() || '');
    if (title.indexOf('HR_COPY_SETUP_BUTTON') !== 0) return;

    if (/API 키/.test(description)) {
      image.setAnchorCell(sheet.getRange('B10')).setWidth(300).setHeight(104);
    } else if (/학교 최초 설정/.test(description)) {
      image.setAnchorCell(sheet.getRange('F10')).setWidth(300).setHeight(104);
    } else if (/공개 신청 링크|공개 링크/.test(description)) {
      image.setAnchorCell(sheet.getRange('J10')).setWidth(300).setHeight(104);
    }
  });
}

function unifiedMoveCopySetupButtonsToTop_(sheet) {
  if (!sheet || sheet.getLastRow() < 8) return;

  var lastRow = sheet.getLastRow();
  var lastColumn = Math.max(sheet.getLastColumn(), 1);
  var values = sheet.getRange(1, 1, lastRow, lastColumn).getDisplayValues();
  var startRow = 0;
  var endRow = 0;

  values.forEach(function(row, index) {
    var text = row.join(' ').trim();

    if (!startRow && /설정 버튼 모음/.test(text)) {
      startRow = index + 1;
    }

    if (/공개 신청 링크 등록/.test(text)) {
      endRow = Math.max(endRow, index + 1);
    }
  });

  if (!startRow || !endRow || endRow < startRow || startRow <= 8) {
    return;
  }

  var endMerged = sheet
    .getRange(endRow, 1, 1, lastColumn)
    .getMergedRanges();

  endMerged.forEach(function(range) {
    endRow = Math.max(endRow, range.getLastRow());
  });

  endRow = Math.min(lastRow, endRow + 1);

  sheet.moveRows(
    sheet.getRange(
      startRow,
      1,
      endRow - startRow + 1,
      sheet.getMaxColumns()
    ),
    7
  );
}


function seriesLeadingCore_(value) {
  var clean = seriesCleanText_(value);
  if (!clean) return null;

  var stripped = clean.replace(/^\s*(?:[\(\[【][^)\]】]{2,60}[\)\]】]\s*)+/, '').trim();
  if (!stripped || stripped === clean) return null;

  var parsed = seriesParseTitle_(stripped);
  if (!parsed || parsed.key.length < 2) return null;
  return parsed;
}


function seriesFamilySearchText_(value) {
  var text = seriesCleanText_(value);
  if (!text) return '';

  text = text
    .replace(/(?:학습\s*)?만화로\s*(?:보는|읽는|배우는)/gi, ' ')
    .replace(/그림으로\s*(?:보는|읽는|배우는)/gi, ' ')
    .replace(/(?:어린이|초등학생|청소년)을?\s*위한/gi, ' ')
    .replace(/학습만화|그래픽\s*노블|그림책|코믹스?|양장본?|반양장|보드북|팝업북|리커버|개정판|특별판/gi, ' ')
    .replace(/\s+/g, ' ')
    .trim();

  return seriesNormalizeKey_(text).length >= 4 ? text : seriesCleanText_(value);
}


function seriesFamilyKey_(value) {
  return seriesNormalizeKey_(seriesFamilySearchText_(value));
}


function seriesFormatKey_(value) {
  var key = seriesNormalizeKey_(value);
  var tags = [];

  if (/학습만화/.test(key)) tags.push('학습만화');
  else if (/만화로보는|만화로읽는|코믹|만화/.test(key)) tags.push('만화');

  if (/그림으로보는|그림으로읽는|그림책|그래픽노블/.test(key)) tags.push('그림');
  if (/양장|하드커버/.test(key)) tags.push('양장');
  if (/반양장|소프트커버/.test(key)) tags.push('반양장');
  if (/보드북/.test(key)) tags.push('보드북');
  if (/팝업북/.test(key)) tags.push('팝업북');
  if (/큰글자|대활자/.test(key)) tags.push('큰글자');
  if (/리커버/.test(key)) tags.push('리커버');
  if (/개정판|개정증보/.test(key)) tags.push('개정판');
  if (/컬러|올컬러/.test(key)) tags.push('컬러');

  return tags.length ? tags.join('+') : '일반';
}

/* ---------- 🏫 소장도서 「출판사」 칸 ---------- */

// 소장도서 탭 머리줄에 「출판사」 칸이 없으면 바로 오른쪽 빈 칸에 만듭니다. (이미 있으면 아무것도 하지 않습니다)
// 시리즈 누락 점검(seriesReadHoldings_)이 이 칸을 읽어 같은 이름의 다른 출판사 시리즈를 구분합니다.
function unifiedEnsureHoldingsPublisherColumn_(ss) {
  ss = ss || seriesSpreadsheet_();
  var sheet = ss.getSheetByName(SERIES_HOLDINGS_SHEET);
  if (!sheet) return false;
  var lastRow = sheet.getLastRow();
  var lastColumn = sheet.getLastColumn();
  if (lastRow < 1 || lastColumn < 1) return false;

  // 머리줄 찾기: 맨 위 10줄 안에서 ISBN 칸이 있는 줄 (seriesReadHoldings_ 와 같은 방식)
  var grid = sheet.getRange(1, 1, Math.min(10, lastRow), lastColumn).getDisplayValues();
  var headerRow = 0;
  var lastHeaderColumn = 0;
  for (var r = 0; r < grid.length && !headerRow; r++) {
    var hasIsbn = grid[r].some(function(value) { return /^ISBN(13)?$/.test(String(value || '').replace(/\s+/g, '').toUpperCase()); });
    if (!hasIsbn) continue;
    headerRow = r + 1;
    for (var c = 0; c < grid[r].length; c++) {
      var header = String(grid[r][c] || '').replace(/\s+/g, '');
      if (/^(출판사|발행자|출판자|발행처)$/.test(header)) return false;   // 이미 있음
      if (header) lastHeaderColumn = c + 1;
    }
  }
  if (!headerRow || !lastHeaderColumn) return false;

  var column = lastHeaderColumn + 1;
  if (sheet.getMaxColumns() < column) sheet.insertColumnsAfter(sheet.getMaxColumns(), column - sheet.getMaxColumns());
  var target = sheet.getRange(headerRow, column);
  if (String(target.getDisplayValue() || '').trim()) return false;   // 그 칸에 다른 글이 있으면 건드리지 않습니다.

  sheet.getRange(headerRow, lastHeaderColumn).copyFormatToRange(sheet, column, column, headerRow, headerRow);
  target.setValue('출판사').setNote('DLS 소장 목록의 출판사를 붙여 넣어 주세요. (도서명·ISBN과 같은 줄 순서) 시리즈 누락 점검이 출판사로 시리즈를 구분합니다.');
  try { sheet.setColumnWidth(column, Math.max(sheet.getColumnWidth(lastHeaderColumn), 160)); } catch (error) {}

  // 머리줄에 필터가 걸려 있으면 새 칸까지 넓힙니다.
  try {
    var filter = sheet.getFilter();
    if (filter) {
      var range = filter.getRange();
      if (range.getRow() === headerRow && range.getColumn() <= lastHeaderColumn && range.getLastColumn() < column) {
        var top = range.getRow(), left = range.getColumn(), rows = Math.max(range.getNumRows(), sheet.getLastRow() - top + 1);
        filter.remove();
        sheet.getRange(top, left, rows, column - left + 1).createFilter();
      }
    }
  } catch (error) {
    console.warn('소장도서 필터 넓히기 실패: ' + error);
  }
  return true;
}
