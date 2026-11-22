/**
 * 희망도서 신청 사이드바
 *
 * 기존 Code.gs의 processRow(), getHeaders(), getColumnMap()을 그대로 사용해
 * YES24 검색, 소장도서·중복 확인, 청구기호 자동입력을 실행합니다.
 */

var HR_TARGET_SHEET_NAME = '📚 희망도서 목록';
var HR_LEGACY_TARGET_SHEET_NAME = '2차 희망도서';
var HR_LIBRARY_SHEET_NAME = '🏫 우리학교 소장도서';
var HR_COPY_SETUP_SHEET_NAME = '⚙️ 사본 최초설정';
var HR_SIDEBAR_TRIGGER_FUNCTION = 'HR_openSidebarOnSpreadsheetOpen';
var HR_SPREADSHEET_ID_PROPERTY = 'HR_TARGET_SPREADSHEET_ID';
var HR_YES24_KEY_PROPERTY = 'HR_YES24_API_KEY';
var HR_NLK_KEY_PROPERTY = 'HR_NLK_API_KEY';
var HR_SCHOOL_NAME_PROPERTY = 'HR_SCHOOL_NAME';
var HR_MANAGER_NAME_PROPERTY = 'HR_MANAGER_NAME';
var HR_MANAGER_EXTENSION_PROPERTY = 'HR_MANAGER_EXTENSION';
var HR_SETUP_COMPLETE_PROPERTY = 'HR_SETUP_COMPLETE';
var HR_PUBLIC_WEB_APP_URL_PROPERTY = 'HR_PUBLIC_WEB_APP_URL';
var HR_REQUEST_BACKUP_SHEET_NAME = '🔒 신청접수백업';
var HR_REQUEST_CANCEL_BACKUP_SHEET_NAME = '🔒 신청취소복구';
var HR_REQUEST_LOCK_WAIT_MS = 30000;
var HR_YES24_API_MANAGE_URL =
  'https://developers.yes24.com/';

/**
 * 권한 승인 없이 사용하는 공개 신청 웹앱 화면입니다.
 */
function doGet(e) {
  // 주소가 열리는지 확인하는 요청(?probe=1)에는 화면을 만들지 않고 짧게 답합니다.
  if (e && e.parameter && e.parameter.probe === '1') {
    return ContentService.createTextOutput('ok');
  }
  var properties = PropertiesService.getScriptProperties();
  var audience = HR_audienceFromParameter_(e && e.parameter && e.parameter.type);
  var schoolName = String(
    properties.getProperty(HR_SCHOOL_NAME_PROPERTY) || ''
  ).trim();
  var managerName = String(
    properties.getProperty(HR_MANAGER_NAME_PROPERTY) || ''
  ).trim();
  var managerExtension = String(
    properties.getProperty(HR_MANAGER_EXTENSION_PROPERTY) || ''
  ).trim();
  var siteTitle = schoolName
    ? schoolName + ' 희망도서 신청'
    : '희망도서 신청';
  var template = HtmlService.createTemplateFromFile(
    'RequestSidebar'
  );

  template.siteTitle = siteTitle;
  template.audience = audience;
  template.fromGate = String(e && e.parameter && e.parameter.g || '') === '1';
  template.baseUrl = HR_getWebAppBaseUrl_();
  template.buildId = HR_BUILD_ID;
  template.limitStaff = HR_getRequestLimit_('staff');
  template.limitParent = HR_getRequestLimit_('parent');
  template.staffCodeRequired = HR_isStaffCodeEnabled_();
  template.maxClasses = HR_getMaxClasses_();
  template.managerName = managerName;
  template.managerExtension = managerExtension;

  return template
    .evaluate()
    .setTitle('📚 ' + siteTitle)
    .addMetaTag(
      'viewport',
      'width=device-width, initial-scale=1'
    );
}

// 이 코드가 만들어진 때(빌드가 채워 줍니다). 화면에 보여 주고, 열어 둔 화면이 옛 버전인지 알아보는 데 씁니다.
var HR_BUILD_ID = '__BUILD_ID__';

function HR_getBuildId() {
  return HR_BUILD_ID;
}

// 웹 앱 주소가 실제로 열리는지 확인합니다.
// 'ok' 열림 · 'notfound' 없는 주소(404) · 'restricted' 로그인이 필요함(공개 아님) · 'unknown' 확인 못 함
function HR_probeWebAppUrl_(url) {
  try {
    var response = UrlFetchApp.fetch(String(url) + '?probe=1', {
      muteHttpExceptions: true,
      followRedirects: false
    });
    var code = response.getResponseCode();
    if (code === 404) return 'notfound';
    if (code === 401 || code === 403) return 'restricted';
    if (code >= 300 && code < 400) {
      var headers = response.getHeaders() || {};
      var location = String(headers.Location || headers.location || '');
      return /accounts\.google\.com/.test(location) ? 'restricted' : 'ok';
    }
    return code >= 200 && code < 300 ? 'ok' : 'unknown';
  } catch (error) {
    return 'unknown';
  }
}

// 확인 결과를 사용자에게 보여 줄 말로 바꿉니다. (열리면 빈 글자)
function HR_probeProblemMessage_(status) {
  if (status === 'notfound') {
    return '이 주소로는 접속이 안 돼요(페이지를 찾을 수 없음). 주소가 중간에 잘렸거나 지워진 배포일 수 있어요. ' +
      'Apps Script 「배포 → 배포 관리」에서 현재 웹 앱 주소를 처음부터 끝까지 다시 복사해 주세요.';
  }
  if (status === 'restricted') {
    return '이 주소는 로그인이 필요해요. Apps Script 「배포 → 배포 관리」에서 웹 앱의 액세스 권한을 「모든 사용자」로 바꿔 주세요.';
  }
  return '';
}

// 웹 앱 주소(…/exec) 모양인지 확인합니다.
function HR_isWebAppExecUrl_(url) {
  return /^https:\/\/script\.google\.com\/(?:a\/macros\/[^/]+|macros)\/s\/[^/]+\/exec$/.test(String(url || ''));
}

// 「공개 신청 링크 등록」에 저장한 주소(검사를 통과한 것만). 없으면 빈 글자입니다.
function HR_getRegisteredWebAppUrl_() {
  var saved = String(
    PropertiesService.getScriptProperties().getProperty(HR_PUBLIC_WEB_APP_URL_PROPERTY) || ''
  ).trim().replace(/[?#].*$/, '');
  return HR_isWebAppExecUrl_(saved) ? saved : '';
}

// 구글이 알려 주는 배포 주소. 못 구하면 빈 글자입니다.
function HR_getDetectedWebAppUrl_() {
  try {
    var url = String(ScriptApp.getService().getUrl() || '');
    return HR_isWebAppExecUrl_(url) ? url : '';
  } catch (error) {
    return '';
  }
}

// 이 웹 앱의 주소. 「공개 신청 링크 등록」에 넣은 주소가 항상 먼저이고, 없을 때만 구글이 알려 주는 주소를 씁니다.
function HR_getWebAppBaseUrl_() {
  return HR_getRegisteredWebAppUrl_() || HR_getDetectedWebAppUrl_();
}

/**
 * 최초 설정에서 저장한 학교 스프레드시트를 사용합니다.
 */
function HR_getSpreadsheet_() {
  
  var properties = PropertiesService.getScriptProperties();
  var spreadsheetId = String(
    properties.getProperty(HR_SPREADSHEET_ID_PROPERTY) || ''
  ).trim();

  if (!spreadsheetId) {
    throw new Error(
      '학교 최초 설정이 필요합니다. ' +
      '⚙️ 사본 최초설정 탭에서 설정을 완료해 주세요.'
    );
  }

  return SpreadsheetApp.openById(spreadsheetId);
}

/**
 * 사용설명서 버튼에는 소유자 전용 /dev 주소가 아니라
 * 누구나 접속 가능한 공개 /exec 주소를 우선 사용합니다.
 */
function HR_getConfiguredPublicWebAppUrl_() {
  var properties = PropertiesService.getScriptProperties();
  var savedUrl = String(
    properties.getProperty(HR_PUBLIC_WEB_APP_URL_PROPERTY) || ''
  ).trim();

  return savedUrl;
}

/**
 * 이전 사본의 시트 이름을 새 이름으로 한 번만 바꿉니다.
 */
function HR_renameLegacyTargetSheet() {
  var spreadsheet = HR_getSetupSpreadsheet_();
  var currentSheet = spreadsheet.getSheetByName(
    HR_TARGET_SHEET_NAME
  );
  var legacySheet;

  if (currentSheet) {
    return {
      success: true,
      message: '희망도서 목록 시트 이름이 이미 적용되어 있습니다.'
    };
  }

  legacySheet = spreadsheet.getSheetByName(
    HR_LEGACY_TARGET_SHEET_NAME
  );

  if (!legacySheet) {
    throw new Error(
      '희망도서 목록으로 바꿀 기존 시트를 찾지 못했습니다.'
    );
  }

  legacySheet.setName(HR_TARGET_SHEET_NAME);

  return {
    success: true,
    message: '시트 이름을 「희망도서 목록」으로 변경했습니다.'
  };
}

/**
 * 학교별로 저장한 YES24 API Key를 가져옵니다.
 * 기존 프로젝트는 Code.gs의 YES24_KEY를 예비값으로 사용합니다.
 */
function HR_getYes24Key_() {
  var properties = PropertiesService.getScriptProperties();
  var savedKey = String(
    properties.getProperty(HR_YES24_KEY_PROPERTY) || ''
  ).trim();

  if (savedKey) {
    return savedKey;
  }

  return (
    typeof YES24_KEY !== 'undefined'
      ? String(YES24_KEY || '').trim()
      : ''
  );
}

/**
 * 국립중앙도서관 API 키를 가져옵니다.
 * KDC 참고용 선택 기능이며, 키가 없으면 빈 값으로 둡니다.
 */
function HR_getNlkKey_() {
  var properties = PropertiesService.getScriptProperties();
  var savedKey = String(
    properties.getProperty(HR_NLK_KEY_PROPERTY) || ''
  ).trim();
  var legacyKey = '';

  if (savedKey) {
    return savedKey;
  }

  // 코드에 적혀 있던 기본 키는 처음 쓸 때 스크립트 속성으로 옮깁니다.
  legacyKey = typeof NLK_KEY !== 'undefined'
    ? String(NLK_KEY || '').trim()
    : '';
  if (legacyKey) {
    properties.setProperty(HR_NLK_KEY_PROPERTY, legacyKey);
  }
  return legacyKey;
}

/**
 * 희망도서 신청 사이드바를 엽니다.
 */
function HR_showRequestSidebar() {
  var html = HtmlService
    .createHtmlOutputFromFile('RequestSidebar')
    .setTitle('📚 희망도서 신청');

  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * 스프레드시트가 열릴 때 실행되는 설치형 트리거입니다.
 * 신청 메뉴를 추가하고 희망도서 목록 탭이면 사이드바를 자동으로 엽니다.
 */
function HR_openSidebarOnSpreadsheetOpen() {
  // 예전에는 시트를 열 때마다 「희망도서 신청 바로가기」 창을 띄웠습니다.
  // 지금은 창을 띄우지 않고, 남아 있는 옛 트리거를 지웁니다.
  HR_removeLegacyOpenTrigger_();
}

/**
 * 옛 사본에 남은 「열 때 신청 창 띄우기」 트리거를 지웁니다.
 */
function HR_removeLegacyOpenTrigger_() {
  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    if (trigger.getHandlerFunction() === HR_SIDEBAR_TRIGGER_FUNCTION) {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}

/**
 * 자동 열기 트리거를 한 번만 설치합니다.
 */
function HR_installSidebarOpenTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  var triggerIndex = 0;

  for (
    triggerIndex = 0;
    triggerIndex < triggers.length;
    triggerIndex++
  ) {
    if (
      triggers[triggerIndex].getHandlerFunction() ===
      HR_SIDEBAR_TRIGGER_FUNCTION
    ) {
      ScriptApp.deleteTrigger(triggers[triggerIndex]);
    }
  }

  HR_addCopySetupMenu();

  SpreadsheetApp.getActiveSpreadsheet().toast(
    '희망도서 신청은 공개 웹페이지에서 이용합니다.',
    '도서 신청',
    5
  );
}

/**
 * 신청자와 제목·ISBN을 새 행에 넣고 기존 자동입력을 실행합니다.
 */
function HR_submitBookRequest(formData) {
  var applicant = HR_cleanInput_(
    formData && formData.applicant,
    40
  );
  var requestAudience = HR_audienceKey_(formData && formData.audience);
  HR_requireStaffAccess_(requestAudience, formData && formData.staffCode);
  var configuredManagerName = String(
    PropertiesService.getScriptProperties()
      .getProperty(HR_MANAGER_NAME_PROPERTY) || ''
  ).trim();
  var isManagerApplicant = !!configuredManagerName &&
    applicant.toLowerCase() === configuredManagerName.toLowerCase();
  var query = String(
    formData && formData.query || ''
  ).trim();
  var selectedBook =
    formData && formData.book &&
    typeof formData.book === 'object'
      ? formData.book
      : {};
  // 브라우저가 보낸 도서 정보는 믿지 않고, 정리한 뒤 YES24에서 다시 확인한 값으로 바꿉니다.
  selectedBook = HR_prepareSelectedBook_(selectedBook, query);
  var selectedPrice = Number(
    selectedBook.price ||
    selectedBook.priceStandard ||
    selectedBook.priceStd ||
    selectedBook.priceSales ||
    0
  );
  var selectedYear = String(
    selectedBook.year ||
    selectedBook.pubDate ||
    selectedBook.publicationDate ||
    ''
  ).match(/\d{4}/);
  var normalizedIsbn = query.replace(/[^0-9Xx]/g, '');
  var requestId = HR_cleanInput_(
    formData && formData.requestId,
    80
  ) || Utilities.getUuid();
  var isIsbn = (
    normalizedIsbn.length === 10 ||
    normalizedIsbn.length === 13
  );
  var spreadsheet = HR_getSpreadsheet_();
  var sheet = spreadsheet.getSheetByName(HR_TARGET_SHEET_NAME);
  var documentLock = LockService.getScriptLock();
  var headers = [];
  var columns = {};
  var applicantColumn = 11;
  var sequenceColumn = 1;
  var row = 0;
  var previousRow = 0;
  var previousRowHeight = 0;
  var sequenceValues = [];
  var nextSequence = 1;
  var sequenceIndex = 0;
  var ownedBook = null;
  var requestedBook = null;
  var lockAcquired = false;
  var backupRow = 0;
  var fastBook = {};
  var responseResult = {};
  var priceColumn = 0;
  var yearColumn = 0;

  if (!query) {
    throw new Error('책 제목이나 ISBN을 입력해 주세요.');
  }

  if (!sheet) {
    throw new Error(
      HR_TARGET_SHEET_NAME + ' 시트를 찾지 못했습니다.'
    );
  }

  // 신청 전 학교 소장도서 여부를 먼저 확인합니다.
  ownedBook = HR_findOwnedBook_(
    spreadsheet,
    selectedBook
  );

  if (ownedBook.found && !isManagerApplicant) {
    return {
      success: false,
      owned: true,
      message: '이미 소장중인 도서입니다.',
      title: ownedBook.title || selectedBook.title || query,
      isbn: ownedBook.isbn || selectedBook.isbn || normalizedIsbn,
      status: '이미 소장중인 도서입니다.'
    };
  }

  try {
    // 동시 신청 시 행이 겹치지 않도록 저장 구간만 잠급니다.
    lockAcquired = documentLock.tryLock(
      HR_REQUEST_LOCK_WAIT_MS
    );

    if (!lockAcquired) {
      throw new Error(
        '현재 신청자가 많아 잠시 대기 중입니다. 자동으로 다시 시도해 주세요.'
      );
    }

    headers = getHeaders(sheet);
    columns = getColumnMap(headers);

    // 원본 신청 내용은 숨김 백업 시트에 먼저 남깁니다.
    backupRow = HR_appendRequestBackup_(
      spreadsheet,
      {
        requestId: requestId,
        applicant: applicant,
        query: query,
        title: selectedBook.title || '',
        isbn: selectedBook.isbn || normalizedIsbn,
        status: '접수 확인 중'
      }
    );

    requestedBook = HR_findRequestedBook_(
      sheet,
      selectedBook,
      headers,
      columns
    );

    if (requestedBook.found) {
      HR_updateRequestBackup_(
        spreadsheet,
        backupRow,
        '이미 신청됨',
        '',
        ''
      );
      return {
        success: false,
        duplicateRequest: true,
        message: '이미 신청된 도서입니다.',
        title: requestedBook.title || selectedBook.title || query,
        isbn: requestedBook.isbn || selectedBook.isbn || normalizedIsbn,
        status: '이미 신청된 도서입니다.'
      };
    }

    applicantColumn = HR_findHeaderColumn_(
      headers,
      '신청자',
      11
    );
    sequenceColumn = HR_findHeaderColumn_(
      headers,
      '순번',
      1
    );
    row = Math.max(sheet.getLastRow() + 1, FIRST_DATA_ROW);
    previousRow = Math.max(FIRST_DATA_ROW, row - 1);
    previousRowHeight = sheet.getRowHeight(previousRow);

    if (row > FIRST_DATA_ROW) {
      sheet
        .getRange(
          previousRow,
          1,
          1,
          sheet.getLastColumn()
        )
        .copyTo(
          sheet.getRange(
            row,
            1,
            1,
            sheet.getLastColumn()
          ),
          SpreadsheetApp.CopyPasteType.PASTE_FORMAT,
          false
        );

      sheet.setRowHeight(row, previousRowHeight);

      sequenceValues = sheet
        .getRange(FIRST_DATA_ROW, sequenceColumn, row - FIRST_DATA_ROW, 1)
        .getValues()
        .flat();

      for (
        sequenceIndex = 0;
        sequenceIndex < sequenceValues.length;
        sequenceIndex++
      ) {
        var sequenceNumber = Number(
          sequenceValues[sequenceIndex]
        );

        if (
          isFinite(sequenceNumber) &&
          sequenceNumber >= nextSequence
        ) {
          nextSequence = sequenceNumber + 1;
        }
      }
    }

    sheet
      .getRange(row, sequenceColumn)
      .setValue(nextSequence);

    HR_setSafeText_(
      sheet.getRange(row, applicantColumn),
      applicant
    );
    sheet
      .getRange(row, HR_ensureAudienceColumn_(sheet))
      .setValue(HR_audienceLabel_(requestAudience));

    // 검색 화면에서 이미 받은 YES24 결과를 그대로 사용합니다.
    // 같은 책을 다시 API로 조회하지 않아 접수 시간을 줄입니다.
    fastBook = {
      title: selectedBook.title || (!isIsbn ? query : ''),
      author: selectedBook.author || '',
      publisher: selectedBook.publisher || '',
      isbn: selectedBook.isbn || normalizedIsbn,
      cover: selectedBook.cover || '',
      priceStd: selectedPrice,
      year: selectedYear ? selectedYear[0] : '',
      category: selectedBook.category || '',
      callNumber: selectedBook.callNumber || ''
    };

    fillFinalData(
      sheet,
      row,
      fastBook,
      headers,
      '✅ 접수 완료' + (ownedBook.found && isManagerApplicant ? ' · 관리자 소장도서 예외' : '') + HR_unverifiedNote_(selectedBook)
    );

    // 웹 신청에서도 정가와 출판연도가 반드시 시트에 남도록
    // 실제 머리글을 다시 찾아 직접 기록합니다.
    priceColumn = HR_findAnyHeaderColumn_(
      headers,
      ['정가', '도서정가', '정가(원)', '가격']
    );
    yearColumn = HR_findAnyHeaderColumn_(
      headers,
      ['연도', '출판연도', '발행연도', '출간연도']
    );

    if (priceColumn && selectedPrice > 0) {
      sheet
        .getRange(row, priceColumn)
        .setValue(selectedPrice)
        .setNumberFormat('#,##0');
    }

    if (yearColumn && fastBook.year) {
      sheet
        .getRange(row, yearColumn)
        .setValue(fastBook.year)
        .setNumberFormat('0');
    }

    SpreadsheetApp.flush();

    responseResult = {
      success: true,
      row: row,
      title: fastBook.title || query,
      isbn: fastBook.isbn || normalizedIsbn,
      author: fastBook.author,
      publisher: fastBook.publisher,
      price: selectedPrice,
      year: fastBook.year,
      status: '✅ 접수 완료' + (ownedBook.found && isManagerApplicant ? ' · 관리자 소장도서 예외' : '') + HR_unverifiedNote_(selectedBook),
      requestId: requestId
    };

    HR_updateRequestBackup_(
      spreadsheet,
      backupRow,
      '처리 완료' + (ownedBook.found && isManagerApplicant ? ' · 관리자 소장도서 예외' : ''),
      row,
      ''
    );

    if (isManagerApplicant) {
      try {
        sendHopeRequestsToReview(row);
        buildUnifiedBookstoreExportSheet();
      } catch (routeError) {
        console.error('관리자 희망도서 자동 전달 오류', routeError);
      }
    }

    return responseResult;
  } catch (error) {
    if (
      row > 0 &&
      columns.status
    ) {
      sheet
        .getRange(row, columns.status)
        .setValue(
          '❌ 오류: ' +
          String(error.message || error)
        );
    }

    if (backupRow > 0) {
      HR_updateRequestBackup_(
        spreadsheet,
        backupRow,
        '접수 오류',
        row,
        String(error.message || error)
      );
    }

    throw error;
  } finally {
    if (lockAcquired) {
      documentLock.releaseLock();
    }
  }
}

/**
 * 신청 원본을 별도 숨김 시트에 남깁니다.
 * 자동입력 도중 오류가 나도 신청자·제목·ISBN을 복구할 수 있습니다.
 */
function HR_appendRequestBackup_(spreadsheet, requestData) {
  var sheet = spreadsheet.getSheetByName(
    HR_REQUEST_BACKUP_SHEET_NAME
  );
  var headers = [[
    '접수시각',
    '요청ID',
    '신청자',
    '검색어',
    '도서명',
    'ISBN13',
    '처리상태',
    '희망도서행',
    '오류내용'
  ]];
  var row;

  if (!sheet) {
    sheet = spreadsheet.insertSheet(
      HR_REQUEST_BACKUP_SHEET_NAME
    );
    sheet.getRange(1, 1, 1, headers[0].length)
      .setValues(headers)
      .setFontWeight('bold')
      .setBackground('#E8EEF8');
        sheet.setFrozenRows(1);
    sheet.hideSheet();
      }


  row = Math.max(sheet.getLastRow() + 1, 2);
  // 신청자·검색어·도서명·ISBN은 텍스트로 저장해 수식으로 실행되지 않게 합니다.
  sheet.getRange(row, 1, 1, 9).setValues([[
    new Date(),
    '',
    '',
    '',
    '',
    '',
    requestData.status || '접수',
    '',
    ''
  ]]);
  HR_setSafeText_(sheet.getRange(row, 2), requestData.requestId);
  HR_setSafeText_(sheet.getRange(row, 3), requestData.applicant);
  HR_setSafeText_(sheet.getRange(row, 4), requestData.query);
  HR_setSafeText_(sheet.getRange(row, 5), requestData.title);
  HR_setSafeText_(sheet.getRange(row, 6), requestData.isbn);
  sheet.getRange(row, 1).setNumberFormat(
    'yyyy. m. d. HH:mm:ss'
  );
  return row;
}

function HR_updateRequestBackup_(
  spreadsheet,
  backupRow,
  status,
  targetRow,
  errorMessage
) {
  var sheet;
  if (!backupRow) return;
  sheet = spreadsheet.getSheetByName(
    HR_REQUEST_BACKUP_SHEET_NAME
  );
  if (!sheet) return;

  sheet.getRange(backupRow, 7, 1, 2).setValues([[
    status || '',
    targetRow || ''
  ]]);
  HR_setSafeText_(sheet.getRange(backupRow, 9), errorMessage);
}

/**
 * 장바구니에서 고른 여러 권을 순서대로 신청합니다.
 * 한 번에 최대 20권까지 처리하고, 성공·중복·소장 도서를 각각 알려줍니다.
 */
function HR_submitBookRequests(formData) {
  var applicant = String(
    formData && formData.applicant || ''
  ).trim();
  var selectedBooks =
    formData && Array.isArray(formData.books)
      ? formData.books
      : [];
  var requestAudience = HR_audienceKey_(formData && formData.audience);
  HR_requireStaffAccess_(requestAudience, formData && formData.staffCode);
  var results = [];
  var bookIndex = 0;
  var successCount = 0;
  var duplicateCount = 0;
  var ownedCount = 0;
  var errorCount = 0;
  var requestIdBase = String(
    formData && formData.requestId || Utilities.getUuid()
  ).trim();

  if (!selectedBooks.length) {
    throw new Error('신청할 책을 한 권 이상 선택해 주세요.');
  }

  if (selectedBooks.length > 20) {
    throw new Error(
      '한 번에 최대 20권까지 신청할 수 있습니다.'
    );
  }

  if (requestAudience === 'parent' && !HR_parseParentApplicant_(applicant)) {
    throw new Error('학년·반·번호를 모두 골라 주십시오. (예: 3학년 2반 15번)');
  }

  var limitMessage = HR_requestLimitMessage_(
    requestAudience,
    applicant,
    selectedBooks.length
  );
  if (limitMessage) {
    throw new Error(limitMessage);
  }

  for (
    bookIndex = 0;
    bookIndex < selectedBooks.length;
    bookIndex++
  ) {
    var selectedBook = selectedBooks[bookIndex] || {};
    var result;

    try {
      result = HR_submitBookRequest({
        applicant: applicant,
        audience: requestAudience,
        query: selectedBook.isbn || selectedBook.title,
        book: selectedBook,
        requestId: requestIdBase + '-' + bookIndex
      });

      if (result.success) {
        successCount++;
      } else if (result.owned) {
        ownedCount++;
      } else if (result.duplicateRequest) {
        duplicateCount++;
      } else {
        errorCount++;
      }

      results.push(result);
    } catch (error) {
      errorCount++;
      results.push({
        success: false,
        error: true,
        title: selectedBook.title || '제목 없음',
        isbn: selectedBook.isbn || '',
        message: String(error.message || error)
      });
    }
  }

  return {
    success: successCount > 0,
    successCount: successCount,
    duplicateCount: duplicateCount,
    ownedCount: ownedCount,
    errorCount: errorCount,
    results: results
  };
}

/**
 * YES24에서 책 목록을 검색합니다.
 * 제목 검색은 최대 50권, ISBN 검색은 해당 도서 한 권을 반환합니다.
 */
function HR_searchRequestBooks(query) {
  var keyword = String(query || '').trim();
  if (!keyword) throw new Error('검색할 책 제목이나 ISBN을 입력해 주세요.');
  var isIsbn = /^[0-9Xx\-\s]+$/.test(keyword) && /^(?:[0-9Xx]{10}|[0-9]{13})$/.test(keyword.replace(/[\s-]/g, ''));
  var key = HR_getYes24Key_();
  var items = yes24SearchItems_(keyword, key, 'BOOK', 50, isIsbn);
  var setPattern = /(?:세트|전집|박스(?:세트)?|묶음|전\s*\d+\s*(?:권|종)|총\s*\d+\s*(?:권|종)|\d+\s*(?:~|～|∼|-)\s*\d+\s*권|\d+\s*권\s*(?:세트|완결)|전권|합본)/i;
  function singles(list) {
    return list.filter(function(item) {
      var title = HR_cleanYes24Text_(item && item.title || '');
      var publisher = HR_cleanYes24Text_(item && item.publisher || '');
      var category = HR_cleanYes24Text_(item && item.goodsSortNm || '');
      return !setPattern.test(title) && !/(?:전집|세트)/i.test(publisher + ' ' + category);
    });
  }
  items = singles(items);
  // 권차를 적었다면 해당 단권만 보여 줍니다. 붙여 쓴 시리즈명도 비교합니다.
  var volumeMatch = !isIsbn && keyword.match(/^(.+?)\s+(\d{1,3})\s*(?:권|편|호)?$/);
  if (volumeMatch) {
    var base = volumeMatch[1].trim();
    var baseKey = base.replace(/\s+/g, '').toLowerCase();
    var volume = volumeMatch[2];
    function sameVolume(item) {
      var title = HR_cleanYes24Text_(item && item.title || '').replace(/\s+/g, '').toLowerCase();
      var at = title.indexOf(baseKey);
      return at >= 0 && new RegExp('^' + volume + '(?!\\d)').test(title.slice(at + baseKey.length));
    }
    var matched = items.filter(sameVolume);
    if (!matched.length) {
      matched = singles(yes24SearchItems_(base, key, 'BOOK', 100, false)).filter(sameVolume);
    }
    items = matched;
  } else if (!isIsbn && items.length <= 3) {
    // 띄어 쓰지 않은 시리즈명으로 검색된 단권의 실제 제목을 이용해 다른 권차도 찾습니다.
    var compact = keyword.replace(/\s+/g, '').toLowerCase();
    var sample = items.map(function(item) {
      return HR_cleanYes24Text_(item && item.title || '').match(/^(.+?)\s+\d{1,3}(?:\s|[:：.]|$)/);
    }).filter(Boolean)[0];
    if (sample && sample[1].replace(/\s+/g, '').toLowerCase() === compact && sample[1] !== keyword) {
      var seriesKey = compact;
      items = singles(yes24SearchItems_(sample[1], key, 'BOOK', 100, false)).filter(function(item) {
        var title = HR_cleanYes24Text_(item && item.title || '').replace(/\s+/g, '').toLowerCase();
        return title.indexOf(seriesKey) === 0 && new RegExp('^\\d{1,3}(?!\\d)').test(title.slice(seriesKey.length));
      });
    }
  }
  return items.map(yes24RequestBook_);
}

/**
 * YES24 텍스트를 화면 표시용으로 정리합니다.
 */
function HR_cleanYes24Text_(value) {
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

/**
 * 검색 결과 표지를 큰 이미지 주소로 정리합니다.
 */
function HR_largeYes24Cover_(value) {
  return HR_safeUrl_(value)
    
    ;
}

/**
 * 외부 링크를 안전한 HTTP 주소만 허용합니다.
 */
function HR_safeUrl_(value) {
  var text = String(value || '').trim();

  return /^https?:\/\//i.test(text)
    ? text
    : '';
}

/**
 * 우리학교 소장도서에서 ISBN을 우선 확인하고,
 * ISBN이 없을 때만 제목으로 중복을 확인합니다.
 */
function HR_findOwnedBook_(spreadsheet, book) {
  var librarySheet = spreadsheet.getSheetByName(
    HR_LIBRARY_SHEET_NAME
  );
  var result = {
    found: false,
    title: '',
    isbn: ''
  };
  var lastRow = 0;
  var lastColumn = 0;
  var headers = [];
  var titleColumn = 0;
  var isbnColumn = 0;
  var values = [];
  var targetTitle = HR_normalizeBookTitle_(
    book && book.title
  );
  var targetIsbn = HR_normalizeIsbn_(
    book && book.isbn
  );
  var rowIndex = 0;

  if (!librarySheet) {
    return result;
  }

  lastRow = librarySheet.getLastRow();
  lastColumn = librarySheet.getLastColumn();

  if (lastRow < 2 || lastColumn < 1) {
    return result;
  }

  headers = librarySheet
    .getRange(1, 1, 1, lastColumn)
    .getDisplayValues()[0];

  titleColumn = HR_findAnyHeaderColumn_(
    headers,
    ['책제목', '도서명', '서명', '제목']
  );
  isbnColumn = HR_findAnyHeaderColumn_(
    headers,
    ['ISBN13', 'ISBN']
  );

  if (!titleColumn && !isbnColumn) {
    return result;
  }

  values = librarySheet
    .getRange(2, 1, lastRow - 1, lastColumn)
    .getDisplayValues();

  for (rowIndex = 0; rowIndex < values.length; rowIndex++) {
    var libraryTitle = titleColumn
      ? String(values[rowIndex][titleColumn - 1] || '').trim()
      : '';
    var libraryIsbn = isbnColumn
      ? HR_normalizeIsbn_(
          values[rowIndex][isbnColumn - 1]
        )
      : '';
    var isbnMatches = Boolean(
      targetIsbn &&
      libraryIsbn &&
      targetIsbn === libraryIsbn
    );
    var titleMatches = Boolean(
      !targetIsbn &&
      targetTitle &&
      libraryTitle &&
      targetTitle === HR_normalizeBookTitle_(libraryTitle)
    );

    if (isbnMatches || titleMatches) {
      return {
        found: true,
        title: libraryTitle,
        isbn: libraryIsbn
      };
    }
  }

  return result;
}

/**
 * 희망도서 목록에 같은 책이 이미 신청되어 있는지 확인합니다.
 * ISBN이 있으면 ISBN을 우선 사용하고, ISBN이 없을 때만 제목을 비교합니다.
 */
function HR_findRequestedBook_(sheet, book, headers, columns) {
  var result = {
    found: false,
    title: '',
    isbn: '',
    row: 0
  };
  var lastRow = sheet.getLastRow();
  var titleColumn = columns && columns.title
    ? columns.title
    : HR_findAnyHeaderColumn_(
        headers || [],
        ['도서명', '책제목', '서명', '제목']
      );
  var isbnColumn = columns && columns.isbn
    ? columns.isbn
    : HR_findAnyHeaderColumn_(
        headers || [],
        ['ISBN13', 'ISBN']
      );
  var targetTitle = HR_normalizeBookTitle_(
    book && book.title
  );
  var targetIsbn = HR_normalizeIsbn_(
    book && book.isbn
  );
  var values = [];
  var rowIndex = 0;

  if (
    lastRow < FIRST_DATA_ROW ||
    (!titleColumn && !isbnColumn)
  ) {
    return result;
  }

  values = sheet
    .getRange(
      FIRST_DATA_ROW,
      1,
      lastRow - FIRST_DATA_ROW + 1,
      sheet.getLastColumn()
    )
    .getDisplayValues();

  for (rowIndex = 0; rowIndex < values.length; rowIndex++) {
    var requestedTitle = titleColumn
      ? String(values[rowIndex][titleColumn - 1] || '').trim()
      : '';
    var requestedIsbn = isbnColumn
      ? HR_normalizeIsbn_(
          values[rowIndex][isbnColumn - 1]
        )
      : '';
    var isbnMatches = Boolean(
      targetIsbn &&
      requestedIsbn &&
      targetIsbn === requestedIsbn
    );
    var titleMatches = Boolean(
      !targetIsbn &&
      targetTitle &&
      requestedTitle &&
      targetTitle === HR_normalizeBookTitle_(requestedTitle)
    );

    if (isbnMatches || titleMatches) {
      return {
        found: true,
        title: requestedTitle,
        isbn: requestedIsbn,
        row: rowIndex + FIRST_DATA_ROW
      };
    }
  }

  return result;
}

function HR_findAnyHeaderColumn_(headers, candidates) {
  var headerIndex = 0;
  var candidateIndex = 0;

  for (
    headerIndex = 0;
    headerIndex < headers.length;
    headerIndex++
  ) {
    var normalizedHeader = String(
      headers[headerIndex] || ''
    )
      .replace(/\s+/g, '')
      .toUpperCase();

    for (
      candidateIndex = 0;
      candidateIndex < candidates.length;
      candidateIndex++
    ) {
      var normalizedCandidate = String(
        candidates[candidateIndex] || ''
      )
        .replace(/\s+/g, '')
        .toUpperCase();

      if (
        normalizedHeader === normalizedCandidate ||
        normalizedHeader.indexOf(normalizedCandidate) !== -1
      ) {
        return headerIndex + 1;
      }
    }
  }

  return 0;
}

function HR_normalizeIsbn_(value) {
  return String(value || '')
    .replace(/[^0-9Xx]/g, '')
    .toUpperCase();
}

function HR_normalizeBookTitle_(value) {
  return String(value || '')
    .replace(/\([^)]*\)/g, '')
    .replace(/\[[^\]]*\]/g, '')
    .replace(/[^0-9A-Za-z가-힣]/g, '')
    .toLowerCase();
}

/**
 * 헤더 이름으로 열 번호를 찾습니다.
 */
function HR_findHeaderColumn_(
  headers,
  targetHeader,
  fallbackColumn
) {
  var headerIndex = 0;
  var normalizedTarget = String(targetHeader || '')
    .replace(/\s+/g, '');

  for (
    headerIndex = 0;
    headerIndex < headers.length;
    headerIndex++
  ) {
    var normalizedHeader = String(
      headers[headerIndex] || ''
    ).replace(/\s+/g, '');

    if (
      normalizedHeader.indexOf(normalizedTarget) !== -1
    ) {
      return headerIndex + 1;
    }
  }

  return fallbackColumn;
}

/**
 * 사용설명서의 핵심 안내를 한 화면에 보이도록 정리합니다.
 */
function HR_formatGuideSheet() {
  var spreadsheet = HR_getSetupSpreadsheet_();
  var sheet = spreadsheet.getSheetByName('📗 희망도서 신청안내');
  var webAppUrl = HR_getConfiguredPublicWebAppUrl_();
  var shortcutUrl = webAppUrl || (
    'https://docs.google.com/spreadsheets/d/' +
    spreadsheet.getId() +
    '/edit'
  );
  var titleColor = '#243B64';
  var accentColor = '#5277C7';
  var textColor = '#27344A';
  var lineColor = '#D8E1F0';
  var steps;

  if (!sheet) {
    throw new Error('사용설명서 시트를 찾지 못했습니다.');
  }

  sheet.setHiddenGridlines(true);
  sheet.setColumnWidth(1, 70);
  sheet.setColumnWidths(2, 2, 52);
  sheet.setColumnWidths(4, 9, 103);
  sheet.setFrozenRows(0);
  sheet.setFrozenColumns(0);
  HR_breakMergedRangesSafely_(sheet);
  sheet.getRange('B3:L50').clear();
  sheet
    .getRange('B3:L50')
    .setBackground('#FFFFFF')
    .setFontFamily('Jua')
    .setFontColor(textColor)
    .setVerticalAlignment('middle')
    .setWrap(true);

  sheet.getRange('B3:L4').merge();
  sheet
    .getRange('B3')
    .setValue('📚 희망도서 신청 안내');
  sheet.getRange('B5:L5').merge();
  sheet
    .getRange('B5')
    .setValue('신청 페이지 열기  →  도서 검색  →  여러 권 선택  →  한 번에 신청');

  steps = [
    {
      badge: 'B7:C8',
      text: 'D7:L8',
      number: '1',
      value: '아래 「희망도서 신청하기」 버튼을 누릅니다.'
    },
    {
      badge: 'B10:C11',
      text: 'D10:L11',
      number: '2',
      value: '책 제목 또는 ISBN을 입력해 원하는 책을 검색합니다.'
    },
    {
      badge: 'B13:C14',
      text: 'D13:L14',
      number: '3',
      value:
        '표지·저자·연도·정가를 확인하고 신청할 책을 여러 권 선택한 뒤 「선택한 도서 신청하기」를 누릅니다.'
    },
    {
      badge: 'B16:C17',
      text: 'D16:L17',
      number: '4',
      value:
        '신청한 도서는 「희망도서 목록」 시트에서 확인할 수 있습니다.'
    }
  ];

  HR_drawGuideSteps_(
    sheet,
    steps,
    accentColor,
    '#F3F6FC',
    lineColor
  );

  sheet.getRange('B19:L20').merge();
  sheet
    .getRange('B19')
    .setValue(
      'ℹ️ 신청자 이름은 선택사항이며, 입력하면 우선 대출 혜택을 받을 수 있습니다. 한 번에 최대 20권까지 신청할 수 있습니다.'
    )
    .setBackground('#EDF2FA')
    .setFontColor(titleColor)
    .setFontSize(14)
    .setHorizontalAlignment('left');

  sheet.getRange('B22:L24').merge();
  sheet
    .getRange('B22')
    .setFormula(
      '=HYPERLINK("' +
      shortcutUrl +
      '","📚 희망도서 신청하기")'
    )
    .setBackground(accentColor)
    .setFontColor('#FFFFFF')
    .setFontSize(19)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  sheet
    .getRange('B3:L4')
    .setBackground(titleColor)
    .setFontColor('#FFFFFF')
    .setFontSize(27)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  sheet
    .getRange('B5:L5')
    .setBackground('#E7EDF8')
    .setFontColor(titleColor)
    .setFontSize(14)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  [3, 4].forEach(function(rowNumber) {
    sheet.setRowHeight(rowNumber, 38);
  });
  sheet.setRowHeight(5, 34);
  [7, 8, 10, 11, 13, 14, 16, 17].forEach(function(rowNumber) {
    sheet.setRowHeight(rowNumber, 34);
  });
  [6, 9, 12, 15, 18, 21, 25, 26]
    .forEach(function(rowNumber) {
      sheet.setRowHeight(rowNumber, 10);
    });
  [19, 20].forEach(function(rowNumber) {
    sheet.setRowHeight(rowNumber, 31);
  });
  [22, 23, 24].forEach(function(rowNumber) {
    sheet.setRowHeight(rowNumber, 34);
  });

  sheet.activate();
  sheet.getRange('A3').activate();
}

function HR_formatGuideSheetAdminFirst_() {
  var spreadsheet = HR_getSetupSpreadsheet_();
  var sheet = spreadsheet.getSheetByName('📗 희망도서 신청안내');
  var webAppUrl = HR_getConfiguredPublicWebAppUrl_();
  var shortcutUrl = webAppUrl || (
    'https://docs.google.com/spreadsheets/d/' +
    spreadsheet.getId() +
    '/edit'
  );
  var titleColor = '#243B64';
  var adminColor = '#7A5128';
  var adminAccent = '#DDA333';
  var userAccent = '#5277C7';
  var textColor = '#27344A';
  var lineColor = '#D8E1F0';
  var adminSteps;
  var userSteps;

  if (!sheet) {
    throw new Error('사용설명서 시트를 찾지 못했습니다.');
  }

  sheet.setHiddenGridlines(true);
  sheet.setColumnWidth(1, 70);
  sheet.setColumnWidths(2, 2, 52);
  sheet.setColumnWidths(4, 9, 103);
  sheet.setFrozenRows(0);
  sheet.setFrozenColumns(0);
  HR_breakMergedRangesSafely_(sheet);
  
  
  sheet.getRange('B3:L50').clear();
  sheet
    .getRange('B3:L50')
    .setBackground('#FFFFFF')
    .setFontFamily('Jua')
    .setFontColor(textColor)
    .setVerticalAlignment('middle')
    .setWrap(true);

  sheet.getRange('B3:L4').merge();
  sheet
    .getRange('B3')
    .setValue('🏫 다른 학교에서 사본으로 사용하는 방법');
  sheet.getRange('B5:L5').merge();
  sheet
    .getRange('B5')
    .setValue('처음 한 번만 1번부터 차례대로 따라 하세요.');

  adminSteps = [
    {
      badge: 'B7:C8',
      text: 'D7:L8',
      number: '1',
      value:
        '원본 시트에서 「파일 → 사본 만들기」를 누르고 새 학교 이름으로 저장합니다.'
    },
    {
      badge: 'B10:C11',
      text: 'D10:L11',
      number: '2',
      value:
        '사본의 「⚙️ 사본 최초설정」 탭에서 API 키 설정을 누릅니다. YES24은 필수, 국립중앙도서관은 선택입니다.'
    },
    {
      badge: 'B13:C14',
      text: 'D13:L14',
      number: '3',
      value:
        '「이 학교 최초 설정」을 누르고 학교 이름을 입력한 뒤 실행합니다.'
    },
    {
      badge: 'B16:C17',
      text: 'D16:L17',
      number: '4',
      value:
        '「우리학교 소장도서」 시트에 책 제목과 ISBN13을 붙여넣습니다.'
    },
    {
      badge: 'B19:C20',
      text: 'D19:L20',
      number: '5',
      value:
        '확장 프로그램 → Apps Script → 배포 → 새 배포 → 웹 앱에서 「나 / 모든 사용자」로 배포합니다. 완료 후 설정 탭은 삭제해도 됩니다.'
    }
  ];

  HR_drawGuideSteps_(
    sheet,
    adminSteps,
    adminAccent,
    '#FFF9EF',
    '#E8D2AB'
  );

  sheet.getRange('B22:L24').merge();
  sheet
    .getRange('B22')
    .setValue(
      '꼭 확인하세요! 「우리학교 소장도서」에 책 제목과 ISBN13을 모두 넣어야 “이미 소장중인 도서입니다.” 안내가 정확히 나옵니다.'
    )
    .setBackground('#FFF0F0')
    .setFontColor('#A33C3C')
    .setFontSize(14)
    .setFontWeight('bold')
    .setHorizontalAlignment('left');

  sheet.getRange('B26:L27').merge();
  sheet
    .getRange('B26')
    .setValue(
      'API 안내  ·  YES24은 도서 검색에 꼭 필요합니다. 국립중앙도서관 API는 KDC 참고용이라 입력하지 않아도 됩니다.'
    )
    .setBackground('#EDF4FF')
    .setFontColor('#284E86')
    .setFontSize(14)
    .setFontWeight('bold')
    .setHorizontalAlignment('left');

  sheet.getRange('B30:L31').merge();
  sheet
    .getRange('B30')
    .setValue('📚 희망도서 신청 방법');
  sheet.getRange('B32:L32').merge();
  sheet
    .getRange('B32')
    .setValue('신청 페이지 열기  →  도서 검색  →  원하는 책 선택');

  userSteps = [
    {
      badge: 'B34:C35',
      text: 'D34:L35',
      number: '1',
      value: '아래 「희망도서 신청하기」 버튼을 누릅니다.'
    },
    {
      badge: 'B37:C38',
      text: 'D37:L38',
      number: '2',
      value: '책 제목 또는 ISBN을 입력해 원하는 책을 검색합니다.'
    },
    {
      badge: 'B40:C41',
      text: 'D40:L41',
      number: '3',
      value:
        '표지·저자·연도·정가를 확인한 뒤 「이 도서 선택하기」를 누릅니다.'
    }
  ];

  HR_drawGuideSteps_(
    sheet,
    userSteps,
    userAccent,
    '#F3F6FC',
    lineColor
  );

  sheet.getRange('B43:L44').merge();
  sheet
    .getRange('B43')
    .setValue(
      'ℹ️ 신청자 이름은 선택사항이며, 입력하면 우선 대출 혜택을 받을 수 있습니다. 시리즈는 여러 권을 차례로 선택할 수 있습니다.'
    )
    .setBackground('#EDF2FA')
    .setFontColor(titleColor)
    .setFontSize(14)
    .setHorizontalAlignment('left');

  sheet.getRange('B46:L48').merge();
  sheet
    .getRange('B46')
    .setFormula(
      '=HYPERLINK("' +
      shortcutUrl +
      '","📚 희망도서 신청하기")'
    )
    .setBackground(userAccent)
    .setFontColor('#FFFFFF')
    .setFontSize(19)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  sheet
    .getRange('B3:L4')
    .setBackground(adminColor)
    .setFontColor('#FFFFFF')
    .setFontSize(27)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  sheet
    .getRange('B5:L5')
    .setBackground('#F3E9DA')
    .setFontColor(adminColor)
    .setFontSize(14)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  sheet
    .getRange('B30:L31')
    .setBackground(titleColor)
    .setFontColor('#FFFFFF')
    .setFontSize(25)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  sheet
    .getRange('B32:L32')
    .setBackground('#E7EDF8')
    .setFontColor(titleColor)
    .setFontSize(14)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  [3, 4, 30, 31].forEach(function(rowNumber) {
    sheet.setRowHeight(rowNumber, 38);
  });
  [5, 32].forEach(function(rowNumber) {
    sheet.setRowHeight(rowNumber, 34);
  });
  [
    7, 8, 10, 11, 13, 14, 16, 17, 19, 20,
    34, 35, 37, 38, 40, 41
  ].forEach(function(rowNumber) {
    sheet.setRowHeight(rowNumber, 34);
  });
  [6, 9, 12, 15, 18, 21, 25, 28, 29, 33, 36, 39, 42, 45, 49, 50]
    .forEach(function(rowNumber) {
      sheet.setRowHeight(rowNumber, 10);
    });
  [22, 23, 24].forEach(function(rowNumber) {
    sheet.setRowHeight(rowNumber, 28);
  });
  [26, 27, 43, 44].forEach(function(rowNumber) {
    sheet.setRowHeight(rowNumber, 31);
  });
  [46, 47, 48].forEach(function(rowNumber) {
    sheet.setRowHeight(rowNumber, 34);
  });

  sheet.activate();
  sheet.getRange('A3').activate();
}

function HR_drawGuideSteps_(
  sheet,
  steps,
  badgeColor,
  textBackground,
  borderColor
) {
  steps.forEach(function(step) {
    sheet.getRange(step.badge).merge();
    sheet.getRange(step.text).merge();
    sheet
      .getRange(step.badge)
      .setBackground(badgeColor)
      .setFontColor('#FFFFFF')
      .setFontSize(24)
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setBorder(
        true,
        true,
        true,
        true,
        false,
        false,
        badgeColor,
        SpreadsheetApp.BorderStyle.SOLID
      );
    sheet
      .getRange(step.badge.split(':')[0])
      .setValue(step.number);
    sheet
      .getRange(step.text)
      .setBackground(textBackground)
      .setFontColor('#27344A')
      .setFontSize(15)
      .setFontWeight('bold')
      .setHorizontalAlignment('left')
      .setBorder(
        true,
        true,
        true,
        true,
        false,
        false,
        borderColor,
        SpreadsheetApp.BorderStyle.SOLID
      );
    sheet
      .getRange(step.text.split(':')[0])
      .setValue(step.value);
  });
}

function HR_formatGuideSheetLegacy_() {
  var spreadsheet = HR_getSpreadsheet_();
  var sheet = spreadsheet.getSheetByName('📗 희망도서 신청안내');
  var webAppUrl = HR_getConfiguredPublicWebAppUrl_();
  var shortcutUrl = webAppUrl || (
    'https://docs.google.com/spreadsheets/d/' +
    spreadsheet.getId() +
    '/edit?gid=0#gid=0'
  );

  if (!sheet) {
    throw new Error('사용설명서 시트를 찾지 못했습니다.');
  }

  var guideRange = sheet.getRange('B3:L46');
  var titleColor = '#243B64';
  var accentColor = '#5277C7';
  var paleColor = '#F3F6FC';
  var textColor = '#27344A';
  var lineColor = '#D8E1F0';

  sheet.setHiddenGridlines(true);
  sheet.setColumnWidth(1, 70);
  sheet.setColumnWidths(2, 2, 52);
  sheet.setColumnWidths(4, 9, 103);

  guideRange.breakApart();
  guideRange.clear();
  guideRange
    .setBackground('#FFFFFF')
    .setFontFamily('Jua')
    .setFontColor(textColor)
    .setVerticalAlignment('middle')
    .setWrap(true);

  sheet.getRange('B3:L4').merge();
  sheet
    .getRange('B3')
    .setValue('📚 희망도서 신청 안내');

  sheet.getRange('B5:L5').merge();
  sheet
    .getRange('B5')
    .setValue('신청 페이지 열기  →  도서 검색  →  원하는 책 선택');

  var steps = [
    {
      badgeRange: 'B7:C8',
      textRange: 'D7:L8',
      number: '1',
      text: '아래 「희망도서 신청하기」 버튼을 누릅니다.'
    },
    {
      badgeRange: 'B10:C11',
      textRange: 'D10:L11',
      number: '2',
      text: '책 제목 또는 ISBN을 입력해 원하는 책을 검색합니다.'
    },
    {
      badgeRange: 'B13:C14',
      textRange: 'D13:L14',
      number: '3',
      text: '표지·저자·연도·정가를 확인한 뒤 「이 도서 선택하기」를 누릅니다.'
    }
  ];

  steps.forEach(function(step) {
    sheet.getRange(step.badgeRange).merge();
    sheet.getRange(step.textRange).merge();

    sheet
      .getRange(step.badgeRange)
      .setBackground(accentColor)
      .setFontColor('#FFFFFF')
      .setFontSize(24)
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setBorder(
        true,
        true,
        true,
        true,
        false,
        false,
        accentColor,
        SpreadsheetApp.BorderStyle.SOLID
      );

    sheet
      .getRange(step.badgeRange.split(':')[0])
      .setValue(step.number);

    sheet
      .getRange(step.textRange)
      .setBackground(paleColor)
      .setFontColor(textColor)
      .setFontSize(15)
      .setFontWeight('bold')
      .setHorizontalAlignment('left')
      .setBorder(
        true,
        true,
        true,
        true,
        false,
        false,
        lineColor,
        SpreadsheetApp.BorderStyle.SOLID
      );

    sheet
      .getRange(step.textRange.split(':')[0])
      .setValue(step.text);
  });

  sheet.getRange('B16:L17').merge();
  sheet
    .getRange('B16')
    .setValue(
      'ℹ️ 신청자 이름은 선택사항이며, 입력하면 우선 대출 혜택을 받을 수 있습니다. 시리즈는 여러 권을 차례로 선택할 수 있습니다.'
    );

  sheet.getRange('B19:L20').merge();
  sheet
    .getRange('B19')
    .setFormula(
      '=HYPERLINK("' +
      shortcutUrl +
      '","📚 희망도서 신청하기")'
    );

  sheet.getRange('B22:L23').merge();
  sheet
    .getRange('B22')
    .setValue('🏫 다른 학교에서 사본으로 사용하는 방법');

  sheet.getRange('B25:C26').merge();
  sheet.getRange('D25:L26').merge();
  sheet.getRange('B28:C29').merge();
  sheet.getRange('D28:L29').merge();
  sheet.getRange('B31:C32').merge();
  sheet.getRange('D31:L32').merge();
  sheet.getRange('B34:C35').merge();
  sheet.getRange('D34:L35').merge();
  sheet.getRange('B37:C38').merge();
  sheet.getRange('D37:L38').merge();

  sheet.getRange('B25').setValue('1');
  sheet
    .getRange('D25')
    .setValue(
      '원본 시트에서 「파일 → 사본 만들기」를 누르고 새 학교 이름으로 저장합니다.'
    );
  sheet.getRange('B28').setValue('2');
  sheet
    .getRange('D28')
    .setValue(
      '사본의 「⚙️ 사본 최초설정」 탭에서 API 키 설정을 누릅니다. YES24은 필수, 국립중앙도서관은 선택입니다.'
    );
  sheet.getRange('B31').setValue('3');
  sheet
    .getRange('D31')
    .setValue(
      '「이 학교 최초 설정」을 누르고 학교 이름을 입력한 뒤 실행합니다.'
    );
  sheet.getRange('B34').setValue('4');
  sheet
    .getRange('D34')
    .setValue(
      '「우리학교 소장도서」 시트에 책 제목과 ISBN13을 붙여넣습니다.'
    );
  sheet.getRange('B37').setValue('5');
  sheet
    .getRange('D37')
    .setValue(
      '확장 프로그램 → Apps Script → 배포 → 새 배포 → 웹 앱에서 「나 / 모든 사용자」로 배포합니다. 완료 후 설정 탭은 삭제해도 됩니다.'
    );

  [
    'B25:C26',
    'B28:C29',
    'B31:C32',
    'B34:C35',
    'B37:C38'
  ].forEach(
    function(rangeA1) {
      sheet
        .getRange(rangeA1)
        .setBackground('#D99A32')
        .setFontColor('#FFFFFF')
        .setFontSize(22)
        .setFontWeight('bold')
        .setHorizontalAlignment('center');
    }
  );

  [
    'D25:L26',
    'D28:L29',
    'D31:L32',
    'D34:L35',
    'D37:L38'
  ].forEach(
    function(rangeA1) {
      sheet
        .getRange(rangeA1)
        .setBackground('#FFF8EC')
        .setFontColor(textColor)
        .setFontSize(14)
        .setFontWeight('bold')
        .setHorizontalAlignment('left')
        .setBorder(
          true,
          true,
          true,
          true,
          false,
          false,
          '#E9D4AE',
          SpreadsheetApp.BorderStyle.SOLID
        );
    }
  );

  sheet.getRange('B41:L43').merge();
  sheet
    .getRange('B41')
    .setValue(
      '꼭 확인하세요! 「우리학교 소장도서」에 책 제목과 ISBN13을 모두 넣어야 “이미 소장중인 도서입니다.” 안내가 정확히 나옵니다.'
    );

  sheet.getRange('B45:L46').merge();
  sheet
    .getRange('B45')
    .setValue(
      'API 안내  ·  YES24 API Key는 도서 검색에 꼭 필요합니다. 국립중앙도서관 API는 KDC 분류번호 참고용이라 입력하지 않아도 신청 기능은 정상 작동합니다.'
    );

  sheet
    .getRange('B3:L4')
    .setBackground(titleColor)
    .setFontColor('#FFFFFF')
    .setFontSize(27)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  sheet
    .getRange('B5:L5')
    .setBackground('#E7EDF8')
    .setFontColor(titleColor)
    .setFontSize(14)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  sheet
    .getRange('B16:L17')
    .setBackground('#EDF2FA')
    .setFontColor(titleColor)
    .setFontSize(14)
    .setFontWeight('normal')
    .setHorizontalAlignment('left')
    .setBorder(
      true,
      true,
      true,
      true,
      false,
      false,
      lineColor,
      SpreadsheetApp.BorderStyle.SOLID
    );

  sheet
    .getRange('B19:L20')
    .setBackground(accentColor)
    .setFontColor('#FFFFFF')
    .setFontSize(17)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  sheet
    .getRange('B22:L23')
    .setBackground('#6F4A25')
    .setFontColor('#FFFFFF')
    .setFontSize(18)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  sheet
    .getRange('B41:L43')
    .setBackground('#FFF0F0')
    .setFontColor('#A33C3C')
    .setFontSize(13)
    .setFontWeight('bold')
    .setHorizontalAlignment('left');

  sheet
    .getRange('B45:L46')
    .setBackground('#EDF4FF')
    .setFontColor('#284E86')
    .setFontSize(13)
    .setFontWeight('bold')
    .setHorizontalAlignment('left');

  [3, 4].forEach(function(rowNumber) {
    sheet.setRowHeight(rowNumber, 38);
  });
  sheet.setRowHeight(5, 34);
  [7, 8, 10, 11, 13, 14].forEach(
    function(rowNumber) {
      sheet.setRowHeight(rowNumber, 32);
    }
  );
  [6, 9, 12, 15, 18].forEach(
    function(rowNumber) {
      sheet.setRowHeight(rowNumber, 12);
    }
  );
  [16, 17, 19, 20].forEach(
    function(rowNumber) {
      sheet.setRowHeight(rowNumber, 31);
    }
  );
  [22, 23].forEach(function(rowNumber) {
    sheet.setRowHeight(rowNumber, 30);
  });
  [
    25, 26,
    28, 29,
    31, 32,
    34, 35,
    37, 38
  ].forEach(
    function(rowNumber) {
      sheet.setRowHeight(rowNumber, 31);
    }
  );
  [21, 24, 27, 30, 33, 36, 39, 40, 44].forEach(
    function(rowNumber) {
      sheet.setRowHeight(rowNumber, 10);
    }
  );
  [41, 42, 43].forEach(function(rowNumber) {
    sheet.setRowHeight(rowNumber, 28);
  });
  [45, 46].forEach(function(rowNumber) {
    sheet.setRowHeight(rowNumber, 31);
  });

  sheet.activate();
  sheet.getRange('A3').activate();
}

/**
 * ============================================================
 * 사본 배포용 학교 최초 설정
 * ============================================================
 */

function HR_addCopySetupMenu() {
  SpreadsheetApp.getUi()
    .createMenu('⚙️ 사본 설정')
    .addItem('✅ 설정 한 번에 점검', 'HR_showSystemCheck')
    .addItem('🔗 신청 링크 받기', 'HR_showApplyLinksDialog')
    .addItem('📢 업데이트 소식 보기', 'HR_showUpdateNotice')
    .addItem('📏 신청 한도 설정', 'HR_showRequestLimitDialog')
    .addItem('🔑 교직원 인증번호 설정', 'HR_showStaffCodeDialog')
    .addSeparator()
    .addItem(
      '🔑 API 키 설정 (YES24 필수)',
      'HR_showYes24KeyDialog'
    )
    .addItem('🏫 학교 최초 설정', 'HR_showSchoolSetupDialog')
    .addItem(
      '🔗 공개 신청 링크 등록',
      'HR_showPublicWebAppUrlDialog'
    )
    .addSeparator()
    .addItem('📥 이전 사본에서 데이터 가져오기', 'HR_importFromOldCopy')
    .addItem('🧰 설정 탭 만들기·새로고침', 'HR_createCopySetupSheet')
    .addItem('🗑️ 설정 완료 후 탭 삭제', 'HR_deleteCopySetupSheet')
    .addToUi();
}

function HR_showYes24KeyDialog() {
  try {
    var template = HtmlService.createTemplateFromFile('CopySetup');
    template.mode = 'api';
    var output = template.evaluate().setWidth(680).setHeight(690);
    SpreadsheetApp.getUi().showModalDialog(
      output,
      '🔑 도서 API 키 설정'
    );
  } catch (error) {
    var ui = SpreadsheetApp.getUi();
    var yes24Response = ui.prompt(
      '🔑 YES24 API 키 설정',
      'YES24 API Key를 붙여 넣으세요.',
      ui.ButtonSet.OK_CANCEL
    );
    if (yes24Response.getSelectedButton() !== ui.Button.OK) return;
    var yes24Key = String(yes24Response.getResponseText() || '').trim();
    if (!yes24Key) {
      ui.alert('YES24 API 키를 입력해 주세요.');
      return;
    }
    var nlkResponse = ui.prompt(
      '국립중앙도서관 API 키 (선택)',
      '없으면 빈칸으로 두고 확인을 누르세요.',
      ui.ButtonSet.OK_CANCEL
    );
    var nlkKey = nlkResponse.getSelectedButton() === ui.Button.OK
      ? String(nlkResponse.getResponseText() || '').trim()
      : '';
    try {
      var result = HR_saveYes24ApiKey({
        yes24Key: yes24Key,
        nlkKey: nlkKey
      });
      ui.alert(result.message || 'API 키를 저장했습니다.');
    } catch (saveError) {
      ui.alert(
        'API 키 저장 오류',
        saveError.message || String(saveError),
        ui.ButtonSet.OK
      );
    }
  }
}

function HR_showSchoolSetupDialog() {
  HR_showCopySetupDialog_(
    'school',
    '🏫 학교 최초 설정',
    700,
    680
  );
}

/**
 * 웹 앱 배포 주소를 한 번만 붙여 넣는 간단한 설정창입니다.
 * 저장하면 사용설명서의 「희망도서 신청하기」 버튼도 자동 갱신됩니다.
 */
function HR_showPublicWebAppUrlDialog() {
    var activeSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  if (activeSheet) {
    activeSheet.getRange('A3').activate();
  }

  var savedUrl = HR_getConfiguredPublicWebAppUrl_();
  var html = HtmlService
    .createHtmlOutput(
      '<!DOCTYPE html>' +
      '<html lang="ko"><head><base target="_top">' +
      '<meta charset="UTF-8">' +
      '<style>' +
      '*{box-sizing:border-box}' +
      'body{margin:0;padding:28px;background:#f5f8ff;color:#253858;' +
      'font-family:-apple-system,BlinkMacSystemFont,"Apple SD Gothic Neo",' +
      '"Noto Sans KR",sans-serif}' +
      'h1{margin:0 0 8px;font-size:24px}' +
      '.sub{margin:0 0 22px;color:#5f6f8b;font-size:15px;line-height:1.65}' +
      '.card{padding:22px;border:1px solid #dbe4f3;border-radius:18px;' +
      'background:#fff;box-shadow:0 10px 28px rgba(54,83,130,.08)}' +
      'label{display:block;margin-bottom:9px;font-size:17px;font-weight:800}' +
      'input{width:100%;height:56px;padding:0 16px;border:2px solid #cfdaf0;' +
      'border-radius:13px;outline:none;font-size:16px}' +
      'input:focus{border-color:#5277c7;box-shadow:0 0 0 4px #e8eefb}' +
      '.help{margin:12px 0 0;color:#66758f;font-size:14px;line-height:1.6}' +
      '.message{display:none;margin-top:15px;padding:12px 14px;' +
      'border-radius:11px;font-size:14px;line-height:1.5}' +
      '.message.show{display:block}.message.error{background:#fff0f0;color:#a23f3f}' +
      '.message.ok{background:#edf8f0;color:#27633b}' +
      '.buttons{display:flex;justify-content:flex-end;gap:10px;margin-top:20px}' +
      'button{height:48px;padding:0 20px;border-radius:12px;font-size:16px;' +
      'font-weight:800;cursor:pointer}.cancel{border:1px solid #d6deeb;' +
      'background:#fff;color:#65738a}.save{border:0;background:#5277c7;color:#fff}' +
      'button:disabled{cursor:wait;opacity:.55}' +
      '</style></head><body>' +
      '<h1>🔗 공개 신청 링크 등록</h1>' +
      '<p class="sub">웹 앱을 배포한 뒤 주소를 아래 칸에 한 번만 붙여 넣으세요.<br>' +
      '사용설명서의 신청 버튼은 자동으로 바뀝니다.</p>' +
      '<div class="card">' +
      '<label for="publicUrl">웹 앱 주소</label>' +
      '<input id="publicUrl" type="url" spellcheck="false" ' +
      'placeholder="https://script.google.com/macros/s/.../exec" value="' +
      HR_escapeHtml_(savedUrl) +
      '">' +
      '<p class="help">Apps Script → 배포 → 배포 관리에서 ' +
      '<strong>/exec로 끝나는 웹 앱 URL</strong>을 복사해 붙여 넣으세요.</p>' +
      '<div id="message" class="message"></div>' +
      '<div class="buttons">' +
      '<button class="cancel" type="button" ' +
      'onclick="google.script.host.close()">닫기</button>' +
      '<button id="saveButton" class="save" type="button" ' +
      'onclick="saveUrl()">링크 저장하기</button>' +
      '</div></div>' +
      '<script>' +
      'function saveUrl(){' +
      'var input=document.getElementById("publicUrl");' +
      'var button=document.getElementById("saveButton");' +
      'var message=document.getElementById("message");' +
      'var url=input.value.trim();' +
      'if(!url){showMessage("웹 앱 주소를 붙여 넣어 주세요.","error");return;}' +
      'button.disabled=true;button.textContent="저장 중…";' +
      'google.script.run.withSuccessHandler(function(result){' +
      'showMessage(result.message,"ok");button.textContent="저장 완료 ✓";' +
      'setTimeout(function(){google.script.host.close();},900);' +
      '}).withFailureHandler(function(error){' +
      'button.disabled=false;button.textContent="링크 저장하기";' +
      'showMessage(error&&error.message?error.message:String(error),"error");' +
      '}).HR_savePublicWebAppUrl({url:url});}' +
      'function showMessage(text,type){var box=document.getElementById("message");' +
      'box.textContent=text;box.className="message show "+type;}' +
      '</script></body></html>'
    )
    .setWidth(650)
    .setHeight(500);

  SpreadsheetApp.getUi().showModalDialog(
    html,
    '공개 신청 링크 등록'
  );
}

/**
 * 공개 /exec 주소를 저장하고 사용설명서 신청 버튼에 즉시 반영합니다.
 */
function HR_savePublicWebAppUrl(formData) {
  var url = String(
    formData && formData.url || ''
  ).trim();

  if (!url) {
    throw new Error('웹 앱 주소를 붙여 넣어 주세요.');
  }

  // 복사 과정에서 들어간 공백·쿼리 문자열을 제거합니다.
  url = url.replace(/\s+/g, '');
  url = url.replace(/[?#].*$/, '');

  // 배포 기본 주소만 붙여 넣어도 공개용 /exec 주소로 자동 보정합니다.
  if (
    /^https:\/\/script\.google\.com\/macros\/s\/[^/]+\/?$/.test(url)
  ) {
    url = url.replace(/\/$/, '') + '/exec';
  }

  url = url
    .replace(/\/dev\/?$/, '/exec')
    .replace(/\/exec\/?$/, '/exec');

  if (
    !/^https:\/\/script\.google\.com\/macros\/s\/[^/]+\/exec$/.test(
      url
    )
  ) {
    throw new Error(
      '주소가 올바르지 않습니다. Apps Script 배포 관리의 웹 앱 주소를 붙여 넣어 주세요.'
    );
  }

  // 실제로 열리는 주소인지 확인합니다. 안 열리는 주소는 저장하지 않고 이유를 알려 줍니다.
  var problem = HR_probeProblemMessage_(HR_probeWebAppUrl_(url));
  if (problem) {
    throw new Error(problem);
  }

  var properties = PropertiesService.getScriptProperties();

  // 주소 검사가 끝난 뒤에만 저장하고 안내 시트를 다시 그립니다.
  // 따라서 잘못된 주소를 입력해도 뒤쪽 시트 디자인은 바뀌지 않습니다.
  properties.setProperty(
    HR_PUBLIC_WEB_APP_URL_PROPERTY,
    url
  );

  var updatedButtons = HR_updateGuidePublicLink_(url);

  return {
    success: true,
    message:
      '링크를 저장했습니다. 희망도서 신청 버튼을 새 주소로 연결했습니다.',
    webAppUrl: url
  };
}

/**
 * 공개 링크 저장 시 안내 시트 전체를 다시 만들지 않고
 * 기존 신청 버튼의 링크만 안전하게 바꿉니다.
 * 병합 셀·고정 행 구조를 건드리지 않으므로 디자인이 깨지지 않습니다.
 */
function HR_updateGuidePublicLink_(url) {
  var spreadsheet = HR_getSetupSpreadsheet_();
  var sheet = spreadsheet.getSheetByName('📗 희망도서 신청안내');
  if (!sheet) {
    throw new Error('희망도서 신청안내 시트를 찾지 못했습니다.');
  }

  var linkText = '📚 희망도서 신청하기';
  var linkValue = SpreadsheetApp
    .newRichTextValue()
    .setText(linkText)
    .setLinkUrl(String(url || ''))
    .build();
  var targets = [];
  var seen = {};

  // 이미 만들어진 신청 버튼이 있으면 링크만 바꿉니다.
  var found = sheet
    .createTextFinder('희망도서 신청하기')
    .matchCase(false)
    .matchEntireCell(false)
    .findAll();

  found.forEach(function(cell) {
    var merged = cell.getMergedRanges();
    var buttonRange = merged.length ? merged[0] : cell;
    var target = buttonRange.getCell(1, 1);
    var key = target.getA1Notation();

    if (!seen[key]) {
      seen[key] = true;
      targets.push({
        cell: target,
        range: buttonRange
      });
    }
  });

  // 현재 통합 안내 양식에는 버튼이 없으므로
  // 본문 아래의 새 행에 버튼을 한 번만 만듭니다.
  if (!targets.length) {
    var buttonStartRow = Math.max(26, sheet.getLastRow() + 2);
    var buttonEndRow = buttonStartRow + 2;

    if (sheet.getMaxRows() < buttonEndRow) {
      sheet.insertRowsAfter(
        sheet.getMaxRows(),
        buttonEndRow - sheet.getMaxRows()
      );
    }

    var newButtonRange = sheet.getRange(
      buttonStartRow,
      1,
      3,
      12
    );

    newButtonRange
      .breakApart()
      .clearContent()
      .merge();

    targets.push({
      cell: newButtonRange.getCell(1, 1),
      range: newButtonRange
    });

    sheet.setRowHeight(buttonStartRow - 1, 10);
    sheet.setRowHeights(buttonStartRow, 3, 30);
  }

  targets.forEach(function(item) {
    item.cell.setRichTextValue(linkValue);

    item.range
      .setBackground('#2F765E')
      .setFontColor('#FFFFFF')
      .setFontFamily('Arial')
      .setFontSize(18)
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle')
      .setBorder(
        true,
        true,
        true,
        true,
        false,
        false,
        '#245B49',
        SpreadsheetApp.BorderStyle.SOLID_THICK
      );
  });

  SpreadsheetApp.flush();
  return targets.length;
}

function HR_showCopySetupDialog_(
  mode,
  title,
  width,
  height
) {
  var template = HtmlService.createTemplateFromFile(
    'CopySetup'
  );

  template.mode = mode;

  SpreadsheetApp.getUi().showModalDialog(
    template
      .evaluate()
      .setWidth(width)
      .setHeight(height),
    title
  );
}

/**
 * 설정 탭의 큰 버튼 영역을 선택하면 설정창을 엽니다.
 * 상단 「⚙️ 사본 설정」 메뉴에서도 같은 창을 열 수 있습니다.
 */
function onSelectionChange(e) {
  if (!e || !e.range) {
    return;
  }

  var sheet = e.range.getSheet();

  if (sheet.getName() !== HR_COPY_SETUP_SHEET_NAME) {
    return;
  }

  var activeCell = e.range.getA1Notation().split(':')[0];

  if (activeCell === 'B61') {
    HR_showYes24KeyDialog();
  } else if (activeCell === 'B65') {
    HR_showSchoolSetupDialog();
  } else if (activeCell === 'B69') {
    HR_showPublicWebAppUrlDialog();
  }
}

function HR_getCopySetupState() {
  var spreadsheet = HR_getSetupSpreadsheet_();
  var properties = PropertiesService.getScriptProperties();
  var libraryStatus = HR_getLibrarySetupStatus_(spreadsheet);
  var configuredSpreadsheetId = String(
    properties.getProperty(HR_SPREADSHEET_ID_PROPERTY) || ''
  ).trim();
  var yes24Key = HR_getYes24Key_();
  var nlkKey = HR_getNlkKey_();
  var webAppUrl = HR_getConfiguredPublicWebAppUrl_();

  return {
    spreadsheetId: spreadsheet.getId(),
    configuredSpreadsheetId: configuredSpreadsheetId,
    spreadsheetName: spreadsheet.getName(),
    schoolName: String(
      properties.getProperty(HR_SCHOOL_NAME_PROPERTY) || ''
    ).trim(),
    managerName: String(
      properties.getProperty(HR_MANAGER_NAME_PROPERTY) || ''
    ).trim(),
    managerExtension: String(
      properties.getProperty(HR_MANAGER_EXTENSION_PROPERTY) || ''
    ).trim(),
    maxClasses: String(properties.getProperty(HR_MAX_CLASSES_PROPERTY) || '').trim(),
    yes24Key: '',
    yes24Ready: Boolean(yes24Key),
    nlkKey: '',
    nlkReady: Boolean(nlkKey),
    targetSheetReady: Boolean(
      spreadsheet.getSheetByName(HR_TARGET_SHEET_NAME)
    ),
    librarySheetReady: libraryStatus.sheetReady,
    libraryHeadersReady: libraryStatus.headersReady,
    libraryTitleColumnReady: libraryStatus.titleColumn > 0,
    libraryIsbnColumnReady: libraryStatus.isbnColumn > 0,
    libraryBookCount: libraryStatus.bookCount,
    webAppUrl: webAppUrl,
    setupComplete:
      String(
        properties.getProperty(HR_SETUP_COMPLETE_PROPERTY) || ''
      ) === 'true' &&
      configuredSpreadsheetId === spreadsheet.getId() &&
      Boolean(yes24Key) &&
      libraryStatus.headersReady
  };
}

function HR_saveYes24ApiKey(formData) {
  var key = String(
    formData && formData.yes24Key || HR_getYes24Key_()
  ).trim();
  var nlkKey = String(
    formData && formData.nlkKey || HR_getNlkKey_()
  ).trim();
  var properties = PropertiesService.getScriptProperties();

  if (!key) {
    throw new Error('YES24 API Key를 입력해 주세요.');
  }

  HR_testYes24Key_(key);

  properties.setProperty(HR_YES24_KEY_PROPERTY, key);

  if (nlkKey) {
    properties.setProperty(HR_NLK_KEY_PROPERTY, nlkKey);
  } else {
    properties.deleteProperty(HR_NLK_KEY_PROPERTY);
  }

  return {
    success: true,
    message:
      'YES24 연결을 확인했습니다.' +
      (nlkKey
        ? ' 국립중앙도서관 키도 저장했습니다.'
        : ' KDC 참고용 키는 비워 두었습니다.')
  };
}

function HR_runSchoolInitialSetup(formData) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var properties = PropertiesService.getScriptProperties();
  var schoolName = String(
    formData && formData.schoolName || ''
  ).trim();
  var managerName = String(
    formData && formData.managerName || ''
  ).trim();
  var managerExtension = String(
    formData && formData.managerExtension || ''
  ).trim();
  var maxClassesText = String(
    formData && formData.maxClasses || ''
  ).trim();
  var targetSheet;
  var libraryStatus;
  var yes24Ready = Boolean(HR_getYes24Key_());
  var warnings = [];

  if (!spreadsheet) {
    throw new Error(
      '사본 스프레드시트에서 학교 최초 설정창을 열어 주세요.'
    );
  }

  if (!schoolName) {
    throw new Error('학교 이름을 입력해 주세요.');
  }

  if (maxClassesText && !(/^\d{1,2}$/.test(maxClassesText) && Number(maxClassesText) >= 1 && Number(maxClassesText) <= 30)) {
    throw new Error('한 학년 최대 반 수는 1~30 사이의 숫자로 입력해 주세요. (비워 두면 20반까지 보입니다.)');
  }

  targetSheet = spreadsheet.getSheetByName(
    HR_TARGET_SHEET_NAME
  );

  if (!targetSheet) {
    throw new Error(
      '「' + HR_TARGET_SHEET_NAME +
      '」 시트를 찾지 못했습니다.'
    );
  }

  properties.setProperties({
    HR_TARGET_SPREADSHEET_ID: spreadsheet.getId(),
    HR_SCHOOL_NAME: schoolName,
    HR_MANAGER_NAME: managerName,
    HR_MANAGER_EXTENSION: managerExtension,
    HR_MAX_CLASSES: maxClassesText,
    HR_SETUP_COMPLETE: 'true'
  });

  libraryStatus = HR_getLibrarySetupStatus_(spreadsheet);

  if (!libraryStatus.headersReady) {
    try {
      HR_ensureLibrarySheet_(spreadsheet);
      libraryStatus = HR_getLibrarySetupStatus_(spreadsheet);
    } catch (libraryError) {
      throw new Error(
        '우리학교 소장도서 시트 준비 중 오류가 발생했습니다: ' +
        String(libraryError && libraryError.message || libraryError)
      );
    }
  }

  try {
    HR_installEditTriggerForSpreadsheet_(spreadsheet);
  } catch (triggerError) {
    warnings.push('자동입력 트리거는 상단 메뉴의 설정 점검에서 다시 설치할 수 있습니다.');
  }

  return {
    success: true,
    message:
      schoolName + ' 설정을 완료했습니다.' +
      (yes24Ready
        ? ''
        : ' YES24 API 키는 「API 키 설정」에서 별도로 입력해 주세요.') +
      (warnings.length ? ' ' + warnings.join(' ') : ''),
    libraryHeadersReady: libraryStatus.headersReady,
    libraryBookCount: libraryStatus.bookCount,
    webAppUrl: HR_getConfiguredPublicWebAppUrl_(),
    warnings: warnings
  };
}

function HR_deleteCopySetupSheet() {
  var spreadsheet = HR_getSetupSpreadsheet_();
  var properties = PropertiesService.getScriptProperties();
  var configuredSpreadsheetId = String(
    properties.getProperty(HR_SPREADSHEET_ID_PROPERTY) || ''
  ).trim();
  var setupSheet = spreadsheet.getSheetByName(
    HR_COPY_SETUP_SHEET_NAME
  );
  var guideSheet = spreadsheet.getSheetByName('📗 희망도서 신청안내');

  if (
    configuredSpreadsheetId !== spreadsheet.getId() ||
    !HR_getYes24Key_()
  ) {
    throw new Error(
      'YES24 API 키와 학교 최초 설정을 먼저 완료해 주세요.'
    );
  }

  if (setupSheet) {
    if (guideSheet) {
      spreadsheet.setActiveSheet(guideSheet);
    }

    spreadsheet.deleteSheet(setupSheet);
  }

  return {
    success: true,
    message: '사본 최초설정 탭을 삭제했습니다.'
  };
}

function HR_createCopySetupSheet() {
  var spreadsheet = HR_getSetupSpreadsheet_();

  HR_createCopySetupSheet_(spreadsheet);

  var setupSheet = spreadsheet.getSheetByName(
    HR_COPY_SETUP_SHEET_NAME
  );

  spreadsheet.setActiveSheet(setupSheet);
  setupSheet.getRange('A3').activate();

  return {
    success: true,
    message: '사본 최초설정 탭을 준비했습니다.'
  };
}

/**
 * 기존 배포용 사본에서 이미지 버튼을 다시 만들지 않고
 * 초보자용 안내 문구와 YES24 바로가기만 빠르게 갱신합니다.
 */
function HR_refreshCopySetupBeginnerGuide() {
  var spreadsheet = HR_getSetupSpreadsheet_();
  HR_createCopySetupSheet_(spreadsheet);

  var sheet = spreadsheet.getSheetByName(
    HR_COPY_SETUP_SHEET_NAME
  );

  spreadsheet.setActiveSheet(sheet);
  sheet.getRange('A3').activate();

  return {
    success: true,
    message: '초보자용 8단계 안내를 맨 위로 옮겼습니다.'
  };
}

/**
 * 초보자용 8단계를 가장 먼저 보이게 배치한 사본 최초설정 화면입니다.
 */
function HR_createCopySetupSheet_(spreadsheet) {
  var sheet = spreadsheet.getSheetByName(
    HR_COPY_SETUP_SHEET_NAME
  );
  var libraryStatus = HR_getLibrarySetupStatus_(spreadsheet);
  var properties = PropertiesService.getScriptProperties();
  var configuredId = String(
    properties.getProperty(HR_SPREADSHEET_ID_PROPERTY) || ''
  ).trim();
  var yes24Ready = Boolean(HR_getYes24Key_());
  var nlkReady = Boolean(HR_getNlkKey_());
  var schoolReady = configuredId === spreadsheet.getId();
  var publicUrlReady = /\/exec(?:$|[?#])/.test(
    HR_getConfiguredPublicWebAppUrl_()
  );
  var detailSteps;

  if (!sheet) {
    sheet = spreadsheet.insertSheet(
      HR_COPY_SETUP_SHEET_NAME,
      0
    );
  }

  
  
    sheet.setFrozenRows(0);
  sheet.setFrozenColumns(0);
  HR_breakMergedRangesSafely_(sheet);
sheet.clear();
  sheet.clearConditionalFormatRules();
  sheet.setHiddenGridlines(true);
  sheet.setColumnWidth(1, 58);
  sheet.setColumnWidths(2, 11, 76);
  sheet
    .getRange('B3:L86')
    .setBackground('#FFFFFF')
    .setFontFamily('Jua')
    .setFontColor('#2E3A33')
    .setVerticalAlignment('middle')
    .setWrap(true);

  sheet.getRange('B3:L4').merge();
  sheet
    .getRange('B3')
    .setValue('⚙️ 다른 학교 사본용 최초 설정');
  sheet.getRange('B5:L5').merge();
  sheet
    .getRange('B5')
    .setValue(
      '처음 한 번만 아래 1번부터 8번까지 차례대로 따라 하세요.'
    );

  sheet.getRange('B7:L8').merge();
  sheet
    .getRange('B7')
    .setValue('👣 컴퓨터가 익숙하지 않아도 괜찮아요! 위에서부터 그대로 따라 하세요')
    .setBackground('#183D35')
    .setFontColor('#FFFFFF')
    .setFontSize(21)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  detailSteps = [
    {
      badge: 'B10:C13',
      text: 'D10:L13',
      number: '1',
      value:
        '사본 만들고 시작 버튼 누르기\n' +
        '① 원본 시트에서 「파일 → 사본 만들기」를 누르고 새 학교 이름으로 저장하세요.\n' +
        '② 새 파일이 열리면 5~10초 기다린 뒤 브라우저 새로고침 버튼을 한 번 누르세요.\n' +
        '③ 위쪽 「⚙️ 시작하기」 → 「✅ 프로그램 사용 시작」을 누르세요. 구글 허용 창이 뜨면 내 계정 → 「고급」 → 「이동(안전하지 않음)」 → 「허용」 순서로 누르세요. 처음 한 번만 하면 되고, 경고가 뜨는 건 정상이에요. 몇 초 뒤 「⚙️ 사본 설정」 메뉴가 나타납니다.'
    },
    {
      badge: 'B15:C18',
      text: 'D15:L18',
      number: '2',
      value:
        'YES24 API 키 받기\n' +
        '바로 아래 「YES24 API 키 발급 사이트 열기」를 누르세요. YES24에 로그인한 뒤 인증키 발급·관리 화면에서 API Key 전체를 복사하세요.'
    },
    {
      badge: 'B21:C24',
      text: 'D21:L24',
      number: '3',
      value:
        '복사한 API 키 붙여 넣기\n' +
        '화면 위쪽 「⚙️ 사본 설정」 → 「🔑 API 키 설정 (YES24 필수)」을 누르세요. YES24 API Key 칸에 붙여 넣고 저장하세요. 국립중앙도서관 키는 KDC 참고용이므로 비워도 됩니다.'
    },
    {
      badge: 'B26:C30',
      text: 'D26:L30',
      number: '4',
      value:
        'DLS에서 우리 학교 소장도서 반출하기\n' +
        'DLS 접속 → 「자료관리」 → 오른쪽 아래 「페이징 권수」를 우리 학교가 가진 전체 권수만큼 입력 → 「칼럼」 → 「칼럼 설정」 → 「ISBN」 체크 → 「반출」을 누르세요.\n' +
        '반출한 파일의 책 제목과 ISBN을 복사해 「우리학교 소장도서」 시트 2행부터 붙여 넣으세요. 이 정보가 있어야 중복 도서를 알려 줍니다.'
    },
    {
      badge: 'B32:C35',
      text: 'D32:L35',
      number: '5',
      value:
        '학교 이름 연결하기\n' +
        '화면 위쪽 「⚙️ 사본 설정」 → 「🏫 학교 최초 설정」을 누르세요. 학교 이름을 입력하고 실행한 뒤 완료 안내가 나올 때까지 기다리세요.'
    },
    {
      badge: 'B37:C42',
      text: 'D37:L42',
      number: '6',
      value:
        '신청 웹페이지 만들기\n' +
        '화면 위쪽 「확장 프로그램」 → 「Apps Script」를 누르세요. 새 화면에서 「배포」 → 「새 배포」 → 톱니바퀴 모양 「유형 선택」 → 「웹 앱」을 고릅니다.\n' +
        '실행 사용자는 「나」, 액세스 권한은 「모든 사용자」로 정하고 「배포」를 누르세요. 마지막에 나오는 /exec 주소를 복사하세요.'
    },
    {
      badge: 'B44:C47',
      text: 'D44:L47',
      number: '7',
      value:
        '공개 신청 링크 저장하기\n' +
        '다시 스프레드시트로 돌아와 「⚙️ 사본 설정」 → 「🔗 공개 신청 링크 등록」을 누르세요. 방금 복사한 /exec 주소를 붙여 넣고 「링크 저장하기」를 누르세요.'
    },
    {
      badge: 'B49:C52',
      text: 'D49:L52',
      number: '8',
      value:
        '마지막으로 시험하기\n' +
        '「사용설명서」 탭의 「📚 희망도서 신청하기」를 눌러 보세요. 책 제목을 하나 검색해 신청해 보고, 「희망도서 목록」에 들어오면 성공입니다. 마지막에는 「⚙️ 사본 설정」 → 「✅ 설정 한 번에 점검」도 눌러 보세요.'
    }
  ];

  HR_drawGuideSteps_(
    sheet,
    detailSteps,
    '#244D42',
    '#EAF3ED',
    '#A8C8BB'
  );

  sheet.getRange('B19:L19').merge();
  sheet
    .getRange('B19')
    .setFormula(
      '=HYPERLINK("' +
      HR_YES24_API_MANAGE_URL +
      '","🔗 YES24 API 키 발급 사이트 열기")'
    )
    .setBackground('#FBF1D6')
    .setFontColor('#7A5A12')
    .setFontSize(15)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setBorder(
      true,
      true,
      true,
      true,
      false,
      false,
      '#E3C47A',
      SpreadsheetApp.BorderStyle.SOLID
    );

  sheet.getRange('B54:L56').merge();
  sheet
    .getRange('B54')
    .setValue(
      '⚠️ 꼭 확인하세요\n' +
      '① 「희망도서 목록」 시트 이름은 변경하지 마세요. 새 목록으로 교체할 때는 기존 시트 이름을 바꾼 뒤, 새 시트를 반드시 「희망도서 목록」으로 만드세요.\n' +
      '② 「우리학교 소장도서」에는 책 제목과 ISBN13을 모두 넣어야 중복 소장 안내가 정확히 작동합니다.'
    )
    .setBackground('#FFF6DA')
    .setFontColor('#7A5A12')
    .setFontSize(14)
    .setFontWeight('bold')
    .setHorizontalAlignment('left')
    .setBorder(
      true,
      true,
      true,
      true,
      false,
      false,
      '#E3C47A',
      SpreadsheetApp.BorderStyle.SOLID
    );

  sheet.getRange('B58:L59').merge();
  sheet
    .getRange('B58')
    .setValue('🧰 설정 버튼 모음  ·  안내를 읽은 뒤 필요한 버튼을 누르세요')
    .setBackground('#183D35')
    .setFontColor('#FFFFFF')
    .setFontSize(20)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  sheet.getRange('B61:L63').merge();
  sheet
    .getRange('B61')
    .setValue(
      (yes24Ready ? '✅ ' : '') +
      '🔑 API 키 설정  ·  YES24 필수 / 국립중앙도서관 선택'
    );
  sheet.getRange('B65:L67').merge();
  sheet
    .getRange('B65')
    .setValue(
      (schoolReady ? '✅ ' : '') +
      '🏫 이 학교 최초 설정'
    );
  sheet.getRange('B69:L71').merge();
  sheet
    .getRange('B69')
    .setValue(
      (publicUrlReady ? '✅ ' : '') +
      '🔗 공개 신청 링크 등록'
    );

  ['B61:L63', 'B65:L67', 'B69:L71'].forEach(
    function(rangeA1) {
      sheet
        .getRange(rangeA1)
        .setBackground('#244D42')
        .setFontColor('#FFFFFF')
        .setFontSize(18)
        .setFontWeight('bold')
        .setHorizontalAlignment('center')
        .setBorder(
          true,
          true,
          true,
          true,
          false,
          false,
          '#B9852A',
          SpreadsheetApp.BorderStyle.SOLID_THICK
        );
    }
  );

  sheet.getRange('B73:L75').merge();
  sheet
    .getRange('B73')
    .setValue(
      '현재 점검  ·  YES24 ' +
      (yes24Ready ? '완료' : '필요') +
      '  ·  KDC 참고 API ' +
      (nlkReady ? '사용 중' : '선택') +
      '  ·  신청서 연결 ' +
      (schoolReady ? '완료' : '필요') +
      '  ·  소장도서 제목/ISBN13 ' +
      (libraryStatus.headersReady ? '완료' : '필요') +
      '  ·  공개 링크 ' +
      (publicUrlReady ? '완료' : '필요')
    )
    .setBackground('#F8F5ED')
    .setFontColor('#34564D')
    .setFontSize(13)
    .setHorizontalAlignment('center');

  sheet.getRange('B77:L79').merge();
  sheet
    .getRange('B77')
    .setValue('🗑️ 모든 설정을 마친 뒤 이 탭 삭제')
    .setBackground('#34564D')
    .setFontColor('#FFFFFF')
    .setFontSize(17)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  sheet.getRange('B81:L82').merge();
  sheet
    .getRange('B81')
    .setValue(
      'YES24은 도서 검색에 꼭 필요합니다. 국립중앙도서관 API는 KDC 참고용이라 없어도 됩니다.\n' +
      '이 시트의 프로그램은 저절로 최신 기능으로 바뀝니다. 따로 업데이트하지 않아도 돼요.'
    )
    .setBackground('#EAF3ED')
    .setFontColor('#244D42')
    .setFontSize(13)
    .setFontWeight('bold')
    .setHorizontalAlignment('left');

  sheet
    .getRange('B3:L4')
    .setBackground('#183D35')
    .setFontColor('#FFFFFF')
    .setFontSize(25)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  sheet
    .getRange('B5:L5')
    .setBackground('#F6EBCF')
    .setFontColor('#183D35')
    .setFontSize(14)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  [3, 4].forEach(function(rowNumber) {
    sheet.setRowHeight(rowNumber, 36);
  });
  [5, 19].forEach(function(rowNumber) {
    sheet.setRowHeight(rowNumber, 34);
  });
  [7, 8, 58, 59].forEach(function(rowNumber) {
    sheet.setRowHeight(rowNumber, 34);
  });
  [
    15, 16, 17, 18,
    21, 22, 23, 24, 32, 33, 34, 35,
    44, 45, 46, 47, 49, 50, 51, 52,
    61, 62, 63, 65, 66, 67, 69, 70, 71,
    73, 74, 75, 77, 78, 79
  ].forEach(function(rowNumber) {
    sheet.setRowHeight(rowNumber, 34);
  });
  [10, 11, 12, 13].forEach(function(rowNumber) {
    sheet.setRowHeight(rowNumber, 62);
  });
  [26, 27, 28, 29, 30].forEach(function(rowNumber) {
    sheet.setRowHeight(rowNumber, 40);
  });
  [37, 38, 39, 40, 41, 42].forEach(function(rowNumber) {
    sheet.setRowHeight(rowNumber, 35);
  });
  [54, 55, 56].forEach(function(rowNumber) {
    sheet.setRowHeight(rowNumber, 42);
  });
  [81, 82].forEach(function(rowNumber) {
    sheet.setRowHeight(rowNumber, 38);
  });
  [
    6, 9, 14, 20, 25, 31, 36, 43, 48, 53,
    57, 60, 64, 68, 72, 76, 80, 83, 84, 85, 86
  ].forEach(function(rowNumber) {
    sheet.setRowHeight(rowNumber, 10);
  });

  HR_addSetupButtonImages_(sheet);
}

function HR_createCopySetupSheetBeforeTopGuide_(spreadsheet) {
  var sheet = spreadsheet.getSheetByName(
    HR_COPY_SETUP_SHEET_NAME
  );
  var libraryStatus = HR_getLibrarySetupStatus_(spreadsheet);
  var properties = PropertiesService.getScriptProperties();
  var configuredId = String(
    properties.getProperty(HR_SPREADSHEET_ID_PROPERTY) || ''
  ).trim();
  var yes24Ready = Boolean(HR_getYes24Key_());
  var nlkReady = Boolean(HR_getNlkKey_());
  var schoolReady = configuredId === spreadsheet.getId();
  var publicUrlReady = /\/exec(?:$|[?#])/.test(
    HR_getConfiguredPublicWebAppUrl_()
  );
  var steps;

  if (!sheet) {
    sheet = spreadsheet.insertSheet(
      HR_COPY_SETUP_SHEET_NAME,
      0
    );
  }

  sheet.getRange('B3:L95').breakApart();
  sheet.clear();
  sheet.clearConditionalFormatRules();
  sheet.setHiddenGridlines(true);
  sheet.setColumnWidth(1, 58);
  sheet.setColumnWidths(2, 11, 76);
  sheet
    .getRange('B3:L95')
    .setBackground('#FFFFFF')
    .setFontFamily('Jua')
    .setFontColor('#27344A')
    .setVerticalAlignment('middle')
    .setWrap(true);

  sheet.getRange('B3:L4').merge();
  sheet
    .getRange('B3')
    .setValue('⚙️ 다른 학교 사본용 최초 설정');
  sheet.getRange('B5:L5').merge();
  sheet
    .getRange('B5')
    .setValue(
      '사본을 만든 뒤 이 화면의 1번부터 차례대로 따라 하세요. 메뉴가 안 보이면 새로고침하세요.'
    );

  steps = [
    {
      badge: 'B7:C8',
      text: 'D7:L8',
      number: '1',
      value:
        '원본에서 「파일 → 사본 만들기」를 누릅니다. 새 파일이 열리면 상단의 「⚙️ 사본 설정」 메뉴가 나타날 때까지 잠시 기다립니다.'
    },
    {
      badge: 'B15:C17',
      text: 'D15:L17',
      number: '3',
      value:
        'DLS에서 책 제목과 ISBN을 반출한 뒤, 「우리학교 소장도서」 시트의 2행부터 붙여 넣습니다.'
    },
    {
      badge: 'B23:C25',
      text: 'D23:L25',
      number: '5',
      value:
        '확장 프로그램 → Apps Script → 배포 → 새 배포 → 웹 앱에서 「나 / 모든 사용자」로 배포합니다.'
    },
    {
      badge: 'B27:C29',
      text: 'D27:L29',
      number: '6',
      value:
        '상단 「⚙️ 사본 설정 → 🔗 공개 신청 링크 등록」을 누르고, 배포한 /exec 주소를 붙여 넣습니다.'
    }
  ];

  HR_drawGuideSteps_(
    sheet,
    steps,
    '#DDA333',
    '#FFF9EF',
    '#E8D2AB'
  );

  sheet.getRange('B11:L13').merge();
  sheet
    .getRange('B11')
    .setValue(
      (yes24Ready ? '✅ 2' : '2') +
      '  API 키 설정  ·  YES24 필수 / 국립중앙도서관 선택'
    );

  sheet.getRange('B19:L21').merge();
  sheet
    .getRange('B19')
    .setValue(
      (schoolReady ? '✅ 4' : '4') +
      '  이 학교 최초 설정'
    );

  sheet.getRange('B14:L14').merge();
  sheet
    .getRange('B14')
    .setFormula(
      '=HYPERLINK("' +
      HR_YES24_API_MANAGE_URL +
      '","🔗 YES24 API 키 발급 사이트 열기")'
    )
    .setBackground('#EAF1FF')
    .setFontColor('#315F9F')
    .setFontSize(14)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setBorder(
      true,
      true,
      true,
      true,
      false,
      false,
      '#BFD0EE',
      SpreadsheetApp.BorderStyle.SOLID
    );

  ['B11:L13', 'B19:L21'].forEach(function(rangeA1) {
    sheet
      .getRange(rangeA1)
      .setBackground('#5277C7')
      .setFontColor('#FFFFFF')
      .setFontSize(18)
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setBorder(
        true,
        true,
        true,
        true,
        false,
        false,
        '#3E61AA',
        SpreadsheetApp.BorderStyle.SOLID_THICK
      );
  });

  sheet.getRange('B31:L33').merge();
  sheet
    .getRange('B31')
    .setValue(
      '⚠️ 주의사항\n' +
      '① 「희망도서 목록」 시트 이름은 변경하지 마세요. 새 목록으로 교체할 때는 기존 시트 이름을 바꾼 뒤, 새 시트를 반드시 「희망도서 목록」으로 만드세요.\n' +
      '② 「우리학교 소장도서」에는 책 제목과 ISBN13을 모두 넣어야 중복 소장 안내가 정확히 작동합니다.'
    )
    .setBackground('#FFF0F0')
    .setFontColor('#A33C3C')
    .setFontSize(14)
    .setFontWeight('bold')
    .setHorizontalAlignment('left')
    .setBorder(
      true,
      true,
      true,
      true,
      false,
      false,
      '#E8B7B7',
      SpreadsheetApp.BorderStyle.SOLID
    );

  sheet.getRange('B35:L37').merge();
  sheet
    .getRange('B35')
    .setValue(
      '현재 점검  ·  YES24 ' +
      (yes24Ready ? '완료' : '필요') +
      '  ·  KDC 참고 API ' +
      (nlkReady ? '사용 중' : '선택') +
      '  ·  신청서 연결 ' +
      (schoolReady ? '완료' : '필요') +
      '  ·  소장도서 제목/ISBN13 ' +
      (libraryStatus.headersReady ? '완료' : '필요') +
      '  ·  공개 링크 ' +
      (publicUrlReady ? '완료' : '필요')
    )
    .setBackground('#F4F6FA')
    .setFontColor('#3B4A64')
    .setFontSize(13)
    .setHorizontalAlignment('center');

  sheet.getRange('B39:L41').merge();
  sheet
    .getRange('B39')
    .setValue('🗑️ 모든 설정을 마친 뒤 이 탭 삭제')
    .setBackground('#6F7785')
    .setFontColor('#FFFFFF')
    .setFontSize(17)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  sheet.getRange('B43:L44').merge();
  sheet
    .getRange('B43')
    .setValue(
      'YES24은 도서 검색에 꼭 필요합니다. 국립중앙도서관 API는 KDC 참고용이라 없어도 됩니다.'
    )
    .setBackground('#EDF4FF')
    .setFontColor('#284E86')
    .setFontSize(13)
    .setFontWeight('bold')
    .setHorizontalAlignment('left');

  sheet
    .getRange('B3:L4')
    .setBackground('#7A5128')
    .setFontColor('#FFFFFF')
    .setFontSize(25)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  sheet
    .getRange('B5:L5')
    .setBackground('#F3E9DA')
    .setFontColor('#7A5128')
    .setFontSize(14)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  [3, 4].forEach(function(rowNumber) {
    sheet.setRowHeight(rowNumber, 36);
  });
  sheet.setRowHeight(5, 34);
  [
    7, 8, 11, 12, 13, 15, 16, 17,
    19, 20, 21, 23, 24, 25, 27, 28, 29,
    35, 36, 37, 39, 40, 41
  ].forEach(function(rowNumber) {
    sheet.setRowHeight(rowNumber, 34);
  });
  [31, 32, 33].forEach(function(rowNumber) {
    sheet.setRowHeight(rowNumber, 42);
  });
  sheet.setRowHeight(14, 32);
  [6, 9, 10, 18, 22, 26, 30, 34, 38, 42, 45]
    .forEach(function(rowNumber) {
      sheet.setRowHeight(rowNumber, 10);
    });
  [43, 44].forEach(function(rowNumber) {
    sheet.setRowHeight(rowNumber, 31);
  });

  HR_writeCopySetupBeginnerDetails_(sheet);
  HR_addSetupButtonImages_(sheet);
}

/**
 * 컴퓨터가 익숙하지 않은 사용자를 위한 상세 설명입니다.
 */
function HR_writeCopySetupBeginnerDetails_(sheet) {
  var detailSteps = [
    {
      badge: 'B50:C53',
      text: 'D50:L53',
      number: '1',
      value:
        '사본을 연 직후\n' +
        '새 파일이 열리면 5~10초 정도 기다리세요. 화면 위쪽에 「⚙️ 사본 설정」 메뉴가 보이면 준비 완료입니다. 안 보이면 브라우저의 새로고침 버튼을 한 번 누르세요.'
    },
    {
      badge: 'B55:C58',
      text: 'D55:L58',
      number: '2',
      value:
        'YES24 API 키 받기\n' +
        '위쪽의 파란 「YES24 API 키 발급 사이트 열기」를 누르세요. YES24에 로그인한 뒤 인증키 발급·관리 화면에서 API Key 전체를 복사하세요.'
    },
    {
      badge: 'B60:C63',
      text: 'D60:L63',
      number: '3',
      value:
        '복사한 API 키 붙여 넣기\n' +
        '다시 이 시트로 돌아와 위쪽 「⚙️ 사본 설정」 → 「🔑 API 키 설정 (YES24 필수)」을 누르세요. YES24 API Key 칸에 붙여 넣고 저장하세요. 국립중앙도서관 키는 비워도 됩니다.'
    },
    {
      badge: 'B65:C68',
      text: 'D65:L68',
      number: '4',
      value:
        'DLS에서 우리 학교 소장도서 반출하기\n' +
        'DLS 접속 → 「자료관리」 → 오른쪽 아래 「페이징 권수」를 우리 학교가 가진 전체 권수만큼 입력 → 「칼럼」 → 「칼럼 설정」 → 「ISBN」 체크 → 「반출」을 누르세요.\n' +
        '반출한 파일의 책 제목과 ISBN을 복사해 「우리학교 소장도서」 시트 2행부터 붙여 넣으세요. 이 정보가 있어야 중복 도서를 알려 줍니다.'
    },
    {
      badge: 'B70:C73',
      text: 'D70:L73',
      number: '5',
      value:
        '학교 이름 연결하기\n' +
        '위쪽 「⚙️ 사본 설정」 → 「🏫 학교 최초 설정」을 누르세요. 학교 이름을 입력하고 실행한 뒤, 완료 안내가 나올 때까지 기다리세요.'
    },
    {
      badge: 'B75:C80',
      text: 'D75:L80',
      number: '6',
      value:
        '신청 웹페이지 만들기\n' +
        '위쪽 「확장 프로그램」 → 「Apps Script」를 누르세요. 새 화면에서 「배포」 → 「새 배포」 → 톱니바퀴 모양 「유형 선택」 → 「웹 앱」을 고릅니다.\n' +
        '실행 사용자는 「나」, 액세스 권한은 「모든 사용자」로 정하고 「배포」를 누르세요. 마지막에 나오는 /exec 주소를 복사하세요.'
    },
    {
      badge: 'B82:C85',
      text: 'D82:L85',
      number: '7',
      value:
        '공개 신청 링크 저장하기\n' +
        '다시 스프레드시트로 돌아와 「⚙️ 사본 설정」 → 「🔗 공개 신청 링크 등록」을 누르세요. 방금 복사한 /exec 주소를 붙여 넣고 「링크 저장하기」를 누르세요.'
    },
    {
      badge: 'B87:C90',
      text: 'D87:L90',
      number: '8',
      value:
        '마지막으로 시험하기\n' +
        '「사용설명서」 탭의 「📚 희망도서 신청하기」를 눌러 보세요. 책 제목을 하나 검색해 신청해 보고, 「희망도서 목록」에 들어오면 성공입니다. 마지막에는 「⚙️ 사본 설정」 → 「✅ 설정 한 번에 점검」도 눌러 보세요.'
    }
  ];

  sheet.getRange('B47:L92').breakApart().clear();
  sheet
    .getRange('B47:L92')
    .setBackground('#FFFFFF')
    .setFontFamily('Jua')
    .setFontColor('#27344A')
    .setVerticalAlignment('middle')
    .setWrap(true);

  sheet.getRange('B47:L48').merge();
  sheet
    .getRange('B47')
    .setValue('👣 컴퓨터가 익숙하지 않다면 아래 8단계를 그대로 따라 하세요')
    .setBackground('#243B64')
    .setFontColor('#FFFFFF')
    .setFontSize(21)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  HR_drawGuideSteps_(
    sheet,
    detailSteps,
    '#5277C7',
    '#F3F6FC',
    '#D8E1F0'
  );

  [47, 48].forEach(function(rowNumber) {
    sheet.setRowHeight(rowNumber, 34);
  });

  [
    50, 51, 52, 53,
    55, 56, 57, 58,
    60, 61, 62, 63,
    65, 66, 67, 68,
    70, 71, 72, 73,
    75, 76, 77, 78, 79, 80,
    82, 83, 84, 85,
    87, 88, 89, 90
  ].forEach(function(rowNumber) {
    sheet.setRowHeight(rowNumber, 31);
  });

  [65, 66, 67, 68].forEach(function(rowNumber) {
    sheet.setRowHeight(rowNumber, 40);
  });

  [49, 54, 59, 64, 69, 74, 81, 86, 91, 92]
    .forEach(function(rowNumber) {
      sheet.setRowHeight(rowNumber, 10);
    });
}

function HR_createCopySetupSheetLegacy_(spreadsheet) {
  var sheet = spreadsheet.getSheetByName(
    HR_COPY_SETUP_SHEET_NAME
  );
  var libraryStatus = HR_getLibrarySetupStatus_(spreadsheet);
  var properties = PropertiesService.getScriptProperties();
  var configuredId = String(
    properties.getProperty(HR_SPREADSHEET_ID_PROPERTY) || ''
  ).trim();
  var yes24Ready = Boolean(HR_getYes24Key_());
  var nlkReady = Boolean(HR_getNlkKey_());
  var schoolReady = configuredId === spreadsheet.getId();

  if (!sheet) {
    sheet = spreadsheet.insertSheet(
      HR_COPY_SETUP_SHEET_NAME,
      0
    );
  }

  sheet.clear();
  sheet.clearConditionalFormatRules();
  sheet.setHiddenGridlines(true);
  sheet.setColumnWidth(1, 58);
  sheet.setColumnWidths(2, 11, 76);

  sheet.getRange('B3:L4').merge();
  sheet
    .getRange('B3')
    .setValue('⚙️ 다른 학교 사본용 최초 설정');

  sheet.getRange('B5:L5').merge();
  sheet
    .getRange('B5')
    .setValue(
      '① API 키 설정  →  ② 학교 최초 설정  →  ③ 소장도서 붙여넣기'
    );

  sheet.getRange('B7:L9').merge();
  sheet
    .getRange('B7')
    .setValue(
      (yes24Ready ? '✅' : '1️⃣') +
      ' 클릭: API 키 설정  ·  YES24 필수 / 국립중앙도서관 선택'
    );

  sheet.getRange('B12:L14').merge();
  sheet
    .getRange('B12')
    .setValue(
      (schoolReady ? '✅' : '2️⃣') +
      ' 클릭: 이 학교 최초 설정'
    );

  sheet.getRange('B17:L19').merge();
  sheet
    .getRange('B17')
    .setValue('🗑️ 설정 완료 후 이 탭 삭제');

  sheet.getRange('B21:L23').merge();
  sheet
    .getRange('B21')
    .setValue(
      '필수: 「우리학교 소장도서」 시트에 책 제목과 ISBN13을 입력해야 중복 도서 확인이 정확히 작동합니다.'
    );

  sheet.getRange('B25:L27').merge();
  sheet
    .getRange('B25')
    .setValue(
      '현재 점검  ·  YES24 API ' +
      (yes24Ready ? '완료' : '필요') +
      '  ·  KDC 참고 API ' +
      (nlkReady ? '사용 중' : '선택') +
      '  ·  신청서 연결 ' +
      (schoolReady ? '완료' : '필요') +
      '  ·  소장도서 제목/ISBN 열 ' +
      (libraryStatus.headersReady ? '완료' : '필요')
    );

  sheet
    .getRange('B3:L27')
    .setFontFamily('Jua')
    .setVerticalAlignment('middle')
    .setWrap(true);

  sheet
    .getRange('B3:L4')
    .setBackground('#243B64')
    .setFontColor('#FFFFFF')
    .setFontSize(25)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  sheet
    .getRange('B5:L5')
    .setBackground('#E9EEF8')
    .setFontColor('#243B64')
    .setFontSize(13)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  ['B7:L9', 'B12:L14'].forEach(function(rangeA1) {
    sheet
      .getRange(rangeA1)
      .setBackground('#5277C7')
      .setFontColor('#FFFFFF')
      .setFontSize(18)
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setBorder(
        true,
        true,
        true,
        true,
        false,
        false,
        '#3E61AA',
        SpreadsheetApp.BorderStyle.SOLID_THICK
      );
  });

  sheet
    .getRange('B17:L19')
    .setBackground('#6F7785')
    .setFontColor('#FFFFFF')
    .setFontSize(17)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  sheet
    .getRange('B21:L23')
    .setBackground('#FFF0F0')
    .setFontColor('#A33C3C')
    .setFontSize(15)
    .setFontWeight('bold')
    .setHorizontalAlignment('left')
    .setBorder(
      true,
      true,
      true,
      true,
      false,
      false,
      '#E8B7B7',
      SpreadsheetApp.BorderStyle.SOLID
    );

  sheet
    .getRange('B25:L27')
    .setBackground('#F4F6FA')
    .setFontColor('#3B4A64')
    .setFontSize(13)
    .setHorizontalAlignment('center');

  [3, 4].forEach(function(rowNumber) {
    sheet.setRowHeight(rowNumber, 36);
  });
  [7, 8, 9, 12, 13, 14, 17, 18, 19].forEach(
    function(rowNumber) {
      sheet.setRowHeight(rowNumber, 34);
    }
  );
  [21, 22, 23, 25, 26, 27].forEach(
    function(rowNumber) {
      sheet.setRowHeight(rowNumber, 31);
    }
  );

  HR_addSetupButtonImages_(sheet);
}

/**
 * 투명한 이미지 버튼을 셀 위에 올려 파란 영역 전체를 클릭할 수 있게 합니다.
 * 사본을 만들면 버튼에 연결된 함수도 함께 복사됩니다.
 */
function HR_prepareTopSetupButtons_(sheet) {
  var lastRow = sheet.getLastRow();
  var lastColumn = Math.max(sheet.getLastColumn(), 15);
  var values = sheet.getRange(1, 1, lastRow, lastColumn).getDisplayValues();
  var startRow = 0;
  var endRow = 0;
  var headerText = '🧰 설정 버튼 모음  ·  필요한 버튼을 순서대로 누르세요';
  var apiReady = false;
  var schoolReady = false;
  var publicReady = false;

  values.forEach(function(row, index) {
    row.forEach(function(cellValue) {
      var cellText = String(cellValue || '').trim();
      if (/설정 버튼 모음/.test(cellText)) {
        if (!startRow) startRow = index + 1;
        headerText = cellText;
      }
      if (/API 키 설정/.test(cellText) && /✅/.test(cellText)) apiReady = true;
      if (/(?:이 학교|학교) 최초 설정/.test(cellText) && /✅/.test(cellText)) schoolReady = true;
      if (/공개 (?:신청 )?링크 등록/.test(cellText)) {
        if (/✅/.test(cellText)) publicReady = true;
        endRow = Math.max(endRow, index + 1);
      }
    });
  });

  if (startRow > 8 && endRow >= startRow) {
    var merged = sheet.getRange(endRow, 1, 1, lastColumn).getMergedRanges();
    merged.forEach(function(range) {
      endRow = Math.max(endRow, range.getLastRow());
    });
    endRow = Math.min(lastRow, endRow + 1);
    sheet.moveRows(
      sheet.getRange(startRow, 1, endRow - startRow + 1, sheet.getMaxColumns()),
      7
    );
  }

  HR_breakMergedRangesSafely_(sheet, 'B7:O20');
  sheet.getRange('B7:O20').clearContent();
  sheet.getRange('B7:L8').merge().setValue(headerText);
  sheet.getRange('B10:D13').merge().setValue((apiReady ? '✅ ' : '') + '🔑 API 키 설정');
  sheet.getRange('F10:H13').merge().setValue((schoolReady ? '✅ ' : '') + '🏫 학교 최초 설정');
  sheet.getRange('J10:L13').merge().setValue((publicReady ? '✅ ' : '') + '🔗 공개 신청 링크 등록');

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
}

function HR_styleCopySetupModern_(sheet) {
  if (!sheet) return;

  var lastRow = Math.max(sheet.getLastRow(), 82);
  var lastColumn = Math.max(sheet.getLastColumn(), 12);
  var canvas = sheet.getRange(1, 1, lastRow, lastColumn);

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
    sheet.setColumnWidth(column, column <= 12 ? 86 : 72);
  }
  sheet.setColumnWidths(2, 3, 100);
  sheet.setColumnWidth(5, 28);
  sheet.setColumnWidths(6, 3, 100);
  sheet.setColumnWidth(9, 28);
  sheet.setColumnWidths(10, 3, 100);

  canvas.getMergedRanges().forEach(function(range) {
    var text = String(range.getDisplayValue() || '').trim();

    range
      .setBackground('#FFFFFF')
      .setFontColor('#2E3A33')
      .setFontFamily('Arial')
      .setFontSize(12)
      .setVerticalAlignment('middle')
      .setWrap(true)
      .setBorder(
        true, true, true, true, false, false,
        '#D9E2DC', SpreadsheetApp.BorderStyle.SOLID
      );

    if (!text) {
      range
        .setBackground('#F8F5ED')
        .setBorder(false, false, false, false, false, false);
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

    if (/컴퓨터가 익숙하지 않아도|설정 버튼 모음/.test(text)) {
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

    if (/API 키 설정|이 학교 최초 설정|학교 최초 설정|공개 신청 링크 등록|공개 링크 등록/.test(text)) {
      range
        .setBackground('#527D6B')
        .setFontColor('#FFFFFF')
        .setFontSize(16)
        .setFontWeight('bold')
        .setHorizontalAlignment('center')
        .setBorder(
          true, true, true, true, false, false,
          '#315D4D', SpreadsheetApp.BorderStyle.SOLID_MEDIUM
        );
      return;
    }

    if (/사이트 열기|메뉴 확인/.test(text)) {
      range
        .setBackground('#E7F0EC')
        .setFontColor('#285244')
        .setFontWeight('bold')
        .setHorizontalAlignment('center');
      return;
    }

    if (/꼭 확인|주의|변경 금지|삭제하지/.test(text)) {
      range
        .setBackground('#FCEDEC')
        .setFontColor('#A33F3F')
        .setFontWeight('bold');
      return;
    }

    if (/현재 점검/.test(text)) {
      range
        .setBackground('#EEF4F1')
        .setFontColor('#34564D')
        .setFontWeight('bold')
        .setHorizontalAlignment('center');
    }
  });

  sheet.activate();
  sheet.setActiveRange(sheet.getRange('B3'));
}

function HR_addSetupButtonImages_(sheet) {
  var transparentPngBase64 =
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJ' +
    'AAAACXBIWXMAAAPoAAAD6AG1e1JrAAAADUlEQVR4nGP4' +
    '//8/MwAI/wMBlb/HmAAAAABJRU5ErkJggg==';
  var buttonBlob = Utilities.newBlob(
    Utilities.base64Decode(transparentPngBase64),
    'image/png',
    'copy-setup-button.png'
  );
  var images = sheet.getImages();
  var imageIndex = 0;

  HR_prepareTopSetupButtons_(sheet);

  for (imageIndex = images.length - 1; imageIndex >= 0; imageIndex--) {
    if (
      String(images[imageIndex].getAltTextTitle() || '')
        .indexOf('HR_COPY_SETUP_BUTTON') === 0
    ) {
      images[imageIndex].remove();
    }
  }

  [
    {
      cell: 'B10',
      width: 300,
      height: 104,
      title: 'API 키 설정',
      functionName: 'HR_showYes24KeyDialog'
    },
    {
      cell: 'F10',
      width: 300,
      height: 104,
      title: '학교 최초 설정',
      functionName: 'HR_showSchoolSetupDialog'
    },
    {
      cell: 'J10',
      width: 300,
      height: 104,
      title: '공개 신청 링크 등록',
      functionName: 'HR_showPublicWebAppUrlDialog'
    },
    {
      cell: 'B77',
      width: 836,
      height: 102,
      title: '설정 탭 삭제',
      functionName: 'HR_deleteCopySetupSheet'
    }
  ].forEach(function(config, index) {
    var anchor = sheet.getRange(config.cell);
    var image = sheet.insertImage(
      buttonBlob,
      anchor.getColumn(),
      anchor.getRow()
    );

    image
      .setWidth(config.width)
      .setHeight(config.height)
      .setAltTextTitle('HR_COPY_SETUP_BUTTON_' + (index + 1))
      .setAltTextDescription(config.title)
      .assignScript(config.functionName);
  });

  HR_styleCopySetupModern_(sheet);
}

function HR_getSetupSpreadsheet_() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  if (spreadsheet) {
    return spreadsheet;
  }

  return HR_getSpreadsheet_();
}

function HR_ensureLibrarySheet_(spreadsheet) {
  var sheet = spreadsheet.getSheetByName(
    HR_LIBRARY_SHEET_NAME
  );
  var lastColumn;
  var headers;
  var titleColumn;
  var isbnColumn;

  if (!sheet) {
    sheet = spreadsheet.insertSheet(HR_LIBRARY_SHEET_NAME);
  }

  lastColumn = Math.max(sheet.getLastColumn(), 1);
  headers = sheet
    .getRange(1, 1, 1, lastColumn)
    .getDisplayValues()[0];
  titleColumn = HR_findAnyHeaderColumn_(
    headers,
    ['책제목', '도서명', '서명', '제목']
  );
  isbnColumn = HR_findAnyHeaderColumn_(
    headers,
    ['ISBN13', 'ISBN']
  );

  if (!titleColumn) {
    lastColumn = Math.max(sheet.getLastColumn(), 0) + 1;
    sheet.getRange(1, lastColumn).setValue('책 제목');
  }

  if (!isbnColumn) {
    lastColumn = Math.max(sheet.getLastColumn(), 0) + 1;
    sheet.getRange(1, lastColumn).setValue('ISBN13');
  }

  sheet
    .getRange(1, 1, 1, sheet.getLastColumn())
    .setBackground('#243B64')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  sheet.setFrozenRows(1);
  sheet.setColumnWidths(1, sheet.getLastColumn(), 180);
}

function HR_getLibrarySetupStatus_(spreadsheet) {
  var sheet = spreadsheet.getSheetByName(
    HR_LIBRARY_SHEET_NAME
  );
  var result = {
    sheetReady: Boolean(sheet),
    headersReady: false,
    titleColumn: 0,
    isbnColumn: 0,
    bookCount: 0
  };
  var lastColumn;
  var headers;

  if (!sheet) {
    return result;
  }

  lastColumn = Math.max(sheet.getLastColumn(), 1);
  headers = sheet
    .getRange(1, 1, 1, lastColumn)
    .getDisplayValues()[0];
  result.titleColumn = HR_findAnyHeaderColumn_(
    headers,
    ['책제목', '도서명', '서명', '제목']
  );
  result.isbnColumn = HR_findAnyHeaderColumn_(
    headers,
    ['ISBN13', 'ISBN']
  );
  result.headersReady = Boolean(
    result.titleColumn && result.isbnColumn
  );
  result.bookCount = Math.max(sheet.getLastRow() - 1, 0);

  return result;
}

function HR_installEditTriggerForSpreadsheet_(spreadsheet) {
  var triggers = ScriptApp.getProjectTriggers();
  var triggerIndex = 0;

  for (
    triggerIndex = 0;
    triggerIndex < triggers.length;
    triggerIndex++
  ) {
    if (
      triggers[triggerIndex].getHandlerFunction() ===
      'onEditTriggerAction'
    ) {
      ScriptApp.deleteTrigger(triggers[triggerIndex]);
    }
  }

  ScriptApp
    .newTrigger('onEditTriggerAction')
    .forSpreadsheet(spreadsheet)
    .onEdit()
    .create();
}

function HR_testYes24Key_(key) {
  yes24Get_('/goods/itemList', {query: '아몬드', category: 'BOOK', pageSize: 1}, key);
}

/**
 * 학교별 필수 설정을 한 번에 확인합니다.
 */
function HR_getSystemCheckState_() {
  var spreadsheet = HR_getSetupSpreadsheet_();
  var properties = PropertiesService.getScriptProperties();
  var targetSheet = spreadsheet.getSheetByName(
    HR_TARGET_SHEET_NAME
  );
  var libraryStatus = HR_getLibrarySetupStatus_(spreadsheet);
  var configuredId = String(
    properties.getProperty(HR_SPREADSHEET_ID_PROPERTY) || ''
  ).trim();
  var publicUrl = HR_getConfiguredPublicWebAppUrl_();
  var checks = [];

  checks.push({
    label: 'YES24 API 키',
    ok: Boolean(HR_getYes24Key_()),
    detail: HR_getYes24Key_() ? '설정됨' : '설정 필요'
  });
  checks.push({
    label: '학교 신청서 연결',
    ok: configuredId === spreadsheet.getId(),
    detail: configuredId === spreadsheet.getId()
      ? '현재 파일에 연결됨'
      : '학교 최초 설정 필요'
  });
  checks.push({
    label: '📚 희망도서 목록',
    ok: Boolean(targetSheet),
    detail: targetSheet ? '정상' : '시트 없음'
  });
  checks.push({
    label: '🏫 우리학교 소장도서',
    ok: Boolean(
      libraryStatus.sheetReady &&
      libraryStatus.headersReady
    ),
    detail: libraryStatus.headersReady
      ? libraryStatus.bookCount + '권 확인'
      : '책 제목·ISBN13 열 확인 필요'
  });
  checks.push({
    label: '공개 신청 링크',
    ok: /\/exec(?:$|\?)/.test(publicUrl),
    detail: /\/exec(?:$|\?)/.test(publicUrl)
      ? '누구나 신청 가능한 주소'
      : '공개 /exec 주소 확인 필요'
  });
  return {
    checks: checks,
    allRequiredReady: checks
      .slice(0, 5)
      .every(function(check) {
        return check.ok;
      })
  };
}

function HR_showSystemCheck() {
  var state = HR_getSystemCheckState_();
  var checkHtml = state.checks
    .map(function(check) {
      return (
        '<div class="item ' +
        (check.ok ? 'ok' : 'warn') +
        '">' +
          '<span class="icon">' +
          (check.ok ? '✓' : '!') +
          '</span>' +
          '<div><strong>' +
          HR_escapeHtml_(check.label) +
          '</strong><p>' +
          HR_escapeHtml_(check.detail) +
          '</p></div>' +
        '</div>'
      );
    })
    .join('');
  var html =
    '<!doctype html><html lang="ko"><head><base target="_top">' +
    '<style>' +
    '*{box-sizing:border-box}body{margin:0;padding:22px;background:#f5f7fb;' +
    'font-family:-apple-system,BlinkMacSystemFont,"Apple SD Gothic Neo",' +
    '"Noto Sans KR",sans-serif;color:#24324a}' +
    'h1{margin:0 0 8px;font-size:24px}.lead{margin:0 0 18px;color:#66758e;' +
    'line-height:1.6}.list{display:grid;gap:10px}.item{display:flex;gap:12px;' +
    'align-items:center;padding:14px;border-radius:14px;background:#fff;' +
    'border:1px solid #dce4f2}.icon{display:flex;width:34px;height:34px;' +
    'flex:0 0 auto;align-items:center;justify-content:center;border-radius:12px;' +
    'font-weight:900}.ok .icon{background:#eaf7ef;color:#237a4b}' +
    '.warn .icon{background:#fff3df;color:#a46608}.item strong{font-size:15px}' +
    '.item p{margin:4px 0 0;color:#6b7688;font-size:13px}.summary{margin-top:16px;' +
    'padding:15px;border-radius:14px;font-weight:900;text-align:center}' +
    '.ready{background:#eaf7ef;color:#237a4b}.check{background:#fff3df;color:#8a5a12}' +
    'button{width:100%;height:48px;margin-top:16px;border:0;border-radius:12px;' +
    'background:#4f78d1;color:white;font-size:16px;font-weight:900;cursor:pointer}' +
    '</style></head><body>' +
    '<h1>✅ 설정 한 번에 점검</h1>' +
    '<p class="lead">학교 신청 시스템의 필수 설정을 한눈에 확인합니다.</p>' +
    '<div class="list">' + checkHtml + '</div>' +
    '<div class="summary ' +
    (state.allRequiredReady ? 'ready' : 'check') +
    '">' +
    (state.allRequiredReady
      ? '필수 설정이 모두 정상입니다.'
      : '표시된 항목을 확인해 주세요.') +
    '</div><button onclick="google.script.host.close()">확인</button>' +
    '</body></html>';

  SpreadsheetApp.getUi().showModalDialog(
    HtmlService
      .createHtmlOutput(html)
      .setWidth(520)
      .setHeight(660),
    '설정 점검'
  );
}

function HR_escapeHtml_(value) {
  return String(value || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

/**
 * ============================================================
 * 공개 신청 웹앱 베스트셀러
 * - 종합 / 청소년 / 어린이
 * - YES24 베스트셀러 API
 * - 24시간 단위 갱신
 * ============================================================
 */

/**
 * 공개 신청 페이지에 최근 희망도서 신청 목록을 제공합니다.
 * 신청자 이름은 개인정보 보호를 위해 절대 반환하지 않습니다.
 */
function HR_getSubmittedRequests(audience, staffCode) {
  var audienceKey = HR_audienceKey_(audience);
  HR_requireStaffAccess_(audienceKey, staffCode);
  var spreadsheet = HR_getSpreadsheet_();
  var sheet = spreadsheet.getSheetByName(
    HR_TARGET_SHEET_NAME
  );
  var headers;
  var lastRow;
  var lastColumn;
  var titleColumn;
  var authorColumn;
  var publisherColumn;
  var yearColumn;
  var statusColumn;
  var dateColumn;
  var sequenceColumn;
  var isbnColumn;
  var priceColumn;
  var categoryColumn;
  var pagesColumn;
  var linkColumn;
  var coverColumn;
  var values;
  var formulas;
  var items = [];
  var rowIndex;

  if (!sheet) {
    throw new Error(
      HR_TARGET_SHEET_NAME + ' 시트를 찾지 못했습니다.'
    );
  }

  lastRow = sheet.getLastRow();
  lastColumn = sheet.getLastColumn();

  if (lastRow < FIRST_DATA_ROW || lastColumn < 1) {
    return {
      totalCount: 0,
      shownCount: 0,
      items: [],
      updatedAt: new Date().toISOString()
    };
  }

  headers = getHeaders(sheet);
  titleColumn = HR_findAnyHeaderColumn_(
    headers,
    ['도서명', '책 제목', '책제목', '서명', '제목']
  );
  authorColumn = HR_findAnyHeaderColumn_(
    headers,
    ['저자', '지은이']
  );
  publisherColumn = HR_findAnyHeaderColumn_(
    headers,
    ['출판사', '발행처']
  );
  yearColumn = HR_findAnyHeaderColumn_(
    headers,
    ['연도', '출판연도', '발행연도']
  );
  statusColumn = HR_findAnyHeaderColumn_(
    headers,
    ['처리상태', '처리 상태', '상태']
  );
  dateColumn = HR_findAnyHeaderColumn_(
    headers,
    ['신청일', '신청일자', '접수일', '등록일']
  );
  sequenceColumn = HR_findAnyHeaderColumn_(
    headers,
    ['순번', '번호']
  );
  isbnColumn = HR_findAnyHeaderColumn_(
    headers,
    ['ISBN13', 'ISBN', 'isbn13', 'isbn']
  );
  priceColumn = HR_findAnyHeaderColumn_(
    headers,
    ['정가', '가격', '판매가']
  );
  categoryColumn = HR_findAnyHeaderColumn_(
    headers,
    ['분야', '카테고리', '주제', 'KDC']
  );
  pagesColumn = HR_findAnyHeaderColumn_(
    headers,
    ['쪽수', '페이지', '전체 페이지']
  );
  linkColumn = HR_findAnyHeaderColumn_(
    headers,
    ['YES24 링크', '도서 링크', '링크']
  );
  coverColumn = HR_findAnyHeaderColumn_(
    headers,
    ['표지', '책 표지', '표지 URL', '표지URL', '이미지']
  );

  if (!titleColumn) {
    throw new Error(
      '희망도서 목록에서 도서명 열을 찾지 못했습니다.'
    );
  }

  values = sheet
    .getRange(
      FIRST_DATA_ROW,
      1,
      lastRow - FIRST_DATA_ROW + 1,
      lastColumn
    )
    .getDisplayValues();

  formulas = sheet
    .getRange(
      FIRST_DATA_ROW,
      1,
      lastRow - FIRST_DATA_ROW + 1,
      lastColumn
    )
    .getFormulas();

  var audienceColumn = HR_findAudienceColumn_(headers);

  for (rowIndex = 0; rowIndex < values.length; rowIndex++) {
    var row = values[rowIndex];
    var title = String(
      row[titleColumn - 1] || ''
    ).trim();

    if (!title) {
      continue;
    }

    // 다른 구분(교직원/학부모)의 신청은 보여 주지 않습니다. 구분이 비어 있으면 교직원입니다.
    if (
      (audienceColumn ? HR_audienceKey_(row[audienceColumn - 1]) : 'staff') !==
      audienceKey
    ) {
      continue;
    }

    items.push({
      row: rowIndex + FIRST_DATA_ROW,
      sequence: sequenceColumn
        ? String(row[sequenceColumn - 1] || '').trim()
        : String(rowIndex + 1),
      title: title,
      author: authorColumn
        ? String(row[authorColumn - 1] || '').trim()
        : '',
      publisher: publisherColumn
        ? String(row[publisherColumn - 1] || '').trim()
        : '',
      year: yearColumn
        ? String(row[yearColumn - 1] || '').trim()
        : '',
      status: statusColumn
        ? String(row[statusColumn - 1] || '').trim()
        : '신청 접수',
      requestedAt: dateColumn
        ? String(row[dateColumn - 1] || '').trim()
        : '',
      isbn: isbnColumn
        ? String(row[isbnColumn - 1] || '').trim()
        : '',
      price: priceColumn
        ? String(row[priceColumn - 1] || '').trim()
        : '',
      category: categoryColumn
        ? String(row[categoryColumn - 1] || '').trim()
        : '',
      pages: pagesColumn
        ? String(row[pagesColumn - 1] || '').trim()
        : '',
      link: linkColumn
        ? String(row[linkColumn - 1] || '').trim()
        : '',
      cover: coverColumn
        ? HR_extractRequestCoverUrl_(
            formulas[rowIndex][coverColumn - 1],
            row[coverColumn - 1]
          )
        : ''
    });
  }

  items = items.concat(
    HR_getPendingCancelledItems_(
      spreadsheet,
      headers,
      audienceKey
    )
  );

  return {
    totalCount: items.length,
    shownCount: Math.min(items.length, 100),
    items: items.slice(-100).reverse(),
    updatedAt: new Date().toISOString()
  };
}


function HR_getPendingCancelledItems_(
  spreadsheet,
  sourceHeaders,
  audienceKey
) {
  var backupSheet = spreadsheet.getSheetByName(
    HR_REQUEST_CANCEL_BACKUP_SHEET_NAME
  );
  var headers = Array.isArray(sourceHeaders)
    ? sourceHeaders
    : [];
  var items = [];
  var lastRow;
  var readWidth;
  var values;
  var formulas;
  var titleColumn;
  var authorColumn;
  var publisherColumn;
  var yearColumn;
  var dateColumn;
  var sequenceColumn;
  var isbnColumn;
  var priceColumn;
  var categoryColumn;
  var pagesColumn;
  var linkColumn;
  var coverColumn;
  var rowIndex;

  if (
    !backupSheet ||
    backupSheet.getLastRow() < 2 ||
    backupSheet.getLastColumn() < 8
  ) {
    return items;
  }

  lastRow = backupSheet.getLastRow();
  readWidth = Math.min(
    backupSheet.getLastColumn(),
    8 + headers.length
  );
  values = backupSheet
    .getRange(2, 1, lastRow - 1, readWidth)
    .getDisplayValues();
  formulas = backupSheet
    .getRange(2, 1, lastRow - 1, readWidth)
    .getFormulas();

  titleColumn = HR_findAnyHeaderColumn_(
    headers,
    ['도서명', '책 제목', '책제목', '서명', '제목']
  );
  authorColumn = HR_findAnyHeaderColumn_(
    headers,
    ['저자', '지은이']
  );
  publisherColumn = HR_findAnyHeaderColumn_(
    headers,
    ['출판사', '발행처']
  );
  yearColumn = HR_findAnyHeaderColumn_(
    headers,
    ['연도', '출판연도', '발행연도']
  );
  dateColumn = HR_findAnyHeaderColumn_(
    headers,
    ['신청일', '신청일자', '접수일', '등록일']
  );
  sequenceColumn = HR_findAnyHeaderColumn_(
    headers,
    ['순번', '번호']
  );
  isbnColumn = HR_findAnyHeaderColumn_(
    headers,
    ['ISBN13', 'ISBN', 'isbn13', 'isbn']
  );
  priceColumn = HR_findAnyHeaderColumn_(
    headers,
    ['정가', '가격', '판매가']
  );
  categoryColumn = HR_findAnyHeaderColumn_(
    headers,
    ['분야', '카테고리', '주제', 'KDC']
  );
  pagesColumn = HR_findAnyHeaderColumn_(
    headers,
    ['쪽수', '페이지', '전체 페이지']
  );
  linkColumn = HR_findAnyHeaderColumn_(
    headers,
    ['YES24 링크', '도서 링크', '링크']
  );
  coverColumn = HR_findAnyHeaderColumn_(
    headers,
    ['표지', '책 표지', '표지 URL', '표지URL', '이미지']
  );

  for (
    rowIndex = 0;
    rowIndex < values.length;
    rowIndex++
  ) {
    var backupRow = values[rowIndex];
    var backupFormulaRow = formulas[rowIndex];
    var backupStatus = String(
      backupRow[6] || ''
    ).trim();
    var sourceRow = backupRow.slice(
      8,
      8 + headers.length
    );
    var sourceFormulaRow = backupFormulaRow.slice(
      8,
      8 + headers.length
    );
    var title = String(
      backupRow[2] ||
      (
        titleColumn
          ? sourceRow[titleColumn - 1]
          : ''
      ) ||
      ''
    ).trim();

    if (
      !/^취소/.test(backupStatus) ||
      !title
    ) {
      continue;
    }

    var pendingAudienceColumn = HR_findAudienceColumn_(headers);
    var pendingAudience = pendingAudienceColumn
      ? HR_audienceKey_(sourceRow[pendingAudienceColumn - 1])
      : 'staff';
    if (pendingAudience !== HR_audienceKey_(audienceKey)) {
      continue;
    }

    items.push({
      row: 0,
      sequence: sequenceColumn
        ? String(
          sourceRow[sequenceColumn - 1] || '취소'
        ).trim()
        : '취소',
      title: title,
      author: authorColumn
        ? String(sourceRow[authorColumn - 1] || '').trim()
        : '',
      publisher: publisherColumn
        ? String(sourceRow[publisherColumn - 1] || '').trim()
        : '',
      year: yearColumn
        ? String(sourceRow[yearColumn - 1] || '').trim()
        : '',
      status: '🚫 신청 취소',
      requestedAt: dateColumn
        ? String(sourceRow[dateColumn - 1] || '').trim()
        : String(backupRow[1] || '').trim(),
      isbn: String(
        backupRow[3] ||
        (
          isbnColumn
            ? sourceRow[isbnColumn - 1]
            : ''
        ) ||
        ''
      ).trim(),
      price: priceColumn
        ? String(sourceRow[priceColumn - 1] || '').trim()
        : '',
      category: categoryColumn
        ? String(sourceRow[categoryColumn - 1] || '').trim()
        : '',
      pages: pagesColumn
        ? String(sourceRow[pagesColumn - 1] || '').trim()
        : '',
      link: linkColumn
        ? String(sourceRow[linkColumn - 1] || '').trim()
        : '',
      cover: coverColumn
        ? HR_extractRequestCoverUrl_(
          sourceFormulaRow[coverColumn - 1],
          sourceRow[coverColumn - 1]
        )
        : '',
      cancelId: String(backupRow[0] || '').trim(),
      cancelledBackup: true
    });
  }

  return items;
}

/**
 * 예전 방식으로 상태만 취소 처리된 행을 다시 접수 상태로 바꿉니다.
 */
function HR_restoreLegacyCancelledRequest(requestData) {
  var rowNumber = Number(
    requestData && requestData.row || 0
  );
  var requestedTitle = String(
    requestData && requestData.title || ''
  ).trim();
  var requestedIsbn = HR_normalizeIsbn_(
    requestData && requestData.isbn || ''
  );
  var spreadsheet = HR_getSpreadsheet_();
  var sheet = spreadsheet.getSheetByName(
    HR_TARGET_SHEET_NAME
  );
  var lock = LockService.getScriptLock();
  var lockAcquired = false;
  var headers;
  var titleColumn;
  var isbnColumn;
  var statusColumn;
  var currentTitle;
  var currentIsbn;
  var currentStatus;

  if (!sheet) {
    throw new Error(
      HR_TARGET_SHEET_NAME + ' 시트를 찾지 못했습니다.'
    );
  }

  try {
    lockAcquired = lock.tryLock(
      HR_REQUEST_LOCK_WAIT_MS
    );

    if (!lockAcquired) {
      throw new Error(
        '현재 이용자가 많습니다. 잠시 후 다시 시도해 주세요.'
      );
    }

    if (
      !rowNumber ||
      rowNumber < 2 ||
      rowNumber > sheet.getLastRow()
    ) {
      throw new Error(
        '신청 목록이 변경되었습니다. 새로고침한 뒤 다시 시도해 주세요.'
      );
    }

    headers = getHeaders(sheet);
    titleColumn = HR_findAnyHeaderColumn_(
      headers,
      ['도서명', '책 제목', '책제목', '서명', '제목']
    );
    isbnColumn = HR_findAnyHeaderColumn_(
      headers,
      ['ISBN13', 'ISBN 13', 'ISBN', '국제표준도서번호']
    );
    statusColumn = HR_findAnyHeaderColumn_(
      headers,
      ['처리상태', '처리 상태', '상태']
    );

    if (!titleColumn || !statusColumn) {
      throw new Error(
        '희망도서 목록에서 도서명 또는 처리상태 열을 찾지 못했습니다.'
      );
    }

    currentTitle = String(
      sheet
        .getRange(rowNumber, titleColumn)
        .getDisplayValue() || ''
    ).trim();
    currentIsbn = isbnColumn
      ? HR_normalizeIsbn_(
        sheet
          .getRange(rowNumber, isbnColumn)
          .getDisplayValue()
      )
      : '';
    currentStatus = String(
      sheet
        .getRange(rowNumber, statusColumn)
        .getDisplayValue() || ''
    ).trim();

    if (
      !currentTitle ||
      HR_normalizeBookTitle_(currentTitle) !==
        HR_normalizeBookTitle_(requestedTitle) ||
      (
        requestedIsbn &&
        currentIsbn !== requestedIsbn
      )
    ) {
      throw new Error(
        '신청 목록이 변경되었습니다. 새로고침한 뒤 다시 시도해 주세요.'
      );
    }

    if (!/취소/.test(currentStatus)) {
      return {
        success: true,
        alreadyRestored: true,
        title: currentTitle,
        isbn: currentIsbn,
        status: currentStatus
      };
    }

    sheet
      .getRange(rowNumber, statusColumn)
      .setValue('✅ 접수 완료');

    SpreadsheetApp.flush();

    return {
      success: true,
      restored: true,
      title: currentTitle,
      isbn: currentIsbn,
      row: rowNumber,
      status: '✅ 접수 완료'
    };
  } finally {
    if (lockAcquired) {
      lock.releaseLock();
    }
  }
}

/**
 * 표지 셀의 IMAGE 수식 또는 URL에서 실제 이미지 주소를 꺼냅니다.
 * 저장된 값을 읽기만 하므로 YES24 API를 추가 호출하지 않습니다.
 */
function HR_extractRequestCoverUrl_(
  formula,
  displayValue
) {
  var formulaText = String(formula || '');
  var displayText = String(displayValue || '').trim();
  var match = formulaText.match(
    /IMAGE\("([^"]+)"/i
  );

  if (match && match[1]) {
    return match[1].replace(/""/g, '"');
  }

  return /^https?:\/\//i.test(displayText)
    ? displayText
    : '';
}


function HR_getOrCreateCancelBackup_(
  spreadsheet,
  sourceSheet,
  headers
) {
  var sheet = spreadsheet.getSheetByName(
    HR_REQUEST_CANCEL_BACKUP_SHEET_NAME
  );
  var metadataHeaders = [
    '취소ID',
    '취소시각',
    '도서명',
    'ISBN13',
    '원래행',
    '원본열수',
    '복구상태',
    '복구시각'
  ];
  var sourceHeaders = Array.isArray(headers)
    ? headers
    : [];
  var requiredColumns =
    metadataHeaders.length + sourceHeaders.length;

  if (!sheet) {
    sheet = spreadsheet.insertSheet(
      HR_REQUEST_CANCEL_BACKUP_SHEET_NAME
    );
  }

  if (sheet.getMaxColumns() < requiredColumns) {
    sheet.insertColumnsAfter(
      sheet.getMaxColumns(),
      requiredColumns - sheet.getMaxColumns()
    );
  }

  sheet
    .getRange(1, 1, 1, requiredColumns)
    .setValues([metadataHeaders.concat(sourceHeaders)])
    .setFontWeight('bold')
    .setBackground('#efe8dc');

  sheet.setFrozenRows(1);

  if (!sheet.isSheetHidden()) {
    sheet.hideSheet();
  }

  return sheet;
}

/**
 * 공개 신청 화면에서 해당 도서 한 건만 안전하게 취소합니다.
 * 원본 행은 숨김 백업 시트에 보관한 뒤 목록에서 삭제합니다.
 */
function HR_cancelSubmittedRequest(requestData) {
  HR_requireStaffAccess_(HR_audienceKey_(requestData && requestData.audience), requestData && requestData.staffCode);
  var rowNumber = Number(
    requestData && requestData.row || 0
  );
  var requestedTitle = String(
    requestData && requestData.title || ''
  ).trim();
  var requestedIsbn = HR_normalizeIsbn_(
    requestData && requestData.isbn || ''
  );
  var spreadsheet = HR_getSpreadsheet_();
  var sheet = spreadsheet.getSheetByName(
    HR_TARGET_SHEET_NAME
  );
  var lock = LockService.getScriptLock();
  var lockAcquired = false;
  var headers;
  var titleColumn;
  var isbnColumn;
  var statusColumn;
  var currentTitle;
  var currentIsbn;
  var currentStatus;
  var originalColumnCount;
  var backupSheet;
  var backupRow;
  var cancelId;

  if (!sheet) {
    throw new Error(
      HR_TARGET_SHEET_NAME + ' 시트를 찾지 못했습니다.'
    );
  }

  try {
    lockAcquired = lock.tryLock(
      HR_REQUEST_LOCK_WAIT_MS
    );

    if (!lockAcquired) {
      throw new Error(
        '현재 이용자가 많습니다. 잠시 후 다시 시도해 주세요.'
      );
    }

    if (
      !rowNumber ||
      rowNumber < FIRST_DATA_ROW ||
      rowNumber > sheet.getLastRow()
    ) {
      throw new Error(
        '신청 목록이 변경되었습니다. 새로고침한 뒤 다시 시도해 주세요.'
      );
    }

    headers = getHeaders(sheet);
    titleColumn = HR_findAnyHeaderColumn_(
      headers,
      ['도서명', '책 제목', '책제목', '서명', '제목']
    );
    isbnColumn = HR_findAnyHeaderColumn_(
      headers,
      ['ISBN13', 'ISBN 13', 'ISBN', '국제표준도서번호']
    );
    statusColumn = HR_findAnyHeaderColumn_(
      headers,
      ['처리상태', '처리 상태', '상태']
    );

    if (!titleColumn || !statusColumn) {
      throw new Error(
        '희망도서 목록에서 도서명 또는 처리상태 열을 찾지 못했습니다.'
      );
    }

    currentTitle = String(
      sheet
        .getRange(rowNumber, titleColumn)
        .getDisplayValue() || ''
    ).trim();
    currentIsbn = isbnColumn
      ? HR_normalizeIsbn_(
        sheet
          .getRange(rowNumber, isbnColumn)
          .getDisplayValue()
      )
      : '';
    currentStatus = String(
      sheet
        .getRange(rowNumber, statusColumn)
        .getDisplayValue() || ''
    ).trim();

    if (
      !currentTitle ||
      HR_normalizeBookTitle_(currentTitle) !==
        HR_normalizeBookTitle_(requestedTitle) ||
      (
        requestedIsbn &&
        currentIsbn !== requestedIsbn
      )
    ) {
      throw new Error(
        '신청 목록이 변경되었습니다. 새로고침한 뒤 다시 시도해 주세요.'
      );
    }

    var cancelAudienceColumn = HR_findAudienceColumn_(headers);
    var cancelRowAudience = cancelAudienceColumn
      ? HR_audienceKey_(
        sheet.getRange(rowNumber, cancelAudienceColumn).getDisplayValue()
      )
      : 'staff';
    if (cancelRowAudience !== HR_audienceKey_(requestData && requestData.audience)) {
      throw new Error(
        '다른 구분(교직원/학부모)의 신청은 취소할 수 없습니다.'
      );
    }

    if (/구입|입수|선정|반려/.test(currentStatus)) {
      return {
        success: false,
        locked: true,
        title: currentTitle,
        status: currentStatus,
        message:
          '이미 처리 중이거나 완료된 도서라 취소할 수 없습니다.'
      };
    }

    originalColumnCount = sheet.getLastColumn();
    backupSheet = HR_getOrCreateCancelBackup_(
      spreadsheet,
      sheet,
      headers
    );
    backupRow = Math.max(
      backupSheet.getLastRow() + 1,
      2
    );
    cancelId = Utilities.getUuid();

    if (backupSheet.getMaxRows() < backupRow) {
      backupSheet.insertRowsAfter(
        backupSheet.getMaxRows(),
        backupRow - backupSheet.getMaxRows()
      );
    }

    sheet
      .getRange(
        rowNumber,
        1,
        1,
        originalColumnCount
      )
      .copyTo(
        backupSheet.getRange(
          backupRow,
          9,
          1,
          originalColumnCount
        ),
        SpreadsheetApp.CopyPasteType.PASTE_NORMAL,
        false
      );

    backupSheet
      .getRange(backupRow, 1, 1, 8)
      .setValues([[
        cancelId,
        new Date(),
        currentTitle,
        currentIsbn,
        rowNumber,
        originalColumnCount,
        '취소 백업 완료',
        ''
      ]]);

    try {
      sheet.deleteRow(rowNumber);
    } catch (deleteError) {
      backupSheet
        .getRange(backupRow, 7)
        .setValue('삭제 실패 · 원본 유지');
      throw deleteError;
    }

    backupSheet
      .getRange(backupRow, 7)
      .setValue('취소 완료');

    SpreadsheetApp.flush();

    return {
      success: true,
      deleted: true,
      cancelId: cancelId,
      title: currentTitle,
      isbn: currentIsbn,
      status: '신청 취소'
    };
  } finally {
    if (lockAcquired) {
      lock.releaseLock();
    }
  }
}

/**
 * 방금 취소한 신청 한 건을 숨김 백업에서 원래 목록으로 복구합니다.
 */
function HR_undoCancelledRequest(requestData) {
  HR_requireStaffAccess_(HR_audienceKey_(requestData && requestData.audience), requestData && requestData.staffCode);
  var cancelId = String(
    requestData && requestData.cancelId || ''
  ).trim();
  var spreadsheet = HR_getSpreadsheet_();
  var sheet = spreadsheet.getSheetByName(
    HR_TARGET_SHEET_NAME
  );
  var backupSheet = spreadsheet.getSheetByName(
    HR_REQUEST_CANCEL_BACKUP_SHEET_NAME
  );
  var lock = LockService.getScriptLock();
  var lockAcquired = false;
  var cancelCell;
  var backupRow;
  var originalRow;
  var originalColumnCount;
  var backupStatus;
  var title;
  var isbn;
  var headers;
  var titleColumn;
  var isbnColumn;
  var currentRows;
  var duplicateRow = 0;
  var rowIndex;
  var targetRow;
  var lastRow;

  if (!cancelId) {
    throw new Error('되돌릴 신청 정보가 없습니다.');
  }

  if (!sheet || !backupSheet) {
    throw new Error(
      '신청 취소 백업을 찾지 못했습니다.'
    );
  }

  try {
    lockAcquired = lock.tryLock(
      HR_REQUEST_LOCK_WAIT_MS
    );

    if (!lockAcquired) {
      throw new Error(
        '현재 이용자가 많습니다. 잠시 후 다시 시도해 주세요.'
      );
    }

    if (backupSheet.getLastRow() < 2) {
      throw new Error(
        '되돌릴 신청 정보를 찾지 못했습니다.'
      );
    }

    cancelCell = backupSheet
      .getRange(
        2,
        1,
        backupSheet.getLastRow() - 1,
        1
      )
      .createTextFinder(cancelId)
      .matchEntireCell(true)
      .findNext();

    if (!cancelCell) {
      throw new Error(
        '되돌릴 신청 정보를 찾지 못했습니다.'
      );
    }

    backupRow = cancelCell.getRow();
    originalRow = Number(
      backupSheet.getRange(backupRow, 5).getValue()
    );
    originalColumnCount = Number(
      backupSheet.getRange(backupRow, 6).getValue()
    );
    backupStatus = String(
      backupSheet.getRange(backupRow, 7).getValue() || ''
    );
    title = String(
      backupSheet.getRange(backupRow, 3).getValue() || ''
    ).trim();
    isbn = HR_normalizeIsbn_(
      backupSheet.getRange(backupRow, 4).getValue()
    );

    if (/^복구 완료|^복구 생략/.test(backupStatus)) {
      return {
        success: true,
        alreadyRestored: true,
        title: title,
        isbn: isbn,
        status: backupStatus
      };
    }

    if (!originalColumnCount) {
      throw new Error(
        '백업된 신청 정보가 올바르지 않습니다.'
      );
    }

    headers = getHeaders(sheet);
    var undoAudienceColumn = HR_findAudienceColumn_(headers);
    var undoRowAudience = (undoAudienceColumn && originalColumnCount >= undoAudienceColumn)
      ? HR_audienceKey_(
        backupSheet.getRange(backupRow, 8 + undoAudienceColumn).getDisplayValue()
      )
      : 'staff';
    if (undoRowAudience !== HR_audienceKey_(requestData && requestData.audience)) {
      throw new Error(
        '다른 구분(교직원/학부모)의 신청은 되돌릴 수 없습니다.'
      );
    }
    titleColumn = HR_findAnyHeaderColumn_(
      headers,
      ['도서명', '책 제목', '책제목', '서명', '제목']
    );
    isbnColumn = HR_findAnyHeaderColumn_(
      headers,
      ['ISBN13', 'ISBN 13', 'ISBN', '국제표준도서번호']
    );

    if (!titleColumn) {
      throw new Error(
        '희망도서 목록에서 도서명 열을 찾지 못했습니다.'
      );
    }

    lastRow = sheet.getLastRow();

    if (lastRow >= FIRST_DATA_ROW) {
      currentRows = sheet
        .getRange(
          FIRST_DATA_ROW,
          1,
          lastRow - FIRST_DATA_ROW + 1,
          sheet.getLastColumn()
        )
        .getDisplayValues();

      for (
        rowIndex = 0;
        rowIndex < currentRows.length;
        rowIndex++
      ) {
        var existingTitle = String(
          currentRows[rowIndex][titleColumn - 1] || ''
        ).trim();
        var existingIsbn = isbnColumn
          ? HR_normalizeIsbn_(
            currentRows[rowIndex][isbnColumn - 1]
          )
          : '';

        if (
          (
            isbn &&
            existingIsbn === isbn
          ) ||
          (
            !isbn &&
            HR_normalizeBookTitle_(existingTitle) ===
              HR_normalizeBookTitle_(title)
          )
        ) {
          duplicateRow = rowIndex + FIRST_DATA_ROW;
          break;
        }
      }
    }

    if (duplicateRow) {
      backupSheet
        .getRange(backupRow, 7, 1, 2)
        .setValues([[
          '복구 생략 · 이미 목록에 있음',
          new Date()
        ]]);

      return {
        success: true,
        alreadyRestored: true,
        title: title,
        isbn: isbn,
        row: duplicateRow,
        status: '이미 목록에 있음'
      };
    }

    if (sheet.getMaxColumns() < originalColumnCount) {
      sheet.insertColumnsAfter(
        sheet.getMaxColumns(),
        originalColumnCount - sheet.getMaxColumns()
      );
    }

    lastRow = sheet.getLastRow();
    targetRow = Math.max(
      FIRST_DATA_ROW,
      Math.min(
        originalRow || lastRow + 1,
        lastRow + 1
      )
    );

    if (targetRow <= lastRow) {
      sheet.insertRowBefore(targetRow);
    } else {
      sheet.insertRowAfter(Math.max(lastRow, HEADER_ROW));
      targetRow = Math.max(lastRow + 1, FIRST_DATA_ROW);
    }

    backupSheet
      .getRange(
        backupRow,
        9,
        1,
        originalColumnCount
      )
      .copyTo(
        sheet.getRange(
          targetRow,
          1,
          1,
          originalColumnCount
        ),
        SpreadsheetApp.CopyPasteType.PASTE_NORMAL,
        false
      );

    backupSheet
      .getRange(backupRow, 7, 1, 2)
      .setValues([[
        '복구 완료',
        new Date()
      ]]);

    SpreadsheetApp.flush();

    return {
      success: true,
      restored: true,
      title: title,
      isbn: isbn,
      row: targetRow,
      status: '신청 접수'
    };
  } finally {
    if (lockAcquired) {
      lock.releaseLock();
    }
  }
}

var HR_BESTSELLER_CACHE_VERSION = 'yes24_v1';
var HR_BESTSELLER_CACHE_MS = 24 * 60 * 60 * 1000;
var HR_BESTSELLER_MAX_RESULTS = 30;
var HR_BESTSELLER_CATEGORIES = {
  all: {
    label: '종합',
    categoryId: ''
  },
  youth: {
    label: '청소년',
    categoryId: ''
  },
  children: {
    label: '어린이',
    categoryId: ''
  },
  picture: {
    label: '그림책',
    categoryId: ''
  }
};

/**
 * 공개 신청 화면에서 세 분류의 베스트셀러를 한 번에 가져옵니다.
 * 같은 날에는 저장해 둔 결과를 재사용하므로 화면을 여러 명이 열어도
 * YES24 API를 반복 호출하지 않습니다.
 */
function HR_getBestsellers() {
  var yes24Key = HR_getYes24Key_();
  var properties = PropertiesService.getScriptProperties();
  var lock = LockService.getScriptLock();
  var categoryKeys = ['all', 'youth', 'children', 'picture'];
  var result = {
    source: 'YES24 인터넷서점',
    cacheHours: 24,
    categories: {}
  };
  var lockAcquired = false;

  if (!yes24Key) {
    throw new Error('YES24 API 키가 설정되지 않았습니다.');
  }

  try {
    lockAcquired = lock.tryLock(15000);

    categoryKeys.forEach(function(categoryKey) {
      result.categories[categoryKey] =
        HR_getBestsellerCategory_(
          categoryKey,
          yes24Key,
          properties,
          lockAcquired
        );
    });
  } finally {
    if (lockAcquired) {
      lock.releaseLock();
    }
  }

  return result;
}

/**
 * 분류 한 개의 저장 결과를 확인하고, 24시간이 지났을 때만 새로 조회합니다.
 */
function HR_getBestsellerCategory_(
  categoryKey,
  yes24Key,
  properties,
  allowRefresh
) {
  var category = HR_BESTSELLER_CATEGORIES[categoryKey];
  var cacheKey =
    'HR_BESTSELLER_' +
    HR_BESTSELLER_CACHE_VERSION +
    '_' +
    categoryKey;
  var timeKey = cacheKey + '_AT';
  var cachedAt = Number(
    properties.getProperty(timeKey) || 0
  );
  var cachedBooks = HR_readBestsellerCache_(
    properties,
    cacheKey
  );
  var isFresh =
    cachedBooks.length > 0 &&
    cachedAt > 0 &&
    Date.now() - cachedAt < HR_BESTSELLER_CACHE_MS;

  if (!category) {
    return {
      label: categoryKey,
      books: [],
      updatedAt: '',
      stale: false
    };
  }

  if (isFresh || !allowRefresh) {
    return {
      label: category.label,
      books: cachedBooks,
      updatedAt: cachedAt
        ? new Date(cachedAt).toISOString()
        : '',
      stale: !isFresh && cachedBooks.length > 0
    };
  }

  try {
    cachedBooks = HR_fetchBestsellerCategory_(
      yes24Key,
      category
    );
    cachedAt = Date.now();

    HR_writeBestsellerCache_(
      properties,
      cacheKey,
      cachedBooks
    );
    properties.setProperty(
      timeKey,
      String(cachedAt)
    );

    return {
      label: category.label,
      books: cachedBooks,
      updatedAt: new Date(cachedAt).toISOString(),
      stale: false
    };
  } catch (error) {
    if (cachedBooks.length > 0) {
      return {
        label: category.label,
        books: cachedBooks,
        updatedAt: cachedAt
          ? new Date(cachedAt).toISOString()
          : '',
        stale: true
      };
    }

    throw error;
  }
}

/**
 * YES24 베스트셀러 API에서 한 분류의 베스트셀러를 조회합니다.
 */
function HR_fetchBestsellerCategory_(yes24Key, category) {
  var categoryId = yes24CategoryId_(category.label, yes24Key);
  var data = yes24Get_('/category/bestseller', {categoryId: categoryId, page: 1, pageSize: HR_BESTSELLER_MAX_RESULTS, detail: 'N'}, yes24Key);
  return data.items.slice(0, HR_BESTSELLER_MAX_RESULTS).map(function(item,index) {
    var book = yes24RequestBook_(item); book.rank = index + 1; return book;
  });
}

/**
 * 베스트셀러 30권을 ScriptProperties 용량 제한에 맞춰 나누어 읽습니다.
 */
function HR_readBestsellerCache_(properties, cacheKey) {
  var count = Number(
    properties.getProperty(cacheKey + '_COUNT') || 0
  );
  var text = '';
  var index;

  if (count > 0) {
    for (index = 0; index < count; index++) {
      text += String(
        properties.getProperty(
          cacheKey + '_PART_' + index
        ) || ''
      );
    }
  } else {
    text = String(
      properties.getProperty(cacheKey) || ''
    );
  }

  return HR_parseBestsellerCache_(text);
}

/**
 * 베스트셀러 결과를 안전한 크기의 여러 속성으로 나누어 저장합니다.
 */
function HR_writeBestsellerCache_(
  properties,
  cacheKey,
  books
) {
  var text = JSON.stringify(books || []);
  var chunkSize = 7000;
  var oldCount = Number(
    properties.getProperty(cacheKey + '_COUNT') || 0
  );
  var count = Math.max(
    1,
    Math.ceil(text.length / chunkSize)
  );
  var values = {};
  var index;

  for (index = 0; index < count; index++) {
    values[cacheKey + '_PART_' + index] =
      text.substring(
        index * chunkSize,
        (index + 1) * chunkSize
      );
  }

  values[cacheKey + '_COUNT'] = String(count);
  properties.setProperties(values, false);
  properties.deleteProperty(cacheKey);

  for (index = count; index < oldCount; index++) {
    properties.deleteProperty(
      cacheKey + '_PART_' + index
    );
  }
}

/**
 * 저장된 베스트셀러 JSON을 안전하게 읽습니다.
 */
function HR_parseBestsellerCache_(text) {
  var value;

  if (!text) {
    return [];
  }

  try {
    value = JSON.parse(text);
  } catch (error) {
    return [];
  }

  return Array.isArray(value)
    ? value
    : [];
}


/**
 * 이미 접수된 도서 중 비어 있는 정가·출판연도만 YES24 정보로 보완합니다.
 * 한 번에 최대 20권까지만 조회해 API 사용량을 제한합니다.
 */
function HR_fillMissingPriceAndYear() {
  var spreadsheet = HR_getSpreadsheet_();
  var sheet = spreadsheet.getSheetByName(HR_TARGET_SHEET_NAME);
  var headers;
  var columns;
  var priceColumn;
  var yearColumn;
  var lastRow;
  var row;
  var checkedCount = 0;
  var updatedCount = 0;
  var failedCount = 0;

  if (!sheet) {
    throw new Error(HR_TARGET_SHEET_NAME + ' 시트를 찾지 못했습니다.');
  }

  headers = getHeaders(sheet);
  columns = getColumnMap(headers);
  priceColumn = HR_findAnyHeaderColumn_(
    headers,
    ['정가', '도서정가', '정가(원)', '가격']
  );
  yearColumn = HR_findAnyHeaderColumn_(
    headers,
    ['연도', '출판연도', '발행연도', '출간연도']
  );
  lastRow = sheet.getLastRow();

  if (!priceColumn || !yearColumn) {
    throw new Error('정가 또는 연도 열을 찾지 못했습니다.');
  }

  for (row = 2; row <= lastRow; row++) {
    var priceValue = Number(
      String(sheet.getRange(row, priceColumn).getDisplayValue() || '')
        .replace(/[^0-9.-]/g, '')
    );
    var yearValue = String(
      sheet.getRange(row, yearColumn).getDisplayValue() || ''
    ).trim();
    var isbnValue = columns.isbn
      ? String(sheet.getRange(row, columns.isbn).getDisplayValue() || '').trim()
      : '';
    var titleValue = columns.title
      ? String(sheet.getRange(row, columns.title).getDisplayValue() || '').trim()
      : '';
    var foundBooks;
    var foundBook;
    var rowUpdated = false;

    if ((priceValue > 0 && yearValue) || (!isbnValue && !titleValue)) {
      continue;
    }

    if (checkedCount >= 20) {
      break;
    }

    checkedCount++;

    try {
      foundBooks = HR_searchRequestBooks(isbnValue || titleValue);
      foundBook = foundBooks && foundBooks.length
        ? foundBooks[0]
        : null;

      if (!foundBook) {
        failedCount++;
        continue;
      }

      if (!(priceValue > 0) && Number(foundBook.price || 0) > 0) {
        sheet
          .getRange(row, priceColumn)
          .setValue(Number(foundBook.price))
          .setNumberFormat('#,##0');
        rowUpdated = true;
      }

      if (!yearValue && foundBook.year) {
        sheet
          .getRange(row, yearColumn)
          .setValue(String(foundBook.year))
          .setNumberFormat('0');
        rowUpdated = true;
      }

      if (rowUpdated) {
        updatedCount++;
      }
    } catch (error) {
      failedCount++;
      console.error('정가·연도 보완 오류 (' + row + '행)', error);
    }

    Utilities.sleep(120);
  }

  SpreadsheetApp.flush();

  return {
    checkedCount: checkedCount,
    updatedCount: updatedCount,
    failedCount: failedCount
  };
}


function HR_breakMergedRangesSafely_(sheet, rangeA1) {
  var targetRange = rangeA1
    ? sheet.getRange(rangeA1)
    : sheet.getRange(
        1,
        1,
        Math.max(sheet.getMaxRows(), 1),
        Math.max(sheet.getMaxColumns(), 1)
      );

  targetRange.getMergedRanges().forEach(function(mergedRange) {
    mergedRange.breakApart();
  });
}



/**
 * 웹페이지에서 해석한 엑셀·CSV 도서 목록을 업무용 시트에 저장합니다.
 * 원본 파일 전체는 전송하지 않고 필요한 서지정보만 받습니다.
 */
function HR_importBookListFile(payload) {
  payload = payload || {};

  var books = Array.isArray(payload.books) ? payload.books : [];
  var fileName = String(payload.fileName || '업로드 파일').trim();

  if (!books.length) {
    throw new Error('파일에서 읽은 도서 정보가 없습니다.');
  }
  if (books.length > 2000) {
    throw new Error('한 번에 최대 1,500권까지 업로드할 수 있습니다.');
  }

  return unifiedReplaceUploadRows_(books, fileName);
}


/**
 * ============================================================
 * 공개 신청 입력 정리·검증
 * - 브라우저에서 온 값은 신뢰하지 않습니다.
 * - 문자열은 정리해 길이를 제한하고, 시트에는 텍스트로만 기록합니다.
 * - 도서 정보는 ISBN으로 YES24에서 다시 확인한 값을 우선합니다.
 * ============================================================
 */
function HR_cleanInput_(value, maxLength) {
  var source = String(value === null || value === undefined ? '' : value);
  var text = '';
  var index = 0;
  var code = 0;

  // 제어문자·보이지 않는 문자는 공백으로 바꿉니다.
  for (index = 0; index < source.length; index++) {
    code = source.charCodeAt(index);
    if (
      code < 32 || code === 127 ||
      (code >= 8203 && code <= 8207) ||
      code === 8232 || code === 8233 || code === 65279
    ) {
      text += ' ';
    } else {
      text += source.charAt(index);
    }
  }

  text = text.replace(/\s+/g, ' ').trim();

  if (maxLength && text.length > maxLength) {
    text = text.slice(0, maxLength);
  }
  return text;
}

/**
 * 텍스트 서식을 먼저 지정한 뒤 값을 넣어, 입력이 수식 문자로 시작해도 실행되지 않게 합니다.
 */
function HR_setSafeText_(cell, value) {
  var text = HR_cleanInput_(value, 500);

  cell.setNumberFormat('@');
  if (!text) {
    cell.setValue('');
    return cell;
  }
  // setValue 는 '=' 로 시작하는 문자열을 수식으로 실행하므로, '=' 로 시작하면
  // 항상 텍스트로 저장되는 서식 있는 텍스트를 사용합니다. 그 밖의 문자열은 텍스트 서식만으로 안전합니다.
  if (text.charAt(0) === '=') {
    cell.setRichTextValue(
      SpreadsheetApp.newRichTextValue().setText(text).build()
    );
  } else {
    cell.setValue(text);
  }
  return cell;
}

function HR_cleanPrice_(value) {
  var number = Number(String(value === null || value === undefined ? '' : value).replace(/[^0-9.]/g, ''));
  return isFinite(number) && number > 0 && number <= 500000
    ? Math.round(number)
    : 0;
}

function HR_cleanCallNumber_(value) {
  var text = HR_cleanInput_(value, 20);
  return /^[0-9]{1,3}(?:\.[0-9]{1,8})?$/.test(text) ? text : '';
}

/**
 * 표지는 YES24 주소만 허용합니다. 수식 안에 들어가므로 따옴표·공백·괄호가 있으면 버립니다.
 */
function HR_safeCoverUrl_(value) {
  var text = String(value || '').trim();

  if (text.length > 300) return '';
  if (/["\\\s()]/.test(text)) return '';
  return /^https?:\/\/(?:[a-z0-9-]+\.)*yes24\.com\//i.test(text)
    ? text
    : '';
}

function HR_prepareSelectedBook_(clientBook, query) {
  var source = clientBook && typeof clientBook === 'object' ? clientBook : {};
  var year = (
    HR_cleanInput_(source.year || source.pubDate || source.publicationDate, 20)
      .match(/\d{4}/) || ['']
  )[0];
  var book = {
    title: HR_cleanInput_(source.title, 200),
    author: HR_cleanInput_(source.author, 150),
    publisher: HR_cleanInput_(source.publisher, 100),
    isbn: yes24Isbn13_(source.isbn) || yes24Isbn13_(query) || '',
    cover: '',
    link: '',
    price: HR_cleanPrice_(
      source.price || source.priceStandard || source.priceStd || source.priceSales
    ),
    year: year && Number(year) >= 1800 && Number(year) <= 2100 ? year : '',
    category: HR_cleanInput_(source.category, 100),
    callNumber: HR_cleanCallNumber_(source.callNumber),
    verified: false
  };
  var items;
  var server;

  if (book.isbn) {
    try {
      items = yes24SearchItems_(book.isbn, HR_getYes24Key_(), 'BOOK', 1, true);

      if (items && items.length) {
        server = yes24RequestBook_(items[0]);
        book.title = HR_cleanInput_(server.title, 200) || book.title;
        book.author = HR_cleanInput_(server.author, 150) || book.author;
        book.publisher = HR_cleanInput_(server.publisher, 100) || book.publisher;
        book.isbn = yes24Isbn13_(server.isbn) || book.isbn;
        book.price = HR_cleanPrice_(server.price) || book.price;
        book.year = server.year || book.year;
        book.category = HR_cleanInput_(server.category, 100) || book.category;
        book.cover = HR_safeCoverUrl_(server.cover);
        book.link = HR_safeUrl_(server.link);
        book.verified = true;
      }
    } catch (error) {
      // YES24 연결 문제로 확인하지 못한 경우에는 정리한 입력값으로 접수하고 '미확인'으로 표시합니다.
      console.error('신청 도서 서버 확인 실패', error);
    }
  }

  return book;
}

function HR_unverifiedNote_(book) {
  return book && book.verified ? '' : ' · 도서정보 미확인';
}
