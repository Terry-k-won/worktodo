/** @OnlyCurrentDoc */

// ============================================================================
// Life & Task Management Dashboard - Google Apps Script Backend
// 
// 🚨 [보안 차단 및 구문 오류 해결 가이드]
// 1. 반드시 구글 스프레드시트 내의 [확장 프로그램] -> [Apps Script]에서 작성하세요.
// 2. 기존 코드를 모두 지우고 이 코드 전체를 복사해서 붙여넣으세요.
// 3. 우측 상단 [배포] -> [새 배포] -> 웹 앱 선택
//    - 실행: 나 (Me)
//    - 액세스: 모든 사용자 (Anyone)
// 4. 권한 승인 시 경고 화면이 나오면: [고급] -> [안전하지 않음으로 이동] -> [허용]
// ============================================================================

var SHEET_NAME = 'TaskDB';

function doGet(e) {
  try {
    var action = (e && e.parameter) ? e.parameter.action : '';
    
    if (action === 'ping') {
      return createJsonResponse({ status: 'ok', message: 'Google Apps Script is online' });
    }

    // Direct GET Sync (Guarantees CORS-free writing to Google Sheets from mobile/web)
    if (action === 'sync' || action === 'save') {
      if (e.parameter && e.parameter.data) {
        var itemsData = JSON.parse(e.parameter.data);
        if (itemsData && Array.isArray(itemsData)) {
          saveAllItemsToSheet(itemsData);
        }
      }
      var currentItems = getAllItemsFromSheet();
      return createJsonResponse({ status: 'success', items: currentItems });
    }

    var items = getAllItemsFromSheet();
    return createJsonResponse({ status: 'success', items: items });
  } catch (err) {
    return createJsonResponse({ status: 'error', message: err.toString() });
  }
}

function doPost(e) {
  try {
    var itemsData = null;

    if (e && e.postData && e.postData.contents) {
      try {
        var parsed = JSON.parse(e.postData.contents);
        if (parsed.items) itemsData = parsed.items;
      } catch (err1) {
        // Fallback for form data
      }
    }

    if (!itemsData && e && e.parameter && e.parameter.data) {
      try {
        itemsData = JSON.parse(e.parameter.data);
      } catch (err2) {}
    }

    if (itemsData && Array.isArray(itemsData)) {
      saveAllItemsToSheet(itemsData);
    }

    var updatedItems = getAllItemsFromSheet();
    return createJsonResponse({ status: 'success', items: updatedItems });
  } catch (err) {
    return createJsonResponse({ status: 'error', message: err.toString() });
  }
}

function getOrCreateSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    // Add Headers
    var headers = ['ID', 'Category', 'Content', 'Priority', 'DueDate', 'Note', 'Status', 'CreatedAt', 'UpdatedAt'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    // Format Header Row
    var headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground('#1e293b');
    headerRange.setFontColor('#ffffff');
    headerRange.setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getAllItemsFromSheet() {
  var sheet = getOrCreateSheet();
  var lastRow = sheet.getLastRow();
  
  if (lastRow <= 1) return []; // Only header row

  var data = sheet.getRange(2, 1, lastRow - 1, 9).getValues();
  var items = [];

  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    if (!row[0]) continue; // Skip empty ID

    items.push({
      id: String(row[0]),
      category: String(row[1] || 'today'),
      content: String(row[2] || ''),
      priority: String(row[3] || 'medium'),
      dueDate: row[4] ? formatDateValue(row[4]) : '',
      note: String(row[5] || ''),
      status: String(row[6] || 'active'),
      createdAt: row[7] ? String(row[7]) : new Date().toISOString(),
      updatedAt: row[8] ? String(row[8]) : new Date().toISOString()
    });
  }

  return items;
}

function saveAllItemsToSheet(items) {
  var sheet = getOrCreateSheet();
  
  // Clear old data rows (keep header)
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, 9).clearContent();
  }

  if (!items || items.length === 0) return;

  var rows = [];
  for (var i = 0; i < items.length; i++) {
    var item = items[i];
    rows.push([
      item.id || '',
      item.category || 'today',
      item.content || '',
      item.priority || 'medium',
      item.dueDate || '',
      item.note || '',
      item.status || 'active',
      item.createdAt || new Date().toISOString(),
      item.updatedAt || new Date().toISOString()
    ]);
  }

  sheet.getRange(2, 1, rows.length, 9).setValues(rows);
}

function formatDateValue(val) {
  if (val instanceof Date) {
    var year = val.getFullYear();
    var month = String(val.getMonth() + 1);
    if (month.length < 2) month = '0' + month;
    var day = String(val.getDate());
    if (day.length < 2) day = '0' + day;
    return year + '-' + month + '-' + day;
  }
  return String(val);
}

function createJsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
