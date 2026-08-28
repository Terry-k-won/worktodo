/** @OnlyCurrentDoc */

// ============================================================================
// Life & Task Management Dashboard - Google Apps Script Backend (v3.0)
// ============================================================================

var SHEET_NAME = 'TaskDB';

function doGet(e) {
  try {
    var action = (e && e.parameter) ? e.parameter.action : '';
    
    if (action === 'ping') {
      return createJsonResponse({ status: 'ok', message: 'Google Apps Script is online' });
    }

    // Direct GET Sync (Handles CORS-free reading & writing from any mobile/web browser)
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
      } catch (err1) {}
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
      dueDate: cleanDateString(row[4]),
      note: String(row[5] || ''),
      status: String(row[6] || 'active'),
      createdAt: cleanDateString(row[7]),
      updatedAt: cleanDateString(row[8])
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
      cleanDateString(item.dueDate),
      item.note || '',
      item.status || 'active',
      cleanDateString(item.createdAt),
      cleanDateString(item.updatedAt)
    ]);
  }

  sheet.getRange(2, 1, rows.length, 9).setValues(rows);
}

function cleanDateString(val) {
  if (!val) return '';
  if (val instanceof Date) {
    var year = val.getFullYear();
    var month = String(val.getMonth() + 1);
    if (month.length < 2) month = '0' + month;
    var day = String(val.getDate());
    if (day.length < 2) day = '0' + day;
    return year + '-' + month + '-' + day;
  }
  
  var str = String(val).trim();
  // Strip GMT timezone suffixes if present
  if (str.indexOf('GMT') !== -1) {
    try {
      var d = new Date(str);
      if (!isNaN(d.getTime())) {
        var y = d.getFullYear();
        var m = String(d.getMonth() + 1);
        if (m.length < 2) m = '0' + m;
        var date = String(d.getDate());
        if (date.length < 2) date = '0' + date;
        return y + '-' + m + '-' + date;
      }
    } catch (e) {}
  }
  return str;
}

function createJsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
