const SALES_TAB = 'Sales Record';
const LOG_TAB = 'ScanLog';
const C = {
  ORDER_ID: 4, STATUS: 5, SERVED_BY: 6, SOURCE: 7, FILENAME: 8, BATCH: 9,
  CLIENT_NAME: 10, CLIENT_CONTACT: 11, FRAME: 14, SIZE: 15, BOARD: 16,
  PAY_METHOD: 17, DEPOSIT: 18, TOTAL: 22, REDO: 23
};
const SALES_COLUMN_COUNT = C.REDO + 1;

function doGet(e) {
  const action = e.parameter.action || '';
  let result;
  try {
    if (action === 'lookup') result = fastLookup(e.parameter.orderId);
    else if (action === 'history') result = getHistory(e.parameter.date);
    else if (action === 'ping') result = { ok: true };
    else result = { error: 'Invalid' };
  } catch (error) {
    result = { error: error.message };
  }
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  let result;
  try {
    const body = JSON.parse(e.postData.contents);
    if (body.action === 'logScan') result = logScan(body.orders);
    else if (body.action === 'updateNotify') result = updateNotify(body.rowIndex, body.status);
    else if (body.action === 'deleteOrder') result = deleteOrder(body.rowIndex);
    else result = { error: 'Invalid' };
  } catch (error) {
    result = { error: error.message };
  }
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function fastLookup(orderId) {
  if (!orderId) return { error: 'No orderId' };
  const id = String(orderId).trim();
  const cache = CacheService.getScriptCache();
  const cacheKey = 'order_v2_' + id;
  const cached = cache.get(cacheKey);
  if (cached) return JSON.parse(cached);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SALES_TAB);
  if (!sheet) return { found: false, orderId: id, error: 'Tab "' + SALES_TAB + '" not found' };

  const matches = sheet.getRange('E:E').createTextFinder(id).matchEntireCell(true).findAll();
  if (!matches || !matches.length) return { found: false, orderId: id };

  const rowNum = matches[matches.length - 1].getRow();
  const row = sheet.getRange(rowNum, 1, 1, SALES_COLUMN_COUNT).getValues()[0];
  const total = parseRM(row[C.TOTAL]);
  const deposit = parseRM(row[C.DEPOSIT]);
  const paymentMethod = String(row[C.PAY_METHOD] || '').trim();
  const isRedo = isRedoRow(row);

  let isPaid = /fully\s*paid/i.test(paymentMethod);
  let balance = isPaid ? 0 : total - deposit;
  let redo = {
    groupKey: '',
    orderCount: 0,
    originalTotal: 0,
    balance: balance,
    needsManualReview: false,
    reason: ''
  };

  if (isRedo) {
    redo = calculateRedoBalance(sheet, rowNum, row);
    balance = redo.balance;
    isPaid = !redo.needsManualReview && balance <= 0;
  }

  const result = {
    found: true,
    orderId: id,
    clientName: row[C.CLIENT_NAME] || '',
    clientContact: String(row[C.CLIENT_CONTACT] || ''),
    filename: row[C.FILENAME] || '',
    size: row[C.SIZE] || '',
    servedBy: row[C.SERVED_BY] || '',
    source: row[C.SOURCE] || '',
    batch: row[C.BATCH] || '',
    frame: row[C.FRAME] || '',
    board: row[C.BOARD] || '',
    status: row[C.STATUS] || '',
    paymentMethod: paymentMethod,
    deposit: deposit,
    total: total,
    balance: balance,
    isPaid: isPaid,
    isRedo: isRedo,
    redoGroupKey: redo.groupKey,
    redoOrderCount: redo.orderCount,
    redoOriginalTotal: redo.originalTotal,
    needsManualReview: redo.needsManualReview,
    redoReason: redo.reason
  };
  cache.put(cacheKey, JSON.stringify(result), 600);
  return result;
}

function calculateRedoBalance(sheet, redoRowNumber, redoRow) {
  const phone = normalizePhone(redoRow[C.CLIENT_CONTACT]);
  const deposit = parseRM(redoRow[C.DEPOSIT]);

  if (!phone || deposit <= 0) {
    return redoManualReview('Redo is missing a phone number or deposit amount.');
  }

  const rowCountAbove = redoRowNumber - 2;
  const rowsAbove = rowCountAbove > 0
    ? sheet.getRange(2, 1, rowCountAbove, SALES_COLUMN_COUNT).getValues()
    : [];

  const originals = rowsAbove.filter(function(candidate) {
    return !isRedoRow(candidate) &&
      parseRM(candidate[C.TOTAL]) > 0 &&
      normalizePhone(candidate[C.CLIENT_CONTACT]) === phone &&
      sameMoney(parseRM(candidate[C.DEPOSIT]), deposit);
  });

  if (!originals.length) {
    return redoManualReview('No matching original order was found above this redo.');
  }

  const originalTotal = originals.reduce(function(sum, candidate) {
    return sum + parseRM(candidate[C.TOTAL]);
  }, 0);

  return {
    groupKey: phone + '|' + deposit.toFixed(2),
    orderCount: originals.length,
    originalTotal: originalTotal,
    balance: Math.max(0, originalTotal - deposit),
    needsManualReview: false,
    reason: ''
  };
}

function redoManualReview(reason) {
  return {
    groupKey: '',
    orderCount: 0,
    originalTotal: 0,
    balance: 0,
    needsManualReview: true,
    reason: reason
  };
}

function isRedoRow(row) {
  const value = row[C.REDO];
  return value === true || String(value || '').trim().toUpperCase() === 'TRUE';
}

function normalizePhone(value) {
  return String(value || '').replace(/\D/g, '');
}

function sameMoney(a, b) {
  return Math.abs(a - b) < 0.005;
}

function parseRM(value) {
  if (value == null || value === '') return 0;
  if (typeof value === 'number') return value;
  return parseFloat(String(value).replace(/[^0-9.\-]/g, '')) || 0;
}

// Cols: A:ScanTime B:OrderID C:ClientName D:Phone E:Size F:Salesperson G:Source H:Payment I:Balance J:NotifyStatus K:NotifyTime
function logScan(orders) {
  if (!orders || !orders.length) return { error: 'No orders' };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let logSheet = ss.getSheetByName(LOG_TAB);
  if (!logSheet) {
    logSheet = ss.insertSheet(LOG_TAB);
    logSheet.appendRow(['Scan Time', 'Order ID', 'Client Name', 'Phone', 'Size', 'Salesperson', 'Source', 'Payment', 'Balance', 'Notify Status', 'Notify Time']);
    logSheet.getRange(1, 1, 1, 11).setFontWeight('bold').setBackground('#1D9E75').setFontColor('#fff');
    logSheet.setFrozenRows(1);
  }

  const now = new Date();
  const rows = orders.map(function(order) {
    return [
      now,
      order.orderId || '',
      order.clientName || '',
      order.clientContact || '',
      order.size || '',
      order.servedBy || '',
      order.source || '',
      order.isPaid ? 'Paid' : (order.isRedo ? 'Redo group balance' : 'Balance'),
      order.isPaid ? 0 : (order.balance || 0),
      '',
      ''
    ];
  });
  logSheet.getRange(logSheet.getLastRow() + 1, 1, rows.length, 11).setValues(rows);
  return { success: true, logged: rows.length };
}

function updateNotify(rowIndex, status) {
  if (!rowIndex) return { error: 'No rowIndex' };
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName(LOG_TAB);
  if (!logSheet) return { error: 'No ScanLog' };

  const row = parseInt(rowIndex, 10);
  if (row < 2 || row > logSheet.getLastRow()) return { error: 'Invalid row' };

  if (status === 'Notified') {
    logSheet.getRange(row, 10).setValue('Notified');
    logSheet.getRange(row, 11).setValue(new Date());
  } else {
    logSheet.getRange(row, 10).setValue('');
    logSheet.getRange(row, 11).setValue('');
  }
  return { success: true, row: row, status: status || '' };
}

function deleteOrder(rowIndex) {
  if (!rowIndex) return { error: 'No rowIndex' };
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName(LOG_TAB);
  if (!logSheet) return { error: 'No ScanLog' };

  const row = parseInt(rowIndex, 10);
  if (row < 2 || row > logSheet.getLastRow()) return { error: 'Invalid row' };
  logSheet.deleteRow(row);
  return { success: true, deleted: row };
}

function getHistory(dateStr) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName(LOG_TAB);
  if (!logSheet || logSheet.getLastRow() < 2) return { orders: [], total: 0 };

  const data = logSheet.getDataRange().getValues();
  const results = [];
  const date = dateStr ? new Date(dateStr) : new Date();
  const targetDay = Utilities.formatDate(date, 'Asia/Kuala_Lumpur', 'yyyy-MM-dd');

  for (let i = data.length - 1; i >= 1; i--) {
    const scanTime = data[i][0];
    if (scanTime instanceof Date &&
        Utilities.formatDate(scanTime, 'Asia/Kuala_Lumpur', 'yyyy-MM-dd') === targetDay) {
      results.push({
        rowIndex: i + 1,
        scanTime: Utilities.formatDate(scanTime, 'Asia/Kuala_Lumpur', 'hh:mm a'),
        orderId: data[i][1],
        clientName: data[i][2],
        phone: String(data[i][3] || ''),
        size: data[i][4] || '',
        salesperson: data[i][5] || '',
        source: data[i][6] || '',
        paymentStatus: data[i][7] || '',
        balance: data[i][8] || 0,
        notifyStatus: data[i][9] || '',
        notifyTime: data[i][10] instanceof Date
          ? Utilities.formatDate(data[i][10], 'Asia/Kuala_Lumpur', 'hh:mm a')
          : ''
      });
    }
  }
  return { orders: results, total: results.length, date: targetDay };
}