// app_sheet_utils.gs
// Shared spreadsheet, parsing, normalization, and access helpers.

let SS_CACHE = null;

function _isValidIndex(idx) {
  return typeof idx === 'number' && idx >= 0;
}

function getSpreadsheet() {
  if (!SS_CACHE) {
    SS_CACHE = SpreadsheetApp.getActiveSpreadsheet();
  }
  return SS_CACHE;
}

function resolveSheet(ref) {
  const ss = getSpreadsheet();
  if (typeof ref === 'number') {
    const sheets = ss.getSheets();
    if (ref >= 0 && ref < sheets.length) return sheets[ref];
    return null;
  } else if (typeof ref === 'string') {
    return ss.getSheetByName(ref) || null;
  }
  return null;
}

function findHeaderIndex(headers, candidates) {
  if (!Array.isArray(headers)) return -1;
  const lower = headers.map(h => {
    if (h === null || h === undefined) return '';
    return String(h).trim().toLowerCase();
  });
  for (let i = 0; i < candidates.length; i++) {
    const cand = (candidates[i] || '').toString().trim().toLowerCase();
    if (!cand) continue;
    const idx = lower.findIndex(h => h === cand);
    if (idx !== -1) return idx;
  }
  return -1;
}

function getColumnIndices(headers) {
  headers = headers || [];
  return {
    itemId: findHeaderIndex(headers, ['item id', 'itemid', 'รหัส', 'รหัสรายการ', 'หมายเลข', 'no.', 'id']) !== -1
      ? findHeaderIndex(headers, ['item id', 'itemid', 'รหัส', 'รหัสรายการ', 'หมายเลข', 'no.', 'id'])
      : 0,
    department: findHeaderIndex(headers, ['สำนัก/กอง', 'หน่วยงาน', 'department', 'division', 'dept']),
    work: findHeaderIndex(headers, ['งาน', 'work', 'project', 'task', 'project name']),
    budgetType: findHeaderIndex(headers, ['งบรายจ่าย', 'ประเภทงบ', 'budget type', 'budget category']),
    category: findHeaderIndex(headers, ['หมวดรายจ่าย', 'หมวด', 'category']),
    expenseType: findHeaderIndex(headers, ['ประเภทรายจ่าย', 'ประเภทการใช้จ่าย', 'expense type']),
    item: findHeaderIndex(headers, ['รายการ', 'description', 'item', 'detail', 'item description']),
    budget: findHeaderIndex(headers, ['งบประมาณ', 'budget', 'amount budget', 'amount', 'total budget']),
    used: findHeaderIndex(headers, ['เบิกจ่าย', 'used', 'spent', 'เบิกจ่ายแล้ว', 'amount used']),
    remaining: findHeaderIndex(headers, ['คงเหลือ', 'remaining', 'balance', 'left', 'balance remaining'])
  };
}

function getTransactionLogColumnIndices(headers) {
  const itemId = findHeaderIndex(headers, ['item id', 'itemid', 'item', 'รหัส', 'id']);
  const amount = findHeaderIndex(headers, ['amount']);
  const user = findHeaderIndex(headers, ['user']);
  const type = findHeaderIndex(headers, ['type']);
  const status = findHeaderIndex(headers, ['status']);
  const editedBy = findHeaderIndex(headers, ['edited by', 'editedby']);
  const quantity = findHeaderIndex(headers, ['quantity', 'qty', 'จำนวนเบิกจ่าย', 'จำนวน']);

  return {
    itemId: itemId !== -1 ? itemId : 3,
    amount: amount !== -1 ? amount : 4,
    user: user !== -1 ? user : 2,
    type: type !== -1 ? type : 9,
    status,
    editedBy,
    quantity: quantity !== -1 ? quantity : 8
  };
}

function getTransactionLogHeaders() {
  return [
    'Timestamp',
    'Expense Date',
    'User',
    'Item ID',
    'Amount',
    'Description',
    'New Used',
    'New Remaining',
    'Quantity',
    'Type',
    'Status',
    'Edited By'
  ];
}

function ensureTransactionLogSheet() {
  const ss = getSpreadsheet();
  let logSheet = resolveSheet(CONFIG.SHEETS.TRANSACTION_LOG);
  if (!logSheet) {
    logSheet = ss.insertSheet(CONFIG.SHEETS.TRANSACTION_LOG);
    logSheet.appendRow(getTransactionLogHeaders());
    try { logSheet.getRange('D:D').setNumberFormat('@'); } catch (e) {}
  }
  return logSheet;
}

function getTransactionLogContext(logSheet) {
  const sheet = logSheet || resolveSheet(CONFIG.SHEETS.TRANSACTION_LOG);
  if (!sheet) return null;

  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h) {
    return (h || '').toString().trim();
  });

  return {
    sheet: sheet,
    lastCol: lastCol,
    headers: headers,
    logCols: getTransactionLogColumnIndices(headers)
  };
}

function getTransactionLogRowModel(row, logCols, rowIndex) {
  row = row || [];
  logCols = logCols || getTransactionLogColumnIndices([]);

  return {
    logRowIndex: rowIndex || 0,
    timestamp: row[0],
    expenseDate: row[1],
    user: (row[logCols.user] || '') + '',
    itemId: normalizeItemId((row[logCols.itemId] || '') + ''),
    amount: (typeof row[logCols.amount] === 'number') ? row[logCols.amount] : parseNumberSafe(row[logCols.amount]),
    description: (row[5] || '') + '',
    newUsed: (typeof row[6] === 'number') ? row[6] : parseNumberSafe(row[6]),
    newRemaining: (typeof row[7] === 'number') ? row[7] : parseNumberSafe(row[7]),
    quantity: logCols.quantity !== -1
      ? ((typeof row[logCols.quantity] === 'number') ? row[logCols.quantity] : parseNumberSafe(row[logCols.quantity]))
      : 0,
    type: (row[logCols.type] || '') + '',
    status: logCols.status !== -1 ? (row[logCols.status] || 'ACTIVE') + '' : 'ACTIVE',
    editedBy: logCols.editedBy !== -1 ? (row[logCols.editedBy] || '') + '' : ''
  };
}

function normalizeItemId(id) {
  const idStr = (id || '').toString().trim();
  if (!idStr) return '';
  if (!/^[A-Z0-9\-\s]*$/i.test(idStr)) {
    Logger.log('WARNING: Invalid Item ID characters detected: ' + idStr);
    return '';
  }
  const prefixMatch = idStr.match(/^([A-Za-z0-9]+)[\-\s]*0*(\d+)$/);
  if (prefixMatch) {
    const prefix = prefixMatch[1].toUpperCase();
    const numPart = parseInt(prefixMatch[2], 10);
    if (!isNaN(numPart)) {
      return `${prefix}-${String(numPart).padStart(CONFIG.ITEM_ID_LENGTH, '0')}`;
    }
  }
  const anyDigits = idStr.match(/(\d+)/);
  if (anyDigits) {
    const num = parseInt(anyDigits[1], 10);
    if (!isNaN(num)) {
      return `${CONFIG.ITEM_ID_PREFIX}${String(num).padStart(CONFIG.ITEM_ID_LENGTH, '0')}`;
    }
  }
  return idStr.toUpperCase();
}

function getItemIdTrailingNumber(value) {
  const match = String(value || '').match(/(\d+)\s*$/);
  return match ? parseInt(match[1], 10) : null;
}

function getItemIdPrefix(value) {
  return String(value || '').replace(/[\-\s]*\d+\s*$/, '').trim().toUpperCase();
}

function itemIdsMatch(cellValue, itemId) {
  const searchRaw = String(itemId || '').trim();
  const cellRaw = String(cellValue || '').trim();
  if (!searchRaw || !cellRaw) return false;

  const searchNorm = normalizeItemId(searchRaw);
  const cellNorm = normalizeItemId(cellRaw);
  if (searchNorm && cellNorm && cellNorm.toUpperCase() === searchNorm.toUpperCase()) {
    return true;
  }

  const searchTrailing = getItemIdTrailingNumber(searchRaw);
  const cellTrailing = getItemIdTrailingNumber(cellRaw);
  if (searchTrailing !== null && cellTrailing !== null) {
    const searchPrefix = getItemIdPrefix(searchRaw);
    const cellPrefix = getItemIdPrefix(cellRaw);
    if (searchPrefix) {
      return cellPrefix === searchPrefix && Number(cellTrailing) === Number(searchTrailing);
    }
    return Number(cellTrailing) === Number(searchTrailing);
  }

  return cellRaw.toUpperCase() === searchRaw.toUpperCase();
}

function findBudgetRowIndicesByItemIds(itemIds, data, cols) {
  const searchIds = Array.isArray(itemIds) ? itemIds : [itemIds];
  const targets = searchIds
    .map(id => normalizeItemId((id || '').toString().trim()))
    .filter(Boolean);
  const found = {};

  targets.forEach(target => {
    found[target.toUpperCase()] = null;
  });
  if (!targets.length) return found;

  const budgetData = data || (function() {
    const budgetSheet = resolveSheet(CONFIG.SHEETS.BUDGET);
    return budgetSheet ? budgetSheet.getDataRange().getValues() : [];
  })();
  if (!budgetData || budgetData.length < 2) return found;

  const columnMap = cols || getColumnIndices(budgetData[0]);
  const idColIndex = columnMap.itemId >= 0 ? columnMap.itemId : 0;

  for (let i = 1; i < budgetData.length; i++) {
    const cellVal = (budgetData[i][idColIndex] || '').toString().trim();
    if (!cellVal) continue;

    for (let j = 0; j < targets.length; j++) {
      const target = targets[j];
      const targetKey = target.toUpperCase();
      if (found[targetKey]) continue;
      if (itemIdsMatch(cellVal, target)) {
        found[targetKey] = i + 1;
      }
    }
  }

  return found;
}

function findRowIndexByItemId(itemId) {
  const normalized = normalizeItemId((itemId || '').toString().trim());
  if (!normalized) return null;
  const found = findBudgetRowIndicesByItemIds([normalized]);
  return found[normalized.toUpperCase()] || null;
}

function findRowIndexInSheet(sheetName, itemId, columnMapFn) {
  const normalized = normalizeItemId((itemId || '').toString().trim());
  if (!normalized) return null;
  const sheet = resolveSheet(sheetName);
  if (!sheet) return null;
  const data = sheet.getDataRange().getValues();
  if (!data || data.length < 2) return null;
  const cols = (typeof columnMapFn === 'function') ? columnMapFn(data[0]) : getColumnIndices(data[0]);
  const found = findBudgetRowIndicesByItemIds([normalized], data, cols);
  return found[normalized.toUpperCase()] || null;
}

function mapBudgetRowToItem(row, cols, rowIndex) {
  return {
    itemId:      normalizeItemId((cols.itemId !== -1 && row[cols.itemId]) ? row[cols.itemId].toString().trim() : ''),
    rowIndex:    rowIndex,
    department:  row[cols.department] || '',
    work:        (cols.work        !== -1) ? (row[cols.work]        || '') : '',
    budgetType:  (cols.budgetType  !== -1) ? (row[cols.budgetType]  || '') : '',
    category:    (cols.category    !== -1) ? (row[cols.category]    || '') : '',
    expenseType: (cols.expenseType !== -1) ? (row[cols.expenseType] || '') : '',
    item:        (cols.item        !== -1) ? (row[cols.item]        || '') : '',
    budget:    parseNumberSafe(row[cols.budget]),
    used:      (cols.used      !== -1) ? parseNumberSafe(row[cols.used])      : 0,
    remaining: (cols.remaining !== -1) ? parseNumberSafe(row[cols.remaining]) : 0
  };
}

function computeBudgetAlertLevel(budget, used, remaining) {
  const effective = (remaining <= 0 && budget > 0) ? Math.max(0, budget - used) : remaining;
  const pct = budget > 0 ? (used / budget * 100) : 0;
  let level = null, message = '';
  if (effective <= 0) {
    level = 'critical'; message = 'หมดงบแล้ว';
  } else if (pct >= CONFIG.ALERT_THRESHOLD.critical) {
    level = 'critical'; message = `ใช้ไปแล้ว ${pct.toFixed(1)}% (เหลือ ${effective.toLocaleString('th-TH')} บาท)`;
  } else if (pct >= CONFIG.ALERT_THRESHOLD.high) {
    level = 'high';     message = `ใช้ไปแล้ว ${pct.toFixed(1)}% (เหลือ ${effective.toLocaleString('th-TH')} บาท)`;
  } else if (pct >= CONFIG.ALERT_THRESHOLD.medium) {
    level = 'medium';   message = `ใช้ไปแล้ว ${pct.toFixed(1)}% (เหลือ ${effective.toLocaleString('th-TH')} บาท)`;
  }
  return { level, message, percentage: +pct.toFixed(1), effective };
}

function normalizeDateInput(dateInput) {
  if (!dateInput) return null;
  try {
    const dt = (dateInput instanceof Date) ? dateInput : new Date(dateInput);
    return isNaN(dt.getTime()) ? null : dt;
  } catch (e) {
    return null;
  }
}

function createResponse(success, message = '', data = {}) {
  return { success, message, ...data };
}

function parseNumberSafe(value) {
  if (value == null) return 0;
  try {
    const normalized = String(value)
      .replace(/\u00A0/g, '')
      .replace(/[,\s]+/g, '')
      .replace(/[^\d\.\-]/g, '');
    const num = parseFloat(normalized);
    return isNaN(num) ? 0 : num;
  } catch (e) {
    return 0;
  }
}

function normalizeAccessValue(value) {
  return String(value == null ? '' : value)
    .replace(/\u00A0/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
}

function normalizeRoleValue(role) {
  const normalized = normalizeAccessValue(role);
  if (!normalized) return '';

  if ([
    'admin',
    'administrator',
    'super admin',
    'superadmin',
    'ผู้ดูแลระบบ',
    'แอดมิน'
  ].includes(normalized)) {
    return 'admin';
  }

  if ([
    'viewer',
    'read only',
    'readonly',
    'ดูอย่างเดียว',
    'ผู้อ่าน',
    'อ่านอย่างเดียว'
  ].includes(normalized)) {
    return 'viewer';
  }

  return normalized;
}

function calculateBudgetSafely(budget, used, addAmount) {
  try {
    const budgetCents = Math.round(budget * 100);
    const usedCents = Math.round(used * 100);
    const addCents = Math.round(addAmount * 100);
    const newUsedCents = usedCents + addCents;
    const newRemCents = budgetCents - newUsedCents;
    return {
      newUsed: parseFloat((Math.round(newUsedCents) / 100).toFixed(2)),
      newRemaining: parseFloat((Math.round(newRemCents) / 100).toFixed(2)),
      valid: newRemCents >= 0
    };
  } catch (e) {
    return { newUsed: used, newRemaining: budget - used, valid: false };
  }
}

function hasAccessToRow(user, rowDepartment) {
  if (!user) return false;
  if (normalizeRoleValue(user.role) === 'admin') {
    return Boolean(user.email && String(user.email).trim());
  }

  const userDept = normalizeAccessValue(user.department);
  const rowDept = normalizeAccessValue(rowDepartment);
  return Boolean(userDept && rowDept && userDept === rowDept);
}

function sanitizeHtmlForPDF(html) {
  try {
    if (!html || typeof html !== 'string') return html;
    return html
      .replace(/<script[\s\S]*?<\/script>/gi, '')
      .replace(/\s+on\w+\s*=\s*(['"])(?:(?=(\\?))\2.)*?\1/gi, '')
      .replace(/(href|src)\s*=\s*(['"])\s*javascript:[^'"]*\2/gi, '')
      .replace(/<iframe[\s\S]*?<\/iframe>/gi, '');
  } catch (e) {
    return html;
  }
}

function acquireLockWithRetry(maxRetries = CONFIG.MAX_LOCK_RETRIES) {
  const lock = LockService.getDocumentLock();
  let attempt = 0;
  while (attempt < maxRetries) {
    if (lock.tryLock(CONFIG.LOCK_TIMEOUT_MS)) return lock;
    attempt++;
    if (attempt < maxRetries) Utilities.sleep(Math.random() * (100 * attempt));
  }
  return null;
}
