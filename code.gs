// Code.gs - Main Backend Script (Improved & Optimized & Secured)
// Last Updated: 2026-03-19

// ==================== CONFIGURATION ====================
const CONFIG = {
  ITEM_ID_PREFIX: 'BG69-',
  ITEM_ID_LENGTH: 3,
  SHEETS: {
    BUDGET: 'Budget',
    USERS: 'Users',
    TRANSACTION_LOG: 'Transaction_Log',
    ERROR_LOG: 'Error_Log'
  },
  ADMIN_EMAIL: 'admin@example.com',
  ALERT_THRESHOLD: {
    critical: 95,
    high: 90,
    medium: 80
  },
  TIMEZONE: 'Asia/Bangkok',
  LOCK_TIMEOUT_MS: 5000,
  MAX_LOCK_RETRIES: 3
};

// ==================== CACHE ====================
let SS_CACHE = null;

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

// ==================== ERROR LOGGING ====================

function logErrorToSheet(errorObj) {
  try {
    const ss = getSpreadsheet();
    let errorLogSheet = resolveSheet(CONFIG.SHEETS.ERROR_LOG);
    if (!errorLogSheet) {
      errorLogSheet = ss.insertSheet(CONFIG.SHEETS.ERROR_LOG);
      errorLogSheet.appendRow(['Timestamp','User Email','Function','Error Message','Stack Trace','Context']);
    }
    errorLogSheet.appendRow([
      new Date(),
      getUserEmail() || 'system',
      errorObj.functionName || 'unknown',
      errorObj.message || '',
      errorObj.stack || '',
      JSON.stringify(errorObj.context || {})
    ]);
  } catch (e) {
    Logger.log('Error logging failed: ' + e.toString());
  }
}

function handleError(functionName, error, context = {}) {
  const errorDetails = {
    functionName,
    message: error?.message || error?.toString() || 'Unknown error',
    stack: error?.stack || '',
    timestamp: new Date().toISOString(),
    userId: getUserEmail(),
    context
  };
  Logger.log(JSON.stringify(errorDetails));
  logErrorToSheet(errorDetails);
  return errorDetails;
}

// ==================== HEADER / COLUMN HELPERS ====================

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
    itemId:      findHeaderIndex(headers, ['item id','itemid','รหัส','รหัสรายการ','หมายเลข','no.','id']) !== -1
                   ? findHeaderIndex(headers, ['item id','itemid','รหัส','รหัสรายการ','หมายเลข','no.','id']) : 0,
    department:  findHeaderIndex(headers, ['สำนัก/กอง','หน่วยงาน','department','division','dept']),
    work:        findHeaderIndex(headers, ['งาน','work','project','task','project name']),
    budgetType:  findHeaderIndex(headers, ['งบรายจ่าย','ประเภทงบ','budget type','budget category']),
    category:    findHeaderIndex(headers, ['หมวดรายจ่าย','หมวด','category']),
    expenseType: findHeaderIndex(headers, ['ประเภทรายจ่าย','ประเภทการใช้จ่าย','expense type']),
    item:        findHeaderIndex(headers, ['รายการ','description','item','detail','item description']),
    budget:      findHeaderIndex(headers, ['งบประมาณ','budget','amount budget','amount','total budget']),
    used:        findHeaderIndex(headers, ['เบิกจ่าย','used','spent','เบิกจ่ายแล้ว','amount used']),
    remaining:   findHeaderIndex(headers, ['คงเหลือ','remaining','balance','left','balance remaining'])
  };
}

function getTransactionLogColumnIndices(headers) {
  const itemId = findHeaderIndex(headers, ['item id','itemid','item','รหัส','id']);
  const amount = findHeaderIndex(headers, ['amount']);
  const user = findHeaderIndex(headers, ['user']);
  const type = findHeaderIndex(headers, ['type']);
  const status = findHeaderIndex(headers, ['status']);
  const editedBy = findHeaderIndex(headers, ['edited by','editedby']);
  const quantity = findHeaderIndex(headers, ['quantity','qty','จำนวนเบิกจ่าย','จำนวน']);

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

// ==================== ITEM ID HELPERS ====================

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

function findRowIndexByItemId(itemId) {
  const budgetSheet = resolveSheet(CONFIG.SHEETS.BUDGET);
  if (!budgetSheet) return null;
  const data = budgetSheet.getDataRange().getValues();
  if (!data || data.length < 2) return null;
  const cols = getColumnIndices(data[0]);
  const idColIndex = cols.itemId >= 0 ? cols.itemId : 0;
  const searchNorm = normalizeItemId((itemId || '').toString().trim());
  if (!searchNorm) return null;

  function trailingNumber(s) {
    const m = String(s || '').match(/(\d+)\s*$/);
    return m ? parseInt(m[1], 10) : null;
  }
  function prefixBefore(s) {
    return String(s || '').replace(/[\-\s]*\d+\s*$/, '').trim().toUpperCase();
  }

  const searchTrailing = trailingNumber(itemId);
  const searchPrefix   = prefixBefore(itemId);

  for (let i = 1; i < data.length; i++) {
    const cellVal  = (data[i][idColIndex] || '').toString().trim();
    if (!cellVal) continue;
    const cellNorm = normalizeItemId(cellVal);
    if (cellNorm.toUpperCase() === searchNorm.toUpperCase()) return i + 1;
    const cellTrailing = trailingNumber(cellVal);
    if (searchTrailing !== null && cellTrailing !== null) {
      const cellPrefix = prefixBefore(cellVal);
      if (searchPrefix) {
        if (cellPrefix === searchPrefix && Number(cellTrailing) === Number(searchTrailing)) return i + 1;
      } else {
        if (Number(cellTrailing) === Number(searchTrailing)) return i + 1;
      }
    }
  }
  return null;
}

function createResponse(success, message = '', data = {}) {
  return { success, message, ...data };
}

// ==================== INPUT VALIDATION ====================

function validateExpenseInput(itemId, amount, expenseDate) {
  const errors = [];
  if (!itemId || !String(itemId).trim()) errors.push('เธฃเธซเธฑเธชเธฃเธฒเธขเธเธฒเธฃ (Item ID) เนเธกเนเนเธ”เนเธฃเธฐเธเธธ');
  const amt = parseFloat(String(amount || '').replace(/[^\d\.\-]/g, '')) || 0;
  if (isNaN(amt) || amt <= 0) errors.push('เธเธณเธเธงเธเน€เธเธดเธเธ•เนเธญเธเธกเธฒเธเธเธงเนเธฒ 0');
  if (expenseDate) {
    const dt = (expenseDate instanceof Date) ? expenseDate : new Date(expenseDate);
    if (isNaN(dt.getTime())) errors.push('เธฃเธนเธเนเธเธเธงเธฑเธเธ—เธตเนเนเธกเนเธ–เธนเธเธ•เนเธญเธ');
  }
  return { valid: errors.length === 0, errors, sanitizedAmount: amt };
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
    'เธเธนเนเธ”เธนเนเธฅเธฃเธฐเธเธ',
    'เนเธญเธ”เธกเธดเธ'
  ].includes(normalized)) {
    return 'admin';
  }

  if ([
    'viewer',
    'read only',
    'readonly',
    'เธ”เธนเธญเธขเนเธฒเธเน€เธ”เธตเธขเธง',
    'เธเธนเนเธ”เธน',
    'เธญเนเธฒเธเธญเธขเนเธฒเธเน€เธ”เธตเธขเธง'
  ].includes(normalized)) {
    return 'viewer';
  }

  return normalized;
}

// ==================== BUDGET CALCULATION ====================

function calculateBudgetSafely(budget, used, addAmount) {
  try {
    const budgetCents    = Math.round(budget * 100);
    const usedCents      = Math.round(used * 100);
    const addCents       = Math.round(addAmount * 100);
    const newUsedCents   = usedCents + addCents;
    const newRemCents    = budgetCents - newUsedCents;
    return {
      newUsed:      parseFloat((Math.round(newUsedCents) / 100).toFixed(2)),
      newRemaining: parseFloat((Math.round(newRemCents)  / 100).toFixed(2)),
      valid: newRemCents >= 0
    };
  } catch (e) {
    return { newUsed: used, newRemaining: budget - used, valid: false };
  }
}

// ==================== ACCESS CONTROL ====================

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

// ==================== WEB APP ====================

function doGet(e) {
  try {
    const page = (e && e.parameter && e.parameter.page)
      ? String(e.parameter.page).toLowerCase().trim() : '';
    const templateName = (page === 'support') ? 'SupportIndex' : 'Index';
    return HtmlService.createTemplateFromFile(templateName)
      .evaluate()
      .setTitle('ระบบสารสนเทศในการบริหารงบประมาณ')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (err) {
    handleError('doGet', err, { page: e?.parameter?.page });
    try { return HtmlService.createTemplateFromFile('Index').evaluate(); } catch (e) {}
    return HtmlService.createHtmlOutput('Error: ' + (err.message || err.toString()));
  }
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ==================== USER MANAGEMENT ====================

function getUserEmail() {
  try { return Session.getActiveUser().getEmail(); } catch (e) { return ''; }
}

function getUserPermission() {
  try {
    const email = (Session.getActiveUser().getEmail() || '').toString().trim().toLowerCase();
    const ss = getSpreadsheet();
    let usersSheet = resolveSheet(CONFIG.SHEETS.USERS);
    if (!usersSheet) {
      usersSheet = ss.insertSheet(CONFIG.SHEETS.USERS);
      usersSheet.appendRow(['Email', 'เธชเธณเธเธฑเธ/เธเธญเธ', 'Role']);
      return null;
    }
    const data = usersSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if ((data[i][0] || '').toString().trim().toLowerCase() === email) {
        return {
          email: (data[i][0] || '').toString().trim(),
          department: (data[i][1] || '').toString().trim(),
          role: normalizeRoleValue(data[i][2]),
          rawRole: (data[i][2] || '').toString().trim()
        };
      }
    }
    return null;
  } catch (error) {
    handleError('getUserPermission', error);
    return null;
  }
}

// ==================== BUDGET ITEMS ====================

function getBudgetItems() {
  try {
    const user = getUserPermission();
    if (!user) return createResponse(false, `เนเธกเนเธเธเธเนเธญเธกเธนเธฅเธเธนเนเนเธเนเนเธเธฃเธฐเธเธ (Email: ${getUserEmail()})`);
    const budgetSheet = resolveSheet(CONFIG.SHEETS.BUDGET);
    if (!budgetSheet) return createResponse(false, 'เนเธกเนเธเธ Sheet เธเธเธเธฃเธฐเธกเธฒเธ“');
    const data = budgetSheet.getDataRange().getValues();
    if (!data || data.length < 2) return createResponse(false, 'Sheet เธเธเธเธฃเธฐเธกเธฒเธ“เนเธกเนเธกเธตเธเนเธญเธกเธนเธฅ');
    const cols = getColumnIndices(data[0]);
    if (cols.department === -1 || cols.budget === -1) {
      return createResponse(false, 'เนเธกเนเธเธ column เธ—เธตเนเธเธณเน€เธเนเธ (เธชเธณเธเธฑเธ/เธเธญเธ เธซเธฃเธทเธญ เธเธเธเธฃเธฐเธกเธฒเธ“)');
    }
    const items = data.slice(1)
      .map((row, i) => ({ row, originalIndex: i + 2 }))
      .filter(({ row }) => {
        const deptVal = row[cols.department] || '';
        const itemVal = (cols.item !== -1) ? (row[cols.item] || '') : '';
        if (!deptVal && !itemVal) return false;
        return hasAccessToRow(user, deptVal);
      })
      .map(({ row, originalIndex }) => ({
        itemId:      normalizeItemId((cols.itemId !== -1 && row[cols.itemId]) ? row[cols.itemId].toString().trim() : ''),
        rowIndex:    originalIndex,
        department:  row[cols.department] || '',
        work:        (cols.work        !== -1) ? (row[cols.work]        || '') : '',
        budgetType:  (cols.budgetType  !== -1) ? (row[cols.budgetType]  || '') : '',
        category:    (cols.category    !== -1) ? (row[cols.category]    || '') : '',
        expenseType: (cols.expenseType !== -1) ? (row[cols.expenseType] || '') : '',
        item:        (cols.item        !== -1) ? (row[cols.item]        || '') : '',
        budget:    parseNumberSafe(row[cols.budget]),
        used:      (cols.used      !== -1) ? parseNumberSafe(row[cols.used])      : 0,
        remaining: (cols.remaining !== -1) ? parseNumberSafe(row[cols.remaining]) : 0
      }));
    return createResponse(true, '', { user, items });
  } catch (error) {
    handleError('getBudgetItems', error);
    return createResponse(false, 'เน€เธเธดเธ”เธเนเธญเธเธดเธ”เธเธฅเธฒเธ”: ' + error.toString());
  }
}

// ==================== โ… NEW: getInitialData (เธฃเธงเธก items + alerts เนเธ 1 call) ====================

/**
 * เนเธซเธฅเธ”เธเนเธญเธกเธนเธฅเธ—เธฑเนเธเธซเธกเธ”เธ—เธตเน frontend เธ•เนเธญเธเธเธฒเธฃเธ•เธญเธเน€เธเธดเธ”เธซเธเนเธฒ เนเธ 1 round-trip
 * เนเธ—เธเธ—เธตเนเธเธฒเธฃเน€เธฃเธตเธขเธ getBudgetItems() + checkBudgetAlerts() เนเธขเธเธเธฑเธ
 */
function getInitialData() {
  try {
    // โ”€โ”€ 1. Auth โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€
    const user = getUserPermission();
    if (!user) return createResponse(false, `เนเธกเนเธเธเธเนเธญเธกเธนเธฅเธเธนเนเนเธเนเนเธเธฃเธฐเธเธ (Email: ${getUserEmail()})`);

    // โ”€โ”€ 2. เธญเนเธฒเธ sheet เธเธฃเธฑเนเธเน€เธ”เธตเธขเธง โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€โ”€
    const budgetSheet = resolveSheet(CONFIG.SHEETS.BUDGET);
    if (!budgetSheet) return createResponse(false, 'เนเธกเนเธเธ Sheet เธเธเธเธฃเธฐเธกเธฒเธ“');

    const data = budgetSheet.getDataRange().getValues();
    if (!data || data.length < 2) return createResponse(false, 'Sheet เธเธเธเธฃเธฐเธกเธฒเธ“เนเธกเนเธกเธตเธเนเธญเธกเธนเธฅ');

    const cols = getColumnIndices(data[0]);
    if (cols.department === -1 || cols.budget === -1) {
      return createResponse(false, 'เนเธกเนเธเธ column เธ—เธตเนเธเธณเน€เธเนเธ (เธชเธณเธเธฑเธ/เธเธญเธ เธซเธฃเธทเธญ เธเธเธเธฃเธฐเธกเธฒเธ“)');
    }

    const items  = [];
    const alerts = [];

    // โ”€โ”€ 3. เธงเธเธฅเธนเธเธเธฃเธฑเนเธเน€เธ”เธตเธขเธง เธชเธฃเนเธฒเธเธ—เธฑเนเธ items เนเธฅเธฐ alerts โ”€โ”€โ”€โ”€โ”€โ”€
    for (let i = 1; i < data.length; i++) {
      const row  = data[i];
      const dept = (row[cols.department] || '').toString().trim();
      const itemVal = (cols.item !== -1) ? (row[cols.item] || '') : '';

      if (!dept && !itemVal) continue;
      if (!hasAccessToRow(user, dept)) continue;
      const itemId = normalizeItemId((cols.itemId !== -1 && row[cols.itemId]) ? row[cols.itemId].toString().trim() : '');

      const budget    = parseNumberSafe(row[cols.budget]);
      const used      = (cols.used      !== -1) ? parseNumberSafe(row[cols.used])      : 0;
      const remaining = (cols.remaining !== -1) ? parseNumberSafe(row[cols.remaining]) : 0;

      // items
      items.push({
        itemId,
        rowIndex:    i + 1,
        department:  dept,
        work:        (cols.work        !== -1) ? (row[cols.work]        || '') : '',
        budgetType:  (cols.budgetType  !== -1) ? (row[cols.budgetType]  || '') : '',
        category:    (cols.category    !== -1) ? (row[cols.category]    || '') : '',
        expenseType: (cols.expenseType !== -1) ? (row[cols.expenseType] || '') : '',
        item:        (cols.item        !== -1) ? (row[cols.item]        || '') : '',
        budget, used, remaining
      });

      // alerts โ€” เนเธเธฅเธนเธเน€เธ”เธตเธขเธงเธเธฑเธ เนเธกเนเธ•เนเธญเธเธงเธเธเนเธณ
      const pct = budget > 0 ? (used / budget * 100) : 0;
      const eff = (remaining <= 0 && budget > 0) ? Math.max(0, budget - used) : remaining;
      let level = null, message = '';

      if (eff <= 0) {
        level = 'critical'; message = 'หมดงบแล้ว';
      } else if (pct >= CONFIG.ALERT_THRESHOLD.critical) {
        level = 'critical'; message = `ใช้ไปแล้ว ${pct.toFixed(1)}% (เหลือ ${eff.toLocaleString('th-TH')} บาท)`;
      } else if (pct >= CONFIG.ALERT_THRESHOLD.high) {
        level = 'high';     message = `ใช้ไปแล้ว ${pct.toFixed(1)}% (เหลือ ${eff.toLocaleString('th-TH')} บาท)`;
      } else if (pct >= CONFIG.ALERT_THRESHOLD.medium) {
        level = 'medium';   message = `ใช้ไปแล้ว ${pct.toFixed(1)}% (เหลือ ${eff.toLocaleString('th-TH')} บาท)`;
      }

      if (level) {
        alerts.push({
          itemId: String(itemId || '').trim(),
          work:   (cols.work  !== -1) ? String(row[cols.work]  || 'ไม่ระบุ').trim() : 'ไม่ระบุ',
          item:   (cols.item  !== -1) ? String(row[cols.item]  || 'ไม่ระบุ').trim() : 'ไม่ระบุ',
          budget: +budget, used: +used, remaining: +eff,
          percentage: +pct.toFixed(1), level, message
        });
      }
    }

    return createResponse(true, '', { user, items, alerts });
  } catch (error) {
    handleError('getInitialData', error);
    return createResponse(false, 'เน€เธเธดเธ”เธเนเธญเธเธดเธ”เธเธฅเธฒเธ”: ' + error.toString());
  }
}

// ==================== EXPENSE RECORDING ====================

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

function recordExpense(itemId, amount, description, expenseDate) {
  const startTime  = new Date();
  const currentUser = getUserPermission();           // โ เน€เธฃเธตเธขเธเธเธฃเธฑเนเธเน€เธ”เธตเธขเธง
  if (!currentUser || currentUser.role === 'viewer') {
    return createResponse(false, 'เนเธกเนเธกเธตเธชเธดเธ—เธเธดเนเธเธฑเธเธ—เธถเธเธฃเธฒเธขเธเธฒเธฃ');
  }
 
  const validation = validateExpenseInput(itemId, amount, expenseDate);
  if (!validation.valid) return createResponse(false, 'เธเนเธญเธเธดเธ”เธเธฅเธฒเธ”: ' + validation.errors.join(', '));
 
  const amt = validation.sanitizedAmount;
  const normalizedItemId = normalizeItemId(String(itemId).trim());
  if (!normalizedItemId) return createResponse(false, 'เนเธกเนเธชเธฒเธกเธฒเธฃเธ– normalize Item ID เนเธ”เน');
 
  let parsedDate = null;
  if (expenseDate) {
    try {
      parsedDate = (expenseDate instanceof Date) ? expenseDate : new Date(expenseDate);
      if (isNaN(parsedDate.getTime())) parsedDate = null;
    } catch (e) { parsedDate = null; }
  }
 
  const lock = acquireLockWithRetry();
  if (!lock) return createResponse(false, 'เธฃเธฐเธเธเธเธณเธฅเธฑเธเธเธฃเธฑเธเธเธฃเธธเธเธเนเธญเธกเธนเธฅ เธเธฃเธธเธ“เธฒเธฅเธญเธเนเธซเธกเนเธญเธตเธเธเธฃเธฑเนเธ');
 
  try {
    const rowIndex = findRowIndexByItemId(normalizedItemId);
    if (typeof rowIndex !== 'number' || rowIndex <= 1) {
      return createResponse(false, 'เนเธกเนเธเธ Item ID: ' + normalizedItemId);
    }
 
    const budgetSheet = resolveSheet(CONFIG.SHEETS.BUDGET);
    const lastCol     = budgetSheet.getLastColumn();
    const headers     = budgetSheet.getRange(1, 1, 1, lastCol).getValues()[0];
    const cols        = getColumnIndices(headers);
    if (!cols || cols.itemId === undefined || cols.itemId === -1) {
      return createResponse(false, 'เนเธกเนเธชเธฒเธกเธฒเธฃเธ–เธซเธฒเธเธญเธฅเธฑเธกเธเน itemId เนเธ”เน');
    }
 
    const values    = budgetSheet.getRange(rowIndex, 1, 1, lastCol).getValues()[0];
    const sheetId   = (values[cols.itemId]    || '').toString().trim();
    const deptInRow = (values[cols.department] || '').toString().trim();
    // โ… เนเธเน currentUser เนเธ—เธ getUserPermission() เธเนเธณ (เธฅเธ” 1 API call)
 
    if (sheetId.toUpperCase() !== normalizedItemId.toUpperCase()) {
      return createResponse(false, 'เธเนเธญเธกเธนเธฅเนเธกเนเธชเธญเธ”เธเธฅเนเธญเธเธเธฑเธ: เธฃเธซเธฑเธชเธฃเธฒเธขเธเธฒเธฃเนเธกเนเธ•เธฃเธเธเธฑเธเนเธ–เธงเธ—เธตเนเธเนเธเธเธ');
    }
    if (!hasAccessToRow(currentUser, deptInRow)) {
      return createResponse(false, 'เธเธธเธ“เนเธกเนเธกเธตเธชเธดเธ—เธเธดเนเน€เธเธดเธเธเนเธฒเธขเธเธฒเธเนเธเธเธเธเธตเน');
    }
 
    const currentUsed = parseNumberSafe(values[cols.used]   || 0);
    const budget      = parseNumberSafe(values[cols.budget] || 0);
    const calc = calculateBudgetSafely(budget, currentUsed, amt);
    if (!calc.valid) return createResponse(false, 'เธขเธญเธ”เน€เธเธดเธเธเนเธฒเธขเน€เธเธดเธเธเธเธเธฃเธฐเธกเธฒเธ“เธ—เธตเนเธ•เธฑเนเธเนเธงเน');
 
    const { newUsed, newRemaining } = calc;
 
    try {
      // โ… batch write: เธ–เนเธฒ used เธเธฑเธ remaining เธญเธขเธนเนเธ•เธดเธ”เธเธฑเธ โ’ 1 setValues() เนเธ—เธ 2 setValue()
      if (Math.abs(cols.remaining - cols.used) === 1) {
        const startCol = Math.min(cols.used, cols.remaining) + 1;
        budgetSheet.getRange(rowIndex, startCol, 1, 2).setValues([
          cols.used < cols.remaining ? [newUsed, newRemaining] : [newRemaining, newUsed]
        ]);
      } else {
        // เธเธญเธฅเธฑเธกเธเนเนเธกเนเธ•เธดเธ”เธเธฑเธ โ€” เธเธณเน€เธเนเธเธ•เนเธญเธ write เนเธขเธ 2 เธเธฃเธฑเนเธ
        budgetSheet.getRange(rowIndex, cols.used      + 1).setValue(newUsed);
        budgetSheet.getRange(rowIndex, cols.remaining + 1).setValue(newRemaining);
      }
    } catch (writeErr) {
      handleError('recordExpense - sheet write', writeErr, { rowIndex });
      return createResponse(false, 'เน€เธเธดเธ”เธเนเธญเธเธดเธ”เธเธฅเธฒเธ”เนเธเธเธฒเธฃเธเธฑเธเธ—เธถเธเธเนเธญเธกเธนเธฅ');
    }
 
    try {
      logTransaction(normalizedItemId, amt, description || '', parsedDate || '', newUsed, newRemaining);
    } catch (logErr) {
      handleError('recordExpense - transaction log', logErr);
    }
 
    Logger.log('recordExpense completed in %sms', new Date() - startTime);
    return createResponse(true, 'เธเธฑเธเธ—เธถเธเธชเธณเน€เธฃเนเธ', {
      newUsed, newRemaining, timestamp: new Date().toISOString()
    });
 
  } catch (err) {
    handleError('recordExpense', err, { itemId: normalizedItemId, amount: amt });
    return createResponse(false, 'เน€เธเธดเธ”เธเนเธญเธเธดเธ”เธเธฅเธฒเธ”: ' + err.toString());
  } finally {
    try { if (lock) lock.releaseLock(); } catch (e) {}
  }
}

// ==================== BUDGET TRANSFER ====================

function transferBudget(fromItemId, toItemId, amount, reason) {
  const currentUser = getUserPermission();
  // โ… เน€เธเธฅเธตเนเธขเธ: เธ—เธธเธ role เนเธเนเนเธ”เน (เธฅเธ admin check เธญเธญเธ)
  if (!currentUser) {
    return createResponse(false, 'เนเธกเนเธเธเธเนเธญเธกเธนเธฅเธเธนเนเนเธเนเนเธเธฃเธฐเธเธ');
  }
  if (currentUser.role === 'viewer') {
    return createResponse(false, 'เธเธนเนเนเธเนเธเธฃเธฐเน€เธ เธ— Viewer เนเธกเนเธชเธฒเธกเธฒเธฃเธ–เนเธญเธเธเธเนเธ”เน');
  }
 
  const errors   = [];
  const normFrom = normalizeItemId(String(fromItemId || '').trim());
  const normTo   = normalizeItemId(String(toItemId   || '').trim());
  const amt      = parseNumberSafe(amount);
 
  if (!normFrom)              errors.push('เนเธกเนเธฃเธฐเธเธธ Item ID เธ•เนเธเธ—เธฒเธ');
  if (!normTo)                errors.push('เนเธกเนเธฃเธฐเธเธธ Item ID เธเธฅเธฒเธขเธ—เธฒเธ');
  if (normFrom === normTo)    errors.push('เธ•เนเธเธ—เธฒเธเนเธฅเธฐเธเธฅเธฒเธขเธ—เธฒเธเธ•เนเธญเธเนเธกเนเนเธเนเธฃเธฒเธขเธเธฒเธฃเน€เธ”เธตเธขเธงเธเธฑเธ');
  if (isNaN(amt) || amt <= 0) errors.push('เธเธณเธเธงเธเน€เธเธดเธเธ•เนเธญเธเธกเธฒเธเธเธงเนเธฒ 0');
  if (!reason || !String(reason).trim()) errors.push('เธเธฃเธธเธ“เธฒเธฃเธฐเธเธธเน€เธซเธ•เธธเธเธฅเธเธฒเธฃเนเธญเธเธเธ');
  if (errors.length) return createResponse(false, 'เธเนเธญเธเธดเธ”เธเธฅเธฒเธ”: ' + errors.join(', '));
 
  const lock = acquireLockWithRetry();
  if (!lock) return createResponse(false, 'เธฃเธฐเธเธเธเธณเธฅเธฑเธเธเธฃเธฐเธกเธงเธฅเธเธฅ เธเธฃเธธเธ“เธฒเธฅเธญเธเนเธซเธกเนเธญเธตเธเธเธฃเธฑเนเธ');
 
  try {
    const budgetSheet = resolveSheet(CONFIG.SHEETS.BUDGET);
    if (!budgetSheet) return createResponse(false, 'เนเธกเนเธเธ Sheet เธเธเธเธฃเธฐเธกเธฒเธ“');
 
    // เธญเนเธฒเธ data เธเธฃเธฑเนเธเน€เธ”เธตเธขเธง
    const allData = budgetSheet.getDataRange().getValues();
    const cols    = getColumnIndices(allData[0]);
 
    if (cols.budget === -1 || cols.used === -1 || cols.remaining === -1) {
      return createResponse(false, 'เนเธกเนเธเธเธเธญเธฅเธฑเธกเธเนเธ—เธตเนเธเธณเน€เธเนเธ');
    }
 
    // เธซเธฒ row เธ—เธฑเนเธเธเธนเนเนเธ loop เน€เธ”เธตเธขเธง
    let fromRow = null, toRow = null;
    for (let i = 1; i < allData.length; i++) {
      const cellId = normalizeItemId((allData[i][cols.itemId] || '').toString());
      if (fromRow === null && cellId.toUpperCase() === normFrom.toUpperCase()) fromRow = i + 1;
      if (toRow   === null && cellId.toUpperCase() === normTo.toUpperCase())   toRow   = i + 1;
      if (fromRow !== null && toRow !== null) break;
    }
 
    if (!fromRow) return createResponse(false, 'เนเธกเนเธเธ Item ID เธ•เนเธเธ—เธฒเธ: '  + normFrom);
    if (!toRow)   return createResponse(false, 'เนเธกเนเธเธ Item ID เธเธฅเธฒเธขเธ—เธฒเธ: ' + normTo);
 
    const fromValues = allData[fromRow - 1];
    const toValues   = allData[toRow   - 1];
 
    // โ… เธ•เธฃเธงเธเธชเธดเธ—เธเธดเน: เธ•เนเธญเธเธกเธตเธชเธดเธ—เธเธดเนเน€เธเนเธฒเธ–เธถเธเธ—เธฑเนเธเธ•เนเธเธ—เธฒเธเนเธฅเธฐเธเธฅเธฒเธขเธ—เธฒเธ
    const fromDept = (fromValues[cols.department] || '').toString().trim();
    const toDept   = (toValues[cols.department]   || '').toString().trim();
 
    if (!hasAccessToRow(currentUser, fromDept)) {
      return createResponse(false, 'เธเธธเธ“เนเธกเนเธกเธตเธชเธดเธ—เธเธดเนเนเธญเธเธเธเธเธฒเธเธฃเธฒเธขเธเธฒเธฃเธ•เนเธเธ—เธฒเธ (เธซเธเนเธงเธขเธเธฒเธ: ' + fromDept + ')');
    }
    if (!hasAccessToRow(currentUser, toDept)) {
      return createResponse(false, 'เธเธธเธ“เนเธกเนเธกเธตเธชเธดเธ—เธเธดเนเนเธญเธเธเธเนเธเธขเธฑเธเธฃเธฒเธขเธเธฒเธฃเธเธฅเธฒเธขเธ—เธฒเธ (เธซเธเนเธงเธขเธเธฒเธ: ' + toDept + ')');
    }
 
    const fromBudget    = parseNumberSafe(fromValues[cols.budget]);
    const fromUsed      = parseNumberSafe(fromValues[cols.used]);
    const fromRemaining = parseNumberSafe(fromValues[cols.remaining]);
    const toBudget      = parseNumberSafe(toValues[cols.budget]);
    const toUsed        = parseNumberSafe(toValues[cols.used]);
 
    if (amt > fromRemaining) {
      return createResponse(false,
        `เธเธเธเธเน€เธซเธฅเธทเธญเธเธญเธเธ•เนเธเธ—เธฒเธเนเธกเนเน€เธเธตเธขเธเธเธญ (เธเธเน€เธซเธฅเธทเธญ: ${fromRemaining.toLocaleString('th-TH')} เธเธฒเธ—)`);
    }
 
    const newFromBudget    = parseFloat((fromBudget    - amt).toFixed(2));
    const newFromRemaining = parseFloat((fromRemaining - amt).toFixed(2));
    const newToBudget      = parseFloat((toBudget      + amt).toFixed(2));
    const newToRemaining   = parseFloat((newToBudget   - toUsed).toFixed(2));
 
    // batch write
    const budgetCol    = cols.budget    + 1;
    const remainingCol = cols.remaining + 1;
 
    if (Math.abs(cols.budget - cols.remaining) === 1) {
      const startCol = Math.min(budgetCol, remainingCol);
      budgetSheet.getRange(fromRow, startCol, 1, 2).setValues([
        cols.budget < cols.remaining
          ? [newFromBudget, newFromRemaining]
          : [newFromRemaining, newFromBudget]
      ]);
      budgetSheet.getRange(toRow, startCol, 1, 2).setValues([
        cols.budget < cols.remaining
          ? [newToBudget, newToRemaining]
          : [newToRemaining, newToBudget]
      ]);
    } else {
      budgetSheet.getRange(fromRow, budgetCol).setValue(newFromBudget);
      budgetSheet.getRange(fromRow, remainingCol).setValue(newFromRemaining);
      budgetSheet.getRange(toRow,   budgetCol).setValue(newToBudget);
      budgetSheet.getRange(toRow,   remainingCol).setValue(newToRemaining);
    }
 
    const note = `[TRANSFER] ${reason} (${normFrom} โ’ ${normTo})`;
    logTransaction(normFrom, -amt, note, new Date(), fromUsed, newFromRemaining, 'TRANSFER_OUT');
    logTransaction(normTo,    amt, note, new Date(), toUsed,   newToRemaining,   'TRANSFER_IN');
 
    return createResponse(true,
      `เนเธญเธเธเธเธชเธณเน€เธฃเนเธ: ${normFrom} โ’ ${normTo} เธเธณเธเธงเธ ${amt.toLocaleString('th-TH')} เธเธฒเธ—`, {
        from: { itemId: normFrom, newBudget: newFromBudget, newRemaining: newFromRemaining },
        to:   { itemId: normTo,   newBudget: newToBudget,   newRemaining: newToRemaining },
        amount: amt, transferredBy: currentUser.email, timestamp: new Date().toISOString()
      });
 
  } catch (err) {
    handleError('transferBudget', err, { fromItemId: normFrom, toItemId: normTo, amount: amt });
    return createResponse(false, 'เน€เธเธดเธ”เธเนเธญเธเธดเธ”เธเธฅเธฒเธ”: ' + err.toString());
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

// ==================== TRANSACTION LOG ====================

function normalizeDateForSheet(dateInput) {
  if (!dateInput) return null;
  try {
    const dt = (dateInput instanceof Date) ? dateInput : new Date(dateInput);
    return isNaN(dt.getTime()) ? null : dt;
  } catch (e) { return null; }
}

function logTransaction(itemId, amount, description, expenseDate, newUsed, newRemaining, type, quantity, status, editedBy) {
  try {
    const ss = getSpreadsheet();
    let logSheet = resolveSheet(CONFIG.SHEETS.TRANSACTION_LOG);
    if (!logSheet) {
      logSheet = ss.insertSheet(CONFIG.SHEETS.TRANSACTION_LOG);
      logSheet.appendRow([
        'Timestamp','Expense Date','User','Item ID','Amount',
        'Description','New Used','New Remaining','Quantity','Type',
        'Status','Edited By'
      ]);
      try { logSheet.getRange('D:D').setNumberFormat('@'); } catch (e) {}
    }
    logSheet.appendRow([
      new Date(),
      normalizeDateForSheet(expenseDate),
      getUserEmail(),
      normalizeItemId(itemId),
      Number(amount || 0) || 0,
      description || '',
      newUsed,
      newRemaining,
      Number(quantity || 0) || 0,
      type     || '',
      status   || 'ACTIVE',
      editedBy || ''
    ]);
  } catch (error) {
    handleError('logTransaction', error, { itemId, amount });
  }
}

function getTransactionHistory(itemId) {
  try {
    const inputId = normalizeItemId(itemId || '');
    if (!inputId) return [];
    const logSheet = resolveSheet(CONFIG.SHEETS.TRANSACTION_LOG);
    if (!logSheet) return [];
    const data = logSheet.getDataRange().getValues();
    if (!data || data.length < 2) return [];
 
    const headers = (data[0] || []).map(h => (h || '').toString().trim());
    const logCols = getTransactionLogColumnIndices(headers);
 
    const tz = CONFIG.TIMEZONE;
 
    function formatDate(val) {
      if (!val && val !== 0) return '';
      try {
        if (val instanceof Date) return Utilities.formatDate(val, tz, 'yyyy-MM-dd');
        if (typeof val === 'number') {
          const d = new Date(Math.round((val - 25569) * 86400 * 1000));
          if (!isNaN(d.getTime())) return Utilities.formatDate(d, tz, 'yyyy-MM-dd');
        }
        const p = new Date(String(val));
        if (!isNaN(p.getTime())) return Utilities.formatDate(p, tz, 'yyyy-MM-dd');
      } catch (e) {}
      return '';
    }
    function formatTs(val) {
      if (!val && val !== 0) return '';
      try {
        if (val instanceof Date) return val.toISOString();
        if (typeof val === 'number') return new Date(val).toISOString();
        const p = new Date(String(val));
        if (!isNaN(p.getTime())) return p.toISOString();
      } catch (e) {}
      return '';
    }
 
    const history = [];
    for (let i = data.length - 1; i >= 1; i--) {
      const row = data[i] || [];
      const logIdNorm = normalizeItemId((row[logCols.itemId] || '') + '');
      if (!logIdNorm || logIdNorm.toUpperCase() !== inputId.toUpperCase()) continue;
 
      const type = (row[logCols.type] || '') + '';

      // Filter internal reversal rows; keep active/edited/cancelled visible.
      if (type === 'REVERSAL') continue;
 
      const status = logCols.status !== -1 ? (row[logCols.status] || 'ACTIVE') + '' : 'ACTIVE';
      const editedBy = logCols.editedBy !== -1 ? (row[logCols.editedBy] || '') + '' : '';
      const quantity = logCols.quantity !== -1
        ? ((typeof row[logCols.quantity] === 'number') ? row[logCols.quantity] : parseNumberSafe(row[logCols.quantity]))
        : 0;
 
      history.push({
        logRowIndex:  i + 1,
        timestamp:    formatTs(row[0]),
        expenseDate:  formatDate(row[1]),
        user:         (row[logCols.user] || '') + '',
        amount:       (typeof row[logCols.amount] === 'number') ? row[logCols.amount] : parseNumberSafe(row[logCols.amount]),
        description:  (row[5] || '') + '',
        newUsed:      (typeof row[6] === 'number') ? row[6] : parseNumberSafe(row[6]),
        newRemaining: (typeof row[7] === 'number') ? row[7] : parseNumberSafe(row[7]),
        quantity,
        type,
        status,
        editedBy
      });
    }
    return history;
  } catch (error) {
    handleError('getTransactionHistory', error, { itemId });
    return [];
  }
}

function getTransferTotalsByItemId() {
  try {
    const logSheet = resolveSheet(CONFIG.SHEETS.TRANSACTION_LOG);
    if (!logSheet) return {};

    const data = logSheet.getDataRange().getValues();
    if (!data || data.length < 2) return {};

    const headers = (data[0] || []).map(h => (h || '').toString().trim());
    const logCols = getTransactionLogColumnIndices(headers);

    const totals = {};
    for (let i = 1; i < data.length; i++) {
      const row = data[i] || [];
      const status = logCols.status !== -1 ? String(row[logCols.status] || 'ACTIVE').trim().toUpperCase() : 'ACTIVE';
      if (status === 'CANCELLED') continue;

      const type = String(row[logCols.type] || '').trim().toUpperCase();
      if (type !== 'TRANSFER_IN' && type !== 'TRANSFER_OUT') continue;

      const itemId = normalizeItemId((row[logCols.itemId] || '') + '');
      if (!itemId) continue;

      const amount = parseNumberSafe(row[logCols.amount]);
      if (!amount) continue;

      if (!totals[itemId]) {
        totals[itemId] = { transferIn: 0, transferOut: 0, netTransfer: 0 };
      }

      if (type === 'TRANSFER_IN') {
        totals[itemId].transferIn = parseFloat((totals[itemId].transferIn + Math.abs(amount)).toFixed(2));
      } else {
        totals[itemId].transferOut = parseFloat((totals[itemId].transferOut + Math.abs(amount)).toFixed(2));
      }
      totals[itemId].netTransfer = parseFloat((totals[itemId].transferIn - totals[itemId].transferOut).toFixed(2));
    }

    return totals;
  } catch (error) {
    handleError('getTransferTotalsByItemId', error);
    return {};
  }
}

// ==================== DASHBOARD ====================

function getDashboardData() {
  try {
    const user = getUserPermission();
    if (!user) return createResponse(false, 'เนเธกเนเธเธเธเนเธญเธกเธนเธฅเธเธนเนเนเธเน');
    const budgetSheet = resolveSheet(CONFIG.SHEETS.BUDGET);
    if (!budgetSheet) return createResponse(false, 'เนเธกเนเธเธ Sheet เธเธเธเธฃเธฐเธกเธฒเธ“');
    const data = budgetSheet.getDataRange().getValues();
    if (!data || data.length < 2) return createResponse(false, 'Sheet เธเธเธเธฃเธฐเธกเธฒเธ“เนเธกเนเธกเธตเธเนเธญเธกเธนเธฅ');
    const cols = getColumnIndices(data[0]);
    const transferTotals = getTransferTotalsByItemId();
    const workSummary = {};
    for (let i = 1; i < data.length; i++) {
      const row  = data[i];
      const dept = row[cols.department] || '';
      if (!hasAccessToRow(user, dept)) continue;
      const work = row[cols.work] || 'เนเธกเนเธฃเธฐเธเธธ';
      if (!workSummary[work]) workSummary[work] = { work, totalBudget:0, totalUsed:0, totalTransferIn:0, totalTransferOut:0, totalRemaining:0, items:0 };
      const itemId = normalizeItemId((cols.itemId !== -1 && row[cols.itemId]) ? row[cols.itemId].toString().trim() : '');
      const transfer = transferTotals[itemId] || {};
      workSummary[work].totalBudget    += parseNumberSafe(row[cols.budget]);
      workSummary[work].totalUsed      += parseNumberSafe(row[cols.used]);
      workSummary[work].totalTransferIn += parseNumberSafe(transfer.transferIn || 0);
      workSummary[work].totalTransferOut += parseNumberSafe(transfer.transferOut || 0);
      workSummary[work].totalRemaining += parseNumberSafe(row[cols.remaining]);
      workSummary[work].items          += 1;
    }
    return createResponse(true, '', { data: Object.values(workSummary), user });
  } catch (error) {
    handleError('getDashboardData', error);
    return createResponse(false, 'เน€เธเธดเธ”เธเนเธญเธเธดเธ”เธเธฅเธฒเธ”: ' + error.toString());
  }
}

function getWorkDetails(workName) {
  try {
    const user = getUserPermission();
    if (!user) return createResponse(false, 'เนเธกเนเธเธเธเนเธญเธกเธนเธฅเธเธนเนเนเธเน');
    const budgetSheet = resolveSheet(CONFIG.SHEETS.BUDGET);
    if (!budgetSheet) return createResponse(false, 'เนเธกเนเธเธ Sheet เธเธเธเธฃเธฐเธกเธฒเธ“');
    const data = budgetSheet.getDataRange().getValues();
    if (!data || data.length < 2) return createResponse(false, 'Sheet เธเธเธเธฃเธฐเธกเธฒเธ“เนเธกเนเธกเธตเธเนเธญเธกเธนเธฅ');
    const cols = getColumnIndices(data[0]);
    const transferTotals = getTransferTotalsByItemId();
    const detailedData = [];
    for (let i = 1; i < data.length; i++) {
      const row  = data[i];
      const dept = row[cols.department] || '';
      if (!hasAccessToRow(user, dept) || row[cols.work] !== workName) continue;
      const itemId = normalizeItemId((cols.itemId !== -1 && row[cols.itemId]) ? row[cols.itemId].toString().trim() : '');
      const transfer = transferTotals[itemId] || {};
      detailedData.push({
        work:        row[cols.work]        || 'เนเธกเนเธฃเธฐเธเธธ',
        budgetType:  (cols.budgetType  !== -1) ? (row[cols.budgetType]  || 'เนเธกเนเธฃเธฐเธเธธ') : 'เนเธกเนเธฃเธฐเธเธธ',
        category:    (cols.category    !== -1) ? (row[cols.category]    || 'เนเธกเนเธฃเธฐเธเธธ') : 'เนเธกเนเธฃเธฐเธเธธ',
        expenseType: (cols.expenseType !== -1) ? (row[cols.expenseType] || 'เนเธกเนเธฃเธฐเธเธธ') : 'เนเธกเนเธฃเธฐเธเธธ',
        item:        (cols.item        !== -1) ? (row[cols.item]        || 'เนเธกเนเธฃเธฐเธเธธ') : 'เนเธกเนเธฃเธฐเธเธธ',
        budget:    parseNumberSafe(row[cols.budget]),
        used:      parseNumberSafe(row[cols.used]),
        transferIn: parseNumberSafe(transfer.transferIn || 0),
        transferOut: parseNumberSafe(transfer.transferOut || 0),
        remaining: parseNumberSafe(row[cols.remaining])
      });
    }
    return createResponse(true, '', { detailedData });
  } catch (error) {
    handleError('getWorkDetails', error, { workName });
    return createResponse(false, 'เน€เธเธดเธ”เธเนเธญเธเธดเธ”เธเธฅเธฒเธ”: ' + error.toString());
  }
}

// ==================== ALERT & NOTIFICATION ====================

function checkBudgetAlerts() {
  try {
    const user = getUserPermission();
    if (!user) return { success: false, message: 'เนเธกเนเธเธเธเนเธญเธกเธนเธฅเธเธนเนเนเธเน', alerts: [] };
    const budgetSheet = resolveSheet(CONFIG.SHEETS.BUDGET);
    if (!budgetSheet) return { success: false, message: 'เนเธกเนเธเธ Sheet เธเธเธเธฃเธฐเธกเธฒเธ“', alerts: [] };
    const data = budgetSheet.getDataRange().getValues();
    if (!data || data.length < 2) return { success: true, message: '', alerts: [] };
    const cols   = getColumnIndices(data[0]);
    const alerts = [];
    for (let i = 1; i < data.length; i++) {
      const row  = data[i];
      const dept = (cols.department != null) ? (row[cols.department] || '').toString().trim() : '';
      if (!hasAccessToRow(user, dept)) continue;
      const budget    = parseNumberSafe(cols.budget    != null ? row[cols.budget]    : 0);
      const used      = parseNumberSafe(cols.used      != null ? row[cols.used]      : 0);
      let remaining   = parseNumberSafe(cols.remaining != null ? row[cols.remaining] : 0);
      if (remaining === 0 && budget > 0) remaining = Math.max(0, budget - used);
      const pct = budget > 0 ? (used / budget * 100) : 0;
      let level = null, message = '';
      if (remaining <= 0)                              { level='critical'; message='หมดงบแล้ว'; }
      else if (pct >= CONFIG.ALERT_THRESHOLD.critical) { level='critical'; message=`ใช้ไปแล้ว ${pct.toFixed(1)}% (เหลือ ${remaining.toLocaleString('th-TH')} บาท)`; }
      else if (pct >= CONFIG.ALERT_THRESHOLD.high)     { level='high';     message=`ใช้ไปแล้ว ${pct.toFixed(1)}% (เหลือ ${remaining.toLocaleString('th-TH')} บาท)`; }
      else if (pct >= CONFIG.ALERT_THRESHOLD.medium)   { level='medium';   message=`ใช้ไปแล้ว ${pct.toFixed(1)}% (เหลือ ${remaining.toLocaleString('th-TH')} บาท)`; }
      if (level) alerts.push({
        itemId: String(normalizeItemId(cols.itemId != null ? row[cols.itemId] : '') || '').trim(),
        department: dept,
        work:   String(cols.work != null ? row[cols.work] || 'ไม่ระบุ' : 'ไม่ระบุ').trim(),
        item:   String(cols.item != null ? row[cols.item] || 'ไม่ระบุ' : 'ไม่ระบุ').trim(),
        budget: +budget, used: +used, remaining: +remaining,
        percentage: +pct.toFixed(1), level, message
      });
    }
    return { success: true, message: '', alerts };
  } catch (error) {
    handleError('checkBudgetAlerts', error);
    return { success: false, message: 'เน€เธเธดเธ”เธเนเธญเธเธดเธ”เธเธฅเธฒเธ”: ' + (error?.message || error?.toString()), alerts: [] };
  }
}

function getBudgetAlertRecipients() {
  try {
    const usersSheet = resolveSheet(CONFIG.SHEETS.USERS);
    if (!usersSheet) return [];
    const data = usersSheet.getDataRange().getValues();
    if (!data || data.length < 2) return [];

    const recipients = [];
    for (let i = 1; i < data.length; i++) {
      const email = (data[i][0] || '').toString().trim();
      const department = (data[i][1] || '').toString().trim();
      const role = normalizeRoleValue(data[i][2]);
      if (!email || !role) continue;
      recipients.push({ email, department, role });
    }
    return recipients;
  } catch (error) {
    handleError('getBudgetAlertRecipients', error);
    return [];
  }
}

function buildBudgetAlertEmailBody(recipientEmail, alerts, dateStr) {
  const critical = alerts.filter(a => a.level === 'critical');
  const high = alerts.filter(a => a.level === 'high');
  const medium = alerts.filter(a => a.level === 'medium');

  const section = (arr, color, emoji, label) => {
    if (!arr.length) return '';
    return `<h3 style="color:${color};margin:18px 0 8px;">${emoji} ${label}</h3><ul style="padding-left:18px;">` +
      arr.map(a => {
        const deptText = a.department ? `${a.department} / ` : '';
        return `<li style="margin-bottom:10px;"><strong>${a.itemId}</strong> - ${deptText}${a.work} > ${a.item}<br><span style="color:${color}">${a.message}</span></li>`;
      }).join('') +
      '</ul>';
  };

  return `<html><body style="font-family:Sarabun,Arial,sans-serif;">
    <h2 style="color:#667eea;">๐”” เธฃเธฒเธขเธเธฒเธเธชเธ–เธฒเธเธฐเธเธเธเธฃเธฐเธกเธฒเธ“เธเธฃเธฐเธเธณเธงเธฑเธ (${dateStr})</h2>
    <p>เน€เธฃเธตเธขเธ ${recipientEmail}</p>
    <p>เธเธเธฃเธฒเธขเธเธฒเธฃเนเธเนเธเน€เธ•เธทเธญเธเธ—เธฑเนเธเธซเธกเธ” <strong>${alerts.length}</strong> เธฃเธฒเธขเธเธฒเธฃ</p>
    ${section(critical, '#ff4444', '๐จ', 'เน€เธฃเนเธเธ”เนเธงเธ (เธซเธกเธ”เธเธเธซเธฃเธทเธญเน€เธซเธฅเธทเธญเธเนเธญเธขเธเธงเนเธฒ 5%)')}
    ${section(high, '#FF9800', 'โ ๏ธ', 'เธเธงเธฃเธฃเธฐเธงเธฑเธ (เนเธเน 90-94%)')}
    ${section(medium, '#FFC107', '๐“', 'เธ•เธดเธ”เธ•เธฒเธก (เนเธเน 80-89%)')}
    <hr><p style="font-size:12px;color:#666;">เธชเนเธเธเธฒเธเธฃเธฐเธเธเธเธฑเธเธ—เธถเธเธเธฒเธฃเนเธเนเธเธเธเธฃเธฐเธกเธฒเธ“เธญเธฑเธ•เนเธเธกเธฑเธ•เธด</p>
  </body></html>`;
}

function sendDailyBudgetAlertLegacy() { // deprecated
  try {
    const result = checkBudgetAlerts();
    if (!result || !result.success || !result.alerts.length) return;
    const user      = getUserPermission();
    const recipient = (user && user.email) ? user.email : CONFIG.ADMIN_EMAIL;
    if (!recipient) return;
    const critical = result.alerts.filter(a => a.level === 'critical');
    const high     = result.alerts.filter(a => a.level === 'high');
    const medium   = result.alerts.filter(a => a.level === 'medium');
    const dateStr  = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'dd/MM/yyyy');
    let body = `<html><body style="font-family:Sarabun,Arial,sans-serif;">
      <h2 style="color:#667eea;">๐”” เธฃเธฒเธขเธเธฒเธเธชเธ–เธฒเธเธฐเธเธเธเธฃเธฐเธกเธฒเธ“เธเธฃเธฐเธเธณเธงเธฑเธ (${dateStr})</h2>
      <p>เน€เธฃเธตเธขเธ ${recipient}</p>`;
    const section = (arr, color, emoji, label) => {
      if (!arr.length) return '';
      return `<h3 style="color:${color};">${emoji} ${label}</h3><ul>` +
        arr.map(a => `<li><strong>${a.itemId}</strong> โ€ข ${a.work} > ${a.item}<br><span style="color:${color}">${a.message}</span></li>`).join('') +
        '</ul>';
    };
    body += section(critical, '#ff4444', '๐จ', 'เน€เธฃเนเธเธ”เนเธงเธ (เธซเธกเธ”เธเธเธซเธฃเธทเธญเน€เธซเธฅเธทเธญเธเนเธญเธขเธเธงเนเธฒ 5%)');
    body += section(high,     '#FF9800', 'โ ๏ธ', 'เธเธงเธฃเธฃเธฐเธงเธฑเธ (เนเธเน 90-94%)');
    body += section(medium,   '#FFC107', '๐“', 'เธ•เธดเธ”เธ•เธฒเธก (เนเธเน 80-89%)');
    body += '<hr><p style="font-size:12px;color:#666;">เธชเนเธเธเธฒเธเธฃเธฐเธเธเธเธฑเธเธ—เธถเธเธเธฒเธฃเนเธเนเธเธเธเธฃเธฐเธกเธฒเธ“เธญเธฑเธ•เนเธเธกเธฑเธ•เธด</p></body></html>';
    MailApp.sendEmail({
      to: recipient,
      subject: `๐”” เนเธเนเธเน€เธ•เธทเธญเธเธชเธ–เธฒเธเธฐเธเธเธเธฃเธฐเธกเธฒเธ“ - ${dateStr}`,
      htmlBody: body,
      name: 'เธฃเธฐเธเธเธเธฑเธเธ—เธถเธเธเธฒเธฃเนเธเนเธเธเธเธฃเธฐเธกเธฒเธ“'
    });
  } catch (error) {
    handleError('sendDailyBudgetAlert', error);
  }
}

function sendDailyBudgetAlert() {
  try {
    const result = checkBudgetAlerts();
    if (!result || !result.success || !result.alerts.length) return;

    const alerts = result.alerts.slice();
    const recipients = getBudgetAlertRecipients();
    const dateStr = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'dd/MM/yyyy');
    const sent = {};

    recipients
      .slice()
      .sort((a, b) => {
        const ra = normalizeRoleValue(a.role);
        const rb = normalizeRoleValue(b.role);
        if (ra === rb) return 0;
        if (ra === 'admin') return -1;
        if (rb === 'admin') return 1;
        if (ra === 'head') return -1;
        if (rb === 'head') return 1;
        return 0;
      })
      .forEach(recipient => {
        const email = (recipient.email || '').toString().trim();
        const emailKey = email.toLowerCase();
        const role = normalizeRoleValue(recipient.role);
        const dept = normalizeAccessValue(recipient.department);
        if (!email || sent[emailKey]) return;

        let scopedAlerts = [];
        if (role === 'admin') {
          scopedAlerts = alerts;
        } else if (role === 'head') {
          scopedAlerts = alerts.filter(a => normalizeAccessValue(a.department) === dept);
        } else {
          return;
        }

        if (!scopedAlerts.length) return;

        MailApp.sendEmail({
          to: email,
          subject: `๐”” เนเธเนเธเน€เธ•เธทเธญเธเธชเธ–เธฒเธเธฐเธเธเธเธฃเธฐเธกเธฒเธ“ - ${dateStr}`,
          htmlBody: buildBudgetAlertEmailBody(email, scopedAlerts, dateStr),
          name: 'เธฃเธฐเธเธเธเธฑเธเธ—เธถเธเธเธฒเธฃเนเธเนเธเธเธเธฃเธฐเธกเธฒเธ“'
        });
        sent[emailKey] = true;
      });

    if (!Object.keys(sent).length && CONFIG.ADMIN_EMAIL) {
      MailApp.sendEmail({
        to: CONFIG.ADMIN_EMAIL,
        subject: `๐”” เนเธเนเธเน€เธ•เธทเธญเธเธชเธ–เธฒเธเธฐเธเธเธเธฃเธฐเธกเธฒเธ“ - ${dateStr}`,
        htmlBody: buildBudgetAlertEmailBody(CONFIG.ADMIN_EMAIL, alerts, dateStr),
        name: 'เธฃเธฐเธเธเธเธฑเธเธ—เธถเธเธเธฒเธฃเนเธเนเธเธเธเธฃเธฐเธกเธฒเธ“'
      });
    }
  } catch (error) {
    handleError('sendDailyBudgetAlert', error);
  }
}

// ==================== PDF EXPORT ====================

function exportDashboardToPDF(htmlContent, reportTitle) {
  try {
    htmlContent  = sanitizeHtmlForPDF(String(htmlContent  || '<html><body></body></html>'));
    reportTitle  = reportTitle || 'เธฃเธฒเธขเธเธฒเธเธชเธฃเธธเธเธเธเธเธฃเธฐเธกเธฒเธ“';
    const date   = new Date().toLocaleDateString('th-TH',{year:'numeric',month:'2-digit',day:'2-digit'}).replace(/\//g,'-');
    const safe   = reportTitle.replace(/[^a-z0-9\u0E00-\u0E7F]/gi,'_').substring(0,50);
    const fname  = `${safe}_${date}_${Date.now()}.pdf`;
    const tmp    = DriveApp.createFile(Utilities.newBlob(htmlContent,'text/html').setName(`Tmp_${Date.now()}.html`));
    const doc    = tmp.makeCopy('Tmp_Convert', DriveApp.getRootFolder());
    tmp.setTrashed(true);
    const pdf    = DriveApp.getRootFolder().createFile(doc.getAs(MimeType.PDF).setName(fname));
    doc.setTrashed(true);
    return pdf.getUrl();
  } catch (e) {
    handleError('exportDashboardToPDF', e, { reportTitle });
    throw new Error('เนเธกเนเธชเธฒเธกเธฒเธฃเธ–เธชเธฃเนเธฒเธเนเธเธฅเน PDF เนเธ”เน: ' + e.toString());
  }
}

function savePdfBase64ToDrive(base64, fileName) {
  try {
    if (!base64) throw new Error('เนเธกเนเธกเธตเธเนเธญเธกเธนเธฅ PDF (base64)');
    const file = DriveApp.getRootFolder().createFile(
      Utilities.newBlob(Utilities.base64Decode(base64), 'application/pdf', fileName)
    );
    return file.getUrl();
  } catch (e) {
    handleError('savePdfBase64ToDrive', e, { fileName });
    throw e;
  }
}

// ==================== TEST FUNCTIONS ====================

function runTests() {
  Logger.log('=== Running Tests ===');
  ['BG69-001','BG69 1','BG69_1','BG691','SP69-005','123',''].forEach(id =>
    Logger.log(`normalizeItemId("${id}") -> "${normalizeItemId(id)}"`)
  );
  [{budget:10000,used:5000,add:3500},{budget:1000.50,used:500.25,add:249.75},{budget:100,used:50,add:60}]
    .forEach(({budget,used,add}) => {
      const r = calculateBudgetSafely(budget,used,add);
      Logger.log(`calc(${budget},${used},${add}) -> valid:${r.valid} used:${r.newUsed} rem:${r.newRemaining}`);
    });
  Logger.log('=== Tests Complete ===');
}

// ==================== NEW: cancelExpense ====================
// เธขเธเน€เธฅเธดเธเธฃเธฒเธขเธเธฒเธฃ: reverse เธขเธญเธ”เนเธ Budget sheet + mark log เธงเนเธฒ CANCELLED
 
function cancelExpense(logRowIndex, reason) {
  const currentUser = getUserPermission();
  if (!currentUser) return createResponse(false, 'เนเธกเนเธเธเธเนเธญเธกเธนเธฅเธเธนเนเนเธเน');
  if (!logRowIndex || typeof logRowIndex !== 'number') return createResponse(false, 'เนเธกเนเธฃเธฐเธเธธ logRowIndex');
  if (!reason || !String(reason).trim()) return createResponse(false, 'เธเธฃเธธเธ“เธฒเธฃเธฐเธเธธเน€เธซเธ•เธธเธเธฅเธเธฒเธฃเธขเธเน€เธฅเธดเธ');
 
  const lock = acquireLockWithRetry();
  if (!lock) return createResponse(false, 'เธฃเธฐเธเธเธเธณเธฅเธฑเธเธเธฃเธฐเธกเธงเธฅเธเธฅ เธเธฃเธธเธ“เธฒเธฅเธญเธเนเธซเธกเน');
 
  try {
    const logSheet   = resolveSheet(CONFIG.SHEETS.TRANSACTION_LOG);
    if (!logSheet) return createResponse(false, 'เนเธกเนเธเธ Transaction Log');
 
    const lastLogCol = logSheet.getLastColumn();
    const headers    = logSheet.getRange(1, 1, 1, lastLogCol).getValues()[0].map(h => (h||'').toString().trim());
    const logCols = getTransactionLogColumnIndices(headers);
 
    const logRow    = logSheet.getRange(logRowIndex, 1, 1, lastLogCol).getValues()[0];
    const logStatus = logCols.status !== -1 ? (logRow[logCols.status] || 'ACTIVE') + '' : 'ACTIVE';
 
    if (logStatus === 'CANCELLED') return createResponse(false, 'เธฃเธฒเธขเธเธฒเธฃเธเธตเนเธ–เธนเธเธขเธเน€เธฅเธดเธเนเธเนเธฅเนเธง');
 
    const logOwner = (logRow[logCols.user] || '') + '';
    const isOwner  = logOwner.toLowerCase() === currentUser.email.toLowerCase();
    if (!isOwner && currentUser.role !== 'admin') {
      return createResponse(false, 'เธเธธเธ“เนเธกเนเธกเธตเธชเธดเธ—เธเธดเนเธขเธเน€เธฅเธดเธเธฃเธฒเธขเธเธฒเธฃเธเธญเธเธเธนเนเธญเธทเนเธ');
    }
 
    const logType = (logRow[logCols.type] || '') + '';
    if (['TRANSFER_OUT','TRANSFER_IN','REVERSAL'].includes(logType)) {
      return createResponse(false, 'เนเธกเนเธชเธฒเธกเธฒเธฃเธ–เธขเธเน€เธฅเธดเธเธฃเธฒเธขเธเธฒเธฃเธเธฃเธฐเน€เธ เธ— ' + logType);
    }
 
    const itemId = normalizeItemId((logRow[logCols.itemId] || '') + '');
    const amount = (typeof logRow[logCols.amount] === 'number') ? logRow[logCols.amount] : parseNumberSafe(logRow[logCols.amount]);
    if (!itemId || !amount) return createResponse(false, 'เธเนเธญเธกเธนเธฅเธฃเธฒเธขเธเธฒเธฃเนเธกเนเธเธฃเธ');
 
    const budgetSheet = resolveSheet(CONFIG.SHEETS.BUDGET);
    if (!budgetSheet) return createResponse(false, 'เนเธกเนเธเธ Sheet เธเธเธเธฃเธฐเธกเธฒเธ“');
    const allData = budgetSheet.getDataRange().getValues();
    const cols    = getColumnIndices(allData[0]);
    if (cols.used === -1 || cols.remaining === -1) return createResponse(false, 'เนเธกเนเธเธเธเธญเธฅเธฑเธกเธเนเธ—เธตเนเธเธณเน€เธเนเธ');
 
    let budgetRow = null;
    for (let i = 1; i < allData.length; i++) {
      if (normalizeItemId((allData[i][cols.itemId]||'')+'').toUpperCase() === itemId.toUpperCase()) { budgetRow = i+1; break; }
    }
    if (!budgetRow) return createResponse(false, 'เนเธกเนเธเธ Item ID: ' + itemId);
 
    const budgetRowData = allData[budgetRow - 1];
    const currentUsed   = parseNumberSafe(budgetRowData[cols.used]);
    const budget        = parseNumberSafe(budgetRowData[cols.budget]);
    const newUsed       = parseFloat((currentUsed - amount).toFixed(2));
    const newRemaining  = parseFloat((budget - newUsed).toFixed(2));
 
    if (newUsed < 0) return createResponse(false, 'เนเธกเนเธชเธฒเธกเธฒเธฃเธ– reverse เนเธ”เน: เธขเธญเธ” used เธเธฐเธ•เธดเธ”เธฅเธ');
 
    if (Math.abs(cols.used - cols.remaining) === 1) {
      const startCol = Math.min(cols.used, cols.remaining) + 1;
      budgetSheet.getRange(budgetRow, startCol, 1, 2).setValues([
        cols.used < cols.remaining ? [newUsed, newRemaining] : [newRemaining, newUsed]
      ]);
    } else {
      budgetSheet.getRange(budgetRow, cols.used      + 1).setValue(newUsed);
      budgetSheet.getRange(budgetRow, cols.remaining + 1).setValue(newRemaining);
    }
 
    if (logCols.status   !== -1) logSheet.getRange(logRowIndex, logCols.status   + 1).setValue('CANCELLED');
    if (logCols.editedBy !== -1) logSheet.getRange(logRowIndex, logCols.editedBy + 1).setValue(currentUser.email);
 
  const note = `[CANCEL] ${reason} (ยกเลิกรายการ row ${logRowIndex})`;
  logTransaction(
    itemId,
    -amount,
    note,
    new Date(),
    newUsed,
    newRemaining,
    'CANCEL',
    0,
    'ACTIVE',
    currentUser.email
  );
 
    return createResponse(true, 'เธขเธเน€เธฅเธดเธเธฃเธฒเธขเธเธฒเธฃเธชเธณเน€เธฃเนเธ', {
      itemId, reversedAmount: amount, newUsed, newRemaining,
      cancelledBy: currentUser.email, timestamp: new Date().toISOString()
    });
 
  } catch (err) {
    handleError('cancelExpense', err, { logRowIndex });
    return createResponse(false, 'เน€เธเธดเธ”เธเนเธญเธเธดเธ”เธเธฅเธฒเธ”: ' + err.toString());
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

// ==================== NEW: editExpense ====================
// เนเธเนเนเธเธฃเธฒเธขเธเธฒเธฃ: reverse เธขเธญเธ”เน€เธเนเธฒ เนเธฅเนเธงเธเธฑเธเธ—เธถเธเธขเธญเธ”เนเธซเธกเน
 
function editExpense(logRowIndex, newAmount, newDescription, newExpenseDate) {
  const currentUser = getUserPermission();
  if (!currentUser) return createResponse(false, 'เนเธกเนเธเธเธเนเธญเธกเธนเธฅเธเธนเนเนเธเน');
  if (!logRowIndex || typeof logRowIndex !== 'number') return createResponse(false, 'เนเธกเนเธฃเธฐเธเธธ logRowIndex');
 
  const amt = parseNumberSafe(newAmount);
  if (isNaN(amt) || amt <= 0) return createResponse(false, 'เธเธณเธเธงเธเน€เธเธดเธเธ•เนเธญเธเธกเธฒเธเธเธงเนเธฒ 0');
 
  const lock = acquireLockWithRetry();
  if (!lock) return createResponse(false, 'เธฃเธฐเธเธเธเธณเธฅเธฑเธเธเธฃเธฐเธกเธงเธฅเธเธฅ เธเธฃเธธเธ“เธฒเธฅเธญเธเนเธซเธกเน');
 
  try {
    const logSheet   = resolveSheet(CONFIG.SHEETS.TRANSACTION_LOG);
    if (!logSheet) return createResponse(false, 'เนเธกเนเธเธ Transaction Log');
 
    const lastLogCol = logSheet.getLastColumn();
    const headers    = logSheet.getRange(1, 1, 1, lastLogCol).getValues()[0].map(h => (h||'').toString().trim());
    const logCols = getTransactionLogColumnIndices(headers);
 
    const logRow    = logSheet.getRange(logRowIndex, 1, 1, lastLogCol).getValues()[0];
    const logStatus = logCols.status !== -1 ? (logRow[logCols.status] || 'ACTIVE') + '' : 'ACTIVE';
 
    if (logStatus === 'CANCELLED') return createResponse(false, 'เนเธกเนเธชเธฒเธกเธฒเธฃเธ–เนเธเนเนเธเธฃเธฒเธขเธเธฒเธฃเธ—เธตเนเธขเธเน€เธฅเธดเธเนเธฅเนเธง');
 
    const logOwner = (logRow[logCols.user] || '') + '';
    const isOwner  = logOwner.toLowerCase() === currentUser.email.toLowerCase();
    if (!isOwner && currentUser.role !== 'admin') {
      return createResponse(false, 'เธเธธเธ“เนเธกเนเธกเธตเธชเธดเธ—เธเธดเนเนเธเนเนเธเธฃเธฒเธขเธเธฒเธฃเธเธญเธเธเธนเนเธญเธทเนเธ');
    }
 
    const logType = (logRow[logCols.type] || '') + '';
    if (['TRANSFER_OUT','TRANSFER_IN','REVERSAL'].includes(logType)) {
      return createResponse(false, 'เนเธกเนเธชเธฒเธกเธฒเธฃเธ–เนเธเนเนเธเธฃเธฒเธขเธเธฒเธฃเธเธฃเธฐเน€เธ เธ— ' + logType);
    }
 
    const itemId    = normalizeItemId((logRow[logCols.itemId] || '') + '');
    const oldAmount = (typeof logRow[logCols.amount] === 'number') ? logRow[logCols.amount] : parseNumberSafe(logRow[logCols.amount]);
    const diff      = amt - oldAmount;
    if (!itemId) return createResponse(false, 'เนเธกเนเธเธ Item ID เนเธเธฃเธฒเธขเธเธฒเธฃเธเธตเน');
 
    const budgetSheet = resolveSheet(CONFIG.SHEETS.BUDGET);
    if (!budgetSheet) return createResponse(false, 'เนเธกเนเธเธ Sheet เธเธเธเธฃเธฐเธกเธฒเธ“');
    const allData = budgetSheet.getDataRange().getValues();
    const cols    = getColumnIndices(allData[0]);
    if (cols.used === -1 || cols.remaining === -1) return createResponse(false, 'เนเธกเนเธเธเธเธญเธฅเธฑเธกเธเนเธ—เธตเนเธเธณเน€เธเนเธ');
 
    let budgetRow = null;
    for (let i = 1; i < allData.length; i++) {
      if (normalizeItemId((allData[i][cols.itemId]||'')+'').toUpperCase() === itemId.toUpperCase()) { budgetRow = i+1; break; }
    }
    if (!budgetRow) return createResponse(false, 'เนเธกเนเธเธ Item ID: ' + itemId);
 
    const budgetRowData = allData[budgetRow - 1];
    const currentUsed   = parseNumberSafe(budgetRowData[cols.used]);
    const budget        = parseNumberSafe(budgetRowData[cols.budget]);
    const newUsed       = parseFloat((currentUsed + diff).toFixed(2));
    const newRemaining  = parseFloat((budget - newUsed).toFixed(2));
 
    if (newUsed < 0)      return createResponse(false, 'เธขเธญเธ”เธ—เธตเนเนเธเนเนเธเธ—เธณเนเธซเนเธขเธญเธ” used เธ•เธดเธ”เธฅเธ');
    if (newRemaining < 0) return createResponse(false, 'เธขเธญเธ”เน€เธเธดเธเธเนเธฒเธขเน€เธเธดเธเธเธเธเธฃเธฐเธกเธฒเธ“เธ—เธตเนเธ•เธฑเนเธเนเธงเน');
 
    if (Math.abs(cols.used - cols.remaining) === 1) {
      const startCol = Math.min(cols.used, cols.remaining) + 1;
      budgetSheet.getRange(budgetRow, startCol, 1, 2).setValues([
        cols.used < cols.remaining ? [newUsed, newRemaining] : [newRemaining, newUsed]
      ]);
    } else {
      budgetSheet.getRange(budgetRow, cols.used      + 1).setValue(newUsed);
      budgetSheet.getRange(budgetRow, cols.remaining + 1).setValue(newRemaining);
    }
 
    if (logCols.status   !== -1) logSheet.getRange(logRowIndex, logCols.status   + 1).setValue('EDITED');
    if (logCols.editedBy !== -1) logSheet.getRange(logRowIndex, logCols.editedBy + 1).setValue(currentUser.email);
 
    let parsedDate = null;
    if (newExpenseDate) {
      try {
        parsedDate = (newExpenseDate instanceof Date) ? newExpenseDate : new Date(newExpenseDate);
        if (isNaN(parsedDate.getTime())) parsedDate = null;
      } catch (e) { parsedDate = null; }
    }
 
    const originalDesc = (logRow[5] || '') + '';
    const finalDesc = newDescription || originalDesc;
    const note = `[EDIT] row ${logRowIndex} | เดิม ${oldAmount} -> ใหม่ ${amt} | ${finalDesc}`;

    logTransaction(
      itemId,
      amt,
      note,
      parsedDate || logRow[1] || new Date(),
      newUsed,
      newRemaining,
      'EDIT',
      0,
      'ACTIVE',
      currentUser.email
    );
 
    return createResponse(true, 'เนเธเนเนเธเธฃเธฒเธขเธเธฒเธฃเธชเธณเน€เธฃเนเธ', {
      itemId, oldAmount, newAmount: amt, diff,
      newUsed, newRemaining,
      editedBy: currentUser.email, timestamp: new Date().toISOString()
    });
 
  } catch (err) {
    handleError('editExpense', err, { logRowIndex, newAmount });
    return createResponse(false, 'เน€เธเธดเธ”เธเนเธญเธเธดเธ”เธเธฅเธฒเธ”: ' + err.toString());
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}



