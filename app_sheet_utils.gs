// app_sheet_utils.gs
// Shared spreadsheet, parsing, normalization, and access helpers.

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
  const searchPrefix = prefixBefore(itemId);

  for (let i = 1; i < data.length; i++) {
    const cellVal = (data[i][idColIndex] || '').toString().trim();
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
