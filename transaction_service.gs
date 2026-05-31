// transaction_service.gs
// Expense, transfer, and transaction log service functions.

function recordExpense(itemId, amount, description, expenseDate) {
  const startTime  = new Date();
  const currentUser = getUserPermission();
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
      if (Math.abs(cols.remaining - cols.used) === 1) {
        const startCol = Math.min(cols.used, cols.remaining) + 1;
        budgetSheet.getRange(rowIndex, startCol, 1, 2).setValues([
          cols.used < cols.remaining ? [newUsed, newRemaining] : [newRemaining, newUsed]
        ]);
      } else {
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

function transferBudget(fromItemId, toItemId, amount, reason) {
  const currentUser = getUserPermission();
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

    const allData = budgetSheet.getDataRange().getValues();
    const cols    = getColumnIndices(allData[0]);

    if (cols.budget === -1 || cols.used === -1 || cols.remaining === -1) {
      return createResponse(false, 'เนเธกเนเธเธเธเธญเธฅเธฑเธกเธเนเธ—เธตเนเธเธณเน€เธเนเธ');
    }

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
