// transaction_service.gs
// Expense, transfer, and transaction log service functions.

function validateExpenseInput(itemId, amount, expenseDate) {
  const errors = [];
  if (!itemId || !String(itemId).trim()) errors.push('กรุณาระบุรหัสรายการ (Item ID)');
  const amt = parseFloat(String(amount || '').replace(/[^\d\.\-]/g, '')) || 0;
  if (isNaN(amt) || amt <= 0) errors.push('จำนวนเงินต้องมากกว่า 0');
  if (expenseDate) {
    const dt = (expenseDate instanceof Date) ? expenseDate : new Date(expenseDate);
    if (isNaN(dt.getTime())) errors.push('รูปแบบวันที่ไม่ถูกต้อง');
  }
  return { valid: errors.length === 0, errors, sanitizedAmount: amt };
}

function recordExpense(itemId, amount, description, expenseDate) {
  const startTime  = new Date();
  const currentUser = getUserPermission();
  if (!currentUser || currentUser.role === 'viewer') {
    return createResponse(false, 'ไม่มีสิทธิ์บันทึกรายการ');
  }

  const validation = validateExpenseInput(itemId, amount, expenseDate);
  if (!validation.valid) return createResponse(false, 'ข้อผิดพลาด: ' + validation.errors.join(', '));

  const amt = validation.sanitizedAmount;
  const normalizedItemId = normalizeItemId(String(itemId).trim());
  if (!normalizedItemId) return createResponse(false, 'ไม่สามารถ normalize Item ID ได้');

  const parsedDate = normalizeDateInput(expenseDate);

  const lock = acquireLockWithRetry();
  if (!lock) return createResponse(false, 'ระบบกำลังปรับปรุงข้อมูล กรุณาลองใหม่อีกครั้ง');

  try {
    const rowIndex = findRowIndexByItemId(normalizedItemId);
    if (typeof rowIndex !== 'number' || rowIndex <= 1) {
      return createResponse(false, 'ไม่พบ Item ID: ' + normalizedItemId);
    }

    const budgetSheet = resolveSheet(CONFIG.SHEETS.BUDGET);
    const lastCol     = budgetSheet.getLastColumn();
    const headers     = budgetSheet.getRange(1, 1, 1, lastCol).getValues()[0];
    const cols        = getColumnIndices(headers);
    if (!cols || cols.itemId === undefined || cols.itemId === -1) {
      return createResponse(false, 'ไม่สามารถหาคอลัมน์ itemId ได้');
    }

    const values    = budgetSheet.getRange(rowIndex, 1, 1, lastCol).getValues()[0];
    const sheetId   = (values[cols.itemId]    || '').toString().trim();
    const deptInRow = (values[cols.department] || '').toString().trim();

    if (sheetId.toUpperCase() !== normalizedItemId.toUpperCase()) {
      return createResponse(false, 'ข้อมูลไม่สอดคล้องกัน: รหัสรายการไม่ตรงกับแถวที่ค้นพบ');
    }
    if (!hasAccessToRow(currentUser, deptInRow)) {
      return createResponse(false, 'คุณไม่มีสิทธิ์เบิกจ่ายจากแผนกนี้');
    }

    const currentUsed = parseNumberSafe(values[cols.used]   || 0);
    const budget      = parseNumberSafe(values[cols.budget] || 0);
    const calc = calculateBudgetSafely(budget, currentUsed, amt);
    if (!calc.valid) return createResponse(false, 'ยอดเบิกจ่ายเกินงบประมาณที่ตั้งไว้');

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
      return createResponse(false, 'เกิดข้อผิดพลาดในการบันทึกข้อมูล');
    }

    try {
      logTransaction(normalizedItemId, amt, description || '', parsedDate || '', newUsed, newRemaining);
    } catch (logErr) {
      handleError('recordExpense - transaction log', logErr);
    }

    Logger.log('recordExpense completed in %sms', new Date() - startTime);
    return createResponse(true, 'บันทึกสำเร็จ', {
      newUsed, newRemaining, timestamp: new Date().toISOString()
    });

  } catch (err) {
    handleError('recordExpense', err, { itemId: normalizedItemId, amount: amt });
    return createResponse(false, 'เกิดข้อผิดพลาด: ' + err.toString());
  } finally {
    try { if (lock) lock.releaseLock(); } catch (e) {}
  }
}

function transferBudget(fromItemId, toItemId, amount, reason) {
  const currentUser = getUserPermission();
  if (!currentUser) {
    return createResponse(false, 'ไม่พบข้อมูลผู้ใช้ในระบบ');
  }
  if (currentUser.role === 'viewer') {
    return createResponse(false, 'ผู้ใช้ประเภท Viewer ไม่สามารถโอนงบได้');
  }

  const errors   = [];
  const normFrom = normalizeItemId(String(fromItemId || '').trim());
  const normTo   = normalizeItemId(String(toItemId   || '').trim());
  const amt      = parseNumberSafe(amount);

  if (!normFrom)              errors.push('ไม่ระบุ Item ID ต้นทาง');
  if (!normTo)                errors.push('ไม่ระบุ Item ID ปลายทาง');
  if (normFrom === normTo)    errors.push('ต้นทางและปลายทางต้องไม่ใช่รายการเดียวกัน');
  if (isNaN(amt) || amt <= 0) errors.push('จำนวนเงินต้องมากกว่า 0');
  if (!reason || !String(reason).trim()) errors.push('กรุณาระบุเหตุผลการโอนงบ');
  if (errors.length) return createResponse(false, 'ข้อผิดพลาด: ' + errors.join(', '));

  const lock = acquireLockWithRetry();
  if (!lock) return createResponse(false, 'ระบบกำลังประมวลผล กรุณาลองใหม่');

  try {
    const budgetSheet = resolveSheet(CONFIG.SHEETS.BUDGET);
    if (!budgetSheet) return createResponse(false, 'ไม่พบ Sheet งบประมาณ');

    const allData = budgetSheet.getDataRange().getValues();
    const cols    = getColumnIndices(allData[0]);

    if (cols.budget === -1 || cols.used === -1 || cols.remaining === -1) {
      return createResponse(false, 'ไม่พบคอลัมน์ที่จำเป็น');
    }

    const foundRows = findBudgetRowIndicesByItemIds([normFrom, normTo], allData, cols);
    const fromRow = foundRows[normFrom.toUpperCase()];
    const toRow = foundRows[normTo.toUpperCase()];

    if (!fromRow) return createResponse(false, 'ไม่พบ Item ID ต้นทาง: '  + normFrom);
    if (!toRow)   return createResponse(false, 'ไม่พบ Item ID ปลายทาง: ' + normTo);

    const fromValues = allData[fromRow - 1];
    const toValues   = allData[toRow   - 1];

    const fromDept = (fromValues[cols.department] || '').toString().trim();
    const toDept   = (toValues[cols.department]   || '').toString().trim();

    if (!hasAccessToRow(currentUser, fromDept)) {
      return createResponse(false, 'คุณไม่มีสิทธิ์โอนงบจากรายการต้นทาง (หน่วยงาน: ' + fromDept + ')');
    }
    if (!hasAccessToRow(currentUser, toDept)) {
      return createResponse(false, 'คุณไม่มีสิทธิ์โอนงบไปยังรายการปลายทาง (หน่วยงาน: ' + toDept + ')');
    }

    const fromBudget    = parseNumberSafe(fromValues[cols.budget]);
    const fromUsed      = parseNumberSafe(fromValues[cols.used]);
    const fromRemaining = parseNumberSafe(fromValues[cols.remaining]);
    const toBudget      = parseNumberSafe(toValues[cols.budget]);
    const toUsed        = parseNumberSafe(toValues[cols.used]);

    if (amt > fromRemaining) {
      return createResponse(false,
        `งบคงเหลือของต้นทางไม่เพียงพอ (คงเหลือ: ${fromRemaining.toLocaleString('th-TH')} บาท)`);
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

    const note = `[TRANSFER] ${reason} (${normFrom} → ${normTo})`;
    logTransaction(normFrom, -amt, note, new Date(), fromUsed, newFromRemaining, 'TRANSFER_OUT');
    logTransaction(normTo,    amt, note, new Date(), toUsed,   newToRemaining,   'TRANSFER_IN');

    return createResponse(true,
      `โอนงบสำเร็จ: ${normFrom} → ${normTo} จำนวน ${amt.toLocaleString('th-TH')} บาท`, {
        from: { itemId: normFrom, newBudget: newFromBudget, newRemaining: newFromRemaining },
        to:   { itemId: normTo,   newBudget: newToBudget,   newRemaining: newToRemaining },
        amount: amt, transferredBy: currentUser.email, timestamp: new Date().toISOString()
      });

  } catch (err) {
    handleError('transferBudget', err, { fromItemId: normFrom, toItemId: normTo, amount: amt });
    return createResponse(false, 'เกิดข้อผิดพลาด: ' + err.toString());
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

function logTransaction(itemId, amount, description, expenseDate, newUsed, newRemaining, type, quantity, status, editedBy) {
  try {
    const logSheet = ensureTransactionLogSheet();
    logSheet.appendRow([
      new Date(),
      normalizeDateInput(expenseDate),
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
    const currentUser = getUserPermission();
    if (!currentUser) return [];
    const inputId = normalizeItemId(itemId || '');
    if (!inputId) return [];

    // Resolve department for the item to check access
    const isSupportItem = String(itemId).toUpperCase().indexOf('SP') === 0;
    if (isSupportItem) {
      const sRow = findRowIndexInSheetSupport(inputId);
      if (sRow) {
        const supportSheet = resolveSheet(SUPPORT_SHEET_NAME);
        const sHeaders = supportSheet.getRange(1, 1, 1, supportSheet.getLastColumn()).getValues()[0];
        const sMap = mapSupportColumns(sHeaders);
        const sVals = supportSheet.getRange(sRow, 1, 1, supportSheet.getLastColumn()).getValues()[0];
        const dept = _isValidIndex(sMap.department) ? String(sVals[sMap.department] || '').trim() : '';
        if (dept && !hasAccessToRow(currentUser, dept)) return [];
      }
    } else {
      const budgetSheet = resolveSheet(CONFIG.SHEETS.BUDGET);
      if (budgetSheet) {
        const allData = budgetSheet.getDataRange().getValues();
        const cols = getColumnIndices(allData[0]);
        const foundRows = findBudgetRowIndicesByItemIds([inputId], allData, cols);
        const budgetRow = foundRows[inputId.toUpperCase()];
        if (budgetRow !== null && budgetRow !== undefined) {
          const rowData = allData[budgetRow - 1];
          const dept = cols.department !== -1 ? String(rowData[cols.department] || '').trim() : '';
          if (dept && !hasAccessToRow(currentUser, dept)) return [];
        }
      }
    }

    const logContext = getTransactionLogContext();
    if (!logContext) return [];
    const data = logContext.sheet.getDataRange().getValues();
    if (!data || data.length < 2) return [];
    const logCols = logContext.logCols;

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
      const entry = getTransactionLogRowModel(data[i], logCols, i + 1);
      if (!entry.itemId || entry.itemId.toUpperCase() !== inputId.toUpperCase()) continue;
      if (entry.type === 'REVERSAL') continue;

      history.push({
        logRowIndex:  entry.logRowIndex,
        timestamp:    formatTs(entry.timestamp),
        expenseDate:  formatDate(entry.expenseDate),
        user:         entry.user,
        amount:       entry.amount,
        description:  entry.description,
        newUsed:      entry.newUsed,
        newRemaining: entry.newRemaining,
        quantity:     entry.quantity,
        type:         entry.type,
        status:       entry.status,
        editedBy:     entry.editedBy
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
  if (!currentUser) return createResponse(false, 'ไม่พบข้อมูลผู้ใช้');
  if (!logRowIndex || typeof logRowIndex !== 'number') return createResponse(false, 'ไม่ระบุ logRowIndex');
  if (!reason || !String(reason).trim()) return createResponse(false, 'กรุณาระบุเหตุผลการยกเลิก');

  const lock = acquireLockWithRetry();
  if (!lock) return createResponse(false, 'ระบบกำลังประมวลผล กรุณาลองใหม่');

  try {
    const logContext = getTransactionLogContext();
    if (!logContext) return createResponse(false, 'ไม่พบ Transaction Log');

    const logSheet = logContext.sheet;
    const logCols = logContext.logCols;
    const logRow = logSheet.getRange(logRowIndex, 1, 1, logContext.lastCol).getValues()[0];
    const logEntry = getTransactionLogRowModel(logRow, logCols, logRowIndex);

    if (logEntry.status === 'CANCELLED') return createResponse(false, 'รายการนี้ถูกยกเลิกไปแล้ว');

    const isOwner  = logEntry.user.toLowerCase() === currentUser.email.toLowerCase();
    if (!isOwner && currentUser.role !== 'admin') {
      return createResponse(false, 'คุณไม่มีสิทธิ์ยกเลิกรายการของผู้อื่น');
    }

    if (['TRANSFER_OUT','TRANSFER_IN','REVERSAL'].includes(logEntry.type)) {
      return createResponse(false, 'ไม่สามารถยกเลิกรายการประเภท ' + logEntry.type);
    }

    const itemId = logEntry.itemId;
    const amount = logEntry.amount;
    if (!itemId || !amount) return createResponse(false, 'ข้อมูลรายการไม่ครบ');

    const isSupportItem = String(itemId).toUpperCase().indexOf('SP') === 0;

    let sheet, row, currentUsed, budget, newUsed, newRemaining;
    if (isSupportItem) {
      const supportSheet = resolveSheet(SUPPORT_SHEET_NAME);
      if (!supportSheet) return createResponse(false, 'ไม่พบ Sheet งบเงินอุดหนุน');
      const sHeaders = supportSheet.getRange(1, 1, 1, supportSheet.getLastColumn()).getValues()[0];
      const sMap = mapSupportColumns(sHeaders);
      const sRow = findRowIndexInSheetSupport(itemId);
      if (!sRow) return createResponse(false, 'ไม่พบ Item ID: ' + itemId);
      const sVals = supportSheet.getRange(sRow, 1, 1, supportSheet.getLastColumn()).getValues()[0];

      // Compute all new values before writing
      const logQty = parseNumberSafe(logEntry.quantity || 0);
      let newUsedQty = null;
      if (_isValidIndex(sMap.usedQty) && logQty > 0) {
        const currentQty = parseNumberSafe(sVals[sMap.usedQty] || 0);
        newUsedQty = parseFloat((currentQty - logQty).toFixed(2));
        if (newUsedQty < 0) return createResponse(false, 'ไม่สามารถ reverse ได้: ยอด quantity จะติดลบ');
      }

      currentUsed = _isValidIndex(sMap.usedMoney) ? parseNumberSafe(sVals[sMap.usedMoney] || 0) : 0;
      budget = _isValidIndex(sMap.budgetMoney) ? parseNumberSafe(sVals[sMap.budgetMoney] || 0) : 0;
      newUsed = parseFloat((currentUsed - amount).toFixed(2));
      newRemaining = parseFloat((budget - newUsed).toFixed(2));
      if (newUsed < 0) return createResponse(false, 'ไม่สามารถ reverse ได้: ยอด used จะติดลบ');

      // All validations passed — write now
      if (_isValidIndex(sMap.usedMoney)) supportSheet.getRange(sRow, sMap.usedMoney + 1).setValue(newUsed);
      if (_isValidIndex(sMap.remainingMoney)) supportSheet.getRange(sRow, sMap.remainingMoney + 1).setValue(newRemaining);
      if (_isValidIndex(sMap.usedQty) && newUsedQty !== null) {
        supportSheet.getRange(sRow, sMap.usedQty + 1).setValue(newUsedQty);
      }
      sheet = supportSheet;
      row = sRow;
    } else {
      const budgetSheet = resolveSheet(CONFIG.SHEETS.BUDGET);
      if (!budgetSheet) return createResponse(false, 'ไม่พบ Sheet งบประมาณ');
      const allData = budgetSheet.getDataRange().getValues();
      const cols    = getColumnIndices(allData[0]);
      if (cols.used === -1 || cols.remaining === -1) return createResponse(false, 'ไม่พบคอลัมน์ที่จำเป็น');
      const foundRows = findBudgetRowIndicesByItemIds([itemId], allData, cols);
      const budgetRow = foundRows[itemId.toUpperCase()];
      if (!budgetRow) return createResponse(false, 'ไม่พบ Item ID: ' + itemId);
      const budgetRowData = allData[budgetRow - 1];
      const cu   = parseNumberSafe(budgetRowData[cols.used]);
      const b    = parseNumberSafe(budgetRowData[cols.budget]);
      newUsed       = parseFloat((cu - amount).toFixed(2));
      newRemaining  = parseFloat((b - newUsed).toFixed(2));
      if (newUsed < 0) return createResponse(false, 'ไม่สามารถ reverse ได้: ยอด used จะติดลบ');
      if (Math.abs(cols.used - cols.remaining) === 1) {
        const startCol = Math.min(cols.used, cols.remaining) + 1;
        budgetSheet.getRange(budgetRow, startCol, 1, 2).setValues([
          cols.used < cols.remaining ? [newUsed, newRemaining] : [newRemaining, newUsed]
        ]);
      } else {
        budgetSheet.getRange(budgetRow, cols.used      + 1).setValue(newUsed);
        budgetSheet.getRange(budgetRow, cols.remaining + 1).setValue(newRemaining);
      }
      sheet = budgetSheet;
      row = budgetRow;
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

    return createResponse(true, 'ยกเลิกรายการสำเร็จ', {
      itemId, reversedAmount: amount, newUsed, newRemaining,
      cancelledBy: currentUser.email, timestamp: new Date().toISOString()
    });

  } catch (err) {
    handleError('cancelExpense', err, { logRowIndex });
    return createResponse(false, 'เกิดข้อผิดพลาด: ' + err.toString());
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

function editExpense(logRowIndex, newAmount, newDescription, newExpenseDate, newQuantity) {
  const currentUser = getUserPermission();
  if (!currentUser) return createResponse(false, 'ไม่พบข้อมูลผู้ใช้');
  if (!logRowIndex || typeof logRowIndex !== 'number') return createResponse(false, 'ไม่ระบุ logRowIndex');

  const amt = parseNumberSafe(newAmount);
  if (isNaN(amt) || amt <= 0) return createResponse(false, 'จำนวนเงินต้องมากกว่า 0');

  const lock = acquireLockWithRetry();
  if (!lock) return createResponse(false, 'ระบบกำลังประมวลผล กรุณาลองใหม่');

  try {
    const logContext = getTransactionLogContext();
    if (!logContext) return createResponse(false, 'ไม่พบ Transaction Log');

    const logSheet = logContext.sheet;
    const logCols = logContext.logCols;
    const logRow = logSheet.getRange(logRowIndex, 1, 1, logContext.lastCol).getValues()[0];
    const logEntry = getTransactionLogRowModel(logRow, logCols, logRowIndex);

    if (logEntry.status === 'CANCELLED') return createResponse(false, 'ไม่สามารถแก้ไขรายการที่ยกเลิกแล้ว');

    const isOwner  = logEntry.user.toLowerCase() === currentUser.email.toLowerCase();
    if (!isOwner && currentUser.role !== 'admin') {
      return createResponse(false, 'คุณไม่มีสิทธิ์แก้ไขรายการของผู้อื่น');
    }

    if (['TRANSFER_OUT','TRANSFER_IN','REVERSAL'].includes(logEntry.type)) {
      return createResponse(false, 'ไม่สามารถแก้ไขรายการประเภท ' + logEntry.type);
    }

    const itemId    = logEntry.itemId;
    const oldAmount = logEntry.amount;
    const diff      = amt - oldAmount;
    if (!itemId) return createResponse(false, 'ไม่พบ Item ID ในรายการนี้');

    const isSupportItem = String(itemId).toUpperCase().indexOf('SP') === 0;

    let newUsed, newRemaining;
    if (isSupportItem) {
      const supportSheet = resolveSheet(SUPPORT_SHEET_NAME);
      if (!supportSheet) return createResponse(false, 'ไม่พบ Sheet งบเงินอุดหนุน');
      const sHeaders = supportSheet.getRange(1, 1, 1, supportSheet.getLastColumn()).getValues()[0];
      const sMap = mapSupportColumns(sHeaders);
      const sRow = findRowIndexInSheetSupport(itemId);
      if (!sRow) return createResponse(false, 'ไม่พบ Item ID: ' + itemId);
      const sVals = supportSheet.getRange(sRow, 1, 1, supportSheet.getLastColumn()).getValues()[0];

      // Compute money diff
      const currentUsed = _isValidIndex(sMap.usedMoney) ? parseNumberSafe(sVals[sMap.usedMoney] || 0) : 0;
      const budget = _isValidIndex(sMap.budgetMoney) ? parseNumberSafe(sVals[sMap.budgetMoney] || 0) : 0;
      newUsed = parseFloat((currentUsed + diff).toFixed(2));
      newRemaining = parseFloat((budget - newUsed).toFixed(2));
      if (newUsed < 0) return createResponse(false, 'ยอดที่แก้ไขทำให้ยอด used ติดลบ');
      if (newRemaining < 0) return createResponse(false, 'ยอดเบิกจ่ายเกินงบประมาณที่ตั้งไว้');

      // Compute quantity diff
      let newUsedQty = null;
      if (_isValidIndex(sMap.usedQty)) {
        const oldQty = parseNumberSafe(logEntry.quantity || 0);
        const newQty = parseNumberSafe(newQuantity || 0);
        const qtyDiff = newQty - oldQty;
        const currentQty = parseNumberSafe(sVals[sMap.usedQty] || 0);
        newUsedQty = parseFloat((currentQty + qtyDiff).toFixed(2));
        if (newUsedQty < 0) return createResponse(false, 'ยอดที่แก้ไขทำให้จำนวนหน่วยติดลบ');
      }

      // All validations passed — write now
      if (_isValidIndex(sMap.usedMoney)) supportSheet.getRange(sRow, sMap.usedMoney + 1).setValue(newUsed);
      if (_isValidIndex(sMap.remainingMoney)) supportSheet.getRange(sRow, sMap.remainingMoney + 1).setValue(newRemaining);
      if (_isValidIndex(sMap.usedQty) && newUsedQty !== null) {
        supportSheet.getRange(sRow, sMap.usedQty + 1).setValue(newUsedQty);
      }
    } else {
      const budgetSheet = resolveSheet(CONFIG.SHEETS.BUDGET);
      if (!budgetSheet) return createResponse(false, 'ไม่พบ Sheet งบประมาณ');
      const allData = budgetSheet.getDataRange().getValues();
      const cols    = getColumnIndices(allData[0]);
      if (cols.used === -1 || cols.remaining === -1) return createResponse(false, 'ไม่พบคอลัมน์ที่จำเป็น');
      const foundRows = findBudgetRowIndicesByItemIds([itemId], allData, cols);
      const budgetRow = foundRows[itemId.toUpperCase()];
      if (!budgetRow) return createResponse(false, 'ไม่พบ Item ID: ' + itemId);
      const budgetRowData = allData[budgetRow - 1];
      const currentUsed   = parseNumberSafe(budgetRowData[cols.used]);
      const budget        = parseNumberSafe(budgetRowData[cols.budget]);
      newUsed       = parseFloat((currentUsed + diff).toFixed(2));
      newRemaining  = parseFloat((budget - newUsed).toFixed(2));
      if (newUsed < 0)      return createResponse(false, 'ยอดที่แก้ไขทำให้ยอด used ติดลบ');
      if (newRemaining < 0) return createResponse(false, 'ยอดเบิกจ่ายเกินงบประมาณที่ตั้งไว้');
      if (Math.abs(cols.used - cols.remaining) === 1) {
        const startCol = Math.min(cols.used, cols.remaining) + 1;
        budgetSheet.getRange(budgetRow, startCol, 1, 2).setValues([
          cols.used < cols.remaining ? [newUsed, newRemaining] : [newRemaining, newUsed]
        ]);
      } else {
        budgetSheet.getRange(budgetRow, cols.used      + 1).setValue(newUsed);
        budgetSheet.getRange(budgetRow, cols.remaining + 1).setValue(newRemaining);
      }
    }

    if (logCols.status   !== -1) logSheet.getRange(logRowIndex, logCols.status   + 1).setValue('EDITED');
    if (logCols.editedBy !== -1) logSheet.getRange(logRowIndex, logCols.editedBy + 1).setValue(currentUser.email);

    const parsedDate = normalizeDateInput(newExpenseDate);

    const originalDesc = (logRow[5] || '') + '';
    const finalDesc = newDescription || originalDesc;
    const note = `[EDIT] row ${logRowIndex} | เดิม ${oldAmount} -> ใหม่ ${amt} | ${finalDesc}`;

    const finalQuantity = isSupportItem ? parseNumberSafe(newQuantity || 0) : 0;

    logTransaction(
      itemId,
      amt,
      note,
      parsedDate || logRow[1] || new Date(),
      newUsed,
      newRemaining,
      'EDIT',
      finalQuantity,
      'ACTIVE',
      currentUser.email
    );

    return createResponse(true, 'แก้ไขรายการสำเร็จ', {
      itemId, oldAmount, newAmount: amt, diff,
      newUsed, newRemaining,
      editedBy: currentUser.email, timestamp: new Date().toISOString()
    });

  } catch (err) {
    handleError('editExpense', err, { logRowIndex, newAmount });
    return createResponse(false, 'เกิดข้อผิดพลาด: ' + err.toString());
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

function reverseTransfer(logRowIndex) {
  const currentUser = getUserPermission();
  if (!currentUser) return createResponse(false, 'ไม่พบข้อมูลผู้ใช้');
  if (!logRowIndex || typeof logRowIndex !== 'number') return createResponse(false, 'ไม่ระบุ logRowIndex');

  const lock = acquireLockWithRetry();
  if (!lock) return createResponse(false, 'ระบบกำลังประมวลผล กรุณาลองใหม่');

  try {
    const logContext = getTransactionLogContext();
    if (!logContext) return createResponse(false, 'ไม่พบ Transaction Log');

    const logSheet = logContext.sheet;
    const logCols = logContext.logCols;
    const logRow = logSheet.getRange(logRowIndex, 1, 1, logContext.lastCol).getValues()[0];
    const logEntry = getTransactionLogRowModel(logRow, logCols, logRowIndex);

    // Reject if already cancelled
    if (logEntry.status === 'CANCELLED') return createResponse(false, 'รายการนี้ถูกยกเลิกไปแล้ว');

    // Owner/admin authorization check
    const isOwner = logEntry.user.toLowerCase() === currentUser.email.toLowerCase();
    if (!isOwner && currentUser.role !== 'admin') {
      return createResponse(false, 'คุณไม่มีสิทธิ์ยกเลิกรายการของผู้อื่น');
    }

    const entryType = (logEntry.type || '').toUpperCase();
    if (entryType !== 'TRANSFER_OUT' && entryType !== 'TRANSFER_IN') {
      return createResponse(false, 'รายการนี้ไม่ใช่รายการโอนงบ');
    }

    const absAmount = Math.abs(logEntry.amount);
    if (absAmount <= 0) return createResponse(false, 'จำนวนเงินไม่ถูกต้อง');

    // Parse description for both item IDs: [TRANSFER] reason (FROM → TO)
    const desc = logEntry.description || '';
    const descMatch = desc.match(/\((.+?)\s*→\s*(.+?)\)/);
    if (!descMatch) return createResponse(false, 'ไม่พบข้อมูลรายการโอนในคำอธิบาย');

    const fromId = normalizeItemId(descMatch[1].trim());
    const toId = normalizeItemId(descMatch[2].trim());
    if (!fromId || !toId) return createResponse(false, 'ไม่พบ Item ID ในรายการ');

    // Check sibling entry is not already cancelled
    const siblingType = entryType === 'TRANSFER_OUT' ? 'TRANSFER_IN' : 'TRANSFER_OUT';
    const allLogData = logSheet.getDataRange().getValues();
    const siblingRows = [];
    for (let i = 1; i < allLogData.length; i++) {
      const row = allLogData[i];
      const rType = (row[logCols.type] || '').toUpperCase();
      if (rType !== siblingType) continue;
      const rDesc = String(row[5] || '');
      if (rDesc !== desc) continue;
      const rStatus = logCols.status !== -1 ? String(row[logCols.status] || '').toUpperCase() : '';
      if (rStatus === 'CANCELLED') return createResponse(false, 'รายการคู่โอนถูกยกเลิกไปแล้ว');
      siblingRows.push(i + 1);
    }

    const budgetSheet = resolveSheet(CONFIG.SHEETS.BUDGET);
    if (!budgetSheet) return createResponse(false, 'ไม่พบ Sheet งบประมาณ');

    const allData = budgetSheet.getDataRange().getValues();
    const cols = getColumnIndices(allData[0]);
    if (cols.budget === -1 || cols.used === -1 || cols.remaining === -1) {
      return createResponse(false, 'ไม่พบคอลัมน์ที่จำเป็น');
    }

    const foundRows = findBudgetRowIndicesByItemIds([fromId, toId], allData, cols);
    const fromRow = foundRows[fromId.toUpperCase()];
    const toRow = foundRows[toId.toUpperCase()];
    if (!fromRow) return createResponse(false, 'ไม่พบ Item ID ต้นทาง: ' + fromId);
    if (!toRow)   return createResponse(false, 'ไม่พบ Item ID ปลายทาง: ' + toId);

    const fromVals = allData[fromRow - 1];
    const toVals   = allData[toRow - 1];

    const fromCurBudget = parseNumberSafe(fromVals[cols.budget]);
    const fromCurUsed   = parseNumberSafe(fromVals[cols.used]);
    const toCurBudget   = parseNumberSafe(toVals[cols.budget]);
    const toCurUsed     = parseNumberSafe(toVals[cols.used]);

    // Reverse: source gets budget back, destination loses budget
    const fromNewBudget = parseFloat((fromCurBudget + absAmount).toFixed(2));
    const fromNewRemaining = parseFloat((fromNewBudget - fromCurUsed).toFixed(2));
    const toNewBudget = parseFloat((toCurBudget - absAmount).toFixed(2));
    const toNewRemaining = parseFloat((toNewBudget - toCurUsed).toFixed(2));

    if (toNewBudget < 0) return createResponse(false, 'งบประมาณปลายทางจะติดลบหลังยกเลิก');
    if (toNewRemaining < 0) return createResponse(false, 'ยอดเบิกจ่ายปลายทางมากกว่างบที่จะคืน');

    function writeBudgetRow(budgetSheet, rowIdx, newBudget, newRemaining, cols) {
      if (Math.abs(cols.budget - cols.remaining) === 1) {
        const startCol = Math.min(cols.budget, cols.remaining) + 1;
        budgetSheet.getRange(rowIdx, startCol, 1, 2).setValues([
          cols.budget < cols.remaining ? [newBudget, newRemaining] : [newRemaining, newBudget]
        ]);
      } else {
        budgetSheet.getRange(rowIdx, cols.budget + 1).setValue(newBudget);
        budgetSheet.getRange(rowIdx, cols.remaining + 1).setValue(newRemaining);
      }
    }

    writeBudgetRow(budgetSheet, fromRow, fromNewBudget, fromNewRemaining, cols);
    writeBudgetRow(budgetSheet, toRow,   toNewBudget,   toNewRemaining,   cols);

    // Mark the clicked entry as CANCELLED
    if (logCols.status !== -1) logSheet.getRange(logRowIndex, logCols.status + 1).setValue('CANCELLED');
    if (logCols.editedBy !== -1) logSheet.getRange(logRowIndex, logCols.editedBy + 1).setValue(currentUser.email);

    // Mark sibling entries as CANCELLED
    siblingRows.forEach(function(sibRow) {
      if (logCols.status !== -1) logSheet.getRange(sibRow, logCols.status + 1).setValue('CANCELLED');
      if (logCols.editedBy !== -1) logSheet.getRange(sibRow, logCols.editedBy + 1).setValue(currentUser.email);
    });

    // Log reversal entry
    const note = '[REVERSE_TRANSFER] ยกเลิกรายการ row ' + logRowIndex
      + ' (' + fromId + ' ↔ ' + toId + ')';
    logTransaction(fromId, 0, note, new Date(), fromCurUsed, fromNewRemaining, 'REVERSAL', 0, 'ACTIVE', currentUser.email);

    return createResponse(true, 'ยกเลิกรายการโอนงบสำเร็จ', {
      from: { itemId: fromId, newBudget: fromNewBudget, newRemaining: fromNewRemaining },
      to:   { itemId: toId,   newBudget: toNewBudget,   newRemaining: toNewRemaining },
      amount: absAmount, reversedBy: currentUser.email, timestamp: new Date().toISOString()
    });

  } catch (err) {
    handleError('reverseTransfer', err, { logRowIndex });
    return createResponse(false, 'เกิดข้อผิดพลาด: ' + err.toString());
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}
