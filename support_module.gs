/**
 * support_module.gs
 * Support-side data access and transaction helpers.
 */

const SUPPORT_SHEET_NAME = 'Support';

function normalizeSupportItemModel(raw) {
  raw = raw || {};

  const model = {
    itemId: normalizeItemId(raw.itemId || raw.id || ''),
    rowIndex: Number(raw.rowIndex || 0) || 0,
    department: (raw.department || '').toString().trim(),
    work: (raw.work || raw.area || '').toString().trim(),
    budgetType: (raw.budgetType || raw.budgetCategory || '').toString().trim(),
    expenseType: (raw.expenseType || '').toString().trim(),
    item: (raw.item || raw.itemName || '').toString().trim(),
    allocatedQuantity: Number(
      raw.allocatedQuantity != null
        ? raw.allocatedQuantity
        : (raw.allocatedQty != null ? raw.allocatedQty : 0)
    ) || 0,
    usedQuantity: Number(
      raw.usedQuantity != null
        ? raw.usedQuantity
        : (raw.quantityUsed != null
            ? raw.quantityUsed
            : (raw.qtyUsed != null ? raw.qtyUsed : 0))
    ) || 0,
    budget: Number(raw.budget != null ? raw.budget : (raw.budgetMoney != null ? raw.budgetMoney : 0)) || 0,
    used: Number(raw.used != null ? raw.used : (raw.usedMoney != null ? raw.usedMoney : 0)) || 0,
    remaining: Number(raw.remaining != null ? raw.remaining : (raw.remainingMoney != null ? raw.remainingMoney : 0)) || 0,
    note: (raw.note || '').toString().trim()
  };

  // Backward-compatible aliases for existing client/report code.
  model.area = model.work;
  model.budgetCategory = model.budgetType;
  model.allocatedQty = model.allocatedQuantity;
  model.quantityUsed = model.usedQuantity;
  model.qtyUsed = model.usedQuantity;
  model.itemName = model.item;

  return model;
}

function _isValidIndex(idx) {
  return typeof idx === 'number' && idx >= 0;
}

function ensureSupportSheetExists() {
  const ss = getSpreadsheet();
  let sheet = resolveSheet(SUPPORT_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SUPPORT_SHEET_NAME);
    const headers = [[
      'Item ID',
      'สำนัก/กอง',
      'งบรายจ่าย',
      'ด้าน',
      'ประเภทรายจ่าย',
      'รายการ',
      'จำนวนจัดสรร',
      'จำนวนเบิกจ่าย',
      'งบประมาณจัดสรร',
      'งบประมาณเบิกจ่าย',
      'งบประมาณคงเหลือ',
      'หมายเหตุ'
    ]];
    sheet.getRange(1, 1, 1, headers[0].length).setValues(headers);
    try {
      sheet.getRange('A:A').setNumberFormat('@');
    } catch (e) {}
  }
  return sheet;
}

function mapSupportColumns(headers) {
  headers = headers || [];
  const lower = headers.map(function(h) {
    return (h || '').toString().trim().toLowerCase();
  });

  function findAny(candidates) {
    for (let i = 0; i < candidates.length; i++) {
      const idx = lower.indexOf((candidates[i] || '').toString().trim().toLowerCase());
      if (idx !== -1) return idx;
    }
    return -1;
  }

  return {
    itemId: findAny(['item id', 'itemid', 'id', 'รหัส', 'รหัสรายการ', 'item', 'บาร์โค้ด']),
    department: findAny(['สำนัก/กอง', 'หน่วยงาน', 'department', 'office', 'สำนัก', 'กอง']),
    budgetCategory: findAny(['งบรายจ่าย', 'ประเภทงบ', 'budgetcategory', 'budget type']),
    area: findAny(['ด้าน', 'area', 'work', 'งาน']),
    expenseType: findAny(['ประเภทรายจ่าย', 'expense type', 'ประเภท']),
    item: findAny(['รายการ', 'description', 'item', 'detail']),
    allocatedQty: findAny(['จำนวนจัดสรร', 'จำนวนจัดสรร(หน่วย)', 'allocatedqty', 'qty', 'quantity']),
    usedQty: findAny(['จำนวนเบิกจ่าย', 'จำนวนเบิก', 'usedqty', 'quantity_used']),
    budgetMoney: findAny(['งบประมาณจัดสรร', 'งบประมาณ', 'budget', 'งบจัดสรร']),
    usedMoney: findAny(['งบประมาณเบิกจ่าย', 'usedmoney', 'amount_used', 'เบิกจ่าย']),
    remainingMoney: findAny(['งบประมาณคงเหลือ', 'คงเหลือ', 'remaining', 'balance']),
    note: findAny(['หมายเหตุ', 'note', 'comments'])
  };
}

function getSupportItemsSupport() {
  try {
    const user = getUserPermission();
    if (!user) return createResponse(false, 'ไม่พบข้อมูลผู้ใช้');

    const sheet = resolveSheet(SUPPORT_SHEET_NAME) || ensureSupportSheetExists();
    const data = sheet.getDataRange().getValues();
    if (!data || data.length < 2) {
      return createResponse(true, '', { user: user, items: [] });
    }

    const headers = data[0];
    const map = mapSupportColumns(headers);
    const items = [];

    for (let i = 1; i < data.length; i++) {
      const row = data[i] || [];
      const idRaw = _isValidIndex(map.itemId) ? row[map.itemId] : row[0];
      const dept = _isValidIndex(map.department) ? row[map.department] : (row[1] || '');
      const item = normalizeSupportItemModel({
        itemId: idRaw || '',
        rowIndex: i + 1,
        department: dept || '',
        budgetType: _isValidIndex(map.budgetCategory) ? row[map.budgetCategory] : '',
        work: _isValidIndex(map.area) ? row[map.area] : '',
        expenseType: _isValidIndex(map.expenseType) ? row[map.expenseType] : '',
        item: _isValidIndex(map.item) ? row[map.item] : '',
        allocatedQuantity: _isValidIndex(map.allocatedQty) ? row[map.allocatedQty] : '',
        usedQuantity: _isValidIndex(map.usedQty) ? Number(row[map.usedQty] || 0) : 0,
        budget: _isValidIndex(map.budgetMoney) ? Number(row[map.budgetMoney] || 0) : 0,
        used: _isValidIndex(map.usedMoney) ? Number(row[map.usedMoney] || 0) : 0,
        remaining: _isValidIndex(map.remainingMoney) ? Number(row[map.remainingMoney] || 0) : 0,
        note: _isValidIndex(map.note) ? row[map.note] : ''
      });

      if (
        !item.department ||
        user.role === 'admin' ||
        normalizeAccessValue(item.department) === normalizeAccessValue(user.department)
      ) {
        items.push(item);
      }
    }

    return createResponse(true, '', { user: user, items: items });
  } catch (e) {
    Logger.log('getSupportItemsSupport error: ' + e.toString());
    return createResponse(false, 'เกิดข้อผิดพลาด: ' + e.toString());
  }
}

function findRowIndexInSheetSupport(itemId) {
  try {
    const sheet = resolveSheet(SUPPORT_SHEET_NAME);
    if (!sheet) return null;

    const data = sheet.getDataRange().getValues();
    if (!data || data.length < 2) return null;

    const headers = (data[0] || []).map(function(h) {
      return (h || '').toString();
    });
    const map = mapSupportColumns(headers);
    const idCol = _isValidIndex(map.itemId) ? map.itemId : 0;

    const searchRaw = (itemId || '').toString().trim();
    if (!searchRaw) return null;
    const searchNorm = normalizeItemId(searchRaw).toUpperCase();

    function trailingNumber(value) {
      if (!value) return null;
      const match = String(value).match(/(\d+)\s*$/);
      return match ? parseInt(match[1], 10) : null;
    }

    function prefixBeforeTrailingNumber(value) {
      if (!value) return '';
      return String(value).replace(/[-_\s]*\d+\s*$/, '').trim().toUpperCase();
    }

    const searchTrailing = trailingNumber(searchRaw);
    const searchPrefix = prefixBeforeTrailingNumber(searchRaw);
    const tried = [];

    for (let i = 1; i < data.length; i++) {
      const row = data[i] || [];
      const cellVal = (row[idCol] || '').toString().trim();
      if (!cellVal) {
        tried.push({ row: i + 1, cell: '', match: false });
        continue;
      }

      const cellNorm = normalizeItemId(cellVal).toUpperCase();
      tried.push({ row: i + 1, cell: cellVal, cellNorm: cellNorm });

      if (searchNorm && cellNorm === searchNorm) {
        Logger.log('findRowIndexInSheetSupport: exact normalized match row=%s cell="%s"', i + 1, cellVal);
        return i + 1;
      }

      const cellTrailing = trailingNumber(cellVal);
      if (searchTrailing !== null && cellTrailing !== null) {
        const cellPrefix = prefixBeforeTrailingNumber(cellVal);
        if (searchPrefix) {
          if (cellPrefix === searchPrefix && Number(cellTrailing) === Number(searchTrailing)) {
            Logger.log('findRowIndexInSheetSupport: prefix+number match row=%s cell="%s"', i + 1, cellVal);
            return i + 1;
          }
        } else if (Number(cellTrailing) === Number(searchTrailing)) {
          Logger.log('findRowIndexInSheetSupport: number-only match row=%s cell="%s"', i + 1, cellVal);
          return i + 1;
        }
      }

      if (cellVal.toUpperCase() === searchRaw.toUpperCase()) {
        Logger.log('findRowIndexInSheetSupport: case-insensitive raw match row=%s cell="%s"', i + 1, cellVal);
        return i + 1;
      }
    }

    try {
      const sample = tried.slice(0, 12).map(function(t) {
        return JSON.stringify(t);
      }).join(' | ');
      Logger.log('findRowIndexInSheetSupport: NOT FOUND for "%s" (norm="%s"). Tried sample: %s', searchRaw, searchNorm, sample);
    } catch (e) {
      Logger.log('findRowIndexInSheetSupport: NOT FOUND (logging failed)');
    }

    return null;
  } catch (e) {
    Logger.log('findRowIndexInSheetSupport error: ' + e.toString());
    return null;
  }
}

function recordSupportExpenseSupport(itemId, amount, description, expenseDate, quantity) {
  const lock = LockService.getScriptLock();
  try {
    if (!lock.tryLock(5000)) {
      return createResponse(false, 'ระบบกำลังปรับปรุงข้อมูล กรุณาลองใหม่อีกครั้ง');
    }

    const sheet = resolveSheet(SUPPORT_SHEET_NAME) || ensureSupportSheetExists();
    const lastCol = sheet.getLastColumn();
    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    const map = mapSupportColumns(headers);
    const rowIndex = findRowIndexInSheetSupport(itemId);

    Logger.log('recordSupportExpenseSupport: findRowIndexInSheetSupport returned: %s for itemId="%s"', rowIndex, itemId);

    if (!rowIndex) {
      const normId = normalizeItemId(itemId || '');
      Logger.log('recordSupportExpenseSupport: Unable to locate itemId="%s" (normalized="%s") in Support sheet.', itemId, normId);
      return createResponse(
        false,
        'ไม่พบ Item ID: ' + (itemId || '') + ' กรุณาตรวจสอบรหัสที่แสดงในหน้า Support และค่าในชีต'
      );
    }

    const rowVals = sheet.getRange(rowIndex, 1, 1, lastCol).getValues()[0];
    const addQty = Number(quantity || 0) || 0;
    const addAmt = Number(amount || 0) || 0;

    let newUsedQty = null;
    if (_isValidIndex(map.usedQty)) {
      const currentQty = Number(rowVals[map.usedQty] || 0) || 0;
      newUsedQty = Number((currentQty + addQty).toFixed(2));
      sheet.getRange(rowIndex, map.usedQty + 1).setValue(newUsedQty);
    }

    let newUsedMoney = null;
    let newRemainingMoney = null;
    if (_isValidIndex(map.usedMoney)) {
      const currentUsed = Number(rowVals[map.usedMoney] || 0) || 0;
      const budget = _isValidIndex(map.budgetMoney) ? Number(rowVals[map.budgetMoney] || 0) : 0;
      newUsedMoney = Number((currentUsed + addAmt).toFixed(2));
      sheet.getRange(rowIndex, map.usedMoney + 1).setValue(newUsedMoney);

      if (_isValidIndex(map.budgetMoney) && _isValidIndex(map.remainingMoney)) {
        newRemainingMoney = Number((budget - newUsedMoney).toFixed(2));
        sheet.getRange(rowIndex, map.remainingMoney + 1).setValue(newRemainingMoney);
      }
    }

    try {
      if (typeof logTransaction === 'function') {
        try {
          logTransaction(itemId, addAmt, description, expenseDate, newUsedMoney, newRemainingMoney, 'support', addQty);
        } catch (e) {
          const ss = getSpreadsheet();
          let logSheet = resolveSheet(CONFIG.SHEETS.TRANSACTION_LOG);
          if (!logSheet) logSheet = ss.insertSheet(CONFIG.SHEETS.TRANSACTION_LOG);
          logSheet.appendRow([
            new Date(),
            expenseDate,
            getUserEmail(),
            normalizeItemId(itemId),
            addAmt,
            description,
            newUsedMoney,
            newRemainingMoney,
            addQty,
            'support'
          ]);
        }
      } else {
        const ss2 = getSpreadsheet();
        let logSheet2 = resolveSheet(CONFIG.SHEETS.TRANSACTION_LOG);
        if (!logSheet2) logSheet2 = ss2.insertSheet(CONFIG.SHEETS.TRANSACTION_LOG);
        logSheet2.appendRow([
          new Date(),
          expenseDate,
          getUserEmail(),
          normalizeItemId(itemId),
          addAmt,
          description,
          newUsedMoney,
          newRemainingMoney,
          addQty,
          'support'
        ]);
      }
    } catch (e) {
      Logger.log('recordSupportExpenseSupport log error: ' + e.toString());
    }

    return createResponse(true, 'บันทึกสำเร็จ', {
      newUsedQty: newUsedQty,
      newUsedMoney: newUsedMoney,
      newRemainingMoney: newRemainingMoney
    });
  } catch (e) {
    Logger.log('recordSupportExpenseSupport error: ' + e.toString());
    return createResponse(false, 'เกิดข้อผิดพลาด: ' + e.toString());
  } finally {
    try {
      lock.releaseLock();
    } catch (e) {}
  }
}

function getSupportQuarterlyReport(year, fiscalStartMonth) {
  try {
    year = Number(year) || (new Date()).getFullYear();
    fiscalStartMonth = Number(fiscalStartMonth) || 10;

    const sheet = resolveSheet(SUPPORT_SHEET_NAME);
    if (!sheet) return createResponse(false, 'ไม่พบชีต Support');

    const data = sheet.getDataRange().getValues();
    if (!data || data.length < 2) return createResponse(false, 'ไม่มีข้อมูล Support');

    const headers = data[0];
    const map = mapSupportColumns(headers);
    const itemsMap = {};

    for (let r = 1; r < data.length; r++) {
      const row = data[r] || [];
      const idRaw = _isValidIndex(map.itemId) ? row[map.itemId] : row[0] || '';
      const id = normalizeItemId(idRaw);
      itemsMap[id] = normalizeSupportItemModel({
        itemId: id,
        department: _isValidIndex(map.department) ? (row[map.department] || '') : '',
        work: _isValidIndex(map.area) ? (row[map.area] || '') : '',
        budgetType: _isValidIndex(map.budgetCategory) ? (row[map.budgetCategory] || '') : '',
        expenseType: _isValidIndex(map.expenseType) ? (row[map.expenseType] || '') : '',
        item: _isValidIndex(map.item) ? (row[map.item] || '') : '',
        budget: _isValidIndex(map.budgetMoney) ? Number(row[map.budgetMoney] || 0) : 0,
        used: _isValidIndex(map.usedMoney) ? Number(row[map.usedMoney] || 0) : 0,
        remaining: _isValidIndex(map.remainingMoney) ? Number(row[map.remainingMoney] || 0) : 0,
        quarters: { Q1: 0, Q2: 0, Q3: 0, Q4: 0 }
      });
    }

    const byArea = {};
    const byExpenseType = {};
    const logSheet = resolveSheet(CONFIG.SHEETS.TRANSACTION_LOG);

    if (logSheet) {
      const logData = logSheet.getDataRange().getValues();
      if (logData && logData.length > 1) {
        const logHeaders = logData[0].map(function(h) {
          return (h || '').toString().trim().toLowerCase();
        });
        const colTimestamp = logHeaders.indexOf('timestamp') >= 0 ? logHeaders.indexOf('timestamp') : 0;
        const colItem = logHeaders.indexOf('item id') >= 0 ? logHeaders.indexOf('item id') : 3;
        const colAmount = logHeaders.indexOf('amount') >= 0 ? logHeaders.indexOf('amount') : 4;
        const colType = logHeaders.indexOf('type') >= 0 ? logHeaders.indexOf('type') : (logHeaders.length - 1);

        for (let i = 1; i < logData.length; i++) {
          try {
            const row = logData[i];
            const typeVal = (row[colType] || '').toString().trim().toLowerCase();
            if (typeVal !== 'support') continue;

            const ts = new Date(row[colTimestamp]);
            if (isNaN(ts) || ts.getFullYear() !== year) continue;

            const month = ts.getMonth() + 1;
            const offset = ((month - fiscalStartMonth + 12) % 12);
            const qIndex = Math.floor(offset / 3);
            const qName = 'Q' + (qIndex + 1);
            const itemId = normalizeItemId(row[colItem] || '');
            const amt = Number(row[colAmount] || 0) || 0;

            let meta = itemsMap[itemId];
            if (!meta) {
              meta = normalizeSupportItemModel({
                itemId: itemId,
                department: 'ไม่ระบุ',
                work: 'ไม่ระบุ',
                budgetType: 'ไม่ระบุ',
                expenseType: 'ไม่ระบุ',
                item: '',
                budget: 0,
                used: 0,
                remaining: 0,
                quarters: { Q1: 0, Q2: 0, Q3: 0, Q4: 0 }
              });
              itemsMap[itemId] = meta;
            }

            meta.quarters[qName] = (meta.quarters[qName] || 0) + amt;

            if (!byArea[meta.work]) {
              byArea[meta.work] = { Q1: 0, Q2: 0, Q3: 0, Q4: 0, total: 0 };
            }
            byArea[meta.work][qName] += amt;
            byArea[meta.work].total += amt;

            if (!byExpenseType[meta.expenseType]) {
              byExpenseType[meta.expenseType] = { Q1: 0, Q2: 0, Q3: 0, Q4: 0, total: 0 };
            }
            byExpenseType[meta.expenseType][qName] += amt;
            byExpenseType[meta.expenseType].total += amt;
          } catch (e) {
            Logger.log('support log parse error: ' + e.toString());
          }
        }
      }
    }

    const items = Object.keys(itemsMap).map(function(k) {
      return itemsMap[k];
    });

    return createResponse(true, '', {
      year: year,
      fiscalStartMonth: fiscalStartMonth,
      byArea: byArea,
      byExpenseType: byExpenseType,
      items: items
    });
  } catch (e) {
    Logger.log('getSupportQuarterlyReport error: ' + e.toString());
    return createResponse(false, 'เกิดข้อผิดพลาด: ' + e.toString());
  }
}
