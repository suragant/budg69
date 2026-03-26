/**
 * support_module.gs
 * - Safe index checks for mapSupportColumns results (ensure index >= 0 before using)
 * - Minor robustness improvements when interacting with sheets and logging
 */

const SUPPORT_SHEET_NAME = 'Support';

function ensureSupportSheetExists() {
  const ss = getSpreadsheet();
  let sheet = resolveSheet(SUPPORT_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SUPPORT_SHEET_NAME);
    const headers = [
      'Item ID','สำนัก/กอง','งบรายจ่าย','ด้าน','ประเภทรายจ่าย','รายการ',
      'จำนวนจัดสรร','จำนวนเบิกจ่าย','งบประมาณจัดสรร','งบประมาณเบิกจ่าย','งบประมาณคงเหลือ','หมายเหตุ'
    ];
    sheet.getRange(1,1,1,headers.length).setValues([headers]);
    try { sheet.getRange('A:A').setNumberFormat('@'); } catch(e){}
  }
  return sheet;
}

function mapSupportColumns(headers) {
  headers = headers || [];
  const lower = headers.map(h => (h||'').toString().trim().toLowerCase());
  function findAny(list){ for (const s of list) { const idx = lower.indexOf((s||'').toString().trim().toLowerCase()); if (idx !== -1) return idx; } return -1; }
  return {
    itemId: findAny(['item id','itemid','id','รหัส','รหัสรายการ','item','บาร์โค้ด']),
    department: findAny(['สำนัก/กอง','หน่วยงาน','department','office','สำนัก','กอง']),
    budgetCategory: findAny(['งบรายจ่าย','ประเภทงบ','budgetcategory','budget type']),
    area: findAny(['ด้าน','area','work','งาน']),
    expenseType: findAny(['ประเภทรายจ่าย','expense type','ประเภท']),
    item: findAny(['รายการ','description','item','detail']),
    allocatedQty: findAny(['จำนวนจัดสรร','จำนวนจัดสรร(หน่วย)','allocatedqty','qty','quantity']),
    usedQty: findAny(['จำนวนเบิกจ่าย','จำนวนเบิก','usedqty','quantity_used']),
    budgetMoney: findAny(['งบประมาณจัดสรร','งบประมาณ','budget','งบจัดสรร']),
    usedMoney: findAny(['งบประมาณเบิกจ่าย','งบประมาณเบิกจ่าย','usedmoney','amount_used','เบิกจ่าย']),
    remainingMoney: findAny(['งบประมาณคงเหลือ','คงเหลือ','remaining','balance']),
    note: findAny(['หมายเหตุ','note','comments'])
  };
}

/* helper: safe check for index presence */
function _isValidIndex(idx) {
  return (typeof idx === 'number') && idx >= 0;
}

function getSupportItemsSupport() {
  try {
    const user = getUserPermission();
    if (!user) return createResponse(false, 'ไม่พบข้อมูลผู้ใช้');

    const sheet = resolveSheet(SUPPORT_SHEET_NAME) || ensureSupportSheetExists();
    const dataRange = sheet.getDataRange();
    const data = dataRange ? dataRange.getValues() : [];
    if (!data || data.length < 2) return createResponse(true, '', { user, items: [] });

    const headers = data[0];
    const map = mapSupportColumns(headers);

    const items = [];
    for (let i = 1; i < data.length; i++) {
      const row = data[i] || [];
      const idRaw = _isValidIndex(map.itemId) ? row[map.itemId] : row[0];
      const id = normalizeItemId(idRaw || '');
      const dept = _isValidIndex(map.department) ? row[map.department] : (row[1] || '');
      const obj = {
        itemId: id,
        rowIndex: i + 1,
        department: dept || '',
        budgetCategory: _isValidIndex(map.budgetCategory) ? row[map.budgetCategory] : '',
        area: _isValidIndex(map.area) ? row[map.area] : '',
        expenseType: _isValidIndex(map.expenseType) ? row[map.expenseType] : '',
        item: _isValidIndex(map.item) ? row[map.item] : '',
        allocatedQty: _isValidIndex(map.allocatedQty) ? row[map.allocatedQty] : '',
        quantityUsed: _isValidIndex(map.usedQty) ? Number(row[map.usedQty] || 0) : 0,
        budget: _isValidIndex(map.budgetMoney) ? Number(row[map.budgetMoney] || 0) : 0,
        used: _isValidIndex(map.usedMoney) ? Number(row[map.usedMoney] || 0) : 0,
        remaining: _isValidIndex(map.remainingMoney) ? Number(row[map.remainingMoney] || 0) : 0,
        note: _isValidIndex(map.note) ? row[map.note] : ''
      };
      // permission filter: admin sees all, otherwise only own department (normalize spaces/case)
      if (
        !obj.department ||
        user.role === 'admin' ||
        normalizeAccessValue(obj.department) === normalizeAccessValue(user.department)
      ) {
        items.push(obj);
      }
    }

    return createResponse(true, '', { user, items });
  } catch (e) {
    Logger.log('getSupportItemsSupport error: ' + e.toString());
    return createResponse(false, 'เกิดข้อผิดพลาด: ' + e.toString());
  }
}

function findRowIndexInSheetSupport(itemId) {
  try {
    var sheet = resolveSheet(SUPPORT_SHEET_NAME);
    if (!sheet) return null;
    var data = sheet.getDataRange().getValues();
    if (!data || data.length < 2) return null;
    var headers = (data[0] || []).map(function(h){ return (h||'').toString(); });

    // build column map
    var map = mapSupportColumns(headers);
    var idCol = (_isValidIndex(map.itemId) ? map.itemId : 0);

    var searchRaw = (itemId || '').toString().trim();
    if (!searchRaw) return null;
    var searchNorm = normalizeItemId(searchRaw).toUpperCase();

    // helper: trailing number and prefix
    function trailingNumber(s) {
      if (!s) return null;
      var m = String(s).match(/(\d+)\s*$/);
      return m ? parseInt(m[1], 10) : null;
    }
    function prefixBeforeTrailingNumber(s) {
      if (!s) return '';
      return String(s).replace(/[-_\s]*\d+\s*$/, '').trim().toUpperCase();
    }

    var searchTrailing = trailingNumber(searchRaw);
    var searchPrefix = prefixBeforeTrailingNumber(searchRaw);

    // collect diagnostics for logging
    var tried = [];

    for (var i = 1; i < data.length; i++) {
      var row = data[i] || [];
      var cellVal = (row[idCol] || '').toString().trim();
      if (!cellVal) { tried.push({ row: i+1, cell: '', match: false }); continue; }

      var cellNorm = normalizeItemId(cellVal).toUpperCase();
      tried.push({ row: i+1, cell: cellVal, cellNorm: cellNorm });

      // 1) exact normalized match (preferred)
      if (searchNorm && cellNorm && cellNorm === searchNorm) {
        Logger.log('findRowIndexInSheetSupport: exact normalized match row=%s cell="%s"', i+1, cellVal);
        return i + 1;
      }

      // 2) numeric+prefix fallback: if both have trailing numbers, compare numbers and (if search prefix exists) compare prefix
      var cellTrailing = trailingNumber(cellVal);
      if (searchTrailing !== null && cellTrailing !== null) {
        var cellPrefix = prefixBeforeTrailingNumber(cellVal);
        if (searchPrefix) {
          if (cellPrefix === searchPrefix && Number(cellTrailing) === Number(searchTrailing)) {
            Logger.log('findRowIndexInSheetSupport: prefix+number match row=%s cell="%s"', i+1, cellVal);
            return i + 1;
          }
        } else {
          if (Number(cellTrailing) === Number(searchTrailing)) {
            Logger.log('findRowIndexInSheetSupport: number-only match row=%s cell="%s"', i+1, cellVal);
            return i + 1;
          }
        }
      }

      // 3) case-insensitive raw equality (last resort)
      if (cellVal.toUpperCase() === searchRaw.toUpperCase()) {
        Logger.log('findRowIndexInSheetSupport: case-insensitive raw match row=%s cell="%s"', i+1, cellVal);
        return i + 1;
      }
    }

    // not found — log diagnostics (limited sample)
    try {
      var sample = tried.slice(0, 12).map(function(t){ return JSON.stringify(t); }).join(' | ');
      Logger.log('findRowIndexInSheetSupport: NOT FOUND for "%s" (norm="%s"). Tried sample: %s', searchRaw, searchNorm, sample);
    } catch(e){ Logger.log('findRowIndexInSheetSupport: NOT FOUND (logging failed)'); }

    return null;
  } catch (e) {
    Logger.log('findRowIndexInSheetSupport error: ' + e.toString());
    return null;
  }
}


function recordSupportExpenseSupport(itemId, amount, description, expenseDate, quantity) {
  var lock = LockService.getScriptLock();
  var got = false;
  try {
    if (!lock.tryLock(5000)) return createResponse(false, 'ระบบกำลังปรับปรุงข้อมูล กรุณาลองใหม่อีกครั้ง');

    var sheet = resolveSheet(SUPPORT_SHEET_NAME) || ensureSupportSheetExists();
    var lastCol = sheet.getLastColumn();
    var headers = sheet.getRange(1,1,1,lastCol).getValues()[0];
    var map = mapSupportColumns(headers);

    // robust row search
    var rowIndex = findRowIndexInSheetSupport(itemId);
    Logger.log('recordSupportExpenseSupport: findRowIndexInSheetSupport returned: %s for itemId="%s"', rowIndex, itemId);

    if (!rowIndex) {
      var normId = normalizeItemId(itemId || '');
      Logger.log('recordSupportExpenseSupport: Unable to locate itemId="%s" (normalized="%s") in Support sheet.', itemId, normId);
      return createResponse(false, 'ไม่พบ Item ID: ' + (itemId || '') + ' — กรุณาตรวจสอบรหัสที่แสดงในหน้า Support และค่าใน Sheet (ตัวพิมพ์/ช่องว่าง/Prefix/จำนวนหลักอาจต่างกัน)');
    }

    var rowVals = sheet.getRange(rowIndex, 1, 1, lastCol).getValues()[0];
    var addQty = Number(quantity || 0) || 0;
    var addAmt = Number(amount || 0) || 0;

    var newUsedQty = null;
    if (_isValidIndex(map.usedQty)) {
      var curQty = Number(rowVals[map.usedQty] || 0) || 0;
      newUsedQty = Number((curQty + addQty).toFixed(2));
      sheet.getRange(rowIndex, map.usedQty + 1).setValue(newUsedQty);
    }

    var newUsedMoney = null;
    var newRemainingMoney = null;
    if (_isValidIndex(map.usedMoney)) {
      var curUsed = Number(rowVals[map.usedMoney] || 0) || 0;
      var budget = _isValidIndex(map.budgetMoney) ? Number(rowVals[map.budgetMoney] || 0) : 0;
      newUsedMoney = Number((curUsed + addAmt).toFixed(2));
      sheet.getRange(rowIndex, map.usedMoney + 1).setValue(newUsedMoney);
      if (_isValidIndex(map.budgetMoney) && _isValidIndex(map.remainingMoney)) {
        newRemainingMoney = Number((budget - newUsedMoney).toFixed(2));
        sheet.getRange(rowIndex, map.remainingMoney + 1).setValue(newRemainingMoney);
      }
    }

    // append to transaction log (use normalized item id)
    try {
      if (typeof logTransaction === 'function') {
        try {
          logTransaction(itemId, addAmt, description, expenseDate, newUsedMoney, newRemainingMoney, 'support', addQty);
        } catch (e) {
          var ss = getSpreadsheet();
          var logSheet = resolveSheet(CONFIG.SHEETS.TRANSACTION_LOG);
          if (!logSheet) logSheet = ss.insertSheet(CONFIG.SHEETS.TRANSACTION_LOG);
          logSheet.appendRow([ new Date(), expenseDate, getUserEmail(), normalizeItemId(itemId), addAmt, description, newUsedMoney, newRemainingMoney, addQty, 'support' ]);
        }
      } else {
        var ss2 = getSpreadsheet();
        var logSheet2 = resolveSheet(CONFIG.SHEETS.TRANSACTION_LOG);
        if (!logSheet2) logSheet2 = ss2.insertSheet(CONFIG.SHEETS.TRANSACTION_LOG);
        logSheet2.appendRow([ new Date(), expenseDate, getUserEmail(), normalizeItemId(itemId), addAmt, description, newUsedMoney, newRemainingMoney, addQty, 'support' ]);
      }
    } catch(e) {
      Logger.log('recordSupportExpenseSupport log error: ' + e.toString());
    }

    return createResponse(true, 'บันทึกสำเร็จ', { newUsedQty: newUsedQty, newUsedMoney: newUsedMoney, newRemainingMoney: newRemainingMoney });

  } catch (e) {
    Logger.log('recordSupportExpenseSupport error: ' + e.toString());
    return createResponse(false, 'เกิดข้อผิดพลาด: ' + e.toString());
  } finally {
    try { lock.releaseLock(); } catch(e){}
  }
}

/* Quarterly aggregation for support (reads Transaction_Log type='support') */
function getSupportQuarterlyReport(year, fiscalStartMonth) {
  try {
    year = Number(year) || (new Date()).getFullYear();
    fiscalStartMonth = Number(fiscalStartMonth) || 10; // default Oct

    const sheet = resolveSheet(SUPPORT_SHEET_NAME);
    if (!sheet) return createResponse(false, 'ไม่พบ sheet Support');
    const data = sheet.getDataRange().getValues();
    if (!data || data.length < 2) return createResponse(false, 'ไม่มีข้อมูล Support');
    const headers = data[0];
    const map = mapSupportColumns(headers);

    const itemsMap = {};
    for (let r = 1; r < data.length; r++) {
      const row = data[r] || [];
      const idRaw = _isValidIndex(map.itemId) ? row[map.itemId] : row[0] || '';
      const id = normalizeItemId(idRaw);
      itemsMap[id] = {
        itemId: id,
        area: _isValidIndex(map.department) ? (row[map.department] || '') : '',
        expenseType: _isValidIndex(map.expenseType) ? (row[map.expenseType] || '') : '',
        itemName: _isValidIndex(map.item) ? (row[map.item] || '') : '',
        budget: _isValidIndex(map.budgetMoney) ? Number(row[map.budgetMoney] || 0) : 0,
        used: _isValidIndex(map.usedMoney) ? Number(row[map.usedMoney] || 0) : 0,
        remaining: _isValidIndex(map.remainingMoney) ? Number(row[map.remainingMoney] || 0) : 0,
        quarters: { Q1:0, Q2:0, Q3:0, Q4:0 }
      };
    }

    const byArea = {};
    const byExpenseType = {};
    const logSheet = resolveSheet(CONFIG.SHEETS.TRANSACTION_LOG);
    if (logSheet) {
      const logData = logSheet.getDataRange().getValues();
      if (logData && logData.length > 1) {
        const logHeaders = logData[0].map(h => (h||'').toString().trim().toLowerCase());
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
              meta = { itemId: itemId, area: 'ไม่ระบุ', expenseType: 'ไม่ระบุ', itemName: '', budget:0, used:0, remaining:0, quarters:{Q1:0,Q2:0,Q3:0,Q4:0} };
              itemsMap[itemId] = meta;
            }
            meta.quarters[qName] = (meta.quarters[qName] || 0) + amt;

            if (!byArea[meta.area]) byArea[meta.area] = { Q1:0,Q2:0,Q3:0,Q4:0, total:0 };
            byArea[meta.area][qName] += amt; byArea[meta.area].total += amt;

            if (!byExpenseType[meta.expenseType]) byExpenseType[meta.expenseType] = { Q1:0,Q2:0,Q3:0,Q4:0, total:0 };
            byExpenseType[meta.expenseType][qName] += amt; byExpenseType[meta.expenseType].total += amt;
          } catch(e) {
            Logger.log('support log parse error: ' + e.toString());
          }
        }
      }
    }

    const items = Object.keys(itemsMap).map(k => itemsMap[k]);
    return createResponse(true, '', { year: year, fiscalStartMonth: fiscalStartMonth, byArea: byArea, byExpenseType: byExpenseType, items: items });
  } catch (e) {
    Logger.log('getSupportQuarterlyReport error: ' + e.toString());
    return createResponse(false, 'เกิดข้อผิดพลาด: ' + e.toString());
  }
}
