// budget_service.gs
// Budget reads, dashboard summaries, and alert-related data helpers.

function getBudgetItems() {
  try {
    const user = getUserPermission();
    if (!user) return createResponse(false, `ไม่พบข้อมูลผู้ใช้ในระบบ (Email: ${getUserEmail()})`);
    const budgetSheet = resolveSheet(CONFIG.SHEETS.BUDGET);
    if (!budgetSheet) return createResponse(false, 'ไม่พบ Sheet งบประมาณ');
    const data = budgetSheet.getDataRange().getValues();
    if (!data || data.length < 2) return createResponse(false, 'Sheet งบประมาณไม่มีข้อมูล');
    const cols = getColumnIndices(data[0]);
    if (cols.department === -1 || cols.budget === -1) {
      return createResponse(false, 'ไม่พบคอลัมน์ที่จำเป็น (สำนัก/กอง หรือ งบประมาณ)'); 
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
    return createResponse(false, 'เกิดข้อผิดพลาด: ' + error.toString());
  }
}

function getInitialData() {
  try {
    const user = getUserPermission();
    if (!user) return createResponse(false, `ไม่พบข้อมูลผู้ใช้ในระบบ (Email: ${getUserEmail()})`);
    const budgetSheet = resolveSheet(CONFIG.SHEETS.BUDGET);
    if (!budgetSheet) return createResponse(false, 'ไม่พบ Sheet งบประมาณ');
    const data = budgetSheet.getDataRange().getValues();
    if (!data || data.length < 2) return createResponse(false, 'Sheet งบประมาณไม่มีข้อมูล');
    const cols = getColumnIndices(data[0]);
    if (cols.department === -1 || cols.budget === -1) {
      return createResponse(false, 'ไม่พบคอลัมน์ที่จำเป็น (สำนัก/กอง หรือ งบประมาณ)'); 
    }

    const items  = [];
    const alerts = [];
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
    return createResponse(false, 'เกิดข้อผิดพลาด: ' + error.toString());
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

function getDashboardData() {
  try {
    const user = getUserPermission();
    if (!user) return createResponse(false, 'ไม่พบข้อมูลผู้ใช้');
    const budgetSheet = resolveSheet(CONFIG.SHEETS.BUDGET);
    if (!budgetSheet) return createResponse(false, 'ไม่พบ Sheet งบประมาณ');
    const data = budgetSheet.getDataRange().getValues();
    if (!data || data.length < 2) return createResponse(false, 'Sheet งบประมาณไม่มีข้อมูล');
    const cols = getColumnIndices(data[0]);
    const transferTotals = getTransferTotalsByItemId();
    const workSummary = {};
    for (let i = 1; i < data.length; i++) {
      const row  = data[i];
      const dept = row[cols.department] || '';
      if (!hasAccessToRow(user, dept)) continue;
      const work = row[cols.work] || 'ไม่ระบุ';
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
    return createResponse(false, 'เกิดข้อผิดพลาด: ' + error.toString());
  }
}

function getWorkDetails(workName) {
  try {
    const user = getUserPermission();
    if (!user) return createResponse(false, 'ไม่พบข้อมูลผู้ใช้');
    const budgetSheet = resolveSheet(CONFIG.SHEETS.BUDGET);
    if (!budgetSheet) return createResponse(false, 'ไม่พบ Sheet งบประมาณ');
    const data = budgetSheet.getDataRange().getValues();
    if (!data || data.length < 2) return createResponse(false, 'Sheet งบประมาณไม่มีข้อมูล');
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
        work:        row[cols.work]        || 'ไม่ระบุ',
        budgetType:  (cols.budgetType  !== -1) ? (row[cols.budgetType]  || 'ไม่ระบุ') : 'ไม่ระบุ',
        category:    (cols.category    !== -1) ? (row[cols.category]    || 'ไม่ระบุ') : 'ไม่ระบุ',
        expenseType: (cols.expenseType !== -1) ? (row[cols.expenseType] || 'ไม่ระบุ') : 'ไม่ระบุ',
        item:        (cols.item        !== -1) ? (row[cols.item]        || 'ไม่ระบุ') : 'ไม่ระบุ',
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
    return createResponse(false, 'เกิดข้อผิดพลาด: ' + error.toString());
  }
}

function checkBudgetAlerts() {
  try {
    const user = getUserPermission();
    if (!user) return { success: false, message: 'ไม่พบข้อมูลผู้ใช้', alerts: [] };
    const budgetSheet = resolveSheet(CONFIG.SHEETS.BUDGET);
    if (!budgetSheet) return { success: false, message: 'ไม่พบ Sheet งบประมาณ', alerts: [] };
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
    return { success: false, message: 'เกิดข้อผิดพลาด: ' + (error?.message || error?.toString()), alerts: [] };
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
