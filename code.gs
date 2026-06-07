// Code.gs - Main Backend Script (Improved & Optimized & Secured)
// Last Updated: 2026-03-19

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

// ==================== EXPENSE RECORDING ====================

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



