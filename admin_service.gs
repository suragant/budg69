// admin_service.gs
// Admin alerts, export helpers, and maintenance utilities.

function buildBudgetAlertEmailBody(recipientEmail, alerts, dateStr) {
  const critical = alerts.filter(a => a.level === 'critical');
  const high = alerts.filter(a => a.level === 'high');
  const medium = alerts.filter(a => a.level === 'medium');
  const alertsUrl = buildBudgetAppDeepLink({ open: 'alerts' });
  const safeAlertsUrl = escapeHtmlAttribute(alertsUrl);

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
    <h2 style="color:#667eea;">รายงานสถานะงบประมาณประจำวัน (${dateStr})</h2>
    <p>เรียน ${recipientEmail}</p>
    <p>พบรายการแจ้งเตือนทั้งหมด <strong>${alerts.length}</strong> รายการ</p>
    ${safeAlertsUrl ? `
      <div style="margin:16px 0 20px;">
        <a href="${safeAlertsUrl}" target="_blank" style="display:inline-block;background:#667eea;color:#fff;text-decoration:none;padding:10px 16px;border-radius:8px;font-weight:700;">
          เปิดระบบและดูแจ้งเตือน
        </a>
        <div style="margin-top:8px;font-size:12px;color:#666;">หากปุ่มไม่ทำงาน สามารถเปิดลิงก์นี้ได้โดยตรง: ${safeAlertsUrl}</div>
      </div>` : ''}
    ${section(critical, '#ff4444', 'ด่วน', 'เร่งด่วน (หมดงบหรือเหลือน้อยกว่า 5%)')}
    ${section(high, '#FF9800', 'เตือน', 'ควรระวัง (ใช้ 90-94%)')}
    ${section(medium, '#FFC107', 'ติดตาม', 'ติดตาม (ใช้ 80-89%)')}
    <hr><p style="font-size:12px;color:#666;">ส่งจากระบบบันทึกการใช้งบประมาณอัตโนมัติ</p>
  </body></html>`;
}

function buildBudgetAppDeepLink(params) {
  const appUrl = getBudgetAppUrl();
  if (!appUrl) return '';

  const paramKeys = Object.keys(params || {});
  const query = paramKeys
    .filter(key => params[key] !== undefined && params[key] !== null && String(params[key]).trim() !== '')
    .map(key => `${encodeURIComponent(key)}=${encodeURIComponent(String(params[key]).trim())}`)
    .join('&');

  if (!query) return appUrl;

  const hashIndex = appUrl.indexOf('#');
  const fullBase = hashIndex === -1 ? appUrl : appUrl.slice(0, hashIndex);
  const hash = hashIndex === -1 ? '' : appUrl.slice(hashIndex);
  const queryIndex = fullBase.indexOf('?');
  const path = queryIndex === -1 ? fullBase : fullBase.slice(0, queryIndex);
  const existingQuery = queryIndex === -1 ? '' : fullBase.slice(queryIndex + 1);
  const existing = existingQuery
    ? existingQuery.split('&').filter(part => {
        const key = decodeURIComponent(String(part.split('=')[0] || '').replace(/\+/g, ' '));
        return paramKeys.indexOf(key) === -1;
      })
    : [];
  const combinedQuery = existing.concat(query).filter(Boolean).join('&');
  return path + (combinedQuery ? '?' + combinedQuery : '') + hash;
}

function getBudgetAppUrl() {
  try {
    const props = PropertiesService.getScriptProperties();
    const propUrl = String(props.getProperty('WEB_APP_URL') || '').trim();
    if (propUrl) return propUrl;

    const configuredUrl = String(CONFIG.WEB_APP_URL || '').trim();
    if (configuredUrl) return configuredUrl;

    const url = ScriptApp.getService().getUrl();
    return url ? String(url).trim() : '';
  } catch (error) {
    return '';
  }
}

function escapeHtmlAttribute(value) {
  return String(value || '')
    .replace(/&/g, '&amp;')
    .replace(/"/g, '&quot;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
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
      <h2 style="color:#667eea;">รายงานสถานะงบประมาณประจำวัน (${dateStr})</h2>
      <p>เรียน ${recipient}</p>`;
    const section = (arr, color, emoji, label) => {
      if (!arr.length) return '';
      return `<h3 style="color:${color};">${emoji} ${label}</h3><ul>` +
        arr.map(a => `<li><strong>${a.itemId}</strong> - ${a.work} > ${a.item}<br><span style="color:${color}">${a.message}</span></li>`).join('') +
        '</ul>';
    };
    body += section(critical, '#ff4444', 'ด่วน', 'เร่งด่วน (หมดงบหรือเหลือน้อยกว่า 5%)');
    body += section(high,     '#FF9800', 'เตือน', 'ควรระวัง (ใช้ 90-94%)');
    body += section(medium,   '#FFC107', 'ติดตาม', 'ติดตาม (ใช้ 80-89%)');
    body += '<hr><p style="font-size:12px;color:#666;">ส่งจากระบบบันทึกการใช้งบประมาณอัตโนมัติ</p></body></html>';
    MailApp.sendEmail({
      to: recipient,
      subject: `แจ้งเตือนสถานะงบประมาณ - ${dateStr}`,
      htmlBody: body,
      name: 'ระบบบันทึกการใช้งบประมาณ'
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
          subject: `แจ้งเตือนสถานะงบประมาณ - ${dateStr}`,
          htmlBody: buildBudgetAlertEmailBody(email, scopedAlerts, dateStr),
          name: 'ระบบบันทึกการใช้งบประมาณ'
        });
        sent[emailKey] = true;
      });

    if (!Object.keys(sent).length && CONFIG.ADMIN_EMAIL) {
      MailApp.sendEmail({
        to: CONFIG.ADMIN_EMAIL,
        subject: `แจ้งเตือนสถานะงบประมาณ - ${dateStr}`,
        htmlBody: buildBudgetAlertEmailBody(CONFIG.ADMIN_EMAIL, alerts, dateStr),
        name: 'ระบบบันทึกการใช้งบประมาณ'
      });
    }
  } catch (error) {
    handleError('sendDailyBudgetAlert', error);
  }
}

function exportDashboardToPDF(htmlContent, reportTitle) {
  try {
    htmlContent  = sanitizeHtmlForPDF(String(htmlContent  || '<html><body></body></html>'));
    reportTitle  = reportTitle || 'รายงานสรุปงบประมาณ';
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
    throw new Error('ไม่สามารถสร้างไฟล์ PDF ได้: ' + e.toString());
  }
}

function savePdfBase64ToDrive(base64, fileName) {
  try {
    if (!base64) throw new Error('ไม่มีข้อมูล PDF (base64)');
    const file = DriveApp.getRootFolder().createFile(
      Utilities.newBlob(Utilities.base64Decode(base64), 'application/pdf', fileName)
    );
    return file.getUrl();
  } catch (e) {
    handleError('savePdfBase64ToDrive', e, { fileName });
    throw e;
  }
}

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
