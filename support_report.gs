/**
 * support_report.gs (core)
 *
 * Contains:
 * - _normalizeReportForTemplate(reportType, reportData)
 * - _getReportData (fallback server aggregator stub)
 * - exportSupportReportPdf_inline (sheet-side preferred, fallback to server agg)
 *
 * Ensure SupportReportCaptureTemplate.html exists.
 */

/* Simple dispatcher fallback (keep existing server aggregators if present) */
function _getReportData(reportType, params) {
  params = params || {};
  var year = Number(params.year) || (new Date()).getFullYear();
  var month = params.month ? Number(params.month) : null;
  var fiscal = Number(params.fiscalStartMonth) || 10;

  // Normalize reportType values
  var rt = (reportType || '').toString().trim();

  // If sheet-side aggregator exists, use it for monthly/byItem/quarterly per-item outputs
  if (typeof buildItemsFromAgg === 'function' && (rt === 'byitem' || rt === 'byItem' || rt === 'by_item' || rt === 'monthly' || rt === 'quarterly')) {
    try {
      var items = buildItemsFromAgg(year, fiscal);
      return { success: true, data: { year: year, fiscalStartMonth: fiscal, items: items } };
    } catch (eAgg) {
      try { Logger.log(' _getReportData: buildItemsFromAgg failed: %s', String(eAgg)); } catch(e){}
      // fall through to server aggregator below
    }
  }

  // Fallback: call server-side aggregator functions if available
  if (rt === 'monthly') {
    if (typeof getSupportMonthlyReport === 'function') {
      return getSupportMonthlyReport(year, month, fiscal);
    } else if (typeof getSupportQuarterlyReport === 'function') {
      // If monthly aggregator missing but quarterly returns per-item months/quarters, try it as fallback
      return getSupportQuarterlyReport(year, fiscal);
    } else {
      return { success: false, message: 'Unknown reportType or server aggregator missing: monthly' };
    }
  }

  if (rt === 'quarterly') {
    if (typeof getSupportQuarterlyReport === 'function') return getSupportQuarterlyReport(year, fiscal);
    if (typeof getSupportSummaryByItem === 'function') return getSupportSummaryByItem(year, fiscal);
    return { success: false, message: 'Unknown reportType or server aggregator missing: quarterly' };
  }

  if (rt === 'byitem' || rt === 'byItem') {
    if (typeof getSupportSummaryByItem === 'function') return getSupportSummaryByItem(year, fiscal);
    if (typeof getSupportQuarterlyReport === 'function') return getSupportQuarterlyReport(year, fiscal);
    return { success: false, message: 'Unknown reportType or server aggregator missing: byItem' };
  }

  return { success: false, message: 'Unknown reportType: ' + reportType };
}

/**
 * Ensure report object has safe properties for the template.
 * - Convert items mapping -> array
 * - Ensure months arrays length 12 and numeric coercions
 */
function _normalizeReportForTemplate(reportType, reportData) {
  if (!reportData) return reportData;

  // If items is an object (map), convert to array and preserve map as itemsMap
  if (reportData.items && !Array.isArray(reportData.items) && typeof reportData.items === 'object') {
    try {
      var itemsObj = reportData.items;
      var arr = Object.keys(itemsObj).map(function(k){
        var it = itemsObj[k] || {};
        it.itemId = it.itemId || it.item || it.itemName || k;
        return it;
      });
      reportData.itemsMap = reportData.itemsMap || itemsObj;
      reportData.items = arr;
    } catch (e) {
      reportData.items = Array.isArray(reportData.items) ? reportData.items : [];
      reportData.itemsMap = reportData.itemsMap || {};
    }
  }

  // Ensure items is array and normalize fields
  if (Array.isArray(reportData.items)) {
    reportData.items = reportData.items.map(function(it, idx){
      it = it || {};
      it.itemId = it.itemId || it.item || it.itemName || ('item_' + idx);
      it.item = it.item || it.itemName || '';
      it.itemName = it.itemName || it.item || '';
      it.work = it.work || it.area || '';
      it.area = it.area || it.work || '';
      it.budgetType = it.budgetType || it.budgetCategory || '';
      it.budgetCategory = it.budgetCategory || it.budgetType || '';
      it.budget = Number(it.budget || 0) || 0;
      it.used = Number(it.used || 0) || 0;
      it.allocatedQuantity = Number(it.allocatedQuantity != null ? it.allocatedQuantity : (it.allocatedQty != null ? it.allocatedQty : (it.allocated != null ? it.allocated : 0))) || 0;
      it.allocatedQty = it.allocatedQuantity;
      it.usedQuantity = Number(it.usedQuantity != null ? it.usedQuantity : (it.qtyUsed != null ? it.qtyUsed : (it.quantityUsed != null ? it.quantityUsed : 0))) || 0;
      it.qtyUsed = it.usedQuantity;
      it.quantityUsed = it.usedQuantity;
      it.remaining = Number(it.remaining || 0) || 0;
      // months: ensure length 12
      if (Array.isArray(it.months)) {
        var months = it.months.slice(0,12);
        while (months.length < 12) months.push(0);
        it.months = months.map(function(v){ return Number(v || 0) || 0; });
      } else {
        // fallback: convert quarters to months evenly, or zero
        if (it.quarters && typeof it.quarters === 'object') {
          var q = it.quarters || {};
          var qvals = [Number(q.Q1||0)||0, Number(q.Q2||0)||0, Number(q.Q3||0)||0, Number(q.Q4||0)||0];
          var monthsArr = [];
          qvals.forEach(function(qv){
            var part = Number(qv || 0) / 3;
            monthsArr.push(part); monthsArr.push(part); monthsArr.push(part);
          });
          it.months = monthsArr.slice(0,12);
        } else {
          it.months = Array.apply(null, Array(12)).map(Number.prototype.valueOf,0);
        }
      }

      // normalize transactions array
      if (Array.isArray(it.transactions)) {
        it.transactions = it.transactions.map(function(t){
          return {
            timestamp: t && t.timestamp ? t.timestamp : '',
            expenseDate: (t && (t.expenseDate || t.timestamp)) || '',
            amount: Number((t && (t.amount || t.newUsed)) || 0) || 0,
            quantity: Number((t && t.quantity) || 0) || 0,
            user: (t && t.user) || '',
            description: (t && t.description) || ''
          };
        });
      } else {
        it.transactions = [];
      }
      return it;
    });
  } else {
    reportData.items = [];
  }

  // Ensure byMonth shape if present
  if (reportType === 'monthly') {
    reportData.byMonth = reportData.byMonth || {};
    Object.keys(reportData.byMonth).forEach(function(k){
      var m = reportData.byMonth[k] || {};
      m.totalAmount = Number(m.totalAmount || 0) || 0;
      m.totalQty = Number(m.totalQty || 0) || 0;
      m.byQuarter = m.byQuarter || { Q1:0, Q2:0, Q3:0, Q4:0 };
      m.byArea = m.byArea || {};
      m.byExpenseType = m.byExpenseType || {};
      reportData.byMonth[k] = m;
    });
  }

  return reportData;
}

/**
 * exportSupportReportPdf_inline
 *
 * - Uses sheet-side aggregation when available (buildItemsFromAgg) for per-item monthly reports.
 * - Falls back to _getReportData when appropriate.
 * - Normalizes report data via _normalizeReportForTemplate before rendering template.
 * - Creates PDF via HtmlService, converts to base64 and returns { success, data: { fileName, base64 } }.
 * - Guards against very large PDF blobs (SIZE_LIMIT).
 */
function exportSupportReportPdf_inline(reportType, params) {
  try {
    params = params || {};
    var year = Number(params.year) || (new Date()).getFullYear();
    var fiscalStartMonth = Number(params.fiscalStartMonth) || 10;
    var includeLogo = !!params.includeLogo;
    var fileName = params.fileName || ('SupportReport_' + (reportType || 'report') + '_' + year + '.pdf');

    Logger.log('exportSupportReportPdf_inline: start reportType=%s year=%s fiscal=%s', reportType, year, fiscalStartMonth);

    // Build reportData (prefer sheet-side)
    var reportData = null;
    try {
      if (typeof buildItemsFromAgg === 'function' && (reportType === 'byItem' || reportType === 'quarterly' || reportType === 'monthly')) {
        var itemsAgg = buildItemsFromAgg(year, fiscalStartMonth);
        Logger.log('exportSupportReportPdf_inline: buildItemsFromAgg returned count=%s', (itemsAgg && itemsAgg.length) ? itemsAgg.length : 0);
        if (itemsAgg && Array.isArray(itemsAgg) && itemsAgg.length > 0) {
          reportData = { year: year, fiscalStartMonth: fiscalStartMonth, items: itemsAgg };
        }
      }
    } catch (eAgg) {
      Logger.log('exportSupportReportPdf_inline: buildItemsFromAgg error: %s', String(eAgg));
      reportData = null;
    }

    // Fallback to server aggregator if no sheet-side data
    if (!reportData) {
      if (typeof _getReportData === 'function') {
        var resp = _getReportData(reportType, { year: year, fiscalStartMonth: fiscalStartMonth, month: params.month });
        Logger.log('exportSupportReportPdf_inline: _getReportData resp=%s', resp && resp.success ? 'OK' : ('ERR:' + (resp && resp.message ? resp.message : 'no resp')));
        if (!resp || !resp.success) {
          var msg = (resp && resp.message) ? resp.message : 'Failed to build report data';
          Logger.log('exportSupportReportPdf_inline: data build failed: %s', msg);
          return { success: false, message: msg };
        }
        reportData = (resp.data || resp);
        // attach agg if items absent
        try {
          if ((!reportData.items || (Array.isArray(reportData.items) && reportData.items.length === 0)) && typeof buildItemsFromAgg === 'function') {
            var itemsAgg2 = buildItemsFromAgg(year, fiscalStartMonth);
            Logger.log('exportSupportReportPdf_inline: attach itemsAgg2 count=%s', (itemsAgg2 && itemsAgg2.length) ? itemsAgg2.length : 0);
            if (itemsAgg2 && Array.isArray(itemsAgg2) && itemsAgg2.length > 0) reportData.items = itemsAgg2;
          }
        } catch (eAttach) {
          Logger.log('exportSupportReportPdf_inline: attach agg failed: %s', String(eAttach));
        }
      } else {
        Logger.log('exportSupportReportPdf_inline: no data source available');
        return { success: false, message: 'No data source available for report' };
      }
    }

    // Log basic info about reportData
    try {
      var cnt = (reportData && reportData.items && Array.isArray(reportData.items)) ? reportData.items.length : 0;
      Logger.log('exportSupportReportPdf_inline: prepared reportData items=%s', cnt);
    } catch (e) { Logger.log('exportSupportReportPdf_inline: error logging reportData count: %s', String(e)); }

    // Normalize for template
    try {
      if (typeof _normalizeReportForTemplate === 'function') {
        reportData = _normalizeReportForTemplate(reportType, reportData) || reportData;
      }
    } catch (eNorm) {
      Logger.log('exportSupportReportPdf_inline: normalization error: %s', String(eNorm));
    }

    // Prepare and render template
    var tpl;
    try {
      tpl = HtmlService.createTemplateFromFile('SupportReportCaptureTemplate');
      tpl.reportType = reportType;
      tpl.params = { year: year, month: params.month || null, fiscalStartMonth: fiscalStartMonth, includeLogo: includeLogo };
      tpl.report = reportData;
      tpl.generatedAt = new Date();
      try { tpl.generatedBy = (typeof getUserEmail === 'function') ? getUserEmail() : Session.getActiveUser().getEmail(); } catch(eUser) { tpl.generatedBy = ''; }
    } catch (eTpl) {
      Logger.log('exportSupportReportPdf_inline: createTemplateFromFile error: %s', String(eTpl));
      return { success: false, message: 'Template creation failed: ' + String(eTpl) };
    }

    var htmlOutput;
    try {
      htmlOutput = tpl.evaluate().setWidth(1200).setHeight(1600);
    } catch (eEval) {
      Logger.log('exportSupportReportPdf_inline: tpl.evaluate() error: %s', String(eEval));
      return { success: false, message: 'Template evaluation failed: ' + String(eEval) };
    }

    var pdfBlob;
    try {
      pdfBlob = htmlOutput.getAs('application/pdf').setName(fileName);
    } catch (ePdf) {
      Logger.log('exportSupportReportPdf_inline: getAs(application/pdf) error: %s', String(ePdf));
      return { success: false, message: 'PDF conversion failed: ' + String(ePdf) };
    }

    // Get bytes and base64
    var bytes = [];
    try {
      bytes = pdfBlob.getBytes();
      Logger.log('exportSupportReportPdf_inline: pdfBlob name=%s size=%s bytes', pdfBlob.getName(), bytes.length);
    } catch (eBytes) {
      Logger.log('exportSupportReportPdf_inline: pdfBlob.getBytes error: %s', String(eBytes));
      return { success: false, message: 'Failed to read PDF bytes: ' + String(eBytes) };
    }

    // size guard
    var SIZE_LIMIT = 6 * 1024 * 1024;
    if (bytes.length > SIZE_LIMIT) {
      Logger.log('exportSupportReportPdf_inline: blob too large (%s bytes) > %s', bytes.length, SIZE_LIMIT);
      return { success: false, message: 'Report too large for inline export (use capture/export to Drive)' };
    }

    var b64;
    try {
      b64 = Utilities.base64Encode(bytes);
      Logger.log('exportSupportReportPdf_inline: base64 length=%s', b64 ? b64.length : 0);
    } catch (eEnc) {
      Logger.log('exportSupportReportPdf_inline: base64Encode error: %s', String(eEnc));
      return { success: false, message: 'Failed to encode PDF: ' + String(eEnc) };
    }

    // Return plain object (avoid transformation by createResponse)
    return { success: true, message: 'OK', data: { fileName: pdfBlob.getName(), base64: b64, size: bytes.length } };

  } catch (err) {
    try { Logger.log('exportSupportReportPdf_inline unexpected error: %s', String(err)); } catch(e){}
    return { success: false, message: String(err) };
  }
}
