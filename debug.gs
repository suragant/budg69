/**
 * Debug helpers - run these from Apps Script Editor -> Run
 * 1) validateSupportAgg()  -> ดู rows ใน Support_Agg
 * 2) testBuildItemsFromAgg() -> เรียก buildItemsFromAgg(...) และดู items ที่คืนมา
 * 3) debugReportData(reportType, year, fiscalStartMonth) -> เรียก _getReportData แล้ว normalize และ log preview
 * 4) debug_exportInline_toDrive() -> เรียก exportSupportReportPdf_inline และเขียนไฟล์ลง Drive (ถ้ามี base64)
 */

function validateSupportAgg() {
  try {
    var ss = SpreadsheetApp.getActive();
    var sh = ss.getSheetByName('Support_Agg');
    if (!sh) {
      Logger.log('validateSupportAgg: Support_Agg sheet NOT FOUND');
      return;
    }
    var rows = sh.getDataRange().getValues();
    Logger.log('validateSupportAgg: rows=%s', rows.length);
    // log first 20 rows for inspection
    var preview = rows.slice(0, 20);
    Logger.log('validateSupportAgg preview: %s', JSON.stringify(preview));
    return preview;
  } catch (e) {
    Logger.log('validateSupportAgg ERROR: %s', String(e));
    throw e;
  }
}

function testBuildItemsFromAgg() {
  try {
    // adjust year/fiscal as you normally call export
    var year = new Date().getFullYear();
    var fiscal = 10;
    Logger.log('testBuildItemsFromAgg: calling buildItemsFromAgg(%s,%s)', year, fiscal);
    if (typeof buildItemsFromAgg !== 'function') {
      Logger.log('testBuildItemsFromAgg: buildItemsFromAgg not found');
      return;
    }
    var items = buildItemsFromAgg(year, fiscal);
    Logger.log('testBuildItemsFromAgg: items count=%s', items && items.length ? items.length : 0);
    if (items && items.length > 0) {
      Logger.log('testBuildItemsFromAgg sample (first 10): %s', JSON.stringify(items.slice(0,10)));
    }
    return items;
  } catch (e) {
    Logger.log('testBuildItemsFromAgg ERROR: %s', String(e));
    throw e;
  }
}

function debugReportData(reportType, year, fiscalStartMonth) {
  try {
    reportType = reportType || 'monthly';
    year = Number(year) || (new Date()).getFullYear();
    fiscalStartMonth = Number(fiscalStartMonth) || 10;
    Logger.log('debugReportData: calling _getReportData(%s, {year:%s,fiscal:%s})', reportType, year, fiscalStartMonth);
    if (typeof _getReportData !== 'function') {
      Logger.log('debugReportData: _getReportData not found');
      return;
    }
    var resp = _getReportData(reportType, { year: year, fiscalStartMonth: fiscalStartMonth });
    Logger.log('debugReportData: _getReportData returned success=%s', resp && resp.success);
    Logger.log('debugReportData: raw resp = %s', JSON.stringify(resp));
    // Normalize if possible
    if (typeof _normalizeReportForTemplate === 'function') {
      var report = _normalizeReportForTemplate(reportType, (resp.data || resp));
      Logger.log('debugReportData: after normalize, items count=%s', (report && report.items && report.items.length) ? report.items.length : 0);
      if (report && report.items && report.items.length > 0) {
        Logger.log('debugReportData: sample item[0] = %s', JSON.stringify(report.items[0]));
      }
      return report;
    } else {
      Logger.log('debugReportData: normalizer not found, returning raw resp');
      return resp;
    }
  } catch (e) {
    Logger.log('debugReportData ERROR: %s', String(e));
    throw e;
  }
}

function debug_exportInline_toDrive() {
  try {
    Logger.log('debug_exportInline_toDrive: calling exportSupportReportPdf_inline');
    if (typeof exportSupportReportPdf_inline !== 'function') {
      Logger.log('debug_exportInline_toDrive: exportSupportReportPdf_inline not found');
      return;
    }
    var res = exportSupportReportPdf_inline('monthly', { year: new Date().getFullYear(), fiscalStartMonth: 10, fileName: 'debug_support.pdf' });
    Logger.log('debug_exportInline_toDrive: res=%s', JSON.stringify(res));
    if (!res || !res.success) {
      Logger.log('debug_exportInline_toDrive: server returned error: %s', res && res.message);
      return res;
    }
    var b64 = (res.data && res.data.base64) ? res.data.base64 : null;
    if (!b64) {
      Logger.log('debug_exportInline_toDrive: base64 missing in response');
      return res;
    }
    var bytes = Utilities.base64Decode(b64);
    var blob = Utilities.newBlob(bytes, 'application/pdf', res.data.fileName || 'debug_support.pdf');
    var file = DriveApp.createFile(blob);
    Logger.log('debug_exportInline_toDrive: file created url=%s id=%s size=%s', file.getUrl(), file.getId(), file.getSize());
    return { fileUrl: file.getUrl(), fileId: file.getId(), size: file.getSize() };
  } catch (e) {
    Logger.log('debug_exportInline_toDrive ERROR: %s', String(e));
    throw e;
  }
}