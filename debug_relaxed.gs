// debug_relaxed.gs
// ชั่วคราว: สร้างรายการจาก Support_Agg โดยไม่กรอง fiscal window
function buildItemsFromAgg_relaxed(year, fiscalStartMonth) {
  var ss = SpreadsheetApp.getActive();
  var aggSheet = ss.getSheetByName('Support_Agg');
  if (!aggSheet) {
    Logger.log('buildItemsFromAgg_relaxed: Support_Agg sheet not found');
    return [];
  }
  var rows = aggSheet.getDataRange().getValues();
  Logger.log('buildItemsFromAgg_relaxed: total rows=%s', rows.length);
  var itemsMap = {};
  for (var i = 1; i < rows.length; i++) {
    try {
      var itemId = String(rows[i][0] || '').trim();
      var ym = String(rows[i][1] || '').trim();
      var sumVal = Number(rows[i][2] || 0) || 0;
      if (!itemId || !ym) continue;
      var parts = ym.split('-');
      var calYear = Number(parts[0]) || null;
      var calMonth = Number(parts[1]) || null;
      var key = itemId;
      if (!itemsMap[key]) {
        itemsMap[key] = { itemId: key, itemName: '', months: Array.apply(null, Array(12)).map(Number.prototype.valueOf,0), budget:0, remaining:0 };
      }
      // map by calendar month (just put into month index 0..11 based on calMonth-1)
      if (calMonth >=1 && calMonth <=12) {
        itemsMap[key].months[calMonth - 1] = (itemsMap[key].months[calMonth - 1] || 0) + sumVal;
      }
      // also attach sample ym value for debugging
      if (!itemsMap[key].sample) itemsMap[key].sample = [];
      itemsMap[key].sample.push({ ym: ym, value: sumVal });
    } catch(e) {
      Logger.log('buildItemsFromAgg_relaxed row %s error: %s', i+1, String(e));
    }
  }
  var items = [];
  for (var k in itemsMap) if (Object.prototype.hasOwnProperty.call(itemsMap,k)) items.push(itemsMap[k]);
  Logger.log('buildItemsFromAgg_relaxed: items count=%s', items.length);
  if (items.length > 0) Logger.log('buildItemsFromAgg_relaxed sample: %s', JSON.stringify(items[0]));
  return items;
}

// runner
function runBuildItemsFromAgg_relaxed() {
  var items = buildItemsFromAgg_relaxed(new Date().getFullYear(), 10);
  Logger.log('runBuildItemsFromAgg_relaxed returned %s items', items.length);
  return items;
}