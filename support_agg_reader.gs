/**
 * buildItemsFromAgg (include all Support items + sort by itemId)
 *
 * - อ่าน Support_Agg เพื่อรวมยอด per item per yyyy-mm (fiscal mapping as before)
 * - เติม (ensure) ให้ครอบคลุมทุกรายการที่อยู่ใน sheet "Support"
 * - คืน array ของ items ที่มี fields: itemId, itemName, months[12], budget, remaining, sample
 * - ผลลัพธ์จะถูกเรียงตาม itemId (รหัส) ก่อนคืน
 */
function buildItemsFromAgg(year, fiscalStartMonth) {
  year = Number(year) || (new Date()).getFullYear();
  fiscalStartMonth = Number(fiscalStartMonth) || 10; // 1..12

  var ss = SpreadsheetApp.getActive();
  var aggSheet = ss.getSheetByName('Support_Agg');
  var supportSheet = ss.getSheetByName('Support');

  // read raw agg rows (may be headerless)
  var aggRows = [];
  if (aggSheet) aggRows = aggSheet.getDataRange().getValues() || [];

  // header detection for aggRows (same heuristic as before)
  var startIndex = 0;
  if (aggRows && aggRows.length > 0) {
    var firstRowSecondCol = (aggRows[0] && aggRows[0][1]) ? String(aggRows[0][1]).trim() : '';
    var yyyyMmRegex = /^\d{4}-\d{2}$/;
    if (!firstRowSecondCol || !yyyyMmRegex.test(firstRowSecondCol)) startIndex = 1;
  }

  // helper to aggregate using fiscalStartYear
  function aggregateWithFiscalStart(fiscalStartYear) {
    var fiscalStart = new Date(fiscalStartYear, fiscalStartMonth - 1, 1);
    var fiscalEnd = new Date(fiscalStart.getFullYear() + 1, fiscalStart.getMonth(), 1);

    var itemsMap = {}; // key -> { itemId, itemName, months, budget, remaining, sample }

    for (var i = startIndex; i < aggRows.length; i++) {
      try {
        var row = aggRows[i] || [];
        var itemId = String(row[0] || '').trim();
        var ym = String(row[1] || '').trim(); // "2025-11"
        var sumVal = Number(row[2] || 0) || 0;
        if (!itemId || !ym) continue;

        var parts = ym.split('-');
        if (parts.length < 2) continue;
        var calYear = Number(parts[0]);
        var calMonth = Number(parts[1]);
        if (isNaN(calYear) || isNaN(calMonth)) continue;

        var dt = new Date(calYear, calMonth - 1, 1);
        if (!(dt >= fiscalStart && dt < fiscalEnd)) continue;

        var fiscalIndex = ((calMonth - fiscalStartMonth + 12) % 12); // 0..11

        if (!itemsMap[itemId]) {
          itemsMap[itemId] = {
            itemId: itemId,
            itemName: '',
            months: Array.apply(null, Array(12)).map(Number.prototype.valueOf,0),
            budget: 0,
            remaining: 0,
            sample: []
          };
        }
        itemsMap[itemId].months[fiscalIndex] = Number(itemsMap[itemId].months[fiscalIndex] || 0) + sumVal;
        itemsMap[itemId].sample.push({ rowIndex: i+1, ym: ym, value: sumVal, fiscalIndex: fiscalIndex });
      } catch (eRow) {
        Logger.log('buildItemsFromAgg: agg row parse error at %s: %s', i+1, String(eRow));
      }
    }

    return itemsMap;
  } // end aggregateWithFiscalStart

  // Try interpretation: year as fiscal end (default), fallback to year as fiscal start
  var fiscalStartYearDefault = (fiscalStartMonth > 1) ? (year - 1) : year;
  var itemsMap = aggregateWithFiscalStart(fiscalStartYearDefault);
  Logger.log('buildItemsFromAgg: try fiscalStartYear=%s -> itemsMapKeys=%s', fiscalStartYearDefault, Object.keys(itemsMap).length);

  if (!itemsMap || Object.keys(itemsMap).length === 0) {
    var altFiscalStart = year;
    var altMap = aggregateWithFiscalStart(altFiscalStart);
    Logger.log('buildItemsFromAgg: fallback try fiscalStartYear=%s -> itemsMapKeys=%s', altFiscalStart, Object.keys(altMap).length);
    if (Object.keys(altMap).length > 0) itemsMap = altMap;
  }

  // --- NOW ensure all items from Support sheet are present ---
  if (supportSheet) {
    try {
      var sData = supportSheet.getDataRange().getValues();
      if (sData && sData.length > 0) {
        // find candidate columns (flexible detection)
        var sHeaders = sData[0].map(function(h){ return (h||'').toString().trim().toLowerCase(); });
        var findIdx = function(choices){
          for (var k=0;k<choices.length;k++){
            var name = choices[k].toString().toLowerCase();
            var idx = sHeaders.indexOf(name);
            if (idx !== -1) return idx;
          }
          return -1;
        };
        var colItemId = findIdx(['item id','itemid','id','รหัส','รหัสรายการ','item']);
        if (colItemId === -1) colItemId = 0; // fallback to col A
        var colItemName = findIdx(['item','item name','รายการ','description','detail']);
        var colBudget = findIdx(['budget','งบประมาณ','งบประมาณจัดสรร','budgetmoney','budget_money']);
        var colRemaining = findIdx(['remaining','balance','งบประมาณคงเหลือ']);

        // iterate support rows and insert missing items with zeros
        for (var r = 1; r < sData.length; r++) {
          try {
            var rawId = String(sData[r][colItemId] || '').trim();
            if (!rawId) continue;
            if (!itemsMap[rawId]) {
              var name = (colItemName >= 0) ? String(sData[r][colItemName] || '') : '';
              var budgetVal = (colBudget >= 0) ? Number(sData[r][colBudget] || 0) : 0;
              var remainingVal = (colRemaining >= 0) ? Number(sData[r][colRemaining] || 0) : 0;
              itemsMap[rawId] = {
                itemId: rawId,
                itemName: name || '',
                months: Array.apply(null, Array(12)).map(Number.prototype.valueOf,0),
                budget: budgetVal || 0,
                remaining: remainingVal || 0,
                sample: []
              };
              Logger.log('buildItemsFromAgg: added missing item from Support sheet rawId=%s name=%s', rawId, name);
            } else {
              // enrich name/budget/remaining if present
              if (colItemName >= 0 && !itemsMap[rawId].itemName) itemsMap[rawId].itemName = String(sData[r][colItemName] || '') || itemsMap[rawId].itemName;
              if (colBudget >= 0 && (!itemsMap[rawId].budget || itemsMap[rawId].budget === 0)) itemsMap[rawId].budget = Number(sData[r][colBudget] || itemsMap[rawId].budget || 0);
              if (colRemaining >= 0 && (!itemsMap[rawId].remaining || itemsMap[rawId].remaining === 0)) itemsMap[rawId].remaining = Number(sData[r][colRemaining] || itemsMap[rawId].remaining || 0);
            }
          } catch (enr) {
            Logger.log('buildItemsFromAgg: support enrich row %s error: %s', r+1, String(enr));
          }
        } // end support rows
      } // end sData
    } catch (eSup) {
      Logger.log('buildItemsFromAgg: support sheet read failed: %s', String(eSup));
    }
  } // end if supportSheet

  // convert itemsMap -> items array and sort by itemId
  var items = [];
  for (var key in itemsMap) {
    if (Object.prototype.hasOwnProperty.call(itemsMap, key)) {
      var it = itemsMap[key];
      if (!it.itemName) it.itemName = it.itemId || key;
      if (!Array.isArray(it.months)) it.months = Array.apply(null, Array(12)).map(Number.prototype.valueOf,0);
      while (it.months.length < 12) it.months.push(0);
      items.push(it);
    }
  }

  // Sort items by itemId (lexicographically). If you want numeric-aware sort, adjust comparer.
  items.sort(function(a,b){
    var A = String(a.itemId || '').toUpperCase();
    var B = String(b.itemId || '').toUpperCase();
    return A.localeCompare(B, 'en', { numeric: true });
  });

  Logger.log('buildItemsFromAgg: final items count=%s', items.length);
  return items;
}