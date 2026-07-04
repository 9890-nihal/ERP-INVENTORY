/* =========================
   INVENTORY MODULE
========================= */

const INVENTORY_STOCK_TYPES = ['NEW', 'OLD', 'REPAIRED', 'FAULTY', 'SCRAP'];
const INVENTORY_HEADER_ROWS = 3;

function getInventoryInitialData(payload) {
  payload = payload || {};
  const periodInfo = getInventoryPeriodInfo_(payload.period);
  const rows = buildInventoryRows_(periodInfo);
  const stats = buildInventoryStats_(rows);

  // Fast UI loading: do not write Sheet4 on every screen open/auto refresh.
  // Sheet4 sync runs only when user clicks Refresh / Sync Sheet4.
  if (payload.syncSheet === true) {
    syncInventorySheet_(periodInfo, rows, stats);
  }

  return {
    success: true,
    period: periodInfo.period,
    periodLabel: periodInfo.label,
    fromDate: toDisplayDateDashFromDate_(periodInfo.start),
    toDate: toDisplayDateDashFromDate_(new Date(periodInfo.end.getTime() - 24 * 60 * 60 * 1000)),
    rows: rows,
    stats: stats,
    stockTypes: INVENTORY_STOCK_TYPES
  };
}

function updateInventoryStorageLocation(payload) {
  if (!payload || !payload.itemKey) {
    throw new Error('Invalid inventory item');
  }

  const itemKey = clean_(payload.itemKey);
  const location = clean_(payload.location);
  const periodInfo = getInventoryPeriodInfo_(payload.period);

  const sheet = getOrCreateInventorySheet_();
  const values = sheet.getDataRange().getDisplayValues();

  if (values.length >= INVENTORY_HEADER_ROWS) {
    const headers = values[INVENTORY_HEADER_ROWS - 1].map(clean_);
    const keyCol = findHeaderCol_(headers, 'ITEM KEY');
    const locationCol = findHeaderCol_(headers, 'STORAGE LOCATION');

    if (keyCol !== -1 && locationCol !== -1) {
      for (let r = INVENTORY_HEADER_ROWS; r < values.length; r++) {
        if (clean_(values[r][keyCol]) === itemKey) {
          sheet.getRange(r + 1, locationCol + 1).setValue(location);
          break;
        }
      }
    }
  }

  const rows = buildInventoryRows_(periodInfo);
  const stats = buildInventoryStats_(rows);
  syncInventorySheet_(periodInfo, rows, stats);

  return {
    success: true,
    message: 'Storage location updated',
    period: periodInfo.period,
    periodLabel: periodInfo.label,
    fromDate: toDisplayDateDashFromDate_(periodInfo.start),
    toDate: toDisplayDateDashFromDate_(new Date(periodInfo.end.getTime() - 24 * 60 * 60 * 1000)),
    rows: rows,
    stats: stats,
    stockTypes: INVENTORY_STOCK_TYPES
  };
}

function buildInventoryRows_(periodInfo) {
  const masterItems = getInventoryMasterItems_();
  // Sheet4 is also used as a light editable register.
  // Manual fields preserved from Sheet4: MINIMUM STOCK QTY and STORAGE LOCATION.
  const manualMap = getInventoryManualMap_();

  const openingInward = aggregateInwardByPeriod_(null, periodInfo.start);
  const openingOutward = aggregateOutwardByPeriod_(null, periodInfo.start);
  const periodInward = aggregateInwardByPeriod_(periodInfo.start, periodInfo.end);
  const periodOutward = aggregateOutwardByPeriod_(periodInfo.start, periodInfo.end);

  const itemMap = {};

  masterItems.forEach(function(item) {
    itemMap[item.itemKey] = item;
  });

  [openingInward, openingOutward, periodInward, periodOutward].forEach(function(map) {
    Object.keys(map).forEach(function(itemKey) {
      if (!itemMap[itemKey]) {
        itemMap[itemKey] = {
          itemKey: itemKey,
          goodsType: 'UNMAPPED',
          mainGroup: '',
          subGroup: '',
          goodsDescription: map[itemKey].itemName || itemKey,
          uom: '',
          minimumStockQty: 0,
          brand: '',
          model: '',
          unmapped: true
        };
      }
    });
  });

  const rows = Object.keys(itemMap).sort(function(a, b) {
    const x = itemMap[a];
    const y = itemMap[b];
    return [x.mainGroup, x.subGroup, x.goodsDescription, x.brand, x.model].join('|')
      .localeCompare([y.mainGroup, y.subGroup, y.goodsDescription, y.brand, y.model].join('|'));
  }).map(function(itemKey) {
    const item = itemMap[itemKey];
    const opening = zeroStock_();
    const inward = zeroStock_();
    const outward = zeroStock_();
    const current = zeroStock_();

    INVENTORY_STOCK_TYPES.forEach(function(stockType) {
      opening[stockType] = getStockQty_(openingInward, itemKey, stockType) - getStockQty_(openingOutward, itemKey, stockType);
      inward[stockType] = getStockQty_(periodInward, itemKey, stockType);
      outward[stockType] = getStockQty_(periodOutward, itemKey, stockType);
      current[stockType] = opening[stockType] + inward[stockType] - outward[stockType];
    });

    const currentStockLevel = totalStock_(current);
    const manual = manualMap[item.itemKey] || {};
    const minimumStockQty = manual.hasMinimumStockQty
      ? Number(manual.minimumStockQty) || 0
      : Number(item.minimumStockQty) || 0;
    const reorderRequired = minimumStockQty > 0 && currentStockLevel <= minimumStockQty;

    return {
      itemKey: item.itemKey,
      goodsType: item.goodsType,
      mainGroup: item.mainGroup,
      subGroup: item.subGroup,
      goodsDescription: item.goodsDescription,
      uom: item.uom,
      minimumStockQty: minimumStockQty,
      brand: item.brand,
      model: item.model,
      opening: opening,
      inward: inward,
      outward: outward,
      current: current,
      currentStockLevel: currentStockLevel,
      reorderIndicator: reorderRequired ? 'REORDER REQUIRED' : 'NOT REQUIRED',
      reorderRequired: reorderRequired,
      storageLocation: manual.storageLocation || '',
      unmapped: !!item.unmapped
    };
  });

  return rows;
}

function getInventoryMasterItems_() {
  const sheet = getSheet_(ERP_CONFIG.MASTER_SHEET_NAME);
  const values = sheet.getDataRange().getDisplayValues();
  const data = [];
  const seen = {};

  if (values.length < 2) return data;

  for (let r = 1; r < values.length; r++) {
    const row = values[r];

    const goodsType = clean_(row[0]);              // A
    const mainGroup = clean_(row[1]);              // B
    const subGroup = clean_(row[2]);               // C
    const goodsDescription = clean_(row[3]);       // D
    const uom = clean_(row[4]);                    // E
    const minimumStockQty = getNumber_(row[5]);    // F
    const brand = clean_(row[6]) || 'Not Available'; // G
    const model = clean_(row[7]) || 'Not Available'; // H

    if (!goodsType && !mainGroup && !subGroup && !goodsDescription) continue;
    if (!goodsDescription) continue;

    const itemKey = makeInventoryItemKey_(mainGroup, subGroup, goodsDescription);
    if (seen[itemKey]) continue;
    seen[itemKey] = true;

    data.push({
      itemKey: itemKey,
      goodsType: goodsType,
      mainGroup: mainGroup,
      subGroup: subGroup,
      goodsDescription: goodsDescription,
      uom: uom,
      minimumStockQty: minimumStockQty,
      brand: brand,
      model: model
    });
  }

  return data;
}

function aggregateInwardByPeriod_(startDate, endDate) {
  const sheet = getOrCreateInwardSheet_();
  const values = sheet.getDataRange().getDisplayValues();
  const result = {};

  if (values.length < 2) return result;

  const headers = values[0].map(clean_);

  const dateCol = findHeaderCol_(headers, 'DATE');
  const mainGroupCol = findHeaderCol_(headers, 'MAIN GROUP');
  const subGroupCol = findHeaderCol_(headers, 'SUB GROUP');
  const descCol = findHeaderCol_(headers, 'GOODS DESCRIPTION');
  const qtyCol = findHeaderCol_(headers, 'INWARD QTY');
  const stockTypeCol = findHeaderCol_(headers, 'STOCK TYPE');

  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const txDate = parseInventoryDate_(row[dateCol]);

    if (!isDateInInventoryRange_(txDate, startDate, endDate)) continue;

    const mainGroup = clean_(row[mainGroupCol]);
    const subGroup = clean_(row[subGroupCol]);
    const itemName = clean_(row[descCol]);
    const qty = getNumber_(row[qtyCol]);
    const stockType = normalizeInventoryStockType_(row[stockTypeCol]);

    if (!itemName || qty <= 0) continue;

    addInventoryQty_(result, makeInventoryItemKey_(mainGroup, subGroup, itemName), stockType, qty, itemName);
  }

  return result;
}

function aggregateOutwardByPeriod_(startDate, endDate) {
  const sheet = getSheet_(ERP_CONFIG.OUTWARD_SHEET_NAME);
  const values = sheet.getDataRange().getDisplayValues();
  const result = {};

  if (values.length < 2) return result;

  const headers = values[0].map(clean_);

  const issueDateCol = findInventoryHeaderColByAny_(headers, ['Issue Date', 'Issue_Date', 'ISSUE DATE', 'ISSUE_DATE']);
  const requestDateCol = findInventoryHeaderColByAny_(headers, ['Date', 'DATE']);
  const mainGroupCol = findInventoryHeaderColByAny_(headers, ['Main_Group', 'Main Group', 'MAIN_GROUP', 'MAIN GROUP']);
  const subGroupCol = findInventoryHeaderColByAny_(headers, ['Sub_Group', 'Sub Group', 'SUB_GROUP', 'SUB GROUP']);
  const itemCol = findInventoryHeaderColByAny_(headers, ['Item_Name', 'Item Name', 'ITEM_NAME', 'ITEM NAME']);
  const issueQtyCol = findInventoryHeaderColByAny_(headers, ['Issue Qty', 'Issue_Qty', 'ISSUE QTY', 'ISSUE_QTY']);
  const requestQtyCol = findInventoryHeaderColByAny_(headers, ['Quantity', 'QUANTITY', 'Qty', 'QTY']);
  const stockTypeCol = findInventoryHeaderColByAny_(headers, ['STOCK TYPE', 'Stock Type', 'Stock_Type', 'STOCK_TYPE']);
  const typeCol = findInventoryHeaderColByAny_(headers, ['Type', 'TYPE', 'Status', 'STATUS', 'Request Status', 'REQUEST STATUS']);
  const transactionCol = findInventoryHeaderColByAny_(headers, ['Transaction_ID', 'Transaction ID', 'TRANSACTION_ID', 'TRANSACTION ID']);
  const dcMiCol = findInventoryHeaderColByAny_(headers, ['DC / MI', 'DC/MI', 'DC_MI', 'DC MI']);

  for (let r = 1; r < values.length; r++) {
    const row = values[r];

    const issueDateValue = issueDateCol !== -1 ? clean_(row[issueDateCol]) : '';
    const requestDateValue = requestDateCol !== -1 ? clean_(row[requestDateCol]) : '';
    const displayDate = issueDateValue || requestDateValue;
    const txDate = parseInventoryDate_(displayDate);

    if (!isDateInInventoryRange_(txDate, startDate, endDate)) continue;

    const mainGroup = mainGroupCol !== -1 ? clean_(row[mainGroupCol]) : '';
    const subGroup = subGroupCol !== -1 ? clean_(row[subGroupCol]) : '';
    const itemName = itemCol !== -1 ? clean_(row[itemCol]) : '';

    let qty = issueQtyCol !== -1 ? getNumber_(row[issueQtyCol]) : 0;
    if (qty <= 0 && requestQtyCol !== -1) {
      qty = getNumber_(row[requestQtyCol]);
    }

    const stockType = normalizeInventoryStockType_(stockTypeCol !== -1 ? row[stockTypeCol] : 'NEW');

    if (!itemName || qty <= 0) continue;

    // Return rows are added back to stock by using negative outward.
    // Current = Opening + Inward - Outward, so negative outward increases current stock.
    const isReturn = isInventoryReturnOutwardRow_(row, typeCol, transactionCol, dcMiCol);
    const signedQty = isReturn ? -qty : qty;

    addInventoryQty_(
      result,
      makeInventoryItemKey_(mainGroup, subGroup, itemName),
      stockType,
      signedQty,
      itemName
    );
  }

  return result;
}

function isInventoryReturnOutwardRow_(row, typeCol, transactionCol, dcMiCol) {
  const type = typeCol !== -1 ? clean_(row[typeCol]).toUpperCase() : '';
  const transactionId = transactionCol !== -1 ? clean_(row[transactionCol]).toUpperCase() : '';
  const dcMi = dcMiCol !== -1 ? clean_(row[dcMiCol]).toUpperCase() : '';

  return (
    type === 'RETURNED' ||
    transactionId.indexOf('_RETURN_') !== -1 ||
    dcMi === 'RETURN'
  );
}

function findInventoryHeaderColByAny_(headers, headerNames) {
  const aliases = headerNames.map(function(name) {
    return normalizeInventoryHeaderAlias_(name);
  });

  for (let i = 0; i < headers.length; i++) {
    if (aliases.indexOf(normalizeInventoryHeaderAlias_(headers[i])) !== -1) {
      return i;
    }
  }

  return -1;
}

function normalizeInventoryHeaderAlias_(value) {
  return String(value || '')
    .trim()
    .toUpperCase()
    .replace(/[^A-Z0-9]/g, '');
}

function syncInventorySheet_(periodInfo, rows, stats) {
  const sheet = getOrCreateInventorySheet_();
  const headers = getInventorySheetHeaders_();
  const totalCols = headers.length;

  // Important: Sheet4 has grouped headers. Before recreating the register,
  // remove old merges first; otherwise freeze/frozen boundaries can fail.
  // We also show hidden columns before writing because ITEM KEY is hidden after sync.
  try {
    sheet.showColumns(1, sheet.getMaxColumns());
  } catch (e) {}
  sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).breakApart();
  sheet.clear({contentsOnly: false});

  // Do not merge row 1 across all columns. A full-width merged title breaks
  // frozen columns because any frozen boundary becomes part of that merged cell.
  sheet.getRange(1, 1, 1, totalCols)
    .setBackground('#f4b183')
    .setFontWeight('bold')
    .setFontSize(13)
    .setHorizontalAlignment('center');
  sheet.getRange(1, 1).setValue('Inventory Summary - ' + periodInfo.label + ' | Opening = previous closing, Inward/Outward = selected month only');

  const groupHeaders = new Array(totalCols).fill('');
  groupHeaders[0] = 'ITEM MASTER - SHEET2 (A:H)';
  groupHeaders[8] = 'OP.STK';
  groupHeaders[13] = 'INWARD';
  groupHeaders[18] = 'OUTWARD';
  groupHeaders[23] = 'CURRENT STK';
  groupHeaders[28] = 'REORDER / LOCATION';
  sheet.getRange(2, 1, 1, totalCols).setValues([groupHeaders]);

  sheet.getRange(2, 1, 1, 8).merge().setBackground('#dbeafe').setFontWeight('bold').setHorizontalAlignment('center');
  sheet.getRange(2, 9, 1, 5).merge().setBackground('#dbeafe').setFontWeight('bold').setHorizontalAlignment('center');
  sheet.getRange(2, 14, 1, 5).merge().setBackground('#dbeafe').setFontWeight('bold').setHorizontalAlignment('center');
  sheet.getRange(2, 19, 1, 5).merge().setBackground('#dbeafe').setFontWeight('bold').setHorizontalAlignment('center');
  sheet.getRange(2, 24, 1, 5).merge().setBackground('#fde68a').setFontWeight('bold').setHorizontalAlignment('center');
  sheet.getRange(2, 29, 1, 4).merge().setBackground('#dbeafe').setFontWeight('bold').setHorizontalAlignment('center');

  sheet.getRange(3, 1, 1, totalCols).setValues([headers])
    .setFontWeight('bold')
    .setBackground('#e5eef7')
    .setHorizontalAlignment('center')
    .setWrap(true);

  const body = rows.map(function(row) {
    return getInventorySheetRow_(row);
  });

  if (body.length) {
    sheet.getRange(4, 1, body.length, totalCols).setValues(body);
    sheet.getRange(4, 9, body.length, 20).setHorizontalAlignment('right');
    sheet.getRange(4, 29, body.length, 1).setHorizontalAlignment('right');

    for (let i = 0; i < rows.length; i++) {
      const indicatorCell = sheet.getRange(4 + i, 30);
      if (rows[i].reorderRequired) {
        indicatorCell.setBackground('#fee2e2').setFontColor('#991b1b').setFontWeight('bold');
      } else {
        indicatorCell.setBackground('#dcfce7').setFontColor('#166534').setFontWeight('bold');
      }

      if (rows[i].unmapped) {
        sheet.getRange(4 + i, 1, 1, totalCols).setBackground('#fff7ed');
      }
    }
  }

  sheet.setFrozenRows(3);
  // Do not freeze columns in Sheet4. With wide merged group headers, frozen
  // columns create the Google Sheets 'current window is too small' warning.
  sheet.setFrozenColumns(0);
  applyInventoryColumnWidths_(sheet, totalCols);
  sheet.hideColumns(totalCols);

  SpreadsheetApp.flush();
}

function getInventorySheetRow_(row) {
  return [
    row.goodsType,
    row.mainGroup,
    row.subGroup,
    row.goodsDescription,
    row.uom,
    row.minimumStockQty,
    row.brand,
    row.model,
    row.opening.NEW,
    row.opening.OLD,
    row.opening.REPAIRED,
    row.opening.FAULTY,
    row.opening.SCRAP,
    row.inward.NEW,
    row.inward.OLD,
    row.inward.REPAIRED,
    row.inward.FAULTY,
    row.inward.SCRAP,
    row.outward.NEW,
    row.outward.OLD,
    row.outward.REPAIRED,
    row.outward.FAULTY,
    row.outward.SCRAP,
    row.current.NEW,
    row.current.OLD,
    row.current.REPAIRED,
    row.current.FAULTY,
    row.current.SCRAP,
    row.currentStockLevel,
    row.reorderIndicator,
    row.storageLocation,
    row.itemKey
  ];
}

function getInventorySheetHeaders_() {
  return [
    'GOODS TYPE',
    'MAIN GROUP',
    'SUB GROUP',
    'GOODS DESCRIPTION',
    'UOM',
    'MINIMUM STOCK QTY',
    'BRAND',
    'MODEL',
    'OP NEW',
    'OP OLD',
    'OP REPAIRED',
    'OP FAULTY',
    'OP SCRAP',
    'IN NEW',
    'IN OLD',
    'IN REPAIRED',
    'IN FAULTY',
    'IN SCRAP',
    'OUT NEW',
    'OUT OLD',
    'OUT REPAIRED',
    'OUT FAULTY',
    'OUT SCRAP',
    'CUR NEW',
    'CUR OLD',
    'CUR REPAIRED',
    'CUR FAULTY',
    'CUR SCRAP',
    'CURRENT STOCK LEVEL',
    'REORDER INDICATOR',
    'STORAGE LOCATION',
    'ITEM KEY'
  ];
}

function getInventorySheetName_() {
  return (typeof ERP_CONFIG !== 'undefined' && ERP_CONFIG.INVENTORY_SHEET_NAME)
    ? ERP_CONFIG.INVENTORY_SHEET_NAME
    : 'Sheet4';
}

function getInventoryManualMap_() {
  const ss = getSpreadsheet_();
  const sheet = ss.getSheetByName(getInventorySheetName_());
  const map = {};

  if (!sheet || sheet.getLastRow() < INVENTORY_HEADER_ROWS + 1) return map;

  const values = sheet.getDataRange().getDisplayValues();
  const headers = values[INVENTORY_HEADER_ROWS - 1].map(clean_);
  const keyCol = findHeaderCol_(headers, 'ITEM KEY');
  const minCol = findHeaderCol_(headers, 'MINIMUM STOCK QTY');
  const locationCol = findHeaderCol_(headers, 'STORAGE LOCATION');

  if (keyCol === -1) return map;

  for (let r = INVENTORY_HEADER_ROWS; r < values.length; r++) {
    const itemKey = clean_(values[r][keyCol]);
    if (!itemKey) continue;

    map[itemKey] = {
      hasMinimumStockQty: minCol !== -1,
      minimumStockQty: minCol !== -1 ? getNumber_(values[r][minCol]) : 0,
      storageLocation: locationCol !== -1 ? clean_(values[r][locationCol]) : ''
    };
  }

  return map;
}

function applyInventoryColumnWidths_(sheet, totalCols) {
  const widths = [
    115, // GOODS TYPE
    150, // MAIN GROUP
    150, // SUB GROUP
    280, // GOODS DESCRIPTION
    70,  // UOM
    110, // MINIMUM STOCK QTY
    140, // BRAND
    130  // MODEL
  ];

  for (let c = 1; c <= totalCols; c++) {
    let width = widths[c - 1] || 72;

    if (c >= 9 && c <= 28) width = 68;      // stock type quantity columns
    if (c === 29) width = 115;              // CURRENT STOCK LEVEL
    if (c === 30) width = 150;              // REORDER INDICATOR
    if (c === 31) width = 190;              // STORAGE LOCATION
    if (c === 32) width = 1;                // hidden ITEM KEY

    sheet.setColumnWidth(c, width);
  }
}

function getOrCreateInventorySheet_() {
  const ss = getSpreadsheet_();
  let sheet = ss.getSheetByName(getInventorySheetName_());

  if (!sheet) {
    sheet = ss.insertSheet(getInventorySheetName_());
  }

  return sheet;
}

function buildInventoryStats_(rows) {
  let totalCurrentQty = 0;
  let totalInwardQty = 0;
  let totalOutwardQty = 0;
  let reorderRequired = 0;
  let negativeStockItems = 0;

  rows.forEach(function(row) {
    totalCurrentQty += row.currentStockLevel;
    totalInwardQty += totalStock_(row.inward);
    totalOutwardQty += totalStock_(row.outward);
    if (row.reorderRequired) reorderRequired++;
    if (row.currentStockLevel < 0) negativeStockItems++;
  });

  return {
    totalItems: rows.length,
    totalCurrentQty: totalCurrentQty,
    totalInwardQty: totalInwardQty,
    totalOutwardQty: totalOutwardQty,
    reorderRequired: reorderRequired,
    negativeStockItems: negativeStockItems
  };
}

function getInventoryPeriodInfo_(period) {
  const tz = Session.getScriptTimeZone();
  const today = new Date();
  let year = Number(Utilities.formatDate(today, tz, 'yyyy'));
  let month = Number(Utilities.formatDate(today, tz, 'MM'));

  if (period && /^\d{4}-\d{2}$/.test(String(period))) {
    const parts = String(period).split('-');
    year = Number(parts[0]);
    month = Number(parts[1]);
  }

  const start = new Date(year, month - 1, 1);
  const end = new Date(year, month, 1);

  return {
    period: String(year) + '-' + String(month).padStart(2, '0'),
    label: Utilities.formatDate(start, tz, 'MMM yyyy'),
    start: start,
    end: end
  };
}

function makeInventoryItemKey_(mainGroup, subGroup, itemName) {
  return [mainGroup, subGroup, itemName]
    .map(function(v) { return clean_(v).toUpperCase(); })
    .join('|');
}

function zeroStock_() {
  const obj = {};
  INVENTORY_STOCK_TYPES.forEach(function(stockType) {
    obj[stockType] = 0;
  });
  return obj;
}

function addInventoryQty_(result, itemKey, stockType, qty, itemName) {
  if (!result[itemKey]) {
    result[itemKey] = {
      itemName: itemName,
      stock: zeroStock_()
    };
  }

  result[itemKey].stock[stockType] += qty;
}

function getStockQty_(map, itemKey, stockType) {
  if (!map[itemKey]) return 0;
  return Number(map[itemKey].stock[stockType]) || 0;
}

function totalStock_(stockObj) {
  return INVENTORY_STOCK_TYPES.reduce(function(sum, stockType) {
    return sum + (Number(stockObj[stockType]) || 0);
  }, 0);
}

function normalizeInventoryStockType_(value) {
  const stockType = clean_(value).toUpperCase();
  return INVENTORY_STOCK_TYPES.indexOf(stockType) !== -1 ? stockType : 'NEW';
}

function findHeaderCol_(headers, headerName) {
  const target = String(headerName || '').trim().toUpperCase();
  return headers.findIndex(function(header) {
    return String(header || '').trim().toUpperCase() === target;
  });
}

function parseInventoryDate_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return new Date(value.getFullYear(), value.getMonth(), value.getDate());
  }

  const text = clean_(value);
  if (!text) return null;

  let match = text.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (match) {
    return new Date(Number(match[1]), Number(match[2]) - 1, Number(match[3]));
  }

  match = text.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/);
  if (match) {
    return new Date(Number(match[3]), Number(match[2]) - 1, Number(match[1]));
  }

  const parsed = new Date(text);
  if (!isNaN(parsed.getTime())) {
    return new Date(parsed.getFullYear(), parsed.getMonth(), parsed.getDate());
  }

  return null;
}

function isDateInInventoryRange_(date, startDate, endDate) {
  if (!date) return false;
  if (startDate && date < startDate) return false;
  if (endDate && date >= endDate) return false;
  return true;
}

function getNumber_(value) {
  if (value === null || value === undefined || value === '') return 0;
  return Number(String(value).replace(/,/g, '')) || 0;
}

function toDisplayDateDashFromDate_(date) {
  const dd = String(date.getDate()).padStart(2, '0');
  const mm = String(date.getMonth() + 1).padStart(2, '0');
  const yyyy = date.getFullYear();
  return dd + '-' + mm + '-' + yyyy;
}
