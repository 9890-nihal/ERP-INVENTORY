/* =========================
   INWARD MODULE (Optimized)
========================= */

function getInwardInitialData() {
  return {
    masterData: getInwardMasterData_(),
    recentEntries: getInwardRecentEntries_(),
    suggestions: getInwardSuggestions_(),
    stats: getInwardStats_()
  };
}

function getInwardMasterData_() {
  const sheet = getSheet_(ERP_CONFIG.MASTER_SHEET_NAME);
  const values = sheet.getDataRange().getDisplayValues();
  if (values.length < 2) return [];

  const data = [];
  const seen = {};
  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const goodsType = clean_(row[0]);
    const mainGroup = clean_(row[1]);
    const subGroup = clean_(row[2]);
    const goodsDescription = clean_(row[3]);
    const uom = clean_(row[4]);
    const brand = clean_(row[6]) || 'Not Available'; 
    const model = clean_(row[7]) || 'Not Available';

    if (!mainGroup || !subGroup || !goodsDescription) continue;

    const key = [goodsType, mainGroup, subGroup, goodsDescription, uom, brand, model].join('|').toLowerCase();
    if (seen[key]) continue;

    seen[key] = true;
    data.push({
      goodsType: goodsType, mainGroup: mainGroup, subGroup: subGroup,
      goodsDescription: goodsDescription, uom: uom, brand: brand, model: model
    });
  }
  return data;
}

function saveInwardEntry(payload) {
  validateInwardPayload_(payload);
  const sheet = getOrCreateInwardSheet_();
  const row = [
    toDisplayDateDash_(payload.date), payload.goodsType || '', payload.mainGroup || '',
    payload.subGroup || '', payload.brand || 'Not Available', payload.goodsDescription || '',
    payload.uom || '', payload.inwardQty || '', payload.model || 'Not Available',
    payload.stockType || '', payload.goodsSerialNumber || '', payload.vendorName || '',
    payload.location || '', payload.deliveryChallan || '', payload.invoiceNo || '',
    payload.remark || '', new Date()
  ];
  sheet.appendRow(row);
  SpreadsheetApp.flush();
  return {
    success: true, message: 'Inward entry saved successfully',
    recentEntries: getInwardRecentEntries_(),
    suggestions: getInwardSuggestions_(), stats: getInwardStats_()
  };
}

function updateInwardEntry(payload) {
  const sheet = getOrCreateInwardSheet_();
  const rowNumber = Number(payload.rowNumber);
  if (!rowNumber || rowNumber < 2) throw new Error('Invalid inward row number');

  // Edit mode me ab complete inward row editable hai.
  validateInwardPayload_(payload);

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0].map(clean_);
  const rowRange = sheet.getRange(rowNumber, 1, 1, headers.length);
  const rowData = rowRange.getValues()[0];

  function updateInCol(headerName, val) {
    const idx = headers.indexOf(headerName);
    if (idx !== -1) rowData[idx] = val;
  }

  updateInCol('DATE', toDisplayDateDash_(payload.date));
  updateInCol('GOODS TYPE', payload.goodsType || '');
  updateInCol('MAIN GROUP', payload.mainGroup || '');
  updateInCol('SUB GROUP', payload.subGroup || '');
  updateInCol('BRAND', payload.brand || 'Not Available');
  updateInCol('GOODS DESCRIPTION', payload.goodsDescription || '');
  updateInCol('UOM', payload.uom || '');
  updateInCol('INWARD QTY', payload.inwardQty || '');
  updateInCol('MODEL', payload.model || 'Not Available');
  updateInCol('STOCK TYPE', payload.stockType || '');
  updateInCol('GOODS SERIAL NUMBER', payload.goodsSerialNumber || '');
  updateInCol('VENDOR NAME / PERSON NAME', payload.vendorName || '');
  updateInCol('LOCATION', payload.location || '');
  updateInCol('DELIVERY CHALLAN', payload.deliveryChallan || '');
  updateInCol('INVOICE NO', payload.invoiceNo || '');
  updateInCol('DELIVERY CHALLAN/INVOICE NO', payload.deliveryChallan || '');
  updateInCol('REMARK', payload.remark || '');

  rowRange.setValues([rowData]);
  SpreadsheetApp.flush();

  return {
    success: true, message: 'Inward entry updated successfully',
    recentEntries: getInwardRecentEntries_(),
    suggestions: getInwardSuggestions_(), stats: getInwardStats_()
  };
}

function getInwardRecentEntries_() {
  const sheet = getOrCreateInwardSheet_();
  const values = sheet.getDataRange().getDisplayValues();
  if (values.length < 2) return [];

  const headers = values[0].map(clean_);
  const data = [];
  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const isBlank = row.every(function(cell) { return clean_(cell) === ''; });
    if (isBlank) continue;

    const obj = { rowNumber: r + 1 };
    headers.forEach(function(header, c) { obj[header] = clean_(row[c]); });
    data.push(obj);
  }
  return data.reverse().slice(0, 50);
}

function getInwardSuggestions_() {
  const sheet = getOrCreateInwardSheet_();
  const values = sheet.getDataRange().getDisplayValues();
  if (values.length < 2) return { vendors: [], locations: [] };

  const headers = values[0].map(clean_);
  const vendorCol = headers.indexOf('VENDOR NAME / PERSON NAME');
  const locationCol = headers.indexOf('LOCATION');
  const vendors = [];
  const locations = [];

  for (let r = 1; r < values.length; r++) {
    if (vendorCol !== -1 && clean_(values[r][vendorCol])) vendors.push(clean_(values[r][vendorCol]));
    if (locationCol !== -1 && clean_(values[r][locationCol])) locations.push(clean_(values[r][locationCol]));
  }

  return { vendors: unique_(vendors), locations: unique_(locations) };
}

function getInwardStats_() {
  const sheet = getOrCreateInwardSheet_();
  const values = sheet.getDataRange().getDisplayValues();
  if (values.length < 2) return { totalEntries: 0, todayEntries: 0, assetEntries: 0, totalQty: 0 };

  const headers = values[0].map(clean_);
  const dateCol = headers.indexOf('DATE');
  const goodsTypeCol = headers.indexOf('GOODS TYPE');
  const qtyCol = headers.indexOf('INWARD QTY');
  const today = todayDash_();

  let totalEntries = 0, todayEntries = 0, assetEntries = 0, totalQty = 0;
  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    if (row.every(function(cell) { return clean_(cell) === ''; })) continue;
    
    totalEntries++;
    if (dateCol !== -1 && clean_(row[dateCol]) === today) todayEntries++;
    if (goodsTypeCol !== -1 && clean_(row[goodsTypeCol]).toUpperCase() === 'ASSET') assetEntries++;
    if (qtyCol !== -1) totalQty += Number(String(row[qtyCol]).replace(/,/g, '')) || 0;
  }

  return { totalEntries: totalEntries, todayEntries: todayEntries, assetEntries: assetEntries, totalQty: totalQty };
}

function validateInwardPayload_(payload) {
  if (!payload.date) throw new Error('Date required hai');
  if (!payload.mainGroup) throw new Error('Main Group required hai');
  if (!payload.subGroup) throw new Error('Sub Group required hai');
  if (!payload.brand) throw new Error('Brand required hai');
  if (!payload.goodsDescription) throw new Error('Goods Description required hai');
  if (!payload.goodsType) throw new Error('Goods Type autofill nahi hua');
  if (!payload.uom) throw new Error('UOM autofill nahi hua');
  if (!payload.inwardQty || Number(payload.inwardQty) <= 0) throw new Error('Valid Inward Qty required hai');
  if (!payload.stockType) throw new Error('Stock Type required hai');
  if (String(payload.goodsType).toUpperCase() === 'ASSET' && !payload.goodsSerialNumber) {
    throw new Error('Asset item ke liye Goods Serial Number required hai');
  }
  if (!payload.vendorName) throw new Error('Vendor Name / Person Name required hai');
  if (!payload.location) throw new Error('Location required hai');
}

function getOrCreateInwardSheet_() {
  const ss = getSpreadsheet_();
  let sheet = ss.getSheetByName(ERP_CONFIG.INWARD_SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(ERP_CONFIG.INWARD_SHEET_NAME);

  ensureInwardSheetHeaders_(sheet);
  return sheet;
}

function ensureInwardSheetHeaders_(sheet) {
  const requiredHeaders = [
    'DATE', 'GOODS TYPE', 'MAIN GROUP', 'SUB GROUP', 'BRAND', 'GOODS DESCRIPTION', 'UOM',
    'INWARD QTY', 'MODEL', 'STOCK TYPE', 'GOODS SERIAL NUMBER', 'VENDOR NAME / PERSON NAME',
    'LOCATION', 'DELIVERY CHALLAN', 'INVOICE NO', 'REMARK', 'CREATED AT'
  ];

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(requiredHeaders);
    return;
  }

  let lastColumn = Math.max(sheet.getLastColumn(), requiredHeaders.length);
  let headers = sheet.getRange(1, 1, 1, lastColumn).getDisplayValues()[0].map(clean_);
  const blankHeader = headers.every(function(cell) { return clean_(cell) === ''; });

  if (blankHeader) {
    sheet.getRange(1, 1, 1, requiredHeaders.length).setValues([requiredHeaders]);
    return;
  }

  // Old structure me N column ka header combined tha: DELIVERY CHALLAN/INVOICE NO.
  // New structure: N = DELIVERY CHALLAN, O = INVOICE NO, P = REMARK, Q = CREATED AT.
  const combinedIdx = headers.indexOf('DELIVERY CHALLAN/INVOICE NO');
  const invoiceIdx = headers.indexOf('INVOICE NO');

  if (combinedIdx !== -1) {
    sheet.getRange(1, combinedIdx + 1).setValue('DELIVERY CHALLAN');
    headers[combinedIdx] = 'DELIVERY CHALLAN';
  }

  // If INVOICE NO missing, insert it immediately after DELIVERY CHALLAN.
  headers = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), requiredHeaders.length)).getDisplayValues()[0].map(clean_);
  const deliveryIdx = headers.indexOf('DELIVERY CHALLAN');
  if (headers.indexOf('INVOICE NO') === -1) {
    const insertAfter = deliveryIdx !== -1 ? deliveryIdx + 1 : 14;
    sheet.insertColumnAfter(insertAfter);
    sheet.getRange(1, insertAfter + 1).setValue('INVOICE NO');
  }

  // Final pass: ensure all required headers exist and are in expected first columns.
  // Existing data columns are preserved where already correctly present.
  const finalLastColumn = Math.max(sheet.getLastColumn(), requiredHeaders.length);
  headers = sheet.getRange(1, 1, 1, finalLastColumn).getDisplayValues()[0].map(clean_);

  for (let i = 0; i < requiredHeaders.length; i++) {
    const targetHeader = requiredHeaders[i];
    const current = clean_(sheet.getRange(1, i + 1).getDisplayValue());
    if (!current) {
      sheet.getRange(1, i + 1).setValue(targetHeader);
    }
  }

  // Force exact expected header names for first 17 columns. This matches user's new Sheet3/Inward Register structure.
  sheet.getRange(1, 1, 1, requiredHeaders.length).setValues([requiredHeaders]);
}
