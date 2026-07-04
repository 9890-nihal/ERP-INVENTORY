/* =========================
   MATERIAL LIST MODULE
   Sheet: Material List
   Columns:
   A Goods Type
   B Main Group
   C Sub Group
   D Item Name
   E UOM
   F MOQ / Minimum Stock Qty
   G BRAND
   H MODEL
========================= */

function getMaterialListInitialData() {
  const sheet = getMaterialSheet_();
  ensureMaterialHeaders_(sheet);

  const rows = getMaterialRows_();

  return {
    success: true,
    rows: rows,
    stats: getMaterialStats_(rows),
    filters: getMaterialFilters_(rows)
  };
}

function saveMaterialItem(payload) {
  payload = payload || {};

  const sheet = getMaterialSheet_();
  ensureMaterialHeaders_(sheet);

  const rowNumber = Number(payload.rowNumber) || 0;

  const goodsType = materialClean_(payload.goodsType || 'CONSUMABLE').toUpperCase();
  const mainGroup = materialClean_(payload.mainGroup);
  const subGroup = materialClean_(payload.subGroup);
  const itemName = materialClean_(payload.itemName);
  const uom = materialClean_(payload.uom || 'NOS').toUpperCase();
  const moq = payload.moq === '' || payload.moq === null || payload.moq === undefined
    ? ''
    : Number(payload.moq);
  const brand = materialClean_(payload.brand);
  const model = materialClean_(payload.model);

  if (!goodsType) throw new Error('Goods Type required hai');
  if (!mainGroup) throw new Error('Main Group required hai');
  if (!subGroup) throw new Error('Sub Group required hai');
  if (!itemName) throw new Error('Item Name required hai');
  if (!uom) throw new Error('UOM required hai');

  if (payload.moq !== '' && payload.moq !== null && payload.moq !== undefined && isNaN(moq)) {
    throw new Error('MOQ / Minimum Stock Qty number hona chahiye');
  }

  const rowData = [
    goodsType,
    mainGroup,
    subGroup,
    itemName,
    uom,
    moq,
    brand,
    model
  ];

  const existingRows = getMaterialRows_();
  const newKey = materialItemKey_(rowData);

  const duplicate = existingRows.find(function(row) {
    return row.itemKey === newKey && Number(row.rowNumber) !== Number(rowNumber);
  });

  if (duplicate) {
    throw new Error('Duplicate material already exists at row ' + duplicate.rowNumber);
  }

  if (rowNumber && rowNumber >= 2) {
    sheet.getRange(rowNumber, 1, 1, rowData.length).setValues([rowData]);
  } else {
    sheet.appendRow(rowData);
  }

  SpreadsheetApp.flush();

  const rows = getMaterialRows_();

  return {
    success: true,
    message: rowNumber ? 'Material updated successfully' : 'Material added successfully',
    rows: rows,
    stats: getMaterialStats_(rows),
    filters: getMaterialFilters_(rows)
  };
}

function deleteMaterialItem(payload) {
  payload = payload || {};

  const rowNumber = Number(payload.rowNumber);

  if (!rowNumber || rowNumber < 2) {
    throw new Error('Invalid material row number');
  }

  const sheet = getMaterialSheet_();
  ensureMaterialHeaders_(sheet);

  if (rowNumber > sheet.getLastRow()) {
    throw new Error('Material row not found');
  }

  sheet.deleteRow(rowNumber);
  SpreadsheetApp.flush();

  const rows = getMaterialRows_();

  return {
    success: true,
    message: 'Material deleted successfully',
    rows: rows,
    stats: getMaterialStats_(rows),
    filters: getMaterialFilters_(rows)
  };
}

function getMaterialRows_() {
  const sheet = getMaterialSheet_();
  ensureMaterialHeaders_(sheet);

  const values = sheet.getDataRange().getDisplayValues();

  if (values.length < 2) {
    return [];
  }

  const rows = [];

  for (let r = 1; r < values.length; r++) {
    const row = values[r];

    const isBlank = row.every(function(cell) {
      return materialClean_(cell) === '';
    });

    if (isBlank) continue;

    const goodsType = materialClean_(row[0]);
    const mainGroup = materialClean_(row[1]);
    const subGroup = materialClean_(row[2]);
    const itemName = materialClean_(row[3]);
    const uom = materialClean_(row[4]);
    const moq = materialClean_(row[5]);
    const brand = materialClean_(row[6]);
    const model = materialClean_(row[7]);

    rows.push({
      rowNumber: r + 1,
      itemKey: materialItemKey_([goodsType, mainGroup, subGroup, itemName, uom, moq, brand, model]),
      goodsType: goodsType,
      mainGroup: mainGroup,
      subGroup: subGroup,
      itemName: itemName,
      uom: uom,
      moq: moq,
      brand: brand,
      model: model
    });
  }

  return rows;
}

function getMaterialStats_(rows) {
  rows = rows || [];

  const totalItems = rows.length;
  const assetItems = rows.filter(function(row) {
    return String(row.goodsType).toUpperCase() === 'ASSET';
  }).length;
  const consumableItems = rows.filter(function(row) {
    return String(row.goodsType).toUpperCase() === 'CONSUMABLE';
  }).length;
  const blankMoq = rows.filter(function(row) {
    return materialClean_(row.moq) === '';
  }).length;
  const blankBrand = rows.filter(function(row) {
    return materialClean_(row.brand) === '';
  }).length;

  return {
    totalItems: totalItems,
    assetItems: assetItems,
    consumableItems: consumableItems,
    blankMoq: blankMoq,
    blankBrand: blankBrand
  };
}

function getMaterialFilters_(rows) {
  rows = rows || [];

  return {
    goodsTypes: materialUnique_(rows.map(function(row) { return row.goodsType; })),
    mainGroups: materialUnique_(rows.map(function(row) { return row.mainGroup; })),
    subGroups: materialUnique_(rows.map(function(row) { return row.subGroup; })),
    uoms: materialUnique_(rows.map(function(row) { return row.uom; })),
    brands: materialUnique_(rows.map(function(row) { return row.brand; })),
    models: materialUnique_(rows.map(function(row) { return row.model; }))
  };
}

function getMaterialSheet_() {
  const ss = getMaterialSpreadsheet_();
  const sheetName =
    typeof ERP_CONFIG !== 'undefined' && ERP_CONFIG.MASTER_SHEET_NAME
      ? ERP_CONFIG.MASTER_SHEET_NAME
      : 'Material List';

  let sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }

  return sheet;
}

function getMaterialSpreadsheet_() {
  if (
    typeof ERP_CONFIG !== 'undefined' &&
    ERP_CONFIG.SPREADSHEET_ID
  ) {
    return SpreadsheetApp.openById(ERP_CONFIG.SPREADSHEET_ID);
  }

  return SpreadsheetApp.getActiveSpreadsheet();
}

function ensureMaterialHeaders_(sheet) {
  const headers = [
    'Goods Type',
    'Main Group',
    'Sub Group',
    'Item Name',
    'UOM',
    'MOQ',
    'BRAND',
    'MODEL'
  ];

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#d9ead3');
    return;
  }

  const lastColumn = Math.max(sheet.getLastColumn(), headers.length);
  const currentHeaders = sheet.getRange(1, 1, 1, lastColumn).getDisplayValues()[0];

  const isBlankHeader = currentHeaders.every(function(cell) {
    return materialClean_(cell) === '';
  });

  if (isBlankHeader) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#d9ead3');
    return;
  }

  if (sheet.getLastColumn() < headers.length) {
    const missingHeaders = headers.slice(sheet.getLastColumn());
    sheet.getRange(1, sheet.getLastColumn() + 1, 1, missingHeaders.length).setValues([missingHeaders]);
    sheet.getRange(1, sheet.getLastColumn() - missingHeaders.length + 1, 1, missingHeaders.length)
      .setFontWeight('bold')
      .setBackground('#d9ead3');
  }
}

function materialItemKey_(rowData) {
  return [
    materialClean_(rowData[0]).toUpperCase(),
    materialClean_(rowData[1]).toUpperCase(),
    materialClean_(rowData[2]).toUpperCase(),
    materialClean_(rowData[3]).toUpperCase(),
    materialClean_(rowData[6]).toUpperCase(),
    materialClean_(rowData[7]).toUpperCase()
  ].join('|');
}

function materialUnique_(arr) {
  return [...new Set(
    arr
      .filter(function(value) {
        return materialClean_(value) !== '';
      })
      .map(function(value) {
        return materialClean_(value);
      })
  )].sort();
}

function materialClean_(value) {
  if (value === null || value === undefined) return '';
  return String(value).trim();
}
