/* =========================================================
   OUTWARD MODULE - MATERIAL EDIT + DELETE + WAREHOUSE
   Replace full Outward.gs with this file.
========================================================= */

function getOutwardData() {
  return {
    requests: getOutwardRequests_(),
    masterData: getOutwardMaterialRows_()
  };
}

function getOutwardRequests_() {
  const sheet = getOutwardSheet_();
  ensureOutwardStatusHeader_(sheet);

  const values = sheet.getDataRange().getDisplayValues();
  if (values.length < 2) return [];

  const headers = values[0].map(outwardClean_);
  const data = [];

  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    if (row.every(function(cell) { return outwardClean_(cell) === ''; })) continue;

    const obj = { rowNumber: r + 1 };
    headers.forEach(function(header, c) {
      if (header) obj[header] = outwardClean_(row[c]);
    });
    data.push(obj);
  }

  return data;
}

function getOutwardMaterialRows_() {
  try {
    if (typeof getMaterialRows_ === 'function') {
      return getMaterialRows_().map(function(row) {
        return {
          rowNumber: row.rowNumber,
          goodsType: outwardClean_(row.goodsType),
          mainGroup: outwardClean_(row.mainGroup),
          subGroup: outwardClean_(row.subGroup),
          itemName: outwardClean_(row.itemName || row.goodsDescription),
          uom: outwardClean_(row.uom),
          brand: outwardClean_(row.brand),
          model: outwardClean_(row.model)
        };
      });
    }
  } catch (err) {
    Logger.log('getMaterialRows_ failed in getOutwardMaterialRows_: ' + err);
  }

  const ss = getOutwardSpreadsheet_();
  const sheetName = typeof ERP_CONFIG !== 'undefined' && ERP_CONFIG.MASTER_SHEET_NAME
    ? ERP_CONFIG.MASTER_SHEET_NAME
    : 'Material List';
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];

  const values = sheet.getDataRange().getDisplayValues();
  if (values.length < 2) return [];

  const rows = [];
  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    if (row.every(function(cell) { return outwardClean_(cell) === ''; })) continue;
    rows.push({
      rowNumber: r + 1,
      goodsType: outwardClean_(row[0]),
      mainGroup: outwardClean_(row[1]),
      subGroup: outwardClean_(row[2]),
      itemName: outwardClean_(row[3]),
      uom: outwardClean_(row[4]),
      brand: outwardClean_(row[6]),
      model: outwardClean_(row[7])
    });
  }
  return rows;
}

function generateOutwardTransactionId_(sheet, headPerson) {
  const headers = getOutwardHeaders_(sheet);
  const idCol = findOutwardHeaderIndex_(headers, ['Transaction_ID', 'Transaction ID', 'TRANSACTION_ID']);
  let maxNum = 0;

  if (idCol !== -1 && sheet.getLastRow() > 1) {
    const values = sheet.getRange(2, idCol + 1, sheet.getLastRow() - 1, 1).getDisplayValues();
    values.forEach(function(row) {
      const match = String(row[0] || '').match(/_Request_(\d+)\s*$/i);
      if (match) {
        const num = parseInt(match[1], 10);
        if (num > maxNum) maxNum = num;
      }
    });
  }

  const namePart = outwardClean_(headPerson).split(/\s+/)[0] || 'Store';
  return namePart + '_Request_' + (maxNum + 1);
}

function createOutwardTransaction(payload) {
  payload = payload || {};

  const headPerson = outwardClean_(payload.headPerson);
  const location = outwardClean_(payload.location);
  const purpose = outwardClean_(payload.purpose);
  const technician = outwardClean_(payload.technician);
  const items = Array.isArray(payload.items) ? payload.items : [];

  if (!headPerson) throw new Error('Head Person is required');
  if (!location) throw new Error('Location is required');
  if (!purpose) throw new Error('Purpose is required');
  if (!items.length) throw new Error('Kam se kam ek item add karo');

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);

  try {
    const sheet = getOutwardSheet_();
    ensureOutwardStatusHeader_(sheet);
    const headers = getOutwardHeaders_(sheet);
    const lastCol = sheet.getLastColumn();

    const transactionId = generateOutwardTransactionId_(sheet, headPerson);
    const timeZone = Session.getScriptTimeZone() || 'Asia/Kolkata';
    const displayDate = Utilities.formatDate(new Date(), timeZone, 'dd/MM/yyyy');

    function updateHeaderCol(rowData, headerNames, val) {
      const idx = findOutwardHeaderIndex_(headers, headerNames);
      if (idx !== -1) rowData[idx] = val;
    }

    const newRows = items.map(function(item) {
      const rowData = new Array(lastCol).fill('');

      updateHeaderCol(rowData, ['Transaction_ID', 'Transaction ID', 'TRANSACTION_ID'], transactionId);
      updateHeaderCol(rowData, ['Type', 'TYPE'], 'REQUESTED');
      updateHeaderCol(rowData, ['Date', 'DATE'], displayDate);
      updateHeaderCol(rowData, ['Location', 'LOCATION'], location);
      updateHeaderCol(rowData, ['Purpose', 'PURPOSE'], purpose);
      updateHeaderCol(rowData, ['Head_Person', 'Head Person', 'HEAD_PERSON'], headPerson);
      updateHeaderCol(rowData, ['Technician', 'TECHNICIAN'], technician);

      updateOutwardRowData_(headers, rowData, item);
      return rowData;
    });

    sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, lastCol).setValues(newRows);
    SpreadsheetApp.flush();

    return {
      success: true,
      message: 'New transaction created: ' + transactionId,
      transactionId: transactionId,
      requests: getOutwardRequests_(),
      masterData: getOutwardMaterialRows_()
    };
  } finally {
    lock.releaseLock();
  }
}

function updateOutwardIssue(payload) {
  payload = payload || {};

  const sheet = getOutwardSheet_();
  const rowNumber = Number(payload.rowNumber);
  if (!rowNumber || rowNumber < 2) throw new Error('Invalid outward row number');

  ensureOutwardStatusHeader_(sheet);
  const headers = getOutwardHeaders_(sheet);

  const lastCol = sheet.getLastColumn();
  const rowRange = sheet.getRange(rowNumber, 1, 1, lastCol);
  const rowData = rowRange.getValues()[0];

  updateOutwardRowData_(headers, rowData, payload);

  rowRange.setValues([rowData]);
  SpreadsheetApp.flush();

  return {
    success: true,
    message: 'Outward request updated successfully',
    requests: getOutwardRequests_(),
    masterData: getOutwardMaterialRows_()
  };
}

function updateOutwardManualStatus(payload) {
  payload = payload || {};

  const sheet = getOutwardSheet_();
  const rowNumber = Number(payload.rowNumber);
  const status = outwardClean_(payload.status).toUpperCase();

  if (!rowNumber || rowNumber < 2) throw new Error('Invalid outward row number');
  if (['PENDING', 'PARTIAL', 'SUCCESS'].indexOf(status) === -1) {
    throw new Error('Invalid status. Allowed: PENDING, PARTIAL, SUCCESS');
  }

  ensureOutwardStatusHeader_(sheet);
  const headers = getOutwardHeaders_(sheet);
  const colIndex = findOutwardHeaderIndex_(headers, ['ERP STATUS']);

  if (colIndex !== -1) {
    sheet.getRange(rowNumber, colIndex + 1).setValue(status);
  }

  SpreadsheetApp.flush();

  return {
    success: true,
    message: 'Status updated to ' + status,
    requests: getOutwardRequests_(),
    masterData: getOutwardMaterialRows_()
  };
}
function updateOutwardGroupIssue(payload) {
  payload = payload || {};
  if (!payload.items || !payload.items.length) throw new Error('No transaction items received');

  var lock = LockService.getScriptLock();
  
  // ✅ Timeout badhao - 30 seconds tak wait karo
  try {
    lock.waitLock(30000); // 30 seconds max
  } catch(e) {
    throw new Error('Another update is in progress. Please wait a moment and try again.');
  }

  try {
    var sheet = getOutwardSheet_();
    ensureOutwardStatusHeader_(sheet);

    var headers = getOutwardHeaders_(sheet);
    var lastCol = sheet.getLastColumn();

    var deleteRows = [];
    var updateItems = [];

    // Separate delete and update items
    payload.items.forEach(function(item) {
      var rowNumber = Number(item.rowNumber);
      if (!rowNumber || rowNumber < 2) return;

      if (item.deleteRow === true || 
          String(item.deleteRow).toUpperCase() === 'YES' || 
          String(item.deleteRow).toUpperCase() === 'TRUE') {
        deleteRows.push(rowNumber);
      } else {
        updateItems.push(item);
      }
    });

    // First: Delete rows (bottom to top to maintain row numbers)
    if (deleteRows.length > 0) {
      // Sort descending to delete from bottom
      var uniqueDeleteRows = deleteRows
        .filter(function(rowNumber, index, arr) { 
          return arr.indexOf(rowNumber) === index; 
        })
        .sort(function(a, b) { 
          return b - a; 
        });

      uniqueDeleteRows.forEach(function(rowNumber) {
        if (rowNumber >= 2 && rowNumber <= sheet.getLastRow()) {
          sheet.deleteRow(rowNumber);
        }
      });
      
      // Re-read headers after deletion (columns might shift)
      headers = getOutwardHeaders_(sheet);
      lastCol = sheet.getLastColumn();
    }

    // Then: Update remaining items
    updateItems.forEach(function(item) {
      var rowNumber = Number(item.rowNumber);
      if (!rowNumber || rowNumber < 2 || rowNumber > sheet.getLastRow()) return;

      var rowRange = sheet.getRange(rowNumber, 1, 1, lastCol);
      var rowData = rowRange.getValues()[0];

      updateOutwardRowData_(headers, rowData, item);

      rowRange.setValues([rowData]);
      
      // ✅ Small delay between updates to prevent locking
      SpreadsheetApp.flush();
    });

    SpreadsheetApp.flush();

    var deletedMsg = deleteRows.length ? ' Deleted: ' + deleteRows.length + ' item(s).' : '';

    return {
      success: true,
      message: 'Transaction items updated successfully.' + deletedMsg,
      requests: getOutwardRequests_(),
      masterData: getOutwardMaterialRows_()
    };
    
  } catch(error) {
    throw new Error('Update failed: ' + (error.message || error.toString()));
  } finally {
    // ✅ Always release lock
    try {
      lock.releaseLock();
    } catch(e) {
      Logger.log('Error releasing lock: ' + e);
    }
  }
}

function deleteOutwardItem(payload) {
  payload = payload || {};
  const rowNumber = Number(payload.rowNumber);
  if (!rowNumber || rowNumber < 2) throw new Error('Invalid outward row number');

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);

  try {
    const sheet = getOutwardSheet_();
    if (rowNumber > sheet.getLastRow()) throw new Error('Outward item row not found');
    sheet.deleteRow(rowNumber);
    SpreadsheetApp.flush();

    return {
      success: true,
      message: 'Outward item deleted successfully',
      requests: getOutwardRequests_(),
      masterData: getOutwardMaterialRows_()
    };
  } finally {
    lock.releaseLock();
  }
}

function updateOutwardRowData_(headers, rowData, payload) {
  payload = payload || {};

  function updateCol(headerNames, val) {
    const idx = findOutwardHeaderIndex_(headers, headerNames);
    if (idx !== -1) rowData[idx] = val;
  }

  if (payload.goodsType !== undefined) updateCol(['Goods_Type', 'Goods Type', 'GOODS TYPE'], outwardClean_(payload.goodsType).toUpperCase());
  if (payload.mainGroup !== undefined) updateCol(['Main_Group', 'Main Group', 'MAIN GROUP'], outwardClean_(payload.mainGroup));
  if (payload.subGroup !== undefined) updateCol(['Sub_Group', 'Sub Group', 'SUB GROUP'], outwardClean_(payload.subGroup));
  if (payload.itemName !== undefined) updateCol(['Item_Name', 'Item Name', 'GOODS DESCRIPTION', 'Goods Description'], outwardClean_(payload.itemName));
  if (payload.uom !== undefined) updateCol(['UOM'], outwardClean_(payload.uom).toUpperCase());
  if (payload.model !== undefined) updateCol(['Model', 'MODEL'], outwardClean_(payload.model));
  if (payload.quantity !== undefined) updateCol(['Quantity', 'Req Qty', 'REQUEST QTY', 'Requested Qty'], payload.quantity);

  if (payload.issueDate !== undefined) updateCol(['Issue Date'], toOutwardDisplayDateSlash_(payload.issueDate));
  if (payload.issueQty !== undefined) updateCol(['Issue Qty'], payload.issueQty);
  if (payload.issueFrom !== undefined) updateCol(['Issue From'], outwardClean_(payload.issueFrom).toUpperCase());
  if (payload.stockType !== undefined) updateCol(['STOCK TYPE', 'Stock Type'], outwardClean_(payload.stockType).toUpperCase());
  if (payload.brand !== undefined) updateCol(['Brand', 'BRAND'], outwardClean_(payload.brand));
  if (payload.serialNumber !== undefined) updateCol(['Serial Number', 'GOODS SERIAL NUMBER', 'Serial No'], outwardClean_(payload.serialNumber));
  if (payload.dcMi !== undefined) updateCol(['DC / MI', 'DC/MI', 'DC MI'], outwardClean_(payload.dcMi));

  const status = outwardClean_(payload.status).toUpperCase();
  if (['PENDING', 'PARTIAL', 'SUCCESS'].indexOf(status) !== -1) {
    updateCol(['ERP STATUS'], status);
  }
}

function ensureOutwardStatusHeader_(sheet) {
  const lastColumn = sheet.getLastColumn();
  if (lastColumn === 0) {
    sheet.getRange(1, 1).setValue('ERP STATUS');
    return;
  }

  const headers = sheet.getRange(1, 1, 1, lastColumn).getDisplayValues()[0].map(outwardClean_);
  const existing = findOutwardHeaderIndex_(headers, [
    'ERP STATUS',
    'ERP_STATUS',
    'Manual Status',
    'MANUAL STATUS',
    'Outward Status',
    'OUTWARD STATUS'
  ]);

  if (existing !== -1) {
    if (headers[existing] !== 'ERP STATUS') sheet.getRange(1, existing + 1).setValue('ERP STATUS');
    return;
  }

  sheet.getRange(1, lastColumn + 1).setValue('ERP STATUS');
  sheet.getRange(1, lastColumn + 1).setFontWeight('bold').setBackground('#d9ead3');
}

function getOutwardSheet_() {
  const sheetName = typeof ERP_CONFIG !== 'undefined' && ERP_CONFIG.OUTWARD_SHEET_NAME
    ? ERP_CONFIG.OUTWARD_SHEET_NAME
    : 'Logistics Database';

  if (typeof getSheet_ === 'function') return getSheet_(sheetName);

  const ss = getOutwardSpreadsheet_();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error('Sheet not found: ' + sheetName);
  return sheet;
}

function getOutwardSpreadsheet_() {
  if (typeof ERP_CONFIG !== 'undefined' && ERP_CONFIG.SPREADSHEET_ID) {
    return SpreadsheetApp.openById(ERP_CONFIG.SPREADSHEET_ID);
  }
  return SpreadsheetApp.getActiveSpreadsheet();
}

function getOutwardHeaders_(sheet) {
  return sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0].map(outwardClean_);
}

function findOutwardHeaderIndex_(headers, headerNames) {
  const normalizedHeaders = headers.map(function(header) {
    return outwardNormalizeHeader_(header);
  });

  for (let i = 0; i < headerNames.length; i++) {
    const idx = normalizedHeaders.indexOf(outwardNormalizeHeader_(headerNames[i]));
    if (idx !== -1) return idx;
  }

  return -1;
}

function outwardNormalizeHeader_(value) {
  return outwardClean_(value)
    .toUpperCase()
    .replace(/[_\s]+/g, ' ')
    .replace(/\s*\/\s*/g, ' / ')
    .trim();
}

function outwardClean_(value) {
  if (value === null || value === undefined) return '';
  return String(value).trim();
}

function toOutwardDisplayDateSlash_(inputDate) {
  if (!inputDate) return '';
  const value = String(inputDate);

  if (value.includes('-')) {
    const parts = value.split('-');
    if (parts[0].length === 4) return parts[2] + '/' + parts[1] + '/' + parts[0];
  }

  return value;
}
