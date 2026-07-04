/* =========================
   COMMON HELPERS
========================= */

function getSpreadsheet_() {
  if (ERP_CONFIG.SPREADSHEET_ID) {
    return SpreadsheetApp.openById(ERP_CONFIG.SPREADSHEET_ID);
  }

  return SpreadsheetApp.getActiveSpreadsheet();
}

function getSheet_(sheetName) {
  const ss = getSpreadsheet_();
  const sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    throw new Error('Sheet not found: ' + sheetName);
  }

  return sheet;
}

function setCellByHeader_(sheet, headers, rowNumber, headerName, value) {
  let colIndex = headers.indexOf(headerName);

  if (colIndex === -1) {
    const target = String(headerName).toUpperCase();
    colIndex = headers.findIndex(function(header) {
      return String(header).toUpperCase() === target;
    });
  }

  if (colIndex === -1) {
    throw new Error('Column not found: ' + headerName);
  }

  sheet.getRange(rowNumber, colIndex + 1).setValue(value);
}

function clean_(value) {
  if (value === null || value === undefined) return '';
  return String(value).trim();
}

function unique_(arr) {
  return [...new Set(
    arr
      .filter(function(v) {
        return clean_(v) !== '';
      })
      .map(function(v) {
        return clean_(v);
      })
  )].sort();
}

function toDisplayDateSlash_(inputDate) {
  if (!inputDate) return '';

  const value = String(inputDate);

  if (value.includes('-')) {
    const parts = value.split('-');

    if (parts[0].length === 4) {
      return parts[2] + '/' + parts[1] + '/' + parts[0];
    }
  }

  return value;
}

function toDisplayDateDash_(inputDate) {
  if (!inputDate) return '';

  const value = String(inputDate);

  if (value.includes('-')) {
    const parts = value.split('-');

    if (parts[0].length === 4) {
      return parts[2] + '-' + parts[1] + '-' + parts[0];
    }
  }

  return value;
}

function todayDash_() {
  const date = new Date();
  const dd = String(date.getDate()).padStart(2, '0');
  const mm = String(date.getMonth() + 1).padStart(2, '0');
  const yyyy = date.getFullYear();

  return dd + '-' + mm + '-' + yyyy;
}
