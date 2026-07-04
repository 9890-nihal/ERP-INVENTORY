/**
 * Old_API.gs
 * Purane Android app API code ko safe naam ke saath yahan rakha gaya hai.
 *
 * IMPORTANT:
 * Is file me doGet() / doPost() nahi hai.
 * Code.gs ka router in functions ko call karta hai:
 * - getMaterialListJsonApi_()
 * - saveLogisticsRowsApi_(e)
 */

function saveLogisticsRowsApi_(e) {
  try {
    var sheet = getOrCreateApiSheet_(getApiSheetName_('OUTWARD_SHEET_NAME', 'Logistics Database'));
    ensureLogisticsHeaders_(sheet);

    if (!e || !e.postData || !e.postData.contents) {
      return jsonOutput_({
        status: 'error',
        message: 'No POST data received'
      });
    }

    var data = JSON.parse(e.postData.contents);

    if (!Array.isArray(data)) {
      data = [data];
    }

    var existingData = sheet.getDataRange().getValues();
    var existingKeys = {};

    for (var i = 1; i < existingData.length; i++) {
      var key = cleanApi_(existingData[i][1]);
      if (key) {
        existingKeys[key] = i + 1;
      }
    }

    var inserted = 0;
    var updated = 0;
    var skipped = 0;

    for (var j = 0; j < data.length; j++) {
      var row = data[j];

      var transactionId = cleanApi_(row.id || row.transactionId || row.Transaction_ID);
      var itemName = cleanApi_(row.itemName || row.item || row.Item_Name);

      if (!transactionId || !itemName) {
        skipped++;
        continue;
      }

      var systemKey = cleanApi_(row.systemKey || row.System_Key) || transactionId + '_' + itemName;

      var rowData = [
        transactionId,
        systemKey,
        cleanApi_(row.type || row.Type || 'REQUESTED').toUpperCase(),
        cleanApi_(row.date || row.Date),
        cleanApi_(row.location || row.Location),
        cleanApi_(row.purpose || row.Purpose),
        cleanApi_(row.headPerson || row.Head_Person),
        cleanApi_(row.technician || row.Technician),
        cleanApi_(row.assetReason || row.Asset_Reason || 'N/A'),
        cleanApi_(row.linkedId || row.Linked_Req_ID || row.linkedReqId || 'N/A'),
        cleanApi_(row.goodsType || row.Goods_Type || 'CONSUMABLE').toUpperCase(),
        cleanApi_(row.mainGroup || row.Main_Group),
        cleanApi_(row.subGroup || row.Sub_Group),
        itemName,
        toNumberApi_(row.quantity || row.Quantity),
        cleanApi_(row.uom || row.UOM || 'NOS')
      ];

      if (existingKeys[systemKey]) {
        sheet.getRange(existingKeys[systemKey], 1, 1, rowData.length).setValues([rowData]);
        updated++;
      } else {
        sheet.appendRow(rowData);
        inserted++;
      }
    }

    SpreadsheetApp.flush();

    return jsonOutput_({
      status: 'success',
      inserted: inserted,
      updated: updated,
      skipped: skipped
    });

  } catch (error) {
    return jsonOutput_({
      status: 'error',
      message: String(error && error.message ? error.message : error)
    });
  }
}

function getMaterialListJsonApi_() {
  try {
    var sheet = SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(getApiSheetName_('MASTER_SHEET_NAME', 'Material List'));

    if (!sheet) {
      return jsonOutput_([]);
    }

    var data = sheet.getDataRange().getValues();
    var jsonArray = [];

    for (var i = 1; i < data.length; i++) {
      var row = data[i];

      var goodsType = cleanApi_(row[0]) || 'CONSUMABLE';
      var mainGroup = cleanApi_(row[1]);
      var subGroup = cleanApi_(row[2]);
      var item = cleanApi_(row[3]);
      var uom = cleanApi_(row[4]) || 'NOS';
      var moq = parseInt(row[5], 10) || 1;

      if (!mainGroup && !subGroup && !item) {
        continue;
      }

      jsonArray.push({
        goods_type: goodsType,
        main_group: mainGroup,
        sub_group: subGroup,
        item: item,
        uom: uom,
        moq: moq
      });
    }

    return jsonOutput_(jsonArray);

  } catch (error) {
    return jsonOutput_({
      status: 'error',
      message: String(error && error.message ? error.message : error)
    });
  }
}

function ensureLogisticsHeaders_(sheet) {
  var headers = [
    'Transaction_ID',
    'System_Key',
    'Type',
    'Date',
    'Location',
    'Purpose',
    'Head_Person',
    'Technician',
    'Asset_Reason',
    'Linked_Req_ID',
    'Goods_Type',
    'Main_Group',
    'Sub_Group',
    'Item_Name',
    'Quantity',
    'UOM',
    'Issue Date',
    'Issue Qty',
    'Issue From',
    'STOCK TYPE',
    'Brand',
    'Serial Number',
    'DC / MI'
  ];

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#d9ead3');
    return;
  }

  var lastColumn = sheet.getLastColumn();

  if (lastColumn < headers.length) {
    var missingHeaders = headers.slice(lastColumn);
    sheet.getRange(1, lastColumn + 1, 1, missingHeaders.length).setValues([missingHeaders]);
    sheet.getRange(1, lastColumn + 1, 1, missingHeaders.length).setFontWeight('bold').setBackground('#d9ead3');
  }
}

function getOrCreateApiSheet_(sheetName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }

  return sheet;
}

function getApiSheetName_(configKey, fallbackName) {
  try {
    if (
      typeof ERP_CONFIG !== 'undefined' &&
      ERP_CONFIG &&
      ERP_CONFIG[configKey]
    ) {
      return ERP_CONFIG[configKey];
    }
  } catch (err) {}

  return fallbackName;
}

function cleanApi_(value) {
  if (value === null || value === undefined) return '';
  return String(value).trim();
}

function toNumberApi_(value) {
  if (value === null || value === undefined || value === '') return '';
  var number = Number(String(value).replace(/,/g, ''));
  return isNaN(number) ? '' : number;
}

function jsonOutput_(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}
