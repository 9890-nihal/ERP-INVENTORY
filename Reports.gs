/* =========================
   REPORTS MODULE - PROFESSIONAL LIVE REPORTS
   Replace full Reports.gs with this file.
========================= */

function getReportsInitialData(reportType, filters) {
  filters = filters || {};
  const type = cleanReportValue_(reportType || 'dashboard');
  const periodInfo = getReportPeriodInfo_(filters.period);

  let response;

  if (type === 'stockSummary') {
    response = buildStockSummaryReport_(periodInfo, filters);
  } else if (type === 'reorder') {
    response = buildReorderReport_(periodInfo, filters);
  } else if (type === 'itemLedger') {
    response = buildItemLedgerReport_(periodInfo, filters);
  } else if (type === 'inwardRegister') {
    response = buildInwardRegisterReport_(periodInfo, filters);
  } else if (type === 'outwardRegister') {
    response = buildOutwardRegisterReport_(periodInfo, filters);
  } else if (type === 'assetSerial') {
    response = buildAssetSerialReport_(periodInfo, filters);
  } else if (type === 'exception') {
    response = buildExceptionReport_(periodInfo, filters);
  } else {
    response = buildDashboardReport_(periodInfo, filters);
  }

  response.success = true;
  response.reportType = type;
  response.period = periodInfo.period;
  response.periodLabel = periodInfo.label;
  response.fromDate = formatReportDate_(periodInfo.start);
  response.toDate = formatReportDate_(new Date(periodInfo.end.getTime() - 24 * 60 * 60 * 1000));
  response.generatedAt = formatReportDateTime_(new Date());

  return response;
}

function buildDashboardReport_(periodInfo, filters) {
  const inventory = getReportInventoryData_(periodInfo.period);
  const outward = getReportOutwardRows_();
  const inward = getReportInwardRows_();

  const periodOutward = outward.filter(function(row) {
    return isReportDateInRange_(row.movementDateObj, periodInfo.start, periodInfo.end);
  });

  const periodInward = inward.filter(function(row) {
    return isReportDateInRange_(row.dateObj, periodInfo.start, periodInfo.end);
  });

  const pendingRequests = outward.filter(function(row) {
    return !row.isReturn && row.status !== 'SUCCESS';
  }).length;

  const returnedRows = periodOutward.filter(function(row) {
    return row.isReturn;
  });

  const totalReturnQty = returnedRows.reduce(function(sum, row) {
    return sum + row.qty;
  }, 0);

  const totalIssueQty = periodOutward.filter(function(row) {
    return !row.isReturn;
  }).reduce(function(sum, row) {
    return sum + row.issueQty;
  }, 0);

  const totalInwardQty = periodInward.reduce(function(sum, row) {
    return sum + row.qty;
  }, 0);

  const stats = inventory.stats || {};

  const rows = [
    { metric: 'Total Inventory Items', value: stats.totalItems || 0, detail: 'Sheet2 item master based unique items', status: 'INFO' },
    { metric: 'Current Stock Qty', value: stats.totalCurrentQty || 0, detail: 'Current stock level for selected month', status: 'INFO' },
    { metric: 'Monthly Inward Qty', value: totalInwardQty, detail: 'Sheet3 inward entries for selected month', status: 'OK' },
    { metric: 'Monthly Outward Qty', value: totalIssueQty, detail: 'Sheet1 issued quantity for selected month', status: 'OK' },
    { metric: 'Return Qty', value: totalReturnQty, detail: 'RETURNED rows added back to inventory', status: 'RETURN' },
    { metric: 'Pending Requests', value: pendingRequests, detail: 'Requested rows not fully issued', status: pendingRequests > 0 ? 'ACTION' : 'OK' },
    { metric: 'Reorder Required', value: stats.reorderRequired || 0, detail: 'Current stock <= minimum stock qty', status: (stats.reorderRequired || 0) > 0 ? 'ACTION' : 'OK' },
    { metric: 'Negative Stock Items', value: stats.negativeStockItems || 0, detail: 'Items where current stock is below zero', status: (stats.negativeStockItems || 0) > 0 ? 'CRITICAL' : 'OK' }
  ];

  return {
    title: 'Dashboard Summary',
    subtitle: 'Management overview for inventory, inward, outward, returns and reorder control.',
    columns: ['Metric', 'Value', 'Detail', 'Status'],
    rows: rows,
    kpis: [
      makeKpi_('Total Items', stats.totalItems || 0, 'Inventory master items'),
      makeKpi_('Current Stock Qty', stats.totalCurrentQty || 0, 'Closing stock for selected month'),
      makeKpi_('Inward / Outward', (totalInwardQty || 0) + ' / ' + (totalIssueQty || 0), 'Selected month movement'),
      makeKpi_('Reorder Required', stats.reorderRequired || 0, 'Items needing action'),
      makeKpi_('Pending Requests', pendingRequests, 'Open outward requests'),
      makeKpi_('Return Qty', totalReturnQty, 'Returned stock added back')
    ],
    footerNote: 'Dashboard is calculated live from Sheet1, Sheet2, Sheet3 and inventory logic.'
  };
}

function buildStockSummaryReport_(periodInfo, filters) {
  const inventory = getReportInventoryData_(periodInfo.period);
  const search = normalizeReportSearch_(filters.search);
  const rows = (inventory.rows || []).map(function(row) {
    return {
      goodsType: row.goodsType,
      mainGroup: row.mainGroup,
      subGroup: row.subGroup,
      itemName: row.goodsDescription,
      uom: row.uom,
      brand: row.brand,
      model: row.model,
      openingStock: totalReportStock_(row.opening),
      inward: totalReportStock_(row.inward),
      outward: totalReportStock_(row.outward),
      closingStock: row.currentStockLevel,
      minimumStock: row.minimumStockQty,
      reorderStatus: row.reorderIndicator,
      storageLocation: row.storageLocation || ''
    };
  }).filter(function(row) {
    return reportRowMatchesSearch_(row, search);
  });

  return {
    title: 'Stock Summary',
    subtitle: 'Opening, inward, outward and closing stock for selected month.',
    columns: ['Goods Type', 'Main Group', 'Sub Group', 'Item Name', 'UOM', 'Brand', 'Model', 'Opening Stock', 'Inward', 'Outward', 'Closing Stock', 'Minimum Stock', 'Reorder Status', 'Storage Location'],
    rows: rows,
    kpis: [
      makeKpi_('Total Items', rows.length, 'Visible records'),
      makeKpi_('Opening Stock', sumReportField_(rows, 'openingStock'), 'Opening balance'),
      makeKpi_('Inward Qty', sumReportField_(rows, 'inward'), 'Month inward'),
      makeKpi_('Outward Qty', sumReportField_(rows, 'outward'), 'Month outward'),
      makeKpi_('Closing Stock', sumReportField_(rows, 'closingStock'), 'Current stock'),
      makeKpi_('Reorder Required', rows.filter(function(row) { return row.reorderStatus === 'REORDER REQUIRED'; }).length, 'Action needed')
    ],
    footerNote: 'Return rows are adjusted in outward movement so returned stock is added back.'
  };
}

function buildReorderReport_(periodInfo, filters) {
  const inventory = getReportInventoryData_(periodInfo.period);
  const search = normalizeReportSearch_(filters.search);
  const rows = (inventory.rows || []).filter(function(row) {
    return row.reorderRequired;
  }).map(function(row) {
    const shortQty = Math.max((Number(row.minimumStockQty) || 0) - (Number(row.currentStockLevel) || 0), 0);
    return {
      goodsType: row.goodsType,
      mainGroup: row.mainGroup,
      subGroup: row.subGroup,
      itemName: row.goodsDescription,
      brand: row.brand,
      model: row.model,
      currentStock: row.currentStockLevel,
      minimumStock: row.minimumStockQty,
      shortQty: shortQty,
      storageLocation: row.storageLocation || '',
      suggestedAction: getReorderSuggestedAction_(row, shortQty)
    };
  }).filter(function(row) {
    return reportRowMatchesSearch_(row, search);
  });

  return {
    title: 'Reorder Required',
    subtitle: 'Items where current stock is less than or equal to minimum stock qty.',
    columns: ['Goods Type', 'Main Group', 'Sub Group', 'Item Name', 'Brand', 'Model', 'Current Stock', 'Minimum Stock', 'Short Qty', 'Storage Location', 'Suggested Action'],
    rows: rows,
    kpis: [
      makeKpi_('Reorder Items', rows.length, 'Items needing action'),
      makeKpi_('Total Short Qty', sumReportField_(rows, 'shortQty'), 'Qty below minimum'),
      makeKpi_('Zero/Negative Stock', rows.filter(function(row) { return Number(row.currentStock) <= 0; }).length, 'Urgent items'),
      makeKpi_('Blank Location', rows.filter(function(row) { return !row.storageLocation; }).length, 'Location missing')
    ],
    footerNote: 'Minimum stock comes from Sheet2 / Sheet4 manual minimum stock qty.'
  };
}

function buildInwardRegisterReport_(periodInfo, filters) {
  const search = normalizeReportSearch_(filters.search);
  const rows = getReportInwardRows_().filter(function(row) {
    return isReportDateInRange_(row.dateObj, periodInfo.start, periodInfo.end);
  }).map(function(row) {
    return {
      date: row.date,
      goodsType: row.goodsType,
      mainGroup: row.mainGroup,
      subGroup: row.subGroup,
      brand: row.brand,
      itemName: row.itemName,
      uom: row.uom,
      qty: row.qty,
      stockType: row.stockType,
      serialNumber: row.serialNumber,
      vendorPerson: row.vendorPerson,
      location: row.location,
      dcInvoiceNo: row.dcInvoiceNo,
      remark: row.remark
    };
  }).filter(function(row) {
    return reportRowMatchesSearch_(row, search);
  });

  return {
    title: 'Inward Register',
    subtitle: 'Vendor/person, DC invoice, stock type and item-wise inward register.',
    columns: ['Date', 'Goods Type', 'Main Group', 'Sub Group', 'Brand', 'Item Name', 'UOM', 'Qty', 'Stock Type', 'Serial Number', 'Vendor / Person', 'Location', 'DC / Invoice No', 'Remark'],
    rows: rows,
    kpis: [
      makeKpi_('Entries', rows.length, 'Selected month'),
      makeKpi_('Total Qty', sumReportField_(rows, 'qty'), 'Inward quantity'),
      makeKpi_('Asset Entries', rows.filter(function(row) { return String(row.goodsType).toUpperCase() === 'ASSET'; }).length, 'Asset records'),
      makeKpi_('Blank DC / Invoice', rows.filter(function(row) { return !row.dcInvoiceNo; }).length, 'Control check')
    ],
    footerNote: 'Data source: Sheet3 Inward.'
  };
}

function buildOutwardRegisterReport_(periodInfo, filters) {
  const search = normalizeReportSearch_(filters.search);
  const rows = getReportOutwardRows_().filter(function(row) {
    return isReportDateInRange_(row.movementDateObj, periodInfo.start, periodInfo.end);
  }).map(function(row) {
    return {
      transactionId: row.transactionId,
      type: row.type,
      linkedReqId: row.linkedReqId,
      date: row.date,
      issueDate: row.issueDate,
      location: row.location,
      purpose: row.purpose,
      headPerson: row.headPerson,
      technician: row.technician,
      goodsType: row.goodsType,
      mainGroup: row.mainGroup,
      subGroup: row.subGroup,
      itemName: row.itemName,
      requestedQty: row.requestQty,
      issueQty: row.issueQty,
      uom: row.uom,
      issueFrom: row.issueFrom,
      stockType: row.stockType,
      brand: row.brand,
      serialNumber: row.serialNumber,
      dcMi: row.dcMi,
      status: row.status
    };
  }).filter(function(row) {
    return reportRowMatchesSearch_(row, search);
  });

  return {
    title: 'Outward Register',
    subtitle: 'Location, purpose, head person, technician and item-wise outward/return register.',
    columns: ['Transaction ID', 'Type', 'Linked Req ID', 'Date', 'Issue Date', 'Location', 'Purpose', 'Head Person', 'Technician', 'Goods Type', 'Main Group', 'Sub Group', 'Item Name', 'Requested Qty', 'Issue Qty', 'UOM', 'Issue From', 'Stock Type', 'Brand', 'Serial Number', 'DC / MI', 'Status'],
    rows: rows,
    kpis: [
      makeKpi_('Entries', rows.length, 'Selected month'),
      makeKpi_('Issued Qty', sumReportField_(rows.filter(function(row) { return row.type !== 'RETURNED'; }), 'issueQty'), 'Issued movement'),
      makeKpi_('Return Qty', sumReportField_(rows.filter(function(row) { return row.type === 'RETURNED'; }), 'issueQty'), 'Returned stock'),
      makeKpi_('Pending', rows.filter(function(row) { return row.status === 'PENDING' || row.status === 'PARTIAL'; }).length, 'Open requests')
    ],
    footerNote: 'RETURNED rows use Linked_Req_ID to show against original request.'
  };
}

function buildItemLedgerReport_(periodInfo, filters) {
  const search = normalizeReportSearch_(filters.search);
  const transactions = [];

  getReportInwardRows_().forEach(function(row) {
    if (!isReportDateInRange_(row.dateObj, periodInfo.start, periodInfo.end)) return;
    transactions.push({
      sortDate: row.dateObj,
      itemKey: makeReportItemKey_(row.mainGroup, row.subGroup, row.itemName),
      date: row.date,
      transactionType: 'INWARD',
      referenceNo: row.dcInvoiceNo,
      itemName: row.itemName,
      mainGroup: row.mainGroup,
      subGroup: row.subGroup,
      stockType: row.stockType,
      inwardQty: row.qty,
      outwardQty: 0,
      balanceQty: 0,
      location: row.location,
      party: row.vendorPerson,
      remark: row.remark
    });
  });

  getReportOutwardRows_().forEach(function(row) {
    if (!isReportDateInRange_(row.movementDateObj, periodInfo.start, periodInfo.end)) return;
    const isReturn = row.isReturn;
    transactions.push({
      sortDate: row.movementDateObj,
      itemKey: makeReportItemKey_(row.mainGroup, row.subGroup, row.itemName),
      date: row.issueDate || row.date,
      transactionType: isReturn ? 'RETURN' : 'OUTWARD',
      referenceNo: row.transactionId,
      itemName: row.itemName,
      mainGroup: row.mainGroup,
      subGroup: row.subGroup,
      stockType: row.stockType,
      inwardQty: isReturn ? row.qty : 0,
      outwardQty: isReturn ? 0 : row.issueQty,
      balanceQty: 0,
      location: row.location,
      party: row.technician || row.headPerson,
      remark: isReturn ? ('Against: ' + row.linkedReqId) : row.purpose
    });
  });

  transactions.sort(function(a, b) {
    const ad = a.sortDate ? a.sortDate.getTime() : 0;
    const bd = b.sortDate ? b.sortDate.getTime() : 0;
    if (ad !== bd) return ad - bd;
    return String(a.referenceNo).localeCompare(String(b.referenceNo));
  });

  const balanceMap = {};
  transactions.forEach(function(row) {
    if (!balanceMap[row.itemKey]) balanceMap[row.itemKey] = 0;
    balanceMap[row.itemKey] += (Number(row.inwardQty) || 0) - (Number(row.outwardQty) || 0);
    row.balanceQty = balanceMap[row.itemKey];
  });

  const rows = transactions.filter(function(row) {
    return reportRowMatchesSearch_(row, search);
  });

  return {
    title: 'Item Ledger',
    subtitle: 'Item-wise inward, outward, return and running balance ledger for selected month.',
    columns: ['Date', 'Transaction Type', 'Reference No', 'Item Name', 'Main Group', 'Sub Group', 'Stock Type', 'Inward Qty', 'Outward Qty', 'Balance Qty', 'Location', 'Vendor / Issued To', 'Remark'],
    rows: rows,
    kpis: [
      makeKpi_('Transactions', rows.length, 'Ledger lines'),
      makeKpi_('Inward Qty', sumReportField_(rows, 'inwardQty'), 'Including returns'),
      makeKpi_('Outward Qty', sumReportField_(rows, 'outwardQty'), 'Issued qty'),
      makeKpi_('Unique Items', uniqueReportCount_(rows, 'itemName'), 'Visible items')
    ],
    footerNote: 'Use search box for one item to see cleaner item ledger and running balance.'
  };
}

function buildAssetSerialReport_(periodInfo, filters) {
  const search = normalizeReportSearch_(filters.search);
  const serialMap = {};

  getReportInwardRows_().forEach(function(row) {
    if (String(row.goodsType).toUpperCase() !== 'ASSET') return;
    if (!row.serialNumber) return;
    const key = row.serialNumber.toUpperCase();
    serialMap[key] = {
      assetItem: row.itemName,
      serialNumber: row.serialNumber,
      brand: row.brand,
      model: row.model,
      inwardDate: row.date,
      vendor: row.vendorPerson,
      currentLocation: row.location,
      issuedToTechnician: '',
      currentStatus: 'IN STOCK',
      lastReference: row.dcInvoiceNo
    };
  });

  getReportOutwardRows_().forEach(function(row) {
    if (String(row.goodsType).toUpperCase() !== 'ASSET') return;
    if (!row.serialNumber) return;
    const key = row.serialNumber.toUpperCase();
    if (!serialMap[key]) {
      serialMap[key] = {
        assetItem: row.itemName,
        serialNumber: row.serialNumber,
        brand: row.brand,
        model: '',
        inwardDate: '',
        vendor: '',
        currentLocation: row.location,
        issuedToTechnician: '',
        currentStatus: 'UNMAPPED ASSET',
        lastReference: row.transactionId
      };
    }

    serialMap[key].assetItem = serialMap[key].assetItem || row.itemName;
    serialMap[key].brand = serialMap[key].brand || row.brand;
    serialMap[key].currentLocation = row.location || serialMap[key].currentLocation;
    serialMap[key].issuedToTechnician = row.technician || row.headPerson || serialMap[key].issuedToTechnician;
    serialMap[key].lastReference = row.transactionId;
    serialMap[key].currentStatus = row.isReturn ? 'RETURNED / IN STOCK' : 'ISSUED';
  });

  const rows = Object.keys(serialMap).sort().map(function(key) {
    return serialMap[key];
  }).filter(function(row) {
    return reportRowMatchesSearch_(row, search);
  });

  return {
    title: 'Asset Serial Register',
    subtitle: 'Serial number level asset tracking from inward and outward movement.',
    columns: ['Asset Item', 'Serial Number', 'Brand', 'Model', 'Inward Date', 'Vendor', 'Current Location', 'Issued To / Technician', 'Current Status', 'Last Reference'],
    rows: rows,
    kpis: [
      makeKpi_('Total Serials', rows.length, 'Tracked assets'),
      makeKpi_('In Stock', rows.filter(function(row) { return row.currentStatus.indexOf('IN STOCK') !== -1; }).length, 'Available / returned'),
      makeKpi_('Issued', rows.filter(function(row) { return row.currentStatus === 'ISSUED'; }).length, 'Issued assets'),
      makeKpi_('Unmapped', rows.filter(function(row) { return row.currentStatus === 'UNMAPPED ASSET'; }).length, 'Needs inward link')
    ],
    footerNote: 'For asset accuracy, serial number should be filled in inward and outward rows.'
  };
}

function buildExceptionReport_(periodInfo, filters) {
  const rows = [];
  const inventory = getReportInventoryData_(periodInfo.period);

  (inventory.rows || []).forEach(function(row) {
    if (Number(row.currentStockLevel) < 0) {
      rows.push(makeExceptionRow_('Inventory', 'Negative Stock', row.goodsDescription, row.currentStockLevel, 'Check inward/outward quantity and return entries.'));
    }
    if ((Number(row.minimumStockQty) || 0) <= 0) {
      rows.push(makeExceptionRow_('Inventory', 'Blank Minimum Stock Qty', row.goodsDescription, row.minimumStockQty, 'Update minimum stock qty in Sheet2 or Sheet4.'));
    }
    if (!row.storageLocation) {
      rows.push(makeExceptionRow_('Inventory', 'Blank Storage Location', row.goodsDescription, '', 'Update rack/bin/location in Inventory screen.'));
    }
  });

  getReportInwardRows_().forEach(function(row) {
    if (!isReportDateInRange_(row.dateObj, periodInfo.start, periodInfo.end)) return;
    if (String(row.goodsType).toUpperCase() === 'ASSET' && !row.serialNumber) {
      rows.push(makeExceptionRow_('Inward', 'Asset Serial Missing', row.itemName, row.date, 'Asset inward must have serial number.'));
    }
    if (!row.dcInvoiceNo) {
      rows.push(makeExceptionRow_('Inward', 'DC / Invoice Missing', row.itemName, row.date, 'Fill delivery challan / invoice number.'));
    }
    if (!row.stockType) {
      rows.push(makeExceptionRow_('Inward', 'Stock Type Missing', row.itemName, row.date, 'Fill stock type.'));
    }
  });

  getReportOutwardRows_().forEach(function(row) {
    if (!isReportDateInRange_(row.movementDateObj, periodInfo.start, periodInfo.end)) return;
    if (row.issueQty > 0 && !row.issueDate) {
      rows.push(makeExceptionRow_('Outward', 'Issue Date Missing', row.itemName, row.transactionId, 'Fill issue date for issued request.'));
    }
    if (row.issueQty > 0 && !row.stockType) {
      rows.push(makeExceptionRow_('Outward', 'Stock Type Missing', row.itemName, row.transactionId, 'Fill stock type for issued request.'));
    }
    if (row.isReturn && !row.linkedReqId) {
      rows.push(makeExceptionRow_('Outward Return', 'Linked Req ID Missing', row.itemName, row.transactionId, 'Return row must have Linked_Req_ID.'));
    }
    if (String(row.goodsType).toUpperCase() === 'ASSET' && row.issueQty > 0 && !row.serialNumber) {
      rows.push(makeExceptionRow_('Outward', 'Asset Serial Missing', row.itemName, row.transactionId, 'Asset issue/return must have serial number.'));
    }
  });

  const search = normalizeReportSearch_(filters.search);
  const filteredRows = rows.filter(function(row) {
    return reportRowMatchesSearch_(row, search);
  });

  return {
    title: 'Exception Report',
    subtitle: 'Data mistakes, blank fields, negative stock and control exceptions.',
    columns: ['Area', 'Exception Type', 'Item / Reference', 'Value', 'Suggested Fix'],
    rows: filteredRows,
    kpis: [
      makeKpi_('Total Exceptions', filteredRows.length, 'Visible issues'),
      makeKpi_('Inventory Issues', filteredRows.filter(function(row) { return row.area === 'Inventory'; }).length, 'Stock/master issues'),
      makeKpi_('Inward Issues', filteredRows.filter(function(row) { return row.area === 'Inward'; }).length, 'Inward data issues'),
      makeKpi_('Outward Issues', filteredRows.filter(function(row) { return String(row.area).indexOf('Outward') !== -1; }).length, 'Outward/return issues')
    ],
    footerNote: 'Clear exceptions regularly to keep ERP reports reliable.'
  };
}

function getReportInventoryData_(period) {
  if (typeof getInventoryInitialData === 'function') {
    return getInventoryInitialData({ period: period, syncSheet: false });
  }

  return {
    rows: [],
    stats: {
      totalItems: 0,
      totalCurrentQty: 0,
      reorderRequired: 0,
      negativeStockItems: 0,
      totalInwardQty: 0,
      totalOutwardQty: 0
    }
  };
}

function getReportInwardRows_() {
  const sheet = getReportSheetSafe_(ERP_CONFIG.INWARD_SHEET_NAME);
  if (!sheet) return [];

  const values = sheet.getDataRange().getDisplayValues();
  if (values.length < 2) return [];

  const headers = values[0].map(cleanReportValue_);
  const cols = {
    date: findReportHeaderCol_(headers, ['DATE']),
    goodsType: findReportHeaderCol_(headers, ['GOODS TYPE', 'GOODS_TYPE']),
    mainGroup: findReportHeaderCol_(headers, ['MAIN GROUP', 'MAIN_GROUP']),
    subGroup: findReportHeaderCol_(headers, ['SUB GROUP', 'SUB_GROUP']),
    brand: findReportHeaderCol_(headers, ['BRAND']),
    itemName: findReportHeaderCol_(headers, ['GOODS DESCRIPTION', 'ITEM NAME', 'ITEM_NAME']),
    uom: findReportHeaderCol_(headers, ['UOM']),
    qty: findReportHeaderCol_(headers, ['INWARD QTY', 'INWARD_QTY', 'QTY']),
    model: findReportHeaderCol_(headers, ['MODEL']),
    stockType: findReportHeaderCol_(headers, ['STOCK TYPE', 'STOCK_TYPE']),
    serialNumber: findReportHeaderCol_(headers, ['GOODS SERIAL NUMBER', 'SERIAL NUMBER', 'SERIAL_NUMBER']),
    vendorPerson: findReportHeaderCol_(headers, ['VENDOR NAME / PERSON NAME', 'VENDOR NAME', 'PERSON NAME']),
    location: findReportHeaderCol_(headers, ['LOCATION']),
    dcInvoiceNo: findReportHeaderCol_(headers, ['DELIVERY CHALLAN/INVOICE NO', 'DC / INVOICE NO', 'INVOICE NO']),
    remark: findReportHeaderCol_(headers, ['REMARK'])
  };

  const rows = [];

  for (let r = 1; r < values.length; r++) {
    const source = values[r];
    if (isReportBlankRow_(source)) continue;

    const date = getReportCell_(source, cols.date);
    const itemName = getReportCell_(source, cols.itemName);
    const qty = getReportNumber_(getReportCell_(source, cols.qty));

    rows.push({
      rowNumber: r + 1,
      date: date,
      dateObj: parseReportDate_(date),
      goodsType: getReportCell_(source, cols.goodsType),
      mainGroup: getReportCell_(source, cols.mainGroup),
      subGroup: getReportCell_(source, cols.subGroup),
      brand: getReportCell_(source, cols.brand),
      itemName: itemName,
      uom: getReportCell_(source, cols.uom),
      qty: qty,
      model: getReportCell_(source, cols.model),
      stockType: getReportCell_(source, cols.stockType),
      serialNumber: getReportCell_(source, cols.serialNumber),
      vendorPerson: getReportCell_(source, cols.vendorPerson),
      location: getReportCell_(source, cols.location),
      dcInvoiceNo: getReportCell_(source, cols.dcInvoiceNo),
      remark: getReportCell_(source, cols.remark)
    });
  }

  return rows;
}

function getReportOutwardRows_() {
  const sheet = getReportSheetSafe_(ERP_CONFIG.OUTWARD_SHEET_NAME);
  if (!sheet) return [];

  const values = sheet.getDataRange().getDisplayValues();
  if (values.length < 2) return [];

  const headers = values[0].map(cleanReportValue_);
  const cols = {
    transactionId: findReportHeaderCol_(headers, ['Transaction_ID', 'Transaction ID', 'TRANSACTION_ID']),
    systemKey: findReportHeaderCol_(headers, ['System_Key', 'System Key', 'SYSTEM_KEY']),
    type: findReportHeaderCol_(headers, ['Type', 'TYPE', 'Status', 'STATUS']),
    date: findReportHeaderCol_(headers, ['Date', 'DATE']),
    location: findReportHeaderCol_(headers, ['Location', 'LOCATION']),
    purpose: findReportHeaderCol_(headers, ['Purpose', 'PURPOSE']),
    headPerson: findReportHeaderCol_(headers, ['Head_Person', 'Head Person', 'HEAD_PERSON']),
    technician: findReportHeaderCol_(headers, ['Technician', 'TECHNICIAN']),
    assetReason: findReportHeaderCol_(headers, ['Asset_Reason', 'Asset Reason', 'ASSET_REASON']),
    linkedReqId: findReportHeaderCol_(headers, ['Linked_Req_ID', 'Linked Req ID', 'LINKED_REQ_ID', 'Against Request ID']),
    goodsType: findReportHeaderCol_(headers, ['Goods_Type', 'Goods Type', 'GOODS_TYPE']),
    mainGroup: findReportHeaderCol_(headers, ['Main_Group', 'Main Group', 'MAIN_GROUP']),
    subGroup: findReportHeaderCol_(headers, ['Sub_Group', 'Sub Group', 'SUB_GROUP']),
    itemName: findReportHeaderCol_(headers, ['Item_Name', 'Item Name', 'ITEM_NAME']),
    quantity: findReportHeaderCol_(headers, ['Quantity', 'QUANTITY', 'Qty', 'QTY']),
    uom: findReportHeaderCol_(headers, ['UOM']),
    issueDate: findReportHeaderCol_(headers, ['Issue Date', 'Issue_Date', 'ISSUE_DATE']),
    issueQty: findReportHeaderCol_(headers, ['Issue Qty', 'Issue_Qty', 'ISSUE_QTY']),
    issueFrom: findReportHeaderCol_(headers, ['Issue From', 'Issue_From', 'ISSUE_FROM']),
    stockType: findReportHeaderCol_(headers, ['STOCK TYPE', 'Stock Type', 'Stock_Type', 'STOCK_TYPE']),
    brand: findReportHeaderCol_(headers, ['Brand', 'BRAND']),
    serialNumber: findReportHeaderCol_(headers, ['Serial Number', 'Serial_Number', 'SERIAL_NUMBER']),
    dcMi: findReportHeaderCol_(headers, ['DC / MI', 'DC/MI', 'DC_MI', 'DC MI'])
  };

  const rows = [];

  for (let r = 1; r < values.length; r++) {
    const source = values[r];
    if (isReportBlankRow_(source)) continue;

    const transactionId = getReportCell_(source, cols.transactionId);
    const typeRaw = getReportCell_(source, cols.type);
    const date = getReportCell_(source, cols.date);
    const issueDate = getReportCell_(source, cols.issueDate);
    const requestQty = getReportNumber_(getReportCell_(source, cols.quantity));
    const issueQtyRaw = getReportNumber_(getReportCell_(source, cols.issueQty));
    const issueQty = issueQtyRaw > 0 ? issueQtyRaw : 0;
    const dcMi = getReportCell_(source, cols.dcMi);
    const isReturn = isReportReturnRow_(typeRaw, transactionId, dcMi);
    const type = isReturn ? 'RETURNED' : (typeRaw || 'REQUESTED');
    const movementDate = issueDate || date;

    rows.push({
      rowNumber: r + 1,
      transactionId: transactionId,
      systemKey: getReportCell_(source, cols.systemKey),
      type: type,
      date: date,
      dateObj: parseReportDate_(date),
      issueDate: issueDate,
      movementDate: movementDate,
      movementDateObj: parseReportDate_(movementDate),
      location: getReportCell_(source, cols.location),
      purpose: getReportCell_(source, cols.purpose),
      headPerson: getReportCell_(source, cols.headPerson),
      technician: getReportCell_(source, cols.technician),
      assetReason: getReportCell_(source, cols.assetReason),
      linkedReqId: getReportCell_(source, cols.linkedReqId),
      goodsType: getReportCell_(source, cols.goodsType),
      mainGroup: getReportCell_(source, cols.mainGroup),
      subGroup: getReportCell_(source, cols.subGroup),
      itemName: getReportCell_(source, cols.itemName),
      requestQty: requestQty,
      issueQty: issueQty || requestQty,
      qty: issueQty || requestQty,
      uom: getReportCell_(source, cols.uom),
      issueFrom: getReportCell_(source, cols.issueFrom),
      stockType: getReportCell_(source, cols.stockType),
      brand: getReportCell_(source, cols.brand),
      serialNumber: getReportCell_(source, cols.serialNumber),
      dcMi: dcMi,
      isReturn: isReturn,
      status: getReportOutwardStatus_(requestQty, issueQty, isReturn)
    });
  }

  return rows;
}

function getReportOutwardStatus_(requestQty, issueQty, isReturn) {
  if (isReturn) return 'RETURNED';
  const req = Number(requestQty) || 0;
  const issued = Number(issueQty) || 0;
  if (req > 0 && issued >= req) return 'SUCCESS';
  if (issued > 0 && issued < req) return 'PARTIAL';
  return 'PENDING';
}

function isReportReturnRow_(typeRaw, transactionId, dcMi) {
  const type = cleanReportValue_(typeRaw).toUpperCase();
  const tx = cleanReportValue_(transactionId).toUpperCase();
  const dc = cleanReportValue_(dcMi).toUpperCase();
  return type === 'RETURNED' || tx.indexOf('_RETURN_') !== -1 || dc === 'RETURN';
}

function makeExceptionRow_(area, exceptionType, itemRef, value, suggestedFix) {
  return {
    area: area,
    exceptionType: exceptionType,
    itemReference: itemRef,
    value: value,
    suggestedFix: suggestedFix
  };
}

function getReorderSuggestedAction_(row, shortQty) {
  if (Number(row.currentStockLevel) <= 0) return 'Immediate purchase / inward required';
  if (String(row.goodsType).toUpperCase() === 'ASSET') return 'Asset repair / replacement review';
  if (shortQty > 0) return 'Purchase top-up qty: ' + shortQty;
  return 'Review minimum stock level';
}

function makeKpi_(label, value, hint) {
  return {
    label: label,
    value: value,
    hint: hint || ''
  };
}

function totalReportStock_(stockObj) {
  if (!stockObj) return 0;
  return Object.keys(stockObj).reduce(function(sum, key) {
    return sum + (Number(stockObj[key]) || 0);
  }, 0);
}

function sumReportField_(rows, field) {
  return rows.reduce(function(sum, row) {
    return sum + (Number(row[field]) || 0);
  }, 0);
}

function uniqueReportCount_(rows, field) {
  const map = {};
  rows.forEach(function(row) {
    const value = cleanReportValue_(row[field]);
    if (value) map[value.toUpperCase()] = true;
  });
  return Object.keys(map).length;
}

function reportRowMatchesSearch_(row, search) {
  if (!search) return true;
  const joined = Object.keys(row).map(function(key) {
    return row[key];
  }).join(' ').toLowerCase();
  return joined.indexOf(search) !== -1;
}

function normalizeReportSearch_(value) {
  return cleanReportValue_(value).toLowerCase();
}

function getReportPeriodInfo_(period) {
  const now = new Date();
  let yyyy = now.getFullYear();
  let mm = now.getMonth() + 1;

  if (period && /^\d{4}-\d{2}$/.test(String(period))) {
    const parts = String(period).split('-');
    yyyy = Number(parts[0]);
    mm = Number(parts[1]);
  }

  const start = new Date(yyyy, mm - 1, 1);
  const end = new Date(yyyy, mm, 1);
  const label = Utilities.formatDate(start, getReportTimeZone_(), 'MMM, yyyy');

  return {
    period: yyyy + '-' + String(mm).padStart(2, '0'),
    label: label,
    start: start,
    end: end
  };
}

function parseReportDate_(value) {
  if (!value) return null;
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    return value;
  }

  const text = cleanReportValue_(value);
  if (!text) return null;

  let parts;
  if (/^\d{4}-\d{2}-\d{2}$/.test(text)) {
    parts = text.split('-');
    return new Date(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2]));
  }

  if (/^\d{2}[\/\-]\d{2}[\/\-]\d{4}$/.test(text)) {
    parts = text.split(/[\/\-]/);
    return new Date(Number(parts[2]), Number(parts[1]) - 1, Number(parts[0]));
  }

  const parsed = new Date(text);
  if (!isNaN(parsed.getTime())) return parsed;
  return null;
}

function isReportDateInRange_(dateObj, startDate, endDate) {
  if (!dateObj) return false;
  return dateObj.getTime() >= startDate.getTime() && dateObj.getTime() < endDate.getTime();
}

function formatReportDate_(dateObj) {
  if (!dateObj) return '';
  return Utilities.formatDate(dateObj, getReportTimeZone_(), 'dd/MM/yyyy');
}

function formatReportDateTime_(dateObj) {
  if (!dateObj) return '';
  return Utilities.formatDate(dateObj, getReportTimeZone_(), 'dd/MM/yyyy HH:mm');
}

function getReportTimeZone_() {
  try {
    return Session.getScriptTimeZone() || 'Asia/Kolkata';
  } catch (e) {
    return 'Asia/Kolkata';
  }
}

function getReportSheetSafe_(sheetName) {
  try {
    return getSheet_(sheetName);
  } catch (e) {
    return null;
  }
}

function findReportHeaderCol_(headers, aliases) {
  const normalizedAliases = aliases.map(function(alias) {
    return normalizeReportHeader_(alias);
  });

  for (let i = 0; i < headers.length; i++) {
    if (normalizedAliases.indexOf(normalizeReportHeader_(headers[i])) !== -1) return i;
  }

  return -1;
}

function normalizeReportHeader_(value) {
  return cleanReportValue_(value).toUpperCase().replace(/[^A-Z0-9]/g, '');
}

function getReportCell_(row, colIndex) {
  if (colIndex === -1 || colIndex === null || colIndex === undefined) return '';
  return cleanReportValue_(row[colIndex]);
}

function getReportNumber_(value) {
  const number = Number(cleanReportValue_(value).replace(/,/g, ''));
  return isNaN(number) ? 0 : number;
}

function cleanReportValue_(value) {
  if (value === null || value === undefined) return '';
  return String(value).trim();
}

function isReportBlankRow_(row) {
  return row.every(function(cell) {
    return cleanReportValue_(cell) === '';
  });
}

function makeReportItemKey_(mainGroup, subGroup, itemName) {
  return [mainGroup, subGroup, itemName].map(function(value) {
    return cleanReportValue_(value).toUpperCase();
  }).join('|');
}
