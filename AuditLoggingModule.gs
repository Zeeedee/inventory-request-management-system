const LOG_CONFIG = {
  get SHEET_NAMES() {
    return {
      CHANGE_LOGS: APP_CONFIG.sheets.changeLogs,
      ENCODER_LOGS: APP_CONFIG.sheets.encoderLogs,
      REQUEST_FORM_LOGS: APP_CONFIG.sheets.requestLogs,
      FINALIZED_REQUEST_LOGS: APP_CONFIG.sheets.finalizedRequestLogs
    };
  },
  get HEADERS() {
    return APP_CONFIG.logs.headers;
  },
  get REQUEST_FORM_HEADERS() {
    return ['Timestamp', 'Requested By', 'Actor Email', 'Action', 'Item Name', 'Code', 'Quantity', 'Remarks', 'Source'];
  }
};

function LOG_getHeadersForSheet_(sheetName) {
  return sheetName === LOG_CONFIG.SHEET_NAMES.REQUEST_FORM_LOGS
    ? LOG_CONFIG.REQUEST_FORM_HEADERS
    : LOG_CONFIG.HEADERS;
}

function LOG_getLayout_() {
  return APP_CONFIG.logs.layout;
}

function LOG_getSheetNote_(sheetName) {
  if (sheetName === LOG_CONFIG.SHEET_NAMES.CHANGE_LOGS) {
    return APP_CONFIG.logs.notes.changeLogs;
  }
  if (sheetName === LOG_CONFIG.SHEET_NAMES.ENCODER_LOGS) {
    return APP_CONFIG.logs.notes.encoderLogs;
  }
  if (sheetName === LOG_CONFIG.SHEET_NAMES.REQUEST_FORM_LOGS) {
    return APP_CONFIG.logs.notes.requestLogs;
  }
  if (sheetName === LOG_CONFIG.SHEET_NAMES.FINALIZED_REQUEST_LOGS) {
    return APP_CONFIG.logs.notes.finalizedRequestLogs;
  }
  return APP_CONFIG.logs.notes.default;
}

function LOG_getFriendlyAction_(action) {
  return APP_CONFIG.logs.actions[action] || APP_safeString_(action);
}

function LOG_getFriendlySource_(source) {
  return APP_CONFIG.logs.sources[source] || APP_safeString_(source);
}

function LOG_getDisplayTimestamp_(date) {
  const parts = APP_getTimestampParts_(date || APP_now_());
  return parts.date + ' ' + parts.time;
}

function LOG_buildLogRow_(entry) {
  const safeEntry = entry || {};
  return [
    LOG_getDisplayTimestamp_(safeEntry.timestamp),
    APP_safeString_(safeEntry.requestedBy) || APP_getActor_().display,
    LOG_getFriendlyAction_(safeEntry.action),
    APP_safeString_(safeEntry.itemName),
    APP_isNoCode(safeEntry.code) ? 'NO CODE' : APP_safeString_(safeEntry.code),
    APP_safeString_(safeEntry.quantity),
    APP_safeString_(safeEntry.remarks),
    LOG_getFriendlySource_(safeEntry.source)
  ];
}

function LOG_buildRequestFormLogRow_(entry) {
  const safeEntry = entry || {};
  return [
    LOG_getDisplayTimestamp_(safeEntry.timestamp),
    APP_safeString_(safeEntry.requestedBy) || APP_getActor_().display,
    APP_safeString_(safeEntry.actorEmail),
    LOG_getFriendlyAction_(safeEntry.action),
    APP_safeString_(safeEntry.itemName),
    APP_isNoCode(safeEntry.code) ? 'NO CODE' : APP_safeString_(safeEntry.code),
    APP_safeString_(safeEntry.quantity),
    APP_safeString_(safeEntry.remarks),
    LOG_getFriendlySource_(safeEntry.source)
  ];
}

function LOG_buildLogRowForSheet_(sheetName, entry) {
  return sheetName === LOG_CONFIG.SHEET_NAMES.REQUEST_FORM_LOGS
    ? LOG_buildRequestFormLogRow_(entry)
    : LOG_buildLogRow_(entry);
}

function LOG_migrateLegacySheetStructure_(sheet, headers) {
  const layout = LOG_getLayout_();
  const topRow = sheet.getLastRow() > 0
    ? sheet.getRange(1, 1, 1, headers.length).getValues()[0]
    : [];
  const headerRow = sheet.getLastRow() >= layout.headerRow
    ? sheet.getRange(layout.headerRow, 1, 1, headers.length).getValues()[0]
    : [];
  const legacyHeaderDetected = APP_headersLookValid_(topRow, headers) && !APP_headersLookValid_(headerRow, headers);

  if (legacyHeaderDetected) {
    sheet.insertRowsBefore(1, layout.headerRow - 1);
  }
}

function LOG_applyBodyFormatting_(sheet, rowCount) {
  const headers = LOG_getHeadersForSheet_(sheet.getName());
  const layout = LOG_getLayout_();
  const safeRowCount = Math.max(rowCount || 1, 1);
  const backgrounds = [];

  for (let rowIndex = 0; rowIndex < safeRowCount; rowIndex += 1) {
    const color = rowIndex % 2 === 0 ? APP_CONFIG.system.colors.rowAlt1 : APP_CONFIG.system.colors.rowAlt2;
    const rowColors = [];
    for (let columnIndex = 0; columnIndex < headers.length; columnIndex += 1) {
      rowColors.push(color);
    }
    backgrounds.push(rowColors);
  }

  sheet.getRange(layout.dataStartRow, 1, safeRowCount, headers.length)
    .setBackgrounds(backgrounds)
    .setWrap(true)
    .setVerticalAlignment('middle')
    .setBorder(true, true, true, true, true, true, APP_CONFIG.logs.colors.border, SpreadsheetApp.BorderStyle.SOLID);
}

function LOG_applySheetFormatting_(sheet) {
  const headers = LOG_getHeadersForSheet_(sheet.getName());
  const layout = LOG_getLayout_();
  LOG_migrateLegacySheetStructure_(sheet, headers);

  try {
    sheet.getRange(1, 1, layout.headerRow, headers.length).breakApart();
  } catch (error) {
    console.log('Log unmerge skipped: ' + error.message);
  }

  sheet.getRange(layout.titleRow, 1, 1, headers.length).merge();
  sheet.getRange(layout.helpRow, 1, 1, headers.length).merge();
  sheet.getRange(layout.titleRow, 1)
    .setValue(sheet.getName())
    .setFontSize(15)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setBackground(APP_CONFIG.system.colors.titleBg)
    .setFontColor(APP_CONFIG.system.colors.titleText)
    .setBorder(true, true, true, true, true, true, APP_CONFIG.logs.colors.border, SpreadsheetApp.BorderStyle.SOLID);

  sheet.getRange(layout.helpRow, 1)
    .setValue(LOG_getSheetNote_(sheet.getName()))
    .setWrap(true)
    .setHorizontalAlignment('left')
    .setVerticalAlignment('middle')
    .setBackground(APP_CONFIG.system.colors.noteBg)
    .setBorder(true, true, true, true, true, true, APP_CONFIG.logs.colors.border, SpreadsheetApp.BorderStyle.SOLID);

  sheet.getRange(layout.spacerRow, 1, 1, headers.length)
    .clearContent()
    .setBackground('#ffffff')
    .setBorder(false, false, false, false, false, false);

  sheet.getRange(layout.headerRow, 1, 1, headers.length)
    .setValues([headers])
    .setFontWeight('bold')
    .setBackground(APP_CONFIG.logs.colors.headerBg)
    .setFontColor(APP_CONFIG.logs.colors.headerText)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrap(true)
    .setBorder(true, true, true, true, true, true, APP_CONFIG.logs.colors.border, SpreadsheetApp.BorderStyle.SOLID);

  sheet.setFrozenRows(layout.headerRow);
  sheet.setRowHeight(layout.titleRow, 28);
  sheet.setRowHeight(layout.helpRow, 24);
  sheet.setColumnWidth(1, 165);
  sheet.setColumnWidth(2, 180);
  if (sheet.getName() === LOG_CONFIG.SHEET_NAMES.REQUEST_FORM_LOGS) {
    sheet.setColumnWidth(3, 210);
    sheet.setColumnWidth(4, 165);
    sheet.setColumnWidth(5, 220);
    sheet.setColumnWidth(6, 140);
    sheet.setColumnWidth(7, 90);
    sheet.setColumnWidth(8, 280);
    sheet.setColumnWidth(9, 120);
  } else {
    sheet.setColumnWidth(3, 165);
    sheet.setColumnWidth(4, 220);
    sheet.setColumnWidth(5, 140);
    sheet.setColumnWidth(6, 90);
    sheet.setColumnWidth(7, 280);
    sheet.setColumnWidth(8, 120);
  }

  const dataRowCount = Math.max(sheet.getLastRow() - layout.dataStartRow + 1, 6, 1);
  LOG_applyBodyFormatting_(sheet, dataRowCount);
}

function LOG_ensureLogSheet(sheetName) {
  const ss = APP_getSpreadsheet_();
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }
  const headers = LOG_getHeadersForSheet_(sheetName);
  if (sheet.getMaxColumns() < headers.length) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), headers.length - sheet.getMaxColumns());
  }
  LOG_applySheetFormatting_(sheet);
  return sheet;
}

function LOG_appendEntries(sheetName, entries, options) {
  const safeEntries = entries || [];
  const settings = options || {};
  if (safeEntries.length === 0) {
    return [];
  }
  const sheet = LOG_ensureLogSheet(sheetName);
  const layout = LOG_getLayout_();
  const startRow = Math.max(sheet.getLastRow() + 1, layout.dataStartRow);
  const headers = LOG_getHeadersForSheet_(sheetName);
  const rows = safeEntries.map(function (entry) {
    return LOG_buildLogRowForSheet_(sheetName, entry);
  });
  const range = sheet.getRange(startRow, 1, rows.length, headers.length);
  range.setValues(rows);
  LOG_applyBodyFormatting_(sheet, Math.max(sheet.getLastRow() - layout.dataStartRow + 1, 1));
  if (!settings.skipSync) {
    try {
      const exportKey = typeof SYNC_getExportKeyForSheetName_ === 'function' ? SYNC_getExportKeyForSheetName_(sheetName) : '';
      if (!APP_isBlank_(exportKey) && typeof SYNC_syncExportKey_ === 'function') {
        SYNC_syncExportKey_(exportKey);
      }
    } catch (syncError) {
      console.log('View-only sync warning: ' + syncError.message);
    }
  }
  return rows.map(function (_, index) {
    return startRow + index;
  });
}

function LOG_appendEntry(sheetName, entry, options) {
  const rows = LOG_appendEntries(sheetName, [entry], options);
  return rows.length > 0 ? rows[0] : -1;
}

function LOG_appendChangeLog(entry, options) {
  const safeEntry = Object.assign({ source: 'INVENTORY' }, entry || {});
  return LOG_appendEntry(LOG_CONFIG.SHEET_NAMES.CHANGE_LOGS, safeEntry, options);
}

function LOG_encoderAction(action, searchTerm, resultsCount) {
  return LOG_appendEntry(LOG_CONFIG.SHEET_NAMES.ENCODER_LOGS, {
    requestedBy: APP_getActor_().display,
    action: action,
    itemName: APP_safeString_(searchTerm),
    code: '',
    quantity: resultsCount,
    remarks: 'Matches found: ' + APP_safeString_(resultsCount),
    source: 'ENCODER',
    timestamp: APP_now_()
  });
}

function LOG_requestFormAction(action, context, item, options) {
  const safeItem = item || {};
  return LOG_appendEntry(LOG_CONFIG.SHEET_NAMES.REQUEST_FORM_LOGS, {
    requestedBy: context && context.requestedBy ? context.requestedBy : APP_getActor_().display,
    actorEmail: context && context.actor && context.actor.email ? context.actor.email : APP_getActor_().email,
    action: action,
    itemName: safeItem.itemName || '',
    code: safeItem.code || '',
    quantity: safeItem.quantity || '',
    remarks: safeItem.remarks || '',
    source: 'REQUEST_FORM',
    timestamp: safeItem.timestamp || APP_now_()
  }, options);
}

function LOG_appendFinalizedRequestLogs(requestedBy, items, options) {
  const entries = (items || []).map(function (item) {
    return {
      requestedBy: requestedBy,
      action: item.action,
      itemName: item.itemName,
      code: item.code,
      quantity: item.quantity,
      remarks: item.remarks,
      source: 'REQUEST_FORM',
      timestamp: item.timestamp || APP_now_()
    };
  });
  return LOG_appendEntries(LOG_CONFIG.SHEET_NAMES.FINALIZED_REQUEST_LOGS, entries, options);
}

function LOG_deleteRows_(sheetName, rowNumbers) {
  const sheet = APP_getSpreadsheet_().getSheetByName(sheetName);
  if (!sheet) {
    return 0;
  }
  return APP_deleteRowsByNumbers_(sheet, rowNumbers || []);
}

function LOG_initializeAllLogs() {
  [
    LOG_CONFIG.SHEET_NAMES.CHANGE_LOGS,
    LOG_CONFIG.SHEET_NAMES.ENCODER_LOGS,
    LOG_CONFIG.SHEET_NAMES.REQUEST_FORM_LOGS,
    LOG_CONFIG.SHEET_NAMES.FINALIZED_REQUEST_LOGS
  ].forEach(function (sheetName) {
    LOG_ensureLogSheet(sheetName);
  });
  return APP_result_(true, 'Log sheets initialized.');
}

function LOG_setupInventoryEditTrigger() {
  APP_ensureTrigger_(APP_CONFIG.triggerHandlers.inventoryEdit, 'onEdit');
  return APP_result_(true, 'Inventory edit trigger installed.');
}

function LOG_handleInventoryEdit(e) {
  if (!e || !e.range) {
    return;
  }
  const sheet = e.range.getSheet();
  if (sheet.getName() !== APP_CONFIG.sheets.inventory) {
    return;
  }
  const rowNumber = e.range.getRow();
  const columnNumber = e.range.getColumn();
  if (columnNumber > e.range.getSheet().getLastColumn()) {
    return;
  }

  if (rowNumber === APP_CONFIG.inventory.headerRow) {
    try {
      syncRowToTarget(0);
    } catch (syncError) {
      console.log('View-only sync warning: ' + syncError.message);
    }
    return;
  }

  if (rowNumber < APP_CONFIG.inventory.startRow || columnNumber > APP_CONFIG.inventory.lastDataColumn) {
    return;
  }

  const oldValue = e.oldValue === undefined ? '' : e.oldValue;
  const newValue = e.range.getValue();
  if (APP_safeString_(oldValue) === APP_safeString_(newValue)) {
    return;
  }

  const record = APP_getInventoryRecordByRowNumber_(sheet, rowNumber, true);
  if (!record) {
    return;
  }

  const editTimestamp = APP_now_();
  if (columnNumber === APP_CONFIG.inventory.cols.QTY) {
    APP_writeInventoryStockRows_(sheet, [{
      rowNumber: rowNumber,
      qty: APP_toNumber_(newValue, 0),
      price: record.price,
      dateInv: editTimestamp
    }]);
  }

  APP_invalidateInventorySearchCaches_();
  const headers = APP_getHeaderValues_(sheet, APP_CONFIG.inventory.headerRow, APP_CONFIG.inventory.lastDataColumn);
  const columnName = headers[columnNumber - 1] || ('Column ' + columnNumber);
  LOG_appendChangeLog({
    requestedBy: APP_getActor_().display,
    action: 'MANUAL_EDIT',
    itemName: record.name,
    code: record.code,
    quantity: columnNumber === APP_CONFIG.inventory.cols.QTY ? newValue : '',
    remarks: columnName + ': ' + APP_safeString_(oldValue) + ' -> ' + APP_safeString_(newValue),
    source: 'INVENTORY',
    timestamp: editTimestamp
  });
  try {
    syncRowToTarget(rowNumber);
  } catch (syncError) {
    console.log('View-only sync warning: ' + syncError.message);
  }
}
