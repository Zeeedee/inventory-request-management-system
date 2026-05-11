const APP_RUNTIME_CACHE_ = {
  namespaces: {}
};

function APP_getExecutionCache_(namespace) {
  const key = namespace || 'default';
  if (!APP_RUNTIME_CACHE_.namespaces[key]) {
    APP_RUNTIME_CACHE_.namespaces[key] = {};
  }
  return APP_RUNTIME_CACHE_.namespaces[key];
}

function APP_clearExecutionCache_(namespace) {
  const key = namespace || 'default';
  APP_RUNTIME_CACHE_.namespaces[key] = {};
}

function APP_createProfiler_(label) {
  const now = Date.now();
  return {
    label: label || 'profile',
    startMs: now,
    markMs: now,
    timings: {}
  };
}

function APP_profileMark_(profiler, key) {
  if (!profiler || !key) {
    return 0;
  }
  const now = Date.now();
  const elapsed = now - profiler.markMs;
  profiler.timings[key] = (profiler.timings[key] || 0) + elapsed;
  profiler.markMs = now;
  return elapsed;
}

function APP_profileLog_(profiler, extra) {
  if (!profiler) {
    return null;
  }
  const payload = Object.assign({}, profiler.timings, extra || {}, {
    totalTimeMs: Date.now() - profiler.startMs
  });
  console.log((profiler.label || 'profile') + ': ' + JSON.stringify(payload));
  return payload;
}

function APP_getSpreadsheet_() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function APP_getTimeZone_() {
  return Session.getScriptTimeZone() || SpreadsheetApp.getActive().getSpreadsheetTimeZone() || 'Asia/Manila';
}

function APP_now_() {
  return new Date();
}

function APP_safeString_(value) {
  return value === null || value === undefined ? '' : String(value);
}

function APP_isBlank_(value) {
  return APP_safeString_(value).trim() === '';
}

function APP_toNumber_(value, fallback) {
  const parsed = Number(value);
  return isNaN(parsed) ? (fallback === undefined ? 0 : fallback) : parsed;
}

function APP_normalizeText_(value) {
  return APP_safeString_(value).trim().replace(/\s+/g, ' ').toUpperCase();
}

function APP_isNoCode(code) {
  return APP_CONFIG.noCodeIndicators.indexOf(APP_normalizeText_(code)) !== -1;
}

function APP_splitCodes_(value) {
  if (APP_isBlank_(value) || APP_isNoCode(value)) {
    return [];
  }
  return APP_safeString_(value)
    .split(/[\/,;\n]+/)
    .map(function (part) {
      return APP_normalizeText_(part);
    })
    .filter(function (part) {
      return part !== '';
    });
}

function APP_getActor_() {
  let activeEmail = '';
  let effectiveEmail = '';
  let temporaryKey = '';

  try {
    activeEmail = Session.getActiveUser().getEmail() || '';
  } catch (error1) {
    activeEmail = '';
  }

  try {
    effectiveEmail = Session.getEffectiveUser().getEmail() || '';
  } catch (error2) {
    effectiveEmail = '';
  }

  try {
    temporaryKey = Session.getTemporaryActiveUserKey() || '';
  } catch (error3) {
    temporaryKey = '';
  }

  const email = activeEmail || effectiveEmail || '';
  const actorKey = email || temporaryKey || 'anonymous';
  return {
    display: email || ('user:' + actorKey),
    email: email,
    key: actorKey
  };
}

function APP_getTimestampParts_(date) {
  const value = date || APP_now_();
  const timeZone = APP_getTimeZone_();
  return {
    iso: Utilities.formatDate(value, timeZone, "yyyy-MM-dd'T'HH:mm:ss"),
    date: Utilities.formatDate(value, timeZone, 'MM/dd/yyyy'),
    time: Utilities.formatDate(value, timeZone, 'hh:mm:ss a')
  };
}

function APP_result_(ok, message, data) {
  return {
    ok: ok,
    message: message || '',
    data: data || null
  };
}

function APP_fail_(code, message, details) {
  return {
    ok: false,
    code: code,
    message: message,
    details: details || null
  };
}

function APP_createError_(code, message, details) {
  const error = new Error(message);
  error.code = code;
  error.details = details || null;
  return error;
}

function APP_getSheet_(sheetName, required) {
  const sheet = APP_getSpreadsheet_().getSheetByName(sheetName);
  if (!sheet && required) {
    throw APP_createError_('MISSING_SHEET', 'Required sheet not found: ' + sheetName, { sheetName: sheetName });
  }
  return sheet;
}

function APP_getStructureConfig_(structureKey) {
  const structure = APP_CONFIG.structures[structureKey];
  if (!structure) {
    throw APP_createError_('UNKNOWN_STRUCTURE', 'Unknown structure key: ' + structureKey);
  }
  return structure;
}

function APP_getSheetByStructure_(structureKey, required) {
  const structure = APP_getStructureConfig_(structureKey);
  return APP_getSheet_(structure.sheetName, required);
}

function APP_getLiveHeadersForStructure_(structureKey) {
  const structure = APP_getStructureConfig_(structureKey);
  const sheet = APP_getSheetByStructure_(structureKey, true);
  const width = structure.dynamicHeaderWidth
    ? Math.max(structure.columnCount || 0, sheet.getLastColumn())
    : (structure.columnCount || sheet.getLastColumn());
  if (width <= 0) {
    return [];
  }
  const headers = APP_getHeaderValues_(sheet, structure.headerRow, width);
  let effectiveWidth = headers.length;
  while (effectiveWidth > (structure.columnCount || 0) && APP_isBlank_(headers[effectiveWidth - 1])) {
    effectiveWidth -= 1;
  }
  return headers.slice(0, effectiveWidth);
}

function APP_getTemplateConfig_(templateKey) {
  return APP_CONFIG.templates[templateKey] || {};
}

function APP_getInventoryTemplateHeaders_() {
  const inventoryTemplate = APP_getTemplateConfig_('inventory');
  return (inventoryTemplate.displayHeaders || []).concat(inventoryTemplate.additionalDisplayHeaders || []);
}

function APP_getInventorySheet_() {
  return APP_getSheet_(APP_CONFIG.sheets.inventory, true);
}

function APP_getRequestSheet_() {
  return APP_getSheet_(APP_CONFIG.sheets.requestForm, true);
}

function APP_getEncoderSheet_() {
  return APP_getSheet_(APP_CONFIG.sheets.encoder, true);
}

function APP_getRequestTempStoreSheet_() {
  return APP_getSheet_(APP_CONFIG.sheets.requestTempStore, true);
}

function APP_getDocumentProperties_() {
  try {
    return PropertiesService.getDocumentProperties() || PropertiesService.getScriptProperties();
  } catch (error) {
    return PropertiesService.getScriptProperties();
  }
}

function APP_getHeaderValues_(sheet, rowNumber, width) {
  return sheet.getRange(rowNumber, 1, 1, width).getValues()[0];
}

function APP_headersLookValid_(actualHeaders, expectedHeaders) {
  for (let i = 0; i < expectedHeaders.length; i += 1) {
    if (APP_normalizeText_(actualHeaders[i]) !== APP_normalizeText_(expectedHeaders[i])) {
      return false;
    }
  }
  return true;
}

function APP_buildRequestContextText_(context) {
  if (!context) {
    return '';
  }
  return APP_CONFIG.labels.requestedBy + ': ' + APP_safeString_(context.requestedBy).trim();
}

function APP_safeToast_(message, title) {
  try {
    APP_getSpreadsheet_().toast(message, title || APP_CONFIG.appName);
  } catch (error) {
    console.log((title || APP_CONFIG.appName) + ': ' + message);
  }
}

function APP_hasSheetProtection_(sheet) {
  if (!sheet) {
    return false;
  }
  return sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET).some(function (protection) {
    return !protection.isWarningOnly();
  });
}

function APP_deleteRowsByNumbers_(sheet, rowNumbers) {
  if (!sheet || !rowNumbers || rowNumbers.length === 0) {
    return 0;
  }
  const lastRow = sheet.getLastRow();
  const sortedRows = rowNumbers.filter(function (rowNumber, index, values) {
    return rowNumber > 1 && rowNumber <= lastRow && values.indexOf(rowNumber) === index;
  }).sort(function (a, b) {
    return a - b;
  });
  if (sortedRows.length === 0) {
    return 0;
  }

  const groups = [];
  let groupStart = sortedRows[0];
  let groupCount = 1;
  for (let i = 1; i < sortedRows.length; i += 1) {
    if (sortedRows[i] === sortedRows[i - 1] + 1) {
      groupCount += 1;
      continue;
    }
    groups.push({ startRow: groupStart, count: groupCount });
    groupStart = sortedRows[i];
    groupCount = 1;
  }
  groups.push({ startRow: groupStart, count: groupCount });
  groups.reverse().forEach(function (group) {
    sheet.deleteRows(group.startRow, group.count);
  });
  return sortedRows.length;
}

function APP_writeRowsByNumber_(sheet, startColumn, rowNumbers, rowValues) {
  if (!sheet || !rowNumbers || !rowValues || rowNumbers.length === 0 || rowValues.length === 0) {
    return 0;
  }

  const entries = rowNumbers.map(function (rowNumber, index) {
    return {
      rowNumber: rowNumber,
      values: rowValues[index]
    };
  }).filter(function (entry) {
    return entry.rowNumber > 0 && !!entry.values;
  }).sort(function (a, b) {
    return a.rowNumber - b.rowNumber;
  });

  if (entries.length === 0) {
    return 0;
  }

  const width = entries[0].values.length;
  let groupStart = entries[0].rowNumber;
  let groupValues = [entries[0].values];
  let previousRow = entries[0].rowNumber;
  let writeCount = 0;

  function flushGroup() {
    if (groupValues.length === 0) {
      return;
    }
    sheet.getRange(groupStart, startColumn, groupValues.length, width).setValues(groupValues);
    writeCount += groupValues.length;
  }

  for (let i = 1; i < entries.length; i += 1) {
    if (entries[i].rowNumber === previousRow + 1) {
      groupValues.push(entries[i].values);
      previousRow = entries[i].rowNumber;
      continue;
    }
    flushGroup();
    groupStart = entries[i].rowNumber;
    groupValues = [entries[i].values];
    previousRow = entries[i].rowNumber;
  }

  flushGroup();
  return writeCount;
}

function APP_getInventoryVersion_() {
  return APP_getDocumentProperties_().getProperty(APP_CONFIG.properties.inventoryVersion) || '0';
}

function APP_invalidateInventorySearchCaches_() {
  APP_getDocumentProperties_().setProperty(APP_CONFIG.properties.inventoryVersion, String(Date.now()));
  APP_clearExecutionCache_('inventorySearch');
}

function APP_getInventoryMetadataCell_(sheet, rowNumber) {
  return sheet.getRange(rowNumber, APP_CONFIG.inventory.cols.NAME, 1, 1);
}

function APP_getInventoryItemKeyNote_(cell) {
  const note = APP_safeString_(cell.getNote()).trim();
  const prefix = APP_CONFIG.metadataKeys.inventoryItemKey + ':';
  return note.indexOf(prefix) === 0 ? note.slice(prefix.length).trim() : '';
}

function APP_getInventoryItemKeyFromNote_(note) {
  const raw = APP_safeString_(note).trim();
  const prefix = APP_CONFIG.metadataKeys.inventoryItemKey + ':';
  return raw.indexOf(prefix) === 0 ? raw.slice(prefix.length).trim() : '';
}

function APP_setInventoryItemKeyNote_(cell, itemKey) {
  cell.setNote(APP_CONFIG.metadataKeys.inventoryItemKey + ':' + itemKey);
}

function APP_generateInventoryItemKey_() {
  return 'ITEM::' + Utilities.getUuid();
}

function APP_buildInventoryRecordFromRowValues_(rowNumber, rowValues, itemKey) {
  const values = rowValues || [];
  const name = values[APP_CONFIG.inventory.cols.NAME - 1];
  const code = values[APP_CONFIG.inventory.cols.CODE - 1];
  if (APP_isBlank_(name) && APP_isBlank_(code)) {
    return null;
  }
  return {
    rowNumber: rowNumber,
    key: itemKey || '',
    name: name || '',
    code: code || '',
    qty: APP_toNumber_(values[APP_CONFIG.inventory.cols.QTY - 1], 0),
    price: values[APP_CONFIG.inventory.cols.PRICE - 1] || 0,
    dateInv: values[APP_CONFIG.inventory.cols.DATE_INV - 1] || '',
    remarks: values[APP_CONFIG.inventory.cols.REMARKS - 1] || '',
    rowValues: values
  };
}

function APP_isDefaultNumberFormat_(numberFormat) {
  const rawFormat = APP_safeString_(numberFormat).trim();
  const normalized = APP_normalizeText_(rawFormat);
  return rawFormat === '' || normalized === 'GENERAL' || rawFormat === '@';
}

function APP_ensureInventoryDateTimeFormat_(sheet, rowNumber) {
  const targetSheet = sheet || APP_getInventorySheet_();
  const dateCell = targetSheet.getRange(rowNumber, APP_CONFIG.inventory.cols.DATE_INV);
  if (APP_isDefaultNumberFormat_(dateCell.getNumberFormat())) {
    dateCell.setNumberFormat(APP_CONFIG.inventory.dateTimeNumberFormat);
  }
}

function APP_writeInventoryStockRows_(sheet, updates) {
  const targetSheet = sheet || APP_getInventorySheet_();
  const safeUpdates = (updates || []).filter(function (update) {
    return update && APP_toNumber_(update.rowNumber, 0) >= APP_CONFIG.inventory.startRow;
  }).map(function (update) {
    return {
      rowNumber: APP_toNumber_(update.rowNumber, 0),
      qty: APP_toNumber_(update.qty, 0),
      price: update && update.price !== undefined ? update.price : '',
      dateInv: update && update.dateInv !== undefined ? update.dateInv : ''
    };
  });

  if (safeUpdates.length === 0) {
    return 0;
  }

  const rowNumbers = safeUpdates.map(function (update) {
    return update.rowNumber;
  });
  const rowValues = safeUpdates.map(function (update) {
    return [update.qty, update.price, update.dateInv];
  });

  const writeCount = APP_writeRowsByNumber_(targetSheet, APP_CONFIG.inventory.cols.QTY, rowNumbers, rowValues);
  safeUpdates.forEach(function (update) {
    APP_ensureInventoryDateTimeFormat_(targetSheet, update.rowNumber);
  });
  return writeCount;
}

function APP_getInventoryRecordByRowNumber_(sheet, rowNumber, ensureKey) {
  const targetSheet = sheet || APP_getInventorySheet_();
  const safeRow = APP_toNumber_(rowNumber, 0);
  if (safeRow < APP_CONFIG.inventory.startRow || safeRow > targetSheet.getLastRow()) {
    return null;
  }
  const rowValues = targetSheet.getRange(safeRow, 1, 1, APP_CONFIG.inventory.lastDataColumn).getValues()[0];
  const nameCell = APP_getInventoryMetadataCell_(targetSheet, safeRow);
  let itemKey = APP_getInventoryItemKeyNote_(nameCell);
  if (ensureKey !== false && APP_isBlank_(itemKey)) {
    itemKey = APP_generateInventoryItemKey_();
    APP_setInventoryItemKeyNote_(nameCell, itemKey);
  }
  return APP_buildInventoryRecordFromRowValues_(safeRow, rowValues, itemKey);
}

function APP_getInventorySnapshot_(options) {
  const settings = options || {};
  const cache = APP_getExecutionCache_('inventorySearch');
  const version = APP_getInventoryVersion_();
  if (!settings.forceRefresh && cache.snapshot && cache.version === version) {
    return cache.snapshot;
  }

  const sheet = APP_getInventorySheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < APP_CONFIG.inventory.startRow) {
    const emptySnapshot = { version: version, records: [] };
    cache.version = version;
    cache.snapshot = emptySnapshot;
    cache.indexes = null;
    return emptySnapshot;
  }

  const rowCount = lastRow - APP_CONFIG.inventory.startRow + 1;
  const values = sheet.getRange(APP_CONFIG.inventory.startRow, 1, rowCount, APP_CONFIG.inventory.lastDataColumn).getValues();
  const noteRange = sheet.getRange(APP_CONFIG.inventory.startRow, APP_CONFIG.inventory.cols.NAME, rowCount, 1);
  const notes = noteRange.getNotes();
  let notesChanged = false;
  const records = [];

  for (let offset = 0; offset < values.length; offset += 1) {
    const rowNumber = APP_CONFIG.inventory.startRow + offset;
    let itemKey = APP_getInventoryItemKeyFromNote_(notes[offset][0]);
    const record = APP_buildInventoryRecordFromRowValues_(rowNumber, values[offset], itemKey);
    if (!record) {
      continue;
    }
    if (APP_isBlank_(itemKey)) {
      itemKey = APP_generateInventoryItemKey_();
      notes[offset][0] = APP_CONFIG.metadataKeys.inventoryItemKey + ':' + itemKey;
      notesChanged = true;
      record.key = itemKey;
    }
    records.push(record);
  }

  if (notesChanged) {
    noteRange.setNotes(notes);
  }

  const snapshot = {
    version: version,
    records: records
  };
  cache.version = version;
  cache.snapshot = snapshot;
  cache.indexes = null;
  return snapshot;
}

function APP_getInventoryIndexes_() {
  const cache = APP_getExecutionCache_('inventorySearch');
  const version = APP_getInventoryVersion_();
  if (cache.indexes && cache.version === version) {
    return cache.indexes;
  }

  const records = APP_getInventorySnapshot_().records;
  const indexes = {
    version: version,
    records: records,
    byKey: {},
    byRow: {},
    byCode: {}
  };

  records.forEach(function (record) {
    indexes.byKey[record.key] = record;
    indexes.byRow[record.rowNumber] = record;
    APP_splitCodes_(record.code).forEach(function (token) {
      if (!indexes.byCode[token]) {
        indexes.byCode[token] = [];
      }
      indexes.byCode[token].push(record);
    });
    const exactCode = APP_normalizeText_(record.code);
    if (exactCode !== '' && !APP_isNoCode(exactCode) && APP_splitCodes_(record.code).length === 0) {
      if (!indexes.byCode[exactCode]) {
        indexes.byCode[exactCode] = [];
      }
      indexes.byCode[exactCode].push(record);
    }
  });

  cache.version = version;
  cache.indexes = indexes;
  return indexes;
}

function APP_getInventoryRecords_(options) {
  const settings = options || {};
  if (settings.ensureKey === false) {
    return APP_getInventorySnapshot_(settings).records.map(function (record) {
      const clone = Object.assign({}, record);
      clone.key = '';
      return clone;
    });
  }
  return APP_getInventorySnapshot_(settings).records.slice();
}

function APP_findInventoryRecordByKey_(sheet, itemKey) {
  if (APP_isBlank_(itemKey)) {
    return null;
  }
  const indexes = APP_getInventoryIndexes_();
  return indexes.byKey[itemKey] || null;
}

function APP_findInventoryByRow_(rowNumber) {
  const safeRow = APP_toNumber_(rowNumber, 0);
  if (safeRow < APP_CONFIG.inventory.startRow) {
    return null;
  }
  const indexes = APP_getInventoryIndexes_();
  return indexes.byRow[safeRow] || null;
}

function APP_findInventoryByCode_(searchCode) {
  const normalizedSearch = APP_normalizeText_(searchCode);
  if (normalizedSearch === '' || APP_isNoCode(normalizedSearch)) {
    return [];
  }

  const indexes = APP_getInventoryIndexes_();
  return indexes.byCode[normalizedSearch] ? indexes.byCode[normalizedSearch].slice() : [];
}

function APP_findInventoryRecordsByCode_(sheet, searchCode) {
  return APP_findInventoryByCode_(searchCode);
}

function APP_findInventoryRecordByCode_(sheet, searchCode) {
  const matches = APP_findInventoryByCode_(searchCode);
  return matches.length > 0 ? matches[0] : null;
}

function APP_searchInventoryByName_(searchTerm, limit) {
  const normalizedTerm = APP_safeString_(searchTerm).trim().toLowerCase();
  if (normalizedTerm === '') {
    return [];
  }
  const maxResults = limit || APP_CONFIG.search.maxResults;

  return APP_getInventoryRecords_().filter(function (record) {
    return APP_safeString_(record.name).toLowerCase().indexOf(normalizedTerm) !== -1;
  }).slice(0, maxResults);
}

function APP_resolveInventoryRecord_(sheet, itemKey, options) {
  const settings = options || {};
  if (!APP_isBlank_(itemKey)) {
    const byKey = APP_findInventoryRecordByKey_(sheet || APP_getInventorySheet_(), itemKey);
    if (byKey) {
      return byKey;
    }
  }
  if (APP_toNumber_(settings.sourceRow, 0) >= APP_CONFIG.inventory.startRow) {
    const byRow = APP_findInventoryByRow_(settings.sourceRow);
    if (byRow) {
      return byRow;
    }
  }
  if (!APP_isBlank_(settings.code)) {
    const byCode = APP_findInventoryByCode_(settings.code);
    if (byCode.length === 1) {
      return byCode[0];
    }
  }
  return null;
}

function APP_validateAction_(action) {
  return action === APP_CONFIG.actions.ADD || action === APP_CONFIG.actions.SUBTRACT;
}

function APP_getActionDelta_(action, quantity) {
  const safeQuantity = APP_toNumber_(quantity, 0);
  if (!APP_validateAction_(action) || safeQuantity <= 0) {
    return 0;
  }
  return action === APP_CONFIG.actions.ADD ? safeQuantity : (-1 * safeQuantity);
}

function APP_showActionDialog_(config) {
  const safeConfig = config || {};
  const record = safeConfig.record || {};
  APP_reconcilePendingActionDialog_({ force: true });
  APP_registerPendingActionDialog_(safeConfig.cancelPayload || {});
  const template = HtmlService.createTemplateFromFile('ActionDialog');
  template.dialogData = {
    title: safeConfig.title || 'Update Item',
    requestedBy: safeConfig.requestedBy || '',
    itemName: record.name || '',
    code: APP_isNoCode(record.code) ? 'NO CODE' : APP_safeString_(record.code),
    availableQty: APP_toNumber_(record.qty, 0),
    pricePerPiece: record.price === undefined || record.price === null || record.price === '' ? '' : record.price,
    remarks: APP_safeString_(record.remarks),
    quantityLabel: safeConfig.quantityLabel || 'Quantity',
    heartbeatMs: APP_CONFIG.system.dialogs.heartbeatMs,
    submitPayload: safeConfig.submitPayload || {},
    cancelPayload: safeConfig.cancelPayload || {}
  };
  const html = template.evaluate()
    .setWidth(APP_CONFIG.system.dialogs.width)
    .setHeight(APP_CONFIG.system.dialogs.height)
    .setTitle(safeConfig.title || 'Update Item');
  SpreadsheetApp.getUi().showModalDialog(html, safeConfig.title || 'Update Item');
  return APP_result_(true, 'Action dialog opened.', { deferred: true });
}

function APP_getDialogStatePropertyKey_() {
  return APP_CONFIG.system.dialogs.statePropertyPrefix + APP_getActor_().key;
}

function APP_registerPendingActionDialog_(cancelPayload) {
  const safePayload = cancelPayload || {};
  const now = Date.now();
  APP_getDocumentProperties_().setProperty(APP_getDialogStatePropertyKey_(), JSON.stringify({
    payload: safePayload,
    openedAtMs: now,
    lastSeenAtMs: now
  }));
}

function APP_getPendingActionDialog_() {
  const raw = APP_getDocumentProperties_().getProperty(APP_getDialogStatePropertyKey_());
  if (APP_isBlank_(raw)) {
    return null;
  }
  try {
    return JSON.parse(raw);
  } catch (error) {
    APP_getDocumentProperties_().deleteProperty(APP_getDialogStatePropertyKey_());
    return null;
  }
}

function APP_clearPendingActionDialog_() {
  APP_getDocumentProperties_().deleteProperty(APP_getDialogStatePropertyKey_());
}

function APP_touchPendingActionDialog_() {
  const state = APP_getPendingActionDialog_();
  if (!state) {
    return APP_result_(true, 'No pending dialog state to touch.', { active: false });
  }
  state.lastSeenAtMs = Date.now();
  APP_getDocumentProperties_().setProperty(APP_getDialogStatePropertyKey_(), JSON.stringify(state));
  return APP_result_(true, 'Pending dialog state touched.', { active: true });
}

function APP_touchPendingActionDialog() {
  return APP_touchPendingActionDialog_();
}

function APP_applyActionDialogCleanup_(payload) {
  const safePayload = payload || {};
  if (safePayload.target === 'REQUEST_SELECTION') {
    APP_clearRequestSelectionMarker_(safePayload.markerRow);
  }
  if (safePayload.target === 'ENCODER' && safePayload.clearOnCancel) {
    ENC_clearInputFields_(APP_getEncoderSheet_());
  }
}

function APP_reconcilePendingActionDialog_(options) {
  const settings = options || {};
  const state = APP_getPendingActionDialog_();
  if (!state) {
    return APP_result_(true, 'No pending dialog state.');
  }
  const lastSeenAtMs = APP_toNumber_(state.lastSeenAtMs || state.openedAtMs, 0);
  const isStale = (Date.now() - lastSeenAtMs) > APP_CONFIG.system.dialogs.staleMs;
  if (!settings.force && !isStale) {
    return APP_result_(true, 'Pending dialog still active.', { active: true });
  }
  APP_applyActionDialogCleanup_(state.payload || {});
  APP_clearPendingActionDialog_();
  return APP_result_(true, 'Pending dialog state cleaned up.', { cleaned: true, stale: isStale });
}

function APP_clearRequestSelectionMarker_(markerRow) {
  const safeRow = APP_toNumber_(markerRow, 0);
  if (safeRow < APP_CONFIG.request.searchTable.startRow) {
    return;
  }
  APP_getRequestSheet_().getRange(safeRow, APP_CONFIG.request.searchTable.cols.SELECT).setValue('');
}

function APP_cancelActionDialog_(payload) {
  APP_applyActionDialogCleanup_(payload || {});
  APP_clearPendingActionDialog_();
  return APP_result_(true, 'Action dialog cancelled.');
}

function APP_cancelActionDialog(payload) {
  return APP_cancelActionDialog_(payload);
}

function APP_submitActionDialog_(payload) {
  const safePayload = payload || {};
  const quantity = Number(safePayload.quantity);
  const action = APP_normalizeText_(safePayload.action);

  if (!APP_validateAction_(action)) {
    return APP_fail_('INVALID_ACTION', 'Action must be ADD or SUBTRACT.');
  }
  if (!isFinite(quantity) || quantity <= 0) {
    return APP_fail_('INVALID_QUANTITY', 'Quantity must be greater than zero.');
  }

  const record = APP_resolveInventoryRecord_(APP_getInventorySheet_(), safePayload.itemKey, {
    sourceRow: safePayload.sourceRow,
    code: safePayload.code || ''
  });
  if (!record) {
    return APP_fail_('ITEM_NOT_FOUND', 'The selected inventory item could not be resolved.');
  }

  try {
    if (safePayload.target === 'ENCODER') {
      const result = APP_applyInventoryAction_(record.key, action, quantity, {
        allowNegative: false,
        requestedBy: safePayload.requestedBy || APP_getActor_().display,
        source: 'ENCODER',
        secondaryLogSheet: LOG_CONFIG.SHEET_NAMES.ENCODER_LOGS
      });
      if (result.ok) {
        APP_clearPendingActionDialog_();
        ENC_clearEncoderSheet(APP_getEncoderSheet_());
        APP_safeToast_(result.message, 'Encoder');
      }
      return result;
    }

    if (safePayload.target === 'REQUEST_IMMEDIATE' || safePayload.target === 'REQUEST_SELECTION') {
      const stageResult = REQ_stageInventoryRecord_(record, action, quantity, {
        reqSheet: APP_getRequestSheet_(),
        clearEntryFields: safePayload.clearEntryFields !== false
      });
      if (stageResult.ok) {
        if (safePayload.target === 'REQUEST_SELECTION') {
          APP_clearRequestSelectionMarker_(safePayload.markerRow);
        }
        APP_clearPendingActionDialog_();
      }
      APP_safeToast_(stageResult.message, stageResult.ok ? 'Request' : 'Request Error');
      return APP_result_(stageResult.ok, stageResult.message, {
        sessionId: stageResult.data && stageResult.data.sessionId ? stageResult.data.sessionId : '',
        itemName: record.name,
        code: record.code,
        quantity: quantity,
        action: action
      });
    }

    return APP_fail_('UNSUPPORTED_TARGET', 'Unsupported action target.');
  } catch (error) {
    return APP_fail_(error.code || 'ACTION_DIALOG_FAILED', error.message, error.details);
  }
}

function APP_submitActionDialog(payload) {
  return APP_submitActionDialog_(payload);
}

function APP_applyInventoryAction_(itemKey, action, quantity, context) {
  const settings = context || {};
  if (!APP_validateAction_(action)) {
    throw APP_createError_('INVALID_ACTION', 'Action must be ADD or SUBTRACT.');
  }
  const safeQuantity = APP_toNumber_(quantity, 0);
  if (safeQuantity <= 0) {
    throw APP_createError_('INVALID_QUANTITY', 'Quantity must be greater than zero.');
  }

  const delta = APP_getActionDelta_(action, safeQuantity);
  const sheet = APP_getInventorySheet_();
  const result = APP_withScriptLock_(30000, function () {
    const record = APP_findInventoryRecordByKey_(sheet, itemKey);
    if (!record) {
      throw APP_createError_('ITEM_NOT_FOUND', 'Inventory item could not be resolved.');
    }

    const beforeQty = APP_toNumber_(record.qty, 0);
    const beforeDate = record.dateInv;
    const afterQty = beforeQty + delta;
    if (action === APP_CONFIG.actions.SUBTRACT && !settings.allowNegative && afterQty < 0) {
      throw APP_createError_('INSUFFICIENT_STOCK', 'Subtracting this quantity would make stock negative.');
    }

    const now = APP_now_();
    const writeback = {
      rowNumber: record.rowNumber,
      qty: afterQty,
      price: record.price,
      dateInv: now
    };
    const rollback = {
      rowNumber: record.rowNumber,
      qty: beforeQty,
      price: record.price,
      dateInv: beforeDate
    };
    APP_writeInventoryStockRows_(sheet, [writeback]);

    const logEntry = {
      requestedBy: settings.requestedBy || APP_getActor_().display,
      action: action,
      itemName: record.name,
      code: record.code,
      quantity: safeQuantity,
      remarks: record.remarks,
      source: settings.source || 'INVENTORY',
      timestamp: now
    };
    const changeLogRow = LOG_appendChangeLog(logEntry);
    if (changeLogRow <= 0) {
      APP_writeInventoryStockRows_(sheet, [rollback]);
      throw APP_createError_('LOG_WRITE_FAILED', 'Inventory update log could not be written.');
    }

    if (!APP_isBlank_(settings.secondaryLogSheet)) {
      const secondaryRow = LOG_appendEntry(settings.secondaryLogSheet, logEntry);
      if (secondaryRow <= 0) {
        LOG_deleteRows_(LOG_CONFIG.SHEET_NAMES.CHANGE_LOGS, [changeLogRow]);
        APP_writeInventoryStockRows_(sheet, [rollback]);
        throw APP_createError_('LOG_WRITE_FAILED', 'Secondary action log could not be written.');
      }
    }

    APP_invalidateInventorySearchCaches_();

    return APP_result_(true, 'Inventory updated successfully.', {
      rowNumber: record.rowNumber,
      itemKey: record.key,
      itemName: record.name,
      code: record.code,
      quantity: safeQuantity,
      action: action,
      beforeQty: beforeQty,
      afterQty: afterQty,
      remarks: record.remarks
    });
  });

  try {
    if (result.ok && typeof syncRowToTarget === 'function') {
      syncRowToTarget(result.data.rowNumber);
    }
  } catch (syncError) {
    console.log('View-only sync warning: ' + syncError.message);
  }

  return result;
}

function APP_withScriptLock_(timeoutMs, callback) {
  const lock = LockService.getScriptLock();
  const timeout = timeoutMs || 30000;
  if (!lock.tryLock(timeout)) {
    throw APP_createError_('LOCK_TIMEOUT', 'System is busy. Please retry in a moment.');
  }
  try {
    return callback();
  } finally {
    lock.releaseLock();
  }
}

function APP_ensureTrigger_(handlerName, eventType) {
  const spreadsheet = APP_getSpreadsheet_();
  const spreadsheetId = spreadsheet.getId();
  const normalizedEventType = eventType === 'onEdit' ? 'ON_EDIT' : (eventType === 'onChange' ? 'ON_CHANGE' : '');
  if (normalizedEventType === '') {
    throw APP_createError_('INVALID_TRIGGER_TYPE', 'Unsupported trigger type: ' + eventType);
  }

  ScriptApp.getProjectTriggers().forEach(function (trigger) {
    const triggerSourceId = trigger.getTriggerSourceId ? trigger.getTriggerSourceId() : '';
    if (
      trigger.getHandlerFunction() === handlerName &&
      String(trigger.getEventType()) === normalizedEventType &&
      (!triggerSourceId || triggerSourceId === spreadsheetId)
    ) {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  let builder = ScriptApp.newTrigger(handlerName).forSpreadsheet(spreadsheet);
  builder = eventType === 'onEdit' ? builder.onEdit() : builder.onChange();
  builder.create();
}

function APP_getTriggerHandlers_() {
  return ScriptApp.getProjectTriggers().map(function (trigger) {
    return trigger.getHandlerFunction();
  });
}

function APP_verifyHeaderCheck_(report, label, sheetName, rowNumber, expectedHeaders, startColumn) {
  const sheet = APP_getSpreadsheet_().getSheetByName(sheetName);
  if (!sheet) {
    report.checks.push({ name: label, ok: false, details: 'Sheet not found.' });
    report.ok = false;
    return;
  }
  const width = expectedHeaders.length;
  const values = sheet.getRange(rowNumber, startColumn || 1, 1, width).getValues()[0];
  const ok = APP_headersLookValid_(values, expectedHeaders);
  report.checks.push({
    name: label,
    ok: ok,
    details: ok ? 'Headers match expected layout.' : 'Expected: ' + expectedHeaders.join(', ')
  });
  if (!ok) {
    report.ok = false;
  }
}

function APP_verifyRequiredConfig_(report) {
  const uniqueSheetNames = Object.keys(APP_CONFIG.sheets).map(function (key) {
    return APP_CONFIG.sheets[key];
  });
  const duplicateSheetNames = uniqueSheetNames.filter(function (name, index, values) {
    return values.indexOf(name) !== index;
  });

  const checks = [
    {
      name: 'Application name',
      ok: !APP_isBlank_(APP_CONFIG.appName),
      details: APP_CONFIG.appName
    },
    {
      name: 'Unique sheet names',
      ok: duplicateSheetNames.length === 0,
      details: duplicateSheetNames.length === 0 ? 'All configured sheet names are unique.' : duplicateSheetNames.join(', ')
    },
    {
      name: 'Inventory structure rows',
      ok: APP_CONFIG.inventory.startRow > APP_CONFIG.inventory.headerRow,
      details: 'Header row ' + APP_CONFIG.inventory.headerRow + ', data row ' + APP_CONFIG.inventory.startRow
    },
    {
      name: 'Encoder structure rows',
      ok: APP_CONFIG.encoder.resultStartRow > APP_CONFIG.encoder.resultHeaderRow,
      details: 'Header row ' + APP_CONFIG.encoder.resultHeaderRow + ', data row ' + APP_CONFIG.encoder.resultStartRow
    },
    {
      name: 'Request structure rows',
      ok: REQ_CONFIG.SEARCH_RESULTS_TABLE.startRow > REQ_CONFIG.SEARCH_RESULTS_TABLE.headerRow && REQ_CONFIG.ITEMS_TABLE.startRow > REQ_CONFIG.ITEMS_TABLE.headerRow,
      details: 'Search header row ' + REQ_CONFIG.SEARCH_RESULTS_TABLE.headerRow + ', staged header row ' + REQ_CONFIG.ITEMS_TABLE.headerRow
    },
    {
      name: 'Log structure rows',
      ok: APP_CONFIG.logs.layout.dataStartRow > APP_CONFIG.logs.layout.headerRow,
      details: 'Header row ' + APP_CONFIG.logs.layout.headerRow + ', data row ' + APP_CONFIG.logs.layout.dataStartRow
    }
  ];

  checks.forEach(function (check) {
    report.checks.push(check);
    if (!check.ok) {
      report.ok = false;
    }
  });
}

function APP_verifyCallableFunctions_(report) {
  const functionChecks = [
    { name: 'onOpen', fn: onOpen },
    { name: 'APP_initializeSystem', fn: APP_initializeSystem },
    { name: 'handleEncoderEdit', fn: handleEncoderEdit },
    { name: 'REQ_handleEdit', fn: REQ_handleEdit },
    { name: 'REQ_finalizeRequest', fn: REQ_finalizeRequest },
    { name: 'REQ_storeTemporarily', fn: REQ_storeTemporarily },
    { name: 'REQ_pullStoredRequest', fn: REQ_pullStoredRequest },
    { name: 'LOG_handleInventoryEdit', fn: LOG_handleInventoryEdit },
    { name: 'APP_handleSpreadsheetChange', fn: APP_handleSpreadsheetChange }
  ];

  functionChecks.forEach(function (check) {
    const ok = typeof check.fn === 'function';
    report.functionChecks.push({ name: check.name, ok: ok });
    if (!ok) {
      report.ok = false;
    }
  });
}

function APP_verifyMenus_(report) {
  try {
    onOpen();
    report.menuCheck = {
      ok: true,
      menuName: APP_CONFIG.system.menu.root,
      details: 'Custom menu rebuilt successfully.'
    };
  } catch (error) {
    report.menuCheck = {
      ok: false,
      menuName: APP_CONFIG.system.menu.root,
      details: error.message
    };
    report.ok = false;
  }
}

function runHealthCheck() {
  const requiredSheets = [
    APP_CONFIG.sheets.inventory,
    APP_CONFIG.sheets.requestForm,
    APP_CONFIG.sheets.encoder,
    APP_CONFIG.sheets.changeLogs,
    APP_CONFIG.sheets.encoderLogs,
    APP_CONFIG.sheets.requestLogs,
    APP_CONFIG.sheets.finalizedRequestLogs,
    APP_CONFIG.sheets.portfolioNotes,
    APP_CONFIG.sheets.requestTempStore
  ];
  const missingSheets = requiredSheets.filter(function (sheetName) {
    return !APP_getSpreadsheet_().getSheetByName(sheetName);
  });
  const handlers = APP_getTriggerHandlers_();
  const missingTriggers = Object.keys(APP_CONFIG.triggerHandlers).filter(function (key) {
    return handlers.indexOf(APP_CONFIG.triggerHandlers[key]) === -1;
  }).map(function (key) {
    return APP_CONFIG.triggerHandlers[key];
  });

  const report = {
    ok: missingSheets.length === 0 && missingTriggers.length === 0,
    appName: APP_CONFIG.appName,
    generatedAt: APP_getTimestampParts_(APP_now_()).iso,
    missingSheets: missingSheets,
    missingTriggers: missingTriggers,
    checks: [],
    functionChecks: [],
    menuCheck: null
  };

  APP_verifyRequiredConfig_(report);
  APP_verifyHeaderCheck_(report, 'Inventory headers', APP_CONFIG.sheets.inventory, APP_CONFIG.inventory.headerRow, APP_CONFIG.inventory.visibleHeaders, 1);
  APP_verifyHeaderCheck_(report, 'Encoder headers', APP_CONFIG.sheets.encoder, APP_CONFIG.encoder.resultHeaderRow, APP_CONFIG.encoder.visibleHeaders, 1);
  APP_verifyHeaderCheck_(report, 'Request search headers', APP_CONFIG.sheets.requestForm, REQ_CONFIG.SEARCH_RESULTS_TABLE.headerRow, REQ_CONFIG.SEARCH_RESULTS_TABLE.visibleHeaders, REQ_CONFIG.SEARCH_RESULTS_TABLE.cols.NAME);
  APP_verifyHeaderCheck_(report, 'Request staged headers', APP_CONFIG.sheets.requestForm, REQ_CONFIG.ITEMS_TABLE.headerRow, REQ_CONFIG.ITEMS_TABLE.visibleHeaders, REQ_CONFIG.ITEMS_TABLE.cols.ACTION);
  APP_verifyHeaderCheck_(report, 'Change log headers', APP_CONFIG.sheets.changeLogs, APP_CONFIG.logs.layout.headerRow, LOG_getHeadersForSheet_(APP_CONFIG.sheets.changeLogs), 1);
  APP_verifyHeaderCheck_(report, 'Encoder log headers', APP_CONFIG.sheets.encoderLogs, APP_CONFIG.logs.layout.headerRow, LOG_getHeadersForSheet_(APP_CONFIG.sheets.encoderLogs), 1);
  APP_verifyHeaderCheck_(report, 'Request log headers', APP_CONFIG.sheets.requestLogs, APP_CONFIG.logs.layout.headerRow, LOG_getHeadersForSheet_(APP_CONFIG.sheets.requestLogs), 1);
  APP_verifyHeaderCheck_(report, 'Finalized request log headers', APP_CONFIG.sheets.finalizedRequestLogs, APP_CONFIG.logs.layout.headerRow, LOG_getHeadersForSheet_(APP_CONFIG.sheets.finalizedRequestLogs), 1);
  APP_verifyHeaderCheck_(report, 'Temp store headers', APP_CONFIG.sheets.requestTempStore, REQ_CONFIG.TEMP_STORE.headerRow, REQ_CONFIG.TEMP_STORE.headers, 1);
  APP_verifyCallableFunctions_(report);
  APP_verifyMenus_(report);

  return report;
}
