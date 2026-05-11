const REQ_CONFIG = {
  get SHEET_NAME() {
    return APP_CONFIG.sheets.requestForm;
  },
  get REQUEST_FORM_CELLS() {
    return APP_CONFIG.request.cells;
  },
  get SEARCH_RESULTS_TABLE() {
    return APP_CONFIG.request.searchTable;
  },
  get ITEMS_TABLE() {
    return APP_CONFIG.request.itemsTable;
  },
  get TEMP_STORE() {
    return APP_CONFIG.request.tempStore;
  },
  get FLAGS() {
    return APP_CONFIG.request.flags;
  },
  get STORAGE() {
    return APP_CONFIG.request.sessionStorage;
  }
};

function REQ_getRequestContext_(reqSheet) {
  const sheet = reqSheet || APP_getRequestSheet_();
  const actor = APP_getActor_();
  const requestedBy = APP_safeString_(sheet.getRange(REQ_CONFIG.REQUEST_FORM_CELLS.REQUESTED_BY).getValue()).trim();
  return {
    actor: actor,
    requestedBy: requestedBy,
    requestedByKey: APP_normalizeText_(requestedBy),
    contextKey: actor.key + '|' + APP_normalizeText_(requestedBy)
  };
}

function REQ_validateContext_(context) {
  if (!context || APP_isBlank_(context.requestedBy)) {
    return APP_fail_('MISSING_REQUESTED_BY', 'Requested By is required.');
  }
  return APP_result_(true, 'Context valid.', context);
}

function REQ_generateSessionId_(context) {
  return 'REQ-' + Utilities.formatDate(APP_now_(), APP_getTimeZone_(), 'yyyyMMdd-HHmmss') + '-' + context.actor.key.replace(/[^A-Za-z0-9]/g, '').slice(0, 10) + '-' + Utilities.getUuid().slice(0, 6).toUpperCase();
}

function REQ_generateItemId_() {
  return 'ITEM-' + Utilities.getUuid().slice(0, 10).toUpperCase();
}

function REQ_generateTempRecordId_() {
  return 'TMP-' + Utilities.formatDate(APP_now_(), APP_getTimeZone_(), 'yyyyMMdd-HHmmss') + '-' + Utilities.getUuid().slice(0, 6).toUpperCase();
}

function REQ_getSessionPropertyKey_(sessionId) {
  return REQ_CONFIG.STORAGE.sessionPrefix + APP_safeString_(sessionId).trim();
}

function REQ_getContextPropertyKey_(contextKey) {
  return REQ_CONFIG.STORAGE.contextPrefix + APP_safeString_(contextKey).trim();
}

function REQ_normalizeSessionItem_(item) {
  const safeItem = item || {};
  return {
    itemId: safeItem.itemId || REQ_generateItemId_(),
    partKey: safeItem.partKey || '',
    name: safeItem.name || '',
    code: safeItem.code || '',
    quantity: APP_toNumber_(safeItem.quantity, 0),
    price: safeItem.price || 0,
    remarks: safeItem.remarks || '',
    sourceRow: APP_toNumber_(safeItem.sourceRow, 0),
    action: APP_validateAction_(safeItem.action) ? safeItem.action : APP_CONFIG.actions.SUBTRACT,
    createdAt: safeItem.createdAt || APP_now_(),
    updatedAt: safeItem.updatedAt || APP_now_()
  };
}

function REQ_normalizeSessionRecord_(record) {
  const safeRecord = record || {};
  return {
    sessionId: safeRecord.sessionId || '',
    contextKey: safeRecord.contextKey || '',
    requestedBy: safeRecord.requestedBy || '',
    sourceTempRecordId: safeRecord.sourceTempRecordId || '',
    actor: safeRecord.actor || APP_getActor_().display,
    createdAt: safeRecord.createdAt || APP_now_(),
    updatedAt: safeRecord.updatedAt || APP_now_(),
    items: (safeRecord.items || []).map(REQ_normalizeSessionItem_)
  };
}

function REQ_writeSessionRecord_(record) {
  const normalized = REQ_normalizeSessionRecord_(record);
  if (APP_isBlank_(normalized.sessionId)) {
    throw APP_createError_('INVALID_SESSION', 'Missing session ID.');
  }
  const properties = APP_getDocumentProperties_();
  properties.setProperty(REQ_getSessionPropertyKey_(normalized.sessionId), JSON.stringify(normalized));
  if (!APP_isBlank_(normalized.contextKey)) {
    properties.setProperty(REQ_getContextPropertyKey_(normalized.contextKey), normalized.sessionId);
  }
  return normalized;
}

function REQ_getSessionRecord_(sessionId) {
  if (APP_isBlank_(sessionId)) {
    return null;
  }
  const raw = APP_getDocumentProperties_().getProperty(REQ_getSessionPropertyKey_(sessionId));
  if (APP_isBlank_(raw)) {
    return null;
  }
  try {
    return REQ_normalizeSessionRecord_(JSON.parse(raw));
  } catch (error) {
    APP_getDocumentProperties_().deleteProperty(REQ_getSessionPropertyKey_(sessionId));
    return null;
  }
}

function REQ_removeSessionRecord_(sessionId, contextKey) {
  if (APP_isBlank_(sessionId)) {
    return;
  }
  const properties = APP_getDocumentProperties_();
  properties.deleteProperty(REQ_getSessionPropertyKey_(sessionId));
  if (!APP_isBlank_(contextKey)) {
    const existing = properties.getProperty(REQ_getContextPropertyKey_(contextKey));
    if (existing === sessionId) {
      properties.deleteProperty(REQ_getContextPropertyKey_(contextKey));
    }
  }
}

function REQ_findExistingSessionId_(contextKey) {
  if (APP_isBlank_(contextKey)) {
    return '';
  }
  const sessionId = APP_getDocumentProperties_().getProperty(REQ_getContextPropertyKey_(contextKey)) || '';
  return REQ_getSessionRecord_(sessionId) ? sessionId : '';
}

function REQ_getOrCreateSessionId_(context) {
  return REQ_findExistingSessionId_(context.contextKey) || REQ_generateSessionId_(context);
}

function REQ_getCurrentSessionId() {
  const context = REQ_getRequestContext_(APP_getRequestSheet_());
  const validation = REQ_validateContext_(context);
  if (!validation.ok) {
    throw APP_createError_(validation.code, validation.message);
  }
  const sessionId = REQ_findExistingSessionId_(context.contextKey);
  if (APP_isBlank_(sessionId)) {
    throw APP_createError_('NO_SESSION', 'No active request session found.');
  }
  return sessionId;
}

function REQ_getSessionItems(sessionId) {
  const record = REQ_getSessionRecord_(sessionId);
  return record ? record.items : [];
}

function REQ_clearSession(sessionId) {
  return APP_withScriptLock_(30000, function () {
    const record = REQ_getSessionRecord_(sessionId);
    if (!record) {
      return APP_result_(true, 'No active request session to clear.');
    }
    REQ_removeSessionRecord_(record.sessionId, record.contextKey);
    return APP_result_(true, 'Staged request cleared.', { sessionId: sessionId });
  });
}

function REQ_ensureTempStoreSheet() {
  const ss = APP_getSpreadsheet_();
  let sheet = ss.getSheetByName(APP_CONFIG.sheets.requestTempStore);
  if (!sheet) {
    sheet = ss.insertSheet(APP_CONFIG.sheets.requestTempStore);
  }
  const requiredColumns = REQ_CONFIG.TEMP_STORE.headers.length;
  if (sheet.getMaxColumns() < requiredColumns) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), requiredColumns - sheet.getMaxColumns());
  }
  sheet.getRange(REQ_CONFIG.TEMP_STORE.headerRow, 1, 1, requiredColumns).setValues([REQ_CONFIG.TEMP_STORE.headers]);
  sheet.getRange(REQ_CONFIG.TEMP_STORE.headerRow, 1, 1, requiredColumns)
    .setFontWeight('bold')
    .setBackground('#1f1f1f')
    .setFontColor('#ffffff')
    .setWrap(true)
    .setBorder(true, true, true, true, true, true, APP_CONFIG.system.colors.border, SpreadsheetApp.BorderStyle.SOLID);
  sheet.setFrozenRows(1);
  try {
    sheet.hideSheet();
  } catch (error) {
    console.log('Temp store hide skipped: ' + error.message);
  }
  return sheet;
}

function REQ_getTempStoreRows_(tempSheet) {
  const sheet = tempSheet || REQ_ensureTempStoreSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return [];
  }
  const data = sheet.getRange(2, 1, lastRow - 1, REQ_CONFIG.TEMP_STORE.headers.length).getValues();
  return data.map(function (row, index) {
    return {
      rowNumber: index + 2,
      recordId: row[REQ_CONFIG.TEMP_STORE.cols.RECORD_ID - 1],
      requestedBy: row[REQ_CONFIG.TEMP_STORE.cols.REQUESTED_BY - 1],
      requestedByKey: row[REQ_CONFIG.TEMP_STORE.cols.REQUESTED_BY_KEY - 1],
      storedAt: row[REQ_CONFIG.TEMP_STORE.cols.STORED_AT - 1],
      storedBy: row[REQ_CONFIG.TEMP_STORE.cols.STORED_BY - 1],
      itemId: row[REQ_CONFIG.TEMP_STORE.cols.ITEM_ID - 1],
      action: row[REQ_CONFIG.TEMP_STORE.cols.ACTION - 1],
      name: row[REQ_CONFIG.TEMP_STORE.cols.NAME - 1],
      code: row[REQ_CONFIG.TEMP_STORE.cols.CODE - 1],
      quantity: APP_toNumber_(row[REQ_CONFIG.TEMP_STORE.cols.QTY - 1], 0),
      price: row[REQ_CONFIG.TEMP_STORE.cols.PRICE - 1] || 0,
      remarks: row[REQ_CONFIG.TEMP_STORE.cols.REMARKS - 1] || '',
      sourceRow: APP_toNumber_(row[REQ_CONFIG.TEMP_STORE.cols.SOURCE_ROW - 1], 0),
      itemKey: row[REQ_CONFIG.TEMP_STORE.cols.ITEM_KEY - 1] || ''
    };
  }).filter(function (row) {
    return !APP_isBlank_(row.recordId);
  });
}

function REQ_appendTempStoreRows_(rows) {
  const safeRows = rows || [];
  if (safeRows.length === 0) {
    return { startRow: 0, rowCount: 0 };
  }
  const sheet = REQ_ensureTempStoreSheet();
  const startRow = sheet.getLastRow() + 1;
  const values = safeRows.map(function (row) {
    return [
      row.recordId,
      row.requestedBy,
      row.requestedByKey,
      row.storedAt,
      row.storedBy,
      row.itemId,
      row.action,
      row.name,
      row.code,
      row.quantity,
      row.price,
      row.remarks,
      row.sourceRow,
      row.itemKey
    ];
  });
  sheet.getRange(startRow, 1, values.length, REQ_CONFIG.TEMP_STORE.headers.length).setValues(values);
  return { startRow: startRow, rowCount: values.length };
}

function REQ_deleteTempStoreRecord_(recordId) {
  const rows = REQ_getTempStoreRows_().filter(function (row) {
    return row.recordId === recordId;
  }).map(function (row) {
    return row.rowNumber;
  });
  return APP_deleteRowsByNumbers_(REQ_ensureTempStoreSheet(), rows);
}

function REQ_groupTempStoreRows_(rows) {
  const groups = {};
  (rows || []).forEach(function (row) {
    if (!groups[row.recordId]) {
      groups[row.recordId] = {
        recordId: row.recordId,
        requestedBy: row.requestedBy,
        requestedByKey: row.requestedByKey,
        storedAt: row.storedAt,
        storedBy: row.storedBy,
        items: []
      };
    }
    groups[row.recordId].items.push({
      itemId: row.itemId,
      action: row.action,
      name: row.name,
      code: row.code,
      quantity: row.quantity,
      price: row.price,
      remarks: row.remarks,
      sourceRow: row.sourceRow,
      partKey: row.itemKey
    });
  });
  return Object.keys(groups).map(function (recordId) {
    return groups[recordId];
  }).sort(function (a, b) {
    return new Date(b.storedAt).getTime() - new Date(a.storedAt).getTime();
  });
}

function REQ_buildItemsPreview_(items) {
  return (items || []).slice(0, APP_CONFIG.request.tempSelection.previewItemCount).map(function (item) {
    return item.action + ' ' + (item.name || item.code || 'Item') + ' x' + APP_toNumber_(item.quantity, 0);
  }).join(', ');
}

function REQ_buildItemsTakenSummary_(items) {
  return (items || []).map(function (item) {
    return item.action + ' ' + (item.name || item.code || 'Item') + ' x' + APP_toNumber_(item.quantity, 0);
  }).join(', ');
}

function REQ_promptForRequestedByName_(options) {
  const settings = options || {};
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    settings.title || 'Pull Stored Request',
    settings.message || 'Enter the Requested By name to pull:',
    ui.ButtonSet.OK_CANCEL
  );
  if (response.getSelectedButton() !== ui.Button.OK) {
    return APP_fail_('USER_CANCELLED', settings.cancelMessage || 'Stored request prompt cancelled by user.');
  }
  const requestedBy = APP_safeString_(response.getResponseText()).trim();
  if (APP_isBlank_(requestedBy)) {
    return APP_fail_('MISSING_REQUESTED_BY', 'Requested By is required.');
  }
  return APP_result_(true, 'Requested By captured.', { requestedBy: requestedBy });
}

function REQ_promptForStoredRecordSelection_(records, options) {
  const settings = options || {};
  const safeRecords = (records || []).slice(0, APP_CONFIG.request.tempSelection.maxPromptEntries);
  if (safeRecords.length === 0) {
    return APP_fail_('NO_STORED_RECORDS', 'No stored requests were found.');
  }
  if (safeRecords.length === 1) {
    return APP_result_(true, 'Single stored request resolved.', { record: safeRecords[0] });
  }

  const ui = SpreadsheetApp.getUi();
  const lines = [settings.listMessage || 'Choose which stored request to use by typing its number:', ''];
  safeRecords.forEach(function (record, index) {
    lines.push(
      (index + 1) + '. ' +
      LOG_getDisplayTimestamp_(record.storedAt) +
      ' | ' + record.items.length + ' item(s)' +
      ' | ' + REQ_buildItemsPreview_(record.items)
    );
  });
  const response = ui.prompt(settings.title || 'Select Stored Request', lines.join('\n'), ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() !== ui.Button.OK) {
    return APP_fail_('USER_CANCELLED', settings.cancelMessage || 'Stored request selection cancelled by user.');
  }
  const choice = APP_toNumber_(response.getResponseText(), 0);
  if (choice < 1 || choice > safeRecords.length) {
    return APP_fail_('INVALID_SELECTION', 'Choose a valid stored request number.');
  }
  return APP_result_(true, 'Stored request selected.', { record: safeRecords[choice - 1] });
}

function REQ_clearRequestedBy_(reqSheet) {
  (reqSheet || APP_getRequestSheet_()).getRange(REQ_CONFIG.REQUEST_FORM_CELLS.REQUESTED_BY).setValue('');
}

function REQ_getStoredRecordsByRequestedBy_(requestedBy) {
  const requestedByKey = APP_normalizeText_(requestedBy);
  return REQ_groupTempStoreRows_(REQ_getTempStoreRows_().filter(function (row) {
    return row.requestedByKey === requestedByKey;
  }));
}

function REQ_buildStoredRecordDetailLines_(record) {
  return (record.items || []).map(function (item) {
    return [
      item.action,
      item.name || item.code || 'Item',
      'x' + APP_toNumber_(item.quantity, 0),
      APP_isNoCode(item.code) ? 'NO CODE' : item.code
    ].join(' | ');
  });
}

function REQ_buildStoredRecordDialogEntry_(record) {
  return {
    recordId: record.recordId,
    requestedBy: record.requestedBy,
    storedAtDisplay: LOG_getDisplayTimestamp_(record.storedAt),
    storedBy: record.storedBy,
    itemCount: (record.items || []).length,
    preview: REQ_buildItemsPreview_(record.items),
    detailLines: REQ_buildStoredRecordDetailLines_(record)
  };
}

function REQ_getStoredRequestDialogData(query) {
  const normalizedQuery = APP_safeString_(query).trim().toLowerCase();
  const records = REQ_groupTempStoreRows_(REQ_getTempStoreRows_()).filter(function (record) {
    if (normalizedQuery === '') {
      return true;
    }
    const haystack = [
      APP_safeString_(record.requestedBy),
      APP_safeString_(record.storedBy),
      REQ_buildItemsPreview_(record.items),
      REQ_buildStoredRecordDetailLines_(record).join(' ')
    ].join(' ').toLowerCase();
    return haystack.indexOf(normalizedQuery) !== -1;
  }).map(REQ_buildStoredRecordDialogEntry_);
  return {
    ok: true,
    records: records,
    query: APP_safeString_(query)
  };
}

function REQ_openDeleteStoredRequestDialog_() {
  const template = HtmlService.createTemplateFromFile('StoredRequestDeleteDialog');
  template.dialogData = {
    title: 'Delete Temporary Request',
    initialData: REQ_getStoredRequestDialogData('')
  };
  const html = template.evaluate()
    .setWidth(720)
    .setHeight(560)
    .setTitle('Delete Temporary Request');
  SpreadsheetApp.getUi().showModalDialog(html, 'Delete Temporary Request');
  return APP_result_(true, 'Delete stored request dialog opened.', { deferred: true });
}

function REQ_searchStoredRequests(query) {
  return REQ_getStoredRequestDialogData(query);
}

function REQ_deleteStoredRequestById(recordId) {
  const groupedRecords = REQ_groupTempStoreRows_(REQ_getTempStoreRows_());
  let targetRecord = null;
  for (let index = 0; index < groupedRecords.length; index += 1) {
    if (groupedRecords[index].recordId === recordId) {
      targetRecord = groupedRecords[index];
      break;
    }
  }
  if (!targetRecord) {
    return APP_fail_('NO_STORED_RECORD', 'Stored request not found.');
  }

  const deleteCount = REQ_deleteTempStoreRecord_(targetRecord.recordId);
  if (deleteCount <= 0) {
    return APP_fail_('DELETE_FAILED', 'The stored request could not be deleted.');
  }

  const actor = APP_getActor_();
  LOG_requestFormAction('DELETE_TEMP_REQUEST', {
    requestedBy: targetRecord.requestedBy,
    actor: actor
  }, {
    itemName: 'Deleted Stored Request',
    code: '',
    quantity: targetRecord.items.length,
    remarks: REQ_buildItemsPreview_(targetRecord.items),
    timestamp: APP_now_()
  });

  return APP_result_(true, 'Stored request deleted.', {
    recordId: targetRecord.recordId,
    records: REQ_getStoredRequestDialogData('').records
  });
}

function REQ_confirmDeleteStoredRecord_(record) {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'Delete Stored Request?',
    [
      'Requested By: ' + APP_safeString_(record.requestedBy),
      'Stored: ' + LOG_getDisplayTimestamp_(record.storedAt),
      'Items: ' + APP_safeString_(record.items.length),
      'Preview: ' + REQ_buildItemsPreview_(record.items)
    ].join('\n'),
    ui.ButtonSet.YES_NO
  );
  return response === ui.Button.YES
    ? APP_result_(true, 'Stored request deletion confirmed.')
    : APP_fail_('USER_CANCELLED', 'Delete stored request cancelled by user.');
}

function REQ_deleteStoredRequest() {
  return REQ_openDeleteStoredRequestDialog_();
}

function REQ_confirmReplacingCurrentStaging_(itemCount) {
  if (itemCount <= 0) {
    return APP_result_(true, 'No existing staged items to replace.');
  }
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'Replace Current Staging?',
    'There are currently ' + itemCount + ' staged item(s). Pulling a stored request will replace them. Continue?',
    ui.ButtonSet.YES_NO
  );
  return response === ui.Button.YES
    ? APP_result_(true, 'Existing staging can be replaced.')
    : APP_fail_('USER_CANCELLED', 'Pull stored request cancelled by user.');
}

function REQ_buildSearchMatch_(record) {
  return {
    name: record.name,
    code: APP_isNoCode(record.code) ? 'NO CODE' : record.code,
    qty: record.qty,
    price: record.price,
    dateInv: record.dateInv,
    remarks: record.remarks,
    rowNumber: record.rowNumber,
    itemKey: record.key
  };
}

function REQ_prepareSearchBody_(sheet) {
  const targetSheet = sheet || APP_getRequestSheet_();
  const rowCount = REQ_CONFIG.SEARCH_RESULTS_TABLE.maxResults;
  const range = targetSheet.getRange(REQ_CONFIG.SEARCH_RESULTS_TABLE.startRow, REQ_CONFIG.SEARCH_RESULTS_TABLE.cols.NAME, rowCount, REQ_CONFIG.SEARCH_RESULTS_TABLE.visibleWidth);
  const backgrounds = [];
  for (let rowIndex = 0; rowIndex < rowCount; rowIndex += 1) {
    const color = rowIndex % 2 === 0 ? APP_CONFIG.system.colors.rowAlt1 : APP_CONFIG.system.colors.rowAlt2;
    backgrounds.push([
      color,
      color,
      color,
      color,
      color,
      color,
      color,
      APP_CONFIG.system.colors.successBg
    ]);
  }
  range
    .setBackgrounds(backgrounds)
    .setWrap(true)
    .setVerticalAlignment('middle')
    .setBorder(true, true, true, true, true, true, APP_CONFIG.system.colors.border, SpreadsheetApp.BorderStyle.SOLID);
  targetSheet.getRange(REQ_CONFIG.SEARCH_RESULTS_TABLE.startRow, REQ_CONFIG.SEARCH_RESULTS_TABLE.cols.QTY, rowCount, 1).setHorizontalAlignment('center');
  targetSheet.getRange(REQ_CONFIG.SEARCH_RESULTS_TABLE.startRow, REQ_CONFIG.SEARCH_RESULTS_TABLE.cols.PRICE, rowCount, 1).setNumberFormat('#,##0.00');
  targetSheet.getRange(REQ_CONFIG.SEARCH_RESULTS_TABLE.startRow, REQ_CONFIG.SEARCH_RESULTS_TABLE.cols.ROW, rowCount, 1).setHorizontalAlignment('center');
  targetSheet.getRange(REQ_CONFIG.SEARCH_RESULTS_TABLE.startRow, REQ_CONFIG.SEARCH_RESULTS_TABLE.cols.SELECT, rowCount, 1)
    .setHorizontalAlignment('center')
    .setFontWeight('bold')
    .setFontColor(APP_CONFIG.system.colors.successText);
}

function REQ_prepareItemsBody_(sheet) {
  const targetSheet = sheet || APP_getRequestSheet_();
  const rowCount = REQ_CONFIG.SEARCH_RESULTS_TABLE.maxResults;
  const range = targetSheet.getRange(REQ_CONFIG.ITEMS_TABLE.startRow, REQ_CONFIG.ITEMS_TABLE.cols.ACTION, rowCount, REQ_CONFIG.ITEMS_TABLE.visibleWidth);
  const backgrounds = [];
  for (let rowIndex = 0; rowIndex < rowCount; rowIndex += 1) {
    const color = rowIndex % 2 === 0 ? APP_CONFIG.system.colors.rowAlt1 : APP_CONFIG.system.colors.rowAlt2;
    backgrounds.push([
      APP_CONFIG.system.colors.accentBg,
      color,
      color,
      color,
      color,
      color,
      color,
      APP_CONFIG.system.colors.dangerBg
    ]);
  }
  range
    .setBackgrounds(backgrounds)
    .setWrap(true)
    .setVerticalAlignment('middle')
    .setBorder(true, true, true, true, true, true, APP_CONFIG.system.colors.border, SpreadsheetApp.BorderStyle.SOLID);
  targetSheet.getRange(REQ_CONFIG.ITEMS_TABLE.startRow, REQ_CONFIG.ITEMS_TABLE.cols.ACTION, rowCount, 1)
    .setFontWeight('bold')
    .setFontColor(APP_CONFIG.system.colors.accentText)
    .setHorizontalAlignment('center');
  targetSheet.getRange(REQ_CONFIG.ITEMS_TABLE.startRow, REQ_CONFIG.ITEMS_TABLE.cols.QTY, rowCount, 1).setHorizontalAlignment('center');
  targetSheet.getRange(REQ_CONFIG.ITEMS_TABLE.startRow, REQ_CONFIG.ITEMS_TABLE.cols.PRICE, rowCount, 1).setNumberFormat('#,##0.00');
  targetSheet.getRange(REQ_CONFIG.ITEMS_TABLE.startRow, REQ_CONFIG.ITEMS_TABLE.cols.ROW, rowCount, 1).setHorizontalAlignment('center');
  targetSheet.getRange(REQ_CONFIG.ITEMS_TABLE.startRow, REQ_CONFIG.ITEMS_TABLE.cols.REMOVE, rowCount, 1)
    .setHorizontalAlignment('center')
    .setFontWeight('bold')
    .setFontColor(APP_CONFIG.system.colors.dangerText);
}

function REQ_clearSearchResults(reqSheet) {
  const sheet = reqSheet || APP_getRequestSheet_();
  sheet.getRange(REQ_CONFIG.SEARCH_RESULTS_TABLE.startRow, REQ_CONFIG.SEARCH_RESULTS_TABLE.cols.NAME, REQ_CONFIG.SEARCH_RESULTS_TABLE.maxResults, REQ_CONFIG.SEARCH_RESULTS_TABLE.visibleWidth).clearContent();
  sheet.getRange(REQ_CONFIG.SEARCH_RESULTS_TABLE.startRow, REQ_CONFIG.SEARCH_RESULTS_TABLE.cols.ITEM_KEY, REQ_CONFIG.SEARCH_RESULTS_TABLE.maxResults, 1).clearContent();
}

function REQ_displaySearchResults(reqSheet, matches) {
  const sheet = reqSheet || APP_getRequestSheet_();
  REQ_clearSearchResults(sheet);
  const safeMatches = (matches || []).slice(0, REQ_CONFIG.SEARCH_RESULTS_TABLE.maxResults);
  if (safeMatches.length === 0) {
    return;
  }
  const visibleRows = safeMatches.map(function (item) {
    return [
      item.name,
      item.code,
      item.qty,
      item.price,
      item.dateInv,
      item.remarks,
      item.rowNumber,
      ''
    ];
  });
  const hiddenKeys = safeMatches.map(function (item) {
    return [item.itemKey];
  });
  sheet.getRange(REQ_CONFIG.SEARCH_RESULTS_TABLE.startRow, REQ_CONFIG.SEARCH_RESULTS_TABLE.cols.NAME, visibleRows.length, REQ_CONFIG.SEARCH_RESULTS_TABLE.visibleWidth).setValues(visibleRows);
  sheet.getRange(REQ_CONFIG.SEARCH_RESULTS_TABLE.startRow, REQ_CONFIG.SEARCH_RESULTS_TABLE.cols.ITEM_KEY, hiddenKeys.length, 1).setValues(hiddenKeys);
  sheet.hideColumns(REQ_CONFIG.SEARCH_RESULTS_TABLE.cols.ITEM_KEY);
}

function REQ_runInventorySearch_(records) {
  const rawRecords = (records || []).slice(0, REQ_CONFIG.SEARCH_RESULTS_TABLE.maxResults);
  const matches = rawRecords.map(REQ_buildSearchMatch_);
  REQ_displaySearchResults(APP_getRequestSheet_(), matches);
  return APP_result_(true, 'Search completed.', { matches: matches, records: rawRecords });
}

function REQ_searchByName(searchTerm) {
  const normalized = APP_safeString_(searchTerm).trim();
  if (normalized === '') {
    REQ_clearSearchResults(APP_getRequestSheet_());
    return APP_result_(true, 'Search cleared.', { matches: [] });
  }
  return REQ_runInventorySearch_(APP_searchInventoryByName_(normalized, REQ_CONFIG.SEARCH_RESULTS_TABLE.maxResults));
}

function REQ_searchByCode(searchCode) {
  const normalized = APP_normalizeText_(searchCode);
  if (normalized === '' || APP_isNoCode(normalized)) {
    REQ_clearSearchResults(APP_getRequestSheet_());
    return APP_result_(true, 'Search cleared.', { matches: [] });
  }
  return REQ_runInventorySearch_(APP_findInventoryByCode_(normalized));
}

function REQ_searchByRow(rowNumber) {
  const safeRow = APP_toNumber_(rowNumber, 0);
  if (safeRow < APP_CONFIG.inventory.startRow) {
    REQ_clearSearchResults(APP_getRequestSheet_());
    return APP_result_(true, 'Search cleared.', { matches: [] });
  }
  const record = APP_findInventoryByRow_(safeRow);
  return REQ_runInventorySearch_(record ? [record] : []);
}

function REQ_ensureRequestSheetSize_(reqSheet) {
  const sheet = reqSheet || APP_getRequestSheet_();
  const requiredRows = REQ_CONFIG.SEARCH_RESULTS_TABLE.startRow + REQ_CONFIG.SEARCH_RESULTS_TABLE.maxResults + 5;
  const requiredColumns = REQ_CONFIG.ITEMS_TABLE.cols.ITEM_KEY;
  if (sheet.getMaxRows() < requiredRows) {
    sheet.insertRowsAfter(sheet.getMaxRows(), requiredRows - sheet.getMaxRows());
  }
  if (sheet.getMaxColumns() < requiredColumns) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), requiredColumns - sheet.getMaxColumns());
  }
  return sheet;
}

function REQ_ensureRequestFormSheet() {
  const ss = APP_getSpreadsheet_();
  let sheet = ss.getSheetByName(REQ_CONFIG.SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(REQ_CONFIG.SHEET_NAME);
  }
  return REQ_ensureRequestSheetSize_(sheet);
}

function REQ_initializeFormLayout() {
  const sheet = REQ_ensureRequestFormSheet();
  const layout = APP_CONFIG.request.layout;
  try {
    sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).breakApart();
  } catch (error) {
    console.log('Request layout unmerge skipped: ' + error.message);
  }
  sheet.clear();
  sheet.clearFormats();

  sheet.getRange(layout.formTitleRange).merge()
    .setValue(layout.formTitle)
    .setFontSize(15)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setBackground(APP_CONFIG.system.colors.titleBg)
    .setFontColor(APP_CONFIG.system.colors.titleText)
    .setBorder(true, true, true, true, true, true, APP_CONFIG.system.colors.border, SpreadsheetApp.BorderStyle.SOLID);

  sheet.getRange(layout.workflowRange).merge()
    .setValue(layout.workflowText)
    .setWrap(true)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setBackground(APP_CONFIG.system.colors.accentBg)
    .setFontColor(APP_CONFIG.system.colors.accentText)
    .setBorder(true, true, true, true, true, true, APP_CONFIG.system.colors.border, SpreadsheetApp.BorderStyle.SOLID);

  sheet.getRange(layout.stagedTitleRange).merge()
    .setValue(layout.stagedTitleText)
    .setFontSize(14)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setBackground(APP_CONFIG.system.colors.titleBg)
    .setFontColor(APP_CONFIG.system.colors.titleText)
    .setBorder(true, true, true, true, true, true, APP_CONFIG.system.colors.border, SpreadsheetApp.BorderStyle.SOLID);

  sheet.getRange(layout.stagedWorkflowRange).merge()
    .setValue(layout.stagedWorkflowText)
    .setWrap(true)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setBackground(APP_CONFIG.system.colors.accentBg)
    .setFontColor(APP_CONFIG.system.colors.accentText)
    .setBorder(true, true, true, true, true, true, APP_CONFIG.system.colors.border, SpreadsheetApp.BorderStyle.SOLID);

  sheet.getRange(layout.searchHeaderRange).merge()
    .setValue(layout.searchHeaderText)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setBackground(APP_CONFIG.system.colors.sectionBg)
    .setFontColor(APP_CONFIG.system.colors.sectionText)
    .setBorder(true, true, true, true, true, true, APP_CONFIG.system.colors.border, SpreadsheetApp.BorderStyle.SOLID);

  sheet.getRange('A1').setValue(APP_CONFIG.labels.requestedBy);
  sheet.getRange('A3').setValue('Code');
  sheet.getRange('A4').setValue('Name');
  sheet.getRange('A5').setValue('Row');
  sheet.getRange('A1:A5').setFontWeight('bold').setBackground('#f1f3f4').setHorizontalAlignment('center');
  sheet.getRange('A1:B5').setBorder(true, true, true, true, true, true, APP_CONFIG.system.colors.border, SpreadsheetApp.BorderStyle.SOLID).setVerticalAlignment('middle');
  sheet.getRange('B1:B5').setBackground('#ffffff');

  sheet.getRange(layout.searchSectionRange).merge()
    .setValue(layout.searchSectionText)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setBackground(APP_CONFIG.system.colors.sectionBg)
    .setFontColor(APP_CONFIG.system.colors.sectionText)
    .setBorder(true, true, true, true, true, true, APP_CONFIG.system.colors.border, SpreadsheetApp.BorderStyle.SOLID);

  sheet.getRange(layout.searchHelperRange).merge()
    .setValue(layout.searchHelperText)
    .setWrap(true)
    .setBackground(APP_CONFIG.system.colors.noteBg)
    .setHorizontalAlignment('left')
    .setBorder(true, true, true, true, true, true, APP_CONFIG.system.colors.border, SpreadsheetApp.BorderStyle.SOLID);

  sheet.getRange(layout.stagedSectionRange).merge()
    .setValue(layout.stagedSectionText)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setBackground(APP_CONFIG.system.colors.sectionBg)
    .setFontColor(APP_CONFIG.system.colors.sectionText)
    .setBorder(true, true, true, true, true, true, APP_CONFIG.system.colors.border, SpreadsheetApp.BorderStyle.SOLID);

  sheet.getRange(layout.stagedHelperRange).merge()
    .setValue(layout.stagedHelperText)
    .setWrap(true)
    .setBackground('#e8f0fe')
    .setHorizontalAlignment('left')
    .setBorder(true, true, true, true, true, true, APP_CONFIG.system.colors.border, SpreadsheetApp.BorderStyle.SOLID);

  sheet.getRange(1, APP_CONFIG.request.separatorColumn, REQ_CONFIG.SEARCH_RESULTS_TABLE.startRow + REQ_CONFIG.SEARCH_RESULTS_TABLE.maxResults, 1)
    .setBackground(APP_CONFIG.system.colors.separatorBg)
    .setBorder(false, true, false, true, false, false, APP_CONFIG.system.colors.separatorBorder, SpreadsheetApp.BorderStyle.SOLID_THICK);
  sheet.setColumnWidth(APP_CONFIG.request.separatorColumn, 22);

  sheet.setColumnWidth(1, 150);
  sheet.setColumnWidth(2, 220);
  sheet.setColumnWidth(REQ_CONFIG.SEARCH_RESULTS_TABLE.cols.NAME, 220);
  sheet.setColumnWidth(REQ_CONFIG.SEARCH_RESULTS_TABLE.cols.CODE, 140);
  sheet.setColumnWidth(REQ_CONFIG.SEARCH_RESULTS_TABLE.cols.QTY, 70);
  sheet.setColumnWidth(REQ_CONFIG.SEARCH_RESULTS_TABLE.cols.PRICE, 90);
  sheet.setColumnWidth(REQ_CONFIG.SEARCH_RESULTS_TABLE.cols.DATE_INV, 110);
  sheet.setColumnWidth(REQ_CONFIG.SEARCH_RESULTS_TABLE.cols.REMARKS, 240);
  sheet.setColumnWidth(REQ_CONFIG.SEARCH_RESULTS_TABLE.cols.ROW, 60);
  sheet.setColumnWidth(REQ_CONFIG.SEARCH_RESULTS_TABLE.cols.SELECT, 55);

  sheet.setColumnWidth(REQ_CONFIG.ITEMS_TABLE.cols.ACTION, 95);
  sheet.setColumnWidth(REQ_CONFIG.ITEMS_TABLE.cols.NAME, 220);
  sheet.setColumnWidth(REQ_CONFIG.ITEMS_TABLE.cols.CODE, 140);
  sheet.setColumnWidth(REQ_CONFIG.ITEMS_TABLE.cols.QTY, 70);
  sheet.setColumnWidth(REQ_CONFIG.ITEMS_TABLE.cols.PRICE, 90);
  sheet.setColumnWidth(REQ_CONFIG.ITEMS_TABLE.cols.REMARKS, 240);
  sheet.setColumnWidth(REQ_CONFIG.ITEMS_TABLE.cols.ROW, 60);
  sheet.setColumnWidth(REQ_CONFIG.ITEMS_TABLE.cols.REMOVE, 70);

  sheet.setFrozenRows(8);
  sheet.setRowHeight(1, 28);
  sheet.setRowHeight(2, 24);
  sheet.setRowHeight(6, 22);
  sheet.setRowHeight(7, 24);
  return sheet;
}

function REQ_initializeSearchResultsTable() {
  const sheet = REQ_ensureRequestFormSheet();
  sheet.getRange(REQ_CONFIG.SEARCH_RESULTS_TABLE.headerRow, REQ_CONFIG.SEARCH_RESULTS_TABLE.cols.NAME, 1, REQ_CONFIG.SEARCH_RESULTS_TABLE.visibleWidth)
    .setValues([REQ_CONFIG.SEARCH_RESULTS_TABLE.visibleHeaders])
    .setFontWeight('bold')
    .setBackground('#7a1f1f')
    .setFontColor('#ffffff')
    .setHorizontalAlignment('center')
    .setWrap(true)
    .setBorder(true, true, true, true, true, true, APP_CONFIG.system.colors.border, SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange(REQ_CONFIG.SEARCH_RESULTS_TABLE.headerRow, REQ_CONFIG.SEARCH_RESULTS_TABLE.cols.SELECT)
    .setBackground('#137333')
    .setFontColor('#ffffff')
    .setNote('Type X in this column to stage an item.');
  REQ_prepareSearchBody_(sheet);
  sheet.hideColumns(REQ_CONFIG.SEARCH_RESULTS_TABLE.cols.ITEM_KEY);
}

function REQ_initializeItemsTakenTable() {
  const sheet = REQ_ensureRequestFormSheet();
  sheet.getRange(REQ_CONFIG.ITEMS_TABLE.headerRow, REQ_CONFIG.ITEMS_TABLE.cols.ACTION, 1, REQ_CONFIG.ITEMS_TABLE.visibleWidth)
    .setValues([REQ_CONFIG.ITEMS_TABLE.visibleHeaders])
    .setFontWeight('bold')
    .setBackground('#1a73e8')
    .setFontColor('#ffffff')
    .setHorizontalAlignment('center')
    .setWrap(true)
    .setBorder(true, true, true, true, true, true, APP_CONFIG.system.colors.border, SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange(REQ_CONFIG.ITEMS_TABLE.headerRow, REQ_CONFIG.ITEMS_TABLE.cols.REMOVE)
    .setBackground('#a50e0e')
    .setFontColor('#ffffff')
    .setNote('Type X in this column to remove a staged item.');
  REQ_prepareItemsBody_(sheet);
  sheet.hideColumns(REQ_CONFIG.ITEMS_TABLE.cols.ITEM_ID, 3);
}

function REQ_clearEntryFields_(reqSheet) {
  const sheet = reqSheet || APP_getRequestSheet_();
  sheet.getRange(REQ_CONFIG.REQUEST_FORM_CELLS.SEARCH_CODE).setValue('');
  sheet.getRange(REQ_CONFIG.REQUEST_FORM_CELLS.SEARCH_NAME).setValue('');
  sheet.getRange(REQ_CONFIG.REQUEST_FORM_CELLS.SEARCH_ROW).setValue('');
  REQ_clearSearchResults(sheet);
}

function REQ_refreshItemsTaken(sessionId) {
  const sheet = APP_getRequestSheet_();
  sheet.getRange(REQ_CONFIG.ITEMS_TABLE.startRow, REQ_CONFIG.ITEMS_TABLE.cols.ACTION, REQ_CONFIG.SEARCH_RESULTS_TABLE.maxResults, REQ_CONFIG.ITEMS_TABLE.visibleWidth).clearContent();
  sheet.getRange(REQ_CONFIG.ITEMS_TABLE.startRow, REQ_CONFIG.ITEMS_TABLE.cols.ITEM_ID, REQ_CONFIG.SEARCH_RESULTS_TABLE.maxResults, 3).clearContent();

  let resolvedSessionId = sessionId || '';
  if (APP_isBlank_(resolvedSessionId)) {
    try {
      resolvedSessionId = REQ_getCurrentSessionId();
    } catch (error) {
      resolvedSessionId = '';
    }
  }
  if (APP_isBlank_(resolvedSessionId)) {
    return APP_result_(true, 'No active staging session.', { items: [] });
  }

  const items = REQ_getSessionItems(resolvedSessionId);
  if (items.length === 0) {
    return APP_result_(true, 'No staged items.', { items: [] });
  }

  const safeItems = items.slice(0, REQ_CONFIG.SEARCH_RESULTS_TABLE.maxResults);
  const visibleRows = safeItems.map(function (item) {
    return [
      item.action,
      item.name,
      APP_isNoCode(item.code) ? 'NO CODE' : item.code,
      item.quantity,
      item.price,
      item.remarks,
      item.sourceRow,
      ''
    ];
  });
  const hiddenRows = safeItems.map(function (item) {
    return [item.itemId, resolvedSessionId, item.partKey];
  });

  sheet.getRange(REQ_CONFIG.ITEMS_TABLE.startRow, REQ_CONFIG.ITEMS_TABLE.cols.ACTION, visibleRows.length, REQ_CONFIG.ITEMS_TABLE.visibleWidth).setValues(visibleRows);
  sheet.getRange(REQ_CONFIG.ITEMS_TABLE.startRow, REQ_CONFIG.ITEMS_TABLE.cols.ITEM_ID, hiddenRows.length, 3).setValues(hiddenRows);
  sheet.hideColumns(REQ_CONFIG.ITEMS_TABLE.cols.ITEM_ID, 3);
  return APP_result_(true, 'Staged items refreshed.', { items: safeItems });
}

function REQ_openStageActionDialog_(record, options) {
  const settings = options || {};
  return APP_showActionDialog_({
    title: 'Stage Item',
    requestedBy: settings.requestedBy || APP_safeString_(APP_getRequestSheet_().getRange(REQ_CONFIG.REQUEST_FORM_CELLS.REQUESTED_BY).getValue()).trim(),
    record: record,
    quantityLabel: 'Quantity',
    submitPayload: {
      target: settings.target || 'REQUEST_IMMEDIATE',
      itemKey: record.key,
      sourceRow: record.rowNumber,
      clearEntryFields: settings.clearEntryFields !== false,
      markerRow: settings.markerRow || 0
    },
    cancelPayload: {
      target: settings.target || 'REQUEST_IMMEDIATE',
      markerRow: settings.markerRow || 0
    }
  });
}

function REQ_stageInventoryRecord_(record, action, quantity, options) {
  const settings = options || {};
  const sheet = settings.reqSheet || APP_getRequestSheet_();
  const context = REQ_getRequestContext_(sheet);
  const validation = REQ_validateContext_(context);
  if (!validation.ok) {
    return validation;
  }
  if (!APP_validateAction_(action)) {
    return APP_fail_('INVALID_ACTION', 'Action must be ADD or SUBTRACT.');
  }
  const safeQuantity = APP_toNumber_(quantity, 0);
  if (safeQuantity <= 0) {
    return APP_fail_('INVALID_QUANTITY', 'Quantity must be greater than zero.');
  }
  if (action === APP_CONFIG.actions.SUBTRACT && !REQ_CONFIG.FLAGS.allowNegativeStock && safeQuantity > APP_toNumber_(record.qty, 0)) {
    return APP_fail_('INSUFFICIENT_STOCK', 'Cannot subtract more than the available stock.');
  }

  const now = APP_now_();
  const sessionId = REQ_getOrCreateSessionId_(context);
  const previousRecord = REQ_getSessionRecord_(sessionId);
  const nextRecord = previousRecord || {
    sessionId: sessionId,
    contextKey: context.contextKey,
    requestedBy: context.requestedBy,
    actor: context.actor.display,
    createdAt: now,
    updatedAt: now,
    items: []
  };
  const stagedItem = REQ_normalizeSessionItem_({
    itemId: REQ_generateItemId_(),
    partKey: record.key,
    name: record.name,
    code: record.code,
    quantity: safeQuantity,
    price: record.price,
    remarks: record.remarks,
    sourceRow: record.rowNumber,
    action: action,
    createdAt: now,
    updatedAt: now
  });
  nextRecord.items = nextRecord.items.concat([stagedItem]);
  nextRecord.updatedAt = now;

  return APP_withScriptLock_(30000, function () {
    REQ_writeSessionRecord_(nextRecord);
    const logRow = LOG_requestFormAction('STAGE_REQUEST_ITEM', context, {
      itemName: stagedItem.name,
      code: stagedItem.code,
      quantity: stagedItem.quantity,
      remarks: stagedItem.action + ' | ' + stagedItem.remarks,
      timestamp: now
    });
    if (logRow <= 0) {
      if (previousRecord) {
        REQ_writeSessionRecord_(previousRecord);
      } else {
        REQ_removeSessionRecord_(sessionId, context.contextKey);
      }
      return APP_fail_('LOG_WRITE_FAILED', 'Could not log the staged item.');
    }

    REQ_refreshItemsTaken(sessionId);
    if (settings.clearEntryFields !== false) {
      REQ_clearEntryFields_(sheet);
    }
    return APP_result_(true, 'Item staged successfully.', {
      sessionId: sessionId,
      item: stagedItem
    });
  });
}

function REQ_removeStagedItem_(sessionId, itemId, context) {
  return APP_withScriptLock_(30000, function () {
    const previousRecord = REQ_getSessionRecord_(sessionId);
    if (!previousRecord) {
      return APP_fail_('NO_SESSION', 'Staged request not found.');
    }

    let removedItem = null;
    const remainingItems = previousRecord.items.filter(function (item) {
      if (!removedItem && item.itemId === itemId) {
        removedItem = item;
        return false;
      }
      return true;
    });
    if (!removedItem) {
      return APP_fail_('ITEM_NOT_FOUND', 'Staged item not found.');
    }

    if (remainingItems.length === 0) {
      REQ_removeSessionRecord_(previousRecord.sessionId, previousRecord.contextKey);
    } else {
      const nextRecord = Object.assign({}, previousRecord, {
        items: remainingItems,
        updatedAt: APP_now_()
      });
      REQ_writeSessionRecord_(nextRecord);
    }

    const logRow = LOG_requestFormAction('REMOVE_REQUEST_ITEM', context, {
      itemName: removedItem.name,
      code: removedItem.code,
      quantity: removedItem.quantity,
      remarks: removedItem.action + ' | ' + removedItem.remarks,
      timestamp: APP_now_()
    });
    if (logRow <= 0) {
      REQ_writeSessionRecord_(previousRecord);
      return APP_fail_('LOG_WRITE_FAILED', 'Could not log the removed staged item.');
    }

    return APP_result_(true, 'Staged item removed.', removedItem);
  });
}

function REQ_getMarkedSearchSelection_(reqSheet) {
  const sheet = reqSheet || APP_getRequestSheet_();
  const rowCount = REQ_CONFIG.SEARCH_RESULTS_TABLE.maxResults;
  const visibleValues = sheet.getRange(REQ_CONFIG.SEARCH_RESULTS_TABLE.startRow, REQ_CONFIG.SEARCH_RESULTS_TABLE.cols.NAME, rowCount, REQ_CONFIG.SEARCH_RESULTS_TABLE.visibleWidth).getValues();
  const itemKeys = sheet.getRange(REQ_CONFIG.SEARCH_RESULTS_TABLE.startRow, REQ_CONFIG.SEARCH_RESULTS_TABLE.cols.ITEM_KEY, rowCount, 1).getValues();
  const markedRows = visibleValues.map(function (row, index) {
    return {
      rowNumber: REQ_CONFIG.SEARCH_RESULTS_TABLE.startRow + index,
      selectMark: APP_normalizeText_(row[REQ_CONFIG.SEARCH_RESULTS_TABLE.cols.SELECT - 1]),
      sourceRow: APP_toNumber_(row[REQ_CONFIG.SEARCH_RESULTS_TABLE.cols.ROW - 1], 0),
      itemKey: APP_safeString_(itemKeys[index][0]).trim()
    };
  }).filter(function (row) {
    return row.selectMark === 'X';
  });

  if (markedRows.length === 0) {
    return APP_fail_('NO_SELECTION', 'Mark one result row with X first.');
  }
  if (markedRows.length > 1) {
    return APP_fail_('MULTIPLE_SELECTIONS', 'Mark only one result row with X at a time.');
  }
  const selected = markedRows[0];
  const record = APP_resolveInventoryRecord_(APP_getInventorySheet_(), selected.itemKey, { sourceRow: selected.sourceRow });
  if (!record) {
    return APP_fail_('ITEM_NOT_FOUND', 'Selected result row could not be resolved.');
  }
  return APP_result_(true, 'Result row resolved.', {
    record: record,
    markerRow: selected.rowNumber
  });
}

function REQ_detectRemoval(e) {
  if (!e || !e.range) {
    return null;
  }
  const sheet = e.range.getSheet();
  if (sheet.getName() !== REQ_CONFIG.SHEET_NAME) {
    return null;
  }
  const row = e.range.getRow();
  const col = e.range.getColumn();
  if (col !== REQ_CONFIG.ITEMS_TABLE.cols.REMOVE || row < REQ_CONFIG.ITEMS_TABLE.startRow) {
    return null;
  }
  if (APP_normalizeText_(e.range.getValue()) !== 'X') {
    return null;
  }
  e.range.setValue('');

  const itemId = APP_safeString_(sheet.getRange(row, REQ_CONFIG.ITEMS_TABLE.cols.ITEM_ID).getValue()).trim();
  const sessionId = APP_safeString_(sheet.getRange(row, REQ_CONFIG.ITEMS_TABLE.cols.SESSION_ID).getValue()).trim();
  if (APP_isBlank_(itemId) || APP_isBlank_(sessionId)) {
    return APP_fail_('ITEM_NOT_FOUND', 'The staged item could not be removed.');
  }

  const context = REQ_getRequestContext_(sheet);
  const removeResult = REQ_removeStagedItem_(sessionId, itemId, context);
  if (removeResult.ok) {
    REQ_refreshItemsTaken(sessionId);
  }
  return removeResult;
}

function REQ_handleSearchSelectionEdit_(e) {
  if (!e || !e.range) {
    return null;
  }
  const sheet = e.range.getSheet();
  if (sheet.getName() !== REQ_CONFIG.SHEET_NAME) {
    return null;
  }
  if (e.range.getColumn() !== REQ_CONFIG.SEARCH_RESULTS_TABLE.cols.SELECT || e.range.getRow() < REQ_CONFIG.SEARCH_RESULTS_TABLE.startRow) {
    return null;
  }
  if (APP_normalizeText_(e.range.getValue()) !== 'X') {
    return null;
  }

  const selectionResult = REQ_getMarkedSearchSelection_(sheet);
  if (!selectionResult.ok) {
    e.range.setValue('');
    return selectionResult;
  }

  return REQ_openStageActionDialog_(selectionResult.data.record, {
    requestedBy: APP_safeString_(sheet.getRange(REQ_CONFIG.REQUEST_FORM_CELLS.REQUESTED_BY).getValue()).trim(),
    target: 'REQUEST_SELECTION',
    markerRow: selectionResult.data.markerRow,
    clearEntryFields: false
  });
}

function REQ_handleImmediateResolution_(sheet, searchResult, options) {
  const settings = options || {};
  if (!searchResult.ok) {
    REQ_clearSearchResults(sheet);
    return searchResult;
  }
  const rawRecords = searchResult.data.records || [];
  const displayMatches = searchResult.data.matches || rawRecords.map(REQ_buildSearchMatch_);
  REQ_displaySearchResults(sheet, displayMatches);
  if (rawRecords.length === 0) {
    return APP_fail_('ITEM_NOT_FOUND', 'No inventory item matched.');
  }
  if (rawRecords.length > 1) {
    return APP_fail_('MULTIPLE_RESULTS', settings.multipleMessage || 'Multiple items matched. Use ROW or X to choose the correct item.');
  }

  return REQ_openStageActionDialog_(rawRecords[0], {
    requestedBy: APP_safeString_(sheet.getRange(REQ_CONFIG.REQUEST_FORM_CELLS.REQUESTED_BY).getValue()).trim(),
    target: 'REQUEST_IMMEDIATE',
    clearEntryFields: true
  });
}

function REQ_clearItemsTaken() {
  let sessionId = '';
  try {
    sessionId = REQ_getCurrentSessionId();
  } catch (error) {
    REQ_refreshItemsTaken('');
    REQ_clearEntryFields_(APP_getRequestSheet_());
    return APP_result_(true, 'No active staged request to clear.');
  }
  const result = REQ_clearSession(sessionId);
  REQ_refreshItemsTaken('');
  REQ_clearEntryFields_(APP_getRequestSheet_());
  APP_safeToast_(result.message, 'Request');
  return result;
}

function REQ_storeTemporarily() {
  const sheet = APP_getRequestSheet_();
  const context = REQ_getRequestContext_(sheet);
  const validation = REQ_validateContext_(context);
  if (!validation.ok) {
    APP_safeToast_(validation.message, 'Request');
    return validation;
  }

  let sessionId;
  try {
    sessionId = REQ_getCurrentSessionId();
  } catch (error) {
    const result = APP_fail_('NO_SESSION', 'No staged request found to store.');
    APP_safeToast_(result.message, 'Request');
    return result;
  }

  const sessionRecord = REQ_getSessionRecord_(sessionId);
  const items = sessionRecord ? sessionRecord.items : [];
  const previousTempRecordId = sessionRecord && sessionRecord.sourceTempRecordId ? sessionRecord.sourceTempRecordId : '';
  if (items.length === 0) {
    const result = APP_fail_('EMPTY_SESSION', 'No staged items found to store.');
    APP_safeToast_(result.message, 'Request');
    return result;
  }

  const recordId = REQ_generateTempRecordId_();
  const storedAt = APP_now_();
  const rows = items.map(function (item) {
    return {
      recordId: recordId,
      requestedBy: context.requestedBy,
      requestedByKey: context.requestedByKey,
      storedAt: storedAt,
      storedBy: context.actor.display,
      itemId: item.itemId,
      action: item.action,
      name: item.name,
      code: item.code,
      quantity: item.quantity,
      price: item.price,
      remarks: item.remarks,
      sourceRow: item.sourceRow,
      itemKey: item.partKey
    };
  });

  const appendInfo = REQ_appendTempStoreRows_(rows);
  const logRow = LOG_requestFormAction('STORE_TEMP_REQUEST', context, {
    itemName: 'Stored Request',
    code: '',
    quantity: items.length,
    remarks: REQ_buildItemsPreview_(items),
    timestamp: storedAt
  });
  if (logRow <= 0) {
    REQ_deleteTempStoreRecord_(recordId);
    const result = APP_fail_('LOG_WRITE_FAILED', 'Could not log the stored request.');
    APP_safeToast_(result.message, 'Request');
    return result;
  }

  if (!APP_isBlank_(previousTempRecordId) && previousTempRecordId !== recordId) {
    REQ_deleteTempStoreRecord_(previousTempRecordId);
  }

  REQ_clearSession(sessionId);
  REQ_refreshItemsTaken('');
  REQ_clearEntryFields_(sheet);
  REQ_clearRequestedBy_(sheet);
  const result = APP_result_(true, 'Staged request stored temporarily.', {
    recordId: recordId,
    rowCount: appendInfo.rowCount
  });
  APP_safeToast_(result.message, 'Request');
  return result;
}

function REQ_pullStoredRequest() {
  const namePrompt = REQ_promptForRequestedByName_();
  if (!namePrompt.ok) {
    if (namePrompt.code !== 'USER_CANCELLED') {
      APP_safeToast_(namePrompt.message, 'Request');
    }
    return namePrompt;
  }

  const requestedBy = namePrompt.data.requestedBy;
  const groupedRecords = REQ_getStoredRecordsByRequestedBy_(requestedBy);
  if (groupedRecords.length === 0) {
    const result = APP_fail_('NO_STORED_RECORDS', 'No stored request was found for that name.');
    APP_safeToast_(result.message, 'Request');
    return result;
  }

  const selectionResult = REQ_promptForStoredRecordSelection_(groupedRecords);
  if (!selectionResult.ok) {
    if (selectionResult.code !== 'USER_CANCELLED') {
      APP_safeToast_(selectionResult.message, 'Request');
    }
    return selectionResult;
  }

  const reqSheet = APP_getRequestSheet_();
  let existingSessionId = '';
  let existingSessionRecord = null;
  try {
    existingSessionId = REQ_getCurrentSessionId();
    existingSessionRecord = REQ_getSessionRecord_(existingSessionId);
  } catch (error) {
    existingSessionId = '';
    existingSessionRecord = null;
  }

  const replaceCheck = REQ_confirmReplacingCurrentStaging_(existingSessionRecord ? existingSessionRecord.items.length : 0);
  if (!replaceCheck.ok) {
    return replaceCheck;
  }

  const selectedRecord = selectionResult.data.record;
  const previousRequestedBy = APP_safeString_(reqSheet.getRange(REQ_CONFIG.REQUEST_FORM_CELLS.REQUESTED_BY).getValue()).trim();
  reqSheet.getRange(REQ_CONFIG.REQUEST_FORM_CELLS.REQUESTED_BY).setValue(selectedRecord.requestedBy);
  const newContext = REQ_getRequestContext_(reqSheet);
  const newValidation = REQ_validateContext_(newContext);
  if (!newValidation.ok) {
    reqSheet.getRange(REQ_CONFIG.REQUEST_FORM_CELLS.REQUESTED_BY).setValue(previousRequestedBy);
    APP_safeToast_(newValidation.message, 'Request');
    return newValidation;
  }

  const previousTargetSessionId = REQ_findExistingSessionId_(newContext.contextKey);
  const previousTargetSessionRecord = REQ_getSessionRecord_(previousTargetSessionId);
  const restoredSession = {
    sessionId: REQ_generateSessionId_(newContext),
    contextKey: newContext.contextKey,
    requestedBy: newContext.requestedBy,
    sourceTempRecordId: selectedRecord.recordId,
    actor: newContext.actor.display,
    createdAt: APP_now_(),
    updatedAt: APP_now_(),
    items: selectedRecord.items.map(function (item) {
      return REQ_normalizeSessionItem_({
        itemId: REQ_generateItemId_(),
        partKey: item.partKey,
        name: item.name,
        code: item.code,
        quantity: item.quantity,
        price: item.price,
        remarks: item.remarks,
        sourceRow: item.sourceRow,
        action: item.action
      });
    })
  };

  const pullResult = APP_withScriptLock_(30000, function () {
    if (existingSessionRecord) {
      REQ_removeSessionRecord_(existingSessionRecord.sessionId, existingSessionRecord.contextKey);
    }
    if (previousTargetSessionRecord && previousTargetSessionRecord.sessionId !== restoredSession.sessionId) {
      REQ_removeSessionRecord_(previousTargetSessionRecord.sessionId, previousTargetSessionRecord.contextKey);
    }

    REQ_writeSessionRecord_(restoredSession);
    const logRow = LOG_requestFormAction('PULL_TEMP_REQUEST', newContext, {
      itemName: 'Pulled Request',
      code: '',
      quantity: restoredSession.items.length,
      remarks: REQ_buildItemsPreview_(restoredSession.items),
      timestamp: APP_now_()
    });
    if (logRow <= 0) {
      REQ_removeSessionRecord_(restoredSession.sessionId, restoredSession.contextKey);
      if (existingSessionRecord) {
        REQ_writeSessionRecord_(existingSessionRecord);
      }
      if (previousTargetSessionRecord) {
        REQ_writeSessionRecord_(previousTargetSessionRecord);
      }
      return APP_fail_('LOG_WRITE_FAILED', 'Could not log the pulled request.');
    }
    return APP_result_(true, 'Stored request pulled into staging.', { sessionId: restoredSession.sessionId });
  });

  if (!pullResult.ok) {
    reqSheet.getRange(REQ_CONFIG.REQUEST_FORM_CELLS.REQUESTED_BY).setValue(previousRequestedBy);
    APP_safeToast_(pullResult.message, 'Request');
    return pullResult;
  }

  REQ_refreshItemsTaken(restoredSession.sessionId);
  APP_safeToast_(pullResult.message, 'Request');
  return pullResult;
}

function REQ_finalizeRequestCore_() {
  const sheet = APP_getRequestSheet_();
  const context = REQ_getRequestContext_(sheet);
  const validation = REQ_validateContext_(context);
  if (!validation.ok) {
    return validation;
  }

  let sessionId;
  try {
    sessionId = REQ_getCurrentSessionId();
  } catch (error) {
    return APP_fail_('NO_SESSION', 'No active staged request found.');
  }

  const sessionRecord = REQ_getSessionRecord_(sessionId);
  const items = sessionRecord ? sessionRecord.items : [];
  if (items.length === 0) {
    return APP_fail_('EMPTY_SESSION', 'No staged items found.');
  }

  const inventorySheet = APP_getInventorySheet_();
  const rollbackRows = [];
  const rollbackLogRows = [];
  const aggregates = {};
  const finalizedAt = APP_now_();

  items.forEach(function (item) {
    if (!aggregates[item.partKey]) {
      aggregates[item.partKey] = {
        partKey: item.partKey,
        itemName: item.name,
        code: item.code,
        remarks: item.remarks,
        netDelta: 0
      };
    }
    aggregates[item.partKey].netDelta += APP_getActionDelta_(item.action, item.quantity);
  });

  const aggregateList = Object.keys(aggregates).map(function (key) {
    return aggregates[key];
  }).filter(function (entry) {
    return entry.netDelta !== 0;
  });

  try {
    APP_withScriptLock_(30000, function () {
      const updates = aggregateList.map(function (entry) {
        const record = APP_findInventoryRecordByKey_(inventorySheet, entry.partKey);
        if (!record) {
          throw APP_createError_('ITEM_NOT_FOUND', 'A staged item no longer exists in inventory: ' + entry.itemName);
        }
        const beforeQty = APP_toNumber_(record.qty, 0);
        const afterQty = beforeQty + entry.netDelta;
        if (!REQ_CONFIG.FLAGS.allowNegativeStock && afterQty < 0) {
          throw APP_createError_('INSUFFICIENT_STOCK', 'Finalizing would reduce stock below zero for ' + record.name + '.');
        }
        return {
          rowNumber: record.rowNumber,
          beforeQty: beforeQty,
          beforeDate: record.dateInv,
          price: record.price,
          afterQty: afterQty,
          itemName: record.name,
          code: record.code,
          remarks: record.remarks,
          action: entry.netDelta > 0 ? APP_CONFIG.actions.ADD : APP_CONFIG.actions.SUBTRACT,
          quantity: Math.abs(entry.netDelta)
        };
      });

      APP_writeInventoryStockRows_(inventorySheet, updates.map(function (update) {
        return {
          rowNumber: update.rowNumber,
          qty: update.afterQty,
          price: update.price,
          dateInv: finalizedAt
        };
      }));

      updates.forEach(function (update) {
        rollbackRows.push({
          rowNumber: update.rowNumber,
          qty: update.beforeQty,
          price: update.price,
          dateInv: update.beforeDate
        });

        const changeLogRow = LOG_appendChangeLog({
          requestedBy: context.requestedBy,
          action: update.action,
          itemName: update.itemName,
          code: update.code,
          quantity: update.quantity,
          remarks: update.remarks,
          source: 'REQUEST_FORM',
          timestamp: finalizedAt
        }, { skipSync: true });
        if (changeLogRow <= 0) {
          throw APP_createError_('LOG_WRITE_FAILED', 'Could not write the change log entry.');
        }
        rollbackLogRows.push({ sheetName: LOG_CONFIG.SHEET_NAMES.CHANGE_LOGS, rowNumber: changeLogRow });
      });

      APP_invalidateInventorySearchCaches_();
      const requestLogRow = LOG_requestFormAction('FINALIZE_REQUEST', context, {
        itemName: 'Finalized Request',
        code: '',
        quantity: items.length,
        remarks: REQ_buildItemsTakenSummary_(items),
        timestamp: finalizedAt
      }, { skipSync: true });
      if (requestLogRow <= 0) {
        throw APP_createError_('LOG_WRITE_FAILED', 'Could not write the request log entry.');
      }
      rollbackLogRows.push({ sheetName: LOG_CONFIG.SHEET_NAMES.REQUEST_FORM_LOGS, rowNumber: requestLogRow });

      const finalizedRows = LOG_appendFinalizedRequestLogs(context.requestedBy, updates.map(function (update) {
        return {
          action: update.action,
          itemName: update.itemName,
          code: update.code,
          quantity: update.quantity,
          remarks: update.remarks,
          timestamp: finalizedAt
        };
      }), { skipSync: true });
      if (finalizedRows.length !== updates.length) {
        throw APP_createError_('LOG_WRITE_FAILED', 'Could not write finalized request logs.');
      }
      finalizedRows.forEach(function (rowNumber) {
        rollbackLogRows.push({ sheetName: LOG_CONFIG.SHEET_NAMES.FINALIZED_REQUEST_LOGS, rowNumber: rowNumber });
      });
    });
  } catch (error) {
    APP_writeInventoryStockRows_(inventorySheet, rollbackRows);
    APP_invalidateInventorySearchCaches_();
    const grouped = {};
    rollbackLogRows.forEach(function (entry) {
      if (!grouped[entry.sheetName]) {
        grouped[entry.sheetName] = [];
      }
      grouped[entry.sheetName].push(entry.rowNumber);
    });
    Object.keys(grouped).forEach(function (sheetName) {
      LOG_deleteRows_(sheetName, grouped[sheetName]);
    });
    return APP_fail_(error.code || 'FINALIZE_FAILED', error.message, error.details);
  }

  if (!APP_isBlank_(sessionRecord && sessionRecord.sourceTempRecordId)) {
    REQ_deleteTempStoreRecord_(sessionRecord.sourceTempRecordId);
  }
  REQ_clearSession(sessionId);
  REQ_refreshItemsTaken('');
  REQ_clearEntryFields_(sheet);
  REQ_clearRequestedBy_(sheet);
  try {
    if (typeof syncRowToTarget === 'function') {
      syncRowToTarget(0);
    }
  } catch (syncError) {
    console.log('View-only sync warning: ' + syncError.message);
  }
  return APP_result_(true, 'Staged request finalized successfully.', { sessionId: sessionId, itemCount: items.length });
}

function REQ_finalizeRequest() {
  const result = REQ_finalizeRequestCore_();
  APP_safeToast_(result.message, result.ok ? 'Request' : 'Request Error');
  return result;
}

function REQ_setupTriggers() {
  APP_ensureTrigger_(APP_CONFIG.triggerHandlers.requestEdit, 'onEdit');
  return APP_result_(true, 'Request form trigger installed.');
}

function REQ_handleEdit(e) {
  if (!e || !e.range) {
    return;
  }
  APP_reconcilePendingActionDialog_();
  const sheet = e.range.getSheet();
  if (sheet.getName() !== REQ_CONFIG.SHEET_NAME) {
    return;
  }

  const cellAddress = e.range.getA1Notation();
  try {
    if (cellAddress === REQ_CONFIG.REQUEST_FORM_CELLS.REQUESTED_BY) {
      REQ_refreshItemsTaken('');
      return;
    }

    const selectionResult = REQ_handleSearchSelectionEdit_(e);
    if (selectionResult) {
      if (selectionResult.code !== 'USER_CANCELLED' && !(selectionResult.data && selectionResult.data.deferred)) {
        APP_safeToast_(selectionResult.message, selectionResult.ok ? 'Request' : 'Request Error');
      }
      return;
    }

    const removalResult = REQ_detectRemoval(e);
    if (removalResult) {
      if (removalResult.code !== 'USER_CANCELLED') {
        APP_safeToast_(removalResult.message, removalResult.ok ? 'Request' : 'Request Error');
      }
      return;
    }

    if (cellAddress === REQ_CONFIG.REQUEST_FORM_CELLS.SEARCH_NAME) {
      if (APP_isBlank_(e.range.getValue())) {
        REQ_clearSearchResults(sheet);
        return;
      }
      sheet.getRange(REQ_CONFIG.REQUEST_FORM_CELLS.SEARCH_CODE).setValue('');
      sheet.getRange(REQ_CONFIG.REQUEST_FORM_CELLS.SEARCH_ROW).setValue('');
      const result = REQ_searchByName(e.range.getValue());
      if ((result.data.matches || []).length === 0) {
        APP_safeToast_('No matching items found.', 'Request');
      }
      return;
    }

    if (cellAddress === REQ_CONFIG.REQUEST_FORM_CELLS.SEARCH_CODE) {
      if (APP_isBlank_(e.range.getValue())) {
        REQ_clearSearchResults(sheet);
        return;
      }
      sheet.getRange(REQ_CONFIG.REQUEST_FORM_CELLS.SEARCH_NAME).setValue('');
      sheet.getRange(REQ_CONFIG.REQUEST_FORM_CELLS.SEARCH_ROW).setValue('');
      const result = REQ_handleImmediateResolution_(sheet, REQ_searchByCode(e.range.getValue()), {
        multipleMessage: 'Multiple items matched that code. Use ROW or X from the results list to choose one item.'
      });
      if (result.code !== 'USER_CANCELLED' && !(result.data && result.data.deferred)) {
        APP_safeToast_(result.message, result.ok ? 'Request' : 'Request Error');
      }
      return;
    }

    if (cellAddress === REQ_CONFIG.REQUEST_FORM_CELLS.SEARCH_ROW) {
      if (APP_isBlank_(e.range.getValue())) {
        REQ_clearSearchResults(sheet);
        return;
      }
      sheet.getRange(REQ_CONFIG.REQUEST_FORM_CELLS.SEARCH_CODE).setValue('');
      sheet.getRange(REQ_CONFIG.REQUEST_FORM_CELLS.SEARCH_NAME).setValue('');
      const result = REQ_handleImmediateResolution_(sheet, REQ_searchByRow(e.range.getValue()), {
        multipleMessage: 'Use a single valid inventory row.'
      });
      if (result.code !== 'USER_CANCELLED' && !(result.data && result.data.deferred)) {
        APP_safeToast_(result.message, result.ok ? 'Request' : 'Request Error');
      }
    }
  } catch (error) {
    APP_safeToast_(error.message, 'Request Error');
    console.error('REQ_handleEdit error: ' + error.message);
  }
}

function REQ_initialize() {
  REQ_initializeFormLayout();
  REQ_initializeSearchResultsTable();
  REQ_initializeItemsTakenTable();
  REQ_ensureTempStoreSheet();
  REQ_refreshItemsTaken('');
  REQ_setupTriggers();
  return APP_result_(true, 'Request form initialized.');
}
