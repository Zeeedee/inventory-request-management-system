function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu(APP_CONFIG.system.menu.root)
    .addItem(APP_CONFIG.system.menu.initialize, 'APP_initializeSystem')
    .addItem(APP_CONFIG.system.menu.runVerification, 'runHealthCheck')
    .addSeparator()
    .addItem(APP_CONFIG.system.menu.createViewOnly, 'SYNC_createViewOnlySpreadsheet')
    .addItem(APP_CONFIG.system.menu.syncViewOnly, 'manualSyncToTarget')
    .addSeparator()
    .addItem(APP_CONFIG.system.menu.finalizeRequest, 'REQ_finalizeRequest')
    .addItem(APP_CONFIG.system.menu.storeRequest, 'REQ_storeTemporarily')
    .addItem(APP_CONFIG.system.menu.pullRequest, 'REQ_pullStoredRequest')
    .addItem(APP_CONFIG.system.menu.deleteStoredRequest, 'REQ_deleteStoredRequest')
    .addItem(APP_CONFIG.system.menu.clearRequest, 'REQ_clearItemsTaken')
    .addToUi();
}

function createRequestFormMenus() {
  onOpen();
}

function onOpenMainScript() {
  onOpen();
}

function setupEncoderTrigger() {
  APP_ensureTrigger_(APP_CONFIG.triggerHandlers.encoderEdit, 'onEdit');
  return APP_result_(true, 'Encoder trigger installed.');
}

function ENC_getResultBodyRowCount_(sheet) {
  const targetSheet = sheet || APP_getEncoderSheet_();
  return Math.max(APP_CONFIG.encoder.resultLimit, targetSheet.getLastRow() - APP_CONFIG.encoder.resultStartRow + 1, 1);
}

function ENC_applyResultBodyFormatting_(sheet, rowCount) {
  const targetSheet = sheet || APP_getEncoderSheet_();
  const safeRowCount = Math.max(rowCount || APP_CONFIG.encoder.resultLimit, 1);
  const range = targetSheet.getRange(APP_CONFIG.encoder.resultStartRow, 1, safeRowCount, APP_CONFIG.encoder.visibleWidth);
  const backgrounds = [];
  for (let rowIndex = 0; rowIndex < safeRowCount; rowIndex += 1) {
    const color = rowIndex % 2 === 0 ? APP_CONFIG.encoder.colors.altRow1 : APP_CONFIG.encoder.colors.altRow2;
    backgrounds.push([color, color, color, color, color, color, color]);
  }
  range
    .setBackgrounds(backgrounds)
    .setWrap(true)
    .setBorder(true, true, true, true, true, true, APP_CONFIG.encoder.colors.border, SpreadsheetApp.BorderStyle.SOLID)
    .setVerticalAlignment('middle');
  targetSheet.getRange(APP_CONFIG.encoder.resultStartRow, APP_CONFIG.encoder.cols.QTY, safeRowCount, 1).setHorizontalAlignment('center');
  targetSheet.getRange(APP_CONFIG.encoder.resultStartRow, APP_CONFIG.encoder.cols.PRICE, safeRowCount, 1).setNumberFormat('#,##0.00');
  targetSheet.getRange(APP_CONFIG.encoder.resultStartRow, APP_CONFIG.encoder.cols.ROW, safeRowCount, 1).setHorizontalAlignment('center');
}

function ENC_clearResults(encoderSheet) {
  const sheet = encoderSheet || APP_getEncoderSheet_();
  const rowCount = ENC_getResultBodyRowCount_(sheet);
  sheet.getRange(APP_CONFIG.encoder.resultStartRow, 1, rowCount, APP_CONFIG.encoder.visibleWidth).clearContent();
  sheet.getRange(APP_CONFIG.encoder.resultStartRow, APP_CONFIG.encoder.cols.ITEM_KEY, rowCount, 1).clearContent();
}

function ENC_clearInputFields_(encoderSheet) {
  const sheet = encoderSheet || APP_getEncoderSheet_();
  sheet.getRange(APP_CONFIG.encoder.cells.SEARCH_CODE).setValue('');
  sheet.getRange(APP_CONFIG.encoder.cells.SEARCH_NAME).setValue('');
  sheet.getRange(APP_CONFIG.encoder.cells.SEARCH_ROW).setValue('');
}

function ENC_clearEncoderSheet(encoderSheet) {
  ENC_clearInputFields_(encoderSheet || APP_getEncoderSheet_());
  ENC_clearResults(encoderSheet || APP_getEncoderSheet_());
}

function ENC_buildDisplayRow_(record) {
  return [
    record.name || '',
    APP_isNoCode(record.code) ? 'NO CODE' : record.code,
    APP_toNumber_(record.qty, 0),
    record.price || 0,
    record.dateInv || '',
    record.remarks || '',
    record.rowNumber || ''
  ];
}

function ENC_displayResults_(encoderSheet, records) {
  const sheet = encoderSheet || APP_getEncoderSheet_();
  ENC_clearResults(sheet);
  const safeRecords = (records || []).slice(0, APP_CONFIG.encoder.resultLimit);
  if (safeRecords.length === 0) {
    return;
  }

  const visibleRows = safeRecords.map(ENC_buildDisplayRow_);
  const keyRows = safeRecords.map(function (record) {
    return [record.key || ''];
  });
  sheet.getRange(APP_CONFIG.encoder.resultStartRow, 1, visibleRows.length, APP_CONFIG.encoder.visibleWidth).setValues(visibleRows);
  sheet.getRange(APP_CONFIG.encoder.resultStartRow, APP_CONFIG.encoder.cols.ITEM_KEY, keyRows.length, 1).setValues(keyRows);
  sheet.hideColumns(APP_CONFIG.encoder.cols.ITEM_KEY);
}

function ENC_searchByCodeCore_(searchCode) {
  const normalizedSearch = APP_normalizeText_(searchCode);
  if (normalizedSearch === '' || APP_isNoCode(normalizedSearch)) {
    return APP_result_(true, 'Search cleared.', { records: [] });
  }
  const records = APP_findInventoryByCode_(normalizedSearch);
  LOG_encoderAction('SEARCH_BY_CODE', normalizedSearch, records.length);
  if (records.length === 0) {
    return APP_fail_('CODE_NOT_FOUND', APP_CONFIG.encoder.messages.noMatchingItem);
  }
  return APP_result_(true, 'Code search completed.', { records: records });
}

function ENC_searchByNameCore_(searchTerm) {
  const normalizedTerm = APP_safeString_(searchTerm).trim();
  if (normalizedTerm === '') {
    return APP_result_(true, 'Search cleared.', { records: [] });
  }
  const records = APP_searchInventoryByName_(normalizedTerm, APP_CONFIG.encoder.resultLimit);
  LOG_encoderAction('SEARCH_BY_NAME', normalizedTerm, records.length);
  return APP_result_(true, 'Name search completed.', { records: records });
}

function ENC_searchByRowCore_(rowNumber) {
  const safeRow = APP_toNumber_(rowNumber, 0);
  if (safeRow < APP_CONFIG.inventory.startRow) {
    return APP_result_(true, 'Search cleared.', { records: [] });
  }
  const record = APP_findInventoryByRow_(safeRow);
  LOG_encoderAction('SEARCH_BY_ROW', safeRow, record ? 1 : 0);
  if (!record) {
    return APP_fail_('ROW_NOT_FOUND', 'No inventory item exists on that row.');
  }
  return APP_result_(true, 'Row search completed.', { records: [record] });
}

function ENC_getNoResultMessage_() {
  return APP_CONFIG.encoder.messages.noMatchingItem;
}

function ENC_shouldShowNoResultPrompt_(searchValue, result) {
  if (APP_isBlank_(searchValue)) {
    return false;
  }
  if (!result) {
    return false;
  }
  if (!result.ok) {
    return result.code === 'CODE_NOT_FOUND' || result.code === 'ITEM_NOT_FOUND';
  }
  return ((result.data && result.data.records) || []).length === 0;
}

function ENC_openActionDialog_(record, requestedBy, options) {
  const settings = options || {};
  return APP_showActionDialog_({
    title: 'Update Stock',
    requestedBy: requestedBy,
    record: record,
    quantityLabel: 'Quantity',
    submitPayload: {
      target: 'ENCODER',
      itemKey: record.key,
      sourceRow: record.rowNumber,
      requestedBy: requestedBy
    },
    cancelPayload: {
      target: 'ENCODER',
      clearOnCancel: settings.clearOnCancel === true
    }
  });
}

function ENC_handleImmediateResolution_(sheet, result, requestedBy, options) {
  const settings = options || {};
  if (!result.ok) {
    ENC_clearResults(sheet);
    APP_safeToast_(result.message, 'Encoder');
    return result;
  }

  const records = result.data.records || [];
  ENC_displayResults_(sheet, records);
  if (records.length !== 1) {
    const message = records.length > 1
      ? 'Multiple items matched. Use the visible ROW value to choose one item.'
      : ENC_getNoResultMessage_();
    APP_safeToast_(message, 'Encoder');
    return APP_fail_(records.length > 1 ? 'MULTIPLE_RESULTS' : 'ITEM_NOT_FOUND', message);
  }

  return ENC_openActionDialog_(records[0], requestedBy, settings);
}

function handleEncoderEdit(e) {
  if (!e || !e.range) {
    return;
  }
  APP_reconcilePendingActionDialog_();
  const sheet = e.range.getSheet();
  if (sheet.getName() !== APP_CONFIG.sheets.encoder) {
    return;
  }

  const requestedBy = APP_safeString_(sheet.getRange(APP_CONFIG.encoder.cells.REQUESTED_BY).getValue()).trim() || APP_getActor_().display;
  const cellAddress = e.range.getA1Notation();
  try {
    if (cellAddress === APP_CONFIG.encoder.cells.SEARCH_NAME) {
      if (APP_isBlank_(e.range.getValue())) {
        ENC_clearResults(sheet);
        return;
      }
      sheet.getRange(APP_CONFIG.encoder.cells.SEARCH_CODE).setValue('');
      sheet.getRange(APP_CONFIG.encoder.cells.SEARCH_ROW).setValue('');
      const result = ENC_searchByNameCore_(e.range.getValue());
      ENC_displayResults_(sheet, result.data.records);
      if (ENC_shouldShowNoResultPrompt_(e.range.getValue(), result)) {
        APP_safeToast_(ENC_getNoResultMessage_(), 'Encoder');
      }
      return;
    }

    if (cellAddress === APP_CONFIG.encoder.cells.SEARCH_CODE) {
      if (APP_isBlank_(e.range.getValue())) {
        ENC_clearResults(sheet);
        return;
      }
      sheet.getRange(APP_CONFIG.encoder.cells.SEARCH_NAME).setValue('');
      sheet.getRange(APP_CONFIG.encoder.cells.SEARCH_ROW).setValue('');
      ENC_handleImmediateResolution_(sheet, ENC_searchByCodeCore_(e.range.getValue()), requestedBy, { clearOnCancel: true });
      return;
    }

    if (cellAddress === APP_CONFIG.encoder.cells.SEARCH_ROW) {
      if (APP_isBlank_(e.range.getValue())) {
        ENC_clearResults(sheet);
        return;
      }
      sheet.getRange(APP_CONFIG.encoder.cells.SEARCH_CODE).setValue('');
      sheet.getRange(APP_CONFIG.encoder.cells.SEARCH_NAME).setValue('');
      ENC_handleImmediateResolution_(sheet, ENC_searchByRowCore_(e.range.getValue()), requestedBy, { clearOnCancel: true });
    }
  } catch (error) {
    APP_safeToast_(error.message, 'Encoder Error');
    console.error('handleEncoderEdit error: ' + error.message);
  }
}
