const SETUP_CONFIG = {
  statuses: {
    PASS: 'PASS',
    FIXED: 'FIXED',
    FAIL: 'FAIL'
  }
};

function SETUP_newReport_() {
  return {
    ok: true,
    createdSheets: [],
    rebuiltSheets: [],
    installedTriggers: [],
    checks: [],
    summary: ''
  };
}

function SETUP_addCheck_(report, name, status, details) {
  report.checks.push({ name: name, status: status, details: details || '' });
  if (status === SETUP_CONFIG.statuses.FAIL) {
    report.ok = false;
  }
}

function SETUP_styleSectionHeader_(range, background, fontColor) {
  range
    .setFontWeight('bold')
    .setBackground(background || APP_CONFIG.system.colors.sectionBg)
    .setFontColor(fontColor || APP_CONFIG.system.colors.sectionText)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setBorder(true, true, true, true, true, true, APP_CONFIG.system.colors.border, SpreadsheetApp.BorderStyle.SOLID);
}

function SETUP_getOrCreateSheet_(sheetName, minRows, minCols, report) {
  const ss = APP_getSpreadsheet_();
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    report.createdSheets.push(sheetName);
  }
  if (sheet.getMaxRows() < minRows) {
    sheet.insertRowsAfter(sheet.getMaxRows(), minRows - sheet.getMaxRows());
  }
  if (sheet.getMaxColumns() < minCols) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), minCols - sheet.getMaxColumns());
  }
  return sheet;
}

function SETUP_initializeInventorySheet_(report) {
  const inventoryTemplate = APP_getTemplateConfig_('inventory');
  const inventoryHeaders = APP_getInventoryTemplateHeaders_();
  const extraColumnWidths = inventoryTemplate.additionalColumnWidths || [];
  const totalInventoryColumns = Math.max(APP_CONFIG.inventory.lastDataColumn, inventoryHeaders.length);
  const sheet = SETUP_getOrCreateSheet_(APP_CONFIG.sheets.inventory, APP_CONFIG.inventory.startRow + APP_CONFIG.inventory.templateRowCount, totalInventoryColumns, report);
  const existingLastRow = sheet.getLastRow();
  const existingRowCount = existingLastRow >= APP_CONFIG.inventory.startRow ? existingLastRow - APP_CONFIG.inventory.startRow + 1 : 0;
  const existingValues = existingRowCount > 0
    ? sheet.getRange(APP_CONFIG.inventory.startRow, 1, existingRowCount, totalInventoryColumns).getValues()
    : [];
  const existingNotes = existingRowCount > 0
    ? sheet.getRange(APP_CONFIG.inventory.startRow, APP_CONFIG.inventory.cols.NAME, existingRowCount, 1).getNotes()
    : [];
  try {
    sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).breakApart();
  } catch (error1) {
    console.log('Inventory unmerge skipped: ' + error1.message);
  }
  sheet.clear();
  sheet.clearFormats();

  sheet.getRange(APP_CONFIG.inventory.titleRow, 1, 1, totalInventoryColumns).merge();
  sheet.getRange(APP_CONFIG.inventory.titleRow, 1)
    .setValue(inventoryTemplate.title || APP_CONFIG.sheets.inventory)
    .setFontSize(16)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setBackground(APP_CONFIG.system.colors.titleBg)
    .setFontColor(APP_CONFIG.system.colors.titleText)
    .setBorder(true, true, true, true, true, true, APP_CONFIG.system.colors.border, SpreadsheetApp.BorderStyle.SOLID);

  sheet.getRange(APP_CONFIG.inventory.helpRow, 1, 1, totalInventoryColumns).merge();
  sheet.getRange(APP_CONFIG.inventory.helpRow, 1)
    .setValue(inventoryTemplate.helpText || 'Enter inventory items below. REMARKS is optional.')
    .setWrap(true)
    .setHorizontalAlignment('left')
    .setVerticalAlignment('middle')
    .setBackground(APP_CONFIG.system.colors.noteBg)
    .setBorder(true, true, true, true, true, true, APP_CONFIG.system.colors.border, SpreadsheetApp.BorderStyle.SOLID);

  sheet.getRange(APP_CONFIG.inventory.headerRow, 1, 1, inventoryHeaders.length)
    .setValues([inventoryHeaders])
    .setFontWeight('bold')
    .setBackground(inventoryTemplate.styles.headerBg)
    .setFontColor(inventoryTemplate.styles.headerText)
    .setHorizontalAlignment('center')
    .setWrap(true)
    .setBorder(true, true, true, true, true, true, APP_CONFIG.system.colors.border, SpreadsheetApp.BorderStyle.SOLID);

  sheet.setFrozenRows(APP_CONFIG.inventory.headerRow);
  sheet.setRowHeight(APP_CONFIG.inventory.titleRow, 30);
  sheet.setRowHeight(APP_CONFIG.inventory.helpRow, 24);
  sheet.setColumnWidth(APP_CONFIG.inventory.cols.NAME, inventoryTemplate.columnWidths.NAME);
  sheet.setColumnWidth(APP_CONFIG.inventory.cols.CODE, inventoryTemplate.columnWidths.CODE);
  sheet.setColumnWidth(APP_CONFIG.inventory.cols.QTY, inventoryTemplate.columnWidths.QTY);
  sheet.setColumnWidth(APP_CONFIG.inventory.cols.PRICE, inventoryTemplate.columnWidths.PRICE);
  sheet.setColumnWidth(APP_CONFIG.inventory.cols.DATE_INV, inventoryTemplate.columnWidths.DATE_INV);
  sheet.setColumnWidth(APP_CONFIG.inventory.cols.REMARKS, inventoryTemplate.columnWidths.REMARKS);

  const bodyRange = sheet.getRange(APP_CONFIG.inventory.startRow, 1, APP_CONFIG.inventory.templateRowCount, totalInventoryColumns);
  const bodyBackgrounds = [];
  for (let rowIndex = 0; rowIndex < APP_CONFIG.inventory.templateRowCount; rowIndex += 1) {
    const color = rowIndex % 2 === 0 ? APP_CONFIG.system.colors.rowAlt1 : APP_CONFIG.system.colors.rowAlt2;
    const rowColors = [];
    for (let columnIndex = 0; columnIndex < totalInventoryColumns; columnIndex += 1) {
      rowColors.push(color);
    }
    bodyBackgrounds.push(rowColors);
  }
  bodyRange
    .setBackgrounds(bodyBackgrounds)
    .setWrap(true)
    .setVerticalAlignment('middle')
    .setBorder(true, true, true, true, true, true, APP_CONFIG.system.colors.border, SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange(APP_CONFIG.inventory.startRow, APP_CONFIG.inventory.cols.QTY, APP_CONFIG.inventory.templateRowCount, 1).setHorizontalAlignment('center');
  sheet.getRange(APP_CONFIG.inventory.startRow, APP_CONFIG.inventory.cols.PRICE, APP_CONFIG.inventory.templateRowCount, 1).setNumberFormat('#,##0.00');
  sheet.getRange(APP_CONFIG.inventory.startRow, APP_CONFIG.inventory.cols.DATE_INV, APP_CONFIG.inventory.templateRowCount, 1).setNumberFormat(APP_CONFIG.inventory.dateTimeNumberFormat);

  if (existingValues.length > 0) {
    sheet.getRange(APP_CONFIG.inventory.startRow, 1, existingValues.length, totalInventoryColumns).setValues(existingValues);
    sheet.getRange(APP_CONFIG.inventory.startRow, APP_CONFIG.inventory.cols.NAME, existingNotes.length, 1).setNotes(existingNotes);
    sheet.getRange(APP_CONFIG.inventory.startRow, 1, Math.max(existingValues.length, APP_CONFIG.inventory.templateRowCount), totalInventoryColumns)
      .setBorder(true, true, true, true, true, true, APP_CONFIG.system.colors.border, SpreadsheetApp.BorderStyle.SOLID);
    const existingBackgrounds = [];
    const totalRows = Math.max(existingValues.length, APP_CONFIG.inventory.templateRowCount);
    for (let rowIndex = 0; rowIndex < totalRows; rowIndex += 1) {
      const color = rowIndex % 2 === 0 ? APP_CONFIG.system.colors.rowAlt1 : APP_CONFIG.system.colors.rowAlt2;
      const rowColors = [];
      for (let columnIndex = 0; columnIndex < totalInventoryColumns; columnIndex += 1) {
        rowColors.push(color);
      }
      existingBackgrounds.push(rowColors);
    }
    sheet.getRange(APP_CONFIG.inventory.startRow, 1, Math.max(existingValues.length, APP_CONFIG.inventory.templateRowCount), totalInventoryColumns).setBackgrounds(existingBackgrounds);
    sheet.getRange(APP_CONFIG.inventory.startRow, APP_CONFIG.inventory.cols.REMARKS, Math.max(existingValues.length, APP_CONFIG.inventory.templateRowCount), 1).setWrap(true);
    sheet.getRange(APP_CONFIG.inventory.startRow, APP_CONFIG.inventory.cols.QTY, Math.max(existingValues.length, APP_CONFIG.inventory.templateRowCount), 1).setHorizontalAlignment('center');
    sheet.getRange(APP_CONFIG.inventory.startRow, APP_CONFIG.inventory.cols.PRICE, Math.max(existingValues.length, APP_CONFIG.inventory.templateRowCount), 1).setNumberFormat('#,##0.00');
    sheet.getRange(APP_CONFIG.inventory.startRow, APP_CONFIG.inventory.cols.DATE_INV, Math.max(existingValues.length, APP_CONFIG.inventory.templateRowCount), 1).setNumberFormat(APP_CONFIG.inventory.dateTimeNumberFormat);
  }

  APP_invalidateInventorySearchCaches_();
  extraColumnWidths.forEach(function (columnWidth, index) {
    sheet.setColumnWidth(APP_CONFIG.inventory.lastDataColumn + index + 1, columnWidth);
  });
  report.rebuiltSheets.push(APP_CONFIG.sheets.inventory);
  SETUP_addCheck_(report, 'Inventory sheet', SETUP_CONFIG.statuses.FIXED, 'Inventory sheet exists, is bordered, and uses the new six-column visible structure.');
}

function SETUP_initializeEncoderSheet_(report) {
  const sheet = SETUP_getOrCreateSheet_(APP_CONFIG.sheets.encoder, APP_CONFIG.encoder.resultStartRow + APP_CONFIG.encoder.resultLimit, APP_CONFIG.encoder.resultWidth, report);
  const requestedBy = APP_safeString_(sheet.getRange(APP_CONFIG.encoder.cells.REQUESTED_BY).getValue()).trim();
  const layout = APP_CONFIG.encoder.layout;
  try {
    sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).breakApart();
  } catch (error2) {
    console.log('Encoder unmerge skipped: ' + error2.message);
  }
  sheet.clear();
  sheet.clearFormats();

  sheet.getRange(layout.titleRange).merge()
    .setValue(layout.title)
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

  sheet.getRange(layout.searchHeaderRange).merge().setValue(layout.searchHeaderText);
  SETUP_styleSectionHeader_(sheet.getRange(layout.searchHeaderRange));

  sheet.getRange('A1').setValue(APP_CONFIG.labels.requestedBy);
  sheet.getRange('A3').setValue('Code');
  sheet.getRange('A4').setValue('Name');
  sheet.getRange('A5').setValue('Row');
  sheet.getRange('A1:A5').setFontWeight('bold').setBackground(APP_CONFIG.encoder.colors.titleBg).setHorizontalAlignment('center');
  sheet.getRange('A1:B5').setBorder(true, true, true, true, true, true, APP_CONFIG.system.colors.border, SpreadsheetApp.BorderStyle.SOLID).setVerticalAlignment('middle');
  sheet.getRange('B1:B5').setBackground('#ffffff');

  sheet.getRange(layout.resultsTitleRange).merge().setValue(layout.resultsTitleText);
  SETUP_styleSectionHeader_(sheet.getRange(layout.resultsTitleRange));

  sheet.getRange(layout.helperRange).merge()
    .setValue(layout.helperText)
    .setWrap(true)
    .setBackground(APP_CONFIG.encoder.colors.noteBg)
    .setHorizontalAlignment('left')
    .setBorder(true, true, true, true, true, true, APP_CONFIG.system.colors.border, SpreadsheetApp.BorderStyle.SOLID);

  sheet.getRange(APP_CONFIG.encoder.resultHeaderRow, 1, 1, APP_CONFIG.encoder.visibleWidth)
    .setValues([APP_CONFIG.encoder.visibleHeaders])
    .setFontWeight('bold')
    .setBackground(APP_CONFIG.encoder.colors.headerBg)
    .setFontColor(APP_CONFIG.encoder.colors.headerText)
    .setHorizontalAlignment('center')
    .setWrap(true)
    .setBorder(true, true, true, true, true, true, APP_CONFIG.system.colors.border, SpreadsheetApp.BorderStyle.SOLID);

  sheet.setFrozenRows(APP_CONFIG.encoder.resultHeaderRow);
  sheet.setRowHeight(1, 28);
  sheet.setRowHeight(2, 24);
  sheet.setRowHeight(6, 22);
  sheet.setRowHeight(7, 24);
  sheet.setColumnWidth(APP_CONFIG.encoder.cols.NAME, APP_CONFIG.encoder.columnWidths.NAME);
  sheet.setColumnWidth(APP_CONFIG.encoder.cols.CODE, APP_CONFIG.encoder.columnWidths.CODE);
  sheet.setColumnWidth(APP_CONFIG.encoder.cols.QTY, APP_CONFIG.encoder.columnWidths.QTY);
  sheet.setColumnWidth(APP_CONFIG.encoder.cols.PRICE, APP_CONFIG.encoder.columnWidths.PRICE);
  sheet.setColumnWidth(APP_CONFIG.encoder.cols.DATE_INV, APP_CONFIG.encoder.columnWidths.DATE_INV);
  sheet.setColumnWidth(APP_CONFIG.encoder.cols.REMARKS, APP_CONFIG.encoder.columnWidths.REMARKS);
  sheet.setColumnWidth(APP_CONFIG.encoder.cols.ROW, APP_CONFIG.encoder.columnWidths.ROW);

  ENC_applyResultBodyFormatting_(sheet, APP_CONFIG.encoder.resultLimit);
  ENC_clearResults(sheet);
  if (!APP_isBlank_(requestedBy)) {
    sheet.getRange(APP_CONFIG.encoder.cells.REQUESTED_BY).setValue(requestedBy);
  }
  sheet.hideColumns(APP_CONFIG.encoder.cols.ITEM_KEY);
  report.rebuiltSheets.push(APP_CONFIG.sheets.encoder);
  SETUP_addCheck_(report, 'Encoder sheet', SETUP_CONFIG.statuses.FIXED, 'Encoder layout rebuilt with bordered result table and automatic code/row flow.');
}

function SETUP_initializeRequestSheet_(report) {
  const sheet = SETUP_getOrCreateSheet_(APP_CONFIG.sheets.requestForm, REQ_CONFIG.SEARCH_RESULTS_TABLE.startRow + REQ_CONFIG.SEARCH_RESULTS_TABLE.maxResults + 5, REQ_CONFIG.ITEMS_TABLE.cols.ITEM_KEY, report);
  const requestedBy = APP_safeString_(sheet.getRange(REQ_CONFIG.REQUEST_FORM_CELLS.REQUESTED_BY).getValue()).trim();
  REQ_initializeFormLayout();
  REQ_initializeSearchResultsTable();
  REQ_initializeItemsTakenTable();
  REQ_ensureTempStoreSheet();
  if (!APP_isBlank_(requestedBy)) {
    sheet.getRange(REQ_CONFIG.REQUEST_FORM_CELLS.REQUESTED_BY).setValue(requestedBy);
  }
  REQ_refreshItemsTaken('');
  report.rebuiltSheets.push(APP_CONFIG.sheets.requestForm);
  SETUP_addCheck_(report, 'Request form sheet', SETUP_CONFIG.statuses.FIXED, 'Request form layout rebuilt with result/staging separator, bordered tables, and temp-store support.');
}

function SETUP_initializePortfolioNotesSheet_(report) {
  const notes = APP_CONFIG.portfolio.notes || [];
  const sheet = SETUP_getOrCreateSheet_(APP_CONFIG.sheets.portfolioNotes, Math.max(notes.length + 8, 12), 4, report);
  try {
    sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).breakApart();
  } catch (error) {
    console.log('Portfolio notes unmerge skipped: ' + error.message);
  }
  sheet.clear();
  sheet.clearFormats();

  sheet.getRange('A1:D1').merge()
    .setValue(APP_CONFIG.portfolio.title)
    .setFontSize(16)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setBackground(APP_CONFIG.system.colors.titleBg)
    .setFontColor(APP_CONFIG.system.colors.titleText)
    .setBorder(true, true, true, true, true, true, APP_CONFIG.system.colors.border, SpreadsheetApp.BorderStyle.SOLID);

  sheet.getRange('A2:D2').merge()
    .setValue(APP_CONFIG.portfolio.subtitle)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setBackground(APP_CONFIG.system.colors.accentBg)
    .setFontColor(APP_CONFIG.system.colors.accentText)
    .setBorder(true, true, true, true, true, true, APP_CONFIG.system.colors.border, SpreadsheetApp.BorderStyle.SOLID);

  sheet.getRange('A4:A4').setValue('Project');
  sheet.getRange('B4:D4').merge().setValue(APP_CONFIG.appName);
  SETUP_styleSectionHeader_(sheet.getRange('A4:A4'));
  sheet.getRange('B4:D4')
    .setBackground('#ffffff')
    .setBorder(true, true, true, true, true, true, APP_CONFIG.system.colors.border, SpreadsheetApp.BorderStyle.SOLID);

  sheet.getRange('A6:D6').merge().setValue('Overview');
  SETUP_styleSectionHeader_(sheet.getRange('A6:D6'));

  const noteRows = notes.map(function (note) {
    return ['- ' + note];
  });
  if (noteRows.length > 0) {
    noteRows.forEach(function (row, index) {
      const range = sheet.getRange(7 + index, 1, 1, 4).merge();
      range
        .setValue(row[0])
        .setWrap(true)
        .setVerticalAlignment('top')
        .setBackground('#ffffff')
        .setBorder(true, true, true, true, true, true, APP_CONFIG.system.colors.border, SpreadsheetApp.BorderStyle.SOLID);
    });
  }

  sheet.setFrozenRows(2);
  sheet.setColumnWidth(1, 180);
  sheet.setColumnWidth(2, 260);
  sheet.setColumnWidth(3, 260);
  sheet.setColumnWidth(4, 260);
  sheet.setRowHeight(1, 30);
  sheet.setRowHeight(2, 24);

  report.rebuiltSheets.push(APP_CONFIG.sheets.portfolioNotes);
  SETUP_addCheck_(report, 'Portfolio notes sheet', SETUP_CONFIG.statuses.FIXED, 'Portfolio notes sheet added with a concise project summary.');
}

function SETUP_initializeSupportSheets_(report) {
  REQ_ensureTempStoreSheet();
  SETUP_addCheck_(report, 'Hidden temp store sheet', SETUP_CONFIG.statuses.FIXED, 'Hidden temporary request storage sheet is ready.');
}

function SETUP_initializeLogs_(report) {
  LOG_initializeAllLogs();
  SETUP_addCheck_(report, 'Log sheets', SETUP_CONFIG.statuses.FIXED, 'Log sheets are created and formatted with borders and wrapping.');
}

function SETUP_installTriggers_(report) {
  APP_ensureTrigger_(APP_CONFIG.triggerHandlers.encoderEdit, 'onEdit');
  APP_ensureTrigger_(APP_CONFIG.triggerHandlers.requestEdit, 'onEdit');
  APP_ensureTrigger_(APP_CONFIG.triggerHandlers.inventoryEdit, 'onEdit');
  APP_ensureTrigger_(APP_CONFIG.triggerHandlers.spreadsheetChange, 'onChange');
  report.installedTriggers = [
    APP_CONFIG.triggerHandlers.encoderEdit,
    APP_CONFIG.triggerHandlers.requestEdit,
    APP_CONFIG.triggerHandlers.inventoryEdit,
    APP_CONFIG.triggerHandlers.spreadsheetChange
  ];
  SETUP_addCheck_(report, 'Triggers', SETUP_CONFIG.statuses.FIXED, 'Encoder, request form, and inventory edit triggers were installed.');
}

function SETUP_renderSummary_(report) {
  return [
    'Initialization ' + (report.ok ? 'completed' : 'completed with issues') + '.',
    'Created sheets: ' + (report.createdSheets.length ? report.createdSheets.join(', ') : 'none'),
    'Rebuilt sheets: ' + (report.rebuiltSheets.length ? report.rebuiltSheets.join(', ') : 'none'),
    'Installed triggers: ' + (report.installedTriggers.length ? report.installedTriggers.join(', ') : 'none')
  ].join('\n');
}

function VERIFY_newFunctionalReport_() {
  return {
    ok: true,
    generatedAt: APP_getTimestampParts_(APP_now_()).iso,
    checks: []
  };
}

function VERIFY_addCheck_(report, name, ok, details) {
  report.checks.push({
    name: name,
    status: ok ? SETUP_CONFIG.statuses.PASS : SETUP_CONFIG.statuses.FAIL,
    details: details || ''
  });
  if (!ok) {
    report.ok = false;
  }
}

function VERIFY_getSheetRowCount_(sheetName) {
  const sheet = APP_getSpreadsheet_().getSheetByName(sheetName);
  return sheet ? sheet.getLastRow() : 0;
}

function VERIFY_toTimeMs_(value) {
  if (!value) {
    return 0;
  }
  const dateValue = value instanceof Date ? value : new Date(value);
  const timeMs = dateValue.getTime();
  return isNaN(timeMs) ? 0 : timeMs;
}

function runFunctionalVerification() {
  const report = VERIFY_newFunctionalReport_();
  const inventorySheet = APP_getInventorySheet_();
  const inventoryRecords = APP_getInventoryRecords_();
  const stockRecord = inventoryRecords.length > 0 ? inventoryRecords[0] : null;
  const codeRecord = inventoryRecords.find(function (record) {
    return !APP_isBlank_(record.code) && !APP_isNoCode(record.code);
  }) || null;
  const logCountsBefore = {
    changeLogs: VERIFY_getSheetRowCount_(APP_CONFIG.sheets.changeLogs),
    encoderLogs: VERIFY_getSheetRowCount_(APP_CONFIG.sheets.encoderLogs),
    requestLogs: VERIFY_getSheetRowCount_(APP_CONFIG.sheets.requestLogs),
    finalizedLogs: VERIFY_getSheetRowCount_(APP_CONFIG.sheets.finalizedRequestLogs)
  };
  const previousRequestedBy = APP_safeString_(APP_getRequestSheet_().getRange(REQ_CONFIG.REQUEST_FORM_CELLS.REQUESTED_BY).getValue()).trim();
  const verificationRequestedBy = 'SYSTEM VERIFICATION ' + Utilities.formatDate(APP_now_(), APP_getTimeZone_(), 'yyyyMMdd-HHmmss');
  let originalRecord = stockRecord ? APP_findInventoryRecordByKey_(inventorySheet, stockRecord.key) : null;
  let requestFinalizeRaisedQty = false;

  try {
    try {
      onOpen();
      VERIFY_addCheck_(report, 'Menus still load', true, 'Custom menu rebuilt successfully.');
    } catch (error) {
      VERIFY_addCheck_(report, 'Menus still load', false, error.message);
    }

    VERIFY_addCheck_(report, 'Inventory data available', !!stockRecord, stockRecord ? 'Using inventory row ' + stockRecord.rowNumber + ' for verification.' : 'No inventory items found.');
    if (!stockRecord) {
      return report;
    }

    const nameSearch = ENC_searchByNameCore_(stockRecord.name);
    VERIFY_addCheck_(report, 'Encoder search by name', nameSearch.ok && (nameSearch.data.records || []).length > 0, nameSearch.message);

    if (codeRecord) {
      const codeToken = APP_splitCodes_(codeRecord.code)[0] || APP_safeString_(codeRecord.code).trim();
      const codeSearch = ENC_searchByCodeCore_(codeToken);
      VERIFY_addCheck_(report, 'Encoder search by code', codeSearch.ok && (codeSearch.data.records || []).length > 0, codeSearch.message);
    } else {
      VERIFY_addCheck_(report, 'Encoder search by code', false, 'No inventory item with a usable code was available for verification.');
    }

    const noNameValue = 'NO_MATCH_NAME_' + Utilities.getUuid().slice(0, 8).toUpperCase();
    const noNameResult = ENC_searchByNameCore_(noNameValue);
    const noCodeValue = 'NO_MATCH_CODE_' + Utilities.getUuid().slice(0, 8).toUpperCase();
    const noCodeResult = ENC_searchByCodeCore_(noCodeValue);
    VERIFY_addCheck_(
      report,
      'No-result prompt behavior',
      ENC_shouldShowNoResultPrompt_(noNameValue, noNameResult) && !noCodeResult.ok && noCodeResult.message === ENC_getNoResultMessage_(),
      'Encoder uses the message: ' + ENC_getNoResultMessage_()
    );

    const beforeAddRecord = APP_findInventoryRecordByKey_(inventorySheet, stockRecord.key);
    const addResult = APP_applyInventoryAction_(stockRecord.key, APP_CONFIG.actions.ADD, 1, {
      allowNegative: false,
      requestedBy: verificationRequestedBy,
      source: 'SYSTEM'
    });
    const afterAddRecord = APP_findInventoryRecordByKey_(inventorySheet, stockRecord.key);
    const subtractResult = APP_applyInventoryAction_(stockRecord.key, APP_CONFIG.actions.SUBTRACT, 1, {
      allowNegative: false,
      requestedBy: verificationRequestedBy,
      source: 'SYSTEM'
    });
    const afterSubtractRecord = APP_findInventoryRecordByKey_(inventorySheet, stockRecord.key);
    VERIFY_addCheck_(
      report,
      'Stock add/subtract',
      addResult.ok && subtractResult.ok && afterAddRecord && afterSubtractRecord && afterAddRecord.qty === beforeAddRecord.qty + 1 && afterSubtractRecord.qty === beforeAddRecord.qty,
      addResult.message + ' / ' + subtractResult.message
    );
    VERIFY_addCheck_(
      report,
      'DATE INV. update after stock change',
      afterAddRecord && VERIFY_toTimeMs_(afterAddRecord.dateInv) >= VERIFY_toTimeMs_(beforeAddRecord.dateInv),
      'Previous: ' + APP_safeString_(beforeAddRecord.dateInv) + ' | Updated: ' + APP_safeString_(afterAddRecord && afterAddRecord.dateInv)
    );

    const requestSheet = APP_getRequestSheet_();
    requestSheet.getRange(REQ_CONFIG.REQUEST_FORM_CELLS.REQUESTED_BY).setValue(verificationRequestedBy);
    const stageResult = REQ_stageInventoryRecord_(stockRecord, APP_CONFIG.actions.ADD, 1, {
      reqSheet: requestSheet,
      clearEntryFields: false
    });
    const stagedItems = stageResult.ok && stageResult.data && stageResult.data.sessionId ? REQ_getSessionItems(stageResult.data.sessionId) : [];
    const finalizeResult = REQ_finalizeRequestCore_();
    requestFinalizeRaisedQty = finalizeResult.ok;
    const afterFinalizeRecord = APP_findInventoryRecordByKey_(inventorySheet, stockRecord.key);
    VERIFY_addCheck_(
      report,
      'Request Form staging and finalizing',
      stageResult.ok && stagedItems.length > 0 && finalizeResult.ok,
      stageResult.message + ' / ' + finalizeResult.message
    );
    VERIFY_addCheck_(
      report,
      'DATE INV. update after request finalizing',
      afterFinalizeRecord && VERIFY_toTimeMs_(afterFinalizeRecord.dateInv) >= VERIFY_toTimeMs_(afterSubtractRecord.dateInv),
      'Post-finalize DATE INV.: ' + APP_safeString_(afterFinalizeRecord && afterFinalizeRecord.dateInv)
    );

    const logCountsAfter = {
      changeLogs: VERIFY_getSheetRowCount_(APP_CONFIG.sheets.changeLogs),
      encoderLogs: VERIFY_getSheetRowCount_(APP_CONFIG.sheets.encoderLogs),
      requestLogs: VERIFY_getSheetRowCount_(APP_CONFIG.sheets.requestLogs),
      finalizedLogs: VERIFY_getSheetRowCount_(APP_CONFIG.sheets.finalizedRequestLogs)
    };
    VERIFY_addCheck_(
      report,
      'Logging still works',
      logCountsAfter.changeLogs > logCountsBefore.changeLogs && logCountsAfter.encoderLogs > logCountsBefore.encoderLogs && logCountsAfter.requestLogs > logCountsBefore.requestLogs && logCountsAfter.finalizedLogs > logCountsBefore.finalizedLogs,
      JSON.stringify({ before: logCountsBefore, after: logCountsAfter })
    );
  } finally {
    const requestSheet = APP_getRequestSheet_();
    try {
      requestSheet.getRange(REQ_CONFIG.REQUEST_FORM_CELLS.REQUESTED_BY).setValue(verificationRequestedBy);
      const verificationContext = REQ_getRequestContext_(requestSheet);
      const sessionId = REQ_findExistingSessionId_(verificationContext.contextKey);
      if (!APP_isBlank_(sessionId)) {
        REQ_clearSession(sessionId);
      }
    } catch (cleanupError) {
      console.log('Verification request cleanup warning: ' + cleanupError.message);
    }
    requestSheet.getRange(REQ_CONFIG.REQUEST_FORM_CELLS.REQUESTED_BY).setValue(previousRequestedBy);
    REQ_refreshItemsTaken('');

    if (requestFinalizeRaisedQty) {
      try {
        APP_applyInventoryAction_(stockRecord.key, APP_CONFIG.actions.SUBTRACT, 1, {
          allowNegative: false,
          requestedBy: verificationRequestedBy,
          source: 'SYSTEM'
        });
      } catch (restoreError) {
        console.log('Verification stock restore warning: ' + restoreError.message);
      }
    }

    if (originalRecord) {
      try {
        APP_writeInventoryStockRows_(inventorySheet, [{
          rowNumber: originalRecord.rowNumber,
          qty: originalRecord.qty,
          price: originalRecord.price,
          dateInv: originalRecord.dateInv
        }]);
        APP_invalidateInventorySearchCaches_();
      } catch (restoreDateError) {
        console.log('Verification DATE INV. restore warning: ' + restoreDateError.message);
      }
    }
  }

  return report;
}

function APP_initializeSystem() {
  const report = SETUP_newReport_();
  APP_withScriptLock_(30000, function () {
    SETUP_initializeInventorySheet_(report);
    SETUP_initializeRequestSheet_(report);
    SETUP_initializeEncoderSheet_(report);
    SETUP_initializePortfolioNotesSheet_(report);
    SETUP_initializeSupportSheets_(report);
    SETUP_initializeLogs_(report);
    SETUP_installTriggers_(report);
  });
  report.summary = SETUP_renderSummary_(report);
  APP_safeToast_('System initialized.', APP_CONFIG.appName);
  try {
    SpreadsheetApp.getUi().alert('System Initialization Summary', report.summary, SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (error) {
    console.log(report.summary);
  }
  return report;
}

function runOneClickSetup() {
  return APP_initializeSystem();
}

function runFullVerification() {
  return {
    health: runHealthCheck(),
    functional: runFunctionalVerification()
  };
}

function repairLogsAndProtections() {
  return APP_initializeSystem();
}

function installRequiredTriggers() {
  const report = SETUP_newReport_();
  SETUP_installTriggers_(report);
  report.summary = SETUP_renderSummary_(report);
  return report;
}

function validateSyncConfiguration() {
  return SYNC_getTargetStatus_();
}
