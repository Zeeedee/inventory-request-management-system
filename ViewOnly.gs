function SYNC_isEnabled_() {
  const propertyValue = APP_getDocumentProperties_().getProperty(APP_CONFIG.properties.viewOnlySyncEnabled);
  if (propertyValue === null || propertyValue === '') {
    return APP_CONFIG.viewOnly.enabledByDefault !== false;
  }
  return APP_normalizeText_(propertyValue) !== 'FALSE';
}

function setViewOnlySyncEnabled(enabled) {
  APP_getDocumentProperties_().setProperty(APP_CONFIG.properties.viewOnlySyncEnabled, enabled ? 'true' : 'false');
  return APP_result_(true, enabled ? 'View-only sync enabled.' : 'View-only sync disabled.');
}

function SYNC_getTargetSpreadsheetId_() {
  return APP_getDocumentProperties_().getProperty(APP_CONFIG.properties.viewOnlySpreadsheetId) || '';
}

function SYNC_setTargetSpreadsheetId_(spreadsheetId) {
  APP_getDocumentProperties_().setProperty(APP_CONFIG.properties.viewOnlySpreadsheetId, spreadsheetId);
}

function SYNC_clearTargetSpreadsheetId_() {
  APP_getDocumentProperties_().deleteProperty(APP_CONFIG.properties.viewOnlySpreadsheetId);
}

function SYNC_getSourceSpreadsheetOwnerEmail_() {
  if (!APP_isBlank_(APP_CONFIG.viewOnly.ownerEmailOverride || '')) {
    return APP_CONFIG.viewOnly.ownerEmailOverride;
  }
  try {
    return DriveApp.getFileById(APP_getSpreadsheet_().getId()).getOwner().getEmail() || '';
  } catch (error) {
    return '';
  }
}

function SYNC_getCurrentUserEmails_() {
  const emails = [];
  const addEmail = function (email) {
    const normalizedEmail = APP_normalizeText_(email || '');
    if (normalizedEmail !== '' && emails.indexOf(normalizedEmail) === -1) {
      emails.push(normalizedEmail);
    }
  };

  try {
    addEmail(Session.getActiveUser().getEmail());
  } catch (error1) {
    // Apps Script can hide active user email for consumer accounts.
  }

  try {
    addEmail(Session.getEffectiveUser().getEmail());
  } catch (error2) {
    // Fall through to APP_getActor_ below.
  }

  addEmail(APP_getActor_().email || '');
  return emails;
}

function SYNC_getCurrentUserKey_() {
  try {
    return Session.getTemporaryActiveUserKey() || APP_getActor_().key || '';
  } catch (error) {
    return APP_getActor_().key || '';
  }
}

function SYNC_getOwnerCheckDetails_() {
  const ownerEmail = SYNC_getSourceSpreadsheetOwnerEmail_();
  const currentUserEmails = SYNC_getCurrentUserEmails_();
  const currentUserKey = SYNC_getCurrentUserKey_();
  return {
    ownerEmail: ownerEmail,
    currentUserEmails: currentUserEmails,
    currentUser: currentUserEmails.join(', ') || APP_getActor_().display,
    currentUserKey: currentUserKey,
    ownerEmailOverrideConfigured: !APP_isBlank_(APP_CONFIG.viewOnly.ownerEmailOverride || ''),
    ownerTemporaryUserKeyOverrideConfigured: !APP_isBlank_(APP_CONFIG.viewOnly.ownerTemporaryUserKeyOverride || '')
  };
}

function SYNC_isCurrentUserSourceOwner_() {
  const details = SYNC_getOwnerCheckDetails_();
  const ownerEmail = APP_normalizeText_(details.ownerEmail);
  if (ownerEmail !== '' && details.currentUserEmails.indexOf(ownerEmail) !== -1) {
    return true;
  }

  const ownerUserKeyOverride = APP_normalizeText_(APP_CONFIG.viewOnly.ownerTemporaryUserKeyOverride || '').replace(/^USER:/, '');
  const currentUserKey = APP_normalizeText_(details.currentUserKey || '').replace(/^USER:/, '');
  return ownerUserKeyOverride !== '' && currentUserKey !== '' && ownerUserKeyOverride === currentUserKey;
}

function SYNC_assertCurrentUserCanCreate_() {
  if (APP_CONFIG.viewOnly.ownerOnlyCreate === false) {
    return;
  }
  if (!SYNC_isCurrentUserSourceOwner_()) {
    const details = SYNC_getOwnerCheckDetails_();
    throw APP_createError_(
      'OWNER_ONLY_VIEW_ONLY_CREATE',
      'Only the owner of the main spreadsheet can create the view-only spreadsheet.',
      details
    );
  }
}

function SYNC_formatErrorMessage_(error) {
  let message = error.message || String(error);
  if (error.details) {
    message += '\n\nDetails: ' + JSON.stringify(error.details);
  }
  return message;
}

function SYNC_showOwnerCheckDiagnostic() {
  const details = SYNC_getOwnerCheckDetails_();
  SpreadsheetApp.getUi().alert(
    'View-Only Owner Check Diagnostic',
    JSON.stringify(details, null, 2),
    SpreadsheetApp.getUi().ButtonSet.OK
  );
  return APP_result_(true, 'Owner check diagnostic ready.', details);
}

function SYNC_buildSpreadsheetTitle_() {
  const timestamp = Utilities.formatDate(APP_now_(), APP_getTimeZone_(), 'yyyyMMdd-HHmmss');
  return [
    APP_CONFIG.templates.viewOnly.spreadsheetTitlePrefix,
    APP_CONFIG.templates.viewOnly.spreadsheetTitleSuffix,
    timestamp
  ].join(' - ');
}

function SYNC_getExportKeys_() {
  return Object.keys(APP_CONFIG.viewOnly.exports || {});
}

function SYNC_getExportConfig_(exportKey) {
  const exportConfig = APP_CONFIG.viewOnly.exports[exportKey];
  if (!exportConfig) {
    throw APP_createError_('UNKNOWN_VIEW_ONLY_EXPORT', 'Unknown view-only export: ' + exportKey);
  }
  return exportConfig;
}

function SYNC_getExportKeyForStructure_(structureKey) {
  const exportKeys = SYNC_getExportKeys_();
  for (let index = 0; index < exportKeys.length; index += 1) {
    if (SYNC_getExportConfig_(exportKeys[index]).structureKey === structureKey) {
      return exportKeys[index];
    }
  }
  return '';
}

function SYNC_getExportKeyForSheetName_(sheetName) {
  const exportKeys = SYNC_getExportKeys_();
  for (let index = 0; index < exportKeys.length; index += 1) {
    const structure = APP_getStructureConfig_(SYNC_getExportConfig_(exportKeys[index]).structureKey);
    if (structure.sheetName === sheetName) {
      return exportKeys[index];
    }
  }
  return '';
}

function SYNC_openTargetSpreadsheet_() {
  const spreadsheetId = SYNC_getTargetSpreadsheetId_();
  if (APP_isBlank_(spreadsheetId)) {
    throw APP_createError_('MISSING_VIEW_ONLY_SPREADSHEET', 'No linked view-only spreadsheet was found. Create one from the menu first.');
  }
  try {
    return SpreadsheetApp.openById(spreadsheetId);
  } catch (error) {
    throw APP_createError_('MISSING_VIEW_ONLY_SPREADSHEET', 'The linked view-only spreadsheet is missing or inaccessible. Create it again from the menu.');
  }
}

function SYNC_setCreationMetadata_() {
  const actor = APP_getActor_();
  const parts = APP_getTimestampParts_(APP_now_());
  const properties = APP_getDocumentProperties_();
  properties.setProperty(APP_CONFIG.properties.viewOnlyCreatedBy, actor.email || actor.display || 'unknown');
  properties.setProperty(APP_CONFIG.properties.viewOnlyCreatedAt, parts.iso);
}

function SYNC_openUrlInNewTab_(url, title, message) {
  const rawUrl = String(url || '');
  const safeUrl = rawUrl
    .replace(/&/g, '&amp;')
    .replace(/"/g, '&quot;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
  const scriptUrl = JSON.stringify(rawUrl)
    .replace(/</g, '\\u003c')
    .replace(/>/g, '\\u003e');
  const safeTitle = String(title || 'Opening View-Only Spreadsheet')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
  const safeMessage = String(message || 'Your view-only spreadsheet should open in a new tab.')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');

  const html = HtmlService.createHtmlOutput(
    '<!doctype html><html><head><base target="_top"></head>' +
    '<body style="font-family:Arial,sans-serif;padding:18px;">' +
    '<h3 style="margin-top:0;">' + safeTitle + '</h3>' +
    '<p>' + safeMessage + '</p>' +
    '<p>If a new tab did not open, use this link:</p>' +
    '<p><a href="' + safeUrl + '" target="_blank" rel="noopener">Open view-only spreadsheet</a></p>' +
    '<p><button onclick="google.script.host.close()">Close</button></p>' +
    '<script>' +
    'window.open(' + scriptUrl + ', "_blank", "noopener");' +
    '</script>' +
    '</body></html>'
  ).setWidth(380).setHeight(220);

  SpreadsheetApp.getUi().showModalDialog(html, safeTitle);
}

function SYNC_getTargetSheet_(targetSpreadsheet, exportKey) {
  const spreadsheet = targetSpreadsheet || SYNC_openTargetSpreadsheet_();
  const exportConfig = SYNC_getExportConfig_(exportKey);
  let sheet = spreadsheet.getSheetByName(exportConfig.targetSheetName);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(exportConfig.targetSheetName);
  }
  if (exportConfig.tabColor) {
    sheet.setTabColor(exportConfig.tabColor);
  }
  return sheet;
}

function SYNC_canReuseInitialSheet_(sheet) {
  if (!sheet) {
    return false;
  }
  if (sheet.getLastRow() > 1 || sheet.getLastColumn() > 1) {
    return false;
  }
  const cellValue = sheet.getRange(1, 1).getValue();
  return APP_isBlank_(cellValue);
}

function SYNC_hasRowData_(row) {
  return (row || []).some(function (cell) {
    if (cell === null || cell === undefined) {
      return false;
    }
    if (cell instanceof Date) {
      return true;
    }
    if (typeof cell === 'number') {
      return true;
    }
    return APP_safeString_(cell).trim() !== '';
  });
}

function SYNC_ensureTargetCapacity_(targetSheet, width, rowCount) {
  if (targetSheet.getMaxColumns() < width) {
    targetSheet.insertColumnsAfter(targetSheet.getMaxColumns(), width - targetSheet.getMaxColumns());
  }
  if (targetSheet.getMaxRows() < rowCount) {
    targetSheet.insertRowsAfter(targetSheet.getMaxRows(), rowCount - targetSheet.getMaxRows());
  }
}

function SYNC_applySheetProtection_(sheet) {
  if (!sheet) {
    return;
  }
  const protections = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
  const protection = protections.length > 0
    ? protections[0]
    : sheet.protect().setDescription('View-only sheet protection');
  protection.setWarningOnly(false);
  protection.setDomainEdit(false);

  try {
    const ownerEmail = DriveApp.getFileById(sheet.getParent().getId()).getOwner().getEmail();
    const editors = protection.getEditors();
    if (editors.length > 0) {
      protection.removeEditors(editors);
    }
    if (!APP_isBlank_(ownerEmail)) {
      protection.addEditor(ownerEmail);
    }
  } catch (error) {
    console.log('Protection warning: ' + error.message);
  }
}

function SYNC_getSourcePayload_(exportKey) {
  const exportConfig = SYNC_getExportConfig_(exportKey);
  const structure = APP_getStructureConfig_(exportConfig.structureKey);
  const sourceSheet = APP_getSheetByStructure_(exportConfig.structureKey, true);
  const headers = APP_getLiveHeadersForStructure_(exportConfig.structureKey);
  const width = headers.length || structure.columnCount || sourceSheet.getLastColumn();
  const lastRow = sourceSheet.getLastRow();
  const rows = lastRow < structure.dataStartRow
    ? []
    : sourceSheet.getRange(structure.dataStartRow, 1, lastRow - structure.dataStartRow + 1, width).getValues().filter(SYNC_hasRowData_);

  return {
    sourceSheet: sourceSheet,
    structure: structure,
    headers: headers,
    rows: rows,
    width: width
  };
}

function SYNC_applyTargetFormatting_(targetSheet, exportKey, headers, rowCount) {
  const exportConfig = SYNC_getExportConfig_(exportKey);
  const sheet = targetSheet;
  const width = headers.length;
  sheet.clearFormats();
  sheet.getRange(1, 1, 1, width)
    .setValues([headers])
    .setFontWeight('bold')
    .setBackground(APP_CONFIG.logs.colors.headerBg)
    .setFontColor(APP_CONFIG.logs.colors.headerText)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrap(true)
    .setBorder(true, true, true, true, true, true, APP_CONFIG.logs.colors.border, SpreadsheetApp.BorderStyle.SOLID);

  sheet.setFrozenRows(exportConfig.frozenRows || 1);
  const dataRowCount = Math.max(rowCount, 1);
  const backgrounds = [];
  for (let rowIndex = 0; rowIndex < dataRowCount; rowIndex += 1) {
    const color = rowIndex % 2 === 0 ? APP_CONFIG.system.colors.rowAlt1 : APP_CONFIG.system.colors.rowAlt2;
    const rowColors = [];
    for (let columnIndex = 0; columnIndex < width; columnIndex += 1) {
      rowColors.push(color);
    }
    backgrounds.push(rowColors);
  }
  sheet.getRange(2, 1, dataRowCount, width)
    .setBackgrounds(backgrounds)
    .setWrap(true)
    .setVerticalAlignment('middle')
    .setBorder(true, true, true, true, true, true, APP_CONFIG.logs.colors.border, SpreadsheetApp.BorderStyle.SOLID);

  if (exportConfig.columnWidths && exportConfig.columnWidths.length) {
    exportConfig.columnWidths.forEach(function (columnWidth, index) {
      sheet.setColumnWidth(index + 1, columnWidth);
    });
  } else {
    for (let column = 1; column <= width; column += 1) {
      sheet.setColumnWidth(column, 160);
    }
  }

  if (rowCount === 0) {
    sheet.getRange(2, 1).setValue('No data available.');
  }
}

function SYNC_syncExportSheet_(targetSpreadsheet, exportKey) {
  const spreadsheet = targetSpreadsheet || SYNC_openTargetSpreadsheet_();
  const payload = SYNC_getSourcePayload_(exportKey);
  const targetSheet = SYNC_getTargetSheet_(spreadsheet, exportKey);
  targetSheet.clearContents();

  SYNC_ensureTargetCapacity_(targetSheet, payload.width, payload.rows.length + 2);

  SYNC_applyTargetFormatting_(targetSheet, exportKey, payload.headers, payload.rows.length);
  if (payload.rows.length > 0) {
    targetSheet.getRange(2, 1, payload.rows.length, payload.width).setValues(payload.rows);
  }
  SYNC_applySheetProtection_(targetSheet);
  return APP_result_(true, 'Synced ' + exportKey + '.', {
    sheetName: targetSheet.getName(),
    rowCount: payload.rows.length
  });
}

function SYNC_syncInventoryRow_(targetSpreadsheet, rowNumber) {
  const safeRowNumber = APP_toNumber_(rowNumber, 0);
  if (safeRowNumber < APP_CONFIG.inventory.startRow) {
    return SYNC_syncExportSheet_(targetSpreadsheet, 'inventory');
  }

  const sourceSheet = APP_getInventorySheet_();
  const structure = APP_getStructureConfig_('inventory');
  const headers = APP_getLiveHeadersForStructure_('inventory');
  const width = headers.length || structure.columnCount;
  const targetSheet = SYNC_getTargetSheet_(targetSpreadsheet, 'inventory');
  const targetRow = 2 + (safeRowNumber - structure.dataStartRow);

  SYNC_ensureTargetCapacity_(targetSheet, width, targetRow + 1);
  SYNC_applyTargetFormatting_(targetSheet, 'inventory', headers, Math.max(targetSheet.getLastRow() - 1, 1));

  const sourceValues = sourceSheet.getRange(safeRowNumber, 1, 1, width).getValues()[0];
  if (SYNC_hasRowData_(sourceValues)) {
    targetSheet.getRange(targetRow, 1, 1, width).setValues([sourceValues]);
  } else {
    targetSheet.getRange(targetRow, 1, 1, width).clearContent();
  }
  targetSheet.getRange(targetRow, 1, 1, width)
    .setWrap(true)
    .setVerticalAlignment('middle')
    .setBorder(true, true, true, true, true, true, APP_CONFIG.logs.colors.border, SpreadsheetApp.BorderStyle.SOLID);
  SYNC_applySheetProtection_(targetSheet);
  return APP_result_(true, 'Synced inventory row.', {
    rowNumber: safeRowNumber,
    targetRow: targetRow
  });
}

function SYNC_syncExportKey_(exportKey) {
  if (!SYNC_isEnabled_() || APP_isBlank_(SYNC_getTargetSpreadsheetId_())) {
    return APP_result_(true, 'View-only sync skipped.', { exportKey: exportKey || '' });
  }
  if (!APP_CONFIG.viewOnly.exports[exportKey]) {
    return APP_result_(true, 'View-only sync skipped.', { exportKey: exportKey || '' });
  }
  try {
    return SYNC_syncExportSheet_(SYNC_openTargetSpreadsheet_(), exportKey);
  } catch (error) {
    return APP_fail_(error.code || 'VIEW_ONLY_SYNC_FAILED', error.message, error.details);
  }
}

function SYNC_syncExportKey(exportKey) {
  return SYNC_syncExportKey_(exportKey);
}

function SYNC_isEnabled() {
  return SYNC_isEnabled_();
}

function SYNC_initializeTargetSheet_(targetSpreadsheet) {
  const spreadsheet = targetSpreadsheet || SYNC_openTargetSpreadsheet_();
  const exportKeys = SYNC_getExportKeys_();

  if (spreadsheet.getSheets().length === 1 && exportKeys.length > 0 && SYNC_canReuseInitialSheet_(spreadsheet.getSheets()[0])) {
    spreadsheet.getSheets()[0].setName(SYNC_getExportConfig_(exportKeys[0]).targetSheetName);
  }

  exportKeys.forEach(function (exportKey) {
    SYNC_getTargetSheet_(spreadsheet, exportKey);
  });

  spreadsheet.getSheets().forEach(function (sheet) {
    const isManagedSheet = exportKeys.some(function (exportKey) {
      return SYNC_getExportConfig_(exportKey).targetSheetName === sheet.getName();
    });
    if (!isManagedSheet && spreadsheet.getSheets().length > 1 && SYNC_canReuseInitialSheet_(sheet)) {
      spreadsheet.deleteSheet(sheet);
    }
  });
  return APP_result_(true, 'View-only sheets initialized.');
}

function SYNC_rebuildTargetFromSource_() {
  return fullSyncToTargetSheet();
}

function SYNC_syncInventoryItemByKey_() {
  return syncRowToTarget(0);
}

function SYNC_syncInventoryRows_() {
  return syncRowToTarget(0);
}

function SYNC_createViewOnlySpreadsheet() {
  const lock = LockService.getDocumentLock();
  let hasLock = false;

  try {
    lock.waitLock(APP_CONFIG.viewOnly.creationLockWaitMs || 30000);
    hasLock = true;

    const existingStatus = SYNC_getTargetStatus_();

    if (APP_CONFIG.viewOnly.createOnceOnly !== false && existingStatus.linked) {
      if (APP_CONFIG.viewOnly.autoOpenExisting !== false) {
        SYNC_openUrlInNewTab_(
          existingStatus.url,
          'View-Only Spreadsheet Already Exists',
          'The existing view-only spreadsheet will open in a new tab.'
        );
      } else if (APP_CONFIG.viewOnly.showUrlFallbackAlert !== false) {
        SpreadsheetApp.getUi().alert('View-Only Spreadsheet Already Exists', existingStatus.url, SpreadsheetApp.getUi().ButtonSet.OK);
      }
      return APP_result_(true, 'Existing view-only spreadsheet opened.', existingStatus);
    }

    if (APP_CONFIG.viewOnly.createOnceOnly !== false && !existingStatus.linked && !APP_isBlank_(SYNC_getTargetSpreadsheetId_())) {
      if (APP_CONFIG.viewOnly.allowRecreateIfMissing !== true) {
        throw APP_createError_(
          'VIEW_ONLY_ALREADY_CREATED_BUT_MISSING',
          'A view-only spreadsheet was already created before, but it is now missing or inaccessible. Re-creation is disabled in Config.gs.',
          { storedSpreadsheetId: SYNC_getTargetSpreadsheetId_() }
        );
      }
    }

    SYNC_assertCurrentUserCanCreate_();

    const spreadsheet = SpreadsheetApp.create(SYNC_buildSpreadsheetTitle_());
    SYNC_setTargetSpreadsheetId_(spreadsheet.getId());
    SYNC_setCreationMetadata_();
    SYNC_initializeTargetSheet_(spreadsheet);

    const syncResult = fullSyncToTargetSheet();
    const message = syncResult.ok
      ? 'View-only spreadsheet created and synced.'
      : 'View-only spreadsheet created, but initial sync needs attention.';

    if (APP_CONFIG.viewOnly.autoOpenAfterCreate !== false) {
      SYNC_openUrlInNewTab_(spreadsheet.getUrl(), 'View-Only Spreadsheet Ready', message);
    } else if (APP_CONFIG.viewOnly.showUrlFallbackAlert !== false) {
      SpreadsheetApp.getUi().alert('View-Only Spreadsheet Ready', message + '\n\nURL: ' + spreadsheet.getUrl(), SpreadsheetApp.getUi().ButtonSet.OK);
    }

    return APP_result_(syncResult.ok, message, {
      spreadsheetId: spreadsheet.getId(),
      url: spreadsheet.getUrl()
    });
  } catch (error) {
    SpreadsheetApp.getUi().alert('View-Only Spreadsheet Error', SYNC_formatErrorMessage_(error), SpreadsheetApp.getUi().ButtonSet.OK);
    return APP_fail_(error.code || 'VIEW_ONLY_CREATE_FAILED', error.message, error.details);
  } finally {
    if (hasLock) {
      lock.releaseLock();
    }
  }
}

function SYNC_getTargetStatus_() {
  const spreadsheetId = SYNC_getTargetSpreadsheetId_();
  if (APP_isBlank_(spreadsheetId)) {
    return {
      ok: false,
      enabled: SYNC_isEnabled_(),
      linked: false,
      message: 'No linked view-only spreadsheet found.'
    };
  }
  try {
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    return {
      ok: true,
      enabled: SYNC_isEnabled_(),
      linked: true,
      spreadsheetId: spreadsheetId,
      spreadsheetName: spreadsheet.getName(),
      url: spreadsheet.getUrl(),
      message: 'Linked view-only spreadsheet is available.'
    };
  } catch (error) {
    return {
      ok: false,
      enabled: SYNC_isEnabled_(),
      linked: false,
      spreadsheetId: spreadsheetId,
      message: 'The linked view-only spreadsheet is missing or inaccessible.'
    };
  }
}

function SYNC_ownerResetViewOnlyLink() {
  SYNC_assertCurrentUserCanCreate_();
  SYNC_clearTargetSpreadsheetId_();
  APP_getDocumentProperties_().deleteProperty(APP_CONFIG.properties.viewOnlyCreatedBy);
  APP_getDocumentProperties_().deleteProperty(APP_CONFIG.properties.viewOnlyCreatedAt);
  SpreadsheetApp.getUi().alert('View-only link reset. The owner can now create a new one.');
  return APP_result_(true, 'View-only link reset.');
}

function fullSyncToTargetSheet() {
  if (!SYNC_isEnabled_()) {
    return APP_result_(true, 'View-only sync is currently disabled.');
  }
  let targetSpreadsheet;
  try {
    targetSpreadsheet = SYNC_openTargetSpreadsheet_();
  } catch (error) {
    return APP_fail_(error.code || 'VIEW_ONLY_SYNC_UNAVAILABLE', error.message, error.details);
  }

  const exportResults = [];
  try {
    SYNC_getExportKeys_().forEach(function (exportKey) {
      exportResults.push(SYNC_syncExportSheet_(targetSpreadsheet, exportKey).data);
    });
    return APP_result_(true, 'View-only spreadsheet synced.', {
      spreadsheetId: targetSpreadsheet.getId(),
      url: targetSpreadsheet.getUrl(),
      exports: exportResults
    });
  } catch (error) {
    return APP_fail_(error.code || 'VIEW_ONLY_SYNC_FAILED', error.message, error.details);
  }
}

function syncToTargetSheet() {
  return fullSyncToTargetSheet();
}

function syncRowToTarget(rowNumber) {
  if (!SYNC_isEnabled_() || APP_isBlank_(SYNC_getTargetSpreadsheetId_())) {
    return APP_result_(true, 'View-only sync skipped.', { rowNumber: rowNumber || 0 });
  }
  try {
    const targetSpreadsheet = SYNC_openTargetSpreadsheet_();
    return SYNC_syncInventoryRow_(targetSpreadsheet, rowNumber || 0);
  } catch (error) {
    return APP_fail_(error.code || 'VIEW_ONLY_SYNC_FAILED', error.message, error.details);
  }
}

function manualSyncToTarget() {
  const result = fullSyncToTargetSheet();
  APP_safeToast_(result.message, result.ok ? 'View Only' : 'View Only Error');
  return result;
}

function forceAuth() {
  return APP_result_(true, 'Authorization is ready when you create or sync the view-only spreadsheet.');
}

function setupSyncTrigger() {
  return APP_result_(true, 'Automatic view-only updates use the main system flows and do not need separate sync triggers.');
}

function APP_handleSpreadsheetChange(e) {
  if (!SYNC_isEnabled_() || APP_isBlank_(SYNC_getTargetSpreadsheetId_())) {
    return;
  }
  const changeType = e && e.changeType ? String(e.changeType) : '';
  if (['INSERT_ROW', 'REMOVE_ROW', 'INSERT_COLUMN', 'REMOVE_COLUMN', 'INSERT_GRID', 'REMOVE_GRID', 'FORMAT'].indexOf(changeType) === -1) {
    return;
  }
  try {
    syncRowToTarget(0);
  } catch (error) {
    console.log('View-only change sync warning: ' + error.message);
  }
}
