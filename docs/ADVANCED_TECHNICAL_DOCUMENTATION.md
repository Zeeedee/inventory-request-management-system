# Advanced Technical Documentation

This file is for the person maintaining the Apps Script project. It documents what is in this repo now: the bound spreadsheet scripts, the menu, setup helpers, request flow, logging, and view-only sync.

## 1. Project Shape

The project is a bound Google Apps Script project attached to a Google Sheet.

The spreadsheet is both the UI and the data store. Apps Script adds the custom menu, handles edits, writes logs, creates the view-only spreadsheet, and syncs selected source data into it.

Main runtime pieces:

| Piece | Where It Lives |
| --- | --- |
| Configuration | `Config.gs` |
| Shared helpers and inventory logic | `AppConfig.gs` |
| Menu and encoder behavior | `code.gs` |
| Request form behavior | `RequestFormModule.gs` |
| Logs | `AuditLoggingModule.gs` |
| Setup and verification | `SetupVerificationModule.gs` |
| View-only creation and sync | `ViewOnly.gs` |
| Scopes/runtime | `appsscript.json` |

The main user-facing sheets are shown below.

![Inventory sheet](screenshots/Inventory.png)

![Encoder sheet](screenshots/Encoder.png)

![Request form sheet](screenshots/Request_Form.png)

## 2. Apps Script Services Used

The code uses built-in Apps Script services only.

| Service | Used For |
| --- | --- |
| `SpreadsheetApp` | Sheets, ranges, formatting, menus, dialogs, opening spreadsheets. |
| `DriveApp` | Reading the source spreadsheet owner and applying target sheet protection ownership. |
| `PropertiesService` | Document/script properties for stored IDs, sessions, dialog state, and flags. |
| `LockService` | Preventing overlapping writes and duplicate view-only creation. |
| `ScriptApp` | Installing and checking triggers. |
| `Session` | Time zone, current/effective user emails, temporary active user key. |
| `Utilities` | Date formatting and UUIDs. |
| `HtmlService` | Dialog HTML for action prompts and view-only open link. |

The manifest currently requests spreadsheet, Drive, UI, and trigger scopes. Removing scopes will break parts of the system.

## 3. File-By-File Notes

### `Config.gs`

`Config.gs` defines the `APP_CONFIG` object. Most modules read from it.

Important areas:

- `sheets`: sheet tab names.
- `properties`: document property keys.
- `structures`: header rows, data start rows, and column counts.
- `templates`: titles, help text, display headers, and widths.
- `viewOnly`: view-only spreadsheet behavior.
- `inventory`: main inventory layout and column indexes.
- `encoder`: encoder input cells, result table, and layout.
- `request`: request form search table, staged item table, temp store schema, and flags.
- `logs`: log headers, log labels, colors, and action/source names.
- `system.menu`: custom menu labels.
- `triggerHandlers`: function names used when installing triggers.

Be careful with row numbers and column indexes. The scripts read and write by configured positions.

### `AppConfig.gs`

This is the shared helper module. It is larger than the other files because several workflows depend on it.

Useful groups:

| Group | Examples |
| --- | --- |
| Basic helpers | `APP_safeString_`, `APP_isBlank_`, `APP_toNumber_`, `APP_normalizeText_`. |
| Actor/time/result helpers | `APP_getActor_`, `APP_getTimestampParts_`, `APP_result_`, `APP_fail_`, `APP_createError_`. |
| Sheet helpers | `APP_getSheet_`, `APP_getStructureConfig_`, `APP_getLiveHeadersForStructure_`. |
| Inventory lookup | `APP_getInventorySnapshot_`, `APP_getInventoryIndexes_`, `APP_findInventoryByCode_`, `APP_searchInventoryByName_`. |
| Inventory writes | `APP_applyInventoryAction_`, `APP_writeInventoryStockRows_`. |
| Dialog state | `APP_registerPendingActionDialog_`, `APP_submitActionDialog_`, `APP_cancelActionDialog_`. |
| Locks/triggers | `APP_withScriptLock_`, `APP_ensureTrigger_`. |
| Checks | `runHealthCheck`, `APP_verifyRequiredConfig_`, `APP_verifyCallableFunctions_`. |

Do not treat this as a utility dumping ground. Changes here can affect encoder, request, logging, setup, and sync behavior.

### `code.gs`

This file contains the spreadsheet menu and the encoder edit flow.

`onOpen()` creates this menu:

```text
Inventory & Request Management System
```

The menu calls these global functions:

| Menu Label | Handler |
| --- | --- |
| `Initialize System` | `APP_initializeSystem` |
| `Run Verification` | `runHealthCheck` |
| `Open/Create View-Only Spreadsheet` | `SYNC_createViewOnlySpreadsheet` |
| `Sync View-Only Now` | `manualSyncToTarget` |
| `Finalize Staged Request` | `REQ_finalizeRequest` |
| `Store Request Temporarily` | `REQ_storeTemporarily` |
| `Pull Temporary Request` | `REQ_pullStoredRequest` |
| `Delete Temporary Request` | `REQ_deleteStoredRequest` |
| `Clear Staged Request` | `REQ_clearItemsTaken` |

Encoder-related functions use the `ENC_` prefix. `handleEncoderEdit(e)` is the installed edit trigger handler for the encoder sheet.

### `RequestFormModule.gs`

This module handles the request form.

Real responsibilities:

- Reads request context from the request form.
- Searches inventory from request form input fields.
- Stages selected items.
- Stores unfinished requests in a temp store.
- Pulls stored requests back into the form.
- Deletes stored requests.
- Finalizes staged requests.
- Calls logging and view-only sync after finalization.

Important functions:

| Function | Purpose |
| --- | --- |
| `REQ_initialize` | Initializes request form pieces. |
| `REQ_handleEdit` | Edit trigger entry point for request form actions. |
| `REQ_searchByName`, `REQ_searchByCode`, `REQ_searchByRow` | Search helpers called by edit flow. |
| `REQ_stageInventoryRecord_` | Adds selected inventory item to the staged request. |
| `REQ_storeTemporarily` | Saves the current staged request. |
| `REQ_pullStoredRequest` | Restores a stored request. |
| `REQ_deleteStoredRequest` | Removes a stored request. |
| `REQ_finalizeRequest` | Menu entry point for finalizing. |
| `REQ_finalizeRequestCore_` | Main finalization logic. |

The temp store uses the `__REQUEST_TEMP_STORE` sheet and document properties. Do not change that schema casually once users have stored requests.

### `AuditLoggingModule.gs`

This module creates, formats, and writes log sheets.

![Change logs](screenshots/Change_Logs.png)

![Encoder logs](screenshots/Encoder_Logs.png)

![Request form logs](screenshots/Request_Form_Logs.png)

![Finalized request logs](screenshots/Finalized_Request_Logs.png)

Important functions:

| Function | Purpose |
| --- | --- |
| `LOG_ensureLogSheet` | Creates/formats one log sheet. |
| `LOG_appendEntries` | Appends multiple log rows. |
| `LOG_appendEntry` | Appends one log row. |
| `LOG_appendChangeLog` | Writes inventory change logs. |
| `LOG_encoderAction` | Writes encoder logs. |
| `LOG_requestFormAction` | Writes request form logs. |
| `LOG_appendFinalizedRequestLogs` | Writes finalized request rows. |
| `LOG_handleInventoryEdit` | Edit trigger for manual inventory edits. |
| `LOG_initializeAllLogs` | Initializes all log sheets. |

The log layout comes from `APP_CONFIG.logs` and the structures in `APP_CONFIG.structures`.

### `SetupVerificationModule.gs`

This module is what `Initialize System` uses.

Setup flow:

```text
APP_initializeSystem()
  -> SETUP_initializeInventorySheet_()
  -> SETUP_initializeEncoderSheet_()
  -> SETUP_initializeRequestSheet_()
  -> SETUP_initializePortfolioNotesSheet_()
  -> SETUP_initializeSupportSheets_()
  -> SETUP_initializeLogs_()
  -> SETUP_installTriggers_()
  -> summary alert
```

Other useful functions:

| Function | Purpose |
| --- | --- |
| `runOneClickSetup` | Alias for `APP_initializeSystem`. |
| `runFullVerification` | Runs health and functional verification. |
| `repairLogsAndProtections` | Calls setup again. |
| `installRequiredTriggers` | Installs configured triggers. |
| `validateSyncConfiguration` | Returns view-only target status. |

### `ViewOnly.gs`

This module handles the linked view-only spreadsheet.

Important functions:

| Function | Purpose |
| --- | --- |
| `SYNC_createViewOnlySpreadsheet` | Menu entry point. Creates once or opens existing target. |
| `SYNC_getTargetStatus_` | Checks stored target ID and open status. |
| `fullSyncToTargetSheet` | Full export sync to linked target. |
| `syncRowToTarget` | Row-level inventory sync helper. |
| `manualSyncToTarget` | Menu entry for manual sync. |
| `SYNC_ownerResetViewOnlyLink` | Hidden owner-only reset helper. Not on the normal menu. |
| `SYNC_showOwnerCheckDiagnostic` | Diagnostic helper for owner identity problems. |
| `APP_handleSpreadsheetChange` | Change trigger handler that can call sync. |

The view-only target ID is stored under `APP_CONFIG.properties.viewOnlySpreadsheetId`.

### `appsscript.json`

Current relevant settings:

```json
{
  "timeZone": "Asia/Manila",
  "exceptionLogging": "STACKDRIVER",
  "runtimeVersion": "V8"
}
```

The OAuth scopes are required by the current code. If a scope is removed, authorization may still pass for some functions but fail later when a specific service is used.

## 4. Config Reference

This section lists the config areas that are most likely to be changed or accidentally broken.

### Sheet Names

`APP_CONFIG.sheets` must match the actual tab names.

If a sheet is renamed in Google Sheets but not in config, helper functions such as `APP_getSheet_()` will fail.

### Structures

`APP_CONFIG.structures` tells the code where headers and data start.

Example from the inventory structure:

```javascript
inventory: {
  sheetName: 'INVENTORY',
  headerRow: 4,
  dataStartRow: 5,
  columnCount: 6,
  dynamicHeaderWidth: true
}
```

Changing `headerRow` or `dataStartRow` affects reads, writes, formatting, and sync. Test layout changes in a copy first.

### Inventory Columns

`APP_CONFIG.inventory.cols` maps logical fields to sheet columns:

```javascript
cols: {
  NAME: 1,
  CODE: 2,
  QTY: 3,
  PRICE: 4,
  DATE_INV: 5,
  REMARKS: 6
}
```

If these numbers are wrong, stock updates and record building will use the wrong cells.

### Request Form Config

The request form has two table areas:

- `request.searchTable` for search results.
- `request.itemsTable` for staged items.

Both use fixed column mappings. If the sheet layout changes, update config and test staging, removal, temporary storage, and finalization.

### View-Only Config

Current view-only defaults:

```javascript
enabledByDefault: true,
createOnceOnly: true,
ownerOnlyCreate: true,
ownerEmailOverride: '',
ownerTemporaryUserKeyOverride: '',
autoOpenAfterCreate: true,
autoOpenExisting: true,
allowRecreateIfMissing: false,
showUrlFallbackAlert: true,
creationLockWaitMs: 30000
```

Notes:

- Keep `createOnceOnly` true unless you intentionally want repeated target creation.
- Keep `ownerOnlyCreate` true for normal use.
- Use `ownerTemporaryUserKeyOverride` only when Apps Script hides the owner's email.
- Keep `allowRecreateIfMissing` false unless you are recovering from a deleted or inaccessible target.

### Trigger Handler Names

Configured handlers:

```javascript
triggerHandlers: {
  encoderEdit: 'handleEncoderEdit',
  requestEdit: 'REQ_handleEdit',
  inventoryEdit: 'LOG_handleInventoryEdit',
  spreadsheetChange: 'APP_handleSpreadsheetChange'
}
```

The strings must match real global functions.

## 5. Trigger Flow

`onOpen()` is the simple trigger that creates the menu.

Installable triggers are installed during setup by `SETUP_installTriggers_()`, which calls `APP_ensureTrigger_()`.

Current trigger handlers:

| Handler | Event | What It Handles |
| --- | --- | --- |
| `handleEncoderEdit` | edit | Encoder search/action cells. |
| `REQ_handleEdit` | edit | Request form staging and removals. |
| `LOG_handleInventoryEdit` | edit | Manual inventory edit logging. |
| `APP_handleSpreadsheetChange` | change | Structural changes that may need view-only sync. |

When debugging triggers, check Apps Script `Executions` first. Trigger errors usually show the handler name and line number.

## 6. View-Only Creation Logic

The main entry point is `SYNC_createViewOnlySpreadsheet()`.

Current flow:

```text
Acquire document lock
Check stored target status
If target opens, open that target and return
If target ID exists but target is missing/inaccessible, block recreation unless allowed
Check owner permission
Create spreadsheet
Store target ID
Store creator and creation timestamp
Initialize target sheets
Run full sync
Show open-link dialog
Release lock
```

The stored target ID is not cleared just because `openById()` fails. That is intentional. It lets the code know a target was already created and prevents accidental replacement when `allowRecreateIfMissing` is false.

### Owner Check

The owner check compares the Drive owner email with visible Apps Script user emails. For Gmail accounts, Apps Script may return no current email. The code then supports `ownerTemporaryUserKeyOverride`.

Diagnostic function:

```javascript
SYNC_showOwnerCheckDiagnostic()
```

Use it from the Apps Script editor if the owner is blocked by the owner-only check.

### Open Link Dialog

`SYNC_openUrlInNewTab_()` tries to open the target with `window.open`. It also keeps a visible fallback link and close button. The dialog does not auto-close because browser popup blocking is common.

## 7. Sync Logic

Full sync entry points:

- `fullSyncToTargetSheet()`
- `syncToTargetSheet()`
- `manualSyncToTarget()`

Row sync entry point:

- `syncRowToTarget(rowNumber)`

The export list comes from `APP_CONFIG.viewOnly.exports`. The current configured export is inventory only.

Target sheet setup uses:

- `SYNC_initializeTargetSheet_()`
- `SYNC_getTargetSheet_()`
- `SYNC_ensureTargetCapacity_()`
- `SYNC_applyTargetFormatting_()`
- `SYNC_applySheetProtection_()`

The source payload is read by `SYNC_getSourcePayload_(exportKey)` using the configured structure.

## 8. Error Handling

The project uses simple result objects and thrown errors.

Result helpers:

```javascript
APP_result_(ok, message, data)
APP_fail_(code, message, details)
APP_createError_(code, message, details)
```

Menu actions generally show a UI alert or toast. Lower-level helpers return result objects where possible.

For view-only creation errors, `SYNC_formatErrorMessage_()` includes `error.details` so owner diagnostics are visible to the user.

## 9. Debugging Checklist

Use this order when something breaks:

1. Check whether the right spreadsheet copy is open.
2. Refresh the spreadsheet and see whether the menu appears.
3. Run `Run Verification`.
4. Open Apps Script `Executions` and look at the latest failed run.
5. Check whether all `.gs` files exist and are saved.
6. Check `appsscript.json` scopes if authorization or service access fails.
7. Check triggers if edit-driven behavior is not running.
8. Compare actual sheet names and header rows with `Config.gs`.
9. For view-only issues, run `validateSyncConfiguration()` or `SYNC_showOwnerCheckDiagnostic()`.

## 10. Changes That Need Extra Care

These changes are easy to get wrong:

| Change | Why It Is Risky |
| --- | --- |
| Renaming sheets | Config must match tab names. |
| Moving headers | Structures and setup ranges depend on row numbers. |
| Moving request form columns | Staging and finalization depend on configured column indexes. |
| Changing property keys | Existing stored target ID, sessions, and state may become unreachable. |
| Deleting triggers | Edit and change handlers stop running. |
| Deleting the view-only target | Stored link still exists, and recreation is blocked by default. |
| Editing log sheets manually | Makes audit history unreliable. |

## 11. Recovery Notes

### Broken Layout

Run `Initialize System` first. If the data itself was damaged, use Google Sheets version history or a backup copy.

### Missing Or Wrong View-Only Link

Use the hidden function only after confirming the old target is wrong, deleted, or unusable:

```javascript
SYNC_ownerResetViewOnlyLink()
```

It clears the stored target ID and creation metadata. It is not on the normal menu.

### Missing Triggers

Run:

```javascript
installRequiredTriggers()
```

or run `Initialize System` from the menu.

### Owner Detection Failure

Run:

```javascript
SYNC_showOwnerCheckDiagnostic()
```

If the owner email is correct but `currentUserEmails` is empty, set `APP_CONFIG.viewOnly.ownerTemporaryUserKeyOverride` to the reported current user key.

## 12. Local Review Notes

These files are Apps Script files, not Node modules. A local JavaScript syntax check can catch basic parse errors, but it cannot run Google services.

Useful local check on Windows PowerShell through `cmd`:

```powershell
cmd /v:on /c "for %f in (*.gs) do @node --check --input-type=commonjs < %f || exit /b 1"
```

Real runtime testing still needs to happen inside the bound spreadsheet.
