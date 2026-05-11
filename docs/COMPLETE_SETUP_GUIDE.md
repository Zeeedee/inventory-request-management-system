# Complete Setup Guide

This guide walks through setting up the spreadsheet from scratch. It is written for the person who will own or maintain the file.

The setup uses Google Sheets and a bound Apps Script project. There is no separate server to install.

## 1. Before You Start

You need:

- A Google account that should own the main spreadsheet.
- Access to Google Drive and Google Sheets.
- A desktop or laptop browser.
- The script files from this project.
- Permission to approve Apps Script authorization prompts.

Files that must be present in Apps Script:

| File | Required |
| --- | --- |
| `Config.gs` | Yes |
| `AppConfig.gs` | Yes |
| `code.gs` | Yes |
| `RequestFormModule.gs` | Yes |
| `AuditLoggingModule.gs` | Yes |
| `SetupVerificationModule.gs` | Yes |
| `ViewOnly.gs` | Yes |
| `appsscript.json` | Yes |

Use one Google account during setup if possible. Multiple signed-in accounts can make authorization and owner checks harder to understand.

## 2. Create Or Choose The Google Account

If you already have the correct Google account, use it. If this is for a team, avoid using an account that only one temporary user controls.

To create a new Google account:

1. Go to `https://accounts.google.com/signup`.
2. Follow the account creation steps.
3. Add recovery options.
4. Sign in with the account before creating the spreadsheet.

For a team file, the owner account matters. The view-only creation check reads the Google Drive owner of the main spreadsheet.

## 3. Set Up A Drive Folder

The folder structure does not affect the script, but it makes the files easier to manage.

Suggested folders:

```text
Inventory System
Inventory System/Main
Inventory System/Backups
Inventory System/Test Copies
Inventory System/Docs
```

Keep the main spreadsheet and the generated view-only spreadsheet somewhere easy to find.

## 4. Create The Spreadsheet

1. Open Google Drive.
2. Open the folder where the file should live.
3. Click `New`.
4. Click `Google Sheets`.
5. Rename the file.

Suggested names:

```text
Inventory & Request Management System
Inventory & Request Management System - Test Copy
```

The spreadsheet can start blank. The setup script creates or repairs the required sheets.

## 5. Open Apps Script

1. Open the spreadsheet.
2. Click `Extensions`.
3. Click `Apps Script`.
4. Wait for the Apps Script editor to open.

In the editor, you will add the project files and paste the code.

## 6. Import The Project Files

### Add `.gs` Files

For each script file:

1. Click the plus icon beside `Files`.
2. Choose `Script`.
3. Name the file to match the project file.
4. Paste the code.
5. Save.

Apps Script may show the file name without `.gs`. That is fine.

### Replace The Default File

If Apps Script created a default `Code.gs`, replace its content with this project's `code.gs`, or delete it and create a new file named `code.gs`.

Do not keep old sample functions if they are not part of this project.

### Add `appsscript.json`

1. Open Apps Script project settings.
2. Turn on `Show appsscript.json manifest file in editor`.
3. Open `appsscript.json`.
4. Replace it with this project's manifest.
5. Save.

The manifest should use V8 runtime and include spreadsheet, Drive, UI, and trigger scopes.

## 7. Review `Config.gs`

Most settings can stay as they are. Review these before first setup.

### Sheet Names

The setup expects these names by default:

| Config Key | Sheet Name |
| --- | --- |
| `inventory` | `INVENTORY` |
| `encoder` | `ENCODER` |
| `requestForm` | `REQUEST FORM` |
| `changeLogs` | `CHANGE LOGS` |
| `encoderLogs` | `ENCODER LOGS` |
| `requestLogs` | `REQUEST FORM LOGS` |
| `finalizedRequestLogs` | `FINALIZED REQUEST LOGS` |
| `portfolioNotes` | `PORTFOLIO NOTES` |
| `requestTempStore` | `__REQUEST_TEMP_STORE` |

Do not rename sheets manually unless you also update `Config.gs`.

### View-Only Settings

Recommended values:

```javascript
createOnceOnly: true,
ownerOnlyCreate: true,
ownerEmailOverride: '',
ownerTemporaryUserKeyOverride: '',
autoOpenAfterCreate: true,
autoOpenExisting: true,
allowRecreateIfMissing: false
```

Leave `ownerTemporaryUserKeyOverride` blank at first. Only fill it if the owner check fails because Apps Script hides your email.

### Settings To Avoid Changing During First Setup

- `APP_CONFIG.properties`
- `APP_CONFIG.structures`
- `APP_CONFIG.inventory.cols`
- `APP_CONFIG.request.searchTable.cols`
- `APP_CONFIG.request.itemsTable.cols`
- `APP_CONFIG.triggerHandlers`

These control how the scripts read and write cells.

## 8. Save And Refresh

After all code is pasted:

1. Click Save in Apps Script.
2. Wait for saving to finish.
3. Go back to the spreadsheet tab.
4. Refresh the spreadsheet.
5. Wait for the custom menu to appear.

The menu name is:

```text
Inventory & Request Management System
```

If it does not appear, refresh again. If it still does not appear, check Apps Script `Executions` for errors.

## 9. Run First Authorization

The first menu action usually asks for authorization.

1. Click `Inventory & Request Management System`.
2. Click `Initialize System`.
3. Choose the owner account.
4. Review the permissions.
5. Approve if this is the code you intended to install.
6. Run `Initialize System` again if the first run only completed authorization.

Google may show a warning that the app is not verified. That is normal for a private Apps Script project. Continue only if you trust the code and you are in the correct spreadsheet.

The script needs permission to:

- Work with spreadsheets.
- Create and open the view-only spreadsheet in Drive.
- Show dialogs.
- Install triggers.

## 10. Run `Initialize System`

Run:

```text
Inventory & Request Management System > Initialize System
```

This creates or repairs:

- `INVENTORY`
- `ENCODER`
- `REQUEST FORM`
- log sheets
- portfolio notes
- temp store sheet
- required triggers

After setup, the main sheets should look like this:

![Inventory sheet](screenshots/Inventory.png)

![Encoder sheet](screenshots/Encoder.png)

![Request form sheet](screenshots/Request_Form.png)

## 11. Check The Setup

Run:

```text
Inventory & Request Management System > Run Verification
```

If verification reports a missing sheet or function, check that all files were pasted and saved.

Also check the bottom tabs. You should see the main sheets and logs.

Log sheets should look similar to these:

![Change logs](screenshots/Change_Logs.png)

![Finalized request logs](screenshots/Finalized_Request_Logs.png)

## 12. Check Triggers

`Initialize System` installs triggers through the script.

Expected handlers:

| Handler | Event |
| --- | --- |
| `handleEncoderEdit` | Edit |
| `REQ_handleEdit` | Edit |
| `LOG_handleInventoryEdit` | Edit |
| `APP_handleSpreadsheetChange` | Change |

To check them:

1. Open Apps Script.
2. Open the Triggers page.
3. Confirm the handlers are listed.

If triggers are missing, run `Initialize System` again or run `installRequiredTriggers` from Apps Script.

## 13. Add Or Check Inventory Data

The `INVENTORY` sheet uses these columns:

| Column | Header |
| --- | --- |
| A | `NAME` |
| B | `CODE` |
| C | `QTY` |
| D | `PRICE PER PIECE` |
| E | `DATE INV.` |
| F | `REMARKS` |

If you are starting fresh, add a few test items before training users.

Example test row:

| NAME | CODE | QTY | PRICE PER PIECE | DATE INV. | REMARKS |
| --- | --- | --- | --- | --- | --- |
| Test Item | TEST001 | 10 | 100 | 01/01/2026 | TEST |

Remove or rename test data before using the sheet for real inventory.

## 14. Test The Main Workflows

Do these tests in a copy if the file already has real data.

### Encoder Test

1. Open `ENCODER`.
2. Search for a known item code.
3. Confirm the result appears.
4. If testing stock actions, use a small test item.
5. Check `ENCODER LOGS` after the action.

### Request Form Test

1. Open `REQUEST FORM`.
2. Enter a requested-by name.
3. Search for a known item.
4. Stage it.
5. Store the request temporarily.
6. Pull it back.
7. Finalize it only if this is a test copy or test item.
8. Check `REQUEST FORM LOGS` and `FINALIZED REQUEST LOGS`.

### Manual Sync Test

Do this after the view-only spreadsheet is created.

1. Change a harmless test value in inventory.
2. Run `Sync View-Only Now`.
3. Open the view-only spreadsheet.
4. Confirm the copied data updated.

## 15. Create The View-Only Spreadsheet

Run this as the owner of the main spreadsheet.

1. Open the main spreadsheet.
2. Click `Inventory & Request Management System`.
3. Click `Open/Create View-Only Spreadsheet`.
4. Wait for the script to finish.
5. If a new tab opens, check the generated spreadsheet.
6. If the tab does not open, click the fallback link in the dialog.

The first successful run stores the target spreadsheet ID. Later clicks open the same target instead of creating another one.

If you get an owner-only error but you are the owner, look at the details. If `currentUserEmails` is empty, copy the reported current user key into:

```javascript
ownerTemporaryUserKeyOverride: 'PASTE_THE_KEY_HERE'
```

Save Apps Script, refresh the spreadsheet, and try again.

## 16. Share The Files

Share the main workbook only with people who need to edit inventory or process requests.

Share the view-only workbook with people who only need to view inventory.

Recommended access:

| Person | Main Workbook | View-Only Workbook |
| --- | --- | --- |
| Owner | Owner | Editor or owner |
| Staff updating inventory | Editor | Viewer or editor if needed |
| Request processors | Editor | Viewer or editor if needed |
| View-only users | No access or viewer | Viewer |

Do not give broad editor access to the main workbook unless that is intentional.

## 17. Backup Before Handing It Over

Before users start relying on the file:

1. Open the spreadsheet.
2. Click `File`.
3. Click `Make a copy`.
4. Save the copy in the backups folder.

Use a clear name with the date.

Example:

```text
Inventory & Request Management System - Backup - 2026-05-12
```

## 18. Quick Troubleshooting

| Problem | Check This First |
| --- | --- |
| Menu missing | Refresh the spreadsheet and confirm the script saved. |
| Authorization fails | Make sure you are using the owner account. |
| Function not found | Check that every `.gs` file was added. |
| Trigger behavior missing | Run `Initialize System` again and check Apps Script triggers. |
| Owner-only creation blocks owner | Use the owner diagnostic details and set `ownerTemporaryUserKeyOverride`. |
| View-only link opens old or wrong file | Check stored target status. Use owner reset only if you are sure. |
| Search results are wrong | Check whether headers or columns were moved. |
| Logs are empty | Confirm triggers exist and test an action that writes logs. |

## 19. Final Checklist

Before telling the team to use the file, confirm this list.

- [ ] All script files are present.
- [ ] `appsscript.json` is updated.
- [ ] Apps Script project is saved.
- [ ] Spreadsheet was refreshed after saving.
- [ ] Custom menu appears.
- [ ] `Initialize System` has been run by the owner.
- [ ] Authorization has been approved by the owner.
- [ ] `Run Verification` has been run.
- [ ] Main sheets exist.
- [ ] Log sheets exist.
- [ ] Triggers are installed.
- [ ] Encoder search works.
- [ ] Request form staging works.
- [ ] Temporary request storage works if your team will use it.
- [ ] View-only spreadsheet was created by the owner.
- [ ] Second view-only click opens the same file.
- [ ] View-only sharing is set correctly.
- [ ] A backup copy exists.
- [ ] Users know whether to use the main workbook or view-only workbook.
