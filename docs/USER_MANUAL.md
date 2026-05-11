# User Manual

This guide is for people using the spreadsheet day to day. It explains what each sheet is for, what the menu items do, and how to avoid the common mistakes that cause bad inventory data.

You do not need to edit Apps Script to use the system.

## 1. The Main Idea

The spreadsheet keeps inventory records and request records in one place.

Use the main workbook when you are allowed to edit inventory or process requests. Use the view-only workbook when you only need to look at inventory.

The main workbook has these important parts:

- `INVENTORY` for the item list.
- `ENCODER` for quick searching and stock actions.
- `REQUEST FORM` for staging and finalizing requests.
- Log sheets for checking what happened.
- A custom menu for setup, sync, and request actions.

## 2. Inventory Sheet

![Inventory sheet](screenshots/Inventory.png)

The `INVENTORY` sheet is the main source of item data.

Current columns:

| Column | Meaning |
| --- | --- |
| `NAME` | Item name. |
| `CODE` | Item code. |
| `QTY` | Current quantity. |
| `PRICE PER PIECE` | Item price. |
| `DATE INV.` | Inventory date or date-time. |
| `REMARKS` | Notes such as status or condition. |

Use simple and consistent entries. For quantity and price, use numbers. For items without a code, use the same no-code wording your team already uses, such as `NO CODE`.

Do not edit the title row, help row, or header row unless you are maintaining the system.

## 3. Encoder Sheet

![Encoder sheet](screenshots/Encoder.png)

The `ENCODER` sheet is for quick lookup and stock actions.

Typical use:

1. Open `ENCODER`.
2. Enter the requester or user name if your team requires it.
3. Search by code, name, or row.
4. Check the result carefully.
5. Follow the prompt if an action dialog appears.

Use code search when possible. It is less likely to match the wrong item than a name search.

## 4. Request Form Sheet

![Request form sheet](screenshots/Request_Form.png)

The `REQUEST FORM` sheet is for requests with one or more items.

Normal request flow:

1. Open `REQUEST FORM`.
2. Enter the requested-by name.
3. Search for an item by code, name, or row.
4. Mark the item to stage it.
5. Confirm the staged item appears in the staged area.
6. Repeat until all requested items are staged.
7. Use `Finalize Staged Request` from the custom menu.

Review the staged items before finalizing. Finalizing a request changes inventory quantities and writes logs.

## 5. Log Sheets

The log sheets are there so the owner or staff can check activity later.

![Change logs](screenshots/Change_Logs.png)

![Encoder logs](screenshots/Encoder_Logs.png)

![Request form logs](screenshots/Request_Form_Logs.png)

![Finalized request logs](screenshots/Finalized_Request_Logs.png)

What each log is for:

| Sheet | What It Records |
| --- | --- |
| `CHANGE LOGS` | Inventory-related changes. |
| `ENCODER LOGS` | Encoder searches and actions. |
| `REQUEST FORM LOGS` | Request form activity. |
| `FINALIZED REQUEST LOGS` | Items from completed requests. |

Do not manually edit logs during normal work. If a log looks wrong, tell the person maintaining the spreadsheet.

## 6. Custom Menu

The menu appears at the top of the spreadsheet after it loads:

```text
Inventory & Request Management System
```

Menu items:

| Item | Use It For |
| --- | --- |
| `Initialize System` | Setup or repair. Usually for the owner or maintainer. |
| `Run Verification` | Check whether the setup looks correct. |
| `Open/Create View-Only Spreadsheet` | Create the view-only workbook once or open the existing one. |
| `Sync View-Only Now` | Copy current inventory data to the linked view-only workbook. |
| `Finalize Staged Request` | Complete the request currently staged on `REQUEST FORM`. |
| `Store Request Temporarily` | Save an unfinished staged request. |
| `Pull Temporary Request` | Load a saved request. |
| `Delete Temporary Request` | Delete a saved temporary request. |
| `Clear Staged Request` | Clear the current staged items. |

If the menu is missing, refresh the spreadsheet and wait a few seconds.

## 7. Storing And Pulling Requests

Use `Store Request Temporarily` when a request is not ready to finish yet.

Example:

1. You stage several items.
2. You need approval before finalizing.
3. You store the request temporarily.
4. Later, you use `Pull Temporary Request` to bring it back.
5. You review it and finalize it.

Use `Delete Temporary Request` only when the saved request is no longer needed.

## 8. View-Only Spreadsheet

The view-only spreadsheet is a separate file. It is meant for people who should see inventory but should not edit the main workbook.

What to expect:

- The owner creates it once.
- Later menu clicks open the same file.
- If the browser blocks the new tab, a dialog shows a link you can click.
- If data looks old, someone with access to the main workbook can run `Sync View-Only Now`.

Do not treat the view-only spreadsheet as the source. The source is always the main workbook.

## 9. Entering Data Safely

Keep entries boring and consistent. That is what makes search and logs useful.

Good examples:

| Field | Good Entry |
| --- | --- |
| `NAME` | `Digital Valve Board` |
| `CODE` | `JD16343094` |
| `QTY` | `296` |
| `PRICE PER PIECE` | `9960.10` |
| `DATE INV.` | `07/22/2025 12:00:00 AM` |
| `REMARKS` | `FOR TESTING` |

Avoid:

- Typing words into quantity cells.
- Sorting only one column.
- Pasting over headers.
- Renaming sheets.
- Deleting rows without checking first.
- Editing hidden or protected support areas.

If you accidentally change the wrong data, stop and report it right away. Recent mistakes are easier to fix.

## 10. Searching Tips

Use the most specific search available.

| Search Type | When To Use |
| --- | --- |
| Code | Best when you know the item code. |
| Name | Useful when you only know part of the item name. |
| Row | Useful when someone gave you an exact row number. |

Before applying an action, check both the item name and code. Similar item names can lead to mistakes.

## 11. Exporting Or Printing

This is just normal Google Sheets behavior.

To download a copy:

1. Click `File`.
2. Click `Download`.
3. Choose PDF, Excel, or CSV.

To print:

1. Click `File`.
2. Click `Print`.
3. Choose the current sheet or selected range.

Downloaded Excel or CSV files will not run the Apps Script automation.

## 12. Common Problems

| Problem | What To Try |
| --- | --- |
| Menu is missing | Refresh the spreadsheet. |
| Script asks for authorization | Ask the owner if you are unsure. First-time authorization is normal. |
| Search finds nothing | Check spelling, try code search, or use fewer name words. |
| Request will not finalize | Check requested-by name, staged items, and available quantity. |
| New tab did not open | Click the fallback link in the dialog. |
| View-only file says no access | Ask the owner to share the view-only spreadsheet with you. |
| Log looks wrong | Do not edit it manually. Report it. |
| Sheet looks broken | Stop editing and ask the owner to run setup or check a backup. |

## 13. Short FAQ

### Can I edit the view-only spreadsheet?

No. Use the main workbook if you have permission to edit inventory or requests.

### Why did Google ask for permission?

The scripts need permission to work with the spreadsheet. If you are not sure, ask the owner before approving.

### Can I add new inventory rows?

Only if your team allows you to edit the source inventory.

### Can I undo a finalized request?

There is no normal undo button for a finalized request. Tell the owner or maintainer as soon as possible.

### Why did a request get blocked?

The most common reason is that the request would reduce stock below zero.

### What should I do before finalizing?

Check the requester name, item names, item codes, actions, and quantities.

### What if I am in the wrong Google account?

Switch to the correct account or use a separate browser profile. Multiple signed-in accounts can confuse Apps Script authorization.

### Who should I contact for help?

Use whatever support process your team has set for this spreadsheet. If none exists yet, ask the spreadsheet owner.
