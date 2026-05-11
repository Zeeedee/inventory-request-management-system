const APP_CONFIG = {
  version: '4.1.0',
  appName: 'Inventory & Request Management System',
  sheets: {
    inventory: 'INVENTORY',
    requestForm: 'REQUEST FORM',
    encoder: 'ENCODER',
    changeLogs: 'CHANGE LOGS',
    encoderLogs: 'ENCODER LOGS',
    requestLogs: 'REQUEST FORM LOGS',
    finalizedRequestLogs: 'FINALIZED REQUEST LOGS',
    portfolioNotes: 'PORTFOLIO NOTES',
    requestTempStore: '__REQUEST_TEMP_STORE'
  },
  properties: {
    inventoryVersion: 'APP_INVENTORY_VERSION',
    viewOnlySpreadsheetId: 'APP_VIEW_ONLY_SPREADSHEET_ID',
    viewOnlySyncEnabled: 'APP_VIEW_ONLY_SYNC_ENABLED',
    viewOnlyCreatedBy: 'APP_VIEW_ONLY_CREATED_BY',
    viewOnlyCreatedAt: 'APP_VIEW_ONLY_CREATED_AT',
    viewOnlyAllowRecreateIfMissing: 'APP_VIEW_ONLY_ALLOW_RECREATE_IF_MISSING'
  },
  metadataKeys: {
    inventoryItemKey: 'INV_ITEM_KEY'
  },
  labels: {
    requestedBy: 'REQUESTED BY',
    remarks: 'REMARKS',
    row: 'ROW'
  },
  actions: {
    ADD: 'ADD',
    SUBTRACT: 'SUBTRACT'
  },
  search: {
    maxResults: 50,
    textFinderFallbackLimit: 25
  },
  structures: {
    inventory: {
      sheetName: 'INVENTORY',
      headerRow: 4,
      dataStartRow: 5,
      columnCount: 6,
      dynamicHeaderWidth: true
    },
    requestForm: {
      sheetName: 'REQUEST FORM',
      headerRow: 8,
      dataStartRow: 9,
      columnCount: 8
    },
    encoderResults: {
      sheetName: 'ENCODER',
      headerRow: 8,
      dataStartRow: 9,
      columnCount: 7
    },
    encoderLogs: {
      sheetName: 'ENCODER LOGS',
      headerRow: 4,
      dataStartRow: 5,
      columnCount: 8
    },
    changeLogs: {
      sheetName: 'CHANGE LOGS',
      headerRow: 4,
      dataStartRow: 5,
      columnCount: 8
    },
    requestLogs: {
      sheetName: 'REQUEST FORM LOGS',
      headerRow: 4,
      dataStartRow: 5,
      columnCount: 9
    },
    finalizedRequestLogs: {
      sheetName: 'FINALIZED REQUEST LOGS',
      headerRow: 4,
      dataStartRow: 5,
      columnCount: 8
    }
  },
  templates: {
    inventory: {
      title: 'INVENTORY',
      helpText: 'Maintain current stock records here. Add or update items as needed.',
      displayHeaders: ['NAME', 'CODE', 'QTY', 'PRICE PER PIECE', 'DATE INV.', 'REMARKS'],
      additionalDisplayHeaders: [],
      columnWidths: {
        NAME: 280,
        CODE: 150,
        QTY: 80,
        PRICE: 100,
        DATE_INV: 120,
        REMARKS: 280
      },
      additionalColumnWidths: [],
      styles: {
        headerBg: '#7a1f1f',
        headerText: '#ffffff'
      }
    },
    viewOnly: {
      spreadsheetTitlePrefix: 'Inventory & Request Management System',
      spreadsheetTitleSuffix: 'View Only',
      inventoryNote: 'This sheet is a read-only view generated from the main inventory workbook.',
      logsNote: 'This sheet is a read-only report generated from the main inventory workbook.'
    }
  },
  viewOnly: {
    enabledByDefault: true,
    createOnceOnly: true,
    ownerOnlyCreate: true,
    ownerEmailOverride: '',
    ownerTemporaryUserKeyOverride: '',
    autoOpenAfterCreate: true,
    autoOpenExisting: true,
    allowRecreateIfMissing: false,
    showUrlFallbackAlert: true,
    creationLockWaitMs: 30000,
    exports: {
      inventory: {
        structureKey: 'inventory',
        targetSheetName: 'INVENTORY',
        frozenRows: 1,
        tabColor: '#7a1f1f',
        noteText: 'This sheet is a read-only view generated from the main inventory workbook.',
        columnWidths: [280, 150, 80, 100, 120, 280]
      }
    }
  },
  inventory: {
    titleRow: 1,
    helpRow: 2,
    headerRow: 4,
    startRow: 5,
    lastDataColumn: 6,
    dateTimeNumberFormat: 'MM/dd/yyyy hh:mm:ss AM/PM',
    visibleHeaders: ['NAME', 'CODE', 'QTY', 'PRICE PER PIECE', 'DATE INV.', 'REMARKS'],
    cols: {
      NAME: 1,
      CODE: 2,
      QTY: 3,
      PRICE: 4,
      DATE_INV: 5,
      REMARKS: 6
    },
    templateRowCount: 25,
    columnWidths: {
      NAME: 280,
      CODE: 150,
      QTY: 80,
      PRICE: 100,
      DATE_INV: 120,
      REMARKS: 280
    }
  },
  encoder: {
    cells: {
      REQUESTED_BY: 'B1',
      SEARCH_CODE: 'B3',
      SEARCH_NAME: 'B4',
      SEARCH_ROW: 'B5'
    },
    resultHeaderRow: 8,
    resultStartRow: 9,
    visibleHeaders: ['NAME', 'CODE', 'QTY', 'PRICE PER PIECE', 'DATE INV.', 'REMARKS', 'ROW'],
    visibleWidth: 7,
    resultWidth: 8,
    cols: {
      NAME: 1,
      CODE: 2,
      QTY: 3,
      PRICE: 4,
      DATE_INV: 5,
      REMARKS: 6,
      ROW: 7,
      ITEM_KEY: 8
    },
    resultLimit: 50,
    columnWidths: {
      NAME: 280,
      CODE: 150,
      QTY: 80,
      PRICE: 100,
      DATE_INV: 120,
      REMARKS: 280,
      ROW: 70
    },
    colors: {
      titleBg: '#f1f3f4',
      noteBg: '#fff8e1',
      headerBg: '#7a1f1f',
      headerText: '#ffffff',
      altRow1: '#ffffff',
      altRow2: '#f6f8fc',
      border: '#5f6368'
    },
    layout: {
      titleRange: 'C1:G1',
      workflowRange: 'C2:G2',
      searchHeaderRange: 'A2:B2',
      helperRange: 'A7:G7',
      resultsTitleRange: 'A6:G6',
      title: 'ENCODER',
      workflowText: 'Search -> Select Item -> Update Stock -> Logged',
      helperText: 'Search an item to update stock instantly.',
      searchHeaderText: 'Search Panel',
      resultsTitleText: 'Search Results'
    },
    messages: {
      noMatchingItem: 'No matching item found.'
    }
  },
  request: {
    cells: {
      REQUESTED_BY: 'B1',
      SEARCH_CODE: 'B3',
      SEARCH_NAME: 'B4',
      SEARCH_ROW: 'B5'
    },
    searchTable: {
      headerRow: 8,
      startRow: 9,
      visibleHeaders: ['NAME', 'CODE', 'QTY', 'PRICE PER PIECE', 'DATE INV.', 'REMARKS', 'ROW', 'STAGE (X)'],
      visibleWidth: 8,
      fullWidth: 9,
      maxResults: 50,
      cols: {
        NAME: 1,
        CODE: 2,
        QTY: 3,
        PRICE: 4,
        DATE_INV: 5,
        REMARKS: 6,
        ROW: 7,
        SELECT: 8,
        ITEM_KEY: 9
      }
    },
    separatorColumn: 10,
    itemsTable: {
      headerRow: 8,
      startRow: 9,
      visibleHeaders: ['ACTION', 'NAME', 'CODE', 'QTY', 'PRICE PER PIECE', 'REMARKS', 'ROW', 'REMOVE (X)'],
      visibleWidth: 8,
      fullWidth: 11,
      cols: {
        ACTION: 11,
        NAME: 12,
        CODE: 13,
        QTY: 14,
        PRICE: 15,
        REMARKS: 16,
        ROW: 17,
        REMOVE: 18,
        ITEM_ID: 19,
        SESSION_ID: 20,
        ITEM_KEY: 21
      }
    },
    tempStore: {
      headerRow: 1,
      headers: [
        'RECORD_ID',
        'REQUESTED_BY',
        'REQUESTED_BY_KEY',
        'STORED_AT',
        'STORED_BY',
        'ITEM_ID',
        'ACTION',
        'NAME',
        'CODE',
        'QTY',
        'PRICE PER PIECE',
        'REMARKS',
        'SOURCE_ROW',
        'ITEM_KEY'
      ],
      cols: {
        RECORD_ID: 1,
        REQUESTED_BY: 2,
        REQUESTED_BY_KEY: 3,
        STORED_AT: 4,
        STORED_BY: 5,
        ITEM_ID: 6,
        ACTION: 7,
        NAME: 8,
        CODE: 9,
        QTY: 10,
        PRICE: 11,
        REMARKS: 12,
        SOURCE_ROW: 13,
        ITEM_KEY: 14
      }
    },
    sessionStorage: {
      sessionPrefix: 'REQ_SESSION_V4::',
      contextPrefix: 'REQ_CONTEXT_V4::'
    },
    flags: {
      allowNegativeStock: false,
      autoClearFormOnSuccess: true
    },
    tempSelection: {
      previewItemCount: 3,
      maxPromptEntries: 9
    },
    layout: {
      formTitleRange: 'C1:H1',
      workflowRange: 'C2:H2',
      searchHeaderRange: 'A2:B2',
      searchSectionRange: 'A6:H6',
      searchHelperRange: 'A7:H7',
      stagedTitleRange: 'K1:R1',
      stagedWorkflowRange: 'K2:R2',
      stagedSectionRange: 'K6:R6',
      stagedHelperRange: 'K7:R7',
      formTitle: 'REQUEST FORM',
      workflowText: 'Search -> Stage Items -> Finalize Request -> Logged',
      searchHeaderText: 'Search Panel',
      searchSectionText: 'Search Results - Mark X to Stage',
      searchHelperText: 'Search items, then type X in STAGE to add them.',
      stagedTitleText: 'Staged Items',
      stagedWorkflowText: 'Review staged items, remove with X, then finalize.',
      stagedSectionText: 'Staged Items - Ready for Submission',
      stagedHelperText: 'Staged items appear here. Type X in REMOVE to take one out.'
    }
  },
  logs: {
    headers: ['Timestamp', 'Requested By', 'Action', 'Item Name', 'Code', 'Quantity', 'Remarks', 'Source'],
    layout: {
      titleRow: 1,
      helpRow: 2,
      spacerRow: 3,
      headerRow: 4,
      dataStartRow: 5
    },
    notes: {
      default: 'All completed actions are recorded automatically.',
      changeLogs: 'Inventory changes are recorded here as an audit trail.',
      encoderLogs: 'Encoder searches and stock updates are recorded automatically.',
      requestLogs: 'Request staging, storage, and finalization activity is recorded here.',
      finalizedRequestLogs: 'Completed request totals are stored here for review.'
    },
    colors: {
      headerBg: '#1f1f1f',
      headerText: '#ffffff',
      border: '#5f6368'
    },
    actions: {
      ADD: 'Add',
      SUBTRACT: 'Subtract',
      SEARCH_BY_CODE: 'Search By Code',
      SEARCH_BY_NAME: 'Search By Name',
      SEARCH_BY_ROW: 'Search By Row',
      STAGE_REQUEST_ITEM: 'Stage Request Item',
      REMOVE_REQUEST_ITEM: 'Remove Request Item',
      STORE_TEMP_REQUEST: 'Store Temporary Request',
      PULL_TEMP_REQUEST: 'Pull Temporary Request',
      DELETE_TEMP_REQUEST: 'Delete Temporary Request',
      FINALIZE_REQUEST: 'Finalize Request',
      MANUAL_EDIT: 'Manual Edit',
      SETUP: 'System Setup'
    },
    sources: {
      INVENTORY: 'Inventory',
      ENCODER: 'Encoder',
      REQUEST_FORM: 'Request Form',
      SYSTEM: 'System'
    }
  },
  system: {
    menu: {
      root: 'Inventory & Request Management System',
      initialize: 'Initialize System',
      runVerification: 'Run Verification',
      createViewOnly: 'Open/Create View-Only Spreadsheet',
      syncViewOnly: 'Sync View-Only Now',
      finalizeRequest: 'Finalize Staged Request',
      storeRequest: 'Store Request Temporarily',
      pullRequest: 'Pull Temporary Request',
      deleteStoredRequest: 'Delete Temporary Request',
      clearRequest: 'Clear Staged Request'
    },
    dialogs: {
      width: 420,
      height: 330,
      statePropertyPrefix: 'APP_DIALOG_STATE::',
      staleMs: 15000,
      heartbeatMs: 5000
    },
    colors: {
      titleBg: '#1f1f1f',
      titleText: '#ffffff',
      noteBg: '#fff8e1',
      separatorBg: '#e0e0e0',
      separatorBorder: '#6b7280',
      border: '#5f6368',
      rowAlt1: '#ffffff',
      rowAlt2: '#f6f8fc',
      sectionBg: '#e8eaed',
      sectionText: '#202124',
      successBg: '#e6f4ea',
      successText: '#137333',
      dangerBg: '#fce8e6',
      dangerText: '#a50e0e',
      accentBg: '#e8f0fe',
      accentText: '#174ea6'
    }
  },
  portfolio: {
    title: 'PORTFOLIO NOTES',
    subtitle: 'Project Snapshot',
    notes: [
      'Built with Google Apps Script and Google Sheets.',
      'Supports inventory search, request staging, finalized requests, and audit logs.',
      'Prevents invalid stock updates that would reduce available quantity below zero.',
      'Records important actions automatically across inventory, encoder, and request workflows.',
      'Designed for a simple internal inventory workflow with practical spreadsheet-based controls.'
    ]
  },
  triggerHandlers: {
    encoderEdit: 'handleEncoderEdit',
    requestEdit: 'REQ_handleEdit',
    inventoryEdit: 'LOG_handleInventoryEdit',
    spreadsheetChange: 'APP_handleSpreadsheetChange'
  },
  noCodeIndicators: ['NO CODE', 'N/A', '-', 'NONE', 'TBA', 'TBD', '']
};
