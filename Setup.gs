/**
 * PROJECT-BASED INVENTORY MANAGEMENT SYSTEM
 * File: Setup.gs
 *
 * Run initializeSystem() once after pasting this file.
 */

const SPREADSHEET_ID = '1wZClxJOvHMOVNLkBMZyFxzK-cNZC9QcUHAU8BmQd5-k';

const SYSTEM_SHEETS = Object.freeze({
  DASHBOARD: 'Dashboard',
  USERS: 'Users',
  ROLES: 'Roles',
  PROJECT_MASTER: 'Project_Master',
  SKU_MASTER: 'SKU_Master',
  CATEGORY_MASTER: 'Category_Master',
  VENDOR_MASTER: 'Vendor_Master',
  FREE_INVENTORY: 'Free_Inventory',
  CUSTOMIZE_INVENTORY: 'Customize_Inventory',
  INWARD: 'Inward',
  OUTWARD: 'Outward',
  TRANSFER_REQUEST: 'Transfer_Request',
  APPROVAL: 'Approval',
  TRANSACTION_LOG: 'Transaction_Log',
  EMAIL_QUEUE: 'Email_Queue',
  SETTINGS: 'Settings'
});

const SHEET_HEADERS = Object.freeze({
  Dashboard: [
    'Metric',
    'Value',
    'Last Updated'
  ],

  Users: [
    'User ID',
    'Employee Name',
    'Email',
    'Mobile',
    'Department',
    'Role',
    'Status',
    'Created At',
    'Updated At'
  ],

  Roles: [
    'Role ID',
    'Role Name',
    'Can Inward',
    'Can Outward',
    'Can Transfer',
    'Can Approve',
    'Can Reports'
  ],

  Project_Master: [
    'Project ID',
    'Project Name',
    'Client Name',
    'Project Code',
    'Status',
    'HOD Email',
    'Start Date',
    'End Date'
  ],

  SKU_Master: [
    'SKU Code',
    'SKU Name',
    'Category',
    'Unit',
    'Default Rate',
    'Status'
  ],

  Category_Master: [
    'Category',
    'Status'
  ],

  Vendor_Master: [
    'Vendor ID',
    'Vendor Name',
    'GST',
    'Contact',
    'Email',
    'Mobile',
    'Status'
  ],

  Free_Inventory: [
    'Stock ID',
    'SKU Code',
    'SKU Name',
    'Category',
    'Available Qty',
    'Reserved Qty',
    'Unit',
    'Rate',
    'Value',
    'Last Movement',
    'Status'
  ],

  Customize_Inventory: [
    'Stock ID',
    'Project ID',
    'Project Name',
    'SKU Code',
    'SKU Name',
    'Category',
    'Qty',
    'Used Qty',
    'Balance Qty',
    'Unit',
    'Rate',
    'Value',
    'Last Movement Date',
    'Days Without Movement',
    'Eligible For Free',
    'Status'
  ],

  Inward: [
    'Inward No',
    'Date',
    'Inventory Type',
    'Vendor',
    'Project',
    'SKU',
    'Qty',
    'Unit',
    'Rate',
    'Amount',
    'Entered By'
  ],

  Outward: [
    'Outward No',
    'Date',
    'Inventory Type',
    'Source Project',
    'Destination Project',
    'SKU',
    'Qty',
    'Approval Required',
    'Approval Status',
    'Request ID',
    'Approved By',
    'Approval Date',
    'Outward By',
    'Outward Date',
    'Remarks'
  ],

  Transfer_Request: [
    'Request ID',
    'Request Date',
    'Source Project',
    'Destination Project',
    'SKU',
    'Qty',
    'Request Type',
    'Requested By',
    'HOD Email',
    'Approval Status',
    'Approval Token',
    'Approval Date',
    'Approved By',
    'Outward Status',
    'Outward No',
    'Remarks'
  ],

  Approval: [
    'Approval ID',
    'Request ID',
    'HOD Email',
    'Action',
    'Action Date',
    'Token',
    'Remarks',
    'IP Address',
    'User Agent'
  ],

  Transaction_Log: [
    'Log ID',
    'Date',
    'User',
    'Module',
    'Action',
    'Reference ID',
    'Details'
  ],

  Email_Queue: [
    'Queue ID',
    'Created At',
    'Email',
    'Subject',
    'Body',
    'Status',
    'Sent Time',
    'Error Message'
  ],

  Settings: [
    'Setting',
    'Value',
    'Description'
  ]
});

/**
 * Main setup function.
 * Run this function once.
 */
function initializeSystem() {
  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(30000);

    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);

    spreadsheet.setSpreadsheetTimeZone('Asia/Kolkata');

    createOrVerifySheets_(spreadsheet);
    insertDefaultRoles_(spreadsheet);
    insertDefaultSettings_(spreadsheet);
    insertDashboardMetrics_(spreadsheet);
    applyDataValidations_(spreadsheet);
    applyFormulas_(spreadsheet);
    applySheetFormatting_(spreadsheet);
    createSystemProperties_();

    SpreadsheetApp.flush();

    Logger.log('Inventory Management System initialized successfully.');

    return {
      success: true,
      message: 'Inventory Management System initialized successfully.',
      spreadsheetId: SPREADSHEET_ID,
      spreadsheetUrl: spreadsheet.getUrl()
    };

  } catch (error) {
    Logger.log(error.stack || error.message);

    throw new Error(
      'System initialization failed: ' +
      (error.message || String(error))
    );

  } finally {
    lock.releaseLock();
  }
}

/**
 * Creates missing sheets and verifies headers.
 * Existing data will not be deleted.
 */
function createOrVerifySheets_(spreadsheet) {
  Object.keys(SHEET_HEADERS).forEach(function (sheetName) {
    let sheet = spreadsheet.getSheetByName(sheetName);

    if (!sheet) {
      sheet = spreadsheet.insertSheet(sheetName);
    }

    const headers = SHEET_HEADERS[sheetName];

    if (!headers || headers.length === 0) {
      return;
    }

    ensureMinimumColumns_(sheet, headers.length);

    const currentHeaders = sheet
      .getRange(1, 1, 1, headers.length)
      .getDisplayValues()[0];

    const isHeaderEmpty = currentHeaders.every(function (value) {
      return String(value).trim() === '';
    });

    if (isHeaderEmpty) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      return;
    }

    headers.forEach(function (header, index) {
      const existingHeader = String(currentHeaders[index] || '').trim();

      if (!existingHeader) {
        sheet.getRange(1, index + 1).setValue(header);
      }
    });
  });
}

/**
 * Adds default role records.
 */
function insertDefaultRoles_(spreadsheet) {
  const sheet = spreadsheet.getSheetByName(SYSTEM_SHEETS.ROLES);

  if (!sheet) {
    throw new Error('Roles sheet not found.');
  }

  const existingRoleNames = getColumnValues_(sheet, 2, 2);

  const roles = [
    ['R001', 'ADMIN', 'YES', 'YES', 'YES', 'YES', 'YES'],
    ['R002', 'HOD', 'NO', 'NO', 'NO', 'YES', 'YES'],
    ['R003', 'DOER', 'YES', 'YES', 'YES', 'NO', 'YES'],
    ['R004', 'VIEWER', 'NO', 'NO', 'NO', 'NO', 'YES']
  ];

  const newRows = roles.filter(function (row) {
    return existingRoleNames.indexOf(row[1]) === -1;
  });

  if (newRows.length > 0) {
    sheet
      .getRange(sheet.getLastRow() + 1, 1, newRows.length, newRows[0].length)
      .setValues(newRows);
  }
}

/**
 * Adds default system settings.
 */
function insertDefaultSettings_(spreadsheet) {
  const sheet = spreadsheet.getSheetByName(SYSTEM_SHEETS.SETTINGS);

  if (!sheet) {
    throw new Error('Settings sheet not found.');
  }

  const existingSettings = getColumnValues_(sheet, 1, 2);

  const settings = [
    [
      'Company Name',
      '',
      'Company name displayed in the application and emails.'
    ],
    [
      'Company Email',
      '',
      'Main company email address.'
    ],
    [
      'Currency',
      'INR',
      'Currency used for inventory valuation.'
    ],
    [
      'GST',
      '18',
      'Default GST percentage.'
    ],
    [
      'Free Eligibility Days',
      '90',
      'Customized stock becomes eligible for Free Inventory review after this number of inactive days.'
    ],
    [
      'Time Zone',
      'Asia/Kolkata',
      'Application date and time zone.'
    ],
    [
      'Web App URL',
      '',
      'Paste the deployed Google Apps Script web app URL here.'
    ],
    [
      'Approval Link Expiry Days',
      '7',
      'Number of days an email approval link remains valid.'
    ],
    [
      'Allow Partial Outward',
      'YES',
      'Allow partial stock quantity to be outwarded.'
    ],
    [
      'System Status',
      'ACTIVE',
      'Controls whether the application is active.'
    ]
  ];

  const newRows = settings.filter(function (row) {
    return existingSettings.indexOf(row[0]) === -1;
  });

  if (newRows.length > 0) {
    sheet
      .getRange(sheet.getLastRow() + 1, 1, newRows.length, 3)
      .setValues(newRows);
  }
}

/**
 * Adds dashboard metric labels.
 */
function insertDashboardMetrics_(spreadsheet) {
  const sheet = spreadsheet.getSheetByName(SYSTEM_SHEETS.DASHBOARD);

  if (!sheet) {
    throw new Error('Dashboard sheet not found.');
  }

  const existingMetrics = getColumnValues_(sheet, 1, 2);

  const metrics = [
    ['Total Free Inventory SKUs', '', ''],
    ['Total Customized Inventory SKUs', '', ''],
    ['Total Free Inventory Value', '', ''],
    ['Total Customized Inventory Value', '', ''],
    ['Pending HOD Approvals', '', ''],
    ['Approved Pending Outward', '', ''],
    ['90-Day Eligible Stock', '', ''],
    ['Active Projects', '', ''],
    ['Active Users', '', ''],
    ['Last Dashboard Refresh', '', '']
  ];

  const newRows = metrics.filter(function (row) {
    return existingMetrics.indexOf(row[0]) === -1;
  });

  if (newRows.length > 0) {
    sheet
      .getRange(sheet.getLastRow() + 1, 1, newRows.length, 3)
      .setValues(newRows);
  }
}

/**
 * Applies dropdown validations.
 */
function applyDataValidations_(spreadsheet) {
  applyListValidation_(
    spreadsheet,
    SYSTEM_SHEETS.USERS,
    6,
    ['ADMIN', 'HOD', 'DOER', 'VIEWER']
  );

  applyListValidation_(
    spreadsheet,
    SYSTEM_SHEETS.USERS,
    7,
    ['ACTIVE', 'INACTIVE']
  );

  applyListValidation_(
    spreadsheet,
    SYSTEM_SHEETS.PROJECT_MASTER,
    5,
    ['ACTIVE', 'COMPLETED', 'ON HOLD', 'CANCELLED']
  );

  applyListValidation_(
    spreadsheet,
    SYSTEM_SHEETS.SKU_MASTER,
    6,
    ['ACTIVE', 'INACTIVE']
  );

  applyListValidation_(
    spreadsheet,
    SYSTEM_SHEETS.CATEGORY_MASTER,
    2,
    ['ACTIVE', 'INACTIVE']
  );

  applyListValidation_(
    spreadsheet,
    SYSTEM_SHEETS.VENDOR_MASTER,
    7,
    ['ACTIVE', 'INACTIVE']
  );

  applyListValidation_(
    spreadsheet,
    SYSTEM_SHEETS.INWARD,
    3,
    ['FREE', 'CUSTOMIZE']
  );

  applyListValidation_(
    spreadsheet,
    SYSTEM_SHEETS.OUTWARD,
    3,
    ['FREE', 'CUSTOMIZE']
  );

  applyListValidation_(
    spreadsheet,
    SYSTEM_SHEETS.OUTWARD,
    8,
    ['YES', 'NO']
  );

  applyListValidation_(
    spreadsheet,
    SYSTEM_SHEETS.OUTWARD,
    9,
    ['NOT REQUIRED', 'PENDING', 'APPROVED', 'REJECTED', 'EXPIRED']
  );

  applyListValidation_(
    spreadsheet,
    SYSTEM_SHEETS.TRANSFER_REQUEST,
    7,
    ['PROJECT TO PROJECT', 'CUSTOMIZE TO FREE']
  );

  applyListValidation_(
    spreadsheet,
    SYSTEM_SHEETS.TRANSFER_REQUEST,
    10,
    ['PENDING', 'APPROVED', 'REJECTED', 'EXPIRED']
  );

  applyListValidation_(
    spreadsheet,
    SYSTEM_SHEETS.TRANSFER_REQUEST,
    14,
    ['PENDING', 'COMPLETED', 'CANCELLED']
  );

  applyListValidation_(
    spreadsheet,
    SYSTEM_SHEETS.APPROVAL,
    4,
    ['APPROVED', 'REJECTED']
  );

  applyListValidation_(
    spreadsheet,
    SYSTEM_SHEETS.EMAIL_QUEUE,
    6,
    ['PENDING', 'SENT', 'FAILED']
  );

  applyListValidation_(
    spreadsheet,
    SYSTEM_SHEETS.FREE_INVENTORY,
    11,
    ['AVAILABLE', 'OUT OF STOCK', 'INACTIVE']
  );

  applyListValidation_(
    spreadsheet,
    SYSTEM_SHEETS.CUSTOMIZE_INVENTORY,
    15,
    ['YES', 'NO']
  );

  applyListValidation_(
    spreadsheet,
    SYSTEM_SHEETS.CUSTOMIZE_INVENTORY,
    16,
    [
      'AVAILABLE',
      'PARTIALLY USED',
      'CONSUMED',
      'TRANSFER PENDING',
      'APPROVED FOR TRANSFER',
      'CONVERTED TO FREE',
      'INACTIVE'
    ]
  );
}

/**
 * Applies formulas to calculated columns.
 */
function applyFormulas_(spreadsheet) {
  const maxRows = 5000;

  const freeInventorySheet =
    spreadsheet.getSheetByName(SYSTEM_SHEETS.FREE_INVENTORY);

  const customizeInventorySheet =
    spreadsheet.getSheetByName(SYSTEM_SHEETS.CUSTOMIZE_INVENTORY);

  const inwardSheet =
    spreadsheet.getSheetByName(SYSTEM_SHEETS.INWARD);

  ensureMinimumRows_(freeInventorySheet, maxRows);
  ensureMinimumRows_(customizeInventorySheet, maxRows);
  ensureMinimumRows_(inwardSheet, maxRows);

  freeInventorySheet
    .getRange('I2')
    .setFormula('=ARRAYFORMULA(IF(A2:A="","",E2:E*H2:H))');

  customizeInventorySheet
    .getRange('I2')
    .setFormula('=ARRAYFORMULA(IF(A2:A="","",G2:G-H2:H))');

  customizeInventorySheet
    .getRange('L2')
    .setFormula('=ARRAYFORMULA(IF(A2:A="","",I2:I*K2:K))');

  customizeInventorySheet
    .getRange('N2')
    .setFormula(
      '=ARRAYFORMULA(IF(A2:A="","",IF(M2:M="","",TODAY()-INT(M2:M))))'
    );

  customizeInventorySheet
    .getRange('O2')
    .setFormula(
      '=ARRAYFORMULA(IF(A2:A="","",IF((I2:I>0)*(N2:N>=90),"YES","NO")))'
    );

  inwardSheet
    .getRange('J2')
    .setFormula('=ARRAYFORMULA(IF(A2:A="","",G2:G*I2:I))');
}

/**
 * Applies basic professional formatting.
 */
function applySheetFormatting_(spreadsheet) {
  Object.keys(SHEET_HEADERS).forEach(function (sheetName) {
    const sheet = spreadsheet.getSheetByName(sheetName);

    if (!sheet) {
      return;
    }

    const headers = SHEET_HEADERS[sheetName];

    if (!headers || headers.length === 0) {
      return;
    }

    const headerRange = sheet.getRange(1, 1, 1, headers.length);

    headerRange
      .setFontWeight('bold')
      .setFontColor('#FFFFFF')
      .setBackground('#1F4E78')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle')
      .setWrap(true);

    sheet.setFrozenRows(1);
    sheet.setRowHeight(1, 42);

    for (let column = 1; column <= headers.length; column++) {
      sheet.setColumnWidth(column, 145);
    }

    sheet.getDataRange().setVerticalAlignment('middle');

    if (sheet.getFilter()) {
      sheet.getFilter().remove();
    }

    if (sheet.getLastRow() >= 1) {
      sheet
        .getRange(1, 1, Math.max(sheet.getMaxRows(), 2), headers.length)
        .createFilter();
    }

    applyDateFormats_(sheetName, sheet);
    applyNumberFormats_(sheetName, sheet);
  });
}

/**
 * Applies date formats to known date columns.
 */
function applyDateFormats_(sheetName, sheet) {
  const dateColumns = {
    Users: [8, 9],
    Project_Master: [7, 8],
    Free_Inventory: [10],
    Customize_Inventory: [13],
    Inward: [2],
    Outward: [2, 12, 14],
    Transfer_Request: [2, 12],
    Approval: [5],
    Transaction_Log: [2],
    Email_Queue: [2, 7],
    Dashboard: [3]
  };

  const columns = dateColumns[sheetName] || [];

  columns.forEach(function (column) {
    sheet
      .getRange(2, column, Math.max(sheet.getMaxRows() - 1, 1), 1)
      .setNumberFormat('dd-MMM-yyyy hh:mm:ss');
  });
}

/**
 * Applies number formats to quantity, rate and value columns.
 */
function applyNumberFormats_(sheetName, sheet) {
  const quantityColumns = {
    Free_Inventory: [5, 6],
    Customize_Inventory: [7, 8, 9],
    Inward: [7],
    Outward: [7],
    Transfer_Request: [6]
  };

  const currencyColumns = {
    SKU_Master: [5],
    Free_Inventory: [8, 9],
    Customize_Inventory: [11, 12],
    Inward: [9, 10]
  };

  (quantityColumns[sheetName] || []).forEach(function (column) {
    sheet
      .getRange(2, column, Math.max(sheet.getMaxRows() - 1, 1), 1)
      .setNumberFormat('0.00');
  });

  (currencyColumns[sheetName] || []).forEach(function (column) {
    sheet
      .getRange(2, column, Math.max(sheet.getMaxRows() - 1, 1), 1)
      .setNumberFormat('₹#,##0.00');
  });
}

/**
 * Saves base script properties.
 */
function createSystemProperties_() {
  PropertiesService.getScriptProperties().setProperties(
    {
      SPREADSHEET_ID: SPREADSHEET_ID,
      SYSTEM_VERSION: '1.0.0',
      SYSTEM_TIME_ZONE: 'Asia/Kolkata',
      SYSTEM_INITIALIZED: 'YES',
      SYSTEM_INITIALIZED_AT: new Date().toISOString()
    },
    false
  );
}

/**
 * Creates a list-based data validation.
 */
function applyListValidation_(
  spreadsheet,
  sheetName,
  columnNumber,
  values
) {
  const sheet = spreadsheet.getSheetByName(sheetName);

  if (!sheet) {
    return;
  }

  ensureMinimumRows_(sheet, 5000);

  const rule = SpreadsheetApp
    .newDataValidation()
    .requireValueInList(values, true)
    .setAllowInvalid(false)
    .build();

  sheet
    .getRange(2, columnNumber, sheet.getMaxRows() - 1, 1)
    .setDataValidation(rule);
}

/**
 * Reads non-empty values from a sheet column.
 */
function getColumnValues_(sheet, columnNumber, startRow) {
  const lastRow = sheet.getLastRow();

  if (lastRow < startRow) {
    return [];
  }

  return sheet
    .getRange(startRow, columnNumber, lastRow - startRow + 1, 1)
    .getDisplayValues()
    .flat()
    .map(function (value) {
      return String(value).trim();
    })
    .filter(function (value) {
      return value !== '';
    });
}

/**
 * Ensures the sheet has enough rows.
 */
function ensureMinimumRows_(sheet, requiredRows) {
  if (!sheet) {
    return;
  }

  const currentRows = sheet.getMaxRows();

  if (currentRows < requiredRows) {
    sheet.insertRowsAfter(currentRows, requiredRows - currentRows);
  }
}

/**
 * Ensures the sheet has enough columns.
 */
function ensureMinimumColumns_(sheet, requiredColumns) {
  const currentColumns = sheet.getMaxColumns();

  if (currentColumns < requiredColumns) {
    sheet.insertColumnsAfter(
      currentColumns,
      requiredColumns - currentColumns
    );
  }
}

/**
 * Optional test function.
 */
function testSpreadsheetConnection() {
  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);

  const result = {
    success: true,
    spreadsheetName: spreadsheet.getName(),
    spreadsheetId: spreadsheet.getId(),
    spreadsheetUrl: spreadsheet.getUrl(),
    sheets: spreadsheet
      .getSheets()
      .map(function (sheet) {
        return sheet.getName();
      })
  };

  Logger.log(JSON.stringify(result, null, 2));

  return result;
}
