/**
 * PROJECT-BASED INVENTORY MANAGEMENT SYSTEM
 * File: Config.gs
 */

const APP_CONFIG = Object.freeze({
  APP_NAME:
    'Project Inventory Management System',

  VERSION:
    '1.0.0',

  SPREADSHEET_ID:
    '1wZClxJOvHMOVNLkBMZyFxzK-cNZC9QcUHAU8BmQd5-k',

  TIME_ZONE:
    'Asia/Kolkata',

  DATE_FORMAT:
    'dd-MMM-yyyy',

  DATE_TIME_FORMAT:
    'dd-MMM-yyyy hh:mm:ss',

  INVENTORY_TYPES:
    Object.freeze({
      FREE:
        'FREE',

      CUSTOMIZE:
        'CUSTOMIZE'
    }),

  USER_ROLES:
    Object.freeze({
      ADMIN:
        'ADMIN',

      HOD:
        'HOD',

      DOER:
        'DOER',

      VIEWER:
        'VIEWER'
    }),

  USER_STATUS:
    Object.freeze({
      ACTIVE:
        'ACTIVE',

      INACTIVE:
        'INACTIVE'
    }),

  APPROVAL_STATUS:
    Object.freeze({
      PENDING:
        'PENDING',

      APPROVED:
        'APPROVED',

      REJECTED:
        'REJECTED',

      EXECUTED:
        'EXECUTED',

      CANCELLED:
        'CANCELLED',

      EXPIRED:
        'EXPIRED'
    }),

  OUTWARD_STATUS:
    Object.freeze({
      PENDING:
        'PENDING',

      COMPLETED:
        'COMPLETED',

      CANCELLED:
        'CANCELLED'
    }),

  CONVERSION_STATUS:
    Object.freeze({
      PENDING:
        'PENDING',

      APPROVED:
        'APPROVED',

      REJECTED:
        'REJECTED',

      EXECUTED:
        'EXECUTED',

      CANCELLED:
        'CANCELLED'
    }),

  REQUEST_TYPES:
    Object.freeze({
      PROJECT_TO_PROJECT:
        'PROJECT TO PROJECT',

      CROSS_PROJECT_OUTWARD:
        'CROSS PROJECT OUTWARD',

      CUSTOMIZE_TO_FREE:
        'CUSTOMIZE TO FREE'
    }),

  EMAIL_STATUS:
    Object.freeze({
      PENDING:
        'PENDING',

      SENT:
        'SENT',

      FAILED:
        'FAILED'
    }),

  FREE_ELIGIBILITY_DAYS:
    90,

  SHEETS:
    Object.freeze({
      DASHBOARD:
        'Dashboard',

      USERS:
        'Users',

      ROLES:
        'Roles',

      PROJECT_MASTER:
        'Project_Master',

      SKU_MASTER:
        'SKU_Master',

      CATEGORY_MASTER:
        'Category_Master',

      VENDOR_MASTER:
        'Vendor_Master',

      FREE_INVENTORY:
        'Free_Inventory',

      CUSTOMIZE_INVENTORY:
        'Customize_Inventory',

      INWARD:
        'Inward',

      OUTWARD:
        'Outward',

      TRANSFER_REQUEST:
        'Transfer_Request',

      APPROVAL:
        'Approval',

      APPROVAL_REQUESTS:
        'Approval_Requests',

      CONVERSION_REQUESTS:
        'Conversion_Requests',

      TRANSACTION_LOG:
        'Transaction_Log',

      EMAIL_QUEUE:
        'Email_Queue',

      SETTINGS:
        'Settings'
    }),

  ID_PREFIX:
    Object.freeze({
      USER:
        'USR',

      PROJECT:
        'PRJ',

      VENDOR:
        'VND',

      STOCK:
        'STK',

      FREE_INVENTORY:
        'FIN',

      CUSTOMIZE_INVENTORY:
        'CIN',

      INWARD:
        'INW',

      OUTWARD:
        'OUT',

      REQUEST:
        'REQ',

      APPROVAL:
        'APR',

      CONVERSION:
        'CNV',

      LOG:
        'LOG',

      EMAIL_QUEUE:
        'EMQ'
    })
});


/**
 * Returns the connected spreadsheet.
 */
function getDatabase() {
  return SpreadsheetApp.openById(
    APP_CONFIG.SPREADSHEET_ID
  );
}


/**
 * Returns a sheet by configured name.
 */
function getSystemSheet(
  sheetName
) {
  if (!sheetName) {
    throw new Error(
      'Sheet name is required.'
    );
  }

  const spreadsheet =
    getDatabase();

  const sheet =
    spreadsheet.getSheetByName(
      sheetName
    );

  if (!sheet) {
    throw new Error(
      'Sheet not found: ' +
      sheetName
    );
  }

  return sheet;
}


/**
 * Returns a sheet using APP_CONFIG.SHEETS key.
 *
 * Example:
 * getSheetByKey('USERS')
 */
function getSheetByKey(
  sheetKey
) {
  const sheetName =
    APP_CONFIG.SHEETS[
      sheetKey
    ];

  if (!sheetName) {
    throw new Error(
      'Invalid sheet key: ' +
      sheetKey
    );
  }

  return getSystemSheet(
    sheetName
  );
}


/**
 * Returns true if a configured sheet exists.
 */
function systemSheetExists(
  sheetName
) {
  if (!sheetName) {
    return false;
  }

  return Boolean(
    getDatabase()
      .getSheetByName(
        sheetName
      )
  );
}


/**
 * Returns an existing sheet or creates it.
 */
function getOrCreateSystemSheet(
  sheetName
) {
  if (!sheetName) {
    throw new Error(
      'Sheet name is required.'
    );
  }

  const spreadsheet =
    getDatabase();

  let sheet =
    spreadsheet.getSheetByName(
      sheetName
    );

  if (!sheet) {
    sheet =
      spreadsheet.insertSheet(
        sheetName
      );
  }

  return sheet;
}


/**
 * Returns current date in configured time zone.
 */
function getCurrentDateText() {
  return Utilities.formatDate(
    new Date(),
    APP_CONFIG.TIME_ZONE,
    APP_CONFIG.DATE_FORMAT
  );
}


/**
 * Returns current date and time in configured time zone.
 */
function getCurrentDateTimeText() {
  return Utilities.formatDate(
    new Date(),
    APP_CONFIG.TIME_ZONE,
    APP_CONFIG.DATE_TIME_FORMAT
  );
}


/**
 * Returns a JavaScript Date object.
 */
function getCurrentDateTime() {
  return new Date();
}


/**
 * Reads a setting from the Settings sheet.
 */
function getSetting(
  settingName,
  defaultValue
) {
  if (!settingName) {
    return defaultValue;
  }

  const sheet =
    getSheetByKey(
      'SETTINGS'
    );

  const lastRow =
    sheet.getLastRow();

  if (lastRow < 2) {
    return defaultValue;
  }

  const lastColumn =
    Math.max(
      sheet.getLastColumn(),
      2
    );

  if (
    sheet.getMaxColumns() <
    lastColumn
  ) {
    sheet.insertColumnsAfter(
      sheet.getMaxColumns(),
      lastColumn -
        sheet.getMaxColumns()
    );
  }

  const values =
    sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        2
      )
      .getDisplayValues();

  const searchName =
    String(settingName)
      .trim()
      .toLowerCase();

  for (
    let index = 0;
    index < values.length;
    index++
  ) {
    const currentName =
      String(
        values[index][0]
      )
        .trim()
        .toLowerCase();

    if (
      currentName ===
      searchName
    ) {
      const value =
        values[index][1];

      return value !== ''
        ? value
        : defaultValue;
    }
  }

  return defaultValue;
}


/**
 * Updates or creates a setting.
 */
function setSetting(
  settingName,
  settingValue,
  description
) {
  if (!settingName) {
    throw new Error(
      'Setting name is required.'
    );
  }

  const lock =
    LockService.getScriptLock();

  try {
    lock.waitLock(30000);

    const sheet =
      getSheetByKey(
        'SETTINGS'
      );

    if (
      sheet.getMaxColumns() < 3
    ) {
      sheet.insertColumnsAfter(
        sheet.getMaxColumns(),
        3 -
          sheet.getMaxColumns()
      );
    }

    const lastRow =
      sheet.getLastRow();

    if (lastRow >= 2) {
      const values =
        sheet
          .getRange(
            2,
            1,
            lastRow - 1,
            1
          )
          .getDisplayValues();

      const searchName =
        String(settingName)
          .trim()
          .toLowerCase();

      for (
        let index = 0;
        index <
          values.length;
        index++
      ) {
        const currentName =
          String(
            values[index][0]
          )
            .trim()
            .toLowerCase();

        if (
          currentName ===
          searchName
        ) {
          const rowNumber =
            index + 2;

          sheet
            .getRange(
              rowNumber,
              2
            )
            .setValue(
              settingValue
            );

          if (
            typeof description !==
            'undefined'
          ) {
            sheet
              .getRange(
                rowNumber,
                3
              )
              .setValue(
                description
              );
          }

          return {
            success:
              true,

            action:
              'UPDATED',

            rowNumber:
              rowNumber
          };
        }
      }
    }

    const targetRow =
      Math.max(
        sheet.getLastRow() + 1,
        2
      );

    if (
      targetRow >
      sheet.getMaxRows()
    ) {
      sheet.insertRowsAfter(
        sheet.getMaxRows(),
        targetRow -
          sheet.getMaxRows()
      );
    }

    sheet
      .getRange(
        targetRow,
        1,
        1,
        3
      )
      .setValues([
        [
          settingName,
          settingValue,
          description || ''
        ]
      ]);

    return {
      success:
        true,

      action:
        'CREATED',

      rowNumber:
        targetRow
    };

  } finally {
    lock.releaseLock();
  }
}


/**
 * Returns the configured 90-day eligibility value.
 */
function getFreeEligibilityDays() {
  const settingValue =
    Number(
      getSetting(
        'Free Eligibility Days',
        APP_CONFIG
          .FREE_ELIGIBILITY_DAYS
      )
    );

  if (
    !Number.isFinite(
      settingValue
    ) ||
    settingValue < 1
  ) {
    return APP_CONFIG
      .FREE_ELIGIBILITY_DAYS;
  }

  return settingValue;
}


/**
 * Returns web app URL from Settings or Script Service.
 */
function getWebAppUrl() {
  const savedUrl =
    String(
      getSetting(
        'Web App URL',
        ''
      )
    ).trim();

  if (savedUrl) {
    return savedUrl;
  }

  try {
    return (
      ScriptApp
        .getService()
        .getUrl() ||
      ''
    );
  } catch (error) {
    return '';
  }
}


/**
 * Verifies all configured sheets.
 */
function testConfiguredSheets() {
  const spreadsheet =
    getDatabase();

  const result = {};

  Object.keys(
    APP_CONFIG.SHEETS
  ).forEach(function (key) {
    const sheetName =
      APP_CONFIG.SHEETS[key];

    const sheet =
      spreadsheet.getSheetByName(
        sheetName
      );

    result[key] = {
      sheetName:
        sheetName,

      exists:
        Boolean(sheet),

      lastRow:
        sheet
          ? sheet.getLastRow()
          : 0,

      lastColumn:
        sheet
          ? sheet.getLastColumn()
          : 0,

      maxRows:
        sheet
          ? sheet.getMaxRows()
          : 0,

      maxColumns:
        sheet
          ? sheet.getMaxColumns()
          : 0
    };
  });

  Logger.log(
    JSON.stringify(
      result,
      null,
      2
    )
  );

  return result;
}


/**
 * Basic configuration test.
 */
function testConfig() {
  const spreadsheet =
    getDatabase();

  const result = {
    success:
      true,

    appName:
      APP_CONFIG.APP_NAME,

    version:
      APP_CONFIG.VERSION,

    spreadsheetId:
      spreadsheet.getId(),

    spreadsheetName:
      spreadsheet.getName(),

    timeZone:
      APP_CONFIG.TIME_ZONE,

    freeEligibilityDays:
      getFreeEligibilityDays(),

    webAppUrl:
      getWebAppUrl(),

    currentDateTime:
      getCurrentDateTimeText(),

    approvalRequestsSheet:
      APP_CONFIG.SHEETS
        .APPROVAL_REQUESTS,

    conversionRequestsSheet:
      APP_CONFIG.SHEETS
        .CONVERSION_REQUESTS
  };

  Logger.log(
    JSON.stringify(
      result,
      null,
      2
    )
  );

  return result;
}
