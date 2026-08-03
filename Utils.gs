/**
 * PROJECT-BASED INVENTORY MANAGEMENT SYSTEM
 * File: Utils.gs
 *
 * Common utility functions used by all modules.
 */

/**
 * Generates a sequential ID.
 *
 * Example:
 * generateNextId_('INWARD', 'INW') => INW000001
 */
function generateNextId_(counterKey, prefix) {
  if (!counterKey) {
    throw new Error('Counter key is required.');
  }

  if (!prefix) {
    throw new Error('ID prefix is required.');
  }

  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(30000);

    const properties = PropertiesService.getScriptProperties();
    const propertyName = 'COUNTER_' + String(counterKey).toUpperCase();

    let currentNumber = Number(properties.getProperty(propertyName) || 0);

    if (!Number.isFinite(currentNumber) || currentNumber < 0) {
      currentNumber = 0;
    }

    const nextNumber = currentNumber + 1;

    properties.setProperty(propertyName, String(nextNumber));

    return String(prefix).toUpperCase() +
      String(nextNumber).padStart(6, '0');

  } finally {
    lock.releaseLock();
  }
}

/**
 * Generates a unique token.
 */
function generateSecureToken_() {
  return Utilities.getUuid().replace(/-/g, '') +
    Utilities.getUuid().replace(/-/g, '');
}

/**
 * Generates a normal UUID.
 */
function generateUuid_() {
  return Utilities.getUuid();
}

/**
 * Returns a normalized string.
 */
function normalizeText_(value) {
  return String(value == null ? '' : value).trim();
}

/**
 * Returns uppercase normalized text.
 */
function normalizeUpper_(value) {
  return normalizeText_(value).toUpperCase();
}

/**
 * Returns lowercase normalized text.
 */
function normalizeLower_(value) {
  return normalizeText_(value).toLowerCase();
}

/**
 * Checks whether a value is blank.
 */
function isBlank_(value) {
  return normalizeText_(value) === '';
}

/**
 * Converts a value to a safe finite number.
 */
function toNumber_(value, defaultValue) {
  const numberValue = Number(value);

  if (!Number.isFinite(numberValue)) {
    return typeof defaultValue === 'number' ? defaultValue : 0;
  }

  return numberValue;
}

/**
 * Converts a value to a positive number.
 */
function toPositiveNumber_(value, fieldName) {
  const numberValue = Number(value);

  if (!Number.isFinite(numberValue) || numberValue <= 0) {
    throw new Error(
      (fieldName || 'Value') + ' must be greater than zero.'
    );
  }

  return numberValue;
}

/**
 * Rounds a number to two decimal places.
 */
function roundTwo_(value) {
  return Math.round((toNumber_(value, 0) + Number.EPSILON) * 100) / 100;
}

/**
 * Returns YES or NO.
 */
function toYesNo_(value) {
  const normalized = normalizeUpper_(value);

  return ['YES', 'Y', 'TRUE', '1'].indexOf(normalized) !== -1
    ? 'YES'
    : 'NO';
}

/**
 * Returns true for common truthy values.
 */
function toBoolean_(value) {
  return toYesNo_(value) === 'YES';
}

/**
 * Validates email address.
 */
function isValidEmail_(email) {
  const value = normalizeLower_(email);

  if (!value) {
    return false;
  }

  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(value);
}

/**
 * Validates Indian mobile number.
 */
function isValidMobile_(mobile) {
  const value = normalizeText_(mobile).replace(/\D/g, '');

  return /^[6-9]\d{9}$/.test(value);
}

/**
 * Validates Indian GST number.
 */
function isValidGst_(gstNumber) {
  const value = normalizeUpper_(gstNumber);

  if (!value) {
    return true;
  }

  return /^[0-9]{2}[A-Z]{5}[0-9]{4}[A-Z][1-9A-Z]Z[0-9A-Z]$/.test(value);
}

/**
 * Validates PAN number.
 */
function isValidPan_(panNumber) {
  const value = normalizeUpper_(panNumber);

  if (!value) {
    return true;
  }

  return /^[A-Z]{5}[0-9]{4}[A-Z]$/.test(value);
}

/**
 * Formats a date.
 */
function formatDate_(dateValue, format) {
  if (!dateValue) {
    return '';
  }

  const date = dateValue instanceof Date
    ? dateValue
    : new Date(dateValue);

  if (isNaN(date.getTime())) {
    return '';
  }

  return Utilities.formatDate(
    date,
    APP_CONFIG.TIME_ZONE,
    format || APP_CONFIG.DATE_FORMAT
  );
}

/**
 * Formats date and time.
 */
function formatDateTime_(dateValue) {
  return formatDate_(
    dateValue,
    APP_CONFIG.DATE_TIME_FORMAT
  );
}

/**
 * Converts input to a Date object.
 */
function parseDate_(value) {
  if (!value) {
    return null;
  }

  if (value instanceof Date && !isNaN(value.getTime())) {
    return value;
  }

  const date = new Date(value);

  return isNaN(date.getTime()) ? null : date;
}

/**
 * Returns number of whole days between two dates.
 */
function daysBetween_(startDate, endDate) {
  const start = parseDate_(startDate);
  const end = parseDate_(endDate || new Date());

  if (!start || !end) {
    return 0;
  }

  const startUtc = Date.UTC(
    start.getFullYear(),
    start.getMonth(),
    start.getDate()
  );

  const endUtc = Date.UTC(
    end.getFullYear(),
    end.getMonth(),
    end.getDate()
  );

  return Math.floor((endUtc - startUtc) / 86400000);
}

/**
 * Formats currency in INR.
 */
function formatCurrency_(value) {
  const amount = toNumber_(value, 0);

  return '₹' + amount.toLocaleString('en-IN', {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2
  });
}

/**
 * Escapes HTML.
 */
function escapeHtml_(value) {
  return normalizeText_(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

/**
 * Returns all rows from a configured sheet as objects.
 */
function getSheetObjects_(sheetName) {
  const sheet = getSystemSheet(sheetName);
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();

  if (lastRow < 2 || lastColumn < 1) {
    return [];
  }

  const values = sheet
    .getRange(1, 1, lastRow, lastColumn)
    .getValues();

  const headers = values[0].map(function (header) {
    return normalizeText_(header);
  });

  return values.slice(1)
    .map(function (row, index) {
      return {
        row: row,
        rowNumber: index + 2
      };
    })
    .filter(function (item) {
      return item.row.some(function (cell) {
        return !isBlank_(cell);
      });
    })
    .map(function (item) {
      const record = {
        _rowNumber: item.rowNumber
      };

      headers.forEach(function (header, columnIndex) {
        if (header) {
          record[header] = item.row[columnIndex];
        }
      });

      return record;
    });
}

/**
 * Returns one row as an object.
 */
function getRowObject_(sheetName, rowNumber) {
  const sheet = getSystemSheet(sheetName);

  if (!rowNumber || rowNumber < 2) {
    return null;
  }

  const lastColumn = sheet.getLastColumn();

  const headers = sheet
    .getRange(1, 1, 1, lastColumn)
    .getDisplayValues()[0]
    .map(normalizeText_);

  const values = sheet
    .getRange(rowNumber, 1, 1, lastColumn)
    .getValues()[0];

  const record = {
    _rowNumber: rowNumber
  };

  headers.forEach(function (header, index) {
    if (header) {
      record[header] = values[index];
    }
  });

  return record;
}

/**
 * Finds first matching row number.
 */
function findRowNumber_(
  sheetName,
  columnHeader,
  searchValue,
  caseSensitive
) {
  const sheet = getSystemSheet(sheetName);
  const headerMap = getHeaderMap_(sheet);

  const columnNumber = headerMap[columnHeader];

  if (!columnNumber) {
    throw new Error(
      'Header not found in ' + sheetName + ': ' + columnHeader
    );
  }

  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return -1;
  }

  const values = sheet
    .getRange(2, columnNumber, lastRow - 1, 1)
    .getDisplayValues()
    .flat();

  const target = caseSensitive
    ? normalizeText_(searchValue)
    : normalizeLower_(searchValue);

  for (let index = 0; index < values.length; index++) {
    const current = caseSensitive
      ? normalizeText_(values[index])
      : normalizeLower_(values[index]);

    if (current === target) {
      return index + 2;
    }
  }

  return -1;
}

/**
 * Finds all matching row numbers.
 */
function findAllRowNumbers_(
  sheetName,
  columnHeader,
  searchValue,
  caseSensitive
) {
  const sheet = getSystemSheet(sheetName);
  const headerMap = getHeaderMap_(sheet);

  const columnNumber = headerMap[columnHeader];

  if (!columnNumber) {
    throw new Error(
      'Header not found in ' + sheetName + ': ' + columnHeader
    );
  }

  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return [];
  }

  const values = sheet
    .getRange(2, columnNumber, lastRow - 1, 1)
    .getDisplayValues()
    .flat();

  const target = caseSensitive
    ? normalizeText_(searchValue)
    : normalizeLower_(searchValue);

  const rows = [];

  values.forEach(function (value, index) {
    const current = caseSensitive
      ? normalizeText_(value)
      : normalizeLower_(value);

    if (current === target) {
      rows.push(index + 2);
    }
  });

  return rows;
}

/**
 * Returns a map of header names to column numbers.
 */
function getHeaderMap_(sheet) {
  const lastColumn = sheet.getLastColumn();

  if (lastColumn < 1) {
    return {};
  }

  const headers = sheet
    .getRange(1, 1, 1, lastColumn)
    .getDisplayValues()[0];

  const map = {};

  headers.forEach(function (header, index) {
    const cleanHeader = normalizeText_(header);

    if (cleanHeader) {
      map[cleanHeader] = index + 1;
    }
  });

  return map;
}

/**
 * Appends an object to a sheet based on header names.
 */
function appendObjectRow_(
  sheetName,
  objectData
) {
  if (
    !objectData ||
    typeof objectData !== 'object'
  ) {
    throw new Error(
      'Row data is required.'
    );
  }

  const sheet =
    getSystemSheet(sheetName);

  const lock =
    LockService.getScriptLock();

  try {
    lock.waitLock(30000);

    let lastColumn =
      sheet.getLastColumn();

    /*
     * A sheet with no detected content can return
     * zero columns. Create enough columns before
     * attempting any range operation.
     */
    if (lastColumn < 1) {
      const objectHeaders =
        Object.keys(objectData);

      const requiredColumns =
        Math.max(
          objectHeaders.length,
          1
        );

      if (
        sheet.getMaxColumns() <
        requiredColumns
      ) {
        sheet.insertColumnsAfter(
          sheet.getMaxColumns(),
          requiredColumns -
            sheet.getMaxColumns()
        );
      }

      sheet
        .getRange(
          1,
          1,
          1,
          requiredColumns
        )
        .setValues([
          objectHeaders
        ]);

      lastColumn =
        requiredColumns;
    }

    /*
     * Protect against a header row extending beyond
     * the sheet's current column dimensions.
     */
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

    const headers =
      sheet
        .getRange(
          1,
          1,
          1,
          lastColumn
        )
        .getDisplayValues()[0]
        .map(normalizeText_);

    if (
      headers.filter(Boolean).length === 0
    ) {
      throw new Error(
        'No headers found in sheet: ' +
        sheetName
      );
    }

    const row =
      headers.map(
        function (header) {
          return Object.prototype
            .hasOwnProperty.call(
              objectData,
              header
            )
            ? objectData[header]
            : '';
        }
      );

    const rowNumber =
      sheet.getLastRow() + 1;

    /*
     * Main fix:
     * If the next row is outside the current sheet
     * dimensions, insert the required rows first.
     */
    if (
      rowNumber >
      sheet.getMaxRows()
    ) {
      sheet.insertRowsAfter(
        sheet.getMaxRows(),
        rowNumber -
          sheet.getMaxRows()
      );
    }

    sheet
      .getRange(
        rowNumber,
        1,
        1,
        row.length
      )
      .setValues([row]);

    return rowNumber;

  } finally {
    lock.releaseLock();
  }
}

/**
 * Updates selected fields in an existing row.
 */
function updateObjectRow_(
  sheetName,
  rowNumber,
  objectData
) {
  if (
    !rowNumber ||
    rowNumber < 2
  ) {
    throw new Error(
      'Valid row number is required.'
    );
  }

  if (
    !objectData ||
    typeof objectData !== 'object'
  ) {
    throw new Error(
      'Update data is required.'
    );
  }

  const sheet =
    getSystemSheet(sheetName);

  const lock =
    LockService.getScriptLock();

  try {
    lock.waitLock(30000);

    /*
     * Ensure the requested row exists before getRange().
     */
    if (
      rowNumber >
      sheet.getMaxRows()
    ) {
      sheet.insertRowsAfter(
        sheet.getMaxRows(),
        rowNumber -
          sheet.getMaxRows()
      );
    }

    const headerMap =
      getHeaderMap_(sheet);

    Object.keys(objectData)
      .forEach(function (header) {
        const columnNumber =
          headerMap[header];

        if (!columnNumber) {
          return;
        }

        /*
         * Ensure the requested column exists before
         * writing to it.
         */
        if (
          columnNumber >
          sheet.getMaxColumns()
        ) {
          sheet.insertColumnsAfter(
            sheet.getMaxColumns(),
            columnNumber -
              sheet.getMaxColumns()
          );
        }

        sheet
          .getRange(
            rowNumber,
            columnNumber
          )
          .setValue(
            objectData[header]
          );
      });

    return true;

  } finally {
    lock.releaseLock();
  }
}

/**
 * Deletes a row safely.
 */
function deleteSheetRow_(sheetName, rowNumber) {
  if (!rowNumber || rowNumber < 2) {
    throw new Error('Valid row number is required.');
  }

  const sheet = getSystemSheet(sheetName);
  sheet.deleteRow(rowNumber);

  return true;
}

/**
 * Checks whether a value already exists.
 */
function valueExists_(
  sheetName,
  columnHeader,
  searchValue,
  caseSensitive
) {
  return findRowNumber_(
    sheetName,
    columnHeader,
    searchValue,
    caseSensitive
  ) !== -1;
}

/**
 * Returns a unique list from a sheet column.
 */
function getUniqueColumnValues_(
  sheetName,
  columnHeader,
  includeBlank
) {
  const sheet = getSystemSheet(sheetName);
  const headerMap = getHeaderMap_(sheet);

  const columnNumber = headerMap[columnHeader];

  if (!columnNumber) {
    throw new Error(
      'Header not found in ' + sheetName + ': ' + columnHeader
    );
  }

  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return [];
  }

  const values = sheet
    .getRange(2, columnNumber, lastRow - 1, 1)
    .getDisplayValues()
    .flat()
    .map(normalizeText_);

  const uniqueValues = Array.from(new Set(values));

  return includeBlank
    ? uniqueValues
    : uniqueValues.filter(function (value) {
        return value !== '';
      });
}

/**
 * Logs an application transaction.
 */
/**
 * Adds one record to Transaction_Log.
 *
 * Transaction_Log columns:
 * A - Log ID
 * B - Date
 * C - User
 * D - Module
 * E - Action
 * F - Reference ID
 * G - Details
 */



/**
 * Generates the next Transaction Log ID.
 */


/**
 * Returns the active user's email where available.
 */
function getCurrentUserEmail_() {
  try {
    const activeUserEmail =
      Session.getActiveUser().getEmail();

    if (activeUserEmail) {
      return normalizeLower_(activeUserEmail);
    }

    const effectiveUserEmail =
      Session.getEffectiveUser().getEmail();

    return normalizeLower_(effectiveUserEmail);

  } catch (error) {
    return '';
  }
}

/**
 * Throws an error if required fields are blank.
 */
function validateRequiredFields_(data, requiredFields) {
  const missingFields = [];

  requiredFields.forEach(function (fieldName) {
    if (
      !Object.prototype.hasOwnProperty.call(data, fieldName) ||
      isBlank_(data[fieldName])
    ) {
      missingFields.push(fieldName);
    }
  });

  if (missingFields.length > 0) {
    throw new Error(
      'Required fields missing: ' + missingFields.join(', ')
    );
  }

  return true;
}

/**
 * Returns a standardized success response.
 */
function successResponse_(message, data) {
  return {
    success: true,
    message: message || 'Operation completed successfully.',
    data: typeof data === 'undefined' ? null : data
  };
}

/**
 * Returns a standardized error response.
 */
function errorResponse_(error, defaultMessage) {
  const message = error && error.message
    ? error.message
    : defaultMessage || 'Something went wrong.';

  return {
    success: false,
    message: message,
    data: null
  };
}

/**
 * Executes a function with error handling.
 */
function safeExecute_(callback, defaultMessage) {
  try {
    return callback();
  } catch (error) {
    console.error(error.stack || error.message || error);

    return errorResponse_(error, defaultMessage);
  }
}

/**
 * Sorts objects by a property.
 */
function sortObjects_(
  records,
  propertyName,
  descending
) {
  return records.sort(function (first, second) {
    const firstValue = first[propertyName];
    const secondValue = second[propertyName];

    if (firstValue === secondValue) {
      return 0;
    }

    if (descending) {
      return firstValue < secondValue ? 1 : -1;
    }

    return firstValue > secondValue ? 1 : -1;
  });
}

/**
 * Returns pagination data.
 */
function paginateRecords_(
  records,
  pageNumber,
  pageSize
) {
  const safePageSize = Math.max(
    1,
    Math.min(toNumber_(pageSize, 20), 500)
  );

  const totalRecords = records.length;
  const totalPages = Math.max(
    1,
    Math.ceil(totalRecords / safePageSize)
  );

  const safePageNumber = Math.max(
    1,
    Math.min(toNumber_(pageNumber, 1), totalPages)
  );

  const startIndex =
    (safePageNumber - 1) * safePageSize;

  return {
    records: records.slice(
      startIndex,
      startIndex + safePageSize
    ),
    pagination: {
      pageNumber: safePageNumber,
      pageSize: safePageSize,
      totalRecords: totalRecords,
      totalPages: totalPages
    }
  };
}

/**
 * Test function for Utils.gs.
 */
function testUtils() {
  const testResult = {
    generatedId: generateNextId_(
      'UTILS_TEST',
      'TST'
    ),
    uuid: generateUuid_(),
    secureTokenLength: generateSecureToken_().length,
    normalizedText: normalizeText_('  Inventory  '),
    normalizedUpper: normalizeUpper_('free'),
    numberValue: toNumber_('125.50'),
    roundedValue: roundTwo_(125.567),
    validEmail: isValidEmail_('test@example.com'),
    validMobile: isValidMobile_('9876543210'),
    validGst: isValidGst_('07ABCDE1234F1Z5'),
    formattedCurrency: formatCurrency_(125000.50),
    currentUserEmail: getCurrentUserEmail_(),
    currentDateTime: formatDateTime_(new Date()),
    freeEligibilityDays: getFreeEligibilityDays()
  };

  Logger.log(JSON.stringify(testResult, null, 2));

  return testResult;
}





/**
 * Adds one log record strictly in Transaction_Log columns A:G.
 *
 * A: Log ID
 * B: Date
 * C: User
 * D: Module
 * E: Action
 * F: Reference ID
 * G: Details
 */



/**
 * Returns next Transaction Log ID.
 * This function only reads column A.
 */



/**
 * Tests Transaction_Log without changing inventory.
 */

/**
 * Ensures Transaction_Log has the final A:G structure.
 *
 * Existing A:G log data is preserved.
 * All data validations are removed.
 * Extra columns after G are deleted.
 *
 * Run once:
 * initializeTransactionLogSheet
 */
function initializeTransactionLogSheet() {
  const sheet = getSystemSheet('Transaction_Log');

  const headers = [
    'Log ID',
    'Date',
    'User',
    'Module',
    'Action',
    'Reference ID',
    'Details'
  ];

  if (sheet.getMaxColumns() < 7) {
    sheet.insertColumnsAfter(
      sheet.getMaxColumns(),
      7 - sheet.getMaxColumns()
    );
  }

  sheet
    .getRange(1, 1, 1, 7)
    .setValues([headers])
    .setFontWeight('bold');

  sheet
    .getRange(
      1,
      1,
      sheet.getMaxRows(),
      sheet.getMaxColumns()
    )
    .clearDataValidations();

  const maxColumns = sheet.getMaxColumns();

  if (maxColumns > 7) {
    sheet.deleteColumns(
      8,
      maxColumns - 7
    );
  }

  sheet
    .getRange(
      2,
      2,
      Math.max(sheet.getMaxRows() - 1, 1),
      1
    )
    .setNumberFormat(
      'dd-mmm-yyyy hh:mm AM/PM'
    );

  sheet.setFrozenRows(1);

  SpreadsheetApp.flush();

  Logger.log(
    'Transaction_Log initialized successfully with columns A:G.'
  );

  return true;
}


/**
 * Adds one record strictly to Transaction_Log columns A:G.
 */
function addTransactionLog_(
  userEmail,
  moduleName,
  actionName,
  referenceId,
  details
) {
  const sheet = getSystemSheet('Transaction_Log');
  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(30000);

    const nextLogId =
      getNextTransactionLogId_(sheet);

    let detailsText = '';

    if (
      details !== null &&
      typeof details !== 'undefined'
    ) {
      detailsText =
        typeof details === 'string'
          ? details
          : JSON.stringify(details);
    }

    const targetRow = Math.max(
      sheet.getLastRow() + 1,
      2
    );

    const rowData = [
      nextLogId,
      new Date(),
      normalizeLower_(userEmail),
      normalizeUpper_(moduleName),
      normalizeUpper_(actionName),
      normalizeText_(referenceId),
      detailsText
    ];

    sheet
      .getRange(targetRow, 1, 1, 7)
      .clearDataValidations()
      .setValues([rowData]);

    return nextLogId;

  } finally {
    lock.releaseLock();
  }
}


/**
 * Returns the next sequential Transaction Log ID.
 */
function getNextTransactionLogId_(sheet) {
  const logSheet =
    sheet ||
    getSystemSheet('Transaction_Log');

  const lastRow = logSheet.getLastRow();
  let largestNumber = 0;

  if (lastRow >= 2) {
    const logIds = logSheet
      .getRange(2, 1, lastRow - 1, 1)
      .getDisplayValues()
      .flat();

    logIds.forEach(function (logId) {
      const match = normalizeUpper_(logId)
        .match(/^LOG(\d+)$/);

      if (match) {
        largestNumber = Math.max(
          largestNumber,
          Number(match[1]) || 0
        );
      }
    });
  }

  return (
    'LOG' +
    String(largestNumber + 1)
      .padStart(6, '0')
  );
}


/**
 * Tests Transaction_Log without changing inventory.
 */
function testTransactionLogDirect() {
  const logId = addTransactionLog_(
    getCurrentUserEmail_() ||
      'mis@anushagroup.com',
    'TEST',
    'DIRECT LOG TEST',
    'TEST-001',
    {
      message:
        'Transaction log test completed.'
    }
  );

  Logger.log(
    'Created Log ID: ' + logId
  );

  return logId;
}



/**
 * Repairs sheet row/column dimensions used by
 * inventory outward transactions.
 *
 * Run once after replacing Utils.gs.
 */
function repairInventorySheetDimensions() {
  const sheetNames = [
    'Outward',
    'Free_Inventory',
    'Customize_Inventory',
    'Transaction_Log'
  ];

  const minimumColumns = {
    'Outward': 20,
    'Free_Inventory': 21,
    'Customize_Inventory': 24,
    'Transaction_Log': 7
  };

  const result = {};

  sheetNames.forEach(
    function (sheetName) {
      const sheet =
        getSystemSheet(sheetName);

      const requiredColumns =
        minimumColumns[sheetName];

      if (
        sheet.getMaxColumns() <
        requiredColumns
      ) {
        sheet.insertColumnsAfter(
          sheet.getMaxColumns(),
          requiredColumns -
            sheet.getMaxColumns()
        );
      }

      /*
       * Keep spare rows available for new outward
       * and transaction records.
       */
      const spareRows =
        sheet.getMaxRows() -
        sheet.getLastRow();

      if (spareRows < 100) {
        sheet.insertRowsAfter(
          sheet.getMaxRows(),
          100 - spareRows
        );
      }

      result[sheetName] = {
        maxRows:
          sheet.getMaxRows(),
        lastRow:
          sheet.getLastRow(),
        maxColumns:
          sheet.getMaxColumns(),
        lastColumn:
          sheet.getLastColumn()
      };
    }
  );

  Logger.log(
    JSON.stringify(
      result,
      null,
      2
    )
  );

  return result;
}
