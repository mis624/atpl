/**
 * PROJECT INVENTORY MANAGEMENT SYSTEM
 * File: InventoryService.gs
 *
 * Core inventory balance service for:
 * - FREE INVENTORY
 * - CUSTOMIZE INVENTORY
 * - Stock balance updates
 * - Negative stock prevention
 * - Average rate calculation
 * - Reserved and damaged stock
 * - Project-wise customized inventory
 * - 90-day free conversion eligibility
 * - Inventory summary
 * - Low-stock report
 * - Transaction logging
 */

const INVENTORY_SERVICE_CONFIG = Object.freeze({
  FREE_SHEET: 'Free_Inventory',
  CUSTOMIZE_SHEET: 'Customize_Inventory',

  INVENTORY_TYPE: Object.freeze({
    FREE: 'FREE INVENTORY',
    CUSTOMIZE: 'CUSTOMIZE INVENTORY'
  }),

  MOVEMENT_TYPE: Object.freeze({
    INWARD: 'INWARD',
    OUTWARD: 'OUTWARD',
    TRANSFER_IN: 'TRANSFER IN',
    TRANSFER_OUT: 'TRANSFER OUT',
    CONVERSION_IN: 'CONVERSION IN',
    CONVERSION_OUT: 'CONVERSION OUT',
    ADJUSTMENT_IN: 'ADJUSTMENT IN',
    ADJUSTMENT_OUT: 'ADJUSTMENT OUT',
    RESERVE: 'RESERVE',
    RELEASE: 'RELEASE',
    DAMAGE: 'DAMAGE',
    DAMAGE_RECOVERY: 'DAMAGE RECOVERY'
  }),

  STATUS: Object.freeze({
  ACTIVE: 'AVAILABLE',
  OUT_OF_STOCK: 'OUT OF STOCK',
  INACTIVE: 'INACTIVE'
}),

  YES: 'YES',
  NO: 'NO',

  FREE_INVENTORY_ID_PREFIX: 'FIN',
  CUSTOMIZE_INVENTORY_ID_PREFIX: 'CIN',

  DEFAULT_FREE_ELIGIBILITY_DAYS: 90,
  LOCK_TIMEOUT_MS: 30000
});

/**
 * Returns Free_Inventory sheet name.
 */
function getFreeInventorySheetName_() {
  if (
    typeof APP_CONFIG !== 'undefined' &&
    APP_CONFIG.SHEETS &&
    APP_CONFIG.SHEETS.FREE_INVENTORY
  ) {
    return APP_CONFIG.SHEETS.FREE_INVENTORY;
  }

  return INVENTORY_SERVICE_CONFIG.FREE_SHEET;
}

/**
 * Returns Customize_Inventory sheet name.
 */
function getCustomizeInventorySheetName_() {
  if (
    typeof APP_CONFIG !== 'undefined' &&
    APP_CONFIG.SHEETS &&
    APP_CONFIG.SHEETS.CUSTOMIZE_INVENTORY
  ) {
    return APP_CONFIG.SHEETS.CUSTOMIZE_INVENTORY;
  }

  return INVENTORY_SERVICE_CONFIG.CUSTOMIZE_SHEET;
}

/**
 * Returns configured free eligibility days.
 */
function getFreeEligibilityDays_() {
  if (
    typeof APP_CONFIG !== 'undefined' &&
    Number(APP_CONFIG.FREE_ELIGIBILITY_DAYS) > 0
  ) {
    return Number(APP_CONFIG.FREE_ELIGIBILITY_DAYS);
  }

  try {
    const settingValue = getSettingValue_(
      'FREE_ELIGIBILITY_DAYS'
    );

    if (Number(settingValue) > 0) {
      return Number(settingValue);
    }
  } catch (error) {
    // Use default value.
  }

  return INVENTORY_SERVICE_CONFIG
    .DEFAULT_FREE_ELIGIBILITY_DAYS;
}

/**
 * Safely reads a setting value from Settings sheet.
 */
function getSettingValue_(settingKey) {
  const settingsSheetName =
    APP_CONFIG.SHEETS.SETTINGS || 'Settings';

  const records = getSheetObjects_(
    settingsSheetName
  );

  const normalizedKey = normalizeUpper_(settingKey);

  const record = records.find(function (item) {
    return (
      normalizeUpper_(
        item['Setting Key'] ||
        item['Key'] ||
        item['Setting']
      ) === normalizedKey
    );
  });

  if (!record) {
    return '';
  }

  return (
    record['Setting Value'] ||
    record['Value'] ||
    ''
  );
}

/**
 * Ensures required inventory headers exist.
 *
 * Existing headers and data are not deleted.
 */
function ensureInventoryHeaders() {
  return safeExecute_(function () {
    requireRole_(APP_CONFIG.USER_ROLES.ADMIN);

    const freeHeaders = [
      'Inventory ID',
      'SKU ID',
      'SKU Code',
      'SKU Name',
      'Category Name',
      'UOM',
      'Available Qty',
      'Reserved Qty',
      'Damaged Qty',
      'Minimum Stock',
      'Reorder Level',
      'Average Rate',
      'Last Inward Date',
      'Last Outward Date',
      'Last Movement Date',
      'Status',
      'Remarks',
      'Created By',
      'Created At',
      'Updated By',
      'Updated At'
    ];

    const customizeHeaders = [
      'Inventory ID',
      'Project ID',
      'Project Code',
      'Project Name',
      'SKU ID',
      'SKU Code',
      'SKU Name',
      'Category Name',
      'UOM',
      'Available Qty',
      'Reserved Qty',
      'Damaged Qty',
      'Average Rate',
      'Last Inward Date',
      'Last Outward Date',
      'Last Movement Date',
      'Eligible for Free',
      'Eligibility Date',
      'Status',
      'Remarks',
      'Created By',
      'Created At',
      'Updated By',
      'Updated At'
    ];

    const freeResult = ensureInventorySheetHeaders_(
      getFreeInventorySheetName_(),
      freeHeaders
    );

    const customizeResult =
      ensureInventorySheetHeaders_(
        getCustomizeInventorySheetName_(),
        customizeHeaders
      );

    return successResponse_(
      'Inventory headers verified successfully.',
      {
        freeInventory: freeResult,
        customizeInventory: customizeResult
      }
    );
  }, 'Unable to verify inventory headers.');
}

/**
 * Adds missing headers without deleting existing data.
 */
function ensureInventorySheetHeaders_(
  sheetName,
  requiredHeaders
) {
  const sheet = getSystemSheet(sheetName);

  const lastColumn = Math.max(
    sheet.getLastColumn(),
    1
  );

  let existingHeaders = sheet
    .getRange(1, 1, 1, lastColumn)
    .getDisplayValues()[0]
    .map(function (header) {
      return normalizeText_(header);
    });

  const missingHeaders = requiredHeaders.filter(
    function (header) {
      return existingHeaders.indexOf(header) === -1;
    }
  );

  if (missingHeaders.length > 0) {
    let startColumn = existingHeaders.length + 1;

    while (
      startColumn > 1 &&
      !normalizeText_(
        sheet.getRange(1, startColumn - 1).getValue()
      )
    ) {
      startColumn--;
    }

    sheet
      .getRange(
        1,
        startColumn,
        1,
        missingHeaders.length
      )
      .setValues([missingHeaders]);
  }

  const finalLastColumn = Math.max(
    sheet.getLastColumn(),
    requiredHeaders.length
  );

  sheet
    .getRange(1, 1, 1, finalLastColumn)
    .setFontWeight('bold');

  sheet.setFrozenRows(1);

  return {
    sheetName: sheetName,
    addedHeaders: missingHeaders
  };
}

/**
 * Main inventory movement function.
 *
 * This function is intended to be used by:
 * - InwardService.gs
 * - OutwardService.gs
 * - TransferService.gs
 * - ConversionService.gs
 *
 * movementData:
 * {
 *   inventoryType: 'FREE INVENTORY',
 *   movementType: 'INWARD',
 *   skuId: 'SKU000001',
 *   projectId: '',
 *   quantity: 10,
 *   rate: 500,
 *   referenceId: 'INW000001',
 *   remarks: ''
 * }
 */
function applyInventoryMovement(movementData) {
  return safeExecute_(function () {
    const session = requireAuthenticatedUser_();

    movementData = movementData || {};

    validateRequiredFields_(
      movementData,
      [
        'inventoryType',
        'movementType',
        'skuId',
        'quantity'
      ]
    );

    const inventoryType = normalizeInventoryType_(
      movementData.inventoryType
    );

    const movementType = normalizeMovementType_(
      movementData.movementType
    );

    const skuId = normalizeUpper_(
      movementData.skuId
    );

    const projectId = normalizeUpper_(
      movementData.projectId
    );

    const quantity = inventoryPositiveNumber_(
      movementData.quantity,
      'Quantity'
    );

    const rate = inventoryNonNegativeNumber_(
      movementData.rate,
      'Rate'
    );

    const referenceId = normalizeUpper_(
      movementData.referenceId
    );

    const remarks = normalizeText_(
      movementData.remarks
    );

    if (
      inventoryType ===
        INVENTORY_SERVICE_CONFIG
          .INVENTORY_TYPE.CUSTOMIZE &&
      !projectId
    ) {
      throw new Error(
        'Project ID is required for customized inventory.'
      );
    }

    const sku = getSkuRecordById_(skuId);

    if (!sku) {
      throw new Error(
        'SKU not found: ' + skuId
      );
    }

    if (
      normalizeUpper_(sku['Status']) !==
      'ACTIVE'
    ) {
      throw new Error(
        'Selected SKU is inactive: ' + skuId
      );
    }

    let project = null;

    if (
      inventoryType ===
      INVENTORY_SERVICE_CONFIG
        .INVENTORY_TYPE.CUSTOMIZE
    ) {
      project = getProjectRecordById_(projectId);

      if (!project) {
        throw new Error(
          'Project not found: ' + projectId
        );
      }

      const projectStatus = normalizeUpper_(
        project['Status']
      );

      if (
        projectStatus !== 'ACTIVE' &&
        movementType !==
          INVENTORY_SERVICE_CONFIG
            .MOVEMENT_TYPE.CONVERSION_OUT &&
        movementType !==
          INVENTORY_SERVICE_CONFIG
            .MOVEMENT_TYPE.TRANSFER_OUT
      ) {
        throw new Error(
          'Inventory cannot be added to a project with status: ' +
          projectStatus
        );
      }
    }

    const lock = LockService.getScriptLock();

    if (
      !lock.tryLock(
        INVENTORY_SERVICE_CONFIG.LOCK_TIMEOUT_MS
      )
    ) {
      throw new Error(
        'Inventory is currently being updated by another user. Please try again.'
      );
    }

    try {
      const result = updateInventoryBalance_({
        inventoryType: inventoryType,
        movementType: movementType,
        sku: sku,
        project: project,
        quantity: quantity,
        rate: rate,
        referenceId: referenceId,
        remarks: remarks,
        updatedBy: session.email
      });

      addTransactionLog_(
        session.email,
        'INVENTORY',
        movementType,
        result.inventoryId,
        {
          inventoryType: inventoryType,
          skuId: skuId,
          projectId: projectId,
          quantity: quantity,
          rate: rate,
          referenceId: referenceId,
          previousAvailableQty:
            result.previousAvailableQty,
          newAvailableQty:
            result.newAvailableQty
        }
      );

      return successResponse_(
        'Inventory movement completed successfully.',
        result
      );
    } finally {
      lock.releaseLock();
    }
  }, 'Unable to update inventory.');
}

/**
 * Internal inventory balance update.
 */
function updateInventoryBalance_(data) {
  const isFree =
    data.inventoryType ===
    INVENTORY_SERVICE_CONFIG.INVENTORY_TYPE.FREE;

  const sheetName = isFree
    ? getFreeInventorySheetName_()
    : getCustomizeInventorySheetName_();

  let inventoryRecord = isFree
    ? getFreeInventoryRecordBySkuId_(
        data.sku['SKU ID']
      )
    : getCustomizeInventoryRecord_(
        data.project['Project ID'],
        data.sku['SKU ID']
      );

  const now = new Date();

  if (!inventoryRecord) {
    if (
      isInventoryReductionMovement_(
        data.movementType
      )
    ) {
      throw new Error(
        'Inventory record does not exist for the selected SKU.'
      );
    }

    inventoryRecord = createInventoryRecord_({
      inventoryType: data.inventoryType,
      sku: data.sku,
      project: data.project,
      rate: data.rate,
      createdBy: data.updatedBy,
      now: now
    });
  }

  const previousAvailableQty =
    inventoryToNumber_(
      inventoryRecord['Available Qty'],
      0
    );

  const previousReservedQty =
    inventoryToNumber_(
      inventoryRecord['Reserved Qty'],
      0
    );

  const previousDamagedQty =
    inventoryToNumber_(
      inventoryRecord['Damaged Qty'],
      0
    );

  const previousAverageRate =
    inventoryToNumber_(
      inventoryRecord['Average Rate'],
      0
    );

  let newAvailableQty = previousAvailableQty;
  let newReservedQty = previousReservedQty;
  let newDamagedQty = previousDamagedQty;
  let newAverageRate = previousAverageRate;

  switch (data.movementType) {
    case INVENTORY_SERVICE_CONFIG
      .MOVEMENT_TYPE.INWARD:

    case INVENTORY_SERVICE_CONFIG
      .MOVEMENT_TYPE.TRANSFER_IN:

    case INVENTORY_SERVICE_CONFIG
      .MOVEMENT_TYPE.CONVERSION_IN:

    case INVENTORY_SERVICE_CONFIG
      .MOVEMENT_TYPE.ADJUSTMENT_IN:

      newAverageRate = calculateWeightedAverageRate_(
        previousAvailableQty,
        previousAverageRate,
        data.quantity,
        data.rate
      );

      newAvailableQty =
        previousAvailableQty + data.quantity;
      break;

    case INVENTORY_SERVICE_CONFIG
      .MOVEMENT_TYPE.OUTWARD:

    case INVENTORY_SERVICE_CONFIG
      .MOVEMENT_TYPE.TRANSFER_OUT:

    case INVENTORY_SERVICE_CONFIG
      .MOVEMENT_TYPE.CONVERSION_OUT:

    case INVENTORY_SERVICE_CONFIG
      .MOVEMENT_TYPE.ADJUSTMENT_OUT:

      if (
        previousAvailableQty < data.quantity
      ) {
        throw new Error(
          'Insufficient available stock. Available: ' +
          previousAvailableQty +
          ', Requested: ' +
          data.quantity
        );
      }

      newAvailableQty =
        previousAvailableQty - data.quantity;
      break;

    case INVENTORY_SERVICE_CONFIG
      .MOVEMENT_TYPE.RESERVE:

      if (
        previousAvailableQty < data.quantity
      ) {
        throw new Error(
          'Insufficient stock for reservation. Available: ' +
          previousAvailableQty
        );
      }

      newAvailableQty =
        previousAvailableQty - data.quantity;

      newReservedQty =
        previousReservedQty + data.quantity;
      break;

    case INVENTORY_SERVICE_CONFIG
      .MOVEMENT_TYPE.RELEASE:

      if (
        previousReservedQty < data.quantity
      ) {
        throw new Error(
          'Reserved stock is lower than the release quantity.'
        );
      }

      newReservedQty =
        previousReservedQty - data.quantity;

      newAvailableQty =
        previousAvailableQty + data.quantity;
      break;

    case INVENTORY_SERVICE_CONFIG
      .MOVEMENT_TYPE.DAMAGE:

      if (
        previousAvailableQty < data.quantity
      ) {
        throw new Error(
          'Insufficient available stock for damage entry.'
        );
      }

      newAvailableQty =
        previousAvailableQty - data.quantity;

      newDamagedQty =
        previousDamagedQty + data.quantity;
      break;

    case INVENTORY_SERVICE_CONFIG
      .MOVEMENT_TYPE.DAMAGE_RECOVERY:

      if (
        previousDamagedQty < data.quantity
      ) {
        throw new Error(
          'Damaged stock is lower than the recovery quantity.'
        );
      }

      newDamagedQty =
        previousDamagedQty - data.quantity;

      newAvailableQty =
        previousAvailableQty + data.quantity;
      break;

    default:
      throw new Error(
        'Unsupported inventory movement type.'
      );
  }

  newAvailableQty =
    inventoryRoundQuantity_(newAvailableQty);

  newReservedQty =
    inventoryRoundQuantity_(newReservedQty);

  newDamagedQty =
    inventoryRoundQuantity_(newDamagedQty);

  newAverageRate =
    inventoryRoundMoney_(newAverageRate);

  const updateData = {
    'Available Qty': newAvailableQty,
    'Reserved Qty': newReservedQty,
    'Damaged Qty': newDamagedQty,
    'Average Rate': newAverageRate,
    'Last Movement Date': now,
    'Status': getInventoryStatus_(
  data.inventoryType,
  newAvailableQty,
  newReservedQty
),
    'Remarks': data.remarks,
    'Updated By': data.updatedBy,
    'Updated At': now
  };

  if (
    isInventoryAdditionMovement_(
      data.movementType
    )
  ) {
    updateData['Last Inward Date'] = now;
  }

  if (
    isInventoryReductionMovement_(
      data.movementType
    )
  ) {
    updateData['Last Outward Date'] = now;
  }

  if (!isFree) {
    const eligibility = calculateFreeEligibility_(
      newAvailableQty,
      now
    );

    updateData['Eligible for Free'] =
      eligibility.eligible;

    updateData['Eligibility Date'] =
      eligibility.eligibilityDate;
  }

  updateObjectRow_(
    sheetName,
    inventoryRecord._rowNumber,
    updateData
  );

  return {
    inventoryId: normalizeUpper_(
      inventoryRecord['Inventory ID']
    ),
    inventoryType: data.inventoryType,
    projectId: data.project
      ? normalizeUpper_(
          data.project['Project ID']
        )
      : '',
    skuId: normalizeUpper_(
      data.sku['SKU ID']
    ),
    skuCode: normalizeUpper_(
      data.sku['SKU Code']
    ),
    skuName: normalizeText_(
      data.sku['SKU Name']
    ),
    movementType: data.movementType,
    quantity: data.quantity,
    rate: data.rate,
    previousAvailableQty:
      previousAvailableQty,
    newAvailableQty: newAvailableQty,
    reservedQty: newReservedQty,
    damagedQty: newDamagedQty,
    averageRate: newAverageRate,
    referenceId: data.referenceId
  };
}

/**
 * Creates a new inventory balance record.
 */
function createInventoryRecord_(data) {
  const isFree =
    data.inventoryType ===
    INVENTORY_SERVICE_CONFIG.INVENTORY_TYPE.FREE;

  const sheetName = isFree
    ? getFreeInventorySheetName_()
    : getCustomizeInventorySheetName_();

  const inventoryId = generateNextId_(
    isFree
      ? 'FREE_INVENTORY'
      : 'CUSTOMIZE_INVENTORY',
    isFree
      ? INVENTORY_SERVICE_CONFIG
          .FREE_INVENTORY_ID_PREFIX
      : INVENTORY_SERVICE_CONFIG
          .CUSTOMIZE_INVENTORY_ID_PREFIX
  );

  let record;

  /*
   * Keep the object fields in the same order as the sheet headers.
   * appendObjectRow_ writes values using object insertion order.
   */
  if (isFree) {
    record = {
      'Inventory ID': inventoryId,
      'SKU ID': normalizeUpper_(
        data.sku['SKU ID']
      ),
      'SKU Code': normalizeUpper_(
        data.sku['SKU Code']
      ),
      'SKU Name': normalizeText_(
        data.sku['SKU Name']
      ),
      'Category Name': normalizeText_(
        data.sku['Category Name']
      ),
      'UOM': normalizeUpper_(
        data.sku['UOM']
      ),
      'Available Qty': 0,
      'Reserved Qty': 0,
      'Damaged Qty': 0,
      'Minimum Stock': inventoryToNumber_(
        data.sku['Minimum Stock'],
        0
      ),
      'Reorder Level': inventoryToNumber_(
        data.sku['Reorder Level'],
        0
      ),
      'Average Rate': inventoryRoundMoney_(
        data.rate
      ),
      'Last Inward Date': '',
      'Last Outward Date': '',
      'Last Movement Date': '',
      'Status':
        INVENTORY_SERVICE_CONFIG
          .STATUS.OUT_OF_STOCK,
      'Remarks': '',
      'Created By': data.createdBy,
      'Created At': data.now,
      'Updated By': data.createdBy,
      'Updated At': data.now
    };
  } else {
    record = {
      'Inventory ID': inventoryId,
      'Project ID': normalizeUpper_(
        data.project['Project ID']
      ),
      'Project Code': normalizeUpper_(
        data.project['Project Code']
      ),
      'Project Name': normalizeText_(
        data.project['Project Name']
      ),
      'SKU ID': normalizeUpper_(
        data.sku['SKU ID']
      ),
      'SKU Code': normalizeUpper_(
        data.sku['SKU Code']
      ),
      'SKU Name': normalizeText_(
        data.sku['SKU Name']
      ),
      'Category Name': normalizeText_(
        data.sku['Category Name']
      ),
      'UOM': normalizeUpper_(
        data.sku['UOM']
      ),
      'Available Qty': 0,
      'Reserved Qty': 0,
      'Damaged Qty': 0,
      'Average Rate': inventoryRoundMoney_(
        data.rate
      ),
      'Last Inward Date': '',
      'Last Outward Date': '',
      'Last Movement Date': '',
      'Eligible for Free':
        INVENTORY_SERVICE_CONFIG.NO,
      'Eligibility Date': '',
      'Status': 'AVAILABLE',
      'Remarks': '',
      'Created By': data.createdBy,
      'Created At': data.now,
      'Updated By': data.createdBy,
      'Updated At': data.now
    };
  }

  const rowNumber = appendObjectRow_(
    sheetName,
    record
  );

  SpreadsheetApp.flush();

  const savedRecord = isFree
    ? getFreeInventoryRecordBySkuId_(
        data.sku['SKU ID']
      )
    : getCustomizeInventoryRecord_(
        data.project['Project ID'],
        data.sku['SKU ID']
      );

  if (!savedRecord) {
    throw new Error(
      'Inventory record could not be created.'
    );
  }

  if (
    !savedRecord._rowNumber &&
    Number(rowNumber) > 0
  ) {
    savedRecord._rowNumber = Number(rowNumber);
  }

  return savedRecord;
}

/**
 * Returns Free Inventory record by SKU ID.
 */
function getFreeInventoryRecordBySkuId_(skuId) {
  const normalizedSkuId =
    normalizeUpper_(skuId);

  if (!normalizedSkuId) {
    return null;
  }

  const records = getSheetObjects_(
    getFreeInventorySheetName_()
  );

  return records.find(function (record) {
    return (
      normalizeUpper_(record['SKU ID']) ===
      normalizedSkuId
    );
  }) || null;
}

/**
 * Returns Customize Inventory record by Project ID and SKU ID.
 */
function getCustomizeInventoryRecord_(
  projectId,
  skuId
) {
  const normalizedProjectId =
    normalizeUpper_(projectId);

  const normalizedSkuId =
    normalizeUpper_(skuId);

  if (
    !normalizedProjectId ||
    !normalizedSkuId
  ) {
    return null;
  }

  const records = getSheetObjects_(
    getCustomizeInventorySheetName_()
  );

  return records.find(function (record) {
    return (
      normalizeUpper_(record['Project ID']) ===
        normalizedProjectId &&
      normalizeUpper_(record['SKU ID']) ===
        normalizedSkuId
    );
  }) || null;
}

/**
 * Returns free inventory stock by SKU.
 */
function getFreeInventoryBySku(skuId) {
  return safeExecute_(function () {
    requireAuthenticatedUser_();

    const record =
      getFreeInventoryRecordBySkuId_(skuId);

    if (!record) {
      throw new Error(
        'Free inventory record not found for SKU: ' +
        skuId
      );
    }

    return successResponse_(
      'Free inventory loaded successfully.',
      mapFreeInventoryRecord_(record)
    );
  }, 'Unable to load free inventory.');
}

/**
 * Returns project customized stock by SKU.
 */
function getCustomizeInventoryBySku(
  projectId,
  skuId
) {
  return safeExecute_(function () {
    requireAuthenticatedUser_();

    const record =
      getCustomizeInventoryRecord_(
        projectId,
        skuId
      );

    if (!record) {
      throw new Error(
        'Customized inventory record not found.'
      );
    }

    return successResponse_(
      'Customized inventory loaded successfully.',
      mapCustomizeInventoryRecord_(record)
    );
  }, 'Unable to load customized inventory.');
}

/**
 * Returns free inventory list with filters.
 */
function getFreeInventory(filters) {
  return safeExecute_(function () {
    requireAuthenticatedUser_();

    filters = filters || {};

    const search = normalizeLower_(
      filters.search
    );

    const status = normalizeUpper_(
      filters.status
    );

    const lowStockOnly =
      filters.lowStockOnly === true ||
      normalizeUpper_(filters.lowStockOnly) ===
        'TRUE';

    let records = getSheetObjects_(
      getFreeInventorySheetName_()
    );

    records = records.filter(function (record) {
      if (
        status &&
        normalizeUpper_(record['Status']) !== status
      ) {
        return false;
      }

      if (lowStockOnly) {
        const availableQty =
          inventoryToNumber_(
            record['Available Qty'],
            0
          );

        const reorderLevel =
          inventoryToNumber_(
            record['Reorder Level'],
            0
          );

        if (
          reorderLevel <= 0 ||
          availableQty > reorderLevel
        ) {
          return false;
        }
      }

      if (search) {
        const searchableText = [
          record['Inventory ID'],
          record['SKU ID'],
          record['SKU Code'],
          record['SKU Name'],
          record['Category Name'],
          record['UOM'],
          record['Status']
        ]
          .map(function (value) {
            return normalizeLower_(value);
          })
          .join(' ');

        if (
          searchableText.indexOf(search) === -1
        ) {
          return false;
        }
      }

      return true;
    });

    records.sort(function (first, second) {
      return normalizeText_(
        first['SKU Name']
      ).localeCompare(
        normalizeText_(second['SKU Name'])
      );
    });

    const mappedRecords = records.map(
      mapFreeInventoryRecord_
    );

    return successResponse_(
      'Free inventory loaded successfully.',
      paginateRecords_(
        mappedRecords,
        filters.pageNumber || 1,
        filters.pageSize || 20
      )
    );
  }, 'Unable to load free inventory.');
}

/**
 * Returns customized inventory list with filters.
 */
function getCustomizeInventory(filters) {
  return safeExecute_(function () {
    const session = requireAuthenticatedUser_();

    filters = filters || {};

    const search = normalizeLower_(
      filters.search
    );

    const projectId = normalizeUpper_(
      filters.projectId
    );

    const status = normalizeUpper_(
      filters.status
    );

    const eligibleOnly =
      filters.eligibleOnly === true ||
      normalizeUpper_(filters.eligibleOnly) ===
        'TRUE';

    let records = getSheetObjects_(
      getCustomizeInventorySheetName_()
    );

    records = records.filter(function (record) {
      if (
        session.role ===
          APP_CONFIG.USER_ROLES.DOER
      ) {
        const project = getProjectRecordById_(
          record['Project ID']
        );

        if (
          !project ||
          normalizeLower_(
            project['Doer Email']
          ) !== session.email
        ) {
          return false;
        }
      }

      if (
        projectId &&
        normalizeUpper_(
          record['Project ID']
        ) !== projectId
      ) {
        return false;
      }

      if (
        status &&
        normalizeUpper_(record['Status']) !== status
      ) {
        return false;
      }

      if (
        eligibleOnly &&
        normalizeUpper_(
          record['Eligible for Free']
        ) !== INVENTORY_SERVICE_CONFIG.YES
      ) {
        return false;
      }

      if (search) {
        const searchableText = [
          record['Inventory ID'],
          record['Project ID'],
          record['Project Code'],
          record['Project Name'],
          record['SKU ID'],
          record['SKU Code'],
          record['SKU Name'],
          record['Category Name'],
          record['Status']
        ]
          .map(function (value) {
            return normalizeLower_(value);
          })
          .join(' ');

        if (
          searchableText.indexOf(search) === -1
        ) {
          return false;
        }
      }

      return true;
    });

    records.sort(function (first, second) {
      const projectCompare = normalizeText_(
        first['Project Name']
      ).localeCompare(
        normalizeText_(second['Project Name'])
      );

      if (projectCompare !== 0) {
        return projectCompare;
      }

      return normalizeText_(
        first['SKU Name']
      ).localeCompare(
        normalizeText_(second['SKU Name'])
      );
    });

    const mappedRecords = records.map(
      mapCustomizeInventoryRecord_
    );

    return successResponse_(
      'Customized inventory loaded successfully.',
      paginateRecords_(
        mappedRecords,
        filters.pageNumber || 1,
        filters.pageSize || 20
      )
    );
  }, 'Unable to load customized inventory.');
}

/**
 * Returns current available stock.
 */
function getAvailableInventoryQuantity(
  inventoryType,
  skuId,
  projectId
) {
  return safeExecute_(function () {
    requireAuthenticatedUser_();

    const normalizedType =
      normalizeInventoryType_(inventoryType);

    let record;

    if (
      normalizedType ===
      INVENTORY_SERVICE_CONFIG
        .INVENTORY_TYPE.FREE
    ) {
      record = getFreeInventoryRecordBySkuId_(
        skuId
      );
    } else {
      record = getCustomizeInventoryRecord_(
        projectId,
        skuId
      );
    }

    const availableQty = record
      ? inventoryToNumber_(
          record['Available Qty'],
          0
        )
      : 0;

    return successResponse_(
      'Available quantity loaded successfully.',
      {
        inventoryType: normalizedType,
        projectId: normalizeUpper_(projectId),
        skuId: normalizeUpper_(skuId),
        availableQty: availableQty,
        reservedQty: record
          ? inventoryToNumber_(
              record['Reserved Qty'],
              0
            )
          : 0,
        damagedQty: record
          ? inventoryToNumber_(
              record['Damaged Qty'],
              0
            )
          : 0
      }
    );
  }, 'Unable to load available quantity.');
}

/**
 * Reserves stock.
 */
function reserveInventory(data) {
  data = data || {};

  data.movementType =
    INVENTORY_SERVICE_CONFIG
      .MOVEMENT_TYPE.RESERVE;

  return applyInventoryMovement(data);
}

/**
 * Releases reserved stock.
 */
function releaseReservedInventory(data) {
  data = data || {};

  data.movementType =
    INVENTORY_SERVICE_CONFIG
      .MOVEMENT_TYPE.RELEASE;

  return applyInventoryMovement(data);
}

/**
 * Moves available stock to damaged stock.
 */
function markInventoryDamaged(data) {
  data = data || {};

  data.movementType =
    INVENTORY_SERVICE_CONFIG
      .MOVEMENT_TYPE.DAMAGE;

  return applyInventoryMovement(data);
}

/**
 * Recovers damaged stock back to available stock.
 */
function recoverDamagedInventory(data) {
  data = data || {};

  data.movementType =
    INVENTORY_SERVICE_CONFIG
      .MOVEMENT_TYPE.DAMAGE_RECOVERY;

  return applyInventoryMovement(data);
}

/**
 * Recalculates customized inventory eligibility.
 *
 * Customized inventory becomes eligible only after
 * no movement for the configured number of days.
 *
 * This does not convert stock automatically.
 */
function refreshCustomizeInventoryEligibility() {
  return safeExecute_(function () {
    const session = requireRole_([
      APP_CONFIG.USER_ROLES.ADMIN,
      APP_CONFIG.USER_ROLES.HOD
    ]);

    const records = getSheetObjects_(
      getCustomizeInventorySheetName_()
    );

    const eligibilityDays =
      getFreeEligibilityDays_();

    const now = new Date();

    let updatedRecords = 0;
    let eligibleRecords = 0;

    records.forEach(function (record) {
      const availableQty =
        inventoryToNumber_(
          record['Available Qty'],
          0
        );

      const lastMovementDate =
        inventoryParseDate_(
          record['Last Movement Date']
        ) ||
        inventoryParseDate_(
          record['Last Inward Date']
        ) ||
        inventoryParseDate_(
          record['Created At']
        );

      let eligible =
        INVENTORY_SERVICE_CONFIG.NO;

      let eligibilityDate = '';

      if (
        availableQty > 0 &&
        lastMovementDate
      ) {
        eligibilityDate = new Date(
          lastMovementDate.getTime() +
          eligibilityDays *
          24 *
          60 *
          60 *
          1000
        );

        if (
          now.getTime() >=
          eligibilityDate.getTime()
        ) {
          eligible =
            INVENTORY_SERVICE_CONFIG.YES;

          eligibleRecords++;
        }
      }

      const currentEligible =
        normalizeUpper_(
          record['Eligible for Free']
        );

      const currentEligibilityDate =
        inventoryParseDate_(
          record['Eligibility Date']
        );

      const dateChanged =
        inventoryDateValue_(
          currentEligibilityDate
        ) !==
        inventoryDateValue_(
          eligibilityDate
        );

      if (
        currentEligible !== eligible ||
        dateChanged
      ) {
        updateObjectRow_(
          getCustomizeInventorySheetName_(),
          record._rowNumber,
          {
            'Eligible for Free': eligible,
            'Eligibility Date': eligibilityDate,
            'Updated By': session.email,
            'Updated At': now
          }
        );

        updatedRecords++;
      }
    });

    addTransactionLog_(
      session.email,
      'INVENTORY',
      'REFRESH ELIGIBILITY',
      'CUSTOMIZE INVENTORY',
      {
        eligibilityDays: eligibilityDays,
        checkedRecords: records.length,
        updatedRecords: updatedRecords,
        eligibleRecords: eligibleRecords
      }
    );

    return successResponse_(
      'Customized inventory eligibility refreshed successfully.',
      {
        eligibilityDays: eligibilityDays,
        checkedRecords: records.length,
        updatedRecords: updatedRecords,
        eligibleRecords: eligibleRecords
      }
    );
  }, 'Unable to refresh inventory eligibility.');
}

/**
 * Returns customized inventory eligible for free conversion.
 */
function getEligibleCustomizeInventory(filters) {
  filters = filters || {};
  filters.eligibleOnly = true;

  return getCustomizeInventory(filters);
}

/**
 * Returns inventory summary.
 */
function getInventorySummary() {
  return safeExecute_(function () {
    requireAuthenticatedUser_();

    const freeRecords = getSheetObjects_(
      getFreeInventorySheetName_()
    );

    const customizeRecords = getSheetObjects_(
      getCustomizeInventorySheetName_()
    );

    const summary = {
      freeInventory: {
        totalSkuRecords: freeRecords.length,
        totalAvailableQty: 0,
        totalReservedQty: 0,
        totalDamagedQty: 0,
        totalStockValue: 0,
        lowStockRecords: 0,
        outOfStockRecords: 0
      },

      customizeInventory: {
        totalRecords: customizeRecords.length,
        totalAvailableQty: 0,
        totalReservedQty: 0,
        totalDamagedQty: 0,
        totalStockValue: 0,
        eligibleForFreeRecords: 0,
        outOfStockRecords: 0
      },

      combined: {
        totalAvailableQty: 0,
        totalReservedQty: 0,
        totalDamagedQty: 0,
        totalStockValue: 0
      }
    };

    freeRecords.forEach(function (record) {
      const availableQty =
        inventoryToNumber_(
          record['Available Qty'],
          0
        );

      const reservedQty =
        inventoryToNumber_(
          record['Reserved Qty'],
          0
        );

      const damagedQty =
        inventoryToNumber_(
          record['Damaged Qty'],
          0
        );

      const averageRate =
        inventoryToNumber_(
          record['Average Rate'],
          0
        );

      const reorderLevel =
        inventoryToNumber_(
          record['Reorder Level'],
          0
        );

      summary.freeInventory
        .totalAvailableQty += availableQty;

      summary.freeInventory
        .totalReservedQty += reservedQty;

      summary.freeInventory
        .totalDamagedQty += damagedQty;

      summary.freeInventory
        .totalStockValue +=
          availableQty * averageRate;

      if (
        reorderLevel > 0 &&
        availableQty <= reorderLevel
      ) {
        summary.freeInventory
          .lowStockRecords++;
      }

      if (availableQty <= 0) {
        summary.freeInventory
          .outOfStockRecords++;
      }
    });

    customizeRecords.forEach(function (record) {
      const availableQty =
        inventoryToNumber_(
          record['Available Qty'],
          0
        );

      const reservedQty =
        inventoryToNumber_(
          record['Reserved Qty'],
          0
        );

      const damagedQty =
        inventoryToNumber_(
          record['Damaged Qty'],
          0
        );

      const averageRate =
        inventoryToNumber_(
          record['Average Rate'],
          0
        );

      summary.customizeInventory
        .totalAvailableQty += availableQty;

      summary.customizeInventory
        .totalReservedQty += reservedQty;

      summary.customizeInventory
        .totalDamagedQty += damagedQty;

      summary.customizeInventory
        .totalStockValue +=
          availableQty * averageRate;

      if (
        normalizeUpper_(
          record['Eligible for Free']
        ) === INVENTORY_SERVICE_CONFIG.YES
      ) {
        summary.customizeInventory
          .eligibleForFreeRecords++;
      }

      if (availableQty <= 0) {
        summary.customizeInventory
          .outOfStockRecords++;
      }
    });

    summary.freeInventory.totalStockValue =
      inventoryRoundMoney_(
        summary.freeInventory.totalStockValue
      );

    summary.customizeInventory.totalStockValue =
      inventoryRoundMoney_(
        summary.customizeInventory
          .totalStockValue
      );

    summary.combined.totalAvailableQty =
      inventoryRoundQuantity_(
        summary.freeInventory
          .totalAvailableQty +
        summary.customizeInventory
          .totalAvailableQty
      );

    summary.combined.totalReservedQty =
      inventoryRoundQuantity_(
        summary.freeInventory
          .totalReservedQty +
        summary.customizeInventory
          .totalReservedQty
      );

    summary.combined.totalDamagedQty =
      inventoryRoundQuantity_(
        summary.freeInventory
          .totalDamagedQty +
        summary.customizeInventory
          .totalDamagedQty
      );

    summary.combined.totalStockValue =
      inventoryRoundMoney_(
        summary.freeInventory
          .totalStockValue +
        summary.customizeInventory
          .totalStockValue
      );

    return successResponse_(
      'Inventory summary loaded successfully.',
      summary
    );
  }, 'Unable to load inventory summary.');
}

/**
 * Returns free inventory low-stock records.
 */
function getLowStockInventory() {
  return getFreeInventory({
    lowStockOnly: true,
    pageNumber: 1,
    pageSize: 500
  });
}

/**
 * Calculates 90-day eligibility for a new movement.
 */
function calculateFreeEligibility_(
  availableQty,
  lastMovementDate
) {
  if (
    Number(availableQty) <= 0 ||
    !lastMovementDate
  ) {
    return {
      eligible: INVENTORY_SERVICE_CONFIG.NO,
      eligibilityDate: ''
    };
  }

  const eligibilityDate = new Date(
    lastMovementDate.getTime() +
    getFreeEligibilityDays_() *
    24 *
    60 *
    60 *
    1000
  );

  return {
    eligible:
      new Date().getTime() >=
      eligibilityDate.getTime()
        ? INVENTORY_SERVICE_CONFIG.YES
        : INVENTORY_SERVICE_CONFIG.NO,

    eligibilityDate: eligibilityDate
  };
}

/**
 * Calculates weighted average stock rate.
 */
function calculateWeightedAverageRate_(
  existingQty,
  existingRate,
  newQty,
  newRate
) {
  existingQty = inventoryToNumber_(
    existingQty,
    0
  );

  existingRate = inventoryToNumber_(
    existingRate,
    0
  );

  newQty = inventoryToNumber_(
    newQty,
    0
  );

  newRate = inventoryToNumber_(
    newRate,
    0
  );

  const totalQty = existingQty + newQty;

  if (totalQty <= 0) {
    return 0;
  }

  return inventoryRoundMoney_(
    (
      existingQty * existingRate +
      newQty * newRate
    ) / totalQty
  );
}

/**
 * Checks whether movement adds available stock.
 */
function isInventoryAdditionMovement_(
  movementType
) {
  return [
    INVENTORY_SERVICE_CONFIG
      .MOVEMENT_TYPE.INWARD,

    INVENTORY_SERVICE_CONFIG
      .MOVEMENT_TYPE.TRANSFER_IN,

    INVENTORY_SERVICE_CONFIG
      .MOVEMENT_TYPE.CONVERSION_IN,

    INVENTORY_SERVICE_CONFIG
      .MOVEMENT_TYPE.ADJUSTMENT_IN
  ].indexOf(movementType) !== -1;
}

/**
 * Checks whether movement reduces available stock.
 */
function isInventoryReductionMovement_(
  movementType
) {
  return [
    INVENTORY_SERVICE_CONFIG
      .MOVEMENT_TYPE.OUTWARD,

    INVENTORY_SERVICE_CONFIG
      .MOVEMENT_TYPE.TRANSFER_OUT,

    INVENTORY_SERVICE_CONFIG
      .MOVEMENT_TYPE.CONVERSION_OUT,

    INVENTORY_SERVICE_CONFIG
      .MOVEMENT_TYPE.ADJUSTMENT_OUT
  ].indexOf(movementType) !== -1;
}

/**
 * Validates inventory type.
 */
function normalizeInventoryType_(inventoryType) {
  const normalizedType =
    normalizeUpper_(inventoryType);

  const validTypes = [
    INVENTORY_SERVICE_CONFIG
      .INVENTORY_TYPE.FREE,

    INVENTORY_SERVICE_CONFIG
      .INVENTORY_TYPE.CUSTOMIZE
  ];

  if (
    validTypes.indexOf(normalizedType) === -1
  ) {
    throw new Error(
      'Invalid inventory type: ' +
      normalizedType
    );
  }

  return normalizedType;
}

/**
 * Validates movement type.
 */
function normalizeMovementType_(movementType) {
  const normalizedType =
    normalizeUpper_(movementType);

  const validTypes = Object.keys(
    INVENTORY_SERVICE_CONFIG.MOVEMENT_TYPE
  ).map(function (key) {
    return INVENTORY_SERVICE_CONFIG
      .MOVEMENT_TYPE[key];
  });

  if (
    validTypes.indexOf(normalizedType) === -1
  ) {
    throw new Error(
      'Invalid inventory movement type: ' +
      normalizedType
    );
  }

  return normalizedType;
}

/**
 * Maps Free_Inventory record.
 */
function mapFreeInventoryRecord_(record) {
  const availableQty =
    inventoryToNumber_(
      record['Available Qty'],
      0
    );

  const averageRate =
    inventoryToNumber_(
      record['Average Rate'],
      0
    );

  const reorderLevel =
    inventoryToNumber_(
      record['Reorder Level'],
      0
    );

  return {
    inventoryId: normalizeUpper_(
      record['Inventory ID']
    ),
    inventoryType:
      INVENTORY_SERVICE_CONFIG
        .INVENTORY_TYPE.FREE,
    skuId: normalizeUpper_(
      record['SKU ID']
    ),
    skuCode: normalizeUpper_(
      record['SKU Code']
    ),
    skuName: normalizeText_(
      record['SKU Name']
    ),
    categoryName: normalizeText_(
      record['Category Name']
    ),
    uom: normalizeUpper_(
      record['UOM']
    ),
    availableQty: availableQty,
    reservedQty: inventoryToNumber_(
      record['Reserved Qty'],
      0
    ),
    damagedQty: inventoryToNumber_(
      record['Damaged Qty'],
      0
    ),
    minimumStock: inventoryToNumber_(
      record['Minimum Stock'],
      0
    ),
    reorderLevel: reorderLevel,
    averageRate: averageRate,
    stockValue: inventoryRoundMoney_(
      availableQty * averageRate
    ),
    isLowStock:
      reorderLevel > 0 &&
      availableQty <= reorderLevel,
    lastInwardDate:
      inventoryFormatDateTime_(
        record['Last Inward Date']
      ),
    lastOutwardDate:
      inventoryFormatDateTime_(
        record['Last Outward Date']
      ),
    lastMovementDate:
      inventoryFormatDateTime_(
        record['Last Movement Date']
      ),
    status: normalizeUpper_(
      record['Status']
    ),
    remarks: normalizeText_(
      record['Remarks']
    ),
    createdBy: normalizeLower_(
      record['Created By']
    ),
    createdAt: inventoryFormatDateTime_(
      record['Created At']
    ),
    updatedBy: normalizeLower_(
      record['Updated By']
    ),
    updatedAt: inventoryFormatDateTime_(
      record['Updated At']
    )
  };
}

/**
 * Maps Customize_Inventory record.
 */
function mapCustomizeInventoryRecord_(record) {
  const availableQty =
    inventoryToNumber_(
      record['Available Qty'],
      0
    );

  const averageRate =
    inventoryToNumber_(
      record['Average Rate'],
      0
    );

  return {
    inventoryId: normalizeUpper_(
      record['Inventory ID']
    ),
    inventoryType:
      INVENTORY_SERVICE_CONFIG
        .INVENTORY_TYPE.CUSTOMIZE,
    projectId: normalizeUpper_(
      record['Project ID']
    ),
    projectCode: normalizeUpper_(
      record['Project Code']
    ),
    projectName: normalizeText_(
      record['Project Name']
    ),
    skuId: normalizeUpper_(
      record['SKU ID']
    ),
    skuCode: normalizeUpper_(
      record['SKU Code']
    ),
    skuName: normalizeText_(
      record['SKU Name']
    ),
    categoryName: normalizeText_(
      record['Category Name']
    ),
    uom: normalizeUpper_(
      record['UOM']
    ),
    availableQty: availableQty,
    reservedQty: inventoryToNumber_(
      record['Reserved Qty'],
      0
    ),
    damagedQty: inventoryToNumber_(
      record['Damaged Qty'],
      0
    ),
    averageRate: averageRate,
    stockValue: inventoryRoundMoney_(
      availableQty * averageRate
    ),
    lastInwardDate:
      inventoryFormatDateTime_(
        record['Last Inward Date']
      ),
    lastOutwardDate:
      inventoryFormatDateTime_(
        record['Last Outward Date']
      ),
    lastMovementDate:
      inventoryFormatDateTime_(
        record['Last Movement Date']
      ),
    eligibleForFree: normalizeUpper_(
      record['Eligible for Free']
    ),
    eligibilityDate:
      inventoryFormatDateTime_(
        record['Eligibility Date']
      ),
    status: normalizeUpper_(
      record['Status']
    ),
    remarks: normalizeText_(
      record['Remarks']
    ),
    createdBy: normalizeLower_(
      record['Created By']
    ),
    createdAt: inventoryFormatDateTime_(
      record['Created At']
    ),
    updatedBy: normalizeLower_(
      record['Updated By']
    ),
    updatedAt: inventoryFormatDateTime_(
      record['Updated At']
    )
  };
}

/**
 * Converts value to number.
 */
function inventoryToNumber_(
  value,
  defaultValue
) {
  if (
    value === '' ||
    value === null ||
    typeof value === 'undefined'
  ) {
    return Number(defaultValue) || 0;
  }

  const cleanedValue =
    typeof value === 'string'
      ? value.replace(/,/g, '').trim()
      : value;

  const numberValue = Number(cleanedValue);

  return Number.isFinite(numberValue)
    ? numberValue
    : Number(defaultValue) || 0;
}

/**
 * Validates a positive number.
 */
function inventoryPositiveNumber_(
  value,
  fieldName
) {
  const numberValue =
    inventoryToNumber_(value, NaN);

  if (
    !Number.isFinite(numberValue) ||
    numberValue <= 0
  ) {
    throw new Error(
      (fieldName || 'Value') +
      ' must be greater than zero.'
    );
  }

  return inventoryRoundQuantity_(
    numberValue
  );
}

/**
 * Validates a non-negative number.
 */
function inventoryNonNegativeNumber_(
  value,
  fieldName
) {
  if (
    value === '' ||
    value === null ||
    typeof value === 'undefined'
  ) {
    return 0;
  }

  const numberValue =
    inventoryToNumber_(value, NaN);

  if (
    !Number.isFinite(numberValue) ||
    numberValue < 0
  ) {
    throw new Error(
      (fieldName || 'Value') +
      ' must be zero or greater.'
    );
  }

  return inventoryRoundMoney_(
    numberValue
  );
}

/**
 * Rounds quantity to three decimal places.
 */
function inventoryRoundQuantity_(value) {
  return Math.round(
    inventoryToNumber_(value, 0) * 1000
  ) / 1000;
}

/**
 * Rounds money to two decimal places.
 */
function inventoryRoundMoney_(value) {
  return Math.round(
    inventoryToNumber_(value, 0) * 100
  ) / 100;
}

/**
 * Parses a date safely.
 */
function inventoryParseDate_(value) {
  if (!value) {
    return null;
  }

  if (
    Object.prototype.toString.call(value) ===
      '[object Date]' &&
    !isNaN(value.getTime())
  ) {
    return value;
  }

  const parsedDate = new Date(value);

  return isNaN(parsedDate.getTime())
    ? null
    : parsedDate;
}

/**
 * Returns timestamp value for comparison.
 */
function inventoryDateValue_(value) {
  const date = inventoryParseDate_(value);

  return date ? date.getTime() : 0;
}

/**
 * Formats date-time safely.
 */
function inventoryFormatDateTime_(value) {
  const date = inventoryParseDate_(value);

  if (!date) {
    return '';
  }

  return Utilities.formatDate(
    date,
    APP_CONFIG.TIME_ZONE || 'Asia/Kolkata',
    'dd-MMM-yyyy hh:mm a'
  );
}

/**
 * Creates test free inventory inward.
 *
 * Requires SKU000001 to exist and be ACTIVE.
 */
function testFreeInventoryInward() {
  const result = applyInventoryMovement({
    inventoryType: 'FREE INVENTORY',
    movementType: 'INWARD',
    skuId: 'SKU000001',
    quantity: 25,
    rate: 1250,
    referenceId: 'TEST-FREE-INWARD',
    remarks: 'Created from InventoryService test.'
  });

  Logger.log(
    JSON.stringify(result, null, 2)
  );

  return result;
}

/**
 * Creates test customized inventory inward.
 *
 * Requires:
 * - PRJ000001
 * - SKU000001
 */
function testCustomizeInventoryInward() {
  const result = applyInventoryMovement({
    inventoryType: 'CUSTOMIZE INVENTORY',
    movementType: 'INWARD',
    projectId: 'PRJ000005',
    skuId: 'SKU000001',
    quantity: 10,
    rate: 1250,
    referenceId: 'TEST-CUSTOM-INWARD',
    remarks: 'Created from InventoryService test.'
  });

  Logger.log(
    JSON.stringify(result, null, 2)
  );

  return result;
}

/**
 * Tests inventory listing and summary.
 */
function testInventoryService() {
  const result = {
    freeInventory: getFreeInventory({
      pageNumber: 1,
      pageSize: 10
    }),

    customizeInventory:
      getCustomizeInventory({
        pageNumber: 1,
        pageSize: 10
      }),

    lowStock: getLowStockInventory(),

    eligibleCustomizeInventory:
      getEligibleCustomizeInventory({
        pageNumber: 1,
        pageSize: 10
      }),

    summary: getInventorySummary()
  };

  Logger.log(
    JSON.stringify(result, null, 2)
  );

  return result;
}


function getInventoryStatus_(
  inventoryType,
  availableQty,
  reservedQty
) {
  const isFree =
    inventoryType ===
    INVENTORY_SERVICE_CONFIG.INVENTORY_TYPE.FREE;

  if (isFree) {
    return availableQty > 0
      ? 'AVAILABLE'
      : 'OUT OF STOCK';
  }

  if (availableQty <= 0 && reservedQty <= 0) {
    return 'CONSUMED';
  }

  if (availableQty > 0 && reservedQty > 0) {
    return 'PARTIALLY USED';
  }

  return 'AVAILABLE';
}




function resetInventorySheetsForTesting() {
  const freeSheet = getSystemSheet('Free_Inventory');
  const customizeSheet = getSystemSheet('Customize_Inventory');

  const freeHeaders = [
    'Inventory ID',
    'SKU ID',
    'SKU Code',
    'SKU Name',
    'Category Name',
    'UOM',
    'Available Qty',
    'Reserved Qty',
    'Damaged Qty',
    'Minimum Stock',
    'Reorder Level',
    'Average Rate',
    'Last Inward Date',
    'Last Outward Date',
    'Last Movement Date',
    'Status',
    'Remarks',
    'Created By',
    'Created At',
    'Updated By',
    'Updated At'
  ];

  const customizeHeaders = [
    'Inventory ID',
    'Project ID',
    'Project Code',
    'Project Name',
    'SKU ID',
    'SKU Code',
    'SKU Name',
    'Category Name',
    'UOM',
    'Available Qty',
    'Reserved Qty',
    'Damaged Qty',
    'Average Rate',
    'Last Inward Date',
    'Last Outward Date',
    'Last Movement Date',
    'Eligible for Free',
    'Eligibility Date',
    'Status',
    'Remarks',
    'Created By',
    'Created At',
    'Updated By',
    'Updated At'
  ];

  resetSingleInventorySheet_(freeSheet, freeHeaders);
  resetSingleInventorySheet_(customizeSheet, customizeHeaders);

  applyFreeInventoryValidation_(freeSheet);
  applyCustomizeInventoryValidation_(customizeSheet);

  SpreadsheetApp.flush();

  Logger.log(
    'Free_Inventory and Customize_Inventory reset successfully.'
  );
}

function resetSingleInventorySheet_(sheet, headers) {
  const requiredColumns = headers.length;

  if (sheet.getMaxColumns() < requiredColumns) {
    sheet.insertColumnsAfter(
      sheet.getMaxColumns(),
      requiredColumns - sheet.getMaxColumns()
    );
  }

  sheet.clearContents();
  sheet.clearFormats();
  
  sheet.getDataRange().clearDataValidations();

  sheet
    .getRange(1, 1, 1, requiredColumns)
    .setValues([headers])
    .setFontWeight('bold');

  sheet.setFrozenRows(1);
}

function applyFreeInventoryValidation_(sheet) {
  const statusRule = SpreadsheetApp
    .newDataValidation()
    .requireValueInList(
      [
        'AVAILABLE',
        'OUT OF STOCK',
        'INACTIVE'
      ],
      true
    )
    .setAllowInvalid(false)
    .build();

  sheet
    .getRange(
      2,
      16,
      Math.max(sheet.getMaxRows() - 1, 1),
      1
    )
    .setDataValidation(statusRule);
}

function applyCustomizeInventoryValidation_(sheet) {
  const statusRule = SpreadsheetApp
    .newDataValidation()
    .requireValueInList(
      [
        'AVAILABLE',
        'PARTIALLY USED',
        'CONSUMED',
        'TRANSFER PENDING',
        'APPROVED FOR TRANSFER',
        'CONVERTED TO FREE',
        'INACTIVE'
      ],
      true
    )
    .setAllowInvalid(false)
    .build();

  const yesNoRule = SpreadsheetApp
    .newDataValidation()
    .requireValueInList(
      ['YES', 'NO'],
      true
    )
    .setAllowInvalid(false)
    .build();

  sheet
    .getRange(
      2,
      17,
      Math.max(sheet.getMaxRows() - 1, 1),
      1
    )
    .setDataValidation(yesNoRule);

  sheet
    .getRange(
      2,
      19,
      Math.max(sheet.getMaxRows() - 1, 1),
      1
    )
    .setDataValidation(statusRule);
}



function diagnoseInventoryColumnMapping() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const freeSheet = ss.getSheetByName(
    getFreeInventorySheetName_()
  );

  const customizeSheet = ss.getSheetByName(
    getCustomizeInventorySheetName_()
  );

  const skuSheet = ss.getSheetByName(
    APP_CONFIG.SHEETS.SKU_MASTER || 'SKU_Master'
  );

  const freeLastColumn = freeSheet.getLastColumn();
  const customizeLastColumn =
    customizeSheet.getLastColumn();
  const skuLastColumn = skuSheet.getLastColumn();

  const result = {
    freeInventory: {
      headers: freeSheet
        .getRange(1, 1, 1, freeLastColumn)
        .getDisplayValues()[0],

      rawRow: freeSheet.getLastRow() >= 2
        ? freeSheet
            .getRange(2, 1, 1, freeLastColumn)
            .getValues()[0]
            .map(diagnosticValue_)
        : []
    },

    customizeInventory: {
      headers: customizeSheet
        .getRange(
          1,
          1,
          1,
          customizeLastColumn
        )
        .getDisplayValues()[0],

      rawRow: customizeSheet.getLastRow() >= 2
        ? customizeSheet
            .getRange(
              2,
              1,
              1,
              customizeLastColumn
            )
            .getValues()[0]
            .map(diagnosticValue_)
        : []
    },

    skuMaster: {
      headers: skuSheet
        .getRange(1, 1, 1, skuLastColumn)
        .getDisplayValues()[0],

      rawRow: skuSheet.getLastRow() >= 2
        ? skuSheet
            .getRange(2, 1, 1, skuLastColumn)
            .getValues()[0]
            .map(diagnosticValue_)
        : []
    }
  };

  Logger.log(
    JSON.stringify(result, null, 2)
  );

  return result;
}

function diagnosticValue_(value) {
  if (
    Object.prototype.toString.call(value) ===
    '[object Date]'
  ) {
    return {
      type: 'DATE',
      value: Utilities.formatDate(
        value,
        APP_CONFIG.TIME_ZONE || 'Asia/Kolkata',
        'dd-MMM-yyyy HH:mm:ss'
      ),
      timestamp: value.getTime()
    };
  }

  return {
    type: typeof value,
    value: value
  };
}


/**
 * Fixes inventory sheet column formats.
 */
function fixInventoryColumnFormats() {
  const freeSheet = getSystemSheet(
    getFreeInventorySheetName_()
  );

  const customizeSheet = getSystemSheet(
    getCustomizeInventorySheetName_()
  );

  const freeRows = Math.max(
    freeSheet.getMaxRows() - 1,
    1
  );

  const customizeRows = Math.max(
    customizeSheet.getMaxRows() - 1,
    1
  );

  // FREE INVENTORY numeric columns.
  freeSheet
    .getRange(2, 7, freeRows, 5)
    .setNumberFormat('0.000');

  // Free Average Rate.
  freeSheet
    .getRange(2, 12, freeRows, 1)
    .setNumberFormat('0.00');

  // Free date columns.
  freeSheet
    .getRange(2, 13, freeRows, 3)
    .setNumberFormat('dd-mmm-yyyy hh:mm AM/PM');

  freeSheet
    .getRange(2, 19, freeRows, 1)
    .setNumberFormat('dd-mmm-yyyy hh:mm AM/PM');

  freeSheet
    .getRange(2, 21, freeRows, 1)
    .setNumberFormat('dd-mmm-yyyy hh:mm AM/PM');

  // CUSTOMIZE INVENTORY quantity columns.
  customizeSheet
    .getRange(2, 10, customizeRows, 3)
    .setNumberFormat('0.000');

  // Customize Average Rate.
  customizeSheet
    .getRange(2, 13, customizeRows, 1)
    .setNumberFormat('0.00');

  // Customize movement dates.
  customizeSheet
    .getRange(2, 14, customizeRows, 3)
    .setNumberFormat('dd-mmm-yyyy hh:mm AM/PM');

  // Eligibility Date.
  customizeSheet
    .getRange(2, 18, customizeRows, 1)
    .setNumberFormat('dd-mmm-yyyy hh:mm AM/PM');

  // Created At.
  customizeSheet
    .getRange(2, 22, customizeRows, 1)
    .setNumberFormat('dd-mmm-yyyy hh:mm AM/PM');

  // Updated At.
  customizeSheet
    .getRange(2, 24, customizeRows, 1)
    .setNumberFormat('dd-mmm-yyyy hh:mm AM/PM');

  SpreadsheetApp.flush();

  Logger.log(
    'Inventory column formats fixed successfully.'
  );
}
