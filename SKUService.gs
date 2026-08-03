/**
 * PROJECT INVENTORY MANAGEMENT SYSTEM
 * File: SKUService.gs
 *
 * Handles SKU master operations:
 * - Create SKU
 * - Update SKU
 * - Activate or deactivate SKU
 * - Search and list SKUs
 * - SKU dropdown
 * - Category validation
 * - Duplicate SKU validation
 * - Transaction logging
 */

const SKU_SERVICE_CONFIG = Object.freeze({
  SHEET_NAME: 'SKU_Master',

  STATUS: Object.freeze({
    ACTIVE: 'ACTIVE',
    INACTIVE: 'INACTIVE'
  }),

  DEFAULT_STATUS: 'ACTIVE',
  SKU_ID_PREFIX: 'SKU'
});

/**
 * Returns the configured SKU master sheet name.
 */
function getSkuSheetName_() {
  if (
    typeof APP_CONFIG !== 'undefined' &&
    APP_CONFIG.SHEETS &&
    APP_CONFIG.SHEETS.SKU_MASTER
  ) {
    return APP_CONFIG.SHEETS.SKU_MASTER;
  }

  return SKU_SERVICE_CONFIG.SHEET_NAME;
}

/**
 * Returns the configured category master sheet name.
 */
function getCategorySheetName_() {
  if (
    typeof APP_CONFIG !== 'undefined' &&
    APP_CONFIG.SHEETS &&
    APP_CONFIG.SHEETS.CATEGORY_MASTER
  ) {
    return APP_CONFIG.SHEETS.CATEGORY_MASTER;
  }

  return 'Category_Master';
}

/**
 * Creates a new SKU.
 *
 * Expected skuData:
 * {
 *   skuCode: '',
 *   skuName: '',
 *   categoryId: '',
 *   categoryName: '',
 *   brand: '',
 *   model: '',
 *   description: '',
 *   uom: '',
 *   minimumStock: 0,
 *   reorderLevel: 0,
 *   standardRate: 0,
 *   gstPercent: 18,
 *   status: 'ACTIVE',
 *   remarks: ''
 * }
 */
function createSku(skuData) {
  return safeExecute_(function () {
    const session = requireRole_([
      APP_CONFIG.USER_ROLES.ADMIN,
      APP_CONFIG.USER_ROLES.HOD
    ]);

    skuData = skuData || {};

    validateRequiredFields_(
      skuData,
      [
        'skuName',
        'categoryName',
        'uom'
      ]
    );

    const skuName = normalizeText_(skuData.skuName);
    const categoryId = normalizeUpper_(skuData.categoryId);
    const categoryName = normalizeText_(skuData.categoryName);
    const brand = normalizeText_(skuData.brand);
    const model = normalizeText_(skuData.model);
    const description = normalizeText_(skuData.description);
    const uom = normalizeUpper_(skuData.uom);

    const minimumStock = validateNonNegativeNumber_(
      skuData.minimumStock,
      'Minimum stock'
    );

    const reorderLevel = validateNonNegativeNumber_(
      skuData.reorderLevel,
      'Reorder level'
    );

    const standardRate = validateNonNegativeNumber_(
      skuData.standardRate,
      'Standard rate'
    );

    const gstPercent = validatePercentage_(
      typeof skuData.gstPercent === 'undefined'
        ? 18
        : skuData.gstPercent,
      'GST percentage'
    );

    const status = normalizeSkuStatus_(
      skuData.status || SKU_SERVICE_CONFIG.DEFAULT_STATUS
    );

    const remarks = normalizeText_(skuData.remarks);

    if (skuName.length < 2) {
      throw new Error(
        'SKU name must contain at least 2 characters.'
      );
    }

    if (!categoryName) {
      throw new Error('Category name is required.');
    }

    if (!uom) {
      throw new Error('UOM is required.');
    }

    validateSkuCategory_(categoryId, categoryName);

    if (
      isDuplicateSkuName_(
        skuName,
        brand,
        model,
        null
      )
    ) {
      throw new Error(
        'An SKU with the same name, brand and model already exists.'
      );
    }

    let skuCode = normalizeUpper_(skuData.skuCode);

    if (skuCode) {
      if (valueExists_(
        getSkuSheetName_(),
        'SKU Code',
        skuCode,
        false
      )) {
        throw new Error(
          'SKU code already exists: ' + skuCode
        );
      }
    } else {
      skuCode = generateSkuCode_(
        categoryName,
        skuName
      );
    }

    const skuId = generateNextId_(
      'SKU',
      SKU_SERVICE_CONFIG.SKU_ID_PREFIX
    );

    const now = new Date();

    appendObjectRow_(
      getSkuSheetName_(),
      {
        'SKU ID': skuId,
        'SKU Code': skuCode,
        'SKU Name': skuName,
        'Category ID': categoryId,
        'Category Name': categoryName,
        'Brand': brand,
        'Model': model,
        'Description': description,
        'UOM': uom,
        'Minimum Stock': minimumStock,
        'Reorder Level': reorderLevel,
        'Standard Rate': standardRate,
        'GST Percent': gstPercent,
        'Status': status,
        'Remarks': remarks,
        'Created By': session.email,
        'Created At': now,
        'Updated By': session.email,
        'Updated At': now
      }
    );

    addTransactionLog_(
      session.email,
      'SKU',
      'CREATE',
      skuId,
      {
        skuCode: skuCode,
        skuName: skuName,
        categoryName: categoryName,
        brand: brand,
        model: model,
        uom: uom,
        status: status
      }
    );

    return successResponse_(
      'SKU created successfully.',
      {
        skuId: skuId,
        skuCode: skuCode,
        skuName: skuName,
        categoryName: categoryName,
        status: status
      }
    );
  }, 'Unable to create SKU.');
}

/**
 * Updates an existing SKU.
 */
function updateSku(skuId, skuData) {
  return safeExecute_(function () {
    const session = requireRole_([
      APP_CONFIG.USER_ROLES.ADMIN,
      APP_CONFIG.USER_ROLES.HOD
    ]);

    skuData = skuData || {};

    const sku = getSkuRecordById_(skuId);

    if (!sku) {
      throw new Error('SKU not found: ' + skuId);
    }

    const updatedName =
      Object.prototype.hasOwnProperty.call(
        skuData,
        'skuName'
      )
        ? normalizeText_(skuData.skuName)
        : normalizeText_(sku['SKU Name']);

    const updatedBrand =
      Object.prototype.hasOwnProperty.call(
        skuData,
        'brand'
      )
        ? normalizeText_(skuData.brand)
        : normalizeText_(sku['Brand']);

    const updatedModel =
      Object.prototype.hasOwnProperty.call(
        skuData,
        'model'
      )
        ? normalizeText_(skuData.model)
        : normalizeText_(sku['Model']);

    const updatedCategoryId =
      Object.prototype.hasOwnProperty.call(
        skuData,
        'categoryId'
      )
        ? normalizeUpper_(skuData.categoryId)
        : normalizeUpper_(sku['Category ID']);

    const updatedCategoryName =
      Object.prototype.hasOwnProperty.call(
        skuData,
        'categoryName'
      )
        ? normalizeText_(skuData.categoryName)
        : normalizeText_(sku['Category Name']);

    if (!updatedName) {
      throw new Error('SKU name cannot be blank.');
    }

    if (!updatedCategoryName) {
      throw new Error('Category name cannot be blank.');
    }

    validateSkuCategory_(
      updatedCategoryId,
      updatedCategoryName
    );

    if (
      isDuplicateSkuName_(
        updatedName,
        updatedBrand,
        updatedModel,
        skuId
      )
    ) {
      throw new Error(
        'Another SKU with the same name, brand and model already exists.'
      );
    }

    const updateData = {
      'Updated By': session.email,
      'Updated At': new Date()
    };

    if (
      Object.prototype.hasOwnProperty.call(
        skuData,
        'skuCode'
      )
    ) {
      const skuCode = normalizeUpper_(skuData.skuCode);

      if (!skuCode) {
        throw new Error('SKU code cannot be blank.');
      }

      const existingSku = getSkuRecordByCode_(skuCode);

      if (
        existingSku &&
        normalizeUpper_(existingSku['SKU ID']) !==
          normalizeUpper_(skuId)
      ) {
        throw new Error(
          'Another SKU already uses this SKU code.'
        );
      }

      updateData['SKU Code'] = skuCode;
    }

    if (
      Object.prototype.hasOwnProperty.call(
        skuData,
        'skuName'
      )
    ) {
      updateData['SKU Name'] = updatedName;
    }

    if (
      Object.prototype.hasOwnProperty.call(
        skuData,
        'categoryId'
      )
    ) {
      updateData['Category ID'] = updatedCategoryId;
    }

    if (
      Object.prototype.hasOwnProperty.call(
        skuData,
        'categoryName'
      )
    ) {
      updateData['Category Name'] =
        updatedCategoryName;
    }

    if (
      Object.prototype.hasOwnProperty.call(
        skuData,
        'brand'
      )
    ) {
      updateData['Brand'] = updatedBrand;
    }

    if (
      Object.prototype.hasOwnProperty.call(
        skuData,
        'model'
      )
    ) {
      updateData['Model'] = updatedModel;
    }

    if (
      Object.prototype.hasOwnProperty.call(
        skuData,
        'description'
      )
    ) {
      updateData['Description'] = normalizeText_(
        skuData.description
      );
    }

    if (
      Object.prototype.hasOwnProperty.call(
        skuData,
        'uom'
      )
    ) {
      const uom = normalizeUpper_(skuData.uom);

      if (!uom) {
        throw new Error('UOM cannot be blank.');
      }

      updateData['UOM'] = uom;
    }

    if (
      Object.prototype.hasOwnProperty.call(
        skuData,
        'minimumStock'
      )
    ) {
      updateData['Minimum Stock'] =
        validateNonNegativeNumber_(
          skuData.minimumStock,
          'Minimum stock'
        );
    }

    if (
      Object.prototype.hasOwnProperty.call(
        skuData,
        'reorderLevel'
      )
    ) {
      updateData['Reorder Level'] =
        validateNonNegativeNumber_(
          skuData.reorderLevel,
          'Reorder level'
        );
    }

    if (
      Object.prototype.hasOwnProperty.call(
        skuData,
        'standardRate'
      )
    ) {
      updateData['Standard Rate'] =
        validateNonNegativeNumber_(
          skuData.standardRate,
          'Standard rate'
        );
    }

    if (
      Object.prototype.hasOwnProperty.call(
        skuData,
        'gstPercent'
      )
    ) {
      updateData['GST Percent'] =
        validatePercentage_(
          skuData.gstPercent,
          'GST percentage'
        );
    }

    if (
      Object.prototype.hasOwnProperty.call(
        skuData,
        'status'
      )
    ) {
      updateData['Status'] = normalizeSkuStatus_(
        skuData.status
      );
    }

    if (
      Object.prototype.hasOwnProperty.call(
        skuData,
        'remarks'
      )
    ) {
      updateData['Remarks'] = normalizeText_(
        skuData.remarks
      );
    }

    updateObjectRow_(
      getSkuSheetName_(),
      sku._rowNumber,
      updateData
    );

    addTransactionLog_(
      session.email,
      'SKU',
      'UPDATE',
      normalizeUpper_(skuId),
      updateData
    );

    return successResponse_(
      'SKU updated successfully.',
      getSkuByIdData_(skuId)
    );
  }, 'Unable to update SKU.');
}

/**
 * Returns one SKU by SKU ID.
 */
function getSkuById(skuId) {
  return safeExecute_(function () {
    requireAuthenticatedUser_();

    const sku = getSkuRecordById_(skuId);

    if (!sku) {
      throw new Error('SKU not found: ' + skuId);
    }

    return successResponse_(
      'SKU loaded successfully.',
      mapSkuRecord_(sku)
    );
  }, 'Unable to load SKU.');
}

/**
 * Returns an internal SKU record by SKU ID.
 */
function getSkuRecordById_(skuId) {
  const normalizedId = normalizeUpper_(skuId);

  if (!normalizedId) {
    return null;
  }

  const records = getSheetObjects_(
    getSkuSheetName_()
  );

  for (let index = 0; index < records.length; index++) {
    if (
      normalizeUpper_(records[index]['SKU ID']) ===
      normalizedId
    ) {
      return records[index];
    }
  }

  return null;
}

/**
 * Returns an internal SKU record by SKU Code.
 */
function getSkuRecordByCode_(skuCode) {
  const normalizedCode = normalizeUpper_(skuCode);

  if (!normalizedCode) {
    return null;
  }

  const records = getSheetObjects_(
    getSkuSheetName_()
  );

  for (let index = 0; index < records.length; index++) {
    if (
      normalizeUpper_(records[index]['SKU Code']) ===
      normalizedCode
    ) {
      return records[index];
    }
  }

  return null;
}

/**
 * Returns mapped SKU data internally.
 */
function getSkuByIdData_(skuId) {
  const sku = getSkuRecordById_(skuId);

  return sku ? mapSkuRecord_(sku) : null;
}

/**
 * Returns SKUs with filters and pagination.
 *
 * filters:
 * {
 *   search: '',
 *   categoryName: '',
 *   brand: '',
 *   status: '',
 *   pageNumber: 1,
 *   pageSize: 20
 * }
 */
function getSkus(filters) {
  return safeExecute_(function () {
    requireAuthenticatedUser_();

    filters = filters || {};

    const searchText = normalizeLower_(filters.search);
    const categoryFilter = normalizeLower_(
      filters.categoryName
    );
    const brandFilter = normalizeLower_(filters.brand);
    const statusFilter = normalizeUpper_(filters.status);

    let records = getSheetObjects_(
      getSkuSheetName_()
    );

    records = records.filter(function (sku) {
      if (
        categoryFilter &&
        normalizeLower_(sku['Category Name']) !==
          categoryFilter
      ) {
        return false;
      }

      if (
        brandFilter &&
        normalizeLower_(sku['Brand']) !== brandFilter
      ) {
        return false;
      }

      if (
        statusFilter &&
        normalizeUpper_(sku['Status']) !== statusFilter
      ) {
        return false;
      }

      if (searchText) {
        const searchableText = [
          sku['SKU ID'],
          sku['SKU Code'],
          sku['SKU Name'],
          sku['Category Name'],
          sku['Brand'],
          sku['Model'],
          sku['Description'],
          sku['UOM'],
          sku['Status']
        ]
          .map(normalizeLower_)
          .join(' ');

        if (searchableText.indexOf(searchText) === -1) {
          return false;
        }
      }

      return true;
    });

    records.sort(function (first, second) {
      return normalizeText_(first['SKU Name'])
        .localeCompare(
          normalizeText_(second['SKU Name'])
        );
    });

    const mappedRecords = records.map(
      mapSkuRecord_
    );

    const paginated = paginateRecords_(
      mappedRecords,
      filters.pageNumber || 1,
      filters.pageSize || 20
    );

    return successResponse_(
      'SKUs loaded successfully.',
      paginated
    );
  }, 'Unable to load SKUs.');
}

/**
 * Returns active SKUs for dropdowns.
 */
function getActiveSkuDropdown() {
  return safeExecute_(function () {
    requireAuthenticatedUser_();

    const records = getSheetObjects_(
      getSkuSheetName_()
    )
      .filter(function (sku) {
        return (
          normalizeUpper_(sku['Status']) ===
          SKU_SERVICE_CONFIG.STATUS.ACTIVE
        );
      })
      .sort(function (first, second) {
        return normalizeText_(first['SKU Name'])
          .localeCompare(
            normalizeText_(second['SKU Name'])
          );
      })
      .map(function (sku) {
        const skuCode = normalizeUpper_(
          sku['SKU Code']
        );

        const skuName = normalizeText_(
          sku['SKU Name']
        );

        return {
          value: normalizeUpper_(sku['SKU ID']),
          label:
            skuCode +
            ' - ' +
            skuName,
          skuId: normalizeUpper_(sku['SKU ID']),
          skuCode: skuCode,
          skuName: skuName,
          categoryId: normalizeUpper_(
            sku['Category ID']
          ),
          categoryName: normalizeText_(
            sku['Category Name']
          ),
          brand: normalizeText_(sku['Brand']),
          model: normalizeText_(sku['Model']),
          description: normalizeText_(
            sku['Description']
          ),
          uom: normalizeUpper_(sku['UOM']),
          standardRate: toNumber_(
            sku['Standard Rate'],
            0
          ),
          gstPercent: toNumber_(
            sku['GST Percent'],
            0
          )
        };
      });

    return successResponse_(
      'Active SKU dropdown loaded successfully.',
      records
    );
  }, 'Unable to load active SKU dropdown.');
}

/**
 * Changes SKU status.
 */
function changeSkuStatus(skuId, newStatus, remarks) {
  return safeExecute_(function () {
    const session = requireRole_([
      APP_CONFIG.USER_ROLES.ADMIN,
      APP_CONFIG.USER_ROLES.HOD
    ]);

    const sku = getSkuRecordById_(skuId);

    if (!sku) {
      throw new Error('SKU not found: ' + skuId);
    }

    const oldStatus = normalizeSkuStatus_(
      sku['Status']
    );

    const normalizedStatus = normalizeSkuStatus_(
      newStatus
    );

    if (oldStatus === normalizedStatus) {
      throw new Error(
        'SKU is already in ' +
        normalizedStatus +
        ' status.'
      );
    }

    updateObjectRow_(
      getSkuSheetName_(),
      sku._rowNumber,
      {
        'Status': normalizedStatus,
        'Remarks': normalizeText_(
          remarks || sku['Remarks']
        ),
        'Updated By': session.email,
        'Updated At': new Date()
      }
    );

    addTransactionLog_(
      session.email,
      'SKU',
      'STATUS CHANGE',
      normalizeUpper_(skuId),
      {
        oldStatus: oldStatus,
        newStatus: normalizedStatus,
        remarks: normalizeText_(remarks)
      }
    );

    return successResponse_(
      'SKU status changed successfully.',
      {
        skuId: normalizeUpper_(skuId),
        oldStatus: oldStatus,
        newStatus: normalizedStatus
      }
    );
  }, 'Unable to change SKU status.');
}

/**
 * Activates an SKU.
 */
function activateSku(skuId, remarks) {
  return changeSkuStatus(
    skuId,
    SKU_SERVICE_CONFIG.STATUS.ACTIVE,
    remarks
  );
}

/**
 * Deactivates an SKU.
 */
function deactivateSku(skuId, remarks) {
  return changeSkuStatus(
    skuId,
    SKU_SERVICE_CONFIG.STATUS.INACTIVE,
    remarks
  );
}

/**
 * Validates SKU status.
 */
function normalizeSkuStatus_(status) {
  const normalizedStatus = normalizeUpper_(status);

  const validStatuses = [
    SKU_SERVICE_CONFIG.STATUS.ACTIVE,
    SKU_SERVICE_CONFIG.STATUS.INACTIVE
  ];

  if (
    validStatuses.indexOf(normalizedStatus) === -1
  ) {
    throw new Error(
      'Invalid SKU status: ' + normalizedStatus
    );
  }

  return normalizedStatus;
}

/**
 * Validates a non-negative number.
 */
function validateNonNegativeNumber_(value, fieldName) {
  if (
    value === '' ||
    value === null ||
    typeof value === 'undefined'
  ) {
    return 0;
  }

  const numberValue = Number(value);

  if (
    !Number.isFinite(numberValue) ||
    numberValue < 0
  ) {
    throw new Error(
      (fieldName || 'Value') +
      ' must be zero or greater.'
    );
  }

  return roundTwo_(numberValue);
}

/**
 * Validates percentage from 0 to 100.
 */
function validatePercentage_(value, fieldName) {
  const numberValue = Number(value);

  if (
    !Number.isFinite(numberValue) ||
    numberValue < 0 ||
    numberValue > 100
  ) {
    throw new Error(
      (fieldName || 'Percentage') +
      ' must be between 0 and 100.'
    );
  }

  return roundTwo_(numberValue);
}

/**
 * Validates the selected category.
 *
 * Category_Master supported headers:
 * Category ID
 * Category Name
 * Status
 */
function validateSkuCategory_(
  categoryId,
  categoryName
) {
  const sheet = getSystemSheet(
    getCategorySheetName_()
  );

  if (sheet.getLastRow() < 2) {
    throw new Error(
      'No categories are available in Category_Master.'
    );
  }

  const categories = getSheetObjects_(
    getCategorySheetName_()
  );

  const normalizedId = normalizeUpper_(categoryId);
  const normalizedName = normalizeLower_(
    categoryName
  );

  const category = categories.find(function (record) {
    const idMatches =
      normalizedId &&
      normalizeUpper_(record['Category ID']) ===
        normalizedId;

    const nameMatches =
      normalizeLower_(record['Category Name']) ===
      normalizedName;

    return normalizedId
      ? idMatches && nameMatches
      : nameMatches;
  });

  if (!category) {
    throw new Error(
      'Category was not found in Category_Master: ' +
      categoryName
    );
  }

  const status = normalizeUpper_(
    category['Status']
  );

  if (
    status &&
    status !== SKU_SERVICE_CONFIG.STATUS.ACTIVE
  ) {
    throw new Error(
      'Selected category is inactive: ' +
      categoryName
    );
  }

  return true;
}

/**
 * Checks duplicate SKU combination.
 */
function isDuplicateSkuName_(
  skuName,
  brand,
  model,
  excludeSkuId
) {
  const normalizedName = normalizeLower_(skuName);
  const normalizedBrand = normalizeLower_(brand);
  const normalizedModel = normalizeLower_(model);
  const normalizedExcludeId = normalizeUpper_(
    excludeSkuId
  );

  const records = getSheetObjects_(
    getSkuSheetName_()
  );

  return records.some(function (sku) {
    const currentId = normalizeUpper_(
      sku['SKU ID']
    );

    if (
      normalizedExcludeId &&
      currentId === normalizedExcludeId
    ) {
      return false;
    }

    return (
      normalizeLower_(sku['SKU Name']) ===
        normalizedName &&
      normalizeLower_(sku['Brand']) ===
        normalizedBrand &&
      normalizeLower_(sku['Model']) ===
        normalizedModel
    );
  });
}

/**
 * Generates a unique SKU code.
 *
 * Example:
 * Category: Lighting
 * SKU Name: Smart Dimmer
 * Result: LIG-SMA-001
 */
function generateSkuCode_(
  categoryName,
  skuName
) {
  const categoryPart = createSkuCodePart_(
    categoryName,
    3
  );

  const skuPart = createSkuCodePart_(
    skuName,
    3
  );

  const baseCode =
    categoryPart + '-' + skuPart;

  const existingRecords = getSheetObjects_(
    getSkuSheetName_()
  );

  let highestNumber = 0;

  existingRecords.forEach(function (record) {
    const currentCode = normalizeUpper_(
      record['SKU Code']
    );

    if (
      currentCode.indexOf(baseCode + '-') !== 0
    ) {
      return;
    }

    const match = currentCode.match(/-(\d+)$/);

    if (match) {
      highestNumber = Math.max(
        highestNumber,
        Number(match[1]) || 0
      );
    }
  });

  return (
    baseCode +
    '-' +
    String(highestNumber + 1).padStart(3, '0')
  );
}

/**
 * Creates one part of an SKU code.
 */
function createSkuCodePart_(value, length) {
  const cleanedValue = normalizeUpper_(value)
    .replace(/[^A-Z0-9]/g, '');

  if (!cleanedValue) {
    return 'GEN';
  }

  return cleanedValue
    .substring(0, length || 3)
    .padEnd(length || 3, 'X');
}

/**
 * Maps raw SKU sheet data.
 */
function mapSkuRecord_(sku) {
  return {
    skuId: normalizeUpper_(sku['SKU ID']),
    skuCode: normalizeUpper_(sku['SKU Code']),
    skuName: normalizeText_(sku['SKU Name']),
    categoryId: normalizeUpper_(
      sku['Category ID']
    ),
    categoryName: normalizeText_(
      sku['Category Name']
    ),
    brand: normalizeText_(sku['Brand']),
    model: normalizeText_(sku['Model']),
    description: normalizeText_(
      sku['Description']
    ),
    uom: normalizeUpper_(sku['UOM']),
    minimumStock: toNumber_(
      sku['Minimum Stock'],
      0
    ),
    reorderLevel: toNumber_(
      sku['Reorder Level'],
      0
    ),
    standardRate: toNumber_(
      sku['Standard Rate'],
      0
    ),
    gstPercent: toNumber_(
      sku['GST Percent'],
      0
    ),
    status: normalizeUpper_(sku['Status']),
    remarks: normalizeText_(sku['Remarks']),
    createdBy: normalizeLower_(
      sku['Created By']
    ),
    createdAt: formatDateTime_(
      sku['Created At']
    ),
    updatedBy: normalizeLower_(
      sku['Updated By']
    ),
    updatedAt: formatDateTime_(
      sku['Updated At']
    )
  };
}

/**
 * Returns SKU summary.
 */
function getSkuSummary() {
  return safeExecute_(function () {
    requireAuthenticatedUser_();

    const records = getSheetObjects_(
      getSkuSheetName_()
    );

    const summary = {
      total: records.length,
      active: 0,
      inactive: 0,
      lowStockConfigured: 0,
      categories: 0,
      brands: 0
    };

    const categories = {};
    const brands = {};

    records.forEach(function (sku) {
      const status = normalizeUpper_(
        sku['Status']
      );

      if (
        status === SKU_SERVICE_CONFIG.STATUS.ACTIVE
      ) {
        summary.active++;
      } else if (
        status === SKU_SERVICE_CONFIG.STATUS.INACTIVE
      ) {
        summary.inactive++;
      }

      if (
        toNumber_(sku['Minimum Stock'], 0) > 0 ||
        toNumber_(sku['Reorder Level'], 0) > 0
      ) {
        summary.lowStockConfigured++;
      }

      const category = normalizeUpper_(
        sku['Category Name']
      );

      const brand = normalizeUpper_(
        sku['Brand']
      );

      if (category) {
        categories[category] = true;
      }

      if (brand) {
        brands[brand] = true;
      }
    });

    summary.categories =
      Object.keys(categories).length;

    summary.brands =
      Object.keys(brands).length;

    return successResponse_(
      'SKU summary loaded successfully.',
      summary
    );
  }, 'Unable to load SKU summary.');
}

/**
 * Returns active category dropdown.
 */
function getActiveCategoryDropdown() {
  return safeExecute_(function () {
    requireAuthenticatedUser_();

    const categories = getSheetObjects_(
      getCategorySheetName_()
    )
      .filter(function (category) {
        const status = normalizeUpper_(
          category['Status']
        );

        return (
          !status ||
          status === SKU_SERVICE_CONFIG.STATUS.ACTIVE
        );
      })
      .sort(function (first, second) {
        return normalizeText_(
          first['Category Name']
        ).localeCompare(
          normalizeText_(second['Category Name'])
        );
      })
      .map(function (category) {
        return {
          value: normalizeUpper_(
            category['Category ID']
          ),
          label: normalizeText_(
            category['Category Name']
          ),
          categoryId: normalizeUpper_(
            category['Category ID']
          ),
          categoryName: normalizeText_(
            category['Category Name']
          )
        };
      });

    return successResponse_(
      'Active category dropdown loaded successfully.',
      categories
    );
  }, 'Unable to load category dropdown.');
}

/**
 * Ensures SKU_Master contains all required headers.
 *
 * This function does not delete existing data.
 * Missing headers are added after the last used column.
 */
function ensureSkuMasterHeaders() {
  return safeExecute_(function () {
    requireRole_(APP_CONFIG.USER_ROLES.ADMIN);

    const sheet = getSystemSheet(
      getSkuSheetName_()
    );

    const requiredHeaders = [
      'SKU ID',
      'SKU Code',
      'SKU Name',
      'Category ID',
      'Category Name',
      'Brand',
      'Model',
      'Description',
      'UOM',
      'Minimum Stock',
      'Reorder Level',
      'Standard Rate',
      'GST Percent',
      'Status',
      'Remarks',
      'Created By',
      'Created At',
      'Updated By',
      'Updated At'
    ];

    const lastColumn = Math.max(
      sheet.getLastColumn(),
      1
    );

    const existingHeaders = sheet
      .getRange(1, 1, 1, lastColumn)
      .getDisplayValues()[0]
      .map(normalizeText_);

    const missingHeaders = requiredHeaders.filter(
      function (header) {
        return existingHeaders.indexOf(header) === -1;
      }
    );

    if (missingHeaders.length > 0) {
      const startColumn =
        existingHeaders.filter(Boolean).length + 1;

      sheet
        .getRange(
          1,
          startColumn,
          1,
          missingHeaders.length
        )
        .setValues([missingHeaders]);
    }

    return successResponse_(
      'SKU master headers verified successfully.',
      {
        addedHeaders: missingHeaders,
        requiredHeaders: requiredHeaders
      }
    );
  }, 'Unable to verify SKU master headers.');
}

/**
 * Creates a test category when no category exists.
 */
function createTestCategoryForSku_() {
  const sheet = getSystemSheet(
    getCategorySheetName_()
  );

  const categories = getSheetObjects_(
    getCategorySheetName_()
  );

  const existingCategory = categories.find(
    function (category) {
      return (
        normalizeLower_(
          category['Category Name']
        ) === 'test category'
      );
    }
  );

  if (existingCategory) {
    return {
      categoryId: normalizeUpper_(
        existingCategory['Category ID']
      ),
      categoryName: normalizeText_(
        existingCategory['Category Name']
      )
    };
  }

  const categoryId = generateNextId_(
    'CATEGORY',
    'CAT'
  );

  appendObjectRow_(
    getCategorySheetName_(),
    {
      'Category ID': categoryId,
      'Category Name': 'Test Category',
      'Status': 'ACTIVE',
      'Remarks': 'Created for SKU service testing.',
      'Created By': getCurrentUserEmail_(),
      'Created At': new Date(),
      'Updated By': getCurrentUserEmail_(),
      'Updated At': new Date()
    }
  );

  return {
    categoryId: categoryId,
    categoryName: 'Test Category'
  };
}

/**
 * Creates one test SKU.
 */
function testCreateSku() {
  const category = createTestCategoryForSku_();

  const timestamp = Utilities.formatDate(
    new Date(),
    APP_CONFIG.TIME_ZONE,
    'yyyyMMdd-HHmmss'
  );

  const result = createSku({
    skuName: 'Test Inventory Item ' + timestamp,
    categoryId: category.categoryId,
    categoryName: category.categoryName,
    brand: 'Test Brand',
    model: 'Model ' + timestamp,
    description: 'Created from SKUService.gs test.',
    uom: 'NOS',
    minimumStock: 5,
    reorderLevel: 10,
    standardRate: 1250,
    gstPercent: 18,
    status: 'ACTIVE',
    remarks: 'SKU test record.'
  });

  Logger.log(JSON.stringify(result, null, 2));

  return result;
}

/**
 * Tests SKU listing, dropdown and summary.
 */
function testSkuService() {
  const result = {
    skus: getSkus({
      pageNumber: 1,
      pageSize: 10
    }),
    dropdown: getActiveSkuDropdown(),
    categories: getActiveCategoryDropdown(),
    summary: getSkuSummary()
  };

  Logger.log(JSON.stringify(result, null, 2));

  return result;
}



/**
 * Ensures Category_Master contains all required headers.
 * Existing data is not deleted.
 */
function ensureCategoryMasterHeaders() {
  return safeExecute_(function () {
    requireRole_(APP_CONFIG.USER_ROLES.ADMIN);

    const sheet = getSystemSheet(
      getCategorySheetName_()
    );

    const requiredHeaders = [
      'Category ID',
      'Category Name',
      'Status',
      'Remarks',
      'Created By',
      'Created At',
      'Updated By',
      'Updated At'
    ];

    const currentLastColumn = Math.max(
      sheet.getLastColumn(),
      1
    );

    const existingHeaders = sheet
      .getRange(1, 1, 1, currentLastColumn)
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

    sheet
  .getRange(1, 1, 1, sheet.getLastColumn())
  .setFontWeight('bold');

sheet.setFrozenRows(1);

    SpreadsheetApp.flush();

    return successResponse_(
      'Category master headers verified successfully.',
      {
        addedHeaders: missingHeaders,
        requiredHeaders: requiredHeaders
      }
    );
  }, 'Unable to verify Category_Master headers.');
}

/**
 * Creates or returns the test category.
 */
function createTestCategoryForSku_() {
  const headerResult = ensureCategoryMasterHeaders();

  if (!headerResult.success) {
    throw new Error(headerResult.message);
  }

  const categories = getSheetObjects_(
    getCategorySheetName_()
  );

  const existingCategory = categories.find(
    function (category) {
      return (
        normalizeLower_(
          category['Category Name']
        ) === 'test category'
      );
    }
  );

  if (existingCategory) {
    const categoryId = normalizeUpper_(
      existingCategory['Category ID']
    );

    if (!categoryId) {
      throw new Error(
        'Test Category exists but Category ID is blank.'
      );
    }

    if (
      normalizeUpper_(existingCategory['Status']) !==
      'ACTIVE'
    ) {
      updateObjectRow_(
        getCategorySheetName_(),
        existingCategory._rowNumber,
        {
          'Status': 'ACTIVE',
          'Updated By': getCurrentUserEmail_(),
          'Updated At': new Date()
        }
      );
    }

    return {
      categoryId: categoryId,
      categoryName: normalizeText_(
        existingCategory['Category Name']
      )
    };
  }

  const categoryId = generateNextId_(
    'CATEGORY',
    'CAT'
  );

  const now = new Date();
  const currentUserEmail = getCurrentUserEmail_();

  appendObjectRow_(
    getCategorySheetName_(),
    {
      'Category ID': categoryId,
      'Category Name': 'Test Category',
      'Status': 'ACTIVE',
      'Remarks': 'Created for SKU service testing.',
      'Created By': currentUserEmail,
      'Created At': now,
      'Updated By': currentUserEmail,
      'Updated At': now
    }
  );

  SpreadsheetApp.flush();

  const savedCategory = getSheetObjects_(
    getCategorySheetName_()
  ).find(function (category) {
    return (
      normalizeUpper_(category['Category ID']) ===
      categoryId
    );
  });

  if (!savedCategory) {
    throw new Error(
      'Test Category could not be saved in Category_Master.'
    );
  }

  return {
    categoryId: categoryId,
    categoryName: 'Test Category'
  };
}
