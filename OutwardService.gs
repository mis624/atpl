/**
 * PROJECT INVENTORY MANAGEMENT SYSTEM
 * File: OutwardService.gs
 *
 * Handles:
 * - Free inventory outward
 * - Customized inventory outward
 * - Stock availability validation
 * - Project restrictions
 * - Multiple outward items
 * - Automatic amount calculation
 * - Inventory balance reduction
 * - Transaction listing and summary
 *
 * IMPORTANT BUSINESS RULES:
 *
 * 1. FREE inventory can be issued to any active project.
 *
 * 2. CUSTOMIZE inventory can normally be issued only
 *    from the project to which the stock belongs.
 *
 * 3. Customized stock belonging to one project cannot
 *    be issued to another project directly.
 *
 * 4. Cross-project customized stock requires HOD approval.
 *    That workflow will be handled by ApprovalService.gs.
 */

const OUTWARD_SERVICE_CONFIG = Object.freeze({
  SHEET_NAME: 'Outward',

  INVENTORY_TYPE: Object.freeze({
    FREE: 'FREE',
    CUSTOMIZE: 'CUSTOMIZE'
  }),

  INVENTORY_MOVEMENT_TYPE: Object.freeze({
    FREE: 'FREE INVENTORY',
    CUSTOMIZE: 'CUSTOMIZE INVENTORY'
  }),

  STATUS: Object.freeze({
    COMPLETED: 'COMPLETED',
    CANCELLED: 'CANCELLED'
  }),

  ID_PREFIX: 'OUT',
  MAX_ITEMS: 100
});


/**
 * Returns the configured Outward sheet name.
 */
function getOutwardSheetName_() {
  if (
    typeof APP_CONFIG !== 'undefined' &&
    APP_CONFIG.SHEETS &&
    APP_CONFIG.SHEETS.OUTWARD
  ) {
    return APP_CONFIG.SHEETS.OUTWARD;
  }

  return OUTWARD_SERVICE_CONFIG.SHEET_NAME;
}


/**
 * Creates or verifies Outward sheet headers.
 *
 * Run once before testing OutwardService.gs.
 */
function ensureOutwardHeaders() {
  return safeExecute_(function () {
    requireRole_(APP_CONFIG.USER_ROLES.ADMIN);

    const sheet = getSystemSheet(
      getOutwardSheetName_()
    );

    const requiredHeaders = [
      'Outward No',
      'Date',
      'Inventory Type',
      'Source Project ID',
      'Source Project',
      'Destination Project ID',
      'Destination Project',
      'Reference No',
      'Anusha Invoice No',
      'Issued To',
      'SKU ID',
      'SKU Code',
      'SKU Name',
      'Qty',
      'Unit',
      'Rate',
      'Amount',
      'Status',
      'Remarks',
      'Entered By'
    ];

    ensureExactOutwardHeaders_(
      sheet,
      requiredHeaders
    );

    applyOutwardSheetFormatting_(sheet);

    return successResponse_(
      'Outward headers verified successfully.',
      {
        sheetName: getOutwardSheetName_(),
        headers: requiredHeaders
      }
    );
  }, 'Unable to verify Outward headers.');
}


/**
 * Ensures exact Outward headers.
 */
function ensureExactOutwardHeaders_(
  sheet,
  requiredHeaders
) {
  const requiredColumnCount =
    requiredHeaders.length;

  const currentMaxColumns =
    sheet.getMaxColumns();

  if (
    currentMaxColumns <
    requiredColumnCount
  ) {
    sheet.insertColumnsAfter(
      currentMaxColumns,
      requiredColumnCount -
        currentMaxColumns
    );
  }

  sheet
    .getRange(
      1,
      1,
      1,
      requiredColumnCount
    )
    .setValues([requiredHeaders])
    .setFontWeight('bold');

  sheet.setFrozenRows(1);
}


/**
 * Applies formatting and validations.
 */
function applyOutwardSheetFormatting_(sheet) {
  const requiredColumns = 20;

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

  const maxDataRows =
    Math.max(
      sheet.getMaxRows() - 1,
      1
    );

  sheet
    .getRange(
      2,
      1,
      maxDataRows,
      requiredColumns
    )
    .clearDataValidations();

  const inventoryTypeRule =
    SpreadsheetApp
      .newDataValidation()
      .requireValueInList(
        [
          OUTWARD_SERVICE_CONFIG
            .INVENTORY_TYPE.FREE,

          OUTWARD_SERVICE_CONFIG
            .INVENTORY_TYPE.CUSTOMIZE
        ],
        true
      )
      .setAllowInvalid(false)
      .build();

  const statusRule =
    SpreadsheetApp
      .newDataValidation()
      .requireValueInList(
        [
          OUTWARD_SERVICE_CONFIG
            .STATUS.COMPLETED,

          OUTWARD_SERVICE_CONFIG
            .STATUS.CANCELLED
        ],
        true
      )
      .setAllowInvalid(false)
      .build();

  sheet
    .getRange(
      2,
      3,
      maxDataRows,
      1
    )
    .setDataValidation(
      inventoryTypeRule
    );

  sheet
    .getRange(
      2,
      2,
      maxDataRows,
      1
    )
    .setNumberFormat(
      'dd-mmm-yyyy hh:mm AM/PM'
    );

  // Quantity: Column N.
  sheet
    .getRange(
      2,
      14,
      maxDataRows,
      1
    )
    .setNumberFormat('0.000');

  // Rate and Amount: Columns P:Q.
  sheet
    .getRange(
      2,
      16,
      maxDataRows,
      2
    )
    .setNumberFormat('0.00');

  // Status: Column R.
  sheet
    .getRange(
      2,
      18,
      maxDataRows,
      1
    )
    .setDataValidation(
      statusRule
    );

  sheet
    .getRange(
      1,
      1,
      1,
      requiredColumns
    )
    .setFontWeight('bold');

  sheet.setFrozenRows(1);

  SpreadsheetApp.flush();
}


/**
 * Creates an outward transaction.
 *
 * FREE Inventory request:
 *
 * {
 *   inventoryType: 'FREE',
 *   destinationProjectId: 'PRJ000001',
 *   referenceNo: 'DC-001',
 *   issuedTo: 'Aman',
 *   remarks: '',
 *   items: [
 *     {
 *       skuId: 'SKU000001',
 *       quantity: 2
 *     }
 *   ]
 * }
 *
 * CUSTOMIZE Inventory request:
 *
 * {
 *   inventoryType: 'CUSTOMIZE',
 *   sourceProjectId: 'PRJ000001',
 *   destinationProjectId: 'PRJ000001',
 *   referenceNo: 'DC-002',
 *   issuedTo: 'Aman',
 *   remarks: '',
 *   items: [
 *     {
 *       skuId: 'SKU000001',
 *       quantity: 2
 *     }
 *   ]
 * }
 */
function createOutward(outwardData) {
  return safeExecute_(function () {
    const session =
      requirePermission_('canOutward');

    const payload =
      validateOutwardRequest_(
        outwardData
      );

    const outwardNo =
      generateNextId_(
        'OUTWARD',
        OUTWARD_SERVICE_CONFIG.ID_PREFIX
      );

    const transactionDate =
      new Date();

    const inventoryResults = [];
    const inventoryRollbackItems = [];
    const savedRows = [];

    try {
      /*
       * First reduce inventory.
       */
      payload.items.forEach(
        function (item) {
          const inventoryResult =
            applyInventoryMovement({
              inventoryType:
                payload.inventoryMovementType,

              movementType:
                INVENTORY_SERVICE_CONFIG
                  .MOVEMENT_TYPE.OUTWARD,

              skuId:
                item.sku.skuId,

              projectId:
                payload.sourceProject
                  ? payload.sourceProject.projectId
                  : '',

              quantity:
                item.quantity,

              rate:
                item.rate,

              referenceId:
                outwardNo,

              remarks:
                payload.remarks ||
                (
                  'Stock issued to ' +
                  payload.destinationProject
                    .projectName
                )
            });

          if (
            !inventoryResult ||
            inventoryResult.success !== true
          ) {
            throw new Error(
              inventoryResult &&
              inventoryResult.message
                ? inventoryResult.message
                : 'Inventory outward update failed.'
            );
          }

          inventoryResults.push(
            inventoryResult.data
          );

          inventoryRollbackItems.push({
            inventoryType:
              payload.inventoryMovementType,

            skuId:
              item.sku.skuId,

            projectId:
              payload.sourceProject
                ? payload.sourceProject.projectId
                : '',

            quantity:
              item.quantity,

            rate:
              item.rate,

            referenceId:
              outwardNo
          });
        }
      );

      /*
       * Save outward sheet rows after inventory succeeds.
       */
      payload.items.forEach(
        function (item) {
          const amount =
            roundTwo_(
              item.quantity *
              item.rate
            );

          const rowNumber =
            appendObjectRow_(
              getOutwardSheetName_(),
              {
                'Outward No':
                  outwardNo,

                'Date':
                  transactionDate,

                'Inventory Type':
                  payload.inventoryType,

                'Source Project ID':
                  payload.sourceProject
                    ? payload.sourceProject.projectId
                    : '',

                'Source Project':
                  payload.sourceProject
                    ? payload.sourceProject.projectName
                    : 'FREE INVENTORY',

                'Destination Project ID':
                  payload.destinationProject
                    .projectId,

                'Destination Project':
                  payload.destinationProject
                    .projectName,

                'Reference No':
                  payload.referenceNo,

                'Anusha Invoice No':
                  payload.anushaInvoiceNo,

                'Issued To':
                  payload.issuedTo,

                'SKU ID':
                  item.sku.skuId,

                'SKU Code':
                  item.sku.skuCode,

                'SKU Name':
                  item.sku.skuName,

                'Qty':
                  item.quantity,

                'Unit':
                  item.sku.uom,

                'Rate':
                  item.rate,

                'Amount':
                  amount,

                'Status':
                  OUTWARD_SERVICE_CONFIG
                    .STATUS.COMPLETED,

                'Remarks':
                  payload.remarks,

                'Entered By':
                  session.email
              }
            );

          savedRows.push(rowNumber);
        }
      );

      const totalQuantity =
        payload.items.reduce(
          function (total, item) {
            return total +
              item.quantity;
          },
          0
        );

      const totalAmount =
        payload.items.reduce(
          function (total, item) {
            return total +
              (
                item.quantity *
                item.rate
              );
          },
          0
        );

      addTransactionLog_(
        session.email,
        'OUTWARD',
        'CREATE',
        outwardNo,
        {
          inventoryType:
            payload.inventoryType,

          sourceProjectId:
            payload.sourceProject
              ? payload.sourceProject.projectId
              : '',

          destinationProjectId:
            payload.destinationProject
              .projectId,

          referenceNo:
            payload.referenceNo,

          anushaInvoiceNo:
            payload.anushaInvoiceNo,

          issuedTo:
            payload.issuedTo,

          itemCount:
            payload.items.length,

          totalQuantity:
            roundTwo_(totalQuantity),

          totalAmount:
            roundTwo_(totalAmount)
        }
      );

      return successResponse_(
        'Inventory outward completed successfully.',
        {
          outwardNo: outwardNo,

          date:
            formatDateTime_(
              transactionDate
            ),

          inventoryType:
            payload.inventoryType,

          sourceProject:
            payload.sourceProject,

          destinationProject:
            payload.destinationProject,

          referenceNo:
            payload.referenceNo,

          anushaInvoiceNo:
            payload.anushaInvoiceNo,

          issuedTo:
            payload.issuedTo,

          itemCount:
            payload.items.length,

          totalQuantity:
            roundTwo_(totalQuantity),

          totalAmount:
            roundTwo_(totalAmount),

          items:
            inventoryResults
        }
      );

    } catch (error) {
      rollbackOutwardRows_(
        savedRows
      );

      rollbackOutwardInventory_(
        inventoryRollbackItems
      );

      throw error;
    }
  }, 'Unable to complete inventory outward.');
}


/**
 * Validates outward request.
 */
function validateOutwardRequest_(
  outwardData
) {
  if (
    !outwardData ||
    typeof outwardData !== 'object'
  ) {
    throw new Error(
      'Invalid outward request.'
    );
  }

  const inventoryType =
    normalizeOutwardInventoryType_(
      outwardData.inventoryType
    );

  const referenceNo =
    normalizeText_(
      outwardData.referenceNo
    );

  const anushaInvoiceNo =
    normalizeText_(
      outwardData.anushaInvoiceNo
    );

  const issuedTo =
    normalizeText_(
      outwardData.issuedTo
    );

  const remarks =
    normalizeText_(
      outwardData.remarks
    );

  if (!referenceNo) {
    throw new Error(
      'Reference No is required.'
    );
  }

  if (!issuedTo) {
    throw new Error(
      'Issued To is required.'
    );
  }

  const destinationProjectId =
    normalizeUpper_(
      outwardData.destinationProjectId
    );

  if (!destinationProjectId) {
    throw new Error(
      'Destination Project is required.'
    );
  }

  const destinationProjectRecord =
    getProjectRecordById_(
      destinationProjectId
    );

  if (!destinationProjectRecord) {
    throw new Error(
      'Destination Project not found: ' +
      destinationProjectId
    );
  }

  if (
    normalizeUpper_(
      destinationProjectRecord['Status']
    ) !== 'ACTIVE'
  ) {
    throw new Error(
      'Destination Project is not active.'
    );
  }

  const destinationProject = {
    projectId:
      normalizeUpper_(
        destinationProjectRecord[
          'Project ID'
        ]
      ),

    projectCode:
      normalizeUpper_(
        destinationProjectRecord[
          'Project Code'
        ]
      ),

    projectName:
      normalizeText_(
        destinationProjectRecord[
          'Project Name'
        ]
      )
  };

  let sourceProject = null;

  if (
    inventoryType ===
    OUTWARD_SERVICE_CONFIG
      .INVENTORY_TYPE.CUSTOMIZE
  ) {
    const sourceProjectId =
      normalizeUpper_(
        outwardData.sourceProjectId
      );

    if (!sourceProjectId) {
      throw new Error(
        'Source Project is required for customized inventory.'
      );
    }

    const sourceProjectRecord =
      getProjectRecordById_(
        sourceProjectId
      );

    if (!sourceProjectRecord) {
      throw new Error(
        'Source Project not found: ' +
        sourceProjectId
      );
    }

    sourceProject = {
      projectId:
        normalizeUpper_(
          sourceProjectRecord[
            'Project ID'
          ]
        ),

      projectCode:
        normalizeUpper_(
          sourceProjectRecord[
            'Project Code'
          ]
        ),

      projectName:
        normalizeText_(
          sourceProjectRecord[
            'Project Name'
          ]
        )
    };

    /*
     * Customized stock cannot directly move
     * to another project.
     */
    if (
      sourceProject.projectId !==
      destinationProject.projectId
    ) {
      throw new Error(
        'HOD approval is required to issue customized stock to another project.'
      );
    }
  }

  if (
    !Array.isArray(
      outwardData.items
    ) ||
    outwardData.items.length === 0
  ) {
    throw new Error(
      'At least one outward item is required.'
    );
  }

  if (
    outwardData.items.length >
    OUTWARD_SERVICE_CONFIG.MAX_ITEMS
  ) {
    throw new Error(
      'A maximum of ' +
      OUTWARD_SERVICE_CONFIG.MAX_ITEMS +
      ' items are allowed per outward.'
    );
  }

  const duplicateSkuMap = {};

  const items =
    outwardData.items.map(
      function (
        inputItem,
        index
      ) {
        const lineNumber =
          index + 1;

        if (
          !inputItem ||
          typeof inputItem !== 'object'
        ) {
          throw new Error(
            'Invalid item data at line ' +
            lineNumber +
            '.'
          );
        }

        const skuId =
          normalizeUpper_(
            inputItem.skuId
          );

        if (!skuId) {
          throw new Error(
            'SKU is required at line ' +
            lineNumber +
            '.'
          );
        }

        if (
          duplicateSkuMap[skuId]
        ) {
          throw new Error(
            'Duplicate SKU found: ' +
            skuId
          );
        }

        duplicateSkuMap[skuId] =
          true;

        const skuRecord =
          getSkuRecordById_(
            skuId
          );

        if (!skuRecord) {
          throw new Error(
            'SKU not found at line ' +
            lineNumber +
            ': ' +
            skuId
          );
        }

        if (
          normalizeUpper_(
            skuRecord['Status']
          ) !== 'ACTIVE'
        ) {
          throw new Error(
            'SKU is inactive at line ' +
            lineNumber +
            ': ' +
            skuId
          );
        }

        const quantity =
          toPositiveNumber_(
            inputItem.quantity,
            'Quantity at line ' +
              lineNumber
          );

        let inventoryRecord;

        if (
          inventoryType ===
          OUTWARD_SERVICE_CONFIG
            .INVENTORY_TYPE.FREE
        ) {
          inventoryRecord =
            getFreeInventoryRecordBySkuId_(
              skuId
            );
        } else {
          inventoryRecord =
            getCustomizeInventoryRecord_(
              sourceProject.projectId,
              skuId
            );
        }

        if (!inventoryRecord) {
          throw new Error(
            'Inventory record not found for SKU ' +
            skuId +
            ' at line ' +
            lineNumber +
            '.'
          );
        }

        const availableQty =
          inventoryToNumber_(
            inventoryRecord[
              'Available Qty'
            ],
            0
          );

        if (
          quantity > availableQty
        ) {
          throw new Error(
            'Insufficient stock at line ' +
            lineNumber +
            '. SKU: ' +
            skuId +
            ', Available: ' +
            availableQty +
            ', Requested: ' +
            quantity
          );
        }

        const inventoryRate =
          inventoryToNumber_(
            inventoryRecord[
              'Average Rate'
            ],
            0
          );

        return {
          sku: {
            skuId:
              normalizeUpper_(
                skuRecord['SKU ID']
              ),

            skuCode:
              normalizeUpper_(
                skuRecord['SKU Code']
              ),

            skuName:
              normalizeText_(
                skuRecord['SKU Name']
              ),

            uom:
              normalizeUpper_(
                skuRecord['UOM']
              )
          },

          quantity:
            roundTwo_(quantity),

          rate:
            roundTwo_(
              inventoryRate
            )
        };
      }
    );

  return {
    inventoryType:
      inventoryType,

    inventoryMovementType:
      inventoryType ===
      OUTWARD_SERVICE_CONFIG
        .INVENTORY_TYPE.FREE
        ? OUTWARD_SERVICE_CONFIG
            .INVENTORY_MOVEMENT_TYPE
            .FREE
        : OUTWARD_SERVICE_CONFIG
            .INVENTORY_MOVEMENT_TYPE
            .CUSTOMIZE,

    sourceProject:
      sourceProject,

    destinationProject:
      destinationProject,

    referenceNo:
      referenceNo,

    anushaInvoiceNo:
      anushaInvoiceNo,

    issuedTo:
      issuedTo,

    remarks:
      remarks,

    items:
      items
  };
}


/**
 * Normalizes outward inventory type.
 */
function normalizeOutwardInventoryType_(
  inventoryType
) {
  const normalizedType =
    normalizeUpper_(
      inventoryType
    );

  if (
    normalizedType === 'FREE' ||
    normalizedType ===
      'FREE INVENTORY'
  ) {
    return OUTWARD_SERVICE_CONFIG
      .INVENTORY_TYPE.FREE;
  }

  if (
    normalizedType === 'CUSTOMIZE' ||
    normalizedType ===
      'CUSTOMIZE INVENTORY'
  ) {
    return OUTWARD_SERVICE_CONFIG
      .INVENTORY_TYPE.CUSTOMIZE;
  }

  throw new Error(
    'Invalid outward inventory type: ' +
    normalizedType
  );
}


/**
 * Removes outward rows created during
 * a failed transaction.
 */
function rollbackOutwardRows_(
  rowNumbers
) {
  if (
    !Array.isArray(rowNumbers) ||
    rowNumbers.length === 0
  ) {
    return;
  }

  const sheet =
    getSystemSheet(
      getOutwardSheetName_()
    );

  rowNumbers
    .slice()
    .sort(function (first, second) {
      return second - first;
    })
    .forEach(function (rowNumber) {
      if (
        rowNumber >= 2 &&
        rowNumber <= sheet.getLastRow()
      ) {
        sheet.deleteRow(rowNumber);
      }
    });
}


/**
 * Restores inventory after a failed outward.
 */
function rollbackOutwardInventory_(
  rollbackItems
) {
  if (
    !Array.isArray(rollbackItems) ||
    rollbackItems.length === 0
  ) {
    return;
  }

  rollbackItems
    .slice()
    .reverse()
    .forEach(function (item) {
      try {
        applyInventoryMovement({
          inventoryType:
            item.inventoryType,

          movementType:
            INVENTORY_SERVICE_CONFIG
              .MOVEMENT_TYPE.ADJUSTMENT_IN,

          skuId:
            item.skuId,

          projectId:
            item.projectId,

          quantity:
            item.quantity,

          rate:
            item.rate,

          referenceId:
            item.referenceId +
            '-ROLLBACK',

          remarks:
            'Automatic rollback after failed outward transaction.'
        });
      } catch (rollbackError) {
        console.error(
          'Outward inventory rollback failed:',
          rollbackError
        );
      }
    });
}


/**
 * Returns one outward transaction.
 */
function getOutwardByNumber(
  outwardNo
) {
  return safeExecute_(function () {
    requireAuthenticatedUser_();

    const normalizedOutwardNo =
      normalizeUpper_(
        outwardNo
      );

    if (!normalizedOutwardNo) {
      throw new Error(
        'Outward number is required.'
      );
    }

    const records =
      getSheetObjects_(
        getOutwardSheetName_()
      ).filter(
        function (record) {
          return (
            normalizeUpper_(
              record['Outward No']
            ) === normalizedOutwardNo
          );
        }
      );

    if (records.length === 0) {
      throw new Error(
        'Outward transaction not found: ' +
        normalizedOutwardNo
      );
    }

    return successResponse_(
      'Outward transaction loaded successfully.',
      mapOutwardTransaction_(
        records
      )
    );
  }, 'Unable to load outward transaction.');
}


/**
 * Returns outward transactions.
 */
function getOutwardTransactions(
  filters
) {
  return safeExecute_(function () {
    requireAuthenticatedUser_();

    filters = filters || {};

    const search =
      normalizeLower_(
        filters.search
      );

    const inventoryType =
      filters.inventoryType
        ? normalizeOutwardInventoryType_(
            filters.inventoryType
          )
        : '';

    const destinationProjectId =
      normalizeUpper_(
        filters.destinationProjectId
      );

    const dateFrom =
      parseDate_(
        filters.dateFrom
      );

    const dateTo =
      parseDate_(
        filters.dateTo
      );

    const records =
      getSheetObjects_(
        getOutwardSheetName_()
      ).filter(
        function (record) {
          if (
            !normalizeText_(
              record['Outward No']
            )
          ) {
            return false;
          }

          if (
            inventoryType &&
            normalizeUpper_(
              record[
                'Inventory Type'
              ]
            ) !== inventoryType
          ) {
            return false;
          }

          if (
            destinationProjectId &&
            normalizeUpper_(
              record[
                'Destination Project ID'
              ]
            ) !== destinationProjectId
          ) {
            return false;
          }

          const outwardDate =
            parseDate_(
              record['Date']
            );

          if (
            dateFrom &&
            (
              !outwardDate ||
              outwardDate.getTime() <
                dateFrom.getTime()
            )
          ) {
            return false;
          }

          if (
            dateTo &&
            (
              !outwardDate ||
              outwardDate.getTime() >
                outwardEndOfDay_(
                  dateTo
                ).getTime()
            )
          ) {
            return false;
          }

          if (search) {
            const searchableText = [
              record['Outward No'],
              record['Inventory Type'],
              record['Source Project'],
              record['Destination Project'],
              record['Reference No'],
              record['Issued To'],
              record['SKU Code'],
              record['SKU Name'],
              record['Status'],
              record['Entered By']
            ]
              .map(normalizeLower_)
              .join(' ');

            if (
              searchableText.indexOf(
                search
              ) === -1
            ) {
              return false;
            }
          }

          return true;
        }
      );

    const transactionMap = {};

    records.forEach(
      function (record) {
        const outwardNo =
          normalizeUpper_(
            record['Outward No']
          );

        if (!transactionMap[outwardNo]) {
          transactionMap[outwardNo] = [];
        }

        transactionMap[outwardNo]
          .push(record);
      }
    );

    const transactions =
      Object.keys(transactionMap)
        .map(function (outwardNo) {
          return mapOutwardTransaction_(
            transactionMap[outwardNo]
          );
        })
        .sort(function (first, second) {
          const firstDate =
            parseDate_(first.date) ||
            new Date(0);

          const secondDate =
            parseDate_(second.date) ||
            new Date(0);

          return (
            secondDate.getTime() -
            firstDate.getTime()
          );
        });

    return successResponse_(
      'Outward transactions loaded successfully.',
      paginateRecords_(
        transactions,
        filters.pageNumber || 1,
        filters.pageSize || 20
      )
    );
  }, 'Unable to load outward transactions.');
}


/**
 * Maps outward rows into one transaction.
 */
function mapOutwardTransaction_(
  records
) {
  const firstRecord =
    records[0];

  const items =
    records.map(
      function (record) {
        return {
          skuId:
            normalizeUpper_(
              record['SKU ID']
            ),

          skuCode:
            normalizeUpper_(
              record['SKU Code']
            ),

          skuName:
            normalizeText_(
              record['SKU Name']
            ),

          quantity:
            toNumber_(
              record['Qty'],
              0
            ),

          unit:
            normalizeUpper_(
              record['Unit']
            ),

          rate:
            toNumber_(
              record['Rate'],
              0
            ),

          amount:
            toNumber_(
              record['Amount'],
              0
            )
        };
      }
    );

  return {
    outwardNo:
      normalizeUpper_(
        firstRecord['Outward No']
      ),

    date:
      formatDateTime_(
        firstRecord['Date']
      ),

    inventoryType:
      normalizeUpper_(
        firstRecord[
          'Inventory Type'
        ]
      ),

    sourceProjectId:
      normalizeUpper_(
        firstRecord[
          'Source Project ID'
        ]
      ),

    sourceProject:
      normalizeText_(
        firstRecord[
          'Source Project'
        ]
      ),

    destinationProjectId:
      normalizeUpper_(
        firstRecord[
          'Destination Project ID'
        ]
      ),

    destinationProject:
      normalizeText_(
        firstRecord[
          'Destination Project'
        ]
      ),

    referenceNo:
      normalizeText_(
        firstRecord['Reference No']
      ),

    anushaInvoiceNo:
      normalizeText_(
        firstRecord[
          'Anusha Invoice No'
        ]
      ),

    issuedTo:
      normalizeText_(
        firstRecord['Issued To']
      ),

    status:
      normalizeUpper_(
        firstRecord['Status']
      ),

    remarks:
      normalizeText_(
        firstRecord['Remarks']
      ),

    enteredBy:
      normalizeLower_(
        firstRecord['Entered By']
      ),

    itemCount:
      items.length,

    totalQuantity:
      roundTwo_(
        items.reduce(
          function (total, item) {
            return total +
              item.quantity;
          },
          0
        )
      ),

    totalAmount:
      roundTwo_(
        items.reduce(
          function (total, item) {
            return total +
              item.amount;
          },
          0
        )
      ),

    items:
      items
  };
}


/**
 * Returns end of day.
 */
function outwardEndOfDay_(
  dateValue
) {
  const date =
    new Date(
      dateValue.getTime()
    );

  date.setHours(
    23,
    59,
    59,
    999
  );

  return date;
}


/**
 * Returns outward summary.
 */
function getOutwardSummary() {
  return safeExecute_(function () {
    requireAuthenticatedUser_();

    const records =
      getSheetObjects_(
        getOutwardSheetName_()
      ).filter(function (record) {
        return Boolean(
          normalizeText_(
            record['Outward No']
          )
        );
      });

    const outwardNumbers = {};
    const destinationProjects = {};
    const issuedToUsers = {};

    const summary = {
      totalTransactions: 0,
      totalLines: records.length,
      totalQuantity: 0,
      totalAmount: 0,
      freeQuantity: 0,
      customizeQuantity: 0,
      destinationProjects: 0,
      issuedToUsers: 0
    };

    records.forEach(
      function (record) {
        const outwardNo =
          normalizeUpper_(
            record['Outward No']
          );

        if (outwardNo) {
          outwardNumbers[outwardNo] =
            true;
        }

        const destinationProject =
          normalizeUpper_(
            record[
              'Destination Project'
            ]
          );

        if (destinationProject) {
          destinationProjects[
            destinationProject
          ] = true;
        }

        const issuedTo =
          normalizeUpper_(
            record['Issued To']
          );

        if (issuedTo) {
          issuedToUsers[
            issuedTo
          ] = true;
        }

        const quantity =
          toNumber_(
            record['Qty'],
            0
          );

        const amount =
          toNumber_(
            record['Amount'],
            0
          );

        summary.totalQuantity +=
          quantity;

        summary.totalAmount +=
          amount;

        if (
          normalizeUpper_(
            record['Inventory Type']
          ) ===
          OUTWARD_SERVICE_CONFIG
            .INVENTORY_TYPE.FREE
        ) {
          summary.freeQuantity +=
            quantity;
        } else {
          summary.customizeQuantity +=
            quantity;
        }
      }
    );

    summary.totalTransactions =
      Object.keys(
        outwardNumbers
      ).length;

    summary.destinationProjects =
      Object.keys(
        destinationProjects
      ).length;

    summary.issuedToUsers =
      Object.keys(
        issuedToUsers
      ).length;

    summary.totalQuantity =
      roundTwo_(
        summary.totalQuantity
      );

    summary.totalAmount =
      roundTwo_(
        summary.totalAmount
      );

    summary.freeQuantity =
      roundTwo_(
        summary.freeQuantity
      );

    summary.customizeQuantity =
      roundTwo_(
        summary.customizeQuantity
      );

    return successResponse_(
      'Outward summary loaded successfully.',
      summary
    );
  }, 'Unable to load outward summary.');
}


/**
 * Test FREE Inventory outward.
 *
 * Current expected Free Inventory:
 * 30 units.
 *
 * This test removes 2 units.
 */



function testCreateFreeOutward() {
  const projects =
    getSheetObjects_(
      APP_CONFIG.SHEETS.PROJECT_MASTER
    ).filter(function (record) {
      return (
        normalizeText_(
          record['Project ID']
        ) &&
        normalizeUpper_(
          record['Status']
        ) === 'ACTIVE'
      );
    });

  if (!projects.length) {
    throw new Error(
      'No ACTIVE destination project found.'
    );
  }

  const freeInventory =
    getFreeInventory({
      pageNumber: 1,
      pageSize: 500
    });

  if (
    !freeInventory.success ||
    !freeInventory.data ||
    !freeInventory.data.records ||
    !freeInventory.data.records.length
  ) {
    throw new Error(
      'No Free Inventory stock available.'
    );
  }

  const availableItem =
    freeInventory.data.records.find(
      function (record) {
        return Number(
          record.availableQty || 0
        ) > 0;
      }
    );

  if (!availableItem) {
    throw new Error(
      'No Free Inventory SKU has available quantity.'
    );
  }

  const destinationProjectId =
    normalizeUpper_(
      projects[0]['Project ID']
    );

  const outwardQty =
    Math.min(
      1,
      Number(
        availableItem.availableQty
      )
    );

  const result =
    createOutward({
      inventoryType:
        'FREE',

      destinationProjectId:
        destinationProjectId,

      referenceNo:
        'TEST-FREE-OUT-' +
        Date.now(),

      anushaInvoiceNo:
        'TEST-INV-' +
        Date.now(),

      issuedTo:
        'Test User',

      remarks:
        'Automatic Free Outward test.',

      items: [
        {
          skuId:
            availableItem.skuId,

          quantity:
            outwardQty
        }
      ]
    });

  Logger.log(
    JSON.stringify(
      {
        destinationProjectId:
          destinationProjectId,

        selectedSku:
          availableItem,

        result:
          result
      },
      null,
      2
    )
  );

  return result;
}


function testCreateCustomizeOutward() {
  const projects =
    getSheetObjects_(
      APP_CONFIG.SHEETS.PROJECT_MASTER
    ).filter(function (record) {
      return (
        normalizeText_(
          record['Project ID']
        ) &&
        normalizeUpper_(
          record['Status']
        ) === 'ACTIVE'
      );
    });

  if (!projects.length) {
    throw new Error(
      'No ACTIVE project found in Project_Master.'
    );
  }

  const customizeInventoryResult =
    getCustomizeInventory({
      pageNumber: 1,
      pageSize: 500
    });

  if (
    !customizeInventoryResult ||
    customizeInventoryResult.success !== true ||
    !customizeInventoryResult.data ||
    !Array.isArray(
      customizeInventoryResult.data.records
    )
  ) {
    throw new Error(
      'Customize inventory could not be loaded.'
    );
  }

  const availableItem =
    customizeInventoryResult.data.records.find(
      function (record) {
        return (
          Number(
            record.availableQty || 0
          ) > 0 &&
          normalizeText_(
            record.skuId
          ) &&
          normalizeText_(
            record.projectId
          )
        );
      }
    );

  if (!availableItem) {
    throw new Error(
      'No Customize Inventory SKU has available quantity.'
    );
  }

  const sourceProjectId =
    normalizeUpper_(
      availableItem.projectId
    );

  /*
   * Same source and destination project means
   * direct customized outward without cross-project approval.
   */
  const destinationProjectId =
    sourceProjectId;

  const quantity =
    Math.min(
      1,
      Number(
        availableItem.availableQty
      )
    );

  const result =
    createOutward({
      inventoryType:
        'CUSTOMIZE',

      sourceProjectId:
        sourceProjectId,

      destinationProjectId:
        destinationProjectId,

      referenceNo:
        'TEST-CUSTOM-OUTWARD-' +
        Date.now(),

      anushaInvoiceNo:
        'TEST-ANUSHA-' +
        Date.now(),

      issuedTo:
        'Test User',

      remarks:
        'Automatic Customize Outward test.',

      items: [
        {
          skuId:
            availableItem.skuId,

          quantity:
            quantity
        }
      ]
    });

  Logger.log(
    JSON.stringify(
      {
        sourceProjectId:
          sourceProjectId,

        destinationProjectId:
          destinationProjectId,

        selectedSku:
          availableItem,

        result:
          result
      },
      null,
      2
    )
  );

  return result;
}
