/**
 * PROJECT INVENTORY MANAGEMENT SYSTEM
 * File: ConversionService.gs
 *
 * Handles:
 * - 90-day customized-stock eligibility
 * - Customize Inventory -> Free Inventory conversion request
 * - HOD/Admin approval and rejection
 * - Approved conversion execution by authorized doer
 * - Conversion listing, details, summary, and transaction log
 *
 * BUSINESS FLOW:
 * 1. Customized stock remains unused until its Eligibility Date.
 * 2. Daily/manual scan marks eligible records:
 *      Eligible for Free = YES
 * 3. Doer creates a conversion request.
 * 4. HOD/Admin approves or rejects.
 * 5. Authorized doer executes the approved conversion.
 * 6. Source customized quantity decreases.
 * 7. Free inventory quantity increases.
 */

const CONVERSION_SERVICE_CONFIG = Object.freeze({
  SHEET_NAME: 'Conversion_Requests',

  ID_PREFIX: 'CNV',

  STATUS: Object.freeze({
    PENDING: 'PENDING',
    APPROVED: 'APPROVED',
    REJECTED: 'REJECTED',
    EXECUTED: 'EXECUTED',
    CANCELLED: 'CANCELLED'
  }),

  SOURCE_INVENTORY_TYPE:
    'CUSTOMIZE INVENTORY',

  DESTINATION_INVENTORY_TYPE:
    'FREE INVENTORY',

  MAX_ITEMS: 100
});


/**
 * Returns the Conversion_Requests sheet name.
 */
function getConversionSheetName_() {
  if (
    typeof APP_CONFIG !== 'undefined' &&
    APP_CONFIG.SHEETS &&
    APP_CONFIG.SHEETS.CONVERSION_REQUESTS
  ) {
    return APP_CONFIG.SHEETS.CONVERSION_REQUESTS;
  }

  return CONVERSION_SERVICE_CONFIG
    .SHEET_NAME;
}


/**
 * Creates or verifies Conversion_Requests sheet.
 */
function ensureConversionHeaders() {
  return safeExecute_(function () {
    requireRole_(
      APP_CONFIG.USER_ROLES.ADMIN
    );

    const spreadsheet =
      SpreadsheetApp
        .getActiveSpreadsheet();

    const sheetName =
      getConversionSheetName_();

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

    const headers = [
      'Conversion ID',
      'Request Date',
      'Project ID',
      'Project Code',
      'Project Name',
      'Inventory ID',
      'SKU ID',
      'SKU Code',
      'SKU Name',
      'Qty',
      'Unit',
      'Rate',
      'Amount',
      'Eligibility Date',
      'Days Eligible',
      'Requested By',
      'Requested By Email',
      'Approver Email',
      'Status',
      'Approved By',
      'Approved At',
      'Approval Remarks',
      'Rejected By',
      'Rejected At',
      'Rejection Reason',
      'Executed By',
      'Executed At',
      'Free Inventory ID',
      'Reference No',
      'Remarks',
      'Update History'
    ];

    ensureExactConversionHeaders_(
      sheet,
      headers
    );

    applyConversionSheetFormatting_(
      sheet
    );

    return successResponse_(
      'Conversion request headers verified successfully.',
      {
        sheetName: sheetName,
        headers: headers
      }
    );
  }, 'Unable to verify conversion headers.');
}


/**
 * Ensures exact conversion headers.
 */
function ensureExactConversionHeaders_(
  sheet,
  headers
) {
  if (
    sheet.getMaxColumns() <
    headers.length
  ) {
    sheet.insertColumnsAfter(
      sheet.getMaxColumns(),
      headers.length -
        sheet.getMaxColumns()
    );
  }

  sheet
    .getRange(
      1,
      1,
      1,
      headers.length
    )
    .setValues([headers])
    .setFontWeight('bold');

  sheet.setFrozenRows(1);
}


/**
 * Applies formats and validations.
 */
function applyConversionSheetFormatting_(
  sheet
) {
  const dataRows = Math.max(
    sheet.getMaxRows() - 1,
    1
  );

  sheet
    .getRange(
      2,
      1,
      dataRows,
      31
    )
    .clearDataValidations();

  const statusRule =
    SpreadsheetApp
      .newDataValidation()
      .requireValueInList(
        Object.keys(
          CONVERSION_SERVICE_CONFIG.STATUS
        ).map(function (key) {
          return CONVERSION_SERVICE_CONFIG
            .STATUS[key];
        }),
        true
      )
      .setAllowInvalid(false)
      .build();

  // Qty.
  sheet
    .getRange(
      2,
      10,
      dataRows,
      1
    )
    .setNumberFormat('0.000');

  // Rate and Amount.
  sheet
    .getRange(
      2,
      12,
      dataRows,
      2
    )
    .setNumberFormat('0.00');

  // Date columns.
  [
    2, 14, 21, 24, 27
  ].forEach(function (columnNumber) {
    sheet
      .getRange(
        2,
        columnNumber,
        dataRows,
        1
      )
      .setNumberFormat(
        'dd-mmm-yyyy hh:mm AM/PM'
      );
  });

  // Status column S.
  sheet
    .getRange(
      2,
      19,
      dataRows,
      1
    )
    .setDataValidation(statusRule);

  sheet.setFrozenRows(1);

  SpreadsheetApp.flush();
}


/**
 * Scans customized inventory and updates:
 * - Eligible for Free
 * - Eligibility Date (when missing)
 *
 * Eligibility rule:
 * Available Qty > 0 AND
 * Today >= Eligibility Date.
 */
function refreshCustomizeFreeEligibility() {
  return safeExecute_(function () {
    requireAuthenticatedUser_();

    const sheet =
      getSystemSheet(
        getCustomizeInventorySheetName_()
      );

    const records =
      getSheetObjects_(
        getCustomizeInventorySheetName_()
      ).filter(function (record) {
        return Boolean(
          normalizeText_(
            record['Inventory ID']
          )
        );
      });

    const today = new Date();
    const eligibilityDays =
      getFreeEligibilityDays();

    let eligibleCount = 0;
    let notEligibleCount = 0;
    let updatedCount = 0;

    records.forEach(function (record) {
      const availableQty =
        inventoryToNumber_(
          record['Available Qty'],
          0
        );

      const status =
        normalizeUpper_(
          record['Status']
        );

      const baseDate =
        parseDate_(
          record['Last Movement Date']
        ) ||
        parseDate_(
          record['Last Inward Date']
        ) ||
        parseDate_(
          record['Created At']
        );

      let eligibilityDate =
        parseDate_(
          record['Eligibility Date']
        );

      if (
        !eligibilityDate &&
        baseDate
      ) {
        eligibilityDate =
          new Date(
            baseDate.getTime() +
            eligibilityDays *
            24 * 60 * 60 * 1000
          );
      }

      const eligible =
        availableQty > 0 &&
        status !== 'INACTIVE' &&
        status !== 'CONVERTED TO FREE' &&
        eligibilityDate &&
        today.getTime() >=
          eligibilityDate.getTime();

      const newEligibleValue =
        eligible ? 'YES' : 'NO';

      const currentEligibleValue =
        normalizeUpper_(
          record['Eligible for Free']
        );

      const updateData = {};

      if (
        eligibilityDate &&
        !parseDate_(
          record['Eligibility Date']
        )
      ) {
        updateData['Eligibility Date'] =
          eligibilityDate;
      }

      if (
        currentEligibleValue !==
        newEligibleValue
      ) {
        updateData['Eligible for Free'] =
          newEligibleValue;
      }

      if (
        Object.keys(updateData).length
      ) {
        updateObjectRow_(
          getCustomizeInventorySheetName_(),
          record._rowNumber,
          updateData
        );

        updatedCount++;
      }

      if (eligible) {
        eligibleCount++;
      } else {
        notEligibleCount++;
      }
    });

    SpreadsheetApp.flush();

    return successResponse_(
      'Customized inventory eligibility refreshed successfully.',
      {
        totalRecords: records.length,
        eligibleRecords: eligibleCount,
        notEligibleRecords:
          notEligibleCount,
        updatedRecords: updatedCount,
        eligibilityDays:
          eligibilityDays
      }
    );
  }, 'Unable to refresh customized inventory eligibility.');
}


/**
 * Returns customized inventory eligible for Free conversion.
 */
function getEligibleCustomizeInventory(
  filters
) {
  return safeExecute_(function () {
    requireAuthenticatedUser_();

    filters = filters || {};

    const search =
      normalizeLower_(
        filters.search
      );

    const projectId =
      normalizeUpper_(
        filters.projectId
      );

    const today = new Date();

    const records =
      getSheetObjects_(
        getCustomizeInventorySheetName_()
      )
        .filter(function (record) {
          if (
            !normalizeText_(
              record['Inventory ID']
            )
          ) {
            return false;
          }

          if (
            normalizeUpper_(
              record['Eligible for Free']
            ) !== 'YES'
          ) {
            return false;
          }

          if (
            inventoryToNumber_(
              record['Available Qty'],
              0
            ) <= 0
          ) {
            return false;
          }

          if (
            projectId &&
            normalizeUpper_(
              record['Project ID']
            ) !== projectId
          ) {
            return false;
          }

          if (search) {
            const searchable = [
              record['Inventory ID'],
              record['Project ID'],
              record['Project Code'],
              record['Project Name'],
              record['SKU ID'],
              record['SKU Code'],
              record['SKU Name'],
              record['Category Name']
            ]
              .map(normalizeLower_)
              .join(' ');

            if (
              searchable.indexOf(
                search
              ) === -1
            ) {
              return false;
            }
          }

          return true;
        })
        .map(function (record) {
          const eligibilityDate =
            parseDate_(
              record['Eligibility Date']
            );

          return {
            inventoryId:
              normalizeUpper_(
                record['Inventory ID']
              ),

            projectId:
              normalizeUpper_(
                record['Project ID']
              ),

            projectCode:
              normalizeUpper_(
                record['Project Code']
              ),

            projectName:
              normalizeText_(
                record['Project Name']
              ),

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

            categoryName:
              normalizeText_(
                record['Category Name']
              ),

            uom:
              normalizeUpper_(
                record['UOM']
              ),

            availableQty:
              inventoryToNumber_(
                record['Available Qty'],
                0
              ),

            averageRate:
              inventoryToNumber_(
                record['Average Rate'],
                0
              ),

            eligibilityDate:
              formatDateTime_(
                eligibilityDate
              ),

            daysEligible:
              eligibilityDate
                ? Math.max(
                    0,
                    daysBetween_(
                      eligibilityDate,
                      today
                    )
                  )
                : 0,

            status:
              normalizeUpper_(
                record['Status']
              )
          };
        });

    return successResponse_(
      'Eligible customized inventory loaded successfully.',
      paginateRecords_(
        records,
        filters.pageNumber || 1,
        filters.pageSize || 20
      )
    );
  }, 'Unable to load eligible customized inventory.');
}


/**
 * Creates Customize -> Free conversion request.
 *
 * Request:
 * {
 *   projectId: 'PRJ000001',
 *   approverEmail: 'hod@company.com',
 *   referenceNo: 'CNV-REF-001',
 *   remarks: '',
 *   items: [
 *     { skuId: 'SKU000001', quantity: 1 }
 *   ]
 * }
 */
function createConversionRequest(
  requestData
) {
  return safeExecute_(function () {
    const session =
      requirePermission_('canTransfer');

    const payload =
      validateConversionRequest_(
        requestData
      );

    const conversionId =
      generateNextId_(
        'CONVERSION_REQUEST',
        CONVERSION_SERVICE_CONFIG
          .ID_PREFIX
      );

    const now = new Date();

    const history =
      buildConversionHistoryText_(
        'Conversion request created',
        session.email,
        now
      );

    const savedRows = [];

    try {
      payload.items.forEach(
        function (item) {
          const amount =
            roundTwo_(
              item.quantity *
              item.rate
            );

          const rowNumber =
            appendObjectRow_(
              getConversionSheetName_(),
              {
                'Conversion ID':
                  conversionId,

                'Request Date':
                  now,

                'Project ID':
                  payload.project.projectId,

                'Project Code':
                  payload.project.projectCode,

                'Project Name':
                  payload.project.projectName,

                'Inventory ID':
                  item.inventoryId,

                'SKU ID':
                  item.skuId,

                'SKU Code':
                  item.skuCode,

                'SKU Name':
                  item.skuName,

                'Qty':
                  item.quantity,

                'Unit':
                  item.uom,

                'Rate':
                  item.rate,

                'Amount':
                  amount,

                'Eligibility Date':
                  item.eligibilityDate,

                'Days Eligible':
                  item.daysEligible,

                'Requested By':
                  payload.requestedBy,

                'Requested By Email':
                  session.email,

                'Approver Email':
                  payload.approverEmail,

                'Status':
                  CONVERSION_SERVICE_CONFIG
                    .STATUS.PENDING,

                'Approved By': '',
                'Approved At': '',
                'Approval Remarks': '',

                'Rejected By': '',
                'Rejected At': '',
                'Rejection Reason': '',

                'Executed By': '',
                'Executed At': '',

                'Free Inventory ID': '',

                'Reference No':
                  payload.referenceNo,

                'Remarks':
                  payload.remarks,

                'Update History':
                  history
              }
            );

          savedRows.push(rowNumber);
        }
      );

      addTransactionLog_(
        session.email,
        'CONVERSION',
        'CREATE REQUEST',
        conversionId,
        {
          projectId:
            payload.project.projectId,
          approverEmail:
            payload.approverEmail,
          itemCount:
            payload.items.length,
          totalQuantity:
            roundTwo_(
              payload.items.reduce(
                function (total, item) {
                  return total +
                    item.quantity;
                },
                0
              )
            )
        }
      );

      return successResponse_(
        'Conversion request created successfully.',
        {
          conversionId:
            conversionId,
          status:
            CONVERSION_SERVICE_CONFIG
              .STATUS.PENDING,
          project:
            payload.project,
          approverEmail:
            payload.approverEmail,
          itemCount:
            payload.items.length,
          totalQuantity:
            roundTwo_(
              payload.items.reduce(
                function (total, item) {
                  return total +
                    item.quantity;
                },
                0
              )
            )
        }
      );

    } catch (error) {
      deleteConversionRows_(
        savedRows
      );

      throw error;
    }
  }, 'Unable to create conversion request.');
}


/**
 * Validates conversion request.
 */
function validateConversionRequest_(
  requestData
) {
  if (
    !requestData ||
    typeof requestData !== 'object'
  ) {
    throw new Error(
      'Invalid conversion request.'
    );
  }

  const projectId =
    normalizeUpper_(
      requestData.projectId
    );

  const referenceNo =
    normalizeText_(
      requestData.referenceNo
    );

  const requestedBy =
    normalizeText_(
      requestData.requestedBy ||
      getCurrentUserEmail_()
    );

  const remarks =
    normalizeText_(
      requestData.remarks
    );

  if (!projectId) {
    throw new Error(
      'Project is required.'
    );
  }

  if (!referenceNo) {
    throw new Error(
      'Reference No is required.'
    );
  }

  const projectRecord =
    getProjectRecordById_(
      projectId
    );

  if (!projectRecord) {
    throw new Error(
      'Project not found: ' +
      projectId
    );
  }

  const project = {
    projectId:
      normalizeUpper_(
        projectRecord['Project ID']
      ),

    projectCode:
      normalizeUpper_(
        projectRecord['Project Code']
      ),

    projectName:
      normalizeText_(
        projectRecord['Project Name']
      )
  };

  let approverEmail =
    normalizeLower_(
      requestData.approverEmail
    );

  if (!approverEmail) {
    approverEmail =
      normalizeLower_(
        projectRecord['HOD Email']
      );
  }

  if (!approverEmail) {
    throw new Error(
      'Approver Email is required. Add HOD Email to the project or provide approverEmail.'
    );
  }

  if (
    !isValidEmail_(
      approverEmail
    )
  ) {
    throw new Error(
      'Invalid Approver Email: ' +
      approverEmail
    );
  }

  if (
    !Array.isArray(
      requestData.items
    ) ||
    requestData.items.length === 0
  ) {
    throw new Error(
      'At least one conversion item is required.'
    );
  }

  if (
    requestData.items.length >
    CONVERSION_SERVICE_CONFIG
      .MAX_ITEMS
  ) {
    throw new Error(
      'A maximum of ' +
      CONVERSION_SERVICE_CONFIG
        .MAX_ITEMS +
      ' items are allowed.'
    );
  }

  const duplicateMap = {};
  const today = new Date();

  const items =
    requestData.items.map(
      function (inputItem, index) {
        const lineNumber = index + 1;

        const skuId =
          normalizeUpper_(
            inputItem &&
            inputItem.skuId
          );

        if (!skuId) {
          throw new Error(
            'SKU is required at line ' +
            lineNumber +
            '.'
          );
        }

        if (duplicateMap[skuId]) {
          throw new Error(
            'Duplicate SKU found: ' +
            skuId
          );
        }

        duplicateMap[skuId] = true;

        const inventoryRecord =
          getCustomizeInventoryRecord_(
            projectId,
            skuId
          );

        if (!inventoryRecord) {
          throw new Error(
            'Customized inventory not found for SKU: ' +
            skuId
          );
        }

        if (
          normalizeUpper_(
            inventoryRecord[
              'Eligible for Free'
            ]
          ) !== 'YES'
        ) {
          throw new Error(
            'SKU is not eligible for Free conversion: ' +
            skuId
          );
        }

        const eligibilityDate =
          parseDate_(
            inventoryRecord[
              'Eligibility Date'
            ]
          );

        if (
          !eligibilityDate ||
          today.getTime() <
            eligibilityDate.getTime()
        ) {
          throw new Error(
            'Eligibility date is not complete for SKU: ' +
            skuId
          );
        }

        const availableQty =
          inventoryToNumber_(
            inventoryRecord[
              'Available Qty'
            ],
            0
          );

        const quantity =
          toPositiveNumber_(
            inputItem.quantity,
            'Quantity at line ' +
            lineNumber
          );

        if (
          quantity > availableQty
        ) {
          throw new Error(
            'Insufficient customized stock. SKU: ' +
            skuId +
            ', Available: ' +
            availableQty +
            ', Requested: ' +
            quantity
          );
        }

        return {
          inventoryId:
            normalizeUpper_(
              inventoryRecord[
                'Inventory ID'
              ]
            ),

          skuId:
            normalizeUpper_(
              inventoryRecord['SKU ID']
            ),

          skuCode:
            normalizeUpper_(
              inventoryRecord['SKU Code']
            ),

          skuName:
            normalizeText_(
              inventoryRecord['SKU Name']
            ),

          uom:
            normalizeUpper_(
              inventoryRecord['UOM']
            ),

          quantity:
            roundTwo_(quantity),

          rate:
            roundTwo_(
              inventoryToNumber_(
                inventoryRecord[
                  'Average Rate'
                ],
                0
              )
            ),

          eligibilityDate:
            eligibilityDate,

          daysEligible:
            Math.max(
              0,
              daysBetween_(
                eligibilityDate,
                today
              )
            )
        };
      }
    );

  return {
    project: project,
    approverEmail:
      approverEmail,
    referenceNo:
      referenceNo,
    requestedBy:
      requestedBy,
    remarks:
      remarks,
    items:
      items
  };
}


/**
 * Approves a pending conversion request.
 */
function approveConversionRequest(
  conversionId,
  approvalRemarks
) {
  return safeExecute_(function () {
    const session =
      requirePermission_('canApprove');

    const records =
      getConversionRows_(
        conversionId
      );

    if (!records.length) {
      throw new Error(
        'Conversion request not found: ' +
        conversionId
      );
    }

    assertConversionStatus_(
      records,
      CONVERSION_SERVICE_CONFIG
        .STATUS.PENDING
    );

    validateConversionStock_(records);

    const now = new Date();

    updateConversionRows_(
      records,
      {
        status:
          CONVERSION_SERVICE_CONFIG
            .STATUS.APPROVED,

        values: {
          'Approved By':
            session.email,
          'Approved At':
            now,
          'Approval Remarks':
            normalizeText_(
              approvalRemarks
            )
        },

        history:
          buildConversionHistoryText_(
            'Conversion request approved',
            session.email,
            now
          )
      }
    );

    addTransactionLog_(
      session.email,
      'CONVERSION',
      'APPROVE',
      normalizeUpper_(
        conversionId
      ),
      {
        itemCount:
          records.length
      }
    );

    return successResponse_(
      'Conversion request approved successfully.',
      {
        conversionId:
          normalizeUpper_(
            conversionId
          ),
        status:
          CONVERSION_SERVICE_CONFIG
            .STATUS.APPROVED,
        approvedBy:
          session.email,
        approvedAt:
          formatDateTime_(now)
      }
    );
  }, 'Unable to approve conversion request.');
}


/**
 * Rejects a pending conversion request.
 */
function rejectConversionRequest(
  conversionId,
  rejectionReason
) {
  return safeExecute_(function () {
    const session =
      requirePermission_('canApprove');

    const reason =
      normalizeText_(
        rejectionReason
      );

    if (!reason) {
      throw new Error(
        'Rejection reason is required.'
      );
    }

    const records =
      getConversionRows_(
        conversionId
      );

    if (!records.length) {
      throw new Error(
        'Conversion request not found: ' +
        conversionId
      );
    }

    assertConversionStatus_(
      records,
      CONVERSION_SERVICE_CONFIG
        .STATUS.PENDING
    );

    const now = new Date();

    updateConversionRows_(
      records,
      {
        status:
          CONVERSION_SERVICE_CONFIG
            .STATUS.REJECTED,

        values: {
          'Rejected By':
            session.email,
          'Rejected At':
            now,
          'Rejection Reason':
            reason
        },

        history:
          buildConversionHistoryText_(
            'Conversion request rejected: ' +
            reason,
            session.email,
            now
          )
      }
    );

    addTransactionLog_(
      session.email,
      'CONVERSION',
      'REJECT',
      normalizeUpper_(
        conversionId
      ),
      {
        reason: reason,
        itemCount:
          records.length
      }
    );

    return successResponse_(
      'Conversion request rejected successfully.',
      {
        conversionId:
          normalizeUpper_(
            conversionId
          ),
        status:
          CONVERSION_SERVICE_CONFIG
            .STATUS.REJECTED,
        rejectionReason:
          reason
      }
    );
  }, 'Unable to reject conversion request.');
}


/**
 * Executes an approved conversion.
 *
 * Customized stock decreases.
 * Free stock increases.
 */
function executeApprovedConversion(
  conversionId
) {
  return safeExecute_(function () {
    const session =
      requirePermission_('canTransfer');

    const cleanConversionId =
      normalizeUpper_(
        conversionId
      );

    if (!cleanConversionId) {
      throw new Error(
        'Conversion ID is required.'
      );
    }

    const records =
      getConversionRows_(
        cleanConversionId
      );

    if (!records.length) {
      throw new Error(
        'Conversion request not found: ' +
        cleanConversionId
      );
    }

    assertConversionStatus_(
      records,
      CONVERSION_SERVICE_CONFIG
        .STATUS.APPROVED
    );

    validateConversionStock_(
      records
    );

    const now = new Date();
    const executedItems = [];
    const rollbackItems = [];

    try {
      records.forEach(function (record) {
        const outwardResult =
          applyInventoryMovement({
            inventoryType:
              CONVERSION_SERVICE_CONFIG
                .SOURCE_INVENTORY_TYPE,

            movementType:
              INVENTORY_SERVICE_CONFIG
                .MOVEMENT_TYPE.OUTWARD,

            skuId:
              record.skuId,

            projectId:
              record.projectId,

            quantity:
              record.quantity,

            rate:
              record.rate,

            referenceId:
              cleanConversionId,

            remarks:
              'Converted customized stock to Free Inventory.'
          });

        if (
          !outwardResult ||
          outwardResult.success !== true
        ) {
          throw new Error(
            outwardResult &&
            outwardResult.message
              ? outwardResult.message
              : 'Customized inventory reduction failed.'
          );
        }

        rollbackItems.push({
          type: 'CUSTOMIZE_OUT',
          skuId: record.skuId,
          projectId:
            record.projectId,
          quantity:
            record.quantity,
          rate:
            record.rate
        });

        const inwardResult =
          applyInventoryMovement({
            inventoryType:
              CONVERSION_SERVICE_CONFIG
                .DESTINATION_INVENTORY_TYPE,

            movementType:
              INVENTORY_SERVICE_CONFIG
                .MOVEMENT_TYPE.INWARD,

            skuId:
              record.skuId,

            projectId: '',

            quantity:
              record.quantity,

            rate:
              record.rate,

            referenceId:
              cleanConversionId,

            remarks:
              'Received from customized inventory conversion.'
          });

        if (
          !inwardResult ||
          inwardResult.success !== true
        ) {
          throw new Error(
            inwardResult &&
            inwardResult.message
              ? inwardResult.message
              : 'Free inventory increase failed.'
          );
        }

        rollbackItems.push({
          type: 'FREE_IN',
          skuId: record.skuId,
          projectId: '',
          quantity:
            record.quantity,
          rate:
            record.rate
        });

        executedItems.push({
          skuId:
            record.skuId,
          skuCode:
            record.skuCode,
          quantity:
            record.quantity,
          rate:
            record.rate,
          customizeInventory:
            outwardResult.data,
          freeInventory:
            inwardResult.data
        });
      });

      records.forEach(function (record) {
        const freeRecord =
          getFreeInventoryRecordBySkuId_(
            record.skuId
          );

        updateObjectRow_(
          getConversionSheetName_(),
          record.rowNumber,
          {
            'Status':
              CONVERSION_SERVICE_CONFIG
                .STATUS.EXECUTED,

            'Executed By':
              session.email,

            'Executed At':
              now,

            'Free Inventory ID':
              freeRecord
                ? normalizeUpper_(
                    freeRecord[
                      'Inventory ID'
                    ]
                  )
                : '',

            'Update History': [
              record.updateHistory,
              buildConversionHistoryText_(
                'Conversion executed',
                session.email,
                now
              )
            ]
              .filter(Boolean)
              .join('\n')
          }
        );
      });

      /*
       * Update customized inventory status after conversion.
       */
      records.forEach(function (record) {
        const customizeRecord =
          getCustomizeInventoryRecord_(
            record.projectId,
            record.skuId
          );

        if (!customizeRecord) {
          return;
        }

        const remainingQty =
          inventoryToNumber_(
            customizeRecord[
              'Available Qty'
            ],
            0
          );

        updateObjectRow_(
          getCustomizeInventorySheetName_(),
          customizeRecord._rowNumber,
          {
            'Eligible for Free':
              remainingQty > 0
                ? 'YES'
                : 'NO',

            'Status':
              remainingQty > 0
                ? 'AVAILABLE'
                : 'CONVERTED TO FREE',

            'Updated By':
              session.email,

            'Updated At':
              now
          }
        );
      });

      SpreadsheetApp.flush();

      addTransactionLog_(
        session.email,
        'CONVERSION',
        'EXECUTE',
        cleanConversionId,
        {
          itemCount:
            records.length,
          totalQuantity:
            roundTwo_(
              records.reduce(
                function (total, record) {
                  return total +
                    record.quantity;
                },
                0
              )
            ),
          projectId:
            records[0].projectId
        }
      );

      return successResponse_(
        'Customized inventory converted to Free Inventory successfully.',
        {
          conversionId:
            cleanConversionId,

          status:
            CONVERSION_SERVICE_CONFIG
              .STATUS.EXECUTED,

          projectId:
            records[0].projectId,

          itemCount:
            records.length,

          totalQuantity:
            roundTwo_(
              records.reduce(
                function (total, record) {
                  return total +
                    record.quantity;
                },
                0
              )
            ),

          executedBy:
            session.email,

          executedAt:
            formatDateTime_(now),

          items:
            executedItems
        }
      );

    } catch (error) {
      rollbackConversionInventory_(
        rollbackItems,
        cleanConversionId
      );

      throw error;
    }
  }, 'Unable to execute approved conversion.');
}


/**
 * Attempts to rollback partial conversion.
 */
function rollbackConversionInventory_(
  rollbackItems,
  conversionId
) {
  rollbackItems
    .slice()
    .reverse()
    .forEach(function (item) {
      try {
        if (
          item.type ===
          'FREE_IN'
        ) {
          applyInventoryMovement({
            inventoryType:
              CONVERSION_SERVICE_CONFIG
                .DESTINATION_INVENTORY_TYPE,

            movementType:
              INVENTORY_SERVICE_CONFIG
                .MOVEMENT_TYPE.OUTWARD,

            skuId:
              item.skuId,

            projectId: '',

            quantity:
              item.quantity,

            rate:
              item.rate,

            referenceId:
              conversionId +
              '-ROLLBACK',

            remarks:
              'Automatic conversion rollback.'
          });
        }

        if (
          item.type ===
          'CUSTOMIZE_OUT'
        ) {
          applyInventoryMovement({
            inventoryType:
              CONVERSION_SERVICE_CONFIG
                .SOURCE_INVENTORY_TYPE,

            movementType:
              INVENTORY_SERVICE_CONFIG
                .MOVEMENT_TYPE.INWARD,

            skuId:
              item.skuId,

            projectId:
              item.projectId,

            quantity:
              item.quantity,

            rate:
              item.rate,

            referenceId:
              conversionId +
              '-ROLLBACK',

            remarks:
              'Automatic conversion rollback.'
          });
        }
      } catch (rollbackError) {
        console.error(
          'Conversion rollback failed:',
          rollbackError
        );
      }
    });
}


/**
 * Returns conversion request details.
 */
function getConversionRequest(
  conversionId
) {
  return safeExecute_(function () {
    requireAuthenticatedUser_();

    const records =
      getConversionRows_(
        conversionId
      );

    if (!records.length) {
      throw new Error(
        'Conversion request not found: ' +
        conversionId
      );
    }

    return successResponse_(
      'Conversion request loaded successfully.',
      mapConversionRequest_(
        records
      )
    );
  }, 'Unable to load conversion request.');
}


/**
 * Returns conversion requests.
 */
function getConversionRequests(
  filters
) {
  return safeExecute_(function () {
    requireAuthenticatedUser_();

    filters = filters || {};

    const search =
      normalizeLower_(
        filters.search
      );

    const status =
      normalizeUpper_(
        filters.status
      );

    const projectId =
      normalizeUpper_(
        filters.projectId
      );

    const records =
      getSheetObjects_(
        getConversionSheetName_()
      ).filter(function (record) {
        if (
          !normalizeText_(
            record['Conversion ID']
          )
        ) {
          return false;
        }

        if (
          status &&
          normalizeUpper_(
            record['Status']
          ) !== status
        ) {
          return false;
        }

        if (
          projectId &&
          normalizeUpper_(
            record['Project ID']
          ) !== projectId
        ) {
          return false;
        }

        if (search) {
          const searchable = [
            record['Conversion ID'],
            record['Project ID'],
            record['Project Name'],
            record['SKU ID'],
            record['SKU Code'],
            record['SKU Name'],
            record['Requested By'],
            record['Approver Email'],
            record['Status'],
            record['Reference No']
          ]
            .map(normalizeLower_)
            .join(' ');

          if (
            searchable.indexOf(
              search
            ) === -1
          ) {
            return false;
          }
        }

        return true;
      });

    const requestMap = {};

    records.forEach(function (record) {
      const conversionId =
        normalizeUpper_(
          record['Conversion ID']
        );

      if (!requestMap[conversionId]) {
        requestMap[conversionId] = [];
      }

      requestMap[conversionId]
        .push(record);
    });

    const requests =
      Object.keys(requestMap)
        .map(function (conversionId) {
          return mapConversionRequest_(
            requestMap[conversionId]
          );
        })
        .sort(function (first, second) {
          const firstDate =
            parseDate_(
              first.requestDate
            ) || new Date(0);

          const secondDate =
            parseDate_(
              second.requestDate
            ) || new Date(0);

          return (
            secondDate.getTime() -
            firstDate.getTime()
          );
        });

    return successResponse_(
      'Conversion requests loaded successfully.',
      paginateRecords_(
        requests,
        filters.pageNumber || 1,
        filters.pageSize || 20
      )
    );
  }, 'Unable to load conversion requests.');
}


/**
 * Returns conversion summary.
 */
function getConversionSummary() {
  return safeExecute_(function () {
    requireAuthenticatedUser_();

    const records =
      getSheetObjects_(
        getConversionSheetName_()
      ).filter(function (record) {
        return Boolean(
          normalizeText_(
            record['Conversion ID']
          )
        );
      });

    const map = {};

    records.forEach(function (record) {
      const id =
        normalizeUpper_(
          record['Conversion ID']
        );

      if (!map[id]) {
        map[id] = {
          status:
            normalizeUpper_(
              record['Status']
            ),
          quantity: 0,
          amount: 0
        };
      }

      map[id].quantity +=
        toNumber_(
          record['Qty'],
          0
        );

      map[id].amount +=
        toNumber_(
          record['Amount'],
          0
        );
    });

    const summary = {
      totalRequests: 0,
      pending: 0,
      approved: 0,
      rejected: 0,
      executed: 0,
      cancelled: 0,
      totalRequestedQty: 0,
      totalExecutedQty: 0,
      totalExecutedValue: 0
    };

    Object.keys(map).forEach(
      function (id) {
        const request = map[id];

        summary.totalRequests++;

        summary.totalRequestedQty +=
          request.quantity;

        if (
          request.status ===
          CONVERSION_SERVICE_CONFIG
            .STATUS.PENDING
        ) {
          summary.pending++;
        } else if (
          request.status ===
          CONVERSION_SERVICE_CONFIG
            .STATUS.APPROVED
        ) {
          summary.approved++;
        } else if (
          request.status ===
          CONVERSION_SERVICE_CONFIG
            .STATUS.REJECTED
        ) {
          summary.rejected++;
        } else if (
          request.status ===
          CONVERSION_SERVICE_CONFIG
            .STATUS.EXECUTED
        ) {
          summary.executed++;

          summary.totalExecutedQty +=
            request.quantity;

          summary.totalExecutedValue +=
            request.amount;
        } else if (
          request.status ===
          CONVERSION_SERVICE_CONFIG
            .STATUS.CANCELLED
        ) {
          summary.cancelled++;
        }
      }
    );

    summary.totalRequestedQty =
      roundTwo_(
        summary.totalRequestedQty
      );

    summary.totalExecutedQty =
      roundTwo_(
        summary.totalExecutedQty
      );

    summary.totalExecutedValue =
      roundTwo_(
        summary.totalExecutedValue
      );

    return successResponse_(
      'Conversion summary loaded successfully.',
      summary
    );
  }, 'Unable to load conversion summary.');
}


/**
 * Finds and maps conversion rows.
 */
function getConversionRows_(
  conversionId
) {
  const target =
    normalizeUpper_(
      conversionId
    );

  return getSheetObjects_(
    getConversionSheetName_()
  )
    .filter(function (record) {
      return (
        normalizeUpper_(
          record['Conversion ID']
        ) === target
      );
    })
    .map(mapConversionRow_);
}


/**
 * Maps one conversion row.
 */
function mapConversionRow_(record) {
  return {
    rowNumber:
      record._rowNumber,

    conversionId:
      normalizeUpper_(
        record['Conversion ID']
      ),

    requestDate:
      record['Request Date'],

    projectId:
      normalizeUpper_(
        record['Project ID']
      ),

    projectCode:
      normalizeUpper_(
        record['Project Code']
      ),

    projectName:
      normalizeText_(
        record['Project Name']
      ),

    inventoryId:
      normalizeUpper_(
        record['Inventory ID']
      ),

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

    uom:
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
      ),

    eligibilityDate:
      record['Eligibility Date'],

    daysEligible:
      toNumber_(
        record['Days Eligible'],
        0
      ),

    requestedBy:
      normalizeText_(
        record['Requested By']
      ),

    requestedByEmail:
      normalizeLower_(
        record['Requested By Email']
      ),

    approverEmail:
      normalizeLower_(
        record['Approver Email']
      ),

    status:
      normalizeUpper_(
        record['Status']
      ),

    approvedBy:
      normalizeLower_(
        record['Approved By']
      ),

    approvedAt:
      record['Approved At'],

    approvalRemarks:
      normalizeText_(
        record['Approval Remarks']
      ),

    rejectedBy:
      normalizeLower_(
        record['Rejected By']
      ),

    rejectedAt:
      record['Rejected At'],

    rejectionReason:
      normalizeText_(
        record['Rejection Reason']
      ),

    executedBy:
      normalizeLower_(
        record['Executed By']
      ),

    executedAt:
      record['Executed At'],

    freeInventoryId:
      normalizeUpper_(
        record['Free Inventory ID']
      ),

    referenceNo:
      normalizeText_(
        record['Reference No']
      ),

    remarks:
      normalizeText_(
        record['Remarks']
      ),

    updateHistory:
      normalizeText_(
        record['Update History']
      )
  };
}


/**
 * Maps rows into one conversion request.
 */
function mapConversionRequest_(
  records
) {
  const mapped =
    records.map(function (record) {
      return record.rowNumber
        ? record
        : mapConversionRow_(record);
    });

  const first = mapped[0];

  return {
    conversionId:
      first.conversionId,

    requestDate:
      formatDateTime_(
        first.requestDate
      ),

    projectId:
      first.projectId,

    projectCode:
      first.projectCode,

    projectName:
      first.projectName,

    requestedBy:
      first.requestedBy,

    requestedByEmail:
      first.requestedByEmail,

    approverEmail:
      first.approverEmail,

    status:
      first.status,

    approvedBy:
      first.approvedBy,

    approvedAt:
      formatDateTime_(
        first.approvedAt
      ),

    approvalRemarks:
      first.approvalRemarks,

    rejectedBy:
      first.rejectedBy,

    rejectedAt:
      formatDateTime_(
        first.rejectedAt
      ),

    rejectionReason:
      first.rejectionReason,

    executedBy:
      first.executedBy,

    executedAt:
      formatDateTime_(
        first.executedAt
      ),

    referenceNo:
      first.referenceNo,

    remarks:
      first.remarks,

    itemCount:
      mapped.length,

    totalQuantity:
      roundTwo_(
        mapped.reduce(
          function (total, record) {
            return total +
              record.quantity;
          },
          0
        )
      ),

    totalAmount:
      roundTwo_(
        mapped.reduce(
          function (total, record) {
            return total +
              record.amount;
          },
          0
        )
      ),

    items:
      mapped.map(function (record) {
        return {
          inventoryId:
            record.inventoryId,
          skuId:
            record.skuId,
          skuCode:
            record.skuCode,
          skuName:
            record.skuName,
          quantity:
            record.quantity,
          unit:
            record.uom,
          rate:
            record.rate,
          amount:
            record.amount,
          eligibilityDate:
            formatDateTime_(
              record.eligibilityDate
            ),
          daysEligible:
            record.daysEligible,
          freeInventoryId:
            record.freeInventoryId
        };
      })
  };
}


/**
 * Ensures all rows have expected status.
 */
function assertConversionStatus_(
  records,
  expectedStatus
) {
  const statuses = Array.from(
    new Set(
      records.map(function (record) {
        return record.status;
      })
    )
  );

  if (
    statuses.length !== 1 ||
    statuses[0] !== expectedStatus
  ) {
    throw new Error(
      'Conversion request must have status ' +
      expectedStatus +
      '. Current status: ' +
      statuses.join(', ')
    );
  }
}


/**
 * Revalidates conversion stock.
 */
function validateConversionStock_(
  records
) {
  const errors = [];
  const today = new Date();

  records.forEach(function (record) {
    const inventoryRecord =
      getCustomizeInventoryRecord_(
        record.projectId,
        record.skuId
      );

    if (!inventoryRecord) {
      errors.push(
        'Inventory missing for SKU ' +
        record.skuId
      );
      return;
    }

    const availableQty =
      inventoryToNumber_(
        inventoryRecord[
          'Available Qty'
        ],
        0
      );

    const eligibilityDate =
      parseDate_(
        inventoryRecord[
          'Eligibility Date'
        ]
      );

    const eligibleFlag =
      normalizeUpper_(
        inventoryRecord[
          'Eligible for Free'
        ]
      );

    if (
      record.quantity >
      availableQty
    ) {
      errors.push(
        'SKU ' +
        record.skuId +
        ': Available ' +
        availableQty +
        ', Requested ' +
        record.quantity
      );
    }

    if (
      eligibleFlag !== 'YES' ||
      !eligibilityDate ||
      today.getTime() <
        eligibilityDate.getTime()
    ) {
      errors.push(
        'SKU ' +
        record.skuId +
        ' is no longer eligible.'
      );
    }
  });

  if (errors.length) {
    throw new Error(
      'Conversion validation failed:\n' +
      errors.join('\n')
    );
  }

  return true;
}


/**
 * Updates all conversion rows.
 */
function updateConversionRows_(
  records,
  options
) {
  records.forEach(function (record) {
    const history = [
      record.updateHistory,
      options.history
    ]
      .filter(Boolean)
      .join('\n');

    updateObjectRow_(
      getConversionSheetName_(),
      record.rowNumber,
      Object.assign(
        {},
        options.values || {},
        {
          'Status':
            options.status,
          'Update History':
            history
        }
      )
    );
  });

  SpreadsheetApp.flush();
}


/**
 * Deletes conversion rows after failed create.
 */
function deleteConversionRows_(
  rowNumbers
) {
  if (
    !Array.isArray(rowNumbers) ||
    !rowNumbers.length
  ) {
    return;
  }

  const sheet =
    getSystemSheet(
      getConversionSheetName_()
    );

  rowNumbers
    .slice()
    .sort(function (a, b) {
      return b - a;
    })
    .forEach(function (rowNumber) {
      if (
        rowNumber >= 2 &&
        rowNumber <=
          sheet.getLastRow()
      ) {
        sheet.deleteRow(rowNumber);
      }
    });
}


/**
 * Builds conversion history line.
 */
function buildConversionHistoryText_(
  action,
  email,
  dateValue
) {
  return (
    formatDateTime_(
      dateValue || new Date()
    ) +
    ' | ' +
    normalizeLower_(email) +
    ' | ' +
    normalizeText_(action)
  );
}


/**
 * TEST ONLY:
 * Makes PRJ000001 + SKU000001 eligible immediately.
 *
 * Run only in the test workbook.
 */
function testMarkCustomizeStockEligible() {
  const record =
    getCustomizeInventoryRecord_(
      'PRJ000001',
      'SKU000001'
    );

  if (!record) {
    throw new Error(
      'Test customized inventory record not found.'
    );
  }

  const yesterday = new Date(
    new Date().getTime() -
    24 * 60 * 60 * 1000
  );

  updateObjectRow_(
    getCustomizeInventorySheetName_(),
    record._rowNumber,
    {
      'Eligible for Free': 'YES',
      'Eligibility Date':
        yesterday,
      'Updated By':
        getCurrentUserEmail_(),
      'Updated At':
        new Date()
    }
  );

  Logger.log(
    'Test customized stock marked eligible.'
  );

  return true;
}


/**
 * Test: creates one conversion request.
 */
function testCreateConversionRequest() {
  const result =
    createConversionRequest({
      projectId:
        'PRJ000001',

      approverEmail:
        getCurrentUserEmail_(),

      referenceNo:
        'TEST-CONVERSION-' +
        Date.now(),

      requestedBy:
        'Test User',

      remarks:
        'Created from ConversionService test.',

      items: [
        {
          skuId:
            'SKU000001',
          quantity:
            1
        }
      ]
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
 * Test: approves latest pending conversion.
 */
function testApproveLatestConversion() {
  const pending =
    getConversionRequests({
      status:
        CONVERSION_SERVICE_CONFIG
          .STATUS.PENDING,
      pageNumber: 1,
      pageSize: 10
    });

  if (
    !pending.success ||
    !pending.data.records.length
  ) {
    throw new Error(
      'No pending conversion request found.'
    );
  }

  const result =
    approveConversionRequest(
      pending.data.records[0]
        .conversionId,
      'Approved from test.'
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


/**
 * Test: executes latest approved conversion.
 */
function testExecuteLatestConversion() {
  const approved =
    getConversionRequests({
      status:
        CONVERSION_SERVICE_CONFIG
          .STATUS.APPROVED,
      pageNumber: 1,
      pageSize: 10
    });

  if (
    !approved.success ||
    !approved.data.records.length
  ) {
    throw new Error(
      'No approved conversion request found.'
    );
  }

  const result =
    executeApprovedConversion(
      approved.data.records[0]
        .conversionId
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


/**
 * Tests conversion listing and summary.
 */
function testConversionService() {
  const result = {
    eligible:
      getEligibleCustomizeInventory({
        pageNumber: 1,
        pageSize: 20
      }),

    requests:
      getConversionRequests({
        pageNumber: 1,
        pageSize: 20
      }),

    summary:
      getConversionSummary()
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
