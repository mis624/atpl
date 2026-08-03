/**
 * PROJECT INVENTORY MANAGEMENT SYSTEM
 * File: ApprovalService.gs
 *
 * Handles:
 * - Cross-project customized-stock approval requests
 * - HOD/Admin approval and rejection
 * - Email approval token generation
 * - Approved outward execution by an authorized doer
 * - Stock revalidation before approval and execution
 * - Approval listing, details, and summary
 *
 * BUSINESS FLOW:
 * 1. Doer creates a request to use customized stock on another project.
 * 2. HOD/Admin approves or rejects the request.
 * 3. After approval, an authorized doer executes the outward.
 * 4. Source customized inventory is reduced.
 * 5. Outward sheet records both source and destination projects.
 *
 * EmailService.gs can later call:
 * - approveRequestByToken(token, remarks)
 * - rejectRequestByToken(token, reason)
 */

const APPROVAL_SERVICE_CONFIG = Object.freeze({
  SHEET_NAME: 'Approval_Requests',

  ID_PREFIX: 'APR',

  REQUEST_TYPE: Object.freeze({
    CROSS_PROJECT_OUTWARD: 'CROSS PROJECT OUTWARD'
  }),

  STATUS: Object.freeze({
    PENDING: 'PENDING',
    APPROVED: 'APPROVED',
    REJECTED: 'REJECTED',
    EXECUTED: 'EXECUTED',
    CANCELLED: 'CANCELLED',
    EXPIRED: 'EXPIRED'
  }),

  INVENTORY_TYPE: 'CUSTOMIZE',
  INVENTORY_MOVEMENT_TYPE: 'CUSTOMIZE INVENTORY',

  MAX_ITEMS: 100,
  TOKEN_EXPIRY_DAYS: 7
});


/**
 * Returns Approval_Requests sheet name.
 */
function getApprovalSheetName_() {
  if (
    typeof APP_CONFIG !== 'undefined' &&
    APP_CONFIG.SHEETS &&
    APP_CONFIG.SHEETS.APPROVAL_REQUESTS
  ) {
    return APP_CONFIG.SHEETS.APPROVAL_REQUESTS;
  }

  return APPROVAL_SERVICE_CONFIG.SHEET_NAME;
}


/**
 * Creates/verifies Approval_Requests headers.
 */
function ensureApprovalHeaders() {
  return safeExecute_(function () {
    requireRole_(APP_CONFIG.USER_ROLES.ADMIN);

    const sheet = getSystemSheet(
      getApprovalSheetName_()
    );

    const headers = [
      'Request ID',
      'Request Date',
      'Request Type',
      'Source Project ID',
      'Source Project',
      'Destination Project ID',
      'Destination Project',
      'Reference No',
      'Requested By',
      'Requested By Email',
      'Approver Email',
      'SKU ID',
      'SKU Code',
      'SKU Name',
      'Qty',
      'Unit',
      'Rate',
      'Amount',
      'Status',
      'Approval Token',
      'Token Expiry',
      'Approved By',
      'Approved At',
      'Approval Remarks',
      'Rejected By',
      'Rejected At',
      'Rejection Reason',
      'Executed By',
      'Executed At',
      'Outward No',
      'Remarks',
      'Update History',
      'Approved HOD Name'
    ];

    ensureExactApprovalHeaders_(
      sheet,
      headers
    );

    applyApprovalSheetFormatting_(
      sheet
    );

    return successResponse_(
      'Approval request headers verified successfully.',
      {
        sheetName: getApprovalSheetName_(),
        headers: headers
      }
    );
  }, 'Unable to verify approval request headers.');
}


/**
 * Ensures exact header order.
 */
function ensureExactApprovalHeaders_(
  sheet,
  headers
) {
  if (sheet.getMaxColumns() < headers.length) {
    sheet.insertColumnsAfter(
      sheet.getMaxColumns(),
      headers.length - sheet.getMaxColumns()
    );
  }

  sheet
    .getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setFontWeight('bold');

  sheet.setFrozenRows(1);
}


/**
 * Applies number formats and validations.
 */
function applyApprovalSheetFormatting_(sheet) {
  const dataRows = Math.max(
    sheet.getMaxRows() - 1,
    1
  );

  // Remove any old or incorrect validations.
  sheet
    .getRange(
      2,
      1,
      dataRows,
      33
    )
    .clearDataValidations();

  const statusRule = SpreadsheetApp
    .newDataValidation()
    .requireValueInList(
      Object.keys(
        APPROVAL_SERVICE_CONFIG.STATUS
      ).map(function (key) {
        return APPROVAL_SERVICE_CONFIG.STATUS[key];
      }),
      true
    )
    .setAllowInvalid(false)
    .build();

  // Quantity.
  sheet
    .getRange(2, 15, dataRows, 1)
    .setNumberFormat('0.000');

  // Rate and amount.
  sheet
    .getRange(2, 17, dataRows, 2)
    .setNumberFormat('0.00');

  // Status.
  sheet
    .getRange(2, 19, dataRows, 1)
    .setDataValidation(statusRule);

  // Date/time columns.
  [
    2, 21, 23, 26, 29
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

  sheet.setFrozenRows(1);
  SpreadsheetApp.flush();
}


/**
 * Creates a cross-project customized-stock request.
 *
 * Request:
 * {
 *   sourceProjectId: 'PRJ000001',
 *   destinationProjectId: 'PRJ000002',
 *   referenceNo: 'REQ-001',
 *   requestedBy: 'Aman',
 *   approverEmail: 'hod@company.com', // optional when project HOD exists
 *   remarks: '',
 *   items: [
 *     { skuId: 'SKU000001', quantity: 2 }
 *   ]
 * }
 */
function createCrossProjectApprovalRequest(
  requestData
) {
  return safeExecute_(function () {
    const session =
      requirePermission_('canOutward');

    const payload =
      validateApprovalRequest_(
        requestData
      );

    const requestId =
      generateNextId_(
        'APPROVAL_REQUEST',
        APPROVAL_SERVICE_CONFIG.ID_PREFIX
      );

    const now = new Date();
    const token = generateSecureToken_();
    const tokenExpiry = new Date(
      now.getTime() +
      APPROVAL_SERVICE_CONFIG
        .TOKEN_EXPIRY_DAYS *
      24 * 60 * 60 * 1000
    );

    const history =
      buildApprovalHistoryText_(
        'Approval request created',
        session.email,
        now
      );

    const savedRows = [];

    try {
      payload.items.forEach(
        function (item) {
          const amount = roundTwo_(
            item.quantity * item.rate
          );

          const rowNumber =
            appendObjectRow_(
              getApprovalSheetName_(),
              {
                'Request ID':
                  requestId,

                'Request Date':
                  now,

                'Request Type':
                  APPROVAL_SERVICE_CONFIG
                    .REQUEST_TYPE
                    .CROSS_PROJECT_OUTWARD,

                'Source Project ID':
                  payload.sourceProject.projectId,

                'Source Project':
                  payload.sourceProject.projectName,

                'Destination Project ID':
                  payload.destinationProject.projectId,

                'Destination Project':
                  payload.destinationProject.projectName,

                'Reference No':
                  payload.referenceNo,

                'Requested By':
                  payload.requestedBy,

                'Requested By Email':
                  session.email,

                'Approver Email':
                  payload.approverEmail,

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
                  APPROVAL_SERVICE_CONFIG
                    .STATUS.PENDING,

                'Approval Token':
                  token,

                'Token Expiry':
                  tokenExpiry,

                'Approved By': '',
                'Approved At': '',
                'Approval Remarks': '',

                'Rejected By': '',
                'Rejected At': '',
                'Rejection Reason': '',

                'Executed By': '',
                'Executed At': '',
                'Outward No': '',

                'Remarks':
                  payload.remarks,

                'Update History':
                  history,

                'Approved HOD Name':
                  ''
              }
            );

          savedRows.push(rowNumber);
        }
      );

      addTransactionLog_(
        session.email,
        'APPROVAL',
        'CREATE REQUEST',
        requestId,
        {
          requestType:
            APPROVAL_SERVICE_CONFIG
              .REQUEST_TYPE
              .CROSS_PROJECT_OUTWARD,

          sourceProjectId:
            payload.sourceProject.projectId,

          destinationProjectId:
            payload.destinationProject.projectId,

          approverEmail:
            payload.approverEmail,

          itemCount:
            payload.items.length,

          totalQuantity:
            roundTwo_(
              payload.items.reduce(
                function (total, item) {
                  return total + item.quantity;
                },
                0
              )
            )
        }
      );

      /*
       * EmailService.gs can implement this optional hook.
       * Approval request remains successfully created
       * even when the email service has not been added yet.
       */
      let emailSent = false;
      let emailMessage =
        'Email service is not configured yet.';

      if (
        typeof sendApprovalRequestEmail_ ===
        'function'
      ) {
        try {
          const emailResult =
            sendApprovalRequestEmail_({
              requestId: requestId,
              token: token,
              tokenExpiry: tokenExpiry,
              approverEmail:
                payload.approverEmail,
              sourceProject:
                payload.sourceProject,
              destinationProject:
                payload.destinationProject,
              referenceNo:
                payload.referenceNo,
              requestedBy:
                payload.requestedBy,
              requestedByEmail:
                session.email,
              remarks:
                payload.remarks,
              items:
                payload.items
            });

          emailSent =
            emailResult === true ||
            (
              emailResult &&
              emailResult.success === true
            );

          emailMessage = emailSent
            ? 'Approval email sent successfully.'
            : (
                emailResult &&
                emailResult.message
                  ? emailResult.message
                  : 'Approval email was not sent.'
              );
        } catch (emailError) {
          emailMessage =
            emailError.message ||
            'Approval email failed.';
        }
      }

      return successResponse_(
        'Approval request created successfully.',
        {
          requestId: requestId,
          status:
            APPROVAL_SERVICE_CONFIG
              .STATUS.PENDING,
          approverEmail:
            payload.approverEmail,
          tokenExpiry:
            formatDateTime_(tokenExpiry),
          itemCount:
            payload.items.length,
          totalQuantity:
            roundTwo_(
              payload.items.reduce(
                function (total, item) {
                  return total + item.quantity;
                },
                0
              )
            ),
          emailSent: emailSent,
          emailMessage: emailMessage
        }
      );

    } catch (error) {
      deleteApprovalRows_(savedRows);
      throw error;
    }
  }, 'Unable to create approval request.');
}


/**
 * Validates a cross-project request.
 */
function validateApprovalRequest_(
  requestData
) {
  if (
    !requestData ||
    typeof requestData !== 'object'
  ) {
    throw new Error(
      'Invalid approval request data.'
    );
  }

  const sourceProjectId =
    normalizeUpper_(
      requestData.sourceProjectId
    );

  const destinationProjectId =
    normalizeUpper_(
      requestData.destinationProjectId
    );

  const referenceNo =
    normalizeText_(
      requestData.referenceNo
    );

  const requestedBy =
    normalizeText_(
      requestData.requestedBy
    );

  const remarks =
    normalizeText_(
      requestData.remarks
    );

  if (!sourceProjectId) {
    throw new Error(
      'Source Project is required.'
    );
  }

  if (!destinationProjectId) {
    throw new Error(
      'Destination Project is required.'
    );
  }

  if (
    sourceProjectId ===
    destinationProjectId
  ) {
    throw new Error(
      'Source and Destination Projects must be different.'
    );
  }

  if (!referenceNo) {
    throw new Error(
      'Reference No is required.'
    );
  }

  if (!requestedBy) {
    throw new Error(
      'Requested By is required.'
    );
  }

  const sourceRecord =
    getProjectRecordById_(
      sourceProjectId
    );

  const destinationRecord =
    getProjectRecordById_(
      destinationProjectId
    );

  if (!sourceRecord) {
    throw new Error(
      'Source Project not found: ' +
      sourceProjectId
    );
  }

  if (!destinationRecord) {
    throw new Error(
      'Destination Project not found: ' +
      destinationProjectId
    );
  }

  if (
    normalizeUpper_(
      sourceRecord['Status']
    ) !== 'ACTIVE'
  ) {
    throw new Error(
      'Source Project is not active.'
    );
  }

  if (
    normalizeUpper_(
      destinationRecord['Status']
    ) !== 'ACTIVE'
  ) {
    throw new Error(
      'Destination Project is not active.'
    );
  }

  const sourceProject =
    mapApprovalProject_(sourceRecord);

  const destinationProject =
    mapApprovalProject_(
      destinationRecord
    );

  let approverEmail =
    normalizeLower_(
      requestData.approverEmail
    );

  if (!approverEmail) {
    approverEmail =
      normalizeLower_(
        sourceRecord['HOD Email']
      );
  }

  /*
   * During initial setup, Admin may explicitly supply
   * approverEmail until Project_Master HOD Email is filled.
   */
  if (!approverEmail) {
    throw new Error(
      'Approver Email is required. Add HOD Email in the source project or provide approverEmail.'
    );
  }

  if (!isValidEmail_(approverEmail)) {
    throw new Error(
      'Invalid Approver Email: ' +
      approverEmail
    );
  }

  if (
    !Array.isArray(requestData.items) ||
    requestData.items.length === 0
  ) {
    throw new Error(
      'At least one approval item is required.'
    );
  }

  if (
    requestData.items.length >
    APPROVAL_SERVICE_CONFIG.MAX_ITEMS
  ) {
    throw new Error(
      'A maximum of ' +
      APPROVAL_SERVICE_CONFIG.MAX_ITEMS +
      ' items are allowed.'
    );
  }

  const duplicateMap = {};

  const items =
    requestData.items.map(
      function (inputItem, index) {
        const lineNumber = index + 1;

        if (
          !inputItem ||
          typeof inputItem !== 'object'
        ) {
          throw new Error(
            'Invalid item at line ' +
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

        if (duplicateMap[skuId]) {
          throw new Error(
            'Duplicate SKU found: ' +
            skuId
          );
        }

        duplicateMap[skuId] = true;

        const skuRecord =
          getSkuRecordById_(skuId);

        if (!skuRecord) {
          throw new Error(
            'SKU not found: ' + skuId
          );
        }

        if (
          normalizeUpper_(
            skuRecord['Status']
          ) !== 'ACTIVE'
        ) {
          throw new Error(
            'SKU is inactive: ' + skuId
          );
        }

        const inventoryRecord =
          getCustomizeInventoryRecord_(
            sourceProjectId,
            skuId
          );

        if (!inventoryRecord) {
          throw new Error(
            'Customized inventory was not found for source project and SKU: ' +
            skuId
          );
        }

        const quantity =
          toPositiveNumber_(
            inputItem.quantity,
            'Quantity at line ' +
              lineNumber
          );

        const availableQty =
          inventoryToNumber_(
            inventoryRecord[
              'Available Qty'
            ],
            0
          );

        if (quantity > availableQty) {
          throw new Error(
            'Insufficient customized stock at line ' +
            lineNumber +
            '. SKU: ' +
            skuId +
            ', Available: ' +
            availableQty +
            ', Requested: ' +
            quantity
          );
        }

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
              inventoryToNumber_(
                inventoryRecord[
                  'Average Rate'
                ],
                0
              )
            )
        };
      }
    );

  return {
    sourceProject: sourceProject,
    destinationProject:
      destinationProject,
    referenceNo: referenceNo,
    requestedBy: requestedBy,
    approverEmail: approverEmail,
    remarks: remarks,
    items: items
  };
}


/**
 * Approves a pending request from the application.
 */
function approveApprovalRequest(
  requestId,
  approvalRemarks
) {
  return safeExecute_(function () {
    const session =
      requirePermission_('canApprove');

    return approveApprovalRequestInternal_(
      requestId,
      session.email,
      approvalRemarks,
      ''
    );
  }, 'Unable to approve request.');
}


/**
 * Approves a request using an email token.
 */
function approveRequestByToken(
  token,
  approvalRemarks
) {
  return safeExecute_(function () {
    const cleanToken =
      normalizeText_(token);

    if (!cleanToken) {
      throw new Error(
        'Approval token is required.'
      );
    }

    const records =
      getApprovalRowsByToken_(
        cleanToken
      );

    validateApprovalTokenRecords_(
      records
    );

    const firstRecord = records[0];

    return approveApprovalRequestInternal_(
      firstRecord.requestId,
      firstRecord.approverEmail,
      approvalRemarks,
      cleanToken
    );
  }, 'Unable to approve request.');
}


/**
 * Shared approval logic.
 */

/**
 * Returns the approver's display name from Users.
 * Falls back to the approval email when no user name is found.
 */
function getApprovalHodNameByEmail_(
  email
) {
  const cleanEmail =
    normalizeLower_(email);

  if (!cleanEmail) {
    return '';
  }

  try {
    const users =
      getSheetObjects_(
        APP_CONFIG.SHEETS.USERS
      );

    const user =
      users.find(function (record) {
        return normalizeLower_(
          record['Email'] ||
          record['Email Address'] ||
          record['User Email']
        ) === cleanEmail;
      });

    if (user) {
      return normalizeText_(
        user['Employee Name'] ||
        user['User Name'] ||
        user['Name']
      ) || cleanEmail;
    }
  } catch (error) {
    // Fall back to email.
  }

  return cleanEmail;
}

function approveApprovalRequestInternal_(
  requestId,
  approvedBy,
  approvalRemarks,
  token
) {
  const cleanRequestId =
    normalizeUpper_(requestId);

  if (!cleanRequestId) {
    throw new Error(
      'Request ID is required.'
    );
  }

  const records =
    getApprovalRows_(
      cleanRequestId
    );

  if (records.length === 0) {
    throw new Error(
      'Approval request not found: ' +
      cleanRequestId
    );
  }

  assertSingleApprovalStatus_(
    records,
    APPROVAL_SERVICE_CONFIG
      .STATUS.PENDING
  );

  if (token) {
    validateApprovalTokenRecords_(
      records
    );
  }

  validateApprovalStock_(records);

  const now = new Date();
  const remarks =
    normalizeText_(
      approvalRemarks
    );

  updateApprovalRowsStatus_(
    records,
    {
      status:
        APPROVAL_SERVICE_CONFIG
          .STATUS.APPROVED,

      values: {
        'Approved By':
          normalizeLower_(approvedBy),
        'Approved HOD Name':
          getApprovalHodNameByEmail_(
            approvedBy
          ),
        'Approved At':
          now,
        'Approval Remarks':
          remarks
      },

      historyText:
        buildApprovalHistoryText_(
          'Approval request approved',
          approvedBy,
          now
        )
    }
  );

  addTransactionLog_(
    approvedBy,
    'APPROVAL',
    'APPROVE',
    cleanRequestId,
    {
      itemCount: records.length,
      remarks: remarks
    }
  );

  return successResponse_(
    'Approval request approved successfully. The doer can now execute the outward.',
    {
      requestId: cleanRequestId,
      status:
        APPROVAL_SERVICE_CONFIG
          .STATUS.APPROVED,
      approvedBy:
        normalizeLower_(approvedBy),
      approvedHodName:
        getApprovalHodNameByEmail_(
          approvedBy
        ),
      approvedAt:
        formatDateTime_(now)
    }
  );
}


/**
 * Rejects a pending request from the application.
 */
function rejectApprovalRequest(
  requestId,
  rejectionReason
) {
  return safeExecute_(function () {
    const session =
      requirePermission_('canApprove');

    return rejectApprovalRequestInternal_(
      requestId,
      session.email,
      rejectionReason,
      ''
    );
  }, 'Unable to reject request.');
}


/**
 * Rejects a request using an email token.
 */
function rejectRequestByToken(
  token,
  rejectionReason
) {
  return safeExecute_(function () {
    const cleanToken =
      normalizeText_(token);

    if (!cleanToken) {
      throw new Error(
        'Approval token is required.'
      );
    }

    const records =
      getApprovalRowsByToken_(
        cleanToken
      );

    validateApprovalTokenRecords_(
      records
    );

    return rejectApprovalRequestInternal_(
      records[0].requestId,
      records[0].approverEmail,
      rejectionReason,
      cleanToken
    );
  }, 'Unable to reject request.');
}


/**
 * Shared rejection logic.
 */
function rejectApprovalRequestInternal_(
  requestId,
  rejectedBy,
  rejectionReason,
  token
) {
  const cleanRequestId =
    normalizeUpper_(requestId);

  const reason =
    normalizeText_(
      rejectionReason
    );

  if (!cleanRequestId) {
    throw new Error(
      'Request ID is required.'
    );
  }

  if (!reason) {
    throw new Error(
      'Rejection reason is required.'
    );
  }

  const records =
    getApprovalRows_(
      cleanRequestId
    );

  if (records.length === 0) {
    throw new Error(
      'Approval request not found: ' +
      cleanRequestId
    );
  }

  assertSingleApprovalStatus_(
    records,
    APPROVAL_SERVICE_CONFIG
      .STATUS.PENDING
  );

  if (token) {
    validateApprovalTokenRecords_(
      records
    );
  }

  const now = new Date();

  updateApprovalRowsStatus_(
    records,
    {
      status:
        APPROVAL_SERVICE_CONFIG
          .STATUS.REJECTED,

      values: {
        'Rejected By':
          normalizeLower_(rejectedBy),
        'Rejected At':
          now,
        'Rejection Reason':
          reason
      },

      historyText:
        buildApprovalHistoryText_(
          'Approval request rejected: ' +
          reason,
          rejectedBy,
          now
        )
    }
  );

  addTransactionLog_(
    rejectedBy,
    'APPROVAL',
    'REJECT',
    cleanRequestId,
    {
      reason: reason,
      itemCount: records.length
    }
  );

  return successResponse_(
    'Approval request rejected successfully.',
    {
      requestId: cleanRequestId,
      status:
        APPROVAL_SERVICE_CONFIG
          .STATUS.REJECTED,
      rejectedBy:
        normalizeLower_(rejectedBy),
      rejectedAt:
        formatDateTime_(now),
      rejectionReason: reason
    }
  );
}


/**
 * Executes an approved cross-project outward.
 *
 * HOD only approves.
 * An authorized doer executes the outward after approval.
 */
function executeApprovedOutward(
  requestId
) {
  return safeExecute_(function () {
    const session =
      requirePermission_('canOutward');

    const cleanRequestId =
      normalizeUpper_(requestId);

    if (!cleanRequestId) {
      throw new Error(
        'Request ID is required.'
      );
    }

    const records =
      getApprovalRows_(
        cleanRequestId
      );

    if (records.length === 0) {
      throw new Error(
        'Approval request not found: ' +
        cleanRequestId
      );
    }

    assertSingleApprovalStatus_(
      records,
      APPROVAL_SERVICE_CONFIG
        .STATUS.APPROVED
    );

    validateApprovalStock_(
      records
    );

    const firstRecord = records[0];

    const outwardNo =
      generateNextId_(
        'OUTWARD',
        OUTWARD_SERVICE_CONFIG.ID_PREFIX
      );

    const now = new Date();
    const inventoryResults = [];
    const rollbackItems = [];
    const outwardRows = [];

    try {
      records.forEach(function (record) {
        const movementResult =
          applyInventoryMovement({
            inventoryType:
              APPROVAL_SERVICE_CONFIG
                .INVENTORY_MOVEMENT_TYPE,

            movementType:
              INVENTORY_SERVICE_CONFIG
                .MOVEMENT_TYPE.OUTWARD,

            skuId:
              record.skuId,

            projectId:
              record.sourceProjectId,

            quantity:
              record.quantity,

            rate:
              record.rate,

            referenceId:
              outwardNo,

            remarks:
              'Approved cross-project outward to ' +
              record.destinationProject
          });

        if (
          !movementResult ||
          movementResult.success !== true
        ) {
          throw new Error(
            movementResult &&
            movementResult.message
              ? movementResult.message
              : 'Inventory outward failed.'
          );
        }

        inventoryResults.push(
          movementResult.data
        );

        rollbackItems.push({
          inventoryType:
            APPROVAL_SERVICE_CONFIG
              .INVENTORY_MOVEMENT_TYPE,
          skuId:
            record.skuId,
          projectId:
            record.sourceProjectId,
          quantity:
            record.quantity,
          rate:
            record.rate,
          referenceId:
            outwardNo
        });
      });

      records.forEach(function (record) {
        const rowNumber =
          appendObjectRow_(
            getOutwardSheetName_(),
            {
              'Outward No':
                outwardNo,
              'Date':
                now,
              'Inventory Type':
                APPROVAL_SERVICE_CONFIG
                  .INVENTORY_TYPE,
              'Source Project ID':
                record.sourceProjectId,
              'Source Project':
                record.sourceProject,
              'Destination Project ID':
                record.destinationProjectId,
              'Destination Project':
                record.destinationProject,
              'Reference No':
                record.referenceNo,
              'Issued To':
                record.requestedBy,
              'SKU ID':
                record.skuId,
              'SKU Code':
                record.skuCode,
              'SKU Name':
                record.skuName,
              'Qty':
                record.quantity,
              'Unit':
                record.unit,
              'Rate':
                record.rate,
              'Amount':
                roundTwo_(
                  record.quantity *
                  record.rate
                ),
              'Status':
                OUTWARD_SERVICE_CONFIG
                  .STATUS.COMPLETED,
              'Remarks':
                'Executed against approval ' +
                cleanRequestId +
                (
                  record.remarks
                    ? '. ' + record.remarks
                    : ''
                ),
              'Entered By':
                session.email
            }
          );

        outwardRows.push(rowNumber);
      });

      updateApprovalRowsStatus_(
        records,
        {
          status:
            APPROVAL_SERVICE_CONFIG
              .STATUS.EXECUTED,

          values: {
            'Executed By':
              session.email,
            'Executed At':
              now,
            'Outward No':
              outwardNo
          },

          historyText:
            buildApprovalHistoryText_(
              'Approved outward executed: ' +
              outwardNo,
              session.email,
              now
            )
        }
      );

      addTransactionLog_(
        session.email,
        'APPROVAL',
        'EXECUTE OUTWARD',
        cleanRequestId,
        {
          outwardNo: outwardNo,
          itemCount: records.length,
          sourceProjectId:
            firstRecord.sourceProjectId,
          destinationProjectId:
            firstRecord.destinationProjectId
        }
      );

      return successResponse_(
        'Approved outward executed successfully.',
        {
          requestId:
            cleanRequestId,
          status:
            APPROVAL_SERVICE_CONFIG
              .STATUS.EXECUTED,
          outwardNo:
            outwardNo,
          sourceProjectId:
            firstRecord.sourceProjectId,
          destinationProjectId:
            firstRecord.destinationProjectId,
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
          items:
            inventoryResults
        }
      );

    } catch (error) {
      rollbackOutwardRows_(
        outwardRows
      );

      rollbackOutwardInventory_(
        rollbackItems
      );

      throw error;
    }
  }, 'Unable to execute approved outward.');
}


/**
 * Returns one approval request.
 */
function getApprovalRequest(
  requestId
) {
  return safeExecute_(function () {
    requireAuthenticatedUser_();

    const cleanRequestId =
      normalizeUpper_(requestId);

    if (!cleanRequestId) {
      throw new Error(
        'Request ID is required.'
      );
    }

    const records =
      getApprovalRows_(
        cleanRequestId
      );

    if (records.length === 0) {
      throw new Error(
        'Approval request not found: ' +
        cleanRequestId
      );
    }

    return successResponse_(
      'Approval request loaded successfully.',
      mapApprovalRequest_(records)
    );
  }, 'Unable to load approval request.');
}


/**
 * Returns approval requests with filters.
 */
function getApprovalRequests(
  filters
) {
  return safeExecute_(function () {
    const session =
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

    const sourceProjectId =
      normalizeUpper_(
        filters.sourceProjectId
      );

    const destinationProjectId =
      normalizeUpper_(
        filters.destinationProjectId
      );

    const records =
      getSheetObjects_(
        getApprovalSheetName_()
      ).filter(function (record) {
        if (
          !normalizeText_(
            record['Request ID']
          )
        ) {
          return false;
        }

        const recordStatus =
          normalizeUpper_(
            record['Status']
          );

        if (
          status &&
          recordStatus !== status
        ) {
          return false;
        }

        if (
          sourceProjectId &&
          normalizeUpper_(
            record[
              'Source Project ID'
            ]
          ) !== sourceProjectId
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

        /*
         * Doer can see requests created by them.
         * HOD/Admin/approver can see requests assigned to them.
         */
        if (
          !session.permissions.isAdmin &&
          !session.permissions.canApprove
        ) {
          const requestedByEmail =
            normalizeLower_(
              record[
                'Requested By Email'
              ]
            );

          if (
            requestedByEmail !==
            session.email
          ) {
            return false;
          }
        }

        if (search) {
          const searchable = [
            record['Request ID'],
            record['Source Project'],
            record['Destination Project'],
            record['Reference No'],
            record['Requested By'],
            record['Requested By Email'],
            record['Approver Email'],
            record['SKU Code'],
            record['SKU Name'],
            record['Status'],
            record['Outward No']
          ]
            .map(normalizeLower_)
            .join(' ');

          if (
            searchable.indexOf(search) ===
            -1
          ) {
            return false;
          }
        }

        return true;
      });

    const requestMap = {};

    records.forEach(function (record) {
      const requestId =
        normalizeUpper_(
          record['Request ID']
        );

      if (!requestMap[requestId]) {
        requestMap[requestId] = [];
      }

      requestMap[requestId].push(record);
    });

    const requests =
      Object.keys(requestMap)
        .map(function (id) {
          return mapApprovalRequest_(
            requestMap[id]
          );
        })
        .sort(function (first, second) {
          const firstDate =
            parseDate_(first.requestDate) ||
            new Date(0);

          const secondDate =
            parseDate_(second.requestDate) ||
            new Date(0);

          return (
            secondDate.getTime() -
            firstDate.getTime()
          );
        });

    return successResponse_(
      'Approval requests loaded successfully.',
      paginateRecords_(
        requests,
        filters.pageNumber || 1,
        filters.pageSize || 20
      )
    );
  }, 'Unable to load approval requests.');
}


/**
 * Returns pending approval requests.
 */
function getPendingApprovalRequests() {
  return getApprovalRequests({
    status:
      APPROVAL_SERVICE_CONFIG
        .STATUS.PENDING,
    pageNumber: 1,
    pageSize: 500
  });
}


/**
 * Returns approval summary.
 */
function getApprovalSummary() {
  return safeExecute_(function () {
    requireAuthenticatedUser_();

    const records =
      getSheetObjects_(
        getApprovalSheetName_()
      ).filter(function (record) {
        return Boolean(
          normalizeText_(
            record['Request ID']
          )
        );
      });

    const requestStatusMap = {};

    records.forEach(function (record) {
      const requestId =
        normalizeUpper_(
          record['Request ID']
        );

      if (!requestStatusMap[requestId]) {
        requestStatusMap[requestId] =
          normalizeUpper_(
            record['Status']
          );
      }
    });

    const summary = {
      totalRequests: 0,
      pending: 0,
      approved: 0,
      rejected: 0,
      executed: 0,
      cancelled: 0,
      expired: 0
    };

    Object.keys(requestStatusMap)
      .forEach(function (requestId) {
        const status =
          requestStatusMap[requestId];

        summary.totalRequests++;

        if (
          status ===
          APPROVAL_SERVICE_CONFIG
            .STATUS.PENDING
        ) {
          summary.pending++;
        } else if (
          status ===
          APPROVAL_SERVICE_CONFIG
            .STATUS.APPROVED
        ) {
          summary.approved++;
        } else if (
          status ===
          APPROVAL_SERVICE_CONFIG
            .STATUS.REJECTED
        ) {
          summary.rejected++;
        } else if (
          status ===
          APPROVAL_SERVICE_CONFIG
            .STATUS.EXECUTED
        ) {
          summary.executed++;
        } else if (
          status ===
          APPROVAL_SERVICE_CONFIG
            .STATUS.CANCELLED
        ) {
          summary.cancelled++;
        } else if (
          status ===
          APPROVAL_SERVICE_CONFIG
            .STATUS.EXPIRED
        ) {
          summary.expired++;
        }
      });

    return successResponse_(
      'Approval summary loaded successfully.',
      summary
    );
  }, 'Unable to load approval summary.');
}


/**
 * Finds approval rows by Request ID.
 */
function getApprovalRows_(requestId) {
  const target =
    normalizeUpper_(requestId);

  return getSheetObjects_(
    getApprovalSheetName_()
  )
    .filter(function (record) {
      return (
        normalizeUpper_(
          record['Request ID']
        ) === target
      );
    })
    .map(mapApprovalRow_);
}


/**
 * Finds approval rows by token.
 */
function getApprovalRowsByToken_(token) {
  const target =
    normalizeText_(token);

  return getSheetObjects_(
    getApprovalSheetName_()
  )
    .filter(function (record) {
      return (
        normalizeText_(
          record['Approval Token']
        ) === target
      );
    })
    .map(mapApprovalRow_);
}


/**
 * Maps one sheet row.
 */
function mapApprovalRow_(record) {
  return {
    rowNumber:
      record._rowNumber,

    requestId:
      normalizeUpper_(
        record['Request ID']
      ),

    requestDate:
      record['Request Date'],

    requestType:
      normalizeUpper_(
        record['Request Type']
      ),

    sourceProjectId:
      normalizeUpper_(
        record['Source Project ID']
      ),

    sourceProject:
      normalizeText_(
        record['Source Project']
      ),

    destinationProjectId:
      normalizeUpper_(
        record[
          'Destination Project ID'
        ]
      ),

    destinationProject:
      normalizeText_(
        record[
          'Destination Project'
        ]
      ),

    referenceNo:
      normalizeText_(
        record['Reference No']
      ),

    requestedBy:
      normalizeText_(
        record['Requested By']
      ),

    requestedByEmail:
      normalizeLower_(
        record[
          'Requested By Email'
        ]
      ),

    approverEmail:
      normalizeLower_(
        record['Approver Email']
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
      ),

    status:
      normalizeUpper_(
        record['Status']
      ),

    approvalToken:
      normalizeText_(
        record['Approval Token']
      ),

    tokenExpiry:
      record['Token Expiry'],

    approvedBy:
      normalizeLower_(
        record['Approved By']
      ),

    approvedHodName:
      normalizeText_(
        record['Approved HOD Name']
      ) ||
      getApprovalHodNameByEmail_(
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

    outwardNo:
      normalizeUpper_(
        record['Outward No']
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
 * Maps one project.
 */
function mapApprovalProject_(record) {
  return {
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
      )
  };
}


/**
 * Maps all rows into one request.
 */
function mapApprovalRequest_(records) {
  const normalizedRecords =
    records.map(function (record) {
      return record.rowNumber
        ? record
        : mapApprovalRow_(record);
    });

  const first = normalizedRecords[0];

  return {
    requestId:
      first.requestId,

    requestDate:
      formatDateTime_(
        first.requestDate
      ),

    requestType:
      first.requestType,

    sourceProjectId:
      first.sourceProjectId,

    sourceProject:
      first.sourceProject,

    destinationProjectId:
      first.destinationProjectId,

    destinationProject:
      first.destinationProject,

    referenceNo:
      first.referenceNo,

    requestedBy:
      first.requestedBy,

    requestedByEmail:
      first.requestedByEmail,

    approverEmail:
      first.approverEmail,

    status:
      first.status,

    tokenExpiry:
      formatDateTime_(
        first.tokenExpiry
      ),

    approvedBy:
      first.approvedBy,

    approvedHodName:
      first.approvedHodName,

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

    outwardNo:
      first.outwardNo,

    remarks:
      first.remarks,

    itemCount:
      normalizedRecords.length,

    totalQuantity:
      roundTwo_(
        normalizedRecords.reduce(
          function (total, record) {
            return total +
              record.quantity;
          },
          0
        )
      ),

    totalAmount:
      roundTwo_(
        normalizedRecords.reduce(
          function (total, record) {
            return total +
              record.amount;
          },
          0
        )
      ),

    items:
      normalizedRecords.map(
        function (record) {
          return {
            skuId:
              record.skuId,
            skuCode:
              record.skuCode,
            skuName:
              record.skuName,
            quantity:
              record.quantity,
            unit:
              record.unit,
            rate:
              record.rate,
            amount:
              record.amount
          };
        }
      )
  };
}


/**
 * Ensures all rows have one expected status.
 */
function assertSingleApprovalStatus_(
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
      'Request must have status ' +
      expectedStatus +
      '. Current status: ' +
      statuses.join(', ')
    );
  }
}


/**
 * Validates token and expiry.
 */
function validateApprovalTokenRecords_(
  records
) {
  if (
    !Array.isArray(records) ||
    records.length === 0
  ) {
    throw new Error(
      'Invalid approval token.'
    );
  }

  assertSingleApprovalStatus_(
    records,
    APPROVAL_SERVICE_CONFIG
      .STATUS.PENDING
  );

  const expiry =
    parseDate_(
      records[0].tokenExpiry
    );

  if (
    !expiry ||
    new Date().getTime() >
      expiry.getTime()
  ) {
    throw new Error(
      'Approval token has expired.'
    );
  }
}


/**
 * Revalidates stock before approval/execution.
 */
function validateApprovalStock_(records) {
  const errors = [];

  records.forEach(function (record) {
    const inventoryRecord =
      getCustomizeInventoryRecord_(
        record.sourceProjectId,
        record.skuId
      );

    const availableQty =
      inventoryRecord
        ? inventoryToNumber_(
            inventoryRecord[
              'Available Qty'
            ],
            0
          )
        : 0;

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
  });

  if (errors.length > 0) {
    throw new Error(
      'Insufficient customized stock:\n' +
      errors.join('\n')
    );
  }

  return true;
}


/**
 * Updates all request rows.
 */
function updateApprovalRowsStatus_(
  records,
  options
) {
  const sheet =
    getSystemSheet(
      getApprovalSheetName_()
    );

  records.forEach(function (record) {
    const history = [
      record.updateHistory,
      options.historyText
    ]
      .filter(Boolean)
      .join('\n');

    const updateData =
      Object.assign(
        {},
        options.values || {},
        {
          'Status': options.status,
          'Update History': history
        }
      );

    updateObjectRow_(
      getApprovalSheetName_(),
      record.rowNumber,
      updateData
    );
  });

  SpreadsheetApp.flush();
}


/**
 * Deletes rows saved during a failed request.
 */
function deleteApprovalRows_(
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
      getApprovalSheetName_()
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
 * Builds update history text.
 */
function buildApprovalHistoryText_(
  actionText,
  userEmail,
  dateValue
) {
  return (
    formatDateTime_(
      dateValue || new Date()
    ) +
    ' | ' +
    normalizeLower_(userEmail) +
    ' | ' +
    normalizeText_(actionText)
  );
}


/**
 * Test setup:
 * - PRJ000001 must contain customized SKU000001 stock.
 * - PRJ000002 must exist and be ACTIVE.
 * - Current user must have canOutward/canApprove.
 */
function testCreateApprovalRequest() {
  const result =
    createCrossProjectApprovalRequest({
      sourceProjectId:
        'PRJ000001',

      destinationProjectId:
        'PRJ000002',

      referenceNo:
        'TEST-APPROVAL-001',

      requestedBy:
        'Test User',

      /*
       * Your project currently has blank HOD Email.
       * This explicit email allows the test to run.
       */
      approverEmail:
        getCurrentUserEmail_(),

      remarks:
        'Created from ApprovalService test.',

      items: [
        {
          skuId:
            'SKU000001',
          quantity:
            2
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
 * Approves the latest pending test request.
 */
function testApproveLatestRequest() {
  const pending =
    getPendingApprovalRequests();

  if (
    !pending.success ||
    !pending.data.records.length
  ) {
    throw new Error(
      'No pending approval request found.'
    );
  }

  const requestId =
    pending.data.records[0].requestId;

  const result =
    approveApprovalRequest(
      requestId,
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
 * Executes the latest approved request.
 */
function testExecuteLatestApprovedOutward() {
  const approved =
    getApprovalRequests({
      status:
        APPROVAL_SERVICE_CONFIG
          .STATUS.APPROVED,
      pageNumber: 1,
      pageSize: 10
    });

  if (
    !approved.success ||
    !approved.data.records.length
  ) {
    throw new Error(
      'No approved request found.'
    );
  }

  const requestId =
    approved.data.records[0].requestId;

  const result =
    executeApprovedOutward(
      requestId
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
 * Tests approval listing and summary.
 */
function testApprovalService() {
  const result = {
    requests:
      getApprovalRequests({
        pageNumber: 1,
        pageSize: 10
      }),

    pending:
      getPendingApprovalRequests(),

    summary:
      getApprovalSummary()
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
