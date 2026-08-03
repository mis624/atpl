/**
 * PROJECT INVENTORY MANAGEMENT SYSTEM
 * File: DashboardService.gs
 *
 * Handles:
 * - Main dashboard summary
 * - Free inventory summary
 * - Customized inventory summary
 * - Low-stock summary
 * - 90-day eligible customized stock
 * - Pending approvals
 * - Pending conversions
 * - Today inward and outward
 * - Recent transactions
 * - Project-wise customized inventory
 */

const DASHBOARD_SERVICE_CONFIG = Object.freeze({
  RECENT_LIMIT: 10,
  PROJECT_LIMIT: 20,
  LOW_STOCK_LIMIT: 20,
  ELIGIBLE_LIMIT: 20
});


/**
 * Returns complete dashboard data.
 */
function getDashboardData() {
  return safeExecute_(function () {
    requireAuthenticatedUser_();

    const inventorySummary =
      getInventorySummary();

    const inwardSummary =
      getInwardSummary();

    const outwardSummary =
      getOutwardSummary();

    const approvalSummary =
      getApprovalSummary();

    const conversionSummary =
      getConversionSummary();

    const todayMovement =
      getTodayInventoryMovementSummary_();

    const lowStock =
      getDashboardLowStock_(
        DASHBOARD_SERVICE_CONFIG
          .LOW_STOCK_LIMIT
      );

    const eligibleCustomize =
      getDashboardEligibleCustomize_(
        DASHBOARD_SERVICE_CONFIG
          .ELIGIBLE_LIMIT
      );

    const recentTransactions =
      getDashboardRecentTransactions_(
        DASHBOARD_SERVICE_CONFIG
          .RECENT_LIMIT
      );

    const projectWiseCustomize =
      getProjectWiseCustomizeInventory_(
        DASHBOARD_SERVICE_CONFIG
          .PROJECT_LIMIT
      );

    return successResponse_(
      'Dashboard loaded successfully.',
      {
        generatedAt:
          formatDateTime_(new Date()),

        inventory:
          inventorySummary &&
          inventorySummary.success
            ? inventorySummary.data
            : null,

        inward:
          inwardSummary &&
          inwardSummary.success
            ? inwardSummary.data
            : null,

        outward:
          outwardSummary &&
          outwardSummary.success
            ? outwardSummary.data
            : null,

        approvals:
          approvalSummary &&
          approvalSummary.success
            ? approvalSummary.data
            : null,

        conversions:
          conversionSummary &&
          conversionSummary.success
            ? conversionSummary.data
            : null,

        todayMovement:
          todayMovement,

        lowStock:
          lowStock,

        eligibleCustomize:
          eligibleCustomize,

        recentTransactions:
          recentTransactions,

        projectWiseCustomize:
          projectWiseCustomize
      }
    );
  }, 'Unable to load dashboard.');
}


/**
 * Returns today inward/outward totals.
 */
function getTodayInventoryMovementSummary_() {
  const todayKey =
    formatDashboardDateKey_(new Date());

  const inwardRecords =
    getSheetObjects_(
      getInwardSheetName_()
    ).filter(function (record) {
      return (
        normalizeText_(
          record['Inward No']
        ) &&
        formatDashboardDateKey_(
          record['Date']
        ) === todayKey
      );
    });

  const outwardRecords =
    getSheetObjects_(
      getOutwardSheetName_()
    ).filter(function (record) {
      return (
        normalizeText_(
          record['Outward No']
        ) &&
        formatDashboardDateKey_(
          record['Date']
        ) === todayKey
      );
    });

  const inwardNumbers = {};
  const outwardNumbers = {};

  let inwardQty = 0;
  let inwardAmount = 0;
  let outwardQty = 0;
  let outwardAmount = 0;

  inwardRecords.forEach(function (record) {
    const inwardNo =
      normalizeUpper_(
        record['Inward No']
      );

    if (inwardNo) {
      inwardNumbers[inwardNo] = true;
    }

    inwardQty +=
      toNumber_(
        record['Qty'],
        0
      );

    inwardAmount +=
      toNumber_(
        record['Amount'],
        0
      );
  });

  outwardRecords.forEach(function (record) {
    const outwardNo =
      normalizeUpper_(
        record['Outward No']
      );

    if (outwardNo) {
      outwardNumbers[outwardNo] = true;
    }

    outwardQty +=
      toNumber_(
        record['Qty'],
        0
      );

    outwardAmount +=
      toNumber_(
        record['Amount'],
        0
      );
  });

  return {
    date:
      todayKey,

    inward: {
      transactions:
        Object.keys(
          inwardNumbers
        ).length,

      lines:
        inwardRecords.length,

      quantity:
        roundTwo_(
          inwardQty
        ),

      amount:
        roundTwo_(
          inwardAmount
        )
    },

    outward: {
      transactions:
        Object.keys(
          outwardNumbers
        ).length,

      lines:
        outwardRecords.length,

      quantity:
        roundTwo_(
          outwardQty
        ),

      amount:
        roundTwo_(
          outwardAmount
        )
    },

    netQuantity:
      roundTwo_(
        inwardQty -
        outwardQty
      )
  };
}


/**
 * Returns dashboard low-stock records.
 */
function getDashboardLowStock_(
  limit
) {
  const records =
    getSheetObjects_(
      getFreeInventorySheetName_()
    )
      .filter(function (record) {
        if (
          !normalizeText_(
            record['Inventory ID']
          )
        ) {
          return false;
        }

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

        return (
          reorderLevel > 0 &&
          availableQty <=
            reorderLevel
        );
      })
      .map(function (record) {
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

        return {
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

          categoryName:
            normalizeText_(
              record['Category Name']
            ),

          uom:
            normalizeUpper_(
              record['UOM']
            ),

          availableQty:
            availableQty,

          minimumStock:
            inventoryToNumber_(
              record['Minimum Stock'],
              0
            ),

          reorderLevel:
            reorderLevel,

          shortageQty:
            roundTwo_(
              Math.max(
                0,
                reorderLevel -
                availableQty
              )
            ),

          status:
            normalizeUpper_(
              record['Status']
            )
        };
      })
      .sort(function (first, second) {
        return (
          second.shortageQty -
          first.shortageQty
        );
      });

  return {
    totalRecords:
      records.length,

    records:
      records.slice(
        0,
        Math.max(
          1,
          toNumber_(
            limit,
            DASHBOARD_SERVICE_CONFIG
              .LOW_STOCK_LIMIT
          )
        )
      )
  };
}


/**
 * Returns customized inventory eligible for Free conversion.
 */
function getDashboardEligibleCustomize_(
  limit
) {
  const result =
    getEligibleCustomizeInventory({
      pageNumber: 1,
      pageSize:
        Math.max(
          1,
          toNumber_(
            limit,
            DASHBOARD_SERVICE_CONFIG
              .ELIGIBLE_LIMIT
          )
        )
    });

  if (
    !result ||
    result.success !== true
  ) {
    return {
      totalRecords: 0,
      records: []
    };
  }

  return {
    totalRecords:
      result.data.pagination
        .totalRecords,

    records:
      result.data.records
  };
}


/**
 * Returns recent inward, outward, approval, and conversion activity.
 */
function getDashboardRecentTransactions_(
  limit
) {
  const transactions = [];

  getSheetObjects_(
    getInwardSheetName_()
  ).forEach(function (record) {
    const inwardNo =
      normalizeUpper_(
        record['Inward No']
      );

    if (!inwardNo) {
      return;
    }

    transactions.push({
      type:
        'INWARD',

      referenceId:
        inwardNo,

      date:
        record['Date'],

      inventoryType:
        normalizeUpper_(
          record['Inventory Type']
        ),

      project:
        normalizeText_(
          record['Project']
        ),

      skuCode:
        normalizeUpper_(
          record['SKU']
        ),

      quantity:
        toNumber_(
          record['Qty'],
          0
        ),

      amount:
        toNumber_(
          record['Amount'],
          0
        ),

      status:
        'COMPLETED',

      user:
        normalizeLower_(
          record['Entered By']
        )
    });
  });

  getSheetObjects_(
    getOutwardSheetName_()
  ).forEach(function (record) {
    const outwardNo =
      normalizeUpper_(
        record['Outward No']
      );

    if (!outwardNo) {
      return;
    }

    transactions.push({
      type:
        'OUTWARD',

      referenceId:
        outwardNo,

      date:
        record['Date'],

      inventoryType:
        normalizeUpper_(
          record['Inventory Type']
        ),

      project:
        normalizeText_(
          record[
            'Destination Project'
          ]
        ),

      skuCode:
        normalizeUpper_(
          record['SKU Code']
        ),

      quantity:
        toNumber_(
          record['Qty'],
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

      user:
        normalizeLower_(
          record['Entered By']
        )
    });
  });

  if (
    typeof getApprovalSheetName_ ===
    'function'
  ) {
    getSheetObjects_(
      getApprovalSheetName_()
    ).forEach(function (record) {
      const requestId =
        normalizeUpper_(
          record['Request ID']
        );

      if (!requestId) {
        return;
      }

      transactions.push({
        type:
          'APPROVAL',

        referenceId:
          requestId,

        date:
          record['Request Date'],

        inventoryType:
          'CUSTOMIZE',

        project:
          normalizeText_(
            record[
              'Destination Project'
            ]
          ),

        skuCode:
          normalizeUpper_(
            record['SKU Code']
          ),

        quantity:
          toNumber_(
            record['Qty'],
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

        user:
          normalizeLower_(
            record[
              'Requested By Email'
            ]
          )
      });
    });
  }

  if (
    typeof getConversionSheetName_ ===
    'function'
  ) {
    getSheetObjects_(
      getConversionSheetName_()
    ).forEach(function (record) {
      const conversionId =
        normalizeUpper_(
          record['Conversion ID']
        );

      if (!conversionId) {
        return;
      }

      transactions.push({
        type:
          'CONVERSION',

        referenceId:
          conversionId,

        date:
          record['Request Date'],

        inventoryType:
          'CUSTOMIZE TO FREE',

        project:
          normalizeText_(
            record['Project Name']
          ),

        skuCode:
          normalizeUpper_(
            record['SKU Code']
          ),

        quantity:
          toNumber_(
            record['Qty'],
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

        user:
          normalizeLower_(
            record[
              'Requested By Email'
            ]
          )
      });
    });
  }

  transactions.sort(
    function (first, second) {
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
    }
  );

  return transactions
    .slice(
      0,
      Math.max(
        1,
        toNumber_(
          limit,
          DASHBOARD_SERVICE_CONFIG
            .RECENT_LIMIT
        )
      )
    )
    .map(function (record) {
      return {
        type:
          record.type,

        referenceId:
          record.referenceId,

        date:
          formatDateTime_(
            record.date
          ),

        inventoryType:
          record.inventoryType,

        project:
          record.project,

        skuCode:
          record.skuCode,

        quantity:
          roundTwo_(
            record.quantity
          ),

        amount:
          roundTwo_(
            record.amount
          ),

        status:
          record.status,

        user:
          record.user
      };
    });
}


/**
 * Returns project-wise customized inventory totals.
 */
function getProjectWiseCustomizeInventory_(
  limit
) {
  const map = {};

  getSheetObjects_(
    getCustomizeInventorySheetName_()
  ).forEach(function (record) {
    const inventoryId =
      normalizeUpper_(
        record['Inventory ID']
      );

    if (!inventoryId) {
      return;
    }

    const projectId =
      normalizeUpper_(
        record['Project ID']
      );

    if (!projectId) {
      return;
    }

    if (!map[projectId]) {
      map[projectId] = {
        projectId:
          projectId,

        projectCode:
          normalizeUpper_(
            record['Project Code']
          ),

        projectName:
          normalizeText_(
            record['Project Name']
          ),

        skuRecords:
          0,

        availableQty:
          0,

        reservedQty:
          0,

        damagedQty:
          0,

        stockValue:
          0,

        eligibleRecords:
          0
      };
    }

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

    map[projectId].skuRecords++;

    map[projectId].availableQty +=
      availableQty;

    map[projectId].reservedQty +=
      inventoryToNumber_(
        record['Reserved Qty'],
        0
      );

    map[projectId].damagedQty +=
      inventoryToNumber_(
        record['Damaged Qty'],
        0
      );

    map[projectId].stockValue +=
      availableQty *
      averageRate;

    if (
      normalizeUpper_(
        record['Eligible for Free']
      ) === 'YES'
    ) {
      map[projectId]
        .eligibleRecords++;
    }
  });

  return Object.keys(map)
    .map(function (projectId) {
      const record = map[projectId];

      return {
        projectId:
          record.projectId,

        projectCode:
          record.projectCode,

        projectName:
          record.projectName,

        skuRecords:
          record.skuRecords,

        availableQty:
          roundTwo_(
            record.availableQty
          ),

        reservedQty:
          roundTwo_(
            record.reservedQty
          ),

        damagedQty:
          roundTwo_(
            record.damagedQty
          ),

        stockValue:
          roundTwo_(
            record.stockValue
          ),

        eligibleRecords:
          record.eligibleRecords
      };
    })
    .sort(function (first, second) {
      return (
        second.stockValue -
        first.stockValue
      );
    })
    .slice(
      0,
      Math.max(
        1,
        toNumber_(
          limit,
          DASHBOARD_SERVICE_CONFIG
            .PROJECT_LIMIT
        )
      )
    );
}


/**
 * Returns date key in yyyy-MM-dd format.
 */
function formatDashboardDateKey_(
  value
) {
  const date =
    parseDate_(value);

  if (!date) {
    return '';
  }

  return Utilities.formatDate(
    date,
    APP_CONFIG.TIME_ZONE ||
      'Asia/Kolkata',
    'yyyy-MM-dd'
  );
}


/**
 * Returns compact dashboard counters.
 */
function getDashboardCounters() {
  return safeExecute_(function () {
    requireAuthenticatedUser_();

    const inventory =
      getInventorySummary();

    const approvals =
      getApprovalSummary();

    const conversions =
      getConversionSummary();

    const today =
      getTodayInventoryMovementSummary_();

    const inventoryData =
      inventory &&
      inventory.success
        ? inventory.data
        : {};

    const approvalData =
      approvals &&
      approvals.success
        ? approvals.data
        : {};

    const conversionData =
      conversions &&
      conversions.success
        ? conversions.data
        : {};

    return successResponse_(
      'Dashboard counters loaded successfully.',
      {
        freeAvailableQty:
          inventoryData.freeInventory
            ? inventoryData.freeInventory
                .totalAvailableQty
            : 0,

        freeStockValue:
          inventoryData.freeInventory
            ? inventoryData.freeInventory
                .totalStockValue
            : 0,

        customizeAvailableQty:
          inventoryData.customizeInventory
            ? inventoryData.customizeInventory
                .totalAvailableQty
            : 0,

        customizeStockValue:
          inventoryData.customizeInventory
            ? inventoryData.customizeInventory
                .totalStockValue
            : 0,

        combinedAvailableQty:
          inventoryData.combined
            ? inventoryData.combined
                .totalAvailableQty
            : 0,

        combinedStockValue:
          inventoryData.combined
            ? inventoryData.combined
                .totalStockValue
            : 0,

        lowStockRecords:
          inventoryData.freeInventory
            ? inventoryData.freeInventory
                .lowStockRecords
            : 0,

        pendingApprovals:
          approvalData.pending || 0,

        pendingConversions:
          conversionData.pending || 0,

        todayInwardQty:
          today.inward.quantity,

        todayOutwardQty:
          today.outward.quantity
      }
    );
  }, 'Unable to load dashboard counters.');
}


/**
 * Test complete dashboard.
 */
function testDashboardService() {
  const result = {
    dashboard:
      getDashboardData(),

    counters:
      getDashboardCounters()
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
