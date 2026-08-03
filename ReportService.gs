/**
 * PROJECT INVENTORY MANAGEMENT SYSTEM
 * File: ReportService.gs
 *
 * Handles:
 * - Inventory reports
 * - Free inventory report
 * - Customized inventory report
 * - Inward report
 * - Outward report
 * - Approval report
 * - Conversion report
 * - Project-wise customized stock report
 * - Low-stock report
 * - Eligible-for-free report
 * - Date filtering
 * - CSV export
 */

const REPORT_SERVICE_CONFIG = Object.freeze({
  DEFAULT_PAGE_SIZE: 50,
  MAX_PAGE_SIZE: 500,
  EXPORT_MAX_ROWS: 5000
});


/**
 * Returns complete report dashboard data.
 */
function getReportDashboard(filters) {
  return safeExecute_(function () {
    requireAuthenticatedUser_();

    filters = filters || {};

    return successResponse_(
      'Report dashboard loaded successfully.',
      {
        generatedAt:
          formatDateTime_(new Date()),

        inventorySummary:
          getInventorySummary(),

        inwardSummary:
          getInwardReportSummary_(
            filters
          ),

        outwardSummary:
          getOutwardReportSummary_(
            filters
          ),

        approvalSummary:
          getApprovalReportSummary_(
            filters
          ),

        conversionSummary:
          getConversionReportSummary_(
            filters
          ),

        lowStock:
          getLowStockReport({
            pageNumber: 1,
            pageSize: 20
          }),

        eligibleCustomize:
          getEligibleForFreeReport({
            pageNumber: 1,
            pageSize: 20
          }),

        projectWiseCustomize:
          getProjectWiseCustomizeReport({
            pageNumber: 1,
            pageSize: 20
          })
      }
    );
  }, 'Unable to load report dashboard.');
}


/**
 * Returns Free Inventory report.
 *
 * filters:
 * {
 *   search: '',
 *   status: '',
 *   lowStockOnly: false,
 *   pageNumber: 1,
 *   pageSize: 50
 * }
 */
function getFreeInventoryReport(filters) {
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

    const lowStockOnly =
      toBoolean_(
        filters.lowStockOnly
      );

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

          if (
            status &&
            normalizeUpper_(
              record['Status']
            ) !== status
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

          if (
            lowStockOnly &&
            !(
              reorderLevel > 0 &&
              availableQty <=
                reorderLevel
            )
          ) {
            return false;
          }

          if (search) {
            const searchable = [
              record['Inventory ID'],
              record['SKU ID'],
              record['SKU Code'],
              record['SKU Name'],
              record['Category Name'],
              record['UOM'],
              record['Status'],
              record['Remarks']
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
        .map(mapFreeInventoryReportRow_);

    return successResponse_(
      'Free inventory report loaded successfully.',
      {
        summary:
          summarizeInventoryRows_(
            records
          ),

        records:
          paginateRecords_(
            records,
            getReportPageNumber_(
              filters
            ),
            getReportPageSize_(
              filters
            )
          )
      }
    );
  }, 'Unable to load Free Inventory report.');
}


/**
 * Maps Free Inventory row.
 */
function mapFreeInventoryReportRow_(
  record
) {
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

    reservedQty:
      inventoryToNumber_(
        record['Reserved Qty'],
        0
      ),

    damagedQty:
      inventoryToNumber_(
        record['Damaged Qty'],
        0
      ),

    minimumStock:
      inventoryToNumber_(
        record['Minimum Stock'],
        0
      ),

    reorderLevel:
      reorderLevel,

    averageRate:
      averageRate,

    stockValue:
      roundTwo_(
        availableQty *
        averageRate
      ),

    isLowStock:
      reorderLevel > 0 &&
      availableQty <=
        reorderLevel,

    lastInwardDate:
      formatDateTime_(
        record['Last Inward Date']
      ),

    lastOutwardDate:
      formatDateTime_(
        record['Last Outward Date']
      ),

    lastMovementDate:
      formatDateTime_(
        record['Last Movement Date']
      ),

    status:
      normalizeUpper_(
        record['Status']
      ),

    remarks:
      normalizeText_(
        record['Remarks']
      )
  };
}


/**
 * Returns Customized Inventory report.
 *
 * filters:
 * {
 *   search: '',
 *   projectId: '',
 *   status: '',
 *   eligibleOnly: false,
 *   pageNumber: 1,
 *   pageSize: 50
 * }
 */
function getCustomizeInventoryReport(filters) {
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

    const status =
      normalizeUpper_(
        filters.status
      );

    const eligibleOnly =
      toBoolean_(
        filters.eligibleOnly
      );

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
            projectId &&
            normalizeUpper_(
              record['Project ID']
            ) !== projectId
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
            eligibleOnly &&
            normalizeUpper_(
              record['Eligible for Free']
            ) !== 'YES'
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
              record['Category Name'],
              record['UOM'],
              record['Status'],
              record['Remarks']
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
        .map(mapCustomizeInventoryReportRow_);

    return successResponse_(
      'Customized inventory report loaded successfully.',
      {
        summary:
          summarizeInventoryRows_(
            records
          ),

        records:
          paginateRecords_(
            records,
            getReportPageNumber_(
              filters
            ),
            getReportPageSize_(
              filters
            )
          )
      }
    );
  }, 'Unable to load Customized Inventory report.');
}


/**
 * Maps Customized Inventory row.
 */
function mapCustomizeInventoryReportRow_(
  record
) {
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
      availableQty,

    reservedQty:
      inventoryToNumber_(
        record['Reserved Qty'],
        0
      ),

    damagedQty:
      inventoryToNumber_(
        record['Damaged Qty'],
        0
      ),

    averageRate:
      averageRate,

    stockValue:
      roundTwo_(
        availableQty *
        averageRate
      ),

    lastInwardDate:
      formatDateTime_(
        record['Last Inward Date']
      ),

    lastOutwardDate:
      formatDateTime_(
        record['Last Outward Date']
      ),

    lastMovementDate:
      formatDateTime_(
        record['Last Movement Date']
      ),

    eligibleForFree:
      normalizeUpper_(
        record['Eligible for Free']
      ),

    eligibilityDate:
      formatDateTime_(
        eligibilityDate
      ),

    daysRemaining:
      eligibilityDate
        ? Math.max(
            0,
            daysBetween_(
              new Date(),
              eligibilityDate
            )
          )
        : 0,

    status:
      normalizeUpper_(
        record['Status']
      ),

    remarks:
      normalizeText_(
        record['Remarks']
      )
  };
}


/**
 * Returns Inward report.
 */
function getInwardReport(filters) {
  return safeExecute_(function () {
    requireAuthenticatedUser_();

    filters = filters || {};

    const records =
      filterInwardRowsForReport_(
        filters
      ).map(function (record) {
        return {
          inwardNo:
            normalizeUpper_(
              record['Inward No']
            ),

          date:
            formatDateTime_(
              record['Date']
            ),

          inventoryType:
            normalizeUpper_(
              record['Inventory Type']
            ),

          vendor:
            normalizeText_(
              record['Vendor']
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

          enteredBy:
            normalizeLower_(
              record['Entered By']
            )
        };
      });

    return successResponse_(
      'Inward report loaded successfully.',
      {
        summary:
          summarizeMovementRows_(
            records
          ),

        records:
          paginateRecords_(
            records,
            getReportPageNumber_(
              filters
            ),
            getReportPageSize_(
              filters
            )
          )
      }
    );
  }, 'Unable to load Inward report.');
}


/**
 * Returns Outward report.
 */
function getOutwardReport(filters) {
  return safeExecute_(function () {
    requireAuthenticatedUser_();

    filters = filters || {};

    const records =
      filterOutwardRowsForReport_(
        filters
      ).map(function (record) {
        return {
          outwardNo:
            normalizeUpper_(
              record['Outward No']
            ),

          date:
            formatDateTime_(
              record['Date']
            ),

          inventoryType:
            normalizeUpper_(
              record['Inventory Type']
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

          issuedTo:
            normalizeText_(
              record['Issued To']
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

          remarks:
            normalizeText_(
              record['Remarks']
            ),

          enteredBy:
            normalizeLower_(
              record['Entered By']
            )
        };
      });

    return successResponse_(
      'Outward report loaded successfully.',
      {
        summary:
          summarizeMovementRows_(
            records
          ),

        records:
          paginateRecords_(
            records,
            getReportPageNumber_(
              filters
            ),
            getReportPageSize_(
              filters
            )
          )
      }
    );
  }, 'Unable to load Outward report.');
}


/**
 * Returns Approval report.
 */
function getApprovalReport(filters) {
  return safeExecute_(function () {
    requireAuthenticatedUser_();

    filters = filters || {};

    const result =
      getApprovalRequests({
        search:
          filters.search || '',
        status:
          filters.status || '',
        sourceProjectId:
          filters.sourceProjectId || '',
        destinationProjectId:
          filters.destinationProjectId || '',
        pageNumber:
          getReportPageNumber_(
            filters
          ),
        pageSize:
          getReportPageSize_(
            filters
          )
      });

    return result;
  }, 'Unable to load Approval report.');
}


/**
 * Returns Conversion report.
 */
function getConversionReport(filters) {
  return safeExecute_(function () {
    requireAuthenticatedUser_();

    filters = filters || {};

    return getConversionRequests({
      search:
        filters.search || '',
      status:
        filters.status || '',
      projectId:
        filters.projectId || '',
      pageNumber:
        getReportPageNumber_(
          filters
        ),
      pageSize:
        getReportPageSize_(
          filters
        )
    });
  }, 'Unable to load Conversion report.');
}


/**
 * Returns low-stock report.
 */
function getLowStockReport(filters) {
  filters = filters || {};
  filters.lowStockOnly = true;

  return getFreeInventoryReport(
    filters
  );
}


/**
 * Returns eligible-for-free report.
 */
function getEligibleForFreeReport(filters) {
  filters = filters || {};
  filters.eligibleOnly = true;

  return getCustomizeInventoryReport(
    filters
  );
}


/**
 * Returns project-wise customized inventory report.
 */
function getProjectWiseCustomizeReport(
  filters
) {
  return safeExecute_(function () {
    requireAuthenticatedUser_();

    filters = filters || {};

    const search =
      normalizeLower_(
        filters.search
      );

    const records =
      getProjectWiseCustomizeInventory_(
        REPORT_SERVICE_CONFIG
          .EXPORT_MAX_ROWS
      ).filter(function (record) {
        if (!search) {
          return true;
        }

        const searchable = [
          record.projectId,
          record.projectCode,
          record.projectName
        ]
          .map(normalizeLower_)
          .join(' ');

        return (
          searchable.indexOf(
            search
          ) !== -1
        );
      });

    return successResponse_(
      'Project-wise customized inventory report loaded successfully.',
      {
        summary: {
          totalProjects:
            records.length,

          totalSkuRecords:
            roundTwo_(
              records.reduce(
                function (
                  total,
                  record
                ) {
                  return total +
                    record.skuRecords;
                },
                0
              )
            ),

          totalAvailableQty:
            roundTwo_(
              records.reduce(
                function (
                  total,
                  record
                ) {
                  return total +
                    record.availableQty;
                },
                0
              )
            ),

          totalStockValue:
            roundTwo_(
              records.reduce(
                function (
                  total,
                  record
                ) {
                  return total +
                    record.stockValue;
                },
                0
              )
            ),

          totalEligibleRecords:
            roundTwo_(
              records.reduce(
                function (
                  total,
                  record
                ) {
                  return total +
                    record.eligibleRecords;
                },
                0
              )
            )
        },

        records:
          paginateRecords_(
            records,
            getReportPageNumber_(
              filters
            ),
            getReportPageSize_(
              filters
            )
          )
      }
    );
  }, 'Unable to load project-wise report.');
}


/**
 * Filters Inward rows.
 */
function filterInwardRowsForReport_(
  filters
) {
  const search =
    normalizeLower_(
      filters.search
    );

  const inventoryType =
    normalizeUpper_(
      filters.inventoryType
    );

  const projectName =
    normalizeLower_(
      filters.projectName
    );

  const dateFrom =
    parseDate_(
      filters.dateFrom
    );

  const dateTo =
    parseDate_(
      filters.dateTo
    );

  return getSheetObjects_(
    getInwardSheetName_()
  ).filter(function (record) {
    if (
      !normalizeText_(
        record['Inward No']
      )
    ) {
      return false;
    }

    if (
      inventoryType &&
      normalizeUpper_(
        record['Inventory Type']
      ) !== inventoryType
    ) {
      return false;
    }

    if (
      projectName &&
      normalizeLower_(
        record['Project']
      ) !== projectName
    ) {
      return false;
    }

    if (
      !isDateWithinReportRange_(
        record['Date'],
        dateFrom,
        dateTo
      )
    ) {
      return false;
    }

    if (search) {
      const searchable = [
        record['Inward No'],
        record['Inventory Type'],
        record['Vendor'],
        record['Project'],
        record['SKU'],
        record['Entered By']
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
}


/**
 * Filters Outward rows.
 */
function filterOutwardRowsForReport_(
  filters
) {
  const search =
    normalizeLower_(
      filters.search
    );

  const inventoryType =
    normalizeUpper_(
      filters.inventoryType
    );

  const sourceProjectId =
    normalizeUpper_(
      filters.sourceProjectId
    );

  const destinationProjectId =
    normalizeUpper_(
      filters.destinationProjectId
    );

  const status =
    normalizeUpper_(
      filters.status
    );

  const dateFrom =
    parseDate_(
      filters.dateFrom
    );

  const dateTo =
    parseDate_(
      filters.dateTo
    );

  return getSheetObjects_(
    getOutwardSheetName_()
  ).filter(function (record) {
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
        record['Inventory Type']
      ) !== inventoryType
    ) {
      return false;
    }

    if (
      sourceProjectId &&
      normalizeUpper_(
        record['Source Project ID']
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

    if (
      status &&
      normalizeUpper_(
        record['Status']
      ) !== status
    ) {
      return false;
    }

    if (
      !isDateWithinReportRange_(
        record['Date'],
        dateFrom,
        dateTo
      )
    ) {
      return false;
    }

    if (search) {
      const searchable = [
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
        searchable.indexOf(
          search
        ) === -1
      ) {
        return false;
      }
    }

    return true;
  });
}


/**
 * Date range helper.
 */
function isDateWithinReportRange_(
  value,
  dateFrom,
  dateTo
) {
  const date =
    parseDate_(value);

  if (!date) {
    return false;
  }

  if (
    dateFrom &&
    date.getTime() <
      dateFrom.getTime()
  ) {
    return false;
  }

  if (dateTo) {
    const endDate =
      new Date(
        dateTo.getTime()
      );

    endDate.setHours(
      23,
      59,
      59,
      999
    );

    if (
      date.getTime() >
      endDate.getTime()
    ) {
      return false;
    }
  }

  return true;
}


/**
 * Inventory report summary.
 */
function summarizeInventoryRows_(
  records
) {
  return {
    totalRecords:
      records.length,

    totalAvailableQty:
      roundTwo_(
        records.reduce(
          function (total, record) {
            return total +
              toNumber_(
                record.availableQty,
                0
              );
          },
          0
        )
      ),

    totalReservedQty:
      roundTwo_(
        records.reduce(
          function (total, record) {
            return total +
              toNumber_(
                record.reservedQty,
                0
              );
          },
          0
        )
      ),

    totalDamagedQty:
      roundTwo_(
        records.reduce(
          function (total, record) {
            return total +
              toNumber_(
                record.damagedQty,
                0
              );
          },
          0
        )
      ),

    totalStockValue:
      roundTwo_(
        records.reduce(
          function (total, record) {
            return total +
              toNumber_(
                record.stockValue,
                0
              );
          },
          0
        )
      )
  };
}


/**
 * Movement report summary.
 */
function summarizeMovementRows_(
  records
) {
  const references = {};

  records.forEach(function (record) {
    const reference =
      normalizeUpper_(
        record.inwardNo ||
        record.outwardNo
      );

    if (reference) {
      references[reference] = true;
    }
  });

  return {
    totalTransactions:
      Object.keys(
        references
      ).length,

    totalLines:
      records.length,

    totalQuantity:
      roundTwo_(
        records.reduce(
          function (total, record) {
            return total +
              toNumber_(
                record.quantity,
                0
              );
          },
          0
        )
      ),

    totalAmount:
      roundTwo_(
        records.reduce(
          function (total, record) {
            return total +
              toNumber_(
                record.amount,
                0
              );
          },
          0
        )
      )
  };
}


/**
 * Inward summary for dashboard reports.
 */
function getInwardReportSummary_(
  filters
) {
  const rows =
    filterInwardRowsForReport_(
      filters || {}
    );

  return summarizeMovementRows_(
    rows.map(function (record) {
      return {
        inwardNo:
          record['Inward No'],
        quantity:
          record['Qty'],
        amount:
          record['Amount']
      };
    })
  );
}


/**
 * Outward summary for dashboard reports.
 */
function getOutwardReportSummary_(
  filters
) {
  const rows =
    filterOutwardRowsForReport_(
      filters || {}
    );

  return summarizeMovementRows_(
    rows.map(function (record) {
      return {
        outwardNo:
          record['Outward No'],
        quantity:
          record['Qty'],
        amount:
          record['Amount']
      };
    })
  );
}


/**
 * Approval summary for report dashboard.
 */
function getApprovalReportSummary_(
  filters
) {
  const result =
    getApprovalRequests({
      search:
        filters.search || '',
      status:
        filters.approvalStatus || '',
      sourceProjectId:
        filters.sourceProjectId || '',
      destinationProjectId:
        filters.destinationProjectId || '',
      pageNumber: 1,
      pageSize:
        REPORT_SERVICE_CONFIG
          .EXPORT_MAX_ROWS
    });

  if (
    !result ||
    result.success !== true
  ) {
    return null;
  }

  return {
    totalRequests:
      result.data.pagination
        .totalRecords,

    totalQuantity:
      roundTwo_(
        result.data.records.reduce(
          function (total, record) {
            return total +
              toNumber_(
                record.totalQuantity,
                0
              );
          },
          0
        )
      ),

    totalAmount:
      roundTwo_(
        result.data.records.reduce(
          function (total, record) {
            return total +
              toNumber_(
                record.totalAmount,
                0
              );
          },
          0
        )
      )
  };
}


/**
 * Conversion summary for report dashboard.
 */
function getConversionReportSummary_(
  filters
) {
  const result =
    getConversionRequests({
      search:
        filters.search || '',
      status:
        filters.conversionStatus || '',
      projectId:
        filters.projectId || '',
      pageNumber: 1,
      pageSize:
        REPORT_SERVICE_CONFIG
          .EXPORT_MAX_ROWS
    });

  if (
    !result ||
    result.success !== true
  ) {
    return null;
  }

  return {
    totalRequests:
      result.data.pagination
        .totalRecords,

    totalQuantity:
      roundTwo_(
        result.data.records.reduce(
          function (total, record) {
            return total +
              toNumber_(
                record.totalQuantity,
                0
              );
          },
          0
        )
      ),

    totalAmount:
      roundTwo_(
        result.data.records.reduce(
          function (total, record) {
            return total +
              toNumber_(
                record.totalAmount,
                0
              );
          },
          0
        )
      )
  };
}


/**
 * Returns safe page number.
 */
function getReportPageNumber_(
  filters
) {
  return Math.max(
    1,
    toNumber_(
      filters.pageNumber,
      1
    )
  );
}


/**
 * Returns safe page size.
 */
function getReportPageSize_(
  filters
) {
  return Math.max(
    1,
    Math.min(
      toNumber_(
        filters.pageSize,
        REPORT_SERVICE_CONFIG
          .DEFAULT_PAGE_SIZE
      ),
      REPORT_SERVICE_CONFIG
        .MAX_PAGE_SIZE
    )
  );
}


/**
 * Exports one report as CSV text.
 *
 * reportType:
 * FREE_INVENTORY
 * CUSTOMIZE_INVENTORY
 * INWARD
 * OUTWARD
 * APPROVAL
 * CONVERSION
 * LOW_STOCK
 * ELIGIBLE_FOR_FREE
 * PROJECT_WISE_CUSTOMIZE
 */
function exportReportCsv(
  reportType,
  filters
) {
  return safeExecute_(function () {
    requirePermission_(
      'canReports'
    );

    const type =
      normalizeUpper_(
        reportType
      );

    filters = filters || {};
    filters.pageNumber = 1;
    filters.pageSize =
      REPORT_SERVICE_CONFIG
        .EXPORT_MAX_ROWS;

    let rows = [];
    let headers = [];
    let fileName = '';

    if (
      type === 'FREE_INVENTORY'
    ) {
      const result =
        getFreeInventoryReport(
          filters
        );

      rows =
        result.data.records.records;

      headers = [
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
        'Stock Value',
        'Low Stock',
        'Last Inward Date',
        'Last Outward Date',
        'Last Movement Date',
        'Status',
        'Remarks'
      ];

      rows = rows.map(function (record) {
        return [
          record.inventoryId,
          record.skuId,
          record.skuCode,
          record.skuName,
          record.categoryName,
          record.uom,
          record.availableQty,
          record.reservedQty,
          record.damagedQty,
          record.minimumStock,
          record.reorderLevel,
          record.averageRate,
          record.stockValue,
          record.isLowStock
            ? 'YES'
            : 'NO',
          record.lastInwardDate,
          record.lastOutwardDate,
          record.lastMovementDate,
          record.status,
          record.remarks
        ];
      });

      fileName =
        'Free_Inventory_Report';
    } else if (
      type ===
      'CUSTOMIZE_INVENTORY'
    ) {
      const result =
        getCustomizeInventoryReport(
          filters
        );

      rows =
        result.data.records.records;

      headers = [
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
        'Stock Value',
        'Last Inward Date',
        'Last Outward Date',
        'Last Movement Date',
        'Eligible for Free',
        'Eligibility Date',
        'Days Remaining',
        'Status',
        'Remarks'
      ];

      rows = rows.map(function (record) {
        return [
          record.inventoryId,
          record.projectId,
          record.projectCode,
          record.projectName,
          record.skuId,
          record.skuCode,
          record.skuName,
          record.categoryName,
          record.uom,
          record.availableQty,
          record.reservedQty,
          record.damagedQty,
          record.averageRate,
          record.stockValue,
          record.lastInwardDate,
          record.lastOutwardDate,
          record.lastMovementDate,
          record.eligibleForFree,
          record.eligibilityDate,
          record.daysRemaining,
          record.status,
          record.remarks
        ];
      });

      fileName =
        'Customize_Inventory_Report';
    } else if (
      type === 'INWARD'
    ) {
      const result =
        getInwardReport(
          filters
        );

      rows =
        result.data.records.records;

      headers = [
        'Inward No',
        'Date',
        'Inventory Type',
        'Vendor',
        'Project',
        'SKU Code',
        'Qty',
        'Unit',
        'Rate',
        'Amount',
        'Entered By'
      ];

      rows = rows.map(function (record) {
        return [
          record.inwardNo,
          record.date,
          record.inventoryType,
          record.vendor,
          record.project,
          record.skuCode,
          record.quantity,
          record.unit,
          record.rate,
          record.amount,
          record.enteredBy
        ];
      });

      fileName =
        'Inward_Report';
    } else if (
      type === 'OUTWARD'
    ) {
      const result =
        getOutwardReport(
          filters
        );

      rows =
        result.data.records.records;

      headers = [
        'Outward No',
        'Date',
        'Inventory Type',
        'Source Project ID',
        'Source Project',
        'Destination Project ID',
        'Destination Project',
        'Reference No',
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

      rows = rows.map(function (record) {
        return [
          record.outwardNo,
          record.date,
          record.inventoryType,
          record.sourceProjectId,
          record.sourceProject,
          record.destinationProjectId,
          record.destinationProject,
          record.referenceNo,
          record.issuedTo,
          record.skuId,
          record.skuCode,
          record.skuName,
          record.quantity,
          record.unit,
          record.rate,
          record.amount,
          record.status,
          record.remarks,
          record.enteredBy
        ];
      });

      fileName =
        'Outward_Report';
    } else if (
      type === 'APPROVAL'
    ) {
      const result =
        getApprovalReport(
          filters
        );

      rows =
        result.data.records;

      headers = [
        'Request ID',
        'Request Date',
        'Source Project',
        'Destination Project',
        'Reference No',
        'Requested By',
        'Approver Email',
        'Status',
        'Item Count',
        'Total Quantity',
        'Total Amount',
        'Approved By',
        'Approved At',
        'Rejected By',
        'Rejected At',
        'Executed By',
        'Executed At',
        'Outward No',
        'Remarks'
      ];

      rows = rows.map(function (record) {
        return [
          record.requestId,
          record.requestDate,
          record.sourceProject,
          record.destinationProject,
          record.referenceNo,
          record.requestedBy,
          record.approverEmail,
          record.status,
          record.itemCount,
          record.totalQuantity,
          record.totalAmount,
          record.approvedBy,
          record.approvedAt,
          record.rejectedBy,
          record.rejectedAt,
          record.executedBy,
          record.executedAt,
          record.outwardNo,
          record.remarks
        ];
      });

      fileName =
        'Approval_Report';
    } else if (
      type === 'CONVERSION'
    ) {
      const result =
        getConversionReport(
          filters
        );

      rows =
        result.data.records;

      headers = [
        'Conversion ID',
        'Request Date',
        'Project ID',
        'Project Name',
        'Requested By',
        'Approver Email',
        'Status',
        'Item Count',
        'Total Quantity',
        'Total Amount',
        'Approved By',
        'Approved At',
        'Rejected By',
        'Rejected At',
        'Executed By',
        'Executed At',
        'Reference No',
        'Remarks'
      ];

      rows = rows.map(function (record) {
        return [
          record.conversionId,
          record.requestDate,
          record.projectId,
          record.projectName,
          record.requestedBy,
          record.approverEmail,
          record.status,
          record.itemCount,
          record.totalQuantity,
          record.totalAmount,
          record.approvedBy,
          record.approvedAt,
          record.rejectedBy,
          record.rejectedAt,
          record.executedBy,
          record.executedAt,
          record.referenceNo,
          record.remarks
        ];
      });

      fileName =
        'Conversion_Report';
    } else if (
      type === 'LOW_STOCK'
    ) {
      return exportReportCsv(
        'FREE_INVENTORY',
        Object.assign(
          {},
          filters,
          {
            lowStockOnly: true
          }
        )
      );
    } else if (
      type ===
      'ELIGIBLE_FOR_FREE'
    ) {
      return exportReportCsv(
        'CUSTOMIZE_INVENTORY',
        Object.assign(
          {},
          filters,
          {
            eligibleOnly: true
          }
        )
      );
    } else if (
      type ===
      'PROJECT_WISE_CUSTOMIZE'
    ) {
      const result =
        getProjectWiseCustomizeReport(
          filters
        );

      rows =
        result.data.records.records;

      headers = [
        'Project ID',
        'Project Code',
        'Project Name',
        'SKU Records',
        'Available Qty',
        'Reserved Qty',
        'Damaged Qty',
        'Stock Value',
        'Eligible Records'
      ];

      rows = rows.map(function (record) {
        return [
          record.projectId,
          record.projectCode,
          record.projectName,
          record.skuRecords,
          record.availableQty,
          record.reservedQty,
          record.damagedQty,
          record.stockValue,
          record.eligibleRecords
        ];
      });

      fileName =
        'Project_Wise_Customize_Report';
    } else {
      throw new Error(
        'Invalid report type: ' +
        type
      );
    }

    const csv =
      buildCsvText_(
        headers,
        rows
      );

    return successResponse_(
      'Report CSV generated successfully.',
      {
        reportType:
          type,

        fileName:
          fileName +
          '_' +
          Utilities.formatDate(
            new Date(),
            APP_CONFIG.TIME_ZONE ||
              'Asia/Kolkata',
            'yyyyMMdd_HHmmss'
          ) +
          '.csv',

        mimeType:
          'text/csv',

        totalRows:
          rows.length,

        csv:
          csv
      }
    );
  }, 'Unable to export report CSV.');
}


/**
 * Builds CSV text.
 */
function buildCsvText_(
  headers,
  rows
) {
  const output = [
    headers.map(
      escapeCsvValue_
    ).join(',')
  ];

  rows.forEach(function (row) {
    output.push(
      row.map(
        escapeCsvValue_
      ).join(',')
    );
  });

  return output.join('\n');
}


/**
 * Escapes one CSV value.
 */
function escapeCsvValue_(value) {
  const text =
    normalizeText_(value);

  if (
    /[",\n\r]/.test(text)
  ) {
    return (
      '"' +
      text.replace(
        /"/g,
        '""'
      ) +
      '"'
    );
  }

  return text;
}


/**
 * Test all reports.
 */
function testReportService() {
  const result = {
    dashboard:
      getReportDashboard({}),

    freeInventory:
      getFreeInventoryReport({
        pageNumber: 1,
        pageSize: 10
      }),

    customizeInventory:
      getCustomizeInventoryReport({
        pageNumber: 1,
        pageSize: 10
      }),

    inward:
      getInwardReport({
        pageNumber: 1,
        pageSize: 10
      }),

    outward:
      getOutwardReport({
        pageNumber: 1,
        pageSize: 10
      }),

    approval:
      getApprovalReport({
        pageNumber: 1,
        pageSize: 10
      }),

    conversion:
      getConversionReport({
        pageNumber: 1,
        pageSize: 10
      }),

    lowStock:
      getLowStockReport({
        pageNumber: 1,
        pageSize: 10
      }),

    eligibleForFree:
      getEligibleForFreeReport({
        pageNumber: 1,
        pageSize: 10
      }),

    projectWise:
      getProjectWiseCustomizeReport({
        pageNumber: 1,
        pageSize: 10
      })
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


/**
 * Test CSV export.
 */
function testReportCsvExport() {
  const result =
    exportReportCsv(
      'OUTWARD',
      {
        pageNumber: 1,
        pageSize: 100
      }
    );

  Logger.log(
    JSON.stringify(
      {
        success:
          result.success,

        message:
          result.message,

        fileName:
          result.success
            ? result.data.fileName
            : '',

        totalRows:
          result.success
            ? result.data.totalRows
            : 0,

        csvPreview:
          result.success
            ? result.data.csv
                .substring(0, 500)
            : ''
      },
      null,
      2
    )
  );

  return result;
}
