/**
 * PROJECT INVENTORY MANAGEMENT SYSTEM
 * File: Code.gs
 *
 * Main web app entry point and frontend API layer.
 *
 * Responsibilities:
 * - Web app routing
 * - Email approval routing
 * - Frontend bootstrap data
 * - Project and SKU dropdown data
 * - Dashboard APIs
 * - Inventory APIs
 * - Inward APIs
 * - Outward APIs
 * - Approval APIs
 * - Conversion APIs
 * - Report APIs
 * - JSON doPost endpoint
 *
 * IMPORTANT:
 * Keep only ONE doGet(e) function in the complete Apps Script project.
 */

const MAIN_API_CONFIG = Object.freeze({
  APP_NAME:
    'Project Inventory Management System',

  VERSION:
    '1.0.0',

  HOME_FILE:
    'Index',

  DEFAULT_PAGE_SIZE:
    20,

  MAX_PAGE_SIZE:
    500
});


/**
 * Main web application route.
 *
 * Routes:
 * ?page=approval&action=approve&token=...
 * ?page=approval&action=reject&token=...
 * Default -> Index.html
 */
function doGet(e) {
  try {
    const parameters =
      e && e.parameter
        ? e.parameter
        : {};

    const page =
      normalizeLower_(
        parameters.page
      );

    // Email approval and rejection page.
    if (page === 'approval') {
      return handleApprovalEmailPage(e);
    }

    // Optional public health endpoint.
    if (page === 'health') {
      return ContentService
        .createTextOutput(
          JSON.stringify(
            getPublicAppHealth_(),
            null,
            2
          )
        )
        .setMimeType(
          ContentService.MimeType.JSON
        );
    }

    return serveMainApplication_(e);

  } catch (error) {
    return HtmlService
      .createHtmlOutput(
        buildMainErrorPage_(
          error &&
          error.message
            ? error.message
            : 'Unable to open the application.'
        )
      )
      .setTitle(
        MAIN_API_CONFIG.APP_NAME
      );
  }
}


/**
 * Optional JSON API endpoint.
 *
 * Example POST body:
 * {
 *   "action": "GET_DASHBOARD",
 *   "payload": {}
 * }
 */
function doPost(e) {
  try {
    let request = {};

    if (
      e &&
      e.postData &&
      e.postData.contents
    ) {
      request =
        JSON.parse(
          e.postData.contents
        );
    }

    const action =
      normalizeUpper_(
        request.action
      );

    const payload =
      request.payload || {};

    const result =
      apiDispatch(
        action,
        payload
      );

    return ContentService
      .createTextOutput(
        JSON.stringify(result)
      )
      .setMimeType(
        ContentService.MimeType.JSON
      );

  } catch (error) {
    return ContentService
      .createTextOutput(
        JSON.stringify(
          errorResponse_(
            error,
            'API request failed.'
          )
        )
      )
      .setMimeType(
        ContentService.MimeType.JSON
      );
  }
}


/**
 * Serves Index.html.
 *
 * Until Index.html is created, this returns
 * a basic backend-ready page instead of failing.
 */
function serveMainApplication_(e) {
  const template =
    tryCreateHtmlTemplate_(
      MAIN_API_CONFIG.HOME_FILE
    );

  if (!template) {
    return HtmlService
      .createHtmlOutput(
        buildBackendReadyPage_()
      )
      .setTitle(
        MAIN_API_CONFIG.APP_NAME
      )
      .setXFrameOptionsMode(
        HtmlService
          .XFrameOptionsMode
          .ALLOWALL
      );
  }

  template.APP_NAME =
    MAIN_API_CONFIG.APP_NAME;

  template.APP_VERSION =
    MAIN_API_CONFIG.VERSION;

  template.WEB_APP_URL =
    getMainWebAppUrl_();

  template.initialPage =
    normalizeLower_(
      e &&
      e.parameter
        ? e.parameter.view
        : 'dashboard'
    ) || 'dashboard';

  return template
    .evaluate()
    .setTitle(
      MAIN_API_CONFIG.APP_NAME
    )
    .setXFrameOptionsMode(
      HtmlService
        .XFrameOptionsMode
        .ALLOWALL
    )
    .addMetaTag(
      'viewport',
      'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no'
    );
}


/**
 * Safely creates an HTML template.
 */
function tryCreateHtmlTemplate_(
  fileName
) {
  try {
    return HtmlService
      .createTemplateFromFile(
        fileName
      );
  } catch (error) {
    return null;
  }
}


/**
 * Includes HTML partials.
 *
 * Usage in HTML:
 * <?!= include('Styles'); ?>
 */
function include(fileName) {
  return HtmlService
    .createHtmlOutputFromFile(
      fileName
    )
    .getContent();
}


/**
 * Returns deployed web app URL.
 */
function getMainWebAppUrl_() {
  try {
    if (
      typeof getApprovalWebAppUrl_ ===
      'function'
    ) {
      return getApprovalWebAppUrl_();
    }
  } catch (error) {
    // Continue with fallback.
  }

  try {
    return normalizeText_(
      ScriptApp
        .getService()
        .getUrl()
    );
  } catch (error) {
    return '';
  }
}


/**
 * Frontend bootstrap API.
 *
 * Loads:
 * - Current user
 * - Permissions
 * - Dashboard counters
 * - Projects
 * - SKUs
 * - Application configuration
 */
function apiGetBootstrapData() {
  return safeExecute_(function () {
    const session =
      requireAuthenticatedUser_();

    const counters =
      getDashboardCounters();

    return successResponse_(
      'Application initialized successfully.',
      {
        app: {
          name:
            MAIN_API_CONFIG.APP_NAME,

          version:
            MAIN_API_CONFIG.VERSION,

          timeZone:
            typeof APP_CONFIG !==
              'undefined' &&
            APP_CONFIG.TIME_ZONE
              ? APP_CONFIG.TIME_ZONE
              : 'Asia/Kolkata',

          currentDateTime:
            formatDateTime_(
              new Date()
            ),

          webAppUrl:
            getMainWebAppUrl_(),

          freeEligibilityDays:
            getFreeEligibilityDays()
        },

        user:
          normalizeApiSession_(
            session
          ),

        permissions:
          session.permissions || {},

        dashboardCounters:
          counters &&
          counters.success
            ? counters.data
            : {},

        projects:
          getApiProjectDropdown_(),

        skus:
          getApiSkuDropdown_()
      }
    );
  }, 'Unable to initialize application.');
}


/**
 * Current user API.
 */
function apiGetCurrentUser() {
  return safeExecute_(function () {
    const session =
      requireAuthenticatedUser_();

    return successResponse_(
      'Current user loaded successfully.',
      normalizeApiSession_(
        session
      )
    );
  }, 'Unable to load current user.');
}


/**
 * Normalizes authenticated session.
 */
function normalizeApiSession_(
  session
) {
  session = session || {};

  return {
    userId:
      normalizeUpper_(
        session.userId
      ),

    employeeName:
      normalizeText_(
        session.employeeName
      ),

    email:
      normalizeLower_(
        session.email
      ),

    mobile:
      normalizeText_(
        session.mobile
      ),

    department:
      normalizeUpper_(
        session.department
      ),

    role:
      normalizeUpper_(
        session.role
      ),

    status:
      normalizeUpper_(
        session.status
      ),

    permissions:
      session.permissions || {}
  };
}


/**
 * Dashboard APIs.
 */
function apiGetDashboard() {
  return getFastInventoryDashboard_();
}


function apiGetDashboardCounters() {
  return getDashboardCounters();
}


/**
 * Dropdown APIs.
 */
function apiGetProjectDropdown() {
  return safeExecute_(function () {
    requireAuthenticatedUser_();

    return successResponse_(
      'Project dropdown loaded successfully.',
      getApiProjectDropdown_()
    );
  }, 'Unable to load project dropdown.');
}


function apiGetSkuDropdown() {
  return safeExecute_(function () {
    requireAuthenticatedUser_();

    return successResponse_(
      'SKU dropdown loaded successfully.',
      getApiSkuDropdown_()
    );
  }, 'Unable to load SKU dropdown.');
}


/**
 * Returns current available quantity and rate for outward SKU selection.
 */
function apiGetOutwardSkuInfo(
  payload
) {
  return safeExecute_(function () {
    requireAuthenticatedUser_();

    payload = payload || {};

    const inventoryType =
      normalizeUpper_(
        payload.inventoryType
      );

    const skuId =
      normalizeUpper_(
        payload.skuId
      );

    const sourceProjectId =
      normalizeUpper_(
        payload.sourceProjectId
      );

    if (!skuId) {
      throw new Error(
        'SKU is required.'
      );
    }

    let inventoryRecord = null;

    if (
      inventoryType === 'FREE' ||
      inventoryType ===
        'FREE INVENTORY'
    ) {
      inventoryRecord =
        getFreeInventoryRecordBySkuId_(
          skuId
        );
    } else if (
      inventoryType === 'CUSTOMIZE' ||
      inventoryType ===
        'CUSTOMIZE INVENTORY'
    ) {
      if (!sourceProjectId) {
        throw new Error(
          'Source Project is required for customized inventory.'
        );
      }

      inventoryRecord =
        getCustomizeInventoryRecord_(
          sourceProjectId,
          skuId
        );
    } else {
      throw new Error(
        'Invalid inventory type.'
      );
    }

    if (!inventoryRecord) {
      throw new Error(
        'Inventory record not found for selected SKU.'
      );
    }

    return successResponse_(
      'SKU inventory information loaded successfully.',
      {
        skuId:
          skuId,

        availableQty:
          inventoryToNumber_(
            inventoryRecord[
              'Available Qty'
            ],
            0
          ),

        rate:
          inventoryToNumber_(
            inventoryRecord[
              'Average Rate'
            ],
            0
          )
      }
    );
  }, 'Unable to load SKU inventory information.');
}



/**
 * Returns only SKUs having available stock for Outward.
 *
 * FREE:
 * - Only Free Inventory SKUs with Available Qty > 0.
 *
 * CUSTOMIZE:
 * - Source Project is required.
 * - Only that project's customized SKUs with Available Qty > 0.
 */

/**
 * Returns only projects that currently have
 * Customize Inventory Available Qty greater than zero.
 */

/**
 * Returns all active HOD users for approval selection.
 */
function apiGetHodDropdown() {
  return safeExecute_(function () {
    requireAuthenticatedUser_();

    const users =
      getSheetObjects_(
        APP_CONFIG.SHEETS.USERS
      );

    const hodMap = {};

    users.forEach(function (record) {
      const role =
        normalizeUpper_(
          record['Role'] ||
          record['User Role']
        );

      const status =
        normalizeUpper_(
          record['Status']
        );

      const email =
        normalizeLower_(
          record['Email'] ||
          record['Email Address'] ||
          record['User Email']
        );

      if (
        role !==
          APP_CONFIG.USER_ROLES.HOD ||
        status !== 'ACTIVE' ||
        !email
      ) {
        return;
      }

      const name =
        normalizeText_(
          record['Employee Name'] ||
          record['User Name'] ||
          record['Name']
        );

      hodMap[email] = {
        hodName:
          name || email,

        hodEmail:
          email,

        label:
          [
            name || email,
            name
              ? email
              : ''
          ]
            .filter(Boolean)
            .join(' - ')
      };
    });

    const hods =
      Object.keys(hodMap)
        .map(function (email) {
          return hodMap[email];
        })
        .sort(function (first, second) {
          return (
            first.hodName || ''
          ).localeCompare(
            second.hodName || '',
            undefined,
            {
              sensitivity: 'base'
            }
          );
        });

    return successResponse_(
      'HOD dropdown loaded successfully.',
      hods
    );
  }, 'Unable to load HOD dropdown.');
}


function apiGetAvailableCustomizeProjects() {
  return safeExecute_(function () {
    requireAuthenticatedUser_();

    const reportResult =
      getCustomizeInventoryReport({
        pageNumber: 1,
        pageSize: 5000
      });

    if (
      !reportResult ||
      reportResult.success !== true
    ) {
      throw new Error(
        reportResult &&
        reportResult.message
          ? reportResult.message
          : 'Customize inventory could not be loaded.'
      );
    }

    const pageData =
      reportResult.data &&
      reportResult.data.records
        ? reportResult.data.records
        : {};

    const records =
      Array.isArray(
        pageData.records
      )
        ? pageData.records
        : [];

    const projectMasterMap = {};

    getSheetObjects_(
      APP_CONFIG.SHEETS.PROJECT_MASTER
    ).forEach(function (record) {
      const projectId =
        normalizeUpper_(
          record['Project ID']
        );

      if (!projectId) {
        return;
      }

      projectMasterMap[projectId] = {
        hodName:
          normalizeText_(
            record['HOD Name']
          ),

        hodEmail:
          normalizeLower_(
            record['HOD Email']
          )
      };
    });

    const projectMap = {};

    records.forEach(function (record) {
      const availableQty =
        toNumber_(
          record.availableQty,
          0
        );

      const projectId =
        normalizeUpper_(
          record.projectId
        );

      if (
        availableQty <= 0 ||
        !projectId
      ) {
        return;
      }

      if (!projectMap[projectId]) {
        const master =
          projectMasterMap[projectId] ||
          {};

        projectMap[projectId] = {
          projectId:
            projectId,

          projectCode:
            normalizeUpper_(
              record.projectCode
            ),

          projectName:
            normalizeText_(
              record.projectName
            ),

          hodName:
            normalizeText_(
              master.hodName
            ),

          hodEmail:
            normalizeLower_(
              master.hodEmail
            ),

          availableQty:
            0
        };
      }

      projectMap[projectId]
        .availableQty +=
          availableQty;
    });

    const projects =
      Object.keys(projectMap)
        .map(function (projectId) {
          const project =
            projectMap[projectId];

          project.availableQty =
            roundTwo_(
              project.availableQty
            );

          project.label = [
            project.projectName,
            project.projectId
          ]
            .filter(Boolean)
            .join(' - ');

          return project;
        })
        .sort(function (
          first,
          second
        ) {
          return (
            first.projectName || ''
          ).localeCompare(
            second.projectName || '',
            undefined,
            {
              sensitivity: 'base'
            }
          );
        });

    return successResponse_(
      'Available Customize stock projects loaded successfully.',
      projects
    );
  }, 'Unable to load available Customize stock projects.');
}


function apiGetAvailableOutwardSkus(
  payload
) {
  return safeExecute_(function () {
    requireAuthenticatedUser_();

    payload = payload || {};

    const inventoryType =
      normalizeUpper_(
        payload.inventoryType
      );

    const sourceProjectId =
      normalizeUpper_(
        payload.sourceProjectId
      );

    let reportResult = null;

    if (
      inventoryType === 'FREE' ||
      inventoryType === 'FREE INVENTORY'
    ) {
      reportResult =
        getFreeInventoryReport({
          pageNumber: 1,
          pageSize: 5000
        });

    } else if (
      inventoryType === 'CUSTOMIZE' ||
      inventoryType === 'CUSTOMIZE INVENTORY'
    ) {
      if (!sourceProjectId) {
        return successResponse_(
          'Select Customize Stock Project first.',
          []
        );
      }

      reportResult =
        getCustomizeInventoryReport({
          projectId:
            sourceProjectId,
          pageNumber: 1,
          pageSize: 5000
        });

    } else {
      throw new Error(
        'Invalid inventory type.'
      );
    }

    if (
      !reportResult ||
      reportResult.success !== true
    ) {
      throw new Error(
        reportResult &&
        reportResult.message
          ? reportResult.message
          : 'Inventory could not be loaded.'
      );
    }

    const pageData =
      reportResult.data &&
      reportResult.data.records
        ? reportResult.data.records
        : {};

    const records =
      Array.isArray(
        pageData.records
      )
        ? pageData.records
        : [];

    const availableSkus =
      records
        .filter(function (record) {
          return (
            toNumber_(
              record.availableQty,
              0
            ) > 0 &&
            normalizeText_(
              record.skuId ||
              record.skuCode
            )
          );
        })
        .map(function (record) {
          const availableQty =
            toNumber_(
              record.availableQty,
              0
            );

          const rate =
            toNumber_(
              record.averageRate,
              0
            );

          const skuCode =
            normalizeUpper_(
              record.skuCode
            );

          const skuName =
            normalizeText_(
              record.skuName
            );

          const categoryName =
            normalizeText_(
              record.categoryName
            );

          return {
            skuId:
              normalizeUpper_(
                record.skuId
              ),

            skuCode:
              skuCode,

            skuName:
              skuName,

            categoryName:
              categoryName,

            availableQty:
              roundTwo_(
                availableQty
              ),

            rate:
              roundTwo_(rate),

            uom:
              normalizeUpper_(
                record.uom
              ),

            label: [
              skuCode,
              skuName,
              categoryName,
              'Available: ' +
                roundTwo_(
                  availableQty
                )
            ]
              .filter(Boolean)
              .join(' - ')
          };
        })
        .sort(function (first, second) {
          return first.label.localeCompare(
            second.label
          );
        });

    return successResponse_(
      'Available outward SKUs loaded successfully.',
      availableSkus
    );
  }, 'Unable to load available outward SKUs.');
}


/**
 * Builds active project dropdown directly
 * from the project master sheet.
 */
function getApiProjectDropdown_() {
  const sheetName =
    getApiConfiguredSheetName_(
      [
        'PROJECT_MASTER',
        'PROJECTS'
      ],
      'Project_Master'
    );

  return getSheetObjects_(
    sheetName
  )
    .filter(function (record) {
      return (
        normalizeText_(
          record['Project ID']
        ) &&
        normalizeUpper_(
          record['Status']
        ) === 'ACTIVE'
      );
    })
    .map(function (record) {
      const projectId =
        normalizeUpper_(
          record['Project ID']
        );

      const projectCode =
        normalizeUpper_(
          record['Project Code']
        );

      const projectName =
        normalizeText_(
          record['Project Name']
        );

      const clientName =
        normalizeText_(
          record['Client Name']
        );

      return {
        value:
          projectId,

        label: [
          projectCode,
          projectName,
          clientName
        ]
          .filter(Boolean)
          .join(' - '),

        projectId:
          projectId,

        projectCode:
          projectCode,

        projectName:
          projectName,

        clientName:
          clientName,

        hodName:
          normalizeText_(
            record['HOD Name']
          ),

        hodEmail:
          normalizeLower_(
            record['HOD Email']
          ),

        doerEmail:
          normalizeLower_(
            record['Doer Email']
          )
      };
    })
    .sort(function (first, second) {
      return first.label.localeCompare(
        second.label
      );
    });
}


/**
 * Builds active SKU dropdown directly
 * from SKU_Master.
 */
function getApiSkuDropdown_() {
  const sheetName =
    getApiConfiguredSheetName_(
      [
        'SKU_MASTER',
        'SKUS'
      ],
      'SKU_Master'
    );

  return getSheetObjects_(
    sheetName
  )
    .filter(function (record) {
      return (
        normalizeText_(
          record['SKU ID']
        ) &&
        normalizeUpper_(
          record['Status']
        ) === 'ACTIVE'
      );
    })
    .map(function (record) {
      const skuId =
        normalizeUpper_(
          record['SKU ID']
        );

      const skuCode =
        normalizeUpper_(
          record['SKU Code']
        );

      const skuName =
        normalizeText_(
          record['SKU Name']
        );

      return {
        value:
          skuId,

        label: [
          skuCode,
          skuName
        ]
          .filter(Boolean)
          .join(' - '),

        skuId:
          skuId,

        skuCode:
          skuCode,

        skuName:
          skuName,

        categoryName:
          normalizeText_(
            record['Category Name']
          ),

        brand:
          normalizeText_(
            record['Brand']
          ),

        model:
          normalizeText_(
            record['Model']
          ),

        uom:
          normalizeUpper_(
            record['UOM']
          ),

        standardRate:
          toNumber_(
            record['Standard Rate'],
            0
          ),

        gstPercent:
          toNumber_(
            record['GST Percent'],
            0
          )
      };
    })
    .sort(function (first, second) {
      return first.label.localeCompare(
        second.label
      );
    });
}


/**
 * Finds configured sheet name with fallback.
 */
function getApiConfiguredSheetName_(
  keys,
  fallback
) {
  if (
    typeof APP_CONFIG !== 'undefined' &&
    APP_CONFIG.SHEETS
  ) {
    for (
      let index = 0;
      index < keys.length;
      index++
    ) {
      const value =
        APP_CONFIG.SHEETS[
          keys[index]
        ];

      if (value) {
        return value;
      }
    }
  }

  return fallback;
}


/**
 * Inventory APIs.
 */
function apiGetInventorySummary() {
  return getInventorySummary();
}


function apiGetFreeInventory(
  filters
) {
  return getFreeInventoryReport(
    normalizeApiFilters_(
      filters
    )
  );
}


function apiGetCustomizeInventory(
  filters
) {
  return getCustomizeInventoryReport(
    normalizeApiFilters_(
      filters
    )
  );
}


function apiGetLowStock(
  filters
) {
  return getLowStockReport(
    normalizeApiFilters_(
      filters
    )
  );
}


function apiGetEligibleCustomize(
  filters
) {
  return getEligibleCustomizeInventory(
    normalizeApiFilters_(
      filters
    )
  );
}


function apiRefreshEligibility() {
  return refreshCustomizeFreeEligibility();
}


/**
 * Inward APIs.
 */
function apiCreateInward(
  payload
) {
  return createInward(payload);
}


function apiGetInwardTransactions(
  filters
) {
  return getInwardTransactions(
    normalizeApiFilters_(
      filters
    )
  );
}


function apiGetInwardByNumber(
  inwardNo
) {
  return getInwardByNumber(
    inwardNo
  );
}


function apiGetInwardSummary() {
  return getInwardSummary();
}


/**
 * Outward APIs.
 */
function apiCreateOutward(
  payload
) {
  return createOutward(payload);
}


function apiGetOutwardTransactions(
  filters
) {
  return getOutwardTransactions(
    normalizeApiFilters_(
      filters
    )
  );
}


function apiGetOutwardByNumber(
  outwardNo
) {
  return getOutwardByNumber(
    outwardNo
  );
}


function apiGetOutwardSummary() {
  return getOutwardSummary();
}


/**
 * Approval APIs.
 */
function apiCreateApprovalRequest(
  payload
) {
  return createCrossProjectApprovalRequest(
    payload
  );
}


function apiGetApprovalRequests(
  filters
) {
  return getApprovalRequests(
    normalizeApiFilters_(
      filters
    )
  );
}


function apiGetPendingApprovals() {
  return getPendingApprovalRequests();
}


function apiGetApprovalRequest(
  requestId
) {
  return getApprovalRequest(
    requestId
  );
}


function apiApproveRequest(
  requestId,
  remarks
) {
  return approveApprovalRequest(
    requestId,
    remarks
  );
}


function apiRejectRequest(
  requestId,
  reason
) {
  return rejectApprovalRequest(
    requestId,
    reason
  );
}


function apiExecuteApprovedOutward(
  requestId
) {
  return executeApprovedOutward(
    requestId
  );
}


function apiResendApprovalEmail(
  requestId
) {
  return resendApprovalRequestEmail(
    requestId
  );
}


/**
 * Conversion APIs.
 */
function apiCreateConversionRequest(
  payload
) {
  return createConversionRequest(
    payload
  );
}


function apiGetConversionRequests(
  filters
) {
  return getConversionRequests(
    normalizeApiFilters_(
      filters
    )
  );
}


function apiGetConversionRequest(
  conversionId
) {
  return getConversionRequest(
    conversionId
  );
}


function apiApproveConversion(
  conversionId,
  remarks
) {
  return approveConversionRequest(
    conversionId,
    remarks
  );
}


function apiRejectConversion(
  conversionId,
  reason
) {
  return rejectConversionRequest(
    conversionId,
    reason
  );
}


function apiExecuteConversion(
  conversionId
) {
  return executeApprovedConversion(
    conversionId
  );
}


function apiGetConversionSummary() {
  return getConversionSummary();
}


/**
 * Report APIs.
 */
function apiGetReportDashboard(
  filters
) {
  return getReportDashboard(
    normalizeApiFilters_(
      filters
    )
  );
}


function apiGetReport(
  reportType,
  filters
) {
  const type =
    normalizeUpper_(
      reportType
    );

  const cleanFilters =
    normalizeApiFilters_(
      filters
    );

  if (
    type === 'FREE_INVENTORY'
  ) {
    return getFreeInventoryReport(
      cleanFilters
    );
  }

  if (
    type ===
    'CUSTOMIZE_INVENTORY'
  ) {
    return getCustomizeInventoryReport(
      cleanFilters
    );
  }

  if (type === 'INWARD') {
    return getInwardReport(
      cleanFilters
    );
  }

  if (type === 'OUTWARD') {
    return getOutwardReport(
      cleanFilters
    );
  }

  if (type === 'APPROVAL') {
    return getApprovalReport(
      cleanFilters
    );
  }

  if (type === 'CONVERSION') {
    return getConversionReport(
      cleanFilters
    );
  }

  if (type === 'LOW_STOCK') {
    return getLowStockReport(
      cleanFilters
    );
  }

  if (
    type ===
    'ELIGIBLE_FOR_FREE'
  ) {
    return getEligibleForFreeReport(
      cleanFilters
    );
  }

  if (
    type ===
    'PROJECT_WISE_CUSTOMIZE'
  ) {
    return getProjectWiseCustomizeReport(
      cleanFilters
    );
  }

  return errorResponse_(
    new Error(
      'Invalid report type: ' +
      type
    )
  );
}


function apiExportReportCsv(
  reportType,
  filters
) {
  return exportReportCsv(
    reportType,
    normalizeApiFilters_(
      filters
    )
  );
}


/**
 * Generic frontend API dispatcher.
 *
 * The frontend may call:
 * google.script.run.apiDispatch(action, payload)
 */


/**
 * Fast dashboard API.
 *
 * Reads each required sheet only once and returns:
 * - Existing dashboard counters
 * - SKU-wise combined inventory details
 * - Free/Customize inward totals
 * - Free/Customize outward totals
 *
 * Result is cached briefly to keep the web app fast.
 */
function getFastInventoryDashboard_() {
  return safeExecute_(function () {
    requireAuthenticatedUser_();

    const cache =
      CacheService.getScriptCache();

    const cacheKey =
      'FAST_INVENTORY_DASHBOARD_V5';

    const cached =
      cache.get(cacheKey);

    if (cached) {
      return JSON.parse(cached);
    }

    /*
     * Use the already-tested inventory and movement
     * service functions instead of reading raw sheet
     * columns again.
     */
    const freeResult =
      getFreeInventoryReport({
        pageNumber: 1,
        pageSize: 5000
      });

    const customizeResult =
      getCustomizeInventoryReport({
        pageNumber: 1,
        pageSize: 5000
      });

    const inwardSummaryResult =
      getInwardSummary();

    const outwardSummaryResult =
      getOutwardSummary();

    if (
      !freeResult ||
      freeResult.success !== true
    ) {
      throw new Error(
        freeResult &&
        freeResult.message
          ? freeResult.message
          : 'Free inventory could not be loaded.'
      );
    }

    if (
      !customizeResult ||
      customizeResult.success !== true
    ) {
      throw new Error(
        customizeResult &&
        customizeResult.message
          ? customizeResult.message
          : 'Customize inventory could not be loaded.'
      );
    }

    const freePage =
      freeResult.data &&
      freeResult.data.records
        ? freeResult.data.records
        : {};

    const customizePage =
      customizeResult.data &&
      customizeResult.data.records
        ? customizeResult.data.records
        : {};

    const freeRows =
      Array.isArray(
        freePage.records
      )
        ? freePage.records
        : [];

    const customizeRows =
      Array.isArray(
        customizePage.records
      )
        ? customizePage.records
        : [];

    const freeSummary =
      freeResult.data &&
      freeResult.data.summary
        ? freeResult.data.summary
        : {};

    const customizeSummary =
      customizeResult.data &&
      customizeResult.data.summary
        ? customizeResult.data.summary
        : {};

    const inwardSummary =
      inwardSummaryResult &&
      inwardSummaryResult.success &&
      inwardSummaryResult.data
        ? inwardSummaryResult.data
        : {};

    const outwardSummary =
      outwardSummaryResult &&
      outwardSummaryResult.success &&
      outwardSummaryResult.data
        ? outwardSummaryResult.data
        : {};

    const skuMap = {};

    function skuKey_(
      record
    ) {
      return (
        normalizeUpper_(
          record.skuId
        ) ||
        normalizeUpper_(
          record.skuCode
        )
      );
    }

    function ensureSku_(
      record
    ) {
      const key =
        skuKey_(record);

      if (!key) {
        return null;
      }

      if (!skuMap[key]) {
        skuMap[key] = {
          skuId:
            normalizeUpper_(
              record.skuId
            ),

          skuCode:
            normalizeUpper_(
              record.skuCode
            ),

          category:
            normalizeText_(
              record.categoryName
            ),

          skuName:
            normalizeText_(
              record.skuName
            ),

          minimumQtyFree:
            toNumber_(
              record.minimumStock,
              0
            ),

          freeAvailableQty: 0,
          customizeAvailableQty: 0,

          freeStockValue: 0,
          customizeStockValue: 0,

          freeRate: 0,
          customizeRate: 0
        };
      }

      const item =
        skuMap[key];

      if (!item.skuId) {
        item.skuId =
          normalizeUpper_(
            record.skuId
          );
      }

      if (!item.skuCode) {
        item.skuCode =
          normalizeUpper_(
            record.skuCode
          );
      }

      if (!item.category) {
        item.category =
          normalizeText_(
            record.categoryName
          );
      }

      if (!item.skuName) {
        item.skuName =
          normalizeText_(
            record.skuName
          );
      }

      return item;
    }

    freeRows.forEach(
      function (record) {
        const item =
          ensureSku_(record);

        if (!item) {
          return;
        }

        const availableQty =
          toNumber_(
            record.availableQty,
            0
          );

        const rate =
          toNumber_(
            record.averageRate,
            0
          );

        const stockValue =
          typeof record.stockValue !==
            'undefined'
            ? toNumber_(
                record.stockValue,
                0
              )
            : availableQty * rate;

        item.minimumQtyFree =
          toNumber_(
            record.minimumStock,
            item.minimumQtyFree
          );

        item.freeAvailableQty +=
          availableQty;

        item.freeStockValue +=
          stockValue;

        if (rate > 0) {
          item.freeRate = rate;
        }
      }
    );

    customizeRows.forEach(
      function (record) {
        const item =
          ensureSku_(record);

        if (!item) {
          return;
        }

        const availableQty =
          toNumber_(
            record.availableQty,
            0
          );

        const rate =
          toNumber_(
            record.averageRate,
            0
          );

        const stockValue =
          typeof record.stockValue !==
            'undefined'
            ? toNumber_(
                record.stockValue,
                0
              )
            : availableQty * rate;

        item.customizeAvailableQty +=
          availableQty;

        item.customizeStockValue +=
          stockValue;

        if (rate > 0) {
          item.customizeRate = rate;
        }
      }
    );

    const skuDetails =
      Object.keys(skuMap)
        .map(function (key) {
          const item =
            skuMap[key];

          const totalAvailableQty =
            item.freeAvailableQty +
            item.customizeAvailableQty;

          const totalStockValue =
            item.freeStockValue +
            item.customizeStockValue;

          const combinedRate =
            totalAvailableQty > 0
              ? (
                  totalStockValue /
                  totalAvailableQty
                )
              : (
                  item.freeRate ||
                  item.customizeRate ||
                  0
                );

          return {
            skuId:
              item.skuId,

            skuCode:
              item.skuCode,

            category:
              item.category,

            skuName:
              item.skuName,

            minimumQtyFree:
              roundTwo_(
                item.minimumQtyFree
              ),

            freeAvailableQty:
              roundTwo_(
                item.freeAvailableQty
              ),

            customizeAvailableQty:
              roundTwo_(
                item.customizeAvailableQty
              ),

            totalAvailableQty:
              roundTwo_(
                totalAvailableQty
              ),

            rate:
              roundTwo_(
                combinedRate
              ),

            freeRate:
              roundTwo_(
                item.freeRate
              ),

            customizeRate:
              roundTwo_(
                item.customizeRate
              ),

            freeStockValue:
              roundTwo_(
                item.freeStockValue
              ),

            customizeStockValue:
              roundTwo_(
                item.customizeStockValue
              ),

            totalStockValue:
              roundTwo_(
                totalStockValue
              )
          };
        })
        .sort(function (
          first,
          second
        ) {
          return (
            first.skuCode || ''
          ).localeCompare(
            second.skuCode || ''
          );
        });

    const movementTotals = {
      freeInwardQty:
        toNumber_(
          inwardSummary.freeQuantity,
          0
        ),

      freeInwardAmount:
        toNumber_(
          inwardSummary.freeAmount,
          0
        ),

      customizeInwardQty:
        toNumber_(
          inwardSummary.customizeQuantity,
          0
        ),

      customizeInwardAmount:
        toNumber_(
          inwardSummary.customizeAmount,
          0
        ),

      freeOutwardQty:
        toNumber_(
          outwardSummary.freeQuantity,
          0
        ),

      freeOutwardAmount:
        toNumber_(
          outwardSummary.freeAmount,
          0
        ),

      customizeOutwardQty:
        toNumber_(
          outwardSummary.customizeQuantity,
          0
        ),

      customizeOutwardAmount:
        toNumber_(
          outwardSummary.customizeAmount,
          0
        )
    };

    /*
     * Current summaries may only return a combined
     * totalAmount. Calculate type-wise amounts from
     * the transaction services when those fields are
     * not present.
     */
    if (
      !movementTotals.freeInwardAmount &&
      !movementTotals.customizeInwardAmount
    ) {
      const inwardTransactions =
        getInwardTransactions({
          pageNumber: 1,
          pageSize: 5000
        });

      if (
        inwardTransactions &&
        inwardTransactions.success &&
        inwardTransactions.data &&
        Array.isArray(
          inwardTransactions.data.records
        )
      ) {
        inwardTransactions.data.records
          .forEach(function (
            transaction
          ) {
            const type =
              normalizeUpper_(
                transaction.inventoryType
              );

            const amount =
              toNumber_(
                transaction.totalAmount,
                0
              );

            if (type === 'FREE') {
              movementTotals
                .freeInwardAmount +=
                  amount;
            } else if (
              type === 'CUSTOMIZE'
            ) {
              movementTotals
                .customizeInwardAmount +=
                  amount;
            }
          });
      }
    }

    if (
      !movementTotals.freeOutwardAmount &&
      !movementTotals.customizeOutwardAmount
    ) {
      const outwardTransactions =
        getOutwardTransactions({
          pageNumber: 1,
          pageSize: 5000
        });

      if (
        outwardTransactions &&
        outwardTransactions.success &&
        outwardTransactions.data &&
        Array.isArray(
          outwardTransactions.data.records
        )
      ) {
        outwardTransactions.data.records
          .forEach(function (
            transaction
          ) {
            const type =
              normalizeUpper_(
                transaction.inventoryType
              );

            const amount =
              toNumber_(
                transaction.totalAmount,
                0
              );

            if (type === 'FREE') {
              movementTotals
                .freeOutwardAmount +=
                  amount;
            } else if (
              type === 'CUSTOMIZE'
            ) {
              movementTotals
                .customizeOutwardAmount +=
                  amount;
            }
          });
      }
    }

    Object.keys(movementTotals)
      .forEach(function (key) {
        movementTotals[key] =
          roundTwo_(
            movementTotals[key]
          );
      });

    const inventory = {
      freeInventory: {
        totalSkuRecords:
          toNumber_(
            freeSummary.totalRecords,
            freeRows.length
          ),

        totalAvailableQty:
          roundTwo_(
            skuDetails.reduce(
              function (
                total,
                item
              ) {
                return (
                  total +
                  item.freeAvailableQty
                );
              },
              0
            )
          ),

        totalStockValue:
          roundTwo_(
            skuDetails.reduce(
              function (
                total,
                item
              ) {
                return (
                  total +
                  item.freeStockValue
                );
              },
              0
            )
          )
      },

      customizeInventory: {
        totalRecords:
          toNumber_(
            customizeSummary.totalRecords,
            customizeRows.length
          ),

        totalAvailableQty:
          roundTwo_(
            skuDetails.reduce(
              function (
                total,
                item
              ) {
                return (
                  total +
                  item
                    .customizeAvailableQty
                );
              },
              0
            )
          ),

        totalStockValue:
          roundTwo_(
            skuDetails.reduce(
              function (
                total,
                item
              ) {
                return (
                  total +
                  item
                    .customizeStockValue
                );
              },
              0
            )
          )
      },

      combined: {
        totalAvailableQty:
          roundTwo_(
            skuDetails.reduce(
              function (
                total,
                item
              ) {
                return (
                  total +
                  item.totalAvailableQty
                );
              },
              0
            )
          ),

        totalStockValue:
          roundTwo_(
            skuDetails.reduce(
              function (
                total,
                item
              ) {
                return (
                  total +
                  item.totalStockValue
                );
              },
              0
            )
          )
      }
    };

    const response =
      successResponse_(
        'Dashboard loaded successfully.',
        {
          inventory:
            inventory,

          movementTotals:
            movementTotals,

          skuDetails:
            skuDetails,

          lowStock: {
            totalRecords:
              skuDetails.filter(
                function (item) {
                  return (
                    item.minimumQtyFree >
                      0 &&
                    item.freeAvailableQty <=
                      item.minimumQtyFree
                  );
                }
              ).length
          },

          debug: {
            freeRecords:
              freeRows.length,

            customizeRecords:
              customizeRows.length,

            skuRows:
              skuDetails.length,

            freeInwardQty:
              movementTotals
                .freeInwardQty,

            freeOutwardQty:
              movementTotals
                .freeOutwardQty
          }
        }
      );

    cache.put(
      cacheKey,
      JSON.stringify(response),
      30
    );

    return response;
  }, 'Unable to load inventory dashboard.');
}


/**
 * Clears the short dashboard cache after stock movement.
 */
function clearFastDashboardCache_() {
  const cache =
    CacheService.getScriptCache();

  [
    'FAST_INVENTORY_DASHBOARD_V2',
    'FAST_INVENTORY_DASHBOARD_V3',
    'FAST_INVENTORY_DASHBOARD_V5'
  ].forEach(function (key) {
    cache.remove(key);
  });
}




/**
 * Run this function to verify dashboard data.
 */
function testFastInventoryDashboard() {
  clearFastDashboardCache_();

  const result =
    getFastInventoryDashboard_();

  Logger.log(
    JSON.stringify(
      result,
      null,
      2
    )
  );

  return result;
}

function apiDispatch(
  action,
  payload
) {
  const cleanAction =
    normalizeUpper_(action);

  payload = payload || {};

  const routes = {
    GET_BOOTSTRAP:
      function () {
        return apiGetBootstrapData();
      },

    GET_CURRENT_USER:
      function () {
        return apiGetCurrentUser();
      },

    GET_DASHBOARD:
      function () {
        return apiGetDashboard();
      },

    GET_DASHBOARD_COUNTERS:
      function () {
        return apiGetDashboardCounters();
      },

    GET_PROJECT_DROPDOWN:
      function () {
        return apiGetProjectDropdown();
      },

    GET_SKU_DROPDOWN:
      function () {
        return apiGetSkuDropdown();
      },

    GET_OUTWARD_SKU_INFO:
      function () {
        return apiGetOutwardSkuInfo(
          payload
        );
      },

    GET_AVAILABLE_OUTWARD_SKUS:
      function () {
        return apiGetAvailableOutwardSkus(
          payload
        );
      },

    GET_AVAILABLE_CUSTOMIZE_PROJECTS:
      function () {
        return apiGetAvailableCustomizeProjects();
      },

    GET_HOD_DROPDOWN:
      function () {
        return apiGetHodDropdown();
      },


    GET_INVENTORY_SUMMARY:
      function () {
        return apiGetInventorySummary();
      },

    GET_FREE_INVENTORY:
      function () {
        return apiGetFreeInventory(
          payload.filters || payload
        );
      },

    GET_CUSTOMIZE_INVENTORY:
      function () {
        return apiGetCustomizeInventory(
          payload.filters || payload
        );
      },

    GET_LOW_STOCK:
      function () {
        return apiGetLowStock(
          payload.filters || payload
        );
      },

    GET_ELIGIBLE_CUSTOMIZE:
      function () {
        return apiGetEligibleCustomize(
          payload.filters || payload
        );
      },

    REFRESH_ELIGIBILITY:
      function () {
        return apiRefreshEligibility();
      },

    CREATE_INWARD:
      function () {
        return apiCreateInward(
          payload
        );
      },

    GET_INWARD_TRANSACTIONS:
      function () {
        return apiGetInwardTransactions(
          payload.filters || payload
        );
      },

    GET_INWARD:
      function () {
        return apiGetInwardByNumber(
          payload.inwardNo
        );
      },

    CREATE_OUTWARD:
      function () {
        return apiCreateOutward(
          payload
        );
      },

    GET_OUTWARD_TRANSACTIONS:
      function () {
        return apiGetOutwardTransactions(
          payload.filters || payload
        );
      },

    GET_OUTWARD:
      function () {
        return apiGetOutwardByNumber(
          payload.outwardNo
        );
      },

    CREATE_APPROVAL_REQUEST:
      function () {
        return apiCreateApprovalRequest(
          payload
        );
      },

    GET_APPROVAL_REQUESTS:
      function () {
        return apiGetApprovalRequests(
          payload.filters || payload
        );
      },

    GET_APPROVAL_REQUEST:
      function () {
        return apiGetApprovalRequest(
          payload.requestId
        );
      },

    APPROVE_REQUEST:
      function () {
        return apiApproveRequest(
          payload.requestId,
          payload.remarks
        );
      },

    REJECT_REQUEST:
      function () {
        return apiRejectRequest(
          payload.requestId,
          payload.reason
        );
      },

    EXECUTE_APPROVED_OUTWARD:
      function () {
        return apiExecuteApprovedOutward(
          payload.requestId
        );
      },

    RESEND_APPROVAL_EMAIL:
      function () {
        return apiResendApprovalEmail(
          payload.requestId
        );
      },

    CREATE_CONVERSION_REQUEST:
      function () {
        return apiCreateConversionRequest(
          payload
        );
      },

    GET_CONVERSION_REQUESTS:
      function () {
        return apiGetConversionRequests(
          payload.filters || payload
        );
      },

    GET_CONVERSION_REQUEST:
      function () {
        return apiGetConversionRequest(
          payload.conversionId
        );
      },

    APPROVE_CONVERSION:
      function () {
        return apiApproveConversion(
          payload.conversionId,
          payload.remarks
        );
      },

    REJECT_CONVERSION:
      function () {
        return apiRejectConversion(
          payload.conversionId,
          payload.reason
        );
      },

    EXECUTE_CONVERSION:
      function () {
        return apiExecuteConversion(
          payload.conversionId
        );
      },

    GET_REPORT_DASHBOARD:
      function () {
        return apiGetReportDashboard(
          payload.filters || payload
        );
      },

    GET_REPORT:
      function () {
        return apiGetReport(
          payload.reportType,
          payload.filters || {}
        );
      },

    EXPORT_REPORT_CSV:
      function () {
        return apiExportReportCsv(
          payload.reportType,
          payload.filters || {}
        );
      }
  };

  const route =
    routes[cleanAction];

  if (!route) {
    return errorResponse_(
      new Error(
        'Unknown API action: ' +
        cleanAction
      )
    );
  }

  try {
    return route();
  } catch (error) {
    return errorResponse_(
      error,
      'API operation failed.'
    );
  }
}


/**
 * Normalizes frontend filters.
 */
function normalizeApiFilters_(
  filters
) {
  filters =
    filters &&
    typeof filters === 'object'
      ? Object.assign(
          {},
          filters
        )
      : {};

  filters.pageNumber =
    Math.max(
      1,
      toNumber_(
        filters.pageNumber,
        1
      )
    );

  filters.pageSize =
    Math.max(
      1,
      Math.min(
        toNumber_(
          filters.pageSize,
          MAIN_API_CONFIG
            .DEFAULT_PAGE_SIZE
        ),
        MAIN_API_CONFIG
          .MAX_PAGE_SIZE
      )
    );

  return filters;
}


/**
 * Public health endpoint.
 *
 * Does not expose private inventory data.
 */
function getPublicAppHealth_() {
  return {
    success: true,
    appName:
      MAIN_API_CONFIG.APP_NAME,
    version:
      MAIN_API_CONFIG.VERSION,
    status:
      'ONLINE',
    currentDateTime:
      formatDateTime_(
        new Date()
      )
  };
}


/**
 * Authenticated application health.
 */
function apiGetAppHealth() {
  return safeExecute_(function () {
    const session =
      requireAuthenticatedUser_();

    return successResponse_(
      'Application health check successful.',
      {
        appName:
          MAIN_API_CONFIG.APP_NAME,

        version:
          MAIN_API_CONFIG.VERSION,

        status:
          'ONLINE',

        currentUser:
          normalizeApiSession_(
            session
          ),

        currentDateTime:
          formatDateTime_(
            new Date()
          ),

        webAppUrl:
          getMainWebAppUrl_(),

        services: {
          dashboard:
            typeof getDashboardData ===
            'function',

          inventory:
            typeof getInventorySummary ===
            'function',

          inward:
            typeof createInward ===
            'function',

          outward:
            typeof createOutward ===
            'function',

          approval:
            typeof createCrossProjectApprovalRequest ===
            'function',

          email:
            typeof sendApprovalRequestEmail_ ===
            'function',

          conversion:
            typeof createConversionRequest ===
            'function',

          reports:
            typeof getReportDashboard ===
            'function'
        }
      }
    );
  }, 'Application health check failed.');
}


/**
 * Basic page shown before Index.html is created.
 */
function buildBackendReadyPage_() {
  return (
    '<!DOCTYPE html>' +
    '<html>' +
    '<head>' +
      '<base target="_top">' +
      '<meta name="viewport" content="width=device-width,initial-scale=1">' +
      '<style>' +
        'body{font-family:Arial,sans-serif;background:#f3f4f6;margin:0;padding:24px;color:#111827;}' +
        '.card{max-width:760px;margin:60px auto;background:#fff;border:1px solid #e5e7eb;border-radius:14px;padding:28px;box-shadow:0 12px 32px rgba(0,0,0,.08);}' +
        'h1{margin:0 0 12px;color:#1f4e79;font-size:26px;}' +
        '.ok{display:inline-block;background:#dcfce7;color:#166534;border-radius:999px;padding:7px 12px;font-weight:700;}' +
        'p{line-height:1.6;}' +
        'code{background:#f3f4f6;padding:3px 6px;border-radius:5px;}' +
      '</style>' +
    '</head>' +
    '<body>' +
      '<div class="card">' +
        '<div class="ok">Backend Online</div>' +
        '<h1>' +
          escapeHtml_(
            MAIN_API_CONFIG.APP_NAME
          ) +
        '</h1>' +
        '<p>Main API is working. The next step is to create <code>Index.html</code> and the frontend files.</p>' +
        '<p><strong>Version:</strong> ' +
          escapeHtml_(
            MAIN_API_CONFIG.VERSION
          ) +
        '</p>' +
      '</div>' +
    '</body>' +
    '</html>'
  );
}


/**
 * Main route error page.
 */
function buildMainErrorPage_(
  message
) {
  return (
    '<!DOCTYPE html>' +
    '<html>' +
    '<head>' +
      '<meta name="viewport" content="width=device-width,initial-scale=1">' +
      '<style>' +
        'body{font-family:Arial,sans-serif;background:#f3f4f6;padding:24px;}' +
        '.error{max-width:680px;margin:60px auto;background:#fee2e2;color:#991b1b;border-radius:12px;padding:24px;}' +
      '</style>' +
    '</head>' +
    '<body>' +
      '<div class="error">' +
        '<h2>Application Error</h2>' +
        '<p>' +
          escapeHtml_(message) +
        '</p>' +
      '</div>' +
    '</body>' +
    '</html>'
  );
}


/**
 * Test main API without opening frontend.
 */
function testMainApi() {
  const result = {
    health:
      apiGetAppHealth(),

    bootstrap:
      apiGetBootstrapData(),

    dashboard:
      apiDispatch(
        'GET_DASHBOARD',
        {}
      ),

    projects:
      apiDispatch(
        'GET_PROJECT_DROPDOWN',
        {}
      ),

    skus:
      apiDispatch(
        'GET_SKU_DROPDOWN',
        {}
      )
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
 * Compact test to avoid oversized logs.
 */
function testMainApiCompact() {
  const health =
    apiGetAppHealth();

  const bootstrap =
    apiGetBootstrapData();

  const result = {
    healthSuccess:
      health.success,

    bootstrapSuccess:
      bootstrap.success,

    appName:
      bootstrap.success
        ? bootstrap.data.app.name
        : '',

    userEmail:
      bootstrap.success
        ? bootstrap.data.user.email
        : '',

    projectCount:
      bootstrap.success
        ? bootstrap.data.projects.length
        : 0,

    skuCount:
      bootstrap.success
        ? bootstrap.data.skus.length
        : 0,

    freeAvailableQty:
      bootstrap.success
        ? bootstrap.data
            .dashboardCounters
            .freeAvailableQty
        : 0,

    customizeAvailableQty:
      bootstrap.success
        ? bootstrap.data
            .dashboardCounters
            .customizeAvailableQty
        : 0
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
