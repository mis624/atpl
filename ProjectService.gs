/**
 * PROJECT INVENTORY MANAGEMENT SYSTEM
 * File: ProjectService.gs
 *
 * Handles project master operations:
 * - Create project
 * - Update project
 * - View project
 * - Search projects
 * - Assign HOD and Doer
 * - Activate, close, hold and cancel project
 * - Project dropdown
 * - Soft delete
 * - Transaction logging
 */

const PROJECT_SERVICE_CONFIG = Object.freeze({
  SHEET_NAME: 'Project_Master',

  STATUS: Object.freeze({
    ACTIVE: 'ACTIVE',
    ON_HOLD: 'ON HOLD',
    COMPLETED: 'COMPLETED',
    CLOSED: 'CLOSED',
    CANCELLED: 'CANCELLED',
    INACTIVE: 'INACTIVE'
  }),

  INVENTORY_TYPE: Object.freeze({
    FREE: 'FREE INVENTORY',
    CUSTOMIZE: 'CUSTOMIZE INVENTORY'
  }),

  DEFAULT_STATUS: 'ACTIVE',
  PROJECT_ID_PREFIX: 'PRJ',
  PROJECT_CODE_PREFIX: 'PROJ'
});

/**
 * Returns the configured project sheet name.
 */
function getProjectSheetName_() {
  if (
    typeof APP_CONFIG !== 'undefined' &&
    APP_CONFIG.SHEETS &&
    APP_CONFIG.SHEETS.PROJECT_MASTER
  ) {
    return APP_CONFIG.SHEETS.PROJECT_MASTER;
  }

  return PROJECT_SERVICE_CONFIG.SHEET_NAME;
}

/**
 * Returns all available project statuses.
 */
function getProjectStatuses() {
  return successResponse_(
    'Project statuses loaded successfully.',
    Object.keys(PROJECT_SERVICE_CONFIG.STATUS).map(function (key) {
      return PROJECT_SERVICE_CONFIG.STATUS[key];
    })
  );
}

/**
 * Creates a new project.
 *
 * Expected projectData:
 * {
 *   projectName: '',
 *   clientName: '',
 *   projectLocation: '',
 *   projectManager: '',
 *   hodEmail: '',
 *   doerEmail: '',
 *   startDate: '',
 *   expectedEndDate: '',
 *   status: 'ACTIVE',
 *   remarks: ''
 * }
 */
function createProject(projectData) {
  return safeExecute_(function () {
    const session = requireRole_([
      APP_CONFIG.USER_ROLES.ADMIN,
      APP_CONFIG.USER_ROLES.HOD
    ]);

    projectData = projectData || {};

    validateRequiredFields_(
      projectData,
      [
        'projectName',
        'clientName'
      ]
    );

    const projectName = normalizeText_(projectData.projectName);
    const clientName = normalizeText_(projectData.clientName);
    const projectLocation = normalizeText_(
      projectData.projectLocation
    );
    const projectManager = normalizeText_(
      projectData.projectManager
    );
    const hodEmail = normalizeLower_(projectData.hodEmail);
    const doerEmail = normalizeLower_(projectData.doerEmail);
    const startDate = parseDate_(projectData.startDate);
    const expectedEndDate = parseDate_(
      projectData.expectedEndDate
    );
    const status = normalizeProjectStatus_(
      projectData.status ||
      PROJECT_SERVICE_CONFIG.DEFAULT_STATUS
    );
    const remarks = normalizeText_(projectData.remarks);

    if (projectName.length < 2) {
      throw new Error(
        'Project name must contain at least 2 characters.'
      );
    }

    if (clientName.length < 2) {
      throw new Error(
        'Client name must contain at least 2 characters.'
      );
    }

    validateProjectDates_(startDate, expectedEndDate);

    if (hodEmail) {
      validateProjectAssignedUser_(
        hodEmail,
        APP_CONFIG.USER_ROLES.HOD,
        'HOD'
      );
    }

    if (doerEmail) {
      validateProjectAssignedUser_(
        doerEmail,
        APP_CONFIG.USER_ROLES.DOER,
        'Doer'
      );
    }

    if (
      isDuplicateProjectName_(
        projectName,
        clientName,
        null
      )
    ) {
      throw new Error(
        'A project with the same project name and client name already exists.'
      );
    }

    const projectId = generateNextId_(
      'PROJECT',
      PROJECT_SERVICE_CONFIG.PROJECT_ID_PREFIX
    );

    const projectCode = generateProjectCode_(
      projectName,
      projectId
    );

    const now = new Date();

    appendObjectRow_(
      getProjectSheetName_(),
      {
        'Project ID': projectId,
        'Project Code': projectCode,
        'Project Name': projectName,
        'Client Name': clientName,
        'Project Location': projectLocation,
        'Project Manager': projectManager,
        'HOD Email': hodEmail,
        'Doer Email': doerEmail,
        'Start Date': startDate || '',
        'Expected End Date': expectedEndDate || '',
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
      'PROJECT',
      'CREATE',
      projectId,
      {
        projectCode: projectCode,
        projectName: projectName,
        clientName: clientName,
        hodEmail: hodEmail,
        doerEmail: doerEmail,
        status: status
      }
    );

    return successResponse_(
      'Project created successfully.',
      {
        projectId: projectId,
        projectCode: projectCode,
        projectName: projectName,
        clientName: clientName,
        status: status
      }
    );
  }, 'Unable to create project.');
}

/**
 * Updates an existing project.
 */
function updateProject(projectId, projectData) {
  return safeExecute_(function () {
    const session = requireRole_([
      APP_CONFIG.USER_ROLES.ADMIN,
      APP_CONFIG.USER_ROLES.HOD
    ]);

    projectData = projectData || {};

    const project = getProjectRecordById_(projectId);

    if (!project) {
      throw new Error('Project not found: ' + projectId);
    }

    const currentStatus = normalizeProjectStatus_(
      project['Status']
    );

    if (
      currentStatus ===
        PROJECT_SERVICE_CONFIG.STATUS.CANCELLED &&
      session.role !== APP_CONFIG.USER_ROLES.ADMIN
    ) {
      throw new Error(
        'Only ADMIN can update a cancelled project.'
      );
    }

    const updateData = {
      'Updated By': session.email,
      'Updated At': new Date()
    };

    const updatedProjectName =
      Object.prototype.hasOwnProperty.call(
        projectData,
        'projectName'
      )
        ? normalizeText_(projectData.projectName)
        : normalizeText_(project['Project Name']);

    const updatedClientName =
      Object.prototype.hasOwnProperty.call(
        projectData,
        'clientName'
      )
        ? normalizeText_(projectData.clientName)
        : normalizeText_(project['Client Name']);

    if (!updatedProjectName) {
      throw new Error('Project name cannot be blank.');
    }

    if (!updatedClientName) {
      throw new Error('Client name cannot be blank.');
    }

    if (
      isDuplicateProjectName_(
        updatedProjectName,
        updatedClientName,
        projectId
      )
    ) {
      throw new Error(
        'Another project with the same project name and client name already exists.'
      );
    }

    if (
      Object.prototype.hasOwnProperty.call(
        projectData,
        'projectName'
      )
    ) {
      updateData['Project Name'] = updatedProjectName;
    }

    if (
      Object.prototype.hasOwnProperty.call(
        projectData,
        'clientName'
      )
    ) {
      updateData['Client Name'] = updatedClientName;
    }

    if (
      Object.prototype.hasOwnProperty.call(
        projectData,
        'projectLocation'
      )
    ) {
      updateData['Project Location'] = normalizeText_(
        projectData.projectLocation
      );
    }

    if (
      Object.prototype.hasOwnProperty.call(
        projectData,
        'projectManager'
      )
    ) {
      updateData['Project Manager'] = normalizeText_(
        projectData.projectManager
      );
    }

    if (
      Object.prototype.hasOwnProperty.call(
        projectData,
        'hodEmail'
      )
    ) {
      const hodEmail = normalizeLower_(projectData.hodEmail);

      if (hodEmail) {
        validateProjectAssignedUser_(
          hodEmail,
          APP_CONFIG.USER_ROLES.HOD,
          'HOD'
        );
      }

      updateData['HOD Email'] = hodEmail;
    }

    if (
      Object.prototype.hasOwnProperty.call(
        projectData,
        'doerEmail'
      )
    ) {
      const doerEmail = normalizeLower_(projectData.doerEmail);

      if (doerEmail) {
        validateProjectAssignedUser_(
          doerEmail,
          APP_CONFIG.USER_ROLES.DOER,
          'Doer'
        );
      }

      updateData['Doer Email'] = doerEmail;
    }

    const updatedStartDate =
      Object.prototype.hasOwnProperty.call(
        projectData,
        'startDate'
      )
        ? parseDate_(projectData.startDate)
        : parseDate_(project['Start Date']);

    const updatedExpectedEndDate =
      Object.prototype.hasOwnProperty.call(
        projectData,
        'expectedEndDate'
      )
        ? parseDate_(projectData.expectedEndDate)
        : parseDate_(project['Expected End Date']);

    validateProjectDates_(
      updatedStartDate,
      updatedExpectedEndDate
    );

    if (
      Object.prototype.hasOwnProperty.call(
        projectData,
        'startDate'
      )
    ) {
      updateData['Start Date'] =
        updatedStartDate || '';
    }

    if (
      Object.prototype.hasOwnProperty.call(
        projectData,
        'expectedEndDate'
      )
    ) {
      updateData['Expected End Date'] =
        updatedExpectedEndDate || '';
    }

    if (
      Object.prototype.hasOwnProperty.call(
        projectData,
        'status'
      )
    ) {
      updateData['Status'] = normalizeProjectStatus_(
        projectData.status
      );
    }

    if (
      Object.prototype.hasOwnProperty.call(
        projectData,
        'remarks'
      )
    ) {
      updateData['Remarks'] = normalizeText_(
        projectData.remarks
      );
    }

    updateObjectRow_(
      getProjectSheetName_(),
      project._rowNumber,
      updateData
    );

    addTransactionLog_(
      session.email,
      'PROJECT',
      'UPDATE',
      normalizeUpper_(projectId),
      updateData
    );

    return successResponse_(
      'Project updated successfully.',
      getProjectByIdData_(projectId)
    );
  }, 'Unable to update project.');
}

/**
 * Returns one project by Project ID.
 */
function getProjectById(projectId) {
  return safeExecute_(function () {
    requireAuthenticatedUser_();

    const project = getProjectRecordById_(projectId);

    if (!project) {
      throw new Error('Project not found: ' + projectId);
    }

    return successResponse_(
      'Project loaded successfully.',
      mapProjectRecord_(project)
    );
  }, 'Unable to load project.');
}

/**
 * Internal project lookup by ID.
 */
function getProjectRecordById_(projectId) {
  const normalizedId = normalizeUpper_(projectId);

  if (!normalizedId) {
    return null;
  }

  const projects = getSheetObjects_(
    getProjectSheetName_()
  );

  for (let index = 0; index < projects.length; index++) {
    if (
      normalizeUpper_(projects[index]['Project ID']) ===
      normalizedId
    ) {
      return projects[index];
    }
  }

  return null;
}

/**
 * Internal project lookup by project code.
 */
function getProjectRecordByCode_(projectCode) {
  const normalizedCode = normalizeUpper_(projectCode);

  if (!normalizedCode) {
    return null;
  }

  const projects = getSheetObjects_(
    getProjectSheetName_()
  );

  for (let index = 0; index < projects.length; index++) {
    if (
      normalizeUpper_(projects[index]['Project Code']) ===
      normalizedCode
    ) {
      return projects[index];
    }
  }

  return null;
}

/**
 * Returns mapped project data internally.
 */
function getProjectByIdData_(projectId) {
  const project = getProjectRecordById_(projectId);

  return project ? mapProjectRecord_(project) : null;
}

/**
 * Returns all projects with filters and pagination.
 *
 * filters:
 * {
 *   search: '',
 *   status: '',
 *   hodEmail: '',
 *   doerEmail: '',
 *   pageNumber: 1,
 *   pageSize: 20
 * }
 */
function getProjects(filters) {
  return safeExecute_(function () {
    const session = requireAuthenticatedUser_();

    filters = filters || {};

    const searchText = normalizeLower_(filters.search);
    const statusFilter = normalizeUpper_(filters.status);
    const hodFilter = normalizeLower_(filters.hodEmail);
    const doerFilter = normalizeLower_(filters.doerEmail);

    let projects = getSheetObjects_(
      getProjectSheetName_()
    );

    projects = projects.filter(function (project) {
      const status = normalizeUpper_(project['Status']);

      if (
        session.role === APP_CONFIG.USER_ROLES.DOER &&
        normalizeLower_(project['Doer Email']) !==
          session.email
      ) {
        return false;
      }

      if (
        statusFilter &&
        status !== statusFilter
      ) {
        return false;
      }

      if (
        hodFilter &&
        normalizeLower_(project['HOD Email']) !== hodFilter
      ) {
        return false;
      }

      if (
        doerFilter &&
        normalizeLower_(project['Doer Email']) !== doerFilter
      ) {
        return false;
      }

      if (searchText) {
        const searchableText = [
          project['Project ID'],
          project['Project Code'],
          project['Project Name'],
          project['Client Name'],
          project['Project Location'],
          project['Project Manager'],
          project['HOD Email'],
          project['Doer Email'],
          project['Status']
        ]
          .map(normalizeLower_)
          .join(' ');

        if (searchableText.indexOf(searchText) === -1) {
          return false;
        }
      }

      return true;
    });

    projects.sort(function (first, second) {
      const firstDate =
        parseDate_(first['Created At']) || new Date(0);
      const secondDate =
        parseDate_(second['Created At']) || new Date(0);

      return secondDate.getTime() - firstDate.getTime();
    });

    const mappedProjects = projects.map(
      mapProjectRecord_
    );

    const paginated = paginateRecords_(
      mappedProjects,
      filters.pageNumber || 1,
      filters.pageSize || 20
    );

    return successResponse_(
      'Projects loaded successfully.',
      paginated
    );
  }, 'Unable to load projects.');
}

/**
 * Returns active projects for dropdowns.
 */
function getActiveProjectDropdown() {
  return safeExecute_(function () {
    const session = requireAuthenticatedUser_();

    let projects = getSheetObjects_(
      getProjectSheetName_()
    );

    projects = projects.filter(function (project) {
      const status = normalizeUpper_(project['Status']);

      if (
        status !== PROJECT_SERVICE_CONFIG.STATUS.ACTIVE
      ) {
        return false;
      }

      if (
        session.role === APP_CONFIG.USER_ROLES.DOER &&
        normalizeLower_(project['Doer Email']) !==
          session.email
      ) {
        return false;
      }

      return true;
    });

    projects.sort(function (first, second) {
      return normalizeText_(first['Project Name'])
        .localeCompare(
          normalizeText_(second['Project Name'])
        );
    });

    const dropdown = projects.map(function (project) {
      return {
        value: normalizeUpper_(project['Project ID']),
        label:
          normalizeText_(project['Project Name']) +
          ' - ' +
          normalizeText_(project['Client Name']),
        projectId: normalizeUpper_(
          project['Project ID']
        ),
        projectCode: normalizeUpper_(
          project['Project Code']
        ),
        projectName: normalizeText_(
          project['Project Name']
        ),
        clientName: normalizeText_(
          project['Client Name']
        ),
        hodEmail: normalizeLower_(
          project['HOD Email']
        ),
        doerEmail: normalizeLower_(
          project['Doer Email']
        )
      };
    });

    return successResponse_(
      'Active project dropdown loaded successfully.',
      dropdown
    );
  }, 'Unable to load active project dropdown.');
}

/**
 * Returns active customized-inventory projects.
 */
function getCustomizeInventoryProjectDropdown() {
  return getActiveProjectDropdown();
}

/**
 * Assigns an HOD to a project.
 */
function assignProjectHod(projectId, hodEmail) {
  return safeExecute_(function () {
    const session = requireRole_(
      APP_CONFIG.USER_ROLES.ADMIN
    );

    const project = getProjectRecordById_(projectId);

    if (!project) {
      throw new Error('Project not found: ' + projectId);
    }

    const normalizedEmail = normalizeLower_(hodEmail);

    if (!normalizedEmail) {
      throw new Error('HOD email is required.');
    }

    validateProjectAssignedUser_(
      normalizedEmail,
      APP_CONFIG.USER_ROLES.HOD,
      'HOD'
    );

    updateObjectRow_(
      getProjectSheetName_(),
      project._rowNumber,
      {
        'HOD Email': normalizedEmail,
        'Updated By': session.email,
        'Updated At': new Date()
      }
    );

    addTransactionLog_(
      session.email,
      'PROJECT',
      'ASSIGN HOD',
      normalizeUpper_(projectId),
      {
        hodEmail: normalizedEmail
      }
    );

    return successResponse_(
      'HOD assigned successfully.',
      {
        projectId: normalizeUpper_(projectId),
        hodEmail: normalizedEmail
      }
    );
  }, 'Unable to assign project HOD.');
}

/**
 * Assigns a Doer to a project.
 */
function assignProjectDoer(projectId, doerEmail) {
  return safeExecute_(function () {
    const session = requireRole_([
      APP_CONFIG.USER_ROLES.ADMIN,
      APP_CONFIG.USER_ROLES.HOD
    ]);

    const project = getProjectRecordById_(projectId);

    if (!project) {
      throw new Error('Project not found: ' + projectId);
    }

    const normalizedEmail = normalizeLower_(doerEmail);

    if (!normalizedEmail) {
      throw new Error('Doer email is required.');
    }

    validateProjectAssignedUser_(
      normalizedEmail,
      APP_CONFIG.USER_ROLES.DOER,
      'Doer'
    );

    updateObjectRow_(
      getProjectSheetName_(),
      project._rowNumber,
      {
        'Doer Email': normalizedEmail,
        'Updated By': session.email,
        'Updated At': new Date()
      }
    );

    addTransactionLog_(
      session.email,
      'PROJECT',
      'ASSIGN DOER',
      normalizeUpper_(projectId),
      {
        doerEmail: normalizedEmail
      }
    );

    return successResponse_(
      'Doer assigned successfully.',
      {
        projectId: normalizeUpper_(projectId),
        doerEmail: normalizedEmail
      }
    );
  }, 'Unable to assign project Doer.');
}

/**
 * Changes project status.
 */
function changeProjectStatus(
  projectId,
  newStatus,
  remarks
) {
  return safeExecute_(function () {
    const session = requireRole_([
      APP_CONFIG.USER_ROLES.ADMIN,
      APP_CONFIG.USER_ROLES.HOD
    ]);

    const project = getProjectRecordById_(projectId);

    if (!project) {
      throw new Error('Project not found: ' + projectId);
    }

    const oldStatus = normalizeProjectStatus_(
      project['Status']
    );

    const normalizedNewStatus =
      normalizeProjectStatus_(newStatus);

    if (oldStatus === normalizedNewStatus) {
      throw new Error(
        'Project is already in ' +
        normalizedNewStatus +
        ' status.'
      );
    }

    if (
      normalizedNewStatus ===
        PROJECT_SERVICE_CONFIG.STATUS.CANCELLED &&
      session.role !== APP_CONFIG.USER_ROLES.ADMIN
    ) {
      throw new Error(
        'Only ADMIN can cancel a project.'
      );
    }

    updateObjectRow_(
      getProjectSheetName_(),
      project._rowNumber,
      {
        'Status': normalizedNewStatus,
        'Remarks': normalizeText_(
          remarks || project['Remarks']
        ),
        'Updated By': session.email,
        'Updated At': new Date()
      }
    );

    addTransactionLog_(
      session.email,
      'PROJECT',
      'STATUS CHANGE',
      normalizeUpper_(projectId),
      {
        oldStatus: oldStatus,
        newStatus: normalizedNewStatus,
        remarks: normalizeText_(remarks)
      }
    );

    return successResponse_(
      'Project status changed successfully.',
      {
        projectId: normalizeUpper_(projectId),
        oldStatus: oldStatus,
        newStatus: normalizedNewStatus
      }
    );
  }, 'Unable to change project status.');
}

/**
 * Closes a project.
 */
function closeProject(projectId, remarks) {
  return changeProjectStatus(
    projectId,
    PROJECT_SERVICE_CONFIG.STATUS.CLOSED,
    remarks
  );
}

/**
 * Reopens a project.
 */
function reopenProject(projectId, remarks) {
  return changeProjectStatus(
    projectId,
    PROJECT_SERVICE_CONFIG.STATUS.ACTIVE,
    remarks
  );
}

/**
 * Places a project on hold.
 */
function holdProject(projectId, remarks) {
  return changeProjectStatus(
    projectId,
    PROJECT_SERVICE_CONFIG.STATUS.ON_HOLD,
    remarks
  );
}

/**
 * Marks a project completed.
 */
function completeProject(projectId, remarks) {
  return changeProjectStatus(
    projectId,
    PROJECT_SERVICE_CONFIG.STATUS.COMPLETED,
    remarks
  );
}

/**
 * Soft deletes a project by setting its status to INACTIVE.
 */
function deleteProject(projectId, remarks) {
  return safeExecute_(function () {
    const session = requireRole_(
      APP_CONFIG.USER_ROLES.ADMIN
    );

    const project = getProjectRecordById_(projectId);

    if (!project) {
      throw new Error('Project not found: ' + projectId);
    }

    if (projectHasInventoryTransactions_(projectId)) {
      throw new Error(
        'This project contains inventory transactions and cannot be deleted. Close the project instead.'
      );
    }

    updateObjectRow_(
      getProjectSheetName_(),
      project._rowNumber,
      {
        'Status':
          PROJECT_SERVICE_CONFIG.STATUS.INACTIVE,
        'Remarks': normalizeText_(
          remarks || 'Project soft deleted.'
        ),
        'Updated By': session.email,
        'Updated At': new Date()
      }
    );

    addTransactionLog_(
      session.email,
      'PROJECT',
      'SOFT DELETE',
      normalizeUpper_(projectId),
      {
        previousStatus: normalizeUpper_(
          project['Status']
        ),
        remarks: normalizeText_(remarks)
      }
    );

    return successResponse_(
      'Project deleted successfully.',
      {
        projectId: normalizeUpper_(projectId),
        status:
          PROJECT_SERVICE_CONFIG.STATUS.INACTIVE
      }
    );
  }, 'Unable to delete project.');
}

/**
 * Checks whether a project has any inventory transaction.
 */
function projectHasInventoryTransactions_(projectId) {
  const normalizedProjectId = normalizeUpper_(projectId);

  const sheetNames = [
    APP_CONFIG.SHEETS.CUSTOMIZE_INVENTORY,
    APP_CONFIG.SHEETS.INWARD,
    APP_CONFIG.SHEETS.OUTWARD,
    APP_CONFIG.SHEETS.TRANSFER_REQUEST
  ].filter(function (sheetName) {
    return Boolean(sheetName);
  });

  for (
    let sheetIndex = 0;
    sheetIndex < sheetNames.length;
    sheetIndex++
  ) {
    const sheetName = sheetNames[sheetIndex];

    let sheet;

    try {
      sheet = getSystemSheet(sheetName);
    } catch (error) {
      continue;
    }

    const headerMap = getHeaderMap_(sheet);
    const projectColumn =
      headerMap['Project ID'] ||
      headerMap['From Project ID'] ||
      headerMap['To Project ID'];

    if (!projectColumn || sheet.getLastRow() < 2) {
      continue;
    }

    const values = sheet
      .getRange(
        2,
        projectColumn,
        sheet.getLastRow() - 1,
        1
      )
      .getDisplayValues()
      .flat();

    for (
      let rowIndex = 0;
      rowIndex < values.length;
      rowIndex++
    ) {
      if (
        normalizeUpper_(values[rowIndex]) ===
        normalizedProjectId
      ) {
        return true;
      }
    }

    const fromProjectColumn =
      headerMap['From Project ID'];
    const toProjectColumn =
      headerMap['To Project ID'];

    const additionalColumns = [
      fromProjectColumn,
      toProjectColumn
    ].filter(function (columnNumber) {
      return Boolean(columnNumber);
    });

    for (
      let columnIndex = 0;
      columnIndex < additionalColumns.length;
      columnIndex++
    ) {
      const additionalValues = sheet
        .getRange(
          2,
          additionalColumns[columnIndex],
          sheet.getLastRow() - 1,
          1
        )
        .getDisplayValues()
        .flat();

      if (
        additionalValues.some(function (value) {
          return (
            normalizeUpper_(value) ===
            normalizedProjectId
          );
        })
      ) {
        return true;
      }
    }
  }

  return false;
}

/**
 * Validates an assigned HOD or Doer user.
 */
function validateProjectAssignedUser_(
  email,
  expectedRole,
  label
) {
  if (!isValidEmail_(email)) {
    throw new Error(
      'Please enter a valid ' + label + ' email address.'
    );
  }

  const user = getUserByEmail_(email);

  if (!user) {
    throw new Error(
      label +
      ' is not registered in the Users sheet: ' +
      email
    );
  }

  if (
    normalizeUpper_(user['Status']) !==
    APP_CONFIG.USER_STATUS.ACTIVE
  ) {
    throw new Error(
      label + ' user account is inactive: ' + email
    );
  }

  if (
    normalizeUpper_(user['Role']) !==
    normalizeUpper_(expectedRole)
  ) {
    throw new Error(
      email +
      ' is not assigned the ' +
      expectedRole +
      ' role.'
    );
  }

  return true;
}

/**
 * Validates start and expected completion dates.
 */
function validateProjectDates_(
  startDate,
  expectedEndDate
) {
  if (
    startDate &&
    expectedEndDate &&
    expectedEndDate.getTime() < startDate.getTime()
  ) {
    throw new Error(
      'Expected end date cannot be earlier than start date.'
    );
  }

  return true;
}

/**
 * Validates and normalizes project status.
 */
function normalizeProjectStatus_(status) {
  const normalizedStatus = normalizeUpper_(status);

  const validStatuses = Object.keys(
    PROJECT_SERVICE_CONFIG.STATUS
  ).map(function (key) {
    return PROJECT_SERVICE_CONFIG.STATUS[key];
  });

  if (
    validStatuses.indexOf(normalizedStatus) === -1
  ) {
    throw new Error(
      'Invalid project status: ' + normalizedStatus
    );
  }

  return normalizedStatus;
}

/**
 * Checks duplicate project name and client combination.
 */
function isDuplicateProjectName_(
  projectName,
  clientName,
  excludeProjectId
) {
  const normalizedProjectName =
    normalizeLower_(projectName);
  const normalizedClientName =
    normalizeLower_(clientName);
  const normalizedExcludeId =
    normalizeUpper_(excludeProjectId);

  const projects = getSheetObjects_(
    getProjectSheetName_()
  );

  return projects.some(function (project) {
    const projectId = normalizeUpper_(
      project['Project ID']
    );

    if (
      normalizedExcludeId &&
      projectId === normalizedExcludeId
    ) {
      return false;
    }

    const status = normalizeUpper_(project['Status']);

    if (
      status === PROJECT_SERVICE_CONFIG.STATUS.INACTIVE
    ) {
      return false;
    }

    return (
      normalizeLower_(project['Project Name']) ===
        normalizedProjectName &&
      normalizeLower_(project['Client Name']) ===
        normalizedClientName
    );
  });
}

/**
 * Generates a readable project code.
 *
 * Example:
 * Mumbai Office Automation + PRJ000001
 * becomes MOA-000001
 */
function generateProjectCode_(
  projectName,
  projectId
) {
  const words = normalizeUpper_(projectName)
    .replace(/[^A-Z0-9 ]/g, ' ')
    .split(/\s+/)
    .filter(function (word) {
      return word.length > 0;
    });

  let shortCode = '';

  if (words.length === 1) {
    shortCode = words[0].substring(0, 3);
  } else {
    shortCode = words
      .slice(0, 4)
      .map(function (word) {
        return word.charAt(0);
      })
      .join('');
  }

  if (!shortCode) {
    shortCode =
      PROJECT_SERVICE_CONFIG.PROJECT_CODE_PREFIX;
  }

  const numericPart =
    normalizeText_(projectId).match(/\d+$/);

  return (
    shortCode +
    '-' +
    (numericPart
      ? numericPart[0].padStart(6, '0')
      : String(new Date().getTime()).slice(-6))
  );
}

/**
 * Maps a raw project sheet record.
 */
function mapProjectRecord_(project) {
  return {
    projectId: normalizeUpper_(
      project['Project ID']
    ),
    projectCode: normalizeUpper_(
      project['Project Code']
    ),
    projectName: normalizeText_(
      project['Project Name']
    ),
    clientName: normalizeText_(
      project['Client Name']
    ),
    projectLocation: normalizeText_(
      project['Project Location']
    ),
    projectManager: normalizeText_(
      project['Project Manager']
    ),
    hodEmail: normalizeLower_(
      project['HOD Email']
    ),
    doerEmail: normalizeLower_(
      project['Doer Email']
    ),
    startDate: formatDate_(
      project['Start Date']
    ),
    expectedEndDate: formatDate_(
      project['Expected End Date']
    ),
    status: normalizeUpper_(
      project['Status']
    ),
    remarks: normalizeText_(
      project['Remarks']
    ),
    createdBy: normalizeLower_(
      project['Created By']
    ),
    createdAt: formatDateTime_(
      project['Created At']
    ),
    updatedBy: normalizeLower_(
      project['Updated By']
    ),
    updatedAt: formatDateTime_(
      project['Updated At']
    )
  };
}

/**
 * Returns project summary counts.
 */
function getProjectSummary() {
  return safeExecute_(function () {
    requireAuthenticatedUser_();

    const projects = getSheetObjects_(
      getProjectSheetName_()
    );

    const summary = {
      total: projects.length,
      active: 0,
      onHold: 0,
      completed: 0,
      closed: 0,
      cancelled: 0,
      inactive: 0
    };

    projects.forEach(function (project) {
      const status = normalizeUpper_(
        project['Status']
      );

      switch (status) {
        case PROJECT_SERVICE_CONFIG.STATUS.ACTIVE:
          summary.active++;
          break;

        case PROJECT_SERVICE_CONFIG.STATUS.ON_HOLD:
          summary.onHold++;
          break;

        case PROJECT_SERVICE_CONFIG.STATUS.COMPLETED:
          summary.completed++;
          break;

        case PROJECT_SERVICE_CONFIG.STATUS.CLOSED:
          summary.closed++;
          break;

        case PROJECT_SERVICE_CONFIG.STATUS.CANCELLED:
          summary.cancelled++;
          break;

        case PROJECT_SERVICE_CONFIG.STATUS.INACTIVE:
          summary.inactive++;
          break;
      }
    });

    return successResponse_(
      'Project summary loaded successfully.',
      summary
    );
  }, 'Unable to load project summary.');
}

/**
 * Creates one test project.
 *
 * This test uses the current ADMIN user.
 * HOD and Doer are left blank.
 */
function testCreateProject() {
  const timestamp = Utilities.formatDate(
    new Date(),
    APP_CONFIG.TIME_ZONE,
    'yyyyMMdd-HHmmss'
  );

  const result = createProject({
    projectName: 'Test Inventory Project ' + timestamp,
    clientName: 'Test Client',
    projectLocation: 'Mumbai',
    projectManager: 'System Administrator',
    hodEmail: '',
    doerEmail: '',
    startDate: new Date(),
    expectedEndDate: '',
    status: 'ACTIVE',
    remarks: 'Created from ProjectService.gs test.'
  });

  Logger.log(JSON.stringify(result, null, 2));

  return result;
}

/**
 * Tests project listing and project summary.
 */
function testProjectService() {
  const result = {
    projects: getProjects({
      pageNumber: 1,
      pageSize: 10
    }),
    dropdown: getActiveProjectDropdown(),
    summary: getProjectSummary()
  };

  Logger.log(JSON.stringify(result, null, 2));

  return result;
}
