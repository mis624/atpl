/**
 * PROJECT-BASED INVENTORY MANAGEMENT SYSTEM
 * File: Auth.gs
 *
 * Handles user authentication, authorization and role permissions.
 */

/**
 * Returns the current signed-in user's complete access profile.
 */
function getCurrentUserProfile() {
  return safeExecute_(function () {
    const email = getCurrentUserEmail_();

    if (!email) {
      throw new Error(
        'Unable to identify your Google account. Please open the application using an authorized Google account.'
      );
    }

    const user = getUserByEmail_(email);

    if (!user) {
      throw new Error(
        'Your account is not registered in the Users sheet: ' + email
      );
    }

    const status = normalizeUpper_(user['Status']);

    if (status !== APP_CONFIG.USER_STATUS.ACTIVE) {
      throw new Error('Your account is not active.');
    }

    const roleName = normalizeUpper_(user['Role']);
    const rolePermissions = getRolePermissions_(roleName);

    if (!rolePermissions) {
      throw new Error(
        'Invalid or inactive role assigned to your account: ' + roleName
      );
    }

    return successResponse_(
      'User authenticated successfully.',
      buildUserSession_(user, rolePermissions)
    );
  }, 'Authentication failed.');
}

/**
 * Returns the current user's session data.
 * This function is intended for frontend calls.
 */
function getSessionUser() {
  return getCurrentUserProfile();
}

/**
 * Returns a registered user by email.
 */
function getUserByEmail_(email) {
  const normalizedEmail = normalizeLower_(email);

  if (!normalizedEmail) {
    return null;
  }

  const users = getSheetObjects_(APP_CONFIG.SHEETS.USERS);

  for (let index = 0; index < users.length; index++) {
    const userEmail = normalizeLower_(users[index]['Email']);

    if (userEmail === normalizedEmail) {
      return users[index];
    }
  }

  return null;
}

/**
 * Returns a registered user by User ID.
 */
function getUserById_(userId) {
  const normalizedUserId = normalizeUpper_(userId);

  if (!normalizedUserId) {
    return null;
  }

  const users = getSheetObjects_(APP_CONFIG.SHEETS.USERS);

  for (let index = 0; index < users.length; index++) {
    const currentUserId = normalizeUpper_(users[index]['User ID']);

    if (currentUserId === normalizedUserId) {
      return users[index];
    }
  }

  return null;
}

/**
 * Returns role permission data.
 */
function getRolePermissions_(roleName) {
  const normalizedRole = normalizeUpper_(roleName);

  if (!normalizedRole) {
    return null;
  }

  const roles = getSheetObjects_(APP_CONFIG.SHEETS.ROLES);

  for (let index = 0; index < roles.length; index++) {
    const currentRole = normalizeUpper_(roles[index]['Role Name']);

    if (currentRole === normalizedRole) {
      return {
        roleId: normalizeUpper_(roles[index]['Role ID']),
        roleName: currentRole,
        canInward: toBoolean_(roles[index]['Can Inward']),
        canOutward: toBoolean_(roles[index]['Can Outward']),
        canTransfer: toBoolean_(roles[index]['Can Transfer']),
        canApprove: toBoolean_(roles[index]['Can Approve']),
        canReports: toBoolean_(roles[index]['Can Reports'])
      };
    }
  }

  return null;
}

/**
 * Builds a safe session object for the frontend.
 */
function buildUserSession_(user, permissions) {
  return {
    userId: normalizeUpper_(user['User ID']),
    employeeName: normalizeText_(user['Employee Name']),
    email: normalizeLower_(user['Email']),
    mobile: normalizeText_(user['Mobile']),
    department: normalizeText_(user['Department']),
    role: normalizeUpper_(user['Role']),
    status: normalizeUpper_(user['Status']),
    permissions: {
      canInward: Boolean(permissions.canInward),
      canOutward: Boolean(permissions.canOutward),
      canTransfer: Boolean(permissions.canTransfer),
      canApprove: Boolean(permissions.canApprove),
      canReports: Boolean(permissions.canReports),
      isAdmin:
        normalizeUpper_(user['Role']) === APP_CONFIG.USER_ROLES.ADMIN,
      isHod:
        normalizeUpper_(user['Role']) === APP_CONFIG.USER_ROLES.HOD,
      isDoer:
        normalizeUpper_(user['Role']) === APP_CONFIG.USER_ROLES.DOER,
      isViewer:
        normalizeUpper_(user['Role']) === APP_CONFIG.USER_ROLES.VIEWER
    }
  };
}

/**
 * Requires the current user to be registered and active.
 * Returns the authenticated session object.
 */
function requireAuthenticatedUser_() {
  const email = getCurrentUserEmail_();

  if (!email) {
    throw new Error(
      'Authentication failed. Signed-in Google account could not be identified.'
    );
  }

  const user = getUserByEmail_(email);

  if (!user) {
    throw new Error(
      'Access denied. Your account is not registered: ' + email
    );
  }

  if (
    normalizeUpper_(user['Status']) !== APP_CONFIG.USER_STATUS.ACTIVE
  ) {
    throw new Error('Access denied. Your account is inactive.');
  }

  const roleName = normalizeUpper_(user['Role']);
  const permissions = getRolePermissions_(roleName);

  if (!permissions) {
    throw new Error(
      'Access denied. No valid permissions were found for role: ' +
      roleName
    );
  }

  return buildUserSession_(user, permissions);
}

/**
 * Requires the current user to have one of the given roles.
 *
 * Example:
 * requireRole_(['ADMIN', 'HOD'])
 */
function requireRole_(allowedRoles) {
  const session = requireAuthenticatedUser_();

  const roles = Array.isArray(allowedRoles)
    ? allowedRoles
    : [allowedRoles];

  const normalizedRoles = roles.map(function (role) {
    return normalizeUpper_(role);
  });

  if (normalizedRoles.indexOf(session.role) === -1) {
    throw new Error(
      'Access denied. Required role: ' +
      normalizedRoles.join(' or ')
    );
  }

  return session;
}

/**
 * Requires a specific permission.
 *
 * Supported permissions:
 * canInward
 * canOutward
 * canTransfer
 * canApprove
 * canReports
 */
function requirePermission_(permissionName) {
  const session = requireAuthenticatedUser_();

  if (
    !Object.prototype.hasOwnProperty.call(
      session.permissions,
      permissionName
    )
  ) {
    throw new Error(
      'Invalid permission requested: ' + permissionName
    );
  }

  if (!session.permissions[permissionName]) {
    throw new Error(
      'Access denied. You do not have permission for this action.'
    );
  }

  return session;
}

/**
 * Checks whether the current user has a given role.
 */
function currentUserHasRole_(roleName) {
  try {
    const session = requireAuthenticatedUser_();

    return session.role === normalizeUpper_(roleName);
  } catch (error) {
    return false;
  }
}

/**
 * Checks whether the current user has a given permission.
 */
function currentUserHasPermission_(permissionName) {
  try {
    const session = requireAuthenticatedUser_();

    return Boolean(session.permissions[permissionName]);
  } catch (error) {
    return false;
  }
}

/**
 * Returns true when the current user is ADMIN.
 */
function isCurrentUserAdmin_() {
  return currentUserHasRole_(APP_CONFIG.USER_ROLES.ADMIN);
}

/**
 * Returns true when the current user is HOD.
 */
function isCurrentUserHod_() {
  return currentUserHasRole_(APP_CONFIG.USER_ROLES.HOD);
}

/**
 * Returns true when the current user is DOER.
 */
function isCurrentUserDoer_() {
  return currentUserHasRole_(APP_CONFIG.USER_ROLES.DOER);
}

/**
 * Returns true when the current user is VIEWER.
 */
function isCurrentUserViewer_() {
  return currentUserHasRole_(APP_CONFIG.USER_ROLES.VIEWER);
}

/**
 * Creates the first ADMIN user.
 *
 * Run this function only once before testing Auth.gs.
 * It uses the currently signed-in Google account.
 */
function createFirstAdminUser() {
  return safeExecute_(function () {
    const email = getCurrentUserEmail_();

    if (!email) {
      throw new Error(
        'Signed-in Google account could not be identified.'
      );
    }

    const existingUser = getUserByEmail_(email);

    if (existingUser) {
      const rowNumber = existingUser._rowNumber;

      updateObjectRow_(
        APP_CONFIG.SHEETS.USERS,
        rowNumber,
        {
          'Role': APP_CONFIG.USER_ROLES.ADMIN,
          'Status': APP_CONFIG.USER_STATUS.ACTIVE,
          'Updated At': new Date()
        }
      );

      return successResponse_(
        'Existing user updated as ADMIN.',
        {
          userId: normalizeUpper_(existingUser['User ID']),
          email: email,
          role: APP_CONFIG.USER_ROLES.ADMIN,
          status: APP_CONFIG.USER_STATUS.ACTIVE
        }
      );
    }

    const userId = generateNextId_(
      'USER',
      APP_CONFIG.ID_PREFIX.USER
    );

    const now = new Date();

    appendObjectRow_(
      APP_CONFIG.SHEETS.USERS,
      {
        'User ID': userId,
        'Employee Name': 'System Administrator',
        'Email': email,
        'Mobile': '',
        'Department': 'ADMINISTRATION',
        'Role': APP_CONFIG.USER_ROLES.ADMIN,
        'Status': APP_CONFIG.USER_STATUS.ACTIVE,
        'Created At': now,
        'Updated At': now
      }
    );

    addTransactionLog_(
      email,
      'AUTH',
      'FIRST ADMIN CREATED',
      userId,
      {
        email: email,
        role: APP_CONFIG.USER_ROLES.ADMIN
      }
    );

    return successResponse_(
      'First ADMIN user created successfully.',
      {
        userId: userId,
        email: email,
        role: APP_CONFIG.USER_ROLES.ADMIN,
        status: APP_CONFIG.USER_STATUS.ACTIVE
      }
    );
  }, 'Unable to create the first ADMIN user.');
}

/**
 * Creates a new system user.
 * Only ADMIN can execute this function.
 */
function createUser(userData) {
  return safeExecute_(function () {
    const adminSession = requireRole_(
      APP_CONFIG.USER_ROLES.ADMIN
    );

    validateRequiredFields_(
      userData,
      [
        'employeeName',
        'email',
        'role'
      ]
    );

    const employeeName = normalizeText_(userData.employeeName);
    const email = normalizeLower_(userData.email);
    const mobile = normalizeText_(userData.mobile);
    const department = normalizeText_(userData.department);
    const role = normalizeUpper_(userData.role);
    const status = normalizeUpper_(
      userData.status || APP_CONFIG.USER_STATUS.ACTIVE
    );

    if (!isValidEmail_(email)) {
      throw new Error('Please enter a valid email address.');
    }

    if (mobile && !isValidMobile_(mobile)) {
      throw new Error(
        'Please enter a valid 10-digit Indian mobile number.'
      );
    }

    if (
      Object.values(APP_CONFIG.USER_ROLES).indexOf(role) === -1
    ) {
      throw new Error('Invalid user role: ' + role);
    }

    if (
      [
        APP_CONFIG.USER_STATUS.ACTIVE,
        APP_CONFIG.USER_STATUS.INACTIVE
      ].indexOf(status) === -1
    ) {
      throw new Error('Invalid user status: ' + status);
    }

    if (getUserByEmail_(email)) {
      throw new Error(
        'A user with this email already exists: ' + email
      );
    }

    const rolePermissions = getRolePermissions_(role);

    if (!rolePermissions) {
      throw new Error(
        'Role is not available in the Roles sheet: ' + role
      );
    }

    const userId = generateNextId_(
      'USER',
      APP_CONFIG.ID_PREFIX.USER
    );

    const now = new Date();

    appendObjectRow_(
      APP_CONFIG.SHEETS.USERS,
      {
        'User ID': userId,
        'Employee Name': employeeName,
        'Email': email,
        'Mobile': mobile,
        'Department': department,
        'Role': role,
        'Status': status,
        'Created At': now,
        'Updated At': now
      }
    );

    addTransactionLog_(
      adminSession.email,
      'USER',
      'CREATE',
      userId,
      {
        employeeName: employeeName,
        email: email,
        role: role,
        status: status
      }
    );

    return successResponse_(
      'User created successfully.',
      {
        userId: userId,
        employeeName: employeeName,
        email: email,
        role: role,
        status: status
      }
    );
  }, 'Unable to create user.');
}

/**
 * Updates an existing user.
 * Only ADMIN can execute this function.
 */
function updateUser(userId, userData) {
  return safeExecute_(function () {
    const adminSession = requireRole_(
      APP_CONFIG.USER_ROLES.ADMIN
    );

    const user = getUserById_(userId);

    if (!user) {
      throw new Error('User not found: ' + userId);
    }

    const updateData = {
      'Updated At': new Date()
    };

    if (
      Object.prototype.hasOwnProperty.call(
        userData,
        'employeeName'
      )
    ) {
      const employeeName = normalizeText_(userData.employeeName);

      if (!employeeName) {
        throw new Error('Employee name cannot be blank.');
      }

      updateData['Employee Name'] = employeeName;
    }

    if (
      Object.prototype.hasOwnProperty.call(userData, 'email')
    ) {
      const email = normalizeLower_(userData.email);

      if (!isValidEmail_(email)) {
        throw new Error('Please enter a valid email address.');
      }

      const existingEmailUser = getUserByEmail_(email);

      if (
        existingEmailUser &&
        normalizeUpper_(existingEmailUser['User ID']) !==
          normalizeUpper_(userId)
      ) {
        throw new Error(
          'Another user already uses this email address.'
        );
      }

      updateData['Email'] = email;
    }

    if (
      Object.prototype.hasOwnProperty.call(userData, 'mobile')
    ) {
      const mobile = normalizeText_(userData.mobile);

      if (mobile && !isValidMobile_(mobile)) {
        throw new Error(
          'Please enter a valid 10-digit Indian mobile number.'
        );
      }

      updateData['Mobile'] = mobile;
    }

    if (
      Object.prototype.hasOwnProperty.call(
        userData,
        'department'
      )
    ) {
      updateData['Department'] =
        normalizeText_(userData.department);
    }

    if (
      Object.prototype.hasOwnProperty.call(userData, 'role')
    ) {
      const role = normalizeUpper_(userData.role);

      if (
        Object.values(APP_CONFIG.USER_ROLES).indexOf(role) === -1
      ) {
        throw new Error('Invalid user role: ' + role);
      }

      if (!getRolePermissions_(role)) {
        throw new Error(
          'Role is not available in the Roles sheet: ' + role
        );
      }

      updateData['Role'] = role;
    }

    if (
      Object.prototype.hasOwnProperty.call(userData, 'status')
    ) {
      const status = normalizeUpper_(userData.status);

      if (
        [
          APP_CONFIG.USER_STATUS.ACTIVE,
          APP_CONFIG.USER_STATUS.INACTIVE
        ].indexOf(status) === -1
      ) {
        throw new Error('Invalid user status: ' + status);
      }

      updateData['Status'] = status;
    }

    updateObjectRow_(
      APP_CONFIG.SHEETS.USERS,
      user._rowNumber,
      updateData
    );

    addTransactionLog_(
      adminSession.email,
      'USER',
      'UPDATE',
      normalizeUpper_(userId),
      updateData
    );

    return successResponse_(
      'User updated successfully.',
      {
        userId: normalizeUpper_(userId)
      }
    );
  }, 'Unable to update user.');
}

/**
 * Activates or deactivates a user.
 * Only ADMIN can execute this function.
 */
function setUserStatus(userId, status) {
  return updateUser(userId, {
    status: status
  });
}

/**
 * Returns all system users.
 * Only ADMIN and HOD can access the user list.
 */
function getUsers() {
  return safeExecute_(function () {
    requireRole_([
      APP_CONFIG.USER_ROLES.ADMIN,
      APP_CONFIG.USER_ROLES.HOD
    ]);

    const users = getSheetObjects_(APP_CONFIG.SHEETS.USERS)
      .map(function (user) {
        return {
          userId: normalizeUpper_(user['User ID']),
          employeeName: normalizeText_(user['Employee Name']),
          email: normalizeLower_(user['Email']),
          mobile: normalizeText_(user['Mobile']),
          department: normalizeText_(user['Department']),
          role: normalizeUpper_(user['Role']),
          status: normalizeUpper_(user['Status']),
          createdAt: formatDateTime_(user['Created At']),
          updatedAt: formatDateTime_(user['Updated At'])
        };
      });

    return successResponse_(
      'Users loaded successfully.',
      users
    );
  }, 'Unable to load users.');
}

/**
 * Returns active HOD users.
 * This will later be used for approval email selection.
 */
function getActiveHodUsers_() {
  return getSheetObjects_(APP_CONFIG.SHEETS.USERS)
    .filter(function (user) {
      return (
        normalizeUpper_(user['Role']) ===
          APP_CONFIG.USER_ROLES.HOD &&
        normalizeUpper_(user['Status']) ===
          APP_CONFIG.USER_STATUS.ACTIVE &&
        isValidEmail_(user['Email'])
      );
    })
    .map(function (user) {
      return {
        userId: normalizeUpper_(user['User ID']),
        employeeName: normalizeText_(user['Employee Name']),
        email: normalizeLower_(user['Email']),
        department: normalizeText_(user['Department'])
      };
    });
}

/**
 * Tests Auth.gs.
 *
 * First run createFirstAdminUser().
 * Then run testAuth().
 */
function testAuth() {
  const result = getCurrentUserProfile();

  Logger.log(JSON.stringify(result, null, 2));

  return result;
}
