// Code.gs
/**
 * NepHR Management System - Google App Script Backend
 * Author: Your Name/Company
 * Version: 2.0 (Full HRMS)
 * Purpose: Handles web app serving, authentication, and data interactions with Google Sheets.
 */

// --- Global Constants & Configuration ---
const SPREADSHEET_ID = SpreadsheetApp.getActiveSpreadsheet().getId(); // Make sure this script is bound to a spreadsheet
const OTP_LENGTH = 6; // Length of the One-Time Password (e.g., 6 digits)
const OTP_EXPIRY_SECONDS = 300; // OTP is valid for 5 minutes (300 seconds)
const SESSION_EXPIRY_MINUTES = 60; // Sessions expire after 60 minutes

// Define roles for authorization
const USER_ROLES = {
  EMPLOYEE: 'Employee',
  ADMIN: 'Admin',
  HR: 'HR', // Example additional role
  MANAGER: 'Manager' // Example additional role
};

// --- Sheet Names and Column Mappings (ADJUST THESE TO YOUR SPREADSHEET LAYOUT) ---
// IMPORTANT: Ensure these sheet names match EXACTLY (case-sensitive) with your Google Sheet tab names.
// IMPORTANT: Ensure column indices match EXACTLY (0-indexed) with your Google Sheet column order.
const SHEET_NAMES = {
  USERS: 'Users', // Sheet for user credentials and info
  ACTIVE_SESSIONS: 'ActiveSessions', // Sheet for live user sessions
  APP_CONFIG: 'AppConfig', // Sheet for dynamic app settings (banner, accent color)
  BANNERS: 'Banners', // Sheet for tracking uploaded banners

  // --- HRMS SHEETS ---
  EMPLOYEES: 'Employees', // Main employee directory (for HR/Admin Panel)
  ATTENDANCE: 'Attendance', // Daily check-ins/outs
  TASKS: 'Tasks', // Task assignments and progress
  LEAVES: 'Leaves', // Leave applications and status
  ANNOUNCEMENTS: 'Announcements', // Company-wide announcements
  POLICIES: 'Policies', // Company Policy documents (metadata)
  PAYSLIPS: 'Payslips' // Placeholder for payslip links/metadata if not storing directly
};

// Users Sheet Column Mapping (0-indexed for A, B, C...)
// Headers: USERNAME | PASSWORD_HASH | ROLE | EMPLOYEE_ID | EMAIL | OTP | OTP_EXPIRY | PROFILE_PIC_URL
const USER_COLS = {
  USERNAME: 0,  // Column A: Unique identifier for the user (e.g., username or email)
  PASSWORD_HASH: 1, // Column B: Hashed password
  ROLE: 2,    // Column C: User's role (e.g., Admin, Employee)
  EMPLOYEE_ID: 3,  // Column D: Employee ID
  EMAIL: 4,   // Column E: User's email address (for OTP delivery)
  OTP: 5,    // Column F: Stores the current OTP
  OTP_EXPIRY: 6,    // Column G: Stores the OTP expiry timestamp (milliseconds since epoch)
  PROFILE_PIC_URL: 7 // Column H: User's profile picture URL
};

// ActiveSessions Sheet Column Mapping
// Headers: SESSION_ID | USERNAME | LOGIN_TIME | EXPIRY_TIME | USER_ROLE
const SESSION_COLS = {
  SESSION_ID: 0,  // Column A: Unique session identifier
  USERNAME: 1,  // Column B: Username associated with the session
  LOGIN_TIME: 2,  // Column C: Timestamp of login
  EXPIRY_TIME: 3,  // Column D: Timestamp when session expires
  USER_ROLE: 4    // Column E: User's role for quick access (denormalized from Users sheet)
};

// AppConfig Sheet Column Mapping
// Headers: Key | Value
const APP_CONFIG_COLS = {
  KEY: 0,    // Column A: Configuration key (e.g., "bannerUrl", "accentColor")
  VALUE: 1    // Column B: Configuration value
};

// Banners Sheet Column Mapping
// Headers: FileName | FileURL | FileID | UploadedAt | Uploader
const BANNER_COLS = {
  FILE_NAME: 0,
  FILE_URL: 1,
  FILE_ID: 2,
  UPLOADED_AT: 3,
  UPLOADER: 4
};

// Employees Sheet Column Mapping (For Admin to manage employee details)
// Headers: EmployeeID | Username | FullName | Email | PhoneNumber | Department | Role | JoinDate | Status | ProfilePicURL
const EMPLOYEE_COLS = {
  EMPLOYEE_ID: 0,
  USERNAME: 1,
  FULL_NAME: 2,
  EMAIL: 3,
  PHONE_NUMBER: 4,
  DEPARTMENT: 5,
  ROLE: 6, // Redundant with Users sheet, but useful for admin view
  JOIN_DATE: 7,
  STATUS: 8, // Active/Inactive
  PUBLIC_PROFILE_PIC_URL: 9 // Public URL for profile picture
};

// Attendance Sheet Column Mapping
// Headers: AttendanceID | EmployeeID | Username | Date | Time | Type (Check-in/out) | GeoLocation | Status (On-time, Late, Absent)
const ATTENDANCE_COLS = {
  ATTENDANCE_ID: 0,
  EMPLOYEE_ID: 1,
  USERNAME: 2,
  DATE: 3,
  TIME: 4,
  TYPE: 5, // 'Check-in', 'Check-out'
  GEO_LOCATION: 6, // "lat,long"
  STATUS: 7 // 'On-time', 'Late', 'Absent' (for full day)
};

// Tasks Sheet Column Mapping
// Headers: TaskID | AssignedToUsername | AssignedByUsername | Title | Description | Type (Daily, Monthly, Emergency) | Priority | DueDate | CreatedAt | Status (Pending, In Progress, Completed, Approved) | Progress (0-100%) | AttachmentsURLs | Comments | LastUpdated
const TASK_COLS = {
  TASK_ID: 0,
  ASSIGNED_TO_USERNAME: 1,
  ASSIGNED_BY_USERNAME: 2,
  TITLE: 3,
  DESCRIPTION: 4,
  TYPE: 5, // 'Daily', 'Monthly', 'Emergency'
  PRIORITY: 6, // 'Low', 'Medium', 'High', 'Critical'
  DUE_DATE: 7,
  CREATED_AT: 8,
  STATUS: 9, // 'Pending', 'In Progress', 'Submitted', 'Approved', 'Rejected'
  PROGRESS: 10, // Numeric %
  ATTACHMENTS_URLS: 11, // Comma-separated URLs
  COMMENTS: 12, // Last comment
  LAST_UPDATED: 13
};

// Leaves Sheet Column Mapping
// Headers: LeaveID | EmployeeID | Username | LeaveType | StartDate | EndDate | NumDays | Reason | AttachmentURL | Status (Pending, Approved, Rejected) | AppliedAt | ApprovedByUsername | ApprovedAt | DeniedReason
const LEAVE_COLS = {
  LEAVE_ID: 0,
  EMPLOYEE_ID: 1,
  USERNAME: 2,
  LEAVE_TYPE: 3,
  START_DATE: 4,
  END_DATE: 5,
  NUM_DAYS: 6,
  REASON: 7,
  ATTACHMENT_URL: 8,
  STATUS: 9, // 'Pending', 'Approved', 'Rejected'
  APPLIED_AT: 10,
  APPROVED_BY_USERNAME: 11,
  APPROVED_AT: 12,
  DENIED_REASON: 13
};

// Announcements Sheet Column Mapping
// Headers: AnnouncementID | Title | Content | CreatedByUsername | CreatedAt | ValidUntil
const ANNOUNCEMENT_COLS = {
  ANNOUNCEMENT_ID: 0,
  TITLE: 1,
  CONTENT: 2,
  CREATED_BY_USERNAME: 3,
  CREATED_AT: 4,
  VALID_UNTIL: 5
};

// Payslips Sheet Column Mapping (Simplified example)
// Headers: PayslipID | EmployeeID | Username | MonthYear | DocumentURL | UploadedAt | Comment
const PAYSLIP_COLS = {
  PAYSLIP_ID: 0,
  EMPLOYEE_ID: 1,
  USERNAME: 2,
  MONTH_YEAR: 3,
  DOCUMENT_URL: 4,
  UPLOADED_AT: 5,
  COMMENT: 6
};

// Policies Sheet Column Mapping (Simplified example)
// Headers: PolicyID | Title | Description | DocumentURL | UploadedAt
const POLICIES_COLS = {
  POLICY_ID: 0,
  TITLE: 1,
  DESCRIPTION: 2,
  DOCUMENT_URL: 3,
  UPLOADED_AT: 4
};


// --- Helper Functions ---

/**
 * Retrieves the session ID from the HTTP request headers (specifically the 'Cookie' header).
 * @param {GoogleAppsScript.Events.DoGet | GoogleAppsScript.Events.DoPost} e The event object from doGet or doPost.
 * @returns {string | null} The session ID string if found, otherwise null.
 */
function getSessionIdFromRequest(e) {
  if (e && e.headers && e.headers.Cookie) {
    const cookies = e.headers.Cookie.split(';');
    for (let i = 0; i < cookies.length; i++) {
      const cookie = cookies[i].trim();
      if (cookie.startsWith('nepHR_session=')) {
        return cookie.substring('nepHR_session='.length);
      }
    }
  }
  return null;
}

/**
 * Gets a Google Sheet by its name.
 * @param {string} sheetName The name of the sheet.
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} The sheet object.
 * @throws {Error} If the sheet is not found, or if sheetName is invalid.
 */
function getSheet(sheetName) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    throw new Error(`Sheet "${sheetName}" not found. Please ensure it exists and is named correctly.`);
  }
  return sheet;
}

/**
 * Hashes a given plain text password using SHA-256 and Base64 encoding.
 * @param {string} plainPassword The password to hash.
 * @returns {string} The Base64 encoded SHA-256 hash.
 * @throws {Error} If the password is not a non-empty string.
 */
function hashPassword(plainPassword) {
  const strPassword = String(plainPassword || '');
  if (strPassword.length === 0) {
    throw new Error('Password to hash must be a non-empty string.');
  }
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, strPassword);
  return Utilities.base64Encode(digest);
}

/**
 * Generates a random numeric OTP of specified length.
 * @param {number} length The desired length of the OTP.
 * @returns {string} The generated OTP as a string.
 */
function generateOtp(length) {
  let otp = '';
  for (let i = 0; i < length; i++) {
    otp += Math.floor(Math.random() * 10); // Append a random digit (0-9)
  }
  return otp;
}

/**
 * Generates a unique ID (UUID style).
 * @returns {string} A unique ID.
 */
function generateUuid() {
  return Utilities.getUuid();
}

/**
 * Checks if the authenticated user has the required role.
 * @param {object} sessionInfo - Object containing session details, including role.
 *
 * @param {string | string[]} requiredRoles - A single role string or an array of roles.
 * @returns {boolean} True if the user has the required role, false otherwise.
 */
function checkRole(sessionInfo, requiredRoles) {
  if (!sessionInfo || !sessionInfo.isValid || !sessionInfo.role) {
    console.warn("Role check failed: No valid session or role detected.");
    return false; // No valid session or role
  }
  const userRole = sessionInfo.role;
  if (Array.isArray(requiredRoles)) {
    return requiredRoles.includes(userRole);
  }
  return userRole === requiredRoles;
}

/**
 * Retrieves username, role, employeeID, email, and profile picture from the Users sheet.
 * @param {string} username The username to look up.
 * @returns {object | null} User details (username, role, profilePicUrl) or null if not found.
 */
function getUserDetailsByUsername(username) {
  const usersSheet = getSheet(SHEET_NAMES.USERS);
  const usersData = usersSheet.getDataRange().getValues();
  // Filter for exact username match, case-insensitive
  const userRow = usersData.find(row => row[USER_COLS.USERNAME]?.toString().toLowerCase() === username.toLowerCase());

  if (userRow) {
    return {
      username: userRow[USER_COLS.USERNAME]?.toString() || '',
      role: userRow[USER_COLS.ROLE]?.toString() || '',
      employeeId: userRow[USER_COLS.EMPLOYEE_ID]?.toString() || '',
      email: userRow[USER_COLS.EMAIL]?.toString() || '',
      profilePicUrl: userRow[USER_COLS.PROFILE_PIC_URL]?.toString() || 'https://via.placeholder.com/150/EEEEEE/888888?text=NO+IMG' // Default
    };
  }
  return null;
}

// --- Web App Entry Points ---

/**
 * Handles GET requests to the web app. Serves HTML pages based on session or parameters.
 * @param {GoogleAppsScript.Events.DoGet} e The event object.
 * @returns {GoogleAppsScript.HTML.HtmlOutput} The HTML content to display.
 */
function doGet(e) {
  // Use optional chaining (?.) and nullish coalescing (??) for safer access.
  const pageParam = e?.parameter?.page ?? "login";
  const sessionId = getSessionIdFromRequest(e);
  const sessionInfo = validateSession(sessionId);

  console.log(`doGet: Request for page '${pageParam}'. Session ID: ${sessionId || 'None'}. Session Valid: ${sessionInfo.isValid}. Role: ${sessionInfo.role || 'N/A'}`);
  console.log(`doGet: Received sessionId from request: ${sessionId || 'None'}`);
  console.log(`doGet: Session validation result - isValid: ${sessionInfo.isValid}, Username: ${sessionInfo.username || 'N/A'}, Role: ${sessionInfo.role || 'N/A'}, Message: ${sessionInfo.message || 'N/A'}`);

  // If a specific page is requested and authorized
  if (e && e.parameter && e.parameter.page) {
    // These pages require a valid session
    if (pageParam === 'dashboard' || pageParam === 'admin_dashboard') {
      if (!sessionInfo.isValid) {
        console.log(`doGet: Session invalid for protected page '${pageParam}'. Redirecting to login.`);
        return HtmlService.createHtmlOutputFromFile("login")
          .setTitle('Login')
          .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
      }

      let htmlOutput;

      if (pageParam === 'dashboard') {
        if (checkRole(sessionInfo, [USER_ROLES.EMPLOYEE, USER_ROLES.MANAGER, USER_ROLES.ADMIN, USER_ROLES.HR])) {
          console.log(`doGet: Serving dashboard for ${sessionInfo.username} (${sessionInfo.role}).`);
          htmlOutput = HtmlService.createHtmlOutputFromFile("dashboard")
            .setTitle('Employee Dashboard');
        } else {
          console.warn(`doGet: User ${sessionInfo.username} (Role: ${sessionInfo.role}) attempted to access dashboard but is unauthorized.`);
          htmlOutput = HtmlService.createHtmlOutputFromFile("login") // Unauthorized role, redirect to login
            .setTitle('Login');
        }
      } else if (pageParam === 'admin_dashboard') {
        if (checkRole(sessionInfo, [USER_ROLES.ADMIN, USER_ROLES.HR])) {
          console.log(`doGet: Serving admin_dashboard for ${sessionInfo.username} (${sessionInfo.role}).`);
          htmlOutput = HtmlService.createHtmlOutputFromFile("admin_dashboard")
            .setTitle('Admin Dashboard');
        } else {
          console.warn(`doGet: User ${sessionInfo.username} (Role: ${sessionInfo.role}) attempted to access admin_dashboard but is unauthorized.`);
          htmlOutput = HtmlService.createHtmlOutputFromFile("login") // Unauthorized role, redirect to login
            .setTitle('Login');
        }
      } else {
        // Unknown valid page requested (not dashboard or admin_dashboard)
        console.warn(`doGet: Unknown valid page '${pageParam}' requested. Serving login.`);
        htmlOutput = HtmlService.createHtmlOutputFromFile("login")
          .setTitle('Login');
      }
      return htmlOutput.addMetaTag('viewport', 'width=device-width, initial-scale=1.0');

    } // End of protected page checks
    // If 'page' parameter is present but not 'dashboard' or 'admin_dashboard', still serve login (e.g., 'forgot_password_success')
    console.log(`doGet: Serving login page for unrecognized or unprotected page parameter '${pageParam}'.`);
    return HtmlService.createHtmlOutputFromFile("login")
      .setTitle('Login')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');

  }

  // Default: always serve the login page.
  // This happens if no `page` parameter, or an unrecognized/invalid `page` parameter,
  // or if the session is invalid for protected pages.
  console.log("doGet: Serving login page (default/unrecognized request or invalid session).");
  return HtmlService.createHtmlOutputFromFile("login")
    .setTitle('Login')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

/**
 * Handles POST requests to the web app. Processes form submissions and API calls.
 * @param {GoogleAppsScript.Events.DoPost} e The event object.
 * @returns {GoogleAppsScript.Content.TextOutput} JSON response from the server.
 */
function doPost(e) {
  // Debugging logs for incoming parameters
  console.log("--- doPost Execution Started ---");
  console.log("Received a POST request.");
  if (e && e.parameter) {
    console.log(`Received action: ${e.parameter.action}`);
    console.log(`Received username: ${e.parameter.username}`);
    // Be careful with logging passwords in production environments!
    // console.log("Received password:", e.parameter.password);
    console.log(`Type of e.parameter.password: ${typeof e.parameter.password}`);
    console.log(`Is e.parameter.password an empty string? ${e.parameter.password === ""}`);
  } else {
    console.log("e or e.parameter is null/undefined.");
  }

  if (!e || !e.parameter) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, message: 'Invalid request: Missing event parameters.' })).setMimeType(ContentService.MimeType.JSON);
  }

  const action = e.parameter.action ? String(e.parameter.action) : 'unknown';
  const sessionId = getSessionIdFromRequest(e);
  const sessionInfo = validateSession(sessionId); // Validate session for every request

  let result;
  try {
    // Actions that do NOT require authentication: Login, Forgot Password Workflow, Get App Config
    if (['handleLogin', 'sendPasswordResetEmail', 'validateOtpOnly', 'verifyOtpAndSetPassword', 'getAppConfig'].includes(action)) {
      switch (action) {
        case 'getAppConfig':
          result = getAppConfig();
          break;
        case 'handleLogin':
          if (typeof e.parameter.username !== 'string' || !e.parameter.username ||
            typeof e.parameter.password !== 'string' || !e.parameter.password) {
            result = { success: false, message: 'Missing username or password.' };
          } else {
            result = handleLogin(e.parameter.username, e.parameter.password);
          }
          break;
        case 'sendPasswordResetEmail':
          if (typeof e.parameter.username !== 'string' || !e.parameter.username) {
            result = { success: false, message: 'Missing username for password reset.' };
          } else {
            result = sendPasswordResetEmail(e.parameter.username);
          }
          break;
        case 'validateOtpOnly':
          if (typeof e.parameter.username !== 'string' || !e.parameter.username ||
            typeof e.parameter.otp !== 'string' || !e.parameter.otp) {
            result = { success: false, message: 'Missing username or OTP for validation.' };
          } else {
            result = validateOtpOnly(e.parameter.username, e.parameter.otp);
          }
          break;
        case 'verifyOtpAndSetPassword':
          if (typeof e.parameter.username !== 'string' || !e.parameter.username ||
            typeof e.parameter.otp !== 'string' || !e.parameter.otp ||
            typeof e.parameter.newPassword !== 'string' || !e.parameter.newPassword) {
            result = { success: false, message: 'Missing username, OTP or new password for reset.' };
          } else {
            result = verifyOtpAndSetPassword(e.parameter.username, e.parameter.otp, e.parameter.newPassword);
          }
          break;
        case 'getLoginUrl': // Utility for login.html/reset_password.html to get redirect URL
          result = { success: true, url: ScriptApp.getService().getUrl() };
          break;
      }
    } else { // All actions below this point require authentication
      if (!sessionInfo.isValid) {
        console.warn(`doPost: Authentication failed for action '${action}'. Session ID: ${sessionId} is invalid or expired.`);
        return ContentService.createTextOutput(JSON.stringify({ success: false, message: sessionInfo.message || 'Authentication required.', logout: true })).setMimeType(ContentService.MimeType.JSON);
      }

      // Handle actions based on user role and specific permissions
      switch (action) {
        case 'logout':
          // Logout can be initiated by any valid user
          result = logoutUser(sessionId);
          break;

        // --- EMPLOYEE DASHBOARD ACTIONS (Accessible by Employee, Manager, Admin, HR) ---
        case 'getUserDashboardData':
          if (checkRole(sessionInfo, [USER_ROLES.EMPLOYEE, USER_ROLES.MANAGER, USER_ROLES.ADMIN, USER_ROLES.HR])) {
            result = getUserDashboardData(sessionInfo);
          } else { result = { success: false, message: 'Unauthorized access for this action.' }; }
          break;
        case 'recordAttendance': // For Check-in/Check-out
          if (checkRole(sessionInfo, [USER_ROLES.EMPLOYEE, USER_ROLES.MANAGER, USER_ROLES.ADMIN, USER_ROLES.HR])) {
            // e.parameter.type could be 'Check-in' or 'Check-out'
            result = recordAttendance(sessionInfo.username, e.parameter.type, e.parameter.geoLocation);
          } else { result = { success: false, message: 'Unauthorized access for this action.' }; }
          break;
        case 'getMonthlyAttendance':
          if (checkRole(sessionInfo, [USER_ROLES.EMPLOYEE, USER_ROLES.MANAGER, USER_ROLES.ADMIN, USER_ROLES.HR])) {
            result = getMonthlyAttendance(sessionInfo.username, parseInt(e.parameter.year), parseInt(e.parameter.month));
          } else { result = { success: false, message: 'Unauthorized access for this action.' }; }
          break;
        case 'getLeaveBalances':
          if (checkRole(sessionInfo, [USER_ROLES.EMPLOYEE, USER_ROLES.MANAGER, USER_ROLES.ADMIN, USER_ROLES.HR])) {
            result = getLeaveBalances(sessionInfo.username);
          } else { result = { success: false, message: 'Unauthorized access for this action.' }; }
          break;
        case 'applyLeave':
          if (checkRole(sessionInfo, [USER_ROLES.EMPLOYEE, USER_ROLES.MANAGER, USER_ROLES.ADMIN, USER_ROLES.HR])) {
            result = applyLeave(sessionInfo.username, JSON.parse(e.parameter.leaveData));
          } else { result = { success: false, message: 'Unauthorized access for this action.' }; }
          break;
        case 'getUserLeaveHistory':
          if (checkRole(sessionInfo, [USER_ROLES.EMPLOYEE, USER_ROLES.MANAGER, USER_ROLES.ADMIN, USER_ROLES.HR])) {
            result = getUserLeaveHistory(sessionInfo.username);
          } else { result = { success: false, message: 'Unauthorized access for this action.' }; }
          break;
        case 'getUserTasks':
          if (checkRole(sessionInfo, [USER_ROLES.EMPLOYEE, USER_ROLES.MANAGER, USER_ROLES.ADMIN, USER_ROLES.HR])) {
            result = getUserTasks(sessionInfo.username, e.parameter.filter);
          } else { result = { success: false, message: 'Unauthorized access for this action.' }; }
          break;
        case 'updateTaskProgress':
          if (checkRole(sessionInfo, [USER_ROLES.EMPLOYEE, USER_ROLES.MANAGER, USER_ROLES.ADMIN, USER_ROLES.HR])) {
            result = updateTaskProgress(sessionInfo.username, e.parameter.taskId, e.parameter.progress, e.parameter.comments, e.parameter.status);
          } else { result = { success: false, message: 'Unauthorized access for this action.' }; }
          break;
        case 'getAnnouncements': // Employee and Admin can view
          if (checkRole(sessionInfo, [USER_ROLES.EMPLOYEE, USER_ROLES.MANAGER, USER_ROLES.ADMIN, USER_ROLES.HR])) {
            result = getAnnouncements();
          } else { result = { success: false, message: 'Unauthorized access for this action.' }; }
          break;
        case 'updateProfile':
          if (checkRole(sessionInfo, [USER_ROLES.EMPLOYEE, USER_ROLES.MANAGER, USER_ROLES.ADMIN, USER_ROLES.HR])) {
            result = updateProfile(sessionInfo.username, JSON.parse(e.parameter.profileData));
          } else { result = { success: false, message: 'Unauthorized access for this action.' }; }
          break;
        case 'changePassword':
          if (checkRole(sessionInfo, [USER_ROLES.EMPLOYEE, USER_ROLES.MANAGER, USER_ROLES.ADMIN, USER_ROLES.HR])) {
            result = changePassword(sessionInfo.username, e.parameter.oldPassword, e.parameter.newPassword);
          } else { result = { success: false, message: 'Unauthorized access for this action.' }; }
          break;
        case 'uploadProfilePicture':
          if (checkRole(sessionInfo, [USER_ROLES.EMPLOYEE, USER_ROLES.MANAGER, USER_ROLES.ADMIN, USER_ROLES.HR])) {
            result = uploadProfilePicture(sessionInfo.username, e);
          } else { result = { success: false, message: 'Unauthorized access for this action.' }; }
          break;
        case 'getPayslips':
          if (checkRole(sessionInfo, [USER_ROLES.EMPLOYEE, USER_ROLES.MANAGER, USER_ROLES.ADMIN, USER_ROLES.HR])) {
            result = getPayslips(sessionInfo.username);
          } else { result = { success: false, message: 'Unauthorized access for this action.' }; }
          break;
        case 'getCompanyPolicies':
          if (checkRole(sessionInfo, [USER_ROLES.EMPLOYEE, USER_ROLES.MANAGER, USER_ROLES.ADMIN, USER_ROLES.HR])) {
            result = getCompanyPolicies();
          } else { result = { success: false, message: 'Unauthorized access for this action.' }; }
          break;

        // --- ADMIN DASHBOARD ACTIONS (Accessible by Admin, HR unless specified) ---
        case 'uploadBanner': // Admin specific action
          if (checkRole(sessionInfo, [USER_ROLES.ADMIN])) {
            result = uploadBanner(e);
          } else { result = { success: false, message: 'Unauthorized access. Admin privileges required.' }; }
          break;
        case 'updateAppConfig':
          if (checkRole(sessionInfo, [USER_ROLES.ADMIN])) {
            result = updateAppConfig(e.parameter.key, e.parameter.value);
          } else { result = { success: false, message: 'Unauthorized access. Admin privileges required.' }; }
          break;
        case 'getAllEmployeesData':
          if (checkRole(sessionInfo, [USER_ROLES.ADMIN, USER_ROLES.HR])) {
            result = getAllEmployeesData();
          } else { result = { success: false, message: 'Unauthorized access. Admin/HR privileges required.' }; }
          break;
        case 'addUpdateEmployee':
          if (checkRole(sessionInfo, [USER_ROLES.ADMIN, USER_ROLES.HR])) {
            result = addUpdateEmployee(JSON.parse(e.parameter.employeeData));
          } else { result = { success: false, message: 'Unauthorized access. Admin/HR privileges required.' }; }
          break;
        case 'deleteEmployee':
          if (checkRole(sessionInfo, [USER_ROLES.ADMIN, USER_ROLES.HR])) {
            result = deleteEmployee(e.parameter.username);
          } else { result = { success: false, message: 'Unauthorized access. Admin/HR privileges required.' }; }
          break;
        case 'getAdminAttendanceSummary':
          if (checkRole(sessionInfo, [USER_ROLES.ADMIN, USER_ROLES.HR])) {
            result = getAdminAttendanceSummary(e.parameter.date || new Date().toISOString(), e.parameter.employeeIdFilter);
          } else { result = { success: false, message: 'Unauthorized access. Admin/HR privileges required.' }; }
          break;
        case 'assignTask':
          if (checkRole(sessionInfo, [USER_ROLES.ADMIN, USER_ROLES.MANAGER])) { // Managers can also assign tasks
            result = assignTask(sessionInfo.username, JSON.parse(e.parameter.taskData));
          } else { result = { success: false, message: 'Unauthorized access. Admin/Manager privileges required.' }; }
          break;
        case 'getAllTasks':
          if (checkRole(sessionInfo, [USER_ROLES.ADMIN, USER_ROLES.MANAGER])) { // Managers can also view all tasks
            result = getAllTasks(e.parameter.filter);
          } else { result = { success: false, message: 'Unauthorized access. Admin/Manager privileges required.' }; }
          break;
        case 'approveRejectTask':
          if (checkRole(sessionInfo, [USER_ROLES.ADMIN, USER_ROLES.MANAGER])) {
            result = approveRejectTask(sessionInfo.username, e.parameter.taskId, e.parameter.status, e.parameter.comments);
          } else { result = { success: false, message: 'Unauthorized access. Admin/Manager privileges required.' }; }
          break;
        case 'getAllLeaveApplications':
          if (checkRole(sessionInfo, [USER_ROLES.ADMIN, USER_ROLES.HR])) {
            result = getAllLeaveApplications(e.parameter.filter);
          } else { result = { success: false, message: 'Unauthorized access. Admin/HR privileges required.' }; }
          break;
        case 'approveRejectLeaveApplication':
          if (checkRole(sessionInfo, [USER_ROLES.ADMIN, USER_ROLES.HR])) {
            result = approveRejectLeaveApplication(sessionInfo.username, e.parameter.leaveId, e.parameter.status, e.parameter.deniedReason);
          } else { result = { success: false, message: 'Unauthorized access. Admin/HR privileges required.' }; }
          break;
        case 'createAnnouncement':
          if (checkRole(sessionInfo, [USER_ROLES.ADMIN, USER_ROLES.HR])) {
            result = createAnnouncement(sessionInfo.username, JSON.parse(e.parameter.announcementData));
          } else { result = { success: false, message: 'Unauthorized access. Admin/HR privileges required.' }; }
          break;
        case 'deleteAnnouncement':
          if (checkRole(sessionInfo, [USER_ROLES.ADMIN, USER_ROLES.HR])) {
            result = deleteAnnouncement(e.parameter.announcementId);
          } else { result = { success: false, message: 'Unauthorized access. Admin/HR privileges required.' }; }
          break;
        case 'addUpdatePolicy':
          if (checkRole(sessionInfo, [USER_ROLES.ADMIN, USER_ROLES.HR])) {
            result = addUpdatePolicy(JSON.parse(e.parameter.policyData));
          } else { result = { success: false, message: 'Unauthorized access. Admin/HR privileges required.' }; }
          break;
        case 'deletePolicy':
          if (checkRole(sessionInfo, [USER_ROLES.ADMIN, USER_ROLES.HR])) {
            result = deletePolicy(e.parameter.policyId);
          } else { result = { success: false, message: 'Unauthorized access. Admin/HR privileges required.' }; }
          break;
        case 'uploadPayslip':
          if (checkRole(sessionInfo, [USER_ROLES.ADMIN, USER_ROLES.HR])) {
            // Needs to check if e.parameters contains required data for uploadPayslip
            if (!e.parameters || !e.parameters.username || !e.parameters.monthYear || !e.parameters.file) {
              result = { success: false, message: 'Missing required payslip data (username, monthYear, file).' };
            } else {
              // For some reason, in this specific case withFileUpload, the file blob is passed directly
              // while other params are in e.parameter. We pass the whole e object.
              result = uploadPayslip(e);
            }
          } else { result = { success: false, message: 'Unauthorized access. Admin/HR privileges required.' }; }
          break;
        default:
          result = { success: false, message: `Invalid or unknown action: ${action}` };
      }
    }
  } catch (error) {
    result = { success: false, message: `Server error during action "${action}": ${error.message}` };
    console.error(`ERROR in doPost for action "${action}":`, error);
  }
  console.log("--- doPost Execution Finished ---");
  return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
}

// --- API Handlers for client-side calls ---

/**
 * Retrieves application configuration settings from the 'AppConfig' sheet.
 * @returns {object} An object containing 'bannerUrl' and 'accentColor'.
 */
function getAppConfig() {
  try {
    const configSheet = getSheet(SHEET_NAMES.APP_CONFIG);
    const data = configSheet.getDataRange().getValues();
    const config = {};
    data.forEach(row => {
      const key = row[APP_CONFIG_COLS.KEY];
      const value = row[APP_CONFIG_COLS.VALUE];
      if (key && value !== undefined) {
        config[key] = value;
      }
    });

    // Default banner URL if not found in config or during error
    const defaultBannerUrl = 'https://via.placeholder.com/300x50?text=Company+Banner';
    return {
      success: true,
      bannerUrl: config.bannerUrl || defaultBannerUrl,
      accentColor: config.accentColor || '' // Uses CSS default if not set
    };
  } catch (e) {
    console.error("Error getting app config:", e);
    return { success: false, message: "Could not load app configuration." };
  }
}

/**
 * Updates a specific key-value pair in the AppConfig sheet.
 * @param {string} key The configuration key to update.
 * @param {string} value The new value for the key.
 * @returns {object} Success status and message.
 */
function updateAppConfig(key, value) {
  try {
    const configSheet = getSheet(SHEET_NAMES.APP_CONFIG);
    const data = configSheet.getDataRange().getValues();
    let updated = false;
    for (let i = 0; i < data.length; i++) {
      if (data[i][APP_CONFIG_COLS.KEY] === key) {
        configSheet.getRange(i + 1, APP_CONFIG_COLS.VALUE + 1).setValue(value);
        updated = true;
        break;
      }
    }
    if (!updated) {
      configSheet.appendRow([key, value]);
    }
    return { success: true, message: `App config "${key}" updated successfully.` };
  } catch (e) {
    console.error("Error updating app config:", e);
    return { success: false, message: `Failed to update app config: ${e.message}` };
  }
}

/**
 * Handles user login attempts.
 *
 * @param {string} username The username.
 * @param {string} plainPassword The plain text password.
 * @returns {object} Login result with success, message, and session ID.
 */
function handleLogin(username, plainPassword) {
  try {
    console.log(`handleLogin: Attempting login for username: ${username}`);
    const usersSheet = getSheet(SHEET_NAMES.USERS);
    const usersData = usersSheet.getDataRange().getValues();
    const hashedPassword = hashPassword(plainPassword);

    // Find the user row matching username and hashed password
    const userRow = usersData.find(row =>
      row[USER_COLS.USERNAME]?.toString().toLowerCase() === username.toLowerCase() &&
      row[USER_COLS.PASSWORD_HASH]?.toString() === hashedPassword
    );

    if (!userRow) {
      console.log(`handleLogin: Invalid credentials for username: ${username}`);
      return { success: false, message: 'Invalid username or password.' };
    }

    const sessionId = Utilities.getUuid();
    const currentTime = new Date();
    const expiryTime = new Date(currentTime.getTime() + (SESSION_EXPIRY_MINUTES * 60 * 1000)); // 60 minutes from now
    const userRole = userRow[USER_COLS.ROLE]?.toString(); // Get role from Users sheet

    if (!userRole) {
      console.warn(`handleLogin: User ${username} has no role defined in Users sheet.`);
      return { success: false, message: 'User role not defined. Please contact support.' };
    }

    const sessionsSheet = getSheet(SHEET_NAMES.ACTIVE_SESSIONS);

    // Clear existing sessions for this user to ensure only one active session per user
    const existingSessions = sessionsSheet.getDataRange().getValues();
    // Collect row numbers to delete (1-indexed)
    const rowsToDelete = [];
    for (let i = existingSessions.length - 1; i >= 0; i--) {
        // Ensure username is not null or undefined before calling .toLowerCase()
        if (existingSessions[i][SESSION_COLS.USERNAME] && existingSessions[i][SESSION_COLS.USERNAME].toString().toLowerCase() === username.toLowerCase()) {
            rowsToDelete.push(i + 1);
        }
    }

    // Delete rows starting from the end to prevent index shifting issues
    if (rowsToDelete.length > 0) {
        // Sort in descending order to delete from bottom up
        rowsToDelete.sort((a,b) => b-a).forEach(rowNum => {
            sessionsSheet.deleteRow(rowNum);
        });
        console.log(`handleLogin: Cleaned up ${rowsToDelete.length} old sessions for ${username}.`);
    }


    // Append new session
    sessionsSheet.appendRow([sessionId, username, currentTime.toISOString(), expiryTime.toISOString(), userRole]);
    console.log(`handleLogin: New session created for ${username}: ${sessionId}`);

    // Determine redirect URL based on role
    let redirectUrl;
    if (userRole === USER_ROLES.ADMIN || userRole === USER_ROLES.HR) {
      redirectUrl = ScriptApp.getService().getUrl() + '?page=admin_dashboard';
      console.log(`handleLogin: Redirecting to Admin Dashboard for ${username}.`);
    } else if (userRole === USER_ROLES.EMPLOYEE || userRole === USER_ROLES.MANAGER) {
      redirectUrl = ScriptApp.getService().getUrl() + '?page=dashboard';
      console.log(`handleLogin: Redirecting to Employee Dashboard for ${username}.`);
    } else {
      // Default fallback if role doesn't match above, though roles should be defined
      redirectUrl = ScriptApp.getService().getUrl() + '?page=dashboard';
      console.warn(`handleLogin: Unknown role '${userRole}' for ${username}. Redirecting to default dashboard.`);
    }

    return {
      success: true,
      message: 'Login successful!',
      sessionId: sessionId,
      redirectUrl: redirectUrl
    };

  } catch (e) {
    console.error(`handleLogin Error for username ${username}:`, e);
    return { success: false, message: `An error occurred during login: ${e.message}` };
  }
}

/**
 * Validates a session ID.
 * @param {string} sessionId The session ID to validate.
 * @returns {object} An object indicating validity and user info if valid.
 */
function validateSession(sessionId) {
  if (!sessionId) {
    console.log("validateSession: No session ID provided.");
    return { isValid: false, message: 'No session ID provided.', logout: true };
  }
  try {
    const sessionsSheet = getSheet(SHEET_NAMES.ACTIVE_SESSIONS);
    const sessionsData = sessionsSheet.getDataRange().getValues();
    const currentTime = new Date().getTime();

    for (let i = 0; i < sessionsData.length; i++) {
      const row = sessionsData[i];
      const storedSessionId = row[SESSION_COLS.SESSION_ID]?.toString();
      const storedUsername = row[SESSION_COLS.USERNAME]?.toString();
      const expiryTimestamp = new Date(row[SESSION_COLS.EXPIRY_TIME]).getTime();
      const userRoleInSession = row[SESSION_COLS.USER_ROLE]?.toString();

      if (storedSessionId === sessionId) {
        if (currentTime < expiryTimestamp) {
          // Update expiry time to extend session (sliding window)
          const newExpiryTime = new Date(currentTime + (SESSION_EXPIRY_MINUTES * 60 * 1000));
          sessionsSheet.getRange(i + 1, SESSION_COLS.EXPIRY_TIME + 1).setValue(newExpiryTime.toISOString());
          console.log(`validateSession: Session valid for user ${storedUsername}, role ${userRoleInSession}. Expiry extended.`);
          return {
            isValid: true,
            username: storedUsername,
            role: userRoleInSession,
            sessionId: storedSessionId
          };
        } else {
          // Session expired, remove it from sheet
          console.log(`validateSession: Session ${sessionId} expired for user ${storedUsername}. Deleting.`);
          sessionsSheet.deleteRow(i + 1); // +1 for 1-based indexing
          return { isValid: false, message: 'Session expired. Please log in again.', logout: true };
        }
      }
    }
    console.log(`validateSession: Invalid session ID ${sessionId}. No matching session found.`);
    return { isValid: false, message: 'Invalid session ID. Please log in again.', logout: true };
  } catch (e) {
    console.error("validateSession Error:", e);
    return { isValid: false, message: `Error validating session: ${e.message}`, logout: true };
  }
}

/**
 * Sends an OTP to the user's registered email for password reset.
 * @param {string} username The username (email) to send OTP to.
 * @returns {object} Result of the operation.
 */
function sendPasswordResetEmail(username) {
  try {
    const usersSheet = getSheet(SHEET_NAMES.USERS);
    const usersData = usersSheet.getDataRange().getValues();

    const userRowIndex = usersData.findIndex(row => row[USER_COLS.USERNAME]?.toString().toLowerCase() === username.toLowerCase());

    if (userRowIndex === -1) {
      return { success: false, message: 'No user found with that username.' };
    }

    const userEmail = usersData[userRowIndex][USER_COLS.EMAIL]?.toString();
    if (!userEmail || typeof userEmail !== 'string' || userEmail.trim() === '') {
      return { success: false, message: 'Email address not found or invalid for this user. Cannot send OTP.' };
    }

    const otp = generateOtp(OTP_LENGTH);
    const expiryTime = new Date().getTime() + (OTP_EXPIRY_SECONDS * 1000); // Current time + OTP_EXPIRY_SECONDS

    // Update the Users sheet with the new OTP and its expiry
    const targetRow = userRowIndex + 1; // Convert 0-indexed to 1-indexed

    usersSheet.getRange(targetRow, USER_COLS.OTP + 1).setValue(otp); // +1 for 1-based column
    usersSheet.getRange(targetRow, USER_COLS.OTP_EXPIRY + 1).setValue(expiryTime);

    // Send the email
    MailApp.sendEmail({
      to: userEmail,
      subject: 'NepHR Password Reset OTP',
      htmlBody: `<html><body>
          <p>Dear ${username},</p>
          <p>Your One-Time Password (OTP) for password reset is: <strong>${otp}</strong></p>
          <p>This OTP is valid for ${OTP_EXPIRY_SECONDS / 60} minutes.</p>
          <p>If you did not request this, please ignore this email.</p>
          <p>Sincerely,<br>NepHR Team</p>
          </body></html>`
    });

    return {
      success: true,
      message: 'An OTP has been sent to your registered email.',
      otpExpiryMinutes: OTP_EXPIRY_SECONDS / 60, // Send expiry back to client
      // IMPORTANT: DO NOT SEND THE OTP ITSELF BACK TO THE CLIENT FOR SECURITY.
      // The client-side timer needs only the expiry duration, not the actual OTP.
      serverExpiryTimestamp: expiryTime // Send timestamp for accurate client-side timer
    };

  } catch (e) {
    console.error("Error sending password reset email:", e);
    return { success: false, message: `Failed to send OTP: ${e.message}` };
  }
}

/**
 * Validates an OTP without setting a new password. Used for step validation.
 * @param {string} username The username.
 * @param {string} otp The OTP entered by the user.
 * @returns {object} Validation result.
 */
function validateOtpOnly(username, otp) {
  try {
    const usersSheet = getSheet(SHEET_NAMES.USERS);
    const usersData = usersSheet.getDataRange().getValues();

    const userRowIndex = usersData.findIndex(row => row[USER_COLS.USERNAME]?.toString().toLowerCase() === username.toLowerCase());

    if (userRowIndex === -1) {
      return { success: false, message: 'Invalid OTP or username.' };
    }

    const storedOtp = usersData[userRowIndex][USER_COLS.OTP]?.toString();
    const storedExpiry = usersData[userRowIndex][USER_COLS.OTP_EXPIRY];
    const currentTime = new Date().getTime();

    // Check if OTP matches and is within expiry (storedExpiry is a number/timestamp, convert to number if it's a string)
    if (storedOtp == otp && currentTime < Number(storedExpiry) && storedOtp != '') {
      // OTP is correct and not expired. Don't clear it yet, as it's needed for final password set.
      return { success: true, message: 'OTP verified successfully!' };
    } else {
      // Incorrect or expired OTP. Clear OTP and expiry to prevent brute-force/reuse.
      usersSheet.getRange(userRowIndex + 1, USER_COLS.OTP + 1).setValue('');
      usersSheet.getRange(userRowIndex + 1, USER_COLS.OTP_EXPIRY + 1).setValue('');
      return { success: false, message: 'Invalid or expired OTP.' };
    }
  } catch (e) {
    console.error("Error validating OTP:", e);
    return { success: false, message: `An error occurred during OTP validation: ${e.message}` };
  }
}

/**
 * Verifies the OTP and then sets the new password.
 * @param {string} username The username.
 * @param {string} otp The OTP from the user.
 * @param {string} newPassword The new password to set.
 * @returns {object} Password reset result.
 */
function verifyOtpAndSetPassword(username, otp, newPassword) {
  try {
    const usersSheet = getSheet(SHEET_NAMES.USERS);
    const usersData = usersSheet.getDataRange().getValues();

    const userRowIndex = usersData.findIndex(row => row[USER_COLS.USERNAME]?.toString().toLowerCase() === username.toLowerCase());

    if (userRowIndex === -1) {
      return { success: false, message: 'User not found for password reset.' };
    }

    const storedOtp = usersData[userRowIndex][USER_COLS.OTP]?.toString();
    const storedExpiry = usersData[userRowIndex][USER_COLS.OTP_EXPIRY];
    const currentTime = new Date().getTime();

    // Re-check OTP and expiry. It's crucial for security to re-validate on the server.
    if (storedOtp == otp && currentTime < Number(storedExpiry) && storedOtp != '') {
      // OTP is valid. Now update password and clear OTP fields.
      const hashedPassword = hashPassword(newPassword);
      const targetRow = userRowIndex + 1;

      usersSheet.getRange(targetRow, USER_COLS.PASSWORD_HASH + 1).setValue(hashedPassword);
      usersSheet.getRange(targetRow, USER_COLS.OTP + 1).setValue(''); // Clear OTP after successful reset
      usersSheet.getRange(targetRow, USER_COLS.OTP_EXPIRY + 1).setValue(''); // Clear OTP expiry

      return { success: true, message: 'Password has been reset successfully!' };

    } else {
      // Invalid or expired OTP. Clear everything (even if it was already cleared by validateOtpOnly)
      usersSheet.getRange(userRowIndex + 1, USER_COLS.OTP + 1).setValue('');
      usersSheet.getRange(userRowIndex + 1, USER_COLS.OTP_EXPIRY + 1).setValue('');
      return { success: false, message: 'Invalid or expired OTP. Please restart the reset process.' };
    }
  } catch (e) {
    console.error("Error setting new password:", e);
    return { success: false, message: `An error occurred while setting new password: ${e.message}` };
  }
}

/**
 * Logs out a user by invalidating their session.
 * @param {string} sessionId The session ID to invalidate.
 * @returns {object} Result of the logout operation.
 */
function logoutUser(sessionId) {
  try {
    const sessionsSheet = getSheet(SHEET_NAMES.ACTIVE_SESSIONS);
    const data = sessionsSheet.getDataRange().getValues();

    let userLoggedOut = false;
    for (let i = data.length - 1; i >= 0; i--) { // Iterate backwards to safely delete rows
      if (data[i][SESSION_COLS.SESSION_ID]?.toString() === sessionId) {
        sessionsSheet.deleteRow(i + 1);
        userLoggedOut = true;
        break;
      }
    }

    if (userLoggedOut) {
      return { success: true, message: 'Logged out successfully.', logout: true };
    } else {
      return { success: false, message: 'Session not found or already logged out.', logout: true };
    }
  } catch (e) {
    console.error("Error during logout:", e);
    return { success: false, message: `An error occurred during logout: ${e.message}`, logout: true };
  }
}
/**
 * Retrieves dashboard data for an authenticated user.
 * @param {object} sessionInfo The user's session information.
 * @returns {object} Dashboard data or error if session is invalid.
 */
function getUserDashboardData(sessionInfo) {
  try {
    if (!sessionInfo || !sessionInfo.isValid) {
      return { success: false, message: 'Session invalid or expired.' };
    }

    const username = sessionInfo.username;
    const userDetails = getUserDetailsByUsername(username);

    if (userDetails) {
      const today = new Date();
      // Reset time for comparison
      today.setHours(0,0,0,0);
      const currentMonth = today.getMonth() + 1; // JavaScript months are 0-indexed
      const currentYear = today.getFullYear();

      // Get attendance status for today
      let attendanceStatusToday = 'Not Checked In';
      let checkInTime = null;
      let checkOutTime = null;
      try {
        const attendanceSheet = getSheet(SHEET_NAMES.ATTENDANCE);
        const attendanceData = attendanceSheet.getDataRange().getValues()
          .filter(row => row[ATTENDANCE_COLS.USERNAME]?.toString() === username);

        let latestCheckInEntry = null;
        let latestCheckOutEntry = null;

        // Filter for today's entries and find latest check-in/out
        for (const record of attendanceData) {
          const recordDate = new Date(record[ATTENDANCE_COLS.DATE]);
          recordDate.setHours(0,0,0,0); // Reset time for comparison
          if (recordDate.getTime() === today.getTime()) {
            if (record[ATTENDANCE_COLS.TYPE] === 'Check-in') {
                if (!latestCheckInEntry || record[ATTENDANCE_COLS.TIME] > latestCheckInEntry[ATTENDANCE_COLS.TIME]) {
                    latestCheckInEntry = record;
                }
            } else if (record[ATTENDANCE_COLS.TYPE] === 'Check-out') {
                if (!latestCheckOutEntry || record[ATTENDANCE_COLS.TIME] > latestCheckOutEntry[ATTENDANCE_COLS.TIME]) {
                    latestCheckOutEntry = record;
                }
            }
          }
        }

        if (latestCheckInEntry) {
            checkInTime = latestCheckInEntry[ATTENDANCE_COLS.TIME];
        }
        if (latestCheckOutEntry) {
            checkOutTime = latestCheckOutEntry[ATTENDANCE_COLS.TIME];
        }

        if (checkInTime && checkOutTime) {
            attendanceStatusToday = 'Checked Out';
        } else if (checkInTime) {
            attendanceStatusToday = 'Checked In';
        }

      } catch (e) {
        console.warn(`Could not retrieve attendance for ${username}: ${e.message}`);
        // Continue without attendance data if sheet doesn't exist etc.
      }


      // Get leave balances
      let leaveBalances = {};
      try {
        // Fetch leave balances using the dedicated function
        const balancesResponse = getLeaveBalances(username);
        if (balancesResponse.success) {
            balancesResponse.data.forEach(item => {
                leaveBalances[item.type] = {
                    total: item.total,
                    used: item.used,
                    remaining: item.remaining
                };
            });
        }
      } catch (e) {
        console.warn(`Could not retrieve leave balances for ${username}: ${e.message}`);
      }

      // Get user tasks overview
      let pendingTasks = 0;
      let completedTasks = 0;
      try {
        const tasksSheet = getSheet(SHEET_NAMES.TASKS);
        const tasksData = tasksSheet.getDataRange().getValues();
        const userTasks = tasksData.filter(row => row[TASK_COLS.ASSIGNED_TO_USERNAME]?.toString() === username); // Ensure username comparison
        userTasks.forEach(task => {
          if (task[TASK_COLS.STATUS] === 'Completed' || task[TASK_COLS.STATUS] === 'Approved') {
            completedTasks++;
          } else if (task[TASK_COLS.STATUS] === 'Pending' || task[TASK_COLS.STATUS] === 'In Progress' || task[TASK_COLS.STATUS] === 'Submitted') {
            pendingTasks++;
          }
        });
      } catch (e) {
        console.warn(`Could not retrieve tasks for ${username}: ${e.message}`);
      }

      // Get employee details (full name, department, from Employees sheet)
      // This is accessed by username, so we look up the Employee record
      let employeeFullName = userDetails.username; // Default to username
      let employeeDepartment = 'N/A';
      let employeeRow = null; 
      try {
        const employeesSheet = getSheet(SHEET_NAMES.EMPLOYEES);
        const employeesData = employeesSheet.getDataRange().getValues();
        employeeRow = employeesData.find(row => row[EMPLOYEE_COLS.USERNAME]?.toString() === username);

        if (employeeRow) {
          employeeFullName = employeeRow[EMPLOYEE_COLS.FULL_NAME]?.toString() || employeeFullName;
          employeeDepartment = employeeRow[EMPLOYEE_COLS.DEPARTMENT]?.toString() || employeeDepartment;
        }
      } catch (e) {
        console.warn(`Could not retrieve full employee details for ${username} from Employees sheet: ${e.message}`);
      }

      return {
        success: true,
        message: 'Dashboard data retrieved successfully.',
        data: {
          username: userDetails.username,
          role: userDetails.role,
          employeeId: userDetails.employeeId,
          fullName: employeeFullName,
          email: userDetails.email, // For profile editing
          phoneNumber: employeeRow ? employeeRow[EMPLOYEE_COLS.PHONE_NUMBER]?.toString() : '', // From Employee sheet
          department: employeeDepartment,
          profilePicUrl: userDetails.profilePicUrl,
          attendanceStatusToday: attendanceStatusToday,
          checkInTimeToday: checkInTime,
          checkOutTimeToday: checkOutTime,
          leaveBalances: leaveBalances,
          pendingTasks: pendingTasks,
          completedTasks: completedTasks
        }
      };
    } else {
      return { success: false, message: 'User data not found for session.' };
    }
  } catch (e) {
    console.error("Error retrieving dashboard data:", e);
    return { success: false, message: `Error retrieving dashboard data: ${e.message}` };
  }
}

/**
 * Uploads a banner image to Google Drive and records its details.
 * Also updates `bannerUrl` in AppConfig sheet.
 * @param {GoogleAppsScript.Events.DoPost} e The event object containing file data.
 * @returns {object} Upload result.
 */
function uploadBanner(e) {
  try {
    // For file uploads via google.script.run.withFileUpload()
    // the file blob is often directly in e.parameter.file or e.postData.contents
    const fileBlob = e.parameters.file && e.parameters.file.length > 0 ? e.parameters.file[0] : null;

    if (!fileBlob || typeof fileBlob.getName !== 'function') {
      return { success: false, message: 'No valid file data received.' };
    }

    const folderName = "NepHR Banners";
    let folderIterator = DriveApp.getFoldersByName(folderName);
    let folder;
    if (folderIterator.hasNext()) {
      folder = folderIterator.next();
    } else {
      folder = DriveApp.createFolder(folderName); // Create new folder if not found
    }

    const uploadedFile = folder.createFile(fileBlob);
    const fileUrl = uploadedFile.getUrl(); // Get Google Drive URL
    const fileName = uploadedFile.getName();
    const fileId = uploadedFile.getId();

    // Make the file publicly accessible (needed if displaying directly on web app)
    // IMPORTANT: Be aware of security implications if uploading sensitive files.
    uploadedFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    // Record file details in 'Banners' sheet
    try {
      const bannersSheet = getSheet(SHEET_NAMES.BANNERS);
      // Assuming Banners sheet has columns for: FileName, FileURL, FileID, UploadedAt, Uploader
      bannersSheet.appendRow([fileName, fileUrl, fileId, new Date().toISOString(), e.parameters.uploaderName || "Admin"]);
    } catch (sheetError) {
      console.warn("Could not record banner details in 'Banners' sheet:", sheetError.message);
    }

    // Update the AppConfig with this new banner URL
    try {
      updateAppConfig('bannerUrl', fileUrl); // Re-use the existing updateAppConfig function
    } catch (configError) {
      console.warn("Could not update 'bannerUrl' in AppConfig sheet:", configError.message);
    }

    return { success: true, message: `Banner "${fileName}" uploaded successfully!`, fileUrl: fileUrl, fileId: fileId };

  } catch (error) {
    console.error("Error uploading banner:", error);
    return { success: false, message: `Error uploading banner: ${error.message}` };
  }
}

/**
 * Uploads a profile picture for a user to Google Drive and updates Users sheet.
 * @param {string} username The username for whom the picture is being uploaded.
 * @param {GoogleAppsScript.Events.DoPost} e The event object containing file data.
 * @returns {object} Upload result.
 */
function uploadProfilePicture(username, e) {
  try {
    const fileBlob = e.parameters.file && e.parameters.file.length > 0 ? e.parameters.file[0] : null;

    if (!fileBlob || typeof fileBlob.getName !== 'function') {
      return { success: false, message: 'No valid file data received.' };
    }

    const folderName = "NepHR Profile Pictures";
    let folderIterator = DriveApp.getFoldersByName(folderName);
    let folder;
    if (folderIterator.hasNext()) {
      folder = folderIterator.next();
    } else {
      folder = DriveApp.createFolder(folderName);
    }

    const uploadedFile = folder.createFile(fileBlob);
    const fileUrl = uploadedFile.getUrl();
    uploadedFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    // Update the USER_COLS.PROFILE_PIC_URL in the Users sheet
    const usersSheet = getSheet(SHEET_NAMES.USERS);
    const usersData = usersSheet.getDataRange().getValues();
    const userRowIndex = usersData.findIndex(row => row[USER_COLS.USERNAME]?.toString().toLowerCase() === username.toLowerCase());

    if (userRowIndex === -1) {
      DriveApp.getFileById(uploadedFile.getId()).setTrashed(true); // Clean up uploaded file
      return { success: false, message: 'User not found to update profile picture.' };
    }

    usersSheet.getRange(userRowIndex + 1, USER_COLS.PROFILE_PIC_URL + 1).setValue(fileUrl);

    // Also update in Employees sheet if it exists and username matches
    try {
      const employeesSheet = getSheet(SHEET_NAMES.EMPLOYEES);
      const employeesData = employeesSheet.getDataRange().getValues();
      const employeeRowIndex = employeesData.findIndex(row => row[EMPLOYEE_COLS.USERNAME]?.toString().toLowerCase() === username.toLowerCase());
      if (employeeRowIndex !== -1) {
        employeesSheet.getRange(employeeRowIndex + 1, EMPLOYEE_COLS.PUBLIC_PROFILE_PIC_URL + 1).setValue(fileUrl);
      }
    } catch (e) {
      console.warn("Could not update profile picture in Employees sheet:", e);
      // Not a critical error if Employees sheet update fails but Users succeed
    }

    return { success: true, message: 'Profile picture uploaded and updated successfully!', fileUrl: fileUrl };

  } catch (error) {
    console.error("Error uploading profile picture:", error);
    return { success: false, message: `Error uploading profile picture: ${error.message}` };
  }
}

// --- HRMS BACKEND FUNCTIONS ---

/**
 * Records an attendance entry (Check-in or Check-out).
 * @param {string} username The username.
 * @param {'Check-in'|'Check-out'} type The attendance type.
 * @param {string} geoLocation The geographic location (e.g., "lat,long").
 * @returns {object} Result of the operation.
 */
function recordAttendance(username, type, geoLocation) {
  try {
    const attendanceSheet = getSheet(SHEET_NAMES.ATTENDANCE);
    const userDetails = getUserDetailsByUsername(username);
    const employeeId = userDetails ? userDetails.employeeId : 'N/A';

    const now = new Date();
    const todayDate = new Date(now.getFullYear(), now.getMonth(), now.getDate());
    const currentTime = now.toLocaleTimeString('en-US', { hour12: false }); // E.g., "14:30:00"

    // Prevent multiple check-ins/outs for the same type per day
    const existingEntries = attendanceSheet.getDataRange().getValues()
      .filter(row => row[ATTENDANCE_COLS.USERNAME]?.toString() === username &&
        new Date(row[ATTENDANCE_COLS.DATE]).toLocaleDateString() === todayDate.toLocaleDateString() &&
        row[ATTENDANCE_COLS.TYPE]?.toString() === type);

    if (existingEntries.length > 0) {
      return { success: false, message: `You have already ${type} for today.` };
    }

    // Determine status (simplified, a real system would compare with shift times)
    let status = 'On-time';
    if (type === 'Check-in' && now.getHours() > 9) { // Example: late after 9 AM
      status = 'Late';
      // If a 'Late' check-in prevents successful login, this would need to be handled client-side.
      // Here, it just records the status.
    }

    const attendanceId = generateUuid();
    attendanceSheet.appendRow([
      attendanceId,
      employeeId,
      username,
      todayDate.toISOString(), // Storing as ISO string to preserve full date/time info if needed
      currentTime,
      type,
      geoLocation,
      status
    ]);

    return { success: true, message: `${type} recorded successfully!`, status: status, time: currentTime };
  } catch (e) {
    console.error("Error recording attendance:", e);
    return { success: false, message: `Failed to record attendance: ${e.message}` };
  }
}

/**
 * Retrieves monthly attendance data for a specific user.
 * @param {string} username The username.
 * @param {number} year The year for attendance data.
 * @param {number} month The month (1-12) for attendance data.
 * @returns {object} Monthly attendance data.
 */
function getMonthlyAttendance(username, year, month) {
  try {
    const attendanceSheet = getSheet(SHEET_NAMES.ATTENDANCE);
    const attendanceData = attendanceSheet.getDataRange().getValues();

    const monthlyData = {};
    attendanceData.forEach(row => {
      if (row[ATTENDANCE_COLS.USERNAME]?.toString() === username) {
        const entryDate = new Date(row[ATTENDANCE_COLS.DATE]);
        if (entryDate.getFullYear() === year && entryDate.getMonth() + 1 === month) {
          const dateKey = entryDate.toISOString().split('T')[0]; // YYYY-MM-DD
          if (!monthlyData[dateKey]) {
            monthlyData[dateKey] = {
              date: dateKey,
              checkIn: null,
              checkOut: null,
              status: 'Unknown' // Will be updated
            };
          }
          if (row[ATTENDANCE_COLS.TYPE]?.toString() === 'Check-in') {
            monthlyData[dateKey].checkIn = row[ATTENDANCE_COLS.TIME];
            monthlyData[dateKey].status = row[ATTENDANCE_COLS.STATUS]; // Use status from check-in
          } else if (row[ATTENDANCE_COLS.TYPE]?.toString() === 'Check-out') {
            monthlyData[dateKey].checkOut = row[ATTENDANCE_COLS.TIME];
            // If check-out is recorded, main status is typically 'Present'
            // A more complex system would reconcile check-in status (Late) with check-out
            monthlyData[dateKey].status = 'Present';
          }
        }
      }
    });

    // Determine final daily status (Present, Late, Absent, On Leave)
    const finalAttendance = Object.values(monthlyData).map(entry => {
      let dailyStatus = '';
      if (entry.checkIn && entry.checkOut) {
        dailyStatus = 'Present';
        // If the original check-in status was 'Late', prefer that
        if (entry.status === 'Late') {
          dailyStatus = 'Late';
        }
      } else if (entry.checkIn) {
        dailyStatus = 'Checked In (No Check-out)'; // Still "in good standing" but needs attention
        // If the original check-in status was 'Late', prefer that
        if (entry.status === 'Late') {
          dailyStatus = 'Late';
        }
      } else {
        dailyStatus = 'Absent'; // If no check-in/out for the day
      }

      // INTEGRATE LEAVE DATA: Check if user was on leave for this day
      try {
        const leaveSheet = getSheet(SHEET_NAMES.LEAVES);
        const leaveData = leaveSheet.getDataRange().getValues();
        const userApprovedLeaves = leaveData.filter(row =>
          row[LEAVE_COLS.USERNAME]?.toString() === username &&
          row[LEAVE_COLS.STATUS]?.toString() === 'Approved'
        );

        const currentDay = new Date(entry.date);
        for (const leave of userApprovedLeaves) {
          const leaveStartDate = new Date(leave[LEAVE_COLS.START_DATE]);
          leaveStartDate.setHours(0,0,0,0);
          const leaveEndDate = new Date(leave[LEAVE_COLS.END_DATE]);
          leaveEndDate.setHours(0,0,0,0);

          if (currentDay >= leaveStartDate && currentDay <= leaveEndDate) {
            dailyStatus = 'On Leave';
            break; // Found a leave for this day
          }
        }
      } catch (leaveError) {
        console.warn(`Could not integrate leave data for attendance calendar: ${leaveError.message}`);
      }


      return {
        date: entry.date,
        status: dailyStatus,
        checkIn: entry.checkIn,
        checkOut: entry.checkOut
      };
    });

    return { success: true, data: finalAttendance };
  } catch (e) {
    console.error("Error getting monthly attendance:", e);
    return { success: false, message: `Failed to retrieve attendance: ${e.message}` };
  }
}

/**
 * Retrieves leave balances for a specific user.
 * @param {string} username The username.
 * @returns {object} Leave balance data.
 */
function getLeaveBalances(username) {
  try {
    const leaveSheet = getSheet(SHEET_NAMES.LEAVES);
    const leaveData = leaveSheet.getDataRange().getValues();
    const userApprovedLeaves = leaveData.filter(row =>
      row[LEAVE_COLS.USERNAME]?.toString() === username &&
      row[LEAVE_COLS.STATUS]?.toString() === 'Approved'
    );

    const currentYear = new Date().getFullYear();

    // Define total annual leave types and quotas (can be from AppConfig or another sheet)
    // For simplicity, hardcoded here. In a real system, these would be configurable.
    const totalAnnualLeaves = {
      'Casual Leave': 15,
      'Sick Leave': 10,
      'Annual Leave': 20,
      'Maternity Leave': 90, // Example
      'Paternity Leave': 15 // Example
    };

    const usedLeaves = {};
    for (const type in totalAnnualLeaves) {
      usedLeaves[type] = 0;
    }

    userApprovedLeaves.forEach(leave => {
      const leaveType = leave[LEAVE_COLS.LEAVE_TYPE]?.toString();
      const leaveStartDate = new Date(leave[LEAVE_COLS.START_DATE]);

      // Only count leaves for the current calendar year
      if (leaveStartDate.getFullYear() === currentYear) {
        const numDays = parseFloat(leave[LEAVE_COLS.NUM_DAYS]) || 0;
        if (usedLeaves.hasOwnProperty(leaveType)) {
          usedLeaves[leaveType] += numDays;
        }
      }
    });

    const leaveBalances = [];
    for (const type in totalAnnualLeaves) {
      if (totalAnnualLeaves.hasOwnProperty(type)) {
        leaveBalances.push({
          type: type,
          total: totalAnnualLeaves[type],
          used: usedLeaves[type],
          remaining: totalAnnualLeaves[type] - usedLeaves[type]
        });
      }
    }

    return { success: true, data: leaveBalances };
  } catch (e) {
    console.error("Error getting leave balances:", e);
    return { success: false, message: `Failed to retrieve leave balances: ${e.message}` };
  }
}

/**
 * Submits a leave application.
 * @param {string} username The applying user's username.
 * @param {object} leaveData Object containing leave details (type, startDate, endDate, reason, attachmentUrl).
 * @returns {object} Result of the operation.
 */
function applyLeave(username, leaveData) {
  try {
    const leaveSheet = getSheet(SHEET_NAMES.LEAVES);
    const userDetails = getUserDetailsByUsername(username);
    const employeeId = userDetails ? userDetails.employeeId : 'N/A';

    const startDate = new Date(leaveData.startDate);
    const endDate = new Date(leaveData.endDate);
    const timeDiff = Math.abs(endDate.getTime() - startDate.getTime());
    const diffDays = Math.ceil(timeDiff / (1000 * 3600 * 24)) + 1; // +1 to include both start and end day

    if (diffDays <= 0 || isNaN(diffDays)) {
      return { success: false, message: 'Invalid start or end date.' };
    }

    // Simple validation: ensure start date is not in the past (only future or today)
    // Make sure to zero out time components for date-only comparison
    const today = new Date();
    today.setHours(0,0,0,0);
    const startOfDayStartDate = new Date(startDate);
    startOfDayStartDate.setHours(0,0,0,0);

    if(startOfDayStartDate < today) {
      return { success: false, message: 'Leave start date cannot be in the past.' };
    }

    const leaveId = generateUuid();
    leaveSheet.appendRow([
      leaveId,
      employeeId,
      username,
      leaveData.leaveType,
      startDate.toISOString().split('T')[0], // YYYY-MM-DD
      endDate.toISOString().split('T')[0],  // YYYY-MM-DD
      diffDays,
      leaveData.reason,
      leaveData.attachmentUrl || '',
      'Pending', // Default status
      new Date().toISOString(),
      '', // ApprovedBy
      '', // ApprovedAt
      ''  // DeniedReason
    ]);

    // OPTIONAL: Send email notification to HR/Admin about new leave application
    // Consider adding a configuration flag to enable/disable this in AppConfig
    /*
    MailApp.sendEmail({
      to: 'hr@yourcompany.com', // Replace with dynamic HR email from config or Users sheet
      subject: `New Leave Application from ${username}`,
      htmlBody: `<html><body>
          <p>${username} (Employee ID: ${employeeId}) has applied for ${leaveData.leaveType} from ${startDate.toLocaleDateString()} to ${endDate.toLocaleDateString()}.</p>
          <p>Reason: ${leaveData.reason}</p>
          <p>Please review this application in the HR system.</p>
          </body></html>`
    });
    */

    return { success: true, message: 'Leave application submitted successfully! Waiting for approval.' };
  } catch (e) {
    console.error("Error applying for leave:", e);
    return { success: false, message: `Failed to apply for leave: ${e.message}` };
  }
}

/**
 * Retrieves a user's leave history.
 * @param {string} username The username.
 * @returns {object} User's leave history.
 */
function getUserLeaveHistory(username) {
  try {
    const leaveSheet = getSheet(SHEET_NAMES.LEAVES);
    const allLeaves = leaveSheet.getDataRange().getValues();

    const userLeaves = allLeaves.filter(row => row[LEAVE_COLS.USERNAME]?.toString() === username)
      .map(row => ({
        leaveId: row[LEAVE_COLS.LEAVE_ID]?.toString(),
        leaveType: row[LEAVE_COLS.LEAVE_TYPE]?.toString(),
        startDate: row[LEAVE_COLS.START_DATE]?.toString(),
        endDate: row[LEAVE_COLS.END_DATE]?.toString(),
        numDays: parseFloat(row[LEAVE_COLS.NUM_DAYS]) || 0,
        reason: row[LEAVE_COLS.REASON]?.toString(),
        status: row[LEAVE_COLS.STATUS]?.toString(),
        appliedAt: row[LEAVE_COLS.APPLIED_AT]?.toString()
      }));

    return { success: true, data: userLeaves };
  } catch (e) {
    console.error("Error getting user leave history:", e);
    return { success: false, message: `Failed to retrieve leave history: ${e.message}` };
  }
}

/**
 * Changes a user's password.
 * @param {string} username The username.
 * @param {string} oldPassword The current password.
 * @param {string} newPassword The new password.
 * @returns {object} Result of the operation.
 */
function changePassword(username, oldPassword, newPassword) {
  try {
    const usersSheet = getSheet(SHEET_NAMES.USERS);
    const usersData = usersSheet.getDataRange().getValues();

    const userRowIndex = usersData.findIndex(row => row[USER_COLS.USERNAME]?.toString().toLowerCase() === username.toLowerCase());

    if (userRowIndex === -1) {
      return { success: false, message: 'User not found.' };
    }

    const storedHashedPassword = usersData[userRowIndex][USER_COLS.PASSWORD_HASH]?.toString();
    const hashedOldPassword = hashPassword(oldPassword);

    if (storedHashedPassword !== hashedOldPassword) {
      return { success: false, message: 'Incorrect old password.' };
    }

    const hashedNewPassword = hashPassword(newPassword);
    usersSheet.getRange(userRowIndex + 1, USER_COLS.PASSWORD_HASH + 1).setValue(hashedNewPassword);

    return { success: true, message: 'Password changed successfully!' };

  } catch (e) {
    console.error("Error changing password:", e);
    return { success: false, message: `Failed to change password: ${e.message}` };
  }
}

/**
 * Updates a user's profile information. This updates the Employees sheet.
 * @param {string} username The username of the employee to update.
 * @param {object} profileData Object containing fields to update (e.g., fullName, email, phoneNumber).
 * @returns {object} Success status and message.
 */
function updateProfile(username, profileData) {
  try {
    const employeeSheet = getSheet(SHEET_NAMES.EMPLOYEES);
    const employeeData = employeeSheet.getDataRange().getValues();
    const employeeRowIndex = employeeData.findIndex(row => row[EMPLOYEE_COLS.USERNAME]?.toString().toLowerCase() === username.toLowerCase());

    if (employeeRowIndex === -1) {
      return { success: false, message: 'Employee record not found. Please contact HR.' };
    }

    const targetRow = employeeRowIndex + 1; // 1-indexed
    const currentRowData = employeeData[employeeRowIndex]; // Current data for reference

    // Update fields if provided in profileData AND they are actual changes
    if (profileData.fullName !== undefined && profileData.fullName !== currentRowData[EMPLOYEE_COLS.FULL_NAME]) {
      employeeSheet.getRange(targetRow, EMPLOYEE_COLS.FULL_NAME + 1).setValue(profileData.fullName);
    }
    if (profileData.email !== undefined && profileData.email !== currentRowData[EMPLOYEE_COLS.EMAIL]) {
      employeeSheet.getRange(targetRow, EMPLOYEE_COLS.EMAIL + 1).setValue(profileData.email);
        // Also update email in Users sheet if it changes
        const usersSheet = getSheet(SHEET_NAMES.USERS);
        const usersData = usersSheet.getDataRange().getValues();
        const userRowIndex = usersData.findIndex(row => row[USER_COLS.USERNAME]?.toString().toLowerCase() === username.toLowerCase());
        if (userRowIndex !== -1) {
            usersSheet.getRange(userRowIndex + 1, USER_COLS.EMAIL + 1).setValue(profileData.email);
        }
    }
    if (profileData.phoneNumber !== undefined && profileData.phoneNumber !== currentRowData[EMPLOYEE_COLS.PHONE_NUMBER]) {
      employeeSheet.getRange(targetRow, EMPLOYEE_COLS.PHONE_NUMBER + 1).setValue(profileData.phoneNumber);
    }
    // Add other editable fields as needed

    return { success: true, message: 'Profile updated successfully!' };
  } catch (e) {
    console.error("Error updating profile:", e);
    return { success: false, message: `Failed to update profile: ${e.message}` };
  }
}

/**
 * Retrieves tasks assigned to a specific user.
 * @param {string} username The assigned user's username.
 * @param {string} filter Optional filter ('all', 'pending', 'completed').
 * @returns {object} List of tasks.
 */
function getUserTasks(username, filter = 'all') {
  try {
    const tasksSheet = getSheet(SHEET_NAMES.TASKS);
    const allTasks = tasksSheet.getDataRange().getValues();

    let userTasks = allTasks.filter(row => row[TASK_COLS.ASSIGNED_TO_USERNAME]?.toString() === username); // Ensure string comparison

    if (filter === 'pending') {
      userTasks = userTasks.filter(row => row[TASK_COLS.STATUS] === 'Pending' || row[TASK_COLS.STATUS] === 'In Progress' || row[TASK_COLS.STATUS] === 'Submitted');
    } else if (filter === 'completed') {
      userTasks = userTasks.filter(row => row[TASK_COLS.STATUS] === 'Completed' || row[TASK_COLS.STATUS] === 'Approved');
    }

    const tasks = userTasks.map(row => ({
      taskId: row[TASK_COLS.TASK_ID]?.toString(),
      assignedTo: row[TASK_COLS.ASSIGNED_TO_USERNAME]?.toString(),
      assignedBy: row[TASK_COLS.ASSIGNED_BY_USERNAME]?.toString(),
      title: row[TASK_COLS.TITLE]?.toString(),
      description: row[TASK_COLS.DESCRIPTION]?.toString(),
      type: row[TASK_COLS.TYPE]?.toString(),
      priority: row[TASK_COLS.PRIORITY]?.toString(),
      dueDate: row[TASK_COLS.DUE_DATE]?.toString(),
      createdAt: row[TASK_COLS.CREATED_AT]?.toString(),
      status: row[TASK_COLS.STATUS]?.toString(),
      progress: parseFloat(row[TASK_COLS.PROGRESS]) || 0,
      attachments: row[TASK_COLS.ATTACHMENTS_URLS]?.toString().split(',').filter(url => url) || [],
      comments: row[TASK_COLS.COMMENTS]?.toString(),
      lastUpdated: row[TASK_COLS.LAST_UPDATED]?.toString()
    }));

    return { success: true, data: tasks.sort((a,b) => new Date(a.dueDate) - new Date(b.dueDate)) }; // Sort by due date
  } catch (e) {
    console.error("Error getting user tasks:", e);
    return { success: false, message: `Failed to retrieve tasks: ${e.message}` };
  }
}

/**
 * Updates the progress or status of a specific task.
 * @param {string} username The user updating the task.
 * @param {string} taskId The ID of the task.
 * @param {number} progress The new progress percentage (0-100).
 * @param {string} comments Optional comments.
 * @param {string} status Optional new status ('In Progress', 'Submitted').
 * @returns {object} Result of the operation.
 */
function updateTaskProgress(username, taskId, progress, comments, status) {
  try {
    const tasksSheet = getSheet(SHEET_NAMES.TASKS);
    const tasksData = tasksSheet.getDataRange().getValues();

    const taskRowIndex = tasksData.findIndex(row => row[TASK_COLS.TASK_ID]?.toString() === taskId);

    if (taskRowIndex === -1) {
      return { success: false, message: 'Task not found.' };
    }

    const targetRow = taskRowIndex + 1;
    let currentStatus = tasksData[taskRowIndex][TASK_COLS.STATUS]?.toString();

    // Only allow assigned user to update
    if (tasksData[taskRowIndex][TASK_COLS.ASSIGNED_TO_USERNAME]?.toString() !== username) {
      return { success: false, message: 'You are not authorized to update this task.' };
    }

    // Update progress and status based on action
    if (progress !== undefined && progress !== null) {
      tasksSheet.getRange(targetRow, TASK_COLS.PROGRESS + 1).setValue(Math.min(100, Math.max(0, parseInt(progress))));
    }
    if (comments !== undefined && comments !== null) {
      // Append new comments with timestamp and user
      const existingComments = tasksData[taskRowIndex][TASK_COLS.COMMENTS]?.toString() || '';
      const newCommentEntry = `[${new Date().toLocaleString()}] (${username}): ${comments}\n`;
      tasksSheet.getRange(targetRow, TASK_COLS.COMMENTS + 1).setValue(existingComments + newCommentEntry);
    }
    if (status) {
      // Restrict status progression: Pending -> In Progress -> Submitted
      if (status === 'In Progress' && currentStatus === 'Pending' ||
          status === 'Submitted' && (currentStatus === 'In Progress' || currentStatus === 'Pending')) {
        tasksSheet.getRange(targetRow, TASK_COLS.STATUS + 1).setValue(status);
        currentStatus = status; // Update internal currentStatus
      } else if (status === 'Completed' || status === 'Approved' || status === 'Rejected') {
        // These statuses are typically set by Admin/Manager
        return { success: false, message: 'Invalid status update. Only Admin/Manager can set task to Completed, Approved or Rejected.' };
      }
    }

    // Auto-set status to 'Submitted' if progress is 100% and not yet submitted/approved/rejected
    const newProgressValue = tasksSheet.getRange(targetRow, TASK_COLS.PROGRESS + 1).getValue();
    if(newProgressValue == 100 && (currentStatus === 'In Progress' || currentStatus === 'Pending')) {
      tasksSheet.getRange(targetRow, TASK_COLS.STATUS + 1).setValue('Submitted');
      currentStatus = 'Submitted';
    }


    tasksSheet.getRange(targetRow, TASK_COLS.LAST_UPDATED + 1).setValue(new Date().toISOString());

    return { success: true, message: 'Task updated successfully!', newStatus: currentStatus };

  } catch (e) {
    console.error("Error updating task progress:", e);
    return { success: false, message: `Failed to update task: ${e.message}` };
  }
}

/**
 * Gets all active announcements.
 * @returns {object} List of announcements.
 */
function getAnnouncements() {
  try {
    const announcementsSheet = getSheet(SHEET_NAMES.ANNOUNCEMENTS);
    const allAnnouncements = announcementsSheet.getDataRange().getValues();
    const now = new Date().getTime();

    const activeAnnouncements = allAnnouncements.filter(row => {
      const validUntil = row[ANNOUNCEMENT_COLS.VALID_UNTIL];
      // Filter active (valid_until is empty OR valid_until is in the future)
      return !validUntil || new Date(validUntil).getTime() >= now;
    }).map(row => ({
      announcementId: row[ANNOUNCEMENT_COLS.ANNOUNCEMENT_ID]?.toString(),
      title: row[ANNOUNCEMENT_COLS.TITLE]?.toString(),
      content: row[ANNOUNCEMENT_COLS.CONTENT]?.toString(),
      createdBy: row[ANNOUNCEMENT_COLS.CREATED_BY_USERNAME]?.toString(),
      createdAt: row[ANNOUNCEMENT_COLS.CREATED_AT]?.toString(),
      validUntil: row[ANNOUNCEMENT_COLS.VALID_UNTIL]?.toString()
    }));

    return { success: true, data: activeAnnouncements.sort((a,b) => new Date(b.createdAt) - new Date(a.createdAt)) }; // Newest first
  } catch (e) {
    console.error("Error getting announcements:", e);
    return { success: false, message: `Failed to retrieve announcements: ${e.message}` };
  }
}

/**
 * Retrieves payslip information for a specific user. (Simplified, links to documents)
 * @param {string} username The username.
 * @returns {object} List of payslips.
 */
function getPayslips(username) {
  try {
    const payslipsSheet = getSheet(SHEET_NAMES.PAYSLIPS);
    const allPayslips = payslipsSheet.getDataRange().getValues();

    const userPayslips = allPayslips.filter(row => row[PAYSLIP_COLS.USERNAME]?.toString() === username)
      .map(row => ({
        payslipId: row[PAYSLIP_COLS.PAYSLIP_ID]?.toString(),
        monthYear: row[PAYSLIP_COLS.MONTH_YEAR]?.toString(),
        documentUrl: row[PAYSLIP_COLS.DOCUMENT_URL]?.toString(),
        uploadedAt: row[PAYSLIP_COLS.UPLOADED_AT]?.toString(),
        comment: row[PAYSLIP_COLS.COMMENT]?.toString()
      }));

    return { success: true, data: userPayslips.sort((a,b) => new Date(b.uploadedAt) - new Date(a.uploadedAt)) };
  } catch (e) {
    console.error("Error getting payslips:", e);
    return { success: false, message: `Failed to retrieve payslips: ${e.message}` };
  }
}

/**
 * Retrieves company policies from the Policies sheet.
 * @returns {object} List of policies.
 */
function getCompanyPolicies() {
  try {
    const policiesSheet = getSheet(SHEET_NAMES.POLICIES);
    const allPolicies = policiesSheet.getDataRange().getValues();

    const policies = allPolicies.map(row => ({
      policyId: row[POLICIES_COLS.POLICY_ID]?.toString(), // Assuming ID is first column
      title: row[POLICIES_COLS.TITLE]?.toString(),
      description: row[POLICIES_COLS.DESCRIPTION]?.toString(),
      documentUrl: row[POLICIES_COLS.DOCUMENT_URL]?.toString(),
      uploadedAt: row[POLICIES_COLS.UPLOADED_AT]?.toString()
    }));

    return { success: true, data: policies };
  } catch (e) {
    console.error("Error getting company policies:", e);
    return { success: false, message: `Failed to retrieve policies: ${e.message}` };
  }
}

// ===============================================
// --- ADMIN/HR SPECIFIC FUNCTIONS ---
// ===============================================

/**
 * Retrieves all employee data from the Employees sheet.
 * @returns {object} List of all employee records.
 */
function getAllEmployeesData() {
  try {
    const employeesSheet = getSheet(SHEET_NAMES.EMPLOYEES);
    const allEmployees = employeesSheet.getDataRange().getValues();

    const employees = allEmployees.map(row => ({
      employeeId: row[EMPLOYEE_COLS.EMPLOYEE_ID]?.toString(),
      username: row[EMPLOYEE_COLS.USERNAME]?.toString(),
      fullName: row[EMPLOYEE_COLS.FULL_NAME]?.toString(),
      email: row[EMPLOYEE_COLS.EMAIL]?.toString(),
      phoneNumber: row[EMPLOYEE_COLS.PHONE_NUMBER]?.toString(),
      department: row[EMPLOYEE_COLS.DEPARTMENT]?.toString(),
      role: row[EMPLOYEE_COLS.ROLE]?.toString(),
      joinDate: row[EMPLOYEE_COLS.JOIN_DATE]?.toString(),
      status: row[EMPLOYEE_COLS.STATUS]?.toString(),
      profilePicUrl: row[EMPLOYEE_COLS.PUBLIC_PROFILE_PIC_URL]?.toString()
    }));

    return { success: true, data: employees };
  } catch (e) {
    console.error("Error getting all employees:", e);
    return { success: false, message: `Failed to retrieve employee data: ${e.message}` };
  }
}

/**
 * Adds a new employee or updates an existing one on the Employees and Users sheets.
 * @param {object} employeeData Object containing employee details. Must include username.
 * @returns {object} Result of the operation.
 */
function addUpdateEmployee(employeeData) {
  try {
    const employeesSheet = getSheet(SHEET_NAMES.EMPLOYEES);
    const usersSheet = getSheet(SHEET_NAMES.USERS); // Also update/create user login
    const allEmployees = employeesSheet.getDataRange().getValues();
    const allUsers = usersSheet.getDataRange().getValues();

    const username = (employeeData.username || '').toLowerCase(); // Ensure username is string and lowercase
    if (!username) {
        return { success: false, message: 'Username is required.' };
    }

    const employeeRowIndex = allEmployees.findIndex(row => row[EMPLOYEE_COLS.USERNAME]?.toString().toLowerCase() === username);
    const userRowIndex = allUsers.findIndex(row => row[USER_COLS.USERNAME]?.toString().toLowerCase() === username);

    // Generate EmployeeID if not provided or for new employee
    const employeeId = (employeeData.employeeId && String(employeeData.employeeId).trim() !== ''
        ? String(employeeData.employeeId)
        : `EMP-${(allEmployees.length + 1).toString().padStart(4, '0')}`);

    if (employeeRowIndex === -1) {
      // Add new employee
      employeesSheet.appendRow([
        employeeId,
        username,
        employeeData.fullName || '',
        employeeData.email || '',
        employeeData.phoneNumber || '',
        employeeData.department || '',
        employeeData.role || USER_ROLES.EMPLOYEE,
        employeeData.joinDate || new Date().toISOString().split('T')[0],
        employeeData.status || 'Active',
        employeeData.profilePicUrl || ''
      ]);

      // Create new user login if username doesn't exist in Users sheet
      if (userRowIndex === -1) {
        const defaultPassword = hashPassword('Pa$$word1'); // Set a strong default password
        usersSheet.appendRow([
          username,
          defaultPassword,
          employeeData.role || USER_ROLES.EMPLOYEE,
          employeeId,
          employeeData.email || '',
          '', // OTP
          '', // OTP Exp
          employeeData.profilePicUrl || ''
        ]);
        return { success: true, message: `Employee "${username}" added successfully. Default password for login is "Pa$$word1".` };
      } else {
        // User exists in Users sheet, but not in Employees sheet. Update user's EmployeeID etc.
        const currentUserData = allUsers[userRowIndex];
        usersSheet.getRange(userRowIndex + 1, USER_COLS.ROLE + 1).setValue(employeeData.role || currentUserData[USER_COLS.ROLE]);
        usersSheet.getRange(userRowIndex + 1, USER_COLS.EMPLOYEE_ID + 1).setValue(employeeId);
        usersSheet.getRange(userRowIndex + 1, USER_COLS.EMAIL + 1).setValue(employeeData.email || currentUserData[USER_COLS.EMAIL]);
        return { success: true, message: `Employee "${username}" details added. User login updated.` };
      }
    } else {
      // Update existing employee
      const targetEmpRow = employeeRowIndex + 1;
      const currentEmpData = allEmployees[employeeRowIndex];

      employeesSheet.getRange(targetEmpRow, EMPLOYEE_COLS.FULL_NAME + 1).setValue(employeeData.fullName || currentEmpData[EMPLOYEE_COLS.FULL_NAME]);
      employeesSheet.getRange(targetEmpRow, EMPLOYEE_COLS.EMAIL + 1).setValue(employeeData.email || currentEmpData[EMPLOYEE_COLS.EMAIL]);
      employeesSheet.getRange(targetEmpRow, EMPLOYEE_COLS.PHONE_NUMBER + 1).setValue(employeeData.phoneNumber || currentEmpData[EMPLOYEE_COLS.PHONE_NUMBER]);
      employeesSheet.getRange(targetEmpRow, EMPLOYEE_COLS.DEPARTMENT + 1).setValue(employeeData.department || currentEmpData[EMPLOYEE_COLS.DEPARTMENT]);
      employeesSheet.getRange(targetEmpRow, EMPLOYEE_COLS.ROLE + 1).setValue(employeeData.role || currentEmpData[EMPLOYEE_COLS.ROLE]);
      employeesSheet.getRange(targetEmpRow, EMPLOYEE_COLS.JOIN_DATE + 1).setValue(employeeData.joinDate || currentEmpData[EMPLOYEE_COLS.JOIN_DATE]);
      employeesSheet.getRange(targetEmpRow, EMPLOYEE_COLS.STATUS + 1).setValue(employeeData.status || currentEmpData[EMPLOYEE_COLS.STATUS]);
      employeesSheet.getRange(targetEmpRow, EMPLOYEE_COLS.PUBLIC_PROFILE_PIC_URL + 1).setValue(employeeData.profilePicUrl || currentEmpData[EMPLOYEE_COLS.PUBLIC_PROFILE_PIC_URL]);

      // Also update Users sheet (role, email, employeeId)
      if (userRowIndex !== -1) {
        const currentUserData = allUsers[userRowIndex];
        usersSheet.getRange(userRowIndex + 1, USER_COLS.EMAIL + 1).setValue(employeeData.email || currentUserData[USER_COLS.EMAIL]);
        usersSheet.getRange(userRowIndex + 1, USER_COLS.ROLE + 1).setValue(employeeData.role || currentUserData[USER_COLS.ROLE]);
        usersSheet.getRange(userRowIndex + 1, USER_COLS.EMPLOYEE_ID + 1).setValue(employeeId || currentUserData[USER_COLS.EMPLOYEE_ID]);
      } else {
        // If employee exists but no user entry (e.g., created directly in sheet previously), create one now.
        const defaultPassword = hashPassword('Pa$$word1');
        usersSheet.appendRow([username, defaultPassword, employeeData.role || USER_ROLES.EMPLOYEE, employeeId, employeeData.email || '', '', '', employeeData.profilePicUrl || '']);
        return { success: true, message: `Employee "${username}" updated and a new user login created. Default password is "Pa$$word1".` };
      }

      return { success: true, message: `Employee "${username}" updated successfully!` };
    }
  } catch (e) {
    console.error("Error adding/updating employee:", e);
    return { success: false, message: `Failed to add/update employee: ${e.message}` };
  }
}

/**
 * Deletes an employee and their corresponding user login.
 * @param {string} username The username of the employee to delete.
 * @returns {object} Result of the operation.
 */
function deleteEmployee(username) {
  try {
    const employeesSheet = getSheet(SHEET_NAMES.EMPLOYEES);
    const usersSheet = getSheet(SHEET_NAMES.USERS);

    let employeeDeleted = false;
    let userDeleted = false;
    let employeeProfilePicDeleted = false;

    const employeesData = employeesSheet.getDataRange().getValues();
    const employeeRowIndex = employeesData.findIndex(row => row[EMPLOYEE_COLS.USERNAME]?.toString().toLowerCase() === username.toLowerCase());

    if (employeeRowIndex !== -1) {
      // Get profile pic URL before deletion for cleanup
      const profilePicUrlToDelete = employeesData[employeeRowIndex][EMPLOYEE_COLS.PUBLIC_PROFILE_PIC_URL]?.toString();
      if (profilePicUrlToDelete && profilePicUrlToDelete.includes('drive.google.com/open?id=')) {
        const fileId = profilePicUrlToDelete.split('id=')[1];
        try {
          DriveApp.getFileById(fileId).setTrashed(true);
          employeeProfilePicDeleted = true;
        } catch (e) {
          console.warn(`Could not trash profile picture file ID ${fileId}: ${e.message}`);
        }
      }
      employeesSheet.deleteRow(employeeRowIndex + 1);
      employeeDeleted = true;
    }

    const usersData = usersSheet.getDataRange().getValues();
    const userRowIndex = usersData.findIndex(row => row[USER_COLS.USERNAME]?.toString().toLowerCase() === username.toLowerCase());

    if (userRowIndex !== -1) {
      usersSheet.deleteRow(userRowIndex + 1);
      userDeleted = true;
    }

    if (employeeDeleted || userDeleted) {
      return { success: true, message: `Employee "${username}" and associated user/data deleted. Profile picture trashed: ${employeeProfilePicDeleted}` };
    } else {
      return { success: false, message: `Employee "${username}" not found.` };
    }

  } catch (e) {
    console.error("Error deleting employee:", e);
    return { success: false, message: `Failed to delete employee: ${e.message}` };
  }
}

/**
 * Gets overall attendance summary for admin/HR for a given date or filter by employee.
 * @param {string} dateString Iso date string (YYYY-MM-DD or full ISO)
 * @param {string} employeeIdFilter Optional filter for a specific employee. (Not fully implemented in UI)
 * @returns {object} Aggregated attendance data.
 */
function getAdminAttendanceSummary(dateString, employeeIdFilter) {
  try {
    const attendanceSheet = getSheet(SHEET_NAMES.ATTENDANCE);
    const employeesSheet = getSheet(SHEET_NAMES.EMPLOYEES);

    const allAttendance = attendanceSheet.getDataRange().getValues();
    const allEmployees = employeesSheet.getDataRange().getValues();

    let targetDate = new Date(dateString);
    targetDate.setHours(0, 0, 0, 0); // Normalize to start of day

    const dailyAttendanceMap = new Map(); // Key by username

    allAttendance.forEach(row => {
      const entryDate = new Date(row[ATTENDANCE_COLS.DATE]);
      entryDate.setHours(0, 0, 0, 0); // Normalize this date too

      if (entryDate.getTime() === targetDate.getTime()) {
        const username = row[ATTENDANCE_COLS.USERNAME]?.toString();
        if (!dailyAttendanceMap.has(username)) {
          dailyAttendanceMap.set(username, {
            username: username,
            employeeId: row[ATTENDANCE_COLS.EMPLOYEE_ID]?.toString(),
            checkIn: null,
            checkOut: null,
            status: 'Absent',
            geoLocationIn: null,
            geoLocationOut: null,
            latestOverallStatus: 'Absent' // Status derived from last record for UI
          });
        }

        const currentRecord = dailyAttendanceMap.get(username);
        if (row[ATTENDANCE_COLS.TYPE]?.toString() === 'Check-in') {
          currentRecord.checkIn = row[ATTENDANCE_COLS.TIME];
          currentRecord.geoLocationIn = row[ATTENDANCE_COLS.GEO_LOCATION];
          currentRecord.latestOverallStatus = row[ATTENDANCE_COLS.STATUS];

        } else if (row[ATTENDANCE_COLS.TYPE]?.toString() === 'Check-out') {
          currentRecord.checkOut = row[ATTENDANCE_COLS.TIME];
          currentRecord.geoLocationOut = row[ATTENDANCE_COLS.GEO_LOCATION];
          currentRecord.latestOverallStatus = 'Present'; // A check-out usually means Present
        }
      }
    });

    // Check for leaves on the targetDate for all employees
    const leaveSheet = getSheet(SHEET_NAMES.LEAVES);
    const leaveData = leaveSheet.getDataRange().getValues();
    const employeesOnLeave = new Set(); // Stores usernames of employees on leave today

    const targetDayStart = targetDate.getTime();
    // const targetDayEnd = targetDate.getTime() + (24 * 60 * 60 * 1000); // Not needed with direct comparison

    leaveData.forEach(leave => {
        if (leave[LEAVE_COLS.STATUS]?.toString() === 'Approved') {
            const leaveStartDate = new Date(leave[LEAVE_COLS.START_DATE]);
            leaveStartDate.setHours(0,0,0,0);
            const leaveEndDate = new Date(leave[LEAVE_COLS.END_DATE]);
            leaveEndDate.setHours(0,0,0,0);

            if (targetDayStart >= leaveStartDate.getTime() && targetDayStart <= leaveEndDate.getTime()) {
                employeesOnLeave.add(leave[LEAVE_COLS.USERNAME]?.toString());
            }
        }
    });


    // Determine final status for each employee for the day
    const summary = [];
    allEmployees.forEach(empRow => {
      const empUsername = empRow[EMPLOYEE_COLS.USERNAME]?.toString();
      const empId = empRow[EMPLOYEE_COLS.EMPLOYEE_ID]?.toString();
      const empFullName = empRow[EMPLOYEE_COLS.FULL_NAME]?.toString();
      const empDepartment = empRow[EMPLOYEE_COLS.DEPARTMENT]?.toString();

      if (employeeIdFilter && empId !== employeeIdFilter) {
        return; // Skip if filter is applied and not matching
      }

      const empRecord = dailyAttendanceMap.get(empUsername) || {
        username: empUsername,
        employeeId: empId,
        checkIn: null,
        checkOut: null,
        status: 'Absent',
        geoLocationIn: null,
        geoLocationOut: null,
        latestOverallStatus: 'Absent'
      };

      if (employeesOnLeave.has(empUsername)) {
          empRecord.status = 'On Leave';
      } else if (empRecord.checkIn && empRecord.checkOut) {
          empRecord.status = 'Present'; // Checked In and Out
          // If the check-in was specifically marked 'Late', prefer that status
          if (empRecord.latestOverallStatus === 'Late') {
            empRecord.status = 'Late';
          }
      } else if (empRecord.checkIn) {
          empRecord.status = 'Checked In'; // Checked In, but not Out yet
          // If the check-in was specifically marked 'Late', prefer that status
          if (empRecord.latestOverallStatus === 'Late') {
            empRecord.status = 'Late';
          }
      } else {
          empRecord.status = 'Absent'; // If no check-in/out for the day and not on leave
      }

      summary.push({
        id: empId,
        username: empUsername,
        fullName: empFullName,
        department: empDepartment,
        status: empRecord.status,
        checkInTime: empRecord.checkIn,
        checkOutTime: empRecord.checkOut,
        geoLocationIn: empRecord.geoLocationIn,
        geoLocationOut: empRecord.geoLocationOut
      });
    });

    return { success: true, data: summary, date: targetDate.toISOString().split('T')[0] };
  } catch (e) {
    console.error("Error getting admin attendance summary:", e);
    return { success: false, message: `Failed to retrieve attendance summary: ${e.message}` };
  }
}

/**
 * Assigns a new task to a user or team.
 * @param {string} assignedByUsername The admin/manager assigning the task.
 * @param {object} taskData Object containing task details (assignedToUsername, title, description, type, priority, dueDate, attachments).
 * @returns {object} Result of the operation.
 */
function assignTask(assignedByUsername, taskData) {
  try {
    const tasksSheet = getSheet(SHEET_NAMES.TASKS);
    const usersSheet = getSheet(SHEET_NAMES.USERS);
    const allUsers = usersSheet.getDataRange().getValues();

    // Validate assignedToUsername exists
    const assignedUsernameLower = (taskData.assignedToUsername || '').toLowerCase();
    const assignedUserExists = allUsers.some(row => row[USER_COLS.USERNAME]?.toString().toLowerCase() === assignedUsernameLower);
    if (!assignedUserExists) {
      return { success: false, message: `Assigned user "${taskData.assignedToUsername}" not found.` };
    }

    const taskId = generateUuid();
    tasksSheet.appendRow([
      taskId,
      assignedUsernameLower, // Store as lowercase for consistency
      assignedByUsername,
      taskData.title,
      taskData.description || '',
      taskData.type || 'Daily', // Default to Daily
      taskData.priority || 'Medium', // Default to Medium
      taskData.dueDate, // YYYY-MM-DD string
      new Date().toISOString(),
      'Pending', // Initial status
      0, // Progress 0%
      taskData.attachments || '', // Comma-separated URLs
      '', // Comments
      new Date().toISOString() // Last updated
    ]);

    return { success: true, message: `Task "${taskData.title}" assigned successfully to ${taskData.assignedToUsername}.` };
  } catch (e) {
    console.error("Error assigning task:", e);
    return { success: false, message: `Failed to assign task: ${e.message}` };
  }
}

/**
 * Retrieves all tasks for admin/manager view.
 * @param {string} filter Optional filter ('all', 'pending', 'in progress', 'submitted', 'approved', 'rejected').
 * @returns {object} List of tasks.
 */
function getAllTasks(filter = 'all') {
  try {
    const tasksSheet = getSheet(SHEET_NAMES.TASKS);
    const allTasks = tasksSheet.getDataRange().getValues();

    let filteredTasks = allTasks;
    if (filter !== 'all') {
      filteredTasks = allTasks.filter(row => row[TASK_COLS.STATUS]?.toString().toLowerCase() === filter.toLowerCase());
    }

    const tasks = filteredTasks.map(row => ({
      taskId: row[TASK_COLS.TASK_ID]?.toString(),
      assignedTo: row[TASK_COLS.ASSIGNED_TO_USERNAME]?.toString(),
      assignedBy: row[TASK_COLS.ASSIGNED_BY_USERNAME]?.toString(),
      title: row[TASK_COLS.TITLE]?.toString(),
      description: row[TASK_COLS.DESCRIPTION]?.toString(),
      type: row[TASK_COLS.TYPE]?.toString(),
      priority: row[TASK_COLS.PRIORITY]?.toString(),
      dueDate: row[TASK_COLS.DUE_DATE]?.toString(),
      createdAt: row[TASK_COLS.CREATED_AT]?.toString(),
      status: row[TASK_COLS.STATUS]?.toString(),
      progress: parseFloat(row[TASK_COLS.PROGRESS]) || 0,
      attachments: row[TASK_COLS.ATTACHMENTS_URLS]?.toString().split(',').filter(url => url) || [],
      comments: row[TASK_COLS.COMMENTS]?.toString(),
      lastUpdated: row[TASK_COLS.LAST_UPDATED]?.toString()
    }));

    return { success: true, data: tasks.sort((a,b) => new Date(b.createdAt) - new Date(a.createdAt)) }; // Newest first
  } catch (e) {
    console.error("Error getting all tasks:", e);
    return { success: false, message: `Failed to retrieve all tasks: ${e.message}` };
  }
}

/**
 * Approves or rejects a submitted task. (Admin/Manager action)
 * @param {string} reviewerUsername The admin/manager performing the action.
 * @param {string} taskId The ID of the task.
 * @param {'Approved'|'Rejected'} status The new status for the task.
 * @param {string} comments Optional comments.
 * @returns {object} Result of the operation.
 */
function approveRejectTask(reviewerUsername, taskId, status, comments) {
  try {
    const tasksSheet = getSheet(SHEET_NAMES.TASKS);
    const tasksData = tasksSheet.getDataRange().getValues();

    const taskRowIndex = tasksData.findIndex(row => row[TASK_COLS.TASK_ID]?.toString() === taskId);

    if (taskRowIndex === -1) {
      return { success: false, message: 'Task not found.' };
    }

    const currentStatus = tasksData[taskRowIndex][TASK_COLS.STATUS]?.toString();
    if (currentStatus !== 'Submitted' && currentStatus !== 'Completed') {
      return { success: false, message: 'Task is not in a reviewable state (Needs to be Submitted or Completed by employee).' };
    }

    const targetRow = taskRowIndex + 1;
    tasksSheet.getRange(targetRow, TASK_COLS.STATUS + 1).setValue(status);
    tasksSheet.getRange(targetRow, TASK_COLS.LAST_UPDATED + 1).setValue(new Date().toISOString());

    const existingComments = tasksData[taskRowIndex][TASK_COLS.COMMENTS]?.toString() || '';
    const reviewCommentEntry = `[${new Date().toLocaleString()}] (${reviewerUsername} - ${status}): ${comments || 'No specific comment'}\n`;
    tasksSheet.getRange(targetRow, TASK_COLS.COMMENTS + 1).setValue(existingComments + reviewCommentEntry);

    return { success: true, message: `Task "${tasksData[taskRowIndex][TASK_COLS.TITLE]}" has been ${status.toLowerCase()}.` };
  } catch (e) {
    console.error("Error approving/rejecting task:", e);
    return { success: false, message: `Failed to ${status.toLowerCase()} task: ${e.message}` };
  }
}

/**
 * Retrieves all leave applications for admin/HR review.
 * @param {string} filter Optional filter ('all', 'pending', 'approved', 'rejected').
 * @returns {object} List of leave applications.
 */
function getAllLeaveApplications(filter = 'all') {
  try {
    const leaveSheet = getSheet(SHEET_NAMES.LEAVES);
    const allLeaves = leaveSheet.getDataRange().getValues();

    let filteredLeaves = allLeaves;
    if (filter !== 'all') {
      filteredLeaves = allLeaves.filter(row => row[LEAVE_COLS.STATUS]?.toString().toLowerCase() === filter.toLowerCase());
    }

    const leaves = filteredLeaves.map(row => ({
      leaveId: row[LEAVE_COLS.LEAVE_ID]?.toString(),
      employeeId: row[LEAVE_COLS.EMPLOYEE_ID]?.toString(),
      username: row[LEAVE_COLS.USERNAME]?.toString(),
      leaveType: row[LEAVE_COLS.LEAVE_TYPE]?.toString(),
      startDate: row[LEAVE_COLS.START_DATE]?.toString(),
      endDate: row[LEAVE_COLS.END_DATE]?.toString(),
      numDays: parseFloat(row[LEAVE_COLS.NUM_DAYS]) || 0,
      reason: row[LEAVE_COLS.REASON]?.toString(),
      attachmentUrl: row[LEAVE_COLS.ATTACHMENT_URL]?.toString(),
      status: row[LEAVE_COLS.STATUS]?.toString(),
      appliedAt: row[LEAVE_COLS.APPLIED_AT]?.toString(),
      approvedBy: row[LEAVE_COLS.APPROVED_BY_USERNAME]?.toString(),
      approvedAt: row[LEAVE_COLS.APPROVED_AT]?.toString(),
      deniedReason: row[LEAVE_COLS.DENIED_REASON]?.toString()
    }));

    return { success: true, data: leaves.sort((a,b) => new Date(b.appliedAt) - new Date(a.appliedAt)) };
  } catch (e) {
    console.error("Error getting all leave applications:", e);
    return { success: false, message: `Failed to retrieve leave applications: ${e.message}` };
  }
}

/**
 * Approves or rejects a specific leave application.
 * @param {string} reviewerUsername The admin/HR approving/rejecting.
 * @param {string} leaveId The ID of the leave application.
 * @param {'Approved'|'Rejected'} status The new status.
 * @param {string} deniedReason Optional reason if rejected.
 * @returns {object} Result of the operation.
 */
function approveRejectLeaveApplication(reviewerUsername, leaveId, status, deniedReason = '') {
  try {
    const leaveSheet = getSheet(SHEET_NAMES.LEAVES);
    const leaveData = leaveSheet.getDataRange().getValues();

    const leaveRowIndex = leaveData.findIndex(row => row[LEAVE_COLS.LEAVE_ID]?.toString() === leaveId);

    if (leaveRowIndex === -1) {
      return { success: false, message: 'Leave application not found.' };
    }
    if (leaveData[leaveRowIndex][LEAVE_COLS.STATUS]?.toString() !== 'Pending') {
      return { success: false, message: 'Leave application is no longer pending.' };
    }

    const targetRow = leaveRowIndex + 1;
    leaveSheet.getRange(targetRow, LEAVE_COLS.STATUS + 1).setValue(status);
    leaveSheet.getRange(targetRow, LEAVE_COLS.APPROVED_BY_USERNAME + 1).setValue(reviewerUsername);
    leaveSheet.getRange(targetRow, LEAVE_COLS.APPROVED_AT + 1).setValue(new Date().toISOString());
    leaveSheet.getRange(targetRow, LEAVE_COLS.DENIED_REASON + 1).setValue(deniedReason);

    // OPTIONAL: Send email notification back to the employee
    /*
    const applicantUsername = leaveData[leaveRowIndex][LEAVE_COLS.USERNAME]?.toString();
    const applicantDetails = getUserDetailsByUsername(applicantUsername);
    if (applicantDetails && applicantDetails.email) {
        MailApp.sendEmail({
        to: applicantDetails.email,
        subject: `Your Leave Application for ${leaveData[leaveRowIndex][LEAVE_COLS.LEAVE_TYPE]} is ${status}`,
        htmlBody: `<html><body>
            <p>Dear ${applicantUsername},</p>
            <p>Your leave application for ${leaveData[leaveRowIndex][LEAVE_COLS.LEAVE_TYPE]} from ${new Date(leaveData[leaveRowIndex][LEAVE_COLS.START_DATE]).toLocaleDateString()} to ${new Date(leaveData[leaveRowIndex][LEAVE_COLS.END_DATE]).toLocaleDateString()} has been <strong>${status.toLowerCase()}</strong>.</p>
            ${status === 'Rejected' && deniedReason ? `<p>Reason for denial: ${deniedReason}</p>` : ''}
            <p>Regards,<br>NepHR System</p>
            </body></html>`
        });
    }
    */

    return { success: true, message: `Leave application ${status.toLowerCase()} successfully.` };
  } catch (e) {
    console.error("Error approving/rejecting leave:", e);
    return { success: false, message: `Failed to ${status.toLowerCase()} leave: ${e.message}` };
  }
}

/**
 * Creates a new company-wide announcement.
 * @param {string} createdByUsername The admin/HR creating the announcement.
 * @param {object} announcementData Object containing title, content, and optional validUntil date.
 * @returns {object} Result of the operation.
 */
function createAnnouncement(createdByUsername, announcementData) {
  try {
    const announcementsSheet = getSheet(SHEET_NAMES.ANNOUNCEMENTS);
    const announcementId = generateUuid();
    announcementsSheet.appendRow([
      announcementId,
      announcementData.title,
      announcementData.content,
      createdByUsername,
      new Date().toISOString(),
      announcementData.validUntil || '' // ISO string for valid until
    ]);
    return { success: true, message: `Announcement "${announcementData.title}" created successfully.` };
  } catch (e) {
    console.error("Error creating announcement:", e);
    return { success: false, message: `Failed to create announcement: ${e.message}` };
  }
}

/**
 * Deletes an announcement.
 * @param {string} announcementId The ID of the announcement to delete.
 * @returns {object} Result of the operation.
 */
function deleteAnnouncement(announcementId) {
  try {
    const announcementsSheet = getSheet(SHEET_NAMES.ANNOUNCEMENTS);
    const allAnnouncements = announcementsSheet.getDataRange().getValues();
    const rowIndex = allAnnouncements.findIndex(row => row[ANNOUNCEMENT_COLS.ANNOUNCEMENT_ID]?.toString() === announcementId);

    if (rowIndex === -1) {
      return { success: false, message: 'Announcement not found.' };
    }
    announcementsSheet.deleteRow(rowIndex + 1);
    return { success: true, message: 'Announcement deleted successfully.' };
  } catch (e) {
    console.error("Error deleting announcement:", e);
    return { success: false, message: `Failed to delete announcement: ${e.message}` };
  }
}

/**
 * Adds or updates a company policy.
 * @param {object} policyData Contains policyId (optional for new), title, description, documentUrl.
 * @returns {object} Result of the operation.
 */
function addUpdatePolicy(policyData) {
  try {
    const policiesSheet = getSheet(SHEET_NAMES.POLICIES);
    const allPolicies = policiesSheet.getDataRange().getValues();
    let rowIndex = -1;
    if (policyData.policyId) {
      rowIndex = allPolicies.findIndex(row => row[POLICIES_COLS.POLICY_ID]?.toString() === policyData.policyId); // Assuming PolicyID is first column
    }

    if (rowIndex === -1) {
      // Add new policy
      policiesSheet.appendRow([
        generateUuid(),
        policyData.title || '',
        policyData.description || '',
        policyData.documentUrl || '',
        new Date().toISOString()
      ]);
      return { success: true, message: `Policy "${policyData.title}" added successfully.` };
    } else {
      // Update existing policy
      const targetRow = rowIndex + 1;
      policiesSheet.getRange(targetRow, POLICIES_COLS.TITLE + 1).setValue(policyData.title || allPolicies[rowIndex][POLICIES_COLS.TITLE]);
      policiesSheet.getRange(targetRow, POLICIES_COLS.DESCRIPTION + 1).setValue(policyData.description || allPolicies[rowIndex][POLICIES_COLS.DESCRIPTION]);
      policiesSheet.getRange(targetRow, POLICIES_COLS.DOCUMENT_URL + 1).setValue(policyData.documentUrl || allPolicies[rowIndex][POLICIES_COLS.DOCUMENT_URL]);
      policiesSheet.getRange(targetRow, POLICIES_COLS.UPLOADED_AT + 1).setValue(new Date().toISOString()); // Update 'last updated'
      return { success: true, message: `Policy "${policyData.title}" updated successfully.` };
    }
  } catch (e) {
    console.error("Error adding/updating policy:", e);
    return { success: false, message: `Failed to add/update policy: ${e.message}` };
  }
}

/**
 * Deletes a company policy.
 * @param {string} policyId The ID of the policy to delete.
 * @returns {object} Result of the operation.
 */
function deletePolicy(policyId) {
  try {
    const policiesSheet = getSheet(SHEET_NAMES.POLICIES);
    const allPolicies = policiesSheet.getDataRange().getValues();
    const rowIndex = allPolicies.findIndex(row => row[POLICIES_COLS.POLICY_ID]?.toString() === policyId);

    if (rowIndex === -1) {
      return { success: false, message: 'Policy not found.' };
    }
    policiesSheet.deleteRow(rowIndex + 1);
    return { success: true, message: 'Policy deleted successfully.' };
  } catch (e) {
    console.error("Error deleting policy:", e);
    return { success: false, message: `Failed to delete policy: ${e.message}` };
  }
}

/**
 * Uploads a payslip PDF to Google Drive and records its details.
 * Also links it to a specific employee.
 * @param {GoogleAppsScript.Events.DoPost} e The event object containing file data and form fields (username, monthYear, comment).
 * @returns {object} Upload result.
 */
function uploadPayslip(e) {
  try {
    const fileBlob = e.parameters.file && e.parameters.file.length > 0 ? e.parameters.file[0] : null;
    const username = e.parameters.username?.toString();
    let monthYear = e.parameters.monthYear?.toString();
    const comment = e.parameters.comment?.toString() || '';

    if (!fileBlob || typeof fileBlob.getName !== 'function' || !username || !monthYear) {
      return { success: false, message: 'Missing file data, username, or month/year for payslip upload.' };
    }

    // Validate username exists and get employee ID
    const userDetails = getUserDetailsByUsername(username);
    if (!userDetails) {
      return { success: false, message: `User "${username}" not found.` };
    }
    const employeeId = userDetails.employeeId; // Get employeeId from Users sheet

      // Ensure monthYear is in a consistent format if needed, e.g., "YYYY-MM"
    const parsedDate = new Date(monthYear);
    if (isNaN(parsedDate.getTime())) {
        // If not a recognizable date string, use as is (e.g., "August Payroll")
    } else {
        // Format to "YYYY-MM" or "Month YYYY" for consistency
        // Note: toLocaleDateString might give different formats based on locale.
        // For consistent sheet data, consider a more strict YYYY-MM-DD or YYYY-MM conversion if monthYear is a date.
        monthYear = parsedDate.toLocaleDateString('en-US', { year: 'numeric', month: 'long' });
    }

    const folderName = "NepHR Payslips";
    let folderIterator = DriveApp.getFoldersByName(folderName);
    let folder;
    if (folderIterator.hasNext()) {
      folder = folderIterator.next();
    } else {
      folder = DriveApp.createFolder(folderName);
    }

    const uploadedFile = folder.createFile(fileBlob);
    const fileUrl = uploadedFile.getUrl();
    uploadedFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); // Publicly viewable link for access

    const payslipsSheet = getSheet(SHEET_NAMES.PAYSLIPS);
    const payslipId = generateUuid();
    payslipsSheet.appendRow([
      payslipId,
      employeeId,
      username,
      monthYear,
      fileUrl,
      new Date().toISOString(),
      comment
    ]);

    return { success: true, message: `Payslip for ${username} (${monthYear}) uploaded successfully!`, fileUrl: fileUrl };

  } catch (error) {
    console.error("Error uploading payslip:", error);
    return { success: false, message: `Error uploading payslip: ${error.message}` };
  }
}

// Provided the getLoginUrl function (used by reset_password.html and handleLogout)
function getLoginUrl() {
  return ScriptApp.getService().getUrl();
}


