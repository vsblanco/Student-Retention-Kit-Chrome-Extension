// [2025-12-18 09:30 AM]
// Version: 15.0
/*
* Timestamp: 2025-12-18 09:30 AM
* Version: 15.0
*/

// Centralized configuration for the Submission Checker extension.

/**
 * An enum-like object for the different checker modes.
 */
export const CHECKER_MODES = {
    SUBMISSION: 'submission',
    MISSING: 'missing'
};

/**
 * An enum-like object for the possible states of the extension.
 */
export const EXTENSION_STATES = {
    ON: 'on',
    OFF: 'off',
    PAUSED: 'paused'
};

/**
 * An enum-like object for message types sent between components.
 */
export const MESSAGE_TYPES = {
    INSPECTION_RESULT: 'inspectionResult',
    FOUND_SUBMISSION: 'foundSubmission',
    FOUND_MISSING_ASSIGNMENTS: 'foundMissingAssignments',
    MISSING_CHECK_COMPLETED: 'missingCheckCompleted',
    REQUEST_STORED_LOGS: 'requestingStoredLogs',
    STORED_LOGS: 'storedLogs',
    TEST_CONNECTION_PA: 'test-connection-pa',
    CONNECTION_TEST_RESULT: 'connection-test-result',
    SEND_DEBUG_PAYLOAD: 'send-debug-payload',
    LOG_TO_PANEL: 'logToPanel',
    SHOW_MISSING_ASSIGNMENTS_REPORT: 'showMissingAssignmentsReport',
    SRK_CONNECTOR_ACTIVE: 'SRK_CONNECTOR_ACTIVE',
    SRK_CONNECTOR_HEARTBEAT: 'SRK_CONNECTOR_HEARTBEAT',
    SRK_OFFICE_ADDIN_CONNECTED: 'SRK_OFFICE_ADDIN_CONNECTED',
    SRK_MASTER_LIST_UPDATED: 'SRK_MASTER_LIST_UPDATED',
    SRK_MASTER_LIST_ERROR: 'SRK_MASTER_LIST_ERROR',
    SRK_SELECTED_STUDENTS: 'SRK_SELECTED_STUDENTS',
    SRK_OFFICE_USER_INFO: 'SRK_OFFICE_USER_INFO',
    SRK_SEND_IMPORT_MASTER_LIST: 'SRK_SEND_IMPORT_MASTER_LIST',
    SRK_TASKPANE_PING: 'SRK_TASKPANE_PING',
    SRK_TASKPANE_PONG: 'SRK_TASKPANE_PONG',
    SRK_MANIFEST_INJECTED: 'SRK_MANIFEST_INJECTED'
};

/**
 * An enum-like object for connection types.
 */
export const CONNECTION_TYPES = {
    POWER_AUTOMATE: 'power-automate',
    EXCEL: 'excel'
};

/**
 * Keys for data stored in chrome.storage.local.
 */
export const STORAGE_KEYS = {
    EXTENSION_STATE: 'extensionState',
    FOUND_ENTRIES: 'foundEntries',
    MASTER_ENTRIES: 'masterEntries',
    LAST_UPDATED: 'lastUpdated',
    REFERENCE_DATE: 'referenceDate',
    LOOP_STATUS: 'loopStatus',
    CONNECTIONS: 'connections',
    DEBUG_MODE: 'debugMode',
    LATEST_MISSING_REPORT: 'latestMissingReport',
    // Settings
    CHECKER_MODE: 'checkerMode',
    CONCURRENT_TABS: 'concurrentTabs',
    HIGHLIGHT_COLOR: 'highlightColor',
    CUSTOM_KEYWORD: 'customKeyword',
    LOOPER_DAYS_OUT_FILTER: 'looperDaysOutFilter',
    EMBED_IN_CANVAS: 'embedInCanvas',
    INCLUDE_ALL_ASSIGNMENTS: 'includeAllAssignments',
    CANVAS_API_CACHE: 'canvasApiCache',
    // Scan Filter Settings
    SCAN_FILTER_DAYS_OUT: 'scanFilterDaysOut',
    SCAN_FILTER_INCLUDE_FAILING: 'scanFilterIncludeFailing',
    // Master List Auto-Update Settings
    AUTO_UPDATE_MASTER_LIST: 'autoUpdateMasterList',
    // Power Automate Settings
    POWER_AUTOMATE_URL: 'powerAutomateUrl',
    // Excel Student Sync Settings
    SYNC_ACTIVE_STUDENT: 'syncActiveStudent',
    SEND_MASTER_LIST_TO_EXCEL: 'sendMasterListToExcel',
    // Highlight Student Row Payload Settings
    HIGHLIGHT_STUDENT_ROW_ENABLED: 'highlightStudentRowEnabled',
    HIGHLIGHT_START_COL: 'highlightStartCol',
    HIGHLIGHT_END_COL: 'highlightEndCol',
    HIGHLIGHT_EDIT_COLUMN: 'highlightEditColumn',
    HIGHLIGHT_EDIT_TEXT: 'highlightEditText',
    HIGHLIGHT_TARGET_SHEET: 'highlightTargetSheet',
    HIGHLIGHT_COLOR: 'highlightRowColor',
    // Specific Submission Date Settings
    USE_SPECIFIC_DATE: 'useSpecificDate',
    SPECIFIC_SUBMISSION_DATE: 'specificSubmissionDate',
    // Office User Info
    OFFICE_USER_INFO: 'officeUserInfo',
    // Auto-Sideload Manifest Settings
    AUTO_SIDELOAD_MANIFEST: 'autoSideloadManifest'
};

/**
 * Default values for all extension settings.
 */
export const DEFAULT_SETTINGS = {
    [STORAGE_KEYS.CHECKER_MODE]: CHECKER_MODES.SUBMISSION,
    [STORAGE_KEYS.CONCURRENT_TABS]: 5, // Increased default since API is faster
    [STORAGE_KEYS.HIGHLIGHT_COLOR]: '#ffff00',
    [STORAGE_KEYS.CUSTOM_KEYWORD]: '',
    [STORAGE_KEYS.LOOPER_DAYS_OUT_FILTER]: 'all',
    [STORAGE_KEYS.DEBUG_MODE]: false,
    [STORAGE_KEYS.EMBED_IN_CANVAS]: true,
    [STORAGE_KEYS.INCLUDE_ALL_ASSIGNMENTS]: false,
    [STORAGE_KEYS.SCAN_FILTER_DAYS_OUT]: '>=5',
    [STORAGE_KEYS.SCAN_FILTER_INCLUDE_FAILING]: false,
    [STORAGE_KEYS.AUTO_UPDATE_MASTER_LIST]: 'always', // Options: 'always', 'once-daily', 'never'
    [STORAGE_KEYS.SYNC_ACTIVE_STUDENT]: true, // Enable student sync from Excel add-in by default
    [STORAGE_KEYS.SEND_MASTER_LIST_TO_EXCEL]: true, // Enable sending master list to Excel by default
    // Highlight Student Row Payload Defaults
    [STORAGE_KEYS.HIGHLIGHT_STUDENT_ROW_ENABLED]: true, // Enable student row highlighting by default
    [STORAGE_KEYS.HIGHLIGHT_START_COL]: 'Student Name',
    [STORAGE_KEYS.HIGHLIGHT_END_COL]: 'Outreach',
    [STORAGE_KEYS.HIGHLIGHT_EDIT_COLUMN]: 'Outreach',
    [STORAGE_KEYS.HIGHLIGHT_EDIT_TEXT]: 'Submitted {assignment}',
    [STORAGE_KEYS.HIGHLIGHT_TARGET_SHEET]: 'LDA MM-DD-YYYY',
    [STORAGE_KEYS.HIGHLIGHT_COLOR]: '#92d050',
    // Specific Submission Date Defaults
    [STORAGE_KEYS.USE_SPECIFIC_DATE]: false,
    [STORAGE_KEYS.SPECIFIC_SUBMISSION_DATE]: null,
    // Auto-Sideload Manifest Defaults
    [STORAGE_KEYS.AUTO_SIDELOAD_MANIFEST]: true // Auto-sideload enabled by default
};

/**
 * The name for the network recovery alarm.
 */
export const NETWORK_RECOVERY_ALARM_NAME = 'network_recovery_check';

/**
 * Regular expression for parsing advanced filter queries like '>=5' or '<10'.
 */
export const ADVANCED_FILTER_REGEX = /^\s*([><]=?|=)\s*(\d+)\s*$/;

/**
 * SharePoint URL for the "SharePoint" button in the settings tab.
 */
export const SHAREPOINT_URL = "https://edukgroup3_sharepoint.com/sites/SM-StudentServices/SitePages/CollabHome.aspx";

/**
 * Canvas LMS domain URL.
 */
export const CANVAS_DOMAIN = "https://nuc.instructure.com";

/**
 * Generic avatar URL used by Canvas for users without custom avatars.
 */
export const GENERIC_AVATAR_URL = "https://nuc.instructure.com/images/messages/avatar-50.png";

/**
 * Guide resources available in the extension.
 * Each guide has a name and path to the PDF file.
 */
export const GUIDES = [
    {
        name: 'LDA Guide',
        path: '../../assets/LDA-Guide.pdf',
        icon: 'fa-book'
    },
	{
        name: 'Student Population Report Guide',
        path: '../../assets/Download-Student-Population-Report.pdf',
        icon: 'fa-book'
    }
];

/**
 * Normalizes a field name for comparison by:
 * - Converting to lowercase
 * - Removing spaces, hyphens, underscores
 * - Removing special characters
 *
 * Examples:
 * - "Student Number" → "studentnumber"
 * - "StudentNumber" → "studentnumber"
 * - "student-number" → "studentnumber"
 * - "STUDENT_NUMBER" → "studentnumber"
 *
 * @param {string} fieldName - The field name to normalize
 * @returns {string} The normalized field name
 */
export function normalizeFieldName(fieldName) {
    if (!fieldName) return '';
    return String(fieldName)
        .toLowerCase()
        .replace(/[\s\-_]/g, '') // Remove spaces, hyphens, underscores
        .replace(/[^a-z0-9]/g, ''); // Remove any remaining special characters
}

/**
 * Converts an Excel date serial number to a JavaScript Date object.
 * Excel dates are stored as the number of days since January 1, 1900.
 *
 * @param {number} excelDate - The Excel date serial number
 * @returns {Date} JavaScript Date object
 */
export function excelDateToJSDate(excelDate) {
    // Excel date serial number starts from 1/1/1900
    // Use December 30, 1899 as base to account for Excel's leap year bug
    const excelEpoch = new Date(1899, 11, 30);
    const daysOffset = excelDate;
    const jsDate = new Date(excelEpoch.getTime() + daysOffset * 24 * 60 * 60 * 1000);
    return jsDate;
}

/**
 * Formats a JavaScript Date object to MM-DD-YY format.
 *
 * @param {Date} date - JavaScript Date object
 * @returns {string} Date formatted as MM-DD-YY
 */
export function formatDateToMMDDYY(date) {
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const day = String(date.getDate()).padStart(2, '0');
    const year = String(date.getFullYear()).slice(-2);
    return `${month}-${day}-${year}`;
}

/**
 * Checks if a value appears to be an Excel date serial number.
 * Excel dates are positive integers typically in the range 1-100000.
 *
 * @param {*} value - The value to check
 * @returns {boolean} True if value appears to be an Excel date number
 */
export function isExcelDateNumber(value) {
    if (typeof value !== 'number') return false;
    if (!Number.isInteger(value)) return false;
    // Reasonable range: 1 (1/1/1900) to 100000 (~year 2173)
    if (value < 1 || value > 100000) return false;
    return true;
}

/**
 * Converts an Excel date serial number to MM-DD-YY format string.
 * If the value is not an Excel date number, returns it unchanged.
 *
 * @param {*} value - The value to convert
 * @returns {string|*} Formatted date string or original value
 */
export function convertExcelDate(value) {
    if (isExcelDateNumber(value)) {
        const jsDate = excelDateToJSDate(value);
        return formatDateToMMDDYY(jsDate);
    }
    return value;
}

/**
 * Parses a date string in various formats and returns a JavaScript Date object.
 * Supports formats like: MM-DD-YY, MM/DD/YY, MM-DD-YYYY, MM/DD/YYYY
 *
 * @param {string|number} dateValue - The date value to parse
 * @returns {Date|null} JavaScript Date object or null if parsing fails
 */
export function parseDate(dateValue) {
    if (!dateValue) return null;

    // If it's an Excel date number, convert it first
    if (isExcelDateNumber(dateValue)) {
        return excelDateToJSDate(dateValue);
    }

    // If it's already a Date object, return it
    if (dateValue instanceof Date) {
        return dateValue;
    }

    // Convert to string and try to parse
    const dateStr = String(dateValue).trim();
    if (!dateStr) return null;

    // Try parsing common date formats: MM-DD-YY, MM/DD/YY, MM-DD-YYYY, MM/DD/YYYY
    const patterns = [
        /^(\d{1,2})[-\/](\d{1,2})[-\/](\d{2})$/,     // MM-DD-YY or MM/DD/YY
        /^(\d{1,2})[-\/](\d{1,2})[-\/](\d{4})$/      // MM-DD-YYYY or MM/DD/YYYY
    ];

    for (const pattern of patterns) {
        const match = dateStr.match(pattern);
        if (match) {
            const month = parseInt(match[1], 10);
            const day = parseInt(match[2], 10);
            let year = parseInt(match[3], 10);

            // Convert 2-digit year to 4-digit year
            if (year < 100) {
                year += year < 50 ? 2000 : 1900;
            }

            // Validate month and day
            if (month >= 1 && month <= 12 && day >= 1 && day <= 31) {
                return new Date(year, month - 1, day);
            }
        }
    }

    // Try standard Date parsing as fallback
    const parsed = new Date(dateStr);
    return isNaN(parsed.getTime()) ? null : parsed;
}

/**
 * Calculates the number of days between two dates.
 * Returns the absolute difference in days, rounded down.
 *
 * @param {Date} date1 - First date
 * @param {Date} date2 - Second date
 * @returns {number} Number of days between the dates
 */
export function daysBetween(date1, date2) {
    if (!date1 || !date2) return 0;

    const msPerDay = 24 * 60 * 60 * 1000;
    const diffMs = Math.abs(date2.getTime() - date1.getTime());
    return Math.floor(diffMs / msPerDay);
}

/**
 * Calculates days since last attendance based on LDA value and reference date.
 * Used for both Excel and CSV imports.
 *
 * @param {string|number} ldaValue - The Last Day of Attendance value
 * @param {Date|number} referenceDate - Reference date (file creation date or current date)
 * @returns {number} Number of days since last attendance, or 0 if calculation fails
 */
export function calculateDaysSinceLastAttendance(ldaValue, referenceDate) {
    if (!ldaValue || !referenceDate) return 0;

    // Parse the LDA date
    const ldaDate = parseDate(ldaValue);
    if (!ldaDate) return 0;

    // Ensure referenceDate is a Date object
    let refDate = referenceDate;
    if (typeof referenceDate === 'number') {
        refDate = new Date(referenceDate);
    }
    if (!(refDate instanceof Date) || isNaN(refDate.getTime())) {
        return 0;
    }

    // Calculate days between LDA and reference date
    return daysBetween(ldaDate, refDate);
}

/**
 * Field aliases for semantic field matching in file imports and incoming payloads.
 * Maps internal field names to semantically equivalent field names (actual aliases).
 *
 * NOTE: You don't need to specify case/space variations (e.g., "StudentNumber" vs "student number")
 * as these are handled automatically by normalizeFieldName().
 * Only specify true semantic aliases here (e.g., "Student ID" is an alias for "StudentNumber").
 */
export const FIELD_ALIASES = {
    name: ['studentname', 'student'],
    phone: ['primaryphone', 'phonenumber', 'mobile', 'cell', 'cellphone', 'contact', 'telephone', 'otherphone'],
    grade: ['gradelevel', 'level'],
    StudentNumber: ['studentid', 'sisid'],
    SyStudentId: ['studentsis'],
    daysOut: ['daysinactive', 'days'],
    url: ['gradebook', 'gradeBookUrl', 'canvasUrl', 'studentUrl'],
	studentEmail: ['email', 'studentsemail', 'studentemails', 'studentsemails'],
	personalEmail: ['otheremail', 'studentspersonalemail', 'personalstudentemail', 'othermemails'],
	lda: ['lastdayofattendance', 'lastattendance', 'lastdateofattendance', 'lastdayattended']
};

/**
 * Excel Export Column Configurations
 * Master List columns used for both CSV/Excel export and Excel import payload.
 * Modify this array to customize what columns appear in exports and imports.
 */

/**
 * Master List sheet columns
 * Each object defines: { header: 'Column Name', field: 'propertyName', fallback: optional fallback field, width: optional column width }
 * width: Custom column width in characters (wch). If not specified, column will auto-fit based on content.
 */
export const MASTER_LIST_COLUMNS = [
    { header: 'Student Name', field: 'name', width: 25 },
    { header: 'Student Number', field: 'StudentNumber', width: 15 },
    { header: 'Grade Book', field: 'url', hyperlink: true, hyperlinkText: 'Grade Book', width: 12 },
    { header: 'Grade', field: 'grade', fallback: 'currentGrade', conditionalFormatting: 'grade', width: 8 },
    { header: 'Missing Assignments', field: 'missingCount', width: 10 },
    { header: 'LDA', field: 'lda', width: 12 },
    { header: 'Days Out', field: 'daysOut', width: 10 },
    { header: 'Gender', field: 'gender', hidden: true },
    { header: 'Shift', field: 'shift', width: 10 },
    { header: 'Program Version', field: 'programVersion', width: 30 },
    { header: 'SyStudentId', field: 'SyStudentId', width: 15 },
    { header: 'Phone', field: 'phone', fallback: 'primaryPhone', width: 15 },
    { header: 'Other Phone', field: 'otherPhone', width: 15 },
    { header: 'Work Phone', field: 'workPhone', hidden: true },
    { header: 'Mobile Number', field: 'mobileNumber', hidden: true },
    { header: 'Student Email', field: 'studentEmail', width: 25 },
    { header: 'Personal Email', field: 'personalEmail', width: 25 },
    { header: 'ExpStartDate', field: 'expStartDate', width: 12 },
    { header: 'AmRep', field: 'amRep', width: 15 },
    { header: 'Hold', field: 'hold', width: 10 },
    { header: 'Photo', field: 'photo', hidden: true },
    { header: 'AdSAPStatus', field: 'adSAPStatus', width: 15 }
];

/**
 * Missing Assignments sheet columns
 * Use 'student.' prefix for student fields and 'assignment.' prefix for assignment fields
 * Standardized field names: assignmentTitle, assignmentLink, submissionLink
 * Note: Some columns use hyperlink: true to indicate they should be exported as HYPERLINK formulas
 * width: Custom column width in characters (wch). If not specified, column will auto-fit based on content.
 */
export const EXPORT_MISSING_ASSIGNMENTS_COLUMNS = [
    { header: 'Student', field: 'student.name', width: 25 },
	{ header: 'Grade Book', field: 'student.url', hyperlink: true, hyperlinkText: 'Grade Book', width: 12 },
	{ header: 'Overall Grade', field: 'student.currentGrade', fallback: 'student.grade', conditionalFormatting: 'grade', width: 8 },
    { header: 'Assignment', field: 'assignment.assignmentTitle', hyperlinkField: 'assignment.assignmentLink', hyperlink: true, width: 35 },
    { header: 'Due Date', field: 'assignment.dueDate', width: 12 },
    { header: 'Score', field: 'assignment.score', width: 8 },
    { header: 'Submission', field: 'assignment.submissionLink', hyperlink: true, hyperlinkText: 'Missing', width: 12 }
];

/**
 * LDA Sheet visible columns whitelist
 * Only columns listed here will be visible in the LDA sheet
 * Columns are identified by their field name from MASTER_LIST_COLUMNS
 */
export const LDA_VISIBLE_COLUMNS = [
    'name',
    'StudentNumber',
    'url',
    'grade',
    'missingCount',
    'lda',
    'daysOut',
    'shift',
    'programVersion',
    'phone',
    'otherPhone',
    'studentEmail',
    'personalEmail'
];
