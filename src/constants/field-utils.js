/**
 * Field normalization, alias matching, and date conversion utilities.
 * Used primarily by file-handler.js for CSV/XLSX import processing.
 */

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
	lda: ['lastdayofattendance', 'lastattendance', 'lastdateofattendance', 'lastdayattended'],
	campus: ['location', 'site', 'school', 'campusname']
};

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
